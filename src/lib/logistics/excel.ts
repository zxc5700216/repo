import * as XLSX from "xlsx";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import type {
  AProductRow,
  AWorkbookSummary,
  BWorkbookSummary,
  CWorkbookSummary,
  DWorkbookSummary,
  NamedWorkbookExportResult,
  PdfSummary,
  SaihuWorkbookSummary,
  WorkbookExportResult,
} from "@/lib/logistics/types";
import { parseNumber, toText } from "@/lib/logistics/utils";

type WorkbookWithSheets = {
  workbook: XLSX.WorkBook;
  fileName: string;
};

type ParseAWorkbookOptions = {
  skipImages?: boolean;
};

type ExcelJsMediaItem = {
  index?: number;
  buffer?: ArrayBuffer;
  extension?: string;
};

type ExcelImageExtension = "jpeg" | "png" | "gif";

function normalizeImageExtension(extension: string): ExcelImageExtension {
  const normalized = extension.toLowerCase();
  if (normalized === "jpg" || normalized === "jpeg") {
    return "jpeg";
  }
  if (normalized === "png" || normalized === "gif") {
    return normalized;
  }
  return "png";
}

export async function readWorkbook(file: File): Promise<WorkbookWithSheets> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellHTML: false, cellFormula: true });
  return { workbook, fileName: file.name };
}

function getSheetMatrix(sheet: XLSX.WorkSheet) {
  return XLSX.utils.sheet_to_json<(string | number | null)[]>(sheet, {
    header: 1,
    raw: false,
    defval: null,
  });
}

function columnNumberToName(column: number) {
  let current = column;
  let name = "";

  while (current > 0) {
    const remainder = (current - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    current = Math.floor((current - 1) / 26);
  }

  return name;
}

function buildCellRef(row: number, col: number) {
  return `${columnNumberToName(col)}${row}`;
}

function parseBoxNumberFromLabel(label: string, fallback: number) {
  const matched = label.match(/B(\d+)/iu);
  if (!matched) {
    return fallback;
  }
  const parsed = Number(matched[1]);
  return Number.isFinite(parsed) ? parsed : fallback;
}

async function resolveRowImage(row: AProductRow) {
  if (row.imageAsset) {
    return {
      buffer: row.imageAsset.data,
      extension: normalizeImageExtension(row.imageAsset.extension),
    };
  }

  const imageUrl = row.image.trim();
  if (!/^https?:\/\//iu.test(imageUrl)) {
    return null;
  }

  try {
    const response = await fetch(imageUrl);
    if (!response.ok) {
      return null;
    }
    const contentType = response.headers.get("content-type") || "";
    const extension =
      contentType.includes("png") ? "png" :
      contentType.includes("webp") ? "png" :
      "jpeg";
    return {
      buffer: new Uint8Array(await response.arrayBuffer()),
      extension: normalizeImageExtension(extension),
    };
  } catch {
    return null;
  }
}

const thinBorder = {
  top: { style: "thin" as const },
  left: { style: "thin" as const },
  bottom: { style: "thin" as const },
  right: { style: "thin" as const },
};

const warehouseColorPalette = [
  "FFD9EAF7",
  "FFEADCF8",
  "FFFCE4D6",
  "FFE2F0D9",
  "FFFFF2CC",
  "FFF4CCCC",
  "FFDDEBF7",
];

const templateColumnWidths = [15.1640625, 13, 11.5, 10.83203125];
const templateBoxColumnWidth = 13;
const templateRowHeights = {
  1: 16,
  2: 16,
  3: 28,
  4: 17,
  5: 35,
  17: 33,
  18: 29,
  19: 29,
  20: 28,
  21: 28,
};

function upsertNumericCell(sheetXml: string, cellRef: string, value: number) {
  const rowNumber = Number(cellRef.replace(/^[A-Z]+/u, ""));
  const cellRegex = new RegExp(`(<c[^>]*r="${cellRef}"[^>]*>)([\\s\\S]*?)(</c>)`, "u");

  if (cellRegex.test(sheetXml)) {
    return sheetXml.replace(cellRegex, (_match, startTag, innerXml, endTag) => {
      if (/<v>[\s\S]*?<\/v>/u.test(innerXml)) {
        return `${startTag}${innerXml.replace(/<v>[\s\S]*?<\/v>/u, `<v>${value}</v>`)}${endTag}`;
      }

      if (/<f>[\s\S]*?<\/f>/u.test(innerXml)) {
        return `${startTag}${innerXml}<v>${value}</v>${endTag}`;
      }

      return `${startTag}<v>${value}</v>${endTag}`;
    });
  }

  const rowRegex = new RegExp(`(<row[^>]*r="${rowNumber}"[^>]*>)([\\s\\S]*?)(</row>)`, "u");
  if (rowRegex.test(sheetXml)) {
    return sheetXml.replace(rowRegex, (_match, startTag, innerXml, endTag) => `${startTag}${innerXml}<c r="${cellRef}"><v>${value}</v></c>${endTag}`);
  }

  return sheetXml;
}

function parseBoxHeader(value: unknown) {
  const text = toText(value);
  const matched = text.match(/第?\s*(\d+)\s*箱?/u);
  if (!matched) {
    return null;
  }

  const parsed = Number(matched[1]);
  return Number.isFinite(parsed) ? parsed : null;
}

function findPackageCount(matrix: (string | number | null)[][]) {
  const rowIndex = matrix.findIndex((row) => row.some((cell) => toText(cell).includes("包装箱总数")));
  if (rowIndex < 0) {
    return null;
  }

  const row = matrix[rowIndex] ?? [];
  const labelIndex = row.findIndex((cell) => toText(cell).includes("包装箱总数"));
  if (labelIndex < 0) {
    return null;
  }

  for (let index = labelIndex + 1; index < row.length; index += 1) {
    const parsed = parseNumber(row[index]);
    if (parsed !== null) {
      return parsed;
    }
  }

  return null;
}

export async function parseAWorkbook(file: File, options: ParseAWorkbookOptions = {}): Promise<AWorkbookSummary> {
  const { workbook } = await readWorkbook(file);
  const latestSheetName = workbook.SheetNames.at(-1) ?? "";

  if (!latestSheetName) {
    throw new Error("A表没有可读取的 sheet");
  }

  const sheet = workbook.Sheets[latestSheetName];
  const matrix = getSheetMatrix(sheet);
  const excelJsWorkbook = new ExcelJS.Workbook();
  const imageMap = new Map<number, { extension: string; data: Uint8Array }>();

  if (!options.skipImages) {
    await excelJsWorkbook.xlsx.load(await file.arrayBuffer());
    const excelJsSheet = excelJsWorkbook.getWorksheet(latestSheetName);

    if (excelJsSheet?.getImages) {
      for (const image of excelJsSheet.getImages()) {
        const imageId = Number(image.imageId);
        const media = (excelJsWorkbook.model.media as ExcelJsMediaItem[] | undefined)?.find((item) => item.index === imageId);
        if (!media?.buffer || !media.extension) {
          continue;
        }
        imageMap.set(image.range.tl.nativeRow + 1, {
          extension: media.extension,
          data: new Uint8Array(media.buffer),
        });
      }
    }
  }

  const headerRow = (matrix[2] ?? []) as (string | number | null)[];
  const boxHeaders = headerRow
    .map((value, index) => ({ boxNo: parseNumber(value), index }))
    .filter((item) => item.index >= 24 && item.boxNo !== null)
    .map((item) => item.boxNo as number);

  const rows: AProductRow[] = [];

  for (let rowIndex = 3; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const sku = toText(row[4]);
    const totalShipment = parseNumber(row[23]) ?? 0;

    if (!sku && totalShipment === 0) {
      continue;
    }

    const boxMap: Record<number, number> = {};

    boxHeaders.forEach((boxNo, offset) => {
      const value = parseNumber(row[24 + offset]);
      if (value !== null && value !== 0) {
        boxMap[boxNo] = value;
      }
    });

    rows.push({
      rowIndex: rowIndex + 1,
      image: toText(row[0]),
      imageAsset: imageMap.get(rowIndex + 1) ?? null,
      productName: toText(row[1]),
      asin: toText(row[2]),
      fnsku: toText(row[3]),
      sku,
      packageSize: toText(row[5]),
      packageWeightKg: parseNumber(row[6]),
      hsCode: toText(row[7]),
      purchaseCost: parseNumber(row[12]),
      totalShipment,
      boxMap,
    });
  }

  const inlineMetricRow = (matrix[1] ?? []) as (string | number | null)[];
  const boxWeightRow = matrix.find((row) => toText(row[4]).toLowerCase() === "weight of box (kg)");
  const boxLengthRow = matrix.find((row) => toText(row[4]).toLowerCase() === "box length (cm)");
  const boxWidthRow = matrix.find((row) => toText(row[4]).toLowerCase() === "box width (cm)");
  const boxHeightRow = matrix.find((row) => toText(row[4]).toLowerCase() === "box height (cm)");

  const mapBoxMetric = (metricRow?: (string | number | null)[], inlineStartIndex?: number) =>
    Object.fromEntries(
      boxHeaders.map((boxNo, offset) => {
        const value =
          parseNumber(metricRow?.[24 + offset]) ??
          (inlineStartIndex !== undefined ? parseNumber(inlineMetricRow[inlineStartIndex + offset]) : null) ??
          parseNumber(metricRow?.[7 + boxNo - 1]) ??
          0;
        return [boxNo, value];
      }),
    ) as Record<number, number>;

  return {
    latestSheetName,
    totalRows: rows.length,
    totalBoxes: boxHeaders.length,
    totalShipment: rows.reduce((sum, row) => sum + row.totalShipment, 0),
    imageParsingSkipped: Boolean(options.skipImages),
    boxHeaders,
    boxWeightKgMap: mapBoxMetric(boxWeightRow, 24),
    boxLengthCmMap: mapBoxMetric(boxLengthRow),
    boxWidthCmMap: mapBoxMetric(boxWidthRow),
    boxHeightCmMap: mapBoxMetric(boxHeightRow),
    rows,
  };
}

export async function parseBWorkbook(): Promise<BWorkbookSummary> {
  return {
    templateType: "amazon-official",
    sheetNames: ["Create workflow – template"],
    templateSheetName: "Create workflow – template",
    headerRow: 6,
  };
}

export async function parseCWorkbook(file: File): Promise<CWorkbookSummary> {
  const { workbook } = await readWorkbook(file);
  const sheetName =
    workbook.SheetNames.find((name) => name === "Pack List") ??
    workbook.SheetNames.find((name) => name === "包装箱包装信息") ??
    workbook.SheetNames[0] ??
    "";

  if (!sheetName) {
    throw new Error("C表没有可读取的 sheet");
  }

  const sheet = workbook.Sheets[sheetName];
  const matrix = getSheetMatrix(sheet);
  const headerRowIndex = matrix.findIndex((row) => row.some((cell) => toText(cell) === "SKU" || toText(cell) === "第1箱" || toText(cell) === "第 1 箱"));
  const headerRow = headerRowIndex >= 0 ? matrix[headerRowIndex] ?? [] : matrix[1] ?? [];
  const detectedBoxHeaders = headerRow
    .map((value, index) => ({ boxNo: parseBoxHeader(value), index }))
    .filter((item) => item.index >= 0 && item.boxNo !== null)
    .map((item) => item.boxNo as number);

  const packageCount = findPackageCount(matrix);

  const isNewAmazonPackSheet = sheetName === "包装箱包装信息";
  const boxHeaders =
    isNewAmazonPackSheet && packageCount && packageCount > 0
      ? Array.from({ length: packageCount }, (_, index) => index + 1)
      : detectedBoxHeaders;

  const skuRows = matrix
    .slice(Math.max(headerRowIndex + 1, 0))
    .map((row) => {
      const skuIndex = headerRow.findIndex((cell) => toText(cell) === "SKU");
      return toText(row[skuIndex >= 0 ? skuIndex : 4]);
    })
    .filter(Boolean);

  return {
    sheetName,
    totalBoxes: boxHeaders.length,
    skuRows,
    boxHeaders,
  };
}

export async function parseSaihuWorkbook(file: File): Promise<SaihuWorkbookSummary> {
  const { workbook } = await readWorkbook(file);
  const sheetName = workbook.SheetNames[0] ?? null;
  if (!sheetName) {
    throw new Error("赛狐模板没有可读取的 sheet");
  }

  const sheet = workbook.Sheets[sheetName];
  const matrix = getSheetMatrix(sheet);
  const defaultStore = toText(matrix[1]?.[0]) || toText(matrix[0]?.[0]) || "";

  return {
    sheetNames: workbook.SheetNames,
    sheetName,
    defaultStore,
  };
}

export async function parseDWorkbook(file: File): Promise<DWorkbookSummary> {
  const { workbook } = await readWorkbook(file);

  return {
    sheetNames: workbook.SheetNames,
    templateSheetName: workbook.SheetNames.find((name) => name === "Sheet1") ?? workbook.SheetNames[0] ?? null,
    warehouseSheetName: workbook.SheetNames.find((name) => name === "FBA仓库地址") ?? null,
  };
}

export async function buildBWorkbook(aSummary: AWorkbookSummary): Promise<WorkbookExportResult> {
  const response = await fetch("/logistics-templates/ManifestFileUpload_Template_IncludeCasePack_IncludeExpirationDate_IncludeMLC_MPL.xlsx");
  if (!response.ok) {
    throw new Error("未找到内置的亚马逊官方 Create workflow 模板");
  }

  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = "Create workflow – template";
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    throw new Error("B表缺少 Create workflow – template");
  }

  const rows = aSummary.rows.filter((row) => row.totalShipment > 0);

  rows.forEach((row, index) => {
    XLSX.utils.sheet_add_aoa(
      sheet,
      [[row.sku, row.totalShipment]],
      { origin: { r: 6 + index, c: 0 } },
    );
  });

  const output = XLSX.write(workbook, { type: "array", bookType: "xlsx" });

  return {
    fileName: "ManifestFileUpload_Template_IncludeCasePack_IncludeExpirationDate_IncludeMLC_MPL.xlsx",
    blob: new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  };
}

export async function buildCWorkbook(file: File, aSummary: AWorkbookSummary): Promise<WorkbookExportResult> {
  const buffer = await file.arrayBuffer();
  const xlsxWorkbook = XLSX.read(buffer, { type: "array", cellDates: true, cellHTML: false, cellFormula: true });
  const sheetName =
    xlsxWorkbook.SheetNames.find((name) => name === "Pack List") ??
    xlsxWorkbook.SheetNames.find((name) => name === "包装箱包装信息");

  if (!sheetName) {
    throw new Error("C表缺少 Pack List 或 包装箱包装信息");
  }

  const sheet = xlsxWorkbook.Sheets[sheetName];
  const matrix = getSheetMatrix(sheet);
  const headerRowIndex = matrix.findIndex((row) => row.some((cell) => toText(cell) === "SKU" || toText(cell) === "第1箱" || toText(cell) === "第 1 箱"));
  const headerRow = headerRowIndex >= 0 ? matrix[headerRowIndex] ?? [] : matrix[1] ?? [];
  const skuColIndex = headerRow.findIndex((cell) => toText(cell) === "SKU");
  const boxColIndexMap = new Map<number, number>();
  const isNewAmazonPackSheet = sheetName === "包装箱包装信息";

  if (isNewAmazonPackSheet) {
    const packageCount = findPackageCount(matrix);

    if (!packageCount || packageCount <= 0) {
      throw new Error("新版 C 表未识别到包装箱总数");
    }

    for (let boxNo = 1; boxNo <= packageCount; boxNo += 1) {
      boxColIndexMap.set(boxNo, 13 + boxNo - 1);
    }
  } else {
    headerRow.forEach((cell, index) => {
      const boxNo = parseBoxHeader(cell);
      if (boxNo !== null) {
        boxColIndexMap.set(boxNo, index + 1);
      }
    });
  }

  const skuRowNumberMap = new Map<string, number>();

  matrix.forEach((row, index) => {
    const sku = toText(row[skuColIndex >= 0 ? skuColIndex : 4]);
    if (sku) {
      skuRowNumberMap.set(sku, index + 1);
    }
  });

  const metricRowMap = {
    weight: matrix.findIndex((row) => row.some((cell) => toText(cell).includes("包装箱重量") || toText(cell).toLowerCase() === "weight of box (kg)")) + 1,
    width: matrix.findIndex((row) => row.some((cell) => toText(cell).includes("包装箱宽度") || toText(cell).toLowerCase() === "box width (cm)")) + 1,
    length: matrix.findIndex((row) => row.some((cell) => toText(cell).includes("包装箱长度") || toText(cell).toLowerCase() === "box length (cm)")) + 1,
    height: matrix.findIndex((row) => row.some((cell) => toText(cell).includes("包装箱高度") || toText(cell).toLowerCase() === "box height (cm)")) + 1,
  };
  const zip = await JSZip.loadAsync(buffer);
  const targetSheetPath = sheetName === "包装箱包装信息" ? "xl/worksheets/sheet2.xml" : "xl/worksheets/sheet1.xml";
  const originalSheetXml = await zip.file(targetSheetPath)?.async("string");

  if (!originalSheetXml) {
    throw new Error("未找到 C 表对应的 worksheet xml");
  }

  let patchedSheetXml = originalSheetXml;

  aSummary.rows.forEach((product) => {
    const sheetRowNumber = skuRowNumberMap.get(product.sku);
    if (sheetRowNumber === undefined) {
      return;
    }

    for (const [boxNoText, qty] of Object.entries(product.boxMap)) {
      const boxNo = Number(boxNoText);
      const colIndex = boxColIndexMap.get(boxNo);
      if (colIndex === undefined) {
        continue;
      }
      patchedSheetXml = upsertNumericCell(patchedSheetXml, buildCellRef(sheetRowNumber, colIndex), qty);
    }
  });

  aSummary.boxHeaders.forEach((boxNo) => {
    const colIndex = boxColIndexMap.get(boxNo);
    if (colIndex === undefined) {
      return;
    }
    if (metricRowMap.weight > 0) {
      const weightLb = Math.round((aSummary.boxWeightKgMap[boxNo] ?? 0) / 0.454);
      patchedSheetXml = upsertNumericCell(patchedSheetXml, buildCellRef(metricRowMap.weight, colIndex), weightLb);
    }
    if (metricRowMap.width > 0) {
      patchedSheetXml = upsertNumericCell(patchedSheetXml, buildCellRef(metricRowMap.width, colIndex), 18);
    }
    if (metricRowMap.length > 0) {
      patchedSheetXml = upsertNumericCell(patchedSheetXml, buildCellRef(metricRowMap.length, colIndex), 17);
    }
    if (metricRowMap.height > 0) {
      patchedSheetXml = upsertNumericCell(patchedSheetXml, buildCellRef(metricRowMap.height, colIndex), 16);
    }
  });

  zip.file(targetSheetPath, patchedSheetXml);
  const output = await zip.generateAsync({ type: "uint8array" });
  const outputBytes = Array.from(output);

  return {
    fileName: file.name,
    blob: new Blob([new Uint8Array(outputBytes)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  };
}

export async function buildSaihuWorkbook(file: File, aSummary: AWorkbookSummary): Promise<WorkbookExportResult> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellHTML: false, cellFormula: true });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("赛狐模板缺少可用工作表");
  }

  const sheet = workbook.Sheets[sheetName];
  const matrix = getSheetMatrix(sheet);
  const defaultStore = toText(matrix[1]?.[0]) || toText(matrix[0]?.[0]) || "";
  const rows = aSummary.rows.filter((row) => row.totalShipment > 0);

  const dataRows = rows.map((row) => [defaultStore, row.sku, row.totalShipment, ""]);
  const rebuiltSheet = XLSX.utils.aoa_to_sheet([
    matrix[0] ?? ["*店铺", "*MSKU", "*申报数", "商品有效期（YYYY/MM/DD）"],
    ...dataRows,
  ]);
  workbook.Sheets[sheetName] = rebuiltSheet;

  const output = XLSX.write(workbook, { type: "array", bookType: "xlsx" });

  return {
    fileName: file.name,
    blob: new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  };
}

export async function buildSummaryWorkbook(aSummary: AWorkbookSummary, pdfSummary: PdfSummary | null): Promise<WorkbookExportResult> {
  const workbook = XLSX.utils.book_new();
  const boxHeaders = aSummary.boxHeaders.map((boxNo) => String(boxNo));
  const header = [
    "SKU",
    "品名",
    "最终发货",
    "海关编码",
    "采购成本",
    "仓库名称",
    "货件号",
    "FBA编号",
    "物流渠道",
    "重量/kg",
    ...boxHeaders,
  ];

  const data = aSummary.rows.map((row) => [
    row.sku,
    row.productName,
    row.totalShipment,
    row.hsCode,
    row.purchaseCost ?? "",
    pdfSummary?.warehouseCode ?? "",
    pdfSummary?.shipmentName ?? "",
    pdfSummary?.fbaCode ?? "",
    pdfSummary?.channelName ?? "",
    Object.entries(row.boxMap)
      .map(([boxNo]) => aSummary.boxWeightKgMap[Number(boxNo)] ?? "")
      .filter(Boolean)
      .join(" / "),
    ...aSummary.boxHeaders.map((boxNo) => row.boxMap[boxNo] ?? ""),
  ]);

  const sheet = XLSX.utils.aoa_to_sheet([header, ...data]);
  XLSX.utils.book_append_sheet(workbook, sheet, "装箱汇总表");

  const output = XLSX.write(workbook, { type: "array", bookType: "xlsx" });

  return {
    fileName: `装箱汇总表_${Date.now()}.xlsx`,
    blob: new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  };
}

export async function buildComparisonWorkbook(aSummary: AWorkbookSummary, pdfSummaries: PdfSummary[]): Promise<WorkbookExportResult> {
  const response = await fetch("/logistics-templates/6.30-%E5%88%9B%E8%B4%A7%E4%BB%B6%E5%AF%B9%E6%AF%94%E8%A1%A8.xlsx");
  if (!response.ok) {
    throw new Error("未找到内置的创货件对比表模板");
  }

  const buffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.getWorksheet("优化货件") ?? workbook.worksheets[0];

  if (!worksheet) {
    throw new Error("创货件对比表模板缺少可用工作表");
  }

  const shippedRows = aSummary.rows.filter((row) => row.totalShipment > 0);
  const pages = pdfSummaries.flatMap((pdf) =>
    pdf.pages.map((page) => {
      const columnLabel = page.positionCode || `P1-B${page.pageNumber}`;
      return {
        ...page,
        warehouseCode: page.warehouseCode || pdf.warehouseCode,
        shipmentName: page.shipmentName || pdf.shipmentName,
        weightKg: aSummary.boxWeightKgMap[page.pageNumber] ?? 0,
        columnLabel,
        channelName: pdf.channelName || "",
        boxNo: parseBoxNumberFromLabel(columnLabel, page.pageNumber),
      };
    }),
  );
  const sortedBoxNos = [...aSummary.boxHeaders].sort((left, right) => left - right);
  const pageByBoxNo = new Map<number, (typeof pages)[number]>();
  pages.forEach((page) => {
    if (!pageByBoxNo.has(page.boxNo)) {
      pageByBoxNo.set(page.boxNo, page);
    }
  });

  const boxColumns = sortedBoxNos.map((boxNo) => {
    const page = pageByBoxNo.get(boxNo);
    return {
      boxNo,
      weightKg: aSummary.boxWeightKgMap[boxNo] ?? 0,
      warehouseCode: page?.warehouseCode || "",
      shipmentName: page?.shipmentName || "",
      channelName: page?.channelName || "",
      columnLabel: page?.columnLabel || `P1-B${boxNo}`,
    };
  });
  const warehouseFillMap = new Map<string, string>();
  let paletteIndex = 0;
  boxColumns.forEach((column) => {
    const key = column.warehouseCode || "__empty__";
    if (!warehouseFillMap.has(key)) {
      warehouseFillMap.set(key, warehouseColorPalette[paletteIndex % warehouseColorPalette.length]);
      paletteIndex += 1;
    }
  });

  const headerEndColumn = 4 + boxColumns.length;
  const dataStartRow = 6;
  const templateDataStartRow = 6;
  const templateDataEndRow = 16;
  const templateTailStartRow = 17;
  const baseColumnCount = 4;
  const templateMaxColumn = Math.max(worksheet.columnCount, 45);
  const lastTemplateDataColumn = Math.max(templateMaxColumn, headerEndColumn);
  const extraRowCount = Math.max(shippedRows.length - (templateDataEndRow - templateDataStartRow + 1), 0);

  if (extraRowCount > 0) {
    worksheet.spliceRows(templateTailStartRow, 0, ...Array.from({ length: extraRowCount }, () => []));
  }

  const tailStartRow = templateTailStartRow + extraRowCount;
  const logisticsChannelRow = tailStartRow;
  const logisticsPriceRow = tailStartRow + 1;
  const logisticsFeeRow = tailStartRow + 2;
  const inboundFeeRow = tailStartRow + 3;
  const totalFeeRow = tailStartRow + 4;

  for (let columnNumber = 1; columnNumber <= Math.max(templateMaxColumn, headerEndColumn); columnNumber += 1) {
    const currentWidth = worksheet.getColumn(columnNumber).width;
    if (columnNumber <= baseColumnCount) {
      worksheet.getColumn(columnNumber).width = currentWidth ?? templateColumnWidths[columnNumber - 1] ?? 12;
      continue;
    }
    worksheet.getColumn(columnNumber).width = currentWidth ?? templateBoxColumnWidth;
  }

  worksheet.getCell("D1").value = "重量/kg";
  worksheet.getCell("D2").value = "数量";
  worksheet.getCell("D3").value = "仓库名称";
  worksheet.getCell("D4").value = "货件号";
  worksheet.getCell("A5").value = "产品图片";
  worksheet.getCell("B5").value = "品名";
  worksheet.getCell("C5").value = "SKU";
  worksheet.getCell("D5").value = "最终发货";

  for (let columnNumber = 5; columnNumber <= templateMaxColumn; columnNumber += 1) {
    const columnName = columnNumberToName(columnNumber);
    const isUsed = columnNumber <= headerEndColumn;
    worksheet.getCell(`${columnName}1`).value = isUsed ? "" : null;
    worksheet.getCell(`${columnName}2`).value = isUsed ? "" : null;
    worksheet.getCell(`${columnName}3`).value = isUsed ? "" : null;
    worksheet.getCell(`${columnName}4`).value = isUsed ? "" : null;
    worksheet.getCell(`${columnName}5`).value = isUsed ? "" : null;
  }

  boxColumns.forEach((column, pageIndex) => {
    const columnNumber = 5 + pageIndex;
    const columnName = columnNumberToName(columnNumber);
    worksheet.getCell(`${columnName}1`).value = column.weightKg || "";
    worksheet.getCell(`${columnName}2`).value = {
      formula: `SUM(${columnName}${dataStartRow}:${columnName}${dataStartRow + shippedRows.length - 1})`,
    };
    worksheet.getCell(`${columnName}3`).value = column.warehouseCode || "--";
    worksheet.getCell(`${columnName}4`).value = column.shipmentName || "--";
    worksheet.getCell(`${columnName}5`).value = column.columnLabel;
  });

  for (let rowNumber = dataStartRow; rowNumber < tailStartRow; rowNumber += 1) {
    worksheet.getCell(`A${rowNumber}`).value = "";
    worksheet.getCell(`B${rowNumber}`).value = "";
    worksheet.getCell(`C${rowNumber}`).value = "";
    worksheet.getCell(`D${rowNumber}`).value = "";
    for (let columnNumber = 5; columnNumber <= lastTemplateDataColumn; columnNumber += 1) {
      worksheet.getCell(rowNumber, columnNumber).value = "";
    }
  }

  shippedRows.forEach((row, index) => {
    const rowNumber = dataStartRow + index;
    worksheet.getCell(`A${rowNumber}`).value = "";
    worksheet.getCell(`B${rowNumber}`).value = row.productName || "";
    worksheet.getCell(`C${rowNumber}`).value = row.sku || "";
    worksheet.getCell(`D${rowNumber}`).value = {
      formula: `SUM(E${rowNumber}:${columnNumberToName(headerEndColumn)}${rowNumber})`,
    };
    boxColumns.forEach(({ boxNo }, pageIndex) => {
      worksheet.getCell(rowNumber, 5 + pageIndex).value = row.boxMap[boxNo] ?? "";
    });
  });

  for (let rowNumber = tailStartRow; rowNumber <= totalFeeRow; rowNumber += 1) {
    for (let columnNumber = 5; columnNumber <= templateMaxColumn; columnNumber += 1) {
      worksheet.getCell(rowNumber, columnNumber).value = columnNumber <= headerEndColumn ? worksheet.getCell(rowNumber, columnNumber).value : "";
    }
  }

  worksheet.getCell(`D${logisticsChannelRow}`).value = "物流渠道";
  worksheet.getCell(`D${logisticsPriceRow}`).value = "物流单价";
  worksheet.getCell(`D${logisticsFeeRow}`).value = "物流费用";
  worksheet.getCell(`D${inboundFeeRow}`).value = "入库配置费（美金）";
  worksheet.getCell(`D${totalFeeRow}`).value = "总费用";

  boxColumns.forEach((column, pageIndex) => {
    const columnNumber = 5 + pageIndex;
    const columnName = columnNumberToName(columnNumber);
    worksheet.getCell(`${columnName}${logisticsChannelRow}`).value = column.channelName || "--";
    worksheet.getCell(`${columnName}${logisticsPriceRow}`).value = "";
    worksheet.getCell(`${columnName}${logisticsFeeRow}`).value = {
      formula: `${columnName}${logisticsPriceRow}*${columnName}1`,
    };
    worksheet.getCell(`${columnName}${inboundFeeRow}`).value = pageIndex === 0 ? 0 : "";
  });

  if (boxColumns.length) {
    worksheet.getCell(`E${totalFeeRow}`).value = {
      formula: `SUM(E${logisticsFeeRow}:${columnNumberToName(headerEndColumn)}${logisticsFeeRow})+E${inboundFeeRow}*7.2`,
    };
  }

  Object.entries(templateRowHeights).forEach(([rowNumber, height]) => {
    worksheet.getRow(Number(rowNumber)).height = height;
  });

  shippedRows.forEach((_, index) => {
    const rowNumber = dataStartRow + index;
    const cycleHeights = [69, 79, 87];
    worksheet.getRow(rowNumber).height = worksheet.getRow(rowNumber).height ?? cycleHeights[index % cycleHeights.length];
  });

  for (let rowNumber = templateDataEndRow + 1; rowNumber < tailStartRow; rowNumber += 1) {
    worksheet.getRow(rowNumber).height = worksheet.getRow(templateDataEndRow).height ?? 72;
  }

  worksheet.getRow(logisticsChannelRow).height = worksheet.getRow(templateTailStartRow).height ?? 33;
  worksheet.getRow(logisticsPriceRow).height = worksheet.getRow(templateTailStartRow + 1).height ?? 29;
  worksheet.getRow(logisticsFeeRow).height = worksheet.getRow(templateTailStartRow + 2).height ?? 29;
  worksheet.getRow(inboundFeeRow).height = worksheet.getRow(templateTailStartRow + 3).height ?? 28;
  worksheet.getRow(totalFeeRow).height = worksheet.getRow(templateTailStartRow + 4).height ?? 28;

  boxColumns.forEach((column, pageIndex) => {
    const columnNumber = 5 + pageIndex;
    const color = warehouseFillMap.get(column.warehouseCode || "__empty__") || "FFE2F0D9";
    [3, 4].forEach((rowNumber) => {
      worksheet.getCell(rowNumber, columnNumber).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color },
      };
    });
    worksheet.getCell(5, columnNumber).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF92D050" },
    };
    worksheet.getCell(5, columnNumber).numFmt = "@";
  });

  for (let rowNumber = 1; rowNumber <= totalFeeRow; rowNumber += 1) {
    for (let columnNumber = 4; columnNumber <= headerEndColumn; columnNumber += 1) {
      worksheet.getCell(rowNumber, columnNumber).border = thinBorder;
    }
  }

  for (let columnNumber = 1; columnNumber <= 3; columnNumber += 1) {
    for (let rowNumber = 5; rowNumber < tailStartRow; rowNumber += 1) {
      worksheet.getCell(rowNumber, columnNumber).border = thinBorder;
    }
  }

  worksheet.getCell(`D${totalFeeRow}`).font = { bold: true, color: { argb: "FFFF0000" } };
  if (boxColumns.length) {
    worksheet.getCell(`E${totalFeeRow}`).font = { bold: true, color: { argb: "FFFF0000" } };
  }

  for (const [rowIndex, row] of shippedRows.entries()) {
    const resolvedImage = await resolveRowImage(row);
    if (!resolvedImage) {
      continue;
    }
    const imageId = workbook.addImage({
      base64: Buffer.from(resolvedImage.buffer).toString("base64"),
      extension: resolvedImage.extension,
    });
    const targetRow = 6 + rowIndex;
    worksheet.addImage(imageId, {
      tl: { col: 0.12, row: targetRow - 1 + 0.08 },
      ext: { width: 56, height: 56 },
      editAs: "oneCell",
    });
  }

  return {
    fileName: `创货件对比表_${Date.now()}.xlsx`,
    blob: new Blob([await workbook.xlsx.writeBuffer()], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  };
}

function inferEnglishName(productName: string, sku: string) {
  const samples: Record<string, string> = {
    "0E-7MD5-T78E": "Sand Bag",
    "PP-4VL0-SDK2": "Softball Sleeve Holders",
    "LJ-98F8-THD5": "Soccer Armband",
    "U3-UBDP-WLQY": "Softball Sleeve Holders",
    "5J-APYY-A5SU": "Golf Cart Flag",
    "9M-JWM6-20L1": "Sand Bag",
    "KA-USA Flag": "Golf Cart Flag",
    "KA-pickleball court": "Pickleball Court Marking Kit",
    "B8-FISN-297Q": "Diving Flag Mount",
    "KA-Handicap Flag": "Golf Cart Flag",
    "T-SANDBAG-BLK-L": "Sand Bag",
  };

  if (samples[sku]) {
    return samples[sku];
  }

  return productName || "";
}

function buildDLinesForPdf(aSummary: AWorkbookSummary, pdfSummary: PdfSummary) {
  const lines: Array<{
    boxNo: number;
    fbaBoxCode: string;
    weightKg: number;
    hsCode: string;
    productNameCn: string;
    productNameEn: string;
    qtyPerBox: number;
    declarePrice: number | "";
    image: string;
  }> = [];

  pdfSummary.pages.forEach((page, pageIndex) => {
    const boxNo = page.pageNumber;
    const weightKg = aSummary.boxWeightKgMap[boxNo] ?? 0;

    const matchedProducts = aSummary.rows.filter((row) => (row.boxMap[boxNo] ?? 0) > 0);

    if (!matchedProducts.length) {
      lines.push({
        boxNo,
        fbaBoxCode: page.fbaBoxCode,
        weightKg,
        hsCode: "",
        productNameCn: "",
        productNameEn: "",
        qtyPerBox: page.qty ?? 0,
        declarePrice: "",
        image: "",
      });
      return;
    }

    matchedProducts.forEach((product) => {
      lines.push({
        boxNo,
        fbaBoxCode: page.fbaBoxCode,
        weightKg,
        hsCode: product.hsCode,
        productNameCn: product.productName,
        productNameEn: inferEnglishName(product.productName, product.sku),
        qtyPerBox: product.boxMap[boxNo] ?? 0,
        declarePrice: product.purchaseCost ? Number((product.purchaseCost / 8).toFixed(2)) : "",
        image: product.image,
      });
    });

    if (page.skuType === "Single SKU" && page.sku) {
      const hasMatchedSku = matchedProducts.some((product) => product.sku.replace(/\s+/g, "") === page.sku.replace(/\s+/g, ""));
      if (!hasMatchedSku) {
        // Current version trusts A table as source of truth but keeps the page mapping by box number.
        void pageIndex;
      }
    }
  });

  return lines;
}

export async function buildDWorkbooks(
  file: File,
  aSummary: AWorkbookSummary,
  pdfSummaries: PdfSummary[],
): Promise<NamedWorkbookExportResult[]> {
  const buffer = await file.arrayBuffer();
  const exports: NamedWorkbookExportResult[] = [];

  for (const pdfSummary of pdfSummaries) {
    const workbook = XLSX.read(buffer.slice(0), { type: "array" });
    const sheet = workbook.Sheets["Sheet1"];

    if (!sheet) {
      throw new Error("D表缺少 Sheet1");
    }

    XLSX.utils.sheet_add_aoa(sheet, [[pdfSummary.fbaCode]], { origin: "B3" });
    XLSX.utils.sheet_add_aoa(sheet, [[pdfSummary.warehouseCode]], { origin: "E3" });
    XLSX.utils.sheet_add_aoa(sheet, [[pdfSummary.channelName || ""]], { origin: "B4" });
    XLSX.utils.sheet_add_aoa(sheet, [["美国"]], { origin: "B5" });
    XLSX.utils.sheet_add_aoa(sheet, [[pdfSummary.pages.length]], { origin: "B6" });
    XLSX.utils.sheet_add_aoa(sheet, [["一般报关"]], { origin: "B7" });
    XLSX.utils.sheet_add_aoa(sheet, [["是"]], { origin: "B8" });
    XLSX.utils.sheet_add_aoa(sheet, [["否"]], { origin: "B11" });

    const lines = buildDLinesForPdf(aSummary, pdfSummary);
    const dataRows = lines.map((line) => [
      line.boxNo,
      line.fbaBoxCode,
      "",
      line.weightKg || "",
      50,
      40,
      40,
      line.hsCode,
      line.productNameCn,
      line.productNameEn,
      line.qtyPerBox || "",
      line.declarePrice,
      "无",
      "无",
      "尼龙/Nylon",
      "户外/Outdoor",
      line.image,
    ]);

    if (dataRows.length) {
      XLSX.utils.sheet_add_aoa(sheet, dataRows, { origin: { r: 15, c: 0 } });
    }

    const output = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
    const fileName = pdfSummary.renamedFileName.replace(/\.pdf$/i, ".xlsx");

    exports.push({
      key: pdfSummary.fileNameBase,
      fileName,
      blob: new Blob([output], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    });
  }

  return exports;
}
