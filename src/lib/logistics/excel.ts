import * as XLSX from "xlsx";
import JSZip from "jszip";
import type {
  AProductRow,
  AWorkbookSummary,
  BWorkbookSummary,
  CWorkbookSummary,
  DWorkbookSummary,
  NamedWorkbookExportResult,
  PdfSummary,
  WorkbookExportResult,
} from "@/lib/logistics/types";
import { parseNumber, toText } from "@/lib/logistics/utils";

type WorkbookWithSheets = {
  workbook: XLSX.WorkBook;
  fileName: string;
};

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

export async function parseAWorkbook(file: File): Promise<AWorkbookSummary> {
  const { workbook } = await readWorkbook(file);
  const latestSheetName = workbook.SheetNames.at(-1) ?? "";

  if (!latestSheetName) {
    throw new Error("A表没有可读取的 sheet");
  }

  const sheet = workbook.Sheets[latestSheetName];
  const matrix = getSheetMatrix(sheet);
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
    boxHeaders,
    boxWeightKgMap: mapBoxMetric(boxWeightRow, 24),
    boxLengthCmMap: mapBoxMetric(boxLengthRow),
    boxWidthCmMap: mapBoxMetric(boxWidthRow),
    boxHeightCmMap: mapBoxMetric(boxHeightRow),
    rows,
  };
}

export async function parseBWorkbook(file: File): Promise<BWorkbookSummary> {
  const { workbook } = await readWorkbook(file);
  const templateSheetName = workbook.SheetNames.find((name) => name.includes("Create workflow")) ?? null;

  return {
    templateType: workbook.SheetNames.includes("Create workflow – template") ? "amazon-official" : "unknown",
    sheetNames: workbook.SheetNames,
    templateSheetName,
    headerRow: workbook.SheetNames.includes("Create workflow – template") ? 6 : null,
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

export async function parseDWorkbook(file: File): Promise<DWorkbookSummary> {
  const { workbook } = await readWorkbook(file);

  return {
    sheetNames: workbook.SheetNames,
    templateSheetName: workbook.SheetNames.find((name) => name === "Sheet1") ?? workbook.SheetNames[0] ?? null,
    warehouseSheetName: workbook.SheetNames.find((name) => name === "FBA仓库地址") ?? null,
  };
}

export async function buildBWorkbook(file: File, aSummary: AWorkbookSummary): Promise<WorkbookExportResult> {
  const buffer = await file.arrayBuffer();
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
    fileName: `B_已填写_${Date.now()}.xlsx`,
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
