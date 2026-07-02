import * as XLSX from "xlsx";
import type {
  AProductRow,
  AWorkbookSummary,
  BWorkbookSummary,
  CWorkbookSummary,
  DWorkbookSummary,
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

  const boxWeightRow = matrix.find((row) => toText(row[4]).toLowerCase() === "weight of box (kg)");
  const boxLengthRow = matrix.find((row) => toText(row[4]).toLowerCase() === "box length (cm)");
  const boxWidthRow = matrix.find((row) => toText(row[4]).toLowerCase() === "box width (cm)");
  const boxHeightRow = matrix.find((row) => toText(row[4]).toLowerCase() === "box height (cm)");

  const mapBoxMetric = (metricRow?: (string | number | null)[]) =>
    Object.fromEntries(
      boxHeaders.map((boxNo, offset) => {
        const value = parseNumber(metricRow?.[24 + offset]) ?? parseNumber(metricRow?.[7 + boxNo - 1]) ?? 0;
        return [boxNo, value];
      }),
    ) as Record<number, number>;

  return {
    latestSheetName,
    totalRows: rows.length,
    totalBoxes: boxHeaders.length,
    totalShipment: rows.reduce((sum, row) => sum + row.totalShipment, 0),
    boxHeaders,
    boxWeightKgMap: mapBoxMetric(boxWeightRow),
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
  const sheetName = workbook.SheetNames.find((name) => name === "Pack List") ?? workbook.SheetNames[0] ?? "";

  if (!sheetName) {
    throw new Error("C表没有可读取的 sheet");
  }

  const sheet = workbook.Sheets[sheetName];
  const matrix = getSheetMatrix(sheet);
  const headerRow = matrix[1] ?? [];
  const boxHeaders = headerRow
    .map((value, index) => ({ boxNo: parseNumber(toText(value).replace("第", "").replace("箱", "")), index }))
    .filter((item) => item.index >= 7 && item.boxNo !== null)
    .map((item) => item.boxNo as number);

  const skuRows = matrix
    .slice(2)
    .map((row) => toText(row[4]))
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
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets["Pack List"];

  if (!sheet) {
    throw new Error("C表缺少 Pack List");
  }

  const matrix = getSheetMatrix(sheet);
  const skuRowIndexMap = new Map<string, number>();

  matrix.forEach((row, index) => {
    const sku = toText(row[4]);
    if (sku) {
      skuRowIndexMap.set(sku, index);
    }
  });

  aSummary.rows.forEach((product) => {
    const sheetRowIndex = skuRowIndexMap.get(product.sku);
    if (sheetRowIndex === undefined) {
      return;
    }

    for (const [boxNoText, qty] of Object.entries(product.boxMap)) {
      const boxNo = Number(boxNoText);
      const colIndex = 7 + boxNo;
      XLSX.utils.sheet_add_aoa(sheet, [[qty]], { origin: { r: sheetRowIndex, c: colIndex } });
    }
  });

  aSummary.boxHeaders.forEach((boxNo) => {
    const weightLb = Number(((aSummary.boxWeightKgMap[boxNo] ?? 0) / 0.454).toFixed(2));
    const colIndex = 7 + boxNo;
    XLSX.utils.sheet_add_aoa(sheet, [[weightLb]], { origin: { r: 14, c: colIndex } });
    XLSX.utils.sheet_add_aoa(sheet, [[17]], { origin: { r: 15, c: colIndex } });
    XLSX.utils.sheet_add_aoa(sheet, [[18]], { origin: { r: 16, c: colIndex } });
    XLSX.utils.sheet_add_aoa(sheet, [[16]], { origin: { r: 17, c: colIndex } });
  });

  const output = XLSX.write(workbook, { type: "array", bookType: "xlsx" });

  return {
    fileName: `C_已填写_${Date.now()}.xlsx`,
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
