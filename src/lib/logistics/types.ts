export type UploadedFileState = {
  name: string;
  size: number;
  uploadedAt: string;
};

export type LogisticsStatusTone = "green" | "amber" | "red" | "gray" | "blue";

export type LogisticsLogLevel = "info" | "success" | "warning" | "error";

export type LogisticsLogEntry = {
  id: string;
  level: LogisticsLogLevel;
  message: string;
};

export type AProductRow = {
  rowIndex: number;
  image: string;
  productName: string;
  asin: string;
  fnsku: string;
  sku: string;
  packageSize: string;
  packageWeightKg: number | null;
  hsCode: string;
  purchaseCost: number | null;
  totalShipment: number;
  boxMap: Record<number, number>;
};

export type AWorkbookSummary = {
  latestSheetName: string;
  totalRows: number;
  totalBoxes: number;
  totalShipment: number;
  boxHeaders: number[];
  boxWeightKgMap: Record<number, number>;
  boxLengthCmMap: Record<number, number>;
  boxWidthCmMap: Record<number, number>;
  boxHeightCmMap: Record<number, number>;
  rows: AProductRow[];
};

export type BWorkbookSummary = {
  templateType: "amazon-official" | "unknown";
  sheetNames: string[];
  templateSheetName: string | null;
  headerRow: number | null;
};

export type CWorkbookSummary = {
  sheetName: string;
  totalBoxes: number;
  skuRows: string[];
  boxHeaders: number[];
};

export type DWorkbookSummary = {
  sheetNames: string[];
  templateSheetName: string | null;
  warehouseSheetName: string | null;
};

export type PdfPageSummary = {
  pageNumber: number;
  shipmentName: string;
  warehouseCode: string;
  fbaCode: string;
  totalBoxes: number | null;
  fbaBoxCode: string;
  positionCode: string;
  skuType: string;
  sku: string;
  qty: number | null;
};

export type PdfSummary = {
  fileNameBase: string;
  shipmentName: string;
  warehouseCode: string;
  fbaCode: string;
  totalBoxes: number | null;
  channelName: string;
  renamedFileName: string;
  pages: PdfPageSummary[];
};

export type WorkbookExportResult = {
  fileName: string;
  blob: Blob;
};

export type LogisticsWorkspaceState = {
  aFile: UploadedFileState | null;
  bFile: UploadedFileState | null;
  cFile: UploadedFileState | null;
  dFile: UploadedFileState | null;
  pdfFiles: UploadedFileState[];
  aSummary: AWorkbookSummary | null;
  bSummary: BWorkbookSummary | null;
  cSummary: CWorkbookSummary | null;
  dSummary: DWorkbookSummary | null;
  pdfSummaries: PdfSummary[];
  bExport: WorkbookExportResult | null;
  cExport: WorkbookExportResult | null;
  summaryExport: WorkbookExportResult | null;
  logs: LogisticsLogEntry[];
};
