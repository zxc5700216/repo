import * as XLSX from "xlsx";
import type { AdjustmentDraft, DraftValidationResult, HeaderMap } from "@/lib/types";

type Worksheet = XLSX.WorkSheet;
type Workbook = XLSX.WorkBook;

const fieldHeaderCandidates = {
  bid: ["竞价", "Bid"],
  state: ["状态", "State"],
  operation: ["操作", "Operation"],
};

function normalizeHeader(value: string) {
  return value.toLowerCase().replace(/[\s()[\]_\-:：（）]/g, "");
}

function columnName(columnIndex: number) {
  let column = "";
  let index = columnIndex + 1;

  while (index > 0) {
    const remainder = (index - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    index = Math.floor((index - 1) / 26);
  }

  return column;
}

function getSheetRange(sheet: Worksheet) {
  return sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null;
}

export function buildHeaderMap(sheet: Worksheet): HeaderMap {
  const range = getSheetRange(sheet);
  const headerMap: HeaderMap = {};

  if (!range) {
    return headerMap;
  }

  for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: range.s.r, c: columnIndex });
    const headerName = String(sheet[address]?.v ?? "").trim();

    if (!headerName) {
      continue;
    }

    headerMap[normalizeHeader(headerName)] = {
      headerName,
      columnIndex,
      excelColumn: columnName(columnIndex),
    };
  }

  return headerMap;
}

export function getHeaderEntry(headerMap: HeaderMap, field: "bid" | "state" | "operation") {
  const candidates = fieldHeaderCandidates[field].map(normalizeHeader);
  return candidates.map((candidate) => headerMap[candidate]).find(Boolean);
}

export function getCellByField(
  _sheet: Worksheet,
  headerMap: HeaderMap,
  rowIndex: number,
  field: "bid" | "state" | "operation",
) {
  const entry = getHeaderEntry(headerMap, field);

  if (!entry) {
    return null;
  }

  return XLSX.utils.encode_cell({ r: rowIndex - 1, c: entry.columnIndex });
}

export function validateDraftCellTarget(workbook: Workbook, draft: AdjustmentDraft): DraftValidationResult {
  if (!draft.selected) {
    return {
      draftId: draft.id,
      valid: false,
      status: "blocked",
      message: "草稿未勾选，不允许写回。",
    };
  }

  if (!draft.sheetName || !draft.sourceRowIndex || !draft.field) {
    return {
      draftId: draft.id,
      valid: false,
      status: "blocked",
      message: "草稿缺少 Sheet、原始行号或写回字段。",
    };
  }

  const sheet = workbook.Sheets[draft.sheetName];

  if (!sheet) {
    return {
      draftId: draft.id,
      valid: false,
      status: "blocked",
      message: "目标 Sheet 不存在。",
      sheetName: draft.sheetName,
      sourceRowIndex: draft.sourceRowIndex,
      headerName: draft.headerName,
    };
  }

  const range = getSheetRange(sheet);

  if (!range || draft.sourceRowIndex < 1 || draft.sourceRowIndex > range.e.r + 1) {
    return {
      draftId: draft.id,
      valid: false,
      status: "blocked",
      message: "原始行号不存在。",
      sheetName: draft.sheetName,
      sourceRowIndex: draft.sourceRowIndex,
      headerName: draft.headerName,
    };
  }

  const headerMap = buildHeaderMap(sheet);
  const cellAddress = getCellByField(sheet, headerMap, draft.sourceRowIndex, draft.field);

  if (!cellAddress) {
    return {
      draftId: draft.id,
      valid: false,
      status: "blocked",
      message: "目标写回列不存在。",
      sheetName: draft.sheetName,
      sourceRowIndex: draft.sourceRowIndex,
      headerName: draft.headerName,
    };
  }

  const currentValue = sheet[cellAddress]?.v ?? null;

  if (String(currentValue ?? "") !== String(draft.oldValue ?? "")) {
    return {
      draftId: draft.id,
      valid: false,
      status: "conflict",
      message: "当前单元格值与草稿原值不一致，已阻止写回。",
      sheetName: draft.sheetName,
      sourceRowIndex: draft.sourceRowIndex,
      headerName: draft.headerName,
    };
  }

  return {
    draftId: draft.id,
    valid: true,
    status: "valid",
    message: "可安全写回。",
    sheetName: draft.sheetName,
    sourceRowIndex: draft.sourceRowIndex,
    headerName: draft.headerName,
  };
}

export function applyDraftToWorkbook(workbook: Workbook, draft: AdjustmentDraft): Workbook {
  const validation = validateDraftCellTarget(workbook, draft);

  if (!validation.valid || !draft.sheetName || !draft.sourceRowIndex || !draft.field) {
    return workbook;
  }

  const sheet = workbook.Sheets[draft.sheetName];
  const headerMap = buildHeaderMap(sheet);
  const targetCell = getCellByField(sheet, headerMap, draft.sourceRowIndex, draft.field);
  const operationCell = getCellByField(sheet, headerMap, draft.sourceRowIndex, "operation");

  if (targetCell) {
    sheet[targetCell] = { t: typeof draft.newValue === "number" ? "n" : "s", v: draft.newValue };
  }

  if (operationCell) {
    sheet[operationCell] = { t: "s", v: "Update" };
  }

  return workbook;
}

export function exportSelectedDrafts(input: {
  workbookBuffer: ArrayBuffer;
  drafts: AdjustmentDraft[];
  fileName?: string;
}) {
  const workbook = XLSX.read(input.workbookBuffer, { type: "array" });
  const selectedDrafts = input.drafts.filter((draft) => draft.selected);
  const validations = selectedDrafts.map((draft) => validateDraftCellTarget(workbook, draft));

  for (const draft of selectedDrafts) {
    const validation = validations.find((item) => item.draftId === draft.id);

    if (validation?.valid) {
      applyDraftToWorkbook(workbook, draft);
    }
  }

  const writableCount = validations.filter((item) => item.valid).length;
  const output = XLSX.write(workbook, { bookType: "xlsx", type: "array" });

  return {
    data: output,
    fileName: input.fileName ?? "modified-bulk-operations.xlsx",
    validations,
    writableCount,
    blockedCount: validations.filter((item) => item.status === "blocked").length,
    conflictCount: validations.filter((item) => item.status === "conflict").length,
  };
}
