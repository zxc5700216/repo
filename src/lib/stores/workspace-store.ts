"use client";

import { create } from "zustand";
import { campaignGroups, dataBatches, defaultRules, performanceRows as mockPerformanceRows } from "@/data/mock-data";
import {
  deleteWorkspaceSnapshot,
  readWorkspaceSnapshot,
  writeWorkspaceSnapshot,
} from "@/lib/repositories/workspace-repository";
import { runRuleEngine } from "@/lib/rule-engine/engine";
import type {
  AdjustmentDraft,
  CampaignGroup,
  CampaignSheetGroup,
  LifecycleGroupId,
  ParseJobStatus,
  PerformanceRow,
  RecentAdDataMatchSummary,
  RecentAdDataRow,
  RecentAdDataStatus,
  Rule,
  WorkspaceUnit,
} from "@/lib/types";

type SheetRow = Record<string, string | number | boolean | null | undefined>;

type ParseDiagnostics = {
  totalRows: number;
  sponsoredProductRows: number;
  rowsWithAdGroup: number;
  keywordRows: number;
  rowsWithBid: number;
  executableRows: number;
  sampleHeaders: string[];
  sampleEntities: string[];
};

interface WorkspaceState {
  rules: Rule[];
  campaignGroups: CampaignGroup[];
  campaignSheetGroups: CampaignSheetGroup[];
  workspaceUnits: WorkspaceUnit[];
  performanceRows: PerformanceRow[];
  activeCampaignGroupId: string;
  activeWorkspaceUnitId?: string;
  activeLifecycleGroupId?: LifecycleGroupId;
  workspaceMode: "campaign" | "lifecycle" | "workspace-unit";
  openTabIds: string[];
  selectedDraftIds: string[];
  parseStatus: ParseJobStatus;
  parseProgress: number;
  uploadedFileName?: string;
  originalWorkbookBuffer?: ArrayBuffer;
  activeBatchId?: string;
  parsedRowCount: number;
  parsedSheets: string[];
  parseError?: string;
  parseDiagnostics: ParseDiagnostics;
  recentAdDataFileName?: string;
  recentAdDataRows: RecentAdDataRow[];
  recentAdDataStatus: RecentAdDataStatus;
  recentAdDataError?: string;
  recentAdDataMatchSummary: RecentAdDataMatchSummary;
  adjustmentDrafts: AdjustmentDraft[];
  persistenceStatus: "loading" | "ready" | "saving" | "saved" | "failed";
  persistenceError?: string;
  setRules: (rules: Rule[]) => void;
  upsertRule: (rule: Rule) => void;
  deleteRule: (ruleId: string) => void;
  setActiveCampaignGroup: (campaignGroupId: string) => void;
  openCampaignGroup: (campaignGroupId: string) => void;
  mergeCampaignGroupsIntoWorkspaceUnit: (sourceCampaignGroupId: string, targetCampaignGroupId: string) => void;
  removeCampaignGroupFromWorkspaceUnit: (campaignGroupId: string) => void;
  setActiveWorkspaceUnit: (workspaceUnitId: string) => void;
  setActiveLifecycleGroup: (lifecycleGroupId?: LifecycleGroupId) => void;
  assignLifecycleGroup: (campaignGroupId: string, lifecycleGroupId: LifecycleGroupId) => void;
  importGroupingStatusCsv: (fileName: string, text: string) => { importedRows: number; workspaceUnitCount: number };
  runRulesForActiveGroup: () => void;
  runRulesForActiveLifecycleGroup: () => void;
  runRulesForActiveWorkspaceUnit: () => void;
  toggleDraft: (draftId: string) => void;
  setDraftSelected: (draftId: string, selected: boolean) => void;
  selectAllDrafts: () => void;
  invertDraftSelection: () => void;
  clearDraftSelection: () => void;
  setParseStarted: (fileName: string, originalWorkbookBuffer: ArrayBuffer) => void;
  setParseProgress: (progress: number, sheets?: string[]) => void;
  ingestParsedRows: (sheetName: string, rows: SheetRow[], startRowIndex: number) => void;
  setParseCompleted: (rowCount: number, sheets: string[]) => void;
  setParseFailed: (message: string) => void;
  ingestRecentAdDataCsv: (fileName: string, text: string, scopeCampaignGroupIds: string[]) => void;
  hydratePersistedWorkspace: () => Promise<void>;
  clearPersistedWorkspace: () => Promise<void>;
}

type WorkspaceSnapshot = Pick<
  WorkspaceState,
  | "rules"
  | "campaignGroups"
  | "campaignSheetGroups"
  | "workspaceUnits"
  | "performanceRows"
  | "activeCampaignGroupId"
  | "activeWorkspaceUnitId"
  | "activeLifecycleGroupId"
  | "workspaceMode"
  | "openTabIds"
  | "selectedDraftIds"
  | "parseStatus"
  | "parseProgress"
  | "uploadedFileName"
  | "originalWorkbookBuffer"
  | "activeBatchId"
  | "parsedRowCount"
  | "parsedSheets"
  | "parseError"
  | "parseDiagnostics"
  | "recentAdDataFileName"
  | "recentAdDataRows"
  | "recentAdDataStatus"
  | "recentAdDataError"
  | "recentAdDataMatchSummary"
  | "adjustmentDrafts"
>;

const initialActiveId = campaignGroups[0]?.id ?? "";

const emptyDiagnostics: ParseDiagnostics = {
  totalRows: 0,
  sponsoredProductRows: 0,
  rowsWithAdGroup: 0,
  keywordRows: 0,
  rowsWithBid: 0,
  executableRows: 0,
  sampleHeaders: [],
  sampleEntities: [],
};

const emptyRecentAdDataMatchSummary: RecentAdDataMatchSummary = {
  totalRows: 0,
  matchedRows: 0,
  unmatchedRows: 0,
  ambiguousRows: 0,
  matchedCampaignGroups: 0,
  scopedCampaignGroups: 0,
};

function takeWorkspaceSnapshot(state: WorkspaceState): WorkspaceSnapshot {
  return {
    rules: state.rules,
    campaignGroups: state.campaignGroups,
    campaignSheetGroups: state.campaignSheetGroups,
    workspaceUnits: state.workspaceUnits,
    performanceRows: state.performanceRows,
    activeCampaignGroupId: state.activeCampaignGroupId,
    activeWorkspaceUnitId: state.activeWorkspaceUnitId,
    activeLifecycleGroupId: state.activeLifecycleGroupId,
    workspaceMode: state.workspaceMode,
    openTabIds: state.openTabIds,
    selectedDraftIds: state.selectedDraftIds,
    parseStatus: state.parseStatus,
    parseProgress: state.parseProgress,
    uploadedFileName: state.uploadedFileName,
    originalWorkbookBuffer: state.originalWorkbookBuffer,
    activeBatchId: state.activeBatchId,
    parsedRowCount: state.parsedRowCount,
    parsedSheets: state.parsedSheets,
    parseError: state.parseError,
    parseDiagnostics: state.parseDiagnostics,
    recentAdDataFileName: state.recentAdDataFileName,
    recentAdDataRows: state.recentAdDataRows,
    recentAdDataStatus: state.recentAdDataStatus,
    recentAdDataError: state.recentAdDataError,
    recentAdDataMatchSummary: state.recentAdDataMatchSummary,
    adjustmentDrafts: state.adjustmentDrafts,
  };
}

const fieldCandidates = {
  entity: ["实体层级", "实体", "记录类型", "Entity", "Record Type"],
  campaignName: [
    "广告活动名称（仅供参考）",
    "广告活动名称(仅供参考)",
    "广告活动名称",
    "Campaign Name (Informational only)",
    "Campaign Name",
  ],
  adGroupName: [
    "广告组名称（仅供参考）",
    "广告组名称(仅供参考)",
    "Ad Group Name (Informational only)",
    "Ad Group Name",
    "广告组名称",
    "广告组",
  ],
  keyword: ["关键词文本", "关键字文本", "关键词", "关键字", "Keyword Text", "Keyword"],
  target: ["商品定位表达式", "投放对象", "Product Targeting Expression", "Targeting Expression", "Target"],
  matchType: ["匹配类型", "Match Type"],
  bid: ["竞价", "出价", "关键词竞价", "关键字竞价", "Bid", "Max Bid", "Keyword Bid"],
  state: ["状态", "投放状态", "State", "Status"],
  impressions: ["展示量", "曝光量", "Impressions"],
  clicks: ["点击量", "点击次数", "Clicks"],
  spend: ["花费", "支出", "Spend"],
  sales: ["销售额", "销售", "Sales", "7 Day Total Sales", "14 Day Total Sales"],
  orders: ["订单数量", "订单", "Orders", "Purchases", "7 Day Total Orders", "14 Day Total Orders"],
};

const recentFieldCandidates = {
  sheetName: ["Sheet 名", "Sheet", "Portfolio", "Campaign Type"],
  campaignName: ["广告活动名称", "广告活动名称（仅供参考）", "Campaign Name", "Campaign"],
  adGroupName: ["广告组名称", "广告组名称（仅供参考）", "Ad Group Name", "Ad Group Name (Informational only)"],
  keyword: ["关键词文本", "关键词", "Keyword Text", "Customer Search Term", "Search Term", "Keyword"],
  target: ["商品投放表达式", "投放对象", "Targeting Expression", "Product Targeting Expression", "Target"],
  matchType: ["匹配类型", "Match Type"],
  impressions: ["展示量", "曝光量", "Impressions"],
  clicks: ["点击量", "点击次数", "Clicks"],
  orders: ["订单量", "订单数量", "Orders", "Purchases"],
  sales: ["销售额", "销售", "Sales", "7 Day Total Sales", "14 Day Total Sales"],
  spend: ["花费", "支出", "Spend"],
  acos: ["ACOS", "ACoS", "Advertising Cost of Sales"],
  roas: ["ROAS", "Return on Ad Spend"],
};

function normalizeHeader(value: string) {
  return value
    .toLowerCase()
    .replace(/\uFEFF/g, "")
    .replace(/[\s()[\]_\-:：,，.。/\\（）]/g, "");
}

function readColumn(row: SheetRow, candidates: string[]) {
  const entries = Object.entries(row).filter(([key]) => !key.startsWith("__"));
  const normalizedEntries = entries.map(([key, value]) => [normalizeHeader(key), value] as const);

  for (const candidate of candidates.map(normalizeHeader)) {
    const exactEntry = normalizedEntries.find(([key]) => key === candidate);

    if (exactEntry) {
      const value = exactEntry[1];
      return value === null || value === undefined ? "" : String(value).trim();
    }
  }

  for (const candidate of candidates.map(normalizeHeader)) {
    const fuzzyEntry = normalizedEntries.find(
      ([key]) => key.includes(candidate) || candidate.includes(key),
    );

    if (fuzzyEntry) {
      const value = fuzzyEntry[1];
      return value === null || value === undefined ? "" : String(value).trim();
    }
  }

  return "";
}

function readNumber(row: SheetRow, candidates: string[]) {
  const value = readColumn(row, candidates).replace(/[$,%￥,]/g, "");
  const numericValue = Number(value);

  return Number.isFinite(numericValue) ? numericValue : 0;
}

function parseCsvLine(line: string) {
  const cells: string[] = [];
  let current = "";
  let quoted = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const nextChar = line[index + 1];

    if (char === '"' && quoted && nextChar === '"') {
      current += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      cells.push(current);
      current = "";
    } else {
      current += char;
    }
  }

  cells.push(current);
  return cells.map((cell) => cell.trim());
}

function parseCsv(text: string): SheetRow[] {
  const lines = text
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .filter((line) => line.trim() !== "");
  const headers = parseCsvLine(lines[0] ?? "");

  return lines.slice(1).map((line) => {
    const values = parseCsvLine(line);

    return headers.reduce<SheetRow>((row, header, index) => {
      row[header] = values[index] ?? "";
      return row;
    }, {});
  });
}

function readLooseColumn(row: SheetRow, candidates: string[]) {
  const entries = Object.entries(row).filter(([key]) => !key.startsWith("__"));
  const normalizedEntries = entries.map(([key, value]) => [normalizeHeader(key), value] as const);

  for (const candidate of candidates.map(normalizeHeader)) {
    const entry = normalizedEntries.find(([key]) => key === candidate);

    if (entry) {
      const value = entry[1];
      return value === null || value === undefined ? "" : String(value).trim();
    }
  }

  return "";
}

function normalizeMatchValue(value: string | undefined) {
  return (value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function buildKeywordMatchKey(keyword: string | undefined, matchType: string | undefined) {
  return `${normalizeMatchValue(keyword)}::${normalizeMatchValue(matchType)}`;
}

function buildRecentAdDataRows(
  fileName: string,
  text: string,
  scopeCampaignGroupIds: string[],
  performanceRows: PerformanceRow[],
) {
  const fileId = `recent-${Date.now()}`;
  const csvRows = parseCsv(text);
  const rows = csvRows
    .map((row, index): RecentAdDataRow | null => {
      const keyword = readColumn(row, recentFieldCandidates.keyword);
      const matchType = readColumn(row, recentFieldCandidates.matchType);

      if (!keyword || !matchType) {
        return null;
      }

      const spend = readNumber(row, recentFieldCandidates.spend);
      const sales = readNumber(row, recentFieldCandidates.sales);
      const acosValue = readNumber(row, recentFieldCandidates.acos);
      const roasValue = readNumber(row, recentFieldCandidates.roas);

      return {
        id: `${fileId}-${index}`,
        fileId,
        scopeCampaignGroupIds,
        sheetName: readColumn(row, recentFieldCandidates.sheetName) || undefined,
        campaignName: readColumn(row, recentFieldCandidates.campaignName) || undefined,
        adGroupName: readColumn(row, recentFieldCandidates.adGroupName) || undefined,
        keyword,
        target: readColumn(row, recentFieldCandidates.target),
        matchType,
        impressions: readNumber(row, recentFieldCandidates.impressions),
        clicks: readNumber(row, recentFieldCandidates.clicks),
        orders: readNumber(row, recentFieldCandidates.orders),
        sales,
        spend,
        acos: acosValue || (sales > 0 ? (spend / sales) * 100 : undefined),
        roas: roasValue || (spend > 0 ? sales / spend : undefined),
        matchStatus: "unmatched",
      };
    })
    .filter((row): row is RecentAdDataRow => Boolean(row));

  return matchRecentAdDataRows(rows, scopeCampaignGroupIds, performanceRows, fileName);
}

function matchRecentAdDataRows(
  rows: RecentAdDataRow[],
  scopeCampaignGroupIds: string[],
  performanceRows: PerformanceRow[],
  fileName: string,
) {
  const scopedPerformanceRows = performanceRows.filter((row) => scopeCampaignGroupIds.includes(row.campaignGroupId));
  const bulkRowsByKeywordAndMatchType = scopedPerformanceRows.reduce<Map<string, PerformanceRow[]>>((map, row) => {
    const key = buildKeywordMatchKey(row.keyword, row.matchType);
    const existingRows = map.get(key) ?? [];

    existingRows.push(row);
    map.set(key, existingRows);
    return map;
  }, new Map());
  const matchedRows = rows.map((row) => {
    const candidates = bulkRowsByKeywordAndMatchType.get(buildKeywordMatchKey(row.keyword, row.matchType)) ?? [];
    const candidateCampaignGroupIds = Array.from(new Set(candidates.map((candidate) => candidate.campaignGroupId)));

    if (candidateCampaignGroupIds.length === 1) {
      const matchedCampaignGroupId = candidateCampaignGroupIds[0];
      const matchedBulkRow = candidates.find((candidate) => candidate.campaignGroupId === matchedCampaignGroupId);

      return {
        ...row,
        sheetName: row.sheetName ?? matchedBulkRow?.sheetName,
        campaignName: row.campaignName ?? matchedBulkRow?.campaignName,
        adGroupName: row.adGroupName ?? matchedBulkRow?.adGroupName,
        campaignGroupId: matchedCampaignGroupId,
        matchStatus: "matched" as const,
        matchError: undefined,
      };
    }

    if (candidateCampaignGroupIds.length > 1) {
      return { ...row, matchStatus: "ambiguous" as const, matchError: "关键词和匹配类型命中多个广告组，请在单个广告组页面上传" };
    }

    return { ...row, matchStatus: "unmatched" as const, matchError: "找不到关键词 + 匹配类型完全一致的 Bulk 行" };
  });
  const matchedCampaignGroups = new Set(matchedRows.flatMap((row) => (row.campaignGroupId ? [row.campaignGroupId] : []))).size;
  const summary: RecentAdDataMatchSummary = {
    totalRows: matchedRows.length,
    matchedRows: matchedRows.filter((row) => row.matchStatus === "matched").length,
    unmatchedRows: matchedRows.filter((row) => row.matchStatus === "unmatched").length,
    ambiguousRows: matchedRows.filter((row) => row.matchStatus === "ambiguous").length,
    matchedCampaignGroups,
    scopedCampaignGroups: scopeCampaignGroupIds.length,
  };

  return {
    fileName,
    rows: matchedRows,
    summary,
  };
}

function buildCampaignGroupId(sheetName: string, adGroupName: string) {
  return `${sheetName}::${adGroupName}`.toLowerCase().replace(/[^a-z0-9\u4e00-\u9fff]+/g, "-");
}

function createWorkspaceUnitId() {
  return `workspace-unit-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function buildWorkspaceUnitName(groups: CampaignGroup[]) {
  if (groups.length === 0) {
    return "Manual Workspace Unit";
  }

  if (groups.length === 1) {
    return groups[0].adGroupName;
  }

  return `${groups[0].adGroupName} +${groups.length - 1} 个分组`;
}

function upsertWorkspaceUnitForCampaign(
  workspaceUnits: WorkspaceUnit[],
  campaignGroups: CampaignGroup[],
  campaignGroupId: string,
  workspaceUnitId: string,
) {
  const groupIds = new Set<string>();

  return workspaceUnits.map((unit) => {
    if (unit.id === workspaceUnitId) {
      unit.campaignGroupIds.forEach((id) => groupIds.add(id));
      groupIds.add(campaignGroupId);
      const nextCampaignGroupIds = Array.from(groupIds);
      const nextGroups = campaignGroups.filter((group) => nextCampaignGroupIds.includes(group.id));

      return {
        ...unit,
        campaignGroupIds: nextCampaignGroupIds,
        name: buildWorkspaceUnitName(nextGroups),
        updatedAt: new Date().toISOString(),
      };
    }

    return {
      ...unit,
      campaignGroupIds: unit.campaignGroupIds.filter((id) => id !== campaignGroupId),
    };
  }).filter((unit) => unit.campaignGroupIds.length > 0);
}

function detachCampaignGroupFromWorkspaceUnits(workspaceUnits: WorkspaceUnit[], campaignGroupId: string) {
  return workspaceUnits
    .map((unit) => {
      const nextCampaignGroupIds = unit.campaignGroupIds.filter((id) => id !== campaignGroupId);

      return {
        ...unit,
        campaignGroupIds: nextCampaignGroupIds,
      };
    })
    .filter((unit) => unit.campaignGroupIds.length > 0);
}

function rebuildWorkspaceUnits(workspaceUnits: WorkspaceUnit[], campaignGroups: CampaignGroup[]) {
  return workspaceUnits
    .map((unit) => {
      const groups = campaignGroups.filter((group) => unit.campaignGroupIds.includes(group.id));

      return {
        ...unit,
        name: buildWorkspaceUnitName(groups),
        updatedAt: new Date().toISOString(),
      };
    })
    .filter((unit) => unit.campaignGroupIds.length > 1);
}

function isLifecycleGroupId(value: string): value is LifecycleGroupId {
  return ["launch", "mature", "decline", "clearance"].includes(value);
}

function buildCampaignGroupLookup(groups: CampaignGroup[]) {
  return groups.reduce<Map<string, string>>((lookup, group) => {
    [
      group.id,
      group.adGroupName,
      `${group.sheetName ?? ""}::${group.adGroupName}`,
      `${group.campaignName}::${group.adGroupName}`,
    ].forEach((value) => {
      const normalized = normalizeMatchValue(value);

      if (normalized) {
        lookup.set(normalized, group.id);
      }
    });

    return lookup;
  }, new Map());
}

function buildWorkspaceUnitsFromGroupingRows(rows: SheetRow[], groups: CampaignGroup[]) {
  const groupLookup = buildCampaignGroupLookup(groups);
  const unitRows = new Map<string, string[]>();
  const lifecycleByCampaignGroupId = new Map<string, LifecycleGroupId | undefined>();
  let importedRows = 0;

  rows.forEach((row) => {
    const campaignGroupKey =
      readLooseColumn(row, ["campaignGroupId", "campaign_group_id", "id"]) ||
      readLooseColumn(row, ["adGroupName", "ad_group_name", "Ad Group Name"]) ||
      readLooseColumn(row, ["campaignName", "campaign_name", "Campaign Name"]);
    const campaignGroupId = groupLookup.get(normalizeMatchValue(campaignGroupKey));

    if (!campaignGroupId) {
      return;
    }

    importedRows += 1;

    const rawLifecycle = readLooseColumn(row, ["lifecycleGroup", "lifecycleGroupId", "lifecycle", "生命周期分组"]).toLowerCase();
    lifecycleByCampaignGroupId.set(campaignGroupId, isLifecycleGroupId(rawLifecycle) ? rawLifecycle : undefined);

    const workspaceUnitName = readLooseColumn(row, ["workspaceUnit", "workspaceUnitId", "workspace", "分组"]);

    if (!workspaceUnitName) {
      return;
    }

    const workspaceUnitKey = normalizeMatchValue(workspaceUnitName);
    const currentIds = unitRows.get(workspaceUnitKey) ?? [];
    unitRows.set(workspaceUnitKey, Array.from(new Set([...currentIds, campaignGroupId])));
  });

  const now = new Date().toISOString();
  const workspaceUnits = Array.from(unitRows.entries()).flatMap(([key, campaignGroupIds], index): WorkspaceUnit[] => {
    const validIds = campaignGroupIds.filter((id) => groups.some((group) => group.id === id));

    if (validIds.length < 2) {
      return [];
    }

    const unitGroups = groups.filter((group) => validIds.includes(group.id));

    return [
      {
        id: `imported-workspace-unit-${index + 1}-${key.replace(/[^a-z0-9\u4e00-\u9fff]+/g, "-").slice(0, 32)}`,
        name: buildWorkspaceUnitName(unitGroups),
        campaignGroupIds: validIds,
        createdAt: now,
        updatedAt: now,
      },
    ];
  });

  return { importedRows, lifecycleByCampaignGroupId, workspaceUnits };
}

function findWorkspaceUnitByCampaignGroupId(workspaceUnits: WorkspaceUnit[], campaignGroupId: string) {
  return workspaceUnits.find((unit) => unit.campaignGroupIds.includes(campaignGroupId));
}

function buildSheetGroups(groups: CampaignGroup[]) {
  return groups.reduce<CampaignSheetGroup[]>((sheetGroups, group) => {
    const sheetName = group.sheetName ?? "Mock Sheet";
    const existingSheet = sheetGroups.find((item) => item.sheetName === sheetName);

    if (existingSheet) {
      existingSheet.groups.push(group);
    } else {
      sheetGroups.push({ sheetName, groups: [group] });
    }

    return sheetGroups;
  }, []);
}

function isSponsoredProductsCampaignSheet(sheetName: string) {
  const normalized = normalizeHeader(sheetName);

  return (
    normalized.includes(normalizeHeader("商品推广活动")) ||
    normalized.includes(normalizeHeader("Sponsored Products Campaigns")) ||
    normalized.includes(normalizeHeader("Sponsored Products"))
  );
}

function isKeywordEntity(entity: string) {
  const normalized = normalizeHeader(entity);

  return (
    normalized === normalizeHeader("关键词") ||
    normalized === normalizeHeader("关键字") ||
    normalized === "keyword" ||
    normalized === "keywords"
  );
}

function isExecutableKeywordRow(row: SheetRow, sheetName: string) {
  const entity = readColumn(row, fieldCandidates.entity);
  const adGroupName = readColumn(row, fieldCandidates.adGroupName);
  const bid = readColumn(row, fieldCandidates.bid);

  return isSponsoredProductsCampaignSheet(sheetName) && isKeywordEntity(entity) && adGroupName !== "" && bid !== "";
}

function toPerformanceRow(row: SheetRow, sheetName: string, batchId: string, fallbackSourceRowIndex: number): PerformanceRow | null {
  const sourceRowIndex = Number(row.__sourceRowIndex ?? fallbackSourceRowIndex);
  const adGroupName = readColumn(row, fieldCandidates.adGroupName);

  if (!isExecutableKeywordRow(row, sheetName)) {
    return null;
  }

  const campaignName = readColumn(row, fieldCandidates.campaignName) || sheetName;
  const keyword = readColumn(row, fieldCandidates.keyword) || readColumn(row, fieldCandidates.target) || "未命名关键词";
  const campaignGroupId = buildCampaignGroupId(sheetName, adGroupName);
  const currentBid = readNumber(row, fieldCandidates.bid);
  const state = readColumn(row, fieldCandidates.state);

  return {
    id: `${batchId}-${sheetName}-${sourceRowIndex}`,
    batchId,
    sheetName,
    sourceRowIndex,
    sourceRowNumber: sourceRowIndex,
    campaignGroupId,
    entity: readColumn(row, fieldCandidates.entity),
    adGroupNameRef: adGroupName,
    campaignName,
    adGroupName,
    keyword,
    target: readColumn(row, fieldCandidates.target) || keyword,
    matchType: readColumn(row, fieldCandidates.matchType) || "-",
    currentBid,
    impressions: readNumber(row, fieldCandidates.impressions),
    clicks: readNumber(row, fieldCandidates.clicks),
    orders: readNumber(row, fieldCandidates.orders),
    sales: readNumber(row, fieldCandidates.sales),
    spend: readNumber(row, fieldCandidates.spend),
    topOfSearchShare: 0,
    advertisedProductOrders: 0,
    otherProductOrders: 0,
    viewableImpressions: 0,
    status: state.includes("暂停") || state.toLowerCase() === "paused" ? "paused" : "enabled",
  };
}

function collectDiagnostics(sheetName: string, rows: SheetRow[], current: ParseDiagnostics): ParseDiagnostics {
  const sampleHeaders = current.sampleHeaders.length
    ? current.sampleHeaders
    : Object.keys(rows[0] ?? {})
        .filter((key) => !key.startsWith("__"))
        .slice(0, 16);
  const sampleEntities = new Set(current.sampleEntities);
  let sponsoredProductRows = current.sponsoredProductRows;
  let rowsWithAdGroup = current.rowsWithAdGroup;
  let keywordRows = current.keywordRows;
  let rowsWithBid = current.rowsWithBid;
  let executableRows = current.executableRows;

  for (const row of rows) {
    const entity = readColumn(row, fieldCandidates.entity);
    const hasAdGroup = readColumn(row, fieldCandidates.adGroupName) !== "";
    const hasBid = readColumn(row, fieldCandidates.bid) !== "";
    const isSpSheet = isSponsoredProductsCampaignSheet(sheetName);
    const isKeyword = isKeywordEntity(entity);

    if (entity && sampleEntities.size < 10) {
      sampleEntities.add(entity);
    }
    if (isSpSheet) sponsoredProductRows += 1;
    if (hasAdGroup) rowsWithAdGroup += 1;
    if (isKeyword) keywordRows += 1;
    if (hasBid) rowsWithBid += 1;
    if (isSpSheet && isKeyword && hasAdGroup && hasBid) executableRows += 1;
  }

  return {
    totalRows: current.totalRows + rows.length,
    sponsoredProductRows,
    rowsWithAdGroup,
    keywordRows,
    rowsWithBid,
    executableRows,
    sampleHeaders,
    sampleEntities: Array.from(sampleEntities),
  };
}

function buildParseFailureMessage(diagnostics: ParseDiagnostics) {
  const headers = diagnostics.sampleHeaders.length ? diagnostics.sampleHeaders.join("、") : "未读取到表头";
  const entities = diagnostics.sampleEntities.length ? diagnostics.sampleEntities.join("、") : "未读取到实体层级值";

  return [
    "已解析文件，但未找到可执行关键词行。",
    "MVP 仅处理“商品推广活动 / Sponsored Products Campaigns”中实体层级为“关键词/关键字/Keyword”且竞价不为空的行。",
    `诊断：总行 ${diagnostics.totalRows}，商品推广 Sheet 行 ${diagnostics.sponsoredProductRows}，有广告组 ${diagnostics.rowsWithAdGroup}，关键词实体 ${diagnostics.keywordRows}，有竞价 ${diagnostics.rowsWithBid}，可执行 ${diagnostics.executableRows}。`,
    `识别到的表头示例：${headers}。`,
    `实体层级示例：${entities}。`,
  ].join(" ");
}

function buildGroupsFromRows(existingGroups: CampaignGroup[], rows: PerformanceRow[]) {
  const groupMap = new Map<string, CampaignGroup>();

  for (const group of existingGroups) {
    groupMap.set(group.id, group);
  }

  for (const row of rows) {
    const existingGroup = groupMap.get(row.campaignGroupId);
    groupMap.set(row.campaignGroupId, {
      id: row.campaignGroupId,
      sheetName: row.sheetName,
      campaignName: existingGroup?.campaignName ?? row.campaignName,
      adGroupName: row.adGroupName,
      lifecycleGroupId: existingGroup?.lifecycleGroupId,
      keywordCount: (existingGroup?.keywordCount ?? 0) + 1,
      lastUpdated: new Date().toISOString().slice(0, 10),
    });
  }

  return Array.from(groupMap.values()).sort((left, right) => {
    const sheetCompare = (left.sheetName ?? "").localeCompare(right.sheetName ?? "", "zh-CN");
    return sheetCompare || left.adGroupName.localeCompare(right.adGroupName, "zh-CN");
  });
}

export const useWorkspaceStore = create<WorkspaceState>((set, get) => ({
  rules: defaultRules,
  campaignGroups,
  campaignSheetGroups: buildSheetGroups(campaignGroups),
  workspaceUnits: [],
  performanceRows: mockPerformanceRows,
  activeCampaignGroupId: initialActiveId,
  activeWorkspaceUnitId: undefined,
  activeLifecycleGroupId: undefined,
  workspaceMode: "campaign",
  openTabIds: campaignGroups.slice(0, 4).map((group) => group.id),
  selectedDraftIds: [],
  parseStatus: "idle",
  parseProgress: 0,
  uploadedFileName: undefined,
  originalWorkbookBuffer: undefined,
  activeBatchId: undefined,
  parsedRowCount: 0,
  parsedSheets: [],
  parseError: undefined,
  parseDiagnostics: emptyDiagnostics,
  recentAdDataFileName: undefined,
  recentAdDataRows: [],
  recentAdDataStatus: "idle",
  recentAdDataError: undefined,
  recentAdDataMatchSummary: emptyRecentAdDataMatchSummary,
  adjustmentDrafts: [],
  persistenceStatus: "loading",
  persistenceError: undefined,
  setRules: (rules) => set({ rules }),
  upsertRule: (rule) =>
    set((state) => ({
      rules: state.rules.some((existingRule) => existingRule.id === rule.id)
        ? state.rules.map((existingRule) => (existingRule.id === rule.id ? rule : existingRule))
        : [...state.rules, rule],
    })),
  deleteRule: (ruleId) =>
    set((state) => ({
      rules: state.rules.filter((rule) => rule.id !== ruleId),
    })),
  setActiveCampaignGroup: (campaignGroupId) =>
    set((state) => ({
      activeCampaignGroupId: campaignGroupId,
      activeWorkspaceUnitId: undefined,
      workspaceMode: "campaign",
      openTabIds: state.openTabIds.includes(campaignGroupId)
        ? state.openTabIds
        : [...state.openTabIds, campaignGroupId],
    })),
  openCampaignGroup: (campaignGroupId) =>
    set((state) => ({
      activeCampaignGroupId: campaignGroupId,
      activeWorkspaceUnitId: undefined,
      activeLifecycleGroupId: undefined,
      workspaceMode: "campaign",
      openTabIds: state.openTabIds.includes(campaignGroupId)
        ? state.openTabIds
        : [campaignGroupId, ...state.openTabIds].slice(0, 12),
      adjustmentDrafts: [],
      selectedDraftIds: [],
    })),
  mergeCampaignGroupsIntoWorkspaceUnit: (sourceCampaignGroupId, targetCampaignGroupId) =>
    set((state) => {
      if (sourceCampaignGroupId === targetCampaignGroupId) {
        return state;
      }

      const sourceGroup = state.campaignGroups.find((group) => group.id === sourceCampaignGroupId);
      const targetGroup = state.campaignGroups.find((group) => group.id === targetCampaignGroupId);

      if (!sourceGroup || !targetGroup) {
        return state;
      }

      const targetUnit = findWorkspaceUnitByCampaignGroupId(state.workspaceUnits, targetCampaignGroupId);
      const sourceUnit = findWorkspaceUnitByCampaignGroupId(state.workspaceUnits, sourceCampaignGroupId);
      let workspaceUnits = detachCampaignGroupFromWorkspaceUnits(state.workspaceUnits, sourceCampaignGroupId);

      if (targetUnit) {
        workspaceUnits = upsertWorkspaceUnitForCampaign(workspaceUnits, state.campaignGroups, sourceCampaignGroupId, targetUnit.id);
      } else {
        const id = sourceUnit?.id ?? createWorkspaceUnitId();
        const campaignGroupIds = Array.from(new Set([targetCampaignGroupId, sourceCampaignGroupId]));
        const groups = state.campaignGroups.filter((group) => campaignGroupIds.includes(group.id));

        workspaceUnits.push({
          id,
          name: buildWorkspaceUnitName(groups),
          campaignGroupIds,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        });
      }

      const activeWorkspaceUnit =
        findWorkspaceUnitByCampaignGroupId(workspaceUnits, targetCampaignGroupId) ??
        findWorkspaceUnitByCampaignGroupId(workspaceUnits, sourceCampaignGroupId);

      return {
        workspaceUnits,
        activeWorkspaceUnitId: activeWorkspaceUnit?.id ?? state.activeWorkspaceUnitId,
        activeCampaignGroupId: targetCampaignGroupId,
        activeLifecycleGroupId: undefined,
        workspaceMode: activeWorkspaceUnit ? "workspace-unit" : state.workspaceMode,
        openTabIds: Array.from(new Set([targetCampaignGroupId, sourceCampaignGroupId, ...state.openTabIds])).slice(0, 12),
        adjustmentDrafts: [],
        selectedDraftIds: [],
      };
    }),
  removeCampaignGroupFromWorkspaceUnit: (campaignGroupId) =>
    set((state) => {
      const existingUnit = findWorkspaceUnitByCampaignGroupId(state.workspaceUnits, campaignGroupId);

      if (!existingUnit) {
        return state;
      }

      const workspaceUnits = rebuildWorkspaceUnits(
        detachCampaignGroupFromWorkspaceUnits(state.workspaceUnits, campaignGroupId),
        state.campaignGroups,
      );
      const nextActiveWorkspaceUnit =
        state.activeWorkspaceUnitId && existingUnit.id === state.activeWorkspaceUnitId
          ? workspaceUnits.find((unit) => unit.id === existingUnit.id)
          : undefined;
      const nextActiveCampaignGroupId =
        nextActiveWorkspaceUnit?.campaignGroupIds[0] ??
        (state.activeCampaignGroupId === campaignGroupId ? existingUnit.campaignGroupIds.find((id) => id !== campaignGroupId) : undefined) ??
        state.activeCampaignGroupId;

      return {
        workspaceUnits,
        activeWorkspaceUnitId: nextActiveWorkspaceUnit?.id,
        activeCampaignGroupId: nextActiveCampaignGroupId,
        workspaceMode: nextActiveWorkspaceUnit ? "workspace-unit" : "campaign",
        openTabIds: state.openTabIds.filter((id) => id !== campaignGroupId),
        adjustmentDrafts: [],
        selectedDraftIds: [],
      };
    }),
  setActiveWorkspaceUnit: (workspaceUnitId) =>
    set((state) => {
      const workspaceUnit = state.workspaceUnits.find((unit) => unit.id === workspaceUnitId);
      const activeCampaignGroupId = workspaceUnit?.campaignGroupIds[0] ?? state.activeCampaignGroupId;

      return workspaceUnit
        ? {
            activeWorkspaceUnitId: workspaceUnitId,
            activeCampaignGroupId,
            activeLifecycleGroupId: undefined,
            workspaceMode: "workspace-unit" as const,
            openTabIds: Array.from(new Set([...workspaceUnit.campaignGroupIds, ...state.openTabIds])).slice(0, 12),
            adjustmentDrafts: [],
            selectedDraftIds: [],
          }
        : state;
    }),
  setActiveLifecycleGroup: (lifecycleGroupId) =>
    set((state) => {
      if (!lifecycleGroupId) {
        return {
          activeLifecycleGroupId: undefined,
          activeWorkspaceUnitId: undefined,
          workspaceMode: "campaign" as const,
          adjustmentDrafts: [],
          selectedDraftIds: [],
        };
      }

      const assignedGroups = state.campaignGroups.filter((group) => group.lifecycleGroupId === lifecycleGroupId);
      const activeCampaignGroupId = assignedGroups.some((group) => group.id === state.activeCampaignGroupId)
        ? state.activeCampaignGroupId
        : assignedGroups[0]?.id ?? "";

      return {
        activeLifecycleGroupId: lifecycleGroupId,
        activeCampaignGroupId,
        activeWorkspaceUnitId: undefined,
        workspaceMode: "lifecycle",
        openTabIds: assignedGroups.map((group) => group.id),
        adjustmentDrafts: [],
        selectedDraftIds: [],
      };
    }),
  assignLifecycleGroup: (campaignGroupId, lifecycleGroupId) =>
    set((state) => {
      const campaignGroups = state.campaignGroups.map((group) =>
        group.id === campaignGroupId ? { ...group, lifecycleGroupId } : group,
      );
      const assignedGroups = campaignGroups.filter((group) => group.lifecycleGroupId === lifecycleGroupId);

      return {
        campaignGroups,
        campaignSheetGroups: buildSheetGroups(campaignGroups),
        activeLifecycleGroupId: state.workspaceMode === "lifecycle" ? lifecycleGroupId : state.activeLifecycleGroupId,
        openTabIds:
          state.workspaceMode === "lifecycle"
            ? assignedGroups.map((group) => group.id)
            : state.openTabIds.includes(campaignGroupId)
              ? state.openTabIds
              : [...state.openTabIds, campaignGroupId],
        activeCampaignGroupId: campaignGroupId,
      };
    }),
  importGroupingStatusCsv: (_fileName, text) => {
    const rows = parseCsv(text);
    const result = buildWorkspaceUnitsFromGroupingRows(rows, get().campaignGroups);

    set((state) => {
      const campaignGroups = state.campaignGroups.map((group) =>
        result.lifecycleByCampaignGroupId.has(group.id)
          ? { ...group, lifecycleGroupId: result.lifecycleByCampaignGroupId.get(group.id) }
          : group,
      );
      const activeWorkspaceUnit = result.workspaceUnits.find((unit) => unit.campaignGroupIds.includes(state.activeCampaignGroupId));

      return {
        campaignGroups,
        campaignSheetGroups: buildSheetGroups(campaignGroups),
        workspaceUnits: result.workspaceUnits,
        activeWorkspaceUnitId: activeWorkspaceUnit?.id,
        workspaceMode: activeWorkspaceUnit ? "workspace-unit" : "campaign",
        adjustmentDrafts: [],
        selectedDraftIds: [],
      };
    });

    return {
      importedRows: result.importedRows,
      workspaceUnitCount: result.workspaceUnits.length,
    };
  },
  runRulesForActiveGroup: () => {
    const state = get();
    const campaignGroup = state.campaignGroups.find((group) => group.id === state.activeCampaignGroupId);

    if (!campaignGroup) {
      return;
    }

    const scopedBatchIds = dataBatches
      .filter((batch) => batch.campaignGroupId === campaignGroup.id)
      .slice(-1)
      .map((batch) => batch.id);
    const rows = state.performanceRows.filter((row) => {
      const isActiveGroup = row.campaignGroupId === campaignGroup.id;
      const isImportedBatch = state.activeBatchId ? row.batchId === state.activeBatchId : false;
      const isMockBatch = scopedBatchIds.includes(row.batchId);

      return isActiveGroup && (isImportedBatch || isMockBatch);
    });
    const drafts = runRuleEngine({
      campaignGroup,
      rows,
      recentAdDataRows: state.recentAdDataRows,
      rules: state.rules,
    });

    set({
      adjustmentDrafts: drafts,
      selectedDraftIds: drafts.filter((draft) => draft.selected).map((draft) => draft.id),
    });
  },
  runRulesForActiveLifecycleGroup: () => {
    const state = get();
    const lifecycleGroupId = state.activeLifecycleGroupId;

    if (!lifecycleGroupId) {
      return;
    }

    const campaignGroups = state.campaignGroups.filter((group) => group.lifecycleGroupId === lifecycleGroupId);
    const drafts = campaignGroups.flatMap((campaignGroup) => {
      const scopedBatchIds = dataBatches
        .filter((batch) => batch.campaignGroupId === campaignGroup.id)
        .slice(-1)
        .map((batch) => batch.id);
      const rows = state.performanceRows.filter((row) => {
        const isActiveGroup = row.campaignGroupId === campaignGroup.id;
        const isImportedBatch = state.activeBatchId ? row.batchId === state.activeBatchId : false;
        const isMockBatch = scopedBatchIds.includes(row.batchId);

        return isActiveGroup && (isImportedBatch || isMockBatch);
      });

      return runRuleEngine({
        campaignGroup,
        rows,
        recentAdDataRows: state.recentAdDataRows,
        rules: state.rules,
      });
    });

    set({
      adjustmentDrafts: drafts,
      selectedDraftIds: drafts.filter((draft) => draft.selected).map((draft) => draft.id),
    });
  },
  runRulesForActiveWorkspaceUnit: () => {
    const state = get();
    const workspaceUnit = state.workspaceUnits.find((unit) => unit.id === state.activeWorkspaceUnitId);

    if (!workspaceUnit) {
      return;
    }

    const drafts = workspaceUnit.campaignGroupIds.flatMap((campaignGroupId) => {
      const campaignGroup = state.campaignGroups.find((group) => group.id === campaignGroupId);

      if (!campaignGroup) {
        return [];
      }

      const scopedBatchIds = dataBatches
        .filter((batch) => batch.campaignGroupId === campaignGroup.id)
        .slice(-1)
        .map((batch) => batch.id);
      const rows = state.performanceRows.filter((row) => {
        const isActiveGroup = row.campaignGroupId === campaignGroup.id;
        const isImportedBatch = state.activeBatchId ? row.batchId === state.activeBatchId : false;
        const isMockBatch = scopedBatchIds.includes(row.batchId);

        return isActiveGroup && (isImportedBatch || isMockBatch);
      });

      return runRuleEngine({
        campaignGroup,
        rows,
        recentAdDataRows: state.recentAdDataRows.filter(
          (row) => !row.campaignGroupId || workspaceUnit.campaignGroupIds.includes(row.campaignGroupId),
        ),
        rules: state.rules,
      });
    });

    set({
      adjustmentDrafts: drafts,
      selectedDraftIds: drafts.filter((draft) => draft.selected).map((draft) => draft.id),
    });
  },
  toggleDraft: (draftId) =>
    set((state) => {
      const selectedDraftIds = state.selectedDraftIds.includes(draftId)
        ? state.selectedDraftIds.filter((id) => id !== draftId)
        : [...state.selectedDraftIds, draftId];

      return {
        selectedDraftIds,
        adjustmentDrafts: state.adjustmentDrafts.map((draft) => ({
          ...draft,
          selected: selectedDraftIds.includes(draft.id),
        })),
      };
    }),
  setDraftSelected: (draftId, selected) =>
    set((state) => {
      const selectedDraftIds = selected
        ? state.selectedDraftIds.includes(draftId)
          ? state.selectedDraftIds
          : [...state.selectedDraftIds, draftId]
        : state.selectedDraftIds.filter((id) => id !== draftId);

      return {
        selectedDraftIds,
        adjustmentDrafts: state.adjustmentDrafts.map((draft) => ({
          ...draft,
          selected: selectedDraftIds.includes(draft.id),
        })),
      };
    }),
  selectAllDrafts: () =>
    set((state) => ({
      selectedDraftIds: state.adjustmentDrafts.map((draft) => draft.id),
      adjustmentDrafts: state.adjustmentDrafts.map((draft) => ({ ...draft, selected: true })),
    })),
  invertDraftSelection: () =>
    set((state) => {
      const selectedDraftIds = state.adjustmentDrafts
        .filter((draft) => !state.selectedDraftIds.includes(draft.id))
        .map((draft) => draft.id);

      return {
        selectedDraftIds,
        adjustmentDrafts: state.adjustmentDrafts.map((draft) => ({
          ...draft,
          selected: selectedDraftIds.includes(draft.id),
        })),
      };
    }),
  clearDraftSelection: () =>
    set((state) => ({
      selectedDraftIds: [],
      adjustmentDrafts: state.adjustmentDrafts.map((draft) => ({ ...draft, selected: false })),
    })),
  setParseStarted: (fileName, originalWorkbookBuffer) => {
    const batchId = `batch-${Date.now()}`;
    set({
      campaignGroups: [],
      campaignSheetGroups: [],
      workspaceUnits: [],
      rules: defaultRules,
      performanceRows: [],
      activeCampaignGroupId: "",
      activeWorkspaceUnitId: undefined,
      activeLifecycleGroupId: undefined,
      workspaceMode: "campaign",
      openTabIds: [],
      adjustmentDrafts: [],
      selectedDraftIds: [],
      parseStatus: "parsing",
      parseProgress: 0,
      uploadedFileName: fileName,
      originalWorkbookBuffer,
      activeBatchId: batchId,
      parsedRowCount: 0,
      parsedSheets: [],
      parseError: undefined,
      parseDiagnostics: emptyDiagnostics,
      recentAdDataFileName: undefined,
      recentAdDataRows: [],
      recentAdDataStatus: "idle",
      recentAdDataError: undefined,
      recentAdDataMatchSummary: emptyRecentAdDataMatchSummary,
    });
  },
  setParseProgress: (progress, sheets) =>
    set((state) => ({
      parseStatus: "parsing",
      parseProgress: progress,
      parsedSheets: sheets?.length ? sheets : state.parsedSheets,
    })),
  ingestParsedRows: (sheetName, rows, startRowIndex) =>
    set((state) => {
      const batchId = state.activeBatchId ?? `batch-${Date.now()}`;
      const executableRows = rows
        .map((row, index) => toPerformanceRow(row, sheetName, batchId, startRowIndex + index + 2))
        .filter((row): row is PerformanceRow => Boolean(row));
      const performanceRows = [...state.performanceRows, ...executableRows];
      const campaignGroups = buildGroupsFromRows(state.campaignGroups, executableRows);
      const activeCampaignGroupId =
        state.activeCampaignGroupId && campaignGroups.some((group) => group.id === state.activeCampaignGroupId)
          ? state.activeCampaignGroupId
          : campaignGroups[0]?.id ?? "";

      return {
        performanceRows,
        campaignGroups,
        campaignSheetGroups: buildSheetGroups(campaignGroups),
        activeCampaignGroupId,
        parseDiagnostics: collectDiagnostics(sheetName, rows, state.parseDiagnostics),
        openTabIds: activeCampaignGroupId
          ? state.openTabIds.includes(activeCampaignGroupId)
            ? state.openTabIds
            : [activeCampaignGroupId, ...state.openTabIds].slice(0, 4)
          : [],
      };
    }),
  setParseCompleted: (rowCount, sheets) =>
    set((state) => {
      const recentMatch = state.recentAdDataRows.length
        ? matchRecentAdDataRows(
            state.recentAdDataRows,
            state.recentAdDataRows[0]?.scopeCampaignGroupIds ?? state.campaignGroups.map((group) => group.id),
            state.performanceRows,
            state.recentAdDataFileName ?? "",
          )
        : undefined;

      return {
        parseStatus: "completed",
        parseProgress: 100,
        parsedRowCount: rowCount,
        parsedSheets: sheets,
        parseError: state.campaignGroups.length ? undefined : buildParseFailureMessage(state.parseDiagnostics),
        campaignSheetGroups: buildSheetGroups(state.campaignGroups),
        recentAdDataRows: recentMatch?.rows ?? state.recentAdDataRows,
        recentAdDataMatchSummary: recentMatch?.summary ?? state.recentAdDataMatchSummary,
      };
    }),
  setParseFailed: (message) =>
    set({
      parseStatus: "failed",
      parseError: message,
    }),
  ingestRecentAdDataCsv: (fileName, text, scopeCampaignGroupIds) =>
    set((state) => {
      try {
        const result = buildRecentAdDataRows(fileName, text, scopeCampaignGroupIds, state.performanceRows);
        const hasFatalMismatch = result.summary.totalRows > 0 && result.summary.matchedRows === 0;

        return {
          recentAdDataFileName: result.fileName,
          recentAdDataRows: result.rows,
          recentAdDataStatus: hasFatalMismatch ? "failed" : "matched",
          recentAdDataError: hasFatalMismatch ? "近期广告数据未匹配到任何 Bulk 行，请检查关键词和匹配类型是否完全一致。" : undefined,
          recentAdDataMatchSummary: result.summary,
          adjustmentDrafts: [],
          selectedDraftIds: [],
        };
      } catch (error) {
        return {
          recentAdDataFileName: fileName,
          recentAdDataRows: [],
          recentAdDataStatus: "failed",
          recentAdDataError: error instanceof Error ? error.message : "近期广告数据 CSV 解析失败。",
          recentAdDataMatchSummary: emptyRecentAdDataMatchSummary,
        };
      }
    }),
  hydratePersistedWorkspace: async () => {
    try {
      const persisted = await readWorkspaceSnapshot<WorkspaceSnapshot>();

      if (!persisted?.snapshot) {
        set({ persistenceStatus: "ready", persistenceError: undefined });
        return;
      }

      set({
        ...persisted.snapshot,
        parseStatus: persisted.snapshot.parseStatus === "parsing" ? "completed" : persisted.snapshot.parseStatus,
        parseProgress: persisted.snapshot.parseStatus === "parsing" ? 100 : persisted.snapshot.parseProgress,
        campaignSheetGroups: buildSheetGroups(persisted.snapshot.campaignGroups),
        persistenceStatus: "ready",
        persistenceError: undefined,
      });
    } catch (error) {
      set({
        persistenceStatus: "failed",
        persistenceError: error instanceof Error ? error.message : "恢复本地工作区失败。",
      });
    }
  },
  clearPersistedWorkspace: async () => {
    try {
      await deleteWorkspaceSnapshot();
      set({
        campaignGroups,
        campaignSheetGroups: buildSheetGroups(campaignGroups),
        workspaceUnits: [],
        rules: defaultRules,
        performanceRows: mockPerformanceRows,
        activeCampaignGroupId: initialActiveId,
        activeWorkspaceUnitId: undefined,
        activeLifecycleGroupId: undefined,
        workspaceMode: "campaign",
        openTabIds: campaignGroups.slice(0, 4).map((group) => group.id),
        selectedDraftIds: [],
        parseStatus: "idle",
        parseProgress: 0,
        uploadedFileName: undefined,
        originalWorkbookBuffer: undefined,
        activeBatchId: undefined,
        parsedRowCount: 0,
        parsedSheets: [],
        parseError: undefined,
        parseDiagnostics: emptyDiagnostics,
        recentAdDataFileName: undefined,
        recentAdDataRows: [],
        recentAdDataStatus: "idle",
        recentAdDataError: undefined,
        recentAdDataMatchSummary: emptyRecentAdDataMatchSummary,
        adjustmentDrafts: [],
        persistenceStatus: "ready",
        persistenceError: undefined,
      });
    } catch (error) {
      set({
        persistenceStatus: "failed",
        persistenceError: error instanceof Error ? error.message : "清空本地工作区失败。",
      });
    }
  },
}));

let saveTimer: ReturnType<typeof setTimeout> | undefined;
let hydrated = false;

if (typeof window !== "undefined") {
  useWorkspaceStore.getState().hydratePersistedWorkspace().finally(() => {
    hydrated = true;
  });

  useWorkspaceStore.subscribe((state, previousState) => {
    if (!hydrated || state.persistenceStatus === "loading") {
      return;
    }

    const shouldSave =
      state.campaignGroups !== previousState.campaignGroups ||
      state.rules !== previousState.rules ||
      state.workspaceUnits !== previousState.workspaceUnits ||
      state.performanceRows !== previousState.performanceRows ||
      state.activeCampaignGroupId !== previousState.activeCampaignGroupId ||
      state.activeLifecycleGroupId !== previousState.activeLifecycleGroupId ||
      state.workspaceMode !== previousState.workspaceMode ||
      state.openTabIds !== previousState.openTabIds ||
      state.selectedDraftIds !== previousState.selectedDraftIds ||
      state.parseStatus !== previousState.parseStatus ||
      state.parseProgress !== previousState.parseProgress ||
      state.uploadedFileName !== previousState.uploadedFileName ||
      state.originalWorkbookBuffer !== previousState.originalWorkbookBuffer ||
      state.activeBatchId !== previousState.activeBatchId ||
      state.parsedRowCount !== previousState.parsedRowCount ||
      state.parsedSheets !== previousState.parsedSheets ||
      state.parseError !== previousState.parseError ||
      state.parseDiagnostics !== previousState.parseDiagnostics ||
      state.recentAdDataFileName !== previousState.recentAdDataFileName ||
      state.recentAdDataRows !== previousState.recentAdDataRows ||
      state.recentAdDataStatus !== previousState.recentAdDataStatus ||
      state.recentAdDataError !== previousState.recentAdDataError ||
      state.recentAdDataMatchSummary !== previousState.recentAdDataMatchSummary ||
      state.adjustmentDrafts !== previousState.adjustmentDrafts;

    if (!shouldSave) {
      return;
    }

    if (saveTimer) {
      clearTimeout(saveTimer);
    }

    saveTimer = setTimeout(() => {
      useWorkspaceStore.setState({ persistenceStatus: "saving", persistenceError: undefined });
      writeWorkspaceSnapshot(takeWorkspaceSnapshot(useWorkspaceStore.getState()))
        .then(() => {
          useWorkspaceStore.setState({ persistenceStatus: "saved", persistenceError: undefined });
        })
        .catch((error) => {
          useWorkspaceStore.setState({
            persistenceStatus: "failed",
            persistenceError: error instanceof Error ? error.message : "自动保存失败。",
          });
        });
    }, 500);
  });
}
