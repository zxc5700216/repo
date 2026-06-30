import { enrichMetric } from "@/lib/metrics";
import type {
  AdjustmentDraft,
  Condition,
  ConditionGroup,
  OptimizationEngine,
  OptimizationInput,
  PerformanceRow,
  RecentAdDataRow,
  Rule,
  RuleAction,
} from "@/lib/types";

type RuleEvaluationContext = {
  bulkRow: PerformanceRow;
  recentRow?: RecentAdDataRow;
  campaignRecentRows: RecentAdDataRow[];
  orderShare?: number;
  isCoreKeyword: boolean;
};

function isConditionGroup(value: Condition | ConditionGroup): value is ConditionGroup {
  return "logic" in value;
}

function normalizeMatchValue(value: string) {
  return value.trim().toLowerCase().replace(/\s+/g, " ");
}

function buildKeywordMatchKey(keyword: string, matchType: string) {
  return `${normalizeMatchValue(keyword)}::${normalizeMatchValue(matchType)}`;
}

function getRecentMetric(row: RecentAdDataRow | undefined, metric: string): number | undefined {
  if (!row) {
    return undefined;
  }

  switch (metric) {
    case "acos":
      return row.acos ?? (row.sales > 0 ? (row.spend / row.sales) * 100 : undefined);
    case "roas":
      return row.roas ?? (row.spend > 0 ? row.sales / row.spend : undefined);
    case "ctr":
      return row.impressions > 0 ? (row.clicks / row.impressions) * 100 : undefined;
    case "cpc":
      return row.clicks > 0 ? row.spend / row.clicks : undefined;
    case "cvr":
      return row.clicks > 0 ? (row.orders / row.clicks) * 100 : undefined;
    case "cpa":
      return row.orders > 0 ? row.spend / row.orders : undefined;
    default:
      return Number(row[metric as keyof RecentAdDataRow] ?? 0);
  }
}

function findRecentRow(row: PerformanceRow, recentRows: RecentAdDataRow[]) {
  const matchedRows = recentRows.filter(
    (recentRow) => recentRow.matchStatus === "matched" && recentRow.campaignGroupId === row.campaignGroupId,
  );
  const rowMatchKey = buildKeywordMatchKey(row.keyword, row.matchType);

  return (
    matchedRows.find((recentRow) => buildKeywordMatchKey(recentRow.keyword, recentRow.matchType) === rowMatchKey) ??
    matchedRows[0]
  );
}

function calcOrderShare(recentRow: RecentAdDataRow | undefined, campaignRecentRows: RecentAdDataRow[]) {
  if (!recentRow) {
    return undefined;
  }

  const totalOrders = campaignRecentRows.reduce((sum, row) => sum + row.orders, 0);

  return totalOrders > 0 ? (recentRow.orders / totalOrders) * 100 : undefined;
}

function isCoreKeywordCandidate(row: PerformanceRow) {
  const normalizedCampaignName = row.campaignName.trim().toLowerCase();
  const normalizedAdGroupName = row.adGroupName.trim().toLowerCase();
  const normalizedMatchType = row.matchType.trim().toLowerCase();
  const normalizedTarget = row.target.trim().toLowerCase();

  return (
    normalizedCampaignName.includes("core") ||
    normalizedCampaignName.includes("核心") ||
    normalizedAdGroupName.includes("core") ||
    normalizedAdGroupName.includes("核心") ||
    (normalizedMatchType === "exact" && !normalizedTarget.startsWith("asin="))
  );
}

function getConditionActualValue(context: RuleEvaluationContext, condition: Condition): number | undefined {
  const dataSource = condition.dataSource ?? "bulk";

  if (dataSource === "recent") {
    return getRecentMetric(context.recentRow, condition.metric);
  }

  if (dataSource === "derived") {
    if (condition.metric === "orderShare") {
      return context.orderShare;
    }

    if (condition.metric === "isCoreKeyword") {
      return context.isCoreKeyword ? 1 : 0;
    }

    return undefined;
  }

  if (dataSource === "comparison") {
    const bulkValue = enrichMetric(context.bulkRow, condition.compareMetric ?? condition.metric);
    const recentValue = getRecentMetric(context.recentRow, condition.metric);

    if (bulkValue === 0 || bulkValue === undefined || recentValue === undefined) {
      return undefined;
    }

    return ((recentValue - bulkValue) / Math.abs(bulkValue)) * 100;
  }

  return enrichMetric(context.bulkRow, condition.metric);
}

function evaluateCondition(context: RuleEvaluationContext, condition: Condition): boolean {
  const actual = getConditionActualValue(context, condition);

  if (actual === undefined || Number.isNaN(actual)) {
    return false;
  }

  switch (condition.operator) {
    case "eq":
      return actual === condition.value;
    case "neq":
      return actual !== condition.value;
    case "gt":
      return actual > Number(condition.value);
    case "gte":
      return actual >= Number(condition.value);
    case "lt":
      return actual < Number(condition.value);
    case "lte":
      return actual <= Number(condition.value);
    case "between":
      return actual >= Number(condition.min) && actual <= Number(condition.max);
    case "increase_by":
      return actual >= Number(condition.value);
    case "decrease_by":
      return actual <= -Math.abs(Number(condition.value));
    default:
      return false;
  }
}

export function evaluateConditionGroup(context: RuleEvaluationContext, group: ConditionGroup): boolean {
  const results = group.conditions.map((item) =>
    isConditionGroup(item) ? evaluateConditionGroup(context, item) : evaluateCondition(context, item),
  );

  return group.logic === "AND" ? results.every(Boolean) : results.some(Boolean);
}

function applyBidAction(currentBid: number, action: RuleAction): number {
  const value = Number(action.value ?? 0);

  switch (action.type) {
    case "increase_bid_percent":
      return currentBid * (1 + value / 100);
    case "decrease_bid_percent":
      return currentBid * (1 - value / 100);
    case "increase_bid_fixed":
      return currentBid + value;
    case "decrease_bid_fixed":
      return currentBid - value;
    case "set_bid":
      return value;
    default:
      return currentBid;
  }
}

function buildReason(rule: Rule, action: RuleAction): string {
  const actionText: Record<string, string> = {
    increase_bid_percent: `竞价提高 ${action.value}%`,
    decrease_bid_percent: `竞价降低 ${action.value}%`,
    increase_bid_fixed: `竞价增加 $${action.value}`,
    decrease_bid_fixed: `竞价减少 $${action.value}`,
    set_bid: `设置固定竞价 $${action.value}`,
    pause_keyword: "暂停关键词",
    enable_keyword: "启用关键词",
    add_negative_keyword: "添加否定关键词",
    add_label: `添加标签 ${action.label ?? ""}`,
    mark_pending: "标记待处理",
    no_change: "保持不变",
  };

  return `${rule.name}: ${actionText[action.type]}`;
}

function createBidDraft(row: PerformanceRow, rule: Rule, action: RuleAction): AdjustmentDraft {
  const rawBid = applyBidAction(row.currentBid, action);
  const suggestedBid = Math.max(0.02, Number(rawBid.toFixed(2)));
  const deltaPercent =
    row.currentBid > 0 ? Number((((suggestedBid - row.currentBid) / row.currentBid) * 100).toFixed(1)) : 0;

  return {
    id: `${row.id}-${rule.id}-${action.id}`,
    batchId: row.batchId,
    sheetName: row.sheetName,
    sourceRowIndex: row.sourceRowIndex,
    sourceRowNumber: row.sourceRowNumber,
    campaignGroupId: row.campaignGroupId,
    rowId: row.id,
    field: "bid",
    headerName: "竞价",
    oldValue: row.currentBid,
    newValue: suggestedBid,
    keyword: row.keyword,
    target: row.target,
    currentBid: row.currentBid,
    suggestedBid,
    deltaPercent,
    reason: buildReason(rule, action),
    matchedRule: rule.name,
    selected: true,
  };
}

function isBidAction(action: RuleAction) {
  return [
    "increase_bid_percent",
    "decrease_bid_percent",
    "increase_bid_fixed",
    "decrease_bid_fixed",
    "set_bid",
  ].includes(action.type);
}

function shouldBlockFurtherRules(actions: RuleAction[]) {
  return actions.some((action) => action.type === "no_change");
}

export function runRuleEngine(input: OptimizationInput): AdjustmentDraft[] {
  const scopedRows = input.rows.filter((row) => row.campaignGroupId === input.campaignGroup.id);
  const recentAdDataRows = input.recentAdDataRows ?? [];
  const rules = input.rules
    .filter((rule) => rule.enabled && rule.lifecycleGroupId === input.campaignGroup.lifecycleGroupId)
    .sort((a, b) => a.priority - b.priority);

  const drafts: AdjustmentDraft[] = [];
  const touchedRows = new Set<string>();

  for (const rule of rules) {
    for (const row of scopedRows) {
      if (touchedRows.has(row.id)) {
        continue;
      }

      const recentRow = findRecentRow(row, recentAdDataRows);
      const campaignRecentRows = recentAdDataRows.filter(
        (item) => item.matchStatus === "matched" && item.campaignGroupId === row.campaignGroupId,
      );
      const context = {
        bulkRow: row,
        recentRow,
        campaignRecentRows,
        orderShare: calcOrderShare(recentRow, campaignRecentRows),
        isCoreKeyword: isCoreKeywordCandidate(row),
      };

      if (evaluateConditionGroup(context, rule.conditionGroup)) {
        const action = rule.actions.find(isBidAction);

        if (action) {
          drafts.push(createBidDraft(row, rule, action));
          touchedRows.add(row.id);
          continue;
        }

        if (shouldBlockFurtherRules(rule.actions)) {
          touchedRows.add(row.id);
        }
      }
    }
  }

  return drafts;
}

export const ruleOptimizationEngine: OptimizationEngine = {
  id: "rule-engine-v1",
  name: "规则优化引擎",
  run: runRuleEngine,
};
