import { AdvancedFilterPanel } from "@/components/dashboard/advanced-filter-panel";
import { TrendChart } from "@/components/charts/trend-chart";
import { AppShell } from "@/components/app-shell/app-shell";
import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { campaignGroups, metricSnapshots } from "@/data/mock-data";
import { formatCurrency, formatNumber, formatPercent } from "@/lib/metrics";

export default function DashboardPage() {
  const totals = metricSnapshots.reduce(
    (acc, item) => {
      acc.sales += item.sales;
      acc.spend += item.spend;
      acc.orders += item.orders;
      acc.clicks += item.clicks;
      acc.impressions += item.impressions;
      return acc;
    },
    { sales: 0, spend: 0, orders: 0, clicks: 0, impressions: 0 },
  );
  const acos = totals.sales > 0 ? (totals.spend / totals.sales) * 100 : 0;

  return (
    <AppShell title="Dashboard" subtitle="KPI summary, campaign prioritization, and action-oriented analytics">
      <div className="space-y-5">
        <AdvancedFilterPanel />
        <div className="grid grid-cols-2 gap-4 xl:grid-cols-4">
          {[
            ["总销售额", formatCurrency(totals.sales), "+14.8%"],
            ["总花费", formatCurrency(totals.spend), "-2.6%"],
            ["ACOS", formatPercent(acos), "-4.3%"],
            ["订单量", formatNumber(totals.orders), "+126"],
          ].map(([label, value, delta]) => (
            <Card key={label}>
              <CardContent className="p-5">
                <p className="text-xs font-bold text-muted">{label}</p>
                <div className="mt-4 flex items-end justify-between">
                  <p className="metric-tabular text-3xl font-black text-foreground">{value}</p>
                  <Badge tone="green">{delta}</Badge>
                </div>
              </CardContent>
            </Card>
          ))}
        </div>
        <div className="grid grid-cols-1 gap-5 xl:grid-cols-[1fr_420px]">
          <Card>
            <CardHeader>
              <CardTitle>Sales / Spend 趋势</CardTitle>
            </CardHeader>
            <CardContent>
              <TrendChart />
            </CardContent>
          </Card>
          <Card>
            <CardHeader>
              <CardTitle>Opportunity Queue</CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              {campaignGroups.slice(0, 5).map((group, index) => (
                <div key={group.id} className="flex items-center justify-between rounded-md border border-border p-3">
                  <div className="min-w-0">
                    <p className="truncate text-sm font-bold text-foreground">{group.adGroupName}</p>
                    <p className="text-xs font-medium text-muted">{group.keywordCount} keywords pending prioritization</p>
                  </div>
                  <Badge tone={index % 2 === 0 ? "amber" : "blue"}>{index % 2 === 0 ? "高 ACOS" : "低曝光"}</Badge>
                </div>
              ))}
            </CardContent>
          </Card>
        </div>
      </div>
    </AppShell>
  );
}
