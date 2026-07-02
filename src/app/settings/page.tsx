import { Bot, Database, FileOutput, ShieldCheck } from "lucide-react";
import { AppShell } from "@/components/app-shell/app-shell";
import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

const settings = [
  {
    icon: ShieldCheck,
    title: "数据安全策略",
    text: "规则演算默认仅存在工作区，未点击提交禁止修改原始数据。",
    badge: "强制开启",
  },
  {
    icon: Database,
    title: "广告组隔离",
    text: "Selectors、Rule Runs、Export Jobs 全部绑定 campaignGroupId。",
    badge: "最高优先级",
  },
  {
    icon: Bot,
    title: "AI Bid Optimization Engine",
    text: "预留统一 OptimizationEngine 接口，AI 输出同样进入 AdjustmentDraft。",
    badge: "预留",
  },
  {
    icon: FileOutput,
    title: "导出字段映射",
    text: "复制原始 Workbook 后应用勾选 patch，生成已修改广告数据.xlsx。",
    badge: "安全写回",
  },
];

export default function SettingsPage() {
  return (
    <AppShell title="Settings" subtitle="安全边界、导出映射与 AI 引擎预留配置">
      <div className="grid grid-cols-1 gap-5 xl:grid-cols-2">
        {settings.map((item) => {
          const Icon = item.icon;

          return (
            <Card key={item.title}>
              <CardHeader className="flex flex-row items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-surface-muted text-brand">
                    <Icon className="h-5 w-5" />
                  </div>
                  <CardTitle>{item.title}</CardTitle>
                </div>
                <Badge tone="blue">{item.badge}</Badge>
              </CardHeader>
              <CardContent>
                <p className="text-sm leading-6 text-muted">{item.text}</p>
              </CardContent>
            </Card>
          );
        })}
      </div>
    </AppShell>
  );
}
