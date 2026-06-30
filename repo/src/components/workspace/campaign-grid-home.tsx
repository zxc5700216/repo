"use client";

import { ArrowUpRight, Search, SlidersHorizontal, UploadCloud } from "lucide-react";
import { useEffect, useMemo, useState } from "react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { useBulkUpload } from "@/lib/hooks/use-bulk-upload";
import { useWorkspaceStore } from "@/lib/stores/workspace-store";
import { workspacePanelAnchorId } from "@/lib/workspace-events";

const pageSize = 16;
const dragMimeType = "application/x-campaign-group-id";

type PageItem = number | "...";

function buildPageItems(currentPage: number, totalPages: number) {
  if (totalPages <= 7) {
    return Array.from({ length: totalPages }, (_, index) => index + 1) as PageItem[];
  }

  if (currentPage <= 4) {
    return [1, 2, 3, 4, 5, "...", totalPages];
  }

  if (currentPage >= totalPages - 3) {
    return [1, "...", totalPages - 4, totalPages - 3, totalPages - 2, totalPages - 1, totalPages];
  }

  return [1, "...", currentPage - 1, currentPage, currentPage + 1, "...", totalPages];
}

export function CampaignGridHome() {
  const [query, setQuery] = useState("");
  const [page, setPage] = useState(1);
  const { fileInputRef, handleFileSelected } = useBulkUpload();
  const {
    campaignGroups,
    workspaceUnits,
    activeCampaignGroupId,
    activeLifecycleGroupId,
    openCampaignGroup,
    mergeCampaignGroupsIntoWorkspaceUnit,
    setActiveWorkspaceUnit,
  } = useWorkspaceStore();
  const groupedCampaignIds = useMemo(
    () => new Set(workspaceUnits.flatMap((unit) => unit.campaignGroupIds)),
    [workspaceUnits],
  );

  const filteredCampaigns = useMemo(
    () =>
      campaignGroups.filter((campaign) => {
        const grouped = groupedCampaignIds.has(campaign.id);
        const unitContainsCampaign = workspaceUnits.some((unit) => unit.campaignGroupIds[0] === campaign.id);
        const matchesQuery = `${campaign.sheetName ?? ""} ${campaign.campaignName} ${campaign.adGroupName}`
          .toLowerCase()
          .includes(query.toLowerCase());
        const matchesLifecycle = activeLifecycleGroupId ? campaign.lifecycleGroupId === activeLifecycleGroupId : true;

        return matchesQuery && matchesLifecycle && (!grouped || unitContainsCampaign);
      }),
    [activeLifecycleGroupId, campaignGroups, groupedCampaignIds, query, workspaceUnits],
  );
  const totalPages = Math.max(1, Math.ceil(filteredCampaigns.length / pageSize));
  const pagedCampaigns = filteredCampaigns.slice((page - 1) * pageSize, page * pageSize);
  const pageItems = buildPageItems(page, totalPages);

  useEffect(() => {
    setPage(1);
  }, [query, activeLifecycleGroupId]);

  useEffect(() => {
    if (page > totalPages) {
      setPage(totalPages);
    }
  }, [page, totalPages]);

  function openWorkspaceDetail(campaignGroupId: string, workspaceUnitId?: string) {
    if (workspaceUnitId) {
      setActiveWorkspaceUnit(workspaceUnitId);
    } else {
      openCampaignGroup(campaignGroupId);
    }

    requestAnimationFrame(() => {
      document.getElementById(workspacePanelAnchorId)?.scrollIntoView({
        behavior: "smooth",
        block: "start",
      });
    });
  }

  return (
    <section className="space-y-4">
      <div className="flex flex-wrap items-center gap-3 rounded-xl border border-border bg-white p-4">
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls,.xlsm,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv"
          className="hidden"
          aria-hidden
          onChange={(event) => void handleFileSelected(event.target.files?.[0])}
        />
        <Button onClick={() => fileInputRef.current?.click()}>
          <UploadCloud className="h-4 w-4" />
          Upload Bulk
        </Button>
        <div className="flex h-10 min-w-[240px] flex-1 items-center gap-2 rounded-lg border border-border px-3">
          <Search className="h-4 w-4 text-muted" />
          <input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="Search campaign, ad group, or sheet"
            className="min-w-0 flex-1 bg-transparent text-sm outline-none"
          />
        </div>
        <div className="inline-flex items-center gap-2 rounded-lg border border-border bg-surface-muted px-3 py-2 text-xs font-semibold text-muted">
          <SlidersHorizontal className="h-4 w-4 text-brand" />
          Lifecycle Filter {activeLifecycleGroupId ? `· ${activeLifecycleGroupId}` : "· All"}
        </div>
        <div className="inline-flex items-center gap-2 rounded-lg border border-border bg-surface-muted px-3 py-2 text-xs font-semibold text-muted">
          Sort · Last Updated
        </div>
        <div className="inline-flex items-center gap-2 rounded-lg border border-border bg-surface-muted px-3 py-2 text-xs font-semibold text-muted">
          View · Campaign Grid
        </div>
      </div>

      <div className="flex items-center justify-between gap-3">
        <div className="text-sm font-semibold text-muted">
          {filteredCampaigns.length} campaigns · Page {page} / {totalPages}
        </div>
        <div className="flex flex-wrap items-center gap-2">
          {pageItems.map((item, index) =>
            item === "..." ? (
              <span key={`ellipsis-${index}`} className="px-1 text-sm font-semibold text-muted">
                ...
              </span>
            ) : (
              <button
                key={item}
                type="button"
                onClick={() => {
                  if (typeof item === "number") {
                    setPage(item);
                  }
                }}
                className={`h-9 min-w-9 rounded-md px-3 text-sm font-semibold transition-colors ${
                  page === item ? "bg-brand text-white" : "border border-border bg-white text-muted hover:text-foreground"
                }`}
              >
                {item}
              </button>
            ),
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-4">
        {pagedCampaigns.map((campaign) => {
          const active = campaign.id === activeCampaignGroupId;
          const workspaceUnit = workspaceUnits.find((unit) => unit.campaignGroupIds[0] === campaign.id);
          const tileCampaigns = workspaceUnit
            ? campaignGroups.filter((group) => workspaceUnit.campaignGroupIds.includes(group.id)).slice(0, 4)
            : [campaign];
          const keywordCount = tileCampaigns.reduce((sum, group) => sum + group.keywordCount, 0);
          const splitLayout = tileCampaigns.length >= 4 ? "grid-cols-2" : tileCampaigns.length >= 2 ? "grid-cols-2" : "grid-cols-1";

          return (
            <button
              key={campaign.id}
              draggable
              onDragStart={(event) => {
                event.dataTransfer.setData(dragMimeType, campaign.id);
                event.dataTransfer.effectAllowed = "move";
              }}
              onDragOver={(event) => {
                event.preventDefault();
                event.dataTransfer.dropEffect = "move";
              }}
              onDrop={(event) => {
                event.preventDefault();
                const sourceCampaignGroupId = event.dataTransfer.getData(dragMimeType);

                if (sourceCampaignGroupId) {
                  mergeCampaignGroupsIntoWorkspaceUnit(sourceCampaignGroupId, campaign.id);
                }
              }}
              onClick={() => {
                if (workspaceUnit) {
                  openWorkspaceDetail(campaign.id, workspaceUnit.id);
                } else {
                  openWorkspaceDetail(campaign.id);
                }
              }}
              className={`rounded-2xl border p-5 text-left transition-all ${
                active ? "border-brand bg-white shadow-[0_12px_30px_rgba(23,107,135,0.12)]" : "border-border bg-white hover:-translate-y-0.5"
              }`}
            >
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <p className="truncate text-sm font-black uppercase tracking-[0.12em] text-muted">
                    {campaign.sheetName ?? "Bulk"}
                  </p>
                  <h3 className="mt-2 line-clamp-2 text-lg font-bold text-foreground">
                    {workspaceUnit ? workspaceUnit.name : campaign.adGroupName}
                  </h3>
                  <p className="mt-1 truncate text-sm text-muted">{campaign.campaignName}</p>
                </div>
                <ArrowUpRight className={`h-5 w-5 shrink-0 ${active ? "text-brand" : "text-muted"}`} />
              </div>

              <div className="mt-4 flex items-center gap-2">
                <Badge tone={campaign.lifecycleGroupId === "mature" ? "green" : campaign.lifecycleGroupId === "decline" ? "amber" : campaign.lifecycleGroupId === "clearance" ? "red" : "blue"}>
                  {campaign.lifecycleGroupId ?? "unassigned"}
                </Badge>
                <span className="text-xs font-semibold text-muted">Updated {campaign.lastUpdated}</span>
              </div>

              <div className={`mt-5 grid gap-3 rounded-xl bg-surface-muted p-3 ${splitLayout}`}>
                {tileCampaigns.map((tileCampaign) => (
                  <div
                    key={tileCampaign.id}
                    draggable
                    onDragStart={(event) => {
                      event.stopPropagation();
                      event.dataTransfer.setData(dragMimeType, tileCampaign.id);
                      event.dataTransfer.effectAllowed = "move";
                    }}
                    className="rounded-lg bg-white/80 p-3"
                  >
                    <p className="truncate text-xs font-bold text-muted">{tileCampaign.adGroupName}</p>
                    <p className="mt-2 text-lg font-black text-foreground">{tileCampaign.keywordCount.toLocaleString("zh-CN")}</p>
                  </div>
                ))}
              </div>

              <div className="mt-3 grid grid-cols-2 gap-3 rounded-xl bg-surface-muted p-3">
                <div>
                  <p className="text-[11px] font-bold uppercase tracking-[0.08em] text-muted">Keywords</p>
                  <p className="mt-1 text-xl font-black text-foreground">{keywordCount.toLocaleString("zh-CN")}</p>
                </div>
                <div>
                  <p className="text-[11px] font-bold uppercase tracking-[0.08em] text-muted">Workspace</p>
                  <p className="mt-1 text-sm font-bold text-foreground">{workspaceUnit ? `${tileCampaigns.length} Up` : "Open"}</p>
                </div>
              </div>
            </button>
          );
        })}
      </div>

      {pagedCampaigns.length === 0 && (
        <div className="rounded-xl border border-dashed border-border bg-white px-5 py-10 text-center text-sm font-medium text-muted">
          当前筛选条件下没有可展示的广告组。
        </div>
      )}
    </section>
  );
}
