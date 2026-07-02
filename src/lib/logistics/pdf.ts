import type { PdfPageSummary, PdfSummary } from "@/lib/logistics/types";
import { inferPdfMetaFromFileName } from "@/lib/logistics/utils";

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function decodePdfString(buffer: ArrayBuffer) {
  const decoder = new TextDecoder("latin1");
  return decoder.decode(buffer);
}

function matchFirst(text: string, pattern: RegExp) {
  const matched = text.match(pattern);
  return matched?.[1] ?? "";
}

function normalizePdfText(text: string) {
  return text.replace(/\0/g, "").replace(/\s+/g, " ").trim();
}

function getPageCount(pdfText: string) {
  const matched = pdfText.match(/\/Count\s+(\d+)/);
  const parsed = matched ? Number(matched[1]) : null;
  return parsed && Number.isFinite(parsed) ? parsed : 1;
}

function extractShipmentTitle(pdfText: string, fallbackName: string) {
  const normalizedText = normalizePdfText(pdfText);
  const titlePatterns = [
    /(货件\d+-[A-Z]{3,5}\d*-[A-Z0-9]+-\d+箱-[^ ]+?)\s+Created:/u,
    /(货件\d+-[A-Z]{3,5}\d*-[A-Z0-9]+-\d+箱-[^\r\n]+?)\s+Created:/u,
    /(货件\d+-[A-Z]{3,5}\d*-[A-Z0-9]+-\d+箱-[^\r\n]+)/u,
  ];

  for (const pattern of titlePatterns) {
    const matched = normalizedText.match(pattern)?.[1]?.trim();
    if (matched) {
      return matched;
    }
  }

  return inferPdfMetaFromFileName(fallbackName).shipmentTitle;
}

function extractGlobalMeta(pdfText: string, fallbackName: string) {
  const titleText = extractShipmentTitle(pdfText, fallbackName);
  const titleParts = titleText ? titleText.split("-") : [];

  const shipmentName = titleParts[0] || matchFirst(pdfText, /(货件\d+)/u);

  const fbaBoxCodes = Array.from(pdfText.matchAll(/(FBA[A-Z0-9]+U\d{6})/g)).map((item) => item[1]);
  const firstFbaBoxCode = fbaBoxCodes[0] ?? "";
  const fbaCode = titleParts[2] || (firstFbaBoxCode ? firstFbaBoxCode.replace(/U\d{6}$/u, "") : matchFirst(pdfText, /(FBA[A-Z0-9]+)/u));

  const warehouseCandidates = new Set<string>();
  for (const match of pdfText.matchAll(/\b([A-Z]{3}\d{1,2})\b/g)) {
    const value = match[1];
    if (value !== "FBA" && value !== "PDT" && value !== "SKU") {
      warehouseCandidates.add(value);
    }
  }
  const warehouseCode = titleParts[1] || (Array.from(warehouseCandidates).find((value) => value !== fbaCode) ?? "");
  const titleBoxCount = titleParts[3] ? Number(titleParts[3].replace(/[^\d]/g, "")) : null;
  const channelNameFromTitle = titleParts.slice(4).join("-").trim();

  const fallbackMeta = inferPdfMetaFromFileName(fallbackName);

  return {
    shipmentTitle: titleText || fallbackMeta.shipmentTitle,
    shipmentName: shipmentName || fallbackMeta.shipmentName,
    warehouseCode: warehouseCode || fallbackMeta.warehouseCode,
    fbaCode: fbaCode || fallbackMeta.fbaCode,
    totalBoxes: titleBoxCount,
    channelName: channelNameFromTitle || fallbackMeta.channelName,
  };
}

function extractPageSummaries(
  pdfText: string,
  sharedMeta: { shipmentName: string; warehouseCode: string; fbaCode: string },
  pageCount: number,
) {
  const pages: PdfPageSummary[] = [];
  const boxMatches = Array.from(pdfText.matchAll(/(FBA[A-Z0-9]+U\d{6})/g)).map((item) => item[1]);
  const positionMatches = Array.from(pdfText.matchAll(/(P\d+\s*-\s*B\d+)/g)).map((item) => item[1]);
  const skuTypeMatches = Array.from(pdfText.matchAll(/(Single SKU|Mixed SKUs)/g)).map((item) => item[1]);
  const qtyMatches = Array.from(pdfText.matchAll(/Qty\s*(\d+)/g)).map((item) => Number(item[1]));

  for (let index = 0; index < pageCount; index += 1) {
    const fbaBoxCode = boxMatches[index] ?? (sharedMeta.fbaCode ? `${sharedMeta.fbaCode}U${String(index + 1).padStart(6, "0")}` : "");

    const skuType = skuTypeMatches[index] ?? "";
    const qty = qtyMatches[index] ?? null;
    const positionCode = positionMatches[index] ?? "";

    let sku = "";
    if (skuType === "Single SKU") {
      const localText = pdfText.slice(Math.max(0, (pdfText.indexOf(fbaBoxCode) || 0) - 400), Math.min(pdfText.length, (pdfText.indexOf(fbaBoxCode) || 0) + 800));
      const skuPattern = /Single SKU[\s\S]{0,120}?([A-Z0-9][A-Z0-9\- ]{2,})[\r\n]/;
      sku = matchFirst(localText, skuPattern).trim();
    }

    pages.push({
      pageNumber: index + 1,
      shipmentName: sharedMeta.shipmentName,
      warehouseCode: sharedMeta.warehouseCode,
      fbaCode: sharedMeta.fbaCode,
      totalBoxes: pageCount,
      fbaBoxCode,
      positionCode,
      skuType,
      sku,
      qty,
    });
  }

  return pages;
}

export async function parsePdfFile(file: File): Promise<PdfSummary> {
  const buffer = await file.arrayBuffer();
  const pdfText = decodePdfString(buffer);
  const parsedPageCount = getPageCount(pdfText);
  const sharedMeta = extractGlobalMeta(pdfText, file.name);
  const pageCount = sharedMeta.totalBoxes && Number.isFinite(sharedMeta.totalBoxes) ? sharedMeta.totalBoxes : parsedPageCount;
  const pages = extractPageSummaries(pdfText, sharedMeta, pageCount);
  const renamedFileName = `${sharedMeta.shipmentTitle || `${sharedMeta.shipmentName}-${sharedMeta.warehouseCode}-${sharedMeta.fbaCode}-${pageCount}箱`}.pdf`
    .replace(/\.pdf\.pdf$/i, ".pdf")
    .replace(/^-+|-+$/g, "");

  return {
    fileNameBase: file.name.replace(/\.pdf$/i, ""),
    shipmentTitle:
      sharedMeta.shipmentTitle ||
      [sharedMeta.shipmentName, sharedMeta.warehouseCode, sharedMeta.fbaCode, `${pageCount}箱`, sharedMeta.channelName].filter(Boolean).join("-"),
    shipmentName: sharedMeta.shipmentName,
    warehouseCode: sharedMeta.warehouseCode,
    fbaCode: sharedMeta.fbaCode,
    totalBoxes: pageCount,
    channelName: sharedMeta.channelName,
    renamedFileName: renamedFileName === ".pdf" ? file.name : renamedFileName,
    pages,
  };
}

export function isPdfSummaryContentBased(summary: PdfSummary) {
  const expected = `${summary.shipmentTitle || `${summary.shipmentName}-${summary.warehouseCode}-${summary.fbaCode}-${summary.totalBoxes ?? 0}箱`}.pdf`;
  return new RegExp(`^${escapeRegExp(expected)}$`, "u").test(summary.renamedFileName);
}
