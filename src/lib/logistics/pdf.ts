import type { PdfPageSummary, PdfSummary } from "@/lib/logistics/types";
import { inferPdfMetaFromFileName } from "@/lib/logistics/utils";
import { getDocument } from "pdfjs-dist";
import { inflate } from "pako";
import { recognize } from "tesseract.js";

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function decodePdfString(buffer: ArrayBuffer) {
  const decoder = new TextDecoder("latin1");
  return decoder.decode(buffer);
}

function decodeUtf16BeWithFallback(value: string) {
  const bytes = new Uint8Array(value.length);
  for (let index = 0; index < value.length; index += 1) {
    bytes[index] = value.charCodeAt(index) & 0xff;
  }

  return new TextDecoder("utf-16be", { fatal: false }).decode(bytes);
}

function extractPdfStreamCandidates(buffer: ArrayBuffer) {
  const pdfBytes = new Uint8Array(buffer);
  const pdfText = decodePdfString(buffer);
  const matches = pdfText.matchAll(/stream\r?\n([\s\S]*?)\r?\nendstream/g);
  const candidates: string[] = [];

  for (const match of matches) {
    const streamBody = match[1];
    if (!streamBody) {
      continue;
    }

    const streamStart = match.index ? match.index + match[0].indexOf(streamBody) : -1;
    if (streamStart < 0) {
      continue;
    }

    const streamBytes = pdfBytes.slice(streamStart, streamStart + streamBody.length);

    let inflatedText = "";
    try {
      inflatedText = new TextDecoder("latin1").decode(inflate(streamBytes));
    } catch {
      continue;
    }

    if (!inflatedText.includes("Tj") && !inflatedText.includes("TJ")) {
      continue;
    }

    const literalMatches = Array.from(inflatedText.matchAll(/\((.*?)\)Tj/g)).map((item) => item[1]);
    if (!literalMatches.length) {
      continue;
    }

    literalMatches.forEach((item) => {
      candidates.push(item);
      candidates.push(decodeUtf16BeWithFallback(item));
    });
  }

  return candidates.filter(Boolean);
}

function findShipmentTitleInText(text: string) {
  const normalized = normalizePdfText(text).replace(/[—–]/g, "-");
  const patterns = [
    /货件\d+-[A-Z]{3,5}\d*-FBA[A-Z0-9]+-\d+箱-[\u4e00-\u9fffA-Za-z0-9]+/u,
    /货件\d+-[A-Z]{3,5}\d*-FBA[A-Z0-9]+-\d+箱-[^\s()]+/u,
  ];

  for (const pattern of patterns) {
    const matched = normalized.match(pattern)?.[0];
    if (matched) {
      return matched;
    }
  }

  return "";
}

function extractShipmentTitleFromStream(buffer: ArrayBuffer) {
  const candidates = extractPdfStreamCandidates(buffer);
  for (const candidate of candidates) {
    const matched = findShipmentTitleInText(candidate);
    if (matched) {
      return matched;
    }
  }

  return "";
}

async function extractPdfText(buffer: ArrayBuffer) {
  const fallbackText = decodePdfString(buffer);

  try {
    const loadingTask = getDocument({ data: buffer });
    const pdf = await loadingTask.promise;
    const pageTexts: string[] = [];

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      const text = content.items
        .map((item) => ("str" in item ? item.str : ""))
        .filter(Boolean)
        .join(" ");
      pageTexts.push(text);
    }

    await loadingTask.destroy();
    const extractedText = pageTexts.join("\n");
    return [extractedText, fallbackText].filter((item) => item.trim()).join("\n");
  } catch {
    return fallbackText;
  }
}

async function extractShipmentNameViaOcr(buffer: ArrayBuffer) {
  const title = await extractShipmentTitleViaOcr(buffer);
  const parsed = title ? parseShipmentTitle(title) : null;
  return parsed?.shipmentName ?? "";
}

async function renderFirstPageCanvas(buffer: ArrayBuffer, scale = 3) {
  if (typeof window === "undefined") {
    return null;
  }

  try {
    const loadingTask = getDocument({ data: buffer });
    const pdf = await loadingTask.promise;
    const page = await pdf.getPage(1);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");

    if (!context) {
      await loadingTask.destroy();
      return null;
    }

    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);

    await page.render({
      canvasContext: context,
      viewport,
    }).promise;
    await loadingTask.destroy();

    return canvas;
  } catch {
    return null;
  }
}

function hasSuspiciousShipmentText(value: string) {
  return /[�&〃廄拈]/u.test(value);
}

async function extractShipmentTitleViaOcr(buffer: ArrayBuffer) {
  const canvas = await renderFirstPageCanvas(buffer, 4);
  if (!canvas) {
    return "";
  }

  const cropCanvas = document.createElement("canvas");
  const cropContext = cropCanvas.getContext("2d");
  if (!cropContext) {
    return "";
  }

  const sourceX = Math.floor(canvas.width * 0.02);
  const sourceY = Math.floor(canvas.height * 0.26);
  const sourceWidth = Math.floor(canvas.width * 0.6);
  const sourceHeight = Math.floor(canvas.height * 0.11);

  cropCanvas.width = sourceWidth * 2;
  cropCanvas.height = sourceHeight * 2;
  cropContext.drawImage(
    canvas,
    sourceX,
    sourceY,
    sourceWidth,
    sourceHeight,
    0,
    0,
    cropCanvas.width,
    cropCanvas.height,
  );

  const result = await recognize(cropCanvas, "chi_sim+eng");
  const normalized = normalizePdfText(result.data.text).replace(/[—–]/g, "-");
  const matched = normalized.match(/货件\d+-[A-Z]{3,5}\d*-FBA[A-Z0-9]+-\d+箱-[\u4e00-\u9fffA-Za-z0-9]+/u);
  return matched?.[0] ?? "";
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

function extractShipmentNameToken(text: string) {
  const patterns = [/(货件\s*\d+)/u, /(货\s*件\s*\d+)/u];

  for (const pattern of patterns) {
    const matched = text.match(pattern)?.[1];
    if (matched) {
      return matched.replace(/\s+/g, "");
    }
  }

  return "";
}

function extractShipmentTitle(pdfText: string, fallbackName: string) {
  const matched = findShipmentTitleInText(pdfText);
  if (matched) {
    return matched;
  }

  return inferPdfMetaFromFileName(fallbackName).shipmentTitle;
}

function parseShipmentTitle(titleText: string) {
  const normalizedTitle = titleText.trim().replace(/\.pdf$/i, "");
  const matched = normalizedTitle.match(
    /^(?:(货件\d+)-)?([A-Z]{3,5}\d*)-(FBA[A-Z0-9]+)-(\d+)箱(?:-(.+))?$/u,
  );

  if (!matched) {
    return null;
  }

  return {
    shipmentName: matched[1] ?? "",
    warehouseCode: matched[2] ?? "",
    fbaCode: matched[3] ?? "",
    totalBoxes: matched[4] ? Number(matched[4]) : null,
    channelName: matched[5]?.trim() ?? "",
  };
}

function extractGlobalMeta(pdfText: string, fallbackName: string) {
  const normalizedText = normalizePdfText(pdfText);
  const titleText = extractShipmentTitle(pdfText, fallbackName);
  const parsedTitle = parseShipmentTitle(titleText);

  const shipmentName =
    parsedTitle?.shipmentName ||
    extractShipmentNameToken(normalizedText) ||
    matchFirst(pdfText, /(货件\d+)/u);

  const fbaBoxCodes = Array.from(pdfText.matchAll(/(FBA[A-Z0-9]+U\d{6})/g)).map((item) => item[1]);
  const firstFbaBoxCode = fbaBoxCodes[0] ?? "";
  const fbaCode =
    parsedTitle?.fbaCode ||
    (firstFbaBoxCode ? firstFbaBoxCode.replace(/U\d{6}$/u, "") : matchFirst(pdfText, /(FBA[A-Z0-9]+)/u));

  const warehouseCandidates = new Set<string>();
  for (const match of pdfText.matchAll(/\b([A-Z]{3}\d{1,2})\b/g)) {
    const value = match[1];
    if (value !== "FBA" && value !== "PDT" && value !== "SKU") {
      warehouseCandidates.add(value);
    }
  }
  const warehouseCode = parsedTitle?.warehouseCode || (Array.from(warehouseCandidates).find((value) => value !== fbaCode) ?? "");
  const titleBoxCount = parsedTitle?.totalBoxes ?? null;
  const channelNameFromTitle = parsedTitle?.channelName ?? "";

  const fallbackMeta = inferPdfMetaFromFileName(fallbackName);
  const resolvedShipmentName = shipmentName || fallbackMeta.shipmentName;
  const resolvedWarehouseCode = warehouseCode || fallbackMeta.warehouseCode;
  const resolvedFbaCode = fbaCode || fallbackMeta.fbaCode;
  const resolvedTotalBoxes = titleBoxCount ?? fallbackMeta.totalBoxes;
  const resolvedChannelName = channelNameFromTitle || fallbackMeta.channelName;
  const resolvedShipmentTitle = [
    resolvedShipmentName,
    resolvedWarehouseCode,
    resolvedFbaCode,
    resolvedTotalBoxes ? `${resolvedTotalBoxes}箱` : "",
    resolvedChannelName,
  ]
    .filter(Boolean)
    .join("-");

  return {
    shipmentTitle: resolvedShipmentTitle || titleText || fallbackMeta.shipmentTitle,
    shipmentName: resolvedShipmentName,
    warehouseCode: resolvedWarehouseCode,
    fbaCode: resolvedFbaCode,
    totalBoxes: resolvedTotalBoxes,
    channelName: resolvedChannelName,
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
  const streamShipmentTitle = extractShipmentTitleFromStream(buffer);
  const pdfText = await extractPdfText(buffer);
  const parsedPageCount = getPageCount(pdfText);
  let sharedMeta = extractGlobalMeta(streamShipmentTitle ? `${streamShipmentTitle}\n${pdfText}` : pdfText, file.name);

  if ((!sharedMeta.shipmentName || hasSuspiciousShipmentText(sharedMeta.shipmentTitle) || hasSuspiciousShipmentText(sharedMeta.channelName)) && !streamShipmentTitle) {
    const shipmentTitleFromOcr = await extractShipmentTitleViaOcr(buffer);
    if (shipmentTitleFromOcr) {
      sharedMeta = extractGlobalMeta(`${shipmentTitleFromOcr}\n${pdfText}`, file.name);
    } else if (!sharedMeta.shipmentName) {
      const shipmentNameFromOcr = await extractShipmentNameViaOcr(buffer);
      if (shipmentNameFromOcr) {
        sharedMeta = extractGlobalMeta(`${shipmentNameFromOcr}\n${pdfText}`, file.name);
      }
    }
  }

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
