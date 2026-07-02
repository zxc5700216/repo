import type { PdfPageSummary, PdfSummary } from "@/lib/logistics/types";
import { inferPdfMetaFromFileName } from "@/lib/logistics/utils";

function buildMockPageSummaries(meta: ReturnType<typeof inferPdfMetaFromFileName>, pageCount: number) {
  const pages: PdfPageSummary[] = [];

  for (let index = 0; index < pageCount; index += 1) {
    pages.push({
      pageNumber: index + 1,
      shipmentName: meta.shipmentName,
      warehouseCode: meta.warehouseCode,
      fbaCode: meta.fbaCode,
      totalBoxes: meta.totalBoxes,
      fbaBoxCode: meta.fbaCode ? `${meta.fbaCode}U${String(index + 1).padStart(6, "0")}` : "",
      positionCode: `P1 - B${index + 1}`,
      skuType: index % 2 === 0 ? "Mixed SKUs" : "Single SKU",
      sku: "",
      qty: null,
    });
  }

  return pages;
}

export async function parsePdfFile(file: File): Promise<PdfSummary> {
  const meta = inferPdfMetaFromFileName(file.name);

  // Browser-only lightweight fallback:
  // current version infers metadata from file name and creates page placeholders.
  // If the file name contains total boxes, use it as page count; otherwise default to 1.
  const pageCount = meta.totalBoxes ?? 1;
  const pages = buildMockPageSummaries(meta, pageCount);

  return {
    ...meta,
    pages,
  };
}
