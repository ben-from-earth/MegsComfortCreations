import type { MediaItemForm } from './collector-form/media-item-form-schema';
import { PLACEHOLDER_MEDIA_IMAGE_URL } from './collector-form/media-item-form-schema';
import { blocksInExportOrder, getImageForPNG } from './png-export-images';

export type CollectorOrderDraft = {
  customerName: string;
  orderNumber: string;
  template: 3 | 5;
  repeatCount: number;
  coverIds: string[];
};

export function buildCollectorOrderDraft(input: {
  customerName: string;
  orderNumber: string;
  template: 3 | 5;
  repeatCount: number;
  collectedData: MediaItemForm[];
  pngExportBlockIds?: readonly string[];
}):
  | { draft: CollectorOrderDraft; missingCoverTitles: [] }
  | { draft: null; missingCoverTitles: string[] } {
  const coverIds: string[] = [];
  const missingCoverTitles: string[] = [];

  for (const block of blocksInExportOrder(
    input.collectedData,
    input.pngExportBlockIds,
  )) {
    const image = getImageForPNG(block);
    if (!image || image.url === PLACEHOLDER_MEDIA_IMAGE_URL) {
      continue;
    }
    if (!image.id) {
      missingCoverTitles.push(block.blockInfo.title);
      continue;
    }
    coverIds.push(image.id);
  }

  if (missingCoverTitles.length > 0) {
    return { draft: null, missingCoverTitles };
  }

  return {
    draft: {
      customerName: input.customerName.trim(),
      orderNumber: input.orderNumber.trim(),
      template: input.template,
      repeatCount: input.repeatCount,
      coverIds,
    },
    missingCoverTitles: [],
  };
}
