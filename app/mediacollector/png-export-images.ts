import { spotifyCodeUrl } from 'lib/trpc/routers/collect/actions/spotify-code';
import type { MediaItemForm } from './collector-form/media-item-form-schema';

export type PNGExportImage = {
  url: string;
  type: MediaItemForm['type'];
  spineColor: string;
  codeUrl?: string;
  console?: string;
};

function getImageForPNG(block: MediaItemForm) {
  const selectedImage = block.images.find((image) => image.selected);
  if (selectedImage) {
    return selectedImage;
  }

  const defaultImage = block.images.find((image) => image.isDefault);
  return defaultImage ?? block.images[0];
}

export function buildPNGExportImages(
  collectedData: MediaItemForm[],
  pngExportBlockIds?: readonly string[],
): PNGExportImage[] {
  const collectedBlocksById = new Map(
    collectedData.map((block) => [block.blockID, block]),
  );
  const blocksForExport =
    pngExportBlockIds === undefined
      ? collectedData
      : pngExportBlockIds.flatMap((blockID) => {
          const block = collectedBlocksById.get(blockID);
          return block ? [block] : [];
        });

  return blocksForExport.flatMap((block) => {
    const image = getImageForPNG(block);
    if (!image) {
      return [];
    }

    const { spotifyAlbumUri } = block.blockInfo;
    const codeUrl =
      block.type === 'album' && spotifyAlbumUri && image.codeColor
        ? spotifyCodeUrl(spotifyAlbumUri, image.codeColor)
        : null;

    const consoleName =
      block.type === 'videoGame'
        ? (block.blockInfo.console?.trim() ?? '')
        : undefined;

    return [
      {
        url: image.url,
        type: block.type,
        spineColor: image.spineColor,
        ...(codeUrl ? { codeUrl } : {}),
        ...(consoleName !== undefined ? { console: consoleName } : {}),
      },
    ];
  });
}
