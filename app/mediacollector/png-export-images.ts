import type { MediaItemForm } from './collector-form/mediaItemFormSchema';

export type PNGExportImage = {
  url: string;
  type: MediaItemForm['type'];
  spineColor: string;
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
): PNGExportImage[] {
  return collectedData.flatMap((block) => {
    const image = getImageForPNG(block);
    if (!image) {
      return [];
    }

    return [
      {
        url: image.url,
        type: block.type,
        spineColor: image.spineColor,
      },
    ];
  });
}
