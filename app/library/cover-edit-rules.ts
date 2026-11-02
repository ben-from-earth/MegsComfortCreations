type LibraryCoverFlags = {
  selected: boolean;
  isDefault: boolean;
};

export function removeLibraryCover<T extends LibraryCoverFlags>(
  images: T[],
  index: number,
): T[] {
  const removed = images[index];
  if (images.length <= 1 || removed == null) {
    return images;
  }

  const remaining = images.filter((_, imageIndex) => imageIndex !== index);
  return remaining.map((image, imageIndex) => ({
    ...image,
    isDefault: removed.isDefault ? imageIndex === 0 : image.isDefault,
    selected: removed.selected ? imageIndex === 0 : image.selected,
  }));
}

export function toggleKeptCover<T extends LibraryCoverFlags>(
  images: T[],
  index: number,
): T[] {
  if (images[index] == null) {
    return images;
  }

  const toggled = images.map((image, imageIndex) =>
    imageIndex === index ? { ...image, selected: !image.selected } : image,
  );
  const firstKeptIndex = toggled.findIndex((image) => image.selected);
  return toggled.map((image, imageIndex) => ({
    ...image,
    isDefault: imageIndex === firstKeptIndex,
  }));
}
