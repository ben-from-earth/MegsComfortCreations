type LibraryCoverFlags = {
  selected: boolean;
  isDefault: boolean;
};

export function removeLibraryCover<T extends LibraryCoverFlags>(
  images: T[],
  index: number,
): T[] {
  if (images.length <= 1 || index < 0 || index >= images.length) {
    return images;
  }

  const removed = images[index];

  const remaining = images.filter((_, imageIndex) => imageIndex !== index);
  return remaining.map((image, imageIndex) => ({
    ...image,
    isDefault: removed.isDefault ? imageIndex === 0 : image.isDefault,
    selected: removed.selected ? imageIndex === 0 : image.selected,
  }));
}

function findKeptIndexFrom(images: LibraryCoverFlags[], startIndex: number) {
  for (let offset = 0; offset < images.length; offset++) {
    const imageIndex = (startIndex + offset) % images.length;
    if (images[imageIndex]?.selected) {
      return imageIndex;
    }
  }
  return -1;
}

export function toggleKeptCover<T extends LibraryCoverFlags>(
  images: T[],
  index: number,
): T[] {
  if (index < 0 || index >= images.length) {
    return images;
  }

  const toggled = images.map((image, imageIndex) =>
    imageIndex === index ? { ...image, selected: !image.selected } : image,
  );
  const keptDefaultIndex = toggled.findIndex(
    (image) => image.isDefault && image.selected,
  );
  const defaultIndex =
    keptDefaultIndex >= 0 ? keptDefaultIndex : findKeptIndexFrom(toggled, index);
  return toggled.map((image, imageIndex) => ({
    ...image,
    isDefault: imageIndex === defaultIndex,
  }));
}

export type AddBookCoverState<T> = {
  images: T[];
  focusedIndex: number | null;
};

export type AddBookCoverAction =
  | { type: 'focus' | 'toggleKeep' | 'setDefault'; index: number }
  | { type: 'setSpineColor'; spineColor: string };

export function updateAddBookCovers<
  T extends LibraryCoverFlags & { spineColor: string },
>(
  { images, focusedIndex }: AddBookCoverState<T>,
  action: AddBookCoverAction,
): AddBookCoverState<T> {
  switch (action.type) {
    case 'focus':
      return { images, focusedIndex: action.index };
    case 'toggleKeep': {
      const toggled = toggleKeptCover(images, action.index);
      return {
        images: toggled,
        focusedIndex: toggled[action.index]?.selected
          ? action.index
          : focusedIndex,
      };
    }
    case 'setDefault':
      if (!images[action.index]?.selected) {
        return { images, focusedIndex };
      }
      return {
        images: images.map((image, imageIndex) => ({
          ...image,
          isDefault: imageIndex === action.index,
        })),
        focusedIndex,
      };
    case 'setSpineColor':
      return {
        images: images.map((image, imageIndex) =>
          imageIndex === focusedIndex
            ? { ...image, spineColor: action.spineColor }
            : image,
        ),
        focusedIndex,
      };
  }
}
