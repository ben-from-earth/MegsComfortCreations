import type { MediaItemForm } from './collector-form/media-item-form-schema';

export interface DatabaseItemChanges {
  title: boolean;
  author: boolean;
  pubYear: boolean;
  pageCount: boolean;
  genres: boolean;
  defaultCover: boolean;
  coverList: boolean;
  changedSpineColorUrls: string[];
}

export interface DatabaseItemEdit {
  block: MediaItemForm;
  original: MediaItemForm;
}

function hasSameGenres(left: string[], right: string[]) {
  const leftGenres = new Set(left);
  const rightGenres = new Set(right);
  return (
    leftGenres.size === rightGenres.size &&
    [...leftGenres].every((genre) => rightGenres.has(genre))
  );
}

function findDefaultCoverUrl(images: MediaItemForm['images']) {
  return images.find((image) => image.isDefault)?.url;
}

function hasSameCoverUrls(
  left: MediaItemForm['images'],
  right: MediaItemForm['images'],
) {
  const leftUrls = new Set(left.map((image) => image.url));
  const rightUrls = new Set(right.map((image) => image.url));
  return (
    leftUrls.size === rightUrls.size &&
    [...leftUrls].every((url) => rightUrls.has(url))
  );
}

export function diffDatabaseItem(
  original: MediaItemForm,
  current: MediaItemForm,
): DatabaseItemChanges {
  const originalSpineColorByUrl = new Map(
    original.images.map((image) => [image.url, image.spineColor.toLowerCase()]),
  );

  return {
    title: current.blockInfo.title !== original.blockInfo.title,
    author:
      (current.blockInfo.author ?? '') !== (original.blockInfo.author ?? ''),
    pubYear:
      (current.blockInfo.pubYear ?? null) !==
      (original.blockInfo.pubYear ?? null),
    pageCount:
      (current.blockInfo.pageCount ?? null) !==
      (original.blockInfo.pageCount ?? null),
    genres: !hasSameGenres(current.blockInfo.genres, original.blockInfo.genres),
    defaultCover:
      findDefaultCoverUrl(current.images) !==
      findDefaultCoverUrl(original.images),
    coverList: !hasSameCoverUrls(current.images, original.images),
    changedSpineColorUrls: current.images.flatMap((image) => {
      const originalSpineColor = originalSpineColorByUrl.get(image.url);
      return originalSpineColor !== undefined &&
        originalSpineColor !== image.spineColor.toLowerCase()
        ? [image.url]
        : [];
    }),
  };
}

export function countDatabaseItemChanges(changes: DatabaseItemChanges) {
  return [
    changes.title,
    changes.author,
    changes.pubYear,
    changes.pageCount,
    changes.genres,
    changes.defaultCover,
    changes.coverList,
    changes.changedSpineColorUrls.length > 0,
  ].filter(Boolean).length;
}

export function restoreDatabaseItem(
  original: MediaItemForm,
  current: MediaItemForm,
): MediaItemForm {
  const selectedUrl = current.images.find((image) => image.selected)?.url;
  const keepsSelection = original.images.some(
    (image) => image.url === selectedUrl,
  );
  const images = keepsSelection
    ? original.images.map((image) => ({
        ...image,
        selected: image.url === selectedUrl,
      }))
    : original.images;
  const selectedImage = images.find((image) => image.selected);

  return {
    ...original,
    images,
    blockInfo: {
      ...original.blockInfo,
      spineColor: selectedImage?.spineColor ?? original.blockInfo.spineColor,
    },
  };
}
