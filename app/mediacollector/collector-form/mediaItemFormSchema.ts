import { z } from 'zod';
import { MEDIA_TYPES, type MediaType } from 'lib/constants/mediaTypes';
import type {
  MediaImageItem,
  PostSavedMediaItem,
} from 'lib/interfaces/globalInterfaces';

export const PLACEHOLDER_MEDIA_IMAGE_URL = '/images/placeholder-image.png';

export const mediaImageSelectionSchema = z.object({
  url: z.string(),
  selected: z.boolean(),
  isDefault: z.boolean(),
  spineColor: z.string(),
});

export const mediaItemBlockInfoSchema = z.object({
  title: z.string().trim().min(1, 'Title is Required'),
  spineColor: z.string(),
  genres: z.array(z.string()),
  author: z.string().nullable().optional(),
  pubYear: z.number().nullable().optional(),
  pageCount: z.number().nullable().optional(),
});

export const mediaItemFormSchema = z.object({
  type: z.enum(MEDIA_TYPES),
  images: z
    .array(mediaImageSelectionSchema)
    .min(1, 'Cover image is Required'),
  blockInfo: mediaItemBlockInfoSchema,
  blockID: z.string(),
  isDatabase: z.boolean(),
});

export type MediaItemForm = z.infer<typeof mediaItemFormSchema>;

export function toFormImages(
  images: MediaImageItem[] | undefined,
  fallbackSpineColor: string,
): MediaItemForm['images'] {
  const source =
    images && images.length > 0
      ? images
      : [
          {
            url: PLACEHOLDER_MEDIA_IMAGE_URL,
            isDefault: true,
            selected: true,
            spineColor: fallbackSpineColor,
          },
        ];

  const hasDefault = source.some((image) => image.isDefault);
  const mapped = source.map((image, index) => ({
    url: image.url,
    selected: image.selected ?? false,
    isDefault: hasDefault ? image.isDefault : index === 0,
    spineColor: image.spineColor,
  }));

  if (mapped.some((image) => image.selected)) {
    return mapped;
  }

  const defaultIndex = mapped.findIndex((image) => image.isDefault);
  const selectedIndex = defaultIndex >= 0 ? defaultIndex : 0;
  return mapped.map((image, index) => ({
    ...image,
    selected: index === selectedIndex,
  }));
}

export function convertMediaItemToForm({
  item,
  type,
  genres = [],
}: {
  item: PostSavedMediaItem;
  type: MediaType;
  genres?: string[];
}): MediaItemForm {
  const images = toFormImages(item.images, item.spineColor);
  const defaultImage = images.find((image) => image.isDefault) ?? images[0];

  return {
    type,
    images,
    blockInfo: {
      title: item.title,
      spineColor: defaultImage?.spineColor ?? item.spineColor,
      genres,
      author: item.author ?? null,
      pubYear: item.pubYear ?? null,
      pageCount: item.pageCount ?? null,
    },
    blockID: item.id,
    isDatabase: true,
  };
}

export function convertMediaItemFormToDatabaseItem(form: MediaItemForm) {
  return {
    id: form.blockID,
    title: form.blockInfo.title,
    spineColor: form.blockInfo.spineColor,
    images: form.images,
    author: form.blockInfo.author ?? null,
    pageCount: form.blockInfo.pageCount ?? null,
    pubYear: form.blockInfo.pubYear ?? null,
  };
}
