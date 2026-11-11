import { z } from 'zod';
import { MEDIA_TYPES, type MediaType } from 'lib/constants/media-types';
import { MOVIE_FORMATS } from 'lib/constants/movie-formats';
import type {
  MediaImageItem,
  PostSavedMediaItem,
} from 'lib/interfaces/global-interfaces';

export const PLACEHOLDER_MEDIA_IMAGE_URL = '/images/placeholder-image.png';

export const mediaImageSelectionSchema = z.object({
  url: z.string(),
  selected: z.boolean(),
  isDefault: z.boolean(),
  spineColor: z.string(),
  codeColor: z.string().nullable().optional(),
  uploaded: z.boolean().optional(),
  palette: z.array(z.string()).optional(),
});

export const mediaItemBlockInfoSchema = z.object({
  title: z.string().trim().min(1, 'Title is Required'),
  spineColor: z.string(),
  genres: z.array(z.string()),
  author: z.string().nullable().optional(),
  pubYear: z.number().nullable().optional(),
  pageCount: z.number().nullable().optional(),
  movieFormat: z.enum(MOVIE_FORMATS).optional(),
  artist: z.string().nullable().optional(),
  releaseDate: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/)
    .nullable()
    .optional(),
  spotifyAlbumUri: z.string().nullable().optional(),
  codeStatus: z.string().optional(),
  codeGeneratedFor: z
    .object({
      title: z.string(),
      artist: z.string(),
    })
    .nullable()
    .optional(),
});

const mediaItemFormFields = {
  type: z.enum(MEDIA_TYPES),
  images: z.array(mediaImageSelectionSchema),
  blockInfo: mediaItemBlockInfoSchema,
  blockID: z.string(),
  isDatabase: z.boolean(),
};

function rejectMissingCover(
  item: { type: MediaType; images: unknown[] },
  ctx: z.RefinementCtx,
) {
  if (item.type !== 'album' && item.images.length === 0) {
    ctx.addIssue({
      code: 'custom',
      path: ['images'],
      message: 'Cover image is Required',
    });
  }
}

export const mediaItemFormSchema = z
  .object(mediaItemFormFields)
  .superRefine(rejectMissingCover);

export const addBookFormSchema = z
  .object({
    ...mediaItemFormFields,
    blockInfo: mediaItemBlockInfoSchema.extend({
      author: mediaItemBlockInfoSchema.shape.author.refine(
        (author) => typeof author === 'string' && author.trim() !== '',
        'Author is Required',
      ),
    }),
  })
  .superRefine(rejectMissingCover);

export type MediaItemForm = z.infer<typeof mediaItemFormSchema>;

export function toFormImages(
  images: (MediaImageItem & { palette?: string[] })[] | undefined,
  fallbackSpineColor: string,
  options?: { allowEmpty?: boolean },
): MediaItemForm['images'] {
  if (options?.allowEmpty && (!images || images.length === 0)) {
    return [];
  }

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
    codeColor: image.codeColor ?? null,
    palette: image.palette,
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

export function getMediaItemFormDefaultValues(
  item?: MediaItemForm,
): MediaItemForm {
  if (item) {
    return item;
  }
  return {
    type: 'book',
    blockID: crypto.randomUUID(),
    isDatabase: false,
    blockInfo: {
      title: '',
      author: null,
      pubYear: null,
      pageCount: null,
      genres: [],
      spineColor: '#ffffff',
    },
    images: [],
  };
}

export function convertMediaItemToForm({
  item,
  genres,
}: {
  item: PostSavedMediaItem;
  genres?: string[];
}): MediaItemForm {
  const images = toFormImages(item.images, item.spineColor, {
    allowEmpty: item.mediaType === 'album',
  });
  const defaultImage = images.find((image) => image.isDefault) ?? images[0];
  const spineColor = defaultImage?.spineColor ?? item.spineColor;
  const baseBlockInfo = {
    title: item.title,
    spineColor,
    genres: genres ?? item.genres ?? [],
    author: null,
    pubYear: null,
    pageCount: null,
  };

  switch (item.mediaType) {
    case 'book':
      return {
        type: 'book',
        images,
        blockInfo: {
          ...baseBlockInfo,
          author: item.author,
          pubYear: item.pubYear,
          pageCount: item.pageCount,
        },
        blockID: item.id,
        isDatabase: true,
      };
    case 'movie':
      return {
        type: 'movie',
        images,
        blockInfo: {
          ...baseBlockInfo,
          movieFormat: item.movieFormat,
        },
        blockID: item.id,
        isDatabase: true,
      };
    case 'videoGame':
      return {
        type: 'videoGame',
        images,
        blockInfo: baseBlockInfo,
        blockID: item.id,
        isDatabase: true,
      };
    case 'album':
      return {
        type: 'album',
        images,
        blockInfo: {
          ...baseBlockInfo,
          artist: item.artist,
          releaseDate: item.releaseDate,
          spotifyAlbumUri: item.spotifyAlbumUri,
          ...(item.spotifyAlbumUri
            ? {
                codeStatus: 'generated',
                codeGeneratedFor: {
                  title: item.title,
                  artist: item.artist,
                },
              }
            : {}),
        },
        blockID: item.id,
        isDatabase: true,
      };
  }
}

export function savedMediaItemFromForm(form: MediaItemForm): PostSavedMediaItem {
  const shared = {
    id: form.blockID,
    title: form.blockInfo.title,
    spineColor: form.blockInfo.spineColor,
    images: form.images,
    genres: form.blockInfo.genres,
  };

  switch (form.type) {
    case 'book':
      return {
        mediaType: 'book',
        ...shared,
        author: form.blockInfo.author ?? '',
        pageCount: form.blockInfo.pageCount ?? null,
        pubYear: form.blockInfo.pubYear ?? null,
      };
    case 'movie':
      return {
        mediaType: 'movie',
        ...shared,
        movieFormat: form.blockInfo.movieFormat ?? 'dvd',
      };
    case 'videoGame':
      return {
        mediaType: 'videoGame',
        ...shared,
      };
    case 'album':
      return {
        mediaType: 'album',
        ...shared,
        artist: form.blockInfo.artist ?? '',
        releaseDate: form.blockInfo.releaseDate ?? null,
        spotifyAlbumUri: form.blockInfo.spotifyAlbumUri ?? null,
      };
  }
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
    ...(form.blockInfo.movieFormat
      ? { movieFormat: form.blockInfo.movieFormat }
      : {}),
    ...(form.type === 'album'
      ? {
          artist: form.blockInfo.artist ?? null,
          releaseDate: form.blockInfo.releaseDate ?? null,
          ...(form.blockInfo.spotifyAlbumUri
            ? { spotifyAlbumUri: form.blockInfo.spotifyAlbumUri }
            : {}),
        }
      : {}),
  };
}
