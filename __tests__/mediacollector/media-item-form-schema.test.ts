import { collectorFormSchema } from '@/mediacollector/collector-form/collectorFormSchema';
import {
  convertMediaItemFormToDatabaseItem,
  convertMediaItemToForm,
  mediaItemFormSchema,
  PLACEHOLDER_MEDIA_IMAGE_URL,
  toFormImages,
  type MediaItemForm,
} from '@/mediacollector/collector-form/mediaItemFormSchema';
import type { PostSavedMediaItem } from 'lib/interfaces/globalInterfaces';

function createMediaItemForm(
  overrides: Partial<MediaItemForm> = {},
): MediaItemForm {
  return {
    type: 'book',
    images: [
      {
        url: 'https://img/default.png',
        selected: true,
        isDefault: true,
        spineColor: '#111111',
      },
    ],
    blockInfo: {
      title: 'Dune',
      author: 'Frank Herbert',
      pubYear: 1965,
      pageCount: 412,
      spineColor: '#111111',
      genres: ['Science Fiction'],
    },
    blockID: 'BLK-1',
    isDatabase: false,
    ...overrides,
  };
}

describe('mediaItemFormSchema', () => {
  test('parses one collected card', () => {
    const item = createMediaItemForm();

    expect(mediaItemFormSchema.parse(item)).toEqual(item);
  });

  test('rejects string pubYear', () => {
    const result = mediaItemFormSchema.safeParse({
      type: 'book',
      images: [
        {
          url: 'https://img/default.png',
          selected: true,
          isDefault: true,
          spineColor: '#111111',
        },
      ],
      blockInfo: {
        title: 'Dune',
        author: 'Frank Herbert',
        pubYear: '1965',
        pageCount: 412,
        spineColor: '#111111',
        genres: ['Science Fiction'],
      },
      blockID: 'BLK-1',
      isDatabase: false,
    });

    expect(result.success).toBe(false);
  });

  test('collector collectedData is an array of MediaItemForm', () => {
    const parsed = collectorFormSchema.parse({
      orderNumber: '1001',
      customerName: 'Ada Lovelace',
      bookClubRepeat: 1,
      collectionList: {
        book: [],
        movie: [],
        videoGame: [],
        album: [],
      },
      collectedData: [createMediaItemForm()],
      pngFormat: '3',
    });

    expect(parsed.collectedData).toEqual([createMediaItemForm()]);
  });
});

describe('convertMediaItemToForm', () => {
  const persistedBook: PostSavedMediaItem = {
    id: 'book-42',
    title: 'The Left Hand of Darkness',
    spineColor: '#abcdef',
    mediaType: 'book',
    author: 'Ursula K. Le Guin',
    pubYear: 1969,
    pageCount: 304,
    images: [
      {
        url: 'https://img/cover.png',
        isDefault: true,
        spineColor: '#123456',
      },
    ],
  };

  test('maps a persisted database row onto MediaItemForm', () => {
    const form = convertMediaItemToForm({
      item: persistedBook,
      type: 'book',
      genres: ['Science Fiction'],
    });

    expect(form).toEqual({
      type: 'book',
      images: [
        {
          url: 'https://img/cover.png',
          selected: true,
          isDefault: true,
          spineColor: '#123456',
        },
      ],
      blockInfo: {
        title: 'The Left Hand of Darkness',
        spineColor: '#123456',
        genres: ['Science Fiction'],
        author: 'Ursula K. Le Guin',
        pubYear: 1969,
        pageCount: 304,
      },
      blockID: 'book-42',
      isDatabase: true,
    });
    expect(mediaItemFormSchema.parse(form)).toEqual(form);
  });

  test('fills a placeholder image when the persisted row has none', () => {
    const form = convertMediaItemToForm({
      item: {
        id: 'movie-1',
        title: 'Arrival',
        spineColor: '#ffffff',
        images: [],
      },
      type: 'movie',
    });

    expect(form.images).toEqual([
      {
        url: PLACEHOLDER_MEDIA_IMAGE_URL,
        selected: true,
        isDefault: true,
        spineColor: '#ffffff',
      },
    ]);
    expect(mediaItemFormSchema.parse(form)).toEqual(form);
  });
});

describe('toFormImages', () => {
  test('rejects a database-hit block with empty images at export validation', () => {
    const databaseHit = createMediaItemForm({
      isDatabase: true,
      images: [],
    });

    expect(mediaItemFormSchema.safeParse(databaseHit).success).toBe(false);
  });

  test('makes an empty-image collected block valid before form replace', () => {
    const databaseHit = createMediaItemForm({
      isDatabase: true,
      images: [],
    });

    const normalized = {
      ...databaseHit,
      images: toFormImages(databaseHit.images, databaseHit.blockInfo.spineColor),
    };

    expect(mediaItemFormSchema.parse(normalized).images).toEqual([
      {
        url: PLACEHOLDER_MEDIA_IMAGE_URL,
        selected: true,
        isDefault: true,
        spineColor: '#111111',
      },
    ]);
  });
});

describe('convertMediaItemFormToDatabaseItem', () => {
  test('maps MediaItemForm onto the database.edit item payload', () => {
    const form = createMediaItemForm({
      blockID: 'book-42',
      isDatabase: true,
    });

    expect(convertMediaItemFormToDatabaseItem(form)).toEqual({
      id: 'book-42',
      title: 'Dune',
      spineColor: '#111111',
      images: form.images,
      author: 'Frank Herbert',
      pageCount: 412,
      pubYear: 1965,
    });
  });

  test('coerces missing book fields to null for the edit schema', () => {
    const form = createMediaItemForm({
      blockInfo: {
        title: 'Arrival',
        spineColor: '#ffffff',
        genres: [],
      },
    });

    expect(convertMediaItemFormToDatabaseItem(form)).toMatchObject({
      author: null,
      pageCount: null,
      pubYear: null,
    });
  });
});
