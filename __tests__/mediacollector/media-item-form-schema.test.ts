import {
  convertMediaItemFormToDatabaseItem,
  convertMediaItemToForm,
  getMediaItemFormDefaultValues,
  mediaItemFormSchema,
  PLACEHOLDER_MEDIA_IMAGE_URL,
  toFormImages,
  type MediaItemForm,
} from '@/mediacollector/collector-form/media-item-form-schema';
import type { PostSavedMediaItem } from 'lib/interfaces/global-interfaces';

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
  test('rejects an empty title', () => {
    const result = mediaItemFormSchema.safeParse(
      createMediaItemForm({
        blockInfo: {
          title: '   ',
          spineColor: '#111111',
          genres: ['Science Fiction'],
        },
      }),
    );

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(
        result.error.issues.some(
          (issue) =>
            issue.path.join('.') === 'blockInfo.title' &&
            issue.message === 'Title is Required',
        ),
      ).toBe(true);
    }
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

  test('rejects a collected block with empty images', () => {
    const result = mediaItemFormSchema.safeParse(
      createMediaItemForm({ isDatabase: true, images: [] }),
    );

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(
        result.error.issues.some(
          (issue) =>
            issue.path.join('.') === 'images' &&
            issue.message === 'Cover image is Required',
        ),
      ).toBe(true);
    }
  });
});

describe('toFormImages', () => {
  test('inserts a placeholder so an empty-image block can enter the form', () => {
    const emptyBlock = createMediaItemForm({ images: [] });

    expect(
      mediaItemFormSchema.parse({
        ...emptyBlock,
        images: toFormImages(emptyBlock.images, emptyBlock.blockInfo.spineColor),
      }).images,
    ).toEqual([
      {
        url: PLACEHOLDER_MEDIA_IMAGE_URL,
        selected: true,
        isDefault: true,
        spineColor: '#111111',
      },
    ]);
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
    expect(
      convertMediaItemToForm({
        item: persistedBook,
        type: 'book',
        genres: ['Science Fiction'],
      }),
    ).toEqual({
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
  });

  test('fills a placeholder image when the persisted row has none', () => {
    expect(
      convertMediaItemToForm({
        item: {
          id: 'movie-1',
          title: 'Arrival',
          spineColor: '#ffffff',
          images: [],
        },
        type: 'movie',
      }).images,
    ).toEqual([
      {
        url: PLACEHOLDER_MEDIA_IMAGE_URL,
        selected: true,
        isDefault: true,
        spineColor: '#ffffff',
      },
    ]);
  });
});

describe('getMediaItemFormDefaultValues', () => {
  test('returns a blank book that fails parse on empty title and has no covers', () => {
    const defaults = getMediaItemFormDefaultValues();

    expect(defaults).toMatchObject({
      type: 'book',
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
    });
    expect(defaults.blockID).toEqual(expect.any(String));
    expect(defaults.blockID.length).toBeGreaterThan(0);

    const result = mediaItemFormSchema.safeParse(defaults);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(
        result.error.issues.some(
          (issue) =>
            issue.path.join('.') === 'blockInfo.title' &&
            issue.message === 'Title is Required',
        ),
      ).toBe(true);
    }
  });

  test('preserves fields when an item is passed', () => {
    const item = createMediaItemForm();

    expect(getMediaItemFormDefaultValues(item)).toEqual(item);
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
    expect(
      convertMediaItemFormToDatabaseItem(
        createMediaItemForm({
          blockInfo: {
            title: 'Arrival',
            spineColor: '#ffffff',
            genres: [],
          },
        }),
      ),
    ).toMatchObject({
      author: null,
      pageCount: null,
      pubYear: null,
    });
  });
});
