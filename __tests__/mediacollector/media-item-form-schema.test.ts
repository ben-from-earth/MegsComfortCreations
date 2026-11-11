import {
  addBookFormSchema,
  convertMediaItemFormToDatabaseItem,
  convertMediaItemToForm,
  getMediaItemFormDefaultValues,
  mediaItemFormSchema,
  PLACEHOLDER_MEDIA_IMAGE_URL,
  savedMediaItemFromForm,
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

  test('allows an album block with no covers', () => {
    const result = mediaItemFormSchema.safeParse(
      createMediaItemForm({
        type: 'album',
        images: [],
        blockInfo: {
          title: 'Rumours',
          artist: 'Fleetwood Mac',
          spineColor: '#ffffff',
          genres: [],
        },
      }),
    );

    expect(result.success).toBe(true);
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

describe('addBookFormSchema', () => {
  test.each([null, '   '])('rejects author %p', (author) => {
    const form = createMediaItemForm();
    const result = addBookFormSchema.safeParse({
      ...form,
      blockInfo: { ...form.blockInfo, author },
    });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(
        result.error.issues.some(
          (issue) =>
            issue.path.join('.') === 'blockInfo.author' &&
            issue.message === 'Author is Required',
        ),
      ).toBe(true);
    }
  });

  test('accepts a book with an author', () => {
    expect(addBookFormSchema.safeParse(createMediaItemForm()).success).toBe(
      true,
    );
  });
});

describe('mediaItemFormSchema uploaded covers', () => {
  test('keeps an uploaded cover flag through parse', () => {
    const parsed = mediaItemFormSchema.parse(
      createMediaItemForm({
        isDatabase: true,
        images: [
          {
            url: 'https://img/stored.png',
            selected: false,
            isDefault: true,
            spineColor: '#111111',
          },
          {
            url: 'https://img/new.png',
            selected: true,
            isDefault: false,
            spineColor: '#abcdef',
            uploaded: true,
          },
        ],
      }),
    );

    expect(parsed.images[1]?.uploaded).toBe(true);
    expect(parsed.images[0]?.uploaded).toBeUndefined();
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
        codeColor: null,
      },
    ]);
  });

  test('keeps an empty album cover list empty', () => {
    expect(toFormImages([], '#ffffff', { allowEmpty: true })).toEqual([]);
  });

  test('copies a cover code color', () => {
    expect(
      toFormImages(
        [
          {
            url: 'https://img/cover.png',
            isDefault: true,
            selected: true,
            spineColor: '#ffffff',
            codeColor: '#112233',
          },
        ],
        '#ffffff',
      )[0]?.codeColor,
    ).toBe('#112233');
  });

  test('keeps the cover palette from collection', () => {
    expect(
      toFormImages(
        [
          {
            url: 'https://img/cover.png',
            isDefault: true,
            selected: true,
            spineColor: '#b0533a',
            palette: ['#b0533a', '#fafafa'],
          },
        ],
        '#ffffff',
      )[0]?.palette,
    ).toEqual(['#b0533a', '#fafafa']);
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
        item: { ...persistedBook, genres: ['Science Fiction'] },
      }),
    ).toEqual({
      type: 'book',
      images: [
        {
          url: 'https://img/cover.png',
          selected: true,
          isDefault: true,
          spineColor: '#123456',
          codeColor: null,
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

  test('copies a movie format onto the block', () => {
    const vhs = convertMediaItemToForm({
      item: {
        mediaType: 'movie',
        id: 'movie-vhs',
        title: 'Arrival',
        spineColor: '#ffffff',
        images: [],
        movieFormat: 'vhs',
      },
    });

    expect(vhs.blockInfo.movieFormat).toBe('vhs');
    expect(convertMediaItemFormToDatabaseItem(vhs).movieFormat).toBe('vhs');
  });

  test('fills a placeholder image when the persisted row has none', () => {
    expect(
      convertMediaItemToForm({
        item: {
          mediaType: 'movie',
          id: 'movie-1',
          title: 'Arrival',
          spineColor: '#ffffff',
          images: [],
          movieFormat: 'dvd',
        },
      }).images,
    ).toEqual([
      {
        url: PLACEHOLDER_MEDIA_IMAGE_URL,
        selected: true,
        isDefault: true,
        spineColor: '#ffffff',
        codeColor: null,
      },
    ]);
  });

  test('copies album identity and a null code color', () => {
    expect(
      convertMediaItemToForm({
        item: {
          mediaType: 'album',
          id: 'album-1',
          title: 'Rumours',
          spineColor: '#ffffff',
          artist: 'Fleetwood Mac',
          releaseDate: '1977-02-04',
          spotifyAlbumUri: 'spotify:album:rumours',
          images: [
            {
              url: 'https://img/rumours.png',
              isDefault: true,
              spineColor: '#ffffff',
              codeColor: null,
            },
          ],
        },
      }),
    ).toMatchObject({
      type: 'album',
      images: [
        {
          url: 'https://img/rumours.png',
          codeColor: null,
          spineColor: '#ffffff',
        },
      ],
      blockInfo: {
        artist: 'Fleetwood Mac',
        releaseDate: '1977-02-04',
        spotifyAlbumUri: 'spotify:album:rumours',
        codeStatus: 'generated',
        codeGeneratedFor: {
          title: 'Rumours',
          artist: 'Fleetwood Mac',
        },
      },
    });
  });
});

describe('savedMediaItemFromForm', () => {
  test('keeps book and album fields on their own items', () => {
    expect(
      savedMediaItemFromForm(
        createMediaItemForm({
          blockID: 'book-1',
          type: 'book',
          blockInfo: {
            title: 'Dune',
            spineColor: '#111111',
            genres: ['Science Fiction'],
            author: 'Frank Herbert',
            pageCount: 412,
            pubYear: 1965,
          },
        }),
      ),
    ).toEqual({
      mediaType: 'book',
      id: 'book-1',
      title: 'Dune',
      spineColor: '#111111',
      images: [
        {
          url: 'https://img/default.png',
          selected: true,
          isDefault: true,
          spineColor: '#111111',
        },
      ],
      genres: ['Science Fiction'],
      author: 'Frank Herbert',
      pageCount: 412,
      pubYear: 1965,
    });

    expect(
      savedMediaItemFromForm(
        createMediaItemForm({
          blockID: 'album-1',
          type: 'album',
          blockInfo: {
            title: 'Rumours',
            spineColor: '#ffffff',
            genres: [],
            artist: 'Fleetwood Mac',
            releaseDate: '1977-02-04',
            spotifyAlbumUri: 'spotify:album:rumours',
          },
        }),
      ),
    ).toMatchObject({
      mediaType: 'album',
      artist: 'Fleetwood Mac',
      releaseDate: '1977-02-04',
      spotifyAlbumUri: 'spotify:album:rumours',
    });
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
