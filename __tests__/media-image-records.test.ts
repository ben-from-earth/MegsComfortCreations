import {
  coversToPersistForNewItem,
  listCoversToAppend,
  resolveAndPersistImageList,
} from 'lib/media-storage/media-image-records';
import { persistExternalImageToS3 } from 'lib/media-storage/local-image-storage';

jest.mock('lib/media-storage/local-image-storage', () => ({
  persistExternalImageToS3: jest.fn(),
}));

describe('resolveAndPersistImageList', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('accepts string and object URL payloads and reports invalid payloads', async () => {
    const mockedPersistExternalImageToS3 = persistExternalImageToS3 as jest.Mock;
    mockedPersistExternalImageToS3.mockImplementation(async ({ sourceUrl }) => ({
      publicPath: sourceUrl,
      mimeType: 'image/png',
      sizeBytes: 100,
      sourceUrl,
    }));

    const response = await resolveAndPersistImageList(
      { mediaType: 'book', mediaId: 'book-1' },
      [
        'https://example.com/cover-1.png',
        { url: 'https://example.com/cover-2.png' },
        { src: 'https://example.com/cover-3.png' },
        { image: 'invalid' },
      ],
    );

    expect(response.images).toHaveLength(3);
    expect(response.failures).toEqual([
      {
        sourceUrl: '',
        message: 'Invalid image payload. Expected an image URL string.',
      },
    ]);
  });

  test('keeps a cover code color on the persisted image', async () => {
    const mockedPersistExternalImageToS3 = persistExternalImageToS3 as jest.Mock;
    mockedPersistExternalImageToS3.mockImplementation(async ({ sourceUrl }) => ({
      publicPath: sourceUrl,
      mimeType: 'image/png',
      sizeBytes: 100,
      sourceUrl,
    }));

    const response = await resolveAndPersistImageList(
      { mediaType: 'album', mediaId: 'album-1' },
      [
        {
          url: 'https://example.com/cover.png',
          spineColor: '#ffffff',
          codeColor: '#112233',
        },
      ],
    );

    expect(response.images[0]).toMatchObject({
      spineColor: '#ffffff',
      codeColor: '#112233',
    });
  });
});

describe('coversToPersistForNewItem', () => {
  const covers = [
    {
      url: 'https://img/first.png',
      selected: true,
      isDefault: false,
      spineColor: '#111111',
    },
    {
      url: 'https://img/chosen.png',
      selected: true,
      isDefault: true,
      spineColor: '#abcdef',
    },
    {
      url: 'https://img/dropped.png',
      selected: false,
      isDefault: false,
      spineColor: '#999999',
    },
  ];

  test('keeps an explicit default and leaves spine colors unchanged', () => {
    expect(coversToPersistForNewItem(covers)).toEqual([
      {
        url: 'https://img/first.png',
        selected: true,
        isDefault: false,
        spineColor: '#111111',
      },
      {
        url: 'https://img/chosen.png',
        selected: true,
        isDefault: true,
        spineColor: '#abcdef',
      },
    ]);
  });

  test('uses the first kept cover when none is marked default', () => {
    expect(
      coversToPersistForNewItem([
        covers[0],
        { ...covers[1], isDefault: false },
        covers[2],
      ]),
    ).toEqual([
      {
        url: 'https://img/first.png',
        selected: true,
        isDefault: true,
        spineColor: '#111111',
      },
      {
        url: 'https://img/chosen.png',
        selected: true,
        isDefault: false,
        spineColor: '#abcdef',
      },
    ]);
  });

  test('falls back to the first image when nothing is kept', () => {
    expect(
      coversToPersistForNewItem([
        { ...covers[0], selected: false },
        { ...covers[1], selected: false, isDefault: true },
      ]),
    ).toEqual([
      {
        url: 'https://img/first.png',
        selected: true,
        isDefault: true,
        spineColor: '#111111',
      },
    ]);
  });

  test('keeps a single selected cover as the default', () => {
    expect(
      coversToPersistForNewItem([
        { ...covers[0], selected: false, isDefault: true },
        { ...covers[1], isDefault: false },
      ]),
    ).toEqual([
      {
        url: 'https://img/chosen.png',
        selected: true,
        isDefault: true,
        spineColor: '#abcdef',
      },
    ]);
  });
});

describe('listCoversToAppend', () => {
  const storedCover = { url: 'https://img/stored.png' };

  test('returns only covers whose urls are not already stored', () => {
    expect(
      listCoversToAppend(
        [storedCover],
        [
          { url: 'https://img/stored.png', spineColor: '#111111' },
          { url: 'https://img/new.png', spineColor: '#abcdef' },
        ],
      ),
    ).toEqual([
      {
        url: 'https://img/new.png',
        isDefault: false,
        spineColor: '#abcdef',
      },
    ]);
  });

  test('makes the first new cover the default when the book has no stored covers', () => {
    expect(
      listCoversToAppend(
        [],
        [
          { url: '  https://img/new.png  ', spineColor: '#abcdef' },
          { url: 'https://img/new.png', spineColor: '#000000' },
        ],
      ),
    ).toEqual([
      {
        url: 'https://img/new.png',
        isDefault: true,
        spineColor: '#abcdef',
      },
    ]);
  });

  test('skips blank urls', () => {
    expect(
      listCoversToAppend([], [{ url: '   ', spineColor: '#ffffff' }]),
    ).toEqual([]);
  });
});
