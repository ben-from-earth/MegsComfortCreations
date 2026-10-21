import {
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
});

describe('listCoversToAppend', () => {
  const storedCover = { url: 'https://img/stored.png' };

  test('appends a new cover without replacing covers already stored', () => {
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
