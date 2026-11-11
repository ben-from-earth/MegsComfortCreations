import {
  countDatabaseItemChanges,
  diffDatabaseItem,
  restoreDatabaseItem,
} from '@/mediacollector/database-item-changes';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';

const originalBlock: MediaItemForm = {
  type: 'book',
  blockID: 'BOOK-1',
  isDatabase: true,
  blockInfo: {
    title: 'Where the Crawdads Sing',
    author: 'Delia Owens',
    pubYear: 2018,
    pageCount: 368,
    genres: ['Mystery', 'Classic'],
    spineColor: '#3b6a8a',
  },
  images: [
    {
      url: 'https://img/blue.png',
      selected: true,
      isDefault: true,
      spineColor: '#3b6a8a',
    },
    {
      url: 'https://img/orange.png',
      selected: false,
      isDefault: false,
      spineColor: '#c27c4e',
    },
  ],
};

describe('diffDatabaseItem / countDatabaseItemChanges', () => {
  test('reports no changes for the untouched database item', () => {
    const changes = diffDatabaseItem(originalBlock, originalBlock);
    expect(countDatabaseItemChanges(changes)).toBe(0);
  });

  test('ignores picking a different cover for this order', () => {
    const current: MediaItemForm = {
      ...originalBlock,
      images: originalBlock.images.map((image) => ({
        ...image,
        selected: !image.selected,
      })),
    };
    expect(
      countDatabaseItemChanges(diffDatabaseItem(originalBlock, current)),
    ).toBe(0);
  });

  test('compares genres as a set', () => {
    const current: MediaItemForm = {
      ...originalBlock,
      blockInfo: { ...originalBlock.blockInfo, genres: ['Classic', 'Mystery'] },
    };
    expect(diffDatabaseItem(originalBlock, current).genres).toBe(false);
  });

  test('counts each changed field and all spine colors as one change', () => {
    const current: MediaItemForm = {
      ...originalBlock,
      blockInfo: {
        ...originalBlock.blockInfo,
        author: 'D. Owens',
        pageCount: 370,
        genres: ['Mystery'],
      },
      images: [
        { ...originalBlock.images[0], isDefault: false, spineColor: '#111111' },
        { ...originalBlock.images[1], isDefault: true, spineColor: '#222222' },
      ],
    };

    const changes = diffDatabaseItem(originalBlock, current);

    expect(changes).toEqual({
      title: false,
      author: true,
      pubYear: false,
      pageCount: true,
      genres: true,
      defaultCover: true,
      coverList: false,
      changedSpineColorUrls: ['https://img/blue.png', 'https://img/orange.png'],
      artist: false,
      releaseDate: false,
      spotifyAlbumUri: false,
      changedCodeColorUrls: [],
    });
    expect(countDatabaseItemChanges(changes)).toBe(5);
  });

  test('treats spine colors case-insensitively and ignores new uploads', () => {
    const current: MediaItemForm = {
      ...originalBlock,
      images: [
        { ...originalBlock.images[0], spineColor: '#3B6A8A' },
        originalBlock.images[1],
        {
          url: '/uploads/new.png',
          selected: false,
          isDefault: false,
          spineColor: '#000000',
          uploaded: true,
        },
      ],
    };
    expect(
      diffDatabaseItem(originalBlock, current).changedSpineColorUrls,
    ).toEqual([]);
  });

  test('counts an added or removed cover url as one change', () => {
    const added: MediaItemForm = {
      ...originalBlock,
      images: [
        ...originalBlock.images,
        {
          url: '/uploads/new.png',
          selected: false,
          isDefault: false,
          spineColor: '#000000',
          uploaded: true,
        },
      ],
    };
    const addedChanges = diffDatabaseItem(originalBlock, added);
    expect(addedChanges.coverList).toBe(true);
    expect(countDatabaseItemChanges(addedChanges)).toBe(1);

    const removed: MediaItemForm = {
      ...originalBlock,
      images: [originalBlock.images[0]],
    };
    const removedChanges = diffDatabaseItem(originalBlock, removed);
    expect(removedChanges.coverList).toBe(true);
    expect(countDatabaseItemChanges(removedChanges)).toBe(1);
  });

  test('ignores cover order and a repeated url', () => {
    const reordered: MediaItemForm = {
      ...originalBlock,
      images: [originalBlock.images[1], originalBlock.images[0]],
    };
    expect(diffDatabaseItem(originalBlock, reordered).coverList).toBe(false);
    expect(
      countDatabaseItemChanges(diffDatabaseItem(originalBlock, reordered)),
    ).toBe(0);

    const repeated: MediaItemForm = {
      ...originalBlock,
      images: [...originalBlock.images, { ...originalBlock.images[0] }],
    };
    expect(diffDatabaseItem(originalBlock, repeated).coverList).toBe(false);
  });
});

describe('album database item changes', () => {
  const originalAlbum: MediaItemForm = {
    type: 'album',
    blockID: 'ALBUM-1',
    isDatabase: true,
    blockInfo: {
      title: 'Rumours',
      artist: 'Fleetwood Mac',
      releaseDate: '1977-02-04',
      spotifyAlbumUri: 'spotify:album:rumours',
      genres: [],
      spineColor: '#ffffff',
    },
    images: [
      {
        url: 'https://img/rumours.png',
        selected: true,
        isDefault: true,
        spineColor: '#ffffff',
        codeColor: '#112233',
      },
      {
        url: 'https://img/alt.png',
        selected: false,
        isDefault: false,
        spineColor: '#ffffff',
        codeColor: null,
      },
    ],
  };

  test('counts artist, release date, uri, and one cover code color', () => {
    const current: MediaItemForm = {
      ...originalAlbum,
      blockInfo: {
        ...originalAlbum.blockInfo,
        artist: 'Fleetwood',
        releaseDate: '1977-02-05',
        spotifyAlbumUri: 'spotify:album:other',
        genres: ['Rock'],
      },
      images: [
        { ...originalAlbum.images[0], codeColor: '#aabbcc' },
        originalAlbum.images[1],
      ],
    };

    const changes = diffDatabaseItem(originalAlbum, current);

    expect(changes.artist).toBe(true);
    expect(changes.releaseDate).toBe(true);
    expect(changes.spotifyAlbumUri).toBe(true);
    expect(changes.genres).toBe(false);
    expect(changes.changedCodeColorUrls).toEqual(['https://img/rumours.png']);
    expect(countDatabaseItemChanges(changes)).toBe(4);
  });

  test('treats code colors case-insensitively', () => {
    const current: MediaItemForm = {
      ...originalAlbum,
      images: [
        { ...originalAlbum.images[0], codeColor: '#112233' },
        originalAlbum.images[1],
      ],
    };
    current.images[0] = {
      ...current.images[0],
      codeColor: '#112233'.toUpperCase(),
    };

    expect(diffDatabaseItem(originalAlbum, current).changedCodeColorUrls).toEqual(
      [],
    );
  });

  test('restores artist, release date, uri, and code colors', () => {
    const current: MediaItemForm = {
      ...originalAlbum,
      blockInfo: {
        ...originalAlbum.blockInfo,
        artist: 'Someone Else',
        releaseDate: null,
        spotifyAlbumUri: null,
      },
      images: [
        { ...originalAlbum.images[0], codeColor: '#000000' },
        { ...originalAlbum.images[1], selected: true },
      ],
    };

    const restored = restoreDatabaseItem(originalAlbum, current);

    expect(
      countDatabaseItemChanges(diffDatabaseItem(originalAlbum, restored)),
    ).toBe(0);
    expect(restored.blockInfo.artist).toBe('Fleetwood Mac');
    expect(restored.blockInfo.releaseDate).toBe('1977-02-04');
    expect(restored.blockInfo.spotifyAlbumUri).toBe('spotify:album:rumours');
    expect(restored.images[0]?.codeColor).toBe('#112233');
  });
});

describe('restoreDatabaseItem', () => {
  test('restores every database value and keeps the order cover choice', () => {
    const current: MediaItemForm = {
      ...originalBlock,
      blockInfo: {
        ...originalBlock.blockInfo,
        title: 'Crawdads',
        genres: [],
        spineColor: '#999999',
      },
      images: [
        { ...originalBlock.images[0], selected: false, isDefault: false },
        {
          ...originalBlock.images[1],
          selected: true,
          isDefault: true,
          spineColor: '#999999',
        },
        {
          url: '/uploads/new.png',
          selected: false,
          isDefault: false,
          spineColor: '#000000',
          uploaded: true,
        },
      ],
    };

    const restored = restoreDatabaseItem(originalBlock, current);

    expect(
      countDatabaseItemChanges(diffDatabaseItem(originalBlock, restored)),
    ).toBe(0);
    expect(restored.blockInfo).toEqual({
      ...originalBlock.blockInfo,
      spineColor: '#c27c4e',
    });
    expect(restored.images).toEqual([
      { ...originalBlock.images[0], selected: false },
      { ...originalBlock.images[1], selected: true },
    ]);
  });

  test('falls back to the original selection when the chosen cover was an upload', () => {
    const current: MediaItemForm = {
      ...originalBlock,
      images: [
        ...originalBlock.images.map((image) => ({ ...image, selected: false })),
        {
          url: '/uploads/new.png',
          selected: true,
          isDefault: false,
          spineColor: '#000000',
          uploaded: true,
        },
      ],
    };

    expect(restoreDatabaseItem(originalBlock, current)).toEqual(originalBlock);
  });
});
