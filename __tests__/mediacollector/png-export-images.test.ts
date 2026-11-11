import {
  buildPNGExportImages,
  type PNGExportImage,
} from '@/mediacollector/png-export-images';
import type { MediaItemForm } from '@/mediacollector/collector-form/media-item-form-schema';

const bookBlock: MediaItemForm = {
  type: 'book',
  images: [
    {
      url: 'https://img/default.png',
      selected: false,
      isDefault: true,
      spineColor: '#111111',
    },
    {
      url: 'https://img/selected.png',
      selected: true,
      isDefault: false,
      spineColor: '#222222',
    },
  ],
  blockInfo: {
    title: 'Dune',
    spineColor: '#222222',
    genres: [],
  },
  blockID: 'BLK-1',
  isDatabase: false,
};

const databaseBookBlock: MediaItemForm = {
  ...bookBlock,
  blockID: 'BLK-DB',
  isDatabase: true,
  images: [
    {
      url: 'https://img/database.png',
      selected: true,
      isDefault: true,
      spineColor: '#333333',
    },
  ],
};

const newBookBlock: MediaItemForm = {
  ...bookBlock,
  blockID: 'BLK-NEW',
  isDatabase: false,
  images: [
    {
      url: 'https://img/new.png',
      selected: true,
      isDefault: true,
      spineColor: '#444444',
    },
  ],
};

describe('buildPNGExportImages', () => {
  test('uses the selected image when one is chosen', () => {
    expect(buildPNGExportImages([bookBlock])).toEqual<PNGExportImage[]>([
      {
        url: 'https://img/selected.png',
        type: 'book',
        spineColor: '#222222',
      },
    ]);
  });

  test('falls back to the default image when none is selected', () => {
    expect(
      buildPNGExportImages([
        {
          ...bookBlock,
          images: bookBlock.images.map((image) => ({
            ...image,
            selected: false,
          })),
        },
      ]),
    ).toEqual<PNGExportImage[]>([
      {
        url: 'https://img/default.png',
        type: 'book',
        spineColor: '#111111',
      },
    ]);
  });

  test('keeps collectedData order when no export ID list is given', () => {
    expect(
      buildPNGExportImages([newBookBlock, databaseBookBlock]),
    ).toEqual<PNGExportImage[]>([
      {
        url: 'https://img/new.png',
        type: 'book',
        spineColor: '#444444',
      },
      {
        url: 'https://img/database.png',
        type: 'book',
        spineColor: '#333333',
      },
    ]);
  });

  test('orders remaining covers by the export ID list, not collectedData order', () => {
    expect(
      buildPNGExportImages(
        [newBookBlock, databaseBookBlock],
        [databaseBookBlock.blockID, newBookBlock.blockID],
      ),
    ).toEqual<PNGExportImage[]>([
      {
        url: 'https://img/database.png',
        type: 'book',
        spineColor: '#333333',
      },
      {
        url: 'https://img/new.png',
        type: 'book',
        spineColor: '#444444',
      },
    ]);
  });

  test('adds the Spotify code URL for an album with a code color', () => {
    const albumBlock: MediaItemForm = {
      type: 'album',
      images: [
        {
          url: 'https://img/rumours.png',
          selected: true,
          isDefault: true,
          spineColor: '#ffffff',
          codeColor: '#E8DCC4',
        },
      ],
      blockInfo: {
        title: 'Rumours',
        spineColor: '#ffffff',
        genres: [],
        spotifyAlbumUri: 'spotify:album:rumours',
      },
      blockID: 'BLK-ALBUM',
      isDatabase: false,
    };

    expect(buildPNGExportImages([albumBlock])).toStrictEqual<PNGExportImage[]>(
      [
        {
          url: 'https://img/rumours.png',
          type: 'album',
          spineColor: '#ffffff',
          codeUrl:
            'https://scannables.scdn.co/uri/plain/png/e8dcc4/black/640/spotify:album:rumours',
        },
      ],
    );
    expect(
      buildPNGExportImages([
        {
          ...albumBlock,
          images: [{ ...albumBlock.images[0], codeColor: null }],
        },
      ])[0],
    ).not.toHaveProperty('codeUrl');
  });

  test('omits export IDs that are no longer in collectedData', () => {
    expect(
      buildPNGExportImages(
        [newBookBlock],
        [databaseBookBlock.blockID, newBookBlock.blockID],
      ),
    ).toEqual<PNGExportImage[]>([
      {
        url: 'https://img/new.png',
        type: 'book',
        spineColor: '#444444',
      },
    ]);
  });
});
