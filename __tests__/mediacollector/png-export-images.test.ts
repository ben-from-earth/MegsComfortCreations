import {
  buildPNGExportImages,
  type PNGExportImage,
} from '@/mediacollector/png-export-images';
import type { MediaItemForm } from '@/mediacollector/collector-form/mediaItemFormSchema';

function createBlock(overrides: Partial<MediaItemForm> = {}): MediaItemForm {
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
      genres: [],
    },
    blockID: 'BLK-1',
    isDatabase: false,
    ...overrides,
  };
}

describe('buildPNGExportImages', () => {
  test('uses selected image for database blocks instead of the default image', () => {
    const images = buildPNGExportImages([
      createBlock({
        isDatabase: true,
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
      }),
    ]);

    expect(images).toEqual<PNGExportImage[]>([
      {
        url: 'https://img/selected.png',
        type: 'book',
        spineColor: '#222222',
      },
    ]);
  });

  test('uses selected image for newly collected blocks', () => {
    const images = buildPNGExportImages([
      createBlock({
        isDatabase: false,
        type: 'movie',
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
            spineColor: '#333333',
          },
        ],
      }),
    ]);

    expect(images).toEqual<PNGExportImage[]>([
      {
        url: 'https://img/selected.png',
        type: 'movie',
        spineColor: '#333333',
      },
    ]);
  });

  test('falls back to the default image when no image is selected', () => {
    const images = buildPNGExportImages([
      createBlock({
        images: [
          {
            url: 'https://img/first.png',
            selected: false,
            isDefault: false,
            spineColor: '#111111',
          },
          {
            url: 'https://img/default.png',
            selected: false,
            isDefault: true,
            spineColor: '#444444',
          },
        ],
      }),
    ]);

    expect(images).toEqual<PNGExportImage[]>([
      {
        url: 'https://img/default.png',
        type: 'book',
        spineColor: '#444444',
      },
    ]);
  });
});
