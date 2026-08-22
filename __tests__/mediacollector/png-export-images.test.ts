import {
  buildPNGExportImages,
  type PNGExportImage,
} from '@/mediacollector/png-export-images';
import type { MediaItemForm } from '@/mediacollector/collector-form/mediaItemFormSchema';

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
});
