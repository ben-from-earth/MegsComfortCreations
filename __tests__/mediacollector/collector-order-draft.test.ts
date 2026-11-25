import { buildCollectorOrderDraft } from '@/mediacollector/collector-order-draft';
import {
  PLACEHOLDER_MEDIA_IMAGE_URL,
  type MediaItemForm,
} from '@/mediacollector/collector-form/media-item-form-schema';

function block(
  blockID: string,
  images: MediaItemForm['images'],
  type: MediaItemForm['type'] = 'book',
): MediaItemForm {
  return {
    type,
    images,
    blockInfo: { title: blockID, spineColor: '#111111', genres: [] },
    blockID,
    isDatabase: true,
  };
}

function cover(
  id: string | undefined,
  overrides: Partial<MediaItemForm['images'][number]> = {},
): MediaItemForm['images'][number] {
  return {
    ...(id ? { id } : {}),
    url: `https://img/${id ?? 'new'}.png`,
    selected: false,
    isDefault: false,
    spineColor: '#111111',
    ...overrides,
  };
}

const header = {
  customerName: '  Ada Lovelace ',
  orderNumber: ' 1001 ',
  template: 5 as const,
  repeatCount: 3,
};

describe('buildCollectorOrderDraft', () => {
  test('follows print order and picks the cover the print uses', () => {
    const result = buildCollectorOrderDraft({
      ...header,
      collectedData: [
        block('first', [cover('a-default', { isDefault: true }), cover('a-picked', { selected: true })]),
        block('second', [cover('b-other'), cover('b-default', { isDefault: true })]),
      ],
      pngExportBlockIds: ['second', 'first'],
    });

    expect(result.draft).toEqual({
      customerName: 'Ada Lovelace',
      orderNumber: '1001',
      template: 5,
      repeatCount: 3,
      coverIds: ['b-default', 'a-picked'],
    });
  });

  test('skips blocks without a cover and stores repeat once', () => {
    const result = buildCollectorOrderDraft({
      ...header,
      collectedData: [
        block('album', [], 'album'),
        block('placeholder', [
          cover(undefined, { url: PLACEHOLDER_MEDIA_IMAGE_URL, selected: true }),
        ]),
        block('book', [cover('c', { selected: true })]),
      ],
    });

    expect(result.draft?.coverIds).toEqual(['c']);
    expect(result.draft?.repeatCount).toBe(3);
  });

  test('reports a printed cover that has no stored id', () => {
    expect(
      buildCollectorOrderDraft({
        ...header,
        collectedData: [block('unsaved', [cover(undefined, { selected: true })])],
      }),
    ).toEqual({ draft: null, missingCoverTitles: ['unsaved'] });
  });
});
