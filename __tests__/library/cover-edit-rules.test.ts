import {
  removeLibraryCover,
  toggleKeptCover,
  updateAddBookCovers,
} from '@/library/cover-edit-rules';

const covers = [
  { url: 'first', selected: true, isDefault: true },
  { url: 'second', selected: false, isDefault: false },
  { url: 'third', selected: false, isDefault: false },
];

describe('removeLibraryCover', () => {
  test('leaves the only cover in place', () => {
    const only = [covers[0]];
    expect(removeLibraryCover(only, 0)).toBe(only);
  });

  test('moves the default and the selection to the first remaining cover', () => {
    expect(removeLibraryCover(covers, 0)).toEqual([
      { url: 'second', selected: true, isDefault: true },
      { url: 'third', selected: false, isDefault: false },
    ]);
  });

  test('keeps the current selection when the deleted cover was not selected', () => {
    const withLaterDefault = [
      { url: 'first', selected: true, isDefault: false },
      { url: 'second', selected: false, isDefault: true },
    ];
    expect(removeLibraryCover(withLaterDefault, 1)).toEqual([
      { url: 'first', selected: true, isDefault: true },
    ]);
  });

  test('selects the first remaining cover when a non-default selection is deleted', () => {
    const selectedSecond = [
      { url: 'first', selected: false, isDefault: true },
      { url: 'second', selected: true, isDefault: false },
    ];
    expect(removeLibraryCover(selectedSecond, 1)).toEqual([
      { url: 'first', selected: true, isDefault: true },
    ]);
  });
});

describe('toggleKeptCover', () => {
  test('makes the first kept cover the default', () => {
    const noneKept = covers.map((cover) => ({
      ...cover,
      selected: false,
      isDefault: false,
    }));
    expect(toggleKeptCover(noneKept, 1)).toEqual([
      { url: 'first', selected: false, isDefault: false },
      { url: 'second', selected: true, isDefault: true },
      { url: 'third', selected: false, isDefault: false },
    ]);
  });

  test('moves the default to another kept cover when the default is un-kept', () => {
    const twoKept = [
      { url: 'first', selected: true, isDefault: true },
      { url: 'second', selected: true, isDefault: false },
    ];
    expect(toggleKeptCover(twoKept, 0)).toEqual([
      { url: 'first', selected: false, isDefault: false },
      { url: 'second', selected: true, isDefault: true },
    ]);
  });

  test('keeps the existing default when another cover is kept', () => {
    const laterDefault = [
      { url: 'first', selected: false, isDefault: false },
      { url: 'second', selected: true, isDefault: true },
    ];
    expect(toggleKeptCover(laterDefault, 0)).toEqual([
      { url: 'first', selected: true, isDefault: false },
      { url: 'second', selected: true, isDefault: true },
    ]);
  });

  test('moves the default to the next kept cover after the un-kept default', () => {
    const allKept = covers.map((cover) => ({ ...cover, selected: true }));
    const middleDefault = allKept.map((cover, index) => ({
      ...cover,
      isDefault: index === 1,
    }));
    expect(
      toggleKeptCover(middleDefault, 1).map((cover) => cover.isDefault),
    ).toEqual([false, false, true]);
  });

  test('clears the default when the last kept cover is un-kept', () => {
    const oneKept = [
      { url: 'first', selected: true, isDefault: true },
      { url: 'second', selected: false, isDefault: false },
    ];
    expect(toggleKeptCover(oneKept, 0)).toEqual([
      { url: 'first', selected: false, isDefault: false },
      { url: 'second', selected: false, isDefault: false },
    ]);
  });
});

describe('updateAddBookCovers', () => {
  const searched = [
    { url: 'first', selected: true, isDefault: true, spineColor: '#111111' },
    { url: 'second', selected: false, isDefault: false, spineColor: '#222222' },
    { url: 'third', selected: false, isDefault: false, spineColor: '#333333' },
  ];

  test('focusing a cover leaves kept and default state alone', () => {
    const next = updateAddBookCovers(
      { images: searched, focusedIndex: 0 },
      { type: 'focus', index: 2 },
    );
    expect(next).toEqual({ images: searched, focusedIndex: 2 });
  });

  test('keeping a cover also focuses it', () => {
    const next = updateAddBookCovers(
      { images: searched, focusedIndex: 0 },
      { type: 'toggleKeep', index: 1 },
    );
    expect(next.focusedIndex).toBe(1);
    expect(next.images.map((image) => image.selected)).toEqual([
      true,
      true,
      false,
    ]);
    expect(next.images.map((image) => image.isDefault)).toEqual([
      true,
      false,
      false,
    ]);
  });

  test('un-keeping a cover keeps the current focus', () => {
    expect(
      updateAddBookCovers(
        { images: searched, focusedIndex: 0 },
        { type: 'toggleKeep', index: 0 },
      ).focusedIndex,
    ).toBe(0);
  });

  test('sets the default only on a kept cover', () => {
    const twoKept = searched.map((image, index) => ({
      ...image,
      selected: index < 2,
    }));
    expect(
      updateAddBookCovers(
        { images: twoKept, focusedIndex: 1 },
        { type: 'setDefault', index: 1 },
      ).images.map((image) => image.isDefault),
    ).toEqual([false, true, false]);
    expect(
      updateAddBookCovers(
        { images: twoKept, focusedIndex: 2 },
        { type: 'setDefault', index: 2 },
      ).images,
    ).toBe(twoKept);
  });

  test('applies a spine color only to the focused cover', () => {
    expect(
      updateAddBookCovers(
        { images: searched, focusedIndex: 2 },
        { type: 'setSpineColor', spineColor: '#abcdef' },
      ).images.map((image) => image.spineColor),
    ).toEqual(['#111111', '#222222', '#abcdef']);
  });
});
