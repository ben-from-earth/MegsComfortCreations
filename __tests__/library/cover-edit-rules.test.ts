import {
  canAddBook,
  removeLibraryCover,
  toggleKeptCover,
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

describe('canAddBook', () => {
  test('requires a title and one kept cover', () => {
    expect(canAddBook('  ', covers)).toBe(false);
    expect(canAddBook('Crawdads', [{ selected: false }])).toBe(false);
    expect(canAddBook('  Crawdads  ', covers)).toBe(true);
  });
});
