import { getPageWindow } from '@/library/page-window';

describe('getPageWindow', () => {
  test('returns nothing when there are no pages', () => {
    expect(getPageWindow(1, 0)).toEqual([]);
    expect(getPageWindow(1, -1)).toEqual([]);
  });

  test('includes the first page, the last page, the current page, and its neighbors', () => {
    expect(getPageWindow(1, 1)).toEqual([1]);
    expect(getPageWindow(1, 10)).toEqual([1, 2, 'ellipsis', 10]);
    expect(getPageWindow(5, 10)).toEqual([
      1,
      'ellipsis',
      4,
      5,
      6,
      'ellipsis',
      10,
    ]);
    expect(getPageWindow(10, 10)).toEqual([1, 'ellipsis', 9, 10]);
  });

  test('leaves out the ellipsis when the included pages are consecutive', () => {
    expect(getPageWindow(3, 5)).toEqual([1, 2, 3, 4, 5]);
    expect(getPageWindow(2, 4)).toEqual([1, 2, 3, 4]);
  });
});