import { cellVariationFor } from 'lib/helpers/output-png';

describe('cellVariationFor', () => {
  test('places a movie as one full 3:2 cover', () => {
    expect(cellVariationFor(3, 'movie')).toEqual({
      cellWidth: 392,
      imageWidth: 392,
      pairGap: 0,
      mode: 'single',
    });
    expect(cellVariationFor(5, 'movie')).toEqual({
      cellWidth: 486,
      imageWidth: 486,
      pairGap: 0,
      mode: 'single',
    });
  });

  test('keeps books as a spine pair and albums as a single cover', () => {
    expect(cellVariationFor(3, 'book')).toMatchObject({
      mode: 'pair',
      pairGap: 39,
    });
    expect(cellVariationFor(5, 'book')).toMatchObject({
      mode: 'pair',
      pairGap: 57,
    });
    expect(cellVariationFor(3, 'album').mode).toBe('single');
    expect(cellVariationFor(5, 'videoGame').mode).toBe('pair');
  });
});
