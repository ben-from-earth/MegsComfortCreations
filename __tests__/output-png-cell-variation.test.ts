import { cellVariationFor } from 'lib/helpers/output-png';

describe('cellVariationFor', () => {
  test('places a movie as the 5 mm full cover on either template', () => {
    const movieCell = {
      cellWidth: 486,
      imageWidth: 486,
      pairGap: 0,
      mode: 'single',
    };
    expect(cellVariationFor(3, 'movie')).toEqual(movieCell);
    expect(cellVariationFor(5, 'movie')).toEqual(movieCell);
  });

  test('places an album as two 1 inch covers with a gap on either template', () => {
    const albumCell = {
      cellWidth: 618,
      imageWidth: 300,
      pairGap: 18,
      mode: 'album',
    };
    expect(cellVariationFor(3, 'album')).toEqual(albumCell);
    expect(cellVariationFor(5, 'album')).toEqual(albumCell);
  });

  test('keeps books and video games as a spine pair', () => {
    expect(cellVariationFor(3, 'book')).toMatchObject({
      mode: 'pair',
      pairGap: 39,
    });
    expect(cellVariationFor(5, 'book')).toMatchObject({
      mode: 'pair',
      pairGap: 57,
    });
    expect(cellVariationFor(5, 'videoGame').mode).toBe('pair');
  });
});
