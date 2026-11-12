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
    expect(cellVariationFor(3, 'album')).toBe(cellVariationFor(5, 'album'));
    expect(cellVariationFor(3, 'album')).toEqual(albumCell);
    expect(cellVariationFor(5, 'album')).toEqual(albumCell);
  });

  test('keeps books as a spine pair', () => {
    expect(cellVariationFor(3, 'book')).toMatchObject({
      mode: 'pair',
      pairGap: 39,
    });
    expect(cellVariationFor(5, 'book')).toMatchObject({
      mode: 'pair',
      pairGap: 57,
    });
  });

  test('places a video game as the 5 mm full cover on either template', () => {
    const fullCover = {
      cellWidth: 486,
      imageWidth: 486,
      pairGap: 0,
      mode: 'single',
    };
    expect(cellVariationFor(3, 'videoGame', 'Switch')).toEqual(fullCover);
    expect(cellVariationFor(5, 'videoGame', 'PlayStation')).toEqual(fullCover);
    expect(cellVariationFor(5, 'videoGame')).toEqual(fullCover);
    expect(cellVariationFor(3, 'videoGame', 'Nintendo 3DS')).toEqual(fullCover);
  });

  test('places a 3DS game as a 1.64 by 0.77 inch cover', () => {
    const threeDsCell = {
      cellWidth: 492,
      imageWidth: 492,
      pairGap: 0,
      mode: 'single',
    };
    expect(cellVariationFor(3, 'videoGame', '3DS')).toEqual(threeDsCell);
    expect(cellVariationFor(5, 'videoGame', '3ds')).toEqual(threeDsCell);
    expect(cellVariationFor(5, 'videoGame', '  3DS  ')).toEqual(threeDsCell);
  });
});
