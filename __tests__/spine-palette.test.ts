import { extractSpinePalette } from 'lib/helpers/spine-palette';

function pixelsOf(colors: [number, number, number][], channels = 3) {
  return Uint8Array.from(
    colors.flatMap((color) => (channels === 4 ? [...color, 255] : color)),
  );
}

function repeat(
  color: [number, number, number],
  times: number,
): [number, number, number][] {
  return Array.from({ length: times }, () => color);
}

describe('extractSpinePalette', () => {
  test('returns the single color of a flat image', () => {
    expect(extractSpinePalette(pixelsOf(repeat([40, 90, 160], 20)), 3)).toEqual(
      ['#285aa0'],
    );
  });

  test('orders a two-color image by how much of the cover each color covers', () => {
    const pixels = pixelsOf([
      ...repeat([250, 250, 250], 5),
      ...repeat([176, 83, 58], 15),
    ]);
    expect(extractSpinePalette(pixels, 3)).toEqual(['#b0533a', '#fafafa']);
  });

  test('drops near-duplicate shades and caps the palette at five colors', () => {
    const pixels = pixelsOf(
      [
        ...repeat([200, 30, 30], 10),
        ...repeat([210, 40, 35], 9),
        ...repeat([30, 200, 30], 8),
        ...repeat([30, 30, 200], 7),
        ...repeat([240, 240, 240], 6),
        ...repeat([20, 20, 20], 5),
        ...repeat([200, 200, 30], 4),
      ],
      4,
    );
    expect(extractSpinePalette(pixels, 4)).toEqual([
      '#c81e1e',
      '#1ec81e',
      '#1e1ec8',
      '#f0f0f0',
      '#141414',
    ]);
  });

  test('ignores transparent pixels', () => {
    const pixels = Uint8Array.from([
      ...[0, 0, 0, 0],
      ...[0, 0, 0, 0],
      ...[90, 120, 60, 255],
    ]);
    expect(extractSpinePalette(pixels, 4)).toEqual(['#5a783c']);
  });
});
