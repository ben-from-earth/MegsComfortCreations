const MAX_SPINE_COLORS = 5;
const BUCKET_SHIFT = 4;
const MIN_COLOR_DISTANCE = 48;
const MIN_OPAQUE_ALPHA = 128;

type ColorBucket = { red: number; green: number; blue: number; count: number };

const toHexChannel = (value: number) =>
  Math.round(value).toString(16).padStart(2, '0');

/**
 * Returns up to five distinct hex colors from raw RGB or RGBA pixels, most
 * common first. Near-duplicate colors are dropped in favor of the more common one.
 */
export function extractSpinePalette(
  pixels: Uint8Array,
  channels: number,
): string[] {
  const buckets = new Map<number, ColorBucket>();

  for (let offset = 0; offset + 2 < pixels.length; offset += channels) {
    if (channels === 4 && pixels[offset + 3] < MIN_OPAQUE_ALPHA) {
      continue;
    }
    const red = pixels[offset];
    const green = pixels[offset + 1];
    const blue = pixels[offset + 2];
    const key =
      ((red >> BUCKET_SHIFT) << 8) |
      ((green >> BUCKET_SHIFT) << 4) |
      (blue >> BUCKET_SHIFT);
    const bucket = buckets.get(key);
    if (bucket) {
      bucket.red += red;
      bucket.green += green;
      bucket.blue += blue;
      bucket.count += 1;
    } else {
      buckets.set(key, { red, green, blue, count: 1 });
    }
  }

  const averagedColors = [...buckets.values()]
    .toSorted((first, second) => second.count - first.count)
    .map(({ red, green, blue, count }) => ({
      red: red / count,
      green: green / count,
      blue: blue / count,
    }));

  const palette: typeof averagedColors = [];
  for (const color of averagedColors) {
    const isDistinct = palette.every(
      (chosen) =>
        Math.hypot(
          chosen.red - color.red,
          chosen.green - color.green,
          chosen.blue - color.blue,
        ) >= MIN_COLOR_DISTANCE,
    );
    if (isDistinct) {
      palette.push(color);
    }
    if (palette.length === MAX_SPINE_COLORS) {
      break;
    }
  }

  return palette.map(
    ({ red, green, blue }) =>
      `#${toHexChannel(red)}${toHexChannel(green)}${toHexChannel(blue)}`,
  );
}
