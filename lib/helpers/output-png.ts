'use strict';

import sharp, { OverlayOptions, RGBA } from 'sharp';
import axios from 'axios';
import JSZip from 'jszip';

/*
  PNG/ZIP renderer (paginated)

  Exports:
    - outputAuto({ template, images, prefix? }) → { mime, filename, buffer }
        If one page fits → returns a single PNG.
        If multiple pages → returns a ZIP of PNG pages.
    - outputPNGs({ template, images, namePrefix? }) → [{ filename, buffer }]
    - outputZIP({ template, images, prefix? }) → Buffer (zip)

  Input:
  {
    template: 3 | 5,
    images: ImageData[];
  }
*/

export type Template = 3 | 5;

export interface OutputFile {
  filename: string;
  buffer: Buffer;
}

export interface AutoOutput {
  mime: 'image/png' | 'application/zip';
  filename: string;
  buffer: Buffer;
}

export type ImageData = {
  url: string;
  type: 'book' | 'movie' | 'videoGame' | 'album';
  spineColor: string;
  codeUrl?: string;
  console?: string;
};

interface Slot {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface TemplateMetrics {
  rowHeight: number;
  yGap: number;
  xGap: number;
}

interface Variation {
  cellWidth: number;
  imageWidth: number;
  pairGap: number;
  mode: 'pair' | 'single' | 'album';
}

// Fixed canvas
const CANVAS_W = 2025;
const CANVAS_H = 2775;

// Vertical row metrics by template
const TEMPLATE_METRICS: Record<Template, TemplateMetrics> = {
  5: { rowHeight: 324, yGap: 26, xGap: 17 },
  3: { rowHeight: 261, yGap: 18, xGap: 22 },
} as const;

// Outer margins & horizontal gutters
const MARGIN_X = 0;
const MARGIN_Y = 0;

// Canvas background (transparent)
const CANVAS_BG: RGBA = { r: 0, g: 0, b: 0, alpha: 0 };

// Cell layout variations
const variation1: Variation = {
  cellWidth: 387,
  imageWidth: 174,
  pairGap: 39,
  mode: 'pair',
};
const variation3: Variation = {
  cellWidth: 489,
  imageWidth: 216,
  pairGap: 57,
  mode: 'pair',
};
const variation4: Variation = {
  cellWidth: 471,
  imageWidth: 216,
  pairGap: 39,
  mode: 'pair',
};

// The canvas prints at 300 DPI, so an album side is 1 inch.
const ALBUM_SIDE = 300;
const ALBUM_GAP = 18;
const ALBUM_CODE_WIDTH = 240;
const albumVariation: Variation = {
  cellWidth: ALBUM_SIDE * 2 + ALBUM_GAP,
  imageWidth: ALBUM_SIDE,
  pairGap: ALBUM_GAP,
  mode: 'album',
};

function fullCoverVariation(rowHeight: number): Variation {
  const width = Math.round((rowHeight * 3) / 2);
  return {
    cellWidth: width,
    imageWidth: width,
    pairGap: 0,
    mode: 'single',
  };
}

const movieVariation5 = fullCoverVariation(TEMPLATE_METRICS[5].rowHeight);

// A 3DS cover is 1.64 inches wide by 0.77 inches tall. The sheet is 300 DPI.
const THREE_DS_WIDTH = Math.round(1.64 * 300);
const THREE_DS_HEIGHT = Math.round(0.77 * 300);
const threeDsVariation: Variation = {
  cellWidth: THREE_DS_WIDTH,
  imageWidth: THREE_DS_WIDTH,
  pairGap: 0,
  mode: 'single',
};

function is3dsConsole(consoleName: string | undefined) {
  return consoleName?.trim().toLowerCase() === '3ds';
}

// Which variation is used based on format and media type.
// Movies and video games use the 5 mm full cover, on either bleed setting.
// A console of 3DS prints as one 1.64 by 0.77 inch cover.
const COMBOS: Record<
  Template,
  Record<string, Variation> & { default: Variation }
> = {
  3: {
    book: variation1,
    movie: movieVariation5,
    album: albumVariation,
    default: variation1,
  },
  5: {
    book: variation3,
    movie: movieVariation5,
    album: albumVariation,
    default: variation4,
  },
};

export function cellVariationFor(
  template: Template,
  type: string | undefined,
  consoleName?: string,
): Variation {
  if (type === 'videoGame' && is3dsConsole(consoleName)) {
    return threeDsVariation;
  }
  if (type === 'movie' || type === 'videoGame') {
    return movieVariation5;
  }
  return COMBOS[template][type ?? ''] ?? COMBOS[template].default;
}

//Multi Page PNG ZIP

function layoutPage({
  images,
  startIndex,
  template,
  rowHeight,
  yGap,
  canvasW,
  canvasH,
  marginX,
  marginY,
  xGap,
}: {
  images: ImageData[];
  startIndex: number;
  template: Template;
  rowHeight: number;
  yGap: number;
  canvasW: number;
  canvasH: number;
  marginX: number;
  marginY: number;
  xGap: number;
}): { slots: Slot[]; countUsed: number } {
  const slots: Slot[] = [];
  let x = marginX;
  let y = marginY;
  let countUsed = 0;
  // Track the start index of the current row so we can right-align the last cell
  let rowStartIdx = 0;
  const defaultRowAdvance = rowHeight + yGap;
  let rowAdvance = defaultRowAdvance;
  const movieMetrics = TEMPLATE_METRICS[5];

  for (let i = startIndex; i < images.length; i++) {
    const image = images[i];
    const type = image?.type;
    const cfg = cellVariationFor(template, type, image?.console);
    const is3ds = cfg === threeDsVariation;
    const usesFiveMillimeterCover =
      type === 'movie' || (type === 'videoGame' && !is3ds);
    const usesFiveMillimeterGutters = type === 'movie' || type === 'videoGame';
    const w = cfg.cellWidth;
    const h = is3ds
      ? THREE_DS_HEIGHT
      : usesFiveMillimeterCover
        ? movieMetrics.rowHeight
        : type === 'album'
          ? ALBUM_SIDE
          : rowHeight;
    const gapX = usesFiveMillimeterGutters ? movieMetrics.xGap : xGap;
    const cellAdvance = h + (usesFiveMillimeterGutters ? movieMetrics.yGap : yGap);

    if (x > marginX && x + w > canvasW - marginX) {
      // Right-align the last cell of the previous row
      if (slots.length > rowStartIdx) {
        const lastIdx = slots.length - 1;
        const last = slots[lastIdx];
        slots[lastIdx] = { ...last, x: canvasW - marginX - last.w };
      }
      x = marginX;
      y += rowAdvance;
      rowAdvance = defaultRowAdvance;
      if (y + h > canvasH - marginY) break; // full
      // New row starts at current slots length
      rowStartIdx = slots.length;
    }

    if (y + h <= canvasH - marginY) {
      slots.push({ x, y, w, h });
      x += w + gapX;
      rowAdvance = Math.max(rowAdvance, cellAdvance);
      countUsed++;
    } else {
      break;
    }
  }

  // After placing all cells, right-align the last cell of the final row
  if (slots.length > rowStartIdx) {
    const lastIdx = slots.length - 1;
    const last = slots[lastIdx];
    slots[lastIdx] = { ...last, x: canvasW - marginX - last.w };
  }

  return { slots, countUsed };
}

async function renderPage({
  template,
  metrics,
  images,
  startIndex,
}: {
  template: Template;
  metrics: TemplateMetrics;
  images: ImageData[];
  startIndex: number;
}): Promise<{ buffer: Buffer; countUsed: number }> {
  const base = sharp({
    create: {
      width: CANVAS_W,
      height: CANVAS_H,
      channels: 4,
      background: CANVAS_BG,
    },
  });

  const { slots, countUsed } = layoutPage({
    images,
    startIndex,
    template,
    rowHeight: metrics.rowHeight,
    yGap: metrics.yGap,
    canvasW: CANVAS_W,
    canvasH: CANVAS_H,
    marginX: MARGIN_X,
    marginY: MARGIN_Y,
    xGap: metrics.xGap,
  });

  const overlaysNested = await Promise.all(
    slots.map((slot, k) =>
      renderCell(slot, images[startIndex + k], template),
    ),
  );
  const overlays = overlaysNested.flat();

  const buffer = await base.composite(overlays).png().toBuffer();
  return { buffer, countUsed };
}

// Return an array of PNG pages (filename + buffer)
export async function outputPNGs({
  template,
  images = [],
  fileOutputName,
}: {
  template: Template;
  images: ImageData[];
  fileOutputName: string;
}): Promise<OutputFile[]> {
  const metrics = TEMPLATE_METRICS[template];
  if (!metrics) throw new Error('Unsupported template (use 3 or 5)');

  const results: OutputFile[] = [];
  if (images.length === 0) {
    // Optional: single blank page
    const { buffer } = await renderPage({
      template,
      metrics,
      images: [],
      startIndex: 0,
    });
    results.push({ filename: `${fileOutputName}.png`, buffer });
    return results;
  }

  let idx = 0;
  let pageNo = 1;
  while (idx < images.length) {
    const { buffer, countUsed } = await renderPage({
      template,
      metrics,
      images,
      startIndex: idx,
    });
    const filename =
      pageNo === 1 && countUsed >= images.length
        ? `${fileOutputName}.png`
        : `${fileOutputName}_${String(pageNo).padStart(3, '0')}.png`;
    results.push({ filename, buffer });
    if (countUsed === 0) break;
    idx += countUsed;
    pageNo += 1;
  }

  return results;
}

// Return a ZIP buffer containing all PNG pages.
export async function outputZIP({
  template,
  images = [],
  fileOutputName,
}: {
  template: Template;
  images: ImageData[];
  fileOutputName: string;
}): Promise<Buffer> {
  const zip = new JSZip();
  const pages = await outputPNGs({ template, images, fileOutputName });
  for (const { filename, buffer } of pages) {
    zip.file(filename, buffer);
  }
  return await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  });
}

/**
 * Entry point for route: returns either PNG or ZIP based on how many pages fit.
 * - If everything fits on one page → { mime: 'image/png', filename, buffer }
 * - else → { mime: 'application/zip', filename: '<prefix>_pages.zip', buffer }
 */
export async function outputAuto({
  template,
  images = [],
  fileOutputName,
}: {
  template: Template;
  images: ImageData[];
  fileOutputName: string;
}): Promise<AutoOutput> {
  const pages = await outputPNGs({ template, images, fileOutputName });

  if (pages.length === 1) {
    return {
      mime: 'image/png',
      filename: pages[0].filename,
      buffer: pages[0].buffer,
    };
  }
  const zipBuf = await outputZIP({ template, images, fileOutputName });
  return {
    mime: 'application/zip',
    filename: `${fileOutputName}_pages.zip`,
    buffer: zipBuf,
  };
}

// render functions and helpers

async function renderCell(
  slot: Slot,
  img: ImageData | undefined,
  template: Template,
): Promise<OverlayOptions[]> {
  const { url, spineColor, type, codeUrl, console: consoleName } = img || {};
  const cfg = cellVariationFor(template, type, consoleName);

  if (cfg.mode === 'album') {
    const side = cfg.imageWidth;
    const rightX = slot.x + side + cfg.pairGap;
    const [cover, code] = await Promise.all([
      makeBlock(side, side, url),
      codeUrl
        ? fetchImageBufferAxios(codeUrl).then((buffer) =>
            sharp(buffer)
              .resize({ width: ALBUM_CODE_WIDTH })
              .png()
              .toBuffer({ resolveWithObject: true }),
          )
        : null,
    ]);
    const overlays: OverlayOptions[] = [
      { input: cover, left: slot.x, top: slot.y },
      { input: cover, left: rightX, top: slot.y },
    ];
    if (code) {
      overlays.push({
        input: code.data,
        left: rightX + Math.round((side - code.info.width) / 2),
        top: slot.y + Math.round((side - code.info.height) / 2),
      });
    }
    return overlays;
  }

  // Cell background (spine color fills full slot)
  const cellBg = await makeBlock(
    slot.w,
    slot.h,
    (spineColor as string | RGBA) ?? { r: 255, g: 0, b: 0, alpha: 1 },
  );

  if (cfg.mode === 'single') {
    const imageWidth = cfg.imageWidth;
    const imageHeight = slot.h;
    const leftX = slot.x;
    const topY = slot.y;

    const cover = await makeBlock(imageWidth, imageHeight, url);
    return [
      { input: cellBg, left: slot.x, top: slot.y },
      { input: cover, left: leftX, top: topY },
    ];
  }

  // Pair: left rotated 180°, right normal
  const gap = cfg.pairGap;
  const imageWidth = cfg.imageWidth;
  const imageHeight = slot.h;
  const leftX = slot.x;
  const topY = slot.y;

  const [leftImage, rightImage] = await Promise.all([
    makeBlock(imageWidth, imageHeight, url, 180),
    makeBlock(imageWidth, imageHeight, url, 0),
  ]);

  return [
    { input: cellBg, left: slot.x, top: slot.y },
    { input: leftImage, left: leftX, top: topY },
    { input: rightImage, left: leftX + imageWidth + gap, top: topY },
  ];
}

function isHttpUrl(s: unknown): boolean {
  try {
    const u = new URL(String(s));
    return u.protocol === 'http:' || u.protocol === 'https:';
  } catch {
    return false;
  }
}

function getReferer(url: string): string | undefined {
  try {
    const u = new URL(url);
    return `${u.protocol}//${u.host}/`;
  } catch {
    return undefined;
  }
}

function isImageCtype(ctype?: string): boolean {
  if (!ctype) return false;
  const base = String(ctype).split(';')[0].trim().toLowerCase();
  return base.startsWith('image/');
}

async function sniffIsImage(
  buffer: Buffer,
): Promise<{ ok: boolean; mime?: string; ext?: string }> {
  const { fileTypeFromBuffer } = await import('file-type');
  const ft = await fileTypeFromBuffer(buffer).catch(() => null);
  if (ft?.mime?.startsWith('image/'))
    return { ok: true, mime: ft.mime, ext: ft.ext };
  const head = buffer.toString('utf8', 0, 256);
  if (/\<svg[\s>]/i.test(head))
    return { ok: true, mime: 'image/svg+xml', ext: 'svg' };
  return { ok: false };
}

// Fetch an image as Buffer, with retries and MIME sniffing.
export async function fetchImageBufferAxios(
  url: string,
  {
    maxBytes = 8 * 1024 * 1024,
    timeout = 15_000,
    retries = 3,
    backoff = 500,
    requireImage = true,
    sendReferer = true,
  }: {
    maxBytes?: number;
    timeout?: number;
    retries?: number;
    backoff?: number;
    requireImage?: boolean;
    sendReferer?: boolean;
  } = {},
): Promise<Buffer> {
  let attempt = 0;
  const referer = sendReferer ? getReferer(url) : undefined;

  while (true) {
    try {
      const res = await axios.get<ArrayBuffer>(url, {
        responseType: 'arraybuffer',
        maxContentLength: maxBytes,
        maxBodyLength: maxBytes,
        timeout,
        maxRedirects: 3,
        headers: {
          'User-Agent': 'png-renderer/1.0 (+node)',
          Accept: 'image/avif,image/webp,image/*,*/*;q=0.8',
          ...(referer ? { Referer: referer } : {}),
        },
        validateStatus: (s) => s >= 200 && s < 300,
      });

      const buffer = Buffer.from(res.data);
      const ctype = res.headers['content-type'] as string | undefined;

      if (isImageCtype(ctype)) return buffer;

      if (requireImage) {
        const sniff = await sniffIsImage(buffer);
        if (sniff.ok) return buffer;
        throw new Error(`Not an image. Content-Type=${ctype || 'unknown'}`);
      }

      return buffer;
    } catch (err: unknown) {
      // Narrow/cast the error to the shape we expect from Axios
      const e = err as {
        response?: {
          status?: number;
          headers?: Record<string, string | undefined>;
        };
        config?: { url?: string };
        message?: string;
        code?: string | number;
      };

      attempt++;
      const status = e.response?.status;
      const ctype = e.response?.headers?.['content-type'];
      const urlLogged = e.config?.url;
      const message = e.message || String(err);

      console.error(
        `[fetchImageBufferAxios] attempt=${attempt} | message=${message}` +
          (e.code ? ` | code=${e.code}` : '') +
          (status ? ` | status=${status}` : '') +
          (ctype ? ` | ctype=${ctype}` : '') +
          (urlLogged ? ` | url=${urlLogged}` : ''),
      );

      if (attempt > retries) {
        throw new Error(
          `IMAGE_FETCH_FAILED after ${retries} retries: ${e.code || ''} ${message}`,
        );
      }

      const delay = backoff * Math.pow(2, attempt - 1) + Math.random() * 150;
      await new Promise((r) => setTimeout(r, delay));
    }
  }
}

/**
 * Create a block buffer:
 * - If `fill` is an http(s) URL → fetch & fit to w×h
 * - Else treat `fill` as a color (CSS or {r,g,b,alpha})
 * - rotation: if 180 → flip+flop (exact 180°)
 */
async function makeBlock(
  w: number,
  h: number,
  fill?: string | RGBA,
  rotation: 0 | 180 = 0,
): Promise<Buffer> {
  let pipe: sharp.Sharp;

  if (typeof fill === 'string' && isHttpUrl(fill)) {
    try {
      const buf = await fetchImageBufferAxios(fill);
      pipe = sharp(buf).toColourspace('srgb').resize(w, h, { fit: 'fill' });
    } catch (e: unknown) {
      const err = e as { message?: string };

      console.warn(
        '[makeBlock] fetch failed, using placeholder:',
        err.message || String(e),
      );

      pipe = sharp({
        create: {
          width: w,
          height: h,
          channels: 4,
          background: { r: 200, g: 200, b: 200, alpha: 1 },
        },
      });
    }
  } else {
    pipe = sharp({
      create: {
        width: w,
        height: h,
        channels: 4,
        background: (fill as RGBA) ?? { r: 180, g: 180, b: 180, alpha: 1 },
      },
    });
  }

  if (rotation === 180) pipe = pipe.flip().flop();
  return await pipe.png().toBuffer();
}
