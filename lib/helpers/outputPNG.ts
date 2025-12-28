'use strict';

// library imports
import sharp, { OverlayOptions, RGBA } from 'sharp';
import axios from 'axios';
import JSZip from 'jszip';

// interfaces and types
import { ImageData } from 'lib/state/slices/pngCollectionSlice';

/*
  PNG/ZIP renderer (paginated)

  Exports:
    - outputAuto({ template, images, prefix? }) → { mime, filename, buffer }
        If one page fits → returns a single PNG.
        If multiple pages → returns a ZIP of PNG pages.
    - outputPNGs({ template, images, namePrefix? }) → [{ filename, buffer }]
    - outputZIP({ template, images, prefix? }) → Buffer (zip)
    - outputPNG({ template, images }) → Buffer (single page only)

  Input:
  {
    template: 3 | 5,
    images: Array<{ url: string, spineColor: string, type: string }>
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
  mode: 'pair' | 'single';
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
const variation2: Variation = {
  cellWidth: 261,
  imageWidth: 261,
  pairGap: 0,
  mode: 'single',
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
const variation5: Variation = {
  cellWidth: 324,
  imageWidth: 324,
  pairGap: 0,
  mode: 'single',
};

// Which variation is used based on format and media type
const COMBOS: Record<
  Template,
  Record<string, Variation> & { default: Variation }
> = {
  3: {
    book: variation1,
    movie: variation1,
    videoGame: variation1,
    album: variation2,
    default: variation1,
  },
  5: {
    book: variation3,
    movie: variation4,
    videoGame: variation4,
    album: variation5,
    default: variation4,
  },
} as const;

//Single Page PNG
export async function outputPNG({
  template,
  images = [],
}: {
  template: Template;
  images: ImageData[];
}): Promise<Buffer> {
  const metrics = TEMPLATE_METRICS[template];
  if (!metrics) throw new Error('Unsupported template (use 3 or 5)');

  const base = sharp({
    create: {
      width: CANVAS_W,
      height: CANVAS_H,
      channels: 4,
      background: CANVAS_BG,
    },
  });

  const slots = layoutSlots({
    images,
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
    slots.map((slot, i) =>
      renderCell(slot, images[i], template, metrics.rowHeight),
    ),
  );
  const overlays = overlaysNested.flat();

  try {
    return await base.composite(overlays).png().toBuffer();
  } catch (err: unknown) {
    const e = err as { errors?: { message?: string }[]; message?: string };

    if (Array.isArray(e.errors)) {
      for (const sub of e.errors) {
        console.error('composite overlay error', sub?.message || sub);
      }
    }

    console.error('sharp composite error', e.message || err);
    throw err;
  }
}

/**
 * Build slots (single canvas) using explicit cell widths per (template,type).
 * Flow left to right; wrap when adding the next cell would overflow canvas width.
 */
function layoutSlots({
  images,
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
  template: Template;
  rowHeight: number;
  yGap: number;
  canvasW: number;
  canvasH: number;
  marginX: number;
  marginY: number;
  xGap: number;
}): Slot[] {
  const slots: Slot[] = [];
  let x = marginX;
  let y = marginY;

  for (let i = 0; i < images.length; i++) {
    const type = images[i].type;
    const cfg = COMBOS[template][type] ?? COMBOS[template].default;
    const w = cfg.cellWidth;
    const h = rowHeight;

    // Wrap if this cell would overflow the row
    if (x > marginX && x + w > canvasW - marginX) {
      x = marginX;
      y += h + yGap;
      if (y + h > canvasH - marginY) break; // no more vertical room
    }

    slots.push({ x, y, w, h });
    x += w + xGap;
  }

  return slots;
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

  for (let i = startIndex; i < images.length; i++) {
    const type = images[i]?.type;
    const cfg = COMBOS[template][type] ?? COMBOS[template].default;
    const w = cfg.cellWidth;
    const h = rowHeight;

    if (x > marginX && x + w > canvasW - marginX) {
      x = marginX;
      y += h + yGap;
      if (y + h > canvasH - marginY) break; // full
    }

    if (y + h <= canvasH - marginY) {
      slots.push({ x, y, w, h });
      x += w + xGap;
      countUsed++;
    } else {
      break;
    }
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
      renderCell(slot, images[startIndex + k], template, metrics.rowHeight),
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
  namePrefix = 'MMC_Output',
}: {
  template: Template;
  images: ImageData[];
  namePrefix?: string;
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
    results.push({ filename: `${namePrefix}_001.png`, buffer });
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
    const filename = `${namePrefix}_${String(pageNo).padStart(3, '0')}.png`;
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
  prefix = 'grid',
}: {
  template: Template;
  images: ImageData[];
  prefix?: string;
}): Promise<Buffer> {
  const zip = new JSZip();
  const pages = await outputPNGs({ template, images, namePrefix: prefix });
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
  prefix = 'grid',
}: {
  template: Template;
  images: ImageData[];
  prefix?: string;
}): Promise<AutoOutput> {
  const pages = await outputPNGs({ template, images, namePrefix: prefix });

  if (pages.length === 1) {
    return {
      mime: 'image/png',
      filename: pages[0].filename,
      buffer: pages[0].buffer,
    };
  }
  const zipBuf = await outputZIP({ template, images, prefix });
  return {
    mime: 'application/zip',
    filename: `${prefix}_pages.zip`,
    buffer: zipBuf,
  };
}

// render functions and helpers

async function renderCell(
  slot: Slot,
  img: ImageData | undefined,
  template: Template,
  rowHeight: number,
): Promise<OverlayOptions[]> {
  const { url, spineColor, type } = img || {};
  const cfg = COMBOS[template][type ?? ''] ?? COMBOS[template].default;

  // Cell background (spine color fills full slot)
  const cellBg = await makeBlock(
    slot.w,
    slot.h,
    (spineColor as string | RGBA) ?? { r: 255, g: 0, b: 0, alpha: 1 },
  );

  if (cfg.mode === 'single') {
    const imageWidth = cfg.imageWidth;
    const imageHeight = rowHeight;
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
  const imageHeight = rowHeight;
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
async function fetchImageBufferAxios(
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
