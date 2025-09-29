'use strict';

/**
 * PNG/ZIP renderer (paginated)
 *
 * Exports:
 *   - outputAuto({ template, images, prefix? }) → { mime, filename, buffer }
 *       If one page fits → returns a single PNG.
 *       If multiple pages → returns a ZIP of PNG pages.
 *   - outputPNGs({ template, images, namePrefix? }) → [{ filename, buffer }]
 *   - outputZIP({ template, images, prefix? }) → Buffer (zip)
 *   - outputPNG({ template, images }) → Buffer (single page only)
 *
 * Input:
 * {
 *   template: 3 | 5,
 *   images: Array<{ url: string, spine_color: string, type: string }>
 * }
 *
 * Layout rules (unchanged):
 * - Fixed canvas: 2025w x 2775h
 * - template=5 → rowHeight=324, yGap=26
 *   template=3 → rowHeight=261, yGap=18
 * - Per (template,type), specific cell width and inner pair gap.
 * - Flow left→right, wrap to a new row when the next cell would overflow.
 * - Cell:
 *    • Background = spine_color
 *    • Pair cells: left url rotated 180° + right url normal, with pairGap
 *    • Album cells: single centered url (pairGap=0)
 */

const sharp = require('sharp');
const axios = require('axios');
const JSZip = require('jszip');

/* ===========================
 * CONSTANTS & COMBOS
 * =========================== */

// Fixed canvas (portrait would be CANVAS_W=2025, CANVAS_H=2775)
const CANVAS_W = 2025;
const CANVAS_H = 2775;

// Vertical row metrics by template
const TEMPLATE_METRICS = {
  5: { rowHeight: 324, yGap: 26, xGap: 17 },
  3: { rowHeight: 261, yGap: 18, xGap: 22 },
};

// Outer margins & horizontal gutters
const MARGIN_X = 0;
const MARGIN_Y = 0;

// Canvas background (transparent)
const CANVAS_BG = { r: 0, g: 0, b: 0, alpha: 0 };

// Cell layout variations
const variation1 = {
  cellWidth: 387,
  imageWidth: 174,
  pairGap: 39,
  mode: 'pair',
};
const variation2 = {
  cellWidth: 261,
  imageWidth: 261,
  pairGap: 0,
  mode: 'single',
};
const variation3 = {
  cellWidth: 489,
  imageWidth: 216,
  pairGap: 57,
  mode: 'pair',
};
const variation4 = {
  cellWidth: 471,
  imageWidth: 216,
  pairGap: 39,
  mode: 'pair',
};
const variation5 = {
  cellWidth: 324,
  imageWidth: 324,
  pairGap: 0,
  mode: 'single',
};

// Which variation is used based on format and media type
const COMBOS = {
  3: {
    book: variation1,
    movie: variation1,
    video_game: variation1,
    album: variation2,
    default: variation1,
  },
  5: {
    book: variation3,
    movie: variation4,
    video_game: variation4,
    album: variation5,
    default: variation4,
  },
};

/* ===========================
 * CORE (single page)
 * =========================== */

/** Original single-canvas renderer (kept for compatibility). */
async function outputPNG({ template, images = [] }) {
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
      renderCell(slot, images[i], template, metrics.rowHeight)
    )
  );
  const overlays = overlaysNested.flat();

  try {
    return await base.composite(overlays).png().toBuffer();
  } catch (err) {
    if (err && Array.isArray(err.errors)) {
      for (const e of err.errors)
        console.error('composite overlay error', e?.message || e);
    }
    console.error('sharp composite error', err?.message || err);
    throw err;
  }
}

/**
 * Build slots (single canvas) using explicit cell widths per (template,type).
 * Flow left→right; wrap when adding the next cell would overflow canvas width.
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
}) {
  const slots = [];
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

/* ===========================
 * PAGINATION (multi-page)
 * =========================== */

/** Layout a single page starting at startIndex; return slots and countUsed. */
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
}) {
  const slots = [];
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

/** Render a single paginated canvas to a PNG buffer. */
async function renderPage({ template, metrics, images, startIndex }) {
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
      renderCell(slot, images[startIndex + k], template, metrics.rowHeight)
    )
  );
  const overlays = overlaysNested.flat();

  const buffer = await base.composite(overlays).png().toBuffer();
  return { buffer, countUsed };
}

/** Return an array of PNG pages (filename + buffer). */
async function outputPNGs({ template, images = [], namePrefix = 'grid' }) {
  const metrics = TEMPLATE_METRICS[template];
  if (!metrics) throw new Error('Unsupported template (use 3 or 5)');

  const results = [];
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

/** Return a ZIP buffer containing all PNG pages. */
async function outputZIP({ template, images = [], prefix = 'grid' }) {
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
 * Entry point for your route: returns either PNG or ZIP based on how many pages fit.
 * - If everything fits on one page → { mime: 'image/png', filename, buffer }
 * - Else → { mime: 'application/zip', filename: '<prefix>_pages.zip', buffer }
 */
async function outputAuto({ template, images = [], prefix = 'grid' }) {
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

/* ===========================
 * RENDERERS & IMAGE HELPERS
 * =========================== */

// Render one cell according to combo mode ("pair" | "single").
async function renderCell(slot, img, template, rowHeight) {
  const { url, spine_color, type } = img || {};
  const cfg = COMBOS[template][type] ?? COMBOS[template].default;

  // Cell background (spine color fills full slot)
  const cellBg = await makeBlock(
    slot.w,
    slot.h,
    spine_color ?? { r: 255, g: 0, b: 0, alpha: 1 }
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

function isHttpUrl(s) {
  try {
    const u = new URL(String(s));
    return u.protocol === 'http:' || u.protocol === 'https:';
  } catch {
    return false;
  }
}

function getReferer(url) {
  try {
    const u = new URL(url);
    return `${u.protocol}//${u.host}/`;
  } catch {
    return undefined;
  }
}

function isImageCtype(ctype) {
  if (!ctype) return false;
  const base = String(ctype).split(';')[0].trim().toLowerCase();
  return base.startsWith('image/');
}

async function sniffIsImage(buffer) {
  const { fileTypeFromBuffer } = await import('file-type');
  const ft = await fileTypeFromBuffer(buffer).catch(() => null);
  if (ft?.mime?.startsWith('image/'))
    return { ok: true, mime: ft.mime, ext: ft.ext };
  const head = buffer.slice(0, 256).toString('utf8');
  if (/\<svg[\s>]/i.test(head))
    return { ok: true, mime: 'image/svg+xml', ext: 'svg' };
  return { ok: false };
}

// Fetch an image as Buffer, with retries and MIME sniffing.
async function fetchImageBufferAxios(
  url,
  {
    maxBytes = 8 * 1024 * 1024,
    timeout = 15000,
    retries = 3,
    backoff = 500,
    requireImage = true,
    sendReferer = true,
  } = {}
) {
  let attempt = 0;
  const referer = sendReferer ? getReferer(url) : undefined;

  while (true) {
    try {
      const res = await axios.get(url, {
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
      const ctype = res.headers['content-type'];

      if (isImageCtype(ctype)) return buffer;

      if (requireImage) {
        const sniff = await sniffIsImage(buffer);
        if (sniff.ok) return buffer;
        throw new Error(`Not an image. Content-Type=${ctype || 'unknown'}`);
      }

      return buffer;
    } catch (err) {
      attempt++;
      const status = err.response?.status;
      const ctype = err.response?.headers?.['content-type'];
      const urlLogged = err.config?.url;
      console.error(
        `[fetchImageBufferAxios] attempt=${attempt} | message=${err.message}` +
          (err.code ? ` | code=${err.code}` : '') +
          (status ? ` | status=${status}` : '') +
          (ctype ? ` | ctype=${ctype}` : '') +
          (urlLogged ? ` | url=${urlLogged}` : '')
      );

      if (attempt > retries) {
        throw new Error(
          `IMAGE_FETCH_FAILED after ${retries} retries: ${err.code || ''} ${
            err.message
          }`
        );
      }

      const delay = backoff * Math.pow(2, attempt - 1) + Math.random() * 150;
      await new Promise((r) => setTimeout(r, delay));
    }
  }
}

/**
 * Create a block buffer:
 * - If `fill` is an http(s) URL → fetch & fit to w×h (no letterboxing)
 * - Else treat `fill` as a color (CSS or {r,g,b,alpha})
 * - rotation: if 180 → flip+flop (exact 180°)
 */
async function makeBlock(w, h, fill, rotation = 0) {
  let pipe;
  if (typeof fill === 'string' && isHttpUrl(fill)) {
    try {
      const buf = await fetchImageBufferAxios(fill);
      pipe = sharp(buf).toColourspace('srgb').resize(w, h, { fit: 'fill' });
    } catch (e) {
      console.warn('[makeBlock] fetch failed, using placeholder:', e.message);
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
        background: fill ?? { r: 180, g: 180, b: 180, alpha: 1 },
      },
    });
  }

  if (rotation === 180) pipe = pipe.flip().flop();
  return await pipe.png().toBuffer();
}

/* ===========================
 * EXPORTS
 * =========================== */

module.exports = {
  outputAuto, // ← use this in your route
  outputPNGs,
  outputZIP,
  outputPNG, // kept for any legacy single-page callers
};
