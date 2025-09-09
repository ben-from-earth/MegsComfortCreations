"use strict";

/**
 * outputPNG({ template, images })
 *
 * Input:
 * {
 *   template: 3 | 5,
 *   images: Array<{ url: string, spine_color: string, type: string }>
 * }
 *
 * Behavior:
 * - Fixed canvas: 2775w × 2025h
 * - template=5 → rowHeight=324, yGap=26
 *   template=3 → rowHeight=261, yGap=18
 * - For each (template,type) combo, use a specific cell width (px) and inner pair gap (px).
 * - Cells flow left→right and wrap to a new row if the next cell would overflow the canvas width.
 * - Per cell:
 *    • Background = spine_color
 *    • Pair cells: url on left (rotated 180°) + url on right (normal) with combo’s pairGap
 *    • Album cells: single centered url (pairGap=0)
 */

const sharp = require("sharp");
const axios = require("axios");

/* ===========================
 * CONSTANTS
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

/**
 * Explicit per-combo cell widths & inner pair gaps (pixels).
 * Tweak these numbers to get your exact design.
 *
 * Variants (as you defined):
 *  1) template=3, type in {book, movie, video_game}
 *  2) template=3, type === album
 *  3) template=5, type === books
 *  4) template=5, type in {movie, video_game}
 *  5) template=5, type === album
 *
 * Tip: these defaults line up with your earlier 489×324 and 261→~394 aspect.
 */

const variation1 = {
  cellWidth: 387,
  imageWidth: 174,
  pairGap: 39,
  mode: "pair",
};
const variation2 = {
  cellWidth: 261,
  imageWidth: 261,
  pairGap: 0,
  mode: "single",
};
const variation3 = {
  cellWidth: 489,
  imageWidth: 216,
  pairGap: 57,
  mode: "pair",
};
const variation4 = {
  cellWidth: 471,
  imageWidth: 216,
  pairGap: 39,
  mode: "pair",
};
const variation5 = {
  cellWidth: 324,
  imageWidth: 324,
  pairGap: 0,
  mode: "single",
};
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

async function outputPNG({ template, images = [] }) {
  const metrics = TEMPLATE_METRICS[template];
  if (!metrics) throw new Error("Unsupported template (use 3 or 5)");

  // Base canvas
  const base = sharp({
    create: {
      width: CANVAS_W,
      height: CANVAS_H,
      channels: 4,
      background: CANVAS_BG,
    },
  });

  // Compute slot rectangles (x,y,w,h) for each image (variable widths)
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

  // Render overlays for each slot
  const overlaysNested = await Promise.all(
    slots.map((slot, i) =>
      renderCell(slot, images[i], template, metrics.rowHeight)
    )
  );
  const overlays = overlaysNested.flat();

  // Composite
  try {
    return await base.composite(overlays).png().toBuffer();
  } catch (err) {
    if (err && Array.isArray(err.errors)) {
      for (const e of err.errors)
        console.error("composite overlay error", e?.message || e);
    }
    console.error("sharp composite error", err?.message || err);
    throw err;
  }
}

module.exports = outputPNG;

/* ===========================
 * LAYOUT
 * =========================== */

/**
 * Build slots using **explicit cell widths** per (template,type).
 * Flow left→right with wrap when adding the next cell would overflow canvas width.
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
    const cfg = COMBOS[template][type];
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
 * RENDERERS
 * =========================== */

/** Render one cell according to the combo mode ("pair" or "single"). */
async function renderCell(slot, img, template, rowHeight) {
  const { url, spine_color, type } = img || {};
  //{x,y,w,h} = slot
  //template 3 or 5
  const cfg = COMBOS[template][type];

  // Common cell background (spine color fills full slot)
  const cellBg = await makeBlock(
    slot.w,
    slot.h,
    spine_color ?? { r: 255, g: 0, b: 0, alpha: 1 }
  );

  if (cfg.mode === "single") {
    // Single centered cover
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

  // Pair: left rotated 180°, right normal, combo-specific inner gap
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

/* ===========================
 * IMAGE HELPERS
 * =========================== */

function isHttpUrl(s) {
  try {
    const u = new URL(String(s));
    return u.protocol === "http:" || u.protocol === "https:";
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
  const base = String(ctype).split(";")[0].trim().toLowerCase();
  return base.startsWith("image/");
}

async function sniffIsImage(buffer) {
  const { fileTypeFromBuffer } = await import("file-type");
  // Detect from bytes (handles jpg/png/webp/avif/gif/ico/svg-as-xml, etc.)
  const ft = await fileTypeFromBuffer(buffer).catch(() => null);
  if (ft?.mime?.startsWith("image/"))
    return { ok: true, mime: ft.mime, ext: ft.ext };
  // very light fallback for SVG served as text/xml or text/plain
  const head = buffer.slice(0, 256).toString("utf8");
  if (/\<svg[\s>]/i.test(head))
    return { ok: true, mime: "image/svg+xml", ext: "svg" };
  return { ok: false };
}

/**
 * Fetch an image as Buffer, with retries and MIME sniffing.
 */
async function fetchImageBufferAxios(
  url,
  {
    maxBytes = 8 * 1024 * 1024,
    timeout = 15000,
    retries = 3,
    backoff = 500,
    requireImage = true, // set false if you want to accept any binary
    sendReferer = true,
  } = {}
) {
  let attempt = 0;
  const referer = sendReferer ? getReferer(url) : undefined;

  while (true) {
    try {
      const res = await axios.get(url, {
        responseType: "arraybuffer",
        maxContentLength: maxBytes, // Axios v0.x
        maxBodyLength: maxBytes, // Axios v1.x
        timeout,
        maxRedirects: 3,
        headers: {
          "User-Agent": "png-renderer/1.0 (+node)",
          Accept: "image/avif,image/webp,image/*,*/*;q=0.8",
          ...(referer ? { Referer: referer } : {}),
        },
        // Accept only 2xx after redirects; treat 3xx/4xx/etc. as failures
        validateStatus: (s) => s >= 200 && s < 300,
      });

      const buffer = Buffer.from(res.data);
      const ctype = res.headers["content-type"];

      // Fast path: header says image/*
      if (isImageCtype(ctype)) return buffer;

      // If header is missing or octet-stream, sniff the bytes
      if (requireImage) {
        const sniff = await sniffIsImage(buffer);
        if (sniff.ok) return buffer;
        throw new Error(`Not an image. Content-Type=${ctype || "unknown"}`);
      }

      // If you don't require image verification, just return
      return buffer;
    } catch (err) {
      attempt++;
      const status = err.response?.status;
      const ctype = err.response?.headers?.["content-type"];
      const urlLogged = err.config?.url;
      console.error(
        `[fetchImageBufferAxios] attempt=${attempt} | message=${err.message}` +
          (err.code ? ` | code=${err.code}` : "") +
          (status ? ` | status=${status}` : "") +
          (ctype ? ` | ctype=${ctype}` : "") +
          (urlLogged ? ` | url=${urlLogged}` : "")
      );

      // Give up if out of retries
      if (attempt > retries) {
        throw new Error(
          `IMAGE_FETCH_FAILED after ${retries} retries: ${err.code || ""} ${
            err.message
          }`
        );
      }

      // Backoff with jitter
      const delay = backoff * Math.pow(2, attempt - 1) + Math.random() * 150;
      await new Promise((r) => setTimeout(r, delay));
    }
  }
}

/**
 * Create a block buffer:
 * - If `fill` is an http(s) URL → fetch & fit (cover) to w×h
 * - Else treat `fill` as a color (CSS or {r,g,b,alpha})
 * - rotation: if 180 → flip+flop (exact 180°); else → rotate(angle)
 */
async function makeBlock(w, h, fill, rotation = 0) {
  let pipe;
  if (typeof fill === "string" && isHttpUrl(fill)) {
    try {
      const buf = await fetchImageBufferAxios(fill);
      pipe = sharp(buf).toColorspace("srgb").resize(w, h, { fit: "cover" });
    } catch (e) {
      console.warn("[makeBlock] fetch failed, using placeholder:", e.message);
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
