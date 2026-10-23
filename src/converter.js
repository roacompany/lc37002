import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import workerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";
import PptxGenJS from "pptxgenjs";
import { resolveFont, measureTextWidth } from "./fonts.js";
import { createOcrEngine, eraseText } from "./ocr.js";
import { devicePath, bboxOfPath, extractShapes, addShapes } from "./shapes.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

const ASSET_BASE = new URL(`${import.meta.env.BASE_URL}pdfjs/`, window.location.href).href;

/**
 * @typedef {object} ConvertOptions
 * @property {"editable"|"image"} mode  editable: 배경 이미지 + 편집 가능한 텍스트 / image: 페이지 전체를 이미지로
 * @property {number} scale  렌더링 배율 (1 = 72dpi)
 * @property {"png"|"jpeg"} imageFormat
 * @property {number[]|null} pages  변환할 페이지 번호(1부터). null이면 전체
 * @property {boolean} includeLinks  PDF 링크를 슬라이드 하이퍼링크로 옮길지 여부
 * @property {boolean} [shapes]  편집 가능 모드에서 단순한 벡터 도형을 PowerPoint 도형으로 바꿀지 여부
 * @property {"off"|"auto"|"all"} [ocr]  편집 가능 모드의 OCR: auto는 텍스트가 없는 페이지(스캔)만, all은 모든 페이지의 이미지 속 글자까지
 * @property {(info: {page: number, done: number, total: number, status?: string}) => void} [onProgress]
 * @property {AbortSignal} [signal]
 */

export async function openPdf(data) {
  return pdfjsLib.getDocument({
    data,
    cMapUrl: `${ASSET_BASE}cmaps/`,
    cMapPacked: true,
    // 폰트 원래 이름(굵게/기울임 판별, 폰트 매핑용)을 받기 위해 필요
    fontExtraProperties: true,
    standardFontDataUrl: `${ASSET_BASE}standard_fonts/`,
    wasmUrl: `${ASSET_BASE}wasm/`,
    iccUrl: `${ASSET_BASE}iccs/`,
  }).promise;
}

/**
 * PDF를 PPTX로 변환해 Blob으로 돌려준다.
 * @param {import("pdfjs-dist").PDFDocumentProxy} pdf
 * @param {ConvertOptions} options
 */
export async function convertPdfToPptx(pdf, options) {
  const { mode, scale, imageFormat, onProgress, signal } = options;
  const pageNumbers = options.pages ?? Array.from({ length: pdf.numPages }, (_, i) => i + 1);

  const pptx = new PptxGenJS();
  const first = (await pdf.getPage(pageNumbers[0])).getViewport({ scale: 1 });
  // PowerPoint 슬라이드 크기 제한: 1~56인치
  const layout = slideLayoutFor(first.width, first.height);
  pptx.defineLayout({ name: "PDF", width: layout.width, height: layout.height });
  pptx.layout = "PDF";

  const meta = await Promise.resolve().then(() => pdf.getMetadata()).catch(() => null);
  if (meta?.info?.Title) pptx.title = meta.info.Title;
  if (meta?.info?.Author) pptx.author = meta.info.Author;

  const ocrMode = mode === "editable" ? (options.ocr ?? "off") : "off";
  let ocrEngine = null;
  let done = 0;
  const total = pageNumbers.length;
  try {
    for (const pageNumber of pageNumbers) {
      if (signal?.aborted) throw new DOMException("변환이 취소되었습니다.", "AbortError");
      onProgress?.({ page: pageNumber, done, total });
      const page = await pdf.getPage(pageNumber);
      const slide = pptx.addSlide();
      const ocr =
        ocrMode === "off"
          ? null
          : {
              mode: ocrMode,
              // OCR 엔진은 처음 필요할 때 한 번만 불러온다
              getEngine: async () => {
                if (!ocrEngine) {
                  onProgress?.({ page: pageNumber, done, total, status: "OCR 엔진과 언어 데이터 불러오는 중…" });
                  ocrEngine = await createOcrEngine(["kor", "eng"]);
                }
                onProgress?.({ page: pageNumber, done, total, status: `${pageNumber}쪽 글자 인식(OCR) 중…` });
                return ocrEngine;
              },
            };
      await convertPage(page, slide, layout, { mode, scale, imageFormat, includeLinks: options.includeLinks, shapes: options.shapes ?? true, ocr });
      page.cleanup();
      done++;
      onProgress?.({ page: pageNumber, done, total });
    }
  } finally {
    await ocrEngine?.terminate();
  }

  return pptx.write({ outputType: "blob", compression: true });
}

function slideLayoutFor(widthPt, heightPt) {
  let w = widthPt / 72;
  let h = heightPt / 72;
  const max = Math.max(w, h);
  if (max > 56) {
    w *= 56 / max;
    h *= 56 / max;
  }
  const min = Math.min(w, h);
  if (min < 1) {
    w *= 1 / min;
    h *= 1 / min;
  }
  return { width: w, height: h };
}

async function convertPage(page, slide, layout, { mode, scale, imageFormat, includeLinks, shapes: convertShapes, ocr }) {
  const unit = page.getViewport({ scale: 1 });
  // 페이지 크기가 첫 페이지와 다르면 비율을 유지하며 슬라이드 중앙에 맞춘다
  const fit = Math.min(layout.width / (unit.width / 72), layout.height / (unit.height / 72));
  const ptToIn = fit / 72;
  const offX = (layout.width - (unit.width / 72) * fit) / 2;
  const offY = (layout.height - (unit.height / 72) * fit) / 2;
  const box = { x: offX, y: offY, w: (unit.width / 72) * fit, h: (unit.height / 72) * fit };

  const viewport = page.getViewport({ scale });

  if (mode === "image") {
    const { canvas } = await renderPage(page, viewport);
    slide.addImage({ data: canvasToDataUrl(canvas, imageFormat), ...box });
    releaseCanvas(canvas);
  } else {
    // 1단계: 페이지를 그리면서 각 글자·도형·이미지가 어디에, 어떤 순서와 색으로 그려지는지 기록
    const pass1 = await renderPage(page, viewport, { capture: true });
    const canvasSize = { width: pass1.canvas.width, height: pass1.canvas.height };
    releaseCanvas(pass1.canvas);

    const textContent = await page.getTextContent({ disableNormalization: false });
    const lines = buildTextLines(page, textContent, pass1.glyphs, unit, scale);
    const suppressText = new Set();
    for (const line of lines) for (const item of line.items) for (const g of item.glyphs) suppressText.add(g.index);

    const { shapes, suppress: suppressPaths } = convertShapes
      ? extractShapes(pass1.events, suppressText, canvasSize, scale)
      : { shapes: [], suppress: new Set() };

    // 2단계: 텍스트 상자·도형으로 옮긴 것만 빼고 다시 그려 배경 이미지를 만든다
    const suppress = { suppressText, suppressPaths };
    const pass2 = await renderPage(page, viewport, suppress);

    // 3단계(OCR): 남은 배경 이미지 속 글자를 인식해 텍스트로 바꾸고, 그 자리는 배경색으로 지운다
    let ocrLines = [];
    if (ocr && (ocr.mode === "all" || lines.length === 0)) {
      ocrLines = await recognizePage(page, unit, scale, suppress, pass2.canvas, await ocr.getEngine());
    }

    slide.addImage({ data: canvasToDataUrl(pass2.canvas, imageFormat), ...box });
    releaseCanvas(pass2.canvas);

    addShapes(slide, shapes, ptToIn, offX, offY);
    for (const line of [...lines, ...ocrLines]) addTextLine(slide, line, ptToIn, offX, offY);
  }

  if (includeLinks) await addLinks(page, slide, unit, ptToIn, offX, offY);
}

// OCR은 300dpi 근처에서 가장 정확하다. 너무 큰 페이지는 메모리 때문에 긴 변을 제한한다.
const OCR_DPI = 300;
const OCR_MAX_SIDE = 5000;

async function recognizePage(page, unit, scale, suppress, background, engine) {
  const ocrScale = Math.min(OCR_DPI / 72, OCR_MAX_SIDE / Math.max(unit.width, unit.height));
  // 배경과 같은 배율이면 그대로 쓰고, 아니면 OCR용 해상도로 다시 그린다(편집 가능한 텍스트는 제외)
  const source = Math.abs(ocrScale - scale) < 0.01 ? background : (await renderPage(page, page.getViewport({ scale: ocrScale }), suppress)).canvas;
  const found = await engine.recognize(source, ocrScale * 72);
  eraseText(background, found, scale / ocrScale);
  if (source !== background) releaseCanvas(source);

  return found.map((l) => {
    const width = (l.bbox.x1 - l.bbox.x0) / ocrScale;
    const fontSize = l.fontSize / ocrScale;
    return {
      x: l.bbox.x0 / ocrScale,
      y: l.baseline / ocrScale,
      width,
      fontSize,
      items: [
        {
          segments: l.runs.map((r) => ({ text: r.text, color: r.color, alpha: 1 })),
          font: resolveFont("", {}, "sans-serif", l.runs.map((r) => r.text).join("")),
          width,
          fontSize,
          ascent: l.hasCJK ? 0.88 : 0.9,
        },
      ],
    };
  });
}

function canvasToDataUrl(canvas, format) {
  return format === "png" ? canvas.toDataURL("image/png") : canvas.toDataURL("image/jpeg", 0.92);
}

function releaseCanvas(canvas) {
  canvas.width = 0;
  canvas.height = 0;
}

/**
 * 페이지를 캔버스에 렌더링한다.
 * capture: 대상 캔버스에 그려지는 글자(fillText)·경로(fill/stroke)·이미지 등을 그리는 순서대로 기록
 * suppressText / suppressPaths: 해당 순번의 fillText / fill·stroke 호출을 건너뜀 (편집 가능한 요소를 뺀 배경 만들기)
 */
async function renderPage(page, viewport, { capture = false, suppressText = null, suppressPaths = null } = {}) {
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const full = { x0: 0, y0: 0, x1: canvas.width, y1: canvas.height };

  const glyphs = [];
  const events = [];
  const proto = CanvasRenderingContext2D.prototype;
  const originals = new Map();
  let glyphIndex = 0;
  let pathIndex = 0;
  const clipStack = [];
  let clip = null;

  // pdf.js가 그룹·마스크용으로 쓰는 임시 캔버스는 건드리지 않고, 결과 캔버스에 그리는 호출만 가로챈다
  const patch = (name, fn) => {
    const original = proto[name];
    originals.set(name, original);
    proto[name] = function (...args) {
      if (this.canvas !== canvas) return original.apply(this, args);
      return fn.call(this, original, args);
    };
  };

  if (capture || suppressText || suppressPaths) {
    patch("fillText", function (original, args) {
      const i = glyphIndex++;
      if (capture) {
        const m = this.getTransform();
        const [, x, y] = args;
        const glyph = {
          index: i,
          x: m.a * x + m.c * y + m.e,
          y: m.b * x + m.d * y + m.f,
          color: typeof this.fillStyle === "string" ? this.fillStyle : null,
          alpha: this.globalAlpha,
        };
        glyphs.push(glyph);
        events.push({ kind: "glyph", index: i, bbox: glyphBox(this, glyph) });
      }
      if (suppressText?.has(i)) return;
      return original.apply(this, args);
    });
    for (const kind of ["fill", "stroke"]) {
      patch(kind, function (original, args) {
        const i = pathIndex++;
        if (capture) events.push(pathEvent(this, args, kind, i, clip, full));
        if (suppressPaths?.has(i)) return;
        return original.apply(this, args);
      });
    }
  }
  if (capture) {
    const other = (bboxOf) =>
      function (original, args) {
        events.push({ kind: "other", bbox: bboxOf(this, args) ?? full });
        return original.apply(this, args);
      };
    patch("drawImage", other(imageBox));
    patch("fillRect", other((ctx, [x, y, w, h]) => transformRect(ctx.getTransform(), x, y, w, h)));
    patch("strokeRect", other((ctx, [x, y, w, h]) => transformRect(ctx.getTransform(), x, y, w, h)));
    patch("putImageData", other((ctx, [img, x, y]) => ({ x0: x, y0: y, x1: x + img.width, y1: y + img.height })));
    patch("strokeText", other((ctx, [, x, y]) => glyphBox(ctx, devicePoint(ctx.getTransform(), x, y))));
    patch("save", function (original, args) {
      clipStack.push(clip);
      return original.apply(this, args);
    });
    patch("restore", function (original, args) {
      if (clipStack.length) clip = clipStack.pop();
      return original.apply(this, args);
    });
    patch("clip", function (original, args) {
      clip = intersectClip(clip, clipRegion(this, args, full));
      return original.apply(this, args);
    });
  }

  try {
    await page.render({
      canvas,
      viewport,
      annotationMode: pdfjsLib.AnnotationMode.ENABLE,
      background: "rgb(255,255,255)",
    }).promise;
  } finally {
    for (const [name, original] of originals) proto[name] = original;
  }
  return { canvas, glyphs, events };
}

function devicePoint(m, x, y) {
  return { x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f };
}

function transformRect(m, x, y, w, h) {
  const pts = [devicePoint(m, x, y), devicePoint(m, x + w, y), devicePoint(m, x, y + h), devicePoint(m, x + w, y + h)];
  return {
    x0: Math.min(...pts.map((p) => p.x)),
    y0: Math.min(...pts.map((p) => p.y)),
    x1: Math.max(...pts.map((p) => p.x)),
    y1: Math.max(...pts.map((p) => p.y)),
  };
}

function matrixScale(m) {
  return Math.sqrt(Math.abs(m.a * m.d - m.b * m.c));
}

/** 글자가 차지하는 대략적인 영역 (겹침 판단용) */
function glyphBox(ctx, origin) {
  const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(ctx.font)?.[1] ?? "10") * matrixScale(ctx.getTransform());
  return { x0: origin.x - px * 0.2, y0: origin.y - px, x1: origin.x + px, y1: origin.y + px * 0.3 };
}

function imageBox(ctx, args) {
  const [img] = args;
  let rect;
  if (args.length >= 9) rect = args.slice(5, 9);
  else if (args.length >= 5) rect = args.slice(1, 5);
  else rect = [args[1], args[2], img.width ?? img.videoWidth ?? 0, img.height ?? img.videoHeight ?? 0];
  return transformRect(ctx.getTransform(), ...rect);
}

function pathEvent(ctx, args, kind, index, clip, full) {
  const path = args[0] instanceof Path2D ? args[0] : null;
  const m = ctx.getTransform();
  const cmds = path ? devicePath(path, m) : null;
  const style = kind === "fill" ? ctx.fillStyle : ctx.strokeStyle;
  const color = typeof style === "string" ? parseCssColor(style) : null;
  // 그라데이션·패턴·혼합 모드·필터가 들어간 것은 도형으로 재현할 수 없어 이미지로 남긴다
  const plain = color && ctx.globalCompositeOperation === "source-over" && (!ctx.filter || ctx.filter === "none");
  const ev = {
    kind: "path",
    index,
    indices: [index],
    pathRef: path,
    cmds,
    bbox: cmds ? bboxOfPath(cmds) : full,
    clip,
    supported: !!(cmds && plain),
    fill: null,
    stroke: null,
  };
  const paint = color ? { color: color.hex, alpha: ctx.globalAlpha * color.alpha } : null;
  if (kind === "fill") ev.fill = paint;
  else {
    const s = matrixScale(m);
    ev.stroke = { ...paint, width: ctx.lineWidth * s, dash: ctx.getLineDash().map((d) => d * s) };
  }
  return ev;
}

function clipRegion(ctx, args, full) {
  const path = args[0] instanceof Path2D ? args[0] : null;
  const cmds = path ? devicePath(path, ctx.getTransform()) : null;
  if (!cmds) return { bbox: full, rect: false };
  const bbox = bboxOfPath(cmds);
  const corners = cmds.filter((c) => c[0] !== "Z");
  const rect =
    corners.length <= 5 &&
    corners.every((c) => c[0] !== "C" && c[0] !== "Q") &&
    corners.every((c) => (Math.abs(c[1] - bbox.x0) < 0.5 || Math.abs(c[1] - bbox.x1) < 0.5) && (Math.abs(c[2] - bbox.y0) < 0.5 || Math.abs(c[2] - bbox.y1) < 0.5));
  return { bbox, rect };
}

function intersectClip(a, b) {
  if (!a) return b;
  return {
    bbox: { x0: Math.max(a.bbox.x0, b.bbox.x0), y0: Math.max(a.bbox.y0, b.bbox.y0), x1: Math.min(a.bbox.x1, b.bbox.x1), y1: Math.min(a.bbox.y1, b.bbox.y1) },
    rect: a.rect && b.rect,
  };
}

/** 캔버스 색 문자열("#rrggbb" 또는 "rgba(r, g, b, a)")을 PowerPoint 색과 투명도로 */
function parseCssColor(css) {
  if (css.startsWith("#")) return { hex: parseColor(css), alpha: 1 };
  const m = css.match(/rgba?\(([^)]+)\)/);
  if (!m) return null;
  const parts = m[1].split(",").map((v) => parseFloat(v));
  return { hex: parseColor(css), alpha: parts.length > 3 ? parts[3] : 1 };
}

/**
 * getTextContent 결과(정확한 유니코드)와 렌더링 중 기록한 글자 위치/색을 결합해
 * 한 줄을 하나의 텍스트 상자(line)로, 그 안을 글꼴·색이 같은 구간(segment)으로 나눈다.
 * 좌표는 모두 1배율 뷰포트 기준 pt (왼쪽 위가 원점).
 */
function buildTextLines(page, textContent, glyphs, unit, scale) {
  const vt = unit.transform;
  const items = [];
  for (const item of textContent.items) {
    if (!("str" in item) || !item.str.trim()) continue;
    const [a, b, c, d, e, f] = pdfjsLib.Util.transform(vt, item.transform);
    // 회전·뒤집힌 텍스트는 이미지에 남겨둔다
    if (Math.abs(b) > 1e-3 || Math.abs(c) > 1e-3 || a <= 0 || d >= 0) continue;
    const fontSize = Math.abs(d);
    if (fontSize < 0.5) continue;
    items.push({
      str: item.str,
      x: e, // 기준선 왼쪽
      y: f,
      width: item.width,
      fontSize,
      fontName: item.fontName,
      hasEOL: item.hasEOL,
      glyphs: [],
    });
  }

  // 렌더링된 글자를 텍스트 항목에 배정 (글자 원점이 항목의 기준선 근처에 있는지)
  for (const g of glyphs) {
    const gx = g.x / scale;
    const gy = g.y / scale;
    let best = null;
    let bestDist = Infinity;
    for (const it of items) {
      const tol = it.fontSize * 0.25;
      if (gx < it.x - tol || gx > it.x + it.width + tol) continue;
      const dy = Math.abs(gy - it.y);
      if (dy > it.fontSize * 0.35) continue;
      if (dy < bestDist) {
        bestDist = dy;
        best = it;
      }
    }
    best?.glyphs.push({ ...g, x: gx });
  }

  const styles = textContent.styles;
  const lines = [];
  let cur = null;
  for (const it of items) {
    // 화면에 실제로 그려지지 않은 텍스트(숨김 텍스트, 경로로 그려진 글꼴 등)는 이미지에 맡긴다
    if (it.glyphs.length === 0) {
      cur = null;
      continue;
    }
    const fontObj = getFontObject(page, it.fontName);
    const style = styles[it.fontName] ?? {};
    it.font = resolveFont(fontObj?.name ?? "", fontObj ?? {}, style.fontFamily, it.str);
    it.ascent = style.ascent;
    it.segments = splitByColor(it);

    if (cur && canJoin(cur, it)) {
      const gap = it.x - (cur.x + cur.width);
      const last = cur.items.at(-1);
      if (gap > it.fontSize * 0.15 && !/\s$/.test(last.str) && !/^\s/.test(it.str)) {
        last.segments.at(-1).text += " ";
        last.trailingSpace = true;
      }
      cur.items.push(it);
      cur.width = it.x + it.width - cur.x;
    } else {
      cur = { x: it.x, y: it.y, width: it.width, fontSize: it.fontSize, items: [it] };
      lines.push(cur);
    }
    if (it.hasEOL) cur = null;
  }
  return lines;
}

/** 한 텍스트 항목 안에서 글자별 색이 다르면 색 구간으로 나눈다. */
function splitByColor(item) {
  const glyphs = [...item.glyphs].sort((p, q) => p.x - q.x);
  const colorOf = (g) => ({ color: parseColor(g.color), alpha: g.alpha });
  const chars = [...item.str];
  const visible = chars.filter((ch) => /\S/.test(ch)).length;

  if (visible !== glyphs.length) {
    // 글자 수와 그려진 글리프 수가 다르면(합자 등) 가장 많이 쓰인 색 하나로 처리
    const counts = new Map();
    for (const g of glyphs) {
      const key = `${g.color}|${g.alpha}`;
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    const [color, alpha] = [...counts.entries()].sort((p, q) => q[1] - p[1])[0][0].split("|");
    return [{ text: item.str, color: parseColor(color), alpha: Number(alpha) }];
  }

  const segments = [];
  let gi = 0;
  let style = colorOf(glyphs[0]);
  for (const ch of chars) {
    if (/\S/.test(ch)) style = colorOf(glyphs[gi++]);
    const last = segments.at(-1);
    if (last && last.color === style.color && last.alpha === style.alpha) last.text += ch;
    else if (last && /^\s$/.test(ch)) last.text += ch;
    else segments.push({ text: ch, ...style });
  }
  return segments;
}

function canJoin(line, it) {
  if (Math.abs(line.fontSize - it.fontSize) > Math.max(0.5, line.fontSize * 0.05)) return false;
  if (Math.abs(line.y - it.y) > line.fontSize * 0.1) return false;
  const gap = it.x - (line.x + line.width);
  return gap > -line.fontSize * 0.3 && gap < line.fontSize * 1.0;
}

function getFontObject(page, fontName) {
  try {
    if (page.commonObjs.has(fontName)) return page.commonObjs.get(fontName);
  } catch {
    /* 폰트 정보가 없으면 기본값 사용 */
  }
  return null;
}

function parseColor(css) {
  if (!css) return "000000";
  if (css.startsWith("#")) {
    let hex = css.slice(1);
    if (hex.length === 3) hex = [...hex].map((ch) => ch + ch).join("");
    return hex.slice(0, 6).toUpperCase();
  }
  const m = css.match(/rgba?\(([^)]+)\)/);
  if (m) {
    const [r, g, b] = m[1].split(",").map((v) => parseFloat(v));
    return [r, g, b].map((v) => Math.round(v).toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  return "000000";
}

function addTextLine(slide, line, ptToIn, offX, offY) {
  const size = line.fontSize;
  const first = line.items[0];
  const ascent = first.ascent && first.ascent > 0.5 && first.ascent < 1.2 ? first.ascent : 0.9;
  const sizePt = round(size * ptToIn * 72, 2);

  const runs = [];
  for (const it of line.items) {
    // PowerPoint에서 원본과 같은 폭이 되도록 자간 보정 (대체 폰트일 때 줄이 넘치거나 비는 것을 방지)
    let charSpacing;
    const text = it.segments.map((s) => s.text).join("");
    const measured = measureTextWidth(it.trailingSpace ? text.slice(0, -1) : text, it.font, it.fontSize);
    const count = [...text].length - (it.trailingSpace ? 1 : 0);
    if (measured && count > 1) {
      const diff = (it.width - measured) / (count - 1);
      if (Math.abs(diff) > it.fontSize * 0.01) charSpacing = Math.max(-it.fontSize * 0.3, Math.min(it.fontSize * 0.5, diff));
    }
    for (const seg of it.segments) {
      const options = {
        fontFace: it.font.family,
        fontSize: sizePt,
        bold: it.font.bold,
        italic: it.font.italic,
        color: seg.color,
      };
      if (charSpacing !== undefined) options.charSpacing = round(charSpacing * ptToIn * 72, 2);
      if (seg.alpha < 1) options.transparency = Math.round((1 - seg.alpha) * 100);
      runs.push({ text: seg.text, options });
    }
  }

  slide.addText(runs, {
    x: offX + line.x * ptToIn,
    y: offY + (line.y - size * ascent) * ptToIn,
    w: Math.max(line.width * 1.02, size * 0.5) * ptToIn,
    h: size * 1.25 * ptToIn,
    margin: 0,
    valign: "top",
    align: "left",
    wrap: false,
    fit: "none",
    lineSpacingMultiple: 1,
  });
}

async function addLinks(page, slide, unit, ptToIn, offX, offY) {
  const annotations = await page.getAnnotations().catch(() => []);
  for (const annot of annotations) {
    if (annot.subtype !== "Link" || !annot.url) continue;
    const [x1, y1, x2, y2] = pdfjsLib.Util.normalizeRect(annot.rect);
    const [vx1, vy1] = unit.convertToViewportPoint(x1, y2);
    const [vx2, vy2] = unit.convertToViewportPoint(x2, y1);
    slide.addText("", {
      x: offX + Math.min(vx1, vx2) * ptToIn,
      y: offY + Math.min(vy1, vy2) * ptToIn,
      w: Math.abs(vx2 - vx1) * ptToIn,
      h: Math.abs(vy2 - vy1) * ptToIn,
      hyperlink: { url: annot.url },
    });
  }
}

function round(v, d) {
  const p = 10 ** d;
  return Math.round(v * p) / p;
}
