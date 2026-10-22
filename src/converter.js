import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import workerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";
import PptxGenJS from "pptxgenjs";
import { resolveFont, measureTextWidth } from "./fonts.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

const ASSET_BASE = new URL(`${import.meta.env.BASE_URL}pdfjs/`, window.location.href).href;

/**
 * @typedef {object} ConvertOptions
 * @property {"editable"|"image"} mode  editable: 배경 이미지 + 편집 가능한 텍스트 / image: 페이지 전체를 이미지로
 * @property {number} scale  렌더링 배율 (1 = 72dpi)
 * @property {"png"|"jpeg"} imageFormat
 * @property {number[]|null} pages  변환할 페이지 번호(1부터). null이면 전체
 * @property {boolean} includeLinks  PDF 링크를 슬라이드 하이퍼링크로 옮길지 여부
 * @property {(info: {page: number, done: number, total: number}) => void} [onProgress]
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

  let done = 0;
  for (const pageNumber of pageNumbers) {
    if (signal?.aborted) throw new DOMException("변환이 취소되었습니다.", "AbortError");
    onProgress?.({ page: pageNumber, done, total: pageNumbers.length });
    const page = await pdf.getPage(pageNumber);
    const slide = pptx.addSlide();
    await convertPage(page, slide, layout, { mode, scale, imageFormat, includeLinks: options.includeLinks });
    page.cleanup();
    done++;
    onProgress?.({ page: pageNumber, done, total: pageNumbers.length });
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

async function convertPage(page, slide, layout, { mode, scale, imageFormat, includeLinks }) {
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
    // 1단계: 페이지를 그리면서 각 글자가 어디에, 어떤 색으로 그려지는지 기록
    const pass1 = await renderPage(page, viewport, { capture: true });
    releaseCanvas(pass1.canvas);

    const textContent = await page.getTextContent({ disableNormalization: false });
    const lines = buildTextLines(page, textContent, pass1.glyphs, unit, scale);

    // 2단계: 편집 가능한 텍스트로 옮길 글자만 빼고 다시 그려 배경 이미지를 만든다
    const suppress = new Set();
    for (const line of lines) for (const item of line.items) for (const g of item.glyphs) suppress.add(g.index);
    const pass2 = await renderPage(page, viewport, { suppress });
    slide.addImage({ data: canvasToDataUrl(pass2.canvas, imageFormat), ...box });
    releaseCanvas(pass2.canvas);

    for (const line of lines) addTextLine(slide, line, ptToIn, offX, offY);
  }

  if (includeLinks) await addLinks(page, slide, unit, ptToIn, offX, offY);
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
 * capture: 대상 캔버스에 그려지는 fillText 호출(글자)을 순서대로 기록
 * suppress: 해당 순번의 fillText 호출을 건너뜀 (텍스트 없는 배경 만들기)
 */
async function renderPage(page, viewport, { capture = false, suppress = null } = {}) {
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);

  const glyphs = [];
  const proto = CanvasRenderingContext2D.prototype;
  const originalFillText = proto.fillText;
  let index = 0;

  if (capture || suppress) {
    proto.fillText = function (text, x, y, maxWidth) {
      if (this.canvas !== canvas) return originalFillText.apply(this, arguments);
      const i = index++;
      if (capture) {
        const m = this.getTransform();
        glyphs.push({
          index: i,
          x: m.a * x + m.c * y + m.e,
          y: m.b * x + m.d * y + m.f,
          color: typeof this.fillStyle === "string" ? this.fillStyle : null,
          alpha: this.globalAlpha,
        });
      }
      if (suppress?.has(i)) return;
      return originalFillText.apply(this, arguments);
    };
  }

  try {
    await page.render({
      canvas,
      viewport,
      annotationMode: pdfjsLib.AnnotationMode.ENABLE,
      background: "rgb(255,255,255)",
    }).promise;
  } finally {
    proto.fillText = originalFillText;
  }
  return { canvas, glyphs };
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
