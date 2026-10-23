// 스캔·이미지 속 글자를 tesseract.js(WASM)로 인식한다. 모델과 엔진은 같은 사이트에서 받아오므로
// 이미지가 외부로 전송되지 않는다. 첫 사용 시 내려받은 학습 데이터는 브라우저(IndexedDB)에 캐시된다.

const BASE = new URL(`${import.meta.env.BASE_URL}tesseract/`, window.location.href).href;

// 평균 인식률이 이보다 낮은 구간은 잘못 읽었을 가능성이 커서 텍스트로 바꾸지 않고 이미지에 남긴다
const MIN_CONFIDENCE = 45;
const CJK = /[\u1100-\u11ff\u3130-\u318f\uac00-\ud7af\u3040-\u30ff\u4e00-\u9fff]/;
// 표 테두리·세로줄을 글자로 잘못 읽은 경우
const BORDER_WORD = /^[|[\]ㅣ¦]+$/;

export async function createOcrEngine(languages = ["kor", "eng"], onStatus) {
  const { createWorker, OEM, PSM } = await import("tesseract.js");
  const worker = await createWorker(languages, OEM.LSTM_ONLY, {
    workerPath: `${BASE}worker.min.js`,
    corePath: `${BASE}core/`,
    langPath: `${BASE}lang/`,
    gzip: true,
    logger: (m) => onStatus?.(m),
  });
  // 슬라이드는 제목·표·도형 안 글자처럼 흩어진 텍스트가 많아 SPARSE_TEXT가 자동 레이아웃보다 훨씬 많이 찾는다
  await worker.setParameters({ tessedit_pageseg_mode: PSM.SPARSE_TEXT, preserve_interword_spaces: "1" });
  return {
    /**
     * @param {HTMLCanvasElement} canvas
     * @param {number} dpi  캔버스 해상도 (글자 크기 판단에 사용)
     * @returns {Promise<OcrLine[]>} 캔버스 픽셀 좌표의 줄 목록
     */
    async recognize(canvas, dpi) {
      await worker.setParameters({ user_defined_dpi: String(Math.round(dpi)) });
      // 회색조로 바꿔 넣으면 색 글자·색 배경에서 인식률이 확실히 좋아진다
      const gray = toGrayscale(canvas);
      const { data } = await worker.recognize(gray, {}, { blocks: true });
      gray.width = gray.height = 0;
      return extractLines(data, canvas);
    },
    terminate: () => worker.terminate(),
  };
}

function toGrayscale(canvas) {
  const out = document.createElement("canvas");
  out.width = canvas.width;
  out.height = canvas.height;
  const ctx = out.getContext("2d");
  ctx.drawImage(canvas, 0, 0);
  const img = ctx.getImageData(0, 0, out.width, out.height);
  const d = img.data;
  for (let i = 0; i < d.length; i += 4) {
    const y = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
    d[i] = d[i + 1] = d[i + 2] = y;
  }
  ctx.putImageData(img, 0, 0);
  return out;
}

/**
 * @typedef {object} OcrLine
 * @property {{x0:number,y0:number,x1:number,y1:number}} bbox
 * @property {number} baseline  기준선 y (px)
 * @property {number} fontSize  글자 크기 추정치 (px)
 * @property {{text: string, color: string}[]} runs  색이 같은 구간들 ("RRGGBB")
 * @property {{bbox: object, bg: number[]}} erase  지울 영역과 그 자리의 배경색
 * @property {boolean} hasCJK
 */
function extractLines(data, canvas) {
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const lines = [];
  for (const block of data.blocks ?? []) {
    for (const para of block.paragraphs) {
      for (const line of para.lines) {
        for (const words of splitLine(line)) {
          const conf = words.reduce((sum, w) => sum + w.confidence, 0) / words.length;
          const text = words.map((w) => w.text).join("");
          if (conf < MIN_CONFIDENCE || !/[\p{L}\p{N}]/u.test(text)) continue;
          lines.push(buildLine(ctx, words));
        }
      }
    }
  }
  return lines;
}

/** 표 테두리로 읽힌 기호나 큰 간격이 있으면 한 줄을 여러 텍스트 상자로 나눈다 (표 칸마다 위치 유지). */
function splitLine(line) {
  const parts = [];
  let cur = [];
  // 띄어쓰기는 tesseract가 만든 줄 텍스트를 따른다 (한글은 음절 사이 간격이 실제 띄어쓰기만큼 벌어지기도 해서
  // 좌표로는 판단이 어렵다)
  let cursor = 0;
  for (const w of line.words) {
    const at = line.text.indexOf(w.text, cursor);
    const spaceBefore = at >= 0 ? /\s/.test(line.text.slice(cursor, at)) : true;
    if (at >= 0) cursor = at + w.text.length;
    const text = w.text.trim();
    if (!text) continue;
    if (BORDER_WORD.test(text)) {
      if (cur.length) parts.push(cur);
      cur = [];
      continue;
    }
    // 단어 앞뒤에 붙은 테두리 기호는 떼어내고 그 자리에서 상자를 나눈다 ("[1,234" → | 1,234)
    const lead = /^[|[\]ㅣ]+/.test(text);
    const trail = /[|[\]ㅣ]+$/.test(text);
    const cleaned = text.replace(/^[|[\]ㅣ]+|[|[\]ㅣ]+$/g, "");
    const word = { ...w, text: cleaned || text, bbox: lead || trail ? trimBorderSymbols(w) : w.bbox, spaceBefore };
    const prev = cur.at(-1);
    if (prev) {
      const h = Math.max(prev.bbox.y1 - prev.bbox.y0, w.bbox.y1 - w.bbox.y0);
      if (lead || word.bbox.x0 - prev.bbox.x1 > h * 1.5) {
        parts.push(cur);
        cur = [];
      }
    }
    cur.push(word);
    if (trail) {
      parts.push(cur);
      cur = [];
    }
  }
  if (cur.length) parts.push(cur);
  return parts;
}

/** 테두리 기호를 뺀 글자들만의 영역 (글자별 상자가 없으면 원래 단어 상자) */
function trimBorderSymbols(word) {
  const symbols = (word.symbols ?? []).filter((sym) => !BORDER_WORD.test(sym.text.trim()));
  if (!symbols.length) return word.bbox;
  return {
    x0: Math.min(...symbols.map((sym) => sym.bbox.x0)),
    y0: Math.min(...symbols.map((sym) => sym.bbox.y0)),
    x1: Math.max(...symbols.map((sym) => sym.bbox.x1)),
    y1: Math.max(...symbols.map((sym) => sym.bbox.y1)),
  };
}

function buildLine(ctx, words) {
  const bbox = {
    x0: Math.min(...words.map((w) => w.bbox.x0)),
    y0: Math.min(...words.map((w) => w.bbox.y0)),
    x1: Math.max(...words.map((w) => w.bbox.x1)),
    y1: Math.max(...words.map((w) => w.bbox.y1)),
  };
  const metrics = words.map(wordMetrics);
  const fontSize = median(metrics.map((m) => m.size));
  const baseline = median(metrics.map((m) => m.baseline));

  // 단어별 글자색을 구하고, 비슷한 색끼리 이어 붙여 구간(run)으로 만든다
  const sampled = words.map((w) => sampleWord(ctx, w.bbox));
  const runs = [];
  let prevColor = null;
  words.forEach((w, i) => {
    const color = sampled[i].fg ?? prevColor ?? [0, 0, 0];
    const space = i > 0 && w.spaceBefore ? " " : "";
    const last = runs.at(-1);
    if (last && dist(last.rgb, color) < 60) last.text += space + w.text;
    else {
      if (last) last.text += space;
      runs.push({ text: w.text, rgb: color });
    }
    prevColor = color;
  });

  return {
    bbox,
    baseline,
    fontSize,
    runs: runs.map((r) => ({ text: r.text, color: toHex(r.rgb) })),
    // 한글 단어 상자는 부정확하게 겹치거나 비는 경우가 있어, 지울 때는 구간 전체를 하나로 칠한다
    erase: { bbox, bg: sampleWord(ctx, bbox).bg },
    hasCJK: CJK.test(words.map((w) => w.text).join("")),
  };
}

/**
 * 단어 상자 높이로 글자 크기(em)와 기준선을 추정한다. 실측 비율:
 * 한글 음절 높이 ≈ 0.9em(기준선 아래로 약 0.1em), 라틴 대문자·어센더 ≈ 0.73em, 디센더 ≈ 0.2em, 소문자 x높이 ≈ 0.53em.
 */
function wordMetrics(word) {
  const h = word.bbox.y1 - word.bbox.y0;
  const t = word.text;
  if (CJK.test(t)) return { size: h / 0.9, baseline: word.bbox.y1 - (h / 0.9) * 0.1 };
  const tall = /[A-Zbdfhklt0-9()[\]{}/\\|!?'"#$%&@]/.test(t);
  const desc = /[gjpqy,;()[\]{}|]/.test(t);
  const ratio = tall && desc ? 0.93 : tall ? 0.73 : desc ? 0.73 : 0.53;
  const size = h / ratio;
  return { size, baseline: word.bbox.y1 - (desc ? size * 0.2 : 0) };
}

/** 단어 영역 테두리에서 배경색을, 배경과 가장 다른 픽셀들에서 글자색을 구한다. */
function sampleWord(ctx, bbox) {
  const pad = Math.max(2, Math.round((bbox.y1 - bbox.y0) * 0.15));
  const x = Math.max(0, bbox.x0 - pad);
  const y = Math.max(0, bbox.y0 - pad);
  const w = Math.min(ctx.canvas.width - x, bbox.x1 - bbox.x0 + pad * 2);
  const h = Math.min(ctx.canvas.height - y, bbox.y1 - bbox.y0 + pad * 2);
  if (w <= 0 || h <= 0) return { bg: [255, 255, 255], fg: null };
  const { data } = ctx.getImageData(x, y, w, h);

  const border = [];
  for (let i = 0; i < w; i++) border.push(px(data, w, i, 0), px(data, w, i, h - 1));
  for (let j = 0; j < h; j++) border.push(px(data, w, 0, j), px(data, w, w - 1, j));
  const bg = [0, 1, 2].map((c) => median(border.map((p) => p[c])));

  let maxDist = 0;
  const pixels = [];
  for (let j = pad; j < h - pad; j++) {
    for (let i = pad; i < w - pad; i++) {
      const p = px(data, w, i, j);
      const d = dist(p, bg);
      pixels.push([p, d]);
      if (d > maxDist) maxDist = d;
    }
  }
  if (maxDist < 40) return { bg, fg: null };
  // 안티앨리어싱으로 섞인 가장자리 픽셀을 빼고 글자 안쪽 픽셀만 평균
  const core = pixels.filter(([, d]) => d > maxDist * 0.7).map(([p]) => p);
  return { bg, fg: averageColor(core) };
}

/** 인식한 글자 자리를 주변 배경색으로 칠해 지운다. scale: OCR 캔버스 → 대상 캔버스 배율 */
export function eraseText(targetCanvas, lines, scale) {
  const ctx = targetCanvas.getContext("2d");
  for (const line of lines) {
    const { bbox, bg } = line.erase;
    const pad = Math.max(2, (bbox.y1 - bbox.y0) * 0.12);
    ctx.fillStyle = `rgb(${bg.join(",")})`;
    ctx.fillRect(
      (bbox.x0 - pad) * scale,
      (bbox.y0 - pad) * scale,
      (bbox.x1 - bbox.x0 + pad * 2) * scale,
      (bbox.y1 - bbox.y0 + pad * 2) * scale,
    );
  }
}

function px(data, w, i, j) {
  const k = (j * w + i) * 4;
  return [data[k], data[k + 1], data[k + 2]];
}

function dist(a, b) {
  return Math.hypot(a[0] - b[0], a[1] - b[1], a[2] - b[2]);
}

function median(values) {
  const s = [...values].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

function averageColor(list) {
  if (!list.length) return null;
  const sum = [0, 0, 0];
  for (const p of list) for (let c = 0; c < 3; c++) sum[c] += p[c];
  return sum.map((v) => Math.round(v / list.length));
}

function toHex(rgb) {
  return rgb.map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}
