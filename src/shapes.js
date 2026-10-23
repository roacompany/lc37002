// PDF의 벡터 도형(사각형, 선, 원, 둥근 사각형, 단순한 다각형/곡선)을 PowerPoint 도형으로 바꾼다.
// pdf.js는 경로를 Path2D에 moveTo/lineTo/bezierCurveTo로 만든 뒤 ctx.fill/stroke로 그리므로,
// Path2D 메서드를 감싸 경로 좌표를 기록해 두고 그릴 때의 변환·색·선 두께와 함께 읽는다.

const pathCommands = new WeakMap();
const UNKNOWN = Symbol("unknown");

// 한 경로가 이보다 복잡하면(지도, 복잡한 일러스트 등) 이미지로 남긴다
const MAX_SEGMENTS = 200;
// 한 페이지의 도형이 이보다 많으면 차트·도면 같은 복잡한 그래픽으로 보고 페이지 전체를 이미지로 남긴다
const MAX_SHAPES_PER_PAGE = 1500;

installPathRecorder();

function installPathRecorder() {
  if (typeof Path2D === "undefined" || Path2D.prototype.__pdf2pptRecorded) return;
  const proto = Path2D.prototype;
  proto.__pdf2pptRecorded = true;
  const record = (path, cmd) => {
    let cmds = pathCommands.get(path);
    if (!cmds) pathCommands.set(path, (cmds = []));
    if (cmds === UNKNOWN) return;
    if (cmds.length > MAX_SEGMENTS * 2) pathCommands.set(path, UNKNOWN);
    else cmds.push(cmd);
  };
  const wrap = (name, toCmd) => {
    const original = proto[name];
    if (!original) return;
    proto[name] = function (...args) {
      toCmd(this, ...args);
      return original.apply(this, args);
    };
  };
  wrap("moveTo", (p, x, y) => record(p, ["M", x, y]));
  wrap("lineTo", (p, x, y) => record(p, ["L", x, y]));
  wrap("bezierCurveTo", (p, x1, y1, x2, y2, x, y) => record(p, ["C", x1, y1, x2, y2, x, y]));
  wrap("quadraticCurveTo", (p, x1, y1, x, y) => record(p, ["Q", x1, y1, x, y]));
  wrap("closePath", (p) => record(p, ["Z"]));
  wrap("rect", (p, x, y, w, h) => {
    record(p, ["M", x, y]);
    record(p, ["L", x + w, y]);
    record(p, ["L", x + w, y + h]);
    record(p, ["L", x, y + h]);
    record(p, ["Z"]);
  });
  wrap("addPath", (p, src, m) => {
    const cmds = pathCommands.get(src);
    if (!cmds || cmds === UNKNOWN) return pathCommands.set(p, UNKNOWN);
    for (const cmd of cmds) record(p, m ? transformCmd(cmd, m) : cmd);
  });
  for (const name of ["arc", "arcTo", "ellipse", "roundRect"]) wrap(name, (p) => pathCommands.set(p, UNKNOWN));
}

function transformCmd(cmd, m) {
  const out = [cmd[0]];
  for (let i = 1; i < cmd.length; i += 2) {
    const x = cmd[i];
    const y = cmd[i + 1];
    out.push(m.a * x + m.c * y + m.e, m.b * x + m.d * y + m.f);
  }
  return out;
}

/** 경로를 현재 캔버스 변환을 적용한 장치 좌표 명령 목록으로. 기록이 없거나 지원하지 않는 경로면 null */
export function devicePath(path, matrix) {
  const cmds = path ? pathCommands.get(path) : null;
  if (!cmds || cmds === UNKNOWN || !cmds.length) return null;
  return cmds.map((cmd) => transformCmd(cmd, matrix));
}

/**
 * 경로가 실제로 지나가는 영역. 베지어 제어점은 곡선 밖으로 멀리 나갈 수 있어서
 * 제어점 대신 곡선 위의 점들로 계산한다 (겹침 판단이 지나치게 보수적이 되지 않도록).
 */
export function bboxOfPath(cmds) {
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  const add = (x, y) => {
    x0 = Math.min(x0, x);
    x1 = Math.max(x1, x);
    y0 = Math.min(y0, y);
    y1 = Math.max(y1, y);
  };
  let prev = null;
  for (const cmd of cmds) {
    if (cmd[0] === "Z") continue;
    if ((cmd[0] === "C" || cmd[0] === "Q") && prev) {
      for (let t = 0.05; t < 1; t += 0.05) {
        const u = 1 - t;
        if (cmd[0] === "C") {
          add(
            u * u * u * prev[0] + 3 * u * u * t * cmd[1] + 3 * u * t * t * cmd[3] + t * t * t * cmd[5],
            u * u * u * prev[1] + 3 * u * u * t * cmd[2] + 3 * u * t * t * cmd[4] + t * t * t * cmd[6],
          );
        } else {
          add(u * u * prev[0] + 2 * u * t * cmd[1] + t * t * cmd[3], u * u * prev[1] + 2 * u * t * cmd[2] + t * t * cmd[4]);
        }
      }
    }
    prev = [cmd.at(-2), cmd.at(-1)];
    add(prev[0], prev[1]);
  }
  return { x0, y0, x1, y1 };
}

function intersects(a, b, tol = 0) {
  return a.x0 < b.x1 - tol && b.x0 < a.x1 - tol && a.y0 < b.y1 - tol && b.y0 < a.y1 - tol;
}

function inside(inner, outer, tol) {
  return inner.x0 >= outer.x0 - tol && inner.y0 >= outer.y0 - tol && inner.x1 <= outer.x1 + tol && inner.y1 <= outer.y1 + tol;
}

/**
 * 렌더링 중 기록한 그리기 이벤트(순서대로)에서 PowerPoint 도형으로 옮길 수 있는 것을 고른다.
 * 나중에 그려진 것 중 이미지에 남는 것(사진, 그라데이션, 복잡한 경로, 회전된 글자 등)과 겹치는 도형은
 * 슬라이드에서 그 위로 올라가 가리게 되므로 옮기지 않는다.
 *
 * @param {object[]} events  renderPage가 기록한 이벤트
 * @param {Set<number>} extractedGlyphs  텍스트 상자로 옮긴 글자 순번
 * @param {{width:number,height:number}} canvas
 * @param {number} scale  px → pt
 * @returns {{shapes: object[], suppress: Set<number>}}
 */
export function extractShapes(events, extractedGlyphs, canvas, scale) {
  const paints = mergeFillStroke(events);
  const pageArea = canvas.width * canvas.height;
  const blockers = [];
  const picked = [];

  for (let i = paints.length - 1; i >= 0; i--) {
    const ev = paints[i];
    if (ev.kind === "glyph") {
      if (!extractedGlyphs.has(ev.index)) blockers.push(ev.bbox);
      continue;
    }
    if (ev.kind !== "path") {
      blockers.push(ev.bbox);
      continue;
    }
    const shapes = ev.supported ? classify(ev, scale) : null;
    const coversPage = !ev.stroke && (ev.bbox.x1 - ev.bbox.x0) * (ev.bbox.y1 - ev.bbox.y0) > pageArea * 0.9;
    const clipped = ev.clip && !(ev.clip.rect && inside(ev.bbox, ev.clip.bbox, 1));
    const covered = blockers.some((b) => intersects(ev.bbox, b, 0.5));
    if (shapes && !coversPage && !clipped && !covered) picked.push({ ev, shapes });
    else blockers.push(ev.bbox);
  }

  const total = picked.reduce((n, p) => n + p.shapes.length, 0);
  if (total > MAX_SHAPES_PER_PAGE) return { shapes: [], suppress: new Set() };

  picked.reverse();
  const suppress = new Set();
  for (const { ev } of picked) for (const idx of ev.indices) suppress.add(idx);
  return { shapes: picked.flatMap((p) => p.shapes), suppress };
}

/** fillStroke처럼 같은 경로를 연달아 채우고 그은 것을 하나의 도형(채우기+윤곽선)으로 합친다. */
function mergeFillStroke(events) {
  const out = [];
  for (const ev of events) {
    const prev = out.at(-1);
    if (ev.kind === "path" && ev.stroke && prev?.kind === "path" && prev.fill && !prev.stroke && samePath(prev, ev)) {
      prev.stroke = ev.stroke;
      prev.indices.push(...ev.indices);
      prev.supported &&= ev.supported;
      prev.bbox = expand(prev.bbox, ev.stroke.width / 2);
      continue;
    }
    out.push(ev.kind === "path" && ev.stroke ? { ...ev, bbox: expand(ev.bbox, ev.stroke.width / 2) } : ev);
  }
  return out;
}

function samePath(a, b) {
  if (a.pathRef && a.pathRef === b.pathRef) return true;
  if (!a.cmds || !b.cmds || a.cmds.length !== b.cmds.length) return false;
  return a.cmds.every((cmd, i) => cmd.every((v, j) => (j === 0 ? v === b.cmds[i][j] : Math.abs(v - b.cmds[i][j]) < 0.5)));
}

function expand(b, d) {
  return { x0: b.x0 - d, y0: b.y0 - d, x1: b.x1 + d, y1: b.y1 + d };
}

/** 장치 좌표(px) 경로를 pt 단위 도형 목록으로 분류한다. 지원하지 않는 모양이면 null */
function classify(ev, scale) {
  const subpaths = splitSubpaths(ev.cmds.map((cmd) => cmd.map((v, i) => (i === 0 ? v : v / scale))));
  if (!subpaths.length) return null;
  const style = {
    fill: ev.fill ? { ...ev.fill } : null,
    line: ev.stroke ? { ...ev.stroke, width: ev.stroke.width / scale, dash: ev.stroke.dash.map((d) => d / scale) } : null,
  };

  if (subpaths.length === 1) {
    const shape = classifySubpath(subpaths[0], style);
    return shape === SKIP ? [] : shape ? [shape] : null;
  }

  // 여러 조각으로 된 경로: 채우기가 있으면 조각끼리 겹치지 않을 때만(구멍 뚫린 모양이 아닐 때) 옮긴다
  if (style.fill) {
    const boxes = subpaths.map((sp) => bboxOfPath(sp.cmds));
    for (let i = 0; i < boxes.length; i++) for (let j = i + 1; j < boxes.length; j++) if (intersects(boxes[i], boxes[j], 0.01)) return null;
  }
  // 모든 조각이 사각형·원 같은 기본 도형이면(표 칸 배경 등) 각각 따로, 아니면(점선의 조각들 등) 하나의 도형으로
  const parts = subpaths.map((sp) => classifySubpath(sp, style)).filter((sh) => sh !== SKIP);
  if (parts.every((sh) => sh && sh.type !== "custGeom")) return parts;
  const all = subpaths.flatMap((sp) => (sp.closed ? [...sp.cmds, ["Z"]] : sp.cmds));
  if (all.length > MAX_SEGMENTS) return null;
  return [customGeometry(all, bboxOfPath(all), style, subpaths.every((sp) => sp.closed))];
}

const SKIP = Symbol("skip");

function splitSubpaths(cmds) {
  const subpaths = [];
  let cur = null;
  for (const cmd of cmds) {
    if (cmd[0] === "M") {
      cur = { start: [cmd[1], cmd[2]], cmds: [cmd], closed: false };
      subpaths.push(cur);
    } else if (!cur) {
      return [];
    } else if (cmd[0] === "Z") {
      cur.closed = true;
    } else {
      cur.cmds.push(cmd);
    }
  }
  // moveTo만 있는 조각(빈 조각) 제거
  return subpaths.filter((sp) => sp.cmds.length > 1);
}

function classifySubpath(sp, style) {
  const { cmds } = sp;
  const b = bboxOfPath(cmds);
  const w = b.x1 - b.x0;
  const h = b.y1 - b.y0;
  const base = { x: b.x0, y: b.y0, w, h, fill: style.fill, line: style.line };
  const last = cmds.at(-1);
  const end = [last.at(-2), last.at(-1)];
  const closed = sp.closed || (Math.abs(end[0] - sp.start[0]) < 0.01 && Math.abs(end[1] - sp.start[1]) < 0.01);
  const segs = cmds.slice(1);
  const allLines = segs.every((s) => s[0] === "L");

  if (w < 0.05 && h < 0.05) return SKIP;

  if (allLines) {
    // 선 하나
    if (segs.length === 1 && !closed) {
      if (!style.line) return SKIP; // 채우기만 있는 선은 보이지 않음
      const [, x1, y1] = cmds[0];
      const [, x2, y2] = segs[0];
      return { ...base, type: "line", fill: null, flipH: x2 < x1, flipV: y2 < y1 };
    }
    const pts = uniquePoints(cmds);
    if (closed && pts.length === 4 && isAxisAlignedRect(pts, b)) return { ...base, type: "rect" };
  }

  if (closed && w > 0.5 && h > 0.5 && segs.some((s) => s[0] === "C" || s[0] === "Q")) {
    const samples = samplePoints(cmds);
    if (onEllipse(samples, b)) return { ...base, type: "ellipse" };
    const radius = roundRectRadius(cmds, samples, b);
    if (radius !== null) return { ...base, type: "roundRect", radius };
  }

  if (segs.length > MAX_SEGMENTS) return null;
  // 그 밖의 단순한 모양은 사용자 지정 도형(직선·베지어 곡선)으로
  return customGeometry(sp.closed ? [...cmds, ["Z"]] : cmds, b, style, closed);
}

function customGeometry(cmds, b, style, closed) {
  const w = b.x1 - b.x0;
  const h = b.y1 - b.y0;
  const points = cmds.map((cmd) => {
    const rel = (i) => ({ x: cmd[i] - b.x0, y: cmd[i + 1] - b.y0 });
    if (cmd[0] === "M") return { ...rel(1), moveTo: true };
    if (cmd[0] === "L") return rel(1);
    if (cmd[0] === "C") {
      const c1 = rel(1);
      const c2 = rel(3);
      return { ...rel(5), curve: { type: "cubic", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y } };
    }
    if (cmd[0] === "Z") return { close: true };
    const c = rel(1);
    return { ...rel(3), curve: { type: "quadratic", x1: c.x, y1: c.y } };
  });
  // 너비나 높이가 0인 사용자 지정 도형은 PowerPoint에서 그려지지 않으므로 최소 크기를 준다
  return { x: b.x0, y: b.y0, w: Math.max(w, 0.01), h: Math.max(h, 0.01), fill: style.fill, line: style.line, type: "custGeom", points, closed };
}

/** 경로 위의 점들 (꼭짓점 + 곡선의 중간점) */
function samplePoints(cmds) {
  const pts = [];
  let prev = null;
  for (const c of cmds) {
    if (c[0] === "C" && prev) {
      for (const t of [0.25, 0.5, 0.75]) {
        const u = 1 - t;
        pts.push([
          u * u * u * prev[0] + 3 * u * u * t * c[1] + 3 * u * t * t * c[3] + t * t * t * c[5],
          u * u * u * prev[1] + 3 * u * u * t * c[2] + 3 * u * t * t * c[4] + t * t * t * c[6],
        ]);
      }
    } else if (c[0] === "Q" && prev) {
      pts.push([0.25 * prev[0] + 0.5 * c[1] + 0.25 * c[3], 0.25 * prev[1] + 0.5 * c[2] + 0.25 * c[4]]);
    }
    if (c[0] !== "Z") {
      prev = [c.at(-2), c.at(-1)];
      pts.push(prev);
    }
  }
  return pts;
}

/** 모든 점이 경계 상자에 내접하는 타원 위에 있는지 */
function onEllipse(pts, b) {
  const rx = (b.x1 - b.x0) / 2;
  const ry = (b.y1 - b.y0) / 2;
  const cx = b.x0 + rx;
  const cy = b.y0 + ry;
  const tol = Math.max(0.25, Math.min(rx, ry) * 0.03);
  return pts.every(([x, y]) => {
    const dx = (x - cx) / rx;
    const dy = (y - cy) / ry;
    const r = Math.hypot(dx, dy);
    if (r === 0) return false;
    // 타원 위의 가장 가까운 점까지의 거리를 반지름 방향 차이로 근사
    return Math.abs(r - 1) * Math.min(rx, ry) < tol;
  });
}

/** 모든 점이 같은 반지름의 둥근 사각형 테두리 위에 있으면 그 반지름(pt), 아니면 null */
function roundRectRadius(cmds, pts, b) {
  const hw = (b.x1 - b.x0) / 2;
  const hh = (b.y1 - b.y0) / 2;
  const cx = b.x0 + hw;
  const cy = b.y0 + hh;
  // 위쪽 변의 직선이 시작하는 위치로 반지름을 구한다
  const tol = Math.max(0.25, Math.min(hw, hh) * 0.03);
  let r = null;
  let prev = null;
  for (const c of cmds) {
    const end = [c.at(-2), c.at(-1)];
    if (c[0] === "L" && prev && Math.abs(prev[1] - b.y0) < tol && Math.abs(end[1] - b.y0) < tol) {
      r = Math.min(prev[0], end[0]) - b.x0;
      break;
    }
    if (c[0] !== "Z") prev = end;
  }
  if (r === null || r <= 0.1 || r > Math.min(hw, hh) + tol) return null;
  r = Math.min(r, hw, hh);
  const ok = pts.every(([x, y]) => {
    // 둥근 사각형의 부호 있는 거리 함수
    const qx = Math.abs(x - cx) - (hw - r);
    const qy = Math.abs(y - cy) - (hh - r);
    const d = Math.hypot(Math.max(qx, 0), Math.max(qy, 0)) + Math.min(Math.max(qx, qy), 0) - r;
    return Math.abs(d) < tol;
  });
  return ok ? r : null;
}

function uniquePoints(cmds) {
  const pts = [];
  for (const cmd of cmds) {
    const p = [cmd[1], cmd[2]];
    if (!pts.some((q) => Math.abs(q[0] - p[0]) < 0.01 && Math.abs(q[1] - p[1]) < 0.01)) pts.push(p);
  }
  return pts;
}

function isAxisAlignedRect(pts, b) {
  const tol = 0.05;
  return pts.every(([x, y]) => (Math.abs(x - b.x0) < tol || Math.abs(x - b.x1) < tol) && (Math.abs(y - b.y0) < tol || Math.abs(y - b.y1) < tol));
}

/** 분류한 도형을 슬라이드에 추가한다. (좌표는 pt, 슬라이드 단위는 인치) */
export function addShapes(slide, shapes, ptToIn, offX, offY) {
  const pt = (v) => v * ptToIn * 72; // 슬라이드 배율이 적용된 pt
  for (const s of shapes) {
    const opts = {
      x: offX + s.x * ptToIn,
      y: offY + s.y * ptToIn,
      w: Math.max(s.w * ptToIn, 0),
      h: Math.max(s.h * ptToIn, 0),
      fill: s.fill ? { color: s.fill.color, transparency: transparency(s.fill.alpha) } : { type: "none" },
      line: s.line
        ? {
            color: s.line.color,
            width: Math.max(round(pt(s.line.width), 2), 0.25),
            transparency: transparency(s.line.alpha),
            dashType: dashType(s.line),
          }
        : { type: "none" },
    };
    if (s.type === "line") {
      opts.flipH = s.flipH;
      opts.flipV = s.flipV;
      delete opts.fill;
    }
    if (s.type === "roundRect") opts.rectRadius = s.radius * ptToIn;
    if (s.type === "custGeom") {
      opts.points = s.points.map((p) => {
        if (p.close) return p;
        const q = { ...p, x: p.x * ptToIn, y: p.y * ptToIn };
        if (p.curve) {
          q.curve = { ...p.curve };
          for (const k of ["x1", "y1", "x2", "y2"]) if (k in q.curve) q.curve[k] *= ptToIn;
        }
        return q;
      });
      // 열린 경로를 선으로만 그은 경우 채우기가 생기지 않게
      if (!s.closed && !s.fill) opts.fill = { type: "none" };
    }
    slide.addShape(s.type, opts);
  }
}

function transparency(alpha) {
  return Math.round((1 - Math.max(0, Math.min(1, alpha))) * 100);
}

function dashType(line) {
  if (!line.dash.length) return "solid";
  const unit = line.dash[0] / Math.max(line.width, 0.01);
  return unit <= 1.5 ? "sysDot" : unit <= 4 ? "sysDash" : "dash";
}

function round(v, d) {
  const p = 10 ** d;
  return Math.round(v * p) / p;
}
