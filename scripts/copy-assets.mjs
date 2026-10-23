// 런타임에 불러오는 정적 파일을 public/ 아래로 복사한다.
// - pdf.js: CMap(한글 등 CJK 글꼴), 표준 글꼴, WASM 디코더
// - tesseract.js(OCR): 워커, WASM 엔진, 한국어·영어 학습 데이터
import { cpSync, mkdirSync, readdirSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const pkgDir = (name) => dirname(require.resolve(`${name}/package.json`));
// import.meta.dirname은 Node 20.11+ 전용이라 fileURLToPath로 구한다
const publicDir = join(dirname(fileURLToPath(import.meta.url)), "..", "public");

const pdfjsTarget = join(publicDir, "pdfjs");
rmSync(pdfjsTarget, { recursive: true, force: true });
for (const dir of ["cmaps", "standard_fonts", "wasm", "iccs"]) {
  cpSync(join(pkgDir("pdfjs-dist"), dir), join(pdfjsTarget, dir), { recursive: true });
}

const ocrTarget = join(publicDir, "tesseract");
rmSync(ocrTarget, { recursive: true, force: true });
mkdirSync(join(ocrTarget, "core"), { recursive: true });
mkdirSync(join(ocrTarget, "lang"), { recursive: true });
cpSync(join(pkgDir("tesseract.js"), "dist", "worker.min.js"), join(ocrTarget, "worker.min.js"));
// LSTM 엔진만 사용하므로 *-lstm.wasm.js 변형만 복사 (브라우저 SIMD 지원에 따라 하나를 고름)
const coreDir = pkgDir("tesseract.js-core");
for (const file of readdirSync(coreDir).filter((f) => f.endsWith("-lstm.wasm.js"))) {
  cpSync(join(coreDir, file), join(ocrTarget, "core", file));
}
for (const lang of ["kor", "eng"]) {
  const src = join(pkgDir(`@tesseract.js-data/${lang}`), "4.0.0_best_int", `${lang}.traineddata.gz`);
  cpSync(src, join(ocrTarget, "lang", `${lang}.traineddata.gz`));
}

console.log(`static assets copied to ${publicDir}`);
