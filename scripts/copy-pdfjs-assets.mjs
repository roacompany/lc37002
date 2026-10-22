// pdf.js가 런타임에 불러오는 CMap(한글 등 CJK 폰트), 표준 폰트, WASM 디코더를 public/ 아래로 복사한다.
import { cpSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const pdfjsDir = dirname(require.resolve("pdfjs-dist/package.json"));
const target = join(import.meta.dirname, "..", "public", "pdfjs");

rmSync(target, { recursive: true, force: true });
for (const dir of ["cmaps", "standard_fonts", "wasm", "iccs"]) {
  cpSync(join(pdfjsDir, dir), join(target, dir), { recursive: true });
}
console.log(`pdf.js assets copied to ${target}`);
