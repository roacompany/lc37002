import { openPdf, convertPdfToPptx } from "./converter.js";

const $ = (id) => document.getElementById(id);
const fileInput = $("file");
const drop = $("drop");
const convertBtn = $("convert");
const cancelBtn = $("cancel");
const progress = $("progress");
const barFill = $("bar-fill");
const progressText = $("progress-text");
const message = $("message");

let currentFile = null;
let pdf = null;
let controller = null;

fileInput.addEventListener("change", () => fileInput.files[0] && selectFile(fileInput.files[0]));

drop.addEventListener("dragover", (e) => {
  e.preventDefault();
  drop.classList.add("over");
});
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  drop.classList.remove("over");
  const file = e.dataTransfer.files[0];
  if (file) selectFile(file);
});

async function selectFile(file) {
  setMessage("");
  if (!/\.pdf$/i.test(file.name) && file.type !== "application/pdf") {
    setMessage("PDF 파일만 변환할 수 있습니다.", true);
    return;
  }
  await pdf?.destroy();
  pdf = null;
  currentFile = file;
  convertBtn.disabled = true;
  $("file-info").textContent = `${file.name} · 불러오는 중…`;
  try {
    pdf = await openPdf(new Uint8Array(await file.arrayBuffer()));
    $("file-info").textContent = `${file.name} · ${pdf.numPages}페이지 · ${formatSize(file.size)}`;
    convertBtn.disabled = false;
  } catch (err) {
    console.error(err);
    const msg = err?.name === "PasswordException" ? "암호가 걸린 PDF는 지원하지 않습니다." : "PDF를 열 수 없습니다. 손상된 파일인지 확인해 주세요.";
    $("file-info").textContent = file.name;
    setMessage(msg, true);
  }
}

convertBtn.addEventListener("click", async () => {
  if (!pdf) return;
  let pages;
  try {
    pages = parseRange($("range").value, pdf.numPages);
  } catch (err) {
    setMessage(err.message, true);
    return;
  }

  controller = new AbortController();
  convertBtn.disabled = true;
  cancelBtn.hidden = false;
  progress.hidden = false;
  setMessage("");
  updateProgress(0, pages?.length ?? pdf.numPages);

  try {
    const blob = await convertPdfToPptx(pdf, {
      mode: document.querySelector('input[name="mode"]:checked').value,
      scale: Number($("scale").value),
      imageFormat: $("format").value,
      pages,
      includeLinks: $("links").checked,
      signal: controller.signal,
      onProgress: ({ page, done, total }) => updateProgress(done, total, page),
    });
    download(blob, currentFile.name.replace(/\.pdf$/i, "") + ".pptx");
    setMessage(`변환 완료 · ${formatSize(blob.size)}`);
  } catch (err) {
    if (err?.name === "AbortError") setMessage("변환을 취소했습니다.");
    else {
      console.error(err);
      setMessage(`변환 중 오류가 발생했습니다: ${err?.message ?? err}`, true);
    }
  } finally {
    controller = null;
    convertBtn.disabled = false;
    cancelBtn.hidden = true;
  }
});

cancelBtn.addEventListener("click", () => controller?.abort());

function updateProgress(done, total, page) {
  barFill.style.width = `${(done / total) * 100}%`;
  progressText.textContent = done === total ? `${total}/${total} 페이지 완료 · 파일 만드는 중…` : `${done}/${total} 페이지${page ? ` (현재 ${page}쪽)` : ""}`;
}

/** "1-3, 5" 형식을 페이지 번호 배열로. 비어 있으면 null(전체). */
export function parseRange(text, max) {
  const value = text.trim();
  if (!value) return null;
  const pages = [];
  for (const part of value.split(",")) {
    const m = part.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
    if (!m) throw new Error(`페이지 범위 형식이 올바르지 않습니다: "${part.trim()}"`);
    const start = Number(m[1]);
    const end = m[2] ? Number(m[2]) : start;
    if (start < 1 || end > max || start > end) throw new Error(`페이지 범위는 1~${max} 사이여야 합니다.`);
    for (let p = start; p <= end; p++) if (!pages.includes(p)) pages.push(p);
  }
  return pages;
}

function download(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

function setMessage(text, isError = false) {
  message.textContent = text;
  message.classList.toggle("error", isError);
}

function formatSize(bytes) {
  return bytes > 1024 * 1024 ? `${(bytes / 1024 / 1024).toFixed(1)} MB` : `${Math.ceil(bytes / 1024)} KB`;
}
