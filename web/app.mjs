// 絵コンテ作成ソフト（ブラウザ版）の画面。デスクトップ版（ekonte/gui.py）と同じ流れ：
//   ① カット絵PDFを読み取る → ② 要確認を直す → ③ Excelを選ぶ → ④ 絵コンテPDFを書き出す
// 計算は worker.mjs（ブラウザの中の Python）が行う。選んだファイルはどこにも送らない。

const VERSION = document.documentElement.dataset.version;
const APP_NAME = "絵コンテ作成ソフト";
const FORMATS = { 横: "横", 縦: "縦 9:16" };  // ekonte/layout.py の FORMATS と同じ
const CUT_RE = /^(\d+)-(\d+)([a-j]?)$/;       // ekonte/compose.py の CUT_RE と同じ
const STATUS_CLASS = { OK: "st-ok", 修正済み: "st-done", 確認済み: "st-done", 要確認: "st-review", エラー: "st-error", 除外: "st-excluded" };
const NEEDS_REVIEW = ["要確認", "エラー"];

const $ = (id) => document.getElementById(id);
const el = (tag, props = {}, ...children) => {
  const node = Object.assign(document.createElement(tag), props);
  node.append(...children);
  return node;
};

// このブラウザだけに覚えておく設定（使えない環境では覚えないだけ）
const store = {
  get(key) {
    try { return localStorage.getItem(`ekonte.${key}`); } catch { return null; }
  },
  set(key, value) {
    try { localStorage.setItem(`ekonte.${key}`, value); } catch { /* 覚えられなくても動く */ }
  },
};

// ---------- 計算の受け渡し（worker.mjs） ----------

const worker = new Worker(new URL(`worker.mjs?v=${VERSION}`, import.meta.url), { type: "module" });
const pending = new Map();
let nextId = 1;

worker.onmessage = ({ data }) => {
  if (data.event === "status") return showEngine("busy", `${data.text}…`);
  if (data.event === "ready") return showEngine("ready", "準備ができました");
  if (data.event === "fatal") return engineFailed(data.error);
  const job = pending.get(data.id);
  if (!job) return;
  if (data.event === "progress") return job.onProgress?.(data);
  pending.delete(data.id);
  if (data.ok) job.resolve(data.result);
  else job.reject(new Error(data.error));
};
worker.onerror = (e) => {
  e.preventDefault();
  engineFailed("このブラウザでは動かせませんでした");
};

function call(cmd, args = {}, { transfer = [], onProgress } = {}) {
  return new Promise((resolve, reject) => {
    const id = nextId++;
    pending.set(id, { resolve, reject, onProgress });
    worker.postMessage({ id, cmd, args }, transfer);
  });
}

function showEngine(kind, text) {
  const box = $("engine");
  box.className = `engine ${kind}`;
  $("engine-text").textContent = text;
  if (kind === "ready") setTimeout(() => box.classList.add("fade"), 2500);
}

let engineBroken = false;
function engineFailed(message) {
  if (engineBroken) return;
  engineBroken = true;
  showEngine("error", "準備に失敗しました");
  for (const job of pending.values()) job.reject(new Error(message));
  pending.clear();
  alertDialog("準備に失敗しました", `${message}\n\nChrome・Edge・Safari・Firefox の最新版で開き直してください。\nネットがない場所では、アプリ版（下の「アプリ版」から）を使えます。`);
}

// ---------- ① 読み取り ----------

const tbody = document.querySelector("#table tbody");
const results = [];         // 読み取り結果（1ページ1つ）: {page, name, status, warnings, orientation, format}
const images = [];          // 行 → {picture, mark}（画像の blob URL）
const confirmed = new Set(); // 番号を変えずに「この番号で確定」した行
let selected = -1;
let reading = false;
let unsaved = false;        // 読み取ってから、まだ書き出していない

function blobUrl(bytes, type) {
  return bytes ? URL.createObjectURL(new Blob([bytes], { type })) : null;
}

function clearResults() {
  for (const { picture, mark } of images) {
    URL.revokeObjectURL(picture);
    if (mark) URL.revokeObjectURL(mark);
  }
  results.length = 0;
  images.length = 0;
  confirmed.clear();
  selected = -1;
  tbody.replaceChildren();
  $("empty").hidden = false;
  $("summary").textContent = "";
  showImage($("v-picture"), null, "表の行を選ぶと、ここにカット絵が出ます");
  showImage($("v-mark"), null, "読み取った番号のマーク欄");
}

async function readPdf(file) {
  if (reading) return;
  if (!/\.pdf$/i.test(file.name) && file.type !== "application/pdf") {
    return alertDialog("PDFを選んでください", `${file.name} はPDFではありません。カット絵はPDFで選んでください。`);
  }
  clearResults();
  reading = true;
  busy($("b-open"), true, "読み取り中…");
  $("pdf-label").textContent = file.name;
  $("pdf-label").classList.remove("muted");
  const bar = $("progress");
  bar.hidden = false;
  bar.removeAttribute("value");
  try {
    const pdf = new Uint8Array(await file.arrayBuffer());
    await call("read", { pdf }, { transfer: [pdf.buffer], onProgress: ({ page }) => addResult(page) });
  } catch (e) {
    await alertDialog("読み取りエラー", e.message);
  } finally {
    reading = false;
    busy($("b-open"), false);
    bar.hidden = true;
  }
  if (!results.length) return;
  unsaved = true;
  $("pdf-label").textContent = `${file.name}　（${FORMATS[sheetFormat()]}）`;
  refreshStatus();
  selectNextReview(true);
}

function addResult(page) {
  const { picture, mark, total, ...info } = page;
  const i = results.length;
  results.push(info);
  images.push({ picture: blobUrl(picture, "image/png"), mark: blobUrl(mark, "image/png") });
  const input = el("input", { type: "text", value: info.name ?? "", placeholder: "例: 2-3 / 2-3a（空欄で除外）", spellcheck: false, autocomplete: "off" });
  input.setAttribute("aria-label", `p${info.page} の番号`);
  tbody.append(el("tr", {},
    el("td", { className: "c-page", textContent: info.page }),
    el("td", { className: "c-name" }, input),
    el("td", { className: "c-status" }),
    el("td", { className: "c-msg" })));
  $("empty").hidden = true;
  const bar = $("progress");
  bar.max = total;
  bar.value = info.page;
  refreshStatus();
  if (i === 0) select(0);
}

/** 読み取ったカット絵の用紙の種類（最初に読めたページに合わせる。混在は読み取り時に要確認にしている） */
function sheetFormat() {
  return results.find((r) => r.status !== "エラー")?.format ?? "横";
}

// ---------- ② 確認・修正 ----------

/** 行の番号（全角の数字やハイフンは半角にそろえる） */
function nameOf(i) {
  return tbody.rows[i].querySelector("input").value.normalize("NFKC").trim();
}

function nameCounts() {
  const counts = new Map();
  results.forEach((_, i) => {
    const name = nameOf(i);
    if (name) counts.set(name, (counts.get(name) ?? 0) + 1);
  });
  return counts;
}

/** [状態, 内容]。手で直した行は読み取り時の警告を出さない（gui.py の row_status と同じ） */
function rowStatus(i, counts = nameCounts()) {
  const r = results[i];
  const name = nameOf(i);
  if (r.status === "エラー" && !name) return ["エラー", r.warnings.join("；")];
  if (!name) return ["除外", "番号が空欄のため絵コンテに入れません" + (r.warnings.length ? `（${r.warnings.join("；")}）` : "")];
  if (!CUT_RE.test(name)) return ["要確認", "番号の形式が正しくありません（例: 2-3, 2-3a）"];
  if (counts.get(name) > 1) {
    const pages = results.filter((_, j) => nameOf(j) === name).map((x) => x.page);
    return ["要確認", `番号 ${name} が重複しています（p${pages.join(", p")}）`];
  }
  if (name !== (r.name ?? "")) return ["修正済み", `手で修正（読み取り: ${r.name || "なし"}）`];
  if (confirmed.has(i)) return ["確認済み", "読み取った番号を目で確認して確定"];
  const warnings = r.warnings.filter((w) => !w.includes("重複"));
  const note = ["", "正位置"].includes(r.orientation) ? "" : `［${r.orientation}を補正］`;
  if (warnings.length) return ["要確認", warnings.join("；") + note];
  return ["OK", note];
}

function refreshStatus() {
  const counts = nameCounts();
  const tally = new Map();
  results.forEach((_, i) => {
    const [status, msg] = rowStatus(i, counts);
    tally.set(status, (tally.get(status) ?? 0) + 1);
    const tr = tbody.rows[i];
    tr.className = STATUS_CLASS[status] + (i === selected ? " selected" : "");
    tr.dataset.status = status;
    tr.cells[2].textContent = status;
    tr.cells[3].textContent = msg;
  });
  $("summary").textContent = tally.size ? `全 ${results.length} ページ：${[...tally].map(([k, v]) => `${k} ${v}`).join("　")}` : "";
  applyFilter();
}

function applyFilter() {
  const only = $("only-review").checked;
  for (const tr of tbody.rows) tr.hidden = only && !NEEDS_REVIEW.includes(tr.dataset.status);
}

function showImage(frame, url, placeholder) {
  frame.replaceChildren(url ? el("img", { src: url, alt: "" }) : el("span", { textContent: placeholder }));
}

function select(i) {
  tbody.rows[selected]?.classList.remove("selected");
  selected = i;
  tbody.rows[i].classList.add("selected");
  showImage($("v-picture"), images[i].picture, "");
  showImage($("v-mark"), images[i].mark, "（位置合わせマークが見つからないため表示できません）");
}

function focusRow(i) {
  select(i);
  tbody.rows[i].scrollIntoView({ block: "nearest" });
  tbody.rows[i].querySelector("input").focus({ preventScroll: true });
}

function moveSelection(step) {
  for (let i = selected + step; i >= 0 && i < results.length; i += step) {
    if (!tbody.rows[i].hidden) return focusRow(i);
  }
}

function selectNextReview(fromStart = false) {
  const n = results.length;
  const cur = fromStart || selected < 0 ? -1 : selected;
  for (let k = 1; k <= n; k++) {
    const i = (cur + k) % n;
    if (NEEDS_REVIEW.includes(tbody.rows[i].dataset.status)) return focusRow(i);
  }
}

async function confirmSelected() {
  const i = selected;
  if (i < 0 || !nameOf(i)) return;
  const [status, msg] = rowStatus(i);
  // 形式の誤り・重複は確定できない（番号を直す必要がある）
  if (!CUT_RE.test(nameOf(i)) || msg.includes("重複")) {
    await alertDialog(APP_NAME, "番号の形式が正しくないか、他のページと重複しています。\n番号を直してください。");
    return focusRow(i);
  }
  if (status === "要確認") confirmed.add(i);
  refreshStatus();
  selectNextReview();
}

const rowIndex = (node) => node.closest?.("tr")?.sectionRowIndex ?? -1;

tbody.addEventListener("input", (e) => {
  confirmed.delete(rowIndex(e.target));
  unsaved = true;
  refreshStatus();
});
tbody.addEventListener("focusin", (e) => {
  const i = rowIndex(e.target);
  if (i >= 0 && i !== selected) select(i);
});
tbody.addEventListener("click", (e) => {
  const i = rowIndex(e.target);
  if (i >= 0 && e.target.tagName !== "INPUT") focusRow(i);
});
tbody.addEventListener("keydown", (e) => {
  if (e.isComposing || e.keyCode === 229) return;  // 日本語入力の変換中の Enter などは使わない
  if (e.key === "Enter") {
    e.preventDefault();
    confirmSelected();
  } else if (e.key === "ArrowDown" || e.key === "ArrowUp") {
    e.preventDefault();
    moveSelection(e.key === "ArrowDown" ? 1 : -1);
  }
});
// 画像を押すと大きく表示する（もう一度押すか Esc で閉じる）
for (const id of ["v-picture", "v-mark"]) {
  $(id).addEventListener("click", (e) => {
    if (e.target.tagName !== "IMG") return;
    $("zoom-img").src = e.target.src;
    $("zoom").showModal();
  });
}
$("zoom").addEventListener("click", () => $("zoom").close());
$("only-review").addEventListener("change", applyFilter);
$("b-confirm").addEventListener("click", confirmSelected);
$("b-next").addEventListener("click", () => selectNextReview());

// ---------- ③ Excel ----------

const excel = {};   // 種類 → 選んだファイル名
const latest = {};  // 種類 → 最後の操作の番号（古い結果で上書きしないため）

async function setExcel(row, file) {
  const key = row.dataset.key;
  const info = row.querySelector(".info");
  const before = [info.textContent, info.className];
  const token = (latest[key] = (latest[key] ?? 0) + 1);
  info.textContent = `${file.name}　（読み込み中…）`;
  info.className = "info muted";
  try {
    const data = new Uint8Array(await file.arrayBuffer());
    const ext = (/\.[^.]+$/.exec(file.name)?.[0] ?? ".xlsx").toLowerCase();
    const summary = await call("setExcel", { key, data, ext }, { transfer: [data.buffer] });
    if (latest[key] !== token) return;
    excel[key] = file.name;
    info.textContent = `${file.name}　（${summary}）`;
    info.className = "info";
  } catch (e) {
    if (latest[key] !== token) return;
    [info.textContent, info.className] = before;
    await alertDialog("Excelを読み込めません", `${file.name}\n\n${e.message}`);
  }
}

function clearExcel(row) {
  const key = row.dataset.key;
  latest[key] = (latest[key] ?? 0) + 1;
  delete excel[key];
  call("clearExcel", { key }).catch(() => {});
  const info = row.querySelector(".info");
  info.textContent = "未選択";
  info.className = "info muted";
}

for (const row of document.querySelectorAll(".pick")) {
  const input = row.querySelector("input[type=file]");
  row.querySelector(".choose").addEventListener("click", () => input.click());
  input.addEventListener("change", () => {
    const file = input.files[0];
    input.value = "";
    if (file) setExcel(row, file);
  });
  row.querySelector(".clear").addEventListener("click", () => clearExcel(row));
}

// ---------- ④ 書き出し ----------

/** 番号を絵コンテでの書き方にそろえる（"02-03a" → "2-3a"。compose.make_cut と同じ） */
function cutName(name) {
  const [, s, c, sub] = CUT_RE.exec(name);
  return `${Number(s)}-${Number(c)}${sub}`;
}

async function exportPdf() {
  if (reading) return alertDialog(APP_NAME, "読み取りが終わってから書き出してください。");
  if (!results.length) return alertDialog(APP_NAME, "先に ① でカット絵PDFを読み取ってください。");
  const counts = nameCounts();
  const statuses = results.map((_, i) => rowStatus(i, counts)[0]);
  const waiting = results.filter((_, i) => statuses[i] === "要確認").map((r) => r.page);
  if (waiting.length) {
    return alertDialog("要確認のページがあります",
      `p${waiting.join(", p")} が要確認のままです。\n番号を修正するか、空欄にして除外してから書き出してください。`);
  }
  const rows = statuses.flatMap((s, i) => (["OK", "修正済み", "確認済み"].includes(s) ? [[i, nameOf(i)]] : []));
  if (!rows.length && !excel.cuts) return alertDialog(APP_NAME, "絵コンテに入れるカットがありません。");
  if (excel.cuts && !(await confirmCutCheck(rows.map(([, name]) => cutName(name))))) return;
  busy($("b-export"), true, "書き出し中…");
  let r;
  try {
    r = await call("export", { rows, fmt: sheetFormat(), withImages: $("export-images").checked });
  } catch (e) {
    return alertDialog("書き出しエラー", e.message);
  } finally {
    busy($("b-export"), false);
  }
  unsaved = false;
  const excluded = statuses.filter((s) => s === "除外" || s === "エラー").length;
  let msg = `${r.cuts}カット / ${r.pages}ページ`;
  if (excluded) msg += `\n\n※ 除外・エラーの ${excluded} ページは入れていません`;
  if (r.warnings.length) msg += `\n\n⚠ 文字を最小にしても欄に収まらず、はみ出た分を切っています：\n${r.warnings.join("\n")}`;
  const files = [{ bytes: r.pdf, name: "絵コンテ.pdf", label: "PDFを保存" }];
  if (r.images) files.push({ bytes: r.images, name: "絵コンテ_カット絵.zip", label: "画像をZipで保存" });
  await doneDialog("絵コンテができました", msg, files);
}

/** カット表とカット絵が食い違っていれば、内容を見せて続けるか確かめる */
async function confirmCutCheck(names) {
  let r;
  try {
    r = await call("check", { names });
  } catch (e) {
    await alertDialog("カット表を読み込めません", e.message);
    return false;
  }
  if (!r.missing.length && !r.extra.length) return true;
  const lines = [];
  if (r.missing.length) lines.push(`カット絵がないカット（PICTURE を空けて載せます）：\n  ${r.missing.join(", ")}`);
  if (r.extra.length) lines.push(`カット表にないカット（ACTION/SE などは空欄になります）：\n  ${r.extra.join(", ")}`);
  const go = await dialog("カット表とカット絵が一致しません", `${lines.join("\n\n")}\n\nこのまま書き出しますか？`, [
    { label: "やめる", value: false },
    { label: "書き出す", value: true, primary: true },
  ]);
  return go === true;
}

$("b-export").addEventListener("click", exportPdf);

// ---------- 用紙・サンプル ----------

async function makeSheet() {
  const pages = Math.min(999, Math.max(1, Number.parseInt($("sheet-pages").value, 10) || 20));
  $("sheet-pages").value = pages;
  const fmt = $("sheet-format").value;
  store.set("sheetFormat", fmt);
  busy($("b-sheet"), true, "作成中…");
  let pdf;
  try {
    pdf = await call("sheet", { pages, fmt });
  } catch (e) {
    return alertDialog("用紙を作成できません", e.message);
  } finally {
    busy($("b-sheet"), false);
  }
  const name = fmt === "横" ? "カット絵用紙.pdf" : `カット絵用紙_${fmt}.pdf`;
  await doneDialog("カット絵用紙ができました",
    `${FORMATS[fmt]}・${pages}ページ\n\n印刷するときは「実際のサイズ」（100%）で印刷してください。「用紙に合わせる」にすると縮小されます。`,
    [{ bytes: pdf, name, label: "PDFを保存" }]);
}

let samplesReady = null;
function loadSamples() {
  samplesReady ??= (async () => {
    const res = await fetch(`samples/samples.json?v=${VERSION}`);
    if (!res.ok) throw new Error(`一覧を読み込めません（${res.status}）`);
    const { files, zip } = await res.json();
    const list = $("samples-list");
    let group = null;
    for (const f of files) {
      if (f.group !== group) list.append(el("h4", { textContent: (group = f.group) }));
      list.append(el("div", { className: "sample" },
        el("span", { className: "name", textContent: f.name }),
        el("span", { className: "desc", textContent: f.desc }),
        el("a", { className: "btn", href: `samples/${encodeURIComponent(f.name)}`, download: f.name, textContent: "保存" })));
    }
    Object.assign($("samples-zip"), { href: `samples/${encodeURIComponent(zip)}`, download: zip });
  })();
  return samplesReady;
}

async function showSamples() {
  try {
    await loadSamples();
  } catch (e) {
    samplesReady = null;
    return alertDialog("サンプルを読み込めません", e.message);
  }
  $("samples-dlg").showModal();
}

$("b-sheet").addEventListener("click", makeSheet);
$("b-samples").addEventListener("click", showSamples);
$("samples-close").addEventListener("click", () => $("samples-dlg").close());
$("samples-dlg").addEventListener("click", (e) => {
  if (e.target === e.currentTarget && outside(e)) e.currentTarget.close();  // まわりの暗いところで閉じる
});
const savedFormat = store.get("sheetFormat");
if (savedFormat in FORMATS) $("sheet-format").value = savedFormat;

// ---------- 確認・完了の画面 ----------

let dialogs = Promise.resolve();  // 1つずつ順に出す

/** buttons: [{label, value, primary, href, download, newTab, keepOpen}]。押したボタンの value（閉じただけなら null）を返す */
function dialog(title, text, buttons = [{ label: "閉じる", primary: true }]) {
  const shown = dialogs.then(() => new Promise((resolve) => {
    const dlg = $("dlg");
    let answer = null;
    $("dlg-title").textContent = title;
    $("dlg-text").textContent = text;
    $("dlg-buttons").replaceChildren(...buttons.map((b) => {
      const node = b.href ? el("a", { href: b.href }) : el("button", { type: "button" });
      node.className = b.primary ? "btn primary" : "btn";
      node.textContent = b.label;
      if (b.download) node.download = b.download;
      if (b.newTab) Object.assign(node, { target: "_blank", rel: "noopener" });
      node.addEventListener("click", () => {
        if (b.keepOpen) return;
        answer = b.value ?? null;
        dlg.close();
      });
      return node;
    }));
    dlg.addEventListener("close", () => resolve(answer), { once: true });
    dlg.showModal();
    $("dlg-buttons").querySelector(".primary")?.focus();
  }));
  dialogs = shown.catch(() => {});
  return shown;
}

const alertDialog = (title, text) => dialog(title, text);

/** できたファイルを保存・表示する画面。files: [{bytes, name, label}]（1つ目が主なファイル） */
async function doneDialog(title, text, files) {
  const urls = files.map((f) => blobUrl(f.bytes, f.name.endsWith(".pdf") ? "application/pdf" : "application/zip"));
  // 並びは「閉じる・開いて見る・（ほかのファイルの保存）・主なファイルの保存」。主なボタンを右端に置く
  const saves = files.map((f, k) => ({ label: f.label, href: urls[k], download: f.name, primary: k === 0, keepOpen: true }));
  const buttons = [{ label: "閉じる" }, ...(files[0].name.endsWith(".pdf") ? [{ label: "開いて見る", href: urls[0], newTab: true, keepOpen: true }] : []),
    ...saves.slice(1), saves[0]];
  try {
    await dialog(title, text, buttons);
  } finally {
    setTimeout(() => urls.forEach((u) => URL.revokeObjectURL(u)), 60_000);  // 保存が終わるのを待ってから片付ける
  }
}

function outside(e) {
  const r = e.currentTarget.getBoundingClientRect();
  return e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom;
}

function busy(button, on, text) {
  if (on) {
    button.dataset.label = button.textContent;
    button.textContent = text;
  } else if (button.dataset.label) {
    button.textContent = button.dataset.label;
  }
  button.disabled = on;
}

// ---------- ファイルを選ぶ・ドラッグする ----------

$("b-open").addEventListener("click", () => $("f-pdf").click());
$("f-pdf").addEventListener("change", (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (file) readPdf(file);
});

const hasFiles = (e) => [...(e.dataTransfer?.types ?? [])].includes("Files");
let dragDepth = 0;
addEventListener("dragenter", (e) => {
  if (hasFiles(e) && dragDepth++ === 0) document.body.classList.add("dragging");
});
addEventListener("dragleave", (e) => {
  if (hasFiles(e) && --dragDepth === 0) document.body.classList.remove("dragging");
});
addEventListener("dragover", (e) => {
  if (hasFiles(e)) e.preventDefault();
});
addEventListener("drop", (e) => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  document.body.classList.remove("dragging");
  const file = e.dataTransfer.files[0];
  if (!file) return;
  const row = e.target.closest?.(".pick");
  if (row && /\.xls[xm]$/i.test(file.name)) return setExcel(row, file);
  if (/\.pdf$/i.test(file.name)) return readPdf(file);
  alertDialog("このファイルは使えません",
    `${file.name}\n\nカット絵はPDFを画面にドラッグします。Excel（.xlsx）は ③ の香盤表・カット表などの行にドラッグします。`);
});

// 読み取ったのに書き出さずにページを閉じようとしたら、確かめる
addEventListener("beforeunload", (e) => {
  if (unsaved && results.length) e.preventDefault();
});

// ---------- 使い方の案内（画面に重ねる吹き出し） ----------

const GAP = 46;  // ボタンと吹き出しの間
const SVG = "http://www.w3.org/2000/svg";

/** [囲む部品, 文, 吹き出しの位置 "below"/"above"/"left", 横のずらし, 縦のずらし]（gui.py の show_guide と同じ内容） */
const guideTips = () => [
  [[$("b-open")], "① 描いたカット絵PDFを読み取ります → ② 黄色の「要確認」の行だけ確認します", "below", 130, 0],
  [[$("sheet-format"), $("b-sheet")], "最初に用紙を作ります。左で「横 / 縦 9:16」を選んでから。印刷は実際のサイズ（100%）で", "below", -40, 0],
  [[$("b-samples")], "初めてなら、ここの記入例で一度通してみるのがおすすめです", "below", 0, 120],
  [[$("excel-box")], "③ 香盤表・カット表のExcelを選びます（省略可）", "above", -150, 0],
  [[$("b-export")], "④ 絵コンテPDFを書き出します", "left", 0, 0],
];

function svgEl(tag, attrs) {
  const node = document.createElementNS(SVG, tag);
  for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, v);
  return node;
}

const clamp = (v, lo, hi) => Math.min(Math.max(v, lo), Math.max(lo, hi));

function roundRect({ left: x, top: y, right: x2, bottom: y2 }, r) {
  return `M${x + r} ${y}H${x2 - r}A${r} ${r} 0 0 1 ${x2} ${y + r}V${y2 - r}A${r} ${r} 0 0 1 ${x2 - r} ${y2}`
    + `H${x + r}A${r} ${r} 0 0 1 ${x} ${y2 - r}V${y + r}A${r} ${r} 0 0 1 ${x + r} ${y}Z`;
}

function drawGuide() {
  const g = $("guide");
  if (g.hidden) return;
  // ページ全体に重ねる（狭い画面でページが縦に長くても、スクロールすれば全部の案内が見える）
  Object.assign(g.style, { width: "0", height: "0" });
  const W = document.documentElement.scrollWidth;
  const H = document.documentElement.scrollHeight;
  Object.assign(g.style, { width: `${W}px`, height: `${H}px` });
  const svg = svgEl("svg", { width: W, height: H, viewBox: `0 0 ${W} ${H}` });
  g.replaceChildren(svg);
  const tips = guideTips().map(([targets, text, side, dx, dy]) => {
    const rects = targets.map((t) => t.getBoundingClientRect());
    const r = {
      left: Math.min(...rects.map((x) => x.left)) + scrollX - 5, top: Math.min(...rects.map((x) => x.top)) + scrollY - 5,
      right: Math.max(...rects.map((x) => x.right)) + scrollX + 5, bottom: Math.max(...rects.map((x) => x.bottom)) + scrollY + 5,
    };
    return { r, text, side, dx, dy };
  });
  // 案内するボタンのところだけ明るく残す
  svg.append(svgEl("path", { d: `M0 0H${W}V${H}H0Z` + tips.map(({ r }) => roundRect(r, 6)).join(""), "fill-rule": "evenodd", fill: "rgba(0,0,0,.66)" }));
  const placed = [];
  for (const { r, text, side, dx, dy } of tips) {
    svg.append(svgEl("rect", { x: r.left, y: r.top, width: r.right - r.left, height: r.bottom - r.top, rx: 6, fill: "none", stroke: "#ffd54a", "stroke-width": 2 }));
    const box = el("div", { className: "tip", textContent: text });
    g.append(box);
    const w = box.offsetWidth;
    const h = box.offsetHeight;
    const mid = { x: (r.left + r.right) / 2, y: (r.top + r.bottom) / 2 };
    let left = clamp(side === "left" ? r.left - GAP - w : mid.x + dx - w / 2, 12, W - w - 12);
    let top = side === "below" ? r.bottom + GAP + dy : side === "above" ? r.top - GAP - dy - h : mid.y + dy - h / 2;
    top = clamp(top, 12, H - h - 12);
    // 先に置いた吹き出しと重なったら、下（上向きの吹き出しは上）にずらす
    for (let k = 0; k < placed.length * 2; k++) {
      const hit = placed.find((p) => left < p.right + 8 && left + w > p.left - 8 && top < p.bottom + 8 && top + h > p.top - 8);
      if (!hit) break;
      top = side === "above" ? hit.top - h - 10 : hit.bottom + 10;
    }
    placed.push({ left, top, right: left + w, bottom: top + h });
    box.style.left = `${left}px`;
    box.style.top = `${top}px`;
    const [start, end] = side === "below" ? [[left + w / 2, top], [mid.x, r.bottom + 3]]
      : side === "above" ? [[left + w / 2, top + h], [mid.x, r.top - 3]]
        : [[left + w, top + h / 2], [r.left - 3, mid.y]];
    const angle = Math.atan2(end[1] - start[1], end[0] - start[0]);
    const head = [0.45, -0.45].map((a) => `${end[0] - 12 * Math.cos(angle + a)},${end[1] - 12 * Math.sin(angle + a)}`);
    svg.append(svgEl("line", { x1: start[0], y1: start[1], x2: end[0], y2: end[1], stroke: "#ffd54a", "stroke-width": 2.5 }));
    svg.append(svgEl("polygon", { points: `${end.join(",")} ${head.join(" ")}`, fill: "#ffd54a" }));
  }
  g.append(el("div", { className: "title", textContent: "使い方　—　どこかをクリックすると閉じます" }));
}

function showGuide() {
  const g = $("guide");
  scrollTo(0, 0);
  g.hidden = false;
  drawGuide();
  const close = () => {
    g.hidden = true;
    g.replaceChildren();
    removeEventListener("resize", drawGuide);
    removeEventListener("keydown", onKey);
  };
  const onKey = (e) => {
    if (e.key === "Escape") close();
  };
  g.onclick = close;
  addEventListener("resize", drawGuide);
  addEventListener("keydown", onKey);
}

$("b-guide").addEventListener("click", showGuide);
// 初めて開いたときだけ、案内を重ねる（あとからは「？ 使い方」で出せる）
if (!store.get("guideShown")) {
  store.set("guideShown", "1");
  requestAnimationFrame(showGuide);
}
