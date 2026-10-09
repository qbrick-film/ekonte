// ブラウザ版の計算部分を Node で動かし、記入例で 読み取り → 絵コンテ書き出し → 用紙作成 を行う（GitHub の自動ビルドと手元の確認用）。
//   node web/test_engine.mjs dist-web 出力フォルダ
// できたPDFと読み取り結果は、compare.py でデスクトップ版と同じ Python の処理の結果と比べる。
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

const [dist, out] = process.argv.slice(2).map((p) => path.resolve(p));
fs.mkdirSync(out, { recursive: true });
const load = (name) => new Uint8Array(fs.readFileSync(path.join(dist, name)));
const sample = (name) => load(`samples/${name}`);
const start = performance.now();
const lap = (label) => console.log(`${label}（${((performance.now() - start) / 1000).toFixed(1)}秒）`);

const { startEngine } = await import(pathToFileURL(path.join(dist, "engine.mjs")).href);
const engine = await startEngine({ base: dist + "/", fetchBytes: load, onStatus: (s) => console.log(`  ${s}`) });
const info = engine.info();
lap(`準備できました: Python ${info.python} / numpy ${info.numpy} / Pillow ${info.pillow} / pypdfium2 ${info.pypdfium2}`);

const excel = {
  kouban: engine.setExcel("kouban", sample("香盤表_記入例.xlsx"), ".xlsx"),
  cuts: engine.setExcel("cuts", sample("カット表_記入例.xlsx"), ".xlsx"),
};
lap(`Excel: 香盤表 ${excel.kouban} / カット表 ${excel.cuts}`);

const cases = [];
for (const [fmt, pdf] of [["横", "カット絵_記入例_スキャン.pdf"], ["縦", "カット絵_記入例_縦_スキャン.pdf"]]) {
  const pages = [];
  engine.read(sample(pdf), (page) => pages.push(page));
  lap(`${fmt}: ${pages.length}ページ読み取り`);
  // 番号が読めたページをすべて入れる（デスクトップ版の selftest と同じ）
  const rows = pages.flatMap((p, i) => (p.name ? [[i, p.name]] : []));
  const format = pages.find((p) => p.status !== "エラー")?.format ?? "横";
  const check = engine.check(rows.map(([, name]) => name));
  const r = engine.exportPdf(rows, format, true);
  const output = `絵コンテ_${fmt}.pdf`;
  fs.writeFileSync(path.join(out, output), r.pdf);
  fs.writeFileSync(path.join(out, `カット絵画像_${fmt}.zip`), r.images);
  lap(`${fmt}: ${r.cuts}カット / ${r.pages}ページ書き出し（警告 ${r.warnings.length}件）`);
  const previews = pages.every((p) => p.picture?.length && (p.status === "エラー" || p.mark?.length));
  cases.push({ pdf, format, output, rows, check, previews, result: { pages: r.pages, cuts: r.cuts, warnings: r.warnings },
    pages: pages.map(({ picture, mark, ...p }) => p) });
}

const sheets = [];
for (const fmt of ["横", "縦"]) {
  const output = `用紙_${fmt}.pdf`;
  fs.writeFileSync(path.join(out, output), engine.sheet(2, fmt));
  sheets.push({ fmt, pages: 2, output });
}
lap("用紙を作成");

// 読めないファイルを選んだときは、理由の分かるエラーになること
const errors = {};
const expectError = (label, fn) => {
  try {
    fn();
    errors[label] = null;
  } catch (e) {
    errors[label] = e.message;
  }
  console.log(`  ${label} → ${errors[label] ?? "（エラーにならなかった）"}`);
};
expectError("Excelではないファイルを選ぶ", () => engine.setExcel("kouban", sample("カット絵_記入例.pdf"), ".xlsx"));
expectError("カット表ではないExcelをカット表に選ぶ", () => engine.setExcel("cuts", sample("香盤表_記入例.xlsx"), ".xlsx"));
// 読めなかったときは、前に選んだ香盤表・カット表が残っていること（最後に読み取った縦のカット絵で書き出す）
const last = cases.at(-1);
const kept = engine.exportPdf(last.rows, last.format, false);
const keptOk = kept.pages === last.result.pages && kept.cuts === last.result.cuts;
lap(`前の香盤表・カット表のまま書き出し: ${kept.cuts}カット / ${kept.pages}ページ`);
// PDFでないものを読み取ると、前の読み取り結果は消える（画面も同じく表を空にしてから読み取る）
expectError("PDFではないファイルを読み取る", () => engine.read(sample("香盤表_記入例.xlsx"), () => {}));

fs.writeFileSync(path.join(out, "results.json"), JSON.stringify({ info, excel, cases, sheets, errors, keptOk }, null, 1));
