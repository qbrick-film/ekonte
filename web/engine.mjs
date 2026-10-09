// ブラウザ版の計算部分。ブラウザの中で動く Python（Pyodide）に ekonte を読み込み、JavaScript から呼べるようにする。
// worker.mjs（ブラウザ）と test_engine.mjs（Node での動作確認）の両方から使う。
import { loadPyodide } from "./vendor/pyodide/pyodide.mjs";

// 読み込む Python の部品。vendor/pyodide/pyodide-lock.json に指紋（sha256）つきで載せてあり、一致しなければ読み込まれない
const PACKAGES = ["numpy", "pillow", "pypdfium2", "reportlab", "openpyxl", "defusedxml"];

/**
 * @param base       このファイルの置き場所（ブラウザは URL、Node はフォルダのパス。末尾は /）
 * @param fetchBytes 置き場所からの相対パスを受け取り、ファイルの中身を Uint8Array で返す関数
 * @param onStatus   準備の進み具合を知らせる関数
 */
export async function startEngine({ base, fetchBytes, onStatus = () => {} }) {
  onStatus("Python を準備しています");
  const py = await loadPyodide({ indexURL: base + "vendor/pyodide/" });
  onStatus("部品を読み込んでいます");
  await py.loadPackage(PACKAGES, { messageCallback: () => {} });
  onStatus("絵コンテ作成ソフトを読み込んでいます");
  py.unpackArchive(await fetchBytes("app.zip"), "zip", { extractDir: "/app" });
  py.runPython("import sys; sys.path.insert(0, '/app')");
  return new Engine(py.pyimport("ekonte.web"));
}

class Engine {
  constructor(web) {
    this.web = web;
  }

  /** カット絵PDFを読み取る。1ページごとに onPage({page, total, name, status, warnings, orientation, format, picture, mark}) */
  read(pdf, onPage) {
    return call(() => this.web.read(pdf, (info, picture, mark) => {
      onPage({ ...JSON.parse(info), picture: copyBytes(picture), mark: copyBytes(mark) });
    }));
  }

  /** Excelを読み込んで確かめ、要約（「12シーン」など）を返す */
  setExcel(key, data, ext) {
    return call(() => this.web.set_excel(key, data, ext));
  }

  clearExcel(key) {
    call(() => this.web.clear_excel(key));
  }

  /** カット表とカット絵の突き合わせ → {missing, extra} */
  check(names) {
    return JSON.parse(call(() => this.web.check(JSON.stringify(names))));
  }

  /** 絵コンテPDFを作る → {pages, cuts, warnings, pdf, images} */
  exportPdf(rows, fmt, withImages) {
    const result = JSON.parse(call(() => this.web.export(JSON.stringify(rows), fmt, withImages)));
    return { ...result, pdf: takeBytes(call(() => this.web.output("pdf"))), images: takeBytes(call(() => this.web.output("images"))) };
  }

  /** カット絵用紙のPDF */
  sheet(pages, fmt) {
    return takeBytes(call(() => this.web.sheet(pages, fmt)));
  }

  /** 動作確認用：部品の版など */
  info() {
    return JSON.parse(call(() => this.web.info()));
  }
}

// Python の bytes を Uint8Array に写す（空なら null）。copyBytes は呼び出しの間だけ使える引数用、takeBytes は戻り値用
function copyBytes(proxy) {
  const bytes = proxy.toJs();
  return bytes.length ? bytes : null;
}

function takeBytes(proxy) {
  try {
    return copyBytes(proxy);
  } finally {
    proxy.destroy();
  }
}

// Python で起きたエラーは「Traceback ...」の全文なので、画面には最後の行の理由だけを出す（全文は detail に残す）
function call(fn) {
  try {
    return fn();
  } catch (e) {
    if (!e?.type) throw e;
    const last = String(e.message).trim().split("\n").pop();
    const err = new Error(last.replace(/^[A-Za-z_][\w.]*: /, "") || last);
    err.detail = e.message;
    throw err;
  }
}
