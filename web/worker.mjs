// 画面（app.mjs）から頼まれた計算を、画面とは別の流れ（Web Worker）で行う。読み取りに時間がかかっても画面が固まらないようにするため。
import { startEngine } from "./engine.mjs";

const version = new URL(self.location.href).searchParams.get("v") ?? "";
const base = new URL("./", self.location.href).href;

async function fetchBytes(name) {
  const res = await fetch(new URL(`${name}?v=${encodeURIComponent(version)}`, base));
  if (!res.ok) throw new Error(`${name} を読み込めません（${res.status}）`);
  return new Uint8Array(await res.arrayBuffer());
}

const ready = startEngine({ base, fetchBytes, onStatus: (text) => postMessage({ event: "status", text }) });
ready.then(
  () => postMessage({ event: "ready" }),
  (e) => {
    console.error(e?.detail ?? e);
    postMessage({ event: "fatal", error: String(e?.message ?? e) });
  },
);

// 頼まれごと: 名前 → (engine, 引数, 途中経過を知らせる関数) => 結果
const COMMANDS = {
  read: (engine, { pdf }, notify) => engine.read(pdf, (page) => notify({ page }, [page.picture?.buffer, page.mark?.buffer])),
  setExcel: (engine, { key, data, ext }) => engine.setExcel(key, data, ext),
  clearExcel: (engine, { key }) => engine.clearExcel(key),
  check: (engine, { names }) => engine.check(names),
  export: (engine, { rows, fmt, withImages }) => engine.exportPdf(rows, fmt, withImages),
  sheet: (engine, { pages, fmt }) => engine.sheet(pages, fmt),
};

self.onmessage = async ({ data: { id, cmd, args } }) => {
  try {
    const engine = await ready;
    const notify = (data, transfer = []) => postMessage({ id, event: "progress", ...data }, transfer.filter(Boolean));
    const result = await COMMANDS[cmd](engine, args, notify);
    postMessage({ id, ok: true, result }, buffersOf(result));
  } catch (e) {
    console.error(e?.detail ?? e);
    postMessage({ id, ok: false, error: String(e?.message ?? e) });
  }
};

// 結果に含まれる Uint8Array は、写さずにそのまま画面へ渡す
function buffersOf(result) {
  if (!result || typeof result !== "object") return [];
  return Object.values(result).filter((v) => v instanceof Uint8Array).map((v) => v.buffer);
}
