// Roda o Python (Pyodide) fora da tela principal, para a página não travar durante o preenchimento.
importScripts("https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.js");

const PY = `
import io, json, shutil, sys, zipfile
from pathlib import Path
sys.path.insert(0, "/home/pyodide")
import preencher as P

BASE = Path("/work")

def limpar(tipo):
    d = BASE / tipo
    shutil.rmtree(d, ignore_errors=True)
    d.mkdir(parents=True)

def _arquivos(tipo):
    return sorted(p for p in (BASE / tipo).iterdir() if p.is_file())

def parear_js():
    pares, problemas = P.parear(_arquivos("origens"), _arquivos("destinos"))
    return json.dumps({
        "pares": [{"destino": d.name, "origens": [o.name for o in os]} for d, os in pares.items()],
        "problemas": [{"arquivo": a, "motivo": m} for a, m in problemas],
    })

def executar_js(progresso):
    shutil.rmtree(BASE / "destinos" / "preenchidos", ignore_errors=True)
    res = P.preencher_lote(_arquivos("origens"), _arquivos("destinos"), on_progress=progresso)
    saidas = sorted((BASE / "destinos" / "preenchidos").glob("*.xlsx"))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for s in saidas:
            z.write(s, s.name)
    (BASE / "preenchidos.zip").write_bytes(buf.getvalue())
    return json.dumps({"resultados": res, "saidas": [str(s) for s in saidas]}, default=str)
`;

let py;
const pronto = (async () => {
  py = await loadPyodide();
  await py.loadPackage("micropip");
  await py.pyimport("micropip").install("openpyxl");
  const src = await (await fetch("preencher.py", { cache: "no-store" })).text();
  py.FS.writeFile("/home/pyodide/preencher.py", src);
  await py.runPythonAsync(PY);
  postMessage({ tipo: "pronto" });
})().catch((e) => postMessage({ tipo: "falha", msg: String(e) }));

onmessage = async ({ data }) => {
  await pronto;
  try {
    if (data.tipo === "arquivos") {
      py.globals.get("limpar")(data.qual);
      for (const f of data.arquivos) py.FS.writeFile(`/work/${data.qual}/${f.nome}`, new Uint8Array(f.bytes));
      if (data.parear) postMessage({ tipo: "pares", ...JSON.parse(py.globals.get("parear_js")()) });
    } else if (data.tipo === "executar") {
      const cb = (arq, n, atual, total, msg) => postMessage({ tipo: "prog", arq, n, atual, total, msg });
      const out = JSON.parse(py.globals.get("executar_js")(cb));
      const arquivos = out.saidas.map((p) => ({ nome: p.split("/").pop(), bytes: py.FS.readFile(p) }));
      postMessage({ tipo: "fim", resultados: out.resultados, arquivos, zip: py.FS.readFile("/work/preenchidos.zip") });
    }
  } catch (e) {
    postMessage({ tipo: "erro", msg: String(e) });
  }
};
