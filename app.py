"""Interface: escolher as planilhas de origem e destino, conferir os pares, executar e acompanhar."""
import queue
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from preencher import parear, preencher_lote

COR = {
    "fundo": "#F3F5F9", "card": "#FFFFFF", "borda": "#E2E6EE", "texto": "#1F2937",
    "sutil": "#6B7280", "acento": "#2563EB", "acento_hover": "#1D4ED8",
    "ok": "#16A34A", "erro": "#DC2626", "aviso": "#B45309", "log": "#F8FAFC",
}
FONTE = "Helvetica"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Preenchimento EFD REINF")
        self.geometry("820x760")
        self.minsize(700, 640)
        self.configure(bg=COR["fundo"])
        self.fila = queue.Queue()
        self.arquivos = {"origens": [], "destinos": []}
        self.itens = {}  # nome do destino -> id na tabela
        self._estilos()

        raiz = tk.Frame(self, bg=COR["fundo"], padx=24, pady=20)
        raiz.pack(fill="both", expand=True)
        tk.Label(raiz, text="Preenchimento EFD REINF", font=(FONTE, 20, "bold"),
                 bg=COR["fundo"], fg=COR["texto"]).pack(anchor="w")
        tk.Label(raiz, text="Copia as notas com DARF das planilhas UneCont para os Controles REINF, "
                            "pareando pelo CNPJ da empresa.",
                 font=(FONTE, 12), bg=COR["fundo"], fg=COR["sutil"]).pack(anchor="w", pady=(2, 16))

        card = self._card(raiz)
        self._seletor(card, "Planilhas de origem", "UneCont - Serviços Tomados (selecione várias)", "origens")
        tk.Frame(card, height=1, bg=COR["borda"]).pack(fill="x", pady=12)
        self._seletor(card, "Planilhas de destino", "Controle - EFD REINF Mensal (selecione várias)", "destinos")

        pares = self._card(raiz, pady=(12, 0))
        tk.Label(pares, text="Pares encontrados", font=(FONTE, 13, "bold"),
                 bg=COR["card"], fg=COR["texto"]).pack(anchor="w", pady=(0, 8))
        self.tabela = ttk.Treeview(pares, columns=("origem", "destino", "status"), show="headings", height=5)
        for col, titulo, larg in (("origem", "Origem(ns)", 300), ("destino", "Destino", 260), ("status", "Status", 140)):
            self.tabela.heading(col, text=titulo, anchor="w")
            self.tabela.column(col, width=larg, anchor="w")
        for tag, cor in (("ok", COR["ok"]), ("erro", COR["erro"]), ("aviso", COR["aviso"])):
            self.tabela.tag_configure(tag, foreground=cor)
        self.tabela.pack(fill="x")

        botoes = tk.Frame(raiz, bg=COR["fundo"])
        botoes.pack(fill="x", pady=16)
        self.btn = ttk.Button(botoes, text="Executar preenchimento", style="Acento.TButton", command=self._executar)
        self.btn.pack(side="left", fill="x", expand=True)
        self.btn_pasta = ttk.Button(botoes, text="Abrir pasta preenchidos", style="Sec.TButton",
                                    command=self._abrir_saida)
        self.btn_pasta.state(["disabled"])
        self.btn_pasta.pack(side="left", padx=(10, 0), fill="y")

        prog = self._card(raiz, expand=True)
        self.status_geral, self.pct_geral, self.barra_geral = self._barra(prog, "Aguardando planilhas")
        self.status, self.pct, self.barra = self._barra(prog, "Arquivo atual", pequeno=True)
        self.log = tk.Text(prog, height=6, state="disabled", relief="flat", bd=0, wrap="word",
                           font=("Menlo", 11), bg=COR["log"], fg=COR["texto"], padx=10, pady=8,
                           highlightthickness=1, highlightbackground=COR["borda"])
        self.log.pack(fill="both", expand=True, pady=(4, 0))
        for tag in ("ok", "erro", "aviso"):
            self.log.tag_configure(tag, foreground=COR[tag])

    def _estilos(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Acento.TButton", font=(FONTE, 13, "bold"), foreground="white",
                    background=COR["acento"], borderwidth=0, padding=(16, 12), focuscolor=COR["acento"])
        s.map("Acento.TButton", background=[("disabled", "#93B4F5"), ("active", COR["acento_hover"])],
              foreground=[("disabled", "white")])
        s.configure("Sec.TButton", font=(FONTE, 12), foreground=COR["acento"], background=COR["card"],
                    bordercolor=COR["borda"], lightcolor=COR["card"], darkcolor=COR["card"],
                    padding=(12, 6), focuscolor=COR["card"])
        s.map("Sec.TButton", background=[("active", "#EEF3FF")], foreground=[("disabled", COR["sutil"])])
        s.configure("Acento.Horizontal.TProgressbar", troughcolor="#E5E9F2", background=COR["acento"],
                    bordercolor="#E5E9F2", lightcolor=COR["acento"], darkcolor=COR["acento"], thickness=10)
        s.configure("Treeview", font=(FONTE, 11), rowheight=24, background=COR["card"],
                    fieldbackground=COR["card"], bordercolor=COR["borda"])
        s.configure("Treeview.Heading", font=(FONTE, 11, "bold"), background=COR["log"],
                    foreground=COR["sutil"], relief="flat")

    def _card(self, pai, expand=False, pady=0):
        borda = tk.Frame(pai, bg=COR["borda"], padx=1, pady=1)  # moldura de 1px
        borda.pack(fill="both", expand=expand, pady=pady)
        dentro = tk.Frame(borda, bg=COR["card"], padx=18, pady=16)
        dentro.pack(fill="both", expand=True)
        return dentro

    def _barra(self, pai, texto, pequeno=False):
        topo = tk.Frame(pai, bg=COR["card"])
        topo.pack(fill="x")
        f = (FONTE, 11) if pequeno else (FONTE, 12, "bold")
        status = tk.Label(topo, text=texto, font=f, bg=COR["card"], fg=COR["sutil" if pequeno else "texto"], anchor="w")
        status.pack(side="left", fill="x", expand=True)
        pct = tk.Label(topo, text="0%", font=f, bg=COR["card"], fg=COR["acento"])
        pct.pack(side="right")
        barra = ttk.Progressbar(pai, mode="determinate", style="Acento.Horizontal.TProgressbar")
        barra.pack(fill="x", pady=(6, 12))
        return status, pct, barra

    def _seletor(self, pai, titulo, dica, chave):
        linha = tk.Frame(pai, bg=COR["card"])
        linha.pack(fill="x")
        textos = tk.Frame(linha, bg=COR["card"])
        textos.pack(side="left", fill="x", expand=True)
        tk.Label(textos, text=titulo, font=(FONTE, 13, "bold"), bg=COR["card"], fg=COR["texto"]).pack(anchor="w")
        nome = tk.Label(textos, text=dica, font=(FONTE, 11), bg=COR["card"], fg=COR["sutil"], anchor="w")
        nome.pack(anchor="w", fill="x")
        ttk.Button(linha, text="Escolher planilhas", style="Sec.TButton",
                   command=lambda: self._escolher(chave, nome)).pack(side="right")

    def _escolher(self, chave, rotulo):
        caminhos = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx")])
        if not caminhos:
            return
        self.arquivos[chave] = list(caminhos)
        nomes = ", ".join(Path(c).name for c in caminhos)
        rotulo.configure(text=f"✓  {len(caminhos)} planilha(s): " + (nomes if len(nomes) < 70 else nomes[:67] + "..."),
                         fg=COR["ok"])
        if self.arquivos["origens"] and self.arquivos["destinos"]:
            self._mostrar_pares()

    def _mostrar_pares(self):
        self.tabela.delete(*self.tabela.get_children())
        self.itens.clear()
        try:
            pares, problemas = parear(self.arquivos["origens"], self.arquivos["destinos"])
        except Exception as e:
            messagebox.showerror("Erro ao ler as planilhas", str(e))
            return
        for dest, origs in pares.items():
            self.itens[dest.name] = self.tabela.insert(
                "", "end", values=(", ".join(o.name for o in origs), dest.name, "pronto"))
        for arq, motivo in problemas:
            self.tabela.insert("", "end", values=(arq, "-", motivo), tags=("aviso",))
        self.status_geral["text"] = f"{len(pares)} par(es) prontos, {len(problemas)} arquivo(s) serão pulados"

    def _executar(self):
        if not self.arquivos["origens"] or not self.arquivos["destinos"]:
            messagebox.showwarning("Faltam planilhas", "Escolha as planilhas de origem e as de destino.")
            return
        if not self.itens:
            messagebox.showwarning("Nada a fazer", "Nenhum par origem/destino foi encontrado.")
            return
        self.btn.state(["disabled"])
        self.btn_pasta.state(["disabled"])
        self.btn.configure(text="Processando...")
        for b in (self.barra, self.barra_geral):
            b["value"] = 0
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        threading.Thread(target=self._trabalho, daemon=True).start()
        self.after(100, self._atualizar)

    def _trabalho(self):
        try:
            res = preencher_lote(self.arquivos["origens"], self.arquivos["destinos"],
                                 on_progress=lambda *a: self.fila.put(("prog", *a)))
            self.fila.put(("fim", res))
        except Exception as e:  # mostra qualquer erro ao usuário em vez de travar a thread
            self.fila.put(("erro", str(e)))

    def _atualizar(self):
        while not self.fila.empty():
            item = self.fila.get()
            if item[0] == "prog":
                self._progresso(*item[1:])
                continue
            self.btn.state(["!disabled"])
            self.btn.configure(text="Executar preenchimento")
            if item[0] == "fim":
                self._finalizar(item[1])
            else:
                self.status_geral.configure(text="Erro", fg=COR["erro"])
                self._log("ERRO: " + item[1], "erro")
                messagebox.showerror("Erro", item[1])
            return
        self.after(100, self._atualizar)

    def _progresso(self, arq, n_arq, atual, total, msg):
        frac = atual / total if total else 1
        self.barra_geral["value"] = (arq - 1 + frac) / n_arq * 100
        self.pct_geral["text"] = f"{int(self.barra_geral['value'])}%"
        self.status_geral.configure(text=f"Arquivo {arq} de {n_arq}", fg=COR["texto"])
        self.barra["value"] = frac * 100
        self.pct["text"] = f"{int(frac * 100)}%"
        self.status["text"] = msg if len(msg) < 70 else msg[:67] + "..."
        tag = "erro" if msg.startswith("ERRO") else "aviso" if "pulada" in msg or "Não houve" in msg \
            else "ok" if "->" in msg or "Concluído" in msg else ""
        self._log(msg, tag)

    def _finalizar(self, resultados):
        ok = [r for r in resultados if r["status"] == "ok"]
        for r in resultados:
            iid = self.itens.get(r["arquivo"])
            if not iid:
                continue
            if r["status"] == "ok":
                self.tabela.set(iid, "status", f"✓ {r['inseridas']} inserida(s)")
                self.tabela.item(iid, tags=("ok",))
            else:
                self.tabela.set(iid, "status", "✗ erro")
                self.tabela.item(iid, tags=("erro",))
                self._log(f"{r['arquivo']}: {r['motivo']}", "erro")
        pulados = len(resultados) - len(ok)
        notas = sum(r["inseridas"] for r in ok)
        self.barra_geral["value"], self.pct_geral["text"] = 100, "100%"
        self.status_geral.configure(text="Concluído", fg=COR["ok"])
        self.btn_pasta.state(["!disabled"])
        messagebox.showinfo("Concluído", f"{len(ok)} planilha(s) preenchida(s)\n{notas} nota(s) inserida(s)\n"
                                         f"{pulados} arquivo(s) pulado(s) ou com erro")

    def _abrir_saida(self):
        for pasta in {Path(d).parent / "preenchidos" for d in self.arquivos["destinos"]}:
            if pasta.exists():
                subprocess.run(["open", str(pasta)])

    def _log(self, msg, tag=""):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")


if __name__ == "__main__":
    App().mainloop()
