"""Copia notas com DARF da planilha UneCont (origem) para o Controle EFD REINF (destino)."""
import re
from copy import copy
from pathlib import Path

import openpyxl

MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
         "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

# coluna destino -> cabeçalho na origem (só B..F)
MAPA = {
    "B": "Número NFe",
    "C": "Data Competência",
    "D": "Cnpj/Cpf Prestador",
    "E": "Prestador",
    "F": "Valor NFe",
}
COL_CONDICAO = "DARF"
COL_MES = "Data Competência"
COL_CNPJ_EMPRESA = "Cnpj Empresa"
PRIMEIRA_LINHA_DADOS = 6
SEM_REINF = "Não houve reinf"
REF = re.compile(r"(\$?[A-Z]{1,3}\$?)(\d+)(?::(\$?[A-Z]{1,3}\$?)(\d+))?")
CNPJ = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
PERIODO_ARQUIVO = re.compile(r"(\d{4})-(\d{2}) a (\d{4})-(\d{2})")


def _digitos(s):
    return re.sub(r"\D", "", str(s))


def _linhas_origem(caminho):
    """Todas as linhas de nota da origem (dict cabeçalho->valor), sem linhas de total."""
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    try:
        linhas = wb.worksheets[0].iter_rows(values_only=True)
        cab = [str(c).strip() if c is not None else "" for c in next(linhas)]
        faltando = [c for c in list(MAPA.values()) + [COL_CONDICAO] if c not in cab]
        if faltando:
            raise ValueError(f"Colunas não encontradas na origem: {', '.join(faltando)}")
        return [reg for reg in (dict(zip(cab, v)) for v in linhas)
                if reg.get("Número NFe") not in (None, "")]
    finally:
        wb.close()


def _tem_darf(reg):
    return reg.get(COL_CONDICAO) not in (None, "", 0)


def ler_origem(caminho):
    """Linhas da origem que têm valor na coluna DARF."""
    return [r for r in _linhas_origem(caminho) if _tem_darf(r)]


def _meses_da_origem(caminho, linhas):
    """(ano, mês) cobertos pela origem: pelas datas das notas ou, sem notas, pelo nome do arquivo."""
    meses = {(r[COL_MES].year, r[COL_MES].month) for r in linhas if hasattr(r.get(COL_MES), "month")}
    m = PERIODO_ARQUIVO.search(Path(caminho).name)
    if not meses and m:
        a, mm, a2, mm2 = map(int, m.groups())
        while (a, mm) <= (a2, mm2):
            meses.add((a, mm))
            a, mm = (a + 1, 1) if mm == 12 else (a, mm + 1)
    return meses


def cnpj_origem(caminho):
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    try:
        linhas = wb.worksheets[0].iter_rows(values_only=True, max_row=2)
        cab = [str(c).strip() if c is not None else "" for c in next(linhas)]
        if COL_CNPJ_EMPRESA not in cab:
            raise ValueError(f"coluna '{COL_CNPJ_EMPRESA}' não encontrada")
        primeira = next(linhas, None)
        valor = primeira[cab.index(COL_CNPJ_EMPRESA)] if primeira else None
        if not valor:
            raise ValueError("CNPJ da empresa vazio")
        return _digitos(valor)
    finally:
        wb.close()


def cnpj_destino(caminho):
    wb = openpyxl.load_workbook(caminho, read_only=True)
    try:
        for linha in wb.worksheets[0].iter_rows(values_only=True, max_row=5):
            for v in linha:
                m = CNPJ.search(str(v)) if v else None
                if m:
                    return _digitos(m.group())
        raise ValueError("CNPJ não encontrado no título")
    finally:
        wb.close()


def _mes_da_aba(ws):
    titulo = ws.title.strip().upper()
    return next((i for i, m in enumerate(MESES, 1) if titulo.startswith(m)), None)


def _aba_do_mes(wb, ano, mes, criadas=None):
    """Aba do mês; se não existir, cria copiando o design da última aba de mês do arquivo."""
    abas = [ws for ws in wb.worksheets if _mes_da_aba(ws)]
    for ws in abas:
        if _mes_da_aba(ws) == mes:
            return ws
    if not abas:
        raise ValueError("Nenhuma aba de mês no destino para servir de modelo")
    modelo = abas[-1]
    nova = wb.copy_worksheet(modelo)  # copia estilos, larguras, mesclagens e fórmulas
    nova.title = f"{MESES[mes - 1]}-{ano}"
    nova.sheet_view.showGridLines = modelo.sheet_view.showGridLines
    for r in range(PRIMEIRA_LINHA_DADOS, _linha_total(nova)):
        for col in "BCDEF":
            nova[f"{col}{r}"].value = None
        if not str(nova[f"J{r}"].value or "").startswith("="):
            nova[f"J{r}"].value = None  # data de pagamento é do mês do modelo
    # posiciona depois do último mês anterior (ou no início)
    anteriores = [i for i, ws in enumerate(wb.worksheets) if (_mes_da_aba(ws) or 99) < mes]
    wb.move_sheet(nova, (anteriores[-1] + 1 if anteriores else 0) - wb.worksheets.index(nova))
    if criadas is not None:
        criadas.append(nova.title)
    return nova


def _linha_total(ws):
    for r in range(PRIMEIRA_LINHA_DADOS, ws.max_row + 1):
        for col in ("F", "I"):
            v = ws[f"{col}{r}"].value
            if isinstance(v, str) and "SUM(" in v.upper():
                return r
    raise ValueError(f"Linha de TOTAL (SUM) não encontrada na aba {ws.title}")


def _deslocar_formula(formula, inicio, n):
    """Ajusta referências após inserir n linhas na posição `inicio`."""
    def sub(m):
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
        r1 = r1 + n if r1 >= inicio else r1
        if c2 is None:
            return f"{c1}{r1}"
        r2 = int(r2)
        r2 = r2 + n if r2 >= inicio - 1 else r2  # range que terminava na última linha de dados cresce
        return f"{c1}{r1}:{c2}{r2}"
    return REF.sub(sub, formula)


def _inserir_linhas(ws, antes_de, n):
    ws.insert_rows(antes_de, n)
    for row in ws.iter_rows():
        for cel in row:
            if isinstance(cel.value, str) and cel.value.startswith("="):
                cel.value = _deslocar_formula(cel.value, antes_de, n)
    modelo = antes_de - 1
    for r in range(antes_de, antes_de + n):
        ws.row_dimensions[r].height = ws.row_dimensions[modelo].height
        for col in range(2, 12):  # B..K
            src, dst = ws.cell(modelo, col), ws.cell(r, col)
            dst._style = copy(src._style)
            if col >= 10:  # J, K já vêm preenchidos no template
                dst.value = src.value


def _normaliza_nf(v):
    s = str(v).strip()
    return int(s) if s.isdigit() else s


def preencher(origens, destino, on_progress=lambda atual, total, msg: None, saida=None):
    """Preenche o destino com uma ou várias origens e salva uma vez em `saida`
    (padrão: <destino>_preenchido.xlsx). Retorna resumo."""
    if isinstance(origens, (str, Path)):
        origens = [origens]
    on_progress(0, 1, "Lendo origem...")
    notas, meses = [], set()
    for o in origens:
        linhas = _linhas_origem(o)
        meses |= _meses_da_origem(o, linhas)
        notas += [r for r in linhas if _tem_darf(r)]
    total = len(notas)
    on_progress(0, total or 1, f"{total} nota(s) com DARF encontradas. Abrindo destino...")
    wb = openpyxl.load_workbook(destino)
    resumo = {"inseridas": 0, "duplicadas": 0, "linhas_criadas": 0, "sem_reinf": [], "abas_criadas": []}

    for i, nota in enumerate(notas, 1):
        nf = _normaliza_nf(nota["Número NFe"])
        ws = _aba_do_mes(wb, nota[COL_MES].year, nota[COL_MES].month, resumo["abas_criadas"])
        fim = _linha_total(ws)
        existentes = {_normaliza_nf(ws[f"B{r}"].value) for r in range(PRIMEIRA_LINHA_DADOS, fim)
                      if ws[f"B{r}"].value not in (None, "")}
        if nf in existentes:
            resumo["duplicadas"] += 1
            on_progress(i, total, f"NF {nf} já existe em {ws.title.strip()} - pulada")
            continue
        livre = next((r for r in range(PRIMEIRA_LINHA_DADOS, fim) if ws[f"B{r}"].value in (None, "")), None)
        if livre is None:
            _inserir_linhas(ws, fim, 1)
            resumo["linhas_criadas"] += 1
            livre = fim
        for col, campo in MAPA.items():
            cel = ws[f"{col}{livre}"]
            cel.value = nf if col == "B" else nota[campo]
            # mesma formatação (fonte, bordas, R$, data) da 1ª linha de dados do template
            cel._style = copy(ws[f"{col}{PRIMEIRA_LINHA_DADOS}"]._style)
        resumo["inseridas"] += 1
        on_progress(i, total, f"NF {nf} -> {ws.title.strip()} linha {livre}")

    # meses da origem sem nenhuma nota com DARF: marca a 1ª célula de Razão Social
    meses_com_nota = {(n[COL_MES].year, n[COL_MES].month) for n in notas}
    for ano, mes in sorted(meses - meses_com_nota):
        ws = _aba_do_mes(wb, ano, mes, resumo["abas_criadas"])
        r = PRIMEIRA_LINHA_DADOS
        if ws[f"B{r}"].value in (None, "") and ws[f"E{r}"].value in (None, ""):
            ws[f"E{r}"] = SEM_REINF
            resumo["sem_reinf"].append(ws.title.strip())
            on_progress(total or 1, total or 1, f"{ws.title.strip()}: {SEM_REINF}")

    saida = Path(saida) if saida else Path(destino).with_name(Path(destino).stem + "_preenchido.xlsx")
    if resumo["abas_criadas"]:
        on_progress(total, total or 1, "Aba(s) criada(s): " + ", ".join(resumo["abas_criadas"]))
    on_progress(total, total or 1, "Salvando...")
    wb.save(saida)
    resumo["arquivo"] = str(saida)
    on_progress(total or 1, total or 1, f"Concluído: {resumo['inseridas']} inserida(s), "
                                        f"{resumo['duplicadas']} duplicada(s). Salvo em {saida.name}")
    return resumo


def _xlsx(arquivos):
    return [Path(f) for f in arquivos if not Path(f).name.startswith("~$")]


def parear(origens, destinos_arqs):
    """Pareia pelo CNPJ da empresa. Retorna ({destino: [origens]}, [(arquivo, motivo)])."""
    problemas, destinos = [], {}
    for d in _xlsx(destinos_arqs):
        try:
            destinos.setdefault(cnpj_destino(d), []).append(d)
        except Exception as e:
            problemas.append((d.name, f"erro: {e}"))
    pares = {}
    for o in _xlsx(origens):
        try:
            c = cnpj_origem(o)
        except Exception as e:
            problemas.append((o.name, f"erro: {e}"))
            continue
        ds = destinos.get(c, [])
        if not ds:
            problemas.append((o.name, "sem par"))
        elif len(ds) > 1:
            problemas.append((o.name, "conflito: " + ", ".join(x.name for x in ds)))
        else:
            pares.setdefault(ds[0], []).append(o)
    return pares, problemas


def preencher_lote(origens, destinos, on_progress=lambda arq, n_arq, atual, total, msg: None):
    """Preenche todos os pares; cada destino é salvo em <pasta dele>/preenchidos. Um erro não para o lote."""
    pares, problemas = parear(origens, destinos)
    resultados = [{"arquivo": a, "status": "pulado", "motivo": m} for a, m in problemas]
    n = len(pares)
    for i, (dest, origs) in enumerate(pares.items(), 1):
        try:
            saida_dir = dest.parent / "preenchidos"
            saida_dir.mkdir(exist_ok=True)
            r = preencher(origs, dest, lambda a, t, m, i=i: on_progress(i, n, a, t, m),
                          saida=saida_dir / dest.name)
            resultados.append({**r, "saida": r["arquivo"], "arquivo": dest.name, "status": "ok"})
        except Exception as e:  # registra e segue para o próximo destino
            resultados.append({"arquivo": dest.name, "status": "erro", "motivo": str(e)})
            on_progress(i, n, 1, 1, f"ERRO em {dest.name}: {e}")
    return resultados
