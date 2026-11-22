"""Teste com os arquivos reais de exemplo em fixtures/. Rodar: python3 test_preencher.py"""
import tempfile
from pathlib import Path

import openpyxl

from preencher import _deslocar_formula, ler_origem, preencher, preencher_lote

FIX = Path(__file__).parent / "fixtures"
ORIGEM = FIX / "origem.xlsx"
DESTINO = FIX / "destino.xlsx"


def _fixture():
    """NF e CNPJ lidos das planilhas de exemplo (dados de cliente não ficam no código)."""
    wb = openpyxl.load_workbook(ORIGEM, data_only=True)
    ws = wb.worksheets[0]
    cab = [c.value.strip() if isinstance(c.value, str) else c.value for c in ws[1]]
    nf = ler_origem(ORIGEM)[0]["Número NFe"]
    return (int(nf) if str(nf).isdigit() else nf), ws.cell(2, cab.index("Cnpj Empresa") + 1).value


NF, CNPJ_EMPRESA = _fixture()


def test_filtro_darf():
    assert len(ler_origem(ORIGEM)) == 1  # a planilha de exemplo tem 1 nota com DARF


def test_insere_em_destino_limpo_e_preserva_resto():
    tmp = Path(tempfile.mkdtemp())
    limpo = tmp / "limpo.xlsx"
    wb = openpyxl.load_workbook(DESTINO)
    ws = wb["JANEIRO-2025"]
    esperado = [ws[f"{c}6"].value for c in "BCDEF"]
    for c in "BCDEF":
        ws[f"{c}6"] = None
    wb.save(limpo)

    r = preencher(ORIGEM, limpo, saida=tmp / "out.xlsx")
    assert r["inseridas"] == 1 and r["duplicadas"] == 0
    antes = openpyxl.load_workbook(limpo)["JANEIRO-2025"]
    depois = openpyxl.load_workbook(tmp / "out.xlsx")["JANEIRO-2025"]
    assert [depois[f"{c}6"].value for c in "BCDEF"] == esperado
    for row in antes.iter_rows():  # tudo fora de B6:F6 igual, inclusive estilo
        for cel in row:
            d = depois[cel.coordinate]
            if not (cel.row == 6 and cel.column_letter in "BCDEF"):
                assert d.value == cel.value, cel.coordinate
            assert d.style_id == cel.style_id or d.number_format == cel.number_format, cel.coordinate


def test_pula_duplicada():
    r = preencher(ORIGEM, DESTINO, saida=Path(tempfile.mkdtemp()) / "out.xlsx")
    assert r["inseridas"] == 0 and r["duplicadas"] == 1


def test_insere_linha_quando_cheio():
    tmp = Path(tempfile.mkdtemp())
    cheio = tmp / "cheio.xlsx"
    wb = openpyxl.load_workbook(DESTINO)
    ws = wb["JANEIRO-2025"]
    for r in range(6, 10):  # ocupa todas as linhas antes do TOTAL (linha 10)
        ws[f"B{r}"] = 1000 + r
    wb.save(cheio)
    r = preencher(ORIGEM, cheio, saida=tmp / "out.xlsx")
    assert r["inseridas"] == 1 and r["linhas_criadas"] == 1
    ws = openpyxl.load_workbook(tmp / "out.xlsx")["JANEIRO-2025"]
    assert ws["B10"].value == NF
    for c in "BCDEF":  # linha nova com a mesma formatação da linha 6
        assert ws[f"{c}10"].style_id == ws[f"{c}6"].style_id, c
    assert ws["I11"].value == "=SUM(I6:I10)"
    assert ws["H15"].value == "DARF"  # bloco DARF desceu uma linha


def test_deslocar_formula():
    # inserir 1 linha antes da 9: SUM(F6:F8) cresce, refs abaixo descem, acima ficam
    assert _deslocar_formula("=SUM(F6:F8)", 9, 1) == "=SUM(F6:F9)"
    assert _deslocar_formula("=SUM(I6:I9)", 9, 1) == "=SUM(I6:I10)"
    assert _deslocar_formula("=I9", 9, 1) == "=I10"
    assert _deslocar_formula("=J7", 9, 1) == "=J7"
    assert _deslocar_formula("=F7*1.5%", 9, 1) == "=F7*1.5%"


def _origem_modificada(caminho, cnpj, sem_darf=False, mes=None):
    """Cópia da origem com outro CNPJ (data_only: grava valores no lugar das fórmulas)."""
    import datetime
    wb = openpyxl.load_workbook(ORIGEM, data_only=True)
    ws = wb.worksheets[0]
    cab = [c.value.strip() if isinstance(c.value, str) else c.value for c in ws[1]]
    ia, idarf, imes = cab.index("Cnpj Empresa") + 1, cab.index("DARF") + 1, cab.index("Data Competência") + 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, ia).value:
            ws.cell(r, ia).value = cnpj
            if sem_darf:
                ws.cell(r, idarf).value = None
            if mes:
                ws.cell(r, imes).value = datetime.datetime(2026, mes, 15)
    wb.save(caminho)


def _destino_modificado(caminho, cnpj):
    wb = openpyxl.load_workbook(DESTINO)
    wb.worksheets[0]["B2"] = f"OUTRA EMPRESA LTDA - CNPJ {cnpj}"
    for c in "BCDEF":
        wb["JANEIRO-2025"][f"{c}6"] = None
    wb.save(caminho)


def test_lote():
    import shutil
    base = Path(tempfile.mkdtemp())
    po, pd = base / "origens", base / "destinos"
    po.mkdir(), pd.mkdir()
    shutil.copy(DESTINO, pd / "empresa_a.xlsx")                    # já tem a NF -> duplicada
    _destino_modificado(pd / "empresa_b.xlsx", "11.222.333/0001-81")
    shutil.copy(ORIGEM, po / "a_jan.xlsx")
    _origem_modificada(po / "b_jan.xlsx", "11.222.333/0001-81")               # NF entra em JANEIRO
    _origem_modificada(po / "b_fev.xlsx", "11.222.333/0001-81", sem_darf=True, mes=2)  # sem DARF em FEV
    _origem_modificada(po / "c_jan.xlsx", "99.999.999/0001-99")               # sem par
    (po / "quebrado.xlsx").write_bytes(b"isto nao e um xlsx")                 # erro de leitura

    res = {r["arquivo"]: r for r in preencher_lote(sorted(po.iterdir()), sorted(pd.iterdir()))}
    assert res["c_jan.xlsx"]["status"] == "pulado" and res["c_jan.xlsx"]["motivo"] == "sem par"
    assert res["quebrado.xlsx"]["status"] == "pulado" and "erro" in res["quebrado.xlsx"]["motivo"]
    assert res["empresa_a.xlsx"]["status"] == "ok" and res["empresa_a.xlsx"]["duplicadas"] == 1
    b = res["empresa_b.xlsx"]
    assert b["status"] == "ok" and b["inseridas"] == 1 and b["sem_reinf"] == ["FEVEREIRO-2025"]

    out = openpyxl.load_workbook(pd / "preenchidos" / "empresa_b.xlsx")
    assert out["JANEIRO-2025"]["B6"].value == NF
    assert out["FEVEREIRO-2025"]["E6"].value == "Não houve reinf"
    assert out["JANEIRO-2025"]["E6"].value != "Não houve reinf"  # mês com nota não recebe o aviso
    assert openpyxl.load_workbook(pd / "empresa_b.xlsx")["JANEIRO-2025"]["B6"].value is None  # original intacto


def test_sem_reinf_nao_sobrescreve():
    tmp = Path(tempfile.mkdtemp())
    _origem_modificada(tmp / "o.xlsx", CNPJ_EMPRESA, sem_darf=True)
    antes = openpyxl.load_workbook(DESTINO)["JANEIRO-2025"]["E6"].value
    r = preencher(tmp / "o.xlsx", DESTINO, saida=tmp / "out.xlsx")  # JANEIRO já tem nota na linha 6
    assert r["sem_reinf"] == []
    assert openpyxl.load_workbook(tmp / "out.xlsx")["JANEIRO-2025"]["E6"].value == antes


def test_cria_aba_que_nao_existe():
    tmp = Path(tempfile.mkdtemp())
    wb = openpyxl.load_workbook(DESTINO)
    del wb["FEVEREIRO-2025"]
    wb.save(tmp / "sem_fev.xlsx")
    _origem_modificada(tmp / "o.xlsx", CNPJ_EMPRESA, mes=2)  # nota com DARF em fevereiro
    r = preencher(tmp / "o.xlsx", tmp / "sem_fev.xlsx", saida=tmp / "out.xlsx")
    assert r["abas_criadas"] == ["FEVEREIRO-2026"] and r["inseridas"] == 1

    out = openpyxl.load_workbook(tmp / "out.xlsx")
    assert out.sheetnames[1] == "FEVEREIRO-2026"  # logo depois de JANEIRO
    nova, modelo = out["FEVEREIRO-2026"], out.worksheets[-1]
    assert nova["B6"].value == NF and nova["C6"].value.month == 2
    assert nova["J6"].value is None and nova["J7"].value is None  # data de pagamento do modelo limpa
    assert nova["G6"].value == modelo["G6"].value  # códigos mantidos
    assert nova["F9"].value == "=SUM(F6:F8)"
    for cel in ("B2", "B4", "E4", "B6", "F6", "H14"):  # mesmo design
        assert nova[cel].style_id == modelo[cel].style_id, cel
    assert nova.merged_cells.ranges == modelo.merged_cells.ranges
    assert nova.column_dimensions["E"].width == modelo.column_dimensions["E"].width


if __name__ == "__main__":
    for nome, fn in list(globals().items()):
        if nome.startswith("test_"):
            fn()
            print("ok", nome)
