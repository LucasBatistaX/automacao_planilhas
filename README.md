# automacao_planilhas

Automação de preenchimento de planilhas: copia as notas com DARF das planilhas **UneCont - Serviços Tomados** (origem) para os **Controles EFD REINF Mensal** (destino).

**Usar no navegador:** https://lucasbatistax.github.io/automacao_planilhas/

As planilhas são processadas no próprio navegador (Python via [Pyodide](https://pyodide.org)) e não são enviadas para nenhum servidor.

## Como funciona
1. Selecione várias planilhas de origem e várias de destino.
2. O app pareia origem e destino pelo **CNPJ da empresa** (coluna `Cnpj Empresa` da origem × título da linha 2 do destino).
3. Só entram notas com valor na coluna **DARF**. São preenchidas as colunas B a F (Nota, Data, CNPJ e Razão Social do prestador, Valor) com a mesma formatação da linha 6 do template.
4. Cada nota vai para a aba do mês da `Data Competência`. Se a aba não existir, ela é criada com o design da última aba de mês.
5. Mês sem nenhuma nota com DARF recebe "Não houve reinf" na primeira célula de Razão Social.
6. Notas que já estão na aba são puladas. Ao final, baixe os arquivos preenchidos (individualmente ou em .zip).

## Arquivos
- `preencher.py`: toda a regra de preenchimento (usada pela web e pelo app desktop).
- `index.html` + `worker.js`: versão web (GitHub Pages).
- `app.py`: versão desktop (Tkinter), gera o `.app` com PyInstaller.
- `test_preencher.py`: testes; precisam de planilhas de exemplo em `fixtures/` (não versionadas, pois contêm dados de clientes).

## Rodar localmente
```bash
python3 -m http.server 8000   # e abra http://localhost:8000
```
