#!/usr/bin/env python3
"""
Exporta a classificação IBS/CBS/IS da planilha
Classificacao_Produtos_Reforma_Tributaria_IBS_CBS_IS.xlsx para arquivos de importação:

  classificacao_ibs_cbs.csv                CSV ';' com cabeçalho, decimal ',', UTF-8
  classificacao_ibs_cbs_delimitado.txt     TXT '|campo|campo|' (padrão SPED), sem cabeçalho, UTF-8
  classificacao_ibs_cbs_posicional.txt     TXT de largura fixa (207 posições), ANSI (Windows-1252)
  pendencias.csv                           itens NÃO exportados e o motivo

Somente itens válidos (mesmas regras do Status OK da aba Importacao) são exportados.
O leiaute de cada campo está documentado na aba "Layout" da planilha.

Uso:
  python exportar_classificacao.py [planilha.xlsx] [--saida PASTA] [opções]

Requer: Python 3.8+ e openpyxl (pip install openpyxl).
"""
import argparse
import csv
import os
import sys
import unicodedata

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Instale o openpyxl:  pip install openpyxl")

PLANILHA_PADRAO = "Classificacao_Produtos_Reforma_Tributaria_IBS_CBS_IS.xlsx"

# (campo, largura no TXT posicional, tipo)  N = numérico (zeros à esquerda), X = alfanumérico
LEIAUTE = [
    ("COD_ITEM", 20, "X"), ("DESCRICAO", 120, "X"), ("TIPO_ITEM", 1, "X"), ("NCM", 8, "X"),
    ("NBS", 9, "X"), ("CEST", 7, "X"), ("UNIDADE", 6, "X"), ("GTIN", 14, "X"),
    ("CST_IBS_CBS", 3, "N"), ("CCLASSTRIB", 6, "N"), ("PRED_IBS", 5, "N"), ("PRED_CBS", 5, "N"),
    ("SUJEITO_IS", 1, "X"), ("COD_CATEGORIA_IS", 2, "N"),
]
TAM_LINHA = sum(w for _, w, _ in LEIAUTE)  # 207

CATEGORIAS_IS = [
    "Veículos", "Embarcações e aeronaves", "Produtos fumígenos", "Bebidas alcoólicas",
    "Bebidas açucaradas", "Bens minerais", "Concursos de prognósticos e fantasy sport",
]


def txt(v):
    """Valor de célula -> texto limpo (sem ';', '|' e quebras de linha)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).replace(";", ",").replace("|", "/").replace("\r", " ").replace("\n", " ")
    return " ".join(s.split())


def codigo(v, tam):
    s = txt(v)
    return s.zfill(tam) if s.isdigit() else s


def perc(v):
    """Aceita 0,6 / 60 / '60%' e devolve fração (0.6)."""
    if v in (None, ""):
        return 0.0
    if isinstance(v, str):
        v = float(v.replace("%", "").replace(",", ".").strip() or 0)
    v = float(v)
    return v / 100 if v > 1 else v


def ler_tabelas(wb):
    cst = set()
    ws = wb["Tabela_CST"]
    for (a,) in ws.iter_rows(min_row=3, max_col=1, values_only=True):
        c = codigo(a, 3)
        if c.isdigit():
            cst.add(c)
    cct = {}
    ws = wb["Tabela_cClassTrib"]
    for row in ws.iter_rows(min_row=3, max_col=9, values_only=True):
        c = codigo(row[0], 6)
        if len(c) == 6 and c.isdigit():
            cct[c] = {"pred_ibs": perc(row[3]), "pred_cbs": perc(row[4]), "aplicacao": txt(row[7]) or "Ambos"}
    return cst, cct


def validar(it, cst_set, cct):
    """Mesmas regras da coluna INCONSISTENCIAS da aba Importacao."""
    e = []
    if not it["cod"]:
        e.append("Código vazio")
    if not it["desc"]:
        e.append("Descrição vazia")
    if it["tipo"] not in ("Bem", "Serviço"):
        e.append("Tipo deve ser Bem ou Serviço")
    if it["tipo"] == "Bem" and not (len(it["ncm"]) == 8 and it["ncm"].isdigit()):
        e.append("NCM deve ter 8 dígitos")
    if it["tipo"] == "Serviço" and not (len(it["nbs"]) == 9 and it["nbs"].isdigit()):
        e.append("NBS deve ter 9 dígitos")
    if not it["cst"]:
        e.append("CST vazio")
    elif it["cst"] not in cst_set:
        e.append("CST inválido")
    if not it["cct"]:
        e.append("cClassTrib vazio")
    elif it["cct"] not in cct:
        e.append("cClassTrib fora da tabela")
    if it["cct"][:3] != it["cst"]:
        e.append("cClassTrib incompatível com o CST")
    if it["is"] == "S" and not it["cat_is"]:
        e.append("Informar categoria do IS")
    if it["is"] != "S" and it["cat_is"]:
        e.append("Categoria IS preenchida sem IS = S")
    if it["cat_is"] and it["cat_is"] not in CATEGORIAS_IS:
        e.append("Categoria IS inválida")
    ref = cct.get(it["cct"])
    if ref and ref["aplicacao"] not in ("Ambos", it["tipo"]):
        e.append(f"Aviso: cClassTrib previsto para {ref['aplicacao']}")
    return e


def ler_itens(wb):
    ws = wb["Importacao"]
    for n, r in enumerate(ws.iter_rows(min_row=2, max_col=13, values_only=True), start=2):
        it = {
            "linha": n, "cod": txt(r[0]), "desc": txt(r[1]), "tipo": txt(r[2]),
            "ncm": txt(r[3]), "nbs": txt(r[4]), "cest": txt(r[5]),
            "un": txt(r[6]), "gtin": txt(r[7]), "cst": codigo(r[8], 3), "cct": codigo(r[9], 6),
            "is": txt(r[10]).upper(), "cat_is": txt(r[11]),
        }
        if not (it["cod"] or it["desc"] or it["cct"]):
            continue
        yield it


def campos(it, ref, posicional, decimal):
    def p(v):
        n = round(v * 10000)
        return f"{n:05d}" if posicional else f"{n // 100}{decimal}{n % 100:02d}"
    cat = f"{CATEGORIAS_IS.index(it['cat_is']) + 1:02d}" if it["cat_is"] else ("00" if posicional else "")
    return [
        it["cod"], it["desc"], "S" if it["tipo"] == "Serviço" else "B", it["ncm"], it["nbs"],
        it["cest"], it["un"], it["gtin"], it["cst"], it["cct"],
        p(ref["pred_ibs"]), p(ref["pred_cbs"]), "S" if it["is"] == "S" else "N", cat,
    ]


def linha_posicional(vals, sem_acentos):
    out = []
    for v, (nome, w, tipo) in zip(vals, LEIAUTE):
        if sem_acentos:
            v = unicodedata.normalize("NFKD", v).encode("ascii", "ignore").decode()
        v = v[:w]
        out.append(v.rjust(w, "0") if (tipo == "N" and v) else v.ljust(w))
    s = "".join(out)
    assert len(s) == TAM_LINHA, (nome, len(s))
    return s


def main():
    ap = argparse.ArgumentParser(description="Exporta a classificação IBS/CBS/IS para CSV e TXT.")
    ap.add_argument("planilha", nargs="?", default=PLANILHA_PADRAO)
    ap.add_argument("--saida", default=None, help="pasta de saída (padrão: ./exportacao ao lado da planilha)")
    ap.add_argument("--decimal", default=",", help="separador decimal no CSV/TXT delimitado (padrão ',')")
    ap.add_argument("--enc-csv", default="utf-8", help="codificação do CSV (ex.: utf-8, utf-8-sig, cp1252)")
    ap.add_argument("--enc-txt", default="utf-8", help="codificação do TXT delimitado")
    ap.add_argument("--enc-pos", default="cp1252", help="codificação do TXT posicional (padrão ANSI cp1252)")
    ap.add_argument("--sem-cabecalho", action="store_true", help="não grava a linha de cabeçalho no CSV")
    ap.add_argument("--sem-acentos", action="store_true", help="remove acentos no TXT posicional")
    ap.add_argument("--lf", action="store_true", help="quebra de linha LF (padrão CRLF/Windows)")
    a = ap.parse_args()

    if not os.path.exists(a.planilha):
        sys.exit(f"Planilha não encontrada: {a.planilha}")
    saida = a.saida or os.path.join(os.path.dirname(os.path.abspath(a.planilha)), "exportacao")
    os.makedirs(saida, exist_ok=True)
    nl = "\n" if a.lf else "\r\n"

    wb = load_workbook(a.planilha, read_only=True)
    cst_set, cct = ler_tabelas(wb)
    ok, pend = [], []
    for it in ler_itens(wb):
        erros = validar(it, cst_set, cct)
        (pend if erros else ok).append((it, erros))
    wb.close()

    f_csv = os.path.join(saida, "classificacao_ibs_cbs.csv")
    f_del = os.path.join(saida, "classificacao_ibs_cbs_delimitado.txt")
    f_pos = os.path.join(saida, "classificacao_ibs_cbs_posicional.txt")
    f_pen = os.path.join(saida, "pendencias.csv")

    with open(f_csv, "w", encoding=a.enc_csv, errors="replace", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator=nl, quoting=csv.QUOTE_MINIMAL)
        if not a.sem_cabecalho:
            w.writerow([c for c, _, _ in LEIAUTE])
        for it, _ in ok:
            w.writerow(campos(it, cct[it["cct"]], False, a.decimal))
    with open(f_del, "w", encoding=a.enc_txt, errors="replace", newline="") as f:
        for it, _ in ok:
            f.write("|" + "|".join(campos(it, cct[it["cct"]], False, a.decimal)) + "|" + nl)
    with open(f_pos, "w", encoding=a.enc_pos, errors="replace", newline="") as f:
        for it, _ in ok:
            f.write(linha_posicional(campos(it, cct[it["cct"]], True, a.decimal), a.sem_acentos) + nl)
    with open(f_pen, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator=nl)
        w.writerow(["LINHA_PLANILHA", "COD_ITEM", "DESCRICAO", "CST", "CCLASSTRIB", "PENDENCIAS"])
        for it, e in pend:
            w.writerow([it["linha"], it["cod"], it["desc"], it["cst"], it["cct"], "; ".join(e)])

    print(f"Itens exportados: {len(ok)} | Pendentes (não exportados): {len(pend)}")
    for p in (f_csv, f_del, f_pos, f_pen):
        print("  ->", p)
    if pend:
        print("Corrija as pendências na aba Importacao e execute novamente.")


if __name__ == "__main__":
    main()
