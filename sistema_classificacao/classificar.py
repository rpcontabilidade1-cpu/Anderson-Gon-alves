#!/usr/bin/env python3
"""
SISTEMA DE CLASSIFICAÇÃO DE PRODUTOS — IBS / CBS / IS
EC 132/2023 e LC 214/2025 (com alterações posteriores)

Recebe o cadastro de produtos do cliente (Excel .xlsx/.xlsm, CSV ou XMLs de NF-e),
enquadra cada item (CST IBS/CBS, cClassTrib, reduções, Imposto Seletivo) com base na
base_regras.xlsx e devolve um Excel padronizado para o cliente entregar aos
programadores do ERP.

Uso:
    python classificar.py ARQUIVO_DO_CLIENTE [--cliente "Nome"] [--regime "Lucro Real"]
                          [--base base_regras.xlsx] [--saida PASTA]

    ARQUIVO_DO_CLIENTE pode ser .xlsx, .xlsm, .csv, um .xml de NF-e ou uma pasta com XMLs.
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Instale o openpyxl:  pip install openpyxl")

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE_PADRAO = os.path.join(AQUI, "base_regras.xlsx")
NFE_NS = "{http://www.portalfiscal.inf.br/nfe}"

ST_OK, ST_CONF, ST_REV, ST_NCM = "OK", "A CONFIRMAR", "REVISAR", "REVISAR NCM"
CAT_IS = ["Veículos", "Embarcações e aeronaves", "Produtos fumígenos", "Bebidas alcoólicas",
          "Bebidas açucaradas", "Bens minerais", "Concursos de prognósticos e fantasy sport"]


# =============================================================================== utilidades
def norm(s):
    """minúsculas, sem acentos, espaços simples."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return " ".join(re.sub(r"[^a-z0-9/ ]", " ", s).split())


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return " ".join(str(v).split())


def digitos(v):
    return re.sub(r"\D", "", txt(v))


def lista(v):
    return [x.strip() for x in txt(v).split(";") if x.strip()]


def perc(v):
    if v in (None, ""):
        return 0.0
    if isinstance(v, str):
        v = float(v.replace("%", "").replace(",", ".").strip() or 0)
    v = float(v)
    return v / 100 if v > 1 else v


def cod(v, n):
    d = digitos(v)
    return d.zfill(n) if d else ""


# =============================================================================== base de regras
class Base:
    def __init__(self, caminho):
        if not os.path.exists(caminho):
            sys.exit(f"Base de regras não encontrada: {caminho}\n"
                     f"Gere com: python ferramentas/criar_base_regras.py")
        wb = load_workbook(caminho, read_only=True, data_only=True)
        self.regras = self._linhas(wb["Regras_NCM"])
        for r in self.regras:
            r["_ncm"] = [digitos(x) for x in lista(r["NCM"]) if digitos(x)]
            r["_exc"] = [digitos(x) for x in lista(r["EXCETO_NCM"]) if digitos(x)]
            r["_chave"] = [norm(x) for x in lista(r["PALAVRAS_CHAVE"])]
            r["_excl"] = [norm(x) for x in lista(r["PALAVRAS_EXCLUIR"])]
            r["CST"] = cod(r["CST"], 3)
            r["CCLASSTRIB"] = cod(r["CCLASSTRIB"], 6)
            r["_prio"] = int(r["PRIORIDADE"] or 9)
            r["_conf"] = txt(r["CONFERIDO"]).upper() == "S"
        self.regras_is = self._linhas(wb["Regras_IS"])
        for r in self.regras_is:
            r["_ncm"] = [digitos(x) for x in lista(r["NCM"]) if digitos(x)]
            r["_exc"] = [digitos(x) for x in lista(r["EXCETO_NCM"]) if digitos(x)]
        self.cct = {}
        for r in self._linhas(wb["Tabela_cClassTrib"]):
            c = cod(r["CCLASSTRIB"], 6)
            if len(c) == 6:
                self.cct[c] = {"desc": txt(r["DESCRICAO"]), "ri": perc(r["PRED_IBS"]), "rc": perc(r["PRED_CBS"]),
                               "conf": txt(r.get("CONFERIDO", "S")).upper() != "N"}
        self.cst = {cod(r["CST"], 3): txt(r["DESCRICAO"]) for r in self._linhas(wb["Tabela_CST"]) if r["CST"] is not None}
        self.sin = {}
        for r in self._linhas(wb["Sinonimos_Colunas"]):
            for s in lista(r["CABECALHOS_ACEITOS (separados por ;)"]):
                self.sin[norm(s)] = txt(r["CAMPO"])
        self.cfg = {txt(r["PARAMETRO"]): txt(r["VALOR"]) for r in self._linhas(wb["Config"])}
        wb.close()

    @staticmethod
    def _linhas(ws):
        it = ws.iter_rows(values_only=True)
        cab = [txt(c) for c in next(it)]
        out = []
        for row in it:
            if all(v in (None, "") for v in row):
                continue
            out.append(dict(zip(cab, row)))
        return out


# =============================================================================== leitura do cliente
def detectar_cabecalho(linhas, sin):
    """Procura nas 40 primeiras linhas a que mais se parece com um cabeçalho."""
    melhor, idx = {}, -1
    for i, row in enumerate(linhas[:40]):
        mapa = {}
        for j, v in enumerate(row):
            campo = sin.get(norm(v))
            if campo and campo not in mapa:
                mapa[campo] = j
        if ("DESCRICAO" in mapa or "NCM" in mapa) and len(mapa) > len(melhor):
            melhor, idx = mapa, i
    return idx, melhor


def ler_planilha(caminho, base):
    itens, avisos = [], []
    if caminho.lower().endswith(".csv"):
        with open(caminho, encoding="utf-8-sig", errors="replace") as f:
            amostra = f.read(4096); f.seek(0)
            sep = ";" if amostra.count(";") >= amostra.count(",") else ","
            folhas = {"CSV": list(csv.reader(f, delimiter=sep))}
    else:
        wb = load_workbook(caminho, read_only=True, data_only=True)
        folhas = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
        wb.close()
    for nome, linhas in folhas.items():
        idx, mapa = detectar_cabecalho(linhas, base.sin)
        if idx < 0:
            avisos.append(f"Aba '{nome}': cabeçalho não reconhecido (ignorada).")
            continue
        for k, row in enumerate(linhas[idx + 1:], start=idx + 2):
            g = lambda c: row[mapa[c]] if c in mapa and mapa[c] < len(row) else None
            if not (txt(g("CODIGO")) or txt(g("DESCRICAO")) or txt(g("NCM"))):
                continue
            itens.append({
                "codigo": txt(g("CODIGO")), "descricao": txt(g("DESCRICAO")), "ncm_orig": txt(g("NCM")),
                "cest": digitos(g("CEST")), "ean": digitos(g("EAN")), "unidade": txt(g("UNIDADE")),
                "cfop": digitos(g("CFOP")), "cst_cli": cod(g("CST_CLIENTE"), 3),
                "cct_cli": cod(g("CCLASSTRIB_CLIENTE"), 6), "origem": f"{nome}!{k}",
            })
        avisos.append(f"Aba '{nome}': cabeçalho na linha {idx + 1}; colunas reconhecidas: {', '.join(sorted(mapa))}.")
    return itens, avisos


def ler_xmls(caminho):
    arquivos = [caminho] if os.path.isfile(caminho) else [
        os.path.join(d, f) for d, _, fs in os.walk(caminho) for f in fs if f.lower().endswith(".xml")]
    itens, avisos = [], []
    for arq in sorted(arquivos):
        try:
            raiz = ET.parse(arq).getroot()
        except ET.ParseError:
            avisos.append(f"XML inválido: {os.path.basename(arq)}"); continue
        for det in raiz.iter(f"{NFE_NS}det"):
            p = det.find(f"{NFE_NS}prod")
            if p is None:
                continue
            g = lambda t: txt(p.findtext(f"{NFE_NS}{t}"))
            ibs = det.find(f".//{NFE_NS}IBSCBS")
            itens.append({
                "codigo": g("cProd"), "descricao": g("xProd"), "ncm_orig": g("NCM"), "cest": digitos(g("CEST")),
                "ean": digitos(g("cEAN")) if g("cEAN") != "SEM GTIN" else "", "unidade": g("uCom"),
                "cfop": g("CFOP"),
                "cst_cli": cod(ibs.findtext(f"{NFE_NS}CST"), 3) if ibs is not None else "",
                "cct_cli": cod(ibs.findtext(f"{NFE_NS}cClassTrib"), 6) if ibs is not None else "",
                "origem": f"{os.path.basename(arq)} item {det.get('nItem')}",
            })
    avisos.append(f"{len(arquivos)} arquivo(s) XML lido(s).")
    return itens, avisos


# =============================================================================== classificação
def casa(ncm, prefixos, excecoes):
    if any(ncm.startswith(e) for e in excecoes):
        return 0
    return max((len(p) for p in prefixos if ncm.startswith(p)), default=0)


def classificar(item, base):
    ncm, desc = item["ncm"], norm(item["descricao"])
    padrao_cst = cod(base.cfg.get("REGRA_PADRAO_CST", "000"), 3) or "000"
    padrao_cct = cod(base.cfg.get("REGRA_PADRAO_CCLASSTRIB", "000001"), 6) or "000001"
    res = {"cst": padrao_cst, "cct": padrao_cct, "trat": "Tributação integral", "lei": "LC 214/2025, arts. 4º e 11 (regra geral)",
           "cond": "", "anexo": "", "regra": "PADRÃO", "status": ST_OK, "obs": []}

    if not (len(ncm) == 8 and ncm.isdigit()):
        res["status"] = ST_NCM
        res["obs"].append("NCM ausente ou inválido (deve ter 8 dígitos). Classificação padrão aplicada até correção do NCM.")
    else:
        cands = []
        for r in base.regras:
            n = casa(ncm, r["_ncm"], r["_exc"])
            if n:
                cands.append((-n, r["_prio"], r))
        cands.sort(key=lambda x: (x[0], x[1]))
        aplicada = None
        for _, _, r in cands:
            if r["_excl"] and any(p in desc for p in r["_excl"]):
                continue
            if r["_chave"] and not any(p in desc for p in r["_chave"]):
                continue
            aplicada = r
            break
        if aplicada:
            r = aplicada
            res.update(cst=r["CST"], cct=r["CCLASSTRIB"], trat=txt(r["TRATAMENTO"]), lei=txt(r["BASE_LEGAL"]),
                       cond=txt(r["CONDICAO"]), anexo=txt(r["ANEXO"]), regra=txt(r["ID"]))
            if txt(r["OBSERVACAO"]):
                res["obs"].append(txt(r["OBSERVACAO"]))
            if not r["_conf"]:
                res["status"] = ST_CONF
                res["obs"].append(f"Regra {r['ID']} ainda não conferida na base de regras.")
            if not r["CCLASSTRIB"]:
                res["status"] = ST_CONF
                res["obs"].append("cClassTrib a definir pela tabela oficial vigente.")
            elif r["CCLASSTRIB"] not in base.cct:
                res["status"] = ST_CONF
                res["obs"].append(f"cClassTrib {r['CCLASSTRIB']} não consta da Tabela_cClassTrib da base.")
            elif not base.cct[r["CCLASSTRIB"]]["conf"]:
                res["status"] = ST_CONF
                res["obs"].append(f"cClassTrib {r['CCLASSTRIB']} pendente de conferência na tabela oficial.")
        elif cands:
            r = cands[0][2]
            res["status"] = ST_REV
            anexo = f"Anexo {txt(r['ANEXO'])}, " if txt(r["ANEXO"]) not in ("", "-") else ""
            res["obs"].append(f"Possível benefício: {txt(r['DESCRICAO_REGRA'])} ({anexo}regra {r['ID']}) — "
                              f"a descrição não confirma o enquadramento. Aplicada tributação integral (conservador).")
            res["regra"] = f"{r['ID']} (não aplicada)"
    # Imposto Seletivo
    res["is"], res["cat_is"] = "N", ""
    if ncm:
        for r in base.regras_is:
            if casa(ncm, r["_ncm"], r["_exc"]):
                res["is"], res["cat_is"] = "S", txt(r["CATEGORIA_IS"])
                res["obs"].append(f"Sujeito ao IS ({res['cat_is']}). CST-IS/cClassTribIS conforme tabela própria."
                                  + ("" if txt(r.get("CONFERIDO")).upper() == "S" else " Regra IS não conferida."))
                if res["status"] == ST_OK and txt(r.get("CONFERIDO")).upper() != "S":
                    res["status"] = ST_CONF
                break
    # Comparação com a classificação já usada pelo cliente
    if item["cst_cli"] or item["cct_cli"]:
        if (item["cst_cli"], item["cct_cli"]) != (res["cst"], res["cct"]):
            res["obs"].append(f"Cliente usa CST {item['cst_cli'] or '-'} / cClassTrib {item['cct_cli'] or '-'}: diverge da classificação proposta.")
    if res["cct"] and res["cct"][:3] != res["cst"]:
        res["status"] = ST_REV
        res["obs"].append("cClassTrib incompatível com o CST (verificar regra).")
    return res


# =============================================================================== saída Excel
FN = "Arial"
fH = Font(name=FN, bold=True, color="FFFFFF", size=10)
fB = Font(name=FN, size=10)
fBB = Font(name=FN, size=10, bold=True)
fT = Font(name=FN, size=14, bold=True, color="FFFFFF")
fS = Font(name=FN, size=11, bold=True, color="1F3864")
NAVY = PatternFill("solid", fgColor="1F3864")
GREY = PatternFill("solid", fgColor="595959")
LIGHT = PatternFill("solid", fgColor="D9E1F2")
thin = Side(style="thin", color="BFBFBF")
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
WR = Alignment(wrap_text=True, vertical="top")
CT = Alignment(horizontal="center", vertical="center", wrap_text=True)
STATUS_COR = {ST_OK: ("C6EFCE", "006100"), ST_CONF: ("FFEB9C", "9C5700"), ST_REV: ("FFC7CE", "9C0006"), ST_NCM: ("FFC7CE", "9C0006")}


def titulo(ws, texto, ultima_col):
    ws.merge_cells(f"A1:{ultima_col}1")
    c = ws["A1"]; c.value = texto; c.font = fT; c.fill = NAVY
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 28


def cabecalho(ws, linha, cols, fill=NAVY):
    for i, (nome, larg) in enumerate(cols, 1):
        c = ws.cell(row=linha, column=i, value=nome)
        c.font = fH; c.fill = fill if not isinstance(fill, list) else fill[i - 1]; c.alignment = CT; c.border = BD
        ws.column_dimensions[get_column_letter(i)].width = larg
    ws.row_dimensions[linha].height = 32


def cor_status(ws, rng):
    for st, (bg, fg) in STATUS_COR.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{st}"'],
                                      fill=PatternFill("solid", fgColor=bg), font=Font(name=FN, bold=True, color=fg)))


def gerar_excel(itens, base, args, avisos, divergencias, destino):
    wb = Workbook()
    hoje = dt.date.today().strftime("%d/%m/%Y")
    versao = base.cfg.get("VERSAO_TABELA_CCLASSTRIB", "")
    n = len(itens)
    last = n + 1  # última linha de dados no Cadastro_Fiscal

    # ------------------------------------------------------------------ Tabela de Referência
    wr = wb.active; wr.title = "Tabela_Referencia"
    usados = sorted({i["res"]["cct"] for i in itens if i["res"]["cct"]} | {"000001"})
    cabecalho(wr, 1, [("CCLASSTRIB", 12), ("CST", 7), ("DESCRICAO", 70), ("PRED_IBS", 10), ("PRED_CBS", 10), ("CONFERIDO", 11)])
    linhas_ref = sorted(set(usados) | set(base.cct))
    for k, c in enumerate(linhas_ref, 2):
        d = base.cct.get(c, {"desc": "NÃO CONSTA DA TABELA DA BASE", "ri": 0, "rc": 0, "conf": False})
        for j, v in enumerate([c, c[:3], d["desc"], d["ri"], d["rc"], "S" if d["conf"] else "N"], 1):
            x = wr.cell(row=k, column=j, value=v); x.font = fB; x.border = BD
            if j in (1, 2): x.number_format = "@"
            if j in (4, 5): x.number_format = "0%"
    ref_last = len(linhas_ref) + 1
    r0 = ref_last + 3
    wr.cell(row=r0 - 1, column=1, value="TABELA CST IBS/CBS").font = fS
    cabecalho(wr, r0, [("CST", 12), ("CST", 7), ("DESCRICAO", 70)])
    wr.cell(row=r0, column=1, value="CST"); wr.cell(row=r0, column=2, value="")
    csts = sorted(base.cst.items())
    for k, (c, d) in enumerate(csts, r0 + 1):
        wr.cell(row=k, column=1, value=c).number_format = "@"; wr.cell(row=k, column=3, value=d)
        for j in (1, 3): wr.cell(row=k, column=j).font = fB; wr.cell(row=k, column=j).border = BD
    cst_ini, cst_fim = r0 + 1, r0 + len(csts)
    wr.cell(row=cst_fim + 2, column=1, value=f"Versão da tabela: {versao}").font = Font(name=FN, size=9, italic=True)
    wr.freeze_panes = "A2"

    # ------------------------------------------------------------------ Cadastro Fiscal
    wc = wb.create_sheet("Cadastro_Fiscal", 0)
    cols = [("CODIGO_PRODUTO", 14), ("DESCRICAO", 42), ("EAN_GTIN", 15), ("NCM", 10), ("CEST", 9), ("UNIDADE", 8),
            ("CFOP_REF", 8), ("CST_IBS_CBS", 9), ("DESCRICAO_CST", 26), ("CCLASSTRIB", 11), ("DESCRICAO_CCLASSTRIB", 45),
            ("PRED_IBS", 9), ("PRED_CBS", 9), ("TRATAMENTO", 16), ("BASE_LEGAL", 30), ("CONDICAO_BENEFICIO", 34),
            ("SUJEITO_IS", 8), ("CATEGORIA_IS", 20), ("STATUS", 13), ("OBSERVACAO", 60), ("CHECAGEM_CST_X_CCLASSTRIB", 13),
            ("VERSAO_TABELA", 20), ("REGRA_APLICADA", 14), ("CLASSIF_ANTERIOR_CLIENTE", 16), ("ORIGEM_NO_ARQUIVO", 18)]
    calc = {9, 11, 12, 13, 21}
    cabecalho(wc, 1, cols, [GREY if i in calc else NAVY for i in range(1, len(cols) + 1)])
    RT = "Tabela_Referencia"
    for k, it in enumerate(itens, 2):
        r = it["res"]
        ant = "/".join(x for x in (it["cst_cli"], it["cct_cli"]) if x)
        m = f"MATCH($J{k},{RT}!$A$2:$A${ref_last},0)"
        vals = [it["codigo"], it["descricao"], it["ean"], it["ncm"] or it["ncm_orig"], it["cest"], it["unidade"], it["cfop"],
                r["cst"], f'=IFERROR(INDEX({RT}!$C${cst_ini}:$C${cst_fim},MATCH($H{k},{RT}!$A${cst_ini}:$A${cst_fim},0)),"CST não localizado")',
                r["cct"], f'=IF($J{k}="","",IFERROR(INDEX({RT}!$C$2:$C${ref_last},{m}),"cClassTrib não localizado"))',
                f'=IF($J{k}="","",IFERROR(INDEX({RT}!$D$2:$D${ref_last},{m}),""))',
                f'=IF($J{k}="","",IFERROR(INDEX({RT}!$E$2:$E${ref_last},{m}),""))',
                r["trat"], r["lei"], r["cond"], r["is"], r["cat_is"], r["status"], " | ".join(r["obs"]),
                f'=IF($J{k}="","PENDENTE",IF(LEFT($J{k},3)=$H{k},"OK","ERRO"))',
                versao, r["regra"], ant, it["origem"]]
        for j, v in enumerate(vals, 1):
            x = wc.cell(row=k, column=j, value=v); x.font = fB; x.border = BD
            if j in (1, 3, 4, 5, 7, 8, 10): x.number_format = "@"
            if j in (12, 13): x.number_format = "0%"
            if j == 20: x.alignment = WR
    wc.freeze_panes = "C2"
    wc.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(last, 2)}"
    cor_status(wc, f"S2:S{max(last, 2)}")
    wc.conditional_formatting.add(f"U2:U{max(last, 2)}", CellIsRule(operator="equal", formula=['"ERRO"'],
                                  fill=PatternFill("solid", fgColor="FFC7CE")))
    dv1 = DataValidation(type="list", formula1=f"={RT}!$A${cst_ini}:$A${cst_fim}", allow_blank=True)
    dv2 = DataValidation(type="list", formula1=f"={RT}!$A$2:$A${ref_last}", allow_blank=True)
    dv3 = DataValidation(type="list", formula1=f'"{ST_OK},{ST_CONF},{ST_REV},{ST_NCM}"', allow_blank=True)
    for dv, col in ((dv1, "H"), (dv2, "J"), (dv3, "S")):
        wc.add_data_validation(dv); dv.add(f"{col}2:{col}{max(last, 2)}")

    # ------------------------------------------------------------------ Importação ERP
    we = wb.create_sheet("Importacao_ERP", 1)
    ecols = [("COD_ITEM", 14, "A"), ("DESCRICAO", 42, "B"), ("NCM", 10, "D"), ("CEST", 9, "E"), ("GTIN", 15, "C"),
             ("UNIDADE", 8, "F"), ("CST_IBS_CBS", 10, "H"), ("CCLASSTRIB", 11, "J"), ("PRED_IBS_PERC", 10, None),
             ("PRED_CBS_PERC", 10, None), ("SUJEITO_IS", 9, "Q"), ("COD_CATEGORIA_IS", 11, None), ("STATUS", 13, "S")]
    cabecalho(we, 1, [(a, b) for a, b, _ in ecols])
    CF = "Cadastro_Fiscal"
    for k in range(2, last + 1):
        for j, (nome, _, src) in enumerate(ecols, 1):
            if src:
                f = f'=IF({CF}!${src}{k}="","",{CF}!${src}{k}&"")'
            elif nome == "PRED_IBS_PERC":
                f = f'=IF({CF}!$L{k}="","",ROUND({CF}!$L{k}*100,2))'
            elif nome == "PRED_CBS_PERC":
                f = f'=IF({CF}!$M{k}="","",ROUND({CF}!$M{k}*100,2))'
            else:
                f = f'=IF({CF}!$R{k}="","",TEXT(MATCH({CF}!$R{k},Listas!$A$2:$A$8,0),"00"))'
            x = we.cell(row=k, column=j, value=f); x.font = fB; x.border = BD
            if nome in ("PRED_IBS_PERC", "PRED_CBS_PERC"): x.number_format = "0.00"
            else: x.number_format = "@"
    we.freeze_panes = "B2"
    we.auto_filter.ref = f"A1:M{max(last, 2)}"
    cor_status(we, f"M2:M{max(last, 2)}")

    wl = wb.create_sheet("Listas")
    wl["A1"] = "CATEGORIA_IS"; wl["B1"] = "CODIGO"
    for k, c in enumerate(CAT_IS, 2):
        wl.cell(row=k, column=1, value=c); wl.cell(row=k, column=2, value=f"{k - 1:02d}").number_format = "@"
    wl.sheet_state = "hidden"

    # ------------------------------------------------------------------ Resumo
    ws = wb.create_sheet("Resumo", 2)
    titulo(ws, "RESUMO DA CLASSIFICAÇÃO", "E")
    for col, w in zip("ABCDE", [44, 12, 12, 30, 40]):
        ws.column_dimensions[col].width = w
    rng = lambda col: f"{CF}!${col}$2:${col}${max(last, 2)}"
    info = [("Cliente", args.cliente or "-"), ("Regime tributário", args.regime or "não informado"),
            ("Data da classificação", hoje), ("Versão da tabela cClassTrib", versao),
            ("Itens classificados", f"=COUNTA({rng('A')})"), ("Itens sujeitos ao Imposto Seletivo", f'=COUNTIF({rng("Q")},"S")')]
    for k, (a, b) in enumerate(info, 3):
        ws.cell(row=k, column=1, value=a).font = fBB
        x = ws.cell(row=k, column=2, value=b); x.font = fB; x.border = BD
    r = 3 + len(info) + 1
    ws.cell(row=r, column=1, value="Por status").font = fS
    cabecalho(ws, r + 1, [("Status", 44), ("Qtde", 12), ("%", 12), ("Significado", 30)])
    sig = {ST_OK: "Pronto para importação", ST_CONF: "Regra/código pendente de conferência",
           ST_REV: "Possível benefício ou inconsistência: revisar", ST_NCM: "NCM ausente/inválido: corrigir cadastro"}
    for k, st in enumerate(sig, r + 2):
        ws.cell(row=k, column=1, value=st)
        ws.cell(row=k, column=2, value=f'=COUNTIF({rng("S")},A{k})')
        ws.cell(row=k, column=3, value=f"=IF($B$7=0,0,B{k}/$B$7)").number_format = "0.0%"
        ws.cell(row=k, column=4, value=sig[st])
        for j in range(1, 5): ws.cell(row=k, column=j).font = fB; ws.cell(row=k, column=j).border = BD
    cor_status(ws, f"A{r + 2}:A{r + 5}")
    r = r + 7
    ws.cell(row=r, column=1, value="Por tratamento tributário").font = fS
    cabecalho(ws, r + 1, [("Tratamento", 44), ("Qtde", 12), ("%", 12)])
    trats = sorted({i["res"]["trat"] for i in itens})
    for k, t in enumerate(trats, r + 2):
        ws.cell(row=k, column=1, value=t)
        ws.cell(row=k, column=2, value=f'=COUNTIF({rng("N")},A{k})')
        ws.cell(row=k, column=3, value=f"=IF($B$7=0,0,B{k}/$B$7)").number_format = "0.0%"
        for j in range(1, 4): ws.cell(row=k, column=j).font = fB; ws.cell(row=k, column=j).border = BD
    r = r + 2 + len(trats) + 1
    ws.cell(row=r, column=1, value="Por NCM").font = fS
    cabecalho(ws, r + 1, [("NCM", 44), ("Qtde itens", 12), ("Classificações distintas", 12),
                          ("Classificação predominante (CST/cClassTrib)", 30), ("Situação", 40)])
    por_ncm = defaultdict(list)
    for it in itens:
        por_ncm[it["ncm"] or it["ncm_orig"] or "(sem NCM)"].append(f"{it['res']['cst']}/{it['res']['cct']}")
    for k, (ncm, cl) in enumerate(sorted(por_ncm.items()), r + 2):
        cnt = Counter(cl)
        vals = [ncm, f'=COUNTIF({rng("D")},A{k})', len(cnt), cnt.most_common(1)[0][0],
                "Classificações divergentes no mesmo NCM" if len(cnt) > 1 else "Uniforme"]
        for j, v in enumerate(vals, 1):
            x = ws.cell(row=k, column=j, value=v); x.font = fB; x.border = BD
        ws.cell(row=k, column=1).number_format = "@"

    # ------------------------------------------------------------------ Divergências e Riscos
    wd = wb.create_sheet("Divergencias_Riscos", 3)
    cabecalho(wd, 1, [("CODIGO_PRODUTO", 14), ("DESCRICAO", 40), ("NCM", 10), ("STATUS", 13), ("PROBLEMA", 60),
                      ("RISCO", 45), ("ACAO_RECOMENDADA", 50)])
    for k, d in enumerate(divergencias, 2):
        for j, v in enumerate(d, 1):
            x = wd.cell(row=k, column=j, value=v); x.font = fB; x.border = BD; x.alignment = WR
        wd.cell(row=k, column=3).number_format = "@"
    cor_status(wd, f"D2:D{max(len(divergencias) + 1, 2)}")
    wd.freeze_panes = "B2"
    if divergencias:
        wd.auto_filter.ref = f"A1:G{len(divergencias) + 1}"

    # ------------------------------------------------------------------ Leiaute p/ programadores
    wy = wb.create_sheet("Leiaute_Importacao", 4)
    titulo(wy, "LEIAUTE PARA IMPORTAÇÃO NO ERP — ABA Importacao_ERP", "F")
    cabecalho(wy, 3, [("Campo", 20), ("Descrição", 46), ("Tipo", 12), ("Tamanho", 10), ("Obrigatório", 12),
                      ("Destino no documento fiscal / observação", 70)])
    lay = [
        ("COD_ITEM", "Código do produto no ERP do cliente", "Texto", "até 60", "Sim", "Chave para localizar o item no cadastro (cProd)."),
        ("DESCRICAO", "Descrição do produto", "Texto", "até 120", "Sim", "Apenas conferência; não sobrescrever a descrição do ERP."),
        ("NCM", "NCM/SH", "Texto", "8", "Sim", "Somente dígitos. Atualizar no cadastro se divergir."),
        ("CEST", "CEST", "Texto", "7", "Não", "ICMS-ST (mantido na transição)."),
        ("GTIN", "Código de barras", "Texto", "8–14", "Não", "cEAN."),
        ("UNIDADE", "Unidade comercial", "Texto", "até 6", "Não", "uCom."),
        ("CST_IBS_CBS", "CST do IBS e da CBS", "Texto", "3", "Sim", "NF-e/NFC-e: grupo IBSCBS, tag CST (NT 2025.002-RTC). Manter zeros à esquerda."),
        ("CCLASSTRIB", "Código de classificação tributária", "Texto", "6", "Sim", "NF-e/NFC-e: grupo IBSCBS, tag cClassTrib. Os 3 primeiros dígitos = CST."),
        ("PRED_IBS_PERC", "% de redução da alíquota do IBS", "Decimal", "3,2", "Sim", "Ex.: 60.00 = redução de 60%; 100.00 = alíquota zero. Usar no cálculo da alíquota efetiva (grupo de redução da NT)."),
        ("PRED_CBS_PERC", "% de redução da alíquota da CBS", "Decimal", "3,2", "Sim", "Idem para CBS."),
        ("SUJEITO_IS", "Incidência do Imposto Seletivo", "Texto", "1", "Sim", "S/N. Se S, parametrizar grupo IS (CST-IS/cClassTribIS) conforme tabela própria."),
        ("COD_CATEGORIA_IS", "Categoria do IS", "Texto", "2", "Se IS=S", "01 Veículos; 02 Embarcações e aeronaves; 03 Fumígenos; 04 Bebidas alcoólicas; 05 Bebidas açucaradas; 06 Bens minerais; 07 Prognósticos/fantasy sport."),
        ("STATUS", "Situação da classificação", "Texto", "até 12", "-", "Importar somente itens com STATUS = OK. Demais: aguardar validação do escritório contábil."),
    ]
    for k, ln in enumerate(lay, 4):
        for j, v in enumerate(ln, 1):
            x = wy.cell(row=k, column=j, value=v); x.font = fB; x.border = BD; x.alignment = WR
    k = 4 + len(lay) + 1
    for t in ["Orientações aos programadores:",
              "• Ler a aba Importacao_ERP a partir da linha 2 (linha 1 = cabeçalho). Todos os códigos são texto: preservar zeros à esquerda (ex.: 000001).",
              "• A chave de atualização é COD_ITEM. Itens com STATUS diferente de OK não devem ser importados.",
              "• As alíquotas de IBS/CBS não constam do arquivo: são definidas por ente (União, UF e Município) e pelo ano de transição.",
              "• Para exportar em CSV: Arquivo > Salvar como > CSV UTF-8 (separador ';' em Excel pt-BR) com a aba Importacao_ERP ativa."]:
        wy.cell(row=k, column=1, value=t).font = fBB if t.endswith(":") else fB; k += 1

    # ------------------------------------------------------------------ Instruções
    wi = wb.create_sheet("Instrucoes", 0)
    titulo(wi, "CLASSIFICAÇÃO TRIBUTÁRIA DE PRODUTOS — IBS / CBS / IS", "C")
    wi.column_dimensions["A"].width = 3; wi.column_dimensions["B"].width = 32; wi.column_dimensions["C"].width = 100
    escr = base.cfg.get("ESCRITORIO", ""); resp = base.cfg.get("RESPONSAVEL", "")
    blocos = [
        ("IDENTIFICAÇÃO", [("Cliente", args.cliente or "-"), ("Regime tributário", args.regime or "não informado"),
                           ("Escritório / responsável", " – ".join(x for x in (escr, resp) if x) or "-"),
                           ("Data", hoje), ("Arquivo de origem", os.path.basename(args.arquivo)),
                           ("Versão da tabela cClassTrib", versao)]),
        ("BASE NORMATIVA", [("Constituição", "EC 132/2023 (arts. 156-A, 195, V, e 153, VIII, da CF; ADCT)."),
                            ("Lei Complementar", "LC 214/2025, com as alterações da LC 227/2026 e da LC 235/2026 (redação consolidada)."),
                            ("Tabelas", "Informe Técnico 2025.002 do Portal Nacional da NF-e (CST e cClassTrib) e NT 2025.002-RTC (leiaute do grupo IBSCBS).")]),
        ("CONTEÚDO DO ARQUIVO", [("Cadastro_Fiscal", "Classificação item a item: CST, cClassTrib, reduções, tratamento, base legal, condição do benefício, IS, status e observações."),
                                 ("Importacao_ERP", "Aba para os programadores: somente os campos de importação, em formato técnico."),
                                 ("Resumo", "Totais por status, por tratamento e por NCM."),
                                 ("Divergencias_Riscos", "Itens que exigem atenção, risco envolvido e ação recomendada."),
                                 ("Leiaute_Importacao", "Especificação dos campos para o ERP."),
                                 ("Tabela_Referencia", "Tabelas de CST e cClassTrib utilizadas.")]),
        ("LEGENDA DE STATUS", [(ST_OK, sig[ST_OK]), (ST_CONF, sig[ST_CONF]), (ST_REV, sig[ST_REV]), (ST_NCM, sig[ST_NCM])]),
        ("PREMISSAS E LIMITAÇÕES", [
            ("Critério", "Enquadramento pelo NCM informado pelo cliente combinado com a descrição do produto. A correção do NCM é de responsabilidade do cliente."),
            ("Abordagem conservadora", "Quando o NCM consta de anexo com benefício, mas a descrição não permite confirmar o enquadramento, aplica-se tributação integral (CST 000) e o item é marcado para revisão."),
            ("Benefícios condicionados", "Reduções que dependem do adquirente, da finalidade ou de registro (Anvisa, MAPA) estão indicadas na coluna CONDICAO_BENEFICIO e devem ser verificadas em cada operação."),
            ("Atualização", "A classificação deve ser revalidada a cada nova versão da tabela cClassTrib e a cada alteração da LC 214/2025."),
        ]),
    ]
    k = 3
    for sec, linhas in blocos:
        wi.merge_cells(start_row=k, start_column=2, end_row=k, end_column=3)
        c = wi.cell(row=k, column=2, value=sec); c.font = fS; c.fill = LIGHT; k += 1
        for a, b in linhas:
            x = wi.cell(row=k, column=2, value=a); x.font = fBB; x.border = BD; x.alignment = WR
            y = wi.cell(row=k, column=3, value=b); y.font = fB; y.border = BD; y.alignment = WR
            k += 1
        k += 1
    cor_status(wi, f"B3:B{k}")
    wi.sheet_view.showGridLines = False

    wb.calculation.fullCalcOnLoad = True
    wb.save(destino)


# =============================================================================== divergências
def montar_divergencias(itens):
    out = []
    risco = {ST_NCM: ("Rejeição do documento fiscal e enquadramento incorreto em toda a cadeia",
                      "Cliente deve informar/corrigir o NCM; reprocessar o item"),
             ST_REV: ("Recolhimento a maior (benefício não aproveitado) ou glosa/autuação se o benefício for aplicado indevidamente",
                      "Analisar a descrição completa/ficha técnica do produto e confirmar o enquadramento no Anexo"),
             ST_CONF: ("Uso de regra ou código ainda não validado com a tabela oficial vigente",
                       "Escritório deve conferir a regra/código na base de regras antes da importação")}
    for it in itens:
        st = it["res"]["status"]
        if st != ST_OK:
            out.append([it["codigo"], it["descricao"], it["ncm"] or it["ncm_orig"], st, " | ".join(it["res"]["obs"]), *risco[st]])
    # mesmo código com NCM diferente
    por_cod = defaultdict(set)
    for it in itens:
        if it["codigo"]:
            por_cod[it["codigo"]].add(it["ncm"] or it["ncm_orig"])
    for c, ncms in por_cod.items():
        if len(ncms) > 1:
            out.append([c, "", "; ".join(sorted(ncms)), ST_REV, "Mesmo código de produto com NCMs diferentes no arquivo do cliente",
                        "Cadastro inconsistente; risco de classificação divergente entre documentos", "Cliente deve unificar o NCM do produto"])
    # mesmo NCM com classificações diferentes
    por_ncm = defaultdict(set)
    for it in itens:
        if it["ncm"]:
            por_ncm[it["ncm"]].add((it["res"]["cst"], it["res"]["cct"]))
    for ncm, cl in por_ncm.items():
        if len(cl) > 1:
            out.append(["(vários)", "", ncm, ST_REV, "Mesmo NCM com classificações diferentes: " + "; ".join(f"{a}/{b}" for a, b in sorted(cl)),
                        "Pode ser correto (descrições distintas), mas exige justificativa item a item",
                        "Conferir a descrição de cada item frente ao texto do Anexo"])
    # classificação anterior do cliente divergente
    for it in itens:
        if (it["cst_cli"] or it["cct_cli"]) and (it["cst_cli"], it["cct_cli"]) != (it["res"]["cst"], it["res"]["cct"]):
            out.append([it["codigo"], it["descricao"], it["ncm"], it["res"]["status"],
                        f"Cliente já usa {it['cst_cli']}/{it['cct_cli']}; proposta {it['res']['cst']}/{it['res']['cct']}",
                        "Documentos já emitidos podem estar com classificação incorreta",
                        "Validar e, se for o caso, corrigir o cadastro e avaliar documentos emitidos"])
    return out


# =============================================================================== main
def main():
    ap = argparse.ArgumentParser(description="Classificação de produtos IBS/CBS/IS a partir do cadastro do cliente.")
    ap.add_argument("arquivo", help="Excel/CSV do cliente, XML de NF-e ou pasta com XMLs")
    ap.add_argument("--cliente", default="", help="nome do cliente")
    ap.add_argument("--regime", default="", help="regime tributário do cliente")
    ap.add_argument("--base", default=BASE_PADRAO, help="base de regras (padrão: base_regras.xlsx ao lado do script)")
    ap.add_argument("--saida", default=None, help="pasta de saída (padrão: pasta 'saida' ao lado do arquivo do cliente)")
    args = ap.parse_args()

    if not os.path.exists(args.arquivo):
        sys.exit(f"Arquivo não encontrado: {args.arquivo}")
    base = Base(args.base)
    if os.path.isdir(args.arquivo) or args.arquivo.lower().endswith(".xml"):
        itens, avisos = ler_xmls(args.arquivo)
    elif args.arquivo.lower().endswith((".xlsx", ".xlsm", ".csv")):
        itens, avisos = ler_planilha(args.arquivo, base)
    else:
        sys.exit("Formato não suportado. Use .xlsx, .xlsm, .csv ou XML de NF-e (arquivos .xls: salve antes como .xlsx).")
    if not itens:
        sys.exit("Nenhum produto encontrado. " + " ".join(avisos))

    # deduplicação: mesmo código + mesmo NCM = um item
    vistos, unicos = set(), []
    for it in itens:
        it["ncm"] = digitos(it["ncm_orig"])
        chave = (it["codigo"] or it["descricao"], it["ncm"])
        if chave in vistos:
            continue
        vistos.add(chave); unicos.append(it)
    for it in unicos:
        it["res"] = classificar(it, base)
    div = montar_divergencias(unicos)

    nome_cli = re.sub(r"[^A-Za-z0-9_-]+", "_", args.cliente or os.path.splitext(os.path.basename(args.arquivo.rstrip("/\\")))[0]).strip("_")
    pasta = args.saida or os.path.join(os.path.dirname(os.path.abspath(args.arquivo.rstrip("/\\"))), "saida")
    os.makedirs(pasta, exist_ok=True)
    destino = os.path.join(pasta, f"{nome_cli}_Classificacao_IBS_CBS_{dt.date.today():%Y%m%d}.xlsx")
    gerar_excel(unicos, base, args, avisos, div, destino)

    cont = Counter(i["res"]["status"] for i in unicos)
    print("\n".join(avisos))
    print(f"\nItens lidos: {len(itens)} | únicos classificados: {len(unicos)}")
    for st in (ST_OK, ST_CONF, ST_REV, ST_NCM):
        print(f"  {st:<12} {cont.get(st, 0)}")
    print(f"Divergências/riscos listados: {len(div)}")
    print(f"\nArquivo gerado: {destino}")


if __name__ == "__main__":
    main()
