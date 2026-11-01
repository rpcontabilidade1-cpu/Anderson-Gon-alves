#!/usr/bin/env python3
"""
SISTEMA DE CLASSIFICAÇÃO DE PRODUTOS E SERVIÇOS — IBS / CBS / IS
EC 132/2023 e LC 214/2025 (com alterações posteriores)

Recebe o cadastro do cliente (Excel .xlsx/.xlsm, CSV ou XMLs de NF-e), enquadra cada
item (CST IBS/CBS, cClassTrib, reduções, Imposto Seletivo) com base na base_regras.xlsx
e devolve um Excel padronizado para o cliente entregar aos programadores do ERP.

A versão para navegador (web/classificador.html) aplica exatamente as mesmas regras.

Uso:
    python classificar.py ARQUIVO [--cliente "Nome"] [--regime "Simples Nacional - Regime Único"]
                          [--base base_regras.xlsx] [--saida PASTA]
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
    from openpyxl.formatting.rule import CellIsRule
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

REGIMES = {
    "simples": "Simples Nacional – Regime Único e Regime Híbrido (duas visões)",
    "mei": "MEI",
    "presumido": "Lucro Presumido",
    "real": "Lucro Real",
    "": "não informado",
}
NOTAS_REGIME = {
    "simples": [
        ("Duas visões", "Este arquivo traz a classificação do cliente nas duas modalidades do Simples Nacional: Simples puro (regime único, IBS/CBS no DAS) na aba Importacao_ERP_Unico, e Simples híbrido (IBS/CBS pelo regime regular) na aba Importacao_ERP_Hibrido. O ERP deve importar a aba do regime efetivamente adotado no período. A aba Comparativo_Simples resume as diferenças."),
        ("Classificação", "CST, cClassTrib e percentuais de redução são os mesmos nas duas visões: identificam a natureza do item. O que muda é a forma de recolhimento, a apuração e o crédito transferido ao adquirente. Conferir na NT vigente o preenchimento do grupo IBSCBS para o CRT do Simples Nacional em cada modalidade."),
        ("Único – recolhimento", "IBS e CBS recolhidos dentro do DAS (LC 123/2006, com as alterações da LC 214/2025). A classificação desta planilha identifica a natureza de cada item e deve ser informada nos documentos fiscais conforme a NT vigente para o CRT do Simples Nacional."),
        ("Único – benefícios", "Itens com redução ou alíquota zero exigem a segregação da receita no PGDAS-D. Conferir a regulamentação do CGSN vigente sobre o reflexo das reduções no cálculo do DAS."),
        ("Único – crédito do adquirente", "O adquirente sujeito ao regime regular apropria crédito correspondente ao IBS/CBS efetivamente recolhido no DAS na operação, e não ao valor cheio. Avaliar com clientes B2B a opção pelo regime híbrido (LC 214/2025, art. 41, § 3º)."),
        ("Híbrido – recolhimento", "Opção pelo recolhimento do IBS e da CBS pelo regime regular (LC 214/2025, art. 41, § 3º): os demais tributos seguem no DAS; IBS/CBS são apurados fora, de forma não cumulativa."),
        ("Híbrido – aplicação", "A classificação (CST, cClassTrib e reduções) aplica-se integralmente, como nas empresas do regime regular, com crédito pleno ao adquirente e direito a crédito nas aquisições."),
    ],
    "mei": [
        ("Recolhimento", "MEI recolhe valores fixos mensais. Conferir na regulamentação vigente a obrigatoriedade de preenchimento dos grupos de IBS/CBS nos documentos fiscais emitidos."),
    ],
    "presumido": [("Aplicação", "Regime regular de IBS/CBS (não cumulativo), independentemente da forma de tributação do IRPJ/CSLL. A classificação aplica-se integralmente.")],
    "real": [("Aplicação", "Regime regular de IBS/CBS (não cumulativo). A classificação aplica-se integralmente.")],
}


OBS_SIMPLES = ("Simples puro: segregar a receita deste item no PGDAS-D para refletir o benefício. "
               "Simples híbrido: benefício aplicado diretamente na alíquota do item.")

# Colunas próprias de cada visão do Simples nas abas de importação: (cabeçalho, largura, tipo, valor)
ERP_SIMPLES = {
    "Unico": [("REGIME_IBS_CBS", 16, "const", "SIMPLES_UNICO"), ("RECOLHIMENTO_IBS_CBS", 16, "const", "DAS"),
              ("SEGREGAR_RECEITA_PGDAS", 14, "seg", ""), ("CREDITO_ADQUIRENTE", 18, "const", "LIMITADO_AO_DAS")],
    "Hibrido": [("REGIME_IBS_CBS", 16, "const", "SIMPLES_HIBRIDO"), ("RECOLHIMENTO_IBS_CBS", 16, "const", "REGIME_REGULAR"),
                ("APURACAO", 14, "const", "NAO_CUMULATIVA"), ("CREDITO_ADQUIRENTE", 18, "cred", "")],
}
LEIAUTE_SIMPLES = [
    ("REGIME_IBS_CBS", "Modalidade do Simples Nacional da aba", "Texto", "até 16", "Sim", "SIMPLES_UNICO (aba Importacao_ERP_Unico) ou SIMPLES_HIBRIDO (aba Importacao_ERP_Hibrido). Importar só a aba do regime adotado."),
    ("RECOLHIMENTO_IBS_CBS", "Onde o IBS/CBS é recolhido", "Texto", "até 16", "Sim", "DAS (único) ou REGIME_REGULAR (híbrido, fora do DAS)."),
    ("SEGREGAR_RECEITA_PGDAS", "Somente visão Único", "Texto", "3", "Sim", "SIM quando o item tem redução ou alíquota zero: a receita deve ser segregada no PGDAS-D."),
    ("APURACAO", "Somente visão Híbrido", "Texto", "até 16", "Sim", "NAO_CUMULATIVA: IBS/CBS apurados com crédito das aquisições."),
    ("CREDITO_ADQUIRENTE", "Crédito transferido ao comprador", "Texto", "até 16", "Sim", "LIMITADO_AO_DAS (único); INTEGRAL ou SEM_DESTAQUE quando alíquota zero (híbrido)."),
]
COMPARATIVO = [
    ("Recolhimento do IBS/CBS", "Dentro do DAS, conforme anexo e faixa da LC 123/2006", "Fora do DAS, pelas alíquotas do regime regular (LC 214/2025, art. 41, § 3º)"),
    ("Apuração", "Sem crédito de IBS/CBS sobre as aquisições", "Não cumulativa: crédito do IBS/CBS pago nas aquisições"),
    ("Crédito ao adquirente", "Limitado ao IBS/CBS efetivamente recolhido no DAS na operação", "Integral, sobre o IBS/CBS destacado no documento"),
    ("Benefícios por item (reduções, alíquota zero)", "Refletidos pela segregação de receitas no PGDAS-D (conferir regulamentação do CGSN)", "Aplicados diretamente na alíquota do item (cClassTrib)"),
    ("Demais tributos (IRPJ, CSLL, CPP, ICMS/ISS na transição)", "No DAS", "No DAS (sem alteração)"),
    ("Obrigações acessórias", "PGDAS-D", "PGDAS-D + apuração e escrituração do IBS/CBS do regime regular"),
    ("Tende a ser vantajoso quando", "Vendas predominantemente a consumidor final (B2C) e poucas compras com IBS/CBS a creditar", "Vendas a empresas do regime regular (B2B), que valorizam o crédito integral, ou muitas compras com IBS/CBS"),
    ("Opção", "Regra padrão do optante", "Opção expressa pelo regime regular de IBS/CBS; prazos e periodicidade conforme regulamentação vigente"),
]


def regime_chave(v):
    n = norm(v)
    if "mei" in n.split() or "microempreendedor" in n:
        return "mei"
    if "simples" in n or "unico" in n or "hibrid" in n:
        return "simples"
    if "presumido" in n:
        return "presumido"
    if "real" in n:
        return "real"
    return ""


# =============================================================================== utilidades
def norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return " ".join(re.sub(r"[^a-z0-9/ -]", " ", s).split())


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
            sys.exit(f"Base de regras não encontrada: {caminho}")
        wb = load_workbook(caminho, read_only=True, data_only=True)
        self.regras = [self.preparar(r) for r in self._linhas(wb["Regras_NCM"])]
        self.regras_is = []
        for r in self._linhas(wb["Regras_IS"]):
            r["_ncm"] = [digitos(x) for x in lista(r["NCM"]) if digitos(x)]
            r["_exc"] = [digitos(x) for x in lista(r["EXCETO_NCM"]) if digitos(x)]
            r["_conf"] = txt(r.get("CONFERIDO")).upper() == "S"
            self.regras_is.append(r)
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
    def preparar(r):
        r["_ncm"] = [digitos(x) for x in lista(r.get("NCM")) if digitos(x)]
        r["_exc"] = [digitos(x) for x in lista(r.get("EXCETO_NCM")) if digitos(x)]
        r["_nbs"] = [digitos(x) for x in lista(r.get("NBS")) if digitos(x)]
        r["_chave"] = [norm(x) for x in lista(r.get("PALAVRAS_CHAVE"))]
        r["_excl"] = [norm(x) for x in lista(r.get("PALAVRAS_EXCLUIR"))]
        r["CST"] = cod(r.get("CST"), 3)
        r["CCLASSTRIB"] = cod(r.get("CCLASSTRIB"), 6)
        r["_prio"] = int(r.get("PRIORIDADE") or 9)
        r["_conf"] = txt(r.get("CONFERIDO")).upper() == "S"
        r["_serv"] = norm(r.get("TIPO_ITEM")).startswith("serv")
        return r

    @staticmethod
    def _linhas(ws):
        it = ws.iter_rows(values_only=True)
        cab = [txt(c) for c in next(it)]
        return [dict(zip(cab, row)) for row in it if not all(v in (None, "") for v in row)]


# =============================================================================== leitura do cliente
def detectar_cabecalho(linhas, sin):
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


def novo_item(**k):
    base = {"codigo": "", "descricao": "", "ncm_orig": "", "nbs": "", "tipo_cli": "", "cest": "", "ean": "",
            "unidade": "", "cfop": "", "cst_cli": "", "cct_cli": "", "origem": ""}
    base.update(k)
    return base


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
            itens.append(novo_item(
                codigo=txt(g("CODIGO")), descricao=txt(g("DESCRICAO")), ncm_orig=txt(g("NCM")), nbs=digitos(g("NBS")),
                tipo_cli=txt(g("TIPO")), cest=digitos(g("CEST")), ean=digitos(g("EAN")), unidade=txt(g("UNIDADE")),
                cfop=digitos(g("CFOP")), cst_cli=cod(g("CST_CLIENTE"), 3), cct_cli=cod(g("CCLASSTRIB_CLIENTE"), 6),
                origem=f"{nome}!{k}"))
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
            itens.append(novo_item(
                codigo=g("cProd"), descricao=g("xProd"), ncm_orig=g("NCM"), cest=digitos(g("CEST")),
                ean=digitos(g("cEAN")), unidade=g("uCom"), cfop=g("CFOP"),
                cst_cli=cod(ibs.findtext(f"{NFE_NS}CST"), 3) if ibs is not None else "",
                cct_cli=cod(ibs.findtext(f"{NFE_NS}cClassTrib"), 6) if ibs is not None else "",
                origem=f"{os.path.basename(arq)} item {det.get('nItem')}"))
    avisos.append(f"{len(arquivos)} arquivo(s) XML lido(s).")
    return itens, avisos


# =============================================================================== classificação
def casa(cod_item, prefixos, excecoes):
    if not cod_item or any(cod_item.startswith(e) for e in excecoes):
        return 0
    return max((len(p) for p in prefixos if cod_item.startswith(p)), default=0)


def descr_ok(r, desc):
    if r["_excl"] and any(p in desc for p in r["_excl"]):
        return False
    return not r["_chave"] or any(p in desc for p in r["_chave"])


def classificar(item, base, regime):
    ncm, nbs, desc = item["ncm"], item["nbs"], norm(item["descricao"])
    servico_cli = norm(item["tipo_cli"]).startswith("serv") or len(nbs) == 9
    res = {"cst": cod(base.cfg.get("REGRA_PADRAO_CST"), 3) or "000",
           "cct": cod(base.cfg.get("REGRA_PADRAO_CCLASSTRIB"), 6) or "000001",
           "trat": "Tributação integral", "lei": "LC 214/2025, arts. 4º e 11 (regra geral)", "cond": "",
           "regra": "PADRÃO", "status": ST_OK, "obs": [], "tipo": "Bem", "is": "N", "cat_is": ""}
    ncm_ok = len(ncm) == 8
    cands, aplicada = [], None
    if item.get("ncm_ajustado"):
        res["obs"].append(f"NCM informado com 7 dígitos ({item['ncm_orig']}); completado com zero à esquerda: {ncm}.")

    if ncm_ok:
        for r in base.regras:
            if not r["_serv"]:
                n = casa(ncm, r["_ncm"], r["_exc"])
                if n:
                    cands.append((-n, r["_prio"], r))
    else:
        # sem NCM: tenta regras de serviço (por NBS ou palavras-chave)
        for r in base.regras:
            if r["_serv"]:
                n = casa(nbs, r["_nbs"], []) if r["_nbs"] else 0
                if n or (not r["_nbs"] and r["_chave"] and descr_ok(r, desc)):
                    cands.append((-n, r["_prio"], r))
    cands.sort(key=lambda x: (x[0], x[1]))
    for _, _, r in cands:
        if descr_ok(r, desc):
            aplicada = r
            break

    if not ncm_ok and (aplicada or servico_cli):
        res["tipo"] = "Serviço"
    elif not ncm_ok:
        res["status"] = ST_NCM
        res["obs"].append("NCM ausente ou inválido (8 dígitos). Se for serviço, informe a NBS ou o tipo 'Serviço'. Classificação padrão aplicada.")

    if aplicada:
        r = aplicada
        res.update(cst=r["CST"], cct=r["CCLASSTRIB"], trat=txt(r["TRATAMENTO"]), lei=txt(r["BASE_LEGAL"]),
                   cond=txt(r["CONDICAO"]), regra=txt(r["ID"]))
        if txt(r.get("OBSERVACAO")):
            res["obs"].append(txt(r["OBSERVACAO"]))
        if res["status"] != ST_NCM:
            if not r["_conf"]:
                res["status"] = ST_CONF
                res["obs"].append(f"Regra {r['ID']} ainda não conferida na base de regras.")
            if not r["CCLASSTRIB"]:
                res["status"] = ST_CONF
                res["obs"].append("cClassTrib a definir pela tabela oficial vigente.")
            elif r["CCLASSTRIB"] not in base.cct:
                res["status"] = ST_CONF
                res["obs"].append(f"cClassTrib {r['CCLASSTRIB']} não consta da tabela de referência.")
            elif not base.cct[r["CCLASSTRIB"]]["conf"]:
                res["status"] = ST_CONF
                res["obs"].append(f"cClassTrib {r['CCLASSTRIB']} pendente de conferência na tabela oficial.")
    elif cands:
        r = cands[0][2]
        anexo = f"Anexo {txt(r['ANEXO'])}, " if txt(r["ANEXO"]) not in ("", "-") else ""
        res["status"] = ST_REV
        res["regra"] = f"{r['ID']} (não aplicada)"
        res["obs"].append(f"Possível benefício: {txt(r['DESCRICAO_REGRA'])} ({anexo}regra {r['ID']}). A descrição não confirma "
                          f"o enquadramento; aplicada tributação integral (critério conservador).")

    if ncm_ok:
        for r in base.regras_is:
            if casa(ncm, r["_ncm"], r["_exc"]):
                res["is"], res["cat_is"] = "S", txt(r["CATEGORIA_IS"])
                res["obs"].append(f"Sujeito ao IS ({res['cat_is']}); CST-IS/cClassTribIS conforme tabela própria."
                                  + ("" if r["_conf"] else " Regra IS não conferida."))
                if res["status"] == ST_OK and not r["_conf"]:
                    res["status"] = ST_CONF
                break

    if (item["cst_cli"] or item["cct_cli"]) and (item["cst_cli"], item["cct_cli"]) != (res["cst"], res["cct"]):
        res["obs"].append(f"Cliente usa CST {item['cst_cli'] or '-'} / cClassTrib {item['cct_cli'] or '-'}: diverge da proposta.")
    if res["cct"] and res["cct"][:3] != res["cst"]:
        res["status"] = ST_REV
        res["obs"].append("cClassTrib incompatível com o CST (verificar regra).")

    res["obs_regime"] = ""
    if regime == "simples" and res["trat"] != "Tributação integral":
        res["obs_regime"] = OBS_SIMPLES
    elif regime == "mei":
        res["obs_regime"] = "MEI: valores fixos; conferir obrigatoriedade de destaque de IBS/CBS."
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
SIGNIF = {ST_OK: "Pronto para importação", ST_CONF: "Regra/código pendente de conferência pelo escritório",
          ST_REV: "Possível benefício ou inconsistência: revisar", ST_NCM: "NCM ausente/inválido: corrigir cadastro"}

# colunas do Cadastro_Fiscal (chave, cabeçalho, largura, calculada)
CAD = [("cod", "CODIGO_PRODUTO", 14, 0), ("desc", "DESCRICAO", 42, 0), ("tipo", "TIPO_ITEM", 9, 0), ("ncm", "NCM", 10, 0),
       ("nbs", "NBS", 11, 0), ("cest", "CEST", 9, 0), ("ean", "EAN_GTIN", 15, 0), ("un", "UNIDADE", 8, 0),
       ("cfop", "CFOP_REF", 8, 0), ("cst", "CST_IBS_CBS", 9, 0), ("dcst", "DESCRICAO_CST", 26, 1),
       ("cct", "CCLASSTRIB", 11, 0), ("dcct", "DESCRICAO_CCLASSTRIB", 45, 1), ("ri", "PRED_IBS", 9, 1),
       ("rc", "PRED_CBS", 9, 1), ("trat", "TRATAMENTO", 16, 0), ("lei", "BASE_LEGAL", 30, 0),
       ("cond", "CONDICAO_BENEFICIO", 34, 0), ("is", "SUJEITO_IS", 8, 0), ("cat", "CATEGORIA_IS", 20, 0),
       ("st", "STATUS", 13, 0), ("obs", "OBSERVACAO", 60, 0), ("chk", "CHECAGEM_CST_X_CCLASSTRIB", 13, 1),
       ("oreg", "OBS_REGIME", 40, 0), ("regra", "REGRA_APLICADA", 14, 0), ("ant", "CLASSIF_ANTERIOR_CLIENTE", 16, 0),
       ("orig", "ORIGEM_NO_ARQUIVO", 18, 0), ("ver", "VERSAO_TABELA", 22, 0)]
L = {k: get_column_letter(i) for i, (k, *_) in enumerate(CAD, 1)}
TEXTO = {"cod", "ncm", "nbs", "cest", "ean", "cfop", "cst", "cct"}
# Importacao_ERP: (cabeçalho, largura, origem)
ERP = [("COD_ITEM", 14, "cod"), ("DESCRICAO", 42, "desc"), ("TIPO_ITEM", 9, "tipo"), ("NCM", 10, "ncm"), ("NBS", 11, "nbs"),
       ("CEST", 9, "cest"), ("GTIN", 15, "ean"), ("UNIDADE", 8, "un"), ("CST_IBS_CBS", 10, "cst"), ("CCLASSTRIB", 11, "cct"),
       ("PRED_IBS_PERC", 10, "ri%"), ("PRED_CBS_PERC", 10, "rc%"), ("SUJEITO_IS", 9, "is"), ("COD_CATEGORIA_IS", 11, "cat#"),
       ("STATUS", 13, "st")]


def titulo(ws, texto, ultima_col):
    ws.merge_cells(f"A1:{ultima_col}1")
    c = ws["A1"]; c.value = texto; c.font = fT; c.fill = NAVY
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 28


def cabecalho(ws, linha, cols, fills=None):
    for i, (nome, larg) in enumerate(cols, 1):
        c = ws.cell(row=linha, column=i, value=nome)
        c.font = fH; c.fill = fills[i - 1] if fills else NAVY; c.alignment = CT; c.border = BD
        ws.column_dimensions[get_column_letter(i)].width = larg
    ws.row_dimensions[linha].height = 32


def cor_status(ws, rng):
    for st, (bg, fg) in STATUS_COR.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{st}"'],
                                      fill=PatternFill("solid", fgColor=bg), font=Font(name=FN, bold=True, color=fg)))


def gerar_excel(itens, base, args, regime, divergencias, destino):
    wb = Workbook()
    hoje = dt.date.today().strftime("%d/%m/%Y")
    versao = base.cfg.get("VERSAO_TABELA_CCLASSTRIB", "")
    last = max(len(itens) + 1, 2)

    # ---------------- Tabela_Referencia
    wr = wb.active; wr.title = "Tabela_Referencia"
    cabecalho(wr, 1, [("CCLASSTRIB", 12), ("CST", 7), ("DESCRICAO", 70), ("PRED_IBS", 10), ("PRED_CBS", 10), ("CONFERIDO", 11)])
    codigos = sorted({i["res"]["cct"] for i in itens if i["res"]["cct"]} | set(base.cct))
    for k, c in enumerate(codigos, 2):
        d = base.cct.get(c, {"desc": "NÃO CONSTA DA TABELA DA BASE", "ri": 0, "rc": 0, "conf": False})
        for j, v in enumerate([c, c[:3], d["desc"], d["ri"], d["rc"], "S" if d["conf"] else "N"], 1):
            x = wr.cell(row=k, column=j, value=v); x.font = fB; x.border = BD
            if j in (1, 2): x.number_format = "@"
            if j in (4, 5): x.number_format = "0%"
    ref_last = len(codigos) + 1
    r0 = ref_last + 3
    wr.cell(row=r0 - 1, column=1, value="TABELA CST IBS/CBS").font = fS
    for j, h in ((1, "CST"), (3, "DESCRICAO")):
        c = wr.cell(row=r0, column=j, value=h); c.font = fH; c.fill = NAVY; c.border = BD
    csts = sorted(base.cst.items())
    for k, (c, d) in enumerate(csts, r0 + 1):
        a = wr.cell(row=k, column=1, value=c); a.number_format = "@"
        b = wr.cell(row=k, column=3, value=d)
        for x in (a, b): x.font = fB; x.border = BD
    ci, cf = r0 + 1, r0 + len(csts)
    wr.cell(row=cf + 2, column=1, value=f"Versão da tabela: {versao}").font = Font(name=FN, size=9, italic=True)
    wr.freeze_panes = "A2"

    # ---------------- Cadastro_Fiscal
    wc = wb.create_sheet("Cadastro_Fiscal", 0)
    cabecalho(wc, 1, [(h, w) for _, h, w, _ in CAD], [GREY if c else NAVY for *_, c in CAD])
    RT = "Tabela_Referencia"
    for k, it in enumerate(itens, 2):
        r = it["res"]
        m = f"MATCH(${L['cct']}{k},{RT}!$A$2:$A${ref_last},0)"
        v = {"cod": it["codigo"], "desc": it["descricao"], "tipo": r["tipo"], "ncm": it["ncm"] or it["ncm_orig"],
             "nbs": it["nbs"], "cest": it["cest"], "ean": it["ean"], "un": it["unidade"], "cfop": it["cfop"],
             "cst": r["cst"], "cct": r["cct"], "trat": r["trat"], "lei": r["lei"], "cond": r["cond"], "is": r["is"],
             "cat": r["cat_is"], "st": r["status"], "obs": " | ".join(r["obs"]), "oreg": r["obs_regime"],
             "regra": r["regra"], "ant": "/".join(x for x in (it["cst_cli"], it["cct_cli"]) if x),
             "orig": it["origem"], "ver": versao,
             "dcst": f'=IFERROR(INDEX({RT}!$C${ci}:$C${cf},MATCH(${L["cst"]}{k},{RT}!$A${ci}:$A${cf},0)),"CST não localizado")',
             "dcct": f'=IF(${L["cct"]}{k}="","",IFERROR(INDEX({RT}!$C$2:$C${ref_last},{m}),"cClassTrib não localizado"))',
             "ri": f'=IF(${L["cct"]}{k}="","",IFERROR(INDEX({RT}!$D$2:$D${ref_last},{m}),""))',
             "rc": f'=IF(${L["cct"]}{k}="","",IFERROR(INDEX({RT}!$E$2:$E${ref_last},{m}),""))',
             "chk": f'=IF(${L["cct"]}{k}="","PENDENTE",IF(LEFT(${L["cct"]}{k},3)=${L["cst"]}{k},"OK","ERRO"))'}
        for j, (key, *_) in enumerate(CAD, 1):
            x = wc.cell(row=k, column=j, value=v[key] if v[key] != "" else None); x.font = fB; x.border = BD
            if key in TEXTO: x.number_format = "@"
            if key in ("ri", "rc"): x.number_format = "0%"
            if key in ("obs", "oreg"): x.alignment = WR
    wc.freeze_panes = "C2"
    wc.auto_filter.ref = f"A1:{get_column_letter(len(CAD))}{last}"
    cor_status(wc, f"{L['st']}2:{L['st']}{last}")
    wc.conditional_formatting.add(f"{L['chk']}2:{L['chk']}{last}", CellIsRule(operator="equal", formula=['"ERRO"'],
                                  fill=PatternFill("solid", fgColor="FFC7CE")))
    for f1, col in ((f"={RT}!$A${ci}:$A${cf}", "cst"), (f"={RT}!$A$2:$A${ref_last}", "cct"),
                    (f'"{ST_OK},{ST_CONF},{ST_REV},{ST_NCM}"', "st"), ('"Bem,Serviço"', "tipo"), ('"S,N"', "is")):
        dv = DataValidation(type="list", formula1=f1, allow_blank=True)
        wc.add_data_validation(dv); dv.add(f"{L[col]}2:{L[col]}{last}")

    # ---------------- Importacao_ERP (uma aba; ou duas visões para o Simples Nacional)
    CF = "Cadastro_Fiscal"
    abas_erp = [("Importacao_ERP", [])] if regime != "simples" else \
        [("Importacao_ERP_Unico", ERP_SIMPLES["Unico"]), ("Importacao_ERP_Hibrido", ERP_SIMPLES["Hibrido"])]
    for nome_aba, extras in abas_erp:
        we = wb.create_sheet(nome_aba)
        cols = [(h, w, src) for h, w, src in ERP[:-1]] + [(h, w, (tp, val)) for h, w, tp, val in extras] + [ERP[-1]]
        cabecalho(we, 1, [(h, w) for h, w, _ in cols])
        ri = L["ri"]
        for k in range(2, len(itens) + 2):
            for j, (h, _, src) in enumerate(cols, 1):
                fmt = "@"
                if isinstance(src, tuple):
                    tp, val = src
                    if tp == "const":
                        f = val
                    elif tp == "seg":
                        f = f'=IF({CF}!${ri}{k}="","",IF({CF}!${ri}{k}>0,"SIM","NAO"))'
                    else:
                        f = f'=IF({CF}!${ri}{k}="","",IF({CF}!${ri}{k}>=1,"SEM_DESTAQUE","INTEGRAL"))'
                elif src.endswith("%"):
                    col = L[src[:-1]]
                    f = f'=IF({CF}!${col}{k}="","",ROUND({CF}!${col}{k}*100,2))'; fmt = "0.00"
                elif src == "cat#":
                    f = f'=IF({CF}!${L["cat"]}{k}="","",TEXT(MATCH({CF}!${L["cat"]}{k},Listas!$A$2:$A$8,0),"00"))'
                else:
                    f = f'=IF({CF}!${L[src]}{k}="","",{CF}!${L[src]}{k}&"")'
                x = we.cell(row=k, column=j, value=f); x.font = fB; x.border = BD; x.number_format = fmt
        we.freeze_panes = "B2"
        we.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{last}"
        cor_status(we, f"{get_column_letter(len(cols))}2:{get_column_letter(len(cols))}{last}")

    # ---------------- Comparativo_Simples
    if regime == "simples":
        wq = wb.create_sheet("Comparativo_Simples")
        titulo(wq, "SIMPLES NACIONAL — REGIME ÚNICO (PURO) x REGIME HÍBRIDO", "C")
        cabecalho(wq, 3, [("Aspecto", 34), ("Simples puro (regime único)", 55), ("Simples híbrido (IBS/CBS pelo regime regular)", 55)])
        for k, ln in enumerate(COMPARATIVO, 4):
            for j, v in enumerate(ln, 1):
                x = wq.cell(row=k, column=j, value=v); x.font = fBB if j == 1 else fB; x.border = BD; x.alignment = WR
        k = 4 + len(COMPARATIVO) + 1
        wq.cell(row=k, column=1, value="Itens deste cliente").font = fS
        cnt = [("Itens classificados", f"=COUNTA({CF}!${L['cod']}$2:${L['cod']}${last})"),
               ("Itens com redução ou alíquota zero (segregar no PGDAS-D no regime único)", f'=COUNTIF({CF}!${ri}$2:${ri}${last},">0")'),
               ("Itens com alíquota zero (sem crédito a transferir no híbrido)", f'=COUNTIF({CF}!${ri}$2:${ri}${last},">=1")'),
               ("Itens com tributação integral", f'=COUNTIF({CF}!${L["trat"]}$2:${L["trat"]}${last},"Tributação integral")'),
               ("Itens sujeitos ao Imposto Seletivo", f'=COUNTIF({CF}!${L["is"]}$2:${L["is"]}${last},"S")')]
        for i, (a, b) in enumerate(cnt, k + 1):
            x = wq.cell(row=i, column=1, value=a); x.font = fB; x.border = BD; x.alignment = WR
            y = wq.cell(row=i, column=2, value=b); y.font = fBB; y.border = BD
        n = k + 1 + len(cnt) + 1
        wq.cell(row=n, column=1, value=("A escolha entre os regimes depende do faturamento B2B x B2C, das margens e do volume de compras "
                                         "com IBS/CBS a creditar. Para decidir com números, simular os dois cenários com os dados do cliente.")).font = Font(name=FN, size=9, italic=True)
        wq.merge_cells(start_row=n, start_column=1, end_row=n + 1, end_column=3)
        wq.cell(row=n, column=1).alignment = WR

    wl = wb.create_sheet("Listas")
    wl["A1"] = "CATEGORIA_IS"; wl["B1"] = "CODIGO"
    for k, c in enumerate(CAT_IS, 2):
        wl.cell(row=k, column=1, value=c); wl.cell(row=k, column=2, value=f"{k - 1:02d}").number_format = "@"
    wl.sheet_state = "hidden"

    # ---------------- Resumo
    ws = wb.create_sheet("Resumo", 2)
    titulo(ws, "RESUMO DA CLASSIFICAÇÃO", "E")
    for col, w in zip("ABCDE", [44, 12, 12, 34, 40]):
        ws.column_dimensions[col].width = w
    rng = lambda key: f"{CF}!${L[key]}$2:${L[key]}${last}"
    info = [("Cliente", args.cliente or "-"), ("Regime tributário", REGIMES[regime] if regime else (args.regime or "não informado")),
            ("Data da classificação", hoje), ("Versão da tabela cClassTrib", versao),
            ("Itens classificados", f"=COUNTA({rng('cod')})"), ("Itens sujeitos ao Imposto Seletivo", f'=COUNTIF({rng("is")},"S")')]
    for k, (a, b) in enumerate(info, 3):
        ws.cell(row=k, column=1, value=a).font = fBB
        x = ws.cell(row=k, column=2, value=b); x.font = fB; x.border = BD
    r = 10
    ws.cell(row=r, column=1, value="Por status").font = fS
    cabecalho(ws, r + 1, [("Status", 44), ("Qtde", 12), ("%", 12), ("Significado", 34)])
    for k, st in enumerate(SIGNIF, r + 2):
        for j, v in enumerate([st, f'=COUNTIF({rng("st")},A{k})', f"=IF($B$7=0,0,B{k}/$B$7)", SIGNIF[st]], 1):
            x = ws.cell(row=k, column=j, value=v); x.font = fB; x.border = BD
        ws.cell(row=k, column=3).number_format = "0.0%"
    cor_status(ws, f"A{r + 2}:A{r + 5}")
    r = r + 7
    ws.cell(row=r, column=1, value="Por tratamento tributário").font = fS
    cabecalho(ws, r + 1, [("Tratamento", 44), ("Qtde", 12), ("%", 12)])
    trats = sorted({i["res"]["trat"] for i in itens})
    for k, t in enumerate(trats, r + 2):
        for j, v in enumerate([t, f'=COUNTIF({rng("trat")},A{k})', f"=IF($B$7=0,0,B{k}/$B$7)"], 1):
            x = ws.cell(row=k, column=j, value=v); x.font = fB; x.border = BD
        ws.cell(row=k, column=3).number_format = "0.0%"
    r = r + 3 + len(trats)
    ws.cell(row=r, column=1, value="Por NCM (situação na data da classificação)").font = fS
    cabecalho(ws, r + 1, [("NCM / tipo", 44), ("Qtde itens", 12), ("Classificações distintas", 12),
                          ("Classificação predominante (CST/cClassTrib)", 34), ("Situação", 40)])
    por_ncm = defaultdict(list)
    for it in itens:
        chave = it["ncm"] if len(it["ncm"]) == 8 else ("(serviço)" if it["res"]["tipo"] == "Serviço" else "(sem NCM válido)")
        por_ncm[chave].append(f"{it['res']['cst']}/{it['res']['cct']}")
    for k, (ncm, cl) in enumerate(sorted(por_ncm.items()), r + 2):
        cnt = Counter(cl)
        for j, v in enumerate([ncm, len(cl), len(cnt), cnt.most_common(1)[0][0],
                               "Classificações divergentes no mesmo NCM" if len(cnt) > 1 and ncm[0] != "(" else "Uniforme" if len(cnt) == 1 else "-"], 1):
            x = ws.cell(row=k, column=j, value=v); x.font = fB; x.border = BD
        ws.cell(row=k, column=1).number_format = "@"

    # ---------------- Divergencias_Riscos
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

    # ---------------- Leiaute_Importacao
    wy = wb.create_sheet("Leiaute_Importacao", 4)
    titulo(wy, "LEIAUTE PARA IMPORTAÇÃO NO ERP — " + ("ABAS Importacao_ERP_Unico / _Hibrido" if regime == "simples" else "ABA Importacao_ERP"), "F")
    cabecalho(wy, 3, [("Campo", 20), ("Descrição", 46), ("Tipo", 12), ("Tamanho", 10), ("Obrigatório", 12),
                      ("Destino no documento fiscal / observação", 70)])
    lay = LEIAUTE[:-1] + (LEIAUTE_SIMPLES if regime == "simples" else []) + LEIAUTE[-1:]
    for k, ln in enumerate(lay, 4):
        for j, v in enumerate(ln, 1):
            x = wy.cell(row=k, column=j, value=v); x.font = fB; x.border = BD; x.alignment = WR
    k = 4 + len(lay) + 1
    for t in ORIENTACOES:
        wy.cell(row=k, column=1, value=t).font = fBB if t.endswith(":") else fB; k += 1

    # ---------------- Instrucoes
    wi = wb.create_sheet("Instrucoes", 0)
    titulo(wi, "CLASSIFICAÇÃO TRIBUTÁRIA DE PRODUTOS E SERVIÇOS — IBS / CBS / IS", "C")
    wi.column_dimensions["A"].width = 3; wi.column_dimensions["B"].width = 32; wi.column_dimensions["C"].width = 100
    escr, resp = base.cfg.get("ESCRITORIO", ""), base.cfg.get("RESPONSAVEL", "")
    blocos = instrucoes(args.cliente, REGIMES[regime] if regime else (args.regime or "não informado"),
                        " – ".join(x for x in (escr, resp) if x) or "-", hoje, os.path.basename(args.arquivo.rstrip("/\\")),
                        versao, regime)
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

    ordem = ["Instrucoes", "Cadastro_Fiscal", "Importacao_ERP", "Importacao_ERP_Unico", "Importacao_ERP_Hibrido",
             "Comparativo_Simples", "Resumo", "Divergencias_Riscos", "Leiaute_Importacao", "Tabela_Referencia", "Listas"]
    wb._sheets.sort(key=lambda w: ordem.index(w.title))
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(destino)


LEIAUTE = [
    ("COD_ITEM", "Código do produto/serviço no ERP do cliente", "Texto", "até 60", "Sim", "Chave para localizar o item no cadastro (cProd)."),
    ("DESCRICAO", "Descrição", "Texto", "até 120", "Sim", "Apenas conferência; não sobrescrever a descrição do ERP."),
    ("TIPO_ITEM", "Bem ou Serviço", "Texto", "até 7", "Sim", "Bem: NF-e/NFC-e. Serviço: NFS-e."),
    ("NCM", "NCM/SH (bens)", "Texto", "8", "Bens", "Somente dígitos. Atualizar no cadastro se divergir."),
    ("NBS", "NBS (serviços)", "Texto", "9", "Serviços", "Somente dígitos, quando informada pelo cliente."),
    ("CEST", "CEST", "Texto", "7", "Não", "ICMS-ST (mantido na transição)."),
    ("GTIN", "Código de barras", "Texto", "8–14", "Não", "cEAN."),
    ("UNIDADE", "Unidade comercial", "Texto", "até 6", "Não", "uCom."),
    ("CST_IBS_CBS", "CST do IBS e da CBS", "Texto", "3", "Sim", "Grupo IBSCBS, tag CST (NT 2025.002-RTC). Manter zeros à esquerda."),
    ("CCLASSTRIB", "Código de classificação tributária", "Texto", "6", "Sim", "Grupo IBSCBS, tag cClassTrib. Os 3 primeiros dígitos = CST."),
    ("PRED_IBS_PERC", "% de redução da alíquota do IBS", "Decimal", "3,2", "Sim", "60.00 = redução de 60%; 100.00 = alíquota zero. Base para o percentual de redução do grupo de redução da NT."),
    ("PRED_CBS_PERC", "% de redução da alíquota da CBS", "Decimal", "3,2", "Sim", "Idem para CBS."),
    ("SUJEITO_IS", "Incidência do Imposto Seletivo", "Texto", "1", "Sim", "S/N. Se S, parametrizar o grupo IS (CST-IS/cClassTribIS) conforme tabela própria."),
    ("COD_CATEGORIA_IS", "Categoria do IS", "Texto", "2", "Se IS=S", "01 Veículos; 02 Embarcações e aeronaves; 03 Fumígenos; 04 Bebidas alcoólicas; 05 Bebidas açucaradas; 06 Bens minerais; 07 Prognósticos/fantasy sport."),
    ("STATUS", "Situação da classificação", "Texto", "até 12", "-", "Importar somente itens com STATUS = OK."),
]
ORIENTACOES = [
    "Orientações aos programadores:",
    "• Ler a aba Importacao_ERP a partir da linha 2 (linha 1 = cabeçalho). Todos os códigos são texto: preservar zeros à esquerda (ex.: 000001).",
    "• A chave de atualização é COD_ITEM. Itens com STATUS diferente de OK não devem ser importados.",
    "• As alíquotas de IBS/CBS não constam do arquivo: são definidas por ente (União, UF e Município) e pelo ano da transição.",
    "• Para gerar CSV: Arquivo > Salvar como > CSV UTF-8, com a aba Importacao_ERP ativa.",
]


def instrucoes(cliente, regime_txt, escr, hoje, arquivo, versao, regime):
    blocos = [
        ("IDENTIFICAÇÃO", [("Cliente", cliente or "-"), ("Regime tributário", regime_txt), ("Escritório / responsável", escr),
                           ("Data", hoje), ("Arquivo de origem", arquivo), ("Versão da tabela cClassTrib", versao)]),
        ("BASE NORMATIVA", [("Constituição", "EC 132/2023 (arts. 156-A, 195, V, e 153, VIII, da CF; ADCT)."),
                            ("Lei Complementar", "LC 214/2025, com as alterações da LC 227/2026 e da LC 235/2026 (redação consolidada)."),
                            ("Tabelas", "Informe Técnico 2025.002 do Portal Nacional da NF-e (CST e cClassTrib) e NT 2025.002-RTC (leiaute do grupo IBSCBS).")]),
    ]
    if regime in NOTAS_REGIME:
        blocos.append(("REGIME DO CLIENTE", NOTAS_REGIME[regime]))
    blocos += [
        ("CONTEÚDO DO ARQUIVO", [("Cadastro_Fiscal", "Classificação item a item: CST, cClassTrib, reduções, tratamento, base legal, condição do benefício, IS, status e observações."),
                                 ("Importacao_ERP" if regime != "simples" else "Importacao_ERP_Unico / _Hibrido",
                                  "Aba(s) para os programadores: somente os campos de importação, em formato técnico." if regime != "simples" else
                                  "Uma aba por modalidade do Simples (puro e híbrido). Os programadores importam a do regime adotado."),
                                 *([("Comparativo_Simples", "Diferenças entre Simples puro e híbrido e contagem de itens afetados.")] if regime == "simples" else []),
                                 ("Resumo", "Totais por status, por tratamento e por NCM."),
                                 ("Divergencias_Riscos", "Itens que exigem atenção, risco envolvido e ação recomendada."),
                                 ("Leiaute_Importacao", "Especificação dos campos para o ERP."),
                                 ("Tabela_Referencia", "Tabelas de CST e cClassTrib utilizadas.")]),
        ("LEGENDA DE STATUS", [(s, SIGNIF[s]) for s in SIGNIF]),
        ("PREMISSAS E LIMITAÇÕES", [
            ("Critério", "Bens: enquadramento pelo NCM informado combinado com a descrição. Serviços: pela NBS ou pela descrição. A correção do NCM/NBS é de responsabilidade do cliente."),
            ("Abordagem conservadora", "Quando o NCM consta de anexo com benefício, mas a descrição não permite confirmar o enquadramento, aplica-se tributação integral (CST 000) e o item é marcado para revisão."),
            ("Benefícios condicionados", "Reduções que dependem do adquirente, da finalidade ou de registro (Anvisa, MAPA, conselho profissional) estão indicadas na coluna CONDICAO_BENEFICIO."),
            ("Atualização", "A classificação deve ser revalidada a cada nova versão da tabela cClassTrib e a cada alteração da LC 214/2025."),
        ]),
    ]
    return blocos


# =============================================================================== divergências
def montar_divergencias(itens):
    out = []
    risco = {ST_NCM: ("Rejeição do documento fiscal e enquadramento incorreto em toda a cadeia",
                      "Cliente deve informar/corrigir o NCM (bens) ou a NBS (serviços); reprocessar o item"),
             ST_REV: ("Recolhimento a maior (benefício não aproveitado) ou glosa/autuação se o benefício for aplicado indevidamente",
                      "Analisar a descrição completa/ficha técnica e confirmar o enquadramento no anexo"),
             ST_CONF: ("Uso de regra ou código ainda não validado com a tabela oficial vigente",
                       "Escritório deve conferir a regra/código na base de regras antes da importação")}
    for it in itens:
        st = it["res"]["status"]
        if st != ST_OK:
            out.append([it["codigo"], it["descricao"], it["ncm"] or it["ncm_orig"], st, " | ".join(it["res"]["obs"]), *risco[st]])
    por_cod = defaultdict(set)
    for it in itens:
        if it["codigo"]:
            por_cod[it["codigo"]].add(it["ncm"] or it["ncm_orig"])
    for c, ncms in por_cod.items():
        if len(ncms) > 1:
            out.append([c, "", "; ".join(sorted(ncms)), ST_REV, "Mesmo código de produto com NCMs diferentes no arquivo do cliente",
                        "Cadastro inconsistente; risco de classificação divergente entre documentos", "Cliente deve unificar o NCM do produto"])
    por_ncm = defaultdict(set)
    for it in itens:
        if len(it["ncm"]) == 8:
            por_ncm[it["ncm"]].add((it["res"]["cst"], it["res"]["cct"]))
    for ncm, cl in sorted(por_ncm.items()):
        if len(cl) > 1:
            out.append(["(vários)", "", ncm, ST_REV, "Mesmo NCM com classificações diferentes: " + "; ".join(f"{a}/{b}" for a, b in sorted(cl)),
                        "Pode ser correto (descrições distintas), mas exige justificativa item a item",
                        "Conferir a descrição de cada item frente ao texto do anexo"])
    for it in itens:
        if (it["cst_cli"] or it["cct_cli"]) and (it["cst_cli"], it["cct_cli"]) != (it["res"]["cst"], it["res"]["cct"]):
            out.append([it["codigo"], it["descricao"], it["ncm"], it["res"]["status"],
                        f"Cliente já usa {it['cst_cli']}/{it['cct_cli']}; proposta {it['res']['cst']}/{it['res']['cct']}",
                        "Documentos já emitidos podem estar com classificação incorreta",
                        "Validar e, se for o caso, corrigir o cadastro e avaliar documentos emitidos"])
    return out


# =============================================================================== main
def processar(itens, base, regime):
    vistos, unicos = set(), []
    for it in itens:
        it["ncm"] = digitos(it["ncm_orig"])
        it["ncm_ajustado"] = len(it["ncm"]) == 7  # NCM gravado como número perde o zero à esquerda
        if it["ncm_ajustado"]:
            it["ncm"] = "0" + it["ncm"]
        chave = (it["codigo"] or it["descricao"], it["ncm"])
        if chave in vistos:
            continue
        vistos.add(chave); unicos.append(it)
    for it in unicos:
        it["res"] = classificar(it, base, regime)
    return unicos


def main():
    ap = argparse.ArgumentParser(description="Classificação de produtos e serviços IBS/CBS/IS a partir do cadastro do cliente.")
    ap.add_argument("arquivo", help="Excel/CSV do cliente, XML de NF-e ou pasta com XMLs")
    ap.add_argument("--cliente", default="")
    ap.add_argument("--regime", default="", help="Simples Nacional (gera puro e híbrido) | MEI | Lucro Presumido | Lucro Real")
    ap.add_argument("--base", default=BASE_PADRAO)
    ap.add_argument("--saida", default=None)
    args = ap.parse_args()

    if not os.path.exists(args.arquivo):
        sys.exit(f"Arquivo não encontrado: {args.arquivo}")
    base = Base(args.base)
    regime = regime_chave(args.regime)
    if os.path.isdir(args.arquivo) or args.arquivo.lower().endswith(".xml"):
        itens, avisos = ler_xmls(args.arquivo)
    elif args.arquivo.lower().endswith((".xlsx", ".xlsm", ".csv")):
        itens, avisos = ler_planilha(args.arquivo, base)
    else:
        sys.exit("Formato não suportado. Use .xlsx, .xlsm, .csv ou XML de NF-e (.xls: salve antes como .xlsx).")
    if not itens:
        sys.exit("Nenhum item encontrado. " + " ".join(avisos))

    unicos = processar(itens, base, regime)
    div = montar_divergencias(unicos)
    nome_cli = re.sub(r"[^A-Za-z0-9_-]+", "_", args.cliente or os.path.splitext(os.path.basename(args.arquivo.rstrip("/\\")))[0]).strip("_")
    pasta = args.saida or os.path.join(os.path.dirname(os.path.abspath(args.arquivo.rstrip("/\\"))), "saida")
    os.makedirs(pasta, exist_ok=True)
    destino = os.path.join(pasta, f"{nome_cli}_Classificacao_IBS_CBS_{dt.date.today():%Y%m%d}.xlsx")
    gerar_excel(unicos, base, args, regime, div, destino)

    cont = Counter(i["res"]["status"] for i in unicos)
    print("\n".join(avisos))
    print(f"\nRegime: {REGIMES[regime]}")
    print(f"Itens lidos: {len(itens)} | únicos classificados: {len(unicos)}")
    for st in (ST_OK, ST_CONF, ST_REV, ST_NCM):
        print(f"  {st:<12} {cont.get(st, 0)}")
    print(f"Divergências/riscos listados: {len(div)}")
    print(f"\nArquivo gerado: {destino}")


if __name__ == "__main__":
    main()
