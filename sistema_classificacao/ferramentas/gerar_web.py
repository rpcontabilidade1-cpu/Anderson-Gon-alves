#!/usr/bin/env python3
"""
Monta web/classificador.html a partir de:
  - web/template.html  (interface)
  - web/motor.js       (motor de classificação, igual ao classificar.py)
  - base_regras.xlsx   (base embutida na página)
  - exemplos/cliente_exemplo.xlsx (exemplo exibido ao abrir)

Rode sempre que alterar a base de regras para atualizar a base embutida:
    python ferramentas/gerar_web.py
(Na página também é possível carregar uma base_regras.xlsx própria sem regenerar.)
"""
import datetime as dt
import json
import os
from openpyxl import load_workbook

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ABAS_BASE = ["Regras_NCM", "Regras_IS", "Tabela_cClassTrib", "Tabela_CST", "Sinonimos_Colunas", "Config"]


def valor(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return v.isoformat()[:10]
    return v


def base_json(caminho):
    wb = load_workbook(caminho, read_only=True, data_only=True)
    out = {}
    for aba in ABAS_BASE:
        rows = list(wb[aba].iter_rows(values_only=True))
        cab = [str(c) if c is not None else "" for c in rows[0]]
        out[aba] = [{k: valor(v) for k, v in zip(cab, r) if k and v is not None}
                    for r in rows[1:] if any(v not in (None, "") for v in r)]
    wb.close()
    return out


def exemplo_json(caminho):
    wb = load_workbook(caminho, read_only=True, data_only=True)
    out = {ws.title: [[valor(v) for v in r] for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
    wb.close()
    return out


def textos_json():
    import sys
    sys.path.insert(0, RAIZ)
    import classificar as C
    tok = ("{{CLIENTE}}", "{{REGIME}}", "{{ESCR}}", "{{DATA}}", "{{ARQ}}", "{{VERSAO}}")
    return {
        "CAD": C.CAD, "ERP": C.ERP, "LEIAUTE": C.LEIAUTE, "ORIENTACOES": C.ORIENTACOES, "SIGNIF": C.SIGNIF,
        "ERP_SIMPLES": C.ERP_SIMPLES, "LEIAUTE_SIMPLES": C.LEIAUTE_SIMPLES, "COMPARATIVO": C.COMPARATIVO,
        "INSTRUCOES": {k: C.instrucoes(*tok, k) for k in C.REGIMES},
    }


def main():
    web = os.path.join(RAIZ, "web")
    tpl = open(os.path.join(web, "template.html"), encoding="utf-8").read()
    motor = open(os.path.join(web, "motor.js"), encoding="utf-8").read()
    base = json.dumps(base_json(os.path.join(RAIZ, "base_regras.xlsx")), ensure_ascii=False)
    ex = json.dumps(exemplo_json(os.path.join(RAIZ, "exemplos", "cliente_exemplo.xlsx")), ensure_ascii=False)
    html = (tpl.replace("/*__MOTOR__*/", motor)
               .replace("/*__BASE__*/null", base.replace("</", "<\\/"))
               .replace("/*__EXEMPLO__*/null", ex.replace("</", "<\\/"))
               .replace("/*__TEXTOS__*/null", json.dumps(textos_json(), ensure_ascii=False).replace("</", "<\\/")))
    destino = os.path.join(web, "classificador.html")
    open(destino, "w", encoding="utf-8").write(html)
    print("Gerado:", destino, f"({len(html) // 1024} KB)")


if __name__ == "__main__":
    main()
