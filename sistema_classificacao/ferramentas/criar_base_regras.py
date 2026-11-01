#!/usr/bin/env python3
"""
Gera a base_regras.xlsx inicial do sistema de classificação IBS/CBS/IS.

ATENÇÃO: a base inicial foi montada SEM acesso ao texto oficial da LC 214/2025 e à
tabela cClassTrib vigente (IT 2025.002). Por isso TODAS as regras e códigos saem com
CONFERIDO = N, e o sistema marcará os itens enquadrados por elas como "A CONFIRMAR".
Após conferir cada regra com a redação consolidada da LC 214/2025 (alterada pelas
LC 227/2026 e LC 235/2026) e com a tabela oficial, altere CONFERIDO para S na própria
base_regras.xlsx. Não é preciso rodar este script de novo (ele sobrescreve a base).
"""
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

DESTINO = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "base_regras.xlsx")

F = "Arial"
H = Font(name=F, bold=True, color="FFFFFF", size=10)
B = Font(name=F, size=10)
BB = Font(name=F, size=10, bold=True)
FH = PatternFill("solid", fgColor="1F3864")
thin = Side(style="thin", color="BFBFBF")
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
WR = Alignment(wrap_text=True, vertical="top")

# ------------------------------------------------------------------------------------------------
# Regras de enquadramento por NCM
# NCM / EXCETO_NCM: prefixos (2 a 8 dígitos) separados por ';'. Vale o prefixo mais longo.
# PALAVRAS_CHAVE: a descrição deve conter ao menos uma (sem acento, minúsculas). Vazio = qualquer.
# PALAVRAS_EXCLUIR: se a descrição contiver alguma, a regra não se aplica (item vai p/ REVISAR).
# PRIORIDADE: ordem da árvore de decisão (1 imunidade, 2 isenção, 3 alíquota zero,
#             4 redução, 5 diferimento/suspensão, 6 monofásica, 7 regime específico).
# ------------------------------------------------------------------------------------------------
A1 = ("Alíquota zero", "200", "200003", 1, 1, "LC 214/2025, art. 125, Anexo I", 3)
A15 = ("Alíquota zero", "200", "200014", 1, 1, "LC 214/2025, art. 148, Anexo XV", 3)
A8 = ("Redução 60%", "200", "200035", 0.6, 0.6, "LC 214/2025, art. 136, Anexo VIII", 4)
A9 = ("Redução 60%", "200", "200038", 0.6, 0.6, "LC 214/2025, art. 138, Anexo IX", 4)

def R(i, ncm, desc_regra, anexo, base, exceto="", chave="", excluir="", cond="", obs="", nbs="", tipo="Bem"):
    trat, cst, cct, ri, rc, lei, prio = base
    return [i, ncm, exceto, chave, excluir, desc_regra, anexo, trat, cst, cct, ri, rc, lei, cond, prio, "N", obs, nbs, tipo]

A7 = ("Redução 60%", "200", "200034", 0.6, 0.6, "LC 214/2025, art. 135, Anexo VII", 4)
AGRO = ("Redução 60%", "200", "200036", 0.6, 0.6, "LC 214/2025, art. 137", 4)
A4 = ("Redução 60%", "200", "200030", 0.6, 0.6, "LC 214/2025, art. 131, Anexo IV", 4)
A5 = ("Redução 60%", "200", "200031", 0.6, 0.6, "LC 214/2025, art. 132, Anexo V", 4)
A6 = ("Redução 60%", "200", "200033", 0.6, 0.6, "LC 214/2025, art. 134, Anexo VI", 4)

REGRAS = [
    # ---------------- Anexo I – Cesta Básica Nacional (alíquota zero)
    R("R001", "100620;100630;100640", "Arroz", "I", A1),
    R("R002", "04011010;04011090;04012010;04012090;04014010;04015010", "Leite fluido pasteurizado ou UHT", "I", A1, excluir="leite cru;in natura"),
    R("R003", "04021010;04021090;04022110;04022120;04022910;04022920", "Leite em pó", "I", A1),
    R("R004", "19011010;19011090", "Fórmulas infantis", "I", A1, chave="infantil;formula", cond="Fórmulas infantis definidas pela Anvisa"),
    R("R005", "04051000", "Manteiga", "I", A1),
    R("R006", "15171000", "Margarina", "I", A1),
    R("R007", "07133319;07133329;07133399;07133590", "Feijões", "I", A1),
    R("R008", "09011;09012;21011", "Café", "I", A1, excluir="capsula", cond="Exceto café em cápsulas"),
    R("R009", "150790", "Óleo de soja", "I", A1, chave="soja"),
    R("R010", "11062000", "Farinha de mandioca", "I", A1),
    R("R011", "11022000;11031300;11041900;11042300", "Farinha, grumos, sêmolas e grãos de milho", "I", A1, chave="milho;fuba;flocao;canjica;canjiquinha"),
    R("R012", "11010010", "Farinha de trigo", "I", A1),
    R("R013", "17011400;17019900", "Açúcar", "I", A1),
    R("R014", "190211;190219", "Massas alimentícias", "I", A1, excluir="recheada;cozida", cond="Massas não cozidas nem recheadas"),
    R("R015", "19059090", "Pão comum (francês)", "I", A1, chave="pao frances;pao comum;pao de sal;pao carioca", cond="Pão do tipo comum: farinha de cereais, fermento, água e sal"),
    R("R016", "11041200;11029000", "Aveia (grãos e farinha)", "I", A1, chave="aveia"),
    R("R017", "0201;0202;0203;0204;0206;0207;0209;02101;021099", "Carnes bovina, suína, ovina, caprina e de aves e miudezas", "I", A1,
      exceto="02074300;02075300", excluir="foie gras;figado gordo"),
    R("R018", "0302;0303;0304", "Peixes e carnes de peixes", "I", A1, exceto="03021;03023;03031;03034",
      excluir="salmao;truta;atum;bacalhau;hadoque;saithe;ova", cond="Exceto salmonídeos, atuns, bacalhaus, hadoque, saithe e ovas"),
    R("R019", "040610;04062000;04069010;04069020;04069030;04069090", "Queijos (tipos listados no Anexo I)", "I", A1,
      chave="mussarela;mucarela;mozarela;minas;prato;coalho;ricota;requeijao;provolone;parmesao;queijo fresco;reino"),
    R("R020", "25010020;25010090", "Sal de mesa", "I", A1, chave="sal"),
    R("R021", "09030010", "Mate", "I", A1),
    R("R022", "0714", "Raízes e tubérculos (mandioca etc.)", "I", A1, excluir="cozid;congelad;frit"),
    # ---------------- Anexo XV – hortícolas, frutas e ovos (alíquota zero)
    R("R030", "0701;0702;0703;0704;0705;0706;0707;0708;0709", "Produtos hortícolas frescos ou refrigerados", "XV", A15,
      excluir="cozid;conserva;desidratad;frit;temperad", cond="Frescos ou refrigerados; não cozidos"),
    R("R031", "0710", "Produtos hortícolas congelados", "XV", A15, excluir="cozid;frit;temperad", cond="Congelados, não cozidos"),
    R("R032", "0803;0804;0805;0806;0807;0808;0809;0810", "Frutas frescas ou refrigeradas", "XV", A15,
      excluir="seca;desidratad;cristalizad;conserva", cond="Frescas ou refrigeradas"),
    R("R033", "040721;040729;040790", "Ovos", "XV", A15),
    # ---------------- Saúde menstrual (alíquota zero)
    R("R040", "96190000", "Absorventes, tampões e coletores menstruais", "-",
      ("Alíquota zero", "200", "200013", 1, 1, "LC 214/2025, art. 147", 3),
      chave="absorvente;tampao;coletor menstrual;calcinha absorvente"),
    # ---------------- Anexo VIII – higiene pessoal e limpeza (60%)
    R("R050", "34011190", "Sabonetes", "VIII", A8, chave="sabonete"),
    R("R051", "33061000", "Dentifrícios (creme dental)", "VIII", A8),
    R("R052", "96032100", "Escovas de dentes", "VIII", A8),
    R("R053", "48181000", "Papel higiênico", "VIII", A8),
    R("R054", "28289011", "Água sanitária", "VIII", A8),
    R("R055", "34011900", "Sabões em barra", "VIII", A8, chave="sabao"),
    R("R056", "96190000", "Fraldas", "VIII", A8, chave="fralda", obs="Confirmar se fraldas constam do Anexo VIII na redação vigente"),
    # ---------------- Medicamentos (60% – art. 133; zero se Anexo XIV)
    R("R060", "3003;3004", "Medicamentos", "-", ("Redução 60%", "200", "200032", 0.6, 0.6, "LC 214/2025, art. 133", 4),
      cond="Registro na Anvisa ou manipulação. Verificar se o princípio ativo consta do Anexo XIV (alíquota zero, cClassTrib próprio)"),
    # ---------------- Anexo IX – insumos agropecuários (60%)
    R("R070", "3101;3102;3103;3104;3105", "Fertilizantes", "IX", A9, cond="Registro no MAPA / destinação agropecuária"),
    R("R071", "1209", "Sementes para semeadura", "IX", A9, cond="Sementes certificadas/registradas no MAPA"),
    R("R072", "230990", "Rações e preparações para alimentação animal", "IX", A9, chave="racao;suplemento;nucleo;premix", cond="Uso na alimentação animal"),
    # ---------------- Imunidade
    R("R080", "4901;4902;4903", "Livros, jornais e periódicos", "-",
      ("Imunidade", "410", "", 1, 1, "CF, art. 150, VI, d; LC 214/2025, art. 9º", 1),
      obs="Informar o cClassTrib de imunidade da tabela oficial vigente"),
    # ---------------- Monofásica (combustíveis)
    R("R090", "271012;271019;2207;271119", "Combustíveis (gasolina, diesel, etanol, GLP)", "-",
      ("Monofásica", "620", "620001", 0, 0, "LC 214/2025, arts. 172 e ss.", 6),
      chave="gasolina;diesel;etanol;alcool combustivel;glp;gas liquefeito;querosene", cond="Tributação monofásica com alíquota ad rem"),

    # ---------------- Anexo VII – alimentos destinados ao consumo humano (60%)
    R("R100", "0306;0307", "Crustáceos (exceto lagostas e lagostins) e moluscos", "VII", A7, excluir="lagost"),
    R("R101", "0403", "Leite fermentado, iogurtes e bebidas lácteas", "VII", A7),
    R("R102", "22029900", "Bebidas e compostos lácteos", "VII", A7, chave="lacteo;lactea;leite"),
    R("R103", "04090000", "Mel natural", "VII", A7),
    R("R104", "1102;1103;1104", "Farinhas, grumos, sêmolas e grãos de cereais", "VII", A7),
    R("R105", "11081200", "Amido de milho", "VII", A7),
    R("R106", "19030000", "Tapioca e seus sucedâneos", "VII", A7),
    R("R107", "1507;1508;1509;1510;1511;1512;1513;1514;1515", "Óleos vegetais", "VII", A7, cond="Óleo de soja enquadra-se no Anexo I (regra R009)"),
    R("R108", "19059010", "Pão de forma", "VII", A7, chave="pao de forma;pao forma"),
    R("R109", "20029000", "Extrato de tomate", "VII", A7, chave="extrato;polpa;molho"),
    R("R110", "2009", "Sucos naturais de frutas ou hortícolas", "VII", A7, excluir="adocad;com acucar;nectar;refresco",
      cond="Sem adição de açúcar ou edulcorantes e sem conservantes"),
    R("R111", "0811;2008", "Polpas de frutas", "VII", A7, chave="polpa", excluir="com acucar;adocad", cond="Sem adição de açúcar"),
    # ---------------- Produtos agropecuários in natura (60% – art. 137)
    R("R120", "0101;0102;0103;0104;0105;0106", "Animais vivos", "-", AGRO),
    R("R121", "1001;1002;1003;1004;1005;1007;1008;100610", "Cereais em grão", "-", AGRO),
    R("R122", "1201;1202;1204;1205;1206;1207", "Soja e outras sementes oleaginosas em grão", "-", AGRO),
    R("R123", "04011;04012", "Leite cru (in natura)", "-", AGRO, chave="cru;in natura", cond="Leite cru; pasteurizado/UHT segue Anexo I"),
    R("R124", "4401;4403", "Produtos florestais in natura (madeira em bruto, lenha)", "-", AGRO),
    R("R125", "5201", "Algodão não cardado nem penteado", "-", AGRO),
    R("R126", "09011110;09011190", "Café cru em grão", "-", AGRO, cond="Café cru: verificar se prevalece Anexo I (R008)"),
    # ---------------- Dispositivos médicos / acessibilidade / nutrição
    R("R130", "9018;9019;9020;9022", "Instrumentos e aparelhos médicos, odontológicos e hospitalares", "IV", A4,
      cond="Somente NCMs/descrições do Anexo IV; parte consta do Anexo XII (alíquota zero)"),
    R("R131", "9021", "Artigos e aparelhos ortopédicos, próteses, aparelhos auditivos", "IV", A4,
      cond="Verificar Anexos IV, V, XII e XIII conforme o item"),
    R("R132", "30051;30059;3006", "Curativos, materiais de sutura e outros produtos farmacêuticos", "IV", A4),
    R("R133", "40151100;40151200", "Luvas cirúrgicas e de procedimento", "IV", A4),
    R("R134", "8713", "Cadeiras de rodas e veículos para pessoas com deficiência", "V", A5, cond="Verificar Anexo XIII (alíquota zero)"),
    R("R135", "21069030;21069090", "Composições para nutrição enteral e parenteral / fórmulas especiais", "VI", A6,
      chave="enteral;parenteral;dieta;nutricao"),
    # ---------------- Mais higiene e limpeza / insumos
    R("R140", "34022", "Detergentes e preparações para lavagem", "VIII", A8, chave="detergente;lava louca;lava roupa;sabao em po;sabao liquido"),
    R("R141", "3808", "Defensivos agropecuários (inseticidas, herbicidas, fungicidas)", "IX", A9,
      chave="herbicida;fungicida;inseticida agri;defensivo;agrotoxico;acaricida", excluir="domestic;aerossol;repelente",
      cond="Uso agropecuário com registro no MAPA"),
    R("R142", "30024", "Vacinas para uso veterinário", "IX", A9, chave="veterin;bovin;suin;aves;animal"),
    # ---------------- SERVIÇOS (identificados por palavras-chave quando o item não tem NCM, ou por NBS)
    R("S001", "", "Serviços de educação (ensino regular, técnico, superior)", "II",
      ("Redução 60%", "200", "200028", 0.6, 0.6, "LC 214/2025, art. 129, Anexo II", 4),
      chave="mensalidade;matricula;ensino;educacao infantil;creche;pre-escola;graduacao;pos-graduacao;curso tecnico",
      cond="Somente serviços listados no Anexo II", tipo="Serviço"),
    R("S002", "", "Serviços de saúde humana", "III",
      ("Redução 60%", "200", "200029", 0.6, 0.6, "LC 214/2025, art. 130, Anexo III", 4),
      chave="consulta;exame;cirurgi;internacao;fisioterap;odontolog;psicolog;fonoaudiolog;nutricionis;hemodialise;laboratorio clinico;vacinacao;terapia ocupacional",
      cond="Somente serviços listados no Anexo III", tipo="Serviço"),
    R("S003", "", "Profissões intelectuais regulamentadas (30%)", "-",
      ("Redução 30%", "200", "200052", 0.3, 0.3, "LC 214/2025, art. 127", 4),
      chave="honorario;contabil;advocaci;juridic;engenhari;arquitet;economista;administracao de empresas;medico veterinario;agronom;auditoria",
      cond="Profissional/sociedade sujeito a conselho profissional, conforme art. 127 e condições legais", tipo="Serviço"),
    R("S004", "", "Transporte público coletivo de passageiros urbano/metropolitano", "-",
      ("Isenção", "400", "400001", 1, 1, "LC 214/2025, art. 157", 2),
      chave="transporte coletivo;passagem urbana;transporte urbano;metro;trem urbano", tipo="Serviço"),
    R("S005", "", "Produções artísticas, culturais, jornalísticas e audiovisuais", "X",
      ("Redução 60%", "200", "200039", 0.6, 0.6, "LC 214/2025, art. 139, Anexo X", 4),
      chave="espetaculo;teatro;show musical;producao audiovisual;exibicao cinematografica;danca;circo",
      cond="Produções nacionais listadas no Anexo X", tipo="Serviço"),
    R("S006", "", "Atividades desportivas", "-",
      ("Redução 60%", "200", "", 0.6, 0.6, "LC 214/2025, art. 141", 4),
      chave="atividade desportiva;evento esportivo;competicao esportiva",
      cond="Conforme art. 141", obs="Informar o cClassTrib da tabela oficial", tipo="Serviço"),
]
COLS_REGRAS = ["ID", "NCM", "EXCETO_NCM", "PALAVRAS_CHAVE", "PALAVRAS_EXCLUIR", "DESCRICAO_REGRA", "ANEXO",
               "TRATAMENTO", "CST", "CCLASSTRIB", "PRED_IBS", "PRED_CBS", "BASE_LEGAL", "CONDICAO",
               "PRIORIDADE", "CONFERIDO", "OBSERVACAO", "NBS", "TIPO_ITEM"]

IS = [
    ("IS01", "8703", "", "Veículos", "Automóveis de passageiros. Verificar exceções (PcD, taxistas, veículos de carga)"),
    ("IS02", "8802", "", "Embarcações e aeronaves", "Aeronaves"),
    ("IS03", "8903", "", "Embarcações e aeronaves", "Embarcações de recreio/esporte"),
    ("IS04", "2402;2403", "", "Produtos fumígenos", ""),
    ("IS05", "2203;2204;2205;2206;2208", "", "Bebidas alcoólicas", ""),
    ("IS06", "22021000", "", "Bebidas açucaradas", "Bebidas com adição de açúcar"),
    ("IS07", "2601;2709;271111;271121", "", "Bens minerais", "Minério de ferro, petróleo e gás natural (alíquota máxima 0,25%)"),
]

CST = [
    ("000", "Tributação integral"), ("010", "Tributação com alíquotas uniformes"),
    ("011", "Tributação com alíquotas uniformes reduzidas"), ("200", "Alíquota reduzida"),
    ("210", "Redução de alíquota com redutor de base de cálculo"), ("220", "Alíquota fixa"),
    ("221", "Alíquota fixa proporcional"), ("222", "Redução de base de cálculo"), ("400", "Isenção"),
    ("410", "Imunidade e não incidência"), ("510", "Diferimento"), ("515", "Diferimento com redução de alíquota"),
    ("550", "Suspensão"), ("620", "Tributação monofásica"), ("800", "Transferência de crédito"),
    ("810", "Ajuste de IBS na ZFM"), ("811", "Ajustes"), ("820", "Tributação em declaração de regime específico"),
]

CCT = [
    ("000001", "Situações tributadas integralmente pelo IBS e pela CBS", 0, 0, "S"),
    ("200028", "Serviços de educação (Anexo II) – redução de 60%", 0.6, 0.6, "N"),
    ("200029", "Serviços de saúde humana (Anexo III) – redução de 60%", 0.6, 0.6, "N"),
    ("200030", "Dispositivos médicos (Anexo IV) – redução de 60%", 0.6, 0.6, "N"),
    ("200031", "Dispositivos de acessibilidade (Anexo V) – redução de 60%", 0.6, 0.6, "N"),
    ("200033", "Composições para nutrição enteral e parenteral (Anexo VI) – redução de 60%", 0.6, 0.6, "N"),
    ("200036", "Produtos agropecuários, aquícolas, pesqueiros, florestais e extrativistas in natura – redução de 60%", 0.6, 0.6, "N"),
    ("200039", "Produções artísticas, culturais, jornalísticas e audiovisuais (Anexo X) – redução de 60%", 0.6, 0.6, "N"),
    ("200052", "Profissões intelectuais regulamentadas – redução de 30%", 0.3, 0.3, "N"),
    ("400001", "Transporte público coletivo de passageiros urbano e metropolitano – isenção", 1, 1, "N"),
    ("200003", "Cesta Básica Nacional de Alimentos (Anexo I) – alíquota zero", 1, 1, "N"),
    ("200013", "Produtos de cuidados básicos à saúde menstrual – alíquota zero", 1, 1, "N"),
    ("200014", "Produtos hortícolas, frutas e ovos (Anexo XV) – alíquota zero", 1, 1, "N"),
    ("200032", "Medicamentos registrados na Anvisa ou manipulados – redução de 60%", 0.6, 0.6, "N"),
    ("200034", "Alimentos destinados ao consumo humano (Anexo VII) – redução de 60%", 0.6, 0.6, "N"),
    ("200035", "Produtos de higiene pessoal e limpeza (Anexo VIII) – redução de 60%", 0.6, 0.6, "N"),
    ("200038", "Insumos agropecuários e aquícolas (Anexo IX) – redução de 60%", 0.6, 0.6, "N"),
    ("620001", "Tributação monofásica sobre combustíveis", 0, 0, "N"),
]

SINONIMOS = [
    ("CODIGO", "codigo;cod;cod.;codigo produto;cod produto;cod. produto;cprod;sku;referencia;ref;cod item;codigo do item;codigo interno;item"),
    ("DESCRICAO", "descricao;descricao do produto;descricao produto;produto;xprod;nome;nome do produto;mercadoria;descricao item"),
    ("NCM", "ncm;ncm/sh;cod ncm;codigo ncm;classificacao fiscal;class fiscal;ncm sh"),
    ("CEST", "cest;codigo cest"),
    ("EAN", "ean;gtin;cean;codigo de barras;cod barras;ean13;cod. barras"),
    ("UNIDADE", "unidade;un;und;ucom;unid;unidade comercial;um"),
    ("CFOP", "cfop"),
    ("NBS", "nbs;codigo nbs;cod nbs"),
    ("TIPO", "tipo;tipo item;tipo de item;bem/servico;bem ou servico;natureza"),
    ("CST_CLIENTE", "cst ibs/cbs;cst ibs cbs;cst_ibs_cbs;cst ibscbs;cst ibs"),
    ("CCLASSTRIB_CLIENTE", "cclasstrib;cclass trib;classtrib;class trib;classificacao tributaria"),
]


def sheet(wb, nome, cols, linhas, larg, first=False):
    ws = wb.active if first else wb.create_sheet(nome)
    ws.title = nome
    for c, h in enumerate(cols, 1):
        x = ws.cell(row=1, column=c, value=h); x.font = H; x.fill = FH; x.border = BD
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r, ln in enumerate(linhas, 2):
        for c, v in enumerate(ln, 1):
            x = ws.cell(row=r, column=c, value=v); x.font = B; x.border = BD; x.alignment = WR
    for i, w in enumerate(larg, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def main():
    wb = Workbook()
    ws = sheet(wb, "Regras_NCM", COLS_REGRAS, REGRAS,
               [7, 28, 18, 30, 26, 34, 7, 14, 6, 11, 9, 9, 30, 40, 10, 11, 40, 14, 10], first=True)
    for r in range(2, len(REGRAS) + 2):
        for c in (2, 3, 9, 10, 18):
            ws.cell(row=r, column=c).number_format = "@"
        for c in (11, 12):
            ws.cell(row=r, column=c).number_format = "0%"
    dv = DataValidation(type="list", formula1='"S,N"', allow_blank=False)
    ws.add_data_validation(dv); dv.add("P2:P1000")
    red = PatternFill("solid", fgColor="FFC7CE"); grn = PatternFill("solid", fgColor="C6EFCE")
    ws.conditional_formatting.add("P2:P1000", CellIsRule(operator="equal", formula=['"N"'], fill=red))
    ws.conditional_formatting.add("P2:P1000", CellIsRule(operator="equal", formula=['"S"'], fill=grn))

    wi = sheet(wb, "Regras_IS", ["ID", "NCM", "EXCETO_NCM", "CATEGORIA_IS", "OBSERVACAO", "CONFERIDO"],
               [list(x) + ["N"] for x in IS], [7, 30, 18, 30, 60, 11])
    for r in range(2, len(IS) + 2):
        wi.cell(row=r, column=2).number_format = "@"

    wt = sheet(wb, "Tabela_cClassTrib", ["CCLASSTRIB", "CST", "DESCRICAO", "PRED_IBS", "PRED_CBS", "CONFERIDO"],
               [(c, c[:3], d, a, b, k) for c, d, a, b, k in CCT], [12, 7, 70, 10, 10, 11])
    for r in range(2, len(CCT) + 2):
        for c in (1, 2): wt.cell(row=r, column=c).number_format = "@"
        for c in (4, 5): wt.cell(row=r, column=c).number_format = "0%"
    wc = sheet(wb, "Tabela_CST", ["CST", "DESCRICAO"], CST, [7, 60])
    for r in range(2, len(CST) + 2):
        wc.cell(row=r, column=1).number_format = "@"
    sheet(wb, "Sinonimos_Colunas", ["CAMPO", "CABECALHOS_ACEITOS (separados por ;)"], SINONIMOS, [22, 110])

    cfg = sheet(wb, "Config", ["PARAMETRO", "VALOR", "DESCRICAO"], [
        ("ESCRITORIO", "", "Nome do escritório exibido no relatório ao cliente"),
        ("RESPONSAVEL", "", "Contador responsável (nome / CRC)"),
        ("VERSAO_TABELA_CCLASSTRIB", "Base inicial – conferir com IT 2025.002 vigente", "Versão da tabela cClassTrib usada"),
        ("REGRA_PADRAO_CST", "000", "CST aplicado quando nenhuma regra se aplica"),
        ("REGRA_PADRAO_CCLASSTRIB", "000001", "cClassTrib aplicado quando nenhuma regra se aplica"),
    ], [28, 50, 60])
    for r in (5, 6):
        cfg.cell(row=r, column=2).number_format = "@"

    leia = wb.create_sheet("LEIA-ME", 0)
    textos = [
        "BASE DE REGRAS — CLASSIFICAÇÃO IBS/CBS/IS",
        "",
        "Esta base é usada pelo classificar.py para enquadrar os produtos dos clientes. Edite aqui; não é preciso alterar o código.",
        "Regras de SERVIÇO (TIPO_ITEM = Serviço, NCM vazio): aplicadas a itens sem NCM, por NBS (coluna NBS, prefixos) ou pelas PALAVRAS_CHAVE da descrição.",
        "Regras_NCM: uma linha por regra. NCM e EXCETO_NCM aceitam prefixos (2 a 8 dígitos) separados por ';'. Vale a regra de prefixo mais longo; empate → menor PRIORIDADE.",
        "PALAVRAS_CHAVE: a descrição do produto deve conter ao menos uma (sem acentos). PALAVRAS_EXCLUIR: se contiver, a regra NÃO é aplicada e o item vai para REVISAR com CST 000 (abordagem conservadora).",
        "CONFERIDO = S: regra validada pelo contador com a redação consolidada da LC 214/2025 → itens saem com status OK. CONFERIDO = N → itens saem como A CONFIRMAR.",
        "Tabela_cClassTrib: cole aqui a tabela oficial (IT 2025.002 vigente), mantendo as colunas. Código usado por regra e ausente da tabela → A CONFIRMAR.",
        "Regras_IS: NCMs sujeitas ao Imposto Seletivo (art. 409 e Anexo XVII da LC 214/2025). CST-IS/cClassTribIS dependem de tabela própria.",
        "Sinonimos_Colunas: nomes de cabeçalho reconhecidos nos arquivos dos clientes. Acrescente variações conforme aparecerem.",
        "",
        "IMPORTANTE: a base inicial foi elaborada sem acesso ao texto oficial; todas as regras estão com CONFERIDO = N até a validação do contador.",
    ]
    for i, t in enumerate(textos, 1):
        c = leia.cell(row=i, column=1, value=t); c.font = BB if i == 1 else B; c.alignment = WR
    leia.column_dimensions["A"].width = 140

    wb.save(DESTINO)
    print("Base gerada:", DESTINO)


if __name__ == "__main__":
    if os.path.exists(DESTINO) and "--sobrescrever" not in sys.argv:
        sys.exit(f"{DESTINO} já existe. Use --sobrescrever para recriar (as edições serão perdidas).")
    main()
