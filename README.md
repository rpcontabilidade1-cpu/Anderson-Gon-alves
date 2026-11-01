# Sistema de Precificação Tributária — Antes x Reforma (2027–2033)

Planilha (`Precificacao_Tributaria_Reforma_2027_2033.xlsx`) para simular a formação de preço de venda comparando:

- **Regime atual**: Simples Nacional, Lucro Presumido e Lucro Real (PIS, COFINS, ICMS, ISS, IPI).
- **Reforma Tributária (EC 132/2023 e LC 214/2025)**: CBS, IBS e Imposto Seletivo, ano a ano, durante a transição de 2026 a 2033.

## Abas

1. **Capa** — instruções de uso e legenda de cores.
2. **Parâmetros** — alíquotas do regime atual, estimativas de CBS/IBS/IS e cronograma oficial de transição (editável).
3. **Dados de Entrada** — dados do produto/serviço, regime do cliente, custos, despesas e margem desejada.
4. **Cálculo Regime Atual** — carga tributária e preço de venda antes da Reforma.
5. **Cálculo Reforma 2027-2033** — carga tributária efetiva e preço de venda ano a ano durante a transição.
6. **Comparativo Consolidado** — tabela e gráficos da evolução de 2026 a 2033.
7. **Simples - Único x Híbrido** — decisão do cliente do Simples Nacional entre o Regime Único (IBS/CBS por dentro do DAS) e o Regime Híbrido (por fora, crédito pleno ao comprador B2B), conforme LC 214/2025, art. 41, §3º.
8. **Memória de Cálculo** — metodologia, fórmulas, base legal e premissas.

## Como usar

Preencha apenas as células **azuis** (nas abas Parâmetros e Dados de Entrada). Todo o restante é calculado automaticamente por fórmula.

> As alíquotas de CBS/IBS são estimativas técnicas até a definição da alíquota de referência pelo Senado Federal — atualize a aba "Parâmetros" quando os valores oficiais forem publicados.

---

# Classificação de Produtos — Reforma Tributária (IBS / CBS / IS)

Planilha (`Classificacao_Produtos_Reforma_Tributaria_IBS_CBS_IS.xlsx`) de **importação e validação** da classificação tributária de itens conforme a EC 132/2023 e a LC 214/2025 (com alterações posteriores).

## Abas

1. **Instrucoes** — finalidade, base normativa, passo a passo, legenda, regras de validação e riscos.
2. **Importacao** — layout de importação (1.000 linhas): colunas A–M de entrada (código, descrição, tipo, NCM, NBS, CEST, unidade, GTIN, CST IBS/CBS, cClassTrib, IS) e colunas N–V calculadas (descrição do cClassTrib, fundamento legal, % de redução, alíquotas efetivas, inconsistências e status OK/PENDENTE).
3. **Exportacao** — linhas prontas para importação (somente itens com Status OK) em três formatos: CSV (`;`), TXT delimitado (`|campo|`) e TXT posicional (207 posições). Copiar a faixa indicada, colar no Bloco de Notas e salvar como .csv/.txt.
4. **Layout** — especificação técnica de cada campo e leiaute dos arquivos CSV/TXT (posições inicial/final no TXT posicional e códigos das categorias do IS).
5. **Resumo** — totais de itens, pendências, itens sujeitos ao IS e distribuição por CST.
6. **Parametros** — ano de referência e alíquotas estimadas de CBS/IBS na transição 2026–2033.
7. **Tabela_CST** / **Tabela_cClassTrib** — tabelas de referência (substituíveis pela versão oficial vigente).
8. **Anexos_LC214** — Anexos I a XV e XVII com redução aplicável e cClassTrib usual.
9. **Listas** — listas suspensas (tipo de item e categorias do Imposto Seletivo).

## Exportação automática (CSV e TXT)

```bash
pip install openpyxl
python exportar_classificacao.py Classificacao_Produtos_Reforma_Tributaria_IBS_CBS_IS.xlsx
```

Gera na pasta `exportacao/`:

| Arquivo | Formato |
|---|---|
| `classificacao_ibs_cbs.csv` | `;`, cabeçalho, decimal `,`, UTF-8, CRLF |
| `classificacao_ibs_cbs_delimitado.txt` | `\|campo\|campo\|` (padrão SPED), sem cabeçalho, UTF-8 |
| `classificacao_ibs_cbs_posicional.txt` | largura fixa de 207 posições, ANSI (Windows-1252) |
| `pendencias.csv` | itens não exportados e o motivo |

Opções: `--saida`, `--decimal .`, `--enc-csv utf-8-sig`, `--enc-txt cp1252`, `--enc-pos utf-8`, `--sem-cabecalho`, `--sem-acentos`, `--lf`. O script aplica as mesmas regras de validação da planilha e gera saída idêntica à da aba Exportacao.

> Os códigos cClassTrib marcados "CONFERIR" devem ser confirmados na tabela oficial vigente (Portal NF-e / Conformidade Fácil) antes do uso em produção.

---

# Sistema de Classificação de Produtos (pasta `sistema_classificacao/`)

Recebe o cadastro do cliente (Excel, CSV ou XML de NF-e), enquadra automaticamente cada produto (CST IBS/CBS, cClassTrib, reduções, IS) com base numa base de regras editável (`base_regras.xlsx`) e devolve um Excel padronizado com aba de importação para os programadores do ERP. Ver `sistema_classificacao/LEIA-ME.md`.
