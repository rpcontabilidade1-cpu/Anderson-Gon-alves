# Sistema de Classificação de Produtos — IBS / CBS / IS

Recebe o cadastro de produtos do cliente, faz o enquadramento (CST IBS/CBS, cClassTrib, reduções, Imposto Seletivo) e devolve um **Excel padronizado** para o cliente entregar aos programadores do ERP.

## Fluxo de trabalho

```
Cliente envia Excel/CSV/XML  →  classificar.py + base_regras.xlsx  →  revisão do contador  →  Excel devolvido ao cliente
```

1. **Receber** o arquivo do cliente: Excel (.xlsx/.xlsm), CSV, XML de NF-e ou pasta com XMLs. O cabeçalho é localizado automaticamente (aceita títulos acima, nomes variados de coluna, NCM com pontos).
2. **Classificar**: arraste o arquivo sobre `CLASSIFICAR.bat` (Windows) ou execute:
   ```
   python classificar.py arquivo_do_cliente.xlsx --cliente "Nome do Cliente" --regime "Lucro Real"
   ```
   O resultado é gravado na pasta `saida/`, ao lado do arquivo do cliente.
3. **Revisar** no Excel gerado: filtre a coluna STATUS na aba `Cadastro_Fiscal`, corrija CST/cClassTrib quando necessário (as descrições, reduções e a aba `Importacao_ERP` se atualizam sozinhas) e altere o STATUS para OK.
4. **Devolver** o Excel ao cliente. Os programadores usam as abas `Importacao_ERP` e `Leiaute_Importacao`.

## Excel devolvido ao cliente

| Aba | Para quem | Conteúdo |
|---|---|---|
| Instrucoes | Cliente | Identificação, base normativa, legenda, premissas e limitações |
| Cadastro_Fiscal | Escritório/cliente | Classificação completa item a item, com base legal, condição do benefício, status e observações |
| Importacao_ERP | Programadores | Somente os campos de importação, em formato técnico |
| Resumo | Cliente | Totais por status, tratamento e NCM |
| Divergencias_Riscos | Escritório/cliente | Pendências, riscos e ações recomendadas |
| Leiaute_Importacao | Programadores | Especificação dos campos e orientações de importação |
| Tabela_Referencia | Todos | Tabelas de CST e cClassTrib utilizadas |

## Status

| Status | Significado |
|---|---|
| OK | Regra e código conferidos: pronto para importação |
| A CONFIRMAR | Enquadrado por regra ou código ainda não conferido na base |
| REVISAR | NCM consta de anexo com benefício, mas a descrição não confirma: aplicada tributação integral (conservador) |
| REVISAR NCM | NCM ausente ou inválido |

## Base de regras (`base_regras.xlsx`)

O conhecimento técnico fica em uma planilha que você edita sem mexer no código:

- **Regras_NCM**: prefixos de NCM, exceções, palavras-chave obrigatórias/excludentes, CST, cClassTrib, reduções, base legal, condição e `CONFERIDO`.
- **Regras_IS**: NCMs sujeitas ao Imposto Seletivo.
- **Tabela_cClassTrib** e **Tabela_CST**: cole aqui a tabela oficial vigente (IT 2025.002).
- **Sinonimos_Colunas**: nomes de cabeçalho reconhecidos nos arquivos dos clientes.
- **Config**: escritório, responsável, versão da tabela, classificação padrão.

> **Importante:** a base inicial foi montada sem acesso ao texto oficial da LC 214/2025 e à tabela cClassTrib vigente. Todas as regras estão com `CONFERIDO = N`, por isso os itens saem como **A CONFIRMAR**. Depois de validar cada regra com a redação consolidada da lei e com a tabela oficial, mude para `S`: a partir daí os itens enquadrados saem como **OK** para todos os clientes.

## Requisitos

- Python 3.8+ e `openpyxl` (o `CLASSIFICAR.bat` instala automaticamente).
- Arquivos `.xls` antigos: salve como `.xlsx` antes.

## Exemplo

`exemplos/cliente_exemplo.xlsx` → `exemplos/saida/Mercadinho_Exemplo_Classificacao_IBS_CBS_*.xlsx`
