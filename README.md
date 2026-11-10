# Controle de despesas recorrentes (Excel + VBA)

![excel](https://img.shields.io/badge/Excel-f%C3%B3rmulas-217346) ![vba](https://img.shields.io/badge/VBA-Outlook-5C2D91) ![python](https://img.shields.io/badge/python-valida%C3%A7%C3%A3o-3776AB)

Planilha que responde, todo mês, a uma pergunta do contas a pagar: **quais fornecedores recorrentes ainda não enviaram a fatura desta competência, e quem precisa ser cobrado?** Ela cruza o cadastro de despesas recorrentes por PA com os lançamentos do mês e direciona cada pendência ao responsável. Depois, abre no Outlook um e-mail de cobrança por destinatário.

> **Dados fictícios.** Fornecedores, PAs, e-mails (`@coop-exemplo.com.br`) e valores são gerados por script.

## Regras de negócio

| Regra | Como foi implementada |
|---|---|
| Água e energia são cobradas **do próprio PA**; as demais despesas, do **departamento responsável** | Coluna *Destinatário* = `"PA " & PA` ou o departamento; e-mail via `INDEX/MATCH` na aba Destinatarios |
| Exclusões por **fornecedor**, por **PA** ou pela **combinação** dos dois (débito automático, PA em reforma, contrato encerrado) | Três `COUNTIFS` com critério vazio = "todos" |
| Aluguel só entra quando é pago por **boleto** | Cadastrado apenas nos PAs com boleto |
| Status do mês | `EXCLUÍDO` → `RECEBIDO` (há lançamento na competência) → `PENDENTE` |
| Envio seguro | Modo **EXIBIR** por padrão: o e-mail abre para revisão. **ENVIAR** pede confirmação antes. |

## Abas

| Aba | Conteúdo |
|---|---|
| **Painel** | Recebidas, pendentes, excluídas, % recebido e pendências por destinatário (com destaque) |
| **Recorrentes** | Cadastro + colunas calculadas (excluído?, lançamentos, status, destinatário, e-mail) |
| **Lancamentos** | Base do mês importada do sistema contábil |
| **Exclusoes** | Fornecedor e/ou PA fora do controle, com o motivo |
| **Destinatarios** | PA/departamento → e-mail |
| **Parametros** | Competência, modo de envio (lista suspensa) e assinatura |

## Macros (`vba/`)

- **`mdlImportacao.bas`**: importa o CSV de lançamentos (UTF-8, `;`) e **ignora IDs já importados**, o que evita duplicar valores se o arquivo for carregado duas vezes.
- **`mdlAlertas.bas`**: agrupa as pendências por destinatário e monta **um e-mail HTML por destinatário** no Outlook. Usa late binding, então não precisa ativar referência; valida e-mails ausentes antes de começar.

Os módulos ficam como `.bas` no repositório para que o código VBA tenha **controle de versão** (diff legível), em vez de ficar escondido dentro de um `.xlsm`.

## Validação das fórmulas

As fórmulas foram validadas contra uma **implementação de referência em Python** (`scripts/verificar_pendencias.py`). A planilha foi recalculada no LibreOffice e o resultado bateu linha a linha com o Python (32 recebidas, 9 pendentes e 5 excluídas no exemplo). O teste automatizado cobre as regras de exclusão e de roteamento.

## Como usar

```bash
pip install -r requirements.txt
python scripts/construir_planilha.py --competencia 2026-10   # gera a planilha e data/*.csv
python scripts/verificar_pendencias.py 2026-10                # mesma regra, fora do Excel
pytest -q
```

Para usar as macros: abra `Controle_Despesas_Recorrentes.xlsx` → `Alt+F11` → *Arquivo › Importar arquivo* (os dois `.bas`) → salve como **.xlsm**.

## Stack

Excel (COUNTIFS, INDEX/MATCH, validação de dados, formatação condicional) · VBA (Outlook late binding, Scripting.Dictionary) · Python (openpyxl, pandas) · pytest

---
Autor: **Gabriel Antônio Ribeiro**, Analista Financeiro · [LinkedIn](https://www.linkedin.com/in/gabriel-ant%C3%B4nio-ribeiro-017808210)
