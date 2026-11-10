"""Gera dados FICTÍCIOS e monta a planilha de controle de despesas recorrentes.

A lógica fica toda em fórmulas (COUNTIFS/INDEX/MATCH) para funcionar em qualquer Excel;
as macros (pasta vba/) só automatizam importação e envio de alertas.

Uso:  python scripts/construir_planilha.py --competencia 2026-10 --seed 7
Saída: Controle_Despesas_Recorrentes.xlsx  (importe os módulos de vba/ e salve como .xlsm)
"""
from __future__ import annotations

import argparse
import csv
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

AZUL, VERDE = "003641", "00AE9D"
AUTOR = "Gabriel Antônio Ribeiro"
PAS = [f"{i:02d}" for i in range(1, 13)]
DEPARTAMENTOS = {
    "Suporte Operacional": "suporte@coop-exemplo.com.br",
    "Tesouraria": "tesouraria@coop-exemplo.com.br",
    "Jurídico": "juridico@coop-exemplo.com.br",
    "Gente e Cultura": "gente@coop-exemplo.com.br",
    "Marketing": "marketing@coop-exemplo.com.br",
    "Seguros": "seguros@coop-exemplo.com.br",
}
POR_PA = [("Água/Saneamento", "SANEAMENTO ESTADUAL S.A.", "PA"),
          ("Energia elétrica", "ENERGIA CENTRO-OESTE S.A.", "PA"),
          ("Internet/Link", "FIBRA NET TELECOM LTDA", "Suporte Operacional")]
CORPORATIVAS = [("Vigilância", "PROTEGE VIGILÂNCIA LTDA", "Suporte Operacional"),
                ("Limpeza e conservação", "LIMPA BEM SERVIÇOS LTDA", "Suporte Operacional"),
                ("Honorários advocatícios", "SILVA & SOUZA ADVOGADOS", "Jurídico"),
                ("Plano odontológico", "ODONTO MAIS S.A.", "Gente e Cultura"),
                ("Agência de publicidade", "CRIATIVA COMUNICAÇÃO LTDA", "Marketing"),
                ("Seguro patrimonial", "SEGURADORA HORIZONTE S.A.", "Seguros"),
                ("Transporte de valores", "TRANSVALOR LTDA", "Tesouraria")]
ALUGUEL_BOLETO = {"03": "IMOBILIÁRIA CENTRAL LTDA", "08": "JOSÉ PEREIRA (LOCADOR PF)", "11": "PATRIMONIAL INVEST LTDA"}


def gerar_dados(rng: random.Random, comp: str):
    rec = []
    for pa in PAS:
        for tipo, forn, resp in POR_PA:
            rec.append((forn, pa, tipo, resp))
        if pa in ALUGUEL_BOLETO:
            rec.append((ALUGUEL_BOLETO[pa], pa, "Aluguel (boleto)", "Suporte Operacional"))
    for tipo, forn, resp in CORPORATIVAS:
        rec.append((forn, "00", tipo, resp))
    exclusoes = [("ENERGIA CENTRO-OESTE S.A.", "06", "PA 06 em débito automático"),
                 ("", "12", "PA 12 em reforma - fora do controle até 12/2026"),
                 ("CRIATIVA COMUNICAÇÃO LTDA", "", "Contrato encerrado")]
    faltantes = set(rng.sample(range(len(rec)), 9))
    lanc, n = [], 50_000
    for i, (forn, pa, tipo, _) in enumerate(rec):
        if i in faltantes:
            continue
        n += 1
        lanc.append((n, comp, forn, pa, tipo, round(rng.uniform(120, 9800), 2)))
    return rec, exclusoes, lanc


def cabecalho(ws, cols, larguras):
    ws.append(cols)
    for c, w in zip(ws[1], larguras):
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"


def construir(comp: str, seed: int, saida: Path, dados_dir: Path):
    rng = random.Random(seed)
    rec, exc, lanc = gerar_dados(rng, comp)
    wb = Workbook()
    wb.properties.creator = AUTOR; wb.properties.lastModifiedBy = AUTOR
    wb.properties.title = "Controle de despesas recorrentes (dados fictícios)"

    # Parâmetros
    ws = wb.active; ws.title = "Parametros"
    ws["A1"] = "Parâmetros"; ws["A1"].font = Font(bold=True, size=14, color=AZUL)
    for r, (k, v) in enumerate([("Competência (AAAA-MM)", comp), ("Modo de envio", "EXIBIR"),
                                ("Assinatura", "Financeiro - Contas a Pagar")], start=3):
        ws.cell(r, 1, k).font = Font(bold=True); ws.cell(r, 2, v)
    dv = DataValidation(type="list", formula1='"EXIBIR,ENVIAR"', allow_blank=False); ws.add_data_validation(dv); dv.add("B4")
    ws["A7"] = "EXIBIR abre cada e-mail no Outlook para revisão; ENVIAR dispara direto (use com cuidado)."
    ws.column_dimensions["A"].width = 26; ws.column_dimensions["B"].width = 34

    # Destinatários
    wd = wb.create_sheet("Destinatarios"); cabecalho(wd, ["Chave", "E-mail"], [24, 36])
    for pa in PAS:
        wd.append([f"PA {pa}", f"pa{pa}@coop-exemplo.com.br"])
    for dep, mail in DEPARTAMENTOS.items():
        wd.append([dep, mail])

    # Cadastro de recorrentes + status calculado
    wr = wb.create_sheet("Recorrentes", 1)
    cabecalho(wr, ["Fornecedor", "PA", "Tipo de despesa", "Responsável", "Excluído?", "Lançamentos na competência",
                   "Status", "Destinatário do alerta", "E-mail"], [32, 6, 24, 22, 10, 14, 14, 22, 32])
    for i, (forn, pa, tipo, resp) in enumerate(rec, start=2):
        wr.append([forn, pa, tipo, resp])
        wr[f"E{i}"] = (f'=IF(COUNTIFS(Exclusoes!$A:$A,A{i},Exclusoes!$B:$B,"")'
                       f'+COUNTIFS(Exclusoes!$A:$A,"",Exclusoes!$B:$B,B{i})'
                       f'+COUNTIFS(Exclusoes!$A:$A,A{i},Exclusoes!$B:$B,B{i})>0,"SIM","NÃO")')
        wr[f"F{i}"] = f'=COUNTIFS(Lancamentos!$C:$C,A{i},Lancamentos!$D:$D,B{i},Lancamentos!$B:$B,Parametros!$B$3)'
        wr[f"G{i}"] = f'=IF(E{i}="SIM","EXCLUÍDO",IF(F{i}>0,"RECEBIDO","PENDENTE"))'
        wr[f"H{i}"] = f'=IF(D{i}="PA","PA "&B{i},D{i})'
        wr[f"I{i}"] = f'=IFERROR(INDEX(Destinatarios!$B:$B,MATCH(H{i},Destinatarios!$A:$A,0)),"sem e-mail cadastrado")'
    n = len(rec) + 1
    wr.conditional_formatting.add(f"G2:G{n}", CellIsRule(operator="equal", formula=['"PENDENTE"'], fill=PatternFill("solid", fgColor="F8D7DA")))
    wr.conditional_formatting.add(f"G2:G{n}", CellIsRule(operator="equal", formula=['"RECEBIDO"'], fill=PatternFill("solid", fgColor="D9F2E6")))
    wr.conditional_formatting.add(f"G2:G{n}", CellIsRule(operator="equal", formula=['"EXCLUÍDO"'], fill=PatternFill("solid", fgColor="E7E6E6")))
    wr.auto_filter.ref = f"A1:I{n}"
    wr.conditional_formatting.add(f"I2:I{n}", FormulaRule(formula=[f'I2="sem e-mail cadastrado"'], font=Font(color="C00000", bold=True)))

    # Lançamentos (base importada do razão / pré-registros)
    wl = wb.create_sheet("Lancamentos"); cabecalho(wl, ["ID lançamento", "Competência", "Fornecedor", "PA", "Tipo", "Valor"], [14, 12, 32, 6, 24, 14])
    for row in lanc:
        wl.append(list(row))
    for c in wl["F"][1:]:
        c.number_format = '#,##0.00'

    # Exclusões
    we = wb.create_sheet("Exclusoes"); cabecalho(we, ["Fornecedor (vazio = todos)", "PA (vazio = todos)", "Motivo"], [32, 18, 46])
    for row in exc:
        we.append([v if v else None for v in row])

    # Painel
    wp = wb.create_sheet("Painel", 0)
    wp["A1"] = "Controle de despesas recorrentes"; wp["A1"].font = Font(bold=True, size=16, color=AZUL)
    wp["A2"] = '="Competência: "&Parametros!B3'; wp["A2"].font = Font(color="7F8C8D")
    kpis = [("Recorrentes cadastradas", f"=COUNTA(Recorrentes!A2:A{n})"),
            ("Recebidas", f'=COUNTIF(Recorrentes!G2:G{n},"RECEBIDO")'),
            ("Pendentes", f'=COUNTIF(Recorrentes!G2:G{n},"PENDENTE")'),
            ("Excluídas do controle", f'=COUNTIF(Recorrentes!G2:G{n},"EXCLUÍDO")'),
            ("% recebido (sem excluídas)", "=IFERROR(B5/(B4-B7),0)")]
    for r, (k, f) in enumerate(kpis, start=4):
        wp.cell(r, 1, k).font = Font(bold=True); c = wp.cell(r, 2, f); c.font = Font(bold=True, size=13, color=VERDE)
    wp["B8"].number_format = "0.0%"
    wp["A10"] = "Pendências por destinatário"; wp["A10"].font = Font(bold=True, color=AZUL)
    destinos = [f"PA {p}" for p in PAS] + list(DEPARTAMENTOS)
    for r, d in enumerate(destinos, start=11):
        wp.cell(r, 1, d); wp.cell(r, 2, f'=COUNTIFS(Recorrentes!$H$2:$H${n},A{r},Recorrentes!$G$2:$G${n},"PENDENTE")')
    wp.conditional_formatting.add(f"B11:B{10 + len(destinos)}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="F8D7DA")))
    wp.column_dimensions["A"].width = 30; wp.column_dimensions["B"].width = 14

    wb.calculation.fullCalcOnLoad = True   # Excel recalcula tudo ao abrir
    saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(saida)

    # CSVs para quem quiser reproduzir/validar fora do Excel
    dados_dir.mkdir(parents=True, exist_ok=True)
    for nome, cols, rows in (("recorrentes.csv", ["fornecedor", "pa", "tipo", "responsavel"], rec),
                             ("exclusoes.csv", ["fornecedor", "pa", "motivo"], exc),
                             ("lancamentos.csv", ["id", "competencia", "fornecedor", "pa", "tipo", "valor"], lanc)):
        with (dados_dir / nome).open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";"); w.writerow(cols); w.writerows(rows)
    return rec, exc, lanc


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--competencia", default="2026-10")
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--saida", type=Path, default=Path("Controle_Despesas_Recorrentes.xlsx"))
    ap.add_argument("--dados", type=Path, default=Path("data"))
    a = ap.parse_args()
    rec, _, lanc = construir(a.competencia, a.seed, a.saida, a.dados)
    print(f"{len(rec)} recorrentes, {len(lanc)} lançamentos -> {a.saida}")
