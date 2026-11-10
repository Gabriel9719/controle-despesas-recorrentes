"""Implementação de referência em Python da mesma regra das fórmulas da planilha.

Serve para (1) rodar o controle fora do Excel e (2) testar se as fórmulas estão corretas.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd


def calcular_status(dados: Path, competencia: str) -> pd.DataFrame:
    rec = pd.read_csv(dados / "recorrentes.csv", sep=";", dtype=str)
    exc = pd.read_csv(dados / "exclusoes.csv", sep=";", dtype=str).fillna("")
    lan = pd.read_csv(dados / "lancamentos.csv", sep=";", dtype=str)
    lan = lan[lan["competencia"] == competencia]

    def excluido(r) -> bool:
        return any((e.fornecedor in ("", r.fornecedor)) and (e.pa in ("", r.pa)) and (e.fornecedor or e.pa)
                   for e in exc.itertuples())

    recebidos = set(zip(lan["fornecedor"], lan["pa"]))
    rec["status"] = [
        "EXCLUÍDO" if excluido(r) else "RECEBIDO" if (r.fornecedor, r.pa) in recebidos else "PENDENTE"
        for r in rec.itertuples()
    ]
    rec["destinatario"] = rec.apply(lambda r: f"PA {r.pa}" if r.responsavel == "PA" else r.responsavel, axis=1)
    return rec


if __name__ == "__main__":
    comp = sys.argv[1] if len(sys.argv) > 1 else "2026-10"
    st = calcular_status(Path("data"), comp)
    print(st["status"].value_counts().to_string())
    print("\nPendências por destinatário:")
    print(st[st["status"] == "PENDENTE"].groupby("destinatario").size().to_string())
