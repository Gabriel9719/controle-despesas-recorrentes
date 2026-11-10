"""Garante que a regra em Python (referência) produz o resultado esperado e que a
planilha gerada contém as fórmulas nas colunas de status."""
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "scripts"))
from verificar_pendencias import calcular_status  # noqa: E402


def test_planilha_e_regras(tmp_path):
    xlsx, dados = tmp_path / "c.xlsx", tmp_path / "dados"
    subprocess.run([sys.executable, str(RAIZ / "scripts/construir_planilha.py"), "--saida", str(xlsx),
                    "--dados", str(dados)], check=True)
    st = calcular_status(dados, "2026-10")
    cont = st["status"].value_counts()
    assert cont["PENDENTE"] == 9 and cont["EXCLUÍDO"] == 5
    # exclusão por PA inteiro e por fornecedor em todos os PAs
    assert (st.loc[st["pa"] == "12", "status"] == "EXCLUÍDO").all()
    assert (st.loc[st["fornecedor"] == "CRIATIVA COMUNICAÇÃO LTDA", "status"] == "EXCLUÍDO").all()
    # água/energia vão para o PA; demais para o departamento
    agua = st[st["tipo"] == "Água/Saneamento"]
    assert agua["destinatario"].str.startswith("PA ").all()

    ws = load_workbook(xlsx)["Recorrentes"]
    assert str(ws["G2"].value).startswith("=IF(")
    assert "COUNTIFS" in str(ws["F2"].value)
