"""Gera a planilha de teste com casos conhecidos de precedentes e dependentes.

Uso: python tools/gerar_planilha_teste.py [saida.xlsx]
Sem saída indicada, grava docs/planilha-teste.xlsx.
As fórmulas ficam em inglês com vírgula, que é como o arquivo guarda; o Excel em
português exibe SOMA, DESLOC e ponto e vírgula.
"""
import sys
from pathlib import Path

from openpyxl import Workbook

RAIZ = Path(__file__).resolve().parent.parent


def main():
    saida = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "docs" / "planilha-teste.xlsx"
    wb = Workbook()

    premissas = wb.active
    premissas.title = "Premissas"
    premissas["B2"], premissas["C2"] = "Crescimento", 0.10
    premissas["B3"], premissas["C3"] = "Margem EBITDA", 0.25
    for ref in ("C2", "C3"):
        premissas[ref].number_format = "0%"

    modelo = wb.create_sheet("Modelo")
    linhas = [
        ("Receita base", 1000),
        ("Receita ano 1", "=B1*(1+Premissas!C2)"),
        ("Receita ano 2", "=B2*(1+Premissas!C2)"),
        ("EBITDA ano 2", "=B3*Premissas!C3"),
    ]
    for i, (rotulo, valor) in enumerate(linhas, start=1):
        modelo.cell(row=i, column=1, value=rotulo)
        modelo.cell(row=i, column=2, value=valor).number_format = "#,##0.00"

    # Caso do print do Arixcel: =SOMA(O111:DESLOC(O116;-1;;;))
    modelo["N110"] = "Caso do print"
    modelo["O110"] = "=SUM(O111:OFFSET(O116,-1,,,))"
    valores = [-1031327.53, -500000.00, -430505.55, -300000.00, -200000.00, 199245096.56]
    for linha, valor in enumerate(valores, start=111):
        modelo.cell(row=linha, column=15, value=valor)
    for linha in range(110, 117):
        modelo.cell(row=linha, column=15).number_format = "#,##0.00"

    # Abas com espaço e com vírgula no nome testam a leitura dos endereços.
    espaco = wb.create_sheet("Aba com espaço")
    espaco["A1"], espaco["B1"] = "Dobro do EBITDA", "=Modelo!B4*2"

    virgula = wb.create_sheet("Resumo, v2")
    virgula["A1"], virgula["B1"] = "Receita ano 2 + EBITDA", "=Modelo!B3+Modelo!B4"

    for aba in wb.worksheets:
        aba.column_dimensions["A"].width = 26
        aba.column_dimensions["B"].width = 16
    modelo.column_dimensions["N"].width = 16
    modelo.column_dimensions["O"].width = 18

    wb.save(saida)
    print(f"{saida} gerada")


if __name__ == "__main__":
    main()
