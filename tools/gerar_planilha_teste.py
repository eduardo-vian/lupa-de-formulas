"""Gera a planilha de teste com casos conhecidos de precedentes e dependentes.

Uso: python tools/gerar_planilha_teste.py [saida.xlsx]
Sem saída indicada, grava docs/planilha-teste.xlsx.
As fórmulas ficam em inglês com vírgula, que é como o arquivo guarda; o Excel em
português exibe SOMA, DESLOC e ponto e vírgula.
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

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

    # Uma fórmula de cada recurso da janela (funções novas do Excel levam o prefixo _xlfn. no arquivo).
    casos = wb.create_sheet("Casos")
    for linha, (letra, valor) in enumerate([("a", 10), ("b", 20), ("c", 30)], start=1):
        casos.cell(row=linha, column=1, value=letra)
        casos.cell(row=linha, column=2, value=valor)
    exemplos = [
        ("PROCV exato", '=VLOOKUP("b",A1:B3,2,FALSE)'),
        ("SE: ramo em negrito", '=IF(B1>15,"alto","baixo")'),
        ("SOMASES: células filtradas", '=SUMIFS(B1:B3,A1:A3,"<>b")'),
        ("ÍNDICE", "=INDEX(B1:B3,3)"),
        ("CORRESP", '=MATCH("c",A1:A3,0)'),
        ("MÁXIMO", "=MAX(B1:B3)"),
        ("Nome definido Taxa", "=Taxa*2"),
        ("Validação: lista A1:A3", "a"),
        ("PROCX", '=_xlfn.XLOOKUP("c",A1:A3,B1:B3)'),
        ("SES", '=_xlfn.IFS(A1="z",1,A2="b",2)'),
    ]
    for linha, (rotulo, formula) in enumerate(exemplos, start=1):
        casos.cell(row=linha, column=3, value=rotulo)
        casos.cell(row=linha, column=4, value=formula)
    lista = DataValidation(type="list", formula1="$A$1:$A$3", allow_blank=True)
    casos.add_data_validation(lista)
    lista.add("D8")
    taxa = DefinedName("Taxa", attr_text="Casos!$B$2")
    try:
        wb.defined_names["Taxa"] = taxa  # openpyxl 3.1+
    except TypeError:
        wb.defined_names.append(taxa)  # versões antigas
    casos.column_dimensions["C"].width = 28
    casos.column_dimensions["D"].width = 14

    for aba in wb.worksheets:
        if aba.title == "Casos":
            continue
        aba.column_dimensions["A"].width = 26
        aba.column_dimensions["B"].width = 16
    modelo.column_dimensions["N"].width = 16
    modelo.column_dimensions["O"].width = 18

    wb.save(saida)
    print(f"{saida} gerada")


if __name__ == "__main__":
    main()
