from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

ARQUIVO_SAIDA = "controle_pagamentos.xlsx"

MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def aplicar_cabecalho(ws, linha=1):
    fill = PatternFill("solid", fgColor="1F4E78")
    fonte = Font(color="FFFFFF", bold=True)
    for cell in ws[linha]:
        cell.fill = fill
        cell.font = fonte
        cell.alignment = Alignment(horizontal="center", vertical="center")


def aplicar_bordas(ws, min_row, max_row, min_col, max_col):
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = border


def criar_planilha():
    wb = Workbook()

    ws = wb.active
    ws.title = "Lancamentos"

    colunas = [
        "Fornecedor", "B.U", "Serviço", "Natureza", "C/C", "Dia Venc.", "Tipo",
        "Nº Fatura", "Status", "Observações"
    ] + MESES + ["Total Anual", "Média Mensal", "Alerta Vencimento"]

    ws.append(colunas)
    aplicar_cabecalho(ws)

    # Largura das colunas
    larguras = {
        "A": 28, "B": 14, "C": 30, "D": 28, "E": 18, "F": 12, "G": 12,
        "H": 16, "I": 12, "J": 28,
        "K": 12, "L": 12, "M": 12, "N": 12, "O": 12, "P": 12,
        "Q": 12, "R": 12, "S": 12, "T": 12, "U": 12, "V": 12,
        "W": 14, "X": 14, "Y": 20,
    }
    for c, l in larguras.items():
        ws.column_dimensions[c].width = l

    ws.freeze_panes = "A2"

    # Linhas modelo
    inicio = 2
    fim = 300

    # Data validations
    dv_status = DataValidation(type="list", formula1="=Listas!$A$2:$A$3", allow_blank=True)
    dv_tipo = DataValidation(type="list", formula1="=Listas!$B$2:$B$3", allow_blank=True)
    ws.add_data_validation(dv_status)
    ws.add_data_validation(dv_tipo)

    for r in range(inicio, fim + 1):
        # Total Anual: soma meses K:V
        ws[f"W{r}"] = f"=SUM(K{r}:V{r})"
        # Média mensal (apenas meses com valor)
        ws[f"X{r}"] = f"=IFERROR(SUM(K{r}:V{r})/COUNTIF(K{r}:V{r},\">0\"),0)"
        # Alerta vencimento próximo (7 dias)
        ws[f"Y{r}"] = (
            f"=IF(OR(F{r}=\"\",I{r}=\"Pago\"),\"\","
            f"IF(AND(DAY(TODAY())<={r} , FALSE),\"\","
            f"IF(OR(F{r}-DAY(TODAY())<=7,F{r}+30-DAY(TODAY())<=7),\"Vence em até 7 dias\",\"\")))"
        )
        dv_status.add(ws[f"I{r}"])
        dv_tipo.add(ws[f"G{r}"])

    # Formatação monetária
    for col in "KLMNOPQRSTUVWX":
        for r in range(inicio, fim + 1):
            ws[f"{col}{r}"].number_format = 'R$ #,##0.00'

    # Regras de destaque
    pago_fill = PatternFill("solid", fgColor="C6EFCE")
    pendente_fill = PatternFill("solid", fgColor="FFC7CE")
    alerta_fill = PatternFill("solid", fgColor="FFF2CC")

    ws.conditional_formatting.add(
        f"I{inicio}:I{fim}", FormulaRule(formula=[f'I{inicio}="Pago"'], fill=pago_fill)
    )
    ws.conditional_formatting.add(
        f"I{inicio}:I{fim}", FormulaRule(formula=[f'I{inicio}="Pendente"'], fill=pendente_fill)
    )
    ws.conditional_formatting.add(
        f"Y{inicio}:Y{fim}", FormulaRule(formula=[f'Y{inicio}="Vence em até 7 dias"'], fill=alerta_fill)
    )

    aplicar_bordas(ws, 1, fim, 1, len(colunas))

    # Aba Dashboard
    dash = wb.create_sheet("Dashboard")
    dash.append(["Indicador", "Valor"])
    aplicar_cabecalho(dash)
    dash.column_dimensions["A"].width = 34
    dash.column_dimensions["B"].width = 20

    dash["A2"] = "Total Pago"
    dash["B2"] = '=SUMIFS(Lancamentos!W:W,Lancamentos!I:I,"Pago")'
    dash["A3"] = "Total Pendente"
    dash["B3"] = '=SUMIFS(Lancamentos!W:W,Lancamentos!I:I,"Pendente")'
    dash["A4"] = "Previsão do mês atual"
    dash["B4"] = (
        "=SUM(INDEX(Lancamentos!K:V,0,MONTH(TODAY())))"
    )

    for c in ["B2", "B3", "B4"]:
        dash[c].number_format = 'R$ #,##0.00'
    aplicar_bordas(dash, 1, 6, 1, 2)

    # Aba Importacao
    imp = wb.create_sheet("Importacao")
    imp.append(colunas)
    aplicar_cabecalho(imp)
    for c, l in larguras.items():
        imp.column_dimensions[c].width = l
    imp.freeze_panes = "A2"

    # Aba Listas
    listas = wb.create_sheet("Listas")
    listas.append(["Status", "Tipo"])
    listas.append(["Pago", "Fixo"])
    listas.append(["Pendente", "Avulso"])
    listas.column_dimensions["A"].width = 14
    listas.column_dimensions["B"].width = 14

    wb.save(ARQUIVO_SAIDA)
    print(f"Arquivo gerado: {ARQUIVO_SAIDA}")


if __name__ == "__main__":
    criar_planilha()
