import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

def build_financial_workbook():
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # 1. ABA: Configurações
    # -------------------------------------------------------------
    ws_config = wb.active
    ws_config.title = "Configurações"
    ws_config.views.sheetView[0].showGridLines = True
    
    # Estilos Base Config
    navy_header_fill = PatternFill(start_color="16365C", end_color="16365C", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="Segoe UI", size=10)
    bold_font = Font(name="Segoe UI", size=10, bold=True)
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    # Tabela Tipos
    ws_config["A1"] = "Tipos de Lançamento"
    ws_config["A1"].fill = navy_header_fill
    ws_config["A1"].font = header_font
    ws_config["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    tipos = ["Receita", "Despesa"]
    for idx, tipo in enumerate(tipos, start=2):
        cell = ws_config.cell(row=idx, column=1, value=tipo)
        cell.font = cell_font
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="center")

    # Tabela Categorias
    ws_config["C1"] = "Categorias"
    ws_config["C1"].fill = navy_header_fill
    ws_config["C1"].font = header_font
    ws_config["C1"].alignment = Alignment(horizontal="center", vertical="center")
    
    categorias = [
        "Salário",
        "Rendas Extras",
        "Moradia",
        "Alimentação",
        "Transporte",
        "Estudos",
        "Saúde",
        "Lazer",
        "Reserva",
        "Outras Despesas"
    ]
    for idx, cat in enumerate(categorias, start=2):
        cell = ws_config.cell(row=idx, column=3, value=cat)
        cell.font = cell_font
        cell.border = border_thin
        
    # Tabela Meses
    ws_config["E1"] = "Meses"
    ws_config["E1"].fill = navy_header_fill
    ws_config["E1"].font = header_font
    ws_config["E1"].alignment = Alignment(horizontal="center", vertical="center")
    
    meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    for idx, mes in enumerate(meses, start=2):
        cell = ws_config.cell(row=idx, column=5, value=mes)
        cell.font = cell_font
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="center")

    # Tabela Opções de Filtro de Mês
    ws_config["G1"] = "Filtro Mês"
    ws_config["G1"].fill = navy_header_fill
    ws_config["G1"].font = header_font
    ws_config["G1"].alignment = Alignment(horizontal="center", vertical="center")
    
    filtro_meses = ["Todos", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    for idx, f_mes in enumerate(filtro_meses, start=2):
        cell = ws_config.cell(row=idx, column=7, value=f_mes)
        cell.font = cell_font
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="center")

    # Tabela Parâmetros / Metas
    ws_config["I1"] = "Parâmetro"
    ws_config["I1"].fill = navy_header_fill
    ws_config["I1"].font = header_font
    ws_config["J1"] = "Valor"
    ws_config["J1"].fill = navy_header_fill
    ws_config["J1"].font = header_font
    
    ws_config["I2"] = "Meta Reserva (% Receitas)"
    ws_config["I2"].font = cell_font
    ws_config["I2"].border = border_thin
    ws_config["J2"] = 0.20
    ws_config["J2"].font = bold_font
    ws_config["J2"].number_format = "0.0%"
    ws_config["J2"].border = border_thin
    ws_config["J2"].alignment = Alignment(horizontal="center")

    ws_config.column_dimensions["A"].width = 22
    ws_config.column_dimensions["B"].width = 4
    ws_config.column_dimensions["C"].width = 22
    ws_config.column_dimensions["D"].width = 4
    ws_config.column_dimensions["E"].width = 12
    ws_config.column_dimensions["F"].width = 4
    ws_config.column_dimensions["G"].width = 14
    ws_config.column_dimensions["H"].width = 4
    ws_config.column_dimensions["I"].width = 28
    ws_config.column_dimensions["J"].width = 14

    # -------------------------------------------------------------
    # 2. ABA: Lançamentos
    # -------------------------------------------------------------
    ws_lanc = wb.create_sheet(title="Lançamentos")
    ws_lanc.views.sheetView[0].showGridLines = True

    # Título da Tabela de Lançamentos
    ws_lanc.merge_cells("A1:G1")
    ws_lanc["A1"] = "Base de Lançamentos Financeiros"
    ws_lanc["A1"].fill = navy_header_fill
    ws_lanc["A1"].font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    ws_lanc["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_lanc.row_dimensions[1].height = 30

    headers_lanc = ["Data", "Descrição", "Categoria", "Tipo", "Valor (R$)", "Mês", "Observações"]
    for col_num, header in enumerate(headers_lanc, 1):
        cell = ws_lanc.cell(row=2, column=col_num, value=header)
        cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        cell.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin
    ws_lanc.row_dimensions[2].height = 24

    # Dados fiéis à imagem de referência
    raw_data = [
        ("05/01/2025", "Salário - Empresa XPTO", "Salário", "Receita", 4250.00, "Jan", "Salário líquido"),
        ("10/01/2025", "Freelance - Consultoria RH", "Rendas Extras", "Receita", 600.00, "Jan", "Projeto de 5 dias"),
        ("12/01/2025", "Aluguel", "Moradia", "Despesa", 1200.00, "Jan", "Apartamento"),
        ("15/01/2025", "Supermercado Extra", "Alimentação", "Despesa", 320.40, "Jan", "Compras do mês"),
        ("18/01/2025", "Uber / 99", "Transporte", "Despesa", 87.60, "Jan", "Deslocamentos"),
        ("22/01/2025", "Curso Online - Finanças", "Estudos", "Despesa", 197.00, "Jan", "Plataforma Alura"),
        ("25/01/2025", "Livro - Pai Rico, Pai Pobre", "Estudos", "Despesa", 39.90, "Jan", "Compra na Amazon"),
        ("05/02/2025", "Salário - Empresa XPTO", "Salário", "Receita", 4250.00, "Fev", "Salário líquido"),
        ("10/02/2025", "Freelance - Análise Planilhas", "Rendas Extras", "Receita", 250.00, "Fev", "Entrega de planilhas"),
        ("12/02/2025", "Academia Smart Fit", "Saúde", "Despesa", 129.90, "Fev", "Mensalidade"),
        ("14/02/2025", "Cinema - Final de Semana", "Lazer", "Despesa", 78.00, "Fev", "Ingressos + pipoca"),
        ("20/02/2025", "Transferência para Reserva", "Reserva", "Despesa", 500.00, "Fev", "Aplicação Nubank"),
        ("25/02/2025", "Mensalidade - Faculdade", "Estudos", "Despesa", 375.90, "Fev", "2ª parcela"),
        ("05/03/2025", "Salário - Empresa XPTO", "Salário", "Receita", 4250.00, "Mar", "Salário líquido"),
        ("10/03/2025", "Freelance - Dashboard Excel", "Rendas Extras", "Receita", 300.00, "Mar", "Dashboard interativo"),
        ("15/03/2025", "Despesas Diversas / Farmácia", "Outras Despesas", "Despesa", 286.70, "Mar", "Despesas pontuais"),
        ("20/01/2025", "Aporte Reserva Emergência", "Reserva", "Despesa", 350.00, "Jan", "Aplicação CDB"),
        ("02/01/2025", "Saldo Anterior Reserva", "Reserva", "Despesa", 1500.00, "Jan", "Reserva de emergência acumulada"),
    ]

    cat_colors = {
        "Salário": ("E8F8F5", "27AE60"),
        "Rendas Extras": ("FEF9E7", "D35400"),
        "Moradia": ("FDEDEC", "C0392B"),
        "Alimentação": ("FDEDEC", "E74C3C"),
        "Transporte": ("EBF5FB", "2980B9"),
        "Estudos": ("F4ECF7", "8E44AD"),
        "Saúde": ("E8F8F5", "16A085"),
        "Lazer": ("FEF5E7", "D35400"),
        "Reserva": ("EAFAF1", "2E7D32"),
        "Outras Despesas": ("FADBD8", "922B21")
    }

    zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    for row_idx, data in enumerate(raw_data, start=3):
        data_str, desc, cat, tipo, val, mes, obs = data
        
        c1 = ws_lanc.cell(row=row_idx, column=1, value=data_str)
        c2 = ws_lanc.cell(row=row_idx, column=2, value=desc)
        c3 = ws_lanc.cell(row=row_idx, column=3, value=cat)
        c4 = ws_lanc.cell(row=row_idx, column=4, value=tipo)
        c5 = ws_lanc.cell(row=row_idx, column=5, value=val)
        c6 = ws_lanc.cell(row=row_idx, column=6, value=mes)
        c7 = ws_lanc.cell(row=row_idx, column=7, value=obs)
        
        row_fill = zebra_fill if row_idx % 2 == 0 else white_fill
        for c in [c1, c2, c4, c5, c6, c7]:
            c.fill = row_fill
            c.font = cell_font
            c.border = border_thin
            
        c1.alignment = Alignment(horizontal="center")
        c2.alignment = Alignment(horizontal="left")
        
        # Categoria com cor de destaque sutil
        if cat in cat_colors:
            bg_hex, txt_hex = cat_colors[cat]
            c3.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
            c3.font = Font(name="Segoe UI", size=10, bold=True, color=txt_hex)
        else:
            c3.fill = row_fill
            c3.font = bold_font
        c3.alignment = Alignment(horizontal="center")
        c3.border = border_thin

        # Tipo
        if tipo == "Receita":
            c4.font = Font(name="Segoe UI", size=10, bold=True, color="27AE60")
        else:
            c4.font = Font(name="Segoe UI", size=10, bold=True, color="C0392B")
        c4.alignment = Alignment(horizontal="center")

        # Valor
        c5.number_format = "R$ #,##0.00"
        c5.font = bold_font
        c5.alignment = Alignment(horizontal="right")
        
        c6.alignment = Alignment(horizontal="center")
        c7.alignment = Alignment(horizontal="left")
        ws_lanc.row_dimensions[row_idx].height = 20

    # Data Validation para Lançamentos (até linha 500)
    dv_tipo = DataValidation(type="list", formula1="'Configurações'!$A$2:$A$3", allow_blank=True)
    dv_cat = DataValidation(type="list", formula1="'Configurações'!$C$2:$C$11", allow_blank=True)
    dv_mes = DataValidation(type="list", formula1="'Configurações'!$E$2:$E$13", allow_blank=True)

    ws_lanc.add_data_validation(dv_tipo)
    ws_lanc.add_data_validation(dv_cat)
    ws_lanc.add_data_validation(dv_mes)

    dv_tipo.add("D3:D500")
    dv_cat.add("C3:C500")
    dv_mes.add("F3:F500")

    ws_lanc.column_dimensions["A"].width = 14
    ws_lanc.column_dimensions["B"].width = 32
    ws_lanc.column_dimensions["C"].width = 18
    ws_lanc.column_dimensions["D"].width = 12
    ws_lanc.column_dimensions["E"].width = 16
    ws_lanc.column_dimensions["F"].width = 10
    ws_lanc.column_dimensions["G"].width = 30
    
    ws_lanc.auto_filter.ref = f"A2:G{len(raw_data)+2}"

    # -------------------------------------------------------------
    # 3. ABA: Dashboard
    # -------------------------------------------------------------
    ws_dash = wb.create_sheet(title="Dashboard", index=0)
    ws_dash.views.sheetView[0].showGridLines = True

    # 3.1 Cabeçalho Principal (Linha 1)
    ws_dash.merge_cells("A1:K1")
    ws_dash["A1"] = "Controle Financeiro Pessoal - 12 Meses"
    ws_dash["A1"].fill = navy_header_fill
    ws_dash["A1"].font = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
    ws_dash["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # Seletor de Mês no Cabeçalho (Colunas L a N)
    ws_dash.merge_cells("L1:M1")
    ws_dash["L1"] = "Filtrar Mês:"
    ws_dash["L1"].fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    ws_dash["L1"].font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    ws_dash["L1"].alignment = Alignment(horizontal="right", vertical="center")

    ws_dash["N1"] = "Todos"
    ws_dash["N1"].fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    ws_dash["N1"].font = Font(name="Segoe UI", size=12, bold=True, color="16365C")
    ws_dash["N1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash["N1"].border = Border(
        left=Side(style='medium', color='1F4E79'),
        right=Side(style='medium', color='1F4E79'),
        top=Side(style='medium', color='1F4E79'),
        bottom=Side(style='medium', color='1F4E79')
    )

    dv_filtro = DataValidation(type="list", formula1="'Configurações'!$G$2:$G$14", allow_blank=False)
    ws_dash.add_data_validation(dv_filtro)
    dv_filtro.add("N1")

    ws_dash.row_dimensions[1].height = 36
    ws_dash.row_dimensions[2].height = 10

    # 3.2 Título da Seção Resumo Mensal
    ws_dash["A3"] = "Resumo Mensal"
    ws_dash["A3"].font = Font(name="Segoe UI", size=13, bold=True, color="16365C")
    ws_dash.row_dimensions[3].height = 22

    # 3.3 Construção dos 7 Cards de KPI (Linhas 4 a 7)
    card_border_green = Border(
        left=Side(style='thin', color='A3D9B5'),
        right=Side(style='thin', color='A3D9B5'),
        top=Side(style='thin', color='A3D9B5'),
        bottom=Side(style='thin', color='A3D9B5')
    )
    card_border_red = Border(
        left=Side(style='thin', color='F5B7B1'),
        right=Side(style='thin', color='F5B7B1'),
        top=Side(style='thin', color='F5B7B1'),
        bottom=Side(style='thin', color='F5B7B1')
    )
    card_border_blue = Border(
        left=Side(style='thin', color='AED6F1'),
        right=Side(style='thin', color='AED6F1'),
        top=Side(style='thin', color='AED6F1'),
        bottom=Side(style='thin', color='AED6F1')
    )
    card_border_purple = Border(
        left=Side(style='thin', color='D7BDE2'),
        right=Side(style='thin', color='D7BDE2'),
        top=Side(style='thin', color='D7BDE2'),
        bottom=Side(style='thin', color='D7BDE2')
    )

    fill_green = PatternFill(start_color="F4FAF6", end_color="F4FAF6", fill_type="solid")
    fill_red = PatternFill(start_color="FDF4F4", end_color="FDF4F4", fill_type="solid")
    fill_blue = PatternFill(start_color="F2F8FD", end_color="F2F8FD", fill_type="solid")
    fill_purple = PatternFill(start_color="F9F4FB", end_color="F9F4FB", fill_type="solid")

    def apply_card_style(ws, start_col, end_col, start_row, end_row, fill, border):
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = fill
                top_s = border.top if r == start_row else None
                bot_s = border.bottom if r == end_row else None
                left_s = border.left if c == start_col else None
                right_s = border.right if c == end_col else None
                cell.border = Border(top=top_s, bottom=bot_s, left=left_s, right=right_s)

    # Card 1: Salário Líquido (A4:B7)
    apply_card_style(ws_dash, 1, 2, 4, 7, fill_green, card_border_green)
    ws_dash.merge_cells("A4:B4")
    ws_dash["A4"] = "💵  Salário Líquido"
    ws_dash["A4"].font = Font(name="Segoe UI", size=10, bold=True, color="1E8449")
    ws_dash["A4"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("A5:B6")
    ws_dash["A5"] = '=IF($N$1="Todos", SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, "Salário", \'Lançamentos\'!$D$3:$D$500, "Receita"), SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, "Salário", \'Lançamentos\'!$D$3:$D$500, "Receita", \'Lançamentos\'!$F$3:$F$500, $N$1))'
    ws_dash["A5"].font = Font(name="Segoe UI", size=16, bold=True, color="1E8449")
    ws_dash["A5"].number_format = "R$ #,##0.00"
    ws_dash["A5"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("A7:B7")
    ws_dash["A7"] = "(Receita principal)"
    ws_dash["A7"].font = Font(name="Segoe UI", size=8, italic=True, color="566573")
    ws_dash["A7"].alignment = Alignment(horizontal="center", vertical="center")

    # Card 2: Rendas Extras (C4:D7)
    apply_card_style(ws_dash, 3, 4, 4, 7, fill_green, card_border_green)
    ws_dash.merge_cells("C4:D4")
    ws_dash["C4"] = "💼  Rendas Extras (Freelancer)"
    ws_dash["C4"].font = Font(name="Segoe UI", size=10, bold=True, color="1E8449")
    ws_dash["C4"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("C5:D6")
    ws_dash["C5"] = '=IF($N$1="Todos", SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, "Rendas Extras", \'Lançamentos\'!$D$3:$D$500, "Receita"), SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, "Rendas Extras", \'Lançamentos\'!$D$3:$D$500, "Receita", \'Lançamentos\'!$F$3:$F$500, $N$1))'
    ws_dash["C5"].font = Font(name="Segoe UI", size=16, bold=True, color="1E8449")
    ws_dash["C5"].number_format = "R$ #,##0.00"
    ws_dash["C5"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("C7:D7")
    ws_dash["C7"] = "(Projetos e consultoria)"
    ws_dash["C7"].font = Font(name="Segoe UI", size=8, italic=True, color="566573")
    ws_dash["C7"].alignment = Alignment(horizontal="center", vertical="center")

    # Card 3: Total Receitas (E4:F7)
    apply_card_style(ws_dash, 5, 6, 4, 7, fill_green, card_border_green)
    ws_dash.merge_cells("E4:F4")
    ws_dash["E4"] = "Total Receitas"
    ws_dash["E4"].font = Font(name="Segoe UI", size=10, bold=True, color="1E8449")
    ws_dash["E4"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("E5:F6")
    ws_dash["E5"] = '=IF($N$1="Todos", SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$D$3:$D$500, "Receita"), SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$D$3:$D$500, "Receita", \'Lançamentos\'!$F$3:$F$500, $N$1))'
    ws_dash["E5"].font = Font(name="Segoe UI", size=16, bold=True, color="1E8449")
    ws_dash["E5"].number_format = "R$ #,##0.00"
    ws_dash["E5"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("E7:F7")
    ws_dash["E7"] = "(Salário + Extras)"
    ws_dash["E7"].font = Font(name="Segoe UI", size=8, italic=True, color="566573")
    ws_dash["E7"].alignment = Alignment(horizontal="center", vertical="center")

    # Card 4: Total Despesas (G4:H7)
    apply_card_style(ws_dash, 7, 8, 4, 7, fill_red, card_border_red)
    ws_dash.merge_cells("G4:H4")
    ws_dash["G4"] = "⬇️  Total Despesas"
    ws_dash["G4"].font = Font(name="Segoe UI", size=10, bold=True, color="C0392B")
    ws_dash["G4"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("G5:H6")
    ws_dash["G5"] = '=IF($N$1="Todos", SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$D$3:$D$500, "Despesa"), SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$D$3:$D$500, "Despesa", \'Lançamentos\'!$F$3:$F$500, $N$1))'
    ws_dash["G5"].font = Font(name="Segoe UI", size=16, bold=True, color="C0392B")
    ws_dash["G5"].number_format = "R$ #,##0.00"
    ws_dash["G5"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("G7:H7")
    ws_dash["G7"] = "(Todos os gastos)"
    ws_dash["G7"].font = Font(name="Segoe UI", size=8, italic=True, color="7F8C8D")
    ws_dash["G7"].alignment = Alignment(horizontal="center", vertical="center")

    # Card 5: Disponível no Mês (I4:J7)
    apply_card_style(ws_dash, 9, 10, 4, 7, fill_green, card_border_green)
    ws_dash.merge_cells("I4:J4")
    ws_dash["I4"] = "💲  Disponível no Mês"
    ws_dash["I4"].font = Font(name="Segoe UI", size=10, bold=True, color="1E8449")
    ws_dash["I4"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("I5:J6")
    ws_dash["I5"] = '=E5-G5'
    ws_dash["I5"].font = Font(name="Segoe UI", size=16, bold=True, color="1E8449")
    ws_dash["I5"].number_format = "R$ #,##0.00"
    ws_dash["I5"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("I7:J7")
    ws_dash["I7"] = "(Receitas - Despesas)"
    ws_dash["I7"].font = Font(name="Segoe UI", size=8, italic=True, color="566573")
    ws_dash["I7"].alignment = Alignment(horizontal="center", vertical="center")

    # Card 6: Reserva Financeira (K4:L7)
    apply_card_style(ws_dash, 11, 12, 4, 7, fill_blue, card_border_blue)
    ws_dash.merge_cells("K4:L4")
    ws_dash["K4"] = "🏦  Reserva Financeira"
    ws_dash["K4"].font = Font(name="Segoe UI", size=10, bold=True, color="1B4F72")
    ws_dash["K4"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("K5:L5")
    ws_dash["K5"] = '=IF($N$1="Todos", SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, "Reserva"), SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, "Reserva", \'Lançamentos\'!$F$3:$F$500, $N$1))'
    ws_dash["K5"].font = Font(name="Segoe UI", size=14, bold=True, color="1B4F72")
    ws_dash["K5"].number_format = "R$ #,##0.00"
    ws_dash["K5"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("K6:L6")
    ws_dash["K6"] = "(Meta: 20% das receitas)"
    ws_dash["K6"].font = Font(name="Segoe UI", size=8, italic=True, color="566573")
    ws_dash["K6"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("K7:L7")
    ws_dash["K7"] = '="Atingido: " & TEXT(IF(E5>0, K5/(E5*\'Configurações\'!$J$2), 0), "0.0%")'
    ws_dash["K7"].font = Font(name="Segoe UI", size=8, bold=True, color="1E8449")
    ws_dash["K7"].alignment = Alignment(horizontal="center", vertical="center")

    # Card 7: Gasto com Estudos (M4:N7)
    apply_card_style(ws_dash, 13, 14, 4, 7, fill_purple, card_border_purple)
    ws_dash.merge_cells("M4:N4")
    ws_dash["M4"] = "🎓  Gasto com Estudos"
    ws_dash["M4"].font = Font(name="Segoe UI", size=10, bold=True, color="6C3483")
    ws_dash["M4"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("M5:N6")
    ws_dash["M5"] = '=IF($N$1="Todos", SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, "Estudos", \'Lançamentos\'!$D$3:$D$500, "Despesa"), SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, "Estudos", \'Lançamentos\'!$D$3:$D$500, "Despesa", \'Lançamentos\'!$F$3:$F$500, $N$1))'
    ws_dash["M5"].font = Font(name="Segoe UI", size=16, bold=True, color="6C3483")
    ws_dash["M5"].number_format = "R$ #,##0.00"
    ws_dash["M5"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.merge_cells("M7:N7")
    ws_dash["M7"] = "(Cursos, livros, faculdade)"
    ws_dash["M7"].font = Font(name="Segoe UI", size=8, italic=True, color="7F8C8D")
    ws_dash["M7"].alignment = Alignment(horizontal="center", vertical="center")

    ws_dash.row_dimensions[4].height = 18
    ws_dash.row_dimensions[5].height = 16
    ws_dash.row_dimensions[6].height = 16
    ws_dash.row_dimensions[7].height = 16
    ws_dash.row_dimensions[8].height = 10

    # 3.4 Tabela de Lançamentos no Dashboard (Colunas A a G, Linhas 9 a 29)
    ws_dash.merge_cells("A9:G9")
    ws_dash["A9"] = "Lançamentos"
    ws_dash["A9"].fill = navy_header_fill
    ws_dash["A9"].font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    ws_dash["A9"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_dash.row_dimensions[9].height = 24

    headers_dash_lanc = ["Data", "Descrição", "Categoria", "Tipo", "Valor (R$)", "Mês", "Observações"]
    for col_num, h_text in enumerate(headers_dash_lanc, 1):
        cell = ws_dash.cell(row=10, column=col_num, value=h_text)
        cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin
    ws_dash.row_dimensions[10].height = 20

    # Espelhar as linhas de lançamentos com proteção para células vazias
    for r_idx in range(11, 29):
        src_row = r_idx - 8  # Linha correspondente em Lançamentos (ex: 11 -> 3)
        ws_dash.row_dimensions[r_idx].height = 19
        row_fill = zebra_fill if r_idx % 2 == 0 else white_fill
        
        c1 = ws_dash.cell(row=r_idx, column=1, value=f"=IF('Lançamentos'!A{src_row}=\"\",\"\",'Lançamentos'!A{src_row})")
        c2 = ws_dash.cell(row=r_idx, column=2, value=f"=IF('Lançamentos'!B{src_row}=\"\",\"\",'Lançamentos'!B{src_row})")
        c3 = ws_dash.cell(row=r_idx, column=3, value=f"=IF('Lançamentos'!C{src_row}=\"\",\"\",'Lançamentos'!C{src_row})")
        c4 = ws_dash.cell(row=r_idx, column=4, value=f"=IF('Lançamentos'!D{src_row}=\"\",\"\",'Lançamentos'!D{src_row})")
        c5 = ws_dash.cell(row=r_idx, column=5, value=f"=IF('Lançamentos'!E{src_row}=\"\",\"\",'Lançamentos'!E{src_row})")
        c6 = ws_dash.cell(row=r_idx, column=6, value=f"=IF('Lançamentos'!F{src_row}=\"\",\"\",'Lançamentos'!F{src_row})")
        c7 = ws_dash.cell(row=r_idx, column=7, value=f"=IF('Lançamentos'!G{src_row}=\"\",\"\",'Lançamentos'!G{src_row})")
        
        for c in [c1, c2, c3, c4, c5, c6, c7]:
            c.fill = row_fill
            c.font = cell_font
            c.border = border_thin
            
        c1.alignment = Alignment(horizontal="center")
        c2.alignment = Alignment(horizontal="left")
        c3.alignment = Alignment(horizontal="center")
        c3.font = bold_font
        c4.alignment = Alignment(horizontal="center")
        c5.number_format = "R$ #,##0.00"
        c5.alignment = Alignment(horizontal="right")
        c5.font = bold_font
        c6.alignment = Alignment(horizontal="center")
        c7.alignment = Alignment(horizontal="left")

    # 3.5 Tabela Gastos por Categoria (Colunas I a K, Linhas 9 a 20)
    ws_dash.merge_cells("I9:K9")
    ws_dash["I9"] = "Gastos por Categoria"
    ws_dash["I9"].fill = navy_header_fill
    ws_dash["I9"].font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    ws_dash["I9"].alignment = Alignment(horizontal="center", vertical="center")

    cat_subheaders = ["Categoria", "Valor (R$)", "% do Total"]
    for col_offset, sub_h in enumerate(cat_subheaders):
        cell = ws_dash.cell(row=10, column=9 + col_offset, value=sub_h)
        cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin

    despesas_categorias = [
        "Moradia",
        "Alimentação",
        "Transporte",
        "Estudos",
        "Saúde",
        "Lazer",
        "Reserva",
        "Outras Despesas"
    ]

    for idx, cat_name in enumerate(despesas_categorias, start=11):
        ws_dash.row_dimensions[idx].height = 19
        row_fill = zebra_fill if idx % 2 == 0 else white_fill
        
        c_cat = ws_dash.cell(row=idx, column=9, value=cat_name)
        c_val = ws_dash.cell(row=idx, column=10, value=f'=IF($N$1="Todos", SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, I{idx}, \'Lançamentos\'!$D$3:$D$500, "Despesa"), SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, I{idx}, \'Lançamentos\'!$D$3:$D$500, "Despesa", \'Lançamentos\'!$F$3:$F$500, $N$1))')
        c_pct = ws_dash.cell(row=idx, column=11, value=f'=IF($J$19>0, J{idx}/$J$19, 0)')
        
        c_cat.fill = row_fill
        c_cat.font = bold_font
        if cat_name in cat_colors:
            _, txt_c = cat_colors[cat_name]
            c_cat.font = Font(name="Segoe UI", size=10, bold=True, color=txt_c)
        c_cat.border = border_thin
        c_cat.alignment = Alignment(horizontal="left", indent=1)
        
        c_val.fill = row_fill
        c_val.font = cell_font
        c_val.number_format = "R$ #,##0.00"
        c_val.border = border_thin
        c_val.alignment = Alignment(horizontal="right")
        
        c_pct.fill = row_fill
        c_pct.font = cell_font
        c_pct.number_format = "0.0%"
        c_pct.border = border_thin
        c_pct.alignment = Alignment(horizontal="right")

    # Linha Total Despesas (Linha 19)
    ws_dash.row_dimensions[19].height = 20
    c_tot_label = ws_dash.cell(row=19, column=9, value="Total Despesas")
    c_tot_label.font = Font(name="Segoe UI", size=10, bold=True, color="C0392B")
    c_tot_label.fill = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    c_tot_label.border = border_thin
    c_tot_label.alignment = Alignment(horizontal="left", indent=1)

    c_tot_val = ws_dash.cell(row=19, column=10, value="=SUM(J11:J18)")
    c_tot_val.font = Font(name="Segoe UI", size=10, bold=True, color="C0392B")
    c_tot_val.fill = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    c_tot_val.number_format = "R$ #,##0.00"
    c_tot_val.border = border_thin
    c_tot_val.alignment = Alignment(horizontal="right")

    c_tot_pct = ws_dash.cell(row=19, column=11, value="=SUM(K11:K18)")
    c_tot_pct.font = Font(name="Segoe UI", size=10, bold=True, color="C0392B")
    c_tot_pct.fill = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    c_tot_pct.number_format = "0.0%"
    c_tot_pct.border = border_thin
    c_tot_pct.alignment = Alignment(horizontal="right")

    # 3.6 Gráfico de Pizza - Distribuição de Gastos
    pie = PieChart()
    pie.title = "Distribuição de Gastos"
    pie.style = 10
    pie.height = 7.5
    pie.width = 13.5
    
    labels = Reference(ws_dash, min_col=9, min_row=11, max_row=18)
    data = Reference(ws_dash, min_col=10, min_row=10, max_row=18)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    
    # Rótulos de dados no gráfico
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showCatName = False
    
    ws_dash.add_chart(pie, "I21")

    # 3.7 Visão Anual - 12 Meses (Linhas 32 a 37)
    ws_dash.row_dimensions[31].height = 12
    ws_dash.merge_cells("A32:N32")
    ws_dash["A32"] = "Visão Anual - 12 Meses"
    ws_dash["A32"].fill = navy_header_fill
    ws_dash["A32"].font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    ws_dash["A32"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_dash.row_dimensions[32].height = 24

    annual_headers = ["Item", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez", "Total Anual"]
    for col_num, h_text in enumerate(annual_headers, 1):
        cell = ws_dash.cell(row=33, column=col_num, value=h_text)
        cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin
    ws_dash.row_dimensions[33].height = 20

    # Linha 34: Receitas
    ws_dash.row_dimensions[34].height = 20
    c_rec = ws_dash.cell(row=34, column=1, value="Receitas")
    c_rec.font = Font(name="Segoe UI", size=10, bold=True, color="1E8449")
    c_rec.fill = white_fill
    c_rec.border = border_thin
    c_rec.alignment = Alignment(horizontal="left", indent=1)

    for m_col, m_name in enumerate(meses, 2):
        cell = ws_dash.cell(row=34, column=m_col, value=f'=SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$D$3:$D$500, "Receita", \'Lançamentos\'!$F$3:$F$500, {get_column_letter(m_col)}$33)')
        cell.font = cell_font
        cell.fill = white_fill
        cell.number_format = "R$ #,##0.00"
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="right")

    c_tot_rec = ws_dash.cell(row=34, column=14, value="=SUM(B34:M34)")
    c_tot_rec.font = Font(name="Segoe UI", size=10, bold=True, color="1E8449")
    c_tot_rec.fill = PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid")
    c_tot_rec.number_format = "R$ #,##0.00"
    c_tot_rec.border = border_thin
    c_tot_rec.alignment = Alignment(horizontal="right")

    # Linha 35: Despesas
    ws_dash.row_dimensions[35].height = 20
    c_desp = ws_dash.cell(row=35, column=1, value="Despesas")
    c_desp.font = Font(name="Segoe UI", size=10, bold=True, color="C0392B")
    c_desp.fill = zebra_fill
    c_desp.border = border_thin
    c_desp.alignment = Alignment(horizontal="left", indent=1)

    for m_col, m_name in enumerate(meses, 2):
        cell = ws_dash.cell(row=35, column=m_col, value=f'=SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$D$3:$D$500, "Despesa", \'Lançamentos\'!$F$3:$F$500, {get_column_letter(m_col)}$33)')
        cell.font = cell_font
        cell.fill = zebra_fill
        cell.number_format = "R$ #,##0.00"
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="right")

    c_tot_desp = ws_dash.cell(row=35, column=14, value="=SUM(B35:M35)")
    c_tot_desp.font = Font(name="Segoe UI", size=10, bold=True, color="C0392B")
    c_tot_desp.fill = PatternFill(start_color="FDEDEC", end_color="FDEDEC", fill_type="solid")
    c_tot_desp.number_format = "R$ #,##0.00"
    c_tot_desp.border = border_thin
    c_tot_desp.alignment = Alignment(horizontal="right")

    # Linha 36: Saldo do Mês
    ws_dash.row_dimensions[36].height = 20
    c_saldo = ws_dash.cell(row=36, column=1, value="Saldo do Mês")
    c_saldo.font = Font(name="Segoe UI", size=10, bold=True, color="16365C")
    c_saldo.fill = white_fill
    c_saldo.border = border_thin
    c_saldo.alignment = Alignment(horizontal="left", indent=1)

    for m_col in range(2, 14):
        col_let = get_column_letter(m_col)
        cell = ws_dash.cell(row=36, column=m_col, value=f'={col_let}34-{col_let}35')
        cell.font = bold_font
        cell.fill = white_fill
        cell.number_format = "R$ #,##0.00"
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="right")

    c_tot_saldo = ws_dash.cell(row=36, column=14, value="=N34-N35")
    c_tot_saldo.font = Font(name="Segoe UI", size=10, bold=True, color="16365C")
    c_tot_saldo.fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
    c_tot_saldo.number_format = "R$ #,##0.00"
    c_tot_saldo.border = border_thin
    c_tot_saldo.alignment = Alignment(horizontal="right")

    # Linha 37: Reserva do Mês
    ws_dash.row_dimensions[37].height = 20
    c_res = ws_dash.cell(row=37, column=1, value="Reserva do Mês")
    c_res.font = Font(name="Segoe UI", size=10, bold=True, color="1B4F72")
    c_res.fill = zebra_fill
    c_res.border = border_thin
    c_res.alignment = Alignment(horizontal="left", indent=1)

    for m_col in range(2, 14):
        col_let = get_column_letter(m_col)
        cell = ws_dash.cell(row=37, column=m_col, value=f'=SUMIFS(\'Lançamentos\'!$E$3:$E$500, \'Lançamentos\'!$C$3:$C$500, "Reserva", \'Lançamentos\'!$F$3:$F$500, {col_let}$33)')
        cell.font = cell_font
        cell.fill = zebra_fill
        cell.number_format = "R$ #,##0.00"
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="right")

    c_tot_res = ws_dash.cell(row=37, column=14, value="=SUM(B37:M37)")
    c_tot_res.font = Font(name="Segoe UI", size=10, bold=True, color="1B4F72")
    c_tot_res.fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
    c_tot_res.number_format = "R$ #,##0.00"
    c_tot_res.border = border_thin
    c_tot_res.alignment = Alignment(horizontal="right")

    # 3.8 Largura das Colunas no Dashboard
    ws_dash.column_dimensions["A"].width = 13
    ws_dash.column_dimensions["B"].width = 28
    ws_dash.column_dimensions["C"].width = 16
    ws_dash.column_dimensions["D"].width = 10
    ws_dash.column_dimensions["E"].width = 14
    ws_dash.column_dimensions["F"].width = 7
    ws_dash.column_dimensions["G"].width = 24
    ws_dash.column_dimensions["H"].width = 3
    ws_dash.column_dimensions["I"].width = 18
    ws_dash.column_dimensions["J"].width = 15
    ws_dash.column_dimensions["K"].width = 13
    ws_dash.column_dimensions["L"].width = 13
    ws_dash.column_dimensions["M"].width = 13
    ws_dash.column_dimensions["N"].width = 16

    # Salvar o Workbook
    output_filename = "Controle_Financeiro_2025.xlsx"
    wb.save(output_filename)
    print(f"Planilha '{output_filename}' atualizada e salva com sucesso!")

if __name__ == "__main__":
    build_financial_workbook()
