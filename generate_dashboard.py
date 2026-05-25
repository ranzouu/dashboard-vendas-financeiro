import pandas as pd
import xlsxwriter

# Caminhos
INPUT_FILE = '/Users/rayson/Downloads/805d54f9-6d53-4246-bed7-4aa2da615923.xlsx'
OUTPUT_FILE = '/Users/rayson/Vendas_Dashboard_Project/Dashboard_Vendas.xlsx'

def generate():
    # 1. Carregar os dados (Aba 'Bases' ou sheetId=2)
    # Como os nomes das abas têm caracteres especiais, vamos tentar carregar pela ordem ou pelo nome exato se possível.
    # rId2 mapeia para sheet2.xml. Vamos ler todas as abas e procurar a que tem 'Subscriber ID'.
    xls = pd.ExcelFile(INPUT_FILE)
    df = None
    for sheet_name in xls.sheet_names:
        temp_df = pd.read_excel(xls, sheet_name=sheet_name)
        if 'Subscriber ID' in temp_df.columns:
            df = temp_df
            break
    
    if df is None:
        print("Erro: Não foi possível encontrar a aba com os dados de 'Subscriber ID'.")
        return

    # Normalizar nomes de colunas (remover quebras de linha e espaços extras)
    df.columns = df.columns.str.replace('\n', ' ').str.strip()

    # Limpeza básica
    df['Total Value'] = pd.to_numeric(df['Total Value'], errors='coerce').fillna(0)
    df['Coupon Value'] = pd.to_numeric(df['Coupon Value'], errors='coerce').fillna(0)
    df['Subscription Price'] = pd.to_numeric(df['Subscription Price'], errors='coerce').fillna(0)
    df['EA Play Season Pass Price'] = pd.to_numeric(df['EA Play Season Pass Price'], errors='coerce').fillna(0)
    df['Minecraft Season Pass Price'] = pd.to_numeric(df['Minecraft Season Pass Price'], errors='coerce').fillna(0)

    # 2. Cálculos de KPIs
    total_revenue = df['Total Value'].sum()
    avg_ticket = df['Total Value'].mean()
    total_coupons = df['Coupon Value'].sum()
    total_subscribers = len(df)
    extra_revenue = df['EA Play Season Pass Price'].sum() + df['Minecraft Season Pass Price'].sum()

    # Agrupamentos para Gráficos
    rev_by_plan = df.groupby('Plan')['Total Value'].sum().reset_index()
    rev_by_type = df.groupby('Subscription Type')['Total Value'].sum().reset_index()

    # 3. Criar o Excel com o Dashboard
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
    
    # Aba de Dados (para referência e tabelas dinâmicas ocultas)
    df.to_excel(writer, sheet_name='Dados_Base', index=False)
    
    workbook = writer.book
    dashboard = workbook.add_worksheet('Dashboard')
    dashboard.hide_gridlines(2) # Esconder linhas de grade

    # Estilos
    header_fmt = workbook.add_format({
        'bold': True, 'font_size': 24, 'font_name': 'Arial',
        'font_color': '#FFFFFF', 'bg_color': '#0F172A',
        'align': 'center', 'valign': 'vcenter'
    })
    
    kpi_title_fmt = workbook.add_format({
        'bold': True, 'font_size': 12, 'font_name': 'Arial',
        'font_color': '#64748B', 'align': 'center'
    })
    
    kpi_val_fmt = workbook.add_format({
        'bold': True, 'font_size': 20, 'font_name': 'Arial',
        'font_color': '#1E293B', 'align': 'center', 'num_format': '"R$ "#,##0'
    })
    
    kpi_sub_fmt = workbook.add_format({
        'bold': True, 'font_size': 20, 'font_name': 'Arial',
        'font_color': '#1E293B', 'align': 'center', 'num_format': '#,##0'
    })

    # Layout do Dashboard
    # Título
    dashboard.merge_range('B2:K3', 'DASHBOARD FINANCEIRO DE VENDAS', header_fmt)

    # Cartões de KPI
    # Receita Total
    dashboard.write('B5', 'RECEITA TOTAL', kpi_title_fmt)
    dashboard.write('B6', total_revenue, kpi_val_fmt)
    
    # Inscritos
    dashboard.write('E5', 'TOTAL INSCRITOS', kpi_title_fmt)
    dashboard.write('E6', total_subscribers, kpi_sub_fmt)
    
    # Ticket Médio
    dashboard.write('H5', 'TICKET MÉDIO', kpi_title_fmt)
    dashboard.write('H6', avg_ticket, kpi_val_fmt)
    
    # Descontos
    dashboard.write('K5', 'TOTAL DESCONTOS', kpi_title_fmt)
    dashboard.write('K6', total_coupons, kpi_val_fmt)

    # --- Gráfico 1: Receita por Plano ---
    chart_plan = workbook.add_chart({'type': 'column'})
    
    # Adicionar dados de suporte em uma aba oculta ou na Dados_Base
    # Vamos escrever o resumo na Dados_Base a partir da coluna P
    start_row = 1
    dashboard.write('B9', 'Receita por Plano', kpi_title_fmt)
    
    for i, row in rev_by_plan.iterrows():
        dashboard.write(10 + i, 15, row['Plan']) # Coluna P (15)
        dashboard.write(10 + i, 16, row['Total Value']) # Coluna Q (16)
        
    chart_plan.add_series({
        'name': 'Receita por Plano',
        'categories': ['Dashboard', 10, 15, 10 + len(rev_by_plan) - 1, 15],
        'values':     ['Dashboard', 10, 16, 10 + len(rev_by_plan) - 1, 16],
        'fill':       {'color': '#3B82F6'},
        'data_labels': {'value': True, 'font': {'color': '#1E293B'}},
    })
    chart_plan.set_title({'name': 'Distribuição de Receita por Plano', 'name_font': {'size': 14, 'bold': True}})
    chart_plan.set_legend({'none': True})
    chart_plan.set_x_axis({'line': {'none': True}})
    chart_plan.set_y_axis({'visible': False})
    
    dashboard.insert_chart('B10', chart_plan, {'x_scale': 1.2, 'y_scale': 1.2})

    # --- Gráfico 2: Receita por Tipo de Assinatura ---
    chart_type = workbook.add_chart({'type': 'doughnut'})
    
    for i, row in rev_by_type.iterrows():
        dashboard.write(20 + i, 15, row['Subscription Type']) # Coluna P
        dashboard.write(20 + i, 16, row['Total Value']) # Coluna Q

    chart_type.add_series({
        'name': 'Tipo de Assinatura',
        'categories': ['Dashboard', 20, 15, 20 + len(rev_by_type) - 1, 15],
        'values':     ['Dashboard', 20, 16, 20 + len(rev_by_type) - 1, 16],
        'points': [
            {'fill': {'color': '#10B981'}}, # Emerald
            {'fill': {'color': '#F59E0B'}}, # Amber
            {'fill': {'color': '#6366F1'}}, # Indigo
        ],
        'data_labels': {'percentage': True, 'category': True, 'position': 'outside_end'},
    })
    chart_type.set_title({'name': 'Share de Receita por Tipo', 'name_font': {'size': 14, 'bold': True}})
    
    dashboard.insert_chart('G10', chart_type, {'x_scale': 1.2, 'y_scale': 1.2})

    # Fechar e salvar
    writer.close()
    print(f"Dashboard gerado com sucesso em: {OUTPUT_FILE}")

if __name__ == "__main__":
    generate()
