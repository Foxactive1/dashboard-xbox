import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import xlsxwriter

# ------------------------------
# 🎮 Gerar dados simulados de vendas Xbox
# ------------------------------

np.random.seed(42)
num_registros = 200

produtos = ["Xbox Series X", "Xbox Series S", "Controle Xbox", "Assinatura Game Pass", "Headset Xbox"]
regioes = ["Sudeste", "Sul", "Nordeste", "Centro-Oeste", "Norte"]
vendedores = ["Lucas", "Mariana", "João", "Ana", "Carlos"]
canais = ["E-commerce", "Loja Física", "Marketplace"]

dados = {
    "Data da Venda": [datetime.today() - timedelta(days=np.random.randint(0, 365)) for _ in range(num_registros)],
    "Produto": np.random.choice(produtos, num_registros),
    "Região": np.random.choice(regioes, num_registros),
    "Canal": np.random.choice(canais, num_registros),
    "Vendedor": np.random.choice(vendedores, num_registros),
    "Quantidade": np.random.randint(1, 10, num_registros),
    "Preço Unitário": np.random.uniform(199.0, 4999.0, num_registros).round(2)
}

df = pd.DataFrame(dados)
df["Receita Total"] = (df["Quantidade"] * df["Preço Unitário"]).round(2)

# ------------------------------
# 💼 Criar arquivo Excel com Dashboard
# ------------------------------

output_file = "Dashboard_Vendas_Xbox_Completo_InNovaIdeia.xlsx"

with pd.ExcelWriter(output_file, engine='xlsxwriter', datetime_format='yyyy-mm-dd') as writer:
    df.to_excel(writer, sheet_name="Base_Vendas", index=False)
    workbook = writer.book
    worksheet = writer.sheets["Base_Vendas"]

    # 📊 Gráfico 1: Colunas (Receita por Região)
    chart1 = workbook.add_chart({'type': 'column'})
    chart1.add_series({
        'name': 'Receita por Região',
        'categories': f'=Base_Vendas!C2:C{len(df)+1}',
        'values':     f'=Base_Vendas!H2:H{len(df)+1}',
        'gap': 30,
    })
    chart1.set_title({'name': 'Receita Total por Região'})
    chart1.set_x_axis({'name': 'Região'})
    chart1.set_y_axis({'name': 'Receita'})
    chart1.set_style(10)
    worksheet.insert_chart('J2', chart1)

    # 🥧 Gráfico 2: Pizza (Receita por Produto)
    chart2 = workbook.add_chart({'type': 'pie'})
    pivot_produto = df.groupby("Produto")["Receita Total"].sum().reset_index()
    start_row = len(df) + 5
    worksheet.write_row(f'A{start_row}', ['Produto', 'Receita Total'])
    for i, row in enumerate(pivot_produto.values):
        worksheet.write_row(start_row + 1 + i, 0, row)

    chart2.add_series({
        'name': 'Receita por Produto',
        'categories': f'=Base_Vendas!$A${start_row+2}:$A${start_row+1+len(pivot_produto)}',
        'values':     f'=Base_Vendas!$B${start_row+2}:$B${start_row+1+len(pivot_produto)}',
    })
    chart2.set_title({'name': 'Receita por Produto'})
    chart2.set_style(10)
    worksheet.insert_chart('J20', chart2)

    # 📈 Gráfico 3: Linha (Receita por Data)
    chart3 = workbook.add_chart({'type': 'line'})
    df_sorted = df.sort_values("Data da Venda")
    pivot_data = df_sorted.groupby(df_sorted["Data da Venda"].dt.date)["Receita Total"].sum().reset_index()
    start_row2 = start_row + len(pivot_produto) + 5
    worksheet.write_row(f'A{start_row2}', ['Data', 'Receita Total'])
    for i, row in enumerate(pivot_data.values):
        worksheet.write_row(start_row2 + 1 + i, 0, row)

    chart3.add_series({
        'name': 'Tendência de Receita',
        'categories': f'=Base_Vendas!$A${start_row2+2}:$A${start_row2+1+len(pivot_data)}',
        'values':     f'=Base_Vendas!$B${start_row2+2}:$B${start_row2+1+len(pivot_data)}',
    })
    chart3.set_title({'name': 'Tendência de Receita por Data'})
    chart3.set_x_axis({'name': 'Data'})
    chart3.set_y_axis({'name': 'Receita'})
    chart3.set_style(12)
    worksheet.insert_chart('J38', chart3)

print(f"✅ Dashboard gerado: {output_file}")
