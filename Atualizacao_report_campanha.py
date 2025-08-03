import pandas as pd

def load_excel(siprocal):
    df_siprocal = pd.read_excel(siprocal)
    return df_siprocal

def load_csv(verisoft):
    df_verisoft = pd.read_csv(verisoft)
    return df_verisoft

def filtrar_campanha(siprocal, campanhas, colunas_desejadas):
    # Verifica que las columnas existan
    for col in colunas_desejadas:
        if col not in siprocal.columns:
            raise ValueError(f'A coluna "{col}" não exite no arquivo.')

    # Filtrar por Campaña
    regex = '|'.join(map(str, campanhas))
    filtrado = siprocal[siprocal['Campaign Name'].astype(str).str.contains(regex, case=False, na=False)]
    return filtrado[colunas_desejadas].reset_index(drop=True)

def save_data(siprocal, output_path):
    siprocal['Date'] = pd.to_datetime(siprocal['Date'])
    siprocal = siprocal.sort_values(by='Date')
    siprocal['Date'] = siprocal['Date'].dt.strftime('%d/%m/%Y')
    siprocal.to_excel(output_path, index=False)

# === Configurando rutas de los archivos ===
siprocal = r'C:\Users\xxxx\xxxx\Área de Trabalho\raw\julho\23\Report 23.07.25.xlsx'
verisoft = r'C:\Users\xxxx\xxxx\Área de Trabalho\raw\julho\23\report-2025-07-23-09-44-36.csv'
output_path = r'C:\Users\xxxxx\xxxxx\Área de Trabalho\trusted\julho\23\finalizado_23-07.xlsx'

# Filtro de las campañas
campanhas = [
    '3_campanha01_202507',
    '3_campanha02_202507',
    '3_campanha03_202507'
]

colunas_desejadas = [
    'Date',
    'Campaign Name',
    'Impressions',
    'Clicks'
]

# === Ejecución ===
df_siprocal = load_excel(siprocal)
filtrado = filtrar_campanha(df_siprocal, campanhas, colunas_desejadas)
save_data(filtrado, output_path)

print(filtrado.head())
