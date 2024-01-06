import os
import pandas as pd

def read_cc_xls(file_path):
    df = pd.DataFrame()
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, 0, header=26, usecols='B:F', skipfooter=1, converters={'E': float, 'F': float})
    df.drop(df.columns[2], axis=1, inplace=True)
    df = df.dropna(subset=[df.columns[0]])
    df = df.fillna(0)
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y')
    df = df[~df[df.columns[1]].str.contains('SALDO INICIAL')]
    df = df[~df[df.columns[1]].str.contains('SALDO FINAL')]
    df.columns = ['Date', 'Payee', 'Outflow', 'Inflow']
    df.insert(2, 'Memo', '')
    df = df.sort_values(by=['Date'])
    return df

# Cuenta Corrriente
cartolas = os.listdir('../cartolas')
if 'cartola.xls' in cartolas:
    cc_df = read_cc_xls('../cartolas/cartola.xls')
    cc_df.to_csv('output.csv', index=False)