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

def read_tc_xls(file_path):
    df = pd.DataFrame()
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, 0, header=17, usecols='B:K', skipfooter=1, converters={'E': float, 'F': float})
    for col in [8,7,6,5,4,2,1]:
        df.drop(df.columns[col], axis=1, inplace=True)
    df = df.iloc[:, :3]
    df = df.dropna(subset=[df.columns[0]])
    df = df.fillna(0)
    df.columns = ['Date', 'Payee', 'Outflow']
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')
    df.insert(2, 'Memo', '')
    df.insert(4, 'Inflow', 0)
    df = df.sort_values(by=['Date'])
    return df

def remove_duplicates(df, prev_df, ftype='cc'):
    if os.path.exists('ynab_csvs/.last_%s_date' % ftype):
        last_date = open('ynab_csvs/.last_%s_date' % ftype, "r").read()
        last_date = pd.to_datetime(last_date, format='%Y-%m-%d')
        df = df[df['Date'] > last_date]
    if not df.empty:
        last_date = df['Date'].iloc[-1]
        open("ynab_csvs/.last_%s_date" % ftype, "w").write(str(last_date)[:10])
    return df[~df.isin(prev_df)].dropna()

cartolas = os.listdir('cartolas')
ynab_csvs = os.listdir('ynab_csvs')

# Cuenta Corrriente
if 'cartola.xls' in cartolas:
    cc_df = read_cc_xls('cartolas/cartola.xls')
    if 'cartola.csv' in ynab_csvs:
        cc_df = remove_duplicates(cc_df, pd.read_csv('ynab_csvs/cartola.csv').fillna(''))
    cc_df.to_csv('ynab_csvs/cartola.csv', index=False)

# Tarjeta de Credito
if 'Saldo_y_Mov_No_Facturado.xls' in cartolas:
    tc_df = read_tc_xls('cartolas/Saldo_y_Mov_No_Facturado.xls')
    if 'Saldo_y_Mov_No_Facturado.csv' in ynab_csvs:
        tc_df = remove_duplicates(tc_df, pd.read_csv('ynab_csvs/Saldo_y_Mov_No_Facturado.csv').fillna(''), 'tc')
    tc_df.to_csv('ynab_csvs/Saldo_y_Mov_No_Facturado.csv', index=False)

print('Done!')