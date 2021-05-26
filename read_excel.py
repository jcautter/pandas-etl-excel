import numpy as np
import pandas as pd

sheets = ['Norte', 'Sul']

def sheet_etl(path:str, sheet:str):
    df = pd.read_excel(path, sheet_name=sheet, skiprows=3, header=1, usecols='D:Q')
    df.dropna(axis=0, how='all', inplace=True)
    df.drop(columns=['Total'], inplace=True)
    df.rename(columns={'Unnamed: 3':'Abertura2'}, inplace=True)
    df['Abertura1'] = np.nan
    df.iloc[1:5, df.columns.get_loc('Abertura1')] = 'Receita'
    df.iloc[6:9, df.columns.get_loc('Abertura1')] = 'Custo'
    df.iloc[11:14, df.columns.get_loc('Abertura1')] = 'Opex'
    df.iloc[15:16, df.columns.get_loc('Abertura1')] = 'Bad Debt'
    df.dropna(subset = ['Abertura1'], inplace=True)
    df.set_index(['Abertura1', 'Abertura2'], inplace=True)
    df = df.stack().reset_index().rename(columns={'level_2':'Data', 0:'Valor'})
    df.insert(2, 'Região', sheet)
    return df

def report_etl(path:str, sheets:list):
    df = pd.DataFrame()
    for sheet in sheets:
        df = pd.concat([df, sheet_etl(path, sheet)])
    return df

a = report_etl('Report.xlsx', sheets)
