import openpyxl as xl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import NamedStyle
from datetime import datetime
import pyperclip
import pandas as pd
import estilo


# Gera a data de hoje no formato dia-mês
data_hoje = datetime.now().strftime("%d-%m")
arquivo_mb25 = f"C:/Users/prr8ca/Desktop/Projetos/SapAuto/tabelas/mb25/mb25 - {data_hoje}.xlsx"
arquivo_me3m = f"C:/Users/prr8ca/Desktop/Projetos/SapAuto/tabelas/me3m/me3m - {data_hoje}.xlsx"
arquivo_z22k032 = f"C:/Users/prr8ca/Desktop/Projetos/SapAuto/tabelas/z22k032/z22k032 - {data_hoje}.xlsx"
arquivo_z22m051 = f"C:/Users/prr8ca/Desktop/Projetos/SapAuto/tabelas/z22m051/z22m051 - {data_hoje}.xlsx"
arquivo_blocok_depois = "C:/Users/prr8ca/Desktop/Projetos/SapAuto/tabelas/Bloco K Depois.xlsx"

    



def excluir_linhas(tabela, coluna, valores): #Recebe Coluna e Valor, exclui as linhas que o valor esta na coluna
    linhas_a_remover = []
    for row in range(2, tabela.max_row + 1):  # Começa da linha 2 para evitar o cabeçalho
        codigo_valor = tabela[f'{coluna}{row}'].value
        if codigo_valor in valores:
            linhas_a_remover.append(row)
    for row in reversed(linhas_a_remover):
        tabela.delete_rows(row, 1)


def transformar_coluna_em_texto(tabela, coluna):
    for row in range(2, tabela.max_row + 1):  
        cell = tabela[f'{coluna}{row}']
        cell.number_format = '@'


def alterar_valor(tabela, coluna, valor_antes, valor_depois):
    for row in range(2, tabela.max_row + 1):  
        cell = tabela[f'{coluna}{row}']
        if cell.value == valor_antes:
            cell.value = valor_depois


#========================= Bloco K =======================#

def copiar_tabela_entre_arquivos(arquivo_origem, nome_planilha_origem, arquivo_destino, nome_planilha_destino=None):
    df = pd.read_excel(arquivo_origem, sheet_name=nome_planilha_origem)
    with pd.ExcelWriter(arquivo_destino, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=nome_planilha_destino, index=False)


def copiar_lista_material():
    # Carrega a planilha de origem e a planilha de destino
    df_origem = pd.read_excel(arquivo_blocok_depois, sheet_name='MB25')
    df_destino = pd.read_excel(arquivo_blocok_depois, sheet_name='Planilha1')

    # Pega a coluna desejada da planilha de origem (substitua 'NomeColuna' pelo nome da coluna)
    coluna_origem = df_origem[['Material']]

    # Adiciona a coluna na planilha de destino
    df_destino = df_destino.set_index('Material').combine_first(coluna_origem.set_index('Material')).reset_index()

    df_destino.drop_duplicates(subset=['Material'], inplace=True)


    # Salva as alterações de volta no mesmo arquivo Excel
    with pd.ExcelWriter(arquivo_blocok_depois, mode='a', if_sheet_exists='replace') as writer:
        df_origem.to_excel(writer, sheet_name='MB25', index=False)
        df_destino.to_excel(writer, sheet_name='Planilha1', index=False)

def pegar_partnumbers():
    wb = xl.load_workbook(arquivo_blocok_depois)
    ws = wb['Planilha1']
    parnumbers = []

    for cell in ws['A'][1:]:
        if cell.value is not None:
            parnumbers.append(cell.value)
    
    parnumbers_str = "\r\n".join(parnumbers)
    pyperclip.copy(parnumbers_str)
    
    return parnumbers_str


#========================= MB25 =======================#

def alterar_tabela_mb25():

    wb_mb25 = xl.load_workbook(arquivo_mb25)
    tabela_mb25 = wb_mb25['Sheet1']
    header_mb25 = {cell.value: cell.column for cell in tabela_mb25[1]}

    deposito_coluna_idx_mb25 = header_mb25['Depósito']
    deposito_coluna_mb25 = xl.utils.get_column_letter(deposito_coluna_idx_mb25)
    transformar_coluna_em_texto(tabela_mb25, deposito_coluna_mb25)
    excluir_linhas(tabela_mb25, deposito_coluna_mb25, ['9010', '9032', ''])

    UM_coluna_idx_mb25 = header_mb25['UM básica']
    UM_coluna_mb25 = xl.utils.get_column_letter(UM_coluna_idx_mb25)
    excluir_linhas(tabela_mb25, UM_coluna_mb25, ['KG'])

    material_coluna_idx_mb25 = header_mb25['Material']
    material_coluna_mb25 = xl.utils.get_column_letter(material_coluna_idx_mb25)
    transformar_coluna_em_texto(tabela_mb25, material_coluna_mb25)

    wb_mb25.save(arquivo_mb25)

#============================= Bloco K ================================#

    copiar_tabela_entre_arquivos(arquivo_mb25, 'Sheet1', arquivo_blocok_depois,  'MB25')
    copiar_lista_material()
    estilo.formatar_blocok(arquivo_blocok_depois)
    pegar_partnumbers()
    


def alterar_tabelas():

    #========================== me3m ==================================#
    wb_me3m = xl.load_workbook(arquivo_me3m)
    tabela_me3m = wb_me3m['Sheet1']
    header_me3m = {cell.value: cell.column for cell in tabela_me3m[1]}

    codigo_coluna_idx_me3m = header_me3m['Código de eliminação']
    codigo_coluna_me3m = xl.utils.get_column_letter(codigo_coluna_idx_me3m)
    excluir_linhas(tabela_me3m, codigo_coluna_me3m, ['S'])

    materiais_coluna_idx_me3m = header_me3m['Material']
    materiais_coluna_me3m = xl.utils.get_column_letter(materiais_coluna_idx_me3m)
    transformar_coluna_em_texto(tabela_me3m, materiais_coluna_me3m)

    wb_me3m.save(arquivo_me3m)


    #========================== z22k032 ==================================#
    wb_z22k032 = xl.load_workbook(arquivo_z22k032)
    tabela_z22k032 = wb_z22k032['Sheet1']
    header_z22k032 = {cell.value: cell.column for cell in tabela_z22k032[1]}

    materiais_coluna_idx_z22k032 = header_z22k032['Material']
    materiais_coluna_z22k032 = xl.utils.get_column_letter(materiais_coluna_idx_z22k032)
    transformar_coluna_em_texto(tabela_z22k032, materiais_coluna_z22k032)
    
    wb_z22k032.save(arquivo_z22k032)

    #========================== z22m051 ==================================#
    wb_z22m051 = xl.load_workbook(arquivo_z22m051)
    tabela_z22m051 = wb_z22m051['Sheet1']
    header_z22m051 = {cell.value: cell.column for cell in tabela_z22m051[1]}

    UM_coluna_idx_z22m051 = header_z22m051['UM']
    UM_coluna_z22m051 = xl.utils.get_column_letter(UM_coluna_idx_z22m051)
    alterar_valor(tabela_z22m051, UM_coluna_z22m051, 'ST', 'PC')

    materiais_coluna_idx_z22m051 = header_z22m051['Material']
    materiais_coluna_z22m051 = xl.utils.get_column_letter(materiais_coluna_idx_z22m051)
    transformar_coluna_em_texto(tabela_z22m051, materiais_coluna_z22m051)

    wb_z22m051.save(arquivo_z22m051)
