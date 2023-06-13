import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import os


        # PROJETO PARA ATUALIZAÇÃO DA CARTEIRA DOS VENDEDORES E REPRESENTANTES #



def atualizar_relatorioM():

    print('Iniciando...')

    # IMPORTAÇÃO DE TABELAS
    tab_cli = pd.read_excel('BaseClientes.xlsx')
    tab_car = pd.read_excel('BaseCarteira.xlsx')
    tab_stt = pd.read_excel('StatusCM.xlsx')


    hoje = datetime.now()

    # SELEÇÃO DE CAMPOS DA TABELA DE CADASTRO DE CLIENTE
    cad_cli = tab_cli[['ID', 'Nome do Cliente', 'Data Iní. Cad.']]
    cad_cli.loc[cad_cli['Data Iní. Cad.'] == '00/00/0000', 'Data Iní. Cad.'] = ''
    car_rep = tab_car[['ID', 'Cod']]


    # SELEÇÃO DE CAMPOS DO RELATÓRIO ONDE CONTEM AS INFOMAÇÕES DE ÚLTIMA COMPRA
    status_cliente = tab_stt[['ID', 'Data Ultima Compra']]
    status_cliente.loc[status_cliente['Data Ultima Compra'] == '00/00/0000', 'Data Ultima Compra'] = ''
    status_cliente['Data Ultima Compra'] = status_cliente['Data Ultima Compra'].astype('datetime64')


    # UNIÃO DAS TABELAS
    status_car = pd.merge(cad_cli, status_cliente, on='ID', how='left')
    cart_repr = pd.merge(status_car, car_rep, left_on='ID', right_on='ID', how='right')

    cart_repr['StatusM'] = ''
    cart_repr['Operação'] = ''

    # ORGANIZANDO A ORDEM DAS COLUNAS
    cart_repr = cart_repr[['Cod','ID', 'Nome do Cliente', 'Data Iní. Cad.', 'Data Ultima Compra']]
    # FORMATAR CNPJ/CPF
    # cart_repr['CNPJ/CPF'] = cart_repr['CNPJ/CPF'].astype('str')


    df = cart_repr

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')


    #   LÓGICA PARA A COMPARAÇÃO ENTRE AS COLUNAS PARA POVOAR A COLUNA 'Status'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] < hoje - timedelta(days=60)), 'StatusM'] = 'Suspect'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] > hoje - timedelta(days=60)), 'StatusM'] = 'Propect'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] > hoje - timedelta(days=30)), 'StatusM'] = 'Novo'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Ativo'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Atenção'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=60)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Inativo'

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].dt.strftime('%Y-%m-%d')
    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].dt.strftime('%Y-%m-%d')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].astype('datetime64')
    df.rename(columns={'Data Ultima Compra':'Ultima Compra fM'},inplace=True)


    #   EXPORTANDO PARA UM ARQUIVO EXCEL
    df.to_excel('carteiraMa.xlsx')

    carteiraMa = df

    print('Matriz Pronta')

atualizar_relatorioM()


def atualizar_relatorioF2():

    # IMPORTAÇÃO DE TABELAS
    tab_cli = pd.read_excel('BaseClientes.xlsx')
    tab_car = pd.read_excel('BaseCarteira.xlsx')
    tab_stt = pd.read_excel('StatusCF2.xlsx')


    hoje = datetime.now()

    # SELEÇÃO DE CAMPOS DA TABELA DE CADASTRO DE CLIENTE
    cad_cli = tab_cli[['ID', 'Nome do Cliente', 'Data Iní. Cad.']]
    cad_cli.loc[cad_cli['Data Iní. Cad.'] == '00/00/0000', 'Data Iní. Cad.'] = ''
    car_rep = tab_car[['ID', 'Cod']]


    # SELEÇÃO DE CAMPOS DO RELATÓRIO ONDE CONTEM AS INFOMAÇÕES DE ÚLTIMA COMPRA
    status_cliente = tab_stt[['ID', 'Data Ultima Compra']]
    status_cliente.loc[status_cliente['Data Ultima Compra'] == '00/00/0000', 'Data Ultima Compra'] = ''
    status_cliente['Data Ultima Compra'] = status_cliente['Data Ultima Compra'].astype('datetime64')


    # UNIÃO DAS TABELAS
    status_car = pd.merge(cad_cli, status_cliente, on='ID', how='left')
    cart_repr = pd.merge(status_car, car_rep, left_on='ID', right_on='ID', how='right')

    cart_repr['StatusM'] = ''
    cart_repr['Operação'] = ''

    # ORGANIZANDO A ORDEM DAS COLUNAS
    cart_repr = cart_repr[['Cod','ID', 'Nome do Cliente', 'Data Iní. Cad.', 'Data Ultima Compra']]
    # FORMATAR CNPJ/CPF
    # cart_repr['CNPJ/CPF'] = cart_repr['CNPJ/CPF'].astype('str')


    df = cart_repr

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')


    #   LÓGICA PARA A COMPARAÇÃO ENTRE AS COLUNAS PARA POVOAR A COLUNA 'Status'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] < hoje - timedelta(days=60)), 'StatusM'] = 'Suspect'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] > hoje - timedelta(days=60)), 'StatusM'] = 'Propect'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] > hoje - timedelta(days=30)), 'StatusM'] = 'Novo'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Ativo'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Atenção'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=60)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Inativo'

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].dt.strftime('%Y-%m-%d')
    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].dt.strftime('%Y-%m-%d')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].astype('datetime64')
    df.rename(columns={'Data Ultima Compra':'Ultima Compra fM'},inplace=True)


    #   EXPORTANDO PARA UM ARQUIVO EXCEL
    df.to_excel('carteiraF2.xlsx')

    carteiraF2 = df

atualizar_relatorioF2()


def atualizar_relatorioF3():

    # IMPORTAÇÃO DE TABELAS
    tab_cli = pd.read_excel('BaseClientes.xlsx')
    tab_car = pd.read_excel('BaseCarteira.xlsx')
    tab_stt = pd.read_excel('StatusCF3.xlsx')


    hoje = datetime.now()

    # SELEÇÃO DE CAMPOS DA TABELA DE CADASTRO DE CLIENTE
    cad_cli = tab_cli[['ID', 'Nome do Cliente', 'Data Iní. Cad.']]
    cad_cli.loc[cad_cli['Data Iní. Cad.'] == '00/00/0000', 'Data Iní. Cad.'] = ''
    car_rep = tab_car[['ID', 'Cod']]


    # SELEÇÃO DE CAMPOS DO RELATÓRIO ONDE CONTEM AS INFOMAÇÕES DE ÚLTIMA COMPRA
    status_cliente = tab_stt[['ID', 'Data Ultima Compra']]
    status_cliente.loc[status_cliente['Data Ultima Compra'] == '00/00/0000', 'Data Ultima Compra'] = ''
    status_cliente['Data Ultima Compra'] = status_cliente['Data Ultima Compra'].astype('datetime64')


    # UNIÃO DAS TABELAS
    status_car = pd.merge(cad_cli, status_cliente, on='ID', how='left')
    cart_repr = pd.merge(status_car, car_rep, left_on='ID', right_on='ID', how='right')

    cart_repr['StatusM'] = ''
    cart_repr['Operação'] = ''

    # ORGANIZANDO A ORDEM DAS COLUNAS
    cart_repr = cart_repr[['Cod','ID', 'Nome do Cliente', 'Data Iní. Cad.', 'Data Ultima Compra']]
    # FORMATAR CNPJ/CPF
    # cart_repr['CNPJ/CPF'] = cart_repr['CNPJ/CPF'].astype('str')


    df = cart_repr

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')


    #   LÓGICA PARA A COMPARAÇÃO ENTRE AS COLUNAS PARA POVOAR A COLUNA 'Status'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] < hoje - timedelta(days=60)), 'StatusM'] = 'Suspect'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] > hoje - timedelta(days=60)), 'StatusM'] = 'Propect'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] > hoje - timedelta(days=30)), 'StatusM'] = 'Novo'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Ativo'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Atenção'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=60)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Inativo'

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].dt.strftime('%Y-%m-%d')
    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].dt.strftime('%Y-%m-%d')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].astype('datetime64')
    df.rename(columns={'Data Ultima Compra':'Ultima Compra fM'},inplace=True)


    #   EXPORTANDO PARA UM ARQUIVO EXCEL
    df.to_excel('carteiraF3.xlsx')

    carteiraF3 = df

atualizar_relatorioF3()


def atualizar_relatorioF4():

    # IMPORTAÇÃO DE TABELAS
    tab_cli = pd.read_excel('BaseClientes.xlsx')
    tab_car = pd.read_excel('BaseCarteira.xlsx')
    tab_stt = pd.read_excel('StatusCF4.xlsx')


    hoje = datetime.now()

    # SELEÇÃO DE CAMPOS DA TABELA DE CADASTRO DE CLIENTE
    cad_cli = tab_cli[['ID', 'Nome do Cliente', 'Data Iní. Cad.']]
    cad_cli.loc[cad_cli['Data Iní. Cad.'] == '00/00/0000', 'Data Iní. Cad.'] = ''
    car_rep = tab_car[['ID', 'Cod']]


    # SELEÇÃO DE CAMPOS DO RELATÓRIO ONDE CONTEM AS INFOMAÇÕES DE ÚLTIMA COMPRA
    status_cliente = tab_stt[['ID', 'Data Ultima Compra']]
    status_cliente.loc[status_cliente['Data Ultima Compra'] == '00/00/0000', 'Data Ultima Compra'] = ''
    status_cliente['Data Ultima Compra'] = status_cliente['Data Ultima Compra'].astype('datetime64')


    # UNIÃO DAS TABELAS
    status_car = pd.merge(cad_cli, status_cliente, on='ID', how='left')
    cart_repr = pd.merge(status_car, car_rep, left_on='ID', right_on='ID', how='right')

    cart_repr['StatusM'] = ''
    cart_repr['Operação'] = ''

    # ORGANIZANDO A ORDEM DAS COLUNAS
    cart_repr = cart_repr[['Cod','ID', 'Nome do Cliente', 'Data Iní. Cad.', 'Data Ultima Compra']]
    # FORMATAR CNPJ/CPF
    # cart_repr['CNPJ/CPF'] = cart_repr['CNPJ/CPF'].astype('str')


    df = cart_repr

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')


    #   LÓGICA PARA A COMPARAÇÃO ENTRE AS COLUNAS PARA POVOAR A COLUNA 'Status'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] < hoje - timedelta(days=60)), 'StatusM'] = 'Suspect'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] > hoje - timedelta(days=60)), 'StatusM'] = 'Propect'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] > hoje - timedelta(days=30)), 'StatusM'] = 'Novo'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Ativo'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Atenção'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=60)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Inativo'

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].dt.strftime('%Y-%m-%d')
    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].dt.strftime('%Y-%m-%d')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].astype('datetime64')
    df.rename(columns={'Data Ultima Compra':'Ultima Compra fM'},inplace=True)


    #   EXPORTANDO PARA UM ARQUIVO EXCEL
    df.to_excel('carteiraF4.xlsx')

    carteiraF4 = df

atualizar_relatorioF4()


def atualizar_relatorioF5():

    # IMPORTAÇÃO DE TABELAS
    tab_cli = pd.read_excel('BaseClientes.xlsx')
    tab_car = pd.read_excel('BaseCarteira.xlsx')
    tab_stt = pd.read_excel('StatusCF2.xlsx')


    hoje = datetime.now()

    # SELEÇÃO DE CAMPOS DA TABELA DE CADASTRO DE CLIENTE
    cad_cli = tab_cli[['ID', 'Nome do Cliente', 'Data Iní. Cad.']]
    cad_cli.loc[cad_cli['Data Iní. Cad.'] == '00/00/0000', 'Data Iní. Cad.'] = ''
    car_rep = tab_car[['ID', 'Cod']]


    # SELEÇÃO DE CAMPOS DO RELATÓRIO ONDE CONTEM AS INFOMAÇÕES DE ÚLTIMA COMPRA
    status_cliente = tab_stt[['ID', 'Data Ultima Compra']]
    status_cliente.loc[status_cliente['Data Ultima Compra'] == '00/00/0000', 'Data Ultima Compra'] = ''
    status_cliente['Data Ultima Compra'] = status_cliente['Data Ultima Compra'].astype('datetime64')


    # UNIÃO DAS TABELAS
    status_car = pd.merge(cad_cli, status_cliente, on='ID', how='left')
    cart_repr = pd.merge(status_car, car_rep, left_on='ID', right_on='ID', how='right')

    cart_repr['StatusM'] = ''
    cart_repr['Operação'] = ''

    # ORGANIZANDO A ORDEM DAS COLUNAS
    cart_repr = cart_repr[['Cod','ID', 'Nome do Cliente', 'Data Iní. Cad.', 'Data Ultima Compra']]
    # FORMATAR CNPJ/CPF
    # cart_repr['CNPJ/CPF'] = cart_repr['CNPJ/CPF'].astype('str')


    df = cart_repr

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')


    #   LÓGICA PARA A COMPARAÇÃO ENTRE AS COLUNAS PARA POVOAR A COLUNA 'Status'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] < hoje - timedelta(days=60)), 'StatusM'] = 'Suspect'
    df.loc[(df['Data Ultima Compra'] !='0') & (df['Data Iní. Cad.'] > hoje - timedelta(days=60)), 'StatusM'] = 'Propect'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] > hoje - timedelta(days=30)), 'StatusM'] = 'Novo'
    df.loc[(df['Data Ultima Compra'] >= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Ativo'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=30)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Atenção'
    df.loc[(df['Data Ultima Compra'] <= hoje - timedelta(days=60)) & (df['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'StatusM'] = 'Inativo'

    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].dt.strftime('%Y-%m-%d')
    df['Data Iní. Cad.'] = df['Data Iní. Cad.'].astype('datetime64')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].dt.strftime('%Y-%m-%d')
    df['Data Ultima Compra'] = df['Data Ultima Compra'].astype('datetime64')
    df.rename(columns={'Data Ultima Compra':'Ultima Compra fM'},inplace=True)


    #   EXPORTANDO PARA UM ARQUIVO EXCEL
    df.to_excel('carteiraF2.xlsx')

    carteiraF5 = df

atualizar_relatorioF5()


def Atualizar_Clientes():

    hoje = datetime.now()

    clientesM = pd.read_excel('carteiraMa.xlsx')
    clientesF2 = pd.read_excel('carteiraF2.xlsx')
    clientesF3 = pd.read_excel('carteiraF3.xlsx')
    clientesF4 = pd.read_excel('carteiraF4.xlsx')
    clientesF5 = pd.read_excel('carteiraF5.xlsx')


    carteira = clientesM

    #   print(carteira)

    carteirm = pd.merge(clientesM, clientesF2, on='ID', how='right')
    carteirf2 = pd.merge(carteirm, clientesF3, on='ID', how='right')
    carteirf3 = pd.merge(carteirf2, clientesF4, on='ID', how='right')
    carteira = pd.merge(carteirf3, clientesF5, on='ID', how='right')

    #   carteira['Status'] = ''
    carteira['Data ultima Compra'] = '01/01/2000'


    df2 = carteira
    df2 = df2[['Cod', 'ID','Nome do Cliente', 'Data Iní. Cad.', 'Ultima Compra fM','StatusM', 'Ultima Compra f2','Status2', 'Ultima Compra f3','Status3', 'Ultima Compra f4','Status4', 'Ultima Compra f5','Status5', 'Status', 'Data ultima Compra']]
    df2 = df2.drop_duplicates(subset='ID',ignore_index=True)

    df2['Data ultima Compra'] = df2['Data ultima Compra'].astype('datetime64')

    df2.loc[(df2['Ultima Compra fM'] == '') & (df2['Ultima Compra f2'] == '') & (df2['Ultima Compra f3'] == '') & (df2['Ultima Compra f4'] == '') & (df2['Ultima Compra f5'] == ''), 'Data ultima Compra'] = df2['Data ultima Compra']

    df2.loc[(df2['Ultima Compra fM'] >= df2['Ultima Compra f2']), 'Data ultima Compra'] = df2['Ultima Compra fM']
    df2.loc[(df2['Ultima Compra fM'] >= df2['Ultima Compra f3']), 'Data ultima Compra'] = df2['Ultima Compra fM']
    df2.loc[(df2['Ultima Compra fM'] >= df2['Ultima Compra f4']), 'Data ultima Compra'] = df2['Ultima Compra fM']
    df2.loc[(df2['Ultima Compra fM'] >= df2['Ultima Compra f5']), 'Data ultima Compra'] = df2['Ultima Compra fM']
    df2.loc[(df2['Ultima Compra fM'] >= df2['Data ultima Compra']), 'Data ultima Compra'] = df2['Ultima Compra fM']

    df2.loc[(df2['Ultima Compra f2'] >= df2['Ultima Compra fM']), 'Data ultima Compra'] = df2['Ultima Compra f2']
    df2.loc[(df2['Ultima Compra f2'] >= df2['Ultima Compra f3']), 'Data ultima Compra'] = df2['Ultima Compra f2']
    df2.loc[(df2['Ultima Compra f2'] >= df2['Ultima Compra f4']), 'Data ultima Compra'] = df2['Ultima Compra f2']
    df2.loc[(df2['Ultima Compra f2'] >= df2['Ultima Compra f5']), 'Data ultima Compra'] = df2['Ultima Compra f2']
    df2.loc[(df2['Ultima Compra f2'] >= df2['Data ultima Compra']), 'Data ultima Compra'] = df2['Ultima Compra f2']

    df2.loc[(df2['Ultima Compra f3'] >= df2['Ultima Compra fM']), 'Data ultima Compra'] = df2['Ultima Compra f3']
    df2.loc[(df2['Ultima Compra f3'] >= df2['Ultima Compra f2']), 'Data ultima Compra'] = df2['Ultima Compra f3']
    df2.loc[(df2['Ultima Compra f3'] >= df2['Ultima Compra f4']), 'Data ultima Compra'] = df2['Ultima Compra f3']
    df2.loc[(df2['Ultima Compra f3'] >= df2['Ultima Compra f5']), 'Data ultima Compra'] = df2['Ultima Compra f3']
    df2.loc[(df2['Ultima Compra f3'] >= df2['Data ultima Compra']), 'Data ultima Compra'] = df2['Ultima Compra f3']

    df2.loc[(df2['Ultima Compra f4'] >= df2['Ultima Compra fM']), 'Data ultima Compra'] = df2['Ultima Compra f4']
    df2.loc[(df2['Ultima Compra f4'] >= df2['Ultima Compra f2']), 'Data ultima Compra'] = df2['Ultima Compra f4']
    df2.loc[(df2['Ultima Compra f4'] >= df2['Ultima Compra f3']), 'Data ultima Compra'] = df2['Ultima Compra f4']
    df2.loc[(df2['Ultima Compra f4'] >= df2['Ultima Compra f5']), 'Data ultima Compra'] = df2['Ultima Compra f4']
    df2.loc[(df2['Ultima Compra f4'] >= df2['Data ultima Compra']), 'Data ultima Compra'] = df2['Ultima Compra f4']

    df2.loc[(df2['Ultima Compra f5'] >= df2['Ultima Compra fM']), 'Data ultima Compra'] = df2['Ultima Compra f5']
    df2.loc[(df2['Ultima Compra f5'] >= df2['Ultima Compra f2']), 'Data ultima Compra'] = df2['Ultima Compra f5']
    df2.loc[(df2['Ultima Compra f5'] >= df2['Ultima Compra f3']), 'Data ultima Compra'] = df2['Ultima Compra f5']
    df2.loc[(df2['Ultima Compra f5'] >= df2['Ultima Compra f4']), 'Data ultima Compra'] = df2['Ultima Compra f5']
    df2.loc[(df2['Ultima Compra f5'] >= df2['Data ultima Compra']), 'Data ultima Compra'] = df2['Ultima Compra f5']

    #   TRANSFORMANDO A DATA DE SUSPECS E PROSPECTS PARA NULO
    df2.loc[(df2['Data ultima Compra'] == "01/01/2000"), 'Data ultima Compra'] = ''

    #   LÓGICA PARA A COMPARAÇÃO ENTRE AS COLUNAS PARA POVOAR A COLUNA 'Status'
    df2.loc[(df2['Data ultima Compra'] != '0') & (df2['Data Iní. Cad.'] < hoje - timedelta(days=60)), 'Status'] = 'Suspect'
    df2.loc[(df2['Data ultima Compra'] != '0') & (df2['Data Iní. Cad.'] > hoje - timedelta(days=60)), 'Status'] = 'Propect'
    df2.loc[(df2['Data ultima Compra'] >= hoje - timedelta(days=30)) & (df2['Data Iní. Cad.'] > hoje - timedelta(days=30)), 'Status'] = 'Novo'
    df2.loc[(df2['Data ultima Compra'] >= hoje - timedelta(days=30)) & (df2['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'Status'] = 'Ativo'
    df2.loc[(df2['Data ultima Compra'] <= hoje - timedelta(days=30)) & (df2['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'Status'] = 'Atenção'
    df2.loc[(df2['Data ultima Compra'] <= hoje - timedelta(days=60)) & (df2['Data Iní. Cad.'] <= hoje - timedelta(days=30)), 'Status'] = 'Inativo'
 

    df = df2[['Cod', 'ID','Nome do Cliente','Data Iní. Cad.', 'Status', 'Data ultima Compra']]

    df.to_excel('carteiraAtualizada.xlsx')
    print('Arquivos gerados!')

    os.remove('carteiraMa.xlsx')
    os.remove('carteiraF2.xlsx')
    os.remove('carteiraF3.xlsx')
    os.remove('carteiraF4.xlsx')
    os.remove('carteiraF5.xlsx')


    print("Operação Finalizado!")


Atualizar_Clientes()