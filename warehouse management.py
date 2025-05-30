import pandas as pd
from pathlib import Path

df = pd.DataFrame()

BASE_PATH = Path.home() / 'Desktop' / 'Python' / 'Exercicios Teste' / 'Files'
FILE_NAME = 'dataframe_dinamico.xlsx'
FILE_PATH = BASE_PATH / FILE_NAME

def criar_dataframe():
    global df
    
    try:
        num_colunas = int(input("Indicar o nº de colunas a adicionar ao stock do armazém: \n"))
        data = {}
        num_linhas = None
        
        for i in range(num_colunas):
            nome_coluna = input(f"Qual nome quer dar à coluna {i + 1}: \n")
            
            if nome_coluna in data:
                print(f"Nome da coluna - {nome_coluna} já existe no stock de armazém")
                return None
            
            dados = input(f"Introduza os dados a registar na coluna - {nome_coluna} separados por vírgula ','\n").split(',')
            
        if num_linhas == None:
            num_linhas = len(dados)
        elif len(dados) != num_linhas:
            print(f"Dados inválidos, nº de dados nas colunas deve ser igual a {num_linhas} ")
            return None
    
        data[nome_coluna] = dados
        print("Colunas e respetivos dados adicionados ao stock com sucesso\n")
        print()
    
        df = pd.DataFrame(data)
        return df
    
    except ValueError:
        print("Erro! O nº de colunas deve ser inteiro")
        
        
def consultar_dataframe():
    global df
    
    if df.empty:
        print("O dataframe está vazio\n")
        return
    
    print("Dados registados: \n")
    print(df)
    print()
    
    
def adicionar_coluna():
    global df
    if df.empty:
        print("O Data frame está vazio")
        return
    nome_coluna = input("Indique o nome da coluna a adicionar: \n")
    
    if nome_coluna in df.columns:
        print(f"A coluna {nome_coluna} já existe")
        return None
    
    dados = input(f"Introduza os dados a registar na coluna - {nome_coluna} separados por vírgula (',')\n").split(',')
    
    num_linhas_df = len(df.index)
    
    if len(dados) != num_linhas_df:
        print(f"Dados inválidos, nº de dados nas colunas deve ser igual a {num_linhas_df} ")
        return None
    
    df[nome_coluna] = dados
    print(f"Coluna '{nome_coluna}' adicionada com sucesso\n")
    print()
    
def guardar_dataframe():
    global df
    if df.empty:
        print("O Data frame está vazio")
        return
    
    try:
        df.to_excel(FILE_PATH, sheet_name= 'Stock')
        print(f"Stock do armazém guardado no ficheiro {FILE_PATH}\n")
    except Exception as e:
        print(f"Erro do tipo {e} ao gerar o ficheiro")
        
def importar_dataframe():
    global df
    
    try:
        df = pd.read_excel(FILE_PATH, sheet_name= 'Stock')
        print(f"Stock importado com sucesso do ficheiro {FILE_PATH} \n")
    except FileNotFoundError:
        print(f"Erro! Ficheiro não encontrado\n")
    except Exception as e:
        print(f"Erro do tipo {e} ao importar o ficheiro\n")
        
def eliminar_coluna():
    global df
    if df.empty:
        print("O stock está vazio")
        return
    consultar_dataframe()
    
    nome_coluna = input("Introduza o nome da coluna que pretende eliminar: \n")
    
    if nome_coluna not in df.columns:
        print(f"Erro! A coluna {nome_coluna} não está registada\n")
        return None
    
    df = df.drop(columns= [nome_coluna])
    print(f"A coluna {nome_coluna} foi eliminada")
    print()
    

def renomear_coluna():
    global df
    if df.empty:
        print("O stock está vazio")
        return
    consultar_dataframe()
    
    nome_coluna = input("Introduza o nome da coluna que pretende mudar o nome: \n")
    
    if nome_coluna not in df.columns:
        print(f"Erro! A coluna {nome_coluna} não está registada\n")
        return None
    
    while True:
        novo_nome_coluna = input(f"Introduza o novo nome para a coluna {nome_coluna}: ").strip() #strip() para eliminar espaços em branco
        if not novo_nome_coluna:
            print("Erro! O nome da coluna nova não pode ser vazio")
        elif novo_nome_coluna in df.columns:
            print(f"Este nome de coluna - {novo_nome_coluna} já existe")
        else:
            break
        
    df = df.rename(columns={nome_coluna : novo_nome_coluna})
    print(f"A coluna '{nome_coluna}' foi renomeada com sucesso para: {novo_nome_coluna}")
    print()

def filtrar_valores():
    global df
    if df.empty:
        print("O stock está vazio")
        return
    consultar_dataframe()
    
    nome_coluna = input("Introduza o nome da coluna onde pretende filtrar dados.\n")
    
    if nome_coluna not in df.columns:
        print(f"A coluna {nome_coluna} não está registada\n")
        return None
    
    criterio = input(f"Indique o criterio a filtrar no {nome_coluna}\n")
    filtro = df[df[nome_coluna == criterio]]
    
    print(f"Dados na coluna - {nome_coluna} que correspondem a - {criterio}: \n")
    print(filtro[nome_coluna])
    print()
    
            
    
    

while True:
    print("<<<< Bem-vindo ao sistema de gestão de stock >>>>\n")
    print("1 Criar Stock ")
    print("2 Consultar Stock")
    print("3 Adicionar coluna")
    print("4 Eliminar coluna")
    print("5 Renomear coluna")
    print("6 Filtrar dados no stock")
    print("7 Guardar stock em excel")
    print("8 Importar stock de excel")
    print("9 Sair \n")
    
    escolha = input("Indique a opção a executar: \n")
    
    if escolha == "1":
        criar_dataframe()
    elif escolha == "2":
        consultar_dataframe()
    elif escolha == "3":
        adicionar_coluna()
    elif escolha == "4":
        eliminar_coluna()
    elif escolha == "5":
        renomear_coluna()
    elif escolha == "6":
        filtrar_valores()
    elif escolha == "7":
        guardar_dataframe()
    elif escolha == "8":
        importar_dataframe()
    elif escolha == "9":
        print("Obrigado \n")
        break
    else:
        print("Erro! Opção Inválida! \n")                
        
    
    

