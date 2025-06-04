import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from datetime import datetime

df = pd.DataFrame()

BASE_PATH = Path.home() / 'Desktop' / 'Python' / 'Exercicios Teste' / 'Files'
FILE_NAME = 'dataframe_dinamico.xlsx'
FILE_PATH = BASE_PATH / FILE_NAME

def criar_dataframe():
    global df
    
    try:
        num_colunas = int(input("Indicar o nº de colunas a adicionar ao stock do armazém: \n"))
        data = {}
        tipo_dados = {}
        num_linhas = None
       
       
        
        for i in range(num_colunas):
            while True:
                nome_coluna = input(f"Qual nome quer dar à coluna {i + 1}: \n")
                
                if nome_coluna:
                    if nome_coluna not in data:
                        break
                    print(f"A coluna '{nome_coluna}' já está registada no dicionário! ")
                else:
                    print("O nome da coluna não pode ser vazio")
                    
            print("Escolha o tipo de dados a registar na coluna: \n")
            print("1 Texto\n")
            print("2 Nºs inteiros\n")
            print("3 Nºs decimais (reais)\n")
            print("4 Datas (formato: DD/MM/AAAA)\n")
            tipo_escolhido = input("Indique o tipo de dado a registar: \n")
            
            if tipo_escolhido == "1":
                tipo_atual = str
            elif tipo_escolhido == "2":
                tipo_atual = int
            elif tipo_escolhido == "3":
                tipo_atual = float
            elif tipo_escolhido == "4":
                tipo_atual = "date"
            else:
                print("Opção inválida! O tipo de dado a registar será considerado texto!\n")
                tipo_atual = str
         
            
            # if nome_coluna in data:
            #     print(f"Nome da coluna - {nome_coluna} já existe no stock de armazém")
            #     return None
            
            dados = input(f"Introduza os dados a registar na coluna - {nome_coluna} separados por vírgula ','\n").split(',')
            
            if num_linhas is None:
                num_linhas = len(dados)
            elif len(dados) != num_linhas:
                print(f"Número de dados errado! A coluna {nome_coluna} deve ter {num_linhas} linhas! ")
                return None
            
            try:
                if tipo_atual == "date":
                    dados_convertidos = [pd.to_datetime(valor.strip(), format= "%d/%m/%Y") for valor in dados]
                else:
                    dados_convertidos = [tipo_atual(valor.strip()) for valor in dados]
            except ValueError:
                print("Erro ao converter valores\n")
         
            data[nome_coluna] = dados_convertidos
            tipo_dados[nome_coluna] = tipo_atual
            print("Colunas e respetivos dados adicionados ao stock com sucesso\n")
            print()
    
        df = pd.DataFrame(data)
        print("Todos os dados adicionados ao stock com sucesso!\n")
        print(f"Stock criado com {len(data)} e {num_linhas} linhas\n")
        return df
    
    except ValueError:
        print("Erro! O nº de colunas deve ser inteiro")
        return None
        
        
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
    
    while True:
        nome_coluna = input("Introduza o nome da coluna: ").strip()
        if nome_coluna:
            if nome_coluna not in df.columns:
                break
            print(f"A coluna '{nome_coluna}' já está registada no Data Frame! ")
        else:
            print("O Nome da coluna não pode ser vazio! ")     
                    
    print("Escolha o tipo de dados a registar na coluna: ")
    print("1 Texto ")
    print("2 Nºs Inteiros ")
    print("3 Nºs decimais (reais) ")
    print("4 Datas (formato: DD/MM/AAAA) ")
    tipo_escolhido = input("Indique o tipo de dado a registar: ")
            
    if tipo_escolhido == "1":
        tipo_atual = str
    elif tipo_escolhido == "2":
        tipo_atual = int
    elif tipo_escolhido == "3":
        tipo_atual = float
    elif tipo_escolhido == "4":
        tipo_atual = "date"
    else:
        print("Opção Inválida! O tipo de dado a registar será considerado texto! ")
        tipo_atual = str
                
            
    
    dados = input(f"Introduza os dados a registar na coluna - {nome_coluna} separados por vírgula (',')\n").split(',')
    
    num_linhas_df = len(df.index)
    
    if len(dados) != num_linhas_df:
        print(f"Dados inválidos, nº de dados nas colunas deve ser igual a {num_linhas_df} ")
        return None
    
    try:
        if tipo_atual == "date":
            df[nome_coluna] = [pd.to_datetime(valor.strip(), format="%d/%m/%Y") for valor in dados]  
        else:
            df[nome_coluna] = [tipo_atual(valor.strip()) for valor in dados]
        print(f"Coluna '{nome_coluna}' adicionada com sucesso!")    
    except ValueError:
        print("Erro ao converter valores! ")
 
    print()
    
def adicionar_linha():
    global df
    if df.empty:
        print("O Data frame está vazio")
        return
    
    print("\nColunas do Data Frame: ",", ".join(df.columns))
    
    nova_linha = {}
    
    for coluna in df.columns:
        tipo_dado = df[coluna].dtype
        while True:
            if pd.api.types.is_datetime64_any_dtype(tipo_dado):
                dado = input(f"Introduza a data para a coluna '{coluna}' (DD/MM/AAAA): \n")
                try:
                    dado_verif = pd.to_datetime(dado.strip(), format= "%d/%m/%Y")
                    nova_linha[coluna] = dado_verif
                    break
                except ValueError:
                    print("Erro! Formato de data inválido. Use DD/MM/AAAA\n")
            elif pd.api.types.is_numeric_dtype(tipo_dado):
                if 'int' in str(tipo_dado):
                    dado = input(f"Introduza o número inteiro para a coluna '{coluna}': \n")
                    try:
                        dado_verif = int(dado.strip())
                        nova_linha[coluna] = dado_verif
                        break
                    except ValueError:
                        print("Erro! O número deve ser inteiro\n")
                else:
                    dado = input(f"Introduza o número decimal para a coluna '{coluna}': \n")
                    try:
                        dado_verif = float(dado.strip())
                        nova_linha[coluna] = dado_verif
                        break
                    except ValueError:
                        print("Erro! O número deve ser decimal\n")
            else:
                dado = input(f"Introduza o texto para a coluna '{coluna}': \n")
                if dado.strip():
                    nova_linha[coluna] = dado.strip()
                    break
                else:
                    print("Erro! O campo não pode estar vazio\n")
                
        
    
    nova_linha_df = pd.DataFrame([nova_linha])
    df = pd.concat([df, nova_linha_df], ignore_index= True)
    print("Linha adicionada com sucesso! ")
    print(df)
    print()    
    
def remover_linha():
    global df
    if df.empty:
        print("O Data frame está vazio")
        return
    consultar_dataframe()
    
    try:
        linha = int(input("Indique o índice de linha a remover: \n"))
        if linha not in df.index:
            print("Index não encontrado")
            return None
        else:
            df = df.drop(index= linha)
            print(f"Linha nº{linha} removida com sucesso!\n")
    except ValueError:
        print("Erro! O valor do indice deve ser um nº inteiro\n")
        
    
def editar_valor():
    global df
    if df.empty:
        print("O Data frame está vazio")
        return
    consultar_dataframe()
    
    try:
        linha = int(input("Indique o índice da linha a remover: \n"))
        nome_coluna = input("Indique qual a coluna na qual pretende alterar o dado: \n")
        
        if linha not in df.index:
            print(f"Erro! O índice '{linha}' não está registado! \n")
            return None
        
        if nome_coluna not in df.columns:
            print(f"Erro! A coluna '{nome_coluna}, não está registada! \n")
            return None
        
        tipo_dado = df[nome_coluna].dtype.type
        while True:
            dado = input(f"Introduza o dado para a coluna: '{nome_coluna}' ({tipo_dado.__name__}): \n")
            try:
                df.at[linha, nome_coluna] = tipo_dado(dado.strip())
                print("Dado atualizado com sucesso! ")
                break
            except ValueError:
                print(f"Erro! O dado inserido não corresponde a: {tipo_dado.__name__}\n")
    except ValueError:
        print("Erro! O índice da linha deve ser um nº inteiro! ")
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
        df = pd.read_excel(FILE_PATH, sheet_name= 'Stock', index_col= 0)
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
    
    tipo_dado = df[nome_coluna].dtype
    
    if pd.api.types.is_numeric_dtype(df[nome_coluna]):
        try:
            critério = float(input(f"Indique o valor numérico a filtrar em {nome_coluna}: \n"))
        except ValueError:
            print("Erro! O critério de pesquisa não é um valor numérico! \n")
            return 
    elif pd.api.types.is_datetime64_any_dtype(tipo_dado):
        try:
            critério = pd.to_datetime(input(f"Indique a data (DD/MM/AAAA) que pretende filtrar na coluna '{nome_coluna}': "), dayfirst=True)    
        except ValueError:
            print("Erro! O critério não corresponde ao formato de data correto! \n")
            return   
    else:
        critério = input(f"Indique o dado  a filtrar em '{nome_coluna}': \n")         
    
    
    filtro = df[df[nome_coluna] == critério]
    
    print(f"Dados na {nome_coluna} que correspondem a {critério}: \n")
    if filtro.empty:
        print("Nenhum resultado encontrado! ")
    else:    
        print(filtro[nome_coluna])
    print()
    
    
def estatisticas():
    global df
    if df.empty:
        print("O stock está vazio")
        return
    
    colunas_disponiveis = df.select_dtypes(include=['number'])
    if colunas_disponiveis.empty:
        print("Não é possível gerar dados estatísticos porque não existem colunas com valores numéricos no Stock\n")
        return
    
    print("Colunas disponíveis: ", ", ".join(colunas_disponiveis.columns))
    
    coluna = input("Indique em que coluna pretende calcular os dados: \n")
    
    if coluna not in colunas_disponiveis.columns:
        print(f"A coluna {coluna} não existe ou não contém valores numéricos para calcular.\n")
        return
    
    print("Selecione o calculo que pretende efetuar: \n")
    print("1 Média")
    print("2 Soma")
    print("3 Menor valor na coluna")
    print("4 Maior valor na coluna")
    print("5 Mediana")
    
    escolha = input("Indique o cálculo a efetuar: \n")
    
    if escolha == "1":
        print(f"Média dos valores na coluna - {coluna} é: {df[coluna].mean()}\n")
    elif escolha == "2":
        print(f"O resultado total da coluna - {coluna} é: {df[coluna].sum()}\n")
    elif escolha == "3":
        print(f"O menor valor encontrado na coluna - {coluna} é: {df[coluna].min()}\n")
    elif escolha == "4":
        print(f"O maior valor encontrado na coluna - {coluna} é: {df[coluna].max()}\n")
    elif escolha == "5":
        print(f"A mediana dos valores da coluna -  {coluna} é: {df[coluna].media()}\n")
    else:
        print("Erro! Opção inválida")
            
    
    

while True:
    print("<<<< Bem-vindo ao sistema de gestão de stock >>>>\n")
    print("1 Criar Stock ")
    print("2 Consultar Stock")
    print("3 Adicionar coluna")
    print("4 Eliminar coluna")
    print("5 Renomear coluna")
    print("6 Filtrar dados no stock")
    print("7 Adicionar linha ao stock")
    print("8 Remover linha do stock")
    print("9 Editar valor na dataframe")
    print("10 Guardar stock em excel")
    print("11 Importar stock de excel")
    print("12 Sair \n")
    
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
        adicionar_linha()
    elif escolha == "8":
        remover_linha()
    elif escolha == "9":
        editar_valor()
    elif escolha == "10":
        guardar_dataframe()
    elif escolha == "11":
        importar_dataframe()
    elif escolha == "12":
        print("Obrigado \n")
        break
    else:
        print("Erro! Opção Inválida! \n")  
