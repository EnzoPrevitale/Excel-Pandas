# Importando a biblioteca pandas, para manipular dataframes
import pandas as pd

# Criando o dicionário data para armazenar os dados do DataFrame
data = {}

# Solicitando a quantidade de colunas e linhas ao usuário
colunas = int(input("Digite a quantidade de colunas: "))
linhas = int(input("Digite a quantidade de linhas: "))

# Solicitando dados para preencher a planilha
# Colunas
for coluna in range(colunas):
    data[input(f"Digite o conteúdo da coluna {coluna}: ")] = []

# Linhas
for linha in range(linhas):
    for chave in data:
        data[chave].append(input(f"Digite o conteúdo de {chave} {linha}: "))

# Declarando o DataFrame df com os dados de data
df = pd.DataFrame(data)

# Imprimindo o DataFrame
print("Aqui está a sua tabela:")
print(df)

# Verificando se o usuário deseja salvar a tabela em xlsx
print("Deseja salvar a tabela em xlsx?")
xlsx = True if input("S/N: ").upper() == "S" else False

# Salvando o DataFrame em xlsx
if xlsx:
    try:
        # IMPORTANTE: A biblioteca openpyxl precisa estar instalada para o DataFrame ser salvo em xlsx
        print("Salvar no Excel:")
        df.to_excel(f"{input('Digite como você quer salvar o seu arquivo, sem a extensão: ')}.xlsx", sheet_name=input("Digite o nome da planilha: "), index=False)
        print("A planilha foi salva.")
    except PermissionError:
        print("ERRO. Não é possível sobrescrever o arquivo com ele aberto.") # Arquivo não pode estar aberto ao sobrescrever
else:
    print("A tabela não foi salva em xlsx.")
