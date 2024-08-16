import os
import random
import pandas as pd

# Função para gerar dados aleatórios
def gerar_dados():
    meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    dados = {
        "Mês": [],
        "Vendas": [],
        "Custos": [],
        "Lucro": [],
        "Funcionários": [],
        "Satisfação Cliente": []
    }
    
    for _ in range(20):
        mes = random.choice(meses)
        vendas = random.randint(80000, 200000)
        custos = random.randint(50000, vendas)
        lucro = vendas - custos
        funcionarios = random.randint(30, 70)
        satisfacao = round(random.uniform(3.0, 5.0), 1)
        
        dados["Mês"].append(mes)
        dados["Vendas"].append(vendas)
        dados["Custos"].append(custos)
        dados["Lucro"].append(lucro)
        dados["Funcionários"].append(funcionarios)
        dados["Satisfação Cliente"].append(satisfacao)
    
    return pd.DataFrame(dados)

# Função para gerar múltiplos arquivos Excel
def gerar_arquivos_excel(quantidade_arquivos, nome_pasta):
    # Cria a pasta se não existir
    if not os.path.exists(nome_pasta):
        os.makedirs(nome_pasta)
    
    for i in range(quantidade_arquivos):
        df = gerar_dados()
        nome_arquivo = f"{nome_pasta}/dados_financeiros_{i+1}.xlsx"
        df.to_excel(nome_arquivo, index=False)
        print(f"Arquivo gerado: {nome_arquivo}")

# Parâmetros do usuário
quantidade_arquivos = int(input("Quantos arquivos deseja gerar? "))
nome_pasta = input("Nome da pasta para salvar os arquivos: ")

# Gerar os arquivos
gerar_arquivos_excel(quantidade_arquivos, nome_pasta)
