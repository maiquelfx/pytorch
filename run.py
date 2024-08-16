import os
import torch
import torch.nn as nn
import torch.optim as optim
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

# Função para carregar todos os dados XLSX de uma pasta
def carregar_dados(pasta):
    dados = []
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.xlsx'):
            caminho = os.path.join(pasta, arquivo)
            df = pd.read_excel(caminho)
            dados.append(df)
    return pd.concat(dados, ignore_index=True)

# Definir o modelo
class ClienteSatisfacaoModel(nn.Module):
    def __init__(self):
        super(ClienteSatisfacaoModel, self).__init__()
        self.fc1 = nn.Linear(4, 8)
        self.fc2 = nn.Linear(8, 4)
        self.fc3 = nn.Linear(4, 1)
        self.relu = nn.ReLU()
    
    def forward(self, x):
        x = self.relu(self.fc1(x))
        x = self.relu(self.fc2(x))
        x = self.fc3(x)
        return x

# Função para treinar o modelo
def treinar_modelo(X, y):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

    X_train = torch.FloatTensor(X_train)
    y_train = torch.FloatTensor(y_train).reshape(-1, 1)
    X_test = torch.FloatTensor(X_test)
    y_test = torch.FloatTensor(y_test).reshape(-1, 1)

    model = ClienteSatisfacaoModel()
    criterion = nn.MSELoss()
    optimizer = optim.Adam(model.parameters(), lr=0.01)

    num_epochs = 1000
    for epoch in range(num_epochs):
        outputs = model(X_train)
        loss = criterion(outputs, y_train)
        
        optimizer.zero_grad()
        loss.backward()
        optimizer.step()
        
        if (epoch+1) % 100 == 0:
            print(f'Epoch [{epoch+1}/{num_epochs}], Loss: {loss.item():.4f}')

    model.eval()
    with torch.no_grad():
        predictions = model(X_test)
        test_loss = criterion(predictions, y_test)
        print(f'Erro médio quadrático no conjunto de teste: {test_loss.item():.4f}')

    return model, scaler

# Função para prever satisfação do cliente e atualizar planilha
def prever_e_atualizar(modelo, scaler, arquivo_entrada, arquivo_saida):
    df = pd.read_excel(arquivo_entrada)
    X = df[['Vendas', 'Custos', 'Lucro', 'Funcionários']].values
    X_scaled = scaler.transform(X)
    
    X_tensor = torch.FloatTensor(X_scaled)
    with torch.no_grad():
        previsoes = modelo(X_tensor)
    
    df['Satisfação Cliente'] = previsoes.numpy()
    df.to_excel(arquivo_saida, index=False)
    print(f"Arquivo atualizado salvo como: {arquivo_saida}")

# Carregar e preparar dados para treinamento
nome_pasta = input("Nome da pasta com os arquivos de treinamento XLSX: ")
df = carregar_dados(nome_pasta)

X = df[['Vendas', 'Custos', 'Lucro', 'Funcionários']].values
y = df['Satisfação Cliente'].values

scaler = StandardScaler()
X = scaler.fit_transform(X)

# Treinar o modelo
modelo, scaler = treinar_modelo(X, y)

# Prever para novo arquivo
arquivo_entrada = input("Nome do arquivo XLSX para previsão: ")
arquivo_saida = input("Nome do arquivo XLSX de saída com previsões: ")
prever_e_atualizar(modelo, scaler, arquivo_entrada, arquivo_saida)
