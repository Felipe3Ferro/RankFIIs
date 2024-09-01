import requests
import csv
import pandas as pd
import os
from bs4 import BeautifulSoup

# Função para fazer a requisição e salvar o resultado
def fetch_and_save(papel):
    url = f'https://www.fundamentus.com.br/detalhes.php?papel={papel}'

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    # Fazer a requisição HTTP
    page = requests.get(url, headers=headers)
    soup = BeautifulSoup(page.content, 'html.parser')

    # Encontrar todos os elementos <span> com a classe 'txt'
    spans = soup.find_all('span', class_='txt')

    # Escrever os resultados no arquivo
    for span in spans:
        print(span.text)

def fetch_table():
    url = "https://www.fundamentus.com.br/fii_resultado.php"
    
    headersRequest = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    # Fazer a requisição HTTP
    page = requests.get(url, headers=headersRequest)
    soup = BeautifulSoup(page.content, 'html.parser')
                         
    # Encontrar a tabela
    table = soup.find('table')
    
    # Extrair cabeçalhos
    headers = [th.get_text(strip=True) for th in table.find_all('th')]
    

    # Abrir arquivo CSV para escrita
    with open('tabela_resultado.csv', 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        
        # Escrever cabeçalhos no CSV
        writer.writerow(headers)
        
        # Iterar sobre as linhas da tabela
        for row in table.find_all('tr')[1:]:  # Ignorar o cabeçalho
            cells = row.find_all('td')
            row_data = [cell.get_text(strip=True) for cell in cells]
            writer.writerow(row_data)

    print("CSV criado com sucesso!")

def csv_to_excel(csv_file_name, excel_file_name):
    """
    Converte um arquivo CSV em um arquivo Excel.

    :param csv_file_name: Nome do arquivo CSV a ser lido.
    :param excel_file_name: Nome do arquivo Excel a ser salvo.
    """
    # Obtenha o diretório atual (pasta do projeto)
    project_dir = os.path.dirname(os.path.abspath(__file__))

    # Construa os caminhos completos dos arquivos
    csv_file_path = os.path.join(project_dir, csv_file_name)
    excel_file_path = os.path.join(project_dir, excel_file_name)

    # Ler o arquivo CSV em um DataFrame
    df = pd.read_csv(csv_file_path)

    # Salvar o DataFrame em um arquivo Excel
    df.to_excel(excel_file_path, index=False, engine='openpyxl')

    print(f'O arquivo Excel foi salvo em: {excel_file_path}')

def filtra():
    """
    Filtra o arquivo CSV para manter apenas as linhas onde o Dividend Yield é
    igual ou superior a 6% e Liquidez é igual ou superior a 300 mil.
    """
    # Defina o nome do arquivo CSV filtrado
    filtered_csv_file_name = 'tabela_resultado_filtrada.csv'
    
    # Obtenha o diretório atual (pasta do projeto)
    project_dir = os.path.dirname(os.path.abspath(__file__))

    # Construa o caminho completo do arquivo CSV original e filtrado
    csv_file_path = os.path.join(project_dir, 'tabela_resultado.csv')
    filtered_csv_file_path = os.path.join(project_dir, filtered_csv_file_name)

    # Ler o arquivo CSV em um DataFrame
    df = pd.read_csv(csv_file_path)

    # Função para limpar e converter os dados
    
    def clean_and_convert(value):
        newValue=''
        try:
            newValue = value.rstrip('%')
            newValue = newValue.replace('.','')
            newValue = newValue.replace(',', '.')
            # Converte para float
            return float(newValue)
        except ValueError as e:
            print(f"Erro ao converter o valor: '{newValue}'. Erro: {e}")
            return None  # Retorna None para valores problemáticos

    # Aplicar a função de limpeza e conversão
    df['Dividend Yield'] = df['Dividend Yield'].apply(clean_and_convert)
    df['Liquidez'] = df['Liquidez'].str.replace('.', '').str.replace(',', '').astype(float)

    # Filtrar o DataFrame
    filtered_df = df[(df['Dividend Yield'] >= 6) & (df['Liquidez'] >= 300000)]

    # Salvar o DataFrame filtrado em um novo arquivo CSV
    filtered_df.to_csv(filtered_csv_file_path, index=False)

    print(f'Arquivo CSV filtrado criado com sucesso: {filtered_csv_file_path}')

def remove_file(file_path):
    """Remove um arquivo se ele existir."""
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f'Arquivo removido: {file_path}')
    else:
        print(f'Arquivo não encontrado: {file_path}')

def setup_files():
    project_dir = os.path.dirname(os.path.abspath(__file__))
    remove_file(os.path.join(project_dir, 'tabela_resultado.csv'))
    remove_file(os.path.join(project_dir, 'tabela_resultado_filtrada.csv'))
    remove_file(os.path.join(project_dir, 'tabela_excel.xlsx'))

# Prepara os arquivos
setup_files()

# Execute as funções
fetch_table()
filtra()
csv_to_excel('tabela_resultado_filtrada.csv', 'tabela_excel.xlsx')
