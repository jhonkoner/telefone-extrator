
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, ttk

# Função para extrair números de telefone de uma string
def extrair_telefones(texto):
    # Expressão regular para encontrar números de telefone brasileiros (com DDD)
    telefone_pattern = re.compile(r'\\b(?:[1-9][0-9])?\\d{8,9}\\b')
    return telefone_pattern.findall(texto)

# Função para extrair números de telefone de um arquivo XLS/XLSX
def extrair_telefones_xls(caminho_arquivo, progress_var, progress_step):
    dados = []
    df = pd.read_excel(caminho_arquivo, header=None)
    total_rows = len(df)
    for index, row in df.iterrows():
        cpf = row[0]
        nome = row[1]
        telefones = set()
        for coluna in row[2:]:  # Ignorar as duas primeiras colunas (CPF e Nome)
            if coluna:
                telefones.update(extrair_telefones(str(coluna)))
        telefones = list(telefones)
        fones_atuais = telefones[:3]
        outros_fones = telefones[3:]
        dados.append([cpf, nome] + fones_atuais + outros_fones)
        progress_var.set(progress_var.get() + progress_step / total_rows)
    return dados

# Função para abrir a janela de diálogo para selecionar arquivos
def selecionar_arquivos():
    root = tk.Tk()
    root.withdraw()
    caminhos_arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos",
        filetypes=[("Excel files", "*.xls *.xlsx")]
    )
    return caminhos_arquivos

# Função para abrir a janela de diálogo para salvar o arquivo
def salvar_arquivo():
    root = tk.Tk()
    root.withdraw()
    caminho_arquivo = filedialog.asksaveasfilename(
        title="Salvar números de telefone como",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    return caminho_arquivo

# Função principal
def main():
    # Selecionar os arquivos de entrada
    caminhos_arquivos = selecionar_arquivos()
    
    # Lista para armazenar todos os dados extraídos
    todos_dados = []
    
    # Criar a janela de progresso
    progress_window = tk.Tk()
    progress_window.title("Progresso da Extração")
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=100)
    progress_bar.pack(fill=tk.X, expand=1, padx=20, pady=20)
    
    # Calcular o passo de progresso
    total_steps = sum(len(pd.read_excel(caminho_arquivo, header=None)) for caminho_arquivo in caminhos_arquivos if caminho_arquivo.endswith('.xls') or caminho_arquivo.endswith('.xlsx')) if caminhos_arquivos else 1
    progress_step = 100 / total_steps
    
    # Processar cada arquivo selecionado
    for caminho_arquivo in caminhos_arquivos:
        if caminho_arquivo.endswith('.xls') or caminho_arquivo.endswith('.xlsx'):
            dados = extrair_telefones_xls(caminho_arquivo, progress_var, progress_step)
            todos_dados.extend(dados)
        else:
            continue
    
    # Selecionar o arquivo de saída
    caminho_saida = salvar_arquivo()
    
    # Criar um DataFrame com os dados extraídos
    max_telefones = max(len(dados[2:]) for dados in todos_dados)  # Encontrar o máximo de telefones em uma linha
    colunas = ["CPF", "Nome", "Fone Atual 1", "Fone Atual 2", "Fone Atual 3"] + [f"Telefone {i+1}" for i in range(max_telefones - 3)]
    df_novo = pd.DataFrame(todos_dados, columns=colunas)
    
    # Salvar o DataFrame em um arquivo XLSX
    df_novo.to_excel(caminho_saida, index=False)
    
    # Fechar a janela de progresso
    progress_window.destroy()
    
    print("Números de telefone extraídos e salvos com sucesso!")

if __name__ == "__main__":
    main()

# Autor: JhonKoner07
# Versão: 1.0
