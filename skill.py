from tkinter import messagebox
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import ttk

# Inicializa a janela do tkinter
window = tk.Tk()
window.title("Concatenar Relatórios")
window.geometry("400x300")

# Função para selecionar os arquivos
def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(title="Selecione os arquivos", filetypes=(("Arquivos Excel", "*.xls"), ("Todos os arquivos", "*.*")))
    entry_arquivos.delete(0, tk.END)
    entry_arquivos.insert(tk.END, ",".join(arquivos))

# Função para confirmar e executar a operação
def confirmar_execucao():
    arquivos = entry_arquivos.get().split(",")
    linhas_pular = int(entry_linhas_pular.get())
    nome_arquivo_final = entry_nome_arquivo_final.get()
    local_salvamento = filedialog.askdirectory(title="Selecione o local de salvamento")
    executar(arquivos, linhas_pular, nome_arquivo_final, local_salvamento)

# Função para executar a operação
def executar(arquivos, linhas_pular, nome_arquivo_final, local_salvamento):
    # Cria uma lista vazia para armazenar os DataFrames individuais
    dataframes = []

    # Itera sobre os arquivos e importa apenas os dados relevantes em um DataFrame
    for file in arquivos:
        xls = pd.ExcelFile(file)
        sheet_name = xls.sheet_names[0]  # Nome da primeira planilha
        df = pd.read_excel(file, sheet_name=sheet_name, header=None, skiprows=linhas_pular, usecols="C:M")

        dataframes.append(df)

    # Concatena os DataFrames em um único DataFrame
    merged_df = pd.concat(dataframes)
    novos_nomes_colunas = ["Hora", "Nome de açăo", "", "","","Realização", "Lugar", "Utilizador", "Tipo", "Subtipo", "Parada"]
    merged_df.columns = novos_nomes_colunas
    # Substitui os valores NA por vazio
    merged_df.fillna("", inplace=True)

    # Exporta o DataFrame para um arquivo Excel
    caminho_arquivo_final = local_salvamento + "/" + nome_arquivo_final + ".xlsx"
    merged_df.to_excel(caminho_arquivo_final, index=False)

    # Cria uma nova janela para exibir os dados da planilha
    tabela_window = tk.Toplevel(window)
    tabela_window.title("Dados da Planilha")
    tabela_window.geometry("600x500")

    # Cria o widget Treeview para exibir a tabela
    tree = ttk.Treeview(tabela_window)
    tree.pack(fill=tk.BOTH, expand=True)

    # Define as colunas personalizadas da tabela
    colunas = ["Hora", "Nome de açăo", "", "","","Realização", "Lugar", "Utilizador", "Tipo", "Subtipo", "Parada"]
    tree['columns'] = colunas
    for col in colunas:
        tree.heading(col, text=col, anchor=tk.W)
        tree.column(col, anchor=tk.W)

    # Insere os dados na tabela
    for _, row in merged_df.iterrows():
        tree.insert("", tk.END, values=list(row))

    # Cria uma barra de rolagem vertical
    scrollbar_y = ttk.Scrollbar(tabela_window, orient=tk.VERTICAL, command=tree.yview)
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=scrollbar_y.set)

    # Cria uma barra de rolagem horizontal
    scrollbar_x = ttk.Scrollbar(tabela_window, orient=tk.HORIZONTAL, command=tree.xview)
    scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
    tree.configure(xscrollcommand=scrollbar_x.set)

    # Função para confirmar o arquivo
    def confirmar_arquivo():
            tabela_window.destroy()

            #Cria uma janela usando tkinter
            finish_window = tk.Tk()
            finish_window.title("Relatórios Unificados")

            #Cria um rótulo com a mensagem "Relatório Unificado"
            label = tk.Label(finish_window, text="Relatórios Unificados", font=("Arial", 16))
            label.pack(padx=20, pady=20)

            #Cria um rótulo com a mensagem "Created by: Patrick Coelho"
            footer_label = tk.Label(finish_window, text="Created by: Patrick Coelho", font=("Arial", 6))
            footer_label.pack(side="bottom", anchor="se", padx=10, pady=10)

            #Função de callback para fechar a janela
            def close_window():
                finish_window.destroy()

            #Cria um botão para fechar a janela
            button = tk.Button(finish_window, text="Fechar", command=close_window)
            button.pack(pady=10)
            window.destroy()

    # Cria um botão para confirmar o arquivo
    button_confirmar_arquivo = tk.Button(tabela_window, text="Confirmar", command=confirmar_arquivo)
    button_confirmar_arquivo.pack(pady=10)

# Label e Entry para selecionar os arquivos
label_arquivos = tk.Label(window, text="Selecione os arquivos:")
label_arquivos.pack(pady=10)
entry_arquivos = tk.Entry(window, width=40)
entry_arquivos.pack(pady=5)
button_arquivos = tk.Button(window, text="Selecionar", command=selecionar_arquivos)
button_arquivos.pack(pady=5)

# Label e Entry para selecionar a quantidade de linhas a pular
label_linhas_pular = tk.Label(window, text="Quantidade de linhas a pular:")
label_linhas_pular.pack(pady=10)
entry_linhas_pular = tk.Entry(window)
entry_linhas_pular.pack(pady=5)

# Label e Entry para selecionar o nome do arquivo final
label_nome_arquivo_final = tk.Label(window, text="Nome do arquivo final:")
label_nome_arquivo_final.pack(pady=10)
entry_nome_arquivo_final = tk.Entry(window)
entry_nome_arquivo_final.pack(pady=5)

# Botão para confirmar e executar a operação
button_confirmar = tk.Button(window, text="Confirmar", command=confirmar_execucao)
button_confirmar.pack(pady=10)

# Exibe a janela principal
window.mainloop()
