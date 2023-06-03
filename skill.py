from tkinter import messagebox
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Inicializa a janela do tkinter
window = tk.Tk()
window.title("Concatenar Relatórios")
window.geometry("400x300")

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
    window.destroy()  # Fecha a janela

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

    # Substitui os valores NA por vazio
    merged_df.fillna("", inplace=True)

    # Exporta o DataFrame para um arquivo Excel
    caminho_arquivo_final = local_salvamento + "/" + nome_arquivo_final + ".xlsx"
    merged_df.to_excel(caminho_arquivo_final, index=False)

    #Cria uma janela usando tkinter
    window = tk.Tk()
    window.title("Relatório Unificado")

    #Cria um rótulo com a mensagem "Relatório Unificado"
    label = tk.Label(window, text="Relatório Unificado", font=("Arial", 16))
    label.pack(padx=20, pady=20)

    #Cria um rótulo com a mensagem "Created by: Patrick Coelho"
    footer_label = tk.Label(window, text="Created by: Patrick Coelho", font=("Arial", 6))
    footer_label.pack(side="bottom", anchor="se", padx=10, pady=10)

    #Função de callback para fechar a janela
    def close_window():
        window.destroy()

    #Cria um botão para fechar a janela
    button = tk.Button(window, text="Fechar", command=close_window)
    button.pack(pady=10)
    
#    # Exibe uma mensagem de conclusão
#    messagebox.showinfo("Concatenar Relatórios", "Arquivo exportado com sucesso!")

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

# Label de crédito
label_credit = tk.Label(window, text="Created by: Patrick Coelho")
label_credit.pack(side=tk.BOTTOM, pady=5)

# Exibe a janela
window.mainloop()
