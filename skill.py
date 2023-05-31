import pandas as pd
import glob
import tkinter as tk
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

#Lista todos os arquivos XLS no diretório atual
files = glob.glob("*.xls")

#Cria uma lista vazia para armazenar os DataFrames individuais
dataframes = []

#Itera sobre os arquivos XLS e importa apenas os dados relevantes em um DataFrame
for file in files:
    xls = pd.ExcelFile(file)
    sheet_name = xls.sheet_names[0]  # Nome da primeira planilha
    df = pd.read_excel(file, sheet_name=sheet_name, header=None, skiprows=8, usecols="C:M")
    
    #Seleciona as linhas de C9:M9 até o final do arquivo
    start_row = 8  # Linha inicial
    end_row = df.shape[0] - 1  # Última linha
    df = df.iloc[start_row:end_row+1, :]
    
    dataframes.append(df)

#Concatena os DataFrames em um único DataFrame
merged_df = pd.concat(dataframes)

#Substitui os valores NA por vazio
merged_df.fillna("", inplace=True)

#Exibe o DataFrame resultante
print(merged_df)

#Exporta o DataFrame para um arquivo Excel no formato .xlsx
merged_df.to_excel("merged_data.xlsx", index=False)

#Cria uma janela usando tkinter
window = tk.Tk()
window.title("Relatório Unificado")

#Cria um rótulo com a mensagem "Relatório Unificado"
label = tk.Label(window, text="Relatório Unificado", font=("Arial", 16))
label.pack(padx=20, pady=20)

#Cria um rótulo com a mensagem "Created by: Patrick Coelho"
footer_label = tk.Label(window, text="Created by: Patrick Coelho", font=("Arial", 12))
footer_label.pack(side="bottom", pady=10)

#Função de callback para fechar a janela
def close_window():
    window.destroy()

#Cria um botão para fechar a janela
button = tk.Button(window, text="Fechar", command=close_window)
button.pack(pady=10)

#Exibe a janela
window.mainloop()
