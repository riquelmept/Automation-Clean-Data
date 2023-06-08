import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time


def mostrar_previa(df_agrupado):
    window = tk.Toplevel()
    window.title("Prévia da Planilha")
    window.geometry("600x500")

    # Criar o widget Treeview
    treeview = ttk.Treeview(window)
    treeview.pack(fill="both", expand=True)

    # Configurar a barra de rolagem vertical
    scrollbar_y = ttk.Scrollbar(window, orient="vertical", command=treeview.yview)
    scrollbar_y.pack(side="right", fill="y")
    treeview.configure(yscrollcommand=scrollbar_y.set)

    # Configurar a barra de rolagem horizontal
    scrollbar_x = ttk.Scrollbar(window, orient="horizontal", command=treeview.xview)
    scrollbar_x.pack(side="bottom", fill="x")
    treeview.configure(xscrollcommand=scrollbar_x.set)

    # Definir as colunas
    treeview["columns"] = df_agrupado.columns.tolist()

    # Definir o cabeçalho das colunas
    for column in df_agrupado.columns:
        treeview.heading(column, text=column)

    # Adicionar os dados ao Treeview
    for index, row in df_agrupado.iterrows():
        treeview.insert("", "end", values=list(row))

    def confirmar_salvar():
        # Abrir janela de seleção do local de salvamento
        save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
        if save_path:
            df_agrupado.to_excel(save_path, index=False)
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

        else:
            messagebox.showwarning("Nenhum local de salvamento selecionado", "Nenhum local de salvamento selecionado.")

    button_confirmar = tk.Button(window, text="Confirmar e Salvar", command=confirmar_salvar)
    button_confirmar.pack()

    window.mainloop()


# Inicializar a janela do tkinter
root = tk.Tk()
root.withdraw()  # Esconder a janela principal

# Solicitar ao usuário que selecione o arquivo
file_path = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))

# Verificar se um arquivo foi selecionado
if file_path:
    # Carregar o arquivo Excel em um DataFrame
    df = pd.read_excel(file_path)

    # Substituir valores nulos pela string vazia na coluna 'Localização'
    df['Localização'].fillna('', inplace=True)

    # Remover colunas desnecessárias
    colunas_remover = ['Lote', 'Refer.', 'Validade', 'Dt Fabricação', 'Un']
    df = df.drop(colunas_remover, axis=1)

    # Agrupar e somar as quantidades por item e descrição
    df_agrupado = df.groupby(['Item', 'Descrição Item'])['Qtd Liquida'].sum().reset_index()

    # Unir as localizações em uma única coluna
    df_agrupado['Localizações'] = df.groupby(['Item', 'Descrição Item'])['Localização'].apply(lambda x: ', '.join(x.astype(str))).reset_index()['Localização']

    # Remover vírgulas no início das células com string vazia antes do primeiro endereço
    df_agrupado['Localizações'] = df_agrupado['Localizações'].str.replace(r'^,\s*', '', regex=True)

    # Remover vírgulas no início das células com string vazia antes do primeiro endereço
    df_agrupado['Localizações'] = df_agrupado['Localizações'].str.replace(r'^,\s*', '', regex=True)

    # Mostrar a prévia da planilha
    mostrar_previa(df_agrupado)

else:
    print("Nenhum arquivo selecionado.")
    time.sleep(2)
