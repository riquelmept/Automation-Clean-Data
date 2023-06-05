import pandas as pd
from tkinter import Tk, filedialog

# Inicializar a janela do tkinter
root = Tk()
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

    # Exportar o DataFrame tratado para um novo arquivo Excel
    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    if save_path:
        df_agrupado.to_excel(save_path, index=False)
        print("Arquivo salvo com sucesso.")
    else:
        print("Nenhum local de salvamento selecionado.")
else:
    print("Nenhum arquivo selecionado.")

