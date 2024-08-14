import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog


# Função para ler e formatar os dados do arquivo
def ler_e_formatar_dados(caminho_arquivo):
    # Abrir o arquivo e ler todas as linhas
    with open(caminho_arquivo, "r", encoding="latin-1") as arquivo:
        linhas = arquivo.readlines()

    dados = []

    # Definir as colunas e suas posições no arquivo
    colunas_e_posicoes = [
        ("Nb", (0, 7), 5),
        ("Gr", (7, 12), 2),
        ("P", (12, 17), 0),
        ("Q", (17, 22), 0),
        ("Uni", (22, 25), 0),
        ("Mg", (25, 31), 5),
        ("Mt", (33, 39), 4),
        ("Mv", (40, 45), 5),
        ("Me", (47, 52), 4),
        ("Xvd", (53, 57), 0),
        ("Nbc", (57, 61), 0),
    ]

    # Função para extrair colunas de uma linha
    def extrair_colunas(linha, colunas):
        extraido = {}
        for nome_coluna, (inicio, fim), largura in colunas:
            extraido[nome_coluna] = linha[inicio:fim].strip()
        return extraido

    # Processar cada linha do arquivo
    for linha in linhas:
        linha_limpa = linha.strip()
        if linha_limpa.startswith("DMAQ"):
            continue
        if linha_limpa.startswith("(") and not linha_limpa.lstrip().startswith(
            tuple("0123456789")
        ):
            continue
        if linha_limpa:
            dados.append(extrair_colunas(linha, colunas_e_posicoes))

    # Excluir a última linha se for "999999"
    if dados and dados[-1]["Nb"] == "999999":
        dados.pop()

    df = pd.DataFrame(dados)

    # Ordenar o dataframe pelo campo "Nb" em ordem crescente
    df = df.sort_values(by="Nb")

    return df


# Função para salvar o dataframe em um arquivo Excel
def salvar_para_excel(df, caminho_saida):
    with pd.ExcelWriter(caminho_saida) as escritor:
        df.to_excel(escritor, sheet_name="DMAQ.stb", index=False)


# Função principal
def main():
    root = tk.Tk()
    root.withdraw()

    # Abrir janela para selecionar o arquivo de entrada
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo .stb", filetypes=[("Arquivos STB", "*.stb")]
    )
    if not caminho_arquivo:
        print("Nenhum arquivo selecionado. Saindo...")
        return

    # Processar dados
    df = ler_e_formatar_dados(caminho_arquivo)

    # Verificações adicionais para garantir que os dados foram lidos corretamente
    if df.empty:
        print("Aviso: o dataframe está vazio!")

    print("Linhas no dataframe:", len(df))

    # Gerar nome do arquivo de saída com data e hora
    agora = datetime.now()
    nome_arquivo_saida = f"{caminho_arquivo.split('/')[-1].split('.')[0]}_{agora.strftime('%d-%m-%y_%H-%M')}.xlsx"

    # Abrir janela para selecionar o local de salvamento
    caminho_saida = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=nome_arquivo_saida,
        title="Salvar arquivo como",
        filetypes=[("Arquivos Excel", "*.xlsx")],
    )
    if not caminho_saida:
        print("Nenhum local de salvamento selecionado. Saindo...")
        return

    # Salvar dados no Excel
    salvar_para_excel(df, caminho_saida)
    print(f"Arquivo salvo em: {caminho_saida}")


if __name__ == "__main__":
    main()
