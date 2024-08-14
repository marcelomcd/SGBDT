import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from numpy import nan


# Função para ler e formatar os dados do arquivo .stb
def readDMAQ(caminho_arquivo):
    # Abre o arquivo com codificação latin-1 e lê todas as linhas
    with open(caminho_arquivo, "r", encoding="latin-1") as arquivo:
        linhas = arquivo.readlines()

    dados = []

    # Definição das colunas e suas posições no arquivo
    colunas_e_posicoes = [
        ("Nb", (0, 7), 5, int),
        ("Gp", (7, 12), 2, int),
        ("P", (12, 17), 0, float),
        ("Q", (17, 22), 0, float),
        ("Uni", (22, 25), 0, int),
        ("Mg", (25, 31), 5, int),
        ("Mt", (33, 39), 4, str),
        ("Mv", (40, 45), 5, str),
        ("Me", (47, 52), 4, str),
        ("Xvd", (53, 57), 0, float),
        ("Nbc", (57, 61), 0, int),
    ]

    # Função para extrair colunas de uma linha específica
    def extrair_colunas(linha, colunas):
        extraido = {}
        for nome_coluna, (inicio, fim), largura, tipo in colunas:
            try:
                extraido[nome_coluna] = tipo(linha[inicio:fim].strip())
            except ValueError:
                extraido[nome_coluna] = ""
        return extraido

    # Processa cada linha do arquivo
    for linha in linhas:
        linha_strip = linha.strip()
        if linha_strip.startswith("DMAQ"):
            continue
        if linha_strip.startswith("(") and not linha_strip.lstrip().startswith(
            tuple("0123456789")
        ):
            continue
        if linha_strip:
            dados.append(extrair_colunas(linha, colunas_e_posicoes))

    # Exclui a última linha se for "999999"
    if dados and dados[-1]["Nb"] == "999999":
        dados.pop()

    # Cria um DataFrame com os dados extraídos
    df = pd.DataFrame(dados)

    # Ordena o DataFrame pelo campo "Nb" em ordem crescente
    df = df.sort_values(by="Nb")

    return df


# Função para salvar o DataFrame em um arquivo Excel
def salvar_em_excel(df, caminho_saida):
    with pd.ExcelWriter(caminho_saida) as escritor:
        df.to_excel(escritor, sheet_name="DMAQ.stb", index=False)
