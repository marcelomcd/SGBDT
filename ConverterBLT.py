import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from numpy import nan


# Função para ler e formatar dados do arquivo
def ler_e_formatar_dados(caminho_arquivo):
    # Abrir o arquivo e ler todas as linhas
    with open(caminho_arquivo, "r", encoding="latin-1") as arquivo:
        linhas = arquivo.readlines()

    # Inicializar listas para armazenar dados
    dados_1 = []
    dados_2 = []
    dados_3 = []

    # Definir colunas e posições para o primeiro conjunto de dados
    colunas_e_posicoes_1 = [
        ("Mg", (0, 7), 5, int),
        ("CS", (7, 12), 5, int),
        ("Xd", (12, 17), 5, float),
        ("Xq", (17, 22), 5, float),
        ("X'd", (22, 27), 5, float),
        ("X'q", (27, 32), 5, float),
        ('X"d', (32, 37), 5, float),
        ("Xl", (37, 42), 5, float),
        ("T'd", (42, 47), 5, float),
        ("T'q", (47, 52), 5, float),
        ('T"d', (52, 57), 5, float),
        ('T"q', (57, 62), 5, float),
    ]

    # Definir colunas e posições para o segundo conjunto de dados
    colunas_e_posicoes_2 = [
        ("Mg", (0, 7), 5, int),
        ("Ra", (7, 12), 5, float),
        ("H", (12, 17), 5, float),
        ("D", (17, 22), 5, float),
        ("MVA", (22, 27), 5, float),
        ("Fr", (27, 29), 5, int),
        ("C", (29, 31), 5, str),
    ]

    # Definir colunas e posições para o terceiro conjunto de dados
    colunas_e_posicoes_3 = [
        ("CS", (0, 7), 5, int),
        ("T", (7, 9), 3, int),
        ("Y1", (9, 18), 8, float),
        ("Y2", (18, 27), 8, float),
        ("X1", (27, 36), 8, float),
    ]

    # Função para extrair colunas de uma linha
    def extrair_colunas(linha, colunas, flag=2):
        extraido = {}
        for nome_coluna, (inicio, fim), largura, tipo in colunas:
            extraido[nome_coluna] = linha[inicio:fim].strip()
            try:
                extraido[nome_coluna] = tipo(linha[inicio:fim].strip())
            except ValueError:
                extraido[nome_coluna] = ""

        return extraido

    modo = None

    # Processar cada linha do arquivo
    for linha in linhas:
        if linha.startswith("DMDG MD03") or linha.startswith("DMDG MD02"):
            modo = "dados_1"
            continue
        elif linha.startswith("DCST"):
            modo = "dados_3"
            continue
        elif linha.startswith("999999") or linha.startswith("FIM"):
            continue

        if modo == "dados_1":
            if linha.strip() == "" or linha.strip().startswith("("):
                continue
            if "  " in linha:
                if linha.strip().endswith("N"):  # Verifica se a linha termina com "N"
                    dados_2.append(extrair_colunas(linha, colunas_e_posicoes_2, flag=1))
                else:  # Primeiro conjunto de colunas
                    dados_1.append(extrair_colunas(linha, colunas_e_posicoes_1))

        elif modo == "dados_3":
            if linha.strip() == "" or linha.strip().startswith("("):
                continue
            dados_3.append(extrair_colunas(linha, colunas_e_posicoes_3))

    # Criar dataframes a partir dos dados extraídos
    df1 = pd.DataFrame(dados_1)
    df2 = pd.DataFrame(dados_2)
    df3 = pd.DataFrame(dados_3)

    return df1, df2, df3


# Função para mesclar os dataframes
def mesclar_dataframes(df1, df2, df3):
    # Mesclar os dataframes na coluna 'No'
    df_mesclado = pd.merge(df1, df2, on="Mg", how="left", suffixes=("_df1", "_df2"))
    df_mesclado = pd.merge(df_mesclado, df3, on="CS", how="left")

    # Reordenar colunas para corresponder ao formato necessário
    ordem_colunas = [
        "Mg",
        "CS",
        "Xd",
        "Xq",
        "X'd",
        "X'q",
        'X"d',
        "Xl",
        "T'd",
        "T'q",
        'T"d',
        'T"q',
        "Ra",
        "H",
        "D",
        "MVA",
        "Fr",
        "C",
        "T",
        "Y1",
        "Y2",
        "X1",
    ]

    df_mesclado = df_mesclado[ordem_colunas]
    return df_mesclado


# Função para salvar o dataframe mesclado em um arquivo Excel
def salvar_para_excel(df_mesclado, caminho_saida):
    with pd.ExcelWriter(caminho_saida) as escritor:
        df_mesclado.to_excel(escritor, sheet_name="DMDG.blt (Completo)", index=False)


def readDMDG(caminho_arquivo):
    # Processar dados
    df1, df2, df3 = ler_e_formatar_dados(caminho_arquivo)

    # Verificações adicionais para garantir que os dados foram lidos corretamente
    if df1.empty:
        print("Aviso: df1 está vazio!")
    if df2.empty:
        print("Aviso: df2 está vazio!")
    if df3.empty:
        print("Aviso: df3 está vazio!")

    # Mesclar os dataframes
    df_mesclado = mesclar_dataframes(df1, df2, df3)

    return df_mesclado
