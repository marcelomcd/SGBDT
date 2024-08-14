import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from numpy import nan

# Definir os nomes das colunas, suas posições e seus limites de caracteres
colunas_e_posicoes = [
    ("Nb", (0, 6), 5, int),
    ("Gp", (6, 9), 2, int),
    ("Co", (9, 12), 2, int),
    ("Nome", (12, 21), 8, str),
    ("Umn", (21, 25), 2, int),
    ("Umx", (25, 29), 2, int),
    ("Pbas", (29, 35), None, float),
    ("Qbas", (35, 41), None, float),
    ("Pmin", (41, 47), 5, float),
    ("Pmax", (47, 54), 5, float),
    ("Qmin", (54, 59), 6, float),
    ("Qmax", (59, 66), 6, float),
    ("Rtrf", (65, 70), 6, float),
    ("Xtrf", (71, 77), 6, float),
    ("Percent", (77, 80), None, float),
    ("Qmin CS", (54, 59), 6, float),
    ("Qmax CS", (59, 66), 6, float),
]

# Separar colunas, posições e limites
colunas, posicoes, limites_de_caracteres, tipologia = zip(*colunas_e_posicoes)


def processar_valor(coluna, valor):
    """Processa valores específicos baseados na coluna. Realiza ajustes como
    remover caracteres específicos ou alterar valores baseados em condições."""
    if coluna == "Pmax":
        return valor.replace("-", "")
    elif coluna == "Qmin":
        if valor == "-9999" or valor == "9999.":
            return "-9999."
        else:
            return f"-{valor.strip('-')}"
    elif coluna == "Qmax":
        if valor == ". 9999":
            return "9999."
        elif "9999" in valor and "." not in valor:
            return "9999."
        elif "." in valor:
            return valor.strip()
    elif valor and "(" in valor:
        return valor.replace("(", "")
    return valor


def ler_arquivo_largura_fixa(
    caminho_arquivo, posicoes, limites_de_caracteres, codificacao
):
    """Lê um arquivo de largura fixa, processa cada linha e retorna um DataFrame do pandas."""
    dados = []
    iniciar_processamento = False
    with open(caminho_arquivo, "r", encoding=codificacao) as arquivo:
        for linha in arquivo:
            if not iniciar_processamento:
                if linha.strip().startswith("( Nb)"):
                    iniciar_processamento = True
                continue
            elif iniciar_processamento and (linha[0] == "(" or linha == "999999\n"):
                continue

            linha_processada = []
            for (inicio, fim), limite, coluna, tipo in zip(
                posicoes[:-2], limites_de_caracteres[:-2], colunas[:-2], tipologia
            ):
                valor = linha[inicio:fim].strip()
                if limite and len(valor) > limite:
                    valor = valor[:limite]
                valor = processar_valor(coluna, valor)

                try:
                    linha_processada.append(tipo(valor))
                except ValueError:
                    linha_processada.append("")  # linha_processada.append(nan)

            linha_processada += ["", ""]
            dados.append(linha_processada)

    dados = pd.DataFrame(dados, columns=colunas)
    ListCo = dados.loc[dados["Co"] != "", "Co"].unique()

    for curCo in ListCo:

        Qmin_CS, Qmax_CS = dados.loc[
            (dados["Co"] == curCo) & (dados["Pmax"] == ""), ["Qmin", "Qmax"]
        ].values[0]
        # localiza linha com info de geração
        dados.loc[
            (dados["Co"] == curCo) & (dados["Pmax"] != ""), ["Qmin CS", "Qmax CS"]
        ] = [Qmin_CS, Qmax_CS]
        # localiza linha com info de compensacao sincrona
        dados.loc[
            (dados["Co"] == curCo) & (dados["Pmax"] == ""),
            ["Pmin", "Pmax", "Qmin", "Qmax", "Qmin CS", "Qmax CS"],
        ] = dados.loc[
            (dados["Co"] == curCo) & (dados["Pmax"] != ""),
            ["Pmin", "Pmax", "Qmin", "Qmax"],
        ].values[
            0
        ].tolist() + [
            Qmin_CS,
            Qmax_CS,
        ]

    # for idx, tipo in enumerate(tipologia):#
    #    dados.iloc[:, idx] = dados.iloc[:, idx].astype(tipo)

    return dados  # pd.DataFrame(dados, columns=colunas)


def readBNT1(arquivo_entrada):
    """Converte um arquivo de entrada .dat para um arquivo .xlsx, detectando a codificação correta e processando o conteúdo."""
    # Detectar a extensão do arquivo
    _, extensao_arquivo = os.path.splitext(arquivo_entrada)

    if extensao_arquivo.lower() != ".dat":
        raise ValueError("Somente arquivos .DAT são suportados")

    # Ler o arquivo com base na sua extensão
    codificacoes = ["utf-8", "latin1", "ISO-8859-1"]
    for codificacao in codificacoes:
        try:
            dados = ler_arquivo_largura_fixa(
                arquivo_entrada, posicoes, limites_de_caracteres, codificacao
            )
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(
            f"Não foi possível decodificar o arquivo com as codificações disponíveis: {codificacoes}"
        )

    return dados
