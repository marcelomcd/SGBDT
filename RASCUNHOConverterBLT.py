import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog

import pandas as pd

# Definir os nomes das colunas, suas posições e seus limites de caracteres
colunas_e_posicoes_1 = [
    ("No", (0, 5), 5),
    ("CS", (7, 12), 5),
    ("Xd", (12, 17), 5),
    ("Xq", (17, 23), 5),
    ("X'd", (23, 28), 5),
    ("X'q", (28, 33), 5),
    ('X"d', (33, 38), 5),
    ("Xl", (38, 43), 5),
    ("T'd", (43, 48), 5),
    ("T'q", (48, 53), 5),
    ('T"d', (53, 58), 5),
    ('T"q', (58, 62), 4),
]

colunas_e_posicoes_2 = [
    ("No", (0, 4), 4),
    ("Ra", (8, 12), 4),
    ("H", (13, 17), 4),
    ("D", (18, 22), 4),
    ("MVA", (23, 27), 4),
    ("Fr", (28, 29), 4),
    ("C", (30, 31), 4),
]

colunas_e_posicoes_3 = [
    ("No", (0, 4), 4),
    ("T", (8, 9), 1),
    ("Y1", (10, 17), 8),
    ("Y2", (19, 26), 8),
    ("X1", (28, 35), 8),
]

# Separar colunas, posições e limites
colunas_1, posicoes_1, limites_de_caracteres_1 = zip(*colunas_e_posicoes_1)
colunas_2, posicoes_2, limites_de_caracteres_2 = zip(*colunas_e_posicoes_2)
colunas_3, posicoes_3, limites_de_caracteres_3 = zip(*colunas_e_posicoes_3)


def processar_valor(coluna, valor):
    return valor.strip()


def ler_arquivo_largura_fixa(caminho_arquivo, codificacao):
    dados = []
    iniciar_processamento = False
    comecar_regua_3 = False
    with open(caminho_arquivo, "r", encoding=codificacao) as arquivo:
        linha_anterior = None
        for linha in arquivo:
            if not iniciar_processamento:
                if linha.strip().startswith("(No)"):
                    iniciar_processamento = True
                continue

            # Ignorar linhas desnecessárias que começam com "("
            if linha.strip().startswith("("):
                continue

            if linha_anterior is None:
                linha_anterior = linha
                continue

            no_linha_atual = linha[:7].strip()
            no_linha_anterior = linha_anterior[:7].strip()

            if no_linha_atual == no_linha_anterior:
                if comecar_regua_3:
                    linha_processada_3 = []
                    for (inicio, fim), limite, coluna in zip(
                        posicoes_3, limites_de_caracteres_3, colunas_3
                    ):
                        valor = linha[inicio:fim].strip()
                        if limite and len(valor) > limite:
                            valor = valor[:limite]
                        valor = processar_valor(coluna, valor)
                        linha_processada_3.append(valor)

                    dados.append(linha_processada_3)
                else:
                    linha_processada_1 = []
                    for (inicio, fim), limite, coluna in zip(
                        posicoes_1, limites_de_caracteres_1, colunas_1
                    ):
                        valor = linha_anterior[inicio:fim].strip()
                        if limite and len(valor) > limite:
                            valor = valor[:limite]
                        valor = processar_valor(coluna, valor)
                        linha_processada_1.append(valor)

                    linha_processada_2 = []
                    for (inicio, fim), limite, coluna in zip(
                        posicoes_2, limites_de_caracteres_2, colunas_2
                    ):
                        valor = linha[inicio:fim].strip()
                        if limite and len(valor) > limite:
                            valor = valor[:limite]
                        valor = processar_valor(coluna, valor)
                        linha_processada_2.append(valor)

                    dados.append(linha_processada_1 + linha_processada_2)
                linha_anterior = None
            else:
                linha_anterior = linha
                if not comecar_regua_3 and linha.strip().startswith("(No)   T"):
                    comecar_regua_3 = True

    if comecar_regua_3:
        return pd.DataFrame(dados, columns=colunas_3)
    else:
        return pd.DataFrame(dados, columns=colunas_1 + colunas_2)


def converter_arquivo(arquivo_entrada):
    # Detectar a extensão do arquivo
    _, extensao_arquivo = os.path.splitext(arquivo_entrada)

    if extensao_arquivo.lower() != ".blt":
        raise ValueError("Somente arquivos .BLT são suportados")

    # Ler o arquivo com base na sua extensão
    codificacoes = ["utf-8", "latin1", "ISO-8859-1"]
    for codificacao in codificacoes:
        try:
            dados = ler_arquivo_largura_fixa(arquivo_entrada, codificacao)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(
            f"Não foi possível decodificar o arquivo com as codificações disponíveis: {codificacoes}"
        )

    # Criar o nome do arquivo de saída com data e hora
    nome_base = os.path.basename(arquivo_entrada).rsplit(".", 1)[0]
    timestamp = datetime.now().strftime("%d-%m-%y_%H-%M")
    nome_saida = f"{nome_base}_{timestamp}.xlsx"

    # Selecionar o diretório de saída
    diretorio_saida = filedialog.askdirectory(title="Selecione o diretório de saída")
    if not diretorio_saida:
        print("Nenhum diretório de saída selecionado")
        return

    # Definir o caminho completo do arquivo de saída
    caminho_saida = os.path.join(diretorio_saida, nome_saida)

    # Salvar o arquivo como XLSX
    dados.to_excel(caminho_saida, index=False)

    print(f"Arquivo salvo como {caminho_saida}")


def principal():
    # Criar uma janela raiz do Tkinter
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela raiz

    # Selecionar o arquivo de entrada
    arquivo_entrada = filedialog.askopenfilename(
        title="Selecione o arquivo de entrada",
        filetypes=(("Arquivos BLT", "*.blt"),),
    )
    if not arquivo_entrada:
        print("Nenhum arquivo selecionado")
        return

    # Converter o arquivo
    converter_arquivo(arquivo_entrada)


if __name__ == "__main__":
    principal()
