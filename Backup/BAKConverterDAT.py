import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

# Definir os nomes das colunas, suas posições e seus limites de caracteres
colunas_e_posicoes = [
    ("Nb", (0, 6), 5),
    ("Gp", (6, 9), 2),
    ("Co", (9, 12), 2),
    ("Nome", (12, 21), 8),
    ("Umn", (21, 25), 2),
    ("Umx", (25, 29), 2),
    ("Pbas", (29, 35), None),
    ("Qbas", (35, 41), None),
    ("Pmin", (41, 47), 5),
    ("Pmax", (47, 54), 5),
    ("Qmin", (54, 59), 6),
    ("Qmax", (59, 66), 6),
    ("Rtrf", (66, 71), 6),
    ("Xtrf", (71, 77), 6),
    ("Percent", (77, 80), None),
]

# Separar colunas, posições e limites
colunas, posicoes, limites_de_caracteres = zip(*colunas_e_posicoes)


def processar_valor(coluna, valor):
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
    dados = []
    iniciar_processamento = False
    with open(caminho_arquivo, "r", encoding=codificacao) as arquivo:
        for linha in arquivo:
            if not iniciar_processamento:
                if linha.strip().startswith("( Nb)"):
                    iniciar_processamento = True
                continue
            linha_processada = []
            for (inicio, fim), limite, coluna in zip(
                posicoes, limites_de_caracteres, colunas
            ):
                valor = linha[inicio:fim].strip()
                if limite and len(valor) > limite:
                    valor = valor[:limite]
                valor = processar_valor(coluna, valor)
                linha_processada.append(valor)
            dados.append(linha_processada)
    return pd.DataFrame(dados, columns=colunas)


def converter_arquivo(arquivo_entrada):
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
        filetypes=(("Arquivos DAT", "*.dat"),),
    )
    if not arquivo_entrada:
        print("Nenhum arquivo selecionado")
        return

    # Converter o arquivo
    converter_arquivo(arquivo_entrada)


if __name__ == "__main__":
    principal()
