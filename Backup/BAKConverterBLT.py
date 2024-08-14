import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog


def ler_e_formatar_dados(caminho_arquivo):
    with open(caminho_arquivo, "r", encoding="latin-1") as arquivo:
        linhas = arquivo.readlines()

    dados_1 = []
    dados_2 = []
    dados_3 = []

    colunas_e_posicoes_1 = [
        ("No", (0, 7), 5),
        ("CS", (7, 12), 5),
        ("Xd", (12, 17), 5),
        ("Xq", (17, 22), 5),
        ("X'd", (22, 27), 5),
        ("X'q", (27, 32), 5),
        ('X"d', (32, 37), 5),
        ("Xl", (37, 42), 5),
        ("T'd", (42, 47), 5),
        ("T'q", (47, 52), 5),
        ('T"d', (52, 57), 5),
        ('T"q', (57, 62), 5),
    ]

    colunas_e_posicoes_2 = [
        ("No", (0, 7), 5),
        ("Ra", (7, 12), 5),
        ("H", (12, 17), 5),
        ("D", (17, 22), 5),
        ("MVA", (22, 27), 5),
        ("Fr", (27, 29), 5),
        ("C", (29, 31), 5),
    ]

    colunas_e_posicoes_3 = [
        ("No", (0, 7), 5),
        ("T", (7, 9), 3),
        ("Y1", (9, 18), 8),
        ("Y2", (18, 27), 8),
        ("X1", (27, 36), 8),
    ]

    def extrair_colunas(linha, colunas):
        extraido = {}
        for nome_coluna, (inicio, fim), largura in colunas:
            extraido[nome_coluna] = linha[inicio:fim].strip()
        return extraido

    modo = None

    for linha in linhas:
        if linha.startswith("DMDG MD03"):
            modo = "dados_1"
            continue
        elif linha.startswith("DCST"):
            modo = "dados_3"
            continue

        if modo == "dados_1":
            if linha.startswith(" "):
                continue
            if linha.strip().startswith("("):
                continue
            if "  " in linha:
                if linha.strip().endswith("N"):  # Verifica se a linha termina com "N"
                    dados_2.append(extrair_colunas(linha, colunas_e_posicoes_2))
                else:  # Primeiro conjunto de colunas
                    dados_1.append(extrair_colunas(linha, colunas_e_posicoes_1))

        elif modo == "dados_3":
            if linha.startswith(" "):
                continue
            if linha.strip().startswith("("):
                continue
            dados_3.append(extrair_colunas(linha, colunas_e_posicoes_3))

    df1 = pd.DataFrame(dados_1)
    df2 = pd.DataFrame(dados_2)
    df3 = pd.DataFrame(dados_3)

    return df1, df2, df3


def mesclar_dataframes(df1, df2, df3):
    # Mesclar os dataframes na coluna 'No'
    df_mesclado = pd.merge(df1, df2, on="No", como="outer", suffixes=("_df1", "_df2"))
    df_mesclado = pd.merge(df_mesclado, df3, on="No", como="outer")

    # Reordenar colunas para corresponder ao formato necessário
    ordem_colunas = [
        "No",
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


def salvar_para_excel(df_mesclado, caminho_saida):
    with pd.ExcelWriter(caminho_saida) as escritor:
        df_mesclado.to_excel(escritor, sheet_name="DMDG.blt (Completo)", index=False)


def main():
    root = tk.Tk()
    root.withdraw()

    # Abrir janela para selecionar o arquivo de entrada
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo .blt", filetypes=[("Arquivos BLT", "*.blt")]
    )
    if not caminho_arquivo:
        print("Nenhum arquivo selecionado. Saindo...")
        return

    # Processar dados
    df1, df2, df3 = ler_e_formatar_dados(caminho_arquivo)

    # Verificações adicionais para garantir que os dados foram lidos corretamente
    if df1.empty:
        print("Aviso: df1 está vazio!")
    if df2.empty:
        print("Aviso: df2 está vazio!")
    if df3.empty:
        print("Aviso: df3 está vazio!")

    print("Linhas em df1:", len(df1))
    print("Linhas em df2:", len(df2))
    print("Linhas em df3:", len(df3))

    # Mesclar os dataframes
    df_mesclado = mesclar_dataframes(df1, df2, df3)

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
    salvar_para_excel(df_mesclado, caminho_saida)
    print(f"Arquivo salvo em: {caminho_saida}")


if __name__ == "__main__":
    main()
