import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime


# Função para ler arquivos .dat e converter para DataFrame
def ler_arquivo_dat(arquivo):
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
                if not linha.strip() or not linha.strip()[0].isdigit():
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

    codificacoes = ["utf-8", "latin1", "ISO-8859-1"]
    for codificacao in codificacoes:
        try:
            dados = ler_arquivo_largura_fixa(
                arquivo, posicoes, limites_de_caracteres, codificacao
            )
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(
            f"Não foi possível decodificar o arquivo com as codificações disponíveis: {codificacoes}"
        )

    return dados


# Função para ler arquivos .stb e converter para DataFrame
def ler_arquivo_stb(arquivo):
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

    def extract_columns(line, columns):
        extracted = {}
        for col_name, (start, end), width in columns:
            extracted[col_name] = line[start:end].strip()
        return extracted

    with open(arquivo, "r", encoding="latin-1") as file:
        lines = file.readlines()

    data = []
    for line in lines:
        stripped_line = line.strip()
        if stripped_line.startswith("DMAQ"):
            continue
        if stripped_line.startswith("(") and not stripped_line.lstrip().startswith(
            tuple("0123456789")
        ):
            continue
        if stripped_line and stripped_line[0].isdigit():
            data.append(extract_columns(line, colunas_e_posicoes))

    if data and data[-1]["Nb"] == "999999":
        data.pop()

    df = pd.DataFrame(data)
    df = df.sort_values(by="Nb")

    return df


# Função para ler arquivos .blt e converter para DataFrame
def ler_arquivo_blt(arquivo):
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

    def extract_columns(line, columns):
        extracted = {}
        for col_name, (start, end), width in columns:
            extracted[col_name] = line[start:end].strip()
        return extracted

    with open(arquivo, "r", encoding="latin-1") as file:
        lines = file.readlines()

    data_1 = []
    data_2 = []
    data_3 = []
    mode = None

    for line in lines:
        if line.startswith("DMDG MD03"):
            mode = "data_1"
            continue
        elif line.startswith("DCST"):
            mode = "data_3"
            continue

        if mode == "data_1":
            if line.startswith(" "):
                continue
            if line.strip().startswith("("):
                continue
            if "  " in line:
                if line.strip().endswith("N"):
                    data_2.append(extract_columns(line, colunas_e_posicoes_2))
                else:
                    data_1.append(extract_columns(line, colunas_e_posicoes_1))

        elif mode == "data_3":
            if line.startswith(" "):
                continue
            if line.strip().startswith("("):
                continue
            if line.strip() and line.strip()[0].isdigit():
                data_3.append(extract_columns(line, colunas_e_posicoes_3))

    df1 = pd.DataFrame(data_1)
    df2 = pd.DataFrame(data_2)
    df3 = pd.DataFrame(data_3)

    return df1, df2, df3


# Função para mesclar os DataFrames dos arquivos .blt
def merge_dataframes(df1, df2, df3):
    merged_df = pd.merge(df1, df2, on="No", how="outer", suffixes=("_df1", "_df2"))
    merged_df = pd.merge(merged_df, df3, on="No", how="outer")
    columns_order = [
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
        "MVA",
        "Fr",
        "C",
        "T",
        "Y1",
        "Y2",
        "X1",
    ]

    merged_df = merged_df[columns_order]
    return merged_df


# Função principal para converter os arquivos e salvar em Excel
def converter_para_excel():
    root = tk.Tk()
    root.withdraw()

    # Seleção dos arquivos de entrada
    arquivo_dat = filedialog.askopenfilename(
        title="Selecione o arquivo .dat", filetypes=[("DAT files", "*.dat")]
    )
    arquivo_stb = filedialog.askopenfilename(
        title="Selecione o arquivo .stb", filetypes=[("STB files", "*.stb")]
    )
    arquivo_blt = filedialog.askopenfilename(
        title="Selecione o arquivo .blt", filetypes=[("BLT files", "*.blt")]
    )

    # Verificar se todos os arquivos foram selecionados
    if not arquivo_dat or not arquivo_stb or not arquivo_blt:
        print("Todos os arquivos .dat, .stb e .blt devem ser selecionados. Saindo...")
        return

    # Ler e processar os arquivos
    df_dat = ler_arquivo_dat(arquivo_dat)
    df_stb = ler_arquivo_stb(arquivo_stb)
    df1_blt, df2_blt, df3_blt = ler_arquivo_blt(arquivo_blt)
    df_blt = merge_dataframes(df1_blt, df2_blt, df3_blt)

    # Renomear a coluna "No" para "Nb" nos DataFrames do arquivo .blt para alinhar com as outras
    df_blt.rename(columns={"No": "Nb"}, inplace=True)

    # Mesclar os DataFrames com base na coluna "Nb"
    df_final = pd.merge(df_dat, df_stb, on="Nb", how="outer")
    df_final = pd.merge(df_final, df_blt, on="Nb", how="outer")

    # Ordenar o DataFrame final pela coluna "Nb" em ordem crescente
    df_final = df_final.sort_values(by="Nb")

    # Seleção do diretório de saída
    diretorio_saida = filedialog.askdirectory(title="Selecione o diretório de saída")
    if not diretorio_saida:
        print("Nenhum diretório de saída selecionado. Saindo...")
        return

    # Gerar nome do arquivo de saída com data e hora
    timestamp = datetime.now().strftime("%d-%m-%y_%H-%M")
    nome_saida = f"saida_unificada_{timestamp}.xlsx"
    caminho_saida = os.path.join(diretorio_saida, nome_saida)

    # Salvar o DataFrame final em um arquivo Excel
    df_final.to_excel(caminho_saida, index=False)
    print(f"Arquivo salvo como {caminho_saida}")


# Teste da função
if __name__ == "__main__":
    converter_para_excel()
