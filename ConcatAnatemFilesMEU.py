import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import numpy as np

import ConverterDAT
import ConverterSTB
import ConverterBLT

dict_UG = {
    "Nb": "ConjuntoUnidadeGeradoraEstudoEletrico.Barramento.NumeroBarramento",
    "Gp": "ConjuntoUnidadeGeradoraEstudoEletrico.NumeroGrupo",
    "Pmax": "LimiteSuperiorOperativoPotenciaAtiva",
    "Pmax": "PotenciaNominal",
    "Xd": "ReatanciaSincronaEixoDireto",
    "Xq": "ReatanciaSincronaEixoQuadratura",
    "X'd": "ReatanciaTransitoria",
    "X'q": "ReatanciaTransitoriaEixoQuadratura",
    'X"d': "ReatanciaSubtransitoriaEixoDireto",
    "Xl": "ReatanciaSincronaEixoDispersao",
    "T'd": "ConstanteTempoTransitoriaEixoDireto",
    "T'q": "ConstanteTempoTransitoriaEixoQuadratura",
    'T"d': "ConstanteTempoSubTransitoriaEixoDireto",
    'T"q': "ConstanteTempoSubTransitoriaEixoQuadratura",
    "Ra": "ResistenciaEstator",
    "H": "Inercia",
    "D": "ConstanteAmortecimento",
    "MVA": "PotenciaAparentePlaca",
    "T": "TipoCurvaSaturacao",
    "X1": "CoordenadaX1",
    "Y1": "PrimeiraCoordenadaY1",
    "Y2": "SegundaCoordenadaY2",
    "Pmin": "PotenciaMinima",
    "Qmin": "MvarMinimoGerador",
    "Qmax": "MvarMaximoGerador",
    "Qmin CS": "PotenciaReativaMinima",
    "Qmax CS": "PotenciaReativaMaxima",
    "Pmax": "LimiteSuperiorRegulacaoInicialPotenciaAtiva",
}

dict_CS = {
    "Nb": "ConjuntoCompensadorSincronoEstudoEletrico.Barramento.NumeroBarramento",
    "Gp": "ConjuntoCompensadorSincronoEstudoEletrico.NumeroGrupo",
    "Xd": "ReatanciaSincronaEixoDireto",
    "Xq": "ReatanciaSincronaEixoQuadratura",
    "X'd": "ReatanciaTransitoria",
    "X'q": "ReatanciaTransitoriaEixoQuadratura",
    'X"d': "ReatanciaSubtransitoriaEixoDireto",
    "Xl": "ReatanciaSincronaEixoDispersao",
    "T'd": "ConstanteTempoTransitoriaEixoDireto",
    "T'q": "ConstanteTempoTransitoriaEixoQuadratura",
    'T"d': "ConstanteTempoSubTransitoriaEixoDireto",
    'T"q': "ConstanteTempoSubTransitoriaEixoQuadratura",
    "Ra": "ResistenciaEstator",
    "H": "Inercia",
    "D": "ConstanteAmortecimento",
    "MVA": "PotenciaReativaNominal",
    "T": "TipoCurvaSaturacao",
    "X1": "CoordenadaX1",
    "Y1": "PrimeiraCoordenadaY1",
    "Y2": "SegundaCoordenadaY2",
    "Qmin": "LimiteFisicoInferiorPotenciaReativa",
    "Qmax": "LimiteFisicoSuperiorPotenciaReativa",
}


def ComparaUG(
    df_BDAna, df_BDT, dict_BD=dict_UG, init_key_UG=3
):  # init_key_UG = 3 (UG) / 2 (CS)

    df_BDT.dropna(subset=[dict_BD["Nb"]], inplace=True)

    # IDENTIFICA DAS BARRAS QUE TEM NA BDANATEM, AS BARRAS QUE NAO TEM NA BDT
    # df_onlyBDAna = df_BDAna.loc[~df_BDAna['Nb'].isin(df_BDT[dict_BD['Nb']])]

    # IDENTIFICA DAS BARRAS QUE TEM NA BDANATEM, AS BARRAS QUE NAO TEM NA BDT
    # df_onlyBDT   = df_BDT.loc[~df_BDT[dict_BD['Nb']].isin(df_BDAna['Nb'])]

    # IDENTIFICA DAS BARRAS QUE TEM EM AMBAS BD, QUAIS CELULAS ESTAO ERRADAS E CORRIGE
    df_bothBDT = df_BDT.merge(
        df_BDAna,
        left_on=[dict_BD["Nb"], dict_BD["Gp"]],
        right_on=["Nb", "Gp"],
        how="inner",
    ).drop(df_BDAna.columns, axis=1)

    list_new_both = (
        []
    )  # Lista onde serao armazenadas as linhas da BDT que tiveram celulas alteradas

    # Loop atraves das linhas dos sub-dataset da BDT cujas barras e grupos estao contidas na BDAnatem
    for idx, linha_BDT in df_bothBDT.iterrows():

        Nb = linha_BDT[dict_BD["Nb"]]  # Numero da barra
        Gp = linha_BDT[dict_BD["Gp"]]  # Numero do grupo
        linha_Ana = df_BDAna.loc[
            (df_BDAna["Nb"] == Nb) & (df_BDAna["Gp"] == Gp)
        ]  # Linha correspondente do Anatem
        flagChange = False  # Flag (1) se teve alteracao, ou (0) senão

        diffCols = (
            linha_BDT[list(dict_BD.values())[init_key_UG:]].values
            != linha_Ana[list(dict_BD.keys())[init_key_UG:]].values[0]
        )  # Das colunas do dict_BD a partir da 4a chave,
        # quais tem valores diferentes na linha atual da
        # BDT em relacao ao respectivo na BDAnatem.

        if np.any(
            diffCols
        ):  # Se houve diferenca em qualquer coluna, substitui nas células a correção necessaria

            linha_BDT.loc[np.array(list(dict_BD.values()))[init_key_UG:][diffCols]] = (
                linha_Ana[list(dict_BD.keys())[init_key_UG:]].iloc[0, diffCols].values
            )

            flagChange = True

        if init_key_UG == 3:
            # Tratando do campo LimiteSuperiorOperat5ivoPotenciaAtiva, no caso de UG
            Pmax = linha_Ana["Pmax"].values[0]
            if linha_BDT["LimiteSuperiorOperativoPotenciaAtiva"] > Pmax:
                linha_BDT.loc["LimiteSuperiorOperativoPotenciaAtiva"] = Pmax
                flagChange = True

        # Adiciona a linha corrigida, se houver mudança, a Lista com outras linhas reificadas
        if flagChange:
            list_new_both.append(linha_BDT)

    return pd.DataFrame(list_new_both)


def main():

    root = tk.Tk()
    root.withdraw()  # Ocultar a janela raiz

    # Selecionar oS arquivoS de entrada
    arquivo_entrada_BNT1 = filedialog.askopenfilename(
        title="Selecione o arquivo BNT1 de entrada",
        filetypes=(("Arquivos DAT", "*.dat"),),
    )

    arquivo_entrada_DMAQ = filedialog.askopenfilename(
        title="Selecione o arquivo com DMAQ de entrada",
        filetypes=(("Arquivos STB", "*.stb"),),
    )

    arquivo_entrada_BLT = filedialog.askopenfilename(
        title="Selecione o arquivo BLT de entrada",
        filetypes=(("Arquivos BLT", "*.blt"),),
    )

    if not arquivo_entrada_BNT1 or not arquivo_entrada_DMAQ or not arquivo_entrada_BLT:
        print("Algum arquivo nao selecionado.")
    else:
        # Converter o arquivo
        df_BNT1 = ConverterDAT.readBNT1(arquivo_entrada_BNT1)
        df_DMAQ = ConverterSTB.readDMAQ(arquivo_entrada_DMAQ)
        df_DMDG = ConverterBLT.readDMDG(arquivo_entrada_BLT)

        saida = r"C:\Users\marce\OneDrive\Área de Trabalho\SGBDT"

        df_BNT1 = pd.merge(df_BNT1, df_DMAQ, on=["Nb", "Gp"], how="left")
        df_BNT1 = pd.merge(df_BNT1, df_DMDG, on="Mg", how="left")

        df_BNT1.drop_duplicates(subset=["Nb", "Gp"], inplace=True, ignore_index=False)

        # Gera nome do arquivo de saída com data e hora atuais
        agora = datetime.now()
        nome_arquivo_saida = (
            f"{'UnifiedUGDataTable'}_{agora.strftime('%d-%m-%y_%H-%M')}.xlsx"
        )

        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=nome_arquivo_saida,
            title="Salvar arquivo como",
            filetypes=[("Arquivos Excel", "*.xlsx")],
        )

        with pd.ExcelWriter(caminho_saida) as escritor:
            df_BNT1.to_excel(escritor, sheet_name="UG+CS", index=False)

        # INICIO DA ROTINA DE COMPARACAO
        #############################################################################################
        #
        # Planilhas da BDT
        BDT_UG_file = r"C:\Users\marce\OneDrive\Área de Trabalho\SGBDT\Arquivos Base\Export (UG).xls"
        BDT_CS_file = r"C:\Users\marce\OneDrive\Área de Trabalho\SGBDT\Arquivos Base\Export (CS).xls"

        df_BDT_UG = pd.read_excel(BDT_UG_file).fillna("")
        df_BDT_CS = pd.read_excel(BDT_CS_file).fillna("")

        df_fixed_UG = ComparaUG(df_BNT1, df_BDT_UG)
        df_fixed_CS = ComparaUG(df_BNT1, df_BDT_CS, dict_CS, init_key_UG=2)

        with pd.ExcelWriter(saida + r"\Fixed_UG_BDT.xlsx") as escritor:
            pd.concat([df_BDT_UG.head(1), df_fixed_UG], axis=0).to_excel(
                escritor, sheet_name="Fixed", index=False
            )

        with pd.ExcelWriter(saida + r"\Fixed_CS_BDT.xlsx") as escritor:
            pd.concat([df_BDT_CS.head(1), df_fixed_CS], axis=0).to_excel(
                escritor, sheet_name="Fixed", index=False
            )


if __name__ == "__main__":
    main()
