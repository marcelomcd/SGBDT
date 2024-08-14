import sys
import pandas as pd
from datetime import datetime
import numpy as np
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QFileDialog,
    QVBoxLayout,
    QLabel,
    QLineEdit,
    QHBoxLayout,
)

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


def ComparaUG(df_BDAna, df_BDT, dict_BD=dict_UG, init_key_UG=3):
    df_BDT.dropna(subset=[dict_BD["Nb"]], inplace=True)

    df_bothBDT = df_BDT.merge(
        df_BDAna,
        left_on=[dict_BD["Nb"], dict_BD["Gp"]],
        right_on=["Nb", "Gp"],
        how="inner",
    ).drop(df_BDAna.columns, axis=1)

    list_new_both = []

    for idx, linha_BDT in df_bothBDT.iterrows():
        Nb = linha_BDT[dict_BD["Nb"]]
        Gp = linha_BDT[dict_BD["Gp"]]
        linha_Ana = df_BDAna.loc[(df_BDAna["Nb"] == Nb) & (df_BDAna["Gp"] == Gp)]
        flagChange = False

        diffCols = (
            linha_BDT[list(dict_BD.values())[init_key_UG:]].values
            != linha_Ana[list(dict_BD.keys())[init_key_UG:]].values[0]
        )

        if np.any(diffCols):
            linha_BDT.loc[np.array(list(dict_BD.values()))[init_key_UG:][diffCols]] = (
                linha_Ana[list(dict_BD.keys())[init_key_UG:]].iloc[0, diffCols].values
            )

            flagChange = True

        if init_key_UG == 3:
            Pmax = linha_Ana["Pmax"].values[0]
            if linha_BDT["LimiteSuperiorOperativoPotenciaAtiva"] > Pmax:
                linha_BDT.loc["LimiteSuperiorOperativoPotenciaAtiva"] = Pmax
                flagChange = True

        if flagChange:
            list_new_both.append(linha_BDT)

    return pd.DataFrame(list_new_both)


class FileSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.layout = QVBoxLayout()

        self.label_dat = QLabel("Arquivo BNT1 (.dat):")
        self.layout.addWidget(self.label_dat)

        self.line_dat = QLineEdit(self)
        self.layout.addWidget(self.line_dat)

        self.button_dat = QPushButton("Buscar Arquivo .dat", self)
        self.button_dat.clicked.connect(self.select_dat_file)
        self.layout.addWidget(self.button_dat)

        self.label_stb = QLabel("Arquivo DMAQ (.stb):")
        self.layout.addWidget(self.label_stb)

        self.line_stb = QLineEdit(self)
        self.layout.addWidget(self.line_stb)

        self.button_stb = QPushButton("Buscar Arquivo .stb", self)
        self.button_stb.clicked.connect(self.select_stb_file)
        self.layout.addWidget(self.button_stb)

        self.label_blt = QLabel("Arquivo BLT (.blt):")
        self.layout.addWidget(self.label_blt)

        self.line_blt = QLineEdit(self)
        self.layout.addWidget(self.line_blt)

        self.button_blt = QPushButton("Buscar Arquivo .blt", self)
        self.button_blt.clicked.connect(self.select_blt_file)
        self.layout.addWidget(self.button_blt)

        self.label_output = QLabel(
            "Selecione o diretório para salvar a planilha unificada:"
        )
        self.layout.addWidget(self.label_output)

        self.line_output = QLineEdit(self)
        self.layout.addWidget(self.line_output)

        self.button_output = QPushButton("Selecionar diretório", self)
        self.button_output.clicked.connect(self.select_output_dir)
        self.layout.addWidget(self.button_output)

        self.button_process = QPushButton("Processar Arquivos", self)
        self.button_process.clicked.connect(self.show_excel_selection)
        self.layout.addWidget(self.button_process)

        self.setLayout(self.layout)
        self.setWindowTitle("Seleção de Arquivos")
        self.setGeometry(300, 300, 500, 300)

    def select_dat_file(self):
        fname, _ = QFileDialog.getOpenFileName(
            self, "Selecione o arquivo BNT1 (.dat)", "", "Arquivos DAT (*.dat)"
        )
        if fname:
            self.line_dat.setText(fname)

    def select_stb_file(self):
        fname, _ = QFileDialog.getOpenFileName(
            self, "Selecione o arquivo DMAQ (.stb)", "", "Arquivos STB (*.stb)"
        )
        if fname:
            self.line_stb.setText(fname)

    def select_blt_file(self):
        fname, _ = QFileDialog.getOpenFileName(
            self, "Selecione o arquivo BLT (.blt)", "", "Arquivos BLT (*.blt)"
        )
        if fname:
            self.line_blt.setText(fname)

    def select_output_dir(self):
        directory = QFileDialog.getExistingDirectory(
            self, "Selecione o diretório para salvar a planilha unificada:"
        )
        if directory:
            self.line_output.setText(directory)

    def show_excel_selection(self):
        # Criar campos adicionais para selecionar os arquivos UG e CS
        self.label_ug = QLabel("Arquivo Export (UG).xls:")
        self.layout.addWidget(self.label_ug)

        self.line_ug = QLineEdit(self)
        self.layout.addWidget(self.line_ug)

        self.button_ug = QPushButton("Buscar Arquivo Export (UG).xls", self)
        self.button_ug.clicked.connect(self.select_ug_file)
        self.layout.addWidget(self.button_ug)

        self.label_cs = QLabel("Arquivo Export (CS).xls:")
        self.layout.addWidget(self.label_cs)

        self.line_cs = QLineEdit(self)
        self.layout.addWidget(self.line_cs)

        self.button_cs = QPushButton("Buscar Arquivo Export (CS).xls", self)
        self.button_cs.clicked.connect(self.select_cs_file)
        self.layout.addWidget(self.button_cs)

        # Adicionar botão para finalizar o processamento
        self.button_final_process = QPushButton("Finalizar Processamento", self)
        self.button_final_process.clicked.connect(self.process_files)
        self.layout.addWidget(self.button_final_process)

        # Desabilitar o botão de processar enquanto os arquivos não forem escolhidos
        self.button_process.setEnabled(False)

    def select_ug_file(self):
        fname, _ = QFileDialog.getOpenFileName(
            self, "Selecione o arquivo Export (UG).xls", "", "Arquivos Excel (*.xls)"
        )
        if fname:
            self.line_ug.setText(fname)

    def select_cs_file(self):
        fname, _ = QFileDialog.getOpenFileName(
            self, "Selecione o arquivo Export (CS).xls", "", "Arquivos Excel (*.xls)"
        )
        if fname:
            self.line_cs.setText(fname)

    def process_files(self):
        arquivo_entrada_BNT1 = self.line_dat.text()
        arquivo_entrada_DMAQ = self.line_stb.text()
        arquivo_entrada_BLT = self.line_blt.text()
        saida = self.line_output.text()
        arquivo_ug = self.line_ug.text()
        arquivo_cs = self.line_cs.text()

        if (
            not arquivo_entrada_BNT1
            or not arquivo_entrada_DMAQ
            or not arquivo_entrada_BLT
            or not saida
            or not arquivo_ug
            or not arquivo_cs
        ):
            print("Algum arquivo ou diretório não foi selecionado.")
            return

        df_BNT1 = ConverterDAT.readBNT1(arquivo_entrada_BNT1)
        df_DMAQ = ConverterSTB.readDMAQ(arquivo_entrada_DMAQ)
        df_DMDG = ConverterBLT.readDMDG(arquivo_entrada_BLT)

        df_BNT1 = pd.merge(df_BNT1, df_DMAQ, on=["Nb", "Gp"], how="left")
        df_BNT1 = pd.merge(df_BNT1, df_DMDG, on="Mg", how="left")

        df_BNT1.drop_duplicates(subset=["Nb", "Gp"], inplace=True, ignore_index=False)

        agora = datetime.now()
        nome_arquivo_saida = (
            f"{'UnifiedUGDataTable'}_{agora.strftime('%d-%m-%y_%H-%M')}.xlsx"
        )

        caminho_saida = f"{saida}/{nome_arquivo_saida}"

        with pd.ExcelWriter(caminho_saida) as escritor:
            df_BNT1.to_excel(escritor, sheet_name="UG+CS", index=False)

        df_BDT_UG = pd.read_excel(arquivo_ug).fillna("")
        df_BDT_CS = pd.read_excel(arquivo_cs).fillna("")

        df_fixed_UG = ComparaUG(df_BNT1, df_BDT_UG, dict_UG)
        df_fixed_CS = ComparaUG(df_BNT1, df_BDT_CS, dict_CS, init_key_UG=2)

        with pd.ExcelWriter(f"{saida}/Fixed_UG_BDT.xlsx") as escritor:
            pd.concat([df_BDT_UG.head(1), df_fixed_UG], axis=0).to_excel(
                escritor, sheet_name="Fixed", index=False
            )

        with pd.ExcelWriter(f"{saida}/Fixed_CS_BDT.xlsx") as escritor:
            pd.concat([df_BDT_CS.head(1), df_fixed_CS], axis=0).to_excel(
                escritor, sheet_name="Fixed", index=False
            )

        print("Processamento concluído com sucesso!")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = FileSelector()
    ex.show()
    sys.exit(app.exec_())
