import customtkinter as ctk
import pandas as pd
from datetime import datetime
import numpy as np
from tkinter import filedialog, messagebox
import os  # Importar o módulo os para manipular o sistema operacional
import ConverterDAT
import ConverterSTB
import ConverterBLT

ctk.set_appearance_mode("light")  # Modo claro
ctk.set_default_color_theme("blue")  # Tema azul

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


class FileSelector(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Seleção de Arquivos")
        self.geometry("700x450")
        self.create_widgets()

    def create_widgets(self):
        # Configuração do layout
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)

        # Arquivo .dat
        self.label_dat = ctk.CTkLabel(self, text="Arquivo BNT1 (.dat):")
        self.label_dat.grid(row=0, column=0, pady=10, padx=10, sticky="e")

        self.entry_dat = ctk.CTkEntry(self, width=400)
        self.entry_dat.grid(row=0, column=1, pady=10, padx=10, sticky="ew")

        self.button_dat = ctk.CTkButton(
            self, text="Buscar", command=self.select_dat_file
        )
        self.button_dat.grid(row=0, column=2, pady=10, padx=10, sticky="w")

        # Arquivo .stb
        self.label_stb = ctk.CTkLabel(self, text="Arquivo DMAQ (.stb):")
        self.label_stb.grid(row=1, column=0, pady=10, padx=10, sticky="e")

        self.entry_stb = ctk.CTkEntry(self, width=400)
        self.entry_stb.grid(row=1, column=1, pady=10, padx=10, sticky="ew")

        self.button_stb = ctk.CTkButton(
            self, text="Buscar", command=self.select_stb_file
        )
        self.button_stb.grid(row=1, column=2, pady=10, padx=10, sticky="w")

        # Arquivo .blt
        self.label_blt = ctk.CTkLabel(self, text="Arquivo BLT (.blt):")
        self.label_blt.grid(row=2, column=0, pady=10, padx=10, sticky="e")

        self.entry_blt = ctk.CTkEntry(self, width=400)
        self.entry_blt.grid(row=2, column=1, pady=10, padx=10, sticky="ew")

        self.button_blt = ctk.CTkButton(
            self, text="Buscar", command=self.select_blt_file
        )
        self.button_blt.grid(row=2, column=2, pady=10, padx=10, sticky="w")

        # Diretório de saída
        self.label_output = ctk.CTkLabel(self, text="Diretório de saída:")
        self.label_output.grid(row=3, column=0, pady=10, padx=10, sticky="e")

        self.entry_output = ctk.CTkEntry(self, width=400)
        self.entry_output.grid(row=3, column=1, pady=10, padx=10, sticky="ew")

        self.button_output = ctk.CTkButton(
            self, text="Selecionar", command=self.select_output_dir
        )
        self.button_output.grid(row=3, column=2, pady=10, padx=10, sticky="w")

        # Botão para selecionar arquivos Excel
        self.button_process = ctk.CTkButton(
            self,
            text="Processar Arquivos",
            command=self.show_excel_selection,
            fg_color="green",
        )
        self.button_process.grid(row=4, column=0, columnspan=3, pady=20, padx=10)

    def select_dat_file(self):
        fname = filedialog.askopenfilename(filetypes=[("Arquivos DAT", "*.dat")])
        if fname:
            self.entry_dat.delete(0, "end")
            self.entry_dat.insert(0, fname)
            self.adjust_entry_width(self.entry_dat)

    def select_stb_file(self):
        fname = filedialog.askopenfilename(filetypes=[("Arquivos STB", "*.stb")])
        if fname:
            self.entry_stb.delete(0, "end")
            self.entry_stb.insert(0, fname)
            self.adjust_entry_width(self.entry_stb)

    def select_blt_file(self):
        fname = filedialog.askopenfilename(filetypes=[("Arquivos BLT", "*.blt")])
        if fname:
            self.entry_blt.delete(0, "end")
            self.entry_blt.insert(0, fname)
            self.adjust_entry_width(self.entry_blt)

    def select_output_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, directory)
            self.adjust_entry_width(self.entry_output)

    def adjust_entry_width(self, entry):
        text_length = len(entry.get())
        new_width = max(400, text_length * 7)  # Ajuste automático do tamanho da entrada
        entry.configure(width=new_width)

    def show_excel_selection(self):
        # Arquivo UG
        self.label_ug = ctk.CTkLabel(self, text="Arquivo Export (UG).xls:")
        self.label_ug.grid(row=5, column=0, pady=10, padx=10, sticky="e")

        self.entry_ug = ctk.CTkEntry(self, width=400)
        self.entry_ug.grid(row=5, column=1, pady=10, padx=10, sticky="ew")

        self.button_ug = ctk.CTkButton(self, text="Buscar", command=self.select_ug_file)
        self.button_ug.grid(row=5, column=2, pady=10, padx=10, sticky="w")

        # Arquivo CS
        self.label_cs = ctk.CTkLabel(self, text="Arquivo Export (CS).xls:")
        self.label_cs.grid(row=6, column=0, pady=10, padx=10, sticky="e")

        self.entry_cs = ctk.CTkEntry(self, width=400)
        self.entry_cs.grid(row=6, column=1, pady=10, padx=10, sticky="ew")

        self.button_cs = ctk.CTkButton(self, text="Buscar", command=self.select_cs_file)
        self.button_cs.grid(row=6, column=2, pady=10, padx=10, sticky="w")

        # Botão para finalizar o processamento
        self.button_final_process = ctk.CTkButton(
            self,
            text="Finalizar Processamento",
            command=self.process_files,
            fg_color="green",
        )
        self.button_final_process.grid(row=7, column=0, columnspan=3, pady=20, padx=10)

        self.button_process.configure(
            state="disabled"
        )  # Desabilita o botão de processar inicial

        self.update_idletasks()  # Atualiza a interface para que os widgets sejam renderizados
        self.geometry(
            f"{self.winfo_reqwidth()}x{self.winfo_reqheight()}"
        )  # Ajusta a geometria conforme o conteúdo

    def select_ug_file(self):
        fname = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls")])
        if fname:
            self.entry_ug.delete(0, "end")
            self.entry_ug.insert(0, fname)
            self.adjust_entry_width(self.entry_ug)

    def select_cs_file(self):
        fname = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls")])
        if fname:
            self.entry_cs.delete(0, "end")
            self.entry_cs.insert(0, fname)
            self.adjust_entry_width(self.entry_cs)

    def process_files(self):
        arquivo_entrada_BNT1 = self.entry_dat.get()
        arquivo_entrada_DMAQ = self.entry_stb.get()
        arquivo_entrada_BLT = self.entry_blt.get()
        saida = self.entry_output.get()
        arquivo_ug = self.entry_ug.get()
        arquivo_cs = self.entry_cs.get()

        # Verifica se todos os arquivos foram selecionados
        if (
            not arquivo_entrada_BNT1
            or not arquivo_entrada_DMAQ
            or not arquivo_entrada_BLT
            or not saida
            or not arquivo_ug
            or not arquivo_cs
        ):
            messagebox.showerror(
                "Erro", "Algum arquivo ou diretório não foi selecionado."
            )
            return

        df_BNT1 = ConverterDAT.readBNT1(arquivo_entrada_BNT1)
        df_DMAQ = ConverterSTB.readDMAQ(arquivo_entrada_DMAQ)
        df_DMDG = ConverterBLT.readDMDG(arquivo_entrada_BLT)

        df_BNT1 = pd.merge(df_BNT1, df_DMAQ, on=["Nb", "Gp"], how="left")
        df_BNT1 = pd.merge(df_BNT1, df_DMDG, on="Mg", how="left")

        df_BNT1.drop_duplicates(subset=["Nb", "Gp"], inplace=True, ignore_index=False)

        agora = datetime.now()
        timestamp = agora.strftime("%d-%m-%y_%H-%M")
        nome_arquivo_saida = f"{'UnifiedUGDataTable'}_{timestamp}.xlsx"
        caminho_saida = f"{saida}/{nome_arquivo_saida}"

        with pd.ExcelWriter(caminho_saida) as escritor:
            df_BNT1.to_excel(escritor, sheet_name="UG+CS", index=False)

        df_BDT_UG = pd.read_excel(arquivo_ug).fillna("")
        df_BDT_CS = pd.read_excel(arquivo_cs).fillna("")

        df_fixed_UG = ComparaUG(df_BNT1, df_BDT_UG, dict_UG)
        df_fixed_CS = ComparaUG(df_BNT1, df_BDT_CS, dict_CS, init_key_UG=2)

        with pd.ExcelWriter(f"{saida}/Fixed_UG_BDT_{timestamp}.xlsx") as escritor:
            pd.concat([df_BDT_UG.head(1), df_fixed_UG], axis=0).to_excel(
                escritor, sheet_name="Fixed", index=False
            )

        with pd.ExcelWriter(f"{saida}/Fixed_CS_BDT_{timestamp}.xlsx") as escritor:
            pd.concat([df_BDT_CS.head(1), df_fixed_CS], axis=0).to_excel(
                escritor, sheet_name="Fixed", index=False
            )

        print("Processamento concluído com sucesso!")

        # Abrir o explorador de arquivos no diretório de saída
        os.startfile(saida)

        # Desativar o botão de finalizar processamento após a execução
        self.button_final_process.configure(state="disabled")

        # Fechar a interface
        self.after(1000, self.destroy)


if __name__ == "__main__":
    app = FileSelector()
    app.mainloop()
