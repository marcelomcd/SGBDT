import flet as ft
from flet import (
    Page,
    ElevatedButton,
    TextField,
    Switch,
    FilePicker,
    AlertDialog,
    app,
    Container,
    Row,
)
import pandas as pd
from datetime import datetime
import numpy as np
import os
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


def show_error_dialog(page, message):
    dialog = AlertDialog(
        title=ft.Text("Erro"),
        content=ft.Text(message),
        actions=[ft.TextButton("OK", on_click=lambda e: close_dialog(page))],
    )
    page.dialog = dialog
    page.dialog.open = True
    page.update()


def close_dialog(page):
    page.dialog.open = False
    page.update()


def select_file_dialog(page, entry, file_types=None, directory=False):
    def on_file_picked(e):
        if e.files:
            entry.value = e.files[0].path
            page.update()

    def on_directory_picked(e):
        if e.path:
            entry.value = e.path
            page.update()

    file_picker = FilePicker(
        on_result=on_directory_picked if directory else on_file_picked
    )
    page.overlay.append(file_picker)
    page.update()  # Assegura que o controle foi adicionado ao page

    if directory:
        file_picker.get_directory_path()
    else:
        file_picker.pick_files(allowed_extensions=file_types)


def main(page: Page):
    page.title = "Seleção de Arquivos"
    page.window.width = 800
    page.window.height = 600
    page.window.resizable = True
    page.theme_mode = ft.ThemeMode.LIGHT  # Define o tema padrão como Claro

    def toggle_appearance_mode(e):
        if page.theme_mode == ft.ThemeMode.LIGHT:
            page.theme_mode = ft.ThemeMode.DARK
        else:
            page.theme_mode = ft.ThemeMode.LIGHT
        update_button_colors()  # Atualiza as cores do texto dos botões

    # Função para atualizar as cores do texto dos botões
    def update_button_colors():
        text_color = (
            ft.colors.WHITE
            if page.theme_mode == ft.ThemeMode.LIGHT
            else ft.colors.BLACK
        )
        for btn in buttons:
            btn.color = text_color
        btn_process.color = (
            text_color  # Botão "Processar Arquivos" também precisa ser atualizado
        )
        page.update()

    # Configuração do switch de tema
    theme_label = ft.Text("Tema: ")
    light_label = ft.Text("Claro")
    dark_label = ft.Text("Escuro")
    switch_theme = ft.Switch(value=False, on_change=toggle_appearance_mode)

    # Organizando os elementos em uma linha
    theme_row = Row([theme_label, light_label, switch_theme, dark_label])

    def process_files(e):
        arquivo_entrada_BNT1 = entry_dat.value
        arquivo_entrada_DMAQ = entry_stb.value
        arquivo_entrada_BLT = entry_blt.value
        saida = entry_output.value
        arquivo_ug = entry_ug.value
        arquivo_cs = entry_cs.value

        if (
            not arquivo_entrada_BNT1
            or not arquivo_entrada_DMAQ
            or not arquivo_entrada_BLT
            or not saida
            or not arquivo_ug
            or not arquivo_cs
        ):
            show_error_dialog(page, "Algum arquivo ou diretório não foi selecionado.")
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
        os.startfile(saida)
        btn_process.disabled = True
        page.update()

    entry_dat = TextField(label="Arquivo BNT1 (.dat)", width=400, expand=True)
    entry_stb = TextField(label="Arquivo DMAQ (.stb)", width=400, expand=True)
    entry_blt = TextField(label="Arquivo BLT (.blt)", width=400, expand=True)
    entry_output = TextField(label="Diretório de saída", width=400, expand=True)
    entry_cs = TextField(
        label="Arquivo Export (CS).xls", width=400, expand=True
    )  # Invertido com UG
    entry_ug = TextField(label="Arquivo Export (UG).xls", width=400, expand=True)

    # Criação dos botões com fundo azul e atualização das cores do texto
    btn_select_dat = ElevatedButton(
        text="Buscar",
        on_click=lambda e: select_file_dialog(page, entry_dat, file_types=["dat"]),
        bgcolor=ft.colors.BLUE,
    )
    btn_select_stb = ElevatedButton(
        text="Buscar",
        on_click=lambda e: select_file_dialog(page, entry_stb, file_types=["stb"]),
        bgcolor=ft.colors.BLUE,
    )
    btn_select_blt = ElevatedButton(
        text="Buscar",
        on_click=lambda e: select_file_dialog(page, entry_blt, file_types=["blt"]),
        bgcolor=ft.colors.BLUE,
    )
    btn_select_output = ElevatedButton(
        text="Selecionar",
        on_click=lambda e: select_file_dialog(page, entry_output, directory=True),
        bgcolor=ft.colors.BLUE,
    )
    btn_select_cs = ElevatedButton(  # Invertido com UG
        text="Buscar",
        on_click=lambda e: select_file_dialog(page, entry_cs, file_types=["xls"]),
        bgcolor=ft.colors.BLUE,
    )
    btn_select_ug = ElevatedButton(
        text="Buscar",
        on_click=lambda e: select_file_dialog(page, entry_ug, file_types=["xls"]),
        bgcolor=ft.colors.BLUE,
    )

    btn_process = ElevatedButton(
        text="Processar Arquivos", on_click=process_files, bgcolor=ft.colors.GREEN
    )

    # Lista de botões para fácil atualização de cores
    buttons = [
        btn_select_dat,
        btn_select_stb,
        btn_select_blt,
        btn_select_output,
        btn_select_cs,
        btn_select_ug,
    ]

    # Atualiza as cores do texto conforme o tema inicial
    update_button_colors()

    # Adicionando os widgets à página em uma ordem específica
    page.add(
        theme_row,
        Row([entry_dat, btn_select_dat]),
        Row([entry_stb, btn_select_stb]),
        Row([entry_blt, btn_select_blt]),
        Row([entry_output, btn_select_output]),
        Row([entry_cs, btn_select_cs]),  # Invertido com UG
        Row([entry_ug, btn_select_ug]),
        Container(
            btn_process, alignment=ft.alignment.center
        ),  # Botão processar centralizado
    )


app(target=main)
