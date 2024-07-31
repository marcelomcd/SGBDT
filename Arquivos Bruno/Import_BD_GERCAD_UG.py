import pandas as pd

# Caminhos dos arquivos
caminho_base_gercad = r"C:\Users\bruno.pestana\Downloads\Dados_UG.xls"
caminho_anat0 = r"C:\Users\bruno.pestana\Desktop\Anel\BD2303R1\BDados\01_Anat0"
caminho_dmdg = r"C:\Users\bruno.pestana\Desktop\Anel\BD2303R1\BDados\02_MAQ\DMDG.blt"

# Ler o arquivo xls e armazenar os dados em um DataFrame
base_gercad = pd.read_excel(caminho_base_gercad, header = 1, na_filter=False)

colunas = ["ID ONS","Estação","Número da Barra","Número do Grupo","Tipo do Rotor", \
           "Xd(%)","Xq(%)","X'd(%)","X'q(%)","X''d(%)","Xl(%)","T'd0(s)","T'q0(s)","T''d0(s)","T''q0(s)","Ra(%)","H(s)","D(pu/pu)","MVA", \
           "Potência Ativa Mínima(MW)","Potência Ativa Máxima(MW)","Limite Superior de regulação inicial de potência ativa","Mvar Mínimo - Gerador","Mvar Máximo - Gerador","Mvar Mínimo - Síncrono","Mvar Máximo - Síncrono"]

# Crie uma máscara booleana onde True representa valores não-NaN na coluna Número da Barra
mascara = base_gercad[colunas[2]] != ''

# Filtrar o DataFrame usando a máscara
base_gercad_filtrada = base_gercad[mascara]

# Abrir o arquivo DMAQ para leitura
with open(caminho_anat0 + "\DMAQ.stb", 'r') as arquivo:
    dicionario_dmaq = {}  # Criar um dicionário vazio para armazenar os dados

    # Iterar pelas linhas do arquivo
    for linha in arquivo:
        # Ignorar as linhas iniciadas com "("
        if not linha.startswith('('):
            # Remover espaços em branco e quebras de linha da linha atual
            linha_limpa = linha.strip()

            # Dividir a linha em partes usando espaços em branco como separador
            partes = linha_limpa.split()

            # Verificar se a linha tem pelo menos 3 partes (número mínimo para criar a chave)
            if len(partes) >= 4:
                # Criar a chave usando os dois primeiros valores unidos por " - "
                chave = f"{partes[0]} - {partes[1]}"

                # Criar uma lista com os demais valores da linha
                valores = partes[2]

                # Associar a lista como valor para a chave no dicionário
                dicionario_dmaq[chave] = valores
                
with open(caminho_anat0 + "\BNT1.dat", 'r') as arquivo:
        dicionario_bnt1_temp = {}  # Criar um dicionário vazio para armazenar os dados
        
        # Iterar pelas linhas do arquivo
        for linha in arquivo:
            # Ignorar as linhas iniciadas com "("
            if not linha.startswith('('):
                # Processar apenas as linhas válidas (não vazias)
                if linha.strip() and linha.strip()!='999999':
                    # Extrair os campos relevantes usando os índices diretamente
                    nb = linha[0:5].strip()
                    gp = linha[6:8].strip()
                    co = linha[9:11].strip()
                    try:
                        pmax = float(linha[47:53].strip())
                        pmin = float(linha[41:47].strip())
                    except:
                        pmax = linha[47:53].strip()
                        pmin = linha[41:47].strip()
                    qmin = float(linha[53:59].strip())
                    qmax = float(linha[59:65].strip())
    
                    # Criar a chave usando Nb e Gp unidos por " - "
                    chave = f"{nb} - {gp}"
    
                    # Criar a lista com os valores relevantes
                    valores = [co, pmin, pmax, pmax, qmin, qmax]
    
                    # Associar a lista como valor para a chave no dicionário
                    dicionario_bnt1_temp[chave] = valores

dicionario_bnt1 = dicionario_bnt1_temp.copy()
for chave, valor in dicionario_bnt1_temp.items():
    if valor[0] == '':
        dicionario_bnt1[chave] = valor[1:] + ['','']
    elif valor[0] != '' and valor[1] != '':
        for chave_, valor_ in dicionario_bnt1_temp.items():
            if valor_[0] == valor[0] and valor_[1] == '':
                dicionario_bnt1[chave] = valor[1:] + valor_[4:]
                dicionario_bnt1[chave_] = valor[1:] + valor_[4:]

with open(caminho_dmdg, 'r', encoding='latin-1') as arquivo:
        dicionario_dmdg = {}  # Criar um dicionário vazio para armazenar os dados
        count = 0
        # Iterar pelas linhas do arquivo
        for linha in arquivo:
            if linha.strip()[0:4] == 'DCST':
                break
            # Ignorar as linhas iniciadas com "("
            if not linha.startswith('(') and len(linha.strip().split())>=3:                
                # Extrair os campos relevantes usando os índices diretamente
                no = str(int(linha[0:4].strip()))
                if no not in dicionario_dmdg:
                    xd = linha[12:17].strip()
                    xq = linha[17:22].strip()
                    x_d = linha[22:27].strip()
                    x_q = linha[27:32].strip()
                    x__d = linha[32:37].strip()
                    xl = linha[37:42].strip()
                    t_d = linha[42:47].strip()
                    t_q = linha[47:52].strip()
                    t__d = linha[52:57].strip()
                    t__q = linha[57:62].strip()
                    
                    if x_q == '':
                        polo = 'Polo Saliente'
                    else:
                        polo = 'Polo Liso'
    
                    # Criar a lista com os valores relevantes
                    valores = [polo, xd, xq, x_d, x_q, x__d, xl, t_d, t_q, t__d, t__q]
    
                    # Associar a lista como valor para a chave no dicionário
                    dicionario_dmdg[no] = valores
                else:
                    ra = linha[7:12].strip()
                    h = linha[12:17].strip()
                    d = linha[17:22].strip()
                    mva = linha[22:27].strip()
                    
                    # Criar a lista com os valores relevantes
                    valores = [ra, h, d, mva]
    
                    # Associar itens no valor para a chave no dicionário
                    dicionario_dmdg[no] += valores

# Função personalizada para verificar se o valor da coluna 'Chave' está presente no dicionário
def verificar_chave(valor):
    return valor in dicionario_dmaq

base_gercad_filtrada['Chave'] = base_gercad_filtrada[colunas[2]].astype(int).astype(str) + ' - ' + base_gercad_filtrada[colunas[3]].astype(int).astype(str)

# Aplicar a função personalizada à coluna 'Chave' usando o método apply()
# O resultado será uma Series de valores booleanos, que será usada como máscara para filtrar o DataFrame
mascara = base_gercad_filtrada['Chave'].apply(verificar_chave)

# Filtrar o DataFrame usando a máscara
base_gercad_filtrada = base_gercad_filtrada[mascara]

# Criando a base de dados do anatem
base_anatem = {chave: dicionario_dmdg[valor]+dicionario_bnt1[chave] for chave, valor in dicionario_dmaq.items()}

# Função para converter strings para float, se possível
def converter_para_float(valor):
    try:
        return float(valor)
    except ValueError:
        return valor

difs = []
slicer = 4 # Indica a partir de qual variável deve ser comparado gercad x anatem
# Percorrer as linhas do DataFrame
for index, linha in base_gercad_filtrada.iterrows():
    chave = linha['Chave']

    # Se a chave não estiver no dicionário, eliminar a linha do DataFrame
    if chave not in base_anatem:
        base_gercad_filtrada.drop(index, inplace=True)
    else:
        # Se a chave estiver no dicionário, verificar se os valores são iguais
        valores_gercad = [converter_para_float(linha[col]) for col in colunas[slicer:] if col != 'Chave']
        valores_anatem = [converter_para_float(valor) for valor in base_anatem[chave]]
        
        if valores_gercad == valores_anatem:
            base_gercad_filtrada.drop(index, inplace=True)
        else:
            dif = []
            for i in range(len(valores_anatem)):
                try:
                    dif.append(valores_gercad[i] - valores_anatem[i])
                except (TypeError, ValueError):
                    try:
                        dif.append(float(valores_gercad[i]))
                    except:
                        try:
                            dif.append(-float(valores_anatem[i]))
                        except:
                            if valores_gercad[i] == valores_anatem[i]:
                                dif.append('')
                            else:
                                dif.append(valores_anatem[i])
            difs.append(dif)
            dif_df = pd.DataFrame(difs, columns=colunas[slicer:])
            
            # Se os valores forem diferentes, reescrever os valores no DataFrame
            for col, valor in zip(base_gercad_filtrada[colunas[slicer:]], valores_anatem):
                base_gercad_filtrada.at[index, col] = valor

# Elimina a coluna de chaves
base_gercad_filtrada = base_gercad_filtrada.drop('Chave', axis=1)

# Resetar os índices do DataFrame após as operações
base_gercad_filtrada.reset_index(drop=True, inplace=True)

base_gercad_filtrada.to_excel('Dados_UG_Update.xls', index=False)