import glob
import pandas as pd
import numpy as np

Blacklist_df = pd.read_excel(r'M:\Atendimento\Oferta Ativa\MK Bairros\Importante\01 - Contatos excluídos.xlsx', sheet_name='Blacklist')

def lista_blacklist(nome_coluna, data_frame):
    valores_coluna = data_frame[nome_coluna]
    valores_coluna = valores_coluna.dropna()
    valores_coluna = valores_coluna.apply(str)

    blacklist = []
    for i in valores_coluna:
        l = i.replace('.0', '')
        if l != '0':
            blacklist.append(l)
    
    return blacklist

coluna1 = lista_blacklist('Tel 1', Blacklist_df)
coluna2 = lista_blacklist('Tel 2', Blacklist_df)
coluna3 = lista_blacklist('Tel 3', Blacklist_df)
coluna4 = lista_blacklist('Tel 4', Blacklist_df)
coluna5 = lista_blacklist('Cel1', Blacklist_df)

headhunter = coluna1 + coluna2 + coluna3 + coluna4 + coluna5

local_do_arquivo = pd.read_excel('M:\\Atendimento\\Oferta Ativa\\Repique\\Extranet\\Repique Atualizado\\Repique Atualizado_1.xlsx')

for contato in headhunter:

    local_do_arquivo = local_do_arquivo.astype(str)
    tel1 = local_do_arquivo[local_do_arquivo['TELEFONE'].str.contains(str(contato), case=False)]
    
    if not tel1.empty:
        print(f"Excluindo contato {contato}...")
        local_do_arquivo.loc[local_do_arquivo['TELEFONE'].str.contains(str(contato), case=False), 'TELEFONE'] = np.nan
        local_do_arquivo = local_do_arquivo.dropna(subset=['TELEFONE'])

caminho_novo = r'M:\Atendimento\Oferta Ativa\Repique\Extranet\Repique Atualizado\Lista_Repique_Limpa.xlsx'
local_do_arquivo.to_excel(caminho_novo, index=False)

print("Listagem de Repique Limpa")