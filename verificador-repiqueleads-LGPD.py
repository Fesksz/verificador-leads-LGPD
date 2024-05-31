import glob
import pandas as pd
import numpy as np
import time
from datetime import date
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from credentials import login,password
import xlwings as xw
import os
import win32com.client

chrome_options = Options()
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

extranet_url = "https://extranet.lopesrio.com.br/"
navegador.get(extranet_url)
navegador.maximize_window()

for i in login:
    navegador.find_element(By.NAME,"tLogin").send_keys(i)
    time.sleep(0.2)
navegador.find_element(By.NAME,"tSenha").send_keys(password)
time.sleep(1)
navegador.find_element(By.ID,"btLogin").click()

extranet_url = "https://extranet.lopesrio.com.br/CRM/Clientes.aspx"
navegador.get(extranet_url)

data_hoje = date.today()
data_formatada = data_hoje.strftime("%d/%m/%Y")

navegador.find_element(By.ID,"ctl00_ContentPlaceHolder1_tDtIni").send_keys(data_formatada)
navegador.find_element(By.ID,"ctl00_ContentPlaceHolder1_tDtFim").send_keys(data_formatada)
navegador.find_element(By.ID,"ctl00_ContentPlaceHolder1_btExibir").click()

navegador.implicitly_wait(15)
navegador.find_element(By.XPATH,"/html/body/form/main/div/div[1]/div[3]/div/div[2]/div/div/div[3]/div/div/button").click()
time.sleep(60)
navegador.close()

planilha_leads = xw.Book("E:\\Guilherme Felix\\Downloads\\ExportERP.xls")

data_relatorio_leads = data_hoje.strftime("%y-%m-%d")

caminho_leads_novo = r'M:\Atendimento\Oferta Ativa\Repique\Extranet\Repique Atualizado\Leads'

planilha_leads.save(rf'{caminho_leads_novo}\{data_relatorio_leads}.xlsx')

planilha_leads.app.quit()

os.remove("E:\\Guilherme Felix\\Downloads\\ExportERP.xls")

File = win32com.client.Dispatch("Excel.Application")
File.Visible = 1

relatorio_leads = File.Workbooks.Open(r'M:\Atendimento\Oferta Ativa\Repique\Extranet\Repique Atualizado\Repique Atualizado_1.xlsx')
relatorio_leads.RefreshAll ()
time.sleep (60)
relatorio_leads.Save()
File.Quit()

# Salvando Blacklist

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

# Arquivo para ser verificado

local_do_arquivo = pd.read_excel('M:\\Atendimento\\Oferta Ativa\\Repique\\Extranet\\Repique Atualizado\\Repique Atualizado_1.xlsx')

# Exclusão de contatos no repique
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