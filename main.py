import os
import time
from datetime import date

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By

import hidden_data
from config import ConfigGoogleApi, ConfigSelenium

config_selenium = ConfigSelenium()
config_google_api = ConfigGoogleApi()

options = config_selenium.chrome_options
chrome_service = config_selenium.service
driver = webdriver.Chrome(options=options, service=chrome_service)
max_attempts = 3

for attempt in range(max_attempts):
    try:
        driver.get("http://g4.transpotech.com.br/transpotech")

        # Login
        username = driver.find_element(By.NAME, "j_username")
        password = driver.find_element(By.NAME, "j_password")
        login_button = driver.find_element(By.ID, "logar")
        username.send_keys(hidden_data.username)
        password.send_keys(hidden_data.password)
        login_button.click()

        # Access Atendimentos Page
        time.sleep(2)
        atendimentos = driver.find_elements(By.CLASS_NAME, "label_menu")[2]
        atendimentos.click()
        atendimento_selection = driver.find_element(By.ID, "itemMenuAtendimento")
        atendimento_selection.click()

        # Adding filters
        driver.find_element(By.XPATH, '//*[@id="filtrosOS"]/h3').click()

        # Tipo de Serviço
        driver.find_element(By.ID, 'selectTipoAtividade_ms').click()
        driver.find_element(By.XPATH, '/html/body/div[13]/div/ul/li[2]/a/span[2]').click()
        driver.find_element(By.ID, 'ui-multiselect-1-selectTipoAtividade-option-24').click()
        driver.find_element(By.ID, 'ui-multiselect-1-selectTipoAtividade-option-25').click()
        driver.find_element(By.ID, 'ui-multiselect-1-selectTipoAtividade-option-26').click()
        driver.find_element(By.ID, 'ui-multiselect-1-selectTipoAtividade-option-27').click()
        driver.find_element(By.ID, 'ui-multiselect-1-selectTipoAtividade-option-28').click()
        driver.find_element(By.ID, 'ui-multiselect-1-selectTipoAtividade-option-29').click()
        driver.find_element(By.ID, 'selectTipoAtividade_ms').click()

        # Filial
        driver.find_element(By.ID, 'selectUnidades_ms').click()
        driver.find_element(By.XPATH, '/html/body/div[15]/div/ul/li[2]/a/span[1]').click()
        driver.find_element(By.ID, 'ui-multiselect-3-selectUnidades-option-2').click()
        driver.find_element(By.ID, 'selectUnidades_ms').click()

        # Status
        driver.find_element(By.ID, 'selectStatusAtendimento_ms').click()
        driver.find_element(By.XPATH, '/html/body/div[20]/div/ul/li[2]/a/span[1]').click()
        driver.find_element(By.ID, 'ui-multiselect-8-selectStatusAtendimento-option-6').click()
        driver.find_element(By.ID, 'ui-multiselect-8-selectStatusAtendimento-option-7').click()
        driver.find_element(By.ID, 'selectStatusAtendimento_ms').click()

        # Data Conclusão
        input_date = config_selenium.date_range
        driver.find_element(By.ID, 'conclusaoInicial').send_keys(input_date)
        driver.find_element(By.XPATH, '//*[@id="filtrosAtendimento_content"]/div[7]/div[1]/img').click()

        # Search
        search_button = driver.find_element(By.XPATH, '//*[@id="filtros"]/div[3]/div[1]/button/span')
        search_button.click()
        time.sleep(5)

        # Exportar
        driver.find_element(By.XPATH, '//*[@id="exportar"]/button').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="tipoExportacao"]/div/div[4]/button[1]').click()

        break

    except WebDriverException as e:
        print(f'Attempt {attempt} failed: {e}')

        if driver:
            driver.quit()

        time.sleep(2)

else:
    print('Max attempts reached')

# Read excel file
path = config_selenium.folder
lista = os.listdir(path)
while len(lista) == 0 or lista[0].endswith('.crdownload') or lista[0].endswith('.tmp'):
    lista = os.listdir(path)
    time.sleep(1)
time.sleep(3)
excel_path = os.path.join(path, lista[0])
exportacao = pd.read_excel(excel_path)

# Getting data from the spreadsheet
spreadsheet_id = config_google_api.spreadsheet_id
range = config_google_api.range
service = config_google_api.service
result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                             range='HISTORICO_DATA!D:D',
                                             ).execute()

# Data handling
t_exportacao = exportacao.transpose()
atendimentos = result['values']
delta = 0

for i, item in enumerate(atendimentos):
    atendimentos[i] = item[0]

for row in t_exportacao:
    try:
        filial = t_exportacao[row]['FILIAL']
        os_apollo = t_exportacao[row]['CÓDIGO OS APOLLO']
        data_abertura_os = t_exportacao[row]['DATA ABERTURA OS']
        os_g4 = t_exportacao[row]['CÓDIGO OS G4']
        cod_naoprog = t_exportacao[row]['CÓDIGO NÃO PROGRAMADO']
        equipe_atend = t_exportacao[row]['EQUIPE ATENDIMENTO']
        razao_soc = t_exportacao[row]['RAZÃO SOCIAL']
        num_serie = t_exportacao[row]['NÚMERO SÉRIE']
        frota = t_exportacao[row]['FROTA']
        horimetro = t_exportacao[row]['HORÍMETRO']
        tipo_man = t_exportacao[row]['TIPO DE MANUTENÇÃO']
        tipo_oper = t_exportacao[row]['TIPO DE OPERAÇÃO']
        status_os = t_exportacao[row]['STATUS OS']
        status_atendimento = t_exportacao[row]['STATUS ATENDIMENTO']
        intervencao = t_exportacao[row]['INTERVENÇÃO']
        tecnico = t_exportacao[row]['NOME TÉCNICO']
        data_trab = t_exportacao[row]['DATA TRABALHO']
        duracao_ida = t_exportacao[row]['DURAÇÃO IDA']
        duracao_trab = t_exportacao[row]['DURAÇÃO TRABALHO']
        duracao_volta = t_exportacao[row]['DURAÇÃO VOLTA']
        avaliacao = t_exportacao[row]['AVALIAÇÃO']
        pendencia = t_exportacao[row]['PENDÊNCIA']
        comentario = t_exportacao[row]['COMENTÁRIO DO TÉCNICO']
        status_equip = t_exportacao[row]['STATUS DO EQUIPAMENTO']
        km = t_exportacao[row]['KM UTILZADO NO ATENDIMENTO']
        if pd.isna(duracao_ida):
            duracao_ida = ""
        if pd.isna(frota):
            frota = ""
        if pd.isna(duracao_volta):
            duracao_volta = ""
        if pd.isna(equipe_atend):
            equipe_atend = ""
        if pd.isna(cod_naoprog):
            cod_naoprog = ""

        if os_g4 in atendimentos:
            continue

        # writing the data into the sheet
        values = [
            [
                filial, os_apollo, pd.to_datetime(data_abertura_os, format="%d/%m/%Y %H:%M").strftime("%m/%d/%Y %H:%M:%S"), os_g4, cod_naoprog, equipe_atend, razao_soc, num_serie, frota, horimetro,
                tipo_man, tipo_oper, status_os, status_atendimento, intervencao, tecnico, pd.to_datetime(data_trab, format="%d/%m/%Y %H:%M").strftime("%m/%d/%Y %H:%M:%S"), duracao_ida, duracao_trab,
                duracao_volta, avaliacao, pendencia, comentario, status_equip, km
            ],
        ]
        body = {'values': values}
        range_append = 'HISTORICO_DATA!A1'
        result = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id,
                                                        range=range_append,
                                                        valueInputOption='USER_ENTERED',
                                                        body=body
                                                        ).execute()

        if tipo_man == "INSPEÇÃO PREVENTIVA":
            delta += 1

        time.sleep(1)
    except Exception as e:
        print(e)
        continue

# Writing log data into the sheet
try:
    log = [
        [
            str(date.today()), delta
        ]
    ]
    body = {'values': log}
    log_range = 'LOG!A1'
    result = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id,
                                                    range=log_range,
                                                    valueInputOption='USER_ENTERED',
                                                    body=body).execute()
except Exception as e:
    print(e)

driver.quit()

print("Script Completed")
