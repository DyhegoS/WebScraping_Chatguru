import pandas as pd
from pathlib import Path
from time import sleep
import getpass

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.options import Options
import traceback


option = webdriver.ChromeOptions()
option.add_argument(r"--user-data-dir=C:\Users\dyhego.silva\chrome-selenium-profile")
option.add_argument(r'--profile-directory=Default')
driver = webdriver.Chrome(options=option)
    

driver.get('https://s17.chatguru.app/')
TIME_TO_WAIT = 30

login = driver.find_element(By.ID, 'email')
password = driver.find_element(By.ID, 'password')
login.send_keys('suporteti@7oliveiras.com.br')
my_password = getpass.getpass()
password.send_keys(my_password)
driver.find_element(By.CSS_SELECTOR, 'button.FormButton').click()

try:
    access_users = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="/users"]'))
    )

    access_users.click()
except:
    print('link de usuários não encontrado')


complete_sheet = []

try:
    # Espera o <tbody> carregar
    tbody = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.TAG_NAME, "tbody"))
    )

    line_data_users = tbody.find_elements(By.TAG_NAME, "tr")

    data_users = []
    
    for line in line_data_users:
        columns_users = line.find_elements(By.TAG_NAME, "td")
    
        if len(columns_users) >= 4:
            # NOME: dentro de <a>
            name_users = columns_users[1].find_element(By.TAG_NAME, "a").text
            
            # E-MAIL: dentro de <span> no mesmo <td>
            email_users = columns_users[1].find_element(By.TAG_NAME, "span").text
            
            # STATUS LOGIN: dentro de <span>
            status_login_users = columns_users[5].find_element(By.TAG_NAME, "span").text
            
            # ÚLTIMO ACESSO: texto puro
            ultimo_acesso_users = columns_users[6].text.strip()
            
            # VISTO ÚLTIMA VEZ: texto puro
            visto_ultima_vez_users = columns_users[7].text.strip()
        
        data_users.append({
            'Nome': name_users,
            'E-mail': email_users,
            'Status Login': status_login_users,
            'Último Acesso': ultimo_acesso_users,
            'Visto Última Vez': visto_ultima_vez_users
        })

    df = pd.DataFrame(data_users)
    df.to_excel('relatorio_chatguru_beta.xlsx', index=False)

except WebDriverException as e:
    print("Erro ao coletar dados de usuário e gerar .xlsx")
    print(e)

    print("Stacktrace")
    traceback.print_exc()

try:
    access_chats = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="/chats"]'))
    )

    access_chats.click()
except:
    print('link de chats não encontrado')


try:
    tbody_chats = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.TAG_NAME, "tbody"))
    )

    total_chats = []

    line_data_chats = tbody_chats.find_elements(By.TAG_NAME, 'tr')

    for lines in line_data_chats:
        column_chats = lines.find_elements(By.TAG_NAME, 'td')
        if len(column_chats) >= 5:
            name_chats = column_chats[1].text.strip()
            total = column_chats[5].find_element(By.TAG_NAME, 'button').text
        
        total_chats.append({
            'Nomes': name_chats,
            'Conversas Totais': total,
        })
    
    remove = ['Ninguém Delegado', 'Comercial-3', 'E-commerce', 'Financeiro', 'Logistica', 'Comercial-1', 
              'Motoristas', 'Diretoria', 'Administrativo', 'Comercial-2', 'Compras']
    total_chats = [x for x in total_chats if x['Nomes'] not in remove]

    total_chats = sorted(total_chats, key=lambda x: x['Nomes'])
    totais = [
        int(item['Conversas Totais'].replace('.', '').replace(',', '').strip())
        if item['Conversas Totais'].strip().isdigit()
        else 0
        for item in total_chats
    ]

    df2 = pd.read_excel('relatorio_chatguru_beta.xlsx')

    if len(totais) == len(df2):
        df2.insert(loc=5, column='Total', value=totais)
        df2.to_excel('relatorio_chatguru.xlsx', index=False)
    else:
        print(f"Erro: {len(totais)} totais, mas {len(df2)} linhas no DataFrame.")
    
except WebDriverException as e:
    print("Erro ao coletar dados da conversa e gerar .xlsx")
    print(e)

    print("Stacktrace")
    traceback.print_exc()
finally:
    driver.quit()
    sleep(TIME_TO_WAIT)