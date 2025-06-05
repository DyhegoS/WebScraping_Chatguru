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
import traceback

# Chrome Options
# https://peter.sh/experiments/chromium-command-line-switches/
# Doc Selenium
# https://selenium-python.readthedocs.io/locating-elements.html


ROOT_FOLDER = Path(__file__).parent
CHROME_DRIVER_PATH = ROOT_FOLDER / 'driver' / 'chromedriver.exe'


def make_chrome_browser(*options: str) -> webdriver.Chrome:
    chrome_options = webdriver.ChromeOptions()

    # chrome_options.add_argument('--headless')
    if options is not None:
        for option in options:
            chrome_options.add_argument(option)

    chrome_service = Service(
        executable_path=str(CHROME_DRIVER_PATH),
    )

    browser = webdriver.Chrome(
        service=chrome_service,
        options=chrome_options
    )

    return browser


if __name__ == '__main__':
    TIME_TO_WAIT = 700

    options = '--disable-gpu',
    browser = make_chrome_browser(*options)

    browser.get('https://s17.chatguru.app/')

    login = browser.find_element(By.ID, 'email')
    password = browser.find_element(By.ID, 'password')
    login.send_keys('suporteti@7oliveiras.com.br')
    my_password = getpass.getpass()
    password.send_keys(my_password)
    browser.find_element(By.CSS_SELECTOR, 'button.FormButton').click()
try:
    access_users = WebDriverWait(browser, 200).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="/users"]'))
    )

    access_users.click()
except:
    print('link de usuários não encontrado')


complete_sheet = []

try:
    # Espera o <tbody> carregar
    tbody = WebDriverWait(browser, 300).until(
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

except WebDriverException as e:
    print("Erro ao coletar dados de usuário e gerar .xlsx")
    print(e)

    print("Stacktrace")
    traceback.print_exc()

try:
    access_chats = WebDriverWait(browser, 400).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="/chats"]'))
    )

    access_chats.click()
except:
    print('link de chats não encontrado')


try:
    tbody_chats = WebDriverWait(browser, 500).until(
        EC.presence_of_element_located((By.TAG_NAME, "tbody"))
    )

    # pega dado de conversas totais
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

    df = pd.DataFrame(data_users)

    total_chats = sorted(total_chats, key=lambda x: x['Nomes'])
    totais = [item['Conversas Totais'] for item in total_chats]
    
    df['Totais'] = totais
    df.to_excel('dados_extraidos.xlsx', index=False)
except WebDriverException as e:
    print("Erro ao coletar dados da conversa e gerar .xlsx")
    print(e)

    print("Stacktrace")
    traceback.print_exc()
finally:
    browser.quit()
    sleep(TIME_TO_WAIT)