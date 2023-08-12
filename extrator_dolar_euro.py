import pyautogui as py
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
import subprocess

#Escape
py.PAUSE = 0.5
FAILSAFE = True

#Configuração padrão para utilização do Selenium
#Baixar última versão do driver do navegador
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

#Entrar no navegador
driver.get('https://www.google.com')

#Maximizar o navegador
driver.maximize_window()

#Tempo de espera
wait = WebDriverWait(driver, 10)

#Pesquisar no Google o valor do Dolar
wait.until(EC.visibility_of_element_located(('xpath', '//*[@id="APjFqb"]'))).send_keys('Dolar')
py.hotkey('enter')

#Extrair o valor atual do Dolar
valor_dolar = driver.find_element('xpath','//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').text

#Limpar a pesquisa e pesquisar o Euro
driver.find_element('xpath', '//*[@id="APjFqb"]').send_keys('')
py.press('tab')
py.press('enter')
driver.find_element('xpath', '//*[@id="APjFqb"]').send_keys('Euro')
py.hotkey('enter')

#Extrair o valor atual do Euro
valor_euro = driver.find_element('xpath','//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').text

caminho_arquivo = '/home/will/Estudos/Valor dolar euro/valor_dolar_euro.xlsx'
planilha_criada = xlsxwriter.Workbook(caminho_arquivo)
sheet1 = planilha_criada.add_worksheet()

#Abro o arquivo
subprocess.run(['xdg-open', caminho_arquivo])

#Substituir a vírgula por ponto
valor_dolar = valor_dolar.replace(',', '.')
valor_euro = valor_euro.replace(',', '.')

#Converter os valores para o tipo Float
valor_dolar_float = float(valor_dolar)
valor_euro_float = float(valor_euro)

#Escrever nas celulas
sheet1.write('A1', 'Dolar')
sheet1.write('B1', 'Euro')
sheet1.write('A2', valor_dolar_float)
sheet1.write('B2', valor_euro_float)

#fechando a planilha
planilha_criada.close()

py.alert('O BOT foi executado com exito!') 
driver.quit()
