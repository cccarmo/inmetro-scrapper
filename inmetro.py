from urllib.request import urlopen
from bs4 import BeautifulSoup
import codecs
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By


# Open main page
url = "http://www.inmetro.gov.br/prodcert/certificados/busca.asp"
service = Service('C:/Users/RafaelBatistaCarmo/Documents/Personal/inmetro/chromedriver.exe')
driver = webdriver.Chrome(service=service)
driver.get(url)

# Select "Classe de Produto"
d = Select(driver.find_element(By.NAME, "classe_produto"))
d.select_by_value('6')

# Click "Buscar"
search_button = driver.find_element(By.NAME, "btn_enviar")
search_button.click()


workbook = xlsxwriter.Workbook('teste.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, 'Informações gerais')
worksheet.write(0, 1, 'CNPJ/CPF')
worksheet.write(0, 2, 'Razão Social / Nome (PF)')
worksheet.write(0, 3, 'Nome fantasia')
worksheet.write(0, 4, 'Endereço')
worksheet.write(0, 5, 'Status')
worksheet.write(0, 6, 'Papel da empresa')
worksheet.write(0, 7, 'Marca')
worksheet.write(0, 8, 'Modelo')
worksheet.write(0, 9, 'Importado')
worksheet.write(0, 10, 'Descrição')


row = 1
col = 0
basic_info = 0
informacoes_gerais = None
cnpj = None
razao_social = None
nome_fantasia = None
endereco = None
status = None
papel_da_empresa = None


for i in range(38):

    # Wait for the load
    driver.implicitly_wait(30)
    # Gambiarra pra esperar
    page_lenght = len(driver.find_elements(By.CLASS_NAME, "listagem"))
    while page_lenght < 30 and i < 37:
        driver.implicitly_wait(5)
        page_lenght = len(driver.find_elements(By.CLASS_NAME, "listagem"))

    for a in driver.find_elements(By.CLASS_NAME,"listagem"):
        if  'Certificador:' in a.text:
            print(a.text)
            informacoes_gerais = a.text
            basic_info = 0
            continue
        elif basic_info == 0:
            basic_info += 1
            cnpj = a.text
        elif basic_info == 1:
            basic_info += 1
            razao_social = a.text
        elif basic_info == 2:
            basic_info += 1
            nome_fantasia = a.text
        elif basic_info == 3:
            basic_info += 1
            endereco = a.text
        elif basic_info == 4:
            basic_info += 1
            status = a.text
        elif basic_info == 5:
            basic_info += 1
            papel_da_empresa = a.text
        elif basic_info == 6:
            basic_info += 1
            worksheet.write(row, 0, informacoes_gerais)
            worksheet.write(row, 1, cnpj)
            worksheet.write(row, 2, razao_social)
            worksheet.write(row, 3, nome_fantasia)
            worksheet.write(row, 4, endereco)
            worksheet.write(row, 5, status)
            worksheet.write(row, 6, papel_da_empresa)
            col = 7
            worksheet.write(row, col, a.text)
        elif basic_info > 6 and basic_info < 9:
            basic_info += 1
            col += 1
            worksheet.write(row, col, a.text)
        elif basic_info == 9:
            col += 1
            worksheet.write(row, col, a.text)
            row += 1
            basic_info = 6

    next_button = driver.find_element(By.XPATH,  "//*[contains(text(),'Próximo >>')]")
    next_button.click()

workbook.close()
