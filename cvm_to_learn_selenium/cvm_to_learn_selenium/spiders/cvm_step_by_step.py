import scrapy
import time
from scrapy_selenium import SeleniumRequest
from scrapy.selector import Selector
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException
from openpyxl import load_workbook

wb = load_workbook('BD_CADASTRO_NUMERADO_AGO.xlsx')
ws, ws2 = wb['Fundos'], wb['Fundos_Cota']
lista_cnpj = []

for row in ws.iter_rows(values_only=True):
    cnpj = row[1]
    lista_cnpj.append(cnpj)

for row in ws2.iter_rows(values_only=True):
    cnpj = row[1] 
    lista_cnpj.append(cnpj)


class CvmSeleniumSpider(scrapy.Spider):
    name = "cvm_step"
    # allowed_domains = ["cvmweb.cvm.gov.br"]
    # start_urls = ["https://cvmweb.cvm.gov.br/swb/default.asp?sg_sistema=fundosreg"]

    def start_requests(self):
        yield SeleniumRequest(
            url='https://cvmweb.cvm.gov.br/swb/default.asp?sg_sistema=fundosreg',
            wait_time=3,
            callback=self.parse,
        )

    def parse(self, response):
        driver = response.request.meta['driver']

        #pesquisa e troca para o frame correto sem precisar declarar uma variável nova no processo
        WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//frame[@name='Main']"))
        )

        search_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@id='txtCNPJNome']"))
        )
        search_input.clear()
        search_input.send_keys(lista_cnpj[0])
        search_input.send_keys(Keys.ENTER)

        find_name_to_store = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[@id='ddlFundos__ctl0_lnkbtn1']"))
        )
        stored_name = find_name_to_store.text
        find_name_to_store.click()

        yield SeleniumRequest(
            url=response.url,
            wait_time=3,
            callback=self.get_dados_diarios,
            meta={'name': stored_name},
            dont_filter=True
        )

    def get_dados_diarios(self, response):
        driver = response.request.meta['driver']
        driver.switch_to.default_content()

        #pesquisa e troca para o frame correto sem precisar declarar uma variável nova no processo
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//frame[@name='Main']"))
        )

        dados_diarios = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[@id='Hyperlink2']"))
        )
        dados_diarios.click()

        driver.switch_to.default_content()

        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//frame[@name='Main']"))
        )

        value_selected = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//select/option/text()"))
        )

        yield {
            'name': response.meta['name'],
            'value': value_selected
        }