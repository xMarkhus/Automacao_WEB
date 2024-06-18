import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service as ChromeService
from time import sleep
from selenium.webdriver.support import expected_conditions as condicao_esperada
from pathlib import Path


def iniciar_driver():
    chrome_options = Options()

    arguments = ['--lang=en-US', '--window-size=1300,1000',
                 '--incognito']

    for argument in arguments:
        chrome_options.add_argument(argument)

    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'profile.default_content_setting_values.notifications': 2,
        'profile.default_content_setting_values.automatic_downloads': 1

    })

    driver = webdriver.Chrome(service=ChromeService(
        ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(
        driver,
        10,
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )
    return driver, wait

script_dir = Path(__file__).resolve().parent
nome_arquivo = 'produtos.xlsx'
arquivo = script_dir/nome_arquivo

df = pd.read_excel(arquivo)

driver, wait = iniciar_driver()
driver.get('https://cadastroprodutos-devaprender.netlify.app/')

for index, row in df.iterrows():
    produto = row['Produto']
    fornecedor = row['Fornecedor']
    categoria = row['Categoria']
    valor_unitario = row['Valor Unitário']
    
    campo_produto = wait.until(condicao_esperada.element_to_be_clickable((By.XPATH, "//input[@id='campo1']")))
    campo_produto.send_keys(produto)
    sleep(1)

    campo_fornecedor = wait.until(condicao_esperada.element_to_be_clickable((By.XPATH, "//input[@id='campo2']")))
    campo_fornecedor.send_keys(fornecedor)
    sleep(1)

    campo_categoria = wait.until(condicao_esperada.element_to_be_clickable((By.XPATH, "//input[@id='campo3']")))
    campo_categoria.send_keys(categoria)
    sleep(1)

    campo_valor = wait.until(condicao_esperada.element_to_be_clickable((By.XPATH, "//input[@id='campo4']")))
    campo_valor.send_keys(valor_unitario)
    sleep(1)

    btn_cadastrar = wait.until(condicao_esperada.element_to_be_clickable((By.XPATH, "//button[@class='btn btn--radius-2 btn--blue']")))
    btn_cadastrar.click()
    sleep(1)

    alerta = driver.switch_to.alert
    alerta.accept()
    sleep(1.5)

driver.close()
