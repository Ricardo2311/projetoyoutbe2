import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from time import sleep
# Funcao excel
# abrir o excel
# Copiar a célula
# funcao site
email = input('Digite seu email: ')
senha = input('Digite sua senha: ')
# Abrir o site
driver = webdriver.Chrome()
driver.get('https://contabilidade-devaprender.netlify.app')
sleep(5)
# logar no site
campo_email = driver.find_element(By.XPATH, "//input[@id='email']")
sleep(1)
campo_email.send_keys(email)
sleep(1)
campo_senha = driver.find_element(By.XPATH, "//input[@id='senha']")
sleep(1)
campo_senha.send_keys(senha)
sleep(1)
botao_entrar = driver.find_element(By.XPATH, "//button[@id='Entrar']")
sleep(1)
botao_entrar.click()
sleep(5)
# Colar a célula no campo certo

planilha = openpyxl.load_workbook('./empresas.xlsx')
pagina_dados = planilha['dados empresas']

for linha in pagina_dados.iter_rows(min_row=2, values_only=True):
    nome_empresa = linha[0]
    email_empresa = linha[1]
    tel_empresa = linha[2]
    endereco_empresa = linha[3]
    cnpj_empresa = linha[4]
    area_atuacao = linha[5]
    numero_funcionarios = linha[6]
    data_fundacao = linha[7]
    nome = driver.find_element(By.XPATH, "//input[@id='nomeEmpresa']")
    sleep(1)
    nome.send_keys(nome_empresa)
    sleep(1)
    email = driver.find_element(By.XPATH, "//input[@id='emailEmpresa']")
    sleep(1)
    email.send_keys(email_empresa)
    sleep(1)
    telefone = driver.find_element(By.XPATH, "//input[@id='telefoneEmpresa']")
    sleep(1)
    telefone.send_keys(tel_empresa)
    sleep(1)
    endereco = driver.find_element(
        By.XPATH, "//textarea[@id='enderecoEmpresa']")
    sleep(1)
    endereco.send_keys(endereco_empresa)
    sleep(1)
    cnpj = driver.find_element(By.XPATH, "//input[@id='cnpj']")
    sleep(1)
    cnpj.send_keys(cnpj_empresa)
    sleep(1)
    area = Select(driver.find_element(By.XPATH, "//select[@id='areaAtuacao']"))
    sleep(1)
    try:
        area.select_by_visible_text(area_atuacao)
    except:
        area.select_by_visible_text('Selecione...')
    sleep(1)
    
    numero = driver.find_element(By.XPATH, "//input[@id='numeroFuncionarios']")
    sleep(1)
    numero.send_keys(numero_funcionarios)
    sleep(1)
    data = driver.find_element(By.XPATH, "//input[@id='dataFundacao']")
    sleep(1)
    data.send_keys(data_fundacao)
    sleep(1)
    botao_cadastrar = driver.find_element(By.XPATH, "//button[@id='Cadastrar']")
    sleep(1)
    botao_cadastrar.click()
    sleep(2)