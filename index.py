from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import math
import time
import pandas as pd
navegador = webdriver.Chrome()

#guardas os intens inseridos e os que deram erro
erros=[]
inseridos=[]

planilha_medicamentos = './insumos.xlsx'

planilha_insumos = '.planilhas/'

dados = pd.read_excel(planilha_medicamentos)

navegador.get('https://juizdefora-mg.vivver.com/amx/entrada_direta_produto')

acoes = ActionChains(navegador)

while len(navegador.find_elements(By.XPATH,'//*[@id="amx_entrada_direta_produto_search"]')) < 1:
    time.sleep(1)

#botao insert
navegador.find_element(By.XPATH,'//*[@id="amx_entrada_direta_produto_insert"]').click()
time.sleep(0.5)

#campo motivo de entrada
navegador.find_element(By.XPATH,'//*[@id="lookup_key_amx_entrada_direta_produto_codmotivoentr"]').click()
navegador.find_element(By.XPATH,'//*[@id="lookup_key_amx_entrada_direta_produto_codmotivoentr"]').send_keys('9')
time.sleep(0.5)

def apagarTudo():
    acoes.send_keys(Keys.BACKSPACE).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()



for i, row in dados.iterrows():
    produto = row['descricao']
    qtd = row['estoque']

    if qtd == 0 or qtd == '0' or math.isnan(qtd) or qtd == '' :
        continue
 
    #campo produto
    navegador.find_element(By.XPATH,'//*[@id="s2id_amx_entrada_direta_produto_codproduto"]').click()
    time.sleep(0.1)
    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen10_search"]').click()
    time.sleep(0.1)
    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen10_search"]').send_keys(produto)
    time.sleep(0.5)
    acoes.send_keys(Keys.TAB).perform()
    time.sleep(0.1)

    #campo fabricante
    navegador.find_element(By.XPATH, '//*[@id="lookup_key_amx_entrada_direta_produto_codfabricante"]').click()
    time.sleep(0.1)
    apagarTudo()
    time.sleep(0.1)
    navegador.find_element(By.XPATH, '//*[@id="lookup_key_amx_entrada_direta_produto_codfabricante"]').send_keys('396')
    time.sleep(0.1)

    #campo lote
    navegador.find_element(By.XPATH,'//*[@id="amx_entrada_direta_produto_numlote"]').click()
    apagarTudo()
    navegador.find_element(By.XPATH,'//*[@id="amx_entrada_direta_produto_numlote"]').send_keys('123456')
    time.sleep(0.1)

    #data de vencimento
    navegador.find_element(By.XPATH,'//*[@id="amx_entrada_direta_produto_datvalidade"]').click()
    time.sleep(0.5)
    navegador.find_element(By.XPATH,'//*[@id="amx_entrada_direta_produto_datvalidade"]').send_keys('01/02/2025')
    time.sleep(0.5)

    #quantidade
    navegador.find_element(By.XPATH,'//*[@id="amx_entrada_direta_produto_qdeentrada"]').click()
    apagarTudo()
    navegador.find_element(By.XPATH,'//*[@id="amx_entrada_direta_produto_qdeentrada"]').send_keys(int(qtd))
    time.sleep(0.1)

    #seta para baixo salvar
    navegador.find_element(By.XPATH,'//*[@id="btn_add_item_entrada_direta_produto"]').click()
    time.sleep(0.5)

    #verifica se tem alerta de erro
    if len(navegador.find_elements(By.XPATH,'//*[@id="new_amx_entrada_direta_produto"]/div[2]/div[2]/div[2]/div/div[1]/button')) > 0:
        time.sleep(0.5)
        erros.append(produto)
        navegador.find_element(By.XPATH, '//*[@id="new_amx_entrada_direta_produto"]/div[2]/div[2]/div[2]/div/div[1]/button').click()
        time.sleep(0.5  )
    else:
        inseridos.append(produto)


#sucesso
print(f'Medicamentos cadastrados com sucesso: {len(inseridos)}\n --------------------------------------')
for i in inseridos:
    print(i)

#erro
print(f'Medicamentos com ERRO: {len(erros)}\n --------------------------------------')
for i in erros:
    print(i)


