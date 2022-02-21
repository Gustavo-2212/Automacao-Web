from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd


def cotacao(moeda, pagina):
    if moeda != 'ouro':
        pagina.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(moeda)
        pagina.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
        cot = pagina.find_element(By.XPATH, '/html/body/div[7]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div/div[1]/div[1]/div[2]/span[1]').get_attribute('data-value')
        pagina.back()
    else:
       cot = pagina.find_element(By.ID, 'comercial').get_attribute('value') 
    
    return cot

def base_dados(path):
    bd = pd.read_excel(path)
    return bd

def main():
    driver = webdriver.Chrome()
    
    # Capturando a cotação do Dolar e do Euro
    driver.get('https://google.com.br/')
    cotacao_dolar = cotacao('cotação dolar', driver)
    cotacao_euro = cotacao('cotação euro', driver)
    
    # Capturando a Cotação do Ouro
    driver.get('https://www.melhorcambio.com/')
    aba_original = driver.window_handles[0]  
    driver.find_element(By.XPATH, "/html/body/div[15]/div[2]/div/table[2]/tbody/tr[2]/td[2]/a/img").click()
    aba_nova = driver.window_handles[1]
    driver.switch_to.window(aba_nova)
    cotacao_ouro = cotacao('ouro', driver)
    driver.quit()
    
    cotacao_dolar = cotacao_dolar.replace(',', '.')
    cotacao_euro = cotacao_euro.replace(',', '.')
    cotacao_ouro = cotacao_ouro.replace(',', '.')
    
    # Importando a base de dados
    bd = base_dados(r'C:/Users/gusta/OneDrive/Área de Trabalho/Programação - UEMG/Python/Programas Diversos/Projeto 07 - Automação Web/Produtos.xlsx')
    
    # Atualizando as cotações
    bd.loc[bd['Moeda']=='Dólar', 'Cotação'] = float(cotacao_dolar)
    bd.loc[bd['Moeda']=='Euro', 'Cotação'] = float(cotacao_euro)
    bd.loc[bd['Moeda']=='Ouro', 'Cotação'] = float(cotacao_ouro)

    # Recalculando Preço Base Reais, Ajuste, Preço Final
    bd['Preço Base Reais'] = bd['Cotação'] * bd['Preço Base Original']
    bd['Preço Final'] = bd['Ajuste'] * bd['Preço Base Reais']
    
    # Exportando o resultado atualizado para uma planilha
    bd.to_excel('Produtos Atualizados.xlsx', index=False)
    
    
if __name__== '__main__':
    main()