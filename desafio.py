from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import openpyxl

import time


class Empresa:
    def __init__ (self, nome=None, nota=None, categoria=None, reclamacoesRespondidas=None, voltariamFazerNegocio=None, indiceSolucao=None, notaConsumidor=None):
        self.nome = nome
        self.nota = nota
        self.categoria = categoria
        self.reclamacoesRespondidas = reclamacoesRespondidas
        self.voltariamFazerNegocio = voltariamFazerNegocio
        self.indiceSolucao = indiceSolucao
        self.notaConsumidor = notaConsumidor


planilha = pd.DataFrame(columns=['Categoria', 'Empresa', 'Nota', 'Reclamacoes Respondidas', 'Nota Consumidor', 'Voltariam Fazer Negocio', 'Indice Solucao'])
novas_linhas = []        
lista_url_empresas = []      
navegador = webdriver.Chrome()
navegador.get("https://www.reclameaqui.com.br/")
navegador.maximize_window()
botao_piores = navegador.find_element(By.CSS_SELECTOR, "[data-testid='tab-worst']")
botao_melhores = navegador.find_element(By.CSS_SELECTOR, "[data-testid='tab-best']")

botao_selecionar_categoria = navegador.find_element(By.CSS_SELECTOR, ".text-steel.text-sm.py-4.w-full.border-none.focus\\:outline-none.cursor-pointer")
botao_selecionar_categoria.click()

espera = WebDriverWait(navegador, 3)

lista_categoria_moda = navegador.find_elements(By.XPATH, "//span[@title='Moda']/following-sibling::ul[1]/li")

for i in range(len(lista_categoria_moda)):
    lista_categoria_moda = navegador.find_elements(By.XPATH, "//span[@title='Moda']/following-sibling::ul[1]/li")
    categoria = lista_categoria_moda[i]
    espera.until(EC.visibility_of(categoria))
    botao_categoria_nome = categoria.find_element(By.TAG_NAME, "button")
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'})", botao_categoria_nome)
    espera.until(EC.element_to_be_clickable(botao_categoria_nome))
    botao_categoria_nome.click()
    categoria_nome = botao_categoria_nome.text

    try:
        ranking_melhores = espera.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".list.svelte-lzrvt6 a")))
    except TimeoutException:
        ranking_melhores = []
        botao_piores.click()
        
    
    if ranking_melhores:
        for j in range(len(ranking_melhores)):
            empresa = Empresa()
            melhor = ranking_melhores[j] 
            espera.until(EC.visibility_of(melhor))
            
            url_melhor = melhor.get_attribute("href")
            lista_url_empresas.append(url_melhor)
            
            empresa.nome = melhor.find_element(By.CSS_SELECTOR, ".text-sm.text-steel.font-semibold.mr-2").text
            empresa.categoria = categoria_nome

            nova_linha = [empresa.categoria, empresa.nome]
            novas_linhas.append(nova_linha)

            # print(f"Categoria: {empresa.categoria}\nNome: {empresa.nome}\nNota: {empresa.nota}")
            

            if j == len(ranking_melhores) - 1:
                botao_piores.click()

    try:
        ranking_piores = espera.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".list.svelte-lzrvt6 a")))
    except TimeoutException:
        ranking_piores = []
    
    if ranking_piores:
        for k in range(len(ranking_piores)):
            pior = ranking_piores[k]
            empresa = Empresa()
            espera.until(EC.visibility_of(pior))

            url_pior = pior.get_attribute("href")
            lista_url_empresas.append(url_pior)
            
            empresa.nome = pior.find_element(By.CSS_SELECTOR, ".text-sm.text-steel.font-semibold.mr-2").text
            empresa.categoria = categoria_nome

            nova_linha = [empresa.categoria, empresa.nome]
            novas_linhas.append(nova_linha)

            
            # print(f"Categoria: {empresa.categoria}\nNome: {empresa.nome}\nNota: {empresa.nota}")
            
            if k == len(ranking_piores) - 1:
                botao_melhores.click()
              
    print(f"Iteração {i}, Última posição esperada: {len(lista_categoria_moda) - 1}")
    if i != len(lista_categoria_moda) - 1:
        botao_selecionar_categoria.click()

navegador.close()

for i in range (len(lista_url_empresas)):
    navegador = webdriver.Chrome()
    navegador.get(lista_url_empresas[i])
    espera = WebDriverWait(navegador, 3)
    dados_necessarios = espera.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".go2549335548 strong")))
    empresa = Empresa()
    empresa.reclamacoesRespondidas = dados_necessarios[1].text
    empresa.notaConsumidor = dados_necessarios[4].text
    empresa.voltariamFazerNegocio = dados_necessarios[5].text
    empresa.indiceSolucao = dados_necessarios[6].text
    nova_linha = [empresa.reclamacoesRespondidas, empresa.notaConsumidor, empresa.voltariamFazerNegocio, empresa.indiceSolucao]
    novas_linhas[i].extend(nova_linha)
    print(novas_linhas[i])
    print(f"Reclamacoes Respondidas: {empresa.reclamacoesRespondidas}\nNota: {empresa.notaConsumidor}\nVoltariam Negocio: {empresa.voltariamFazerNegocio} \nIndice Solucao: {empresa.indiceSolucao}")
    print (f"Coletando dados da empresa {i+1}, ainda restam {len(lista_url_empresas) - i} para serem coletadas")
    navegador.close()


planilha_novas_linhas = pd.DataFrame(novas_linhas, columns=planilha.columns)
planilha = pd.concat([planilha, planilha_novas_linhas])
planilha.to_excel('pandas_to_excel.xlsx', sheet_name='new_sheet_name', index=False)


# go2549335548

# 1 RECLAMACOES RESPONDIDAS 
# 3 NOTA DO CONSUMIDOR
# 4 VOLTARIAM A FAZER NEGOCIO
# 5 INDICE DE SOLUÇÃO