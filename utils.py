import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
import argparse
import pandas as pd
from selenium import webdriver
from empresa import Empresa



def iniciar_navegador():
    options = Options()
    options.page_load_strategy = 'eager'
    navegador = webdriver.Chrome(service=Service(), options=options)
    navegador.maximize_window()
    return navegador

def clicar_seletor_categorias(navegador):
    botao_selecionar_categoria = navegador.find_element(By.CSS_SELECTOR, ".text-steel.text-sm.py-4.w-full.border-none.focus\\:outline-none.cursor-pointer")
    botao_selecionar_categoria.click()

def validar_total_categorias(valor):
    valor = int(valor)
    if 1 <= valor <= 16:
        return valor
    raise argparse.ArgumentTypeError("O número de categorias deve estar entre 1 e 16.")

def obter_total_categorias():
    parser = argparse.ArgumentParser(description="Define o número de categorias entre 1 e 16.")
    parser.add_argument("--total_categorias", type=validar_total_categorias, default=1, help="Número de categorias (entre 1 e 16, padrão: 1)")
    
    args = parser.parse_args()
    return args.total_categorias

def coletar_dados_empresas_inicial(navegador, total_categorias):
    lista_empresas = []
    espera = WebDriverWait(navegador, 3)

    botao_piores = navegador.find_element(By.CSS_SELECTOR, "[data-testid='tab-worst']")
    botao_melhores = navegador.find_element(By.CSS_SELECTOR, "[data-testid='tab-best']")

    clicar_seletor_categorias(navegador)

    for i in range(total_categorias):
        lista_categoria_moda = navegador.find_elements(By.XPATH, "//span[@title='Moda']/following-sibling::ul[1]/li")
        categoria_atual = lista_categoria_moda[i]
        espera.until(EC.visibility_of(categoria_atual))
        botao_categoria = categoria_atual.find_element(By.TAG_NAME, "button")
        espera.until(EC.element_to_be_clickable(botao_categoria))
        navegador.execute_script("arguments[0].scrollIntoView({block: 'center'})", botao_categoria)
        
        botao_categoria.click()
        categoria_nome = botao_categoria.text

        try:
            ranking_melhores = espera.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".list.svelte-lzrvt6 a")))
        except TimeoutException:
            ranking_melhores = []
            botao_piores.click()

        if ranking_melhores:
            for j in range(len(ranking_melhores)):
                melhor = ranking_melhores[j]
                espera.until(EC.visibility_of(melhor))
                url_melhor = melhor.get_attribute("href")
                nome_empresa = melhor.find_element(By.CSS_SELECTOR, ".text-sm.text-steel.font-semibold.mr-2").text

                empresa = Empresa(nome=nome_empresa, categoria=categoria_nome, url=url_melhor)
                lista_empresas.append(empresa)

                if j == len(ranking_melhores) - 1:
                    botao_piores.click()

        try:
            ranking_piores = espera.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".list.svelte-lzrvt6 a")))
        except TimeoutException:
            ranking_piores = []

        if ranking_piores:
            for k in range(len(ranking_piores)):
                pior = ranking_piores[k]
                espera.until(EC.visibility_of(pior))
                url_pior = pior.get_attribute("href")
                nome_empresa = pior.find_element(By.CSS_SELECTOR, ".text-sm.text-steel.font-semibold.mr-2").text

                empresa = Empresa(nome=nome_empresa, categoria=categoria_nome, url=url_pior)
                lista_empresas.append(empresa)

                if k == len(ranking_piores) - 1:
                    botao_melhores.click()

        clicar_seletor_categorias(navegador)

    return lista_empresas

def coletar_dados_empresas_final(lista_empresas):
    for idx, empresa in enumerate(lista_empresas):
        navegador = iniciar_navegador()
        navegador.get(empresa.url)
        espera = WebDriverWait(navegador, 3)
        dados_necessarios = espera.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".go2549335548 strong")))
        empresa.reclamacoes_respondidas = dados_necessarios[1].text.split()[0]
        empresa.nota_consumidor = dados_necessarios[4].text[:-1]
        empresa.voltariam_fazer_negocio = dados_necessarios[5].text.split()[0]
        empresa.indice_solucao = dados_necessarios[6].text.split()[0]
        print(f"Dados da empresa {idx+1} coletados, ainda restam {len(lista_empresas)-1 - idx} empresas. Aguarde por favor...")
        navegador.close()
    return lista_empresas

def ajustar_largura_colunas(arquivo_excel, nome_pagina='pagina1'):
    planilha = openpyxl.load_workbook(arquivo_excel)
    pagina = planilha[nome_pagina]
    
    for coluna in pagina.columns:
        maior_tamanho = 0
        letra_coluna = coluna[0].column_letter  
        for celula in coluna:
            if celula.value:
                tamanho_celula = len(str(celula.value))
                if tamanho_celula > maior_tamanho:
                    maior_tamanho = tamanho_celula
        largura_ajustada = maior_tamanho + 2
        pagina.column_dimensions[letra_coluna].width = largura_ajustada
    
    planilha.save(arquivo_excel)

def salvar_planilha_atualizada(planilha, lista_empresas):
    print("Salvando dados na planilha...")
    dados = []
    for empresa in lista_empresas:
        dados.append([empresa.categoria, empresa.nome, empresa.reclamacoes_respondidas, empresa.nota_consumidor, empresa.voltariam_fazer_negocio, empresa.indice_solucao])
    planilha_novas_linhas = pd.DataFrame(dados, columns=planilha.columns)
    planilha_final = pd.concat([planilha, planilha_novas_linhas])
    planilha_final.to_excel('dados_empresas.xlsx', sheet_name='pagina1', index=False)
    ajustar_largura_colunas('dados_empresas.xlsx', 'pagina1')
    print("Dados salvos com sucesso!")

