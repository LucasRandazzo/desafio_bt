import pandas as pd
from utils import (iniciar_navegador, 
                   coletar_dados_empresas_inicial, 
                   coletar_dados_empresas_final, 
                   salvar_planilha_atualizada,
                   obter_total_categorias)

def main():
    total_categorias = obter_total_categorias()
    planilha = pd.DataFrame(columns=[
        'Categoria', 'Empresa', 'Reclamacoes Respondidas', 'Nota do Consumidor','Voltariam a Fazer Negocio', 'Indice de Solucao'])
    navegador = iniciar_navegador()
    navegador.get("https://www.reclameaqui.com.br/")
    

    lista_empresas = coletar_dados_empresas_inicial(navegador, total_categorias)
    navegador.close()

    lista_empresas = coletar_dados_empresas_final(lista_empresas)

    salvar_planilha_atualizada(planilha, lista_empresas)

if __name__ == '__main__':
    main()