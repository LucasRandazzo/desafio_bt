class Empresa:
    def __init__(self, nome=None, categoria=None, url=None, reclamacoes_respondidas=None, voltariam_fazer_negocio=None, indice_solucao=None, nota_consumidor=None):
        self.nome = nome
        self.categoria = categoria
        self.url = url
        self.reclamacoes_respondidas = reclamacoes_respondidas
        self.voltariam_fazer_negocio = voltariam_fazer_negocio
        self.indice_solucao = indice_solucao
        self.nota_consumidor = nota_consumidor