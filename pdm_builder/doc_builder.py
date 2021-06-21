from collections import OrderedDict
from docx import Document
from docx.shared import Cm
import pandas as pd

class TableBuilder:

    def __init__(self, ficha, docx=None):

        self.ficha = ficha
        if docx is None:
            docx = Document()
        self.doc = docx

    def build_iniciativas(self, iniciativas):

        lista_textos = []
        for ini in iniciativas:
            texto = f"{ini['id']}) {ini['descricao']}"
            lista_textos.append(texto)
        return '\n'.join(lista_textos)

    def calc_total_orcamento(self, ficha):

        total = 0
        for orc in ficha['orcamento']:
            if str(orc['classif.']).lower() != 'execução':
                continue

            if pd.isnull(orc['Custo TOTAL']):
                custo = 0
            elif type(orc['Custo TOTAL']) is str:
                custo = float(orc['Custo TOTAL'].replace('.', '').replace(',', '.'))
            else:
                custo = orc['Custo TOTAL']
            total += custo

        total = 'R${:,.2f}'.format(total)
        # hackzin engraçado para deixar no formato brasileiro
        total = total.replace('.', 'X').replace(',', '.').replace('X', ',')

        return total

    def build_regionalizacao(self, ficha):

        lista_textos = []

        for reg in ficha['regionalizacao']:
            texto = f"{reg['subprefeitura']} : {reg['projecao_quadrienio']}"
            lista_textos.append(texto)

        return '\n'.join(lista_textos)

    def table_data(self, ficha=None):

        if ficha is None:
            ficha = self.ficha

        tabela_word = OrderedDict()

        # note que o primeiro tem espaço
        tabela_word['Meta '] = str(ficha['ficha_tecnica']['numero_meta'])
        tabela_word['Objetivo estratégico'] = ''
        tabela_word['Meta'] = ficha['ficha_tecnica']['desc_meta']
        tabela_word['Indicador'] = ficha['ficha_tecnica']['indicador']
        tabela_word['Contexto'] = ficha['ficha_tecnica']['contexto']
        tabela_word['Informações Complementares'] = ficha['ficha_tecnica']['info_compl']
        tabela_word['ODS vinculados'] = ''
        tabela_word['Iniciativas'] = self.build_iniciativas(ficha['iniciativas'])
        tabela_word['Secretaria Responsável'] = ficha['ficha_tecnica']['secretaria']
        tabela_word['Total Orçado'] = self.calc_total_orcamento(ficha)
        tabela_word['Regionalização - projeção quadriênio'] = self.build_regionalizacao(ficha)

        return tabela_word

    def build_table(self, table_data=None, docx=None):

        if docx is None:
            docx = self.doc
        if table_data is None:
            table_data = self.table_data(self.ficha)

        table = docx.add_table(rows=len(table_data), cols=2)
        table.autofit = False
        table.allow_autofit = False

        i = 0
        for cell_nom, cell_value in table_data.items():
            try:
                if pd.isnull(cell_value):
                    print(cell_value)
                    cell_value = ''
                row = table.rows[i]
                row.cells[0].text = cell_nom
                row.cells[0].width = Cm(5)
                row.cells[1].text = cell_value
                row.cells[1].width = Cm(10)
                i += 1
            except Exception as e:
                print(cell_nom)
                print(cell_value)
                raise e

        table.style = 'Light List Accent 1'

        return docx

    def __call__(self, table_data=None, docx=None):

        return self.build_table(table_data, docx)