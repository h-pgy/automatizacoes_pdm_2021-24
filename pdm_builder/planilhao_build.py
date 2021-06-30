from docx import Document
import re
import pandas as pd


class DocxParser:

    def __init__(self, path_arquivo):

        self.doc = Document(path_arquivo)

    def limpar_texto(self, text):

        char_zuado = '\xa0'
        text = text.replace(char_zuado, ' ')
        text = text.strip()

        return text

    def read_table(self, table):

        parsed = {}
        for row in table.rows:
            cell_key, cell_value = row.cells
            key = self.limpar_texto(cell_key.text).strip()
            value = self.limpar_texto(cell_value.text).strip()
            if key in parsed:
                key = key + '_'
            parsed[key] = value

        return parsed

    def parse_all_tables(self, doc):

        parsed = []
        tabelas = doc.tables

        for table in tabelas:
            parsed_table = self.read_table(table)
            parsed.append(parsed_table)

        return parsed

    def parse_iniciativas(self, parsed_data):

        iniciativas = parsed_data.pop('Iniciativas')
        iniciativas = iniciativas.split('\n')

        pat = re.compile('^[a-zA-Z]\) ')
        parsed = {}
        for ini in iniciativas:
            m = re.match(pat, ini)
            if m:
                letra = m.group().strip()
                parsed[letra] = ini[2:].strip()
            else:
                print(ini)

        return parsed

    def cast_int(self, valor):

        str_valor = str(valor)

        if '.' in str_valor:
            esquerda, direita = str_valor.split('.')
            if int(direita) > 0:
                return float(str_valor)
            else:
                return int(esquerda)
        else:
            try:
                return int(str_valor)
            except ValueError:
                return valor

    def parse_regionaliz(self, parsed_data):

        regionalizacao = parsed_data.pop('Regionalização - projeção quadriênio')
        regionalizacao = regionalizacao.split('\n')

        parsed = {}
        for reg in regionalizacao:

            splited = reg.split(':')
            if len(splited) == 2:
                subs, valor = splited
                subs = subs.strip()
                valor = valor.strip()
                parsed[subs] = self.cast_int(valor)
            else:
                parsed['comentário'] = reg.strip()

        return parsed

    def __call__(self):

        parsed_data = self.parse_all_tables(self.doc)

        for obj in parsed_data:
            obj['regionalizacao'] = self.parse_regionaliz(obj)
            obj['iniciativas'] = self.parse_iniciativas(obj)

        return parsed_data


class ExcelMaker:

    def __init__(self, data):

        self.data = data

    def subset_data(self, data, target_keys):

        parsed = []
        for dici in data:
            subset = {}
            for key, value in dici.items():
                if key in target_keys:
                    subset[key] = dici[key]

            parsed.append(subset)

        return parsed

    def planilha_principal(self, data):

        coluns = ['Meta',
                  'Meta_',
                  'Objetivo estratégico',
                  'Indicador',
                  'Contexto',
                  'Informações Complementares',
                  'ODS vinculados',
                  'Secretaria Responsável']

        planilha_data = self.subset_data(data, coluns)

        df = pd.DataFrame(planilha_data)
        df.rename({'Meta': 'Número meta',
                   'Meta_': 'Meta'}, axis=1, inplace=True)

        ordem_cols = ['Número meta', 'Meta', 'Indicador', 'Contexto',
                      'Informações Complementares', 'ODS vinculados', 'Secretaria Responsável']

        return df[ordem_cols]

    def planilha_iniciativas(self, data):

        coluns = [
            'Meta',
            'Meta_',
            'iniciativas'
        ]

        subset = self.subset_data(data, coluns)

        parsed_data = []

        for item in subset:
            num_meta = item['Meta']
            desc_meta = item['Meta_']
            for letra, ini in item['iniciativas'].items():
                row = {}
                row['meta_numero'] = num_meta
                row['meta_descricao'] = desc_meta
                row['iniciativa'] = letra.replace(')', '')
                row['iniciativa_descricao'] = ini

                parsed_data.append(row)

        return pd.DataFrame(parsed_data)

    def planilha_regionalizacoes(self, data):

        coluns = [
            'Meta',
            'Meta_',
            'regionalizacao'
        ]

        subset = self.subset_data(data, coluns)

        parsed_data = []

        for item in subset:
            row = {}
            row['meta_numero'] = item['Meta']
            row['meta_descricao'] = item['Meta_']
            for subs, valor in item['regionalizacao'].items():
                row[subs] = valor

            parsed_data.append(row)

        return pd.DataFrame(parsed_data)

    def write_excel(self, file_name):

        writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
        self.planilha_principal(self.data).to_excel(writer, sheet_name='principal', index=False)
        self.planilha_iniciativas(self.data).to_excel(writer, sheet_name='iniciativas', index=False)
        self.planilha_regionalizacoes(self.data).to_excel(writer, sheet_name='regionalizacao', index=False)

        writer.save()

if __name__ == "__main__":
    from pathlib import Path
    import xlsxwriter
    path_onedrive = Path(
        r'C:\Users\h-pgy\one_drive_prefs\OneDrive - Default Directory\Shared Documents\Versão final do Programa de Metas')
    path_arquivo = path_onedrive / 'Fichas FINAL - 30.06.docx'
    parser = DocxParser(path_arquivo)
    data = parser()
    xl = ExcelMaker(data)
    xl.write_excel('planilhao_final.xlsx')