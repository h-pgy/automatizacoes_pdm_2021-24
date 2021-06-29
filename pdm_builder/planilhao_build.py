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
                int(str_valor)
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

