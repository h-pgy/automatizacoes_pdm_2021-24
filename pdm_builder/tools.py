import pandas as pd

def pegar_filtro_alteracoes(path = None):

    if path is None:
        path = 'original_data/Controle das Devolutivas.xlsx'

    controle = pd.read_excel(path, sheet_name='Base')
    filtro_alteracoes = controle['Houve proposta de alteração substancial?'].str.lower( )=='sim'

    metas = list(controle['Meta'][filtro_alteracoes])

    return metas


def sort_fichas(fichas):
    fichas_num = [ficha for ficha in fichas if
                  type(ficha['ficha_tecnica']['numero_meta']) in (int, float)]
    fichas_text = [ficha for ficha in fichas if
                   type(ficha['ficha_tecnica']['numero_meta']) is str]

    fichas_num = sorted(fichas_num, key=lambda ficha: ficha['ficha_tecnica']['numero_meta'])
    fichas_text = sorted(fichas_text, key=lambda ficha: ficha['ficha_tecnica']['numero_meta'])

    ordenadas = fichas_num + fichas_text

    print([ficha for ficha in fichas if ficha not in ordenadas])

    return ordenadas


def checar_metas_presentes(filtro, fichas):
    num_metas = []
    for ficha in fichas:
        num_metas.append(ficha['ficha_tecnica']['numero_meta'])

    return [meta for meta in filtro if meta not in num_metas]

