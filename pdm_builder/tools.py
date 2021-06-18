import pandas as pd
import os

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


def dropar_cols_vazias(df):
    dropar_cols = []
    for col in df.columns:
        if df[col].isnull().all():
            dropar_cols.append(col)
    df.drop(dropar_cols, axis=1, inplace=True)

def set_path(path):

    if not os.path.exists(path):
        os.mkdir(path)

    return path

def check_subs(fichas):

    from pdm_builder.map_builder import DE_PARA_SUBS
    set_subs = set()
    for ficha in fichas:
        subs = [item['subprefeitura'] for item in ficha['regionalizacao']
                if item['subprefeitura'] not in ('CENTRO', 'LESTE', 'OESTE', 'NORTE', 'SUL')]

        set_subs.update(subs)

    fora = [subs for subs in set_subs if subs not in DE_PARA_SUBS]

    if fora:
        raise RuntimeError(f'Subprefeitura {fora} não previstas')

def check_zonas(fichas):

    from pdm_builder.map_builder import ZONAS

    set_zonas = set()
    for ficha in fichas:
        zonas = [item['subprefeitura'] for item in ficha['regionalizacao']
                if item not in DE_PARA_SUBS]

        set_zonas.update(zonas)


    fora = [zona for zona in set_zonas if zona not in ZONAS]

    if fora:
        raise RuntimeError(f'Zona {fora} não previstas')


def get_map_files(num_meta, path_maps = None):

    if path_maps is None:
        path_maps = 'mapas_regionalizacao'

    return [os.path.join(path_maps, file) for file in os.listdir(path_maps)
            if file.startswith(f'{num_meta}_')]




