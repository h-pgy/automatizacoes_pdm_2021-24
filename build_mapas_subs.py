from pathlib import Path
import pandas as pd
from pdm_builder.map_builder import MapBuilder
from pdm_builder.ficha_parsing import ParserFichas

PATH_FILES_ONE_DRIVE = Path(r'C:\Users\h-pgy\one_drive_prefs\OneDrive - Default Directory\Shared Documents\Estruturação do PDM 2021-2024\Elaboração PDM Versão Final')


def pegar_controle_regionalizacao(path_planilha):
    controle = pd.read_excel(path_planilha)
    reg = controle[['Meta', 'Secretaria regionalizou?', 'Tipo de regionalização', 'Tipo do indicador?']]
    reg = reg[~reg['Meta'].isnull()].copy()

    filtro = reg['Secretaria regionalizou?'].isin(['Total', 'Parcial'])

    filtrado = reg[filtro].copy()
    filtrado['Meta'] = filtrado['Meta'].apply(lambda x: str(x))

    return filtrado

def foi_regionalizada(ficha, ctrl_regionalizacao):

    num_meta = ficha['ficha_tecnica']['numero_meta']

    if str(num_meta) in ctrl_regionalizacao['Meta'].unique():
        return True
    else:
        print(f'{num_meta} não está no controle de regionalizacao')


def build_all_maps(path_salvar_jsons, path_salvar_mapas = None):


    parser = ParserFichas(PATH_FILES_ONE_DRIVE / 'Fichas Metas\Devolutiva 11-jun', path_salvar_jsons)
    fichas = parser()

    path_controle = PATH_FILES_ONE_DRIVE / 'Controle das Devolutivas.xlsx'

    builder = MapBuilder(path_salvar=path_salvar_mapas, path_controle=path_controle)

    for ficha in fichas:
        try:
            builder(ficha)
        except Exception as e:
            print(f"ERRO NA FICHA {ficha['file_original']}")
            print(e)


if __name__ == "__main__":

    build_all_maps('rodada_mapas_subs', path_salvar_mapas='mapas_subs_final_v2')