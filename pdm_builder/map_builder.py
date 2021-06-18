import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from pdm_builder.tools import set_path

DE_PARA_SUBS = {
    'Aricanduva/Vila Formosa' : 'ARICANDUVA-FORMOSA-CARRAO',
    'Butantã' : 'BUTANTA',
    'Campo Limpo' : 'CAMPO LIMPO',
    'Capela do Socorro' : 'CAPELA DO SOCORRO',
    'Casa Verde' : 'CASA VERDE-CACHOEIRINHA',
    'Cidade Ademar' : 'CIDADE ADEMAR',
    'Cidade Tiradentes' : 'CIDADE TIRADENTES',
    'Ermelino Matarazzo' : 'ERMELINO MATARAZZO',
    'Freguesia do Ó/Brasilândia' : 'FREGUESIA-BRASILANDIA',
    'Guaianases' : 'GUAIANASES',
    'Ipiranga' : 'IPIRANGA',
    'Itaim Paulista' : 'ITAIM PAULISTA',
    'Itaquera' : 'ITAQUERA',
    'Jabaquara' : 'JABAQUARA',
    'Jaçanã/Tremembé' : 'JACANA-TREMEMBE',
    'Lapa' : 'LAPA',
    "M'Boi Mirim" : "M'BOI MIRIM",
    'Mooca' : 'MOOCA',
    'Parelheiros' : 'PARELHEIROS',
    'Penha' : 'PENHA',
    'Perus' : 'PERUS',
    'Pinheiros' : 'PINHEIROS',
    'Pirituba/Jaraguá' : 'PIRITUBA-JARAGUA',
    'Santana/Tucuruvi' : 'SANTANA-TUCURUVI',
    'Santo Amaro' : 'SANTO AMARO',
    'São Mateus' : 'SAO MATEUS',
    'São Miguel Paulista' : 'SAO MIGUEL',
    'Sapopemba' : 'SAPOPEMBA',
    'Sé' : 'SE',
    'Vila Maria/Vila Guilherme' : 'VILA MARIA-VILA GUILHERME',
    'Vila Mariana' : 'VILA MARIANA',
    'Vila Prudente' : 'VILA PRUDENTE'
}

ZONAS = (
    'CENTRO',
    'SUL',
    'NORTE',
    'OESTE',
    'LESTE'
)


def cmap_plot(geodf, col, f_name = None, path='mapas_subprefeituras_final', tipo_indicador = 'numérico'):

    sns.set()

    if not os.path.exists(path):
        os.makedirs(path)

    if tipo_indicador == 'numérico':
        legend = True
    else:
        legend = False

    ax = geodf.plot(column=col, cmap = 'GnBu',
                    legend_kwds={'orientation': "vertical"},
                    legend=legend,
                    figsize = (10, 15),
                    edgecolor = 'black')

    plt.axis('off')

    fig = ax.get_figure()

    if f_name is None:
        f_name = title +'.png'

    fig.savefig(os.path.join(path, f_name))

class MapBuilder:

    def __init__(self, path_salvar = None, path_mapa_subs = None, path_mapa_zonas = None, path_controle = None):

        if path_mapa_subs is None:
            path_mapa_subs = 'original_data/SIRGAS_SHP_subprefeitura/SIRGAS_SHP_subprefeitura_polygon.shp'
        if path_mapa_zonas is None:
            path_mapa_zonas = 'original_data/SIRGAS_SHP_regiao_5/SIRGAS_REGIAO5.shp'
        if path_controle is None:
            path_controle = 'original_data/Controle das Devolutivas.xlsx'
        if path_salvar is None:
            path_salvar = 'mapas_regionalizacao'

        self.mapa_subs = self.abrir_mapas_geosampa(path_mapa_subs)
        self.mapa_zonas = self.abrir_mapas_geosampa(path_mapa_zonas)
        self.arrumar_nome_zonas()
        self.controle = self.pegar_controle_regionalizacao(path_controle)
        self.path_salvar = set_path(path_salvar)


    def abrir_mapas_geosampa(self, path):
        map_geodf = gpd.read_file(path)
        map_geodf.crs = {'init': 'epsg:31983'}
        map_geodf = map_geodf.to_crs(epsg=4326)

        return map_geodf

    def pegar_controle_regionalizacao(self, path_planilha):

        controle = pd.read_excel(path_planilha)
        reg = controle[['Meta', 'Secretaria regionalizou?', 'Tipo de regionalização', 'Tipo do indicador?']]
        reg = reg[~reg['Meta'].isnull()].copy()

        filtro = reg['Secretaria regionalizou?'].str.lower().isin(['total', 'parcial'])

        return reg[filtro]

    def foi_regionalizada(self, ficha, ctrl_regionalizacao = None):

        if ctrl_regionalizacao is None:
            ctrl_regionalizacao = self.controle

        num_meta = ficha['ficha_tecnica']['numero_meta']

        if num_meta in ctrl_regionalizacao['Meta']:
            return True

    def check_if_reg_ok(self, ficha):

        for item in ficha['regionalizacao']:

            reg = item['subprefeitura']
            if reg not in DE_PARA_SUBS and \
                    reg not in ZONAS:
                raise RuntimeError(f'Regionalização {reg} não prevista!')

        return True

    def get_regionalizacao(self, ficha):

        if self.foi_regionalizada(ficha) and self.check_if_reg_ok(ficha):

            regionalizacao = pd.DataFrame(ficha['regionalizacao'])

            return regionalizacao

    def padronizar_valor(self, val):

        if pd.isnull(val):
            val = 0
        try:
            val = float(val)
        except ValueError:
            val = 'categorico'

        return val

    def padronizar_valores(self, df_reg):

        df_reg['valores_padrao'] = df_reg['projecao_quadrienio'].apply(self.padronizar_valor)

        if 'categorico' in df_reg['valores_padrao']:
            df_reg['tipo_indicador'] = 'categorico'
        else:
            df_reg['tipo_indicador'] = 'numerico'

        return df_reg

    def separar_zonas_de_subs(self, reg):

        zonas = reg[reg['subprefeitura'].isin(ZONAS)]
        subs = reg[~reg['subprefeitura'].isin(ZONAS)]

        return zonas, subs

    def arrumar_nome_zonas(self, mapa_zonas = None):

        if mapa_zonas is None:
            mapa_zonas = self.mapa_zonas

        mapa_zonas['nome_padrao'] = mapa_zonas['NOME'].str.upper()

    def merge_subs(self, reg, mapa_subs = None):

        if mapa_subs is None:
            mapa_subs = self.mapa_subs

        reg['sp_nome'] = reg['subprefeitura'].apply(lambda x: DE_PARA_SUBS[x])

        merged = mapa_subs.merge(reg, how = 'left', on='sp_nome')

        return gpd.GeoDataFrame(merged, geometry = merged['geometry'], crs="EPSG:31983")

    def merge_zonas(self, reg, mapa_zonas = None):

        if mapa_zonas is None:
            mapa_zonas = self.mapa_zonas

        reg = reg.rename({'subprefeitura' : 'nome_padrao'}, axis = 1)
        merged = mapa_zonas.merge(reg, how='left', on='nome_padrao')

        return gpd.GeoDataFrame(merged, geometry=merged['geometry'], crs="EPSG:31983")

    def create_map(self, reg, num_meta, tipo_pol):

        reg = self.padronizar_valores(reg)
        nom_file = f'{num_meta}_{tipo_pol}.png'

        if 'categorico' in reg['tipo_indicador']:
            cmap_plot(reg, 'projecao_quadrienio', nom_file, path = self.path_salvar, tipo_indicador='categorico')
        else:
            cmap_plot(reg, 'projecao_quadrienio', nom_file, path=self.path_salvar, tipo_indicador='numérico')

    def create_maps(self, ficha):

        num_meta = ficha['ficha_tecnica']['numero_meta']
        try:
            reg = self.get_regionalizacao(ficha)
        except ValueError as e:
            print(e)

        if reg is not None and not reg.empty:
            zonas, subs = self.separar_zonas_de_subs(reg)

            if not zonas.empty:
                zonas = self.merge_zonas(zonas)
                self.create_map(zonas, num_meta, 'zonas_sp')

            if not subs.empty:
                subs = self.merge_subs(subs)
                self.create_map(subs, num_meta, 'subs')
        else:
            print(f"Não há mapas para criar: {ficha['ficha_tecnica']['numero_meta']}")







