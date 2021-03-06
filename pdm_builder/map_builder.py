import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
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


def cmap_plot(geodf, col, f_name, base_layer=None, path='mapas_subprefeituras_final', tipo_indicador = 'numérico'):

    if not os.path.exists(path):
        os.makedirs(path)
    fig, ax = plt.subplots(figsize=(10, 15))

    if tipo_indicador == 'numérico':
        print('mapa numerico')
        if geodf[col].max() < 8:
            geodf.plot(ax = ax, column=col, cmap='Blues',
                        legend=True,
                        edgecolor='0.5',
                        vmin=0)
        else:
            try:
                geodf.plot(ax = ax, column=col, cmap='Blues',
                            legend=True,
                            edgecolor='0.5',
                            legend_kwds = {'format':"%.0f"},
                            vmin=0)
            except Exception as e:
                if '__init__() got an unexpected keyword argument' in str(e):
                    print('Bug do legend keyword')
                    geodf[col] = geodf[col].apply(float)
                    geodf.plot(ax = ax, column=col, cmap='Blues',
                                legend=True,
                                edgecolor='0.5',
                                legend_kwds={'format': "%.0f"},
                                vmin=0)
                else:
                    raise(e)

    else:
        print('mapa categorico')
        geodf.plot(ax = ax, column=col, cmap='Blues',
                    legend=False,
                    categorical=True,
                    edgecolor='0.5')

    if base_layer is not None:
        base_layer['geometry'].boundary.plot(ax=ax,
                        color='black')


    plt.axis('off')

    fig.savefig(os.path.join(path, f_name))

    plt.clf()
    plt.close(fig)

class MapBuilder:

    def __init__(self, path_salvar = None, path_mapa_subs = None, path_mapa_zonas = None, path_controle = None, path_contorno_sp = None):

        if path_mapa_subs is None:
            path_mapa_subs = 'original_data/SIRGAS_SHP_subprefeitura/SIRGAS_SHP_subprefeitura_polygon.shp'
        if path_mapa_zonas is None:
            path_mapa_zonas = 'original_data/SIRGAS_SHP_regiao_5/SIRGAS_REGIAO5.shp'
        if path_controle is None:
            path_controle = 'original_data/Controle das Devolutivas.xlsx'
        if path_salvar is None:
            path_salvar = 'mapas_regionalizacao'
        if path_contorno_sp is None:
            path_contorno_sp = 'original_data/SIRGAS_limites_municipios_estado_sao_paulo/SIRGAS_limites_municipais.shp'

        self.mapa_subs = self.abrir_mapas_geosampa(path_mapa_subs)
        self.mapa_zonas = self.abrir_mapas_geosampa(path_mapa_zonas)
        self.mapa_sp = self.pegar_contorno_sp(path_contorno_sp)
        self.arrumar_nome_zonas()
        self.controle = self.pegar_controle_regionalizacao(path_controle)
        self.path_salvar = set_path(path_salvar)


    def abrir_mapas_geosampa(self, path):
        map_geodf = gpd.read_file(path)
        map_geodf.set_crs("epsg:3198", allow_override=True, inplace=True)

        return map_geodf

    def pegar_contorno_sp(self, path):

        map_geodf = self.abrir_mapas_geosampa(path)

        sp =  map_geodf[map_geodf['municipio']=='SÃO PAULO'].copy()
        sp.set_crs("epsg:3198", allow_override=True, inplace=True)

        return sp

    def pegar_controle_regionalizacao(self, path_planilha):

        controle = pd.read_excel(path_planilha)
        reg = controle[['Meta', 'Secretaria regionalizou?', 'Tipo de regionalização', 'Tipo do indicador?']]
        reg = reg[~reg['Meta'].isnull()].copy()

        filtro = reg['Secretaria regionalizou?'].isin(['Total', 'Parcial'])

        filtrado = reg[filtro].copy()
        filtrado['Meta'] = filtrado['Meta'].apply(lambda x: str(x))

        return filtrado

    def foi_regionalizada(self, ficha, ctrl_regionalizacao = None):

        if ctrl_regionalizacao is None:
            ctrl_regionalizacao = self.controle

        num_meta = ficha['ficha_tecnica']['numero_meta']

        if str(num_meta) in ctrl_regionalizacao['Meta'].unique():
            return True
        else:
            print(f'{num_meta} não está no controle de regionalizacao')


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
        else:
            print(f"Não foi regionalizada {ficha['file_original']}")

    def padronizar_valor(self, val):

        if pd.isnull(val):
            val = 0
        elif val == '':
            return 'vazio'
        try:
            if '.' in str(val):
                esquerda, direita = str(val).split('.')

                if int(direita) > 0:
                    return float(val)
                else:
                    return int(esquerda)
            else:
                val = int(val)
        except ValueError:
            val = 'categorico'

        return val

    def padronizar_valores(self, df_reg):

        df_reg['valores_padrao'] = df_reg['projecao_quadrienio'].apply(self.padronizar_valor)
        if (df_reg['valores_padrao'] == 'vazio').all():
            df_reg['flag_dados'] = 'vazio'
        else:
            if 'categorico' in df_reg['valores_padrao']:
                df_reg['flag_dados'] = 'categorico'
            else:
                df_reg['flag_dados'] = 'numerico'

        return df_reg

    def separar_zonas_de_subs(self, reg):

        zonas = reg[reg['subprefeitura'].isin(ZONAS)].copy()
        subs = reg[~reg['subprefeitura'].isin(ZONAS)].copy()

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

        geodf = gpd.GeoDataFrame(merged, geometry = merged['geometry'])
        geodf.set_crs("epsg:3198", allow_override=True, inplace=True)

        return geodf

    def merge_zonas(self, reg, mapa_zonas = None):

        if mapa_zonas is None:
            mapa_zonas = self.mapa_zonas

        reg = reg.rename({'subprefeitura' : 'nome_padrao'}, axis = 1)
        merged = mapa_zonas.merge(reg, how='left', on='nome_padrao')

        geodf = gpd.GeoDataFrame(merged, geometry=merged['geometry'])
        geodf.set_crs("epsg:3198", allow_override=True, inplace=True)

        return geodf

    def ok_cast_int(self, df, col):

        for item in df[col].unique():

            if type(item) is float:
                str_item = str(item)
                if '.' in str_item:
                    esquerda, direita = str_item.split('.')
                    if int(direita) > 0:
                        return False
                else:
                    if pd.isnull(item):
                        return False
            elif type(item) is int:
                pass
            else:
                return False
        return True

    def create_map(self, reg, num_meta, tipo_pol, binario):

        reg = self.padronizar_valores(reg)
        nom_file = f'{num_meta}_{tipo_pol}.png'

        col_values = 'projecao_quadrienio'

        if (reg['flag_dados'] == 'vazio').all():
            print(f'Nenhum dado para a meta {num_meta} e tipo {tipo_pol}')
        else:
            if 'categorico' in reg['flag_dados'].unique() or binario:
                cmap_plot(reg, col_values, nom_file, base_layer=self.mapa_sp,
                          path=self.path_salvar, tipo_indicador='categorico')
            elif (reg['flag_dados'] == 'numerico').all():
                if self.ok_cast_int(reg, col_values):
                    reg[col_values] = reg[col_values].apply(int)

                cmap_plot(reg, col_values, nom_file, base_layer=self.mapa_sp,
                          path=self.path_salvar, tipo_indicador='numérico')
            else:
                print(reg['flag_dados'].unique())

    def indicador_binario(self, num_meta, controle = None):

        if controle is None:
            controle = self.controle

        dados_meta = controle[controle['Meta'] == str(num_meta)]
        if not dados_meta.empty:
            dados_meta.reset_index(drop = True, inplace = True)
            tipo_indi = dados_meta['Tipo do indicador?'].iloc[0]

            if tipo_indi == 'Binário':
                return True
        return False

    def create_maps(self, ficha):

        num_meta = ficha['ficha_tecnica']['numero_meta']
        binario = self.indicador_binario(num_meta)
        try:
            reg = self.get_regionalizacao(ficha)
        except ValueError as e:
            print(ficha['ficha_tecnica']['numero_meta'])
            print(e)

        if reg is not None and not reg.empty:
            zonas, subs = self.separar_zonas_de_subs(reg)

            if not subs.empty:
                subs = self.merge_subs(subs)
                self.create_map(subs, num_meta, 'subs', binario)

            if not zonas.empty:
                zonas = self.merge_zonas(zonas)
                self.create_map(zonas, num_meta, 'zonas_sp', binario)


        else:
            print(f"Não há mapas para criar: {ficha['ficha_tecnica']['numero_meta']}")

    def __call__(self, ficha):

        self.create_maps(ficha)







