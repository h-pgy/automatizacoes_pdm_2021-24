from docx import Document
from pdm_builder.ficha_parsing import ParserFichas
from pdm_builder.doc_builder import TableBuilder
from pdm_builder.map_builder import MapBuilder
from pdm_builder.tools import pegar_filtro_alteracoes, sort_fichas, checar_metas_presentes, get_map_files


def build_docx(filtro = None, sort_func = None):

    parser = ParserFichas('original_data/Devolutiva 11-jun', 'rodada_3')
    fichas = parser()
    if filtro:
        print(f'Metas não encontradas {checar_metas_presentes(filtro, fichas)}')
        fichas = [ficha for ficha in fichas
                  if ficha['ficha_tecnica']['numero_meta'] in filtro]
    if sort_func:
        fichas = sort_func(fichas)
    docx = Document()
    mapbuilder = MapBuilder()
    for ficha in fichas:
        try:
            mapbuilder.create_maps(ficha)
        except Exception as e:
            print(f"Erro ao gerar os mapas na ficha {ficha['file_original']}")
            print(e)

        builder = TableBuilder(ficha, docx)
        docx = builder()

        mapas = get_map_files(ficha['ficha_tecnica']['numero_meta'])
        for mapa in mapas:
            docx.add_picture(mapa)

    docx.save('teste_6.docx')


if __name__ == '__main__':
    try:
        filtro = pegar_filtro_alteracoes()
        build_docx(filtro, sort_fichas)
    except Exception as e:
        print(e)