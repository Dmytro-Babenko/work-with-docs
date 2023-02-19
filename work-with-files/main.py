
from work_with_exel import get_info_from_exel, export_image
from docxtpl import DocxTemplate
from clean_folder.sort import find_free_name
from settings import get_main_user_settings

def main():
    #  не проверять файл с фотками мб поделить проверку, сщщбщения о проверке
    main_settings = get_main_user_settings()
    base_folder = main_settings['base_folder']
    exel_path = main_settings['Exel']
    doc_path = main_settings['Word']
    text_sheet = main_settings['Sheet with information']
    graph_sheet = main_settings['Sheet with charts']

    info = get_info_from_exel(exel_path, text_sheet)
    graphs = export_image(exel_path, graph_sheet)

    doc = DocxTemplate(doc_path)
    doc.render(info)
    for plate, graph in graphs.items():
        doc.replace_pic(plate, graph)
    
    # изменить папку и имя
    doc_result_path = find_free_name(doc_path.stem, base_folder, doc_path.suffix)[1]
    doc.save(doc_result_path)
    pass

if __name__ == '__main__':
    main()