
from work_with_exel import get_info_from_exel, export_image
from docxtpl import DocxTemplate
from clean_folder.sort import find_free_name
from settings import MAIN_SETTINGS, get_settings

def main():
    main_settings = get_settings(MAIN_SETTINGS)
    base_folder = main_settings['base_folder']
    exel_path = main_settings['Exel']
    doc_path = main_settings['Word']
    text_sheet = main_settings['Sheet with information']
    graph_sheet = main_settings['Sheet with charts']
    result_file_name = main_settings['Result file name']
    result_folder = main_settings['Result folder']

    info = get_info_from_exel(exel_path, text_sheet)
    graphs = export_image(exel_path, graph_sheet)

    result_folder_path = base_folder.joinpath(result_folder)
    result_folder_path.mkdir(exist_ok=True)
    new_name = result_folder_path.joinpath(result_file_name)
    doc_result_path = find_free_name(new_name, base_folder, doc_path.suffix)[1]

    doc = DocxTemplate(doc_path)
    doc.render(info)
    doc.save(doc_result_path)
    for plate, graph in graphs.items():
        try:
            doc.replace_pic(plate, graph)
            doc.save(doc_result_path)
        except ValueError:
            continue
    pass

if __name__ == '__main__':
    main()