
import sys
sys.path.append(r'D:\PYHTON\GitHub\home-work2\clean_folder\clean_folder')
from sort import find_free_name
from work_with_exel import get_info_from_exel, export_image
from docxtpl import DocxTemplate
from com_with_user import get_path, get_sheet_name

def main():
    base_folder = get_path('main folder')
    exel_path = get_path('Exel', parent_folder=base_folder)
    doc_path = get_path('template Word', parent_folder=base_folder)
    text_sheet = get_sheet_name(exel_path, 'text info')
    graph_sheet = get_sheet_name(exel_path, 'graphs')

    info = get_info_from_exel(exel_path, text_sheet)
    graphs = export_image(exel_path, graph_sheet)

    doc = DocxTemplate(doc_path)
    doc.render(info)
    for plate, graph in graphs.items():
        doc.replace_pic(plate, graph)
    doc_result_path = find_free_name(doc_path.stem, base_folder, doc_path.suffix)[1]
    doc.save(doc_result_path)
    pass

if __name__ == '__main__':
    main()