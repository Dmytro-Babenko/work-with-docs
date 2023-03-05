from docxtpl import DocxTemplate, InlineImage
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import quote_sheetname, absolute_coordinate
import re
from pathlib import Path
from win32com.client import Dispatch

def a():
    doc = DocxTemplate(r'D:\test\бланк — копия.docx')

    context = {'gr1': InlineImage(doc, r'D:\test\gr\gr1.png'), 'gr2': InlineImage(doc, r'D:\test\gr\gr2.png'),
            'data': [{'name': 'Dima', 'age': 15}, {'name': 'Anton', 'age': 20}]}
    print(context)

    doc.render(context)
    doc.save(r'D:\test\2.docx')


work_book = load_workbook(r'D:\test\бланк тренд 1 — копия.xlsx', data_only=True)
sheet = work_book['main']

# table = sheet.tables['table1']
print(sheet.tables)

# headers = table.column_names

# arr = re.sub(r'(\d):', lambda m: f'{int(m.group(1))+1}:', table.ref)

# dct = {table.name: [{header: str(cell.value).replace('.', ',') if isinstance(cell.value, float) else cell.value
#                     for header, cell in filter(lambda t: t[1].value, zip(headers, row))}
#                     for row in sheet[arr] if row[0].value]}
# # print(sheet[arr])

# defn = work_book.defined_names

# dct2 = {name: work_book[next(obj.destinations)[0]][next(obj.destinations)[1]].value 
#         for name, obj in defn.items()}
# print(dct2)

# for name, obj in defn.items():
#     gen = next(obj.destinations)
#     sheet_name = gen[0]
#     coor = gen[1]
#     value = work_book[sheet_name][coor].value
#     dct[name] = value
#     # print(sheet[coor].value)
# work_book.close()

# doc = DocxTemplate(r'D:\test\бланк — копия — копия.docx')


def export_image(exel_file: Path, sheet_name) -> dict[str:Path]: 
    '''
    Save all charts in Exel sheet to the folder, with Exel_sheet name
    Return dictionary with images name and path
    '''
    graphs = []
    gr_folder = exel_file.parent.joinpath('gr')
    gr_folder.mkdir(exist_ok=True)
    app = Dispatch('Excel.Application')
    wb = app.Workbooks.Open(Filename=exel_file)
    app.DisplayAlerts = False

    i = 1
    gr_sheet = wb.Worksheets(sheet_name)
    for chartObject in gr_sheet.ChartObjects():
        # gr_name = f'{GRAPH_SYMBOL}{i}.png'
        gr_path = gr_folder.joinpath(f'gr{i}.png')
        graphs.append(gr_path)
        chartObject.Chart.Export(gr_path)
        i += 1
    wb.Close(SaveChanges=False, Filename=str(exel_file))

    return graphs

print(export_image(Path(r'D:\test\бланк тренд 1 — копия.xlsx'), 'main'))

context = dct
# print(context)


# doc.render(context)
# doc.save(r'D:\test\2.docx')