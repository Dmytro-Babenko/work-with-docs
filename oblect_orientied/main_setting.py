import re
from setting_class import Setting, DATA_MAIN_SETTINGS
from openpyxl import worksheet, load_workbook, Workbook
from work_with_exel import get_info_from_exel, export_image, GRAPH_SYMBOL
from docxtpl import DocxTemplate, InlineImage
from clean_folder.sort import find_free_name
from pathlib import Path

class ExelWordbook():
    def __init__(self, path: Path, infosheet: str, graphsheet: str) -> None:
        self.path = path
        self.infosheet = infosheet
        self.graphsheet = graphsheet
        pass

    def get_info_from_tables(self, sheet: worksheet) -> dict[str:list]:
        info = {}
        for table in sheet.tables.values():    
            headers = table.column_names
            ref = re.sub(r'(\d):', lambda m: f'{int(m.group(1))+1}:', table.ref)
            info[table.name] = [{header: str(cell.value).replace('.', ',') if isinstance(cell.value, float) else cell.value
                                for header, cell in filter(lambda t: t[1].value, zip(headers, row))}
                                for row in sheet[ref] if row[0].value]
        return info

    def get_info_from_defnames(book: Workbook):
        info = {}
        for name, defn_object in book.defined_names.items():
            destination = next(defn_object.destinations)
            sheet_name, coordinates = destination
            value = book[sheet_name][coordinates].value
            if isinstance(value, float):
                value = str(value).replace('.', ',')
            info[name] = value
        return info


    def get_info_from_exel(self) -> dict[str:any]:
        '''
        Get information from the Exel sheet.
        Return dictionary
        '''
        wb = load_workbook(self.path, data_only=True)
        ws = wb[self.infosheet]
        info = {}

        info = self.get_info_from_tables(ws)
        info.update(self.get_info_from_defnames(wb))
        return info

class MainSettings(Setting):
    def __init__(self, setting_name, data) -> None:
        super().__init__(setting_name, data)

    def exe_program(self):
        exel_path = self.data['exel'].get_value()
        doc_path = self.data['word'].get_value()
        text_sheet = self.data['sheet with information'].get_value()
        graph_sheet = self.data['sheet with charts'].get_value()
        result_file_name = self.data['result file name'].get_value()
        result_folder = self.data['result folder'].get_value()

        doc = DocxTemplate(doc_path)
        info = get_info_from_exel(exel_path, text_sheet)
        graphs = export_image(exel_path, graph_sheet)
        graphs = {x.stem: InlineImage(doc, str(x)) for x in graphs}
        info.update(graphs)
        doc_result_path = find_free_name(result_file_name, result_folder, doc_path.suffix)[1]

        doc.render(info)
        doc.save(doc_result_path)
        # for plate, graph in graphs.items():
        #     try:
        #         doc.replace_pic(plate, graph)
        #         doc.save(doc_result_path)
        #     except ValueError:
        #         continue
        pass


def main():
    main_settings = MainSettings('Main settings', DATA_MAIN_SETTINGS)
    main_root = main_settings.make_setting_root()
    main_settings.make_fields(main_root)
    main_settings.key_element.tk_variable.trace_add('write', main_settings.update)
    main_root.mainloop()

if __name__ == '__main__':
    main()