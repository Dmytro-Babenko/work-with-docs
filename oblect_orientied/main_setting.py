from setting_class import Setting, DATA_MAIN_SETTINGS
from work_with_exel import get_info_from_exel, export_image
from docxtpl import DocxTemplate
from clean_folder.sort import find_free_name

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

        info = get_info_from_exel(exel_path, text_sheet)
        graphs = export_image(exel_path, graph_sheet)
        doc_result_path = find_free_name(result_file_name, result_folder, doc_path.suffix)[1]

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


def main():
    main_settings = MainSettings('Main settings', DATA_MAIN_SETTINGS)
    main_root = main_settings.make_setting_root()
    main_settings.make_fields(main_root)
    main_settings.key_element.tk_variable.trace_add('write', main_settings.update)
    main_root.mainloop()

if __name__ == '__main__':
    main()