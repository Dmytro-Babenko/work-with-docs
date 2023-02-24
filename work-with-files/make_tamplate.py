from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from settings import TEMPLATE_SETTINGS, get_settings
    
def create_image_template(doc: DocxTemplate, folder: Path, symb: str) -> None:
    placeholders = {}
    i = 1
    for image in folder.iterdir():
        if image.suffix == '.png':
            placeholder = InlineImage(doc, str(image.absolute()), Cm(5))
            place = f'{symb}{i}'
            placeholders[place] = placeholder
            i += 1
    return placeholders

def main():
    template_settings = get_settings(TEMPLATE_SETTINGS)
    doc_path = template_settings['template Word']
    temp_image_folder = template_settings['folder with tamplates']
    doc = DocxTemplate(doc_path)
    placeholders = create_image_template(doc, temp_image_folder, temp_image_folder.name)
    doc.render(placeholders)
    doc.save(doc_path)
    pass

if __name__ == '__main__':
    main()
    




