from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from com_with_user import get_path
    
def create_image_template(doc: DocxTemplate, folder: Path, symb='gr') -> None:
    placeholders = {}
    i = 1
    for image in folder.iterdir():
        if image.suffix == '.png':
            placeholder = InlineImage(doc, str(image.absolute()), Cm(5))
            place = f'{symb}{i}'
            placeholders[place] = placeholder
            i += 1
    doc.render(placeholders)
    doc.save('test1.docx')
    pass

def main():
    doc_path = get_path('template Word')
    base_folder = doc_path.parent
    temp_image_folder = get_path('folder with template image', parent_folder=base_folder)
    doc = DocxTemplate('test.docx')
    create_image_template(doc, temp_image_folder)
    pass

if __name__ == '__main__':
    main()
    




