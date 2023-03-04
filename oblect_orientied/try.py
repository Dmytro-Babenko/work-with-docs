from docxtpl import DocxTemplate, InlineImage

doc = DocxTemplate(r'D:\test\бланк — копия.docx')

context = {'gr1': InlineImage(doc, r'D:\test\gr\gr1.png'), 'gr2': InlineImage(doc, r'D:\test\gr\gr2.png')}
print(context)

doc.render(context)
doc.save(r'D:\test\2.docx')