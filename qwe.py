import os
import sys
from docxtpl import DocxTemplate
import comtypes.client
from pdf2image import convert_from_path


class ConvertDoc2Pdf:
    def __init__(self, input_file):
        self.input_file = input_file

    def init_word(self):
        powerpoint = comtypes.client.CreateObject("Word.Application")
        return powerpoint

    def word_to_pdf(self, doc, inputFileName, formatType=17):
        deck = doc.Documents.Open(inputFileName)
        outputFileName = inputFileName[:-5] + '.pdf'
        deck.SaveAs(outputFileName, formatType)
        deck.Close()
        return outputFileName

    def pdf_to_png(self, file):
        pdf_images = convert_from_path(file, poppler_path='/poppler-23.11.0/Library/bin')
        for idx in range(len(pdf_images)):
            pdf_images[idx].save(f'comlete/благодарки/{name}' + '.png', 'PNG')
        print("Successfully converted PDF to images")

    def convert(self):
        comtypes.CoInitialize()
        word_doc = self.init_word()
        pdf = self.word_to_pdf(word_doc, self.input_file)
        word_doc.Quit()
        self.pdf_to_png(pdf)


lines = list(map(str.strip, sys.stdin))

names = {}
for line in lines:
    if line:
        name = ' '.join(line.title().strip().split())
        scores = ''
        names[name] = scores
print(names)
len_names = len(names)
count = 1
for name, scores in names.items():
    context = {
        'item': name
    }



    print(name, '.....',  f'... ({count}/{len_names})')
    count += 1

    doc = DocxTemplate('temp4.docx')
    doc.render(context)
    doc.save(f"comlete/благодарки/{name}.docx")

    path = os.path.abspath(f"comlete/благодарки/{name}.docx")
    convert = ConvertDoc2Pdf(path)
    convert.convert()

print('Хотово!')
