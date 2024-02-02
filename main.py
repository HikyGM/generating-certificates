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
            pdf_images[idx].save(f'comlete/{score}/{name}' + '.png', 'PNG')
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
        name = ' '.join(line.title().strip().split()[:-1])
        scores = float(line.split()[-1])
        if name not in names:
            names[name] = scores
        else:
            if names[name] < scores:
                names[name] = scores
len_names = len(names)
count = 1
for name, scores in names.items():
    context = {
        'item': name
    }

    if scores <= 16:
        temp = "temp.docx"
        score = 'участие'

    if 17 <= scores <= 20:
        temp = "temp3.docx"
        score = '3 место'

    if 21 <= scores <= 24:
        temp = "temp2.docx"
        score = '2 место'

    if 25 <= scores <= 32:
        temp = "temp1.docx"
        score = '1 место'

    print(name, '.....', score, f'... ({count}/{len_names})')
    count += 1

    doc = DocxTemplate(temp)
    doc.render(context)
    doc.save(f"comlete/{score}/{name}.docx")

    path = os.path.abspath(f"comlete/{score}/{name}.docx")
    convert = ConvertDoc2Pdf(path)
    convert.convert()

print('Хотово!')
