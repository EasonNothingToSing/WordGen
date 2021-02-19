from docxtpl import DocxTemplate
import json


def wordgen_temp2word():
    tpl = DocxTemplate("../.Template/WordGen.docx")
    with open("../__info/__ex2js.remodel.json", "r") as fr:
        handle = json.load(fr)

    tpl.render(handle)

    tpl.save("../ListenAI_Doc.docx")


if __name__ == "__main__":
    wordgen_temp2word()
