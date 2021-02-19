import xlrd
import WordGen.Excel2Json as e2j

print(dir())

if __name__ == "__main__":
    wb = xlrd.open_workbook("__info/Venus_SoC_Memory_Mapping.xls")

    e2j.wordgen_excel2json(wb)
    e2j.wordgen_remodel_json()