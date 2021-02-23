import logging
import os
import xlrd
from docx import Document
from docxtpl import DocxTemplate
from . import (Excel2Json, Json2Temp, Temp2Word)

__all__ = ["generate_template", "update_template", "WORDGEN_VERSION", "WORDGEN_XLS_PATH_LIST", "WORDGEN_TPL_PATH_LIST",
           "excel_content_verify"]

WORDGEN_VERSION = "1.1"
__version__ = WORDGEN_VERSION

WORDGEN_ACTION_CHOICES = ["generate", "update"]

LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT)

WORDGEN_XLS_PATH_LIST = ["__info", "Venus_SoC_Memory_Mapping.xls"]
WORDGEN_TPL_PATH_LIST = [".Template", "WordGen.docx"]


def generate_template(search_sheet=None, exclude_sheet=None):
    if search_sheet:
        Excel2Json.Search_Module_List = search_sheet

    if exclude_sheet:
        Excel2Json.Exclude_Module_Tuple = exclude_sheet

    wb = xlrd.open_workbook(os.path.join(os.getcwd(), "__info", "Venus_SoC_Memory_Mapping.xls"))
    logging.debug("Word sheet number: %d" % int(wb.nsheets))
    Excel2Json.wordgen_excel2json(wb)
    Excel2Json.wordgen_remodel_json()

    doc = Document(os.path.join(os.getcwd(), "__info", "Empty_Ref.docx"))
    Json2Temp.wordgen_json2temp(doc)
    doc.save(os.path.join(os.getcwd(), ".Template", "WordGen.docx"))


def update_template(search_sheet=None, exclude_sheet=None):
    if search_sheet:
        Excel2Json.Search_Module_List = search_sheet

    if exclude_sheet:
        Excel2Json.Exclude_Module_Tuple = exclude_sheet

    wb = xlrd.open_workbook(os.path.join(os.getcwd(), "__info", "Venus_SoC_Memory_Mapping.xls"))
    logging.debug("Word sheet number: %d" % int(wb.nsheets))
    Excel2Json.wordgen_excel2json(wb)
    Excel2Json.wordgen_remodel_json()

    tpl = DocxTemplate(os.path.join(os.getcwd(), ".Template", "WordGen.docx"))
    Temp2Word.wordgen_temp2word(tpl)


def excel_content_verify(search_sheet=None, exclude_sheet=None):
    if search_sheet:
        Excel2Json.Search_Module_List = search_sheet

    if exclude_sheet:
        Excel2Json.Exclude_Module_Tuple = exclude_sheet

    wb = xlrd.open_workbook(os.path.join(os.getcwd(), "__info", "Venus_SoC_Memory_Mapping.xls"), formatting_info=True)
    logging.debug("Word sheet number: %d" % int(wb.nsheets))

    Excel2Json.wordgen_excel_verify(wb, True)
