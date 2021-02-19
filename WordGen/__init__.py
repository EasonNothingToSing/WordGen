import logging
import os
import xlrd
from . import (Excel2Json, Json2Temp, Temp2Word)


LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT)

RUNNING_ENV = os.getcwd()


def generate_template():
    wb = xlrd.open_workbook(RUNNING_ENV + "__info/Venus_SoC_Memory_Mapping.xls")
    logging.debug("Word sheet number: %d" % int(wb.nsheets))
    Excel2Json.wordgen_excel2json(wb)
    Excel2Json.wordgen_remodel_json()
