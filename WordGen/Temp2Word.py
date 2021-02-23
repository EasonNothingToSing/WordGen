import json
import os
import logging
from docx import Document
from docx.shared import Inches

__all__ = ["wordgen_temp2word", "WORDGEN_DOC_KEYW", "WORDGEN_DOC_KEYW_WIDTH"]
WORDGEN_DOC_KEYW = ("Name", "Bit", "Type", "Description", "Reset")
WORDGEN_DOC_KEYW_WIDTH = (Inches(1), Inches(0.7), Inches(0.6), Inches(3.5), Inches(0.7))


def __table_check(table):
    try:
        for cell in table.row_cells(0):
            if cell.text in WORDGEN_DOC_KEYW:
                pass
            else:
                return False
    except:
        return False

    return True


def __table_lock_width(table):
    table.autofit = False
    table.allow_autofit = False
    for row in table.rows:
        for idx, width in enumerate(WORDGEN_DOC_KEYW_WIDTH):
            row.cells[idx].width = width


def wordgen_temp2word(tpl):
    with open(os.path.join(os.getcwd(), "__info", "__ex2js.remodel.json"), "r") as fr:
        handle = json.load(fr)
        logging.debug("Read __ex2js.remodel.json file")

    logging.info("Start render template")
    tpl.render(handle)
    logging.info("End render")

    tpl.save(os.path.join(os.getcwd(), "ListenAI_Doc.docx"))

    doc = Document(os.path.join(os.getcwd(), "ListenAI_Doc.docx"))

    logging.info("Remodel cell width \r\n waiting...")
    for table in doc.tables:
        if __table_check(table):
            __table_lock_width(table)
        else:
            pass

    logging.info("Remodel success!!!")
    doc.save(os.path.join(os.getcwd(), "ListenAI_Doc.docx"))
