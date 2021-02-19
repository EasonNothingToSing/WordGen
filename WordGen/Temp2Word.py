import json
import os
import logging


def wordgen_temp2word(tpl):
    with open(os.path.join(os.getcwd(), "__info", "__ex2js.remodel.json"), "r") as fr:
        handle = json.load(fr)
        logging.debug("Read __ex2js.remodel.json file")

    logging.info("Start render template")
    tpl.render(handle)
    logging.info("End render")

    tpl.save(os.path.join(os.getcwd(), "ListenAI_Doc.docx"))


if __name__ == "__main__":
    wordgen_temp2word()
