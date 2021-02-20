import xlrd
from xlutils.copy import copy
import logging
import json
import os
import re
import random

__all__ = ["wordgen_excel2json", "wordgen_remodel_json", "wordgen_excel_verify", "Search_Module_List",
           "Exclude_Module_Tuple"]

Search_Module_List = None
Exclude_Module_Tuple = ("SysAddrMapping", "AP Peripheral AddrMapping", "CP Peripheral AddrMapping", "PDM2PCM")

Register_KeyCell_Column = {"SubAddr": None, "StartBit": None, "EndBit": None, "Default": None, "Property": None,
                           "RegisterName": None, "Description": None, "Alias": None}
Maps_Register_KeyCell = {"Module": None, "Address Start": None}

WordGen_List = []


def wordgen_keycell_locate(sh):
    global Register_KeyCell_Column
    t = 0
    for c in range(sh.ncols):
        v = sh.cell_value(0, c)
        if v == "Sub-Addr\n(Hex)":
            Register_KeyCell_Column["SubAddr"] = c
            t += 1
        elif v == "Start\nBit":
            Register_KeyCell_Column["StartBit"] = c
            t += 2
        elif v == "End\nBit":
            Register_KeyCell_Column["EndBit"] = c
            t += 4
        elif v == "Default\nValue":
            Register_KeyCell_Column["Default"] = c
            t += 8
        elif v == "R/W\nProperty":
            Register_KeyCell_Column["Property"] = c
            t += 16
        elif v == "Register\nName":
            Register_KeyCell_Column["RegisterName"] = c
            t += 32
        elif v == "Register Description":
            Register_KeyCell_Column["Description"] = c
            t += 64
        elif v == "Alias":
            Register_KeyCell_Column["Alias"] = c
            t += 128
        else:
            continue

    if t & 0x7f != 0x7f:
        logging.warning("[Warning] Key word locate error in %s" % (sh.name,))
        return False

    return True


def wordgen_excel2json_set_modules(sh, maps_sh):
    global WordGen_List, Maps_Register_KeyCell
    address = ""

    for c in range(maps_sh.ncols):
        if maps_sh.cell_value(0, c) in Maps_Register_KeyCell.keys():
            Maps_Register_KeyCell[str(maps_sh.cell_value(0, c))] = c

    for r in range(1, maps_sh.nrows):
        if maps_sh.cell_value(r, Maps_Register_KeyCell["Module"]) == sh.name:
            address = str(maps_sh.cell_value(r, Maps_Register_KeyCell["Address Start"]))
            break

    temp_dict = {"module": sh.name, "address": address.replace("_", ""), "registers": []}
    WordGen_List.append(temp_dict)


def alias_or_name(sh, dic, rows):
    if dic["Alias"]:
        return sh.cell_value(rows, dic["Alias"]) or sh.cell_value(rows, dic["RegisterName"])
    else:
        return sh.cell_value(rows, dic["RegisterName"])


def wordgen_excel2json_set_registers(sh):
    global Register_KeyCell_Column, WordGen_List
    temp_dict = {}
    _in_temp_dict = {}

    for r in range(1, sh.nrows):
        if sh.cell_value(r, Register_KeyCell_Column["SubAddr"]) != "":
            if temp_dict:
                # not empty
                WordGen_List[-1]["registers"].append(temp_dict)

            temp_dict = dict(register=alias_or_name(sh, Register_KeyCell_Column, r),
                             offset=sh.cell_value(r, Register_KeyCell_Column["SubAddr"]),
                             default=sh.cell_value(r, Register_KeyCell_Column["Default"]), filed=[])
        else:
            _in_temp_dict = dict(name=alias_or_name(sh, Register_KeyCell_Column, r),
                                 bit="%s:%s" % (int(sh.cell_value(r, Register_KeyCell_Column["EndBit"])),
                                                int(sh.cell_value(r, Register_KeyCell_Column["StartBit"]))),
                                 property=sh.cell_value(r, Register_KeyCell_Column["Property"]),
                                 default=sh.cell_value(r, Register_KeyCell_Column["Default"]),
                                 description=sh.cell_value(r, Register_KeyCell_Column["Description"]))
            temp_dict["filed"].append(_in_temp_dict)

    WordGen_List[-1]["registers"].append(temp_dict)


def wordgen_excel2json(wb):
    global Register_KeyCell_Column, Search_Module_List, Exclude_Module_Tuple, WordGen_List

    for i in range(wb.nsheets):
        # reset
        Register_KeyCell_Column = {"SubAddr": None, "StartBit": None, "EndBit": None, "Default": None, "Property": None,
                                   "RegisterName": None, "Description": None, "Alias": None}
        sh = wb.sheet_by_index(i)
        if Search_Module_List:
            if sh.name in Search_Module_List:
                pass
            else:
                continue

        if sh.name in Exclude_Module_Tuple:
            continue

        logging.debug("The sheet name: %s" % str(sh.name))
        logging.debug("The sheet rows: %d" % int(sh.nrows))
        logging.debug("The sheet column: %d \n" % int(sh.ncols))
        if wordgen_keycell_locate(sh):
            wordgen_excel2json_set_modules(sh, wb.sheet_by_name("AP Peripheral AddrMapping"))
            wordgen_excel2json_set_registers(sh)
        else:
            continue

    with open(os.path.join(os.getcwd(), "__info", "ex2js.json"), "w") as fw:
        json.dump(WordGen_List, fw)
        logging.info("Regenerate ex2js.json file")


def wordgen_remodel_json():
    remodel_dict = {"Module_Register_Col_Labels": ['Name', 'Bit', 'Type', 'Description', 'Reset']}

    with open(os.path.join(os.getcwd(), "__info", "ex2js.json"), "r") as fr:
        handle = json.load(fr)

    for item in handle:
        # module name
        str_key = "%s_Name" % item["module"]
        str_val = item["module"]
        remodel_dict.update({str_key: str_val})

        # module address
        str_key = "%s_Addrss" % item["module"]
        str_val = item["address"]
        remodel_dict.update({str_key: str_val})

        for in_item in item["registers"]:
            # module register name
            str_key = "%s_%s_Name" % (item["module"], in_item["register"])
            str_val = in_item["register"]
            remodel_dict.update({str_key: str_val})

            # module register offset
            str_key = "%s_%s_Offset" % (item["module"], in_item["register"])
            str_val = in_item["offset"]
            remodel_dict.update({str_key: str_val})

            reg_list = []
            # module register contents
            for filed in in_item["filed"]:
                reg_list.append([filed["name"], filed["bit"], filed["property"], filed["description"],
                                 filed["default"]])
            str_key = "%s_%s_contents" % (item["module"], in_item["register"])
            remodel_dict.update({str_key: reg_list})

    with open(os.path.join(os.getcwd(), "__info", "__ex2js.remodel.json"), "w") as fw:
        json.dump(remodel_dict, fw)
        logging.info("Regenerate __ex2js.remodel.json file")


# xlwt API无法直接读取cell的值或者style，所以才用内置static 函数
def __getcell_style(sheet, r, c):
    global Register_KeyCell_Column
    return sheet._Worksheet__rows.get(r)._Row__cells.get(c)


def wordgen_excel_verify(wb, amend=False):
    global Register_KeyCell_Column, Search_Module_List, Exclude_Module_Tuple
    xlu = copy(wb)

    for i in range(wb.nsheets):
        # reset
        Register_KeyCell_Column = {"SubAddr": None, "StartBit": None, "EndBit": None, "Default": None, "Property": None,
                                   "RegisterName": None, "Description": None, "Alias": None}
        sh = wb.sheet_by_index(i)
        xlu_sheet = xlu.get_sheet(i)
        if Search_Module_List:
            if sh.name in Search_Module_List:
                pass
            else:
                continue

        if sh.name in Exclude_Module_Tuple:
            continue

        # 定位关键字
        if wordgen_keycell_locate(sh):

            for r in range(1, sh.nrows):
                style = __getcell_style(xlu_sheet, r, 0).xf_idx

                temp_name = alias_or_name(sh, Register_KeyCell_Column, r)

                if re.match(r"^[a-zA-Z_][a-zA-Z0-9_]*$", temp_name):
                    pass
                else:
                    logging.warning("%s sheet, row: %s, column: %s      Error string: %s" % (sh.name, r+1,
                                                                                             Register_KeyCell_Column
                                                                                             ["RegisterName"],
                                                                                             temp_name))
                    if amend:
                        # 如果为非空格
                        if temp_name:
                            # 将异常命名中的错误去除
                            temp_name = re.sub(r"[\[|<][a-zA-Z0-9_:]+?[>|\]]", "", temp_name.strip())
                        else:
                            temp_name = "Empty%d" % random.randint(1, 100000)

                        logging.warning("Change to %s" % temp_name)

                        # 如果检测出“Alias”为空
                        if not Register_KeyCell_Column["Alias"]:
                            # 在最后一列右侧添加一列
                            xlu_sheet.write(0, sh.ncols, "Alias")
                            __getcell_style(xlu_sheet, 0, sh.ncols).xf_idx = style

                        xlu_sheet.write(r, Register_KeyCell_Column["Alias"] or sh.ncols, temp_name)
                        __getcell_style(xlu_sheet, r, Register_KeyCell_Column["Alias"] or sh.ncols).xf_idx = style

    # 保存结果（将原有一份excel备份后再保存）
    xlu.save(os.path.join(os.getcwd(), "__info", "Venus_SoC_Memory_Mapping.xls"))


if __name__ == "__main__":
    wb = xlrd.open_workbook("../__info/Venus_SoC_Memory_Mapping.xls", formatting_info=True)
    logging.debug("Word sheet number: %d" % int(wb.nsheets))

    wordgen_excel_verify(wb, True)
