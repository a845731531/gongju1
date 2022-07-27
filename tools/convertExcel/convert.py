#coding=utf-8
import os
import sys
import xlrd
import math
import json
import re

SCRIPT_FOLDER, _ = os.path.split(sys.argv[0])
EXCEL_PATH = ".\\Excel"
OUT_CONFIG_PATH = os.path.join(SCRIPT_FOLDER, "../../assets/Script/config")

def format_value(cell_value, cell_type, row):
    try:
        value = str(cell_value).strip()
        if(cell_type == "number"):
            value = float(value)
            if math.floor(value) == math.ceil(value):
                value = math.floor(value)
        elif(cell_type == "boolean"):
            value = bool(value)
        elif(cell_type == "string"):
            value = str(value)
        elif(cell_type == "string[]"):
            if value == "":
                value = []
            else:
                value = value.split("|")
            
        elif(cell_type == "number[]"):
            if value == "":
                value = []
            else:
                value_list = value.split("|")
                value = []
                for item_value in value_list:
                    number_value = float(item_value)
                    if math.floor(number_value) == math.ceil(number_value):
                        number_value = math.floor(number_value)
                    value.append(number_value)
        return value
    except:
        print("\ttype error:" + str(row) + ", " + cell_type)
        return cell_value

def write_config(config_name, config_data, key_names, key_types):
    config_file = os.path.join(OUT_CONFIG_PATH, config_name + ".ts")
    class_name = config_name + "Data"
    with open(config_file, "w", encoding='utf-8') as file_obj:
        file_obj.write("type {0} = {{\n".format(class_name))
        for i in range(0, len(key_names)):
            key_name = key_names[i]
            if(key_name == ""):
                continue
            key_type = key_types[i]
            file_obj.write("\t{0}: {1},\n".format(key_name, key_type))
        
        file_obj.write("}}\nlet {0}: {{ [id: string]: {1} }} = ".format(config_name, class_name))
        file_obj.write(json.dumps(config_data, indent = 4, ensure_ascii=False))
        file_obj.write("\nexport default {0};".format(config_name))

def convert_file(file_path):
    _, file_name = os.path.split(file_path)
    name_list = file_name.split("-")
    if len(name_list) < 3:
        return
    config_name = name_list[2].replace(".xlsx", "")
    if re.match("^[a-zA-Z]+$", config_name) == None:
        return
    
    print("converting " + file_name)
    xlrd.Book.encoding = "utf-8"
    excel_file = xlrd.open_workbook(file_path)
    data_sheet = excel_file.sheet_by_index(0)
    key_sheet = excel_file.sheet_by_index(1)

    key_names = []
    key_types = []
    target_data = {}

    col_num = key_sheet.ncols
    for col in range(col_num):
        key_names.append(str(key_sheet.cell_value(1, col)).strip())  #row 2 is key name
        key_types.append(str(key_sheet.cell_value(2, col)).strip())  #row 3 is key type

    rol_num = data_sheet.nrows
    for row in range(1, rol_num):
        cell_value = data_sheet.cell_value(row, 0)
        target_key = format_value(cell_value, key_types[0], row)
        item_data = {}
        for col in range(col_num):
            if key_names[col] == '':
                continue
            cell_value = format_value(data_sheet.cell_value(row, col), key_types[col], row)
            item_data[key_names[col]] = cell_value
        target_data[target_key] = item_data
    
    write_config(config_name, target_data, key_names, key_types)

def main():
    global EXCEL_PATH
    if len(sys.argv) > 1:
        EXCEL_PATH = sys.argv[1]
    # for parent, folders, files in os.walk(EXCEL_PATH):
    for file in os.listdir(EXCEL_PATH):
        if not file.startswith("~") and file.endswith(".xlsx"):
            full_path = os.path.join(EXCEL_PATH, file)
            convert_file(full_path)


if __name__ == "__main__":
    main()