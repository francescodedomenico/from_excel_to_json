#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import sys
import os
import codecs
import pandas as pd
from tqdm import tqdm
from os import write
from sys import argv


def string_filter(input_string):
    return_buffer = input_string
    if " \r\n " in return_buffer:
        return_buffer = return_buffer.replace(" \r\n ", " ")
    if " \r\n" in return_buffer:
        return_buffer = return_buffer.replace(" \r\n", " ")
    if "\r\n" in return_buffer:
        return_buffer = return_buffer.replace("\r\n", " ")
    if "\\u2019" in return_buffer:
        return_buffer = return_buffer.replace("\\u2019", "'")
    if "\\u00e0" in return_buffer:
        return_buffer = return_buffer.replace("\\u00e0", "á")

    return return_buffer


def get_sheet_data(fname, sheet_name):

    in_filename = fname

    df_sheet_name = pd.read_excel(in_filename, sheet_name=sheet_name)
    df_sheet_name.fillna("null", inplace=True)

    colonne = df_sheet_name.columns
    numero_righe = 0
    for column in colonne:
        df_sheet_name[column] = df_sheet_name[column].astype(str)
        if numero_righe < len(df_sheet_name[column]):
            numero_righe = len(df_sheet_name[column])

    lista_righe = list()
    for i in range(numero_righe):
        temp_riga = dict()
        for column in colonne:
            if isinstance(df_sheet_name[column][i], str):
                temp_riga[string_filter(column)] = string_filter(
                    df_sheet_name[column][i]
                )
            else:
                temp_riga[string_filter(column)] = df_sheet_name[column][i]
        lista_righe.append(temp_riga)
    return lista_righe


def get_excel_data(filename):
    sheet_names = get_sheet_names(filename)
    excel_data = dict()
    for sheet in tqdm(sheet_names):
        excel_data[sheet] = get_sheet_data(filename, sheet)
    return excel_data


def get_sheet_names(fname):
    in_filename = fname
    xls = pd.ExcelFile(in_filename)
    return xls.sheet_names


def write_to_json(to_json, filename):
    with codecs.getwriter("utf8")(open(filename, "wb")) as writeJSON:
        json.dump(eval(str(to_json)), writeJSON, indent=4, ensure_ascii=False)


def convert_excel_into_json(excel_file):
    excel_data = get_excel_data(excel_file)
    output_filename = excel_file.replace(".xlsx", "").replace(".xls", "") + ".json"
    write_to_json(excel_data, output_filename)


convert_excel_into_json(argv[1])
