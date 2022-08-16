import xlrd
from os.path import join, dirname, abspath
import os


def get_sheets(path):
    book = xlrd.open_workbook(path)
    os.environ['date_mode'] = str(book.datemode)
    return book.sheets()


def get_sheet_names(file_name):
    work_book = xlrd.open_workbook(file_name)
    sheets = work_book.sheet_names()
    return {i.strip(): work_book.sheet_by_name(i) for i in sheets}

current_path = abspath(dirname(__file__))


def get_data_path(name):
    return join(current_path, '..', 'data', name)


def clear_json_files():
    for dir_path, dir_names, file_names in os.walk(get_data_path('.')):
        for file_name in file_names:
            name_partitioin = file_name.rpartition('.')
            if name_partitioin[2] == 'json':
                os.remove(get_data_path(file_name))