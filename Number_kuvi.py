import os
import zipfile
from zipfile import ZipFile
from pathlib import Path
import pandas as pd
import warnings
import patoolib
from PyPDF2 import PdfFileReader


def trans_to_pdf(path_to):
    path = Path.cwd()
  # проверяем наличие файлов. Должен быть 1 эксель и несколько пдф
    c_f = true_files(path_to)
    if c_f > 0:
        print("Обработка....")
        find_file(path_to)
        print(f"Файлы обработаны")
    else:
        print(f"\nУвы, ничего не получилось с обработкой файлов... Нет исходных файлов")
    return


def find_kuvi(text):
    kuvi = text.split('\n')
    for sym in kuvi:
        if "КУВИ-" in sym:
            f = sym.split(" ")
    return f[0][:10], f[2][:23]


def find_file(path_to_file):
    p_m = os.getcwd()
    name_kuvi = []
    data_kuvi = []
    name_file = []
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")

        os.chdir(path_to_file)
        files = [x for x in os.listdir() if x.endswith('.pdf')]
        for fs in files:
            name_file.append(fs)
            pdf_document = fs
            with open(pdf_document, "rb") as filehandle:
                pdf = PdfFileReader(filehandle)
                page = pdf.getPage(0)
                text = page.extractText()
            if "КУВИ-" in text:
                data_k, number_k = find_kuvi(text)
                name_kuvi.append(number_k)
                data_kuvi.append(data_k)
            else:
                name_kuvi.append('Не найден номер')
                data_kuvi.append('Не найден номер')

        os.chdir(p_m)
        xl_file = pd.DataFrame()
        xl_file["NUMBER_EGRN"] = name_kuvi
        xl_file['DATE_EGRN'] = data_kuvi
        xl_file['NAME_FILE'] = name_file
        xl_file.to_excel("Реестр номеров.xlsx", encoding='utf-8', index=False)
    return True


def true_files(path_to_f):
    p_m = os.getcwd()
    os.chdir(path_to_f)
    path = os.getcwd()
    fs = [x for x in os.listdir() if x.endswith('.pdf')]
    count_f1 = len(fs)
    if count_f1 == 0:
        print(f"Нет файлов для переименования в папке {path}")
        count_f1 = 0
    else:
        print(f"Исходные файлы дляобработки найдены в {path} в количестве - {count_f1}")
    os.chdir(p_m)
    return count_f1


if __name__ == '__main__':
    path_main = Path.cwd()
    path_files = input("Укажите где лежат файлы с выписками: ")
    print("Составление реестра номеро выписок и PDF... ")
    trans_to_pdf(path_files)
    input("Нажмите любую клавишу ... ")