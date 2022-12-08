import os
from pathlib import Path
import re
import warnings
import comtypes.client
import pandas as pd
from borb.pdf import Document
from borb.pdf import PDF



def xlsx_to_pdf(SOURCE_DIR, TARGET_DIR, name):
    app = comtypes.client.CreateObject('Excel.Application')
    app.Visible = False
    infile = os.path.join(os.path.abspath(SOURCE_DIR), name)
    name_out = f'{name[:-4]}pdf'
    outfile = os.path.join(os.path.abspath(TARGET_DIR), name_out)
    doc = app.Workbooks.Open(infile)
    doc.ExportAsFixedFormat(0, outfile, 1, 0)
    doc.Close()
    app.Quit()
    return True


def translate_pdf(fs1):
    SOURCE_DIR = Path.cwd() / 'input'
    TARGET_DIR = Path.cwd() / 'pdf'
    path = os.getcwd()
    os.chdir(SOURCE_DIR)
    files = os.listdir()
    print("Идет процесс переноса файлов в PDF")
    os.chdir(path)
    for fs in files:
        if fs != fs1:
            print(f'Конвертируется файл - {fs}')
            os.chdir(TARGET_DIR)
            jhgj = f'{fs[:-4]}pdf'
            fil_pdf = [x for x in os.listdir()]
            try:
                if fil_pdf.count(jhgj) == 0:
                    os.chdir(path)
                    xlsx_to_pdf(SOURCE_DIR, TARGET_DIR, fs)
            except BaseException as err:
                pass
    os.chdir(path)


def pd_read(fs):
    # чтение реестра
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        db = pd.read_excel(fs, engine="openpyxl")
        try:

            col_num = [x for x in db.columns.tolist() if x.startswith("NUM")][0]
            col_file = [x for x in db.columns.tolist() if x.startswith("FILE_NAME")][0]
        except BaseException as err:
            print("Несоответствие названий столбцов в реестре ")
            input("Нажмите Enter")
        gb = db[col_num].tolist()
        file_names = db[col_file].tolist()
    return gb, file_names


def read_all(name_f, odds, n_f, fs):
    pl = []
    files = [x for x in os.listdir() if x != name_f]
    reestr_all = {}
    dfses = []
    number_gp = []
    res = {}
    text_gp = ''
    text_ok = ''
    for i_f in files:
        try:
            xl_file = pd.ExcelFile(i_f, engine='openpyxl')
            dfs = {sheet_name: xl_file.parse(sheet_name) for sheet_name in xl_file.sheet_names}
        except BaseException as err:
            print(f"Битый файл {i_f}")
        for i, k in enumerate(dfs.keys()):
            if "№" in k:
                ttt = dfs[k]
                col = ttt.columns.tolist()[0]
                st = ttt.loc[27, col]
                try:
                    gp = re.findall(r'gp\d{10}', st)[0]
                    number_gp.append(gp)
                    res[gp] = [i_f, i, k]
                    text_ok += f'{gp} - {k} - {i} - {i_f}\n'
                except IndexError as err:
                    text_gp += f'{k} - {i} - {i_f}\n'
        dfses.append(dfs)
    for i, x in enumerate(odds):
        try:
            reestr_all[x] = [res[x], n_f[i]]
            pl.append("yes")
        except KeyError as err:
            reestr_all[x] = ['', ['', '', '']]
            pl.append("No")
    sss = os.getcwd()
    os.chdir('..')
    with open('error_gp.txt', 'w') as f:
        f.write(text_gp)
    with open('ok_gp.txt', 'w') as f:
        f.write(text_ok)
    os.chdir(sss)
    save_reestr(pl, fs)
    return reestr_all, dfses


def find_odds(name_reestr):

    print(f"Читаю реестр {name_reestr}")
    odds, file_names = pd_read(name_reestr)
    print("...Формирую словарик по всем файлам ПП")
    reestr_all, dfses = read_all(name_reestr, odds, file_names, name_reestr)
    return odds, file_names, reestr_all


def save_reestr(pl, fs):
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        db = pd.read_excel(fs, engine="openpyxl")
        db["Наличие ПП"] = pl
        print(f"Сохраняем информацию о наличии ПП...{fs}\n Нажмите Enter")
        db.to_excel(fs, index=False)


def pdf_parts(odds, file_names, reestr_all):
    nn = set([reestr_all[x][0][0] for x in reestr_all.keys()])
    print("Разделяем файлы с ПП на отдельные файлы")
    os.chdir('..')
    new_part = Path.cwd()/'pdf'
    os.chdir(new_part)
    for i_n in nn:
        inm = f'{i_n[:-4]}pdf'
        print(f"Разделяем файл {inm}")
        with open(inm, "rb") as pdf_file_handle:
            input_pdf = PDF.loads(pdf_file_handle)
        old_name = {}
        new_name = {}

        for i, odd in enumerate(odds):

            if reestr_all[odd][0][0] == i_n:

                old_name[odd] = reestr_all[odd][0]
                new_name[odd] = reestr_all[odd][1]
                output_pdf_001 = Document()
                path = os.getcwd()
                os.chdir('..')
                source_dir = Path.cwd() / 'output'
                os.chdir(source_dir)
               # Разделение
                output_pdf_001.add_page(input_pdf.get_page(reestr_all[odd][0][1]))
                print(f"Формируем {new_name[odd]}")
                with open(new_name[odd], "wb") as pdf_out_handle:
                     PDF.dumps(pdf_out_handle, output_pdf_001)
                os.chdir('..')
                os.chdir(path)


if __name__ == '__main__':
    path = Path.cwd()
    print("Переносим файлы с ПП из формата xlsx в ПДФ")
    # перенос эксель файло в ПДФ
    name_reestr = input("укажите имя файла с реестром ПП: ")
    translate_pdf(name_reestr)
    path_input = Path.cwd() / 'input'
    os.chdir(path_input)
    odds, file_names, reestr_all = find_odds(name_reestr)
    pdf_parts(odds, file_names, reestr_all)

    print(f"Заберите файлы с директории: {(path / 'output')}")





