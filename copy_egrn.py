import os
from pathlib import Path
import warnings
import shutil

import pandas as pd


class Reestr:
    #RECIPIENT_ID	COMP_ID_PIR	CONTRACT_NUMBER	FILE_NAME_PP	FILE_NAME_EGRN	NUM_ODDS	EGRN_NUMBER	SIGNER_LAST_NAME


    def __init__(self, name_f):
        self.name: str = name_f
        self.count: int = 0
        self.columns: list = []
        self.odds: list = []
        self.file_name: list = []
        self.pr_egrn: list = ['No' for _ in range(self.count)]
        self.egrn_file = {}
        self.pr_pp: list = []
        self.pp_file = {}
        self.lc: list = []
        self.egrn_num: list = []
        self.recipient: list = []
        self.comp: list = []
        self.file_pp: list = []
        self.num_pp: list = []


    def read(self):
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            db = pd.read_excel(self.name, engine='openpyxl')
            db1 = pd.read_excel("Реестр номеров.xlsx", engine='openpyxl')
        self.count = len(db)
        self.columns = db.columns.tolist()
        self.odds = db["NUM_ODDS"].tolist()
        self.lc = db["CONTRACT_NUMBER"].tolist()
        self.file_name = [x[:-8] for x in db["FILE_NAME_PP"].tolist()]
        self.file_name_pp = db["FILE_NAME_PP"].tolist()
        self.file_name_egrn = db["FILE_NAME_EGRN"].tolist()
        self.egrn_num = db["EGRN_NUMBER"].tolist()
        self.recipient = db["RECIPIENT_ID"].tolist()
        self.comp = db["COMP_ID_PIR"].tolist()
        self.kuvi = db1["NUMBER_EGRN"].tolist()
        self.file_kuvi = db1["NUMBER_EGRN"].tolist()
        self.egrn_file = {x: '' for x in self.egrn_num}
        self.pp_file = {x: '' for x in self.odds}
        # 	DATE_EGRN	NAME_FILE

    def save(self):
        db = pd.DataFrame([], columns=self.columns)
        db["RECIPIENT_ID"] = self.recipient
        db["COMP_ID_PIR"] = self.comp
        db["CONTRACT_NUMBER"] = self.lc
        db["FILE_NAME"] = self.file_name
        db["FILE_NAME_PP"] = self.file_name_pp
        db["FILE_NAME_EGRN"] = self.file_name_egrn
        db["NUM_ODDS"] = self.odds
        db["EGRN_NUMBER"] = self.egrn_num
        db["Наличие выписки"] = self.pr_egrn
        db["Наличие ПП"] = self.pr_pp
        db["Старое название файла выписки"] = self.egrn_file
        db.to_excel(self.name, index=False)


class Reestr_kuvi:
    def __init__(self, name):
        self.name = name
        self.kuvi: list = []
        self.file_kuvi: list = []
        self.count = 0

    def read(self):
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            db1 = pd.read_excel("Реестр номеров.xlsx", engine='openpyxl')
        self.kuvi = db1["NUMBER_EGRN"].tolist()
        self.file_kuvi = db1["NAME_FILE"].tolist()
        self.count = len(self.kuvi)


class PP:

    def __init__(self, name, file, count, sod):
        self.name: str = name
        self.name_file: str = file
        self.count: int = count
        self.sod = sod
        self.gp: str = ''


if __name__ == '__main__':
    print("Формирование выписок. В папку с реестрами необходимо положить реестр с "
          "\nперечнем необходимых выписок. При запуске программа запроситназвание реестра и \n"
          "путь до самих файлов с выписками ЕГРН. После подбора и переименования выписки, данные о старых именах и "
          "\nналичии будут сохранены в файл реестра")
    input("Нажмите любую клавишу...")
    try:
        path_main = Path.cwd()
        path_rees = Path.cwd() /'Реестры'
        name_reestr = input("Введите название реестра: ")
        print("Читаю данные...")
        r1 = Reestr(name_reestr)
        r2 = Reestr_kuvi("Реестр номеров.xlsx")
        os.chdir(path_rees)
        r1.read()
        r2.read()
        # соотвествие выписок в наличии
        r1.pr_pp = ['' for _ in range(r1.count)]
        r1.egrn_file = [r2.file_kuvi[r2.kuvi.index(x)] if r2.kuvi.count(x) > 0 else '' for x in r1.egrn_num]
        r1.pr_egrn = ["Yes" if r2.kuvi.count(x) > 0 else 'No' for x in r1.egrn_num]
        os.chdir('..')
        path_pdf = input("Укажите путь до файлов с выписками: ")
        print("Готовлю данные...")
        path_pdf_new = Path.cwd() / 'Выписки ЕГРН'
        print("Копирую выписки")
        ii = 0
        for i, n in enumerate(r1.egrn_file):
            try:
                os.chdir(path_pdf)
                if os.path.exists(n) and n != '':
                    shutil.copy(
                        os.path.join(path_pdf, n),
                        os.path.join(path_pdf_new, r1.file_name_egrn[i])
                    )
                    ii += 1
                else:
                    pass
            except BaseException as err:
                pass
        os.chdir(path_main)
        os.chdir(Path.cwd()/'Реестры')
        r1.save()
    except BaseException as err:
        print(err)
        input("Нажмите любую клавишу")
    print(f"Файлы в количестве {ii} шт с выписками скопированы в папку {path_pdf_new}. \nРеестр обновлен данными")
    input("Нажмите любую клавишу")

