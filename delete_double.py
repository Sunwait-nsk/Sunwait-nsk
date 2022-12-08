import re
from functools import reduce
import pandas as pd
import os
from pathlib import Path
import warnings
import numpy as np
from collections import Counter


def part_family(name):
    f, nn, p = '', '', ''
    # разбиение фамилии на фио
    if name.count(' ') >= 1:
        n = name.split(' ')
        if len(n) == 1:
            f = n[0]
        elif len(n) == 2:
            f = n[0]
            nn = n[1]
        elif len(n) == 3:
            f = n[0]
            nn = n[1]
            p = n[2]
        elif len(n) == 4:
            f = n[0]
            nn = n[1]
            p = f'{n[2]} {n[3]}'
        elif len(n) == 5:
            f = n[0]
            nn = n[1]
            p = f'{n[2]} {n[3]} {n[4]}'
    return [f, nn, p]


def un_d(x):
    if ' ' in x[0]:
        x1 = x[0].replace(' ', '')
    else:
        x1 = x[0]
    x2 = x[1]
    return f'{x1} {x2}'


def match(text, alphabet=set('абвгдеёжзийклмнопрстуфхцчшщъыьэюя')):
    return not alphabet.isdisjoint(text.lower())


def has_latin(text):
    return bool(re.search('[a-zA-Z]', text))


def trans_doc(s_doc, n_doc, date_doc, ii_doc):
    result = ['' for _ in s_doc]
    for i, el in enumerate(s_doc):
        if not (el == '' and n_doc[i] == '' and date_doc[i] == '' and ii_doc[i] == ''):
            if '-' in el:
                el1 = el.split('-')[0]
                el2 = el.split('-')[1]
            else:
                el1 = el
                el2 = ''
            if "ЗАГС" in ii_doc[i] and el2 != '' and has_latin(el1) and match(el2):
                result[i] = 'Свидетельство о рождении'
            elif el2 != '' and has_latin(el1) and match(el2):
                result[i] = 'Свидетельство о рождении'
            elif len(f'{el} {n_doc[i]}') == 12 and int(date_doc[i].split('-')[0]) <= 1991:
                result[i] = "Паспорт СССР"
            elif len(f'{el} {n_doc[i]}') == 12 and int(date_doc[i].split('-')[0]) > 1991:
                result[i] = "Паспорт РФ"
            elif el == '' and 'записи актов гражданского состояния' in ii_doc[i] and int(date_doc[i].split('-')[0]) > 2006:

                result[i] = 'Свидетельство о рождении'
            elif len(f'{el} {n_doc[i]}') == 11 and int(date_doc[i].split('-')[0]) > 1991 and \
                    ("МВД" in ii_doc[i] or "Отделением УФМС" in ii_doc[i] or "УВД" in ii_doc[i] or
                     "Управлением внутренних дел" in ii_doc[i] or "ОВД" in ii_doc[i] or "отделением УФМС" in ii_doc[i] or
                     "Отделом Внутренних дел" in ii_doc[i] or "РОВД" in ii_doc[i] or "УФМС" in ii_doc[i] or
                     "Управление внутренних дел" in ii_doc[i]):
                result[i] = "Паспорт РФ"
            else:
                result[i] = "Паспорт иностранного государства, Прочее"

    return result


def open_file1(name):
    print("Читаются данные с Листа 'Актуальные собственники' ")
    db = pd.read_csv("Актуальные собственники.csv", delimiter=';').replace(np.nan, '')
    print("Читаются данные по ЕГРН")
    db1 = pd.read_csv('ЕГРН.csv', delimiter=';').replace(np.nan, '')
    family = db["ФИО собственника"].tolist()
    fam = [part_family(x) for x in family]
    surname = [x[0] for x in fam]
    name = [x[1] for x in fam]
    part = [x[2] for x in fam]
    birth_day = db["Дата рождения"].tolist()
    address = db["Адрес помещения"].tolist()
    address_reg = db["Адрес фактической регистрации"].tolist()
    type = db["Тип собственности"].tolist()
    dol = db["Доля в собственности"].tolist()
    lc = db["Номер ЛС"].tolist()
    snils = db["СНИЛС"].tolist()
    kuvi = db["Номер актуальной выписки из ЕГРН"].tolist()
    priznak = ['нет' for _ in range(len(db))]

    fio = db1["ФИО"].tolist()
    KUVI2 = db1["registration_number"].tolist()
    date_b = db1["birth_date"].tolist()
    place_b = db1["birth_place"].tolist()
    SNILS2 = db1["snils"].tolist()
    ser_d = []
    for e in db1["document_series"].tolist():
        if e.startswith('серия '):
            ser_d.append(e[6:])
        else:
            ser_d.append(e)
    numv_doc = [un_d(x) for x in zip(ser_d, db1["document_number"].tolist())]

    date_doc = db1["document_date"].tolist()
    iss_doc = db1["document_issuer"].tolist()
    type_doc = trans_doc(ser_d, db1["document_number"].tolist(), date_doc, iss_doc)
    result = [family, surname, name, part, birth_day, lc, snils, kuvi, priznak, address, type, dol,
              address_reg, fio, KUVI2, date_b, place_b, SNILS2, numv_doc, date_doc, iss_doc, type_doc, db]
    return result


def add_date1(bd, f, i, o, b, sn, adr, pl, t_doc, n_doc, d_doc, i_doc):
    n = pd.DataFrame([['', f, i, o, b, '', sn, '', '', '', '', '', '', '', adr, d_doc, t_doc, n_doc, pl, i_doc, '']],
                     columns=['ID', 'LAST_NAME', 'FIRST_NAME', 'MIDDLE_NAME', 'BIRTH_DATE', 'SEX', 'RETIREMENT_INS_NUM',
                              'PHONE', 'WORK_PLACE', 'MODIFIED_TIME', 'DEATH_DATE', 'CITIZENSHIP_ID', 'ID_SOCIAL',
                              'INN', 'REG_ADR', 'ISSUANCE_DATE', 'DOCUMENT_TYPE', 'DOCUMENT_NUMBER', 'HOME_TOWN',
                              'DOCUMENT_ISSUED', 'MODIFIED_TIME.1'])
    bd = pd.concat([bd, n], sort=False, axis=0)
    return bd


def add_journal(bd, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10, param11, param12,
                count, param13, param14, param15, param16, param17, param18, param19, param20):
    n = pd.DataFrame([[param5, param9, f'{param1} {param2} {param3}', param10, param11, param6, param4, param12, param7,
                       param13, param14, param15, param16, param17, param18, param19, param20, count, param8]],
                     columns=['Номер ЛС', 'Адрес помещения', 'ФИО собственника', 'Тип собственности',
                              'Доля в собственности', 'СНИЛС', 'Дата рождения', 'Адрес фактической регистрации',
                              'Номер актуальной выписки из ЕГРН', 'ФИО из ЕГРН', 'Номер выписки', "Дата рождения",
                              'Место рождения', 'СНИЛС', 'Номер документа', "Дата документа", 'Кем выдан',
                              'Число повторов ФИО в Акт собс', 'Признак записи'])
    bd = pd.concat([bd, n], sort=False, axis=0)
    return bd


def add_row(ind_fio, list_parametr12, list_parametr13, list_parametr14, list_parametr15, list_parametr16,
            list_parametr17, list_parametr18, list_parametr19, list_parametr20, list_parametr21, new, journal,
            list_parametr1, list_parametr2, list_parametr3, list_parametr4, list_parametr5, list_parametr6,
            list_parametr7, list_parametr8, list_parametr9, list_parametr10, list_parametr11, count_fam):
    pl_reg = list_parametr12
    if len(ind_fio) != 0:

        famil = list_parametr13[ind_fio[0]]
        kuvi = list_parametr14[ind_fio[0]]
        data_bir = list_parametr15[ind_fio[0]]
        place = list_parametr16[ind_fio[0]]
        snil = list_parametr17[ind_fio[0]]
        number_doc = list_parametr18[ind_fio[0]]
        data_doc = list_parametr19[ind_fio[0]]
        issuer_doc = list_parametr20[ind_fio[0]]
        type_document = list_parametr21[ind_fio[0]]

    else:
        famil, kuvi, data_bir, place, snil, number_doc, data_doc, issuer_doc, type_document = '', '', '', '', '', '', \
                                                                                              '', '', ''
    new = add_date1(new, list_parametr1, list_parametr2, list_parametr3, data_bir, snil, pl_reg, place,
                    type_document, number_doc, data_doc, issuer_doc)
    journal = add_journal(journal, list_parametr1, list_parametr2, list_parametr3, list_parametr4, list_parametr5,
                          list_parametr6, list_parametr7, list_parametr8, list_parametr9, list_parametr10,
                          list_parametr11, list_parametr12, count_fam, famil, kuvi, data_bir, place, snil, number_doc,
                          data_doc, issuer_doc)
    return new, journal


def add_row1(ind_fio, list_parametr12, list_parametr13, list_parametr14, list_parametr15, list_parametr16,
            list_parametr17, list_parametr18, list_parametr19, list_parametr20, journal, list_parametr1,
            list_parametr2, list_parametr3, list_parametr4, list_parametr5, list_parametr6, list_parametr7,
             list_parametr8, list_parametr9, list_parametr10, list_parametr11, count_fam):
    if len(ind_fio) != 0:
        famil = list_parametr13[ind_fio[0]]
        kuvi = list_parametr14[ind_fio[0]]
        data_bir = list_parametr15[ind_fio[0]]
        place = list_parametr16[ind_fio[0]]
        snil = list_parametr17[ind_fio[0]]
        number_doc = list_parametr18[ind_fio[0]]
        data_doc = list_parametr19[ind_fio[0]]
        issuer_doc = list_parametr20[ind_fio[0]]
    else:
        famil, kuvi, data_bir, place, snil, number_doc, data_doc, issuer_doc, type_document = '', '', '', '', '', '', \
                                                                                              '', '', ''
    journal = add_journal(journal, list_parametr1, list_parametr2, list_parametr3, list_parametr4, list_parametr5,
                          list_parametr6, list_parametr7, list_parametr8, list_parametr9, list_parametr10,
                          list_parametr11, list_parametr12, count_fam, famil, kuvi, data_bir, place, snil, number_doc,
                          data_doc, issuer_doc)
    return journal


def find_FIO_kuvi(lst_kuvi, lst_fio, fam, kuvi):
    if fam in lst_fio and kuvi in lst_kuvi:
        res = [i for i, x in enumerate(zip(lst_fio, lst_kuvi)) if x == (fam, kuvi)]
    else:
        res = []
    return res


if __name__ == "__main__":
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        path_main = Path.cwd()
        path_file = input("Введите путь до файл: ")
        name_file = input("Введите название файла: ")
        name_csv = f'{name_file[:-4]}.csv'
        print("Открываем файл для обработки")
        if name_csv in os.listdir():
            print("Найден файл csv с таким именем (Файл формата csv увеличивает скорость анализа данных, "
                  "если его нет, то он будет создан)\n")
            y = input("Использовать найденный файл? y/n: ")
            if y == "y":
                db = pd.read_csv(name_csv, sep=';').replace(np.nan, '')
            else:
                db = pd.read_excel(name_file, engine='openpyxl').replace(np.nan, '')
                db.to_csv(name_csv, sep=';', index=False, encoding='utf-8')
        else:
            db = pd.read_excel(name_file, engine='openpyxl').replace(np.nan, '')
            db.to_csv(name_csv, sep=';', index=False, encoding='utf-8')
        db = pd.read_csv(name_csv, sep=';').replace(np.nan, '').astype(str)
        print(f'Количество записей до удаления копий {len(db)}')
        columns_file = db.columns.tolist()
        res_lst = [' '.join(db.loc[i, :]) for i in range(len(db))]
        count_zap = Counter(res_lst)

        if max([count_zap[x] for x in count_zap]) == 1:
            print("Дублей строк в файле не найдено")
        else:
            dict_f = {x: [] for i, x in enumerate(res_lst) if count_zap[x] > 1}
            for i, x in enumerate(res_lst):
                if x in dict_f:
                    dict_f[x].append(i)

        del_i = reduce(lambda x, y: x+y, [x[1:] for x in dict_f.values()])
        new = pd.DataFrame(columns=columns_file).astype(str)
        for jj, col in enumerate(columns_file):
            per = [str(x) for i, x in enumerate(db[col].tolist()) if not (i in del_i)]
            new[col] = per
        print(f'Количество записей до удаления копий {len(new)}')
        new.to_excel('result.xlsx', index=False, encoding='utf-8')
        input(f"Сформирован файл без дублирующих строк result.xlsx.... Нажмите любую клавишу")





