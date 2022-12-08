import os
import re
import pandas as pd
from pathlib import Path
import datetime


if __name__ == '__main__':
    print("Смотрю поступившие Платежные поручения в папке input")
    number_gp = []
    try:
        path_main = Path.cwd()
        path_xlsx = Path.cwd() / 'input'
        os.chdir(path_xlsx)
        files = [x for x in os.listdir() if x.endswith('.xlsx')]
        for fs in files:
            print(f"Обрабатываю файл {fs}")
            db = pd.ExcelFile(fs, engine='openpyxl')
            dfs = {x: db.parse(x) for x in db.sheet_names}
            for jj in dfs.keys():
                if "№" in jj:
                    col = dfs[jj].columns.tolist()[0]
                    st = dfs[jj].loc[27, col]
                    gp = re.findall(r'gp\d{10}', st)
                    if len(gp) != 0:
                        number_gp.append((gp[0], jj, fs))
                    else:
                        number_gp.append(('', jj, fs))
        print("Сохраняю в файл результаты")
        new = pd.DataFrame()
        new["ОДДС"] = [x[0] for x in number_gp]
        new["имя листа"] = [x[1] for x in number_gp]
        new["файл"] = [x[2] for x in number_gp]
        year, month, day = datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day
        name_new = f"Полученные ПП {day}-{month}-{year}.xlsx"
        new.to_excel(name_new, index=False)
        print(f"Сформирован файл {name_new}")
    except BaseException as err:
        print(err)
        input("Программа завершилась с ошибкой. Нажмите Enter")

    input("Нажмите Enter")
