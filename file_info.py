import pandas as pd
import numpy as np
import os
import warnings

if __name__ == '__main__':
    print("conversion process...")
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        print("Информация о файле Excel.")
        path_to = input("Введите путь до файла: ")
        os.chdir(path_to)

        file_name2 = input("Введите название файла без расширением (.xlsx): ") + '.xlsx'

        try:
            df = pd.ExcelFile(file_name2)
                # pd.read_excel(file_name2, engine='openpyxl').replace(np.nan, '', regex=True)
        except BaseException:
            print("Файл не найден")
        print("Листы книги EXCEL: \t")
        print(df.sheet_names)
        input("Для вывода следующей информации нажмите любую клавишу... ")
        for sheet in df.sheet_names:
            print('*'*20, f"\nИнформация о листе с именем {sheet}")
            db = pd.read_excel(file_name2, sheet, engine='openpyxl')
            print(f"Столбцы. {len(db.columns.tolist())} столбц/ов на листе {sheet}: {db.columns.tolist()}")
            print(f'Записей на листе: {len(db)}\n', '*'*20)
        input("Для закрытия нажмите любую клавишу... ")
