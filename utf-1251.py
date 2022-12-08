import pandas as pd
import os

path_file = input("Ввести путь до файла: ")
os.chdir(path_file)
name = input("Введите название файла")
db = pd.read_csv(name, sep=';', encoding='utf-8')
db.to_csv(name, sep=';', encoding='cp1251')
