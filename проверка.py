import os
import pandas as pd
import re
from collections import Counter


def trans_gr(txt: str, el, fs) -> str:
    t = re.findall(r'#gp\d{10}#', txt)
    if len(t) == 0:
        print(el, fs)
        return ''
    else:
        print(t[0][1:-1])
        return t[0][1:-1]


os.chdir("C:/Users/User/Desktop/work/Сбор пакета документов по оплатам ГП и задолженности-/2 этап. Формирование файлов с ПП\input")
db = pd.read_excel("ОДДС.xlsx", engine='openpyxl')
gb = db['NUM_ODDS'].tolist()
gh = ['No' for _ in range(len(gb))]
files = [x for x in os.listdir() if not x.startswith("ОДДС")]
sad = []
ss = 0
for fs in files:
    xl_file = pd.ExcelFile(fs)
    dfs = {sheet_name: xl_file.parse(sheet_name)
           for sheet_name in xl_file.sheet_names}
    ss += len(dfs)
    sad1 = []
    for el in dfs.keys():

        if el != "Лист1":


            yu = ' '.join([str(x) for x in dfs[el].loc[27, :].tolist()])
            if not 'gp' in yu:
                pass
                # print(el, fs)
            elem = trans_gr(yu, el, fs)
            sad.append(elem)
            gh[gb.index(elem)] = 'Yes'
db["Наличие"] = gh
rr = Counter(sad)
print(sad1)
db.to_excel("new.xlsx", index=False)
# print(sad)
# print(len(sad))
# print(ss)
#     if fs.endswith('_001.pdf'):
#         start_f = fs[:-8]
#         sad[fs[:-8]] = ['001']
#     elif not fs.endswith('_001.pdf') and fs.startswith(start_f):
#         sad[fs[:-8]].append(fs[-6:-4])
#
# for el in sad.keys():
#     if len(sad[el]) < 7:
#         if not '04' in sad[el]:
#             print(sad[el], el)


