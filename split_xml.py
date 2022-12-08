import os
import xml.dom.minidom as minidom
import warnings
from pathlib import Path
import xml.etree.ElementTree as ET


def parseXML(xml_file):
    """
    Парсинг XML используя ElementTree
    """
    tree = ET.ElementTree(file=xml_file)
    root = tree.getroot()
    dict_atrr = {}
    dict_atrr1 = {}
    i = 0
    tags_f = []
    print("Читаю содержимое, копирую теги...")
    et = ET.parse(xml_file)
    for e in et.iter():
        text = ET.tostringlist(e)
        tags_f.append([str(x) for x in text])

    for child in root:
        if child.tag == "WTFORMAT":
            for step_child in child:
                # print(step_child.tag)
                if step_child.tag == "DOCAZKAUBUCASHREQUESTS":
                    for step in step_child:
                        dict_atrr[i] = step.attrib

                        if step.tag == "DOCAZKAUBUCASHREQUEST":
                            for st1 in step:
                                if st1.tag == "LINES":
                                    for st2 in st1:
                                        dict_atrr1[i] = st2.attrib
                        i += 1
    return dict_atrr, dict_atrr1


def create_xml(dict_atrr, dict_attr1, name_file):
    if len(dict_atrr) % 499 == 0:
        count = len(dict_atrr)//499
    else:
        count = len(dict_atrr)//499 +1
    first = 0
    second = 0
    dict_count = {}
    for y in range(count):
        if y == count-1:
            second = len(dict_atrr)
        else:
            second += 499
        dict_count[y] = [first, second]
        first += 499
    print(f'Расчетное количество файлов - {len(dict_count)}')
    for jj in dict_count.keys():
        name = f'{name_file[:-4]}-{jj}.xml'
        print(f"Создаю новый файл {name}, в котором будут документы в диапозоне {dict_count[jj]}")
        root = ET.Element("EXTRWT")
        wtformat = ET.SubElement(root, "WTFORMAT", VERSION='2.4')

        docs = ET.SubElement(wtformat, "DOCAZKAUBUCASHREQUESTS")
        for i in range(dict_count[jj][0], dict_count[jj][1]):
            doc = ET.SubElement(docs, "DOCAZKAUBUCASHREQUEST", dict_atrr[i])
            lines = ET.SubElement(doc, 'LINES')
            line = ET.SubElement(lines, 'LINE', dict_attr1[i])
        etree = ET.ElementTree(root)
        ET.indent(etree, '  ')
        print("Сохраняю ...")
        etree.write(name, encoding='Windows-1251', xml_declaration=True, method='xml')


if __name__ == "__main__":
    with warnings.catch_warnings(record=True):
        print('Нужный для разделения файл, положи в папку input')
        input("Нажми любую клавишу....")
        warnings.simplefilter("always")
        path_main = Path.cwd()
        path_xml = Path.cwd() / 'input'
        path_out = Path.cwd() / 'output'
        os.chdir(path_xml)
        files_xml = [xm for xm in os.listdir() if xm.endswith(".xml")]
        for document in files_xml:
            print(f'Обрабатывается файл {document}')
            a1, a2 = parseXML(document)
            os.chdir(path_out)
            create_xml(a1, a2, document)
            os.chdir(path_xml)
        os.chdir(path_main)
        print(f"Готово. Смотри результат в {path_out}")
