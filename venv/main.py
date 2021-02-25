from docxtpl import DocxTemplate
import openpyxl
import re
from openpyxl import load_workbook
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.Qt import QVBoxLayout
from ui import Ui_Form
import sys

def doc_generate():
    """Генерация документа"""
    doc = DocxTemplate("подряд1.docx")
    number = ui.lineEdit_number.text()
    number_of_month = ui.lineEdit_numberofmonth.text()
    month = ui.lineEdit_month.text()
    year = ui.lineEdit_year.text()
    full_org = ui.lineEdit_fullorg.text()
    director_name = ui.lineEdit_director_name.text()
    adress = ui.lineEdit_adress.text()
    inn = ui.lineEdit_inn.text()
    kpp = ui.lineEdit_kpp.text()
    short_org = ui.lineEdit_short_org.text()
    context = { 'Номер' : number, 'Число_месяца' : number_of_month, 'месяц': month, 'год': year, 'Полное_наименование': full_org,
                'Директор': director_name, 'адрес' : adress, 'ИНН': inn, 'КПП' : kpp, 'Наименование' : short_org}
    doc.render(context)
    doc.save("шаблон-final.docx")

#создание аппликации
app = QtWidgets.QApplication(sys.argv)
FLAG = 0
#инициализация формы
Form = QtWidgets.QWidget()
ui = Ui_Form()
ui.setupUi(Form)
Form.show()
index = 0
"""Подключение к базе xls"""
#book = openpyxl.open("база.xlsx", read_only=True)
wb = load_workbook('база.xlsx')
"""ЗАполнение комбобокса организациями"""
sheet_ranges = wb['Лист1']
column_a = sheet_ranges['A']
org = []
for i in range(len(column_a)):
    print(column_a[i].value)
    org.append(column_a[i].value)
print(org)
"""ПОлное наименование"""
column_b = sheet_ranges['B']
full_org = []
for i in range(len(column_b)):
    print(column_b[i].value)
    full_org.append(column_b[i].value)
print(full_org)
"""Директор"""
column_i = sheet_ranges['I']
director_name = []
for i in range(len(column_i)):
    print(column_i[i].value)
    director_name.append(column_i[i].value)
print(director_name)
"""АДрес"""
column_k = sheet_ranges['K']
adress = []
for i in range(len(column_k)):
    print(column_k[i].value)
    adress.append(column_k[i].value)
print(adress)
"""ИНН"""
column_e = sheet_ranges['E']
inn = []
for i in range(len(column_e)):
    print(column_e[i].value)
    inn.append(column_e[i].value)
print(inn)
"""КПП"""
column_f = sheet_ranges['F']
kpp = []
for i in range(len(column_f)):
    print(column_f[i].value)
    kpp.append(column_f[i].value)
print(kpp)
"""Сокращенное наименование"""
column_d = sheet_ranges['D']
short_org = []
for i in range(len(column_d)):
    print(column_d[i].value)
    short_org.append(column_d[i].value)
print(short_org)
ui.combox(org,full_org,director_name, adress, inn, kpp, short_org, index)
# column_d = sheet_ranges['D']
# name = []
# for i in range(len(column_d)):
#     print(column_d[i].value)
#     name.append(column_d[i].value)
# print(name)
ui.pushButton_generate_podryad.pressed.connect(doc_generate)
sys.exit(app.exec_())