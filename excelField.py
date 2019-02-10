# -*- coding: utf-8 -*-

import arcpy
import os
import xlwt
import xlrd
import numpy as np
import re

#функция, приводящая значения к типу str
def to_unicode(l):
    for i in range(1, len(l)):
        if type(l[i]) == long:
            l[i] = unicode(int(l[i]))
        elif type(l[i]) == float:
            l[i] = unicode(int(l[i]))
        elif type(l[i]) == int:
            l[i] = unicode(l[i])
        elif type(l[i]) == str:
            l[i] = unicode(l[i])       
        elif type(l[i]) == unicode:
            pass
        else:
            arcpy.AddMessage('invalid field type: '+l[i])
            l[i] = '!?: '+str(l[i])
    return l

#функция, удаляющая незначащие значения
def clean(l):
    for i in range(1, len(l)):
        if l[i]=='-' or l[i]=='—' or l[i]==u'МОГ':
            l[i] = ''
    return l

#функция, выявляющая пустые ячейки обязательных для заполнения полей 
def full(l):
    for i in range(1, len(l)):
        if l[i]=='' or l[i] is None:
            l[i] = '!?: '
    return l

#функция, заменяющая значения на соответствующие коды атрибутивного домена
def domen(l):
    for i in range(1, len(l)):
        if l[i].lower().find(u'не баланс') != -1:
            l[i] = 2
        elif l[i].lower().find(u'баланс') != -1:
            l[i] = 1
        elif l[i].lower().find(u'сельск') != -1:
            l[i] = 1
        elif l[i].lower().find(u'город') != -1:
            l[i] = 2
        elif l[i].lower().find(u'договор на то по заявке аб') != -1:
            l[i] = 2
        elif l[i].lower().find(u'договор на то') != -1:
            l[i] = 1
        elif l[i].lower().find(u'без договора на то') != -1:
            l[i] = 4
        elif l[i].lower().find(u'бесхозяйные огх') != -1:
            l[i] = 999
        elif l[i].lower().find(u'огх сторонних организаций') != -1:
            l[i] = 5
        elif l[i].lower().find(u'объекты министерства обороны') != -1:
            l[i] = 6
        elif l[i].lower().find(u'договор аренды') != -1:
            l[i] = 3
        elif l[i] == '':
            pass        
        else:
            arcpy.AddMessage('not in the domain: '+l[i])
            l[i] = '!?: '+l[i]
            
    return l

reestr_path_list = arcpy.GetParameterAsText(0).split(';')
dest_folder = arcpy.GetParameterAsText(1)

wb = xlwt.Workbook()
ws = wb.add_sheet('reestr',cell_overwrite_ok=True)

err_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
err2_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
err3_style = xlwt.easyxf('pattern: pattern solid, fore_colour green;')

tit = ['routeN', 'SGSegmN', 'ITDname', 'RightOwner', 'InventNum', 'commens', 'Contractor', 'passN', 'recordN', 'ContrNum', 'SGplase', 'KEY']

offset = 0 #сдвиг нумерации строк таблицы, присоединяемой к предыдущей
not_first = 0 #параметр, определяющий добавление заголовка только для первой таблицы из всех соединяемых

for file_location in reestr_path_list:
    
    workbook = xlrd.open_workbook(file_location)
    sheet = workbook.sheet_by_index(0)
    
    col_list = []

    col_list.append(to_unicode(full(clean(sheet.col_values(0)))))
    col_list.append(to_unicode(full(clean(sheet.col_values(3)))))
    col_list.append(to_unicode(full(clean(sheet.col_values(5)))))
    col_list.append(full(domen(clean(sheet.col_values(6)))))
    col_list.append(to_unicode(clean(sheet.col_values(7))))
    col_list.append(to_unicode(clean(sheet.col_values(8))))
    col_list.append(domen(clean(sheet.col_values(9))))
    col_list.append(to_unicode(full(clean(sheet.col_values(10)))))
    col_list.append(to_unicode(full(clean(sheet.col_values(11)))))
    col_list.append(to_unicode(clean(sheet.col_values(15))))
    col_list.append(full(domen(clean(sheet.col_values(16)))))

    #копируем значения в новую таблицу, заменяя значения на доменные коды и проверяя заполненность обязательных полей
    for i in range(0, len(col_list)):
        for j in range(0, len(col_list[i])-not_first):
            if unicode(col_list[i][j+not_first]).find('!?: ') != -1:
                ws.write(j+offset, i, col_list[i][j+not_first].replace('!?: ',''), err_style)
            else:
                ws.write(j+offset, i, col_list[i][j+not_first])
       
    #отмечаем желтым цветом несоответствие инвентарного номера балансовой принадлежности
    for i in range(0, len(col_list[3])-not_first):
        if col_list[3][i+not_first] == 1 and col_list[4][i+not_first] == '':
            ws.write(i+offset, 4, col_list[4][i+not_first], err2_style)
        elif col_list[3][i+not_first] == 2 and col_list[4][i+not_first] != '':
            ws.write(i+offset, 4, col_list[4][i+not_first], err2_style)

    #отмечаем желтым цветом несоответствие вида договора балансовой принадлежности        
    for i in range(0, len(col_list[3])-not_first):    
        if col_list[3][i+not_first] == 1 and col_list[6][i+not_first] != '':
            if unicode(col_list[6][i+not_first]).find('!?: ')!= -1:
                ws.write(i+offset, 6, col_list[6][i+not_first].replace('!?: ',''), err2_style)
            else:
                ws.write(i+offset, 6, col_list[6][i+not_first], err2_style)
        elif col_list[3][i+not_first] == 2 and col_list[6][i+not_first] == '':
            ws.write(i+offset, 6, col_list[6][i+not_first], err2_style)

    #отмечаем желтым цветом несоответствие номера договора балансовой принадлежности         
    for i in range(0, len(col_list[3])-not_first):    
        if col_list[3][i+not_first] == 1 and col_list[9][i+not_first] != '':
            ws.write(i+offset, 9, col_list[9][i+not_first], err2_style)
        elif col_list[3][i+not_first] == 2 and col_list[9][i+not_first] == '':
            ws.write(i+offset, 9, col_list[9][i+not_first], err2_style)

    #заполняем поле KEY конкатенацией номера паспорта и номера записи, отмечая зеленым неуникальные значения
    key_list = []
    for i in range(0, len(col_list[7])-not_first):
        key_value = col_list[7][i+not_first]+'+'+col_list[8][i+not_first]
        if key_value in key_list:
            ws.write(i+offset, 11, key_value, err3_style)
        else:
            ws.write(i+offset, 11, key_value)
        key_list.append(key_value)
    
    #меняем заголовок таблицы на имена полей атрибутивной таблицы
    if not_first == 0:
        for i in range(0, len(tit)):
            ws.write(0, i, tit[i])
    
    offset += len(col_list[0]) - not_first
    not_first = 1
    
wb.save(dest_folder+'//reestr_join.xls')