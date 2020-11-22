# -*- coding: utf-8 -*-
"""
Created on Mon Oct 26 18:19:29 2020

@author: ADIV
"""

import pandas as pd
import xlrd



def readData():
    workbook = xlrd.open_workbook('Mahasiswa.xls')
    sheet = workbook.sheet_by_index(0)
    dataTest = [sheet.row_values(rowx) for rowx in range(sheet.nrows)]
    return dataTest

def rumus(a,b,c,d):
    return (a-b)/(c-d)

def low(x,a,b):
    if(x <= a):
        return 1
    elif (a < x < b):
        return rumus(b,x,b,a)
    else:
        return 0

def med(x,a,b,c,d):
    if x <= a or x > d:
        return 0
    elif a < x < b:
        return rumus(x,a,b,a)
    elif b <= x <= c:
        return 1
    else:
        return rumus(d,x,d,c)

def high(x,a,b):
    if x <= a:
        return 0
    elif a < x < b:
        return rumus(x,a,b,a)
    else:
        return 1
    

def FuzzyPenghasilan(dataMhs):
    for data in dataMhs:
        data.append((low(float(data[1]),7,9)))
        data.append((med(float(data[1]),6,8,12,14)))
        data.append((high(float(data[1]),11,16)))
        
def FuzzyPengeluaran(dataMhs):
    for data in dataMhs:
        data.append((low(float(data[2]),4,6)))
        data.append((med(float(data[2]),5,6,8,9)))
        data.append((high(float(data[2]),8,10)))
    
def Inference(dataMhs):
    for data in dataMhs:
        data.append(max(min(data[4],data[6]),min(data[5],data[6]),min(data[5],data[7])))
        data.append(max(min(data[3],data[6]),min(data[4],data[7]),min(data[5],data[8])))
        data.append(max(min(data[3],data[7]),min(data[3],data[8]),min(data[4],data[8])))

    
def deFuzzy(dataMhs):
    for data in dataMhs:
        defuzzy = ((data[9] * 50) + (data[10] * 70) + (data[11] * 100)) / (data[9]+data[10]+data[11])
        data.append(defuzzy)
        
def getdeFuzzy(data):
    return data[12]

dataMhs = readData()
dataMhs.pop(0)
FuzzyPenghasilan(dataMhs)
FuzzyPengeluaran(dataMhs)
Inference(dataMhs)
deFuzzy(dataMhs)
dataMhs.sort(key=getdeFuzzy,reverse=True)
arrHasil = []

for i in dataMhs[0:20]:
    arrHasil.append(int(i[0]))


Hasil = pd.DataFrame(arrHasil)
Hasil.columns = ['Id']
Hasil.to_excel('Bantuan.xls',index=False)
print(Hasil)
print('Data telah diekspor dengan nama file Bantuan.xls')
    
        