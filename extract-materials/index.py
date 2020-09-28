import pandas as pd
import numpy as np
import csv
import os

currentDirectory = os.getcwd()
path = r'C:\Users\engin\dev\disenos-magicos\Cortes'

def Display(rows, cols, width):
	pd.set_option('display.max_rows', rows)
	pd.set_option('display.max_columns', cols)
	pd.set_option('display.width', width)

Display(20000, 10, 2160)

report = []
for filename in os.listdir(path):
    data = []
    table = []
    sets = []
    df = pd.read_excel(path + '\\' + filename)
    df = df.replace(np.nan, '', regex=True)
    cols = [i+1 for i in range(49)]
    df.drop(df.columns[cols], axis=1, inplace=True)
    lookup = df[df.columns[1]]
    columns = len(df.columns)

    for i in range(len(lookup)):
        item = lookup[i]
        if item =='' and i > 0:
            lookup[i] = lookup[i-1]

    counter = 0
    for item in lookup:
        if item == 'Insumo':
            sets.append(counter)
        counter += 1

    counter = 0
    for item in lookup:
        if item == 'Terminos':
            sets.append(counter)
        counter += 1

    for x in range(len(sets)-1):
        start = sets[x]
        end = sets[x+1]
        for i in range(start, end):
            item = lookup[i]
            element = {}
            if (item != 'Insumo') and (item != 'Terminos'):
                for j in range(columns):
                    col = df[df.columns[j]]
                    trait = col[start]
                    value = col[i]
                    if (trait != '') and (trait != 'TotalR'):
                        if (value != '') and (value != 0):
                            element[trait] = value
            table.append(element)

    def checkForTi(element):
        for key in element.keys():
            if 'Ti' in key:
                return True
        return False

    def splitElement(element, k):
        obj = {}
        for key, value in element.items():
            if 'Ti' not in key:
                obj[key] = value
        obj['Talla'] = k
        obj['Cant1'] = element[k]
        return obj

    for element in table:
        if len(element) > 1:
            keys = element.keys()
            if 'Unidad' in keys:
                hasTi = checkForTi(element)
                if hasTi:
                    for key in keys:
                        if 'Ti' in key:
                            newElement = splitElement(element, key)
                            data.append(newElement)
                else:
                    data.append(element)

    def isNumber(s):
        try:
            float(s)
            return True
        except ValueError:
            return False

    def checkForValue(element):
        for value in element.values():
            if isNumber(value):
                return True
        return False

    for element in data:
        hasValue = checkForValue(element)
        if hasValue:
            if 'Talla' not in element.keys():
                element['Talla'] = 'No Aplica'

            if 'Colores' not in element.keys():
                element['Colores'] = 'No Aplica'

            element['OC'] = filename.split('.')[0]
            report.append(element)

    # print(report)

with open('Report.csv', 'w') as output:
    writer = csv.writer(output, lineterminator = '\n')
    headers = ['Insumo', 'Unidad', 'Colores', 'Talla', 'Cant1', 'OC']
    writer.writerow(headers)
    for item in report:
        row = []
        for header in headers:
            row.append(item[header])
        writer.writerow(row)

print('SUCCESS!')
