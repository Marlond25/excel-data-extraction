import pandas as pd
import numpy as np
import os
import sys
import sqlite3 as sql
import csv

currentDirectory = os.getcwd()
addressForSupplies = r'C:\Users\engin\dev\disenos-magicos\calculate-materials\Insumos'
addressForOrders = r'C:\Users\engin\dev\disenos-magicos\calculate-materials\Pedidos.xlsx'
components = {}
accessories = {}
fabrics = {}

def Display(rows, cols, width):
	pd.set_option('display.max_rows', rows)
	pd.set_option('display.max_columns', cols)
	pd.set_option('display.width', width)

def getComponentsAndAccessories(file):
    df = pd.read_excel(file)
    df = df.replace(np.nan, '', regex=True)
    junk = [i for i in range(4)]
    df.drop(df.columns[junk], axis=1, inplace=True)
    criteria = df[df.columns[1]]
    isComponent = df.loc[(criteria == 'Insu') & (criteria != '')]
    isAccessory = df.loc[(criteria != 'Insu') & (criteria != '')]
    criteria = isComponent[isComponent.columns[3]]
    isSizeExtension = isComponent.loc[(criteria != '') & (criteria.str.contains('>') == False)]
    isComponent = isComponent.loc[criteria == '']
    criteria = isAccessory[isAccessory.columns[2]]
    isAccessory = isAccessory.loc[criteria == '']

    rows = len(isComponent.index)
    names = isComponent[isComponent.columns[2]].tolist()
    proportions = isComponent[isComponent.columns[4]].tolist()
    dimentions = isComponent[isComponent.columns[5]].tolist()
    cs = []
    for i in range(rows):
        if (proportions[i] != ''):
            c = {}
            c['name'] = names[i]
            c['proportion'] = proportions[i]
            c['dimention'] = dimentions[i]
            cs.append(c)

    rows = len(isAccessory.index)
    names = isAccessory[isAccessory.columns[1]].tolist()
    proportions = isAccessory[isAccessory.columns[3]].tolist()
    dimentions = isAccessory[isAccessory.columns[4]].tolist()
    ac = []
    for i in range(rows):
        if (proportions[i] != ''):
            a = {}
            a['name'] = names[i]
            a['proportion'] = proportions[i]
            a['dimention'] = dimentions[i]
            ac.append(a)

    componentsAndAccessories = {}
    componentsAndAccessories['components'] = cs
    componentsAndAccessories['accessories'] = ac
    return componentsAndAccessories

def getFabrics(file):
    df = pd.read_excel(file)
    df = df.replace(np.nan, '', regex=True)
    junk = [i for i in range(17)]
    df.drop(df.columns[junk], axis=1, inplace=True)
    criteria = df[df.columns[0]].tolist()
    counter = 1
    f = 0
    fbs = []

    for data in criteria:
        if (data == 'Texto149'):
            f = counter
        counter += 1

    for j in range(5):
        fb = {}
        name = df[df.columns[j]].tolist()[f]
        proportion = df[df.columns[j+10]].tolist()[f]
        if ((name != '') & (proportion != '')):
            fb['name'] = name
            fb['proportion'] = proportion
            fbs.append(fb)
    return fbs

def indexSupplies(components, fabrics):

    for filename in os.listdir(addressForSupplies):
        item = filename.split(' ')[0]
        type = filename.split(' ')[1].split('.')[0]
        file = addressForSupplies + '\\' + filename
        if (type == 'FT'):
            componentsAndAccessories = getComponentsAndAccessories(file)
            components[item] = componentsAndAccessories['components']
            accessories[item] = componentsAndAccessories['accessories']
        if (type == 'TELA'):
            fabrics[item] = getFabrics(file)

Display(20000, 10, 2160)
indexSupplies(components, fabrics)
orders = pd.read_excel(addressForOrders)
numberOfOrders = len(orders.index)

required = []
for i in range(numberOfOrders):
    order = orders.iloc[i]
    ref = order['item']
    requiredComponents = components[str(ref)]
    for com in requiredComponents:
        com['item'] = ref
        com['type'] = 'Component'
        com['total'] = float(order['cant']) * float(com['proportion'])
        com['order'] = order['id']
        com['size'] = order['size']
        required.append(com)

    requiredAccessories = accessories[str(ref)]
    for acc in requiredAccessories:
        acc['item'] = ref
        acc['type'] = 'Accessory'
        acc['total'] = float(order['cant']) * float(acc['proportion'])
        acc['order'] = order['id']
        acc['size'] = order['size']
        required.append(acc)

    requiredFabrics = fabrics[str(ref)]
    for fab in requiredFabrics:
        fab['item'] = ref
        fab['type'] = 'Fabric'
        fab['total'] = float(order['cant']) * float(fab['proportion'])
        fab['order'] = order['id']
        fab['size'] = order['size']
        fab['dimention'] = 'ml'
        required.append(fab)
    # print(required)
with open('Report.csv', 'w') as output:
    writer = csv.writer(output, lineterminator = '\n')
    headers = ['name', 'proportion', 'dimention', 'item', 'type', 'total', 'order', 'size']
    writer.writerow(headers)
    for obj in required:
        row = []
        for header in headers:
            row.append(obj[header])
        writer.writerow(row)

    print('SUCCESS!')
