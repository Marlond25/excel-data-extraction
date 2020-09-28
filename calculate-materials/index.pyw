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
oversizes = {}
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

	rows = len(isSizeExtension.index)
	names = isSizeExtension[isSizeExtension.columns[2]].tolist()
	sis = isSizeExtension[isSizeExtension.columns[3]].tolist()
	ovs = isSizeExtension[isSizeExtension.columns[4]].tolist()
	os = []
	if (rows > 0):
		for i in range(rows):
			o = {}
			o[names[i]] = { sis[i]: ovs[i] }
			os.append(o)

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
	componentsAndAccessories['oversizes'] = os
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
			oversizes[item] = componentsAndAccessories['oversizes']
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

	def getOversize(comps, comp, size):
		for c in comps:
			if (comp in c.keys()):
				if (str(size) in c[comp].keys()):
					return c[comp][str(size)]
		return 'NN'

	for com in requiredComponents:
		add = {}
		add['name'] = com['name']
		add['proportion'] = com['proportion']
		add['dimention'] = com['dimention']
		add['color'] = order['color']
		add['desc'] = order['desc']
		add['cant'] = order['cant']
		add['item'] = ref
		add['type'] = 'Component'
		add['total'] = float(order['cant']) * float(com['proportion'])
		add['order'] = order['id']
		add['size'] = order['size']
		add['oversize'] = getOversize(oversizes[str(ref)], com['name'], order['size'])
		required.append(add)

	requiredAccessories = accessories[str(ref)]
	for acc in requiredAccessories:
		add = {}
		add['name'] = acc['name']
		add['proportion'] = acc['proportion']
		add['dimention'] = acc['dimention']
		add['color'] = order['color']
		add['desc'] = order['desc']
		add['cant'] = order['cant']
		add['item'] = ref
		add['type'] = 'Accessory'
		add['total'] = float(order['cant']) * float(acc['proportion'])
		add['order'] = order['id']
		add['size'] = order['size']
		add['oversize'] = 'NN'
		required.append(add)

	requiredFabrics = fabrics[str(ref)]
	for fab in requiredFabrics:
		add = {}
		add['name'] = fab['name']
		add['proportion'] = fab['proportion']
		add['color'] = order['color']
		add['desc'] = order['desc']
		add['cant'] = order['cant']
		add['item'] = ref
		add['type'] = 'Fabric'
		add['total'] = float(order['cant']) * float(fab['proportion'])
		add['order'] = order['id']
		add['size'] = order['size']
		add['oversize'] = 'NN'
		add['dimention'] = 'ml'
		required.append(add)

with open('Report.csv', 'w') as output:
    writer = csv.writer(output, lineterminator = '\n')
    headers = ['name', 'type', 'dimention', 'item', 'desc', 'color', 'proportion', 'cant', 'total', 'order', 'size', 'oversize']
    writer.writerow(headers)
    for obj in required:
        row = []
        for header in headers:
            row.append(obj[header])
        writer.writerow(row)
