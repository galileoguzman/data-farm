import xlrd

book = xlrd.open_workbook("farm_gdl.xls")

print "The number of worksheets is", book.nsheets
print "Worksheet name(s):", book.sheet_names()



## Imprimir el objeto de la hoja de excel seleccionada
sheet = book.sheet_by_index(0)
print sheet

## Imprimir el numero de columnas
number_col = sheet.ncols
print number_col

for row in range(sheet.nrows):
	if sheet.cell_value(rowx=row, colx=0) == '':
		print 'CADENA VACIA POR GALILEO'
	else:
		print sheet.cell_value(rowx=row, colx=0)
