import xlrd, xlwt, xlutils


def name_city_drug_store(name):
	city = "{} {}".format(name.partition(' ')[0] , name.partition(' ')[1])
	return city

def del_number_giant(name):
	commerce  = name.partition(' ')
	if "GIGANTE" in name:
		return 'GIGANTE'
	elif "FARM" in name:
		return "FARM GUADALAJARA"
	elif "MART" in name:
		return "WAL MART"
	elif "CAYC IMPORTACIONES" in name:
		return "CAYC IMPORTACIONES"
	elif "COBRANZA" in name:
		return "TELEFONICA COBRANZA"
	else:
		return name


def get_type_and_zone_bussiness(commerce):
	if "FARMACIA" in commerce:
		return 1
	elif "GIGANTE" in commerce:
		return 2
	elif "WAL" in commerce:
		return 3
	else:
		return 4

file_xls = "farm_gdl.xls"

book = xlrd.open_workbook(file_xls, formatting_info=True)

print "The number of worksheets is", book.nsheets
print "Worksheet name(s):", book.sheet_names()



## Imprimir el objeto de la hoja de excel seleccionada
sheet = book.sheet_by_index(0)
print sheet

## Imprimir el numero de columnas
number_col = sheet.ncols
print number_col

## New Book
book_new = xlwt.Workbook(encoding="utf-8")
sheet_new = book_new.add_sheet("Hoja 1")

sheet_new.write(0, 0, "POBLACION DEL COMERCIO")
sheet_new.write(0, 1, "NOMBRE DEL COMERCIO")
sheet_new.write(0, 2, "Zona")
sheet_new.write(0, 3, "Tipo")

for row in range(sheet.nrows):
	if sheet.cell_value(rowx=row, colx=0) == '':
		sheet_new.write(row, 0, "VACIO")
		sheet_new.write(row, 1, "VACIO")
		sheet_new.write(row, 2, "VACIO")
		sheet_new.write(row, 3, "VACIO")
	else:
		city = name_city_drug_store(sheet.cell_value(rowx=row, colx=0))
		commerce = del_number_giant(sheet.cell_value(rowx=row, colx=1))
		type_commerce = get_type_and_zone_bussiness(commerce)
		zone_commerce = get_type_and_zone_bussiness(commerce)
		print "Ciudad {c} Comercio {co} Tipo {t} Zona {z}".format(c=city, co=commerce,t=type_commerce,z=zone_commerce)

		sheet_new.write(row, 0, city)
		sheet_new.write(row, 1, commerce)
		sheet_new.write(row, 2, type_commerce)
		sheet_new.write(row, 3, zone_commerce)


book_new.save(file_xls)



