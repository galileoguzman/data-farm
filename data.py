import xlrd

book = xlrd.open_workbook("farm_gdl.xls")

print "The number of worksheets is", book.nsheets
print "Worksheet name(s):", book.sheet_names()