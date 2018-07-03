import xlwings as xw

workbook=xw.Book(r'国开债价格监测表0705.xls')
date_range=workbook.sheets("国开债价格监测表（连续日期）").range(1,1)
print(date_range.column)

workbook.close()