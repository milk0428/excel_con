import xlwings as xw

workbook=xw.Book(r'国开债价格监测表0705.xls')
date_range=workbook.sheets("国开债价格监测表（连续日期）").range(12,1).expand()
print(date_range)

workbook.close()