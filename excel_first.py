import xlwings as xw

workbook=xw.Book(r'国开债价格监测表0705.xls')
thisSheet=workbook.sheets("国开债价格监测表（连续日期）")
# date_range=workbook.sheets("国开债价格监测表（连续日期）").range('a3:D16').value
# print(date_range)

#获取整个表格数据

#最后一行行标
row=thisSheet['c3'].end('down').row
#最后一列列标
column=thisSheet['c3'].end('right').column
print(row,column)


workbook.close()