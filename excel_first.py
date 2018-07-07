import xlwings as xw
import datetime
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
#
# # datetime.timedelta.__divmod__(36000)
#
# #计算两个日期之间相差天数
# days=thisSheet["C5"].value.date()-thisSheet["C4"].value.date()
# print(int(days.total_seconds()/86400))
# print(thisSheet.range((1,1),(row,column)).value)

new_worksheet=workbook.sheets.add()
new_worksheet.range(1,1).value=thisSheet.range((1,1),(row,column)).value

# workbook.close()