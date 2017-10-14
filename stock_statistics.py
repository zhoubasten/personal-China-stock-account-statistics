# coding=utf-8
from openpyxl import load_workbook
import tushare as ts
import chardet
import sys

def update_sheet(sheetname): 
  sheet = wb.get_sheet_by_name(sheetname)
  print("处理sheet: ", sheetname)
  for i in range(2,13,1):
    
    code = str(sheet["V"+str(i)].value) #股票代码
    if code == "None":
      print("%s 查询结束" % sheetname)
      break
    name = sheet["W"+str(i)].value #股票名称
    shares = float(sheet["X"+str(i)].value)

    print("查询到的股票代码", code, name)

    df = ts.get_realtime_quotes(code)
    name = df["name"][0]
    code = df["code"][0]
    price = float(df["price"][0])

    if name == df["name"][0] :
      
      amount = price*shares
      print("经过校验，名字相同, 股价%.2f, 股数%.2f, 总价%.2f" % (price, shares, amount))
      sheet["Y"+str(i)].value = amount
      sheet["Z"+str(i)].value = price
      

    else:
      print("经过校验，股票名字不相同")

if __name__ == "__main__":
  try:
    reload(sys)
    sys.setdefaultencoding('utf-8')
  except:
    pass
  print (sys.getdefaultencoding())
  wb = load_workbook("/Users/roger/OneDrive/股票/2017.xlsx")

  sheets = ("pingan", "huaxi", "huatai")
  for sheet in sheets: 
    update_sheet(sheet)
  wb.save("/Users/roger/OneDrive/股票/2017-1.xlsx")
  print("更新已经完成")


