import xml.etree.ElementTree as ET # 匯入XML函式庫 並命名為ET
import requests # 爬蟲指令函式庫
import time # 時間函式庫
from openpyxl import Workbook

def xml_to_dict(element):
    result = {}
    for child in element:
        if len(child) == 0:
            result[child.tag] = child.text
        else:
            result[child.tag] = xml_to_dict(child)
    return result

def assembleDate(year, month, day):
    if len(str(month)) == 1:
        month = "0" + str(month)
    if len(str(day)) == 1:
        day = "0"+ str(day)
    return f"{year}{month}{day}"

def returnStrDayList(startYear, startMonth, endYear, endMonth, day="01"):
    DayList = list() # 儲存所有年月日日期
    year = startYear
    month = startMonth
    while True:
        if year == endYear and month == endMonth:
            DayList.append(assembleDate(year, month, day))
            return DayList
        DayList.append(assembleDate(year, month, day))
        month += 1
        if month > 12:
            month = 1
            year += 1

def fillSheet(sheet, data, row): #將一行資料填入excel檔案中
    # sheet: 表單名稱
    # data: 資料
    # row: 第幾行
    for column, value in enumerate(data, 1):
        # 將data填在第row行
        sheet.cell(row = row, column = column, value = value)

tree = ET.parse("setting.xml") # 讀取設定檔
root = tree.getroot() # 獲取裡面的資料
data = xml_to_dict(root)
fields = ["日期","成交股數","成交金額","開盤價","最高價","最低價","收盤價","漲跌價差","成交筆數"]
startYear, startMonth = int(data["startYear"]), int(data["startMonth"])
endYear, endMonth = int(data["endYear"]), int(data["endMonth"])
monthList = returnStrDayList(startYear,startMonth,endYear,endMonth)

wb = Workbook() # 打開excel檔案
sheet = wb.active # 啟動excel並建立sheet
sheet.title = f"每日成交統計表"
row = 1
fillSheet(sheet, fields, row)
row += 1

#以下透過monthList抓取每月份的資料
for month in monthList:
    rq = requests.get(data["url"],params={
        "date": month,
        "stockNo": data["stockNo"],
        "response":"json"
    })
    
    print(rq) #確認http狀態碼是否正常
    if str(rq) != "<Response [200]>": #被伺服器擋掉
        break
    jsonData = rq.json()
    dailyPricesList = jsonData.get("data",[]) #從json檔案中抓取data陣列
    for dailyPrices in dailyPricesList:
        fillSheet(sheet, dailyPrices, row)
        row += 1
    time.sleep(3)

name = data["excelName"]
wb.save(name+".xlsx")