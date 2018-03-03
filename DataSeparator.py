from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from style import style_range

# 參數設定區
strFileName = '八月銷售表'  # input('請輸入檔案名稱：')
strDistinctCol = '牙醫診所'  # input('請輸入要建立分頁所依據的欄位名稱：')
strHeadText = '2017 年 8月 請 款 單'  # input('請輸入表頭文字：')
strOutPutFileName = strFileName + '_output'

# 讀入檔案
wb = load_workbook(filename=f'{strFileName}.xlsx')
sheet_ranges = wb['總表']

# 找到要distinct的欄位是在第幾欄
listCols = list(sheet_ranges['1']) # 載入欄位列
listCols = list(filter(lambda x: x.value is not None, listCols))  # 去掉值為None的cell
idxLastCol = listCols[len(listCols)-1].column  # 找到最後一個欄位的英文index
cellDistinctCol = list(filter(lambda x: x.value == strDistinctCol, listCols))[0]  # 找到要distinct的欄位是在那一格

# 讀入要distinct的col
listDistinctCol = sheet_ranges[cellDistinctCol.column]
listDistinctCol = list(filter(lambda x: x.value is not None, listDistinctCol))  # 去掉值為None的cell
listDistinctCol = list(map(lambda x: x.value, listDistinctCol))  # 把每個cell的值都取出來形成list
numRows = len(listDistinctCol)  # 設定有幾列
listDistinctCol.remove(strDistinctCol)  # 移除欄位cell
listDistinctCol = list(set(listDistinctCol))  # distinct欄位值

# 在記憶體中建立新的excel檔案，並且依照listDistictCol不同的cell值建立sheets
wbOutPut = Workbook()
listSheets = list(map(lambda x: wbOutPut.create_sheet(x, 0), listDistinctCol))  # 在記憶體中要輸出的新excel檔裡建立sheets

# 寫入表頭
font = Font(b=True, color="000000", size=28)  # 粗體, 28號字, 黑色
al = Alignment(horizontal="center", vertical="center")  # 置中排列
thin = Side(border_style="thin", color="000000")  # 沒有框
border = Border(top=thin, left=thin, right=thin, bottom=thin)

for sheet in listSheets:
    cellHead = sheet['A1']
    cellHead.value = strHeadText
    style_range(sheet, f'A1:{idxLastCol}3', border=border, fill='', font=font, alignment=al)

# 讀取每一列，根據strDistinctCol的值把列塞到不同的sheet
dictSheetIndex = {}  # 建立sheet字典
idx = 0
for item in listDistinctCol:  # 給字典塞入值，這個數字要用在sheet的索引
    dictSheetIndex[item] = idx
    idx += 1

for numRowIdx in range(2, numRows + 1):
    row = list(sheet_ranges[f'A{numRowIdx}:{idxLastCol}{numRowIdx}'][0])  # 把列讀進來
    idxSheet = dictSheetIndex[sheet_ranges[f'{cellDistinctCol.column}{numRowIdx}'].value]  # 根據要distinct的目標欄位的值，找出要寫入的sheet index
    row = list(map(lambda x: x.value, row))  # 把列的所有cell的摭取出來重組成list
    listSheets[idxSheet].append(row)

# 寫入檔案
wbOutPut.save(filename = f'{strOutPutFileName}.xlsx')

