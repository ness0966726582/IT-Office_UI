#--------固定不變--------#
dirPath = r"\\10.231.250.70\Department\InformationTechnology\Auditing\11. Borrow\NEW"   #照片存放路徑 
Use_bit=20                    #取檔案名位數目的為移除副檔名
Output1="old.txt"            #轉存的文件名
Output2="new.txt"
#--------Borrow檔暫存放區--------#
old=[]                       #讀取舊資料
Borrow_only=[]                  #保留檔名的內容
zbar=0                       #-------------------->About全域變數of完成進度7/17
###################################
#         開啟圖片存放路徑          #
###################################
def open_path():
    import os
    path = dirPath
    os.startfile(path)
    
###################################
#         開啟URL                  #
###################################
def open_URL():
    import webbrowser
    webbrowser.open('https://docs.google.com/spreadsheets/d/1z10OoKJqZUBp4s97PVBXCxb0II5wWi8DGcKzpt__Swc/edit#gid=842949924')  # Go to example.com

def Sendsheet():
    ###################################
    #            程式宣告區             #
    ###################################
    global zbar                       #-------------------->About全域變數of完成進度7/17
    
    import gspread
    from gspread.models import Cell
    from oauth2client.service_account import ServiceAccountCredentials 
    import string as string
    import random
    from pprint import pprint

    ###################################
    #           獲取授權與連結           #
    ###################################
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope) #權限金鑰
    client = gspread.authorize(creds)           #使用金鑰
    #sheet = client.open("2020-ES-ITR/FAF/Borrow/IR").sheet1   #指定googlesheet
    #sheet = client.open("ES-ITR/FAF/Borrow/IR").worksheet('Files Check')#指定googlesheet
    sheet = client.open_by_key("1z10OoKJqZUBp4s97PVBXCxb0II5wWi8DGcKzpt__Swc").worksheet('Files Check')
    ###################################
    #             寫入方式             #
    ###################################
    data=Borrow_only
    x=3   #列暫存初始C(googlesheet)
    y=2   #行暫存初始2(googlesheet)
    z=0   #陣列的初始位置

#    sheet.clear()
    '''------------------------------------------------------------------------------------------>一行一行上傳
    while z < len(data):
        sheet.update_cell(y,x, data[z]) #功能2 更新CELL資料特定位置2行1列
        y = y + 1   #向下加入Sheet
        z = z + 1   #更新計數器
    '''
    
    # Select a range------------------------------------------------------------------------------------------>整筆上傳2020/7/17
    cell_list = sheet.range('C2:C'+str(len(data)+1))   #A2:An  n= 取得字串長度+1
    
    for cell in cell_list:
        
        cell.value = data[z]
        zbar = z + 1            #全域變數
        z = z + 1
    sheet.update_cells(cell_list)
    
#取得路徑內檔名,使用參數: 1.dirPath 2.old=[]
def getName():
    import os
    global old
    old = os.listdir(dirPath)
    #print("路徑內的檔名:\n",os.listdir(dirPath))
    #print("查詢路徑內檔案名"+str(old))#檢查讀取資料

#---------轉出FAF號碼另存.txt--------#    
def outputOld():
    i=0
    with open(Output1, 'w') as outfile:
        while i < len(old):
            outfile.write(old[i]+'\n')
            i = i + 1   #更新計數器
    #print("匯出-舊資料:"+str(old))

#---------移除副檔名暫存---------#
def Borrow_number():
    global Borrow_only
    with open(Output1, 'r') as infile:
        while True:
            data = infile.readline()     # 一次讀一行資料
            if not data:                 # 所有資料讀取完畢
                break
            #print(data[:6]+"\n", end='')          # end='': 不要自動加斷行
            Borrow_only.append(data[:Use_bit]) #使用自訂參數
        #print("暫存-移除副檔名:"+str(Borrow_only))

#---------轉出Borrow號碼另存output.txt--------#    
def outputNew():
    j=0
    with open(Output2, 'w') as outfile:
        while j < len(Borrow_only):
            outfile.write(str(Borrow_only[j])+',')
            j = j + 1
    #print("匯出-新資料:"+str(Borrow_only))    

def bar100():                               #------------------->About全域變數of完成進度7/17
    import time
    import progressbar

    for i in progressbar.progressbar(range(zbar)):
        time.sleep(0.1)
        
open_path()
print("==================")
print("Put into image!!")
print("==================")
print("---->Enter<---- send to google-sheet")
input()
getName()
outputOld()
Borrow_number()
outputNew()
print("Uploading...")
open_URL()
Sendsheet()
bar100()
print("Upload complete")
input()