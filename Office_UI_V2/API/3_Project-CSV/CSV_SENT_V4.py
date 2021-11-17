'''
開發日期:20210818
開發人員:Ness_Huang

主題: 帳號密碼登入+網路修改參數

初始設定
URL:https://docs.google.com/spreadsheets/d/1YvBv43Adb5CDEMngiFn2LPxZXdN9BR0unPYqUVNNaZE/edit#gid=0
ID:1YvBv43Adb5CDEMngiFn2LPxZXdN9BR0unPYqUVNNaZE
PAGE:123
CELL:A15
Path:./參數調整/demo.txt


功能1登入介面ID/PWD:test/123
功能2修改介面可修改sheetAPI服務常用參數URL/ID/PAGE/CELL/Path

'''

#讀取文字檔路徑
_sheet_Key_="./參數調整/creds.json"#金鑰路徑
txtPath=["./參數調整/1_URL.txt","./參數調整/2_ID.txt","./參數調整/3_PAGE.txt","./參數調整/4_CELL.txt","./參數調整/5_Path.txt"]
#下拉選擇+寫入對應.txt
mylist=["1.URL","2.ID","3.Page","4.Cell","5.Path"]  #下拉式清單使用

sh="" #金鑰打包API
URL_Info =[] #網址
ID_Info = [] #GoogleID
Page_Info = []  #分頁
Cell_Info=[] #更新的位置
Path_Info=[] #更新的路徑

#讀取文字當內容
def BTN__Read_All_txt_Info__():
    global URL_Info,ID_Info,Page_Info,Cell_Info,Path_Info
    
    filename = open(txtPath[0],'r',encoding='utf-8')        
    URL_Info = str(filename.read())  
    filename = open(txtPath[1],'r',encoding='utf-8')        
    ID_Info = str(filename.read())    
    filename = open(txtPath[2],'r',encoding='utf-8')        
    Page_Info = str(filename.read())    
    filename = open(txtPath[3],'r',encoding='utf-8')        
    Cell_Info = str(filename.read())
    filename = open(txtPath[4],'r',encoding='utf-8')        
    Path_Info = str(filename.read())
    
#自訂開啟瀏覽器
def BTN__Open_URL__():
    import webbrowser
    webbrowser.open(URL_Info)

#上傳授權與金鑰需上Google cloud platfrom啟用API與服務
def __GoogleService_Key__():
    global sh
    import gspread
    from gspread.models import Cell
    from oauth2client.service_account import ServiceAccountCredentials 
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(_sheet_Key_, scope) #權限金鑰
    client = gspread.authorize(creds)           #使用金鑰
    sh = client.open_by_key(ID_Info).worksheet(Page_Info) #指定頁面 ID + Page_Info   

BTN__Read_All_txt_Info__()
#BTN__Open_URL__()
__GoogleService_Key__()


'''
更新日期:2021-08-19
開發人員:黃亦豪Ness_Huang

【使用前請確認以下】
1.GUI_SheetAPI_V3 第一次使用請編輯 URL / ID / PAGE / CELL /PATH
2.

'''

my_y=100

def __New_GUI__():
    global e1,e2,t
    
    import tkinter as tk  # 使用Tkinter前需要先匯入   
    import tkinter.ttk as ttk 
#建立視窗window
    window = tk.Tk()
    window.title('My Window')
    window.geometry('600x500')   
#直接編輯指定txt
    tk.Label(window, width=20, font=('Arial', 14), text='直接編輯指定.txt').place(x=20, y=0+my_y, anchor='nw')
    tk.Label(window, bg='yellow', width=20, text='CSV路徑:').place(x=10, y=40+my_y, anchor='nw')
    b0 = tk.Button(window, text='Edit path', width=15,height=1, command= BTN__Edit_Path__ ) #引用 BTN__Open_URL__()
    b0.place(x=170, y=40+my_y, anchor='nw')    
    tk.Label(window, bg='yellow', width=20, text='選擇項目:').place(x=10, y=70+my_y, anchor='nw')
    e1 = ttk.Combobox(window,value=[mylist[0],mylist[1],mylist[2],mylist[3],mylist[4]])
    e1.place(x=170, y=70+my_y, anchor='nw')
    tk.Label(window, bg='yellow', width=20, text='調整內容:').place(x=10, y=100+my_y, anchor='nw')
    e2 = tk.Entry(window, show=None, font=('Arial', 14)) #字體與大小
    e2.place(x=170, y=100+my_y,  width=400, anchor='nw')
    b3 = tk.Button(window, text='檢視+確認', width=10,height=3, command= BTN__All_txt_Info__ ) #引用 _DataProcessing_() 
    b3.place(x=400, y=40+my_y, anchor='nw') 
    b4 = tk.Button(window, text='選項+修改', width=10,height=3, command= BTN__Matching_options_modification__ ) #引用 BTN__Open_URL__()
    b4.place(x=490, y=40+my_y, anchor='nw')     

#顯示框
    t = tk.Text(window,bg='black',fg='white', width=800,height=12)
    t.place(x=0, y=135+my_y, anchor='nw')

#按鈕 - 開啟網頁/檔案/發送
    b1 = tk.Button(window, text='Open Google Sheet', width=25,height=2, command= BTN__Open_URL__ ) #引用 BTN__Open_URL__()
    b1.place(x=10, y=120-my_y, anchor='nw')
    b2 = tk.Button(window, text='Open File', width=25,height=2, command= BTN__Open_Path__ ) #引用 BTN__Open_Path__()
    b2.place(x=210, y=120-my_y, anchor='nw')
    b5 = tk.Button(window, text='發送測試', width=25,height=2, command= BNT__Update_Array__ ) #引用 BTN__Open_URL__()
    b5.place(x=410, y=120-my_y, anchor='nw')
    window.mainloop()

#按鈕編輯路徑使用    
def BTN__Edit_Path__(): 
    global Path_Info
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    Path_Info = filedialog.askopenfilename()
    print(Path_Info)
    path = txtPath[4]
    f = open(path, 'w',encoding='utf-8')
    line = Path_Info
    f.writelines(line)
    f.close()

#自訂開啟NAS路徑    
def BTN__Open_Path__():
    import os
    dirPath = Path_Info    
    path = dirPath
    os.startfile(path)
    
#檢視目前的所有狀態
def BTN__All_txt_Info__(): 
    t.delete("1.0","end")
    BTN__Read_All_txt_Info__()
    send_list ="URL:" + str(URL_Info) +"\n"+ "ID: "+ str(ID_Info)  +"\n"+  "PAGE:" + str(Page_Info) +"\n"+  "Cell:"+str(Cell_Info)  +"\n"+  "Path:" + str(Path_Info) +"\n"
    print(send_list)
    t.insert('insert', send_list +"\n")

#比對+寫入e2內容
def BTN__Matching_options_modification__(): 
    if (e1.get()==mylist[0]):
        path = txtPath[0]
        f = open(path, 'w',encoding='utf-8')
        line = e2.get()
        f.writelines(line)
        f.close()
    if (e1.get()==mylist[1]):
        path = txtPath[1]
        f = open(path, 'w',encoding='utf-8')
        line = e2.get()
        f.writelines(line)
        f.close()
    if (e1.get()==mylist[2]):
        path = txtPath[2]
        f = open(path, 'w',encoding='utf-8')
        line = e2.get()
        f.writelines(line)
        f.close()
    if (e1.get()==mylist[3]):
        path = txtPath[3]
        f = open(path, 'w',encoding='utf-8')
        line = e2.get()
        f.writelines(line)
        f.close()
    if (e1.get()==mylist[4]):
        path = txtPath[4]
        f = open(path, 'w',encoding='utf-8')
        line = e2.get()
        f.writelines(line)
        f.close()
    pass

#發送GOOGLE 指定Cell 更新
def BNT__Update_Array__(): 
    __csv_Array__(Path_Info)
    sh.update(Cell_Info, csvList)
    print(csvList)

def __csv_Array__(Path_Info):  #CSV清單製作
    global csvList
    import csv
    with open(Path_Info, newline='',encoding='utf-8') as csvfile:    
        rows = csv.reader(csvfile,delimiter=",")# 讀取 CSV 檔案內容，分隔符號是 Tab "\t"
        csvList = []# 設定一個空陣列      
        for row in rows:# # 以迴圈輸出每一列資料加到 csvlist 陣列裡
            csvList.append(row)
        return csvList

#########################
#~~~~~~~程式開始~~~~~~
#########################
#__New_GUI__()
BNT__Update_Array__()  #發送功能