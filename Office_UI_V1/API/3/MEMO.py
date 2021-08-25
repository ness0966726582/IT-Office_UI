'''
V3
1.未來可以變更gsheet name
舊->sheet = client.open("ES Routine Work").worksheet('Endorsements')#指定googlesheet
改->sheet = client.open_by_key("1z10OoKJqZUBp4s97PVBXCxb0II5wWi8DGcKzpt__Swc").worksheet('Endorsements')#指定googlesheet

2.取得電腦AD
新增->import getpass  #導入通行證 User name

程式說明:
"開啟"-指定MEMO 的 NAS儲存路徑 
GUI 按鈕1 產生陣列並上傳google sheet--------> 範例 : [2020-12-10 , MEMO2 , 貓王害我加班 , ness_huang]
手動填寫 MEMO 與 DESCRIPTION
自動抓取 DATE 與 USER
GUI 按鈕2 寄送通知 Mail or Line------------>還沒做

新智圖URL:
https://coggle.it/diagram/X9CAiVJoLG3m_bsZ/t/it-oms-%E8%BD%89%E7%A7%BB-googlesheet
'''
import tkinter as tk  # 使用Tkinter前需要先匯入   
import tkinter.ttk as ttk 
import datetime#取得time

#dirPath = r"\\10.231.250.70\Department\InformationTechnology\Auditing\02. MEMO AND SPEED LETTER (AUTO)"   #照片存放路徑 
dirPath = r"\\10.231.199.10\Department\InformationTechnology\Auditing\02. MEMO AND SPEED LETTER (AUTO)"   #照片存放路徑 

send_list = []
Date = ""
User = ""
Memo= ""
Description = ""
t = "" #GUI 顯示的換行暫存

###################################
#         開啟圖片存放路徑         #
###################################
def open_path():
    import os
    path = dirPath
    os.startfile(path)

###################################
#         開啟URL                  #
###################################
def URL_DB():
    import webbrowser
    webbrowser.open('https://docs.google.com/spreadsheets/d/1z10OoKJqZUBp4s97PVBXCxb0II5wWi8DGcKzpt__Swc/edit#gid=825972601')  # Go to example.com

def URL_datastudio():
    import webbrowser
    webbrowser.open('https://datastudio.google.com/u/1/reporting/829c365e-916d-46ba-9b55-6778ed119d84/page/Hp9sB')  # Go to example.com

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
    #sheet = client.open("ES-ITR/FAF/Borrow/IR").worksheet('Endorsements')#指定googlesheet
    #sheet = client.open("ES Routine Work").worksheet('Endorsements')#指定googlesheet
    sheet = client.open_by_key("1z10OoKJqZUBp4s97PVBXCxb0II5wWi8DGcKzpt__Swc").worksheet('Endorsements')#指定googlesheet
    #sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1z10OoKJqZUBp4s97PVBXCxb0II5wWi8DGcKzpt__Swc/edit#gid=1838131587")----->不能使用
    ###################################
    #             寫入方式             #
    ###################################
    send_list =[str(Date),str(Memo),str(Description),str(User),str(1)]  #1是用於疊加比數使用
    index=36   #行暫存初始(googlesheet插入的列)
    #row=[1,2,1]
    #sheet.insert_row(row, index)
    sheet.insert_row(send_list, index)
    
#取得路徑內檔名,使用參數: 1.dirPath 2.old=[]

def openTK():
    global e1,e2,t
    
    window = tk.Tk()# 第1步，例項化object，建立視窗window
    window.title('My Window')# 第2步，給視窗的視覺化起名字
    window.geometry('500x300')  # 第3步，設定視窗的大小(長 x 寬)
    #----建立多行文字框--顯示輸出內容--
    t = tk.Text(window,bg='grey', height=4)
    t.place(x=0, y=0, anchor='nw')
    #----建立MEMO標籤+輸入框
    tk.Label(window, bg='yellow', width=20, text='類型 Type:').place(x=50, y=150, anchor='nw')
    #e1 = tk.Entry(window, show=None, font=('Arial', 14))
    e1 = ttk.Combobox(window,value=["Memo","Speed Letter","Other"])
    e1.place(x=200, y=150, anchor='nw')
    
    
    #----建立Description標籤+輸入框
    tk.Label(window, bg='yellow', width=20, text='主旨 Subject:').place(x=50, y=180, anchor='nw')
    e2 = tk.Entry(window, show=None, font=('Arial', 14))
    e2.place(x=200, y=180, anchor='nw')
    
    
    # 發送按鈕
    b1 = tk.Button(window, text='Send - Data 發送資料', width=30,height=1, command=__SendData__)
    b1.place(x=200, y=220, anchor='nw')    
    
    # 網頁檢視清單按鈕
    b2 = tk.Button(window, text='Check - Website'+"\n"+'網頁檢視清單', width=25,height=4, command=openWebsite)
    b2.place(x=250, y=50, anchor='nw')
    
    # 開啟簽GOOGLE到表按鈕
    b3 = tk.Button(window, text='Open - Attendance list'+"\n"+'開啟簽到表', width=25,height=4, command=URL_DB)
    b3.place(x=50, y=50, anchor='nw')
    window.mainloop()

#----定義按鈕函式------------------------------------------------------->可添加新的變數
def __SendData__(): # 在滑鼠焦點處插入輸入內容
    global Date,User,Memo,Description,t
    import getpass  #導入通行證 User name
    
    Date = datetime.datetime.now().strftime("%Y-%m-%d") 
    Memo = e1.get()
    Description = e2.get()
    
    User=getpass.getuser() #import getpass
    #User="IT"    
    send_list = str(Date) +','+ str(Memo)+','+str(Description)+','+str(User)+','+str(1)
    URL_DB()
    print(send_list)
    t.insert('insert', send_list+"\n")
    Sendsheet()

def openWebsite():
    URL_datastudio()

open_path()
print("==================")
print("Put into MEMO!!")
print("==================")
print("---->Enter<---- Fill in info")
openTK()

    

