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
BTN__Open_URL__()
__GoogleService_Key__()





################################
#若轉移用下方GUI_Log-in可以不用複製
################################

def first_window(): 
    global uesr_name,password
    import tkinter as tk
    from tkinter import messagebox

    title_name = "first_window"#視窗命名
    win_h = 120 #外框高度
    win_w = 300 #外框寬度
    win_x = 700 #外框X
    win_y = 500 #外框Y

    initial_position_x = 15 #初始位置X
    initial_position_y = 20 #初始位置Y
    Object_w = 8 #統一物件寬度"按鈕"與"標籤"
    spacing_w = 35  #寬度間距
    spacing_h = 20  #高度間距

    #輸入框初始位置
    entry_x = initial_position_x + 110 
    entry_y = initial_position_y + 0 

    #按鈕初始位置
    BTN_x = initial_position_x + 120
    BTN_y = initial_position_y + 50
    window = tk.Tk()
    window.title(title_name)
    window.geometry('%dx%d+%d+%d'%(win_w,win_h,win_x,win_y))
    user_name = tk.StringVar()
    password= tk.StringVar()

###########標籤+輸入框1##########
    new_entry_x = entry_x #初始位置
    new_entry_y = entry_y #初始位置
    both_spacing_x = 100  #兩者間距左右   
    both_spacing_y = 0    #兩者間距上下
    move_x = 0      #X軸定位
    move_y = 0      #Y軸定位
    if (move_x != 0):     new_entry_x = entry_x+(spacing_w*move_x) #間距*次數調整"按鈕X"
    if (move_y != 0):     new_entry_y = entry_y-(spacing_h*move_y) #間距*次數調整"按鈕Y"
    
    lab_0 = tk.Label(window,width=Object_w,text='用户名',compound='center')
    lab_0.place(   x=new_entry_x-both_spacing_x,  y=new_entry_y-both_spacing_y   )
    entry_0 = tk.Entry(window,textvariable=user_name) #用户名輸入框
    entry_0.pack()
    entry_0.place(   x=new_entry_x,  y=new_entry_y )    
    
###########標籤+輸入框2##########
    new_entry_x = entry_x #初始位置x間隔
    new_entry_y = entry_y #初始位置y間隔
    both_spacing_x = 100  #兩者間距左右 
    both_spacing_y = 0    #兩者間距上下 
    move_x = 0      #X軸定位
    move_y = -1     #Y軸定位
    if (move_x != 0):       new_entry_x = entry_x+(spacing_w*move_x) #間距*次數調整"按鈕X"    
    if (move_y != 0):       new_entry_y = entry_y-(spacing_h*move_y) #間距*次數調整"按鈕Y"
    
    lab_1 = tk.Label(window,width=Object_w,text='密碼',compound='center')
    lab_1.place(  x=new_entry_x-both_spacing_x,  y=new_entry_y-both_spacing_y  )
    entry_1 = tk.Entry(window,show="*",textvariable=password) #密碼顯示*號
    entry_1.pack()
    entry_1.place(  x=new_entry_x,  y=new_entry_y  )
            
###########按鈕1##########
    new_BTN_x = BTN_x #初始位置x間隔
    new_BTN_y = BTN_y #初始位置y間隔
    
    move_x = 0  #等於1向右延伸
    move_y = 0     #等於1向上延伸
    if (move_x != 0):       new_BTN_x = BTN_x+(spacing_w*move_x) #間距*次數調整"按鈕X"
    if (move_y != 0):       new_BTN_y = BTN_y-(spacing_h*move_y) #間距*次數調整"按鈕Y"
    
    btn = tk.Button(window,text='登入',fg="black",width=Object_w,compound='center',\
                      bg = "white",command = lambda :__Judge_id_pwd__(window))
    btn.pack()
    btn.place(x=new_BTN_x,y=new_BTN_y)

###########按鈕2##########
    move_x = 2  #等於1向右延伸
    move_y = 0     #等於1向上延伸
    if (move_x != 0):       new_BTN_x = BTN_x+(spacing_w*move_x) #間距*次數調整"按鈕X"  
    if (move_y != 0):       new_BTN_y = BTN_y-(spacing_h*move_y) #間距*次數調整"按鈕Y"
    
    btn = tk.Button(window,text='註冊送出',fg="black",width=Object_w,compound='center',\
                      bg = "white",command = lambda :__register_id_pwd_and_sendGoogle__(window))
    btn.pack()
    btn.place(x=new_BTN_x,y=new_BTN_y)
    
    def __Judge_id_pwd__(window):  #判斷帳密
        if entry_0.get() != 'test' or  entry_1.get() !='123':
            tk.messagebox.showerror('X_X','密碼錯誤,請重新輸入')
        else:
            tk.messagebox.showinfo('O_O','密碼正確')
            window.destroy()          #關閉登入視窗
###################
#開啟功能視窗>>>>>>>可將自行設計的APP功能放進此處
###################
            __New_GUI__()     
    
    def __register_id_pwd_and_sendGoogle__(window):     #註冊帳密
        ID = entry_0.get()
        PWD = entry_1.get()
        print("請通知資訊部帳號啟用 ID:",ID)
        mylist=[[ID,PWD]]
        print(mylist)
        sh.update(Cell_Info,mylist)
        
    
    window.mainloop()
#檢視目前的所有狀態
def BTN__All_txt_Info__(): 
    t.delete("1.0","end")
    BTN__Read_All_txt_Info__()
    send_list ="URL:" + str(URL_Info) +"\n"+ "ID: "+ str(ID_Info)  +"\n"+  "PAGE:" + str(Page_Info) +"\n"+  "Cell:"+str(Cell_Info)  +"\n"+  "Path:" + str(Path_Info) +"\n"
    print(send_list)
    t.insert('insert', send_list +"\n")
#判斷e1下拉選單相同->開啟路徑->寫入e2內容 
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
   
    pass

def BTN__Edit_Path__(): #按鈕編輯路徑使用
    global Path_Info
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    Path_Info = filedialog.askopenfilename()
    print(Path_Info)
    path = txtPath[4] #指定路徑
    f = open(path, 'w',encoding='utf-8')
    line = Path_Info
    f.writelines(line)
    f.close()

######################
#GUI_調整金鑰使用
######################

def __New_GUI__():
    global e1,e2,t
    
    import tkinter as tk  # 使用Tkinter前需要先匯入   
    import tkinter.ttk as ttk 
#建立視窗window
    window = tk.Tk()
    window.title('My Window')
    window.geometry('800x200+500+500')
    
#標籤+下拉選單
    Label1_x=10;    Label1_y=10;    comb_x=170;    comb_y = 10
    tk.Label(window, bg='yellow', width=20, text='選擇項目:').place(x=Label1_x, y=Label1_y, anchor='nw')
    e1 = ttk.Combobox(window,height=5,value=[mylist[0],mylist[1],mylist[2],mylist[3]])
    e1.place(x=comb_x, y=comb_y, anchor='nw')

#標籤+輸入框    
    Label2_x=10;    Label2_y=40;    en_x=170;    en_y = 40
    tk.Label(window, bg='yellow', width=20, text='調整內容:').place(x=Label2_x, y=Label2_y, anchor='nw')
    e2 = tk.Entry(window, show=None, font=('Arial', 14))
    e2.place(x=en_x, y=en_y,  width=200, anchor='nw')

#按鈕123
    BTN1_x=600;BTN1_y=10;    BTN4_x=600;BTN4_y=40;
    BTN2_x=490;BTN2_y=10;    BTN3_x=400;BTN3_y=10
    
    b1=tk.Button(window, text='Open URL', width=20,height=1, command= BTN__Open_URL__ )
    b1.place(x=BTN1_x, y=BTN1_y, anchor='nw') #引用 BTN__Open_URL__()
    b2=tk.Button(window, text='檢視確認', width=10,height=3, command= BTN__All_txt_Info__ )
    b2.place(x=BTN2_x, y=BTN2_y, anchor='nw')  #引用 _DataProcessing_()
    b3=tk.Button(window, text='暫存調整', width=10,height=3, command= BTN__Matching_options_modification__ )
    b3.place(x=BTN3_x, y=BTN3_y, anchor='nw')   #引用 BTN__Matching_options_modification__()
    b4 = tk.Button(window, text='Edit Path.txt', width=20,height=1, command= BTN__Edit_Path__ ) #引用 BTN__Open_URL__()
    b4.place(x=BTN4_x, y=BTN4_y, anchor='nw')

#顯示框
    t = tk.Text(window,bg='black',fg='white', width=800,height=8)
    t.place(x=0, y=80, anchor='nw')        
    
    window.mainloop()


first_window()
