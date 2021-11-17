'''
程式開發日期:20210510
開發者:Ness_huang

主程式功能讀取ES.txt取得
定時擷取文檔內的OID定時陣列上傳
1.啟用請修改程式內ES或AP
2.下方為上傳的URL
ES網域 : https://docs.google.com/spreadsheets/d/1FApfprfwR77wCJ03ZegJkKIeCJfojUn-YmyuShOnFDU/edit#gid=0
AP網域 : https://docs.google.com/spreadsheets/d/1FApfprfwR77wCJ03ZegJkKIeCJfojUn-YmyuShOnFDU/edit#gid=1481041442
creds.json為金鑰
ES.txt SNMP必要條件請確認自身Community + IP + Port + OID + location + sensor_name (直接在文字檔內新增利用Tab鍵分隔)

修改日期:20210818
-新增金鑰的抓取TXT 19行~74行

'''

#讀取文字檔路徑
_sheet_Key_="./參數調整/creds.json"#金鑰路徑
txtPath=["./參數調整/1_URL.txt","./參數調整/2_ID.txt","./參數調整/3_PAGE.txt","./參數調整/4_CELL.txt","./參數調整/5_Path.txt"]
#下拉選擇+寫入對應.txt
mylist=["1.URL","2.ID","3.Page","4.Cell"]  #下拉式清單使用

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
    
#檢視目前的所有狀態
def BTN__All_txt_Info__(): 
    t.delete("1.0","end")
    BTN__Read_All_txt_Info__()
    send_list ="URL:" + str(URL_Info) +"\n"+ "ID: "+ str(ID_Info)  +"\n"+  "PAGE:" + str(Page_Info) +"\n"+  "Cell:"+str(Cell_Info)  +"\n"+  "Path:" + str(Path_Info) +"\n"
    print(send_list)
    t.insert('insert', send_list +"\n")
    
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

#跑前置作業
BTN__Read_All_txt_Info__()
BTN__Open_URL__()
__GoogleService_Key__()

#################### SNMP處理區域 ####################
from time import *
import time;  # 引入time模块
import os
from apscheduler.schedulers.blocking import BlockingScheduler
from pysnmp.hlapi import *
import sys


#file_name = "./主-參數調整/SNMP.txt"
file_name = Path_Info  #取用Path.txt的路徑

count = 0
result = []

time_now = ""
location = ""
area = ""
IP = ""
oid = ""
oidValue =""
sensor_name = ""
varCommunity = ""
varPort = ""

all_info = ""
new_arr=[]

def reset():
    global count , result , varCommunity ,  varPort 
    global time_now , location , area , IP , oid , oidValue ,sensor_name
    global all_info , new_arr
#-----全歸零-----#
    
    count = 0
    result = []
    
    time_now = ""
    location = ""
    area = ""
    IP = ""
    oid = ""
    oidValue =""
    sensor_name = ""
    varCommunity = ""
    varPort = ""
    
    all_info = ""
    new_arr=[]
    
def now():
    global time_now
    time_now = time.strftime("%Y-%m-%d %H:%M", time.localtime())

def get_TXT():
    global count,result,varCommunity,IP,varPort,oid,location,sensor_name,area

#-----取得每行文字檔資料-----#
    with open(file_name,'r',encoding='utf-8') as f:
        for line in f:
            result.append(list(line.strip('\n').split(',')))

    ###################
    #data.txt內使用行數#
    ###################
    file_dirs = file_name
    filename = open(file_dirs,'r',encoding='utf-8')        #以只读方式打开文件
    file_contents = filename.read()       #读取文档内容到file_contents
    for file_content in file_contents:    #统计文件内容中换行符的数目
        if file_content == '\n':
            count += 1
    if file_contents[-1] != '\n':         #当文件最后一个字符不为换行符时，行数+1
        count += 1
    #print('文件%s總共有%d行' % (file_dirs, count))
        
    #########
    #資料處理#
    #########
    for i in range(1,count):
        i+1
        keep = str(result[i]).split("\\t")
        varCommunity = str(keep[0]).strip("['")
        varPort = str(keep[1])
        location = str(keep[2])
        area = str(keep[3])
        IP = str(keep[4])
        oid = str(keep[5])
        sensor_name =str((keep[6]).strip("']"))
        
        get_OID()
        
def get_OID():
    global o,oidValue,all_info,new_arr

#-----SNMP取得OID-----#
    
    for (errorIndication,errorStatus,errorIndex,varBinds) in getCmd(SnmpEngine(),
        CommunityData(varCommunity, mpModel=0),UdpTransportTarget((IP, varPort)),
        ContextData(),ObjectType(ObjectIdentity(oid))):
        
        if errorIndication or errorStatus:
            #當錯誤時保留錯誤
            print(errorIndication or errorStatus)
            all_info = ['999-99-99 00:00', location , area , IP , oid , '99' ,  sensor_name   ] 
            new_arr.append(all_info)
            break
        else:
            for varBind in varBinds:
                #n=48
                #o='value:'+str(varBind)[45:n]
                #print(o)
                oidValue = (str(varBind).split("="))[1].strip(" ")
                #print(oidValue)
                if (sensor_name == "T-sensor"):
                    oidValue = int(oidValue)/10  #判斷數值處理(因為溫度數值產出的3位數)
                #逐行指令+入陣列
                all_info = [ time_now, location , area , IP ,  oid , oidValue , sensor_name ]
                new_arr.append(all_info)
                #new_arr=[['123','123'],[123,123],[123,123]]

def Sendsheet():
    global all_info

    sh.clear()
  #############
 #上傳資料處理#
#############
    print(new_arr)
    sh.update('A1', [["Time","Location", "Area" , "IP" , "OID"  , "Value", "Sensor"]])
    sh.update(Cell_Info, new_arr)       
     
def tick():
    global count,result,varCommunity,IP,varPort,oid
    reset()
    now()
    get_TXT()
    Sendsheet()

tick()

if __name__ == '__main__':
    scheduler = BlockingScheduler()
    scheduler.add_job(tick, 'cron', hour='00-23',minute='00-59') #每小時的每分鐘取得時間
    
    print('Press Ctrl+{0} to exit'.format('Break' if os.name == 'nt' else 'C    '))
    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        pass
