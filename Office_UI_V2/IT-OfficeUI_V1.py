import time
def __New_GUI__():
    import os
    import tkinter as tk  # 使用Tkinter前需要先匯入   
    import tkinter.ttk as ttk
    window = tk.Tk()
    window.title('My Window')
    window.geometry('800x200+500+500')
    
    #按鈕定位
    Name=0; X=1; Y=2; W=3; H=4; Path=5; #Table表使用;不需要更改
    btn_table1 = [
                ['Name'  ,'X' , 'Y' , 'W' , 'H', 'Path'],
                ['資產自動化上傳\nFAF-Update'   ,'50', '50', '20', '2', r'.\API\1_Project-IT-Office-APPS\FAF.exe'],
                ['借用自動化上傳\nBORROW-Update','50', '100', '20', '2', r'.\API\1_Project-IT-Office-APPS\BORROW.exe'],
                ['公告自動化上傳\nMEMO-Update'  ,'50', '150', '20', '2', r'.\API\1_Project-IT-Office-APPS\MEMO.exe'],
                
                ['需求單&附件\nITR & Attachment ','250', '50', '20', '2', r'\\10.231.250.70\Department\Form\HO-ITD'],
                ['SNMP環控','625', '100', '20', '2', r'.\API\2_Project-SNMP'],
                ['CSV自動化上傳','625', '150', '20', '2', r'.\API\3_Project-CSV']
                ]   
    btn_table2 = [
                ['Name'  ,'X' , 'Y' , 'W' , 'H', 'Path'],
                ['需求單&附件\nITR & Attachment ','250', '50', '20', '2', r'\\10.231.250.70\Department\Form\HO-ITD'],
                ['-','250', '100', '20', '2', r''],
                ['-'  ,'250', '150', '20', '2', r'']
                ]
    btn_table3 = [
                ['Name'  ,'X' , 'Y' , 'W' , 'H', 'Path'],               
                ['SNMP環控','625', '50', '20', '2', r'.\API\2_Project-SNMP'],
                ['CSV自動化上傳','625', '100', '20', '2', r'.\API\3_Project-CSV'],
                ['IoT環境監測'  ,'625', '150', '20', '2', r'.\API\Project-IoT(尚未整合)']
                ]               
## btn_table1
    def __b1_1__():
        Var=1
        os.startfile(btn_table1[Var][Path])
        
    def __b1_2__():
        Var=2
        os.startfile(btn_table1[Var][Path])

    def __b1_3__():
        Var=3
        os.startfile(btn_table1[Var][Path])
# btn_table2
    def __b2_1__():
        Var=1
        os.startfile(btn_table2[Var][Path])
    def __b2_2__():
        Var=2
        os.startfile(btn_table2[Var][Path])
    def __b2_3__():
        Var=3
        os.startfile(btn_table2[Var][Path])
# btn_table3        
    def __b3_1__():
        Var=1
        os.startfile(btn_table3[Var][Path])
    def __b3_2__():
        Var=2
        os.startfile(btn_table3[Var][Path])
    def __b3_3__():
        Var=3
        os.startfile(btn_table3[Var][Path])
    
#第1欄
    tk.Label(window, text="APP ",font=32,bd="16",width="15",height="1").place(x= 50 , y= 0 , anchor='nw') 
    Var=1 ;#請參考Table表;按行數輸入從0開始
    tk.Button(window, text= btn_table1[Var][Name] , width= btn_table1[Var][W], height= btn_table1[Var][H],cursor=("dotbox"), command= __b1_1__ ).place(x= btn_table1[Var][X] , y= btn_table1[Var][Y] , anchor='nw')   
    Var=2 ;
    tk.Button(window, text= btn_table1[Var][Name] , width= btn_table1[Var][W], height= btn_table1[Var][H],cursor="dotbox", command= __b1_2__ ).place(x= btn_table1[Var][X] , y= btn_table1[Var][Y] , anchor='nw')
    Var=3 ;
    tk.Button(window, text= btn_table1[Var][Name] , width= btn_table1[Var][W], height= btn_table1[Var][H],cursor="dotbox", command= __b1_3__ ).place(x= btn_table1[Var][X] , y= btn_table1[Var][Y] , anchor='nw')

#第2欄
    tk.Label(window, text="NAS-PATH",font=32,bd="16",width="15",height="1").place(x= 250 , y= 0 , anchor='nw')
    Var=1 ;
    tk.Button(window, text= btn_table2[Var][Name] , width= btn_table2[Var][W], height= btn_table2[Var][H],cursor="dotbox", command= __b2_1__ ).place(x= btn_table2[Var][X] , y= btn_table2[Var][Y] , anchor='nw')
    Var=2 ;
    tk.Button(window, text= btn_table2[Var][Name] , width= btn_table2[Var][W], height= btn_table2[Var][H],cursor="dotbox", command= __b2_2__ ).place(x= btn_table2[Var][X] , y= btn_table2[Var][Y] , anchor='nw')
    Var=3 ;
    tk.Button(window, text= btn_table2[Var][Name] , width= btn_table2[Var][W], height= btn_table2[Var][H],cursor="dotbox", command= __b2_3__ ).place(x= btn_table2[Var][X] , y= btn_table2[Var][Y] , anchor='nw')
    
#第3欄    
    Var=1 ;
    tk.Button(window, text= btn_table3[Var][Name] , width= btn_table3[Var][W], height= btn_table3[Var][H],cursor="clock", command= __b3_1__ ).place(x= btn_table3[Var][X] , y= btn_table3[Var][Y] , anchor='nw')
    Var=2 ;
    tk.Button(window, text= btn_table3[Var][Name] , width= btn_table3[Var][W], height= btn_table3[Var][H],cursor="clock", command= __b3_2__ ).place(x= btn_table3[Var][X] , y= btn_table3[Var][Y] , anchor='nw')
    Var=3 ;
    tk.Button(window, text= btn_table3[Var][Name] , width= btn_table3[Var][W], height= btn_table3[Var][H],cursor="clock", command= __b3_3__ ).place(x= btn_table3[Var][X] , y= btn_table3[Var][Y] , anchor='nw')
    
    
    tk.Label(window, text="更新日期:2021-8-26\n開發人員:Ness_huang ",font=5,bd="16",width="30",height="1").place(x= 550 , y= 0 , anchor='nw')
    window.mainloop()

__New_GUI__()