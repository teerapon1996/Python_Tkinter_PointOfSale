from audioop import reverse
from socket import timeout
from tkinter import *
from tkinter import ttk
from PIL import Image, ImageTk
from tkinter import font
from tkinter import messagebox
import mysql.connector as mycon
from datetime import datetime
import calendar;
import time;
from googleapiclient.discovery import build
from google.oauth2 import service_account



def connect_googlesheet_api():

    global SERVICE_ACCOUNT_FILE,SCOPES,creds,SAMPLE_SPREADSHEET_ID,service,sheet
    SERVICE_ACCOUNT_FILE = 'keys_marshell.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    SAMPLE_SPREADSHEET_ID = '1xSLPCAYmt0QRUauaPeTWarBeEhgsAMOSpIFdkydVBq8'
    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()

def read_write_sheet():
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range="save_list_marshell!A3:J").execute()
    global record, values_get
    values_get = result.get('values', [])
    
    for record in values_get:
        global count
        tree_save_list.insert(parent='', index='0', iid=count, text='',
                              values=(record[0], record[1], record[2], record[3], record[4],
                                      record[5],record[6], record[7], record[8],record[9])
                              )
        
        count += 1
    
def read_shell_money():

    ########
    Label(window_main, text="กำไรที่ได้", font=("tahoma", 15,"bold"), fg="black", bg="grey95").place(relx=.6, rely=.485)
    used_shell_val = StringVar()
    et_used_shell = Entry(window_main,textvariable=used_shell_val,font=("tahoma", 25,"bold"), width=7, bg="Lime",justify='right')
    et_used_shell.place(relx=.725, rely=.48)
    result_used_shell = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range="calculate!C11").execute()
    values_get_used_shell = result_used_shell.get('values', [])
    et_used_shell.insert(0, values_get_used_shell)

    #########
    Label(window_main, text="Shellเหลือ", font=("tahoma", 15,"bold"), fg="black", bg="grey95").place(relx=.6, rely=.555)
    current_shell_val = StringVar()
    et_current_shell = Entry(window_main,textvariable=current_shell_val,font=("tahoma", 25,"bold"), width=7, bg="Yellow",justify='right')
    et_current_shell.place(relx=.725, rely=.55)
    result_current_shell = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                             range="calculate!A9").execute()
    values_get_current_shell = result_current_shell.get('values', [])
    et_current_shell.insert(0, values_get_current_shell)

    ##############################################################################################

    Label(window_main, text="True Wallet", font=("tahoma", 15,"bold"), fg="black", bg="grey95").place(relx=.11, rely=.485)
    used_shell_val = StringVar()
    et_used_shell = Entry(window_main,textvariable=used_shell_val,font=("tahoma", 22,"bold"), width=9, bg="orange",justify='right')
    et_used_shell.place(relx=.25, rely=.48)
    result_used_shell = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range="calculate!E8").execute()
    values_get_used_shell = result_used_shell.get('values', [])
    et_used_shell.insert(0, values_get_used_shell)

    #########
    Label(window_main, text="K-Bank", font=("tahoma", 15,"bold"), fg="black", bg="grey95").place(relx=.155, rely=.555)
    current_shell_val = StringVar()
    et_current_shell = Entry(window_main,textvariable=current_shell_val,font=("tahoma", 22,"bold"), width=9, bg="green",justify='right')
    et_current_shell.place(relx=.25, rely=.55)
    result_current_shell = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range="calculate!F8").execute()
    values_get_current_shell = result_current_shell.get('values', [])
    et_current_shell.insert(0, values_get_current_shell)


def Add_list():
    ins_id = et_id.get()
    ins_price = et_pri.get()
    ins_use_shell = et_use_shell.get()
    ins_sm_shell = et_sm_shell.get()
    ins_fm_shell = et_fm_shell.get()
    ins_diamond = et_dia.get()
    if ins_id == 0 or ins_price == 0 or ins_use_shell =="" or ins_sm_shell =="" or ins_fm_shell =="" or ins_diamond =="" :
        messagebox.showerror("Failed Log", "กรุณาเลือกราคาก่อน 'บันทึกรายการ' ")
    else :
        ins_id = et_id.get()
        ins_price = et_pri.get()
        ins_use_shell = et_use_shell.get()
        ins_sm_shell = et_sm_shell.get()
        ins_fm_shell = et_fm_shell.get()
        ins_diamond = et_dia.get()

        if rd_val.get() == 1:
            ins_pay_type = "True Money"
        else:
            ins_pay_type = "กสิกรไทย"

        ins_pay_type2 = "ลูกค้าทั่วไป"
        gmt = time.gmtime()
        timestamp = calendar.timegm(gmt)
        get_times = str(datetime.fromtimestamp(timestamp))
        ins_times =get_times

        global inserinsert_to_sheet,request
        insert_to_sheet = [["",ins_id,ins_price,ins_use_shell,ins_sm_shell,ins_fm_shell,
                            ins_diamond,ins_pay_type,ins_pay_type2,ins_times]]

        request = service.spreadsheets().values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                                        range="save_list_marshell!A3:J",valueInputOption="USER_ENTERED",
                                                        insertDataOption="INSERT_ROWS" ,body={"values":insert_to_sheet}).execute()
        clear_treeview()
        clear_entry()
        read_write_sheet()
        read_shell_money()

def GetValue(event):

    et_id.config(state="normal")
    et_id.delete(0, END)
    et_dia.delete(0, END)
    et_pri.delete(0, END)
    et_fm_shell.delete(0, END)
    et_sm_shell.delete(0, END)
    et_use_shell.delete(0, END)
    
    global row_id,select
    row_id = price_tree.selection()[0]
    select = price_tree.set(row_id)
    
    et_id.insert(0, select['Id'])
    et_dia.insert(0, select['Diamond'])
    et_pri.insert(0, select['Price'])
    et_fm_shell.insert(0, select['fml_shell'])
    et_sm_shell.insert(0, select['sm_shell'])
    et_use_shell.insert(0, select['use_shell'])
    et_id.config(state="disabled")

def clear_treeview():
    tree_save_list.delete(*tree_save_list.get_children())

def clear_entry():
    global et_id,et_dia,et_pri,et_fm_shell,et_sm_shell,et_use_shell

    et_id.config(state="normal")
    et_id.delete(0, END)
    et_dia.delete(0, END)
    et_pri.delete(0, END)
    et_fm_shell.delete(0, END)
    et_sm_shell.delete(0, END)
    et_use_shell.delete(0, END)
    et_id.config(state="disabled")


connect_googlesheet_api()
################################ Start Main Window ####################################
window_main = Tk()
window_main = window_main
window_main.geometry("900x1000+0+0")
window_main.title("Main Program")
window_main.configure(bg="black")
#window_main.state('zoomed')
window_main.resizable(False, False) 

################################ Background Main Window ################################
bg = ImageTk.PhotoImage(file=r"images\main_marshell.png")
lbl_bg = Label(window_main, image=bg)
lbl_bg.place(x=0, y=0, relwidth=1, relheight=1)
################################ Lable ################################
Label(window_main, text="Administrator : Bus ", font=("tahoma", 12, "bold"), fg="black", bg="grey95").place(relx=.88, rely=.035, anchor="s")
Label(window_main, text="Developer by @Teerapon Meesuk ", font=("tahoma", 12, "bold"), fg="black", bg="grey95").place(relx=.2,rely=.999999,anchor="s")

################################ Logo Store ################################
img_ms = Image.open(r"images\logo_ms.png")
img_ms = img_ms.resize((125, 125), Image.ANTIALIAS)
photoimg_ms = ImageTk.PhotoImage(img_ms)
lbl_img_ms = Label(window_main, image=photoimg_ms, bg="grey95", borderwidth=0, anchor="n")
lbl_img_ms.place(width=125, height=125, relx=.45, rely=.475)

################ Button Open Save List connected Google Sheet ################
#btn_db = Button(window_main, text="รายการที่บันทึกแล้ว", command = open_window_db, font = ("tahoma", 20,"bold"),
                             #width=15,fg="white",bg="darkgreen").place(relx=.365,rely=.7)

##################### Main Frame General Custumer (ส่วนลูกค้าทั่วไป) ####################
frame_main_sub = Frame(window_main, bg="gold4")
frame_main_sub.place(width=855, height=395, anchor="center", relx=.5, rely=.25)
frame_main = Frame(window_main, bg="bisque2")
frame_main.place(width=830, height=370, anchor="center", relx=.5, rely=.25)
frame_line = Frame(window_main, bg="gold4")
frame_line.place(width=10, height=370, anchor="center", relx=.65, rely=.25)
frame_line = Frame(window_main, bg="gold4")
frame_line.place(width=10, height=370, anchor="center", relx=.35, rely=.25)

############################################################################

################################ Head Column ################################
Label(frame_main, text="เลือกรายการสินค้า", font=("tahoma", 20, "bold"), fg="purple1",bg="bisque2").place(relx=.167,rely=.15,anchor="s")
Label(frame_main, text="ราคาลูกค้าทั่วไป", font=("tahoma", 20, "bold"), fg="Red",bg="bisque2").place(relx=.50,rely=.15,anchor="s")
Label(frame_main, text="สูตรการเติม", font=("tahoma", 20, "bold"), fg="green",bg="bisque2").place(relx=.83,rely=.15,anchor="s")

########## ประกาศชนิดตัวแปรให้ Entry และ Radio ################
id_val = IntVar( )
dia_val = StringVar()
pri_val = IntVar()
fm_shell_val = StringVar()
sm_shell_val = StringVar()
use_shell_val = StringVar()
rd_val = IntVar()
rd_val.set(1)
74
######################## Show data on Entry when bind Click in Treeview (Product Table) #########################
Label(frame_main, text="รหัส", font=("tahoma", 18,"bold"), fg="black", bg="bisque2").place(relx=.365, rely=.24)
et_id = Entry(frame_main,textvariable=id_val, state="disabled" ,font=("tahoma", 25,"bold"), width=8, bg="grey",justify='right')
et_id.place(relx=.435, rely=.23)

Label(frame_main, text="เพชร", font=("tahoma", 18 ,"bold"), fg="black", bg="bisque2").place(relx=.35, rely=.41)
et_dia = Entry(frame_main,textvariable=dia_val, font=("tahoma", 25,"bold"), width=8, bg="cyan2",justify='right')
et_dia.place(relx=.435, rely=.40)

Label(frame_main, text="ราคา", font=("tahoma", 18,"bold"), fg="black", bg="bisque2").place(relx=.35, rely=.58)
et_pri = Entry(frame_main, textvariable=pri_val, bg="green1",font=("tahoma", 25,"bold"), width=8,justify='right')
et_pri.place(relx=.435, rely=.57)

rd_price_true = Radiobutton(frame_main,variable=rd_val, value=1,text="True Money",fg="orange", bg="bisque2", font=("tahoma", 13,"bold"))
rd_price_true.place(relx=.44, rely=.70)
rd_price_tran = Radiobutton(frame_main,variable=rd_val,value=2,text="โอนกสิกรไทย",fg="green", bg="bisque2", font=("tahoma", 13,"bold"))
rd_price_tran.place(relx=.44, rely=.78)

btn_add = Button(frame_main, text="📝 บันทึกรายการ", command = Add_list, font = ("tahoma", 12,"bold"),
                                width=15,fg="white",bg="darkgreen").place(relx=.40,rely=.88)

img_shell = Image.open(r"images\shell.png")
img_shell = img_shell.resize((30, 30), Image.ANTIALIAS)
photoimg_shell = ImageTk.PhotoImage(img_shell)
lbl_img_shell = Label(frame_main, image=photoimg_shell, bg="bisque2", borderwidth=0, anchor="n")
lbl_img_shell.place(width=30, height=30, relx=.93, rely=.06)

Label(frame_main, text="สูตรใช้เชลล์ในID", font=("tahoma", 15 ,"bold"), fg="black", bg="bisque2").place(relx=.74, rely=.20)
et_fm_shell = Entry(frame_main,textvariable=fm_shell_val, font=("tahoma", 20,"bold"), width=14, bg="yellow",justify='center')
et_fm_shell.place(relx=.69, rely=.30)

Label(frame_main, text="ใช้บัตร 30", font=("tahoma", 15,"bold"), fg="black", bg="bisque2").place(relx=.765, rely=.42)
et_sm_shell = Entry(frame_main, textvariable=sm_shell_val, text="ใบ",bg="yellow",font=("tahoma", 20,"bold"), width=5,justify='center')
et_sm_shell.place(relx=.775, rely=.52)

Label(frame_main, text="ใช้เชลล์ในIDไป", font=("tahoma", 15 ,"bold"), fg="black", bg="bisque2").place(relx=.75, rely=.67)
et_use_shell = Entry(frame_main,textvariable=use_shell_val, font=("tahoma", 20,"bold"), width=10, bg="red",justify='center')
et_use_shell.place(relx=.735, rely=.77)

######################## Show product table with Treeview (General Custumer) ############################
frame_tree = Frame(frame_main, bg="black")
frame_tree.place(width=270, height=292, anchor="center", relx=.165, rely=.57)

scroll_x = Scrollbar(frame_tree,orient=HORIZONTAL)
scroll_y = Scrollbar(frame_tree,orient=VERTICAL)
price_tree = ttk.Treeview(frame_tree,xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
scroll_x.pack(side=BOTTOM,fill=X)
scroll_y.pack(side=RIGHT,fill=Y)

scroll_x.config(command=price_tree.xview)
scroll_y.config(command=price_tree.yview)

price_tree['columns'] = ('Id', 'Diamond', 'dia', 'Price', 'bath','use_shell','fml_shell','sm_shell')
price_tree.column('#0', width=0, stretch=NO)
price_tree.column('Id', anchor=CENTER, width=40)
price_tree.column('Diamond', anchor="e", width=90)
price_tree.column('dia', anchor=CENTER, width=42)
price_tree.column('Price', anchor="e", width=45)
price_tree.column('bath', anchor=CENTER, width=40)
price_tree.column('use_shell', anchor=CENTER, width=120)
price_tree.column('fml_shell', anchor=CENTER, width=120)
price_tree.column('sm_shell', anchor=CENTER, width=85)

price_tree.heading('#0', text='', anchor=CENTER)
price_tree.heading('Id', text='รหัส', anchor=CENTER)
price_tree.heading('Diamond', text='สินค้า/เพชร', anchor=CENTER)
price_tree.heading('dia', text='💎', anchor=CENTER)
price_tree.heading('Price', text='ราคา', anchor=CENTER)
price_tree.heading('bath', text='฿', anchor=CENTER)
price_tree.heading('use_shell', text='ใช้ Shell ในID', anchor=CENTER)
price_tree.heading('fml_shell', text='สูตรใช้ Shell', anchor=CENTER)
price_tree.heading('sm_shell', text='ใช้บัตร 30', anchor=CENTER)


style = ttk.Style()
    # style.theme_use("default")
style.configure('Treeview',
                    background="rosybrown1",
                    foreground="rosybrown",
                    font=('Tahoma', 11),
                    rowheight=31,
                    fieldbackground="rosybrown1"
                    )
style.map('Treeview',
              background=[('selected', 'red')])

price_tree.pack(side=BOTTOM,fill=X)

    # insert
data = [
    [1, "รายสัปดาห์", 'บัตร', 66, 'บาท', 100, '100', 0],
    [2, "รายเดือน", 'บัตร', 298, 'บาท', 400, '400', 0],
    [3, "สิทธิ์ขั้นสูง", 'บัตร', 125, 'บาท', 180, '180', 0],
    [4, "12", 'คูปอง', 13, 'บาท', 0, '15', 0],
    [5, "68", 'เพชร', 22, 'บาท', 0, '30', 1],
    [6, "172", 'เพชร', 50, 'บาท', 75, '75', 0],
    [7, "309", 'เพชร', 90, 'บาท', 135, '135', 0],
    [8, "344", 'เพชร', 100, 'บาท', 150, '75+75', 0],
    [9, "481", 'เพชร', 140, 'บาท', 210, '135+75', 0],
    [10, "517", 'เพชร', 149, 'บาท', 225, '225', 0],
    [11, "618", 'เพชร', 180, 'บาท', 270, '135+135', 0],
    [12, "653", 'เพชร', 190, 'บาท', 285, '135+75+75', 0],
    [13, "689", 'เพชร', 199, 'บาท', 300, '225+75', 0],
    [14, "826", 'เพชร', 239, 'บาท', 360, '225+135', 0],
    [15, "861", 'เพชร', 249, 'บาท', 375, '225+75+75', 0],
    [16, "1,052", 'เพชร', 298, 'บาท', 450, '450', 0],
    [17, "1,224", 'เพชร', 348, 'บาท', 525, '450+75', 0],
    [18, "1,361", 'เพชร', 388, 'บาท', 585, '450+135', 0],
    [19, "1,396", 'เพชร', 398, 'บาท', 600, '450+75+75', 0],
    [20, "1,533", 'เพชร', 438, 'บาท', 660, '450+135+75', 0],
    [21, "1,569", 'เพชร', 447, 'บาท', 675, '450+225', 0],
    [22, "1,800", 'เพชร', 498, 'บาท', 750, '750', 0],
    [23, "2,109", 'เพชร', 586, 'บาท', 885, '750+135', 0],
    [24, "2,418", 'เพชร', 674, 'บาท', 1020, '750+135+135', 0],
    [25, "2,798", 'เพชร', 790, 'บาท', 1185, '750+225+135+75', 0],
    [26, "3,697", 'เพชร', 980, 'บาท', 1500, '1500', 0],
    [27, "7,394 ", 'เพชร', 1950, 'บาท', 3000, '1500+1500', 0],
    [28, "11,091", 'เพชร', 2910, 'บาท', 4500, '1500x3', 0],
    ]
count = 0
for record in data:
    price_tree.insert(parent='', index='end', iid=count, text='', values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7]))
    count += 1
    price_tree.bind('<Double-1>', GetValue)
    

def cmd_refresh():
    clear_treeview()
    clear_entry()
    read_write_sheet()
    read_shell_money()
    messagebox.showinfo('Successfuly','Refresh ข้อมูลเรียบร้อย !')


img_refresh = Image.open(r"images\reload.png")
img_refresh = img_refresh.resize((50, 50), Image.ANTIALIAS)
photoimg_refresh = ImageTk.PhotoImage(img_refresh)
btn_img_refresh = Button(window_main, image=photoimg_refresh,command=cmd_refresh ,bg="grey95", borderwidth=0, anchor="n")
btn_img_refresh.place(width=50, height=50, relx=.935, rely=.5695)


tree_save_sub = Frame(window_main, bg="black")
tree_save_sub.place(width=900, height=360, anchor="center", relx=.5, rely=.8)
tree_frame = Frame(tree_save_sub, bg="grey91")
tree_frame.place(width=880, height=340, relx=.01, rely=.025)

scroll_y = Scrollbar(tree_frame)
scroll_y.pack(side=RIGHT, fill=Y)

tree_save_list = ttk.Treeview(tree_frame, yscrollcommand=scroll_y.set)
scroll_y.config(command=tree_save_list.yview)


tree_save_list['columns'] = ('id_list_sold', 'id_product', 'price', 'use_shell', 'small_shell', 'fml_shell' ,
                                 'diamond','payment_type', 'custumer_type', 'sold_time_at')
tree_save_list.column('#0', width=0, stretch=NO)
tree_save_list.column('id_list_sold', anchor=CENTER, width=20)
tree_save_list.column('id_product', anchor=CENTER, width=20)
tree_save_list.column('price', anchor="e", width=30)
tree_save_list.column('use_shell', anchor="e", width=45)
tree_save_list.column('small_shell', anchor=CENTER, width=25)
tree_save_list.column('fml_shell', anchor=CENTER, width=60)
tree_save_list.column('diamond', anchor=CENTER, width=40)
tree_save_list.column('payment_type', anchor=CENTER, width=60)
tree_save_list.column('custumer_type', anchor=CENTER, width=60)
tree_save_list.column('sold_time_at', anchor=CENTER, width=110)


tree_save_list.heading('#0', text='', anchor=CENTER )
tree_save_list.heading('id_list_sold', text='ลำดับ', anchor=CENTER)
tree_save_list.heading('id_product', text='รหัส', anchor=CENTER)
tree_save_list.heading('price', text='ราคา', anchor=CENTER)
tree_save_list.heading('use_shell', text='ใช้เชลล์ไป', anchor=CENTER)
tree_save_list.heading('small_shell', text='บัตรเล็ก', anchor=CENTER)
tree_save_list.heading('fml_shell', text='สูตรเติม', anchor=CENTER)
tree_save_list.heading('diamond', text='เพชร', anchor=CENTER)
tree_save_list.heading('payment_type', text='โอนเงิน', anchor=CENTER)
tree_save_list.heading('custumer_type', text='ประเภทลูกค้า', anchor=CENTER)
tree_save_list.heading('sold_time_at', text='วัน / เวลา', anchor=CENTER)

s = ttk.Style()
    # style.theme_use("default")
s.configure('Treeview.Heading', font=('tahoma',11,'bold'))

tree_save_list.pack(fill=X)

read_write_sheet()
read_shell_money()


window_main.mainloop()