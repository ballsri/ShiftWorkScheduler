from tkinter import *
from tkinter import filedialog, messagebox
from PIL import ImageTk, Image
import os, sys
import pandas as pd
import datetime
from genSchedule import genSchedule as gs

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

root = Tk()
root.title("ShiftWorkScheduler")
# root.geometry("640x480")
root.resizable(False, False)
root.iconphoto(True, ImageTk.PhotoImage(Image.open(resource_path("icon.jpg"))))

# Global variables
root.filename = ""
df = pd.DataFrame()
month = 0
sel_year = 2023

# Create a frame
frame = Frame(root)
frame.pack()

# Create a label
label_frame = Frame(frame)
label_frame.grid(row = 0, column = 0, padx = 10, pady = 10)

label = Label(label_frame, text = "โปรแกรมจัดตารางเวรจากไฟล์ Excel", font = ("Arial", 20))
label.grid(row = 0, column = 0, padx = 10, pady = 10)

year_label = Label(label_frame, text = "อัพเดทปี : 2023", font = ("Arial", 20))
year_label.grid(row = 0, column = 1, padx = 10, pady = 10)

# year_select = Spinbox(label_frame, from_ = sel_year, to = 2026, width = 5, font = ("Arial", 20) )
# year_select.grid(row = 0, column = 2, padx = 10, pady = 10)

# Create an section
input_frame = Frame(frame)
input_frame.grid(row = 1, column = 0, padx = 10, pady = 10)

lFrame = Frame(input_frame)
lFrame.grid(row = 1, column = 0, padx = 10, pady = 10)
rFrame = Frame(input_frame)
rFrame.grid(row = 1, column = 1, padx = 10, pady = 10)

# Create an input button
def openExcel():
    global df
    try:
        root.filename = filedialog.askopenfilename( title = "Select an Excel File", filetypes = (("Excel Worksheet", "*.xlsx"), ("Excel 97-2003 Worksheet", "*.xls")))
        df = pd.read_excel(root.filename, sheet_name = "Sheet1")
        input_btn.config(text = "นำเข้าข้อมูลเรียบร้อยแล้ว", state= DISABLED)
    except:
        messagebox.showinfo("Error", "โปรดเลือกไฟล์ Excel หรือ ปิด Excel ก่อนนำเข้าข้อมูล")
    
input_label = Label(lFrame, text = "กดปุ่มเพื่อนำเข้าข้อมูลจาก Excel", font = ("Arial", 20))
input_label.grid(row = 0, column = 0, padx = 10, pady = 10)

input_btn = Button(lFrame, text = "เลือกไฟล์ข้อมูล Excel", font= ('Arial', 20), command = openExcel, bd=5)
input_btn.grid(row = 1, column = 0, padx = 10, pady = 10)

def resetData():
    # graphical reset
    input_btn.config(text = "เลือกไฟล์ข้อมูล Excel", state = NORMAL)
    input_month_label.config(text = "โปรดเลือกเดือน")
   
    listbox.selection_clear(0, END)
    listbox.activate(0)
    listbox.selection_set(0)
    listbox.selection_anchor(0)

     # data reset
    global df
    global month
    root.filename = ""
    df = pd.DataFrame()
    month = 0


reset_btn = Button(lFrame, text = "รีเซ็ตข้อมูล", font = ('Arial', 20), bd = 5, command = resetData)
reset_btn.grid(row = 2, column = 0, padx = 10, pady = 10)

# Create month select frame
month_frame = Frame(rFrame)
month_frame.grid(row = 1, column = 0, padx = 10, pady = 10)

# Create a lable
month_text = "โปรดเลือกเดือน"
input_month_label = Label(month_frame, text = month_text, font = ("Arial", 20))
input_month_label.grid(row = 0, column = 0, padx = 10, pady = 10)

# Create submit month button

def selectMonth():
    global month

    if listbox.get(ANCHOR) == "":
        messagebox.showinfo("Error", "โปรดเลือกเดือน")
    else:
        input_month_label.config(text = "เดือน : " + listbox.get(ANCHOR))
        month = months.index(listbox.get(ANCHOR)) + 1
        

submit_month_btn = Button(month_frame, text = "ยืนยันเดือน", font = ('Arial', 20), bd= 5, command= selectMonth)
submit_month_btn.grid(row = 0, column = 1, padx = 10, pady = 10)

# Create listbox
months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
listbox = Listbox(rFrame, width = 20, height = 10, font = ('Arial', 20) )
listbox.grid(row = 0, column = 0, padx = 10, pady = 10)
for m in months:
    listbox.insert(END, m)
listbox.activate(0)
listbox.selection_set(0)
listbox.selection_anchor(0)

# Create a generate button
def genSchedule():
    if root.filename == "":
        messagebox.showinfo("Error", "โปรดนำเข้าข้อมูลจาก Excel")
    elif month == 0:
        messagebox.showinfo("Error", "โปรดเลือกเดือน")
    else:
        
        status = gs(df, month,months[month-1], int(root.filename[-9:-5]))
        if status == 0:
            messagebox.showinfo("Success", "สร้างตารางเวรเรียบร้อยแล้ว")
        elif status == 1:
            messagebox.showinfo("Error", "ปิด Excel ก่อนสร้างตารางเวร")
        elif status == 2:
            messagebox.showinfo("Error", "โปรดเลือกไฟล์ให้ถูกต้อง")

submit_label = Label(lFrame, text = "กดปุ่มเพื่อสร้างตารางเวร", font = ("Arial", 20))
submit_label.grid(row = 3, column = 0, padx = 10, pady = 10)

gen_btn = Button(lFrame, text = "สร้างตารางเวร", font = ('Arial', 20), bd = 5, command = genSchedule)
gen_btn.grid(row = 4, column = 0, padx = 10, pady = 10)



root.mainloop()

