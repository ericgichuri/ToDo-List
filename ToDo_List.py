from tkinter import *
import customtkinter as ctk
import os
import time,datetime,openpyxl
from tkinter import ttk,messagebox
from tkcalendar import DateEntry,Calendar

app=Tk()

screen_w=app.winfo_screenwidth()
screen_h=app.winfo_screenheight()
app_width=420
app_height=490
x_cord=(screen_w-app_width)//2
y_cord=(screen_h-app_height)//2
app.geometry("%dx%d+%d+%d"%(app_width,app_height,x_cord,y_cord))
app.resizable(False,False)
app.title("ESS ToDo List")
appicon=PhotoImage(file="icons/todo_list.png")
app.iconphoto(False,appicon)

today=time.strftime("%Y/%m/%d")
today1=datetime.date(int(time.strftime("%Y")),int(time.strftime("%m")),int(time.strftime("%d")))



btn_width=80
entry_width=150

font_title=("Poppins",15,"bold")

col_darkblue="#000088"
col_white="#FFFFFF"
col_black="#000000"
col_labels="#DCE600"
col_button="#DCE600"
col_red="#FF0000"

style=ttk.Style()
style.theme_use('default')
style.map("Treeview",
    background=[("selected",col_darkblue)],
    foreground=[("selected",col_white)]
)

app.configure(bg=col_darkblue)

cur_folder=os.getcwd()
file_folder=f'{cur_folder}\\file_folder'
if not os.path.exists(file_folder):
	os.makedirs(file)

file=f'{file_folder}\\todo_data.xlsx'
headers=['Data','Task']
if not os.path.exists(file):
	workbook=openpyxl.Workbook()
	sheet=workbook.active
	sheet.append(headers)
	try:
		workbook.save(file)
	except Exception as e:
		messagebox.showerror("Error",str(e))

def to_exit():
	msg=messagebox.askyesno("Confirm","Do you want to exit?")
	if msg==True:
		exit()

def func_readfile():
	if os.path.exists(file):
		wb=openpyxl.load_workbook(file)
		activesheet=wb.active
		maxrow=activesheet.max_row
		maxcols=activesheet.max_column
		outter_list=[]
		for i in range(2,maxrow+1):
			inner_list=[]
			if activesheet.cell(i,1).value==today:
				for j in range(1,maxcols+1):
					inner_list.append(activesheet.cell(i,j).value)
				outter_list.append(inner_list)

		for records in table_mytododata.get_children():
			table_mytododata.delete(records)
		for i in outter_list:
			table_mytododata.insert('',END,value=i)
def func_showall_todotask():
	if os.path.exists(file):
		wb=openpyxl.load_workbook(file)
		activesheet=wb.active
		maxrow=activesheet.max_row
		maxcols=activesheet.max_column
		outter_list=[]
		for i in range(2,maxrow+1):
			inner_list=[]
			for j in range(1,maxcols+1):
				inner_list.append(activesheet.cell(i,j).value)
			outter_list.append(inner_list)

		for records in table_mytododata.get_children():
			table_mytododata.delete(records)
		for i in outter_list:
			table_mytododata.insert('',END,value=i)

def func_save_tododata():
	if txt_task.get()=="":
		messagebox.showwarning("Error","Task is empty")
	else:
		todo_task=txt_task.get()
		todo_date=txt_date.get()

		wb=openpyxl.load_workbook(file)
		activesheet=wb.active
		msg=messagebox.askyesno("Confirm","Do you want to add this task?")
		if msg==True:
			maxrow=activesheet.max_row
			maxcols=activesheet.max_column
			adding=True
			for i in range(2,maxrow+1):
				if activesheet.cell(i,1).value==todo_date and activesheet.cell(i,2).value==todo_task:
					adding=False

			if adding==False:
				messagebox.showwarning("Warning","This task exists")
			else:
				activesheet.append([todo_date,todo_task])
				try:
					wb.save(file)
					messagebox.showinfo("Success","ToDo Task added successfully")
					func_clearform()
					func_readfile()
				except Exception as e:
					messagebox.showerror("Error",str(e))

def func_clearform():
	txt_task.delete(0,END)
	txt_date.configure(state='normal')
	txt_date.delete(0,END)
	txt_date.insert(END,today)
	txt_date.configure(state='readonly')

def func_delete_todotask():
	try:
		cur_item=table_mytododata.focus()
		items=table_mytododata.item(cur_item,"values")
		msg=messagebox.askyesno("Confirm",f"Do you want to delete this task? \n {items[1]}")
		if msg==True:
			wb=openpyxl.load_workbook(file)
			activesheet=wb.active
			maxrow=activesheet.max_row
			maxcols=activesheet.max_column
			for i in range(2,maxrow+1):
				if activesheet.cell(i,1).value==str(items[0]) and activesheet.cell(i,2).value==items[1]:
					activesheet.delete_rows(i)
			try:
				wb.save(file)
				messagebox.showinfo("Success","ToDo Task deleted successfully")
				func_readfile()
			except Exception as e:
				messagebox.showerror("Error",str(e))
	except:
		pass

lbl=ctk.CTkLabel(app,text="ESS ToDo List",text_color=col_labels,font=font_title)
lbl.pack(side=TOP,fill=X)
frame_form=ctk.CTkFrame(app,fg_color=col_darkblue)
frame_form.pack(side=TOP,fill=X)
lbl=ctk.CTkLabel(frame_form,text="Task",text_color=col_labels)
lbl.grid(column=0,row=0,pady=(4,4),padx=8)
txt_task=ctk.CTkEntry(frame_form,border_width=1,width=entry_width)
txt_task.grid(column=1,row=0,pady=(4,4),padx=8)
lbl=ctk.CTkLabel(frame_form,text="Date",text_color=col_labels)
lbl.grid(column=0,row=1,pady=(4,4),padx=8)
txt_date=DateEntry(frame_form,date_pattern="yyyy/mm/dd",mindate=today1,width=16,font=("times",12),state='readonly')
txt_date.grid(column=1,row=1,pady=(4,4),padx=8)
txt_date.configure(background='white', foreground='black',
                      selectbackground=col_darkblue, selectforeground='white', bordercolor='black')
btn_save=ctk.CTkButton(frame_form,text="Save",cursor="hand2",width=btn_width,fg_color=col_button,hover_color=col_button,text_color=col_darkblue,command=func_save_tododata)
btn_save.grid(column=2,row=0,pady=(4,4),padx=8)
btn_cancel=ctk.CTkButton(frame_form,text="Cancel",cursor="hand2",width=btn_width,fg_color=col_black,hover_color=col_black,command=func_clearform)
btn_cancel.grid(column=2,row=1,pady=(4,4),padx=8)

frame_viewtable=ctk.CTkFrame(app)
frame_viewtable.pack(side=TOP,fill=BOTH,expand=True)
table_mytododata=ttk.Treeview(frame_viewtable)
table_mytododata.pack(side=LEFT,fill=BOTH,expand=True)
scroll_table=ctk.CTkScrollbar(frame_viewtable,command=table_mytododata.yview)
scroll_table.pack(side=LEFT,fill=Y)
table_mytododata.configure(yscrollcommand=scroll_table.set)
table_mytododata['show']="headings"
table_mytododata['columns']=(0,1)
table_mytododata.heading(0,text="Date")
table_mytododata.heading(1,text="Task")
table_mytododata.column(0,anchor=CENTER)
table_mytododata.column(1)
frame_form1=ctk.CTkFrame(app,fg_color=col_darkblue)
frame_form1.pack(side=TOP)
btn_showall=ctk.CTkButton(frame_form1,text="Show All",cursor="hand2",width=btn_width,fg_color=col_button,hover_color=col_button,text_color=col_darkblue,command=func_showall_todotask)
btn_showall.grid(column=0,row=1,pady=(4,4),padx=8)
btn_showcurrent=ctk.CTkButton(frame_form1,text="Show Today",cursor="hand2",width=btn_width,fg_color=col_button,hover_color=col_button,text_color=col_darkblue,command=func_readfile)
btn_showcurrent.grid(column=1,row=1,pady=(4,4),padx=8)
btn_delete=ctk.CTkButton(frame_form1,text="Delete",cursor="hand2",width=btn_width,fg_color=col_red,hover_color=col_red,command=func_delete_todotask)
btn_delete.grid(column=2,row=1,pady=(4,4),padx=8)

#call functions
func_readfile()
app.protocol("WM_DELETE_WINDOW",to_exit)
app.mainloop()