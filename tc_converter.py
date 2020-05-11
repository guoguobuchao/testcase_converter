from tkinter import *
from tkinter.filedialog import askdirectory,askopenfilename
import os
from read_xls import convert_excel_xml,convert_doc
def selectExcel():
    path_ = askopenfilename(title='Select Excel file', filetypes=[('*.xls', '.xls'), ('*.csv', '.csv'),('All Files', '*')],initialdir=os.path.abspath(os.curdir))
    excel_path.set(path_)
    
def selectWord():
    path_ = askopenfilename(title='Select Excel file', filetypes=[('*.docx', '.docx'), ('*.doc', '.doc'),('All Files', '*')],initialdir=os.path.abspath(os.curdir))
    word_path.set(path_)

def selectPath():
    path_ = askdirectory(title='Select XML output folder',initialdir=os.path.abspath(os.curdir))
    xml_path.set(path_)

def convert():
    msg=''
    word=word_path.get()
    excel=excel_path.get()
    xml=xml_path.get()
    if not excel:
        msg='Please select excel file!\n '
    if not xml:
        msg='Please specify the output folder!\n'
    if word and xml:
        try:
            xml_files=convert_doc(word,xml)
            msg='%s XML Files created'%len(xml_files)
        except Exception as e:
            msg='Error!!:     %s'%e
    if excel and xml:
        try:
            xml_files=convert_excel_xml(excel,xml)
            msg='%s XML Files created'%len(xml_files)
        except Exception as e:
            msg='Error!!:     %s'%e
    
    warning_msg.set(msg)
    
    
    
root = Tk()
option_frame=Frame(root)
info_frame=Frame(root)

excel_path = StringVar()
word_path=StringVar()
xml_path=StringVar()
warning_msg=StringVar()

root.title('Testcase xml converter')
root.resizable(width=False,height=True)
root.geometry('430x300')


word_lable=Label(option_frame,text = "Please select Testplan:")
word_entry=Entry(option_frame, textvariable = word_path,width=30)
word_button=Button(option_frame, text = "Select", command = selectWord)

excel_lable=Label(option_frame,text = "Please select Excel file:")
excel_entry=Entry(option_frame, textvariable = excel_path,width=30)
excel_button=Button(option_frame, text = "Select", command = selectExcel)

xml_lable=Label(option_frame,text = "Select XML output path:")
xml_entry=Entry(option_frame, textvariable = xml_path,width=30)
xml_button=Button(option_frame, text = "Select", command = selectPath)

convert_btn = Button(option_frame, text = "Convert", command = convert)

convert_label=Label(info_frame,textvariable=warning_msg,fg='red',wraplength=300,justify='left')

option_frame.grid(row=2)
info_frame.grid(row=3)

word_lable.grid(row = 0, column = 0,padx=10,pady=10,sticky='W')
word_entry.grid(row = 0, column = 1,padx=5,pady=10)
word_button.grid(row = 0, column = 2,padx=5,pady=10)

excel_lable.grid(row = 1, column = 0,padx=10,pady=10,sticky='W')
excel_entry.grid(row = 1, column = 1,padx=5,pady=10)
excel_button.grid(row = 1, column = 2,padx=5,pady=10)

xml_lable.grid(row = 2, column = 0,padx=10,pady=10,sticky='W')
xml_entry.grid(row = 2, column = 1,padx=5,pady=10)
xml_button.grid(row = 2, column = 2,padx=5,pady=10)
convert_btn.grid(row =3,column=1,pady=20)


convert_label.grid(row=3,sticky='W')


root.mainloop()
