#优雅点名4.0 船新升级！！！！
#去除编辑py文件
#支持从xls文件导入学生信息
#新增导出签到名单
#加入文件检测功能
#输出栏更改为单次输出的结果

import xlrd  #读表格
import xlwt  #写表格
from xlutils.copy import copy #表格备份
from tkinter import * #GUI
import tkinter.messagebox #弹出提示
import time #输出时间
import os #用于检测文件是否存在

def type():
    choose = var.get()

def clean():
    text_box.delete(0, END)
#文本输出
def main():
    out_box.delete(1.0,END)
    if os.path.exists("签到表.xls") == False:
        a=tkinter.messagebox.askokcancel('提示', '您还没有签到表''\n''需要创建吗')
        if a == True:
            create = xlwt.Workbook()
            worksheet = create.add_sheet('名单')
            worksheet.write(0,0,'姓名')
            worksheet.write(1,0,' ')
            create.save('签到表.xls')
            tkinter.messagebox.showinfo('提示','签到表已创建''\n''请前往编辑')
        else:
            tkinter.messagebox.showinfo('提示','操作已取消')
    else:
        read_book = xlrd.open_workbook('签到表.xls', formatting_info=True)
        main_data = read_book.sheets()[0]
        name_list = main_data.col_values(0)

        if main_data.cell(1,0).value == ' ':
            tkinter.messagebox.showinfo('提示','您还未添加学生信息')
        else:
            #取得输入数据
            input_list = text_box.get()


            ontime_list = []
            late_list = []

            member_ontime = 0
            member_late = 0
    
            i = 0
            choose = var.get()
            if choose == 0:
                while i < len(name_list) - 1:
                    if name_list[i] in input_list:
                        txt =name_list[i]+'√''\n'
                        out_box.insert(END,txt)
                
                    else:
                        txt = name_list[i]+'X'+'\n'
                        out_box.insert(END,txt)
                    i = i+1

            elif choose == 1:
                while i < len(name_list) - 1:
                    if name_list[i] in input_list:
                        ontime_list.append(1)
                        ontime_list[member_ontime] = name_list[i]
                        member_ontime = member_ontime + 1
                    else:
                        late_list.append(1)
                        late_list[member_late] = name_list[i]
                        member_late = member_late + 1
                    i = i+1
                out_box.insert(END,'准时签到名单')
                out_box.insert(END,'\n')
                ontime = ' '.join(ontime_list) + '\n'+'\n'
                out_box.insert(END,ontime)
                out_box.insert(END,'\n')
                out_box.insert(END,'迟到名单')
                out_box.insert(END,'\n')
                late = ' '.join(late_list) + '\n'
                out_box.insert(END,late)

        
#表格保存
def save():
    if os.path.exists("签到表.xls") == False:
        a=tkinter.messagebox.askokcancel('提示', '您还没有签到表''\n''需要创建吗')
        if a == True:
            create = xlwt.Workbook()
            worksheet = create.add_sheet('名单')
            worksheet.write(0,0,'姓名')
            worksheet.write(1,0,' ')
            create.save('签到表.xls')
            tkinter.messagebox.showinfo('提示','签到表已创建''\n''请前往编辑')
        else:
            tkinter.messagebox.showinfo('提示','操作已取消')
    else:
        read_book = xlrd.open_workbook('签到表.xls', formatting_info=True)
        main_data = read_book.sheets()[0]
            
        if main_data.cell(1,0).value == ' ':
            tkinter.messagebox.showinfo('提示','您还未添加学生信息')
        else:
            #读取表格
            name_list = main_data.col_values(0)
            write_place = main_data.ncols
            write_high =  main_data.nrows

            #取得输入数据
            input_list = text_box.get()
            state_list = []   #每人签到状态
            i = 0
            
            while i < len(name_list):
                if name_list[i] in input_list:
                    state_list.append(1)
                    state_list[i] = '√'
                else:
                    state_list.append(1)
                    state_list[i] = 'X'
                i = i+1

            new_excel = copy(read_book)
            ws = new_excel.get_sheet(0)
            i = 1
            
            while i <= write_high:
                ws.write(i-1,write_place,state_list[i-1]) #写入签到状态
                i = i + 1
            time_now = time.strftime("%m-%d %H:%M", time.localtime())
            ws.write(0,write_place,time_now)  #写入时间
            new_excel.save('签到表.xls')
    

root = Tk()
root.geometry('460x240')
root.title('优雅点名')

text_box = Entry(root)
text_box.place(x=10, y=10, height=20, width=360)

start_btn = Button(root, text='运行', command=main)
start_btn.place(x=375, y=10, height=20, width=80)

start_btn = Button(root, text='清空输入', command=clean)
start_btn.place(x=375, y=35, height=20, width=80)

out_box = Text(root)
out_box.place(x=10, y=40, height=190, width=360)

title_choose = Label(root,text='模式选择')
title_choose.place(x=390, y=70,)

var = IntVar()
rd1 = Radiobutton(root,text="啰嗦模式",variable=var,value=0,command=type)
rd1.place(x=375, y=100)

rd2 = Radiobutton(root,text="简洁模式",variable=var,value=1,command=type)
rd2.place(x=375, y=120)

start_btn = Button(root, text='输出表格', command=save)
start_btn.place(x=390, y=160, height=20, width=50)

root.mainloop()
