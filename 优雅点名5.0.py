#优雅点名5.0 船新升级！！！！
#支持选择文件进行导入表格
#表格记忆功能上线

import xlrd  #读表格
import xlwt  #写表格
from xlutils.copy import copy #表格备份
from tkinter import * #GUI
import tkinter.messagebox #弹出提示
import time #输出时间
import os #用于检测文件是否存在
import tkinter.font as tkFont
from tkinter import filedialog
from tkinter import ttk
#import tkinter as tk

#引导关闭
def first_exit():
    first.destroy()   
def first2_exit():
    first2.destroy()

#选择状态
def type():
    choose = var.get()

#清除输入
def clean():
    text_box.delete(0, END)
#文本输出
def main():
    out_box.delete(1.0,END)
    if os.path.exists(input_file) == False:
        a=tkinter.messagebox.askokcancel('提示', '您还没有签到表''\n''需要创建吗')
        if a == True:
            create = xlwt.Workbook()
            worksheet = create.add_sheet('名单')
            worksheet.write(0,0,'姓名')
            worksheet.write(1,0,' ')
            create.save(input_file)
            tkinter.messagebox.showinfo('提示','签到表已创建''\n''请前往编辑')
        else:
            tkinter.messagebox.showinfo('提示','操作已取消')
    else:
        read_book = xlrd.open_workbook(input_file, formatting_info=True)
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
                while i < len(name_list):
                    if name_list[i] in input_list:
                        txt =name_list[i]+'√''\n'
                        out_box.insert(END,txt)
                
                    else:
                        txt = name_list[i]+'X'+'\n'
                        out_box.insert(END,txt)
                    i = i+1

            elif choose == 1:
                while i < len(name_list):
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

def select_box():
    comvalue=tkinter.StringVar() 
    select=ttk.Combobox(root,textvariable=comvalue) 
    select["values"]=(user_ini)  
    select.current(0)  
    select.bind("<<ComboboxSelected>>",select_set)
    select.place(x=10, y=40, height=20, width=360)

#表格选择
def select_path():
    global user_ini
    path = filedialog.askopenfilename()
    write_ini(path)
    read_ini()
    select.destroy()
    select_box()
    
#选择选项
def select_set(*args):
    global input_file
    print(select.get())
    input_file = select.get()
    

#写入ini文件
def write_ini(path):
    with open('confing.ini','a') as write_ini:
        write_ini.write('\n'+ path)

  
#读取用户ini文件
def read_ini():
    global user_ini
    with open("confing.ini", "r") as user_ini_r:
        a = 0
        user_ini = []
        for data in user_ini_r.readlines():
            data = data.strip('\n')  #去掉列表中每一个元素的换行符
            user_ini.append(data)
            a = a + 1

        
#表格保存
def save():
    if os.path.exists(input_file) == False:
        a=tkinter.messagebox.askokcancel('提示', '您还没有签到表''\n''需要创建吗')
        if a == True:
            create = xlwt.Workbook()
            worksheet = create.add_sheet('名单')
            worksheet.write(0,0,'姓名')
            worksheet.write(1,0,' ')
            create.save(input_file)
            tkinter.messagebox.showinfo('提示','签到表已创建''\n''请前往编辑')
        else:
            tkinter.messagebox.showinfo('提示','操作已取消')
    else:
        read_book = xlrd.open_workbook(input_file, formatting_info=True)
        main_data = read_book.sheets()[0]
            
        if main_data.cell(1,0).value == ' ':
            tkinter.messagebox.showinfo(input_file,'您还未添加学生信息')
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
            new_excel.save(input_file)

            tkinter.messagebox.showinfo('提示','已保存')


#引导
if os.path.exists("confing.ini") == False:
    first = Tk()
    first.geometry('460x240')
    first.title('优雅点名')
    title_choose1 = Label(first,text='欢迎使用优雅点名')
    title_choose1.pack()
    title_choose1 = Label(first,text='本向导将引导您使用优雅点名')
    title_choose1.pack()
    start_btn = Button(first, text='下一步', command=first_exit)
    start_btn.place(x=370, y=210, height=20, width=80)
    first.mainloop()

    first2 = Tk()
    first2.geometry('460x240')
    first2.title('优雅点名')
    title_choose1 = Label(first2,text='____________________________')
    title_choose1.place(x=7,y=15)
    title_choose1 = Label(first2,text='关于优雅点名')
    title_choose1.place(x=10,y=10)
    about = Text(first2)
    about.place(x=10, y=40, height=160, width=440)
    
    about.insert(END,'本软件使用python编写')
    about.insert(END,'\n')
    about.insert(END,'我的项目开源地址')
    about.insert(END,'\n')
    about.insert(END,'https://github.com/zpxrainbowdash/tencent-meeting-checkin')
    about.insert(END,'\n')
    about.insert(END,'您可访问上方链接取得最新版本')
    about.insert(END,'\n')
    about.insert(END,'非常感谢您使用本软件')
    about.insert(END,'\n')
    about.insert(END,'本软件免费开源请勿用于盈利工具')
    about.insert(END,'\n')
    about.insert(END,'如要美化此软件或进行改造请遵守GPL协议 GNU General Public License')
    about.insert(END,'\n')
    about.insert(END,'本软件要提取输入名单的姓名信息不会进行网络上传请放心')
    about.insert(END,'\n')
    about.insert(END,'如不同意或不接受以上条款 请立即停止使用！')
    start_btn = Button(first2, text='下一步', command=first2_exit)
    start_btn.place(x=370, y=210, height=20, width=80)
    first2.mainloop()

    with open("confing.ini","w") as first_user_ini:
        first_user_ini.write('如果选项过多请适当删除失效路径[请勿删除此句]')

     
#定义路径
user_ini = []
input_file = None
read_ini()

#主窗口
root = Tk()
root.geometry('460x280')
root.title('优雅点名')

text_box = Entry(root)
text_box.place(x=10, y=10, height=20, width=360)

start_btn = Button(root, text='运行', command=main)
start_btn.place(x=375, y=10, height=20, width=80)

start_btn = Button(root, text='选择文件', command=select_path)
start_btn.place(x=375, y=40, height=20, width=80)



comvalue=tkinter.StringVar() 
select=ttk.Combobox(root,textvariable=comvalue) 
select["values"]=(user_ini)  
select.current(0)  
select.bind("<<ComboboxSelected>>",select_set)
select.place(x=10, y=40, height=20, width=360)

#select.destroy()

start_btn = Button(root, text='清空输入', command=clean)
start_btn.place(x=375, y=70, height=20, width=80)

out_box = Text(root)
out_box.place(x=10, y=70, height=190, width=360)

title_choose = Label(root,text='模式选择')
title_choose.place(x=390, y=100,)

var = IntVar()
rd1 = Radiobutton(root,text="啰嗦模式",variable=var,value=0,command=type)
rd1.place(x=375, y=130)

rd2 = Radiobutton(root,text="简洁模式",variable=var,value=1,command=type)
rd2.place(x=375, y=150)

start_btn = Button(root, text='输出表格', command=save)
start_btn.place(x=390, y=190, height=20, width=50)

title_choose = Label(root,text='高一三班')
title_choose.place(x=390, y=215,)

title_choose = Label(root,text='赵埔渲')
title_choose.place(x=395, y=235,)

title_choose = Label(root,text='版本:5.0')
title_choose.place(x=410, y=260,)

out_box.insert(END,'本软件使用python编写')
out_box.insert(END,'\n')
out_box.insert(END,'我的项目开源地址')
out_box.insert(END,'\n')
out_box.insert(END,'https://github.com/zpxrainbowdash/tencent-meeting-checkin')
out_box.insert(END,'\n')
out_box.insert(END,'您可访问上方链接取得最新版本')
out_box.insert(END,'\n')
out_box.insert(END,'非常感谢您使用本软件')
out_box.insert(END,'\n')
out_box.insert(END,'本软件只支持xls文件')
out_box.insert(END,'\n')
out_box.insert(END,'放入其他文件会导致软件错误请勿尝试！！！')
out_box.insert(END,'\n')
out_box.insert(END,'已知BUG：')
out_box.insert(END,'\n')
out_box.insert(END,'第一次运行时无法输出任何内容')
out_box.insert(END,'\n')
out_box.insert(END,'解决办法：')
out_box.insert(END,'\n')
out_box.insert(END,'关闭程序重新打开')
out_box.insert(END,'\n')
out_box.insert(END,'版本号:5.0')

root.mainloop()

