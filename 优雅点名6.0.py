'''
优雅点名6.0
加入了突击点名模式
更改签到方式：加入了签到起始位置
权限获取声明：本软件不会进行用户数据分析请放心
剪切板修改权限：修改或读取剪切板内容
剪切板权限说明：读取签到信息，写入签到起始位置字符
文件读取写入权限：写入文件内容到设备
文件读取写入权限说明：创建/读写软件配置文件(confing.ini)
                      创建/读写签到名单文件(签到表.xls/用户导入的一切文件)
'''

import webbrowser as web #打开网页
import xlrd  #读表格
import xlwt  #写表格
from xlutils.copy import copy #表格备份
from tkinter import * #GUI
import tkinter.messagebox #弹出提示
import time #输出时间
import os #用于检测文件是否存在
import win32clipboard as w #剪切板
import win32con #剪切板
import tkinter.font as tkFont
from tkinter import filedialog
from tkinter import ttk
#import tkinter as tk

#打开网页
def open_github():
    web.open('https://github.com/FeiZhaixiage')

def open_project_web():
    web.open('https://github.com/FeiZhaixiage/tencent-meeting-checkin/releases/latest')

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

#获取剪切板内容
def gettext():
    w.OpenClipboard()
    t = w.GetClipboardData(win32con.CF_UNICODETEXT)
    w.CloseClipboard()
    return t

#写入剪切板内容
def settext():
    a=tkinter.messagebox.askokcancel('提示', '此操作会覆盖剪切板原本复制的内容''\n''是否继续')
    if a == True:
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_UNICODETEXT, write_clipboard)
        w.CloseClipboard()
        tkinter.messagebox.showinfo('提示','成功写入''\n''请前往聊天框粘贴')
    else:
        tkinter.messagebox.showinfo('提示','操作已取消')


#自动运行
def one_step():
    text_box.delete(0, END)
    write_text_box = gettext().split(check_code)#字符串切片
    text_box.insert(END,write_text_box[len(write_text_box)-1])
    main()
    


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
            create.save('签到表.xls')
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
    select["values"]=(user_ini)  
    select.current(0)


#表格选择
def select_path():
    global user_ini
    path = filedialog.askopenfilename()
    write_ini(path)
    read_ini()
    select_box()
    
#选择选项
def select_set(*args):
    global input_file
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
    #第一页
    first = Tk()
    first.geometry('460x240')
    first.title('优雅点名6.0')
    title_choose1 = Label(first,text='欢迎使用优雅点名')
    title_choose1.pack()
    title_choose1 = Label(first,text='本向导将引导您使用优雅点名')
    title_choose1.pack()
    start_btn = Button(first, text='下一步', command=first_exit)
    start_btn.place(x=370, y=210, height=20, width=80)
    first.mainloop()
    
    #第二页
    first2 = Tk()
    first2.geometry('460x240')
    first2.title('优雅点名6.0')
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
    about.insert(END,'https://github.com/FeiZhaixiage/tencent-meeting-checkin')
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

    #第三页
    first2 = Tk()
    first2.geometry('460x240')
    first2.title('优雅点名6.0')
    title_choose1 = Label(first2,text='____________________________')
    title_choose1.place(x=7,y=15)
    title_choose1 = Label(first2,text='权限获取说明')
    title_choose1.place(x=10,y=10)
    about = Text(first2)
    about.place(x=10, y=40, height=160, width=440)
    
    about.insert(END,'权限获取声明：本软件不会进行用户数据分析请放心')
    about.insert(END,'\n')
    about.insert(END,'剪切板修改权限：修改或读取剪切板内容')
    about.insert(END,'\n')
    about.insert(END,'剪切板权限说明：读取签到信息，写入签到起始位置字符')
    about.insert(END,'\n')
    about.insert(END,'文件读取写入权限：写入文件内容到设备')
    about.insert(END,'\n')
    about.insert(END,'文件读取写入权限说明：创建/读写软件配置文件(confing.ini)')
    about.insert(END,'\n')
    about.insert(END,'创建/读写签到名单文件(签到表.xls/用户导入的一切文件)')
    about.insert(END,'\n')
    about.insert(END,'打开电脑内应用权限：打开默认浏览器')
    about.insert(END,'\n')
    about.insert(END,'打开电脑内应用权限：打开网页来获取帮助或更新文件')
    about.insert(END,'\n')
    about.insert(END,'如不接受以上权限获取 请立即停止使用！')
    start_btn = Button(first2, text='下一步', command=first2_exit)
    start_btn.place(x=370, y=210, height=20, width=80)
    first2.mainloop()

    with open("confing.ini","w") as first_user_ini:
        first_user_ini.write('如果选项过多请适当删除失效路径[请勿删除此句]')

     
#定义路径
user_ini = []
input_file = '签到表.xls'
read_ini()

#剪切板内容创建
check_code = "https://github.com/FeiZhaixiage/tencent-meeting-checkin"
write_clipboard = '时间戳:' + str(time.time()) + "\n识别码:" + check_code + '\n现在可以签到了' 

#主窗口
root = Tk()
root.geometry('505x300')
root.title('优雅点名6.0')

text_box = Entry(root)
text_box.place(x=10, y=10, height=20, width=340)

start_btn = Button(root, text='一键运行', command=one_step)
start_btn.place(x=355, y=10, height=20, width=70)

start_btn = Button(root, text='运行', command=main)
start_btn.place(x=430, y=10, height=20, width=70)

start_btn = Button(root, text='选  择  文  件', command=select_path)
start_btn.place(x=355, y=40, height=20, width=145)



comvalue=tkinter.StringVar() 
select=ttk.Combobox(root,textvariable=comvalue) 
select_box() 
select.current(0)  
select.bind("<<ComboboxSelected>>",select_set)
select.place(x=10, y=40, height=20, width=340)

#select.destroy()

start_btn = Button(root, text='写入剪贴板', command=settext)
start_btn.place(x=355, y=70, height=20, width=70)

start_btn = Button(root, text='清空输入', command=clean)
start_btn.place(x=430, y=70, height=20, width=70)

out_box = Text(root)
out_box.place(x=10, y=70, height=220, width=340)

title_choose = Label(root,text='模  式  选  择')
title_choose.place(x=390, y=110,)

var = IntVar()
rd1 = Radiobutton(root,text="啰嗦模式",variable=var,value=0,command=type)
rd1.place(x=385, y=130)

rd2 = Radiobutton(root,text="简洁模式",variable=var,value=1,command=type)
rd2.place(x=385, y=150)

start_btn = Button(root, text='输 出 表 格', command=save)
start_btn.place(x=380, y=180, height=20, width=90)

start_btn = Button(root, text='GitHub', command=open_github)
start_btn.place(x=355, y=275, height=20, width=50)

start_btn = Button(root, text='更新', command=open_project_web)
start_btn.place(x=410, y=275, height=20, width=40)

title_choose = Label(root,text='版本:6.0')
title_choose.place(x=455, y=275,)


out_box.insert(END,'本软件使用python编写')
out_box.insert(END,'\n')
out_box.insert(END,'我的项目开源地址')
out_box.insert(END,'\n')
out_box.insert(END,'https://github.com/FeiZhaixiage/tencent-meeting-checkin/releases/latest')
out_box.insert(END,'\n')
out_box.insert(END,'您可访问上方链接取得最新版本')
out_box.insert(END,'\n')
out_box.insert(END,'非常感谢您使用本软件')
out_box.insert(END,'\n')
out_box.insert(END,'本软件只支持xls文件')
out_box.insert(END,'\n')
out_box.insert(END,'放入其他文件会导致软件错误请勿尝试！！！')
out_box.insert(END,'\n')
out_box.insert(END,'加入了突击点名模式')
out_box.insert(END,'\n')
out_box.insert(END,'更改签到方式：加入了签到起始位置')
out_box.insert(END,'\n')
out_box.insert(END,'加入了突击点名模式')
out_box.insert(END,'\n')
out_box.insert(END,'版本号:6.0')

root.mainloop()

