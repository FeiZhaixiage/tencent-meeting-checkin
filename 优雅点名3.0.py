#优雅点名3.0 船新升级加入了易于操作的GUI！！！！
#加入了简洁模式和唠叨模式emmmm
#简易了软件运行的方式现在可以直接在软件进行粘贴了！
#废除 chat_box.txt

from tkinter import *

def type():
    choose = var.get()

def main():
    input_list = text_box.get()
    text_box.delete(0, END)
    name_list = ['名字']

    ontime_list = []
    late_list = []

    member_ontime = 0
    member_late = 0
    
    i = 0
    choose = var.get()
    if choose == 0:
        while i < len(name_list) - 1:
            if name_list[i] in input_list:
                txt =name_list[i]+'ooo''\n'
                out_box.insert(END,txt)
                
            else:
                txt = name_list[i]+'xxx'+'\n'
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
root = Tk()
root.geometry('460x240')
root.title('优雅点名')

text_box = Entry(root)
text_box.place(x=10, y=10, height=20, width=360)

start_btn = Button(root, text='运行', command=main)
start_btn.place(x=375, y=10, height=20, width=80)

out_box = Text(root)
out_box.place(x=10, y=40, height=190, width=360)

title_choose = Label(root,text='模式选择')
title_choose.place(x=390, y=60,)

var = IntVar()
rd1 = Radiobutton(root,text="啰嗦模式",variable=var,value=0,command=type)
rd1.place(x=375, y=90)

rd2 = Radiobutton(root,text="简洁模式",variable=var,value=1,command=type)
rd2.place(x=375, y=110)

root.mainloop()