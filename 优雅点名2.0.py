#2.0加入了分别输出迟到名单与按时名单
#加入了迟到与准时人数的统计
#减少使用难度 去除了手动输入人数
with open("chat_box.txt",encoding='UTF-8') as chat_box: 
    input_list = chat_box.read().splitlines()

    name_list = ['名字']

    ontime_list = []
    late_list = []

    member_ontime = 0
    member_late = 0
    
    
    i = 0
    while i < len(name_list) - 1:
        if name_list[i] in input_list:
            print(name_list[i],'ooo')
            ontime_list.append(1)
            ontime_list[member_ontime] = name_list[i]
            member_ontime = member_ontime + 1
        else:
            print(name_list[i],'xxx')
            late_list.append(1)
            late_list[member_late] = name_list[i]
            member_late = member_late + 1
        i = i+1
    
    print ('\n')
    print ('按时上课名单')
    print (*ontime_list,end='。')
    print ('共',len(ontime_list),'人')
    
    print ('\n')
    print ('迟到名单')
    print (*late_list,end='。')
    print ('共',len(late_list),'人')

print ('\n')
input('按任意键退出')

