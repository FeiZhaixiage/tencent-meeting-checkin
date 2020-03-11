with open("chat_box.txt",encoding='UTF-8') as chat_box: 
    input_list = chat_box.read().splitlines()
    member_size = 1

    name_list = ['名字']

    member_size = member_size - 1
    i = 0
    while i < member_size - 1:
        if name_list[i] in input_list:
            print(name_list[i],'ooo')
        else:
            print(name_list[i],'xxx')
        i = i+1
input('按任意键退出')
