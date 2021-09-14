import pandas as pd 

df = pd.read_excel('2020年8月份工资核算9.7(1).xls', sheet_name = "结构")

df = df.T

flag = False
show_chart = False #如果想要看到详细步骤，就把这个value设置为True

# 这个for loop会显示给你看具体我走过了哪些数据，最后一个print是回车
# 有实际作用

if show_chart: print("处理之前\n\n")
for name in df.columns:
    for element in df[name]:
        if(element == "序号"): flag = not flag
        if show_chart: print(element, end = ' ')

    if(not flag): del df[name]
    if show_chart: print()

if show_chart: print("\n\n处理之后\n\n")

for name in df.columns:
    for element in df[name]:
        if show_chart: print(element, end = ' ')
    if show_chart: print()
    
# end of print, 显示例子结束

#获取所需要信息以及对应位置
need_list = ["岗位", "7月份人员", "基本工资", "职位工资", "周六加班费"]
map_need = {}

for index in range(len(df[df.columns[0]])):
    top_value = df[df.columns[0]][index]
    if show_chart: print(top_value)
    if top_value in need_list:
        map_need[index] = top_value

if show_chart: print(map_need)
#获取结束

del df[df.columns[0]]
all_people_map = {}
nan_value_flag = False

#输入人员信息
for name in df.columns:
    person_name = ""
    person_info = {}

    for index in range(len(df[name])):
        if(nan_value_flag): break
        if(index in map_need):

            if(map_need[index] == "7月份人员"): person_name = df[name][index]
            else: person_info[map_need[index]] = df[name][index]


        if show_chart: print(df[name][index], end = " ")
    
    all_people_map[person_name] = person_info
    #print (person_name)
    if show_chart: print()

#输入结束
#print output for debug

for person in all_people_map:
    print(person, all_people_map[person])