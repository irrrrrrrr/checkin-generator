import random
import pandas as pd
import subprocess

##########################改这四个
in_hrs=7
in_min=30
out_hrs=17
out_min=0
##########################
def create_2d_list(rows, cols, initial_value=None):
    return [[initial_value for _ in range(cols)] for _ in range(rows)]

in_time= create_2d_list(35, 31, 0)
out_time= create_2d_list(35, 31, 0)
time= create_2d_list(35, 31, 0)
file_path = "D:\\previous versions\\output.xlsx"
excel_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"



def generate_list(input_arr):
    all_number = list(range(1, 36))
    random.shuffle(all_number)
    output_arr=[]
    for i in range(35):
        swap=random.choice(all_number)
        if swap!=input_arr[i]:
            output_arr.append(swap)
            all_number.remove(swap)
        else:
            while True:
                swap=random.choice(all_number)
                if swap!=input_arr[i]:
                    output_arr.append(swap)
                    all_number.remove(swap)
                    break
                elif len(all_number)==1:
                    for j in range(len(output_arr)):
                        if output_arr[j]!=all_number[0]:
                            swap2=output_arr[j]
                            output_arr[j]=all_number[0]
                            output_arr.append(swap2)
                            break
                    break
    return output_arr


def create_matrix():
    rows, cols = 35, 31
    matrix = [[] for _ in range(cols)]  # 转换为按列存储数据

    last_column = [None] * rows  # 初始化上一列

    # 按列生成数据
    for j in range(cols):
        current_column = []

        for i in range(rows):
            if i == 0:  # 列的第一个元素
                new_list = generate_list(last_column)
            else:
                new_list = generate_list(last_column)
            
            current_column.append(new_list[i])
            last_column[i] = new_list[i]  # 更新上一列的值为当前列的值

        matrix[j] = current_column

    # 转置矩阵以符合35行31列的要求
    return list(map(list, zip(*matrix)))



matrix = create_matrix()
for p in range(35):
    for q in range(31):
        temp_hrs=in_hrs
        if matrix[p][q]-in_min>0:
            temp_hrs=in_hrs-1
        if temp_hrs<11:
            temp_hrs='0'+str(temp_hrs)
        if temp_hrs==-1:
            temp_hrs=23
        temp_min=in_min-matrix[p][q]
        if temp_min<0:
            temp_min=60+temp_min
        if temp_min<10:
            temp_min=str('0'+str(temp_min))
        in_time[p][q]=str(temp_hrs)+':'+str(temp_min)

matrix = create_matrix()
for p in range(35):
    for q in range(31):
        temp_hrs=out_hrs
        if (matrix[p][q]+out_min)>60:
            temp_hrs=out_hrs+1
        if temp_hrs<11:
            temp_hrs='0'+str(temp_hrs)
        if temp_hrs==-1:
            temp_hrs=23
        temp_min=out_min+matrix[p][q]
        if temp_min>60:
            temp_min=temp_min-60
        if temp_min<10:
            temp_min=str('0'+str(temp_min))
        out_time[p][q]=str(temp_hrs)+':'+str(temp_min)
        time[p][q]=in_time[p][q]+"\n"+out_time[p][q]


# 创建DataFrame对象
df = pd.DataFrame(time)

# 将DataFrame写入Excel文件
df.to_excel("D:/previous versions/output.xlsx",header=False, index=False)

subprocess.run([excel_path, file_path], check=True)