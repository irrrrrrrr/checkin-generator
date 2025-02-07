import random
import pandas as pd
import numpy as np


while True:
    try:
        in_hrs=int(input("上班时间（小时）："))
        if in_hrs>-1 and in_hrs<25:
            break
        else:
            print("输入错误！请输入正确的值！")
    except ValueError:
        print("输入错误！请输入一个整数！")

while True:
    try:
        in_min=int(input("上班时间（分钟）："))
        if in_min>-1 and in_min<61:
            break
        else:
            print("输入错误！请输入正确的值！")
    except ValueError:
        print("输入错误！请输入一个整数！")

while True:
    try:
        out_hrs=int(input("下班时间（小时）："))
        if out_hrs>-1 and out_hrs<25:
            break
        else:
            print("输入错误！请输入正确的值！")
    except ValueError:
        print("输入错误！请输入一个整数！")

while True:
    try:
        out_min=int(input("下班时间（分钟）："))
        if out_min>-1 and out_min<61:
            break
        else:
            print("输入错误！请输入正确的值！")
    except ValueError:
        print("输入错误！请输入一个整数！")



def create_2d_list(rows, cols, initial_value=None):
    return [[initial_value for _ in range(cols)] for _ in range(rows)]

in_time= create_2d_list(35, 31, 0)
out_time= create_2d_list(35, 31, 0)
time= create_2d_list(35, 31, 0)
file_path="data.xlsx"
excel_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"

def generate_list(q,input_arr):
    all_number = list(range(1, 36))
    random.shuffle(all_number)
    output_arr=[]
    if q==0:
        output_arr=all_number
    else:
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
                        swap2=output_arr[0]
                        output_arr[0]=all_number[0]
                        output_arr.append(swap2)
                        
                        break
    return output_arr

def create_matrix():
    matrix=[]
    temp=[]
    last_list=[]
    for i in range(31):
        temp=generate_list(i,last_list)
        matrix.append(temp)
        last_list=temp
    return np.transpose(matrix)


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



df = pd.DataFrame(time)
df.to_excel(file_path,header=False, index=False)
