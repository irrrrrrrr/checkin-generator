import random
import pandas as pd
import numpy as np

##########################改这四个
in_hrs=7
in_min=30
out_hrs=17
out_min=0
##########################
def create_2d_list(rows, cols, initial_value=None):
    a = [[initial_value for _ in range(cols)] for _ in range(rows)]

    return a


in_time= create_2d_list(35, 31, 0)
out_time= create_2d_list(35, 31, 0)
time= create_2d_list(35, 31, 0)

import random  
  
def find_different_permutation(original_array):  
    # 复制原始数组  
    shuffled_array = original_array[:]  
  
    # 辅助函数，用于递归地寻找不同的排列  
    def backtrack(current_index):  
        nonlocal shuffled_array  # 声明shuffled_array为非局部变量，以便在内部修改  
  
        # 如果已经检查完所有元素，返回True表示找到满足条件的排列  
        if current_index == len(shuffled_array):  
            return True  
  
        # 尝试与当前元素不同的所有可能值  
        for i in range(current_index, len(shuffled_array)):  
            # 如果当前位置的值与原始数组中的值相同，或者与已经放置在新排列中的值相同  
            # 则交换当前位置的值和另一个位置的值  
            if shuffled_array[i] == original_array[current_index] or shuffled_array[i] in shuffled_array[:current_index]:  
                shuffled_array[i], shuffled_array[current_index] = shuffled_array[current_index], shuffled_array[i]  
  
                # 递归检查下一个位置  
                if backtrack(current_index + 1):  
                    return True  
  
                # 如果递归返回False，则撤销交换，尝试下一个可能值  
                shuffled_array[i], shuffled_array[current_index] = shuffled_array[current_index], shuffled_array[i]  
  
        # 如果没有找到满足条件的排列（即所有可能的值都试过且都失败），返回False  
        return False  
  
    # 随机打乱复制的数组（虽然这不是必要的，但可以增加多样性）  
    random.shuffle(shuffled_array)  
  
    # 从第一个位置开始检查并回溯  
    if not backtrack(0):  
        raise ValueError("Cannot find a permutation with no matching elements.")  
  
    return shuffled_array  


def generate_unique_random_list(existing_col):
    arr=[]
    new_list = list(range(1, 36))
    random.shuffle(new_list)

    # 检查条件，并重新洗牌直到满足条件
    while True:
        for i in range(len(new_list)):
            if new_list[i] == existing_col[i]:
                random.shuffle(new_list)
                break
            arr.append(new_list[i])
            new_list.remove(new_list[i])
        if len(new_list)==1 and new_list[0]==existing_col[-1] and new_list[0]!=arr[0] :
            temp=arr[0]
            arr[0]=new_list[0]
            arr.append(temp)
            break
    return arr


def create_matrix():
    rows, cols = 35, 31
    matrix = [[] for _ in range(cols)]  # 转换为按列存储数据
    last_column = [0] * rows  # 初始化上一列

    # 按列生成数据
    for j in range(cols):
        current_column = []

        for i in range(rows):
            if i == 0:  # 列的第一个元素
                new_list = list(range(1,36))
                random.shuffle(new_list)
                current_column=new_list
                break
            else:
                new_list = generate_unique_permutation(last_column)

            
            current_column=new_list
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
        temp_hrs='0'+str(temp_hrs)
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
        temp_min=out_min+matrix[p][q]
        if temp_min>60:
            temp_min=temp_min-60
        if temp_min<10:
            temp_min=str('0'+str(temp_min))
        out_time[p][q]=str(temp_hrs)+':'+str(temp_min)
        time[p][q]=in_time[p][q]+"\n"+out_time[p][q]

#for row in matrix:
#    print(row)
#print(in_hrs_arr)
#print(in_time)
#print(out_time)
#print(len(matrix))
print(matrix)
#print(time)


# 创建DataFrame对象
df = pd.DataFrame(time)

# 将DataFrame写入Excel文件
df.to_excel("D:/previous versions/output.xlsx",header=False, index=False)
