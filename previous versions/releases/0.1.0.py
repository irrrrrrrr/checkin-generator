import random
import pandas as pd

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




def generate_unique_random_list(existing_col):
    arr=[]
    new_list = list(range(1, 36))
    random.shuffle(new_list)

    # 检查条件，并重新洗牌直到满足条件
    while 1==1:
        for i in range(len(new_list)):
            if new_list[i] == existing_col[i]:
                random.shuffle(new_list)
                break
            arr.append(new_list[i])
            new_list.remove(new_list[i])
        if len(new_list)==1 and new_list[0]==existing_col[-1]:
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
                new_list = generate_unique_random_list(last_column)

            
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
print(len(matrix))
print(matrix)
#print(time)


# 创建DataFrame对象
df = pd.DataFrame(time)

# 将DataFrame写入Excel文件
df.to_excel("D:/previous versions/output.xlsx",header=False, index=False)
