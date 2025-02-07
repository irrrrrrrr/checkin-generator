import random
import pandas as pd
import subprocess

##########################改这六个
in_hrs = 7
in_min = 30
out_hrs = 17
out_min = 0
num = 35
rev = 5
##########################

def create_2d_list(rows, cols, initial_value=None):
    return [[initial_value for _ in range(cols)] for _ in range(rows)]

in_time = create_2d_list(num, 31, "")
out_time = create_2d_list(num, 31, "")
time = create_2d_list(num, 31, "")
file_path = "data.xlsx"
excel_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"

def generate_first_array():
    arr = list(range(1, num + rev + 1))
    while True:
        random.shuffle(arr)
        if all(abs(arr[i] - arr[i+1]) > 1 for i in range(num - 1)):
            return arr

def generate_list(last_array):
    new_list = []
    all_numbers = list(range(1, num + rev + 1))
    while len(new_list) < num:
        x = random.choice(all_numbers)
        l = len(new_list)
        if (abs(x - last_array[l]) > 1 and
            abs(x - last_array[max(l-1, 0)]) > 1 and
            abs(x - last_array[min(l+1, num-1)]) > 1):
            new_list.append(x)
            all_numbers.remove(x)
    return new_list

def create_matrix():  
    matrix = []
    matrix.append(generate_first_array())

    while len(matrix) < 31:  
        temp = generate_list(matrix[-1])  
        matrix.append(temp)  
    matrix = list(zip(*matrix))
    matrix = [list(row) for row in matrix]
    return matrix

def generate_time_string(base_hrs, base_min, offset, is_in_time):
    if is_in_time:
        temp_hrs = base_hrs - 1 if base_min - offset < 0 else base_hrs
    else:
        temp_hrs = base_hrs + 1 if base_min + offset >= 60 else base_hrs

    temp_min = (base_min + offset) % 60 if not is_in_time else (base_min - offset) % 60

    temp_hrs = f"{temp_hrs:02d}"
    temp_min = f"{temp_min:02d}"
    return f"{temp_hrs}:{temp_min}"

matrix = create_matrix()
for p in range(num):
    for q in range(31):
        in_time[p][q] = generate_time_string(in_hrs, in_min, matrix[p][q], True)

matrix = create_matrix()
for p in range(num):
    for q in range(31):
        out_time[p][q] = generate_time_string(out_hrs, out_min, matrix[p][q], False)

used_times = set()
for p in range(num):
    for q in range(31):
        in_out_time = f"{in_time[p][q]}\n{out_time[p][q]}"
        while in_out_time in used_times:
            matrix = create_matrix()
            in_time[p][q] = generate_time_string(in_hrs, in_min, random.randint(1, num + rev), True)
            out_time[p][q] = generate_time_string(out_hrs, out_min, random.randint(1, num + rev), False)
            in_out_time = f"{in_time[p][q]}\n{out_time[p][q]}"
        time[p][q] = in_out_time
        used_times.add(in_out_time)

df = pd.DataFrame(time)
df.to_excel(file_path, header=False, index=False)
subprocess.run([excel_path, file_path])
