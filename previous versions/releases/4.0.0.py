import random
import pandas as pd
import subprocess

in_hrs = 6
in_min = 30
out_hrs = 16
out_min = 30
num = 30
rev = 15
#rev为13以上时保证有解，5以上需反复尝试

def create_2d_list(rows, cols):
    return [[None for _ in range(cols)] for _ in range(rows)]


def generate_first_array():
    arr = list(range(1, num + rev + 1))
    while True:
        random.shuffle(arr)
        if all(abs(arr[i] - arr[i + 1]) > 1 for i in range(num - 1)):
            return arr


def generate_in_list(last_array):
    new_list = []
    all_numbers = list(range(1, num + rev + 1))
    while len(new_list) < num:
        x = random.choice(all_numbers)
        temp = len(new_list)
        if (abs(x - last_array[temp]) > 1 and
                abs(x - last_array[max(temp - 1, 0)]) > 1 and
                abs(x - last_array[min(temp + 1, num - 1)]) > 1):
            new_list.append(x)
            all_numbers.remove(x)
    return new_list


def create_in_matrix():
    temp_matrix = [generate_first_array()]

    while len(temp_matrix) < 31:
        temp_matrix.append(generate_in_list(temp_matrix[-1]))
        
    return temp_matrix


def generate_time_string(base_hrs, base_min, offset, is_in_time):
    if is_in_time:
        temp_hrs = base_hrs - 1 if base_min - offset < 0 else base_hrs
    else:
        temp_hrs = base_hrs + 1 if base_min + offset >= 60 else base_hrs

    temp_min = (base_min + offset) % 60 if not is_in_time else (base_min - offset) % 60

    temp_hrs = f"{temp_hrs:02d}"
    temp_min = f"{temp_min:02d}"
    return f"{temp_hrs}:{temp_min}"


def trans_matrix(temp_matrix):
    temp_matrix = list(zip(*temp_matrix))
    temp_matrix = [list(row) for row in temp_matrix]
    return temp_matrix


def generate_out_list(used_times,temp_matrix,last_array):
    new_list = []
    all_numbers = list(range(1, num + rev + 1))
    while len(new_list) < num:
        x = random.choice(all_numbers)
        temp = len(new_list)
        if (abs(x - last_array[temp]) > 1 and
                abs(x - last_array[max(temp - 1, 0)]) > 1 and
                abs(x - last_array[min(temp + 1, num - 1)]) > 1 and
                f"{in_time[len(temp_matrix)][temp]};{x}" not in used_times):
            new_list.append(x)
            all_numbers.remove(x)
            used_times.add(f"{in_time[len(temp_matrix)][temp]};{x}")
    return new_list, used_times

def create_out_matrix(in_time):
    temp_matrix = [generate_first_array()]
    used_times = set()
    for p in range(num):
        used_times.add(f"{in_time[0][p]};{temp_matrix[0][p]}")
    while len(temp_matrix) < 31:
        new_list, used_times = generate_out_list(used_times,temp_matrix,temp_matrix[-1])
        temp_matrix.append(new_list)
    return temp_matrix


def get_merge(in_time, out_time):
    time = create_2d_list(31, num)
    for p in range(31):
        for q in range(num):
            in_time[p][q] = generate_time_string(in_hrs, in_min, in_time[p][q], True)
            out_time[p][q] = generate_time_string(out_hrs, out_min, out_time[p][q], False)
            time[p][q] = f"{in_time[p][q]}\n{out_time[p][q]}"
    return trans_matrix(time)


def write_and_run(time):
    file_path = "data.xlsx"
    excel_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
    df = pd.DataFrame(time)
    df.to_excel(file_path, header=False, index=False)
    subprocess.run([excel_path, file_path])

    
in_time = create_in_matrix()
time = get_merge(in_time,create_out_matrix(in_time))

write_and_run(time)

