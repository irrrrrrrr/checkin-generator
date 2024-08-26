import random
import pandas as pd


def generate(in_hrs, in_min, out_hrs, out_min, num, rev):
    def create_2d_list(rows, cols):
        return [[None for _ in range(cols)] for _ in range(rows)]

    def generate_first_array():
        arr = list(range(1, num + rev + 1))
        while True:
            random.shuffle(arr)
            if all(((abs(arr[i] - arr[i + 1]) > 1) and (arr[i] % 10) != (arr[i+1] % 10)) for i in range(num - 1)):
                return arr

    def generate_in_list(last_array):
        new_list = []
        all_numbers = list(range(1, num + rev + 1))
        while len(new_list) < num:
            available_numbers = [
                x for x in all_numbers
                if (abs(x - last_array[len(new_list)]) > 1 and
                    abs(x - last_array[max(len(new_list) - 1, 0)]) > 1 and
                    abs(x - last_array[min(len(new_list) + 1, num - 1)]) > 1 and
                    x % 10 != last_array[len(new_list)] % 10 and
                    (not new_list or x % 10 != new_list[-1] % 10))
            ]
            if not available_numbers:
                raise ValueError("No valid number found, need to backtrack or change algorithm")
            x = random.choice(available_numbers)
            new_list.append(x)
            all_numbers.remove(x)
        return new_list

    def create_in_matrix():
        temp_matrix = [generate_first_array()]
        while len(temp_matrix) < 31:
            try:
                x = generate_in_list(temp_matrix[-1])
                temp_matrix.append(x)
            except ValueError:
                return create_in_matrix()
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

    def generate_out_list(used_times, temp_matrix, last_array):
        new_list = []
        all_numbers = list(range(1, num + rev + 1))
        while len(new_list) < num:
            available_numbers = [
                x for x in all_numbers
                if (abs(x - last_array[len(new_list)]) > 1 and
                    abs(x - last_array[max(len(new_list) - 1, 0)]) > 1 and
                    abs(x - last_array[min(len(new_list) + 1, num - 1)]) > 1 and
                    x % 10 != last_array[len(new_list)] % 10 and
                    (not new_list or x % 10 != new_list[-1] % 10) and
                    f"{in_time[len(temp_matrix)][len(new_list)]};{x}" not in used_times)
            ]
            if not available_numbers:
                raise ValueError("No valid number found, need to backtrack or change algorithm")
            x = random.choice(available_numbers)
            new_list.append(x)
            all_numbers.remove(x)
            used_times.add(f"{in_time[len(temp_matrix)][len(new_list) - 1]};{x}")
        return new_list, used_times

    def create_out_matrix(in_time):
        temp_matrix = [generate_first_array()]
        used_times = set()
        for p in range(num):
            used_times.add(f"{in_time[0][p]};{temp_matrix[0][p]}")
        while len(temp_matrix) < 31:
            try:
                new_list, used_times = generate_out_list(used_times, temp_matrix, temp_matrix[-1])
                temp_matrix.append(new_list)
            except ValueError:
                return create_out_matrix(in_time)
        return temp_matrix

    def get_merge(in_time, out_time):
        time = create_2d_list(31, num)
        for p in range(31):
            for q in range(num):
                in_time[p][q] = generate_time_string(in_hrs, in_min, in_time[p][q], True)
                out_time[p][q] = generate_time_string(out_hrs, out_min, out_time[p][q], False)
                time[p][q] = f"{in_time[p][q]}\n{out_time[p][q]}"
        return trans_matrix(time)

    in_time = create_in_matrix()
    return get_merge(in_time, create_out_matrix(in_time))


def write(desktop_path, time):
    file_path = f"{desktop_path}\\data.xlsx"
    df = pd.DataFrame(time)
    df.to_excel(file_path, header=False, index=False)
