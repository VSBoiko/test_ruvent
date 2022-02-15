from datetime import datetime, timedelta

import openpyxl


def count_dates(worksheet, column, func):
    cnt = 0
    for val in read_col(worksheet, column):
        cnt += 1 if func(val) else 0

    return cnt


def count_numbers(worksheet, column, func):
    cnt = 0
    for val in read_col(worksheet, column):
        if isinstance(val, str):
            val = upd_str_to_float(val)
        cnt += 1 if func(val) else 0

    return cnt


def is_even_number(num: float):
    return num % 2 == 0


def is_prime_number(num: float):
    for d in range(2, int(num // 2)):
        if num % d == 0:
            return False
    return True


def is_suitable_number(num: float):
    return num < 0.5


def is_tuesday_1(s_date: str):
    return "Tue" in s_date


def is_tuesday_2(s_date: str):
    day = datetime.strptime(s_date, '%Y-%m-%d %H:%M:%S.%f')
    return day.weekday() == 1


def is_last_tue_in_month(s_date: str):
    day = datetime.strptime(s_date, '%m-%d-%Y')
    if day.weekday() != 1:
        return False

    next_tuesday = day + timedelta(days=7)
    if day.month == next_tuesday.month:
        return False

    return True


def upd_str_to_float(s: str):
    s = "".join(s.split())
    s = s.replace(",", ".")
    return float(f'0{s}') if s[0] == "." else float(s)


def read_col(worksheet, column: str, start_row: int = 2):
    for i in range(start_row, worksheet.max_row):
        yield worksheet[column][i].value


if __name__ == '__main__':
    ws_tasks = openpyxl.load_workbook('./task_support.xlsx')["Tasks"]

    print(count_numbers(ws_tasks, "B", is_even_number))             # Сколько четных чисел в этом столбце?
    print(count_numbers(ws_tasks, "C", is_prime_number))            # Сколько простых чисел в этом столбце?
    print(count_numbers(ws_tasks, "D", is_suitable_number))         # Сколько чисел, меньших 0.5 в этом столбце?
    print(count_dates(ws_tasks, "E", is_tuesday_1))                 # Столько вторников в этом столбце?
    print(count_dates(ws_tasks, "F", is_tuesday_2))                 # Столько вторников в этом столбце?
    print(count_dates(ws_tasks, "G", is_last_tue_in_month))         # Сколько последних вторников месяца в этом столбце?
