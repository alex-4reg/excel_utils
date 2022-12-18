import time
import os
from plyer import notification
from openpyxl import load_workbook
from random import randint


def notify(title, timeout):
    notification.notify(title, timeout)


input_file = "./input/word_list.xlsx"
if input_file is None or not os.path.isfile(input_file):
    raise ValueError("'input_file' should be an existing file")

workbook = load_workbook(input_file)
sheet = workbook.active

variables = []
values = []

i: int = 1

while True:
    cell_word = sheet[f"A{i}"].value
    cell_definition = sheet[f"B{i}"].value
    i += 1
    if cell_word is not None:
        variables.append(cell_word)
        values.append(cell_definition)
    else:
        break

print(variables)

while True:
    index = randint(0, len(values) - 1)
    notification.notify(message=values[index], timeout=7)
    time.sleep(2)
    notification.notify(message=variables[index], timeout=7)
    time.sleep(5)
