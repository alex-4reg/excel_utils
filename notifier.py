import time
from plyer import notification
from openpyxl import load_workbook
from random import randint


def notify(title, timeout):
    notification.notify(title, timeout)


workbook = load_workbook("./input/word_list.xlsx")
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
    time.sleep(10)
    notification.notify(message=variables[index], timeout=7)
    time.sleep(120)
