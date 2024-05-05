import openpyxl as OP

wb = OP.load_workbook('HabitTrackerDiscordBot.xlsx')
ws = wb['BaseRef']
my_list = list()

for value in ws.iter_rows(
    min_row=1, max_row=6, min_col=1, max_col=2, 
    values_only=True):
    my_list.append(value)

for i in range(1, 5):
    print(my_list[i * 2] + " : " + my_list[(i * 2) + 1])
