import os
import openpyxl as OP

wb = OP.load_workbook(os.path.join(os.path.dirname(__file__),'HabitTrackerDiscordBot.xlsx'))
ws = wb['BaseRef']

def main():
    Output = []
    for column in ws:
        Output.append(column[0].value)
        Output.append(" : ")
        Output.append(column[1].value)
        Output.append("\n")

    print("".join(map(str,outputtable(1, 1, 2, 5))))

def numtochar(input:int):
    start_index = 1   #  it can start either at 0 or at 1
    letter = ''
    while input > 25 + start_index:
        letter += chr(65 + int((input-start_index)/26) - 1)
        input = input - (int((input-start_index)/26))*26
    letter += chr(65 - start_index + (int(input)))
    return letter

def outputtable(colstart:int, rowstart:int, colsize:int, rowsize:int):
    Output = []
    for row in range(rowstart, rowstart + rowsize):
        for col in range(colstart, colstart + colsize):
            Output.append(ws[numtochar(col) + str(row)].value)
            if((col - colstart) < colsize - 1):
                Output.append(" : ")
        Output.append("\n")
    return Output

main()
