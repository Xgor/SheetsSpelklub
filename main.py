import openpyxl
import random
from tkinter.filedialog import askopenfilename
# Gets name for the selected row in 
def get_game_name(worksheet,number):
    cell_name = "{}{}".format("A", number)
    if worksheet != None:
        print(worksheet[cell_name].value)

def main():
    filename = askopenfilename()
    wb = openpyxl.load_workbook(filename)

    first_sheet = wb.sheetnames[0]
    worksheet = wb[first_sheet]

    game_count = worksheet.max_row

    print('Hur många spel vill du ha:')
    x = input()
    selected_games = random.sample(range(2, game_count), int(x))

    for game in selected_games:
        get_game_name(worksheet,game)
    x = input()

if __name__ == "__main__":
    main()
