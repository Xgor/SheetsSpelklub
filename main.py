import openpyxl
import random

# Gets name for the selected row in 
def get_game_name(worksheet,number):
    cell_name = "{}{}".format("A", number)
    if worksheet != None:
        print(worksheet[cell_name].value)

def main():
    wb = openpyxl.load_workbook('Spel_bokklubb_.xlsx')

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
