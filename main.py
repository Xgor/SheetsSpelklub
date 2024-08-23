import openpyxl
import random
from tkinter.filedialog import askopenfilename

def main():
    # Fråga efter excel dokument
    filename = askopenfilename( filetypes =[('Excel files', '*.xlsx')])
    wb = openpyxl.load_workbook(filename)

    # Öppnar spellistan
    first_sheet = wb.sheetnames[0]
    worksheet = wb[first_sheet]

    # Skapar en array av alla spel
    games = []
    for i in range(2, worksheet.max_row):
        g = worksheet["{}{}".format("A", i)].value
        if g != None:
            games.append(g)
            # Lägger till extra kopior om det har lagts extra röster in i H kolumnen
            copies = worksheet["{}{}".format("H", i)].value
            if copies == None: copies = 0
            for j in range(int(copies)):
                games.append(g)

    print('Hur många spel vill du ha:')
    game_count = input()
    selected_games = []
    # Lägg in antalet spel valda
    for i in range(int(game_count)):
        game = games[random.randrange(0, len(games))]
        selected_games.append(game)
        # Tar bort duplikationerna av valda spelet i listan
        games = [j for j in games if j != game]
        pass

    for game in selected_games:
        print(game)
    x = input()

if __name__ == "__main__":
    main()