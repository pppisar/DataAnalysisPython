import pandas as pd

fileNames = ['Constitution - Ukraine.txt', 'Motrya - Lepky.txt']
alphabet = 'абвгґдеєжзиіїйклмнопрстуфхцчшщьюя '

def Calculations(fileName):
    # Open the file and extract all the text from it
    with open(f"Texts/{fileName}", 'r', encoding="utf-8") as file:
        text = file.read()
    # Create a dictionary with statistics on finding characters in a file
    letters = dict.fromkeys(alphabet, 0)
    # Convert all letters of the text to the same case
    text = text.lower()
    # Remove all characters from the text except letters and spaces
    cleartext = ''
    for item in text:
        if item in alphabet:
            cleartext += item
    # Count the number of times each letter of the alphabet appears in the text
    for l in letters.keys():
        letters[l] = text.count(l) / len(cleartext)
    return letters

def main():
    Text1 = Calculations(fileNames[0])
    Text2 = Calculations(fileNames[1])

    page1 = pd.DataFrame({'Letters': list(Text1.keys()),
                          'Percents': list(Text1.values())})
    page2 = pd.DataFrame({'Letters': list(Text2.keys()),
                          'Percents': list(Text2.values())})

    pages_sheets = {'Text1Percents': page1, 'Text2Percents': page2}
    writer = pd.ExcelWriter('./statistic1.xlsx', engine='xlsxwriter')
    for page_name in pages_sheets.keys():
        pages_sheets[page_name].to_excel(writer, sheet_name=page_name, index=False)
    writer.close()


if __name__ == "__main__":
    main()