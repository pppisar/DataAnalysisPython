import pandas as pd
import itertools

fileNames = ['Constitution - Ukraine.txt', 'Motrya - Lepky.txt']
alphabet = 'абвгґдеєжзиіїйклмнопрстуфхцчшщьюя '


def BigCalculations(fileName):
  # Open the file and extract all the text from it
  with open(f"Texts/{fileName}", 'r', encoding="utf-8") as file:
    text = file.read()
  # Create a dictionary with statistics on finding characters in a file
  BiGrams = {}
  # Convert all letters of the text to the same case
  text = text.lower()
  text = text.replace('\n', ' ')
  # Remove all characters from the text except letters and spaces
  cleartext = ''
  for item in text:
    if item in alphabet:
      cleartext += item
  # A variable for counting the total number of bigrams in the text
  lText = 0
  # Counting the number of bigrams
  for bgr in itertools.product(alphabet[:-1], repeat=2):
    count = cleartext.count(''.join(bgr))
    if count != 0:
      lText += count
      BiGrams[''.join(bgr)] = count
  for key in BiGrams:
    BiGrams[key] = BiGrams[key]/lText
  return BiGrams

def ThreeCalculations(fileName):
  # Open the file and extract all the text from it
  with open(f"Texts/{fileName}", 'r', encoding="utf-8") as file:
    text = file.read()
  # Create a dictionary with statistics on finding characters in a file
  ThreeGrams = {}
  # Convert all letters of the text to the same case
  text = text.lower()
  # Remove all characters from the text except letters and spaces
  cleartext = ''
  for item in text:
    if item in alphabet:
      cleartext += item
  # A variable for counting the total number of trigrams in the text:
  lText = 0
  # Count the number of trigrams
  for bgr in itertools.product(alphabet[:-1], repeat=3):
    count = cleartext.count(''.join(bgr))
    if count != 0:
      lText += count
      ThreeGrams[''.join(bgr)] = count
  for key in ThreeGrams:
    ThreeGrams[key] = ThreeGrams[key]/lText
  return ThreeGrams

def main():
  BigText1 = BigCalculations(fileNames[0])
  ThreeText1 = ThreeCalculations(fileNames[0])
  BigText2 = BigCalculations(fileNames[1])
  ThreeText2 = ThreeCalculations(fileNames[1])


  page1 = pd.DataFrame({'Letters': list(BigText1.keys()),
                        'Percents': list(BigText1.values())})
  page2 = pd.DataFrame({'Letters': list(ThreeText1.keys()),
                        'Percents': list(ThreeText1.values())})
  page3 = pd.DataFrame({'Letters': list(BigText2.keys()),
                        'Percents': list(BigText2.values())})
  page4 = pd.DataFrame({'Letters': list(ThreeText2.keys()),
                        'Percents': list(ThreeText2.values())})

  pages_sheets = {'Text1Bigrams': page1, 'Text1Threegrams': page2, 'Text2Bigrams': page3, 'Text2Threegrams': page4}
  writer = pd.ExcelWriter('./statistic2.xlsx', engine='xlsxwriter')
  for page_name in pages_sheets.keys():
    pages_sheets[page_name].to_excel(writer, sheet_name=page_name, index=False)
  writer.close()


if __name__ == "__main__":
  main()

