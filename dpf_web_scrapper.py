import requests
from bs4 import BeautifulSoup
import xlsxwriter

data = []

url = "https://www.filtry-dpf-fap.pl/katalog-filtrow-dpf-fap-scr-cena"


result = requests.get(url)
doc = BeautifulSoup(result.text, "html.parser")

table = doc.find("table")
row = table.find_all("tr")

for r in row:
    temp = []
    cells = r.find_all("td")
    firstCell = cells[0].get_text(strip=True)
    secondCell = cells[1].get_text(strip=True)
    thirdCell = cells[2]

    thirdCellData = thirdCell.find_all("div")

    pojemnosc = ""
    oznaczenieSilnika = ""
    lataProdukcji = ""
    zamiennikJMJ = ""
    numerOE = ""
    zamiennikBosal = ""
    zamiennikWalker = ""
    zamiennikNovak = ""
    zamiennikBMCatalyst = ""
    numerDephi = ""
    konieMechaniczne = ""
    kilowaty = ""
    długość = ""
    numerVeneporte = ""
    normaEmisjiSpalin = ""

    i = 0
    while i < len(thirdCellData):
        pojemnosc = thirdCellData[0].get_text(strip=True).replace("Pojemność:", "")
        oznaczenieSilnika = thirdCellData[1].get_text(strip=True).replace("Oznaczenie silnika:", "")
        lataProdukcji = thirdCellData[2].get_text(strip=True).replace("Lata produkcji:", "")
        zamiennikJMJ = thirdCellData[3].get_text(strip=True).replace("Zamiennik JMJ:", "")
        numerOE = thirdCellData[4].get_text(strip=True).replace("Numer OE:", "")
        zamiennikBosal = thirdCellData[5].get_text(strip=True).replace("Zamiennik Bosal:", "")
        zamiennikWalker = thirdCellData[6].get_text(strip=True).replace("Zamiennik Walker:", "")
        zamiennikNovak = thirdCellData[7].get_text(strip=True).replace("Zamiennik Novak:", "")
        zamiennikBMCatalyst = thirdCellData[8].get_text(strip=True).replace("Zamiennik BM Catalysts:", "")
        numerDephi = thirdCellData[9].get_text(strip=True).replace("Numer Delphi:", "")
        konieMechaniczne = thirdCellData[10].get_text(strip=True).replace("Konie Mechaniczne:", "")
        kilowaty = thirdCellData[11].get_text(strip=True).replace("Kilowaty:", "")
        długość = thirdCellData[12].get_text(strip=True).replace("Długość:", "")
        numerVeneporte = thirdCellData[13].get_text(strip=True).replace("Numer Veneporte:", "")
        normaEmisjiSpalin = thirdCellData[14].get_text(strip=True).replace("Norma Emisji Spalin:", "")

        i += 1

    temp.append(firstCell)
    temp.append(secondCell)
    temp.append(pojemnosc)
    temp.append(oznaczenieSilnika)
    temp.append(lataProdukcji)
    temp.append(zamiennikJMJ)
    temp.append(numerOE)
    temp.append(zamiennikBosal)
    temp.append(zamiennikWalker)
    temp.append(zamiennikNovak)
    temp.append(zamiennikBMCatalyst)
    temp.append(numerDephi)
    temp.append(konieMechaniczne)
    temp.append(kilowaty)
    temp.append(długość)
    temp.append(numerVeneporte)
    temp.append(normaEmisjiSpalin)

    data.append(temp)

workbook = xlsxwriter.Workbook('dpf.xlsx')
worksheet = workbook.add_worksheet("dpf")

row = 0
col = 0

for marka, model, pojemnosc, oznaczenieSilnika, lataProdukcji, zamiennikJMJ, numerOE, zamiennikBosal, zamiennikWalker, zamiennikNovak, zamiennikBMCatalyst, numerDephi, konieMechaniczne, kilowaty, dlugosc, numerVeneporte, normaEmisjiSpalin in (data):
    worksheet.write(row, col, marka)
    worksheet.write(row, col + 1, model)
    worksheet.write(row, col + 2, pojemnosc)
    worksheet.write(row, col + 3, oznaczenieSilnika)
    worksheet.write(row, col + 4, lataProdukcji)
    worksheet.write(row, col + 5, zamiennikJMJ)
    worksheet.write(row, col + 6, numerOE)
    worksheet.write(row, col + 7, zamiennikBosal)
    worksheet.write(row, col + 8, zamiennikWalker)
    worksheet.write(row, col + 9, zamiennikNovak)
    worksheet.write(row, col + 10, zamiennikBMCatalyst)
    worksheet.write(row, col + 11, numerDephi)
    worksheet.write(row, col + 12, konieMechaniczne)
    worksheet.write(row, col + 13, kilowaty)
    worksheet.write(row, col + 14, dlugosc)
    worksheet.write(row, col + 15, numerVeneporte)
    worksheet.write(row, col + 16, normaEmisjiSpalin)
    row += 1

workbook.close()