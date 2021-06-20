'''
To-do:
- Formatovani
- kopie tabulky
- data ze TXT souboru
- jmena hledat jen podle prijmeni, u duplicitnich vyresit posloupnost

'''

import functions
import io
import msoffcrypto
from openpyxl import Workbook, load_workbook, workbook
from openpyxl.utils import get_column_letter
# from openpyxl.workbook.protection import WorkbookProtection
from csv import reader
from os import path


# file = "Mzdové náklady-1-lock.xlsx"
# decrypted_wb = io.BytesIO()

# with open(file, 'rb') as f:
#     officeFile = msoffcrypto.OfficeFile(f)
#     officeFile.load_key(password='13881744')
#     officeFile.decrypt(decrypted_wb)

# wb = load_workbook(filename=decrypted_wb)



# Vyhleda zadane jmeno v prnim sloupci
def findName(name, depth=100):
    # print(name)
    for row in range(3,depth):
        bunka = ws.cell(column=1, row=row).value
        # print(bunka)
        if bunka is None:
            # print('Run motherfucker')
            continue
        else:
            # if ws.cell(column=1, row=row).value == name:
            if name in ws.cell(column=1, row=row).value:
                return row
    return False


# Vyhleda mesic v prvnim radku s defaultni hloubkou depth=100
def findMonth(month, width, depth=100):
    for column in range(2,depth,width):
        m = functions.strip_accents(ws.cell(column=column,row=1).value).lower()
        if m == month:
            return get_column_letter(column)
            # return column
    return False


def findMoney(name,month,width):

    col = findMonth(month,width)
    row = findName(name)

    if not col:
        # print('Mesic '+month+' nenalezen')
        return 0
    elif not row:
        # print('Jmeno '+name+' nenalezeno')
        return 0
    else:
        return ws[col+str(row)].coordinate


def findFood(name,month):

    col = ord(findMonth(month))-63
    row = findName(name)

    if not col:
        # print('Mesic '+month+' nenalezen')
        return 0
    elif not row:
        # print('Jmeno '+name+' nenalezeno')
        return 0
    else:
        return ws.cell(column=col, row=row).coordinate


def findDup(name):
    count = {}
    # for name in data:
    count.setdefault(name, 0)
    count[name] = count[name] + 1
    print(count)


def formatTxt(names):
    lines = [i for i, s in enumerate(names) if '-------' in s]

    s_lines = sorted(lines, reverse=True)

    for index in s_lines:
        if index == s_lines[-1]: # last element
            del names[:lines[0]+1]
        elif index == s_lines[0]: # first element
            del names[lines[-1]:]
        else:
            del names[index-1:index+1]

    data = {}
    for name in names:

        n = " ".join(name.split())
        splited = n.split(' ')
        mzda, jmeno = splited[-1].replace('.',''), splited[0]

        if '/' in mzda:
            mzda = ''
        else:
            mzda = int(mzda)

        # l = (jmeno, mzda)
        # data.append(l)
        data.setdefault(jmeno,[])
        data[jmeno].append(mzda)
    return data


# updateFile = 'C:/Users/bluem/Documents/Sandbox/python/automate_the_boring_stuff_with_python/excel/leden.csv'
# updateFile = 'leden.csv'

# with open(updateFile, 'r', encoding='windows-1250') as upd:
#     month = path.basename(updateFile).split('.')[0]
#     width = 6

#     csv_reader = reader(upd)
#     list_of_rows = list(map(tuple,csv_reader)) # Output: [('Martin', 1200), ('Jana',1500),...]

updateFile = 'TEXT.TXT'
file = "Mzdové náklady 2021-clean.xlsx"
wb = load_workbook(file)

with open(updateFile,'r',encoding='windows-1250') as f:
    width = 6
    month = 'unor'
    names = f.readlines()
    data = formatTxt(names)

    seznamStredisek = {}

    for worksheet in wb.worksheets:
        if worksheet.title not in ['Celkový součet', 'kontrola', 'List1']:
            print()
            if worksheet.title in ['Úřad práce', 'Úřad práce Cheb', 'Katastrální úřad']:
                width = 4
            else:
                width = 6
            ws = wb[worksheet.title]
            # print(worksheet.title.upper().center(60,'-'))

            # ws = wb['Sklad Dalovice']
            mesic = findMonth(month,width)

            for i in range(1,100):
                bunka = ws.cell(row = i, column = 1)
                # print(f'Bunka: {bunka.coordinate}')
                cele_jmeno = bunka.value
                # print(f'cele_jmeno: {cele_jmeno}')
                if cele_jmeno == 'Zákonné pojištění' or cele_jmeno == 'Mzdové náklady': # Konec jmen
                    break
                if cele_jmeno:
                    cele_jmeno = cele_jmeno.split(' ')
                    # print(f'cele_jmeno[list]: {cele_jmeno}')
                    if len(cele_jmeno) > 1:
                        prijmeni = cele_jmeno[0]
                        seznamStredisek.setdefault(worksheet.title,{})
                        seznamStredisek[worksheet.title].setdefault(prijmeni,[])
                        seznamStredisek[worksheet.title][prijmeni].append(mesic+str(bunka.row))

                        # seznamStredisek.setdefault(prijmeni,{})
                        # seznamStredisek[prijmeni].setdefault(worksheet.title,[])
                        # seznamStredisek[prijmeni][worksheet.title].append(mesic+str(bunka.row))
                        # if prijmeni in data:
                        #     print(f'{prijmeni} ma plat {data[prijmeni]}')

    for stredisko,seznamJmen in seznamStredisek.items():
        ws = wb[stredisko]
        print()
        print(stredisko)
        for jmeno,bunka in seznamJmen.items():
            # print(jmeno, bunka)
            for i, v in enumerate(bunka):
                # print(jmeno, v)
                if jmeno in data:
                    print(f'{jmeno}:{v}={data[jmeno][i]}')
                    del data[jmeno][i]
                    # ws[v].value = data[jmeno][i]

    # for k,v in data.items():
    #     for m in v:
    #         print(f'{k}={m}') # output Fiala=33258; k je jmeno, m je castka



# wb.save(file)
'''
with open(updateFile, 'r', encoding='windows-1250') as f:
    width = 6
    month = 'leden'
    names = f.readlines()
    data = formatTxt(names)

    for worksheet in wb.worksheets:
        if worksheet.title not in ['Celkový součet', 'kontrola', 'List1']:
            # print()
            if worksheet.title in ['Úřad práce', 'Úřad práce Cheb', 'Katastrální úřad']:
                width = 4
                # stravenky = 0

            print(worksheet.title.upper().center(60,'-'))
            ws = wb[worksheet.title]

            # for person in list_of_rows[:]: # Z listu se po kazdem obehu maze aktualizovana osoba
            for person in data[:]: # Z listu se po kazdem obehu maze aktualizovana osoba
                money = findMoney(person[0],month, width) # najde bunku
                count = {}

                if money:
                    if person[1] == '': # Pokud neni vyplnena mzda, je delka tuple 1 a vznika Index error
                        cash = '' # '0' a '' je v excelove tabulce jina hodnota kvuli automatickemu vypsani refundaci
                        ws[money].value=cash
                        # ws[money+2].value=straveky # pridat stravenky do csv
                    else:
                        cash = person[1]
                        ws[money].value=int(cash)
                    message = str('-- Update '+person[0] + ':')+str(cash).rjust(10)
                    message = f'> Update | {person[0]} | {cash:>25}'
                    print(message)
                    data.remove(person)
                # else:
                #     print('-- Name '+person[0]+ ' not found')
                #     continue
    # if len(list_of_rows) > 0:
    if len(data) > 0:
        print('Nenalezeny tyto polozky:')
        # for person in list_of_rows:
        for person in data:
            print(f'{person[0]}:{str(person[1])}')

# wb.security.workbookPassword = 'abc'
# print(WorkbookProtection.workbook_password)
# wb.save(file)

'''
