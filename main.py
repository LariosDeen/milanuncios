from datetime import date

from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
from requests_html import HTMLSession


class Blocked(Exception):
    pass


def get_number(lnk):
    """Returns an integer number of properties on the web page by reference lnk."""
    r = session.get(lnk)
    try:
        about = r.html.find('.ma-ContentListingSummary-label', first=True)
    except Exception:
        return 0

    try:
        target_str = about.text
    except Exception:
        raise Blocked('Your enter is blocked. Wait 60 minutes or change IP')

    if ' anuncios' in target_str:
        target_number = int(target_str[:target_str.find(' anuncios')])
        return target_number
    else:
        return 0


zero = lambda x: '' if x >= 10 else '0'
towns = ('malaga', 'torremolinos')
# price ranges
prises = (('600', '699'), ('700', '799'), ('800', '899'), ('900', '999'))
lst = []
session = HTMLSession()
counter = 0
empty_space = 10

for town in towns:
    for prise_from, prise_to in prises:
        link = f'https://www.milanuncios.com/alquiler-de-pisos-en-{town}-malaga/?' \
               f'desde={prise_from}&hasta={prise_to}&demanda=n&banosd=2&dormd=3'
        # detect the number of objects
        number = get_number(link)
        lst.append(number)
        # visualization of the script process
        counter += 1
        empty_space -= 1
        numb = {1: 'st', 2: 'nd', 3: 'rd'}
        if counter < 4:
            suffix = numb[counter]
        else:
            suffix = 'th'
        print(f'Parsing of {zero(counter)}{counter}-{suffix} page' + ' ' * empty_space + ' ...')

for town in towns:
    link = f'https://www.milanuncios.com/alquiler-de-pisos-en-{town}-malaga/?demanda=n&banosd=1&dormd=4'
    number = get_number(link)
    lst.append(number)
    # visualization of the software process
    counter += 1
    empty_space -= 1
    print(f'Parsing of {zero(counter)}{counter}-th page' + ' ' * empty_space + ' ...')

print(lst)

# working with file Estadistica.xlsm
book = load_workbook(filename='Estadistica.xlsm', read_only=False, keep_vba=True)
sheet = book.active
# next_row - a string to be added to the end of an existing table (Estadistica.xlsm)
next_row = sheet.max_row + 1
# variable for formatting the added string to the current sheet of the file Estadistica.xlsm
medium_border = Border(left=Side(style='medium'),
                       right=Side(style='medium'),
                       top=Side(style='medium'),
                       bottom=Side(style='medium'))

# if the date of the first cell of the last row of the active sheet of the file
# Estadistica.xlsm is different from the current date
if str(sheet[sheet.max_row][0].value)[0:10] != str(date.today()):
    # add the current date to the beginning of the list lst
    lst.insert(0, date.today())
    # add a new line to the active sheet of the Estadistica.xlsm file after the last one
    sheet.append(lst)
    # set the formatting of the added line in the active sheet of the Estadistica.xlsm file
    sheet.row_dimensions[next_row].height = 23.6
    sheet[next_row][0].fill = PatternFill(fill_type='solid', start_color='FFB6FCC5')
    sheet[next_row][9].fill = PatternFill(fill_type='solid', start_color='FFF2F2F2')
    sheet[next_row][10].fill = PatternFill(fill_type='solid', start_color='FFF2F2F2')
    sheet[next_row][0].font = Font(size=10)
    for i in range(1, 12):
        sheet.cell(row=next_row, column=i).border = medium_border
        sheet.cell(row=next_row, column=i).alignment = Alignment(horizontal='center', vertical='center')
    for i in range(1, 11):
        sheet[next_row][i].font = Font(size=18)
    # set the numeric date format (column A) as in the previous cell on top
    sheet[next_row][0].number_format = sheet[next_row - 1][0].number_format

# save the changes to the file
book.save('Estadistica.xlsm')

# open the file Estadistica.xlsm for visual viewing
import os
os.startfile('Estadistica.xlsm')
