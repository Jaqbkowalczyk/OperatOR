#!/usr/bin/env python
# -*- coding: utf-8 -*-

import time
import os
import re
import logging
import pyautogui
from PyPDF3 import PdfFileMerger
import webbrowser as web
import win32com.client as win32
from decimal import Decimal
from docx import Document
import tkinter as tk
from shapely.geometry import Polygon
from tkinter.filedialog import askdirectory, askopenfile
from tkinter import Menu, scrolledtext
from tkinter import messagebox
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.opc.exceptions import PackageNotFoundError
from docx.shared import Pt
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import csv

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
pyautogui.FAILSAFE = True
JEDNOSTKAREJESTROWA = ''
KERG = ''
FOLDER = ''
XYACCURACY = 2
HACCURACY = 2
ANGLEACCURACY = 4
AREAACCURACY = 0

def find_kerg(filename):  # todo regex
    logging.debug(f'{filename}')
    if filename.split('.')[-1] == 'rtf':
        filename = convertrtftodocx(filename)
    text = str(get_text_from_doc(filename))
    logging.debug(f'{text}')
    kerg = re.findall(r"[\d.]+", text)
    print(kerg)
    kerg = set(kerg)
    print(kerg)


def set_kerg(kerg_value: str):
    global KERG
    KERG = kerg_value
    logging.debug(f'KERG został ustawiony na: {KERG}')


def ask_for_kerg():
    pass


def set_folder(folder_value: str):
    global FOLDER
    FOLDER = folder_value
    logging.debug(f'FOLDER został ustawiony na: {FOLDER}')


def open_folder():
    tk.Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askdirectory()  # show an "Open" dialog box and return the path to the selected folder
    return filename


def open_file():
    logging.debug('hej')
    tk.Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfile('r')  # show an "Open" dialog box and return the path to the selected folder
    return filename


def convertrtftodocx(file):
    word = win32.Dispatch("Word.Application")
    wdFormatDocumentDefault = 16
    wdHeaderFooterPrimary = 1
    doc = word.Documents.Open(file)
    for pic in doc.InlineShapes:
        pic.LinkFormat.SavePictureWithDocument = True
    for hPic in doc.sections(1).headers(wdHeaderFooterPrimary).Range.InlineShapes:
        hPic.LinkFormat.SavePictureWithDocument = True
    file = file.split('.')[0]
    doc.SaveAs(str(file + ".docx"), FileFormat=wdFormatDocumentDefault)
    doc.Close()
    word.Quit()
    return str(file + ".docx")


def get_text_from_doc(filename):
    doc = Document(filename)
    fullText = []
    i = 0
    j = 0
    for para in doc.paragraphs:
        fullText.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                fullText.append(cell.text)
                j += 1
            i += 1
            j = 0
        i = 0
    return '\n'.join(fullText)


def find_info():
    for subdir, dirs, files in os.walk(FOLDER):
        for file in files:
            if file == "info_o_materiałach12.rtf":
                logging.debug('Znalazłem info o materiałach!')
                logging.debug(f'{os.path.join(subdir, file)}')
                return os.path.join(subdir, file)
            else:
                pass


def check_project_data():  # todo find a way to extract text from .rtf file
    """Function to check for variables used in all future functions"""
    # 1. Select project folder
    folder = open_folder()
    set_folder(folder)
    # 2. Find data from PODGiK (.gml file, info o materialach)
    # text = get_text_from_doc(find_info())

    print(text)
    # 3. Search through .gml file to find JEDNOSTKAREJESTROWA and OBREB values
    # 4. Find KERG number
    set_kerg('666.2250.2021')
    pass


def write_report():
    """Function to write report file using given values"""
    s = "I love #stackoverflow# because #people# are very #helpful# #helpful#"
    hashtag = re.findall(r"#(\w+)#", s)  # znajdź wszystkie hashtagi w szablonie
    print(set(hashtag))
    document = Document()
    for paragraph in document.paragraphs:
        if 'sea' in paragraph.text:
            print(paragraph.text)
            paragraph.text = 'new text containing ocean'


def createparcelfilefromdoc(file, parcelsfile):
    """This function creates parcel text file from .docx table with parcels"""
    doc = Document(file)
    parcels = []
    i = 0
    j = 0
    newline = lambda x: x + '\n'
    output = open(os.getcwd() + '\\' + parcelsfile, 'w')
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if j == 1:
                    cell_parcels = cell.text.split(',')
                    logging.debug(f'Text komórki: {cell.text}')
                    for parcel in cell_parcels:
                        parcel = parcel.replace(' ', '')
                        parcels.append(parcel)
                j += 1
            i += 1
            j = 0
        i = 0
    for parcel in set(parcels):
        output.write(newline(parcel))
    output.close()
    return output


def filldocxtemplate(templatefile, outputfile, owner=None):
    """function that copies docx template into an end of output file"""
    # select only paragraphs or table nodes
    template = Document(templatefile)
    get_text_from_doc(templatefile)
    try:
        output = Document(outputfile)
    except PackageNotFoundError:
        output = Document()
        output.save(outputfile)
    for child in template.element.body.xpath('w:p | w:tbl'):
        if isinstance(child, CT_P):
            paragraph = Paragraph(child, template)
            outpara = output.add_paragraph()
            s = paragraph.text
            hashtag = re.findall(r"#(\w+)#", s)
            logging.debug(f'Hashtagi: {hashtag}')
            for run in paragraph.runs:
                output_run = outpara.add_run(run.text)
                # Run's bold data
                for parcel in owner.parcels:
                    if parcel in output_run.text:
                        output_run.text = output_run.text.replace(parcel, '')
                        output_run.text = output_run.text.replace(' ,', '')
                        parcel_run = outpara.add_run(', ' + parcel)
                        parcel_run.bold = True
                        parcel_run.font.name = 'Times New Roman'

                output_run.bold = run.bold
                # Run's italic data
                output_run.italic = run.italic
                # Run's underline data
                output_run.underline = run.underline
                # Run's color data
                output_run.font.color.rgb = run.font.color.rgb
                # Run's font
                output_run.font.name = 'Times New Roman'
                output_run.font.size = run.font.size
                # Run's font data
                output_run.style.name = run.style.name
                # Paragraph's alignment data
            outpara.paragraph_format.line_spacing = 1.0
            outpara.paragraph_format.alignment = paragraph.paragraph_format.alignment
            outpara.paragraph_format.first_line_indent = paragraph.paragraph_format.first_line_indent
            outpara.paragraph_format.space_before = 5
            outpara.paragraph_format.space_after = 5

        elif isinstance(child, CT_Tbl):
            table = Table(child, template)
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        logging.debug(f'Text komórki: {paragraph.text}')
                        hashtag = re.findall(r"#(\w+)#", paragraph.text)
                        logging.debug(f'Hashtagi: {hashtag}')
                        if len(hashtag) == 0:
                            pass
                        else:
                            for hash in hashtag:
                                if hash == 'imie':
                                    inline = paragraph.runs
                                    for i in range(len(inline)):
                                        if 'imie' in inline[i].text:
                                            logging.debug(f'Zamieniam #imie# na {owner.name}')
                                            text = inline[i].text.replace('#imie#', owner.name)
                                            inline[i].text = text
                                elif hash == 'nazwisko':
                                    inline = paragraph.runs
                                    for i in range(len(inline)):
                                        if 'nazwisko' in inline[i].text:
                                            logging.debug(f'Zamieniam #nazwisko# na {owner.surname}')
                                            text = inline[i].text.replace('#nazwisko#', owner.surname)
                                            inline[i].text = text
                                elif hash == 'adres':
                                    inline = paragraph.runs
                                    for i in range(len(inline)):
                                        if 'adres' in inline[i].text:
                                            logging.debug(f'Zamieniam #adres# na {owner.address}')
                                            text = inline[i].text.replace('#adres#', owner.address)
                                            text = text.replace(', ', ',\n')
                                            inline[i].text = text
                                elif hash == 'godzina':
                                    inline = paragraph.runs
                                    for i in range(len(inline)):
                                        if 'godzina' in inline[i].text:
                                            try:
                                                text = inline[i].text.replace('#godzina#', owner.hour)
                                                inline[i].text = text
                                                logging.debug(f'Wprowadziłem godzinę dla {owner.name} {owner.surname}:'
                                                              f' {owner.hour} ')
                                            except:
                                                pass

            paragraph = output.add_paragraph()
            paragraph._p.addnext(table._tbl)

            paragraph.paragraph_format.first_line_indent = 1
            paragraph.paragraph_format.space_before = 1
            paragraph.paragraph_format.space_after = 1
            paragraph.paragraph_format.line_spacing = 1
    paragraph = output.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)
    output.save(outputfile)


def findowners(file, parcelspath):
    owners = []
    ownersobj = []
    parcel = ''
    doc = Document(file)
    table = doc.tables[0]
    i = 0
    j = 0
    parcelsfile = open(parcelspath, 'r')
    # find target parcels from file
    parcels = [line.replace('\n', '') for line in parcelsfile.readlines()]
    parcels = set(parcels)
    for row in table.rows:
        if i < 1:
            i += 1
            continue
        elif row.cells[0].text == 'Nr działki' or row.cells[0].text == '':
            i += 1
            continue
        # Find parcel name and chceck with target parcels
        elif row.cells[0].text.split('.')[-1] in parcels:
            for cell in row.cells:
                if j == 0:
                    parcel = cell.text.split('.')
                    parcel = parcel[-1]
                    print(parcel)
                if j == 6:
                    logging.debug(f'Text komórki: {cell.text}')
                    temp = cell.text.split('udział ')
                    temp.pop(0)
                    for udz in temp:
                        temp = udz.split(';')
                        temp[0] = temp[0].replace('Własność: ', '')
                        temp[1] = temp[1].replace('Własność: ', '')
                        logging.debug(f'{temp}')
                        adr = temp[1].replace('\n', '')
                        adr = adr.replace('Władanie: Użytkowanie', '')
                        if temp[0].split(' ')[1] == 'Małż.:':
                            logging.debug(f'Małżeństwo split: {temp}')
                            adr1 = temp[1].split('\n')[0]
                            adr2 = temp[2].split('\n')[0]
                            for item in temp:
                                item = item.replace('\nWłasność: ', '')
                                if '\n' in item:
                                    temp1 = item.split('\n')
                                    temp1.pop(0)
                                    temp1 = temp1[0].split(',')
                                    temp1 = temp1[0].split(' ')
                                    if len(temp1) == 3:
                                        logging.debug(f'Kasuje imię {temp1[1]}')
                                        temp1.pop(1)
                                        logging.debug(f'Współwlasciciel 2: {temp1}  adres: {adr2}')
                                        owners.append((temp1[0], temp1[1], adr2, parcel))
                                    else:
                                        logging.debug(f'Współwlasciciel 2: {temp1} adres: {adr2}')
                                        owners.append(
                                            (temp1[0], ' '.join(temp1[i] for i in range(1, len(temp1))), adr2, parcel))
                            temp = temp[0].split(',')
                            temp = temp[0].split(' ')
                            temp.pop(0)
                            temp.pop(0)
                            if len(temp) == 3 and temp[1][0].isupper():
                                logging.debug(f'Kasuje imię {temp[1]}')
                                temp.pop(1)
                                logging.debug(f'Wspolwlasciciel 1: {temp} adres: {adr1}')
                                owners.append((temp[0], temp[1], adr1, parcel))
                            else:
                                logging.debug(f'Wspolwlasciciel 1: {temp} adres: {adr1}')
                                owners.append((temp[0], ' '.join(temp[i] for i in range(1, len(temp))), adr1, parcel))

                        else:
                            temp = temp[0].split(',')
                            temp = temp[0].split(' ')
                            temp.pop(0)
                            if len(temp) == 3 and temp[1][0].isupper():
                                logging.debug(f'Kasuje imię {temp[1]}')
                                temp.pop(1)
                                logging.debug(f'Wlasciciel: {temp}')
                                owners.append((temp[0], temp[1], adr, parcel))
                            else:
                                logging.debug(f'Wlasciciel: {temp}')
                                owners.append((temp[0], ' '.join(temp[i] for i in range(1, len(temp))), adr, parcel))
                j += 1
            j = 0
        else:
            i += 1
            continue
    i = 0

    for owner in owners:
        alreadyin = False
        if len(ownersobj) == 0:
            o = Owner(owner[0], owner[1], owner[2], owner[3])
            ownersobj.append(o)
        else:
            for obj in ownersobj:
                if obj.name == owner[0] and obj.surname == owner[1] and obj.address == owner[2] and owner[
                    3] not in obj.parcels:
                    obj.addparcels(owner[3])
                    logging.debug(f'dodano parcele do obiektu: {obj.name} {obj.surname}')
                    alreadyin = True
            if not alreadyin:
                o = Owner(owner[0], owner[1], owner[2], owner[3])
                ownersobj.append(o)

    for o in ownersobj:
        print(f'Cześć nazywam się: {o.name} {o.surname}. Mieszkam przy ul.: {o.address},'
              f' jestem właścicielem działki: {o.parcels}')
    return ownersobj


def findkw(file, parcelspath):
    "find KW for specified parcels from .docx file"
    kw = []
    doc = Document(file)
    table = doc.tables[0]
    i = 0
    j = 0
    parcelsfile = open(parcelspath, 'r')
    # find target parcels from file
    parcels = [line.replace('\n', '') for line in parcelsfile.readlines()]
    parcels = set(parcels)
    for row in table.rows:
        if i < 1:
            i += 1
            continue
        elif row.cells[0].text == 'Nr działki' or row.cells[0].text == '':
            i += 1
            continue
        # Find parcel name and chceck with target parcels
        elif row.cells[0].text.split('.')[-1] in parcels:
            for cell in row.cells:
                if j == 0:
                    parcel = cell.text.split('.')
                    parcel = parcel[-1]
                    print(parcel)
                if j == 7:
                    logging.debug(f'Text komórki: {cell.text}')
                    text = cell.text.replace('\t', '')
                    logging.debug(f'{text}')
                    kw.append(calccontrolnumber(text))
                j += 1
            j = 0
    print(set(kw))
    return set(kw)


def calccontrolnumber(kw):
    """Calculate control number for specified KW"""
    dictionary = {'0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9, 'X': 10, 'A': 11,
                  'B': 12, 'C': 13, 'D': 14, 'E': 15, 'F': 16, 'G': 17, 'H': 18, 'I': 19, 'J': 20, 'K': 21, 'L': 22,
                  'M': 23, 'N': 24, 'O': 25, 'P': 26, 'R': 27, 'S': 28, 'T': 29, 'U': 30, 'W': 31, 'Y': 32, 'Z': 33}
    waga = '137137137137'
    sum = 0
    i = 0
    parts = []
    if len(kw.split('/')) == 3:
        return kw
    elif len(kw.split('/')) == 2:
        parts = kw.split('/')
        if len(kw.split('/')[1]) < 8:
            parts[1] = ('0' * (8 - len(parts[1]))) + parts[1]
        kw = parts[0] + parts[1]
    else:
        kw = kw.replace(' ', '')
        kw = kw.replace('KW', '')
        if len(kw) < 8:
            kw = ('0' * (8 - len(kw))) + kw
            kw = 'KR1P' + kw
    for s in kw:
        sum += dictionary[s.upper()] * int(waga[i])
        i += 1
    i = 0
    kw = kw[:4] + '/' + kw[4:] + '/' + str(sum % 10)
    return kw


def namestofile(owners, filename):
    """Create names doc from owners data"""
    i = 0
    doc = Document()
    table = doc.add_table(rows=1, cols=4)
    table.rows[0].cells[0].text = 'Lp.'
    table.rows[0].cells[1].text = 'Imię i Nazwisko'
    table.rows[0].cells[2].text = 'Adres'
    table.rows[0].cells[3].text = 'Numer przesyłki'
    for owner in owners:
        row = table.add_row()
        row.cells[0].text = str(i)
        row.cells[1].text = owner.fullname
        row.cells[2].text = owner.address
        i += 1
    doc.save(filename)


def isthesameowner(owner1, owner2):
    if owner1.name == owner2.name and owner1.surname == owner2.surname and owner1.address == owner2.address:
        return True
    else:
        return False


def removefromlist(list, determine):
    return list


def removeduplicates(file1, file2, outputfile):
    """Remove duplicates from parcel files"""
    parcels = []
    parcelsfile = open(file1, 'r')
    parcelsw = [line.replace('\n', '') for line in parcelsfile.readlines()]
    parcelsw = set(parcelsw)
    parcelsfile.close()
    parcelsfile = open(file2, 'r')
    parcelsu = [line.replace('\n', '') for line in parcelsfile.readlines()]
    parcelsu = set(parcelsu)
    parcelsfile.close()
    for parcel in parcelsw:
        if parcel in parcelsu:
            logging.debug(f'Wywalam: {parcel}')
            continue
        else:
            parcels.append(parcel)
    output = open(outputfile, 'w')
    newline = lambda x: x + '\n'
    for parcel in parcels:
        output.write(newline(parcel))
    output.close()


def createstickers(file, outfile):
    """Create stickers doc from names doc for each owner to put on letter"""
    out = Document()
    stickerstbl = out.add_table(rows=0, cols=1)
    section = out.sections[0]
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')[0]
    cols.set(qn('w:num'), '3')
    doc = Document(file)
    table = doc.tables[0]
    for row in table.rows[1:]:
        text = row.cells[1].text + '\n' + row.cells[2].text
        strow = stickerstbl.add_row()
        strow.cells[0].text = text
    out.save(outfile)


def parcelfinder(parcel, ownersfile):
    with open(ownersfile, 'r', newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter=',')
        for row in reader:
            if parcel in row[3]:
                print(' '.join([row[0], row[1], row[2]]))


def changehash(file, owner):
    doc = Document(file)
    get_text_from_doc(doc)
    for child in doc.element.body.xpath('w:p | w:tbl'):
        if isinstance(child, CT_P):
            paragraph = Paragraph(child, doc)
            s = paragraph.text
            hashtag = re.findall(r"#(\w+)#", s)
            logging.debug(f'Hashtagi: {hashtag}')
            logging.debug(f'Text komórki: {paragraph.text}')
            if len(hashtag) == 0:
                pass
            else:
                for hash in hashtag:
                    if hash == 'imie':
                        inline = paragraph.runs
                        for i in range(len(inline)):
                            if 'imie' in inline[i].text:
                                text = inline[i].text.replace('#imie#', owner.name)
                                inline[i].text = text
                    elif hash == 'nazwisko':
                        inline = paragraph.runs
                        for i in range(len(inline)):
                            if 'nazwisko' in inline[i].text:
                                text = inline[i].text.replace('#nazwisko#', owner.surname)
                                inline[i].text = text
                    elif hash == 'adres':
                        inline = paragraph.runs
                        for i in range(len(inline)):
                            if 'adres' in inline[i].text:
                                text = inline[i].text.replace('#adres#', owner.address)
                                text = text.replace(', ', ',\n')
                                inline[i].text = text
                    elif hash == 'godzina':
                        inline = paragraph.runs
                        for i in range(len(inline)):
                            if 'godzina' in inline[i].text:
                                try:
                                    text = inline[i].text.replace('#godzina#', owner.hour)
                                    inline[i].text = text
                                    logging.debug(f'Wprowadziłem godzinę dla {owner.name} {owner.surname}:'
                                                  f' {owner.hour} ')
                                except:
                                    pass
                    elif hash == 'data':
                        inline = paragraph.runs
                        for i in range(len(inline)):
                            if 'data' in inline[i].text:
                                try:
                                    text = inline[i].text.replace('#data#', owner.date)
                                    inline[i].text = text
                                    logging.debug(f'Wprowadziłem datę dla {owner.name} {owner.surname}:'
                                                  f' {owner.date} ')
                                except:
                                    pass


        elif isinstance(child, CT_Tbl):
            table = Table(child, doc)
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        logging.debug(f'Text komórki: {paragraph.text}')
                        hashtag = re.findall(r"#(\w+)#", paragraph.text)
                        logging.debug(f'Hashtagi: {hashtag}')
                        if len(hashtag) == 0:
                            pass
                        else:
                            for hash in hashtag:
                                if hash == 'imie':
                                    inline = paragraph.runs
                                    for i in range(len(inline)):
                                        if 'imie' in inline[i].text:
                                            text = inline[i].text.replace('#imie#', owner.name)
                                            inline[i].text = text
                                elif hash == 'nazwisko':
                                    inline = paragraph.runs
                                    for i in range(len(inline)):
                                        if 'nazwisko' in inline[i].text:
                                            text = inline[i].text.replace('#nazwisko#', owner.surname)
                                            inline[i].text = text
                                elif hash == 'adres':
                                    inline = paragraph.runs
                                    for i in range(len(inline)):
                                        if 'adres' in inline[i].text:
                                            text = inline[i].text.replace('#adres#', owner.address)
                                            text = text.replace(', ', ',\n')
                                            inline[i].text = text
                                elif hash == 'godzina':
                                    inline = paragraph.runs
                                    for i in range(len(inline)):
                                        if 'godzina' in inline[i].text:
                                            try:
                                                text = inline[i].text.replace('#godzina#', owner.hour)
                                                inline[i].text = text
                                                logging.debug(f'Wprowadziłem godzinę dla {owner.name} {owner.surname}:'
                                                              f' {owner.hour} ')
                                            except:
                                                pass
                                elif hash == 'data':
                                    inline = paragraph.runs
                                    for i in range(len(inline)):
                                        if 'data' in inline[i].text:
                                            try:
                                                text = inline[i].text.replace('#data#', owner.date)
                                                inline[i].text = text
                                                logging.debug(f'Wprowadziłem datę dla {owner.name} {owner.surname}:'
                                                              f' {owner.date} ')
                                            except:
                                                pass
    doc.save(file)


def kwtopdf(kw):
    openweb('https://ekw.ms.gov.pl/eukw_ogol/KsiegiWieczyste')
    if pylocate(os.getcwd() + '\\img\\' + 'kw.png') is not None:
        pyautogui.click(pylocate(os.getcwd() + '\\img\\' + 'kw.png'))
        pyautogui.write(kw.split('/')[0])
        pyautogui.move(40, 0)
        pyautogui.click()
        pyautogui.write(kw.split('/')[1])
        pyautogui.move(100, 0)
        pyautogui.click()
        pyautogui.write(kw.split('/')[2])
        pyautogui.click(pylocate(os.getcwd() + '\\img\\' + 'wyszukaj.png'))
        pyautogui.click(pylocate(os.getcwd() + '\\img\\' + 'przegladanie.png'))
        time.sleep(.5)
        pyautogui.hotkey('ctrlleft', 'p')
        pyautogui.click(pylocate(os.getcwd() + '\\img\\' + 'zapisz.png'))
        time.sleep(.5)
        pyautogui.write(kw.replace('/', '_') + '_1')
        pyautogui.click(pylocate(os.getcwd() + '\\img\\' + 'zapisz2.png'))
        time.sleep(2)
        pyautogui.click(pylocate(os.getcwd() + '\\img\\' + 'dzial2.png'))
        time.sleep(.5)
        pyautogui.hotkey('ctrlleft', 'p')
        pyautogui.click(pylocate(os.getcwd() + '\\img\\' + 'zapisz.png'))
        pyautogui.write(kw.replace('/', '_') + '_2')
        pyautogui.click(pylocate(os.getcwd() + '\\img\\' + 'zapisz2.png'))
        time.sleep(.5)
        pyautogui.hotkey('ctrlleft', 'w')


def pdfmerge(folder):
    for subdir, dirs, files in os.walk(folder):
        for file in files:
            file = file.replace('.pdf', '')
            print(file)
            if file.split('_')[-1] == '1':
                for f2 in files:
                    if f2.split('.')[-1] != 'pdf':
                        continue
                    f2 = f2.replace('.pdf', '')
                    logging.debug(f'{f2.split("_")[1]} vs {file.split("_")[1]} and {file.split("_")[-1]}')
                    if f2.split('_')[1] == file.split('_')[1] and f2.split('_')[-1] != '1':
                        logging.debug('hej')
                        merger = PdfFileMerger()
                        input1 = open(folder + '/' + file + '.pdf', 'rb')
                        input2 = open(folder + '/' + f2 + '.pdf', 'rb')
                        merger.append(input1, pages=(0, 1))
                        merger.append(input2)
                        name = file[:-2] + '.pdf'
                        output = open(name, "wb")
                        merger.write(output.name)


def openweb(url):
    """open desired url in basic browser"""
    web.open_new(url)


def pylocate(img):
    logging.debug(f'{img}')
    i = 0
    s = None
    while i < 5:
        try:
            logging.debug(f'próbuję znaleźć obraz')
            s = pyautogui.locateOnScreen(img, confidence=0.9)
            logging.debug(f'Znalazłem: {s}')
            if s is not None:
                return s
        except:
            time.sleep(3)
        i += 1
    return s


def getfeaturesfromgml(gmlfile, feature):
    content = gmlfile.read()
    contentlist = content.split('<' + feature)[1:]
    for i, item in enumerate(contentlist):
        contentlist[i] = item.split('</gml:featureMember>')[0]
    """for item in contentlist:
        logging.debug(item)
        logging.debug('\n_____________________________________')"""
    return contentlist


def getcontentfromtags(text, tag):
    # Function to grab everything in between two tags, nested tags included
    content = text.split(tag)[1]  # grab text starting from tag
    content = re.sub(r'^.*?>', '', content)  # delete full tag
    content = content.split('</')[:-1]  # delete exit tag
    content = '</'.join(content)  # join every nested tags
    return content


def getinfofromtags(text, tag):
    # Funtion to get meta data from single tags (info) --> <tag/>
    tags = text.split(tag)[1:-1]
    print(tags)
    info = {}
    for item in tags:
        content = item.split('>')[0]
        content = content.split(' ')[1:]
        for meta in content:
            key = meta.split('=')[0]
            data = meta.split('=')[1]
            data = data.replace('"', '')
            info[key] = data
    print(info)
    return info


class Owner:
    def __init__(self, name, surname, address, parcel, hour=None, date=None, source=None):
        self.name = name
        self.surname = surname
        self.address = address
        self.hour = hour
        self.date = date
        self.source = source
        self.parcel = parcel
        self.parcels = []
        self.addparcels(parcel)
        self.fullname = self.name + ' ' + self.surname

    def addparcels(self, parcel):
        self.parcels.append(parcel)

    def zawiadomienie(self):
        pass


class Parcel:
    def __init__(self, id, gmlid, number, points, area, owners=None, kw=None, calc_area=None):
        self.id = id
        self.gmlid = gmlid
        self.number = number
        self.owners = owners
        self.points = points
        self.kw = kw
        self.area = area
        self.calc_area = calc_area

    def calculate_area(self):
        pointlist = []
        for pointobj in self.points:
            pointlist.append((pointobj.x, pointobj.y))
        pgon = Polygon(pointlist)
        logging.debug(f'Parcel calculated Area: {round(pgon.area)}')
        return round(pgon.area)


class Point:
    def __init__(self, id, gmlid, number, x, y, zrd=None, bpp=None, stb=None, rzg=None, operat=None, sporna=None):
        self.id = id
        self.gmlid = gmlid
        self.number = number
        self.x = x
        self.y = y
        self.zrd = zrd
        self.bpp = bpp
        self.stb = stb
        self.rzg = rzg
        self.operat = operat
        self.sporna = sporna


class Navbar(tk.Frame): ...


class Toolbar(tk.Frame): ...


class Statusbar(tk.Frame): ...


class Main(tk.Frame): ...


class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.statusbar = Statusbar(self, ...)
        self.toolbar = Toolbar(self, ...)
        self.navbar = Navbar(self, ...)
        self.main = Main(self, ...)

        self.statusbar.pack(side="bottom", fill="x")
        self.toolbar.pack(side="top", fill="x")
        self.navbar.pack(side="left", fill="y")
        self.main.pack(side="right", fill="both", expand=True)


def main():
    gmlfile = open("Zbiór danych GML.gml", encoding='utf-8')
    pointlist = getfeaturesfromgml(gmlfile, 'egb:EGB_PunktGraniczny')
    pointsobj = []
    for point in pointlist:
        id = getcontentfromtags(point, 'egb:idPunktu')
        number = id.split('.')[-1]
        gmlid = getcontentfromtags(point, 'bt:lokalnyId')
        coordinates = getcontentfromtags(point, 'gml:pos')
        x = float(coordinates.split(' ')[0])
        y = float(coordinates.split(' ')[1])
        try:
            zrd = int(getcontentfromtags(point, 'egb:zrodloDanychZRD'))
        except ValueError:
            zrd = 'brak'
        try:
            bpp = int(getcontentfromtags(point, 'egb:bladPolozeniaWzgledemOsnowy'))
        except ValueError:
            bpp = 'brak'
        try:
            stb = int(getcontentfromtags(point, 'egb:kodStabilizacji'))
        except ValueError:
            stb = 'brak'
        try:
            rzg = int(getcontentfromtags(point, 'egb:kodRzeduGranicy'))
        except ValueError:
            rzg = 'brak'
        new_point = Point(id, gmlid, number, x, y, zrd, bpp, stb, rzg)
        pointsobj.append(new_point)
        logging.debug(f'Utworzyłem nowy punkt o numerze: {new_point.number} id {new_point.id}, '
                      f'gmlid {new_point.gmlid}, wsp: {new_point.x}, {new_point.y}, {new_point.zrd} {new_point.bpp}'
                      f' {new_point.stb} {new_point.rzg}')
    gmlfile.seek(0)
    parcellist = getfeaturesfromgml(gmlfile, 'egb:EGB_DzialkaEwidencyjna')
    for parcel in parcellist:
        id = getcontentfromtags(parcel, 'idDzialki')
        number = id.split('.')[-1]
        gmlid = getcontentfromtags(parcel, 'bt:lokalnyId')
        area = float(getcontentfromtags(parcel, 'egb:powierzchniaEwidencyjna'))
        poslist = getcontentfromtags(parcel, 'gml:posList')
        poslist = poslist.split(' ')
        pointslist = []
        points = []
        for i, coordinate in enumerate(poslist):
            if i%2 == 0:
                pointslist.append((float(poslist[i]), float(poslist[i+1]))) # Only add correct points
        pointslist.pop() #remove last point because it's the same as first
        for point in pointslist:
            for pt in pointsobj:
                if pt.x == point[0] and pt.y == point[1]:
                    points.append(pt)

        new_parcel = Parcel(id, gmlid, number, points, area)
        new_parcel.calc_area = new_parcel.calculate_area()
        logging.debug(f'Utworzyłem nową działkę o numerze: {new_parcel.number} id {new_parcel.id}, '
                      f'gmlid {new_parcel.gmlid}, powierzchni {new_parcel.area}, ')

    # getinfofromtags(parcellist[0], 'gml:Point')
    """point1 = Point(1,1,0,0)
    point2 = Point(2, 2, 13, 0)
    point3 = Point(3, 3, 10, 12)
    point4 = Point(4, 4, 110, 10)
    points = [point1,point2,point3,point4]
    parc = Parcel(1,[],points,101)
    parc.calculate_area()"""
    """root = tk.Tk()
    MainApplication(root).pack(side="top", fill="both", expand=True)
    root.mainloop()"""
    # toplevel = tk.Toplevel(root)

    # create a toplevel menu
    """  menubar = tk.Menu(toplevel)
    menubar.add_command(label="Podział")
    menubar.add_command(label="Quit!", command=root.quit)
    # display the menu
    toplevel.config(menu=menubar)
    root.mainloop()"""
    """ Main program """
    """x = 400
    y = 200
    pyautogui.moveTo(x, y)"""
    # check_project_data()
    # write_report()
    # ConvertRtfToDocx('C:\\Users\\Jurek\\Documents\\Kuba\\Python\\OperatOR\\docs','info_o_materiałach1.rtf')
    # createparcelfilefromdoc(os.getcwd() + '\\docs\\protokol_wyznaczenia_granic.docx', 'parcels.txt')
    """o = Owner('Jakub', 'Kowwalczyk', 'KRakowska 23, 31-102 KRaków', '512')
        filldocxtemplate(os.getcwd() + '\\docs\\Zawiad o wyznaczeniu granic.docx', 'wyzn_granic.docx', o)"""
    # removeduplicates('parcelsw.txt', 'parcelsu.txt', 'parcelsw.txt')
    # ownersfile = open_file()
    # ownersu = findowners(ownersfile, 'parcelsu.txt')
    # ownersw = findowners(ownersfile, 'parcelsw.txt')
    # owners = ownersu + ownersw
    """with open('owners.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter=',')
        for owner in ownersu:
            writer.writerow([owner.name, owner.surname, owner.address, owner.parcels])"""
    # parcelfinder('581/4', ownersfile='owners.csv')
    # namestofile(owners, 'nazwiska i adresy.docx')
    # createstickers('nazwiska i adresy.docx', 'naklejki.docx')
    # file = open_file()
    # find_kerg(file.name)
    # kwlist = findkw('C:\\Users\\Jurek\\Dysk Google\\GEO\\Bibice_Zbożowa\\PODGiK\\właściciele.docx', 'parcelsu.txt')
    # for kw in kwlist:
    #    kwtopdf(kw)
    # kwtopdf('KR1P/00516204/5')
    # pdfmerge(open_folder()) #merge first page of _1 file and _2 file

    """for owner in ownersu:
        filldocxtemplate(os.getcwd() + '\\docs\\Zawiadomienie ustalenie.docx', 'ust_granicKwiatowa.docx', owner)
    for owner in ownersw:
        filldocxtemplate(os.getcwd() + '\\docs\\Zawiadomienie ustalenie.docx', 'ust_granicKwiatowa.docx', owner)"""
    return 0


if __name__ == "__main__":
    main()
