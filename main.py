#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import re
import logging
import pyautogui
import win32com.client as win32
from docx import Document
import tkinter as tk
from tkinter.filedialog import askdirectory, askopenfile
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.opc.exceptions import PackageNotFoundError
from docx.shared import Pt

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
pyautogui.FAILSAFE = True
JEDNOSTKAREJESTROWA = ''
KERG = ''
FOLDER = ''


def find_kerg(filename):
    pass


def set_kerg(kerg_value: str):
    global KERG
    KERG = kerg_value
    logging.debug(f'KERG został ustawiony na: {KERG}')


def set_folder(folder_value: str):
    global FOLDER
    FOLDER = folder_value
    logging.debug(f'FOLDER został ustawiony na: {FOLDER}')


def open_folder():
    tk.Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askdirectory()  # show an "Open" dialog box and return the path to the selected folder
    return filename


def ConvertRtfToDocx(rootDir, file):
    word = win32.Dispatch("Word.Application")
    wdFormatDocumentDefault = 16
    wdHeaderFooterPrimary = 1
    doc = word.Documents.Open(rootDir + "\\" + file)
    for pic in doc.InlineShapes:
        pic.LinkFormat.SavePictureWithDocument = True
    for hPic in doc.sections(1).headers(wdHeaderFooterPrimary).Range.InlineShapes:
        hPic.LinkFormat.SavePictureWithDocument = True
    doc.SaveAs(str(rootDir + "\\" + file + ".docx"), FileFormat=wdFormatDocumentDefault)
    doc.Close()
    word.Quit()


def get_text_from_doc(filename):
    doc = Document(filename)
    fullText = []
    i = 0
    j = 0
    print(len(doc.sections))
    for para in doc.paragraphs:
        fullText.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                print(cell.text)
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


def check_project_data():   # todo find a way to extract text from .rtf file
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


def createparcelfile(file):
    """This function creates parcel text file from .docx table with parcels"""
    doc = Document(file)
    i = 0
    j = 0
    table = doc.tables[0]
    output = open(os.getcwd() + '\\parcels.txt', 'w')
    print(os.getcwd() + '\\parcels.txt')
    for row in table.rows:
        if i < 2:
            i += 1
            continue
        for cell in row.cells:
            if j == 1:
                parcels = cell.text.split(',')
                newline = lambda x: x + '\n'
                for parcel in parcels:
                    parcel = parcel.replace(' ', '')
                    output.write(newline(parcel))
            j += 1
        i += 1
        j = 0
    i = 0
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
            print(hashtag)
            for run in paragraph.runs:
                output_run = outpara.add_run(run.text)
                # Run's bold data
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
            outpara.paragraph_format.space_before = paragraph.paragraph_format.space_before
            outpara.paragraph_format.space_after = paragraph.paragraph_format.space_after


        elif isinstance(child, CT_Tbl):
            table = Table(child, template)
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        s = paragraph.text
                        logging.debug(f'Text komórki: {s}')
                        hashtag = re.findall(r"#(\w+)#", s)
                        print(hashtag)
                        if len(hashtag) == 0:
                            pass
                        else:
                            for hash in hashtag:
                                if hash == 'imie':
                                    s.replace('#imie#', 'KOKOSY')
                                    print(s)
                                elif hash == 'imie':
                                    s.replace('#nazwisko#', owner.surname)
                                elif hash == 'adres':
                                    s.replace('#adres#', owner.address)
                        paragraph.text = s
            paragraph = output.add_paragraph()
            print(table._tbl)
            paragraph._p.addnext(table._tbl)

            paragraph.paragraph_format.first_line_indent = 1
            paragraph.paragraph_format.space_before = 1
            paragraph.paragraph_format.space_after = 1
            paragraph.paragraph_format.line_spacing = 1

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
            #todo find addresses and create owners, then add to list
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
                                        owners.append((temp1[0], ' '.join(temp1[i]for i in range(1, len(temp1))), adr2, parcel))
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
                                owners.append((temp[0], ' '.join(temp[i]for i in range(1, len(temp))), adr1, parcel))

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
                                owners.append((temp[0], ' '.join(temp[i]for i in range(1, len(temp))), adr, parcel))
                j += 1
            j = 0
        else:
            i += 1
            continue
    i = 0
    print(set(owners))

    for owner in owners:
        alreadyin = False
        if len(ownersobj) == 0:
            o = Owner(owner[0], owner[1], owner[2], owner[3])
            ownersobj.append(o)
        else:
            for obj in ownersobj:
                if obj.name == owner[0] and obj.surname == owner[1] and obj.address == owner[2] and owner[3] not in obj.parcels:
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

    def addparcels(self, parcel):
        self.parcels.append(parcel)

    def zawiadomienie(self):
        pass


def main():
    """ Main program """
    """x = 400
    y = 200
    pyautogui.moveTo(x, y)"""
    #check_project_data()
    #write_report()
    #ConvertRtfToDocx('C:\\Users\\Jurek\\Documents\\Kuba\\Python\\OperatOR\\docs','info_o_materiałach1.rtf')
    #print(createparcelfile('C:\\Users\\Jurek\\Dysk Google\\GEO\\Bibice_Zbożowa\\Wyznaczenie\\protokol_wyznaczenia_granic.docx'))

    #owners = findowners('C:\\Users\\Jurek\\Dysk Google\\GEO\\Bibice_Zbożowa\\PODGiK\\właściciele.docx', 'parcels.txt')
    """
    for owner in owners:
        filldocxtemplate(os.getcwd() + '\\docs\\Zawiadomienie o przyj granic.docx', 'out.docx', owner)
    """
    o = Owner('Marek','Zbylski','Konieczna 12', '255/12')
    filldocxtemplate(os.getcwd() + '\\docs\\Zawiadomienie o przyj granic.docx', 'out.docx', o)
    return 0


if __name__ == "__main__":
    """root = tk.Tk()
    root.withdraw()

    toplevel = tk.Toplevel(root)

    # create a toplevel menu
    menubar = tk.Menu(toplevel)
    menubar.add_command(label="Hello!")
    menubar.add_command(label="Quit!", command=root.quit)
    # display the menu
    toplevel.config(menu=menubar)"""
    main()
    #root.mainloop()

