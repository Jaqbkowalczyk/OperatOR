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


def copydocxtemplate(templatefile, outputfile):
    """function that copies docx template into an end of output file"""
    # select only paragraphs or table nodes
    template = Document(templatefile)
    get_text_from_doc(templatefile)
    output = Document(outputfile)
    for child in template.element.body.xpath('w:p | w:tbl'):
        if isinstance(child, CT_P):
            paragraph = Paragraph(child, template)
            outpara = output.add_paragraph()
            outpara._p.addnext(paragraph._p)
        elif isinstance(child, CT_Tbl):
            table = Table(child, template)
            paragraph = output.add_paragraph()
            paragraph._p.addnext(table._tbl)
    output.save(outputfile)


def findowners(file, parcelspath):
    doc = Document(file)
    table = doc.tables[0]
    i = 0
    j = 0
    parcelsfile = open(parcelspath, 'r')
    # find target parcels from file
    parcels = [line.replace('\n', '') for line in parcelsfile.readlines()]
    print(parcels)
    for row in table.rows:
        if i < 1:
            i += 1
            continue
        elif row.cells[0].text == 'Nr działki' or row.cells[0].text == '':
            i += 1
            continue
        # Find parcel name and chceck with target parcels
        elif row.cells[0].text.split('.')[-1] in parcels:
            logging.debug('hej')
            for cell in row.cells:
                if j == 0:
                    parcel = cell.text.split('.')
                    print(parcel[-1])
                j += 1
            j = 0
        else:
            i += 1
            continue
    i = 0


class Zawiadomienie:
    def __init__(self, name, surname, address, hour, date, type):
        self.name = name
        self.surname = surname
        self.address = address
        self.hour = hour
        self.date = date
        self.type = type


def main():
    """ Main program """
    """x = 400
    y = 200
    pyautogui.moveTo(x, y)"""
    #check_project_data()
    #write_report()
    #ConvertRtfToDocx('C:\\Users\\Jurek\\Documents\\Kuba\\Python\\OperatOR\\docs','info_o_materiałach1.rtf')
    #print(createparcelfile('C:\\Users\\Jurek\\Dysk Google\\GEO\\Bibice_Zbożowa\\Wyznaczenie\\protokol_wyznaczenia_granic.docx'))
    #copydocxtemplate('C:\\Users\\Jurek\\Dysk Google\\GEO\\Bibice_Zbożowa\\Wyznaczenie\\protokol_wyznaczenia_granic.docx', 'out.docx')
    findowners('C:\\Users\\Jurek\\Dysk Google\\GEO\\Bibice_Zbożowa\\PODGiK\\właściciele.docx', 'parcels.txt')
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

