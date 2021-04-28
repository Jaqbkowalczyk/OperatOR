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
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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
    parcels =[]
    i = 0
    j = 0
    newline = lambda x: x + '\n'
    output = open(os.getcwd() + '\\parcels.txt', 'w')
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
                                            logging.debug(f'hej: {inline[i].text}')
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


def namestofile(owners, filename):
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
    """Create stickers doc for each owner to put on letter"""
    out = Document()
    stickerstbl = out.add_table(rows=1, cols=1)
    section = out.sections[0]
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')[0]
    cols.set(qn('w:num'), '3')
    doc = Document(file)
    table = doc.tables[0]
    text = ''
    for row in table.rows[1:]:
        for cell in row.cells[1:2]:
            text += '\n' + cell.text
        stickerstbl.add_row()
        row.cells[0].text = text
        text = ''
    out.save(outfile)



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


def main():
    """ Main program """
    """x = 400
    y = 200
    pyautogui.moveTo(x, y)"""
    #check_project_data()
    #write_report()
    #ConvertRtfToDocx('C:\\Users\\Jurek\\Documents\\Kuba\\Python\\OperatOR\\docs','info_o_materiałach1.rtf')
    #print(createparcelfile(os.getcwd() + '\\docs\\protokol_wyznaczenia_granic.docx'))
    """o = Owner('Jakub', 'Kowwalczyk', 'KRakowska 23, 31-102 KRaków', '512')
        filldocxtemplate(os.getcwd() + '\\docs\\Zawiad o wyznaczeniu granic.docx', 'wyzn_granic.docx', o)"""
    """removeduplicates('parcelsw.txt', 'parcelsu.txt', 'parcels.txt')
    ownersu = findowners('C:\\Users\\Jurek\\Dysk Google\\GEO\\Bibice_Zbożowa\\PODGiK\\właściciele.docx', 'parcelsu.txt')
    ownersw = findowners('C:\\Users\\Jurek\\Dysk Google\\GEO\\Bibice_Zbożowa\\PODGiK\\właściciele.docx', 'parcels.txt')
    owners = ownersu + ownersw
    namestofile(owners, 'nazwiska i adresy.docx')"""
    createstickers('nazwiska i adresy.docx', 'naklejki.docx')
    """for owner in owners:
        filldocxtemplate(os.getcwd() + '\\docs\\Zawiadomienie wyznaczenie.docx', 'wyzn_granic.docx', owner)"""
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

