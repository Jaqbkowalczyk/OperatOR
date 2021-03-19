#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import logging
import pyautogui
from docx import Document
import tkinter as tk
from tkinter.filedialog import askdirectory, askopenfile

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



def check_project_data():
    """Function to check for variables used in all future functions"""
    # 1. Select project folder
    folder = open_folder()
    set_folder(folder)
    # 2. Find data from PODGiK (.gml file, info o materialach)
    # 3. Search through .gml file to find JEDNOSTKAREJESTROWA and OBREB values
    # 4. Find KERG number
    
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


def main():
    """ Main program """
    x = 400
    y = 200
    pyautogui.moveTo(x, y)
    check_project_data()
    write_report()
    set_kerg('666.2250.2021')
    return 0


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    toplevel = tk.Toplevel(root)

    # create a toplevel menu
    menubar = tk.Menu(toplevel)
    menubar.add_command(label="Hello!")
    menubar.add_command(label="Quit!", command=root.quit)
    # display the menu
    toplevel.config(menu=menubar)
    main()
    root.mainloop()

