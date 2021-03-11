#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re

import pyautogui
from docx import Document

pyautogui.FAILSAFE = True


def checkprojectdata():
    """Function to check for variables used in all future functions"""
    # 1. Select project folder
    # 2. Find data from PODGiK (.gml file)
    # 3. Search through .gml file to find JEDNOSTKAREJESTROWA and OBREB values
    # 4. Find KERG number
    pass


def writereport():
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
    checkprojectdata()
    writereport()
    return 0


if __name__ == "__main__":
    main()
