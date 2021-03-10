#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pyautogui
import re

def piszsprawozdanie():
    """Funkcja kreowania Sprawozdania Technicznego"""
    s = "I love #stackoverflow# because #people# are very #helpful# #helpful#"
    hash = re.findall(r"#(\w+)#", s)    # znajdź wszystkie hashtagi w szablonie
    print(set(hash))

def main():
    """ Main program """
    x=400
    y=200
    num_seconds = .1
    pyautogui.moveTo(x, y)
    print(pyautogui.size())
    piszsprawozdanie()
    return 0

if __name__ == "__main__":
    main()