#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pyautogui
import re
pyautogui.FAILSAFE = True

class Mdcp:
    pass


class Podzial:
    pass


class Inwentaryzacja:
    pass


def piszsprawozdanie():
    """Funkcja kreowania Sprawozdania Technicznego"""
    s = "I love #stackoverflow# because #people# are very #helpful# #helpful#"
    hashtag = re.findall(r"#(\w+)#", s)    # znajdź wszystkie hashtagi w szablonie
    print(set(hashtag))


def main():
    """ Main program """
    x = 400
    y = 200
    pyautogui.moveTo(x, y)
    print(pyautogui.size())
    piszsprawozdanie()
    return 0


if __name__ == "__main__":
    main()
