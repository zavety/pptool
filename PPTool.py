#!usr/bin/env python
#coding=utf-8

import os
import sys
from win32com.client import Dispatch

def pptool(pptPath):
    p = Dispatch("PowerPoint.Application")
    p.Visible = True
    try:
        path = pptPath.decode("utf-8")
        work = p.Presentations.Open(pptPath)
        pageCount = work.Slides.Count
        pyPath = os.sys.path[0]
        themePath = pyPath + r'\Origin.thmx'
        work.ApplyTheme(themePath)
        work.Save()
        return True
    except Exception as e:
        print themePath
        print e
        return False
    finally:
        p.Quit()


if __name__ == '__main__':
    pptool(str(sys.argv[1]))
