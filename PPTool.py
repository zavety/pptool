#!usr/bin/env python
#coding=utf-8

import os
import sys
import time
import win32gui
import multiprocessing
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

def closerrwindow(*fileBaseNameChars):
    time.sleep(5)
    titles = [u'密码', u'显示修复', u'Microsoft Office PowerPoint']
    #使用多线程传递参数str类型变成了char[]元组，需要重新合并,我也不知道为什么
    title = ''
    for char in fileBaseNameChars:
        title += char
    title = u'文件转换 - ' + str(title).decode('utf-8')
    titles.append(title)
    for t in titles:
        handle = win32gui.FindWindow(Node, t)
        win32gui.SendMessage(handle, 0x0010, 0, '0')
    
if __name__ == '__main__':
    fileBaseName = str(os.path.basename(sys.argv[1]))
    p.multiprocessing.Process(target = closerrwindow, args = (fileBaseName)）
    p.start()
    pptool(str(sys.argv[1]))
    p.join()
