#!usr/bin/env python
#coding=utf-8

#Create by Qiang.D

import unittest
import PPTool

class mytest(unittest.TestCase):

  def setUp(self):
    pass
    
  def tearDown(self):
    pass
  
  def testPPtool(self):
    self.assertEqual(PPTool.pptool(r'D:\test.ppt'), True)
    
if __name__ == '__main__':
  unittest.main()
