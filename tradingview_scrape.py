# -*- coding: utf-8 -*-
"""
Created on Thu Nov  8 18:06:58 2018

@author: naint

windows err 0, corrected wildcard with regex. example BTCUSD -> BTCUSD.*
"""

from selenium import webdriver
import re
import win32com.client, win32gui
import csv
import time


'''
find_element_by_id
find_element_by_name
find_element_by_xpath
find_element_by_link_text
find_element_by_partial_link_text
find_element_by_tag_name
find_element_by_class_name
find_element_by_css_selector
'''

class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name=None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)

drive = webdriver.Chrome()
drive.implicitly_wait(30)
drive.get("https://www.tradingview.com/chart/nN53lJ0J/")

w = WindowMgr()
w.find_window_wildcard("BTCUSD.*")
w.set_foreground()

shell = win32com.client.Dispatch("WScript.Shell")

mylist = []

def ttr(): #0.5 sleep doesn't give err. I guess because of it gives time to load js
    w.find_window_wildcard("BTCUSD")
    w.set_foreground()
    shell.SendKeys('{right}')
    #time.sleep(0.5)

def getd():    
    cOpening = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div[2]/table/tbody/tr[1]/td[2]/div/div[3]/div[1]/div/span[1]/span[2]').text
    cHigh = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div[2]/table/tbody/tr[1]/td[2]/div/div[3]/div[1]/div/span[2]/span[2]').text
    cLow = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div[2]/table/tbody/tr[1]/td[2]/div/div[3]/div[1]/div/span[3]/span[2]').text
    cClose = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div[2]/table/tbody/tr[1]/td[2]/div/div[3]/div[1]/div/span[4]/span[2]').text
    pChange = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div/table/tbody/tr[1]/td[2]/div/div[3]/div[1]/div/span[5]/span[2]').text
    iVolume = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div/table/tbody/tr[1]/td[2]/div/div[3]/div[3]/div/span[1]/span').text
    iStochiB = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div/table/tbody/tr[5]/td[2]/div/div[3]/div/div/span[1]/span').text
    iStochiR= drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div/table/tbody/tr[5]/td[2]/div/div[3]/div/div/span[2]/span').text
    iCCI = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div/table/tbody/tr[7]/td[2]/div/div[3]/div/div/span[1]/span').text
    icciobvRG = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/div[3]/div/div/span[1]/span').text
    icciobvO = drive.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/div[3]/div/div/span[2]/span').text
    d = cOpening, cHigh, cLow, cClose, pChange, iVolume, iStochiB, iStochiR, iCCI, icciobvRG, icciobvO
    global mylist
    mylist.append(d)
    ttr()
    

#for i in range(365):
#    try:
#        getd()
#    except:
#        pass

#
with open('btc.csv', 'w', encoding='utf8') as myfile:
    wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
    wr.writerows(mylist)