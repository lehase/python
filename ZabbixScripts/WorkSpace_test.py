#!/usr/bin/python
# -*- coding: utf-8 -*-

import datetime
import subprocess
import time

from pyvirtualdisplay import Display
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

LogFile = '/usr/lib/zabbix/externalscripts/WorkSpace_test.log'

# profile = webdriver.FirefoxProfile('C:/Ser/1')
profile = webdriver.FirefoxProfile('/usr/lib/zabbix/externalscripts/ffprofile')
profile.add_extension('/usr/lib/zabbix/externalscripts/autoauth@efinke.com.xpi')
profile.set_preference("network.http.phishy-userpass-length", 255)
profile.set_preference("network.negotiate-auth.trusteduris", "cherkizovsky.net")
display = Display(visible=0, size=(800, 600))
display.start()
browser = webdriver.Firefox(profile)
browser.set_page_load_timeout(40)
Log = open(LogFile, 'a')


def Test_1():
    try:
        link = 'http://mow03-opt01.cherkizovsky.net/otcs/cs.exe'
        StartTime = time.time()
        browser.get(link)
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH,
                                                                         '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/form/div/div/div/table/tbody/tr/td/div/div[1]/table/tbody/tr[1]/td/div[1]/table/tbody/tr[3]/td/table/tbody/tr/td/div/div[1]/input')))
        EndTime = time.time()
        Result = round(EndTime - StartTime, 2)
        return Result
    except BaseException as exception:
        Date = (datetime.datetime.today()).strftime('%d.%m.%Y %H:%M:%S')
        Log.write(str(Date) + '\n')
        Log.write('\n')
        return int(-1)


def Test_2():
    try:
        link = 'http://mow03-opt02.cherkizovsky.net/otcs/cs.exe'
        browser.get(link)
        browser.find_element_by_name("otds_username").send_keys("CHERKIZOVSKY\TECH_ZAB_VM")
        browser.find_element_by_name("otds_password").send_keys("7Ns4vn44B447")
        StartTime = time.time()
        browser.find_element_by_class_name("button").submit()
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH,
                                                                         '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/form/div/div/div/table/tbody/tr/td/div/div[1]/table/tbody/tr[1]/td/div[1]/table/tbody/tr[3]/td/table/tbody/tr/td/div/div[1]/input')))
        EndTime = time.time()
        Result = round(EndTime - StartTime, 2)
        return Result
    except BaseException as exception:
        Date = (datetime.datetime.today()).strftime('%d.%m.%Y %H:%M:%S')
        Log.write(str(Date) + '\n')
        Log.write('\n')
        return int(-1)


Test1 = Test_1()
subprocess.call(['zabbix_sender -z 192.168.20.109 -p 10051 -s "MOW03-OPT01" -k web_workplace_load -o ' + str(Test1)],
                shell=True)
Test2 = Test_2()
subprocess.call(['zabbix_sender -z 192.168.20.109 -p 10051 -s "MOW03-OPT02" -k web_workplace_load -o ' + str(Test2)],
                shell=True)

browser.quit()
Log.close()
display.stop()
