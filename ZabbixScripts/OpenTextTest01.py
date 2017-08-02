#!/usr/bin/python
# -*- coding: utf-8 -*-


from selenium import webdriver
import time, datetime, subprocess, sys, re, os
from pyvirtualdisplay import Display
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common import exceptions

Debug = False

LogFile = '/usr/lib/zabbix/externalscripts/Log_OpenTextTest01/' + str(
    (datetime.datetime.today()).strftime('%d_%m_%Y_%H_%M')) + '.log'
HTML_File = '/usr/lib/zabbix/externalscripts/Log_OpenTextTest01/' + str(
    (datetime.datetime.today()).strftime('%d_%m_%Y_%H_%M')) + '.html'
profile = webdriver.FirefoxProfile('/usr/lib/zabbix/externalscripts/ffprofile')
profile.add_extension('/usr/lib/zabbix/externalscripts/autoauth@efinke.com.xpi')
profile.set_preference("network.http.phishy-userpass-length", 255)
profile.set_preference("network.negotiate-auth.trusteduris", "cherkizovsky.net")
if Debug:
    display = Display(visible=1, size=(800, 600))
else:
    display = Display(visible=0, size=(800, 600))

display.start()
Log = open(LogFile, 'a')

browser = webdriver.Firefox(profile)
browser.set_page_load_timeout(40)

HTML = open(HTML_File, 'a')


def Test_11():
    try:
        global Log
        Date = (datetime.datetime.today()).strftime('%d.%m.%Y %H:%M:%S')
        URL = 'http://mow03-opt01/otcs/llisapi.dll?func=ll&objId=51692384&objAction=RunReport'
        if Debug:
            print 'Запуск теста 1.1			' + str(Date)
            print 'Загружаем URL'
            browser.get(URL)
            print 'Поиск кнопки "Поиск"'
            WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@value="Поиск"]')))
            print 'Поиск кнопки "Очистить"'
            browser.find_element_by_xpath('//*[@value="Очистить"]').click()
            time.sleep(0.2)
            print 'Выбор "Исочник: Оба"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr/td[4]/input[3]').click()
            print 'Добавление поля "Штрихкод"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[2]').click()
            print 'Заполнение поля "Штрихкод"'
            browser.find_element_by_xpath('//input[@maxlength="20"]').send_keys("00001002000000107292")
            time.sleep(0.2)
            print 'Запуск таймера'
            StartTime = time.time()
            print 'Нажатие кнопки "Поиск"'
            browser.find_element_by_xpath('//*[@onclick="doSearch(true);"]').click()
            print 'Отчет готов'
            print 'Останавливаем таймер'
            EndTime = time.time()
            Result = round(EndTime - StartTime, 2)
            print 'Тест 1.1 завершен. Время выполнения: ' + str(Result)
            return Result
        else:
            Log.write('Запуск теста 1.1			' + str(Date) + '\n')
            Log.write('Загружаем URL' + '\n')
            browser.get(URL)
            Log.write('Поиск кнопки "Поиск"' + '\n')
            WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@value="Поиск"]')))
            Log.write('Поиск кнопки "Очистить"' + '\n')
            browser.find_element_by_xpath('//*[@value="Очистить"]').click()
            time.sleep(0.2)
            Log.write('Выбор "Источник: Оба"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr/td[4]/input[3]').click()
            Log.write('Добавление поля "Штрихкод"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[2]').click()
            Log.write('Заполнение поля "Штрихкод"' + '\n')
            browser.find_element_by_xpath('//input[@maxlength="20"]').send_keys("00001002000000107292")
            time.sleep(0.2)
            Log.write('Запуск таймера' + '\n')
            Source = browser.page_source.encode('utf-8')
            HTML.write(Source)
            StartTime = time.time()
            Log.write('Нажатие кнопки "Поиск"' + '\n')
            ItemCount = browser.page_source
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[2]/table/tbody/tr/td[1]/input').click()
            Log.write('Отчет готов' + '\n')
            Log.write('Останавливаем таймер' + '\n')
            EndTime = time.time()
            Result = round(EndTime - StartTime, 2)
            Log.write('Тест 1.1 завершен. Время выполнения: ' + str(Result) + '\n')
            return Result
    except:
        if Debug:
            print 'Ошибка во время выполнения теста 1.1'
        Log.write('Ошибка во время выполнения теста 1.1\n')
        return int(-1)


def Test_12():
    try:
        global Log
        Date = (datetime.datetime.today()).strftime('%d.%m.%Y %H:%M:%S')
        URL = 'http://mow03-opt01/otcs/llisapi.dll?func=ll&objId=51692384&objAction=RunReport'
        if Debug:
            print 'Запуск теста 1.2			' + str(Date)
            print 'Загружаем URL'
            browser.get(URL)
            print 'Поиск кнопки "Поиск"'
            WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@value="Поиск"]')))
            print 'Поиск кнопки "Очистить"'
            browser.find_element_by_xpath('//*[@value="Очистить"]').click()
            time.sleep(0.2)
            print 'Выбор "Исочник: Оба"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr/td[4]/input[3]').click()
            print 'Выбор поля "Юридическое лицо"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[4]').click()
            print 'Выбор поля "Отсканирован"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[9]').click()
            print 'Заполнение поля "Юридическое лицо"'
            browser.find_element_by_xpath('//input[@size="5"]').send_keys(u"ЧМПЗ (КПП 774850001)")
            print 'Заполнение поля "Отсканирован"'
            browser.find_element_by_xpath('//input[@field_name="SCANDATEFROM"]').send_keys("01.01.2016")
            browser.find_element_by_xpath('//input[@field_name="SCANDATETO"]').send_keys("01.04.2016")
            time.sleep(0.2)
            print 'Запуск таймера'
            StartTime = time.time()
            print 'Нажатие кнопки "Поиск"'
            browser.find_element_by_xpath('//*[@onclick="doSearch(true);"]').click()
            print 'Отчет готов'
            print 'Останавливаем таймер'
            EndTime = time.time()
            Result = round(EndTime - StartTime, 2)
            print 'Тест 1.2 завершен. Время выполнения: ' + str(Result)
            return Result
        else:
            Log.write('Запуск теста 1.2			' + str(Date) + '\n')
            Log.write('Загружаем URL' + '\n')
            browser.get(URL)
            Log.write('Поиск кнопки "Поиск"' + '\n')
            WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@value="Поиск"]')))
            Log.write('Поиск кнопки "Очистить"' + '\n')
            browser.find_element_by_xpath('//*[@value="Очистить"]').click()
            time.sleep(0.2)
            Log.write('Выбор "Исочник: Оба"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr/td[4]/input[3]').click()
            Log.write('Выбор поля "Юридическое лицо"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[4]').click()
            Log.write('Выбор поля "Отсканирован"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[9]').click()
            Log.write('Заполнение поля "Юридическое лицо"' + '\n')
            browser.find_element_by_xpath('//input[@size="5"]').send_keys(u"ЧМПЗ (КПП 774850001)")
            Log.write('Заполнение поля "Отсканирован"' + '\n')
            browser.find_element_by_xpath('//input[@field_name="SCANDATEFROM"]').send_keys("01.01.2016")
            browser.find_element_by_xpath('//input[@field_name="SCANDATETO"]').send_keys("01.04.2016")
            time.sleep(0.2)
            Log.write('Запуск таймера' + '\n')
            StartTime = time.time()
            Log.write('Нажатие кнопки "Поиск"' + '\n')
            browser.find_element_by_xpath('//*[@onclick="doSearch(true);"]').click()
            Log.write('Отчет готов' + '\n')
            Log.write('Останавливаем таймер' + '\n')
            EndTime = time.time()
            Result = round(EndTime - StartTime, 2)
            Log.write('Тест 1.2 завершен. Время выполнения: ' + str(Result) + '\n')
            return Result
    except:
        if Debug:
            print 'Ошибка во время выполнения теста 1.2'
        Log.write('Ошибка во время выполнения теста 1.2\n')
        return int(-1)


def Test_13():
    try:
        global Log
        Result = []
        Date = (datetime.datetime.today()).strftime('%d.%m.%Y %H:%M:%S')
        URL = 'http://mow03-opt01/otcs/llisapi.dll?func=ll&objId=51692384&objAction=RunReport'
        if Debug:
            print 'Запуск теста 1.3			' + str(Date)
            print 'Загружаем URL'
            browser.get(URL)
            print 'Поиск кнопки "Поиск"'
            WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@value="Поиск"]')))
            print 'Поиск кнопки "Очистить"'
            browser.find_element_by_xpath('//*[@value="Очистить"]').click()
            time.sleep(0.2)
            print 'Выбор поля "Бизнес область"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[3]').click()
            print 'Выбор поля "Юридическое лицо"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[4]').click()
            print 'Выбор поля "Отсканирован"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[9]').click()
            print 'Заполнение поля "Бизнес область"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[4]/div/ul[1]/li/input').send_keys(
                u"ОСН_РЕАЛ")
            print 'Заполнение поля "Юридическое лицо"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td[4]/div/ul[1]/li/input').send_keys(
                u"ЧМПЗ (КПП 774850001)")
            print 'Заполнение поля "Отсканирован"'
            browser.find_element_by_xpath('//input[@field_name="SCANDATEFROM"]').send_keys("01.03.2016")
            browser.find_element_by_xpath('//input[@field_name="SCANDATETO"]').send_keys("01.04.2016")
            time.sleep(0.2)
            print 'Запуск таймера'
            StartTime = time.time()
            print 'Нажатие кнопки "Поиск"'
            browser.find_element_by_xpath('//*[@onclick="doSearch(true);"]').click()
            print 'Отчет готов'
            print 'Останавливаем таймер'
            EndTime = time.time()
            Result.append(round(EndTime - StartTime, 2))
            print 'Извлечение HTML кода страницы'
            ItemCount = browser.page_source
            S_Article = '1-50 of ([\s\S]*?) items'
            S_Compile = re.compile(S_Article)
            print 'Ищем количество элементов'
            S_Findall = S_Compile.findall(ItemCount)
            print 'Строка найдена'
            Result.append(S_Findall[0])
            print 'Тест 1.3 завершен. Время выполнения: ' + str(Result[0])
            print '                   Кол-во элементов: ' + str(Result[1])
            return Result
        else:
            Log.write('Запуск теста 1.3			' + str(Date) + '\n')
            Log.write('Загружаем URL' + '\n')
            browser.get(URL)
            Log.write('Поиск кнопки "Поиск"' + '\n')
            WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@value="Поиск"]')))
            Log.write('Поиск кнопки "Очистить"' + '\n')
            browser.find_element_by_xpath('//*[@value="Очистить"]').click()
            time.sleep(0.2)
            Log.write('Выбор поля "Бизнес область"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[3]').click()
            Log.write('Выбор поля "Юридическое лицо"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[4]').click()
            Log.write('Выбор поля "Отсканирован"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[9]').click()
            Log.write('Заполнение поля "Бизнес область"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[4]/div/ul[1]/li/input').send_keys(
                u"ОСН_РЕАЛ")
            Log.write('Заполнение поля "Юридическое лицо"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td[4]/div/ul[1]/li/input').send_keys(
                u"ЧМПЗ (КПП 774850001)")
            Log.write('Заполнение поля "Отсканирован"' + '\n')
            browser.find_element_by_xpath('//input[@field_name="SCANDATEFROM"]').send_keys("01.03.2016")
            browser.find_element_by_xpath('//input[@field_name="SCANDATETO"]').send_keys("01.04.2016")
            time.sleep(0.2)
            Log.write('Запуск таймера' + '\n')
            StartTime = time.time()
            Log.write('Нажатие кнопки "Поиск"' + '\n')
            browser.find_element_by_xpath('//*[@onclick="doSearch(true);"]').click()
            Log.write('Отчет готов' + '\n')
            Log.write('Останавливаем таймер' + '\n')
            EndTime = time.time()
            Result.append(round(EndTime - StartTime, 2))
            Log.write('Извлечение HTML кода страницы' + '\n')
            ItemCount = browser.page_source
            S_Article = '1-50 of ([\s\S]*?) items'
            S_Compile = re.compile(S_Article)
            Log.write('Ищем количество элементов' + '\n')
            S_Findall = S_Compile.findall(ItemCount)
            Log.write('Строка найдена' + '\n')
            Result.append(S_Findall[0])
            Log.write('Тест 1.3 завершен. Время выполнения: ' + str(Result) + '\n')
            Log.write('                   Кол-во элементов: ' + str(S_Findall[0]) + '\n')
            return Result
    except:
        if Debug:
            print 'Ошибка во время выполнения теста 1.3'
        Log.write('Ошибка во время выполнения теста 1.3\n')
        return [int(-1), 0]


def Test_14():
    try:
        global Log
        Date = (datetime.datetime.today()).strftime('%d.%m.%Y %H:%M:%S')
        URL = 'http://mow03-opt01/otcs/llisapi.dll?func=ll&objId=51692384&objAction=RunReport'
        if Debug:
            print 'Запуск теста 1.4			' + str(Date)
            print 'Загружаем URL'
            browser.get(URL)
            print 'Поиск кнопки "Поиск"'
            WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@value="Поиск"]')))
            print 'Поиск кнопки "Очистить"'
            browser.find_element_by_xpath('//*[@value="Очистить"]').click()
            time.sleep(0.2)
            print 'Выбор "Исочник: Оба"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr/td[4]/input[3]').click()
            print 'Выбор поля "Юридическое лицо"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[4]').click()
            print 'Выбор поля "Фронт-офис"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[7]').click()
            print 'Выбор поля "Отсканирован"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[9]').click()
            time.sleep(0.2)
            print 'Заполнение поля "Юридическое лицо"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[4]/div/ul[1]/li/input').send_keys(
                u"ЧМПЗ (КПП 774850001)")
            print 'Заполнение поля "Фронт-офис"'
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td[4]/div/ul[1]/li/input').send_keys(
                u"Москва-3")
            print 'Заполнение поля "Отсканирован"'
            browser.find_element_by_xpath('//input[@field_name="SCANDATEFROM"]').send_keys("01.03.2016")
            browser.find_element_by_xpath('//input[@field_name="SCANDATETO"]').send_keys("01.04.2016")
            time.sleep(0.2)
            print 'Запуск таймера'
            StartTime = time.time()
            print 'Нажатие кнопки "Поиск"'
            browser.find_element_by_xpath('//*[@onclick="doSearch(true);"]').click()
            print 'Отчет готов'
            print 'Останавливаем таймер'
            EndTime = time.time()
            Result = round(EndTime - StartTime, 2)
            print 'Тест 1.4 завершен. Время выполнения: ' + str(Result)
            return Result
        else:
            Log.write('Запуск теста 1.4			' + str(Date) + '\n')
            Log.write('Загружаем URL' + '\n')
            browser.get(URL)
            Log.write('Поиск кнопки "Поиск"' + '\n')
            WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@value="Поиск"]')))
            Log.write('Поиск кнопки "Очистить"' + '\n')
            browser.find_element_by_xpath('//*[@value="Очистить"]').click()
            time.sleep(0.2)
            Log.write('Выбор "Исочник: Оба"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr/td[4]/input[3]').click()
            Log.write('Выбор поля "Юридическое лицо"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[4]').click()
            Log.write('Выбор поля "Фронт-офис"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[7]').click()
            Log.write('Выбор поля "Отсканирован"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr[1]/td/select/option[9]').click()
            time.sleep(0.2)
            Log.write('Заполнение поля "Юридическое лицо"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[4]/div/ul[1]/li/input').send_keys(
                u"ЧМПЗ (КПП 774850001)")
            Log.write('Заполнение поля "Фронт-офис"' + '\n')
            browser.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[1]/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td[4]/div/ul[1]/li/input').send_keys(
                u"Москва-3")
            Log.write('Заполнение поля "Отсканирован"' + '\n')
            browser.find_element_by_xpath('//input[@field_name="SCANDATEFROM"]').send_keys("01.03.2016")
            browser.find_element_by_xpath('//input[@field_name="SCANDATETO"]').send_keys("01.04.2016")
            time.sleep(0.2)
            Log.write('Запуск таймера' + '\n')
            StartTime = time.time()
            Log.write('Нажатие кнопки "Поиск"' + '\n')
            browser.find_element_by_xpath('//*[@onclick="doSearch(true);"]').click()
            Log.write('Отчет готов' + '\n')
            Log.write('Останавливаем таймер' + '\n')
            EndTime = time.time()
            Result = round(EndTime - StartTime, 2)
            Log.write('Тест 1.4 завершен. Время выполнения: ' + str(Result) + '\n')
            return Result
    except:
        if Debug:
            print 'Ошибка во время выполнения теста 1.4'
        Log.write('Ошибка во время выполнения теста 1.4\n')
        return int(-1)


if Debug:
    Test_11()
    Test_12()
    Test_13()
    Test_14()
else:
    Log.write('#####################################################################\n')
    Test1 = Test_11()
    try:
        subprocess.call(['zabbix_sender -z 192.168.20.109 -p 10051 -s "MOW03-OPT01" -k web_task_1 -o ' + str(Test1)],
                        shell=True)
        Log.write('Данные теста 1.1 успешно отправленны.\n\n')
    except:
        Log.write('Не удалось отправить данные теста 1.1\n\n')

    Log.write('#####################################################################\n')
    time.sleep(5)
    Test2 = Test_12()
    try:
        subprocess.call(['zabbix_sender -z 192.168.20.109 -p 10051 -s "MOW03-OPT01" -k web_task_2 -o ' + str(Test2)],
                        shell=True)
        Log.write('Данные теста 1.2 успешно отправленны.\n\n')
    except:
        Log.write('Не удалось отправить данные теста 1.2\n\n')

    Log.write('#####################################################################\n')
    Test3 = Test_13()
    try:
        subprocess.call(['zabbix_sender -z 192.168.20.109 -p 10051 -s "MOW03-OPT01" -k web_task_3 -o ' + str(Test3[0])],
                        shell=True)
        subprocess.call(
            ['zabbix_sender -z 192.168.20.109 -p 10051 -s "MOW03-OPT01" -k web_task_3h -o ' + str(Test3[1])],
            shell=True)
        Log.write('Данные теста 1.3 успешно отправленны.\n\n')
    except:
        Log.write('Не удалось отправить данные теста 1.3\n\n')

    Log.write('#####################################################################\n')
    Test4 = Test_14()
    try:
        subprocess.call(['zabbix_sender -z 192.168.20.109 -p 10051 -s "MOW03-OPT01" -k web_task_4 -o ' + str(Test4)],
                        shell=True)
        Log.write('Данные теста 1.4 успешно отправленны.\n\n')
    except:
        Log.write('Не удалось отправить данные теста 1.4\n\n')

HTML.close()
browser.quit()
Log.close()
display.stop()
