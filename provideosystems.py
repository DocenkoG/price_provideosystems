# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import sys
import configparser
import time
import shutil
#import openpyxl                      # Для .xlsx
import xlrd                          # для .xls
from   price_tools import getCellXlsx, getCell, quoted, dump_cell, currencyType, openX, sheetByName
import csv
#import requests, lxml.html



def getXlsString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        j = in_columns_j[item]-1
        if item in ('закупка','продажа','цена1') :
            if getCell(row=i, col=j, isDigit='N', sheet=sh).find('Звоните') >=0 :
                impValues[item] = '0.1'
            else :
                impValues[item] = getCell(row=i, col=j, isDigit='Y', sheet=sh)
            #print(sh, i, sh.cell( row=i, column=j).value, sh.cell(row=i, column=j).number_format, currencyType(sh, i, j))
        elif item == 'валюта_по_формату':
            impValues[item] = currencyType(row=i, col=j, sheet=sh)
        else:
            impValues[item] = getCell(row=i, col=j, isDigit='N', sheet=sh)
    return impValues



def getXlsxString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        j = in_columns_j[item]
        if item in ('закупка','продажа','цена','цена1') :
            if getCellXlsx(row=i, col=j, isDigit='N', sheet=sh).find('Звоните') >=0 :
                impValues[item] = '0.1'
            else :
                impValues[item] = getCellXlsx(row=i, col=j, isDigit='Y', sheet=sh)
            #print(sh, i, sh.cell( row=i, column=j).value, sh.cell(row=i, column=j).number_format, currencyType(sh, i, j))
        elif item == 'валюта_по_формату':
            impValues[item] = currencyType(row=i, col=j, sheet=sh)
        else:
            impValues[item] = getCellXlsx(row=i, col=j, isDigit='N', sheet=sh)
    return impValues



def convert_excel2csv(cfg):
    csvFName  = cfg.get('basic','filename_out')
    priceFName= cfg.get('basic','filename_in')
    sheetName = cfg.get('basic','sheetname')
    
    log.debug('Reading file ' + priceFName )
    book = xlrd.open_workbook(priceFName.encode('cp1251'), formatting_info=True)
    sheet = sheetByName(fileName = priceFName, sheetName = sheetName)
    if not sheet :
        log.error("Нет листа "+sheetName+" в файле "+ priceFName)
        return False
    log.debug("Sheet   "+sheetName)
    out_cols = cfg.options("cols_out")
    in_cols  = cfg.options("cols_in")
    out_template = {}
    for vName in out_cols :
         out_template[vName] = cfg.get("cols_out", vName)
    in_cols_j = {}
    for vName in in_cols :
         in_cols_j[vName] = cfg.getint("cols_in",  vName)
    #brands,   discount     = config_read(cfgFName, 'discount')
    #for k in discount.keys():
    #    discount[k] = (100 - int(discount[k]))/100
    #print(discount)

    outFile = open( csvFName, 'w', newline='', encoding='CP1251', errors='replace')
    csvWriter = csv.DictWriter(outFile, fieldnames=out_cols )
    csvWriter.writeheader()

    '''                                     # Блок проверки свойств для распознавания групп      XLSX                                  
    for i in range(2393, 2397):                                                         
        i_last = i
        ccc = sheet.cell( row=i, column=in_cols_j['группа'] )
        print(i, ccc.value)
        print(ccc.font.name, ccc.font.sz, ccc.font.b, ccc.font.i, ccc.font.color.rgb, '------', ccc.fill.fgColor.rgb)
        print('------')
    '''
    '''                                     # Блок проверки свойств для распознавания групп      XLS
    for i in range(1, 12):
        xfx = sheet.cell_xf_index(i, 1)
        book = xlrd.open_workbook(priceFName.encode('cp1251'), formatting_info=True)
        xf  = book.xf_list[xfx]
        bgci  = xf.background.pattern_colour_index
        fonti = xf.font_index
        ccc = sheet.cell(i, 1)
        if ccc.value == None :
            print (i, colSGrp, 'Пусто!!!')
            continue
                                         # Атрибуты шрифта для настройки конфига
        font = book.font_list[fonti]
        print( '---------------------- Строка', i, '-----------------------', sheet.cell(i, 1).value)
        print( 'background_colour_index=',bgci)
        print( 'fonti=', fonti, '           xf.alignment.indent_level=', xf.alignment.indent_level)
        print( 'bold=', font.bold)
        print( 'weight=', font.weight)
        print( 'height=', font.height)
        print( 'italic=', font.italic)
        print( 'colour_index=', font.colour_index )
        print( 'name=', font.name)
    return
    '''

    recOut  ={}
    grp = ''
    subgrp = ''
    subgrp2 = ''
#   for i in range(1, sheet.max_row +1) :                               # xlsx
    for i in range(1, sheet.nrows) :                                     # xls
        i_last = i
        try:
            #impValues = getXlsxString(sheet, i, in_cols_j)              # xlsx
            impValues = getXlsString(sheet, i, in_cols_j)                # xls
            #print( impValues )
            xfx = sheet.cell_xf_index(i, 1)
            xf = book.xf_list[xfx]
            bgci = xf.background.pattern_colour_index
            fonti = xf.font_index

            if (impValues['код_'] in ('', 'Partnumber', 'Part No.') or
                impValues['цена1'] in ('SRP, $','RRP, $', 'Цена MSRP')): # Пустая строка
                continue
            if impValues['цена1'] == '0':                                # Вместо отсутствия цены ставим цену 0.1
                impValues['цена1'] = '0.1'
            if cfg.has_option('cols_in', 'примечание') and impValues['примечание'] != '':      # Примечание
                impValues['примечание'] = ' / (' + impValues['примечание'] + ')'               # обрамляем скобками
            if "\n" in impValues['код_']:                                # В многострочном коде берем
                p = impValues['код_'].rfind("\n")                        # последнюю строку
                impValues['код_'] = impValues['код_'][p+1:]

            if cfg.has_option('cols_in', 'подгруппа') and impValues['код_'] == '' and impValues['подгруппа'] != '':
                subgrp = impValues['подгруппа']                          # Подгруппа
                continue
            elif bgci == 43:                                             # Подгруппа желтая
                subgrp2 = impValues['группа_']
                continue
            elif bgci == 22:                                             # Группа
                subgrp = ''
                grp = impValues['группа_']
            else :                                                       # Обычная строка
                if cfg.has_option('cols_in', 'группа_'):
                    impValues['группа_'] = grp
                if cfg.has_option('cols_in', 'подгруппа'):
                    impValues['подгруппа'] = subgrp
                for outColName in out_template.keys():
                    shablon = out_template[outColName]
                    for key in impValues.keys():
                        if shablon.find(key) >= 0:
                            shablon = shablon.replace(key, impValues[key])
                    if (outColName == 'закупка') and ('*' in shablon):
                        if impValues['цена1'] == '0.1':
                            shablon = '0.1'
                        else:
                            p = shablon.find("*")
                            vvv1 = float(shablon[:p])
                            vvv2 = float(shablon[p+1:])
                            shablon = str(round(vvv1 * vvv2, 2))

                    recOut[outColName] = shablon.strip()

                csvWriter.writerow(recOut)

        except Exception as e:
            print(e)
            if str(e) == "'NoneType' object has no attribute 'rgb'":
                pass
            else:
                log.debug('Exception: <' + str(e) + '> при обработке строки ' + str(i) +'.' )

    log.info('Обработано ' +str(i_last)+ ' строк.')
    outFile.close()



def download(cfg):
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.remote.remote_connection import LOGGER
    LOGGER.setLevel(logging.WARNING)
     
    retCode     = False
    filename1_new= cfg.get('basic','filename1_new')
    filename2_new= cfg.get('basic','filename2_new')
    filename1_old= cfg.get('basic','filename1_old')
    filename2_old= cfg.get('basic','filename2_old')
    login       = cfg.get('download','login'    )
    password    = cfg.get('download','password' )
    url_lk      = cfg.get('download','url_lk'   )
    url_file1   = cfg.get('download','url_file1')
    url_file2   = cfg.get('download','url_file2')

    download_path= os.path.join(os.getcwd(), 'tmp')
    if not os.path.exists(download_path):
        os.mkdir(download_path)

    for fName in os.listdir(download_path) :
        os.remove( os.path.join(download_path, fName))
    dir_befo_download = set(os.listdir(download_path))
        
    if os.path.exists('geckodriver.log') : os.remove('geckodriver.log')
#    try:
    ffprofile = webdriver.FirefoxProfile()
    ffprofile.set_preference("browser.download.dir", download_path)
    ffprofile.set_preference("browser.download.folderList",2);
    ffprofile.set_preference("browser.helperApps.neverAsk.saveToDisk",
                ",application/octet-stream" + 
                ",application/vnd.ms-excel" + 
                ",application/vnd.msexcel" + 
                ",application/x-excel" + 
                ",application/x-msexcel" + 
                ",application/zip" + 
                ",application/xls" + 
                ",application/vnd.ms-excel" +
                ",application/vnd.ms-excel.addin.macroenabled.12" +
                ",application/vnd.ms-excel.sheet.macroenabled.12" +
                ",application/vnd.ms-excel.template.macroenabled.12" +
                ",application/vnd.ms-excelsheet.binary.macroenabled.12" +
                ",application/vnd.ms-fontobject" +
                ",application/vnd.ms-htmlhelp" +
                ",application/vnd.ms-ims" +
                ",application/vnd.ms-lrm" +
                ",application/vnd.ms-officetheme" +
                ",application/vnd.ms-pki.seccat" +
                ",application/vnd.ms-pki.stl" +
                ",application/vnd.ms-word.document.macroenabled.12" +
                ",application/vnd.ms-word.template.macroenabed.12" +
                ",application/vnd.ms-works" +
                ",application/vnd.ms-wpl" +
                ",application/vnd.ms-xpsdocument" +
                ",application/vnd.openofficeorg.extension" +
                ",application/vnd.openxmformats-officedocument.wordprocessingml.document" +
                ",application/vnd.openxmlformats-officedocument.presentationml.presentation" +
                ",application/vnd.openxmlformats-officedocument.presentationml.slide" +
                ",application/vnd.openxmlformats-officedocument.presentationml.slideshw" +
                ",application/vnd.openxmlformats-officedocument.presentationml.template" +
                ",application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" +
                ",application/vnd.openxmlformats-officedocument.spreadsheetml.template" +
                ",application/vnd.openxmlformats-officedocument.wordprocessingml.template" +
                ",application/x-ms-application" +
                ",application/x-ms-wmd" +
                ",application/x-ms-wmz" +
                ",application/x-ms-xbap" +
                ",application/x-msaccess" +
                ",application/x-msbinder" +
                ",application/x-mscardfile" +
                ",application/x-msclip" +
                ",application/x-msdownload" +
                ",application/x-msmediaview" +
                ",application/x-msmetafile" +
                ",application/x-mspublisher" +
                ",application/x-msschedule" +
                ",application/x-msterminal" +
                ",application/x-mswrite" +
                ",application/xml" +
                ",application/xml-dtd" +
                ",application/xop+xml" +
                ",application/xslt+xml" +
                ",application/xspf+xml" +
                ",application/xv+xml" +
                ",application/excel")
    if os.name == 'posix':
            #driver = webdriver.Firefox(ffprofile, executable_path=r'/usr/local/Cellar/geckodriver/0.19.1/bin/geckodriver')
            driver = webdriver.Firefox(ffprofile, executable_path=r'/usr/local/bin/geckodriver')
    elif os.name == 'nt':
            driver = webdriver.Firefox(ffprofile)
    driver.implicitly_wait(10)
    driver.set_page_load_timeout(70)

    driver.get('http://www.provis.ru/partners/dealer/')
    time.sleep(2)
    driver.find_element_by_name("login").click()
    driver.find_element_by_name("login").clear()
    driver.find_element_by_name("login").send_keys(login)
    driver.find_element_by_name("pass").click()
    driver.find_element_by_name("pass").clear()
    driver.find_element_by_name("pass").send_keys(password)
    driver.find_element_by_id("partners_auth_btn").click()
    time.sleep(1)
    try:
        driver.get(url_file1)
    except Exception as e:
        log.debug('Exception: <' + str(e) + '>')
        print('-exept-error-',str(e))

    time.sleep(5)
    driver.set_page_load_timeout(15)
    try:
        driver.get(url_file2)
    except Exception as e:
        log.debug('Exception: <' + str(e) + '>')
        print('-exept-error-', str(e))

    time.sleep(5)
    driver.find_element_by_link_text(u"Выход").click()
    driver.quit()


    dir_afte_download = set(os.listdir(download_path))
    new_files = list( dir_afte_download.difference(dir_befo_download))
    print(new_files)
    if len(new_files) < 2:
        log.error( 'Не удалось скачать файл прайса ')
        retCode= False
    elif len(new_files) > 2:
        log.error( 'Скачалось несколько файлов. Надо разбираться ...')
        retCode= False
    else:   
        new_file = new_files[0]                                                     # загруженo ровно два файл.
        new_ext  = os.path.splitext(new_file)[-1].lower()
        DnewFile = os.path.join( download_path,new_file)
        new_file_date = os.path.getmtime(DnewFile)
        log.info( 'Скачанный файл ' +new_file + ' имеет дату ' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(new_file_date) ) )

        print(new_ext)
        if new_ext in ('.xls','.xlsx','.xlsb','.xlsm','.csv'):
            if os.path.exists( filename1_new) and os.path.exists( filename1_old):
                os.remove( filename1_old)
                os.rename( filename1_new, filename1_old)
            if os.path.exists( filename1_new) :
                os.rename( filename1_new, filename1_old)
            shutil.copy2( DnewFile, filename1_new)
            retCode= True

        new_file = new_files[1]  # загруженo ровно два файл.
        new_ext = os.path.splitext(new_file)[-1].lower()
        DnewFile = os.path.join(download_path, new_file)
        new_file_date = os.path.getmtime(DnewFile)
        log.info('Скачанный файл ' + new_file + ' имеет дату ' + time.strftime("%Y-%m-%d %H:%M:%S",
                                                                               time.localtime(new_file_date)))

        print(new_ext)
        if new_ext in ('.xls', '.xlsx', '.xlsb', '.xlsm', '.csv'):
            if os.path.exists(filename2_new) and os.path.exists(filename2_old):
                os.remove(filename2_old)
                os.rename(filename2_new, filename2_old)
            if os.path.exists(filename2_new):
                os.rename(filename2_new, filename2_old)
            shutil.copy2(DnewFile, filename2_new)
            retCode = True

    return retCode




def config_read( cfgFName ):
    cfg = configparser.ConfigParser(inline_comment_prefixes=('#'))
    if  os.path.exists('private.cfg'):     
        cfg.read('private.cfg', encoding='utf-8')
    if  os.path.exists(cfgFName):     
        cfg.read( cfgFName, encoding='utf-8')
    else: 
        log.debug('Нет файла конфигурации '+cfgFName)
    return cfg



def is_file_fresh(fileName, qty_days):
    if os.path.exists(fileName):
        price_datetime = os.path.getmtime(fileName)
    else:
        log.error('Не найден файл  '+ fileName)
        return False

    file_age = round((time.time() - price_datetime) / 24 / 60 / 60)
    if file_age > qty_days :
        log.error('Файл "' + fileName + '" устарел! Допустимый период ' + str(qty_days)+' дней, а ему ' + str(file_age))
        return False
    else:
        return True



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def processing(cfgFName):
    log.info('----------------------- Processing '+cfgFName )
    cfg = config_read(cfgFName)
    filename_new = cfg.get('download','filename_new')
    
    if cfg.has_section('download'):
        result = download(cfg)
    if is_file_fresh( filename_new, int(cfg.get('basic','срок годности'))):
        #os.system( 'brullov_converter_xlsx.xlsm')
        #convert_csv2csv(cfg)
        convert_excel2csv(cfg)



def main(dealerName):
    """ Обработка прайсов выполняется согласно файлов конфигурации.
    Для этого в текущей папке должны быть файлы конфигурации, описывающие
    свойства файла и правила обработки. По одному конфигу на каждый
    прайс или раздел прайса со своими правилами обработки
    """
    make_loger()
    log.info('          ' + dealerName)

    rc_download = False
    if os.path.exists('getting.cfg'):
        cfg = config_read('getting.cfg')
        filename1_new = cfg.get('basic','filename1_new')

        if cfg.has_section('download'):
            rc_download = download(cfg)
        if not(rc_download==True or is_file_fresh( filename1_new, int(cfg.get('basic','срок годности')))):
            return False

    for cfgFName in os.listdir("."):
        if cfgFName.startswith("cfg") and cfgFName.endswith(".cfg"):
            log.info('----------------------- Processing '+cfgFName )
            cfg = config_read(cfgFName)
            filename_in = cfg.get('basic','filename_in')
            if rc_download==True or is_file_fresh( filename_in, int(cfg.get('basic','срок годности'))):
                convert_excel2csv(cfg)



if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir = os.path.dirname (sys.argv[0])
    print(mydir, myName)
    main(myName)