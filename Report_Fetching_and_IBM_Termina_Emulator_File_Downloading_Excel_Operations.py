import logging
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import urllib.request
import zipfile
import pandas as pd
from Kaizen_timestamp import insert_timestamp
import glob
import stat
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import datetime
import re
from ibm import iseries
import os
import xlwings as xw
from easygui import msgbox
import win32com.client as wc
import sys

st_new = time.gmtime()
cwd = os.getcwd()
sysUser = os.getlogin()
print(sysUser, '\n', cwd)
# cwd = "C:\\Users\\vijaykumar.mane\\Desktop\\Invstat"
dateDownload = datetime.datetime.now().strftime("%Y-%m-%d")
nameDate = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M")
# Macro start time
st = datetime.datetime.now()
try:
    # Check If Chrome driver exists
    chome = os.path.expanduser("~")
    check = os.path.exists(chome + "\\Downloads\\chromedriver.exe")
    if not check:
        # download chromedriver
        # urllib.request.urlretrieve('http://chromedriver.storage.googleapis.com/2.8/chromedriver_win32.zip', chome + "\\Downloads\\chromedriver.zip")
        urllib.request.urlretrieve('https://chromedriver.storage.googleapis.com/2.27/chromedriver_win32.zip', chome + "\\Downloads\\chromedriver.zip")
        # Unzip file
        zip1 = zipfile.ZipFile(chome + "\\Downloads\\chromedriver.zip", 'r')
        zip1.extractall(chome + "\\Downloads\\")
    else:
        pass

    # Open input file
    temp = xw.Book(cwd+'\\Macro Files\\InputFileINV.xlsm')
    sh = temp.sheets('Sheet1')

    # Ceva Dashboard Credentials
    username = sh.range('G8').value
    password = sh.range('G9').value
    signonwindow = sh.range('G10').value
    time.sleep(1)
    try:
        list_of_files = glob.glob(cwd + '\\Daybookf Archive\\*.xls')
        latest_file = max(list_of_files, key=os.path.getctime)
    except:
        latest_file = []
    # Age of file
    def file_age_in_seconds(pathname):
        try:
            return time.time() - os.stat(pathname)[stat.ST_MTIME]
        except:
            return 999999

    age = file_age_in_seconds(latest_file)
    if age < 10800:
        flag1 = 1
    else:
        flag1 = 0
    # Generate file ID from IBM Terminal Emulator

    def vDict(keyValue, valueValue):
        tempDict = {}
        for i, j in zip(keyValue, valueValue):
            if not i in tempDict:
                tempDict[i] = j
            else:
                pass
        return(tempDict)

    userOFS = sh.range('G5').value
    passOFS = sh.range('G6').value
    print(userOFS,' ', passOFS)
    sessionName = sh.range('G11').value #'A'
    profileName = sh.range('G12').value #'OFS-live.WS'
    print('sessionName: ',sessionName,'  profileName: ',profileName )
    try:
        st1 = datetime.datetime.now()
        con = iseries(uid=userOFS, pwd=passOFS, session=sessionName, profile=profileName)
        # con.connect()
        try:
            con.connect()
        except:
            print("Instance of PCOMM is still running, please wait while we kill it for you \nExiting code please rerun")
            os.system("taskkill /IM pcscm.exe /F")
            os.system("taskkill /IM pcsws.exe /F")
            sys.exit()
        con.signon(signonwindow)
        con.test_window()
        con.start_communication()
        con.test_comm()
        con.system_check()
        con.wait_for_text('User',3,2)
        con.set_text(value=userOFS,row=3, col=9)
        con.set_text(value=passOFS,row=4, col=9)
        con.send_keys('enter')
        time.sleep(4)
        # Check screen
        scrCheck = con.get_rect_text(1,28,1,51)
        if "Display Program Messages" in scrCheck or 'Sign-on Information' in scrCheck:
            con.send_keys("enter")
            time.sleep(4)
        scrCheck = con.get_rect_text(1,28,1,51)
        if "Display Program Messages" in scrCheck or 'Sign-on Information' in scrCheck:
            con.send_keys("enter")
        # Select Americas
        con.wait_for_text('Americas',6,10)
        con.wait_for_screen_cursor(20,7)
        con.set_text(value='1',row=20, col=7)
        con.send_keys('enter')
        # Select region
        con.wait_for_text('USA',6,11)
        con.wait_for_screen_cursor(20,7)
        con.set_text(value='1',row=20, col=7)
        con.send_keys('enter')
        time.sleep(2)
        # Report Daybookf
        con.wait_for_text('Forwarding',4,20)
        if flag1 == 0:
            con.set_text(value='DAYBOOKF',row=22, col=23)
            con.send_keys('enter')
            time.sleep(1)

            if 'unmonitored by' in con.get_rect_text(24,2,24,60):
                # msgbox("Function check. QRY5080 unmonitored by BASUXCMD at statement 0000024800, ins\nPlease check might be other using same Credentials to run query.")
                print("Function check. QRY5080 unmonitored by BASUXCMD at statement 0000024800, ins\nPlease check might be other using same Credentials to run query.")
                sys.exit()

            # Modify query
            con.wait_for_text('Select Records',1,34)
            end_date = datetime.datetime.now().strftime("%Y%m%d")

            start_date = (datetime.datetime.now() + datetime.timedelta(-32)).strftime("%Y%m%d")

            con.set_text(value=start_date+' '+end_date ,row=7, col=35)
            print("Query run for date range: ",start_date+' : '+end_date)
            time.sleep(1)
            con.set_text(value="'L' 'G' 'N' 'B' 'E' 'K'" ,row=12, col=35)
            time.sleep(1)
            con.send_keys('enter')
            time.sleep(5)
            # Loop while to get fileID
            print('Query running.....')
            while 'Query running' in con.get_rect_text(24,2,24,60) or '         ' in con.get_rect_text(24,2,24,60):
                if 'unmonitored by' in con.get_rect_text(24,2,24,60):
                    # sys.exit("Function check. QRY5080 unmonitored by BASUXCMD at statement 0000024800, ins\nPlease check might be other using same Credentials to run query.")
                    print("Function check. QRY5080 unmonitored by BASUXCMD at statement 0000024800, ins\nPlease check might be other using same Credentials to run query.")
                    sys.exit()
                time.sleep(2)
                # print(con.get_rect_text(24,2,24,60))
            # Get file ID
            con.wait_for_text('Forwarding',4,20)
            fileID = con.get_rect_text(24,2,24,60)
            fileID = re.findall('DAYBOOKF\s+in\s+(\w+)\s+was',fileID)
            if len(fileID) > 0:
                downloadId = fileID[0]
            else:
                # msgbox('Query Failed.\n Please run again')
                # sys.exit('fileID is Blank: Query Failed.')
                sys.exit()
    except:
        # msgbox('Error Ocurred, please confirm\n1.Please close IBM emulator if open\n2.Create IBM Emulator Profile as "OFS-live.WS"\n3.Someone else might using Credentials to runing query')
        # sys.exit("Error Ocurred, please confirm\n1.Create IBM Emulator Profile as 'OFS-live.WS'")
        sys.exit()

    if flag1 == 0:
        # OFS Close
        print(downloadId)
        if downloadId == None:
            # msgbox('Daybookf report not Generated Please run it again')
            sys.exist('fileID is Blank: Query Failed.')

    # Download file from IBM iSeries
    path1 = cwd+'\\Daybookf Archive\\'
    ip='10.235.108.20'
    uid= userOFS
    pwd= passOFS
    if flag1 == 0:
        path=path1+"Daybookf "+nameDate+'.xls'
        fname= downloadId +'/'+'DAYBOOKF('+uid.upper()+')'
        iseries.iSeries_download(ip, uid, pwd, fname, path)
        listFiles = glob.glob(path1+"*.temp")
        while len(listFiles) > 0:
            time.sleep(5)
            listFiles = glob.glob(path1+"*.temp")
        et1 = datetime.datetime.now()
        ttm1 = (et1 - st1).total_seconds()
        print("Time taken to download Daybookf: Seconds ", ttm1)
        print('Daybookf Downloaded\n',path)
        time.sleep(5)


        def consolidate_excelbook(sourcePath, BooktoConsolidate_in_Wb_Obj):
            temp_wb = BooktoConsolidate_in_Wb_Obj
            print('Source Path :', sourcePath)
            k = 0
            temp_sht = temp_wb.sheets[k]
            # full_inv = xw.Book(sourcePath)
            full_inv = pd.ExcelFile(sourcePath)
            # print(sourcePath)
            # full_sheetCount = full_inv.sheets.count
            full_sheetCount = len(full_inv.sheet_names)
            for i in range(0,full_sheetCount):
                # full_sheet = full_inv.sheets[i]
                df = full_inv.parse(i)
                print("Extracting sheet# : ",i)
                df1 = df[['CREATEDATE','BOKPRT','CSORNADR1','HAWBNO','HBLNO','HOUSENO','AIRSTATUS','FWDSTATUS', 'MATRCDATE', 'MATRDDATE','CSORNO']]
                df = None
                # full_sheet.range('A1').api.CurrentRegion.Copy()
                current_position_data = int(temp_sht.range('A1').api.CurrentRegion.Address.split('$')[-1])
                while not current_position_data <= 983000:
                    k += 1
                    temp_sht = temp_wb.sheets[k]
                    current_position_data = int(temp_sht.range('A1').api.CurrentRegion.Address.split('$')[-1])
                    print(current_position_data)
                if current_position_data == 1:
                    temp_sht.range('A'+str(current_position_data)).options(index=False).value = df1
                    time.sleep(5)
                    temp_sht.range(str(current_position_data)+":"+str(current_position_data)).api.Delete()
                else:
                    temp_sht.range('A'+str(current_position_data+1)).options(index=False).value = df1
                    time.sleep(5)
                    temp_sht.range(str(current_position_data+1)+":"+str(current_position_data+1)).api.Delete()
                    # deleteFilter(temp_sht)


        # path = r'C:\Users\vijaykumar.mane\Desktop\Invstat\Daybookf Archive\Daybookf 2017-03-02-16-36.xls'

        wbt = xw.Book(cwd+'\\Daybookf Archive\\Temp.xlsb')

        consolidate_excelbook(path, wbt)


        dfsht = wbt.sheets[0]
        dfa = xw.Book(cwd+'\\Daybookf Archive\\'+'BookingwithHawbCompleteArchive.xlsb')

        dfasht1 = dfa.sheets[0]
        # dfasht2 = dfa.sheets[1]

        # Last row function
        def lr(shtO):
            lrN = shtO.cells(shtO.range('A:A').rows.count, 1).end('up').row
            return(lrN)

        lrDf = lr(dfsht)

        today = datetime.datetime.now()
        DD = datetime.timedelta(days=32)
        earlier = today - DD
        earlier_str = earlier.strftime("%Y%m%d")
        cond = dfasht1.range("A1").value
        if cond != None:
            dfasht1.range('A1').api.AutoFilter(Field=7, Criteria1='>='+ earlier_str, Operator=7)
            time.sleep(4)
        else:
            dfasht1.range("A1").options(transpose=False).value = ['BOKPRTT','CSORNADR1','CSORNO','HAWBNO','AIRSTATUS','FWDSTATUS','CREATEDATE','MATRCDATE','MATRDDATE']
        column = 'G'
        sheetObj = dfasht1
        strrng = sheetObj.range("B2", sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up')).api.SpecialCells(12).Cells(1,1).Row
        endrng = sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up').row
        if not strrng == 1:
            sheetObj.range(str(strrng)+":"+str(endrng)).api.Delete()
            time.sleep(5)
        dfasht1.range('A1').api.AutoFilter(Field=7)
        time.sleep(5)

        lrDfasht1 = lr(dfasht1)
        if lrDfasht1 > 1000000:
            dfa.sheets.add(earlier_str,dfasht1)
            dfasht1 = dfa.sheets[0]
            dfasht1.range("A1").value = dfa.sheets[1].range("A1:I1").value

        # Create Date
        createDate_df = dfsht.range('A1:A'+str(lrDf)).value
        dfasht1.range('G'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df

        # Booking num
        createDate_df = dfsht.range('B1:B'+str(lrDf)).value
        dfasht1.range('A'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
        # CSORNADR1
        createDate_df = dfsht.range('C1:C'+str(lrDf)).value
        dfasht1.range('B'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
        # CSORNO
        createDate_df = dfsht.range('K1:K'+str(lrDf)).value
        dfasht1.range('C'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df

        # HOUSENO
        createDate_df1 = dfsht.range('D1:D'+str(lrDf)).value
        # HBLNO
        createDate_df2 = dfsht.range('E1:E'+str(lrDf)).value
        # HAWBNO
        createDate_df3 = dfsht.range('F1:F'+str(lrDf)).value
        createDate_df4 = []
        # Merging of columns
        print(len(createDate_df1),' ',len(createDate_df1),' ',len(createDate_df1),' ')
        for i,j,k in zip(createDate_df1, createDate_df2, createDate_df3):
            if i == None:
                if j == None:
                    createDate_df4.append(k)
                else:
                    createDate_df4.append(j)
            else:
                createDate_df4.append(i)

        dfasht1.range('D'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df4

        # AIRSTATUS
        createDate_df = dfsht.range('G1:G'+str(lrDf)).value
        dfasht1.range('E'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
        # 'FWDSTATUS'
        createDate_df = dfsht.range('H1:H'+str(lrDf)).value
        dfasht1.range('F'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
        # MATRCDATE
        createDate_df = dfsht.range('I1:I'+str(lrDf)).value
        dfasht1.range('H'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
        # MATRDDATE
        createDate_df = dfsht.range('J1:J'+str(lrDf)).value
        dfasht1.range('I'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
        # Save Archive Daybookf
        dfa.save()
        time.sleep(1)
        dfa.close()
        wbt.close()
        print("Booking with Hawb Updated...")
        # File Download path

    list_of_files = glob.glob(cwd + '\\Unprinted Archive\\*.csv')
    latest_file2 = max(list_of_files, key=os.path.getctime)

    age2 = file_age_in_seconds(latest_file2)
    if age2 < 10800:
        flag2 = 1
    else:
        flag2 = 0

    if flag2 == 0:
        chromeOptions = webdriver.ChromeOptions()
        prefs = {"download.default_directory" : cwd+"\\Unprinted Archive\\"}
        chromeOptions.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(
            "C:\\Users\\" + sysUser + "\\Downloads\\chromedriver.exe", chrome_options=chromeOptions)
        driver.maximize_window()
        driver.set_page_load_timeout(900)
        username = username.replace('/','%5c')
        password = password.replace('@','%40')
        driver.get("http://"+username+":"+password+"@cndreporting.logistics.corp/matrix_ofs/operations_scorecard.php?setcountry=92")

        element = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, "//*[@id='results']/table/tbody/tr[1]")))

        elem = driver.switch_to_active_element()
        elem.send_keys(Keys.END)
        fileName = cwd + nameDate +".png"
        driver.get_screenshot_as_file(fileName)
        os.remove(fileName)
        time.sleep(2)
        driver.get_screenshot_as_file(filename= cwd +"\\Snapshot Archive\\Screen "+nameDate+".png")
        print("Screenshot captured")
        # Unprinted download URI
        try:
            driver.get("http://cndreporting.logistics.corp/matrix_ofs/report_drilldown.php?metric=unprinted&category1=Total&category2=Total&cluster=NORTAM&cluster_previous=NORTAM&country=92&country_previous=92&region=&region_previous=&station=&station_previous=&product=&product_previous=&subproduct=&subproduct_previous=&date="+ dateDownload +"&date_previous=")
        except:
            try:
                driver.get("http://cndreporting.logistics.corp/matrix_ofs/report_drilldown.php?metric=unprinted&category1=Total&category2=Total&cluster=NORTAM&cluster_previous=NORTAM&country=92&country_previous=92&region=&region_previous=&station=&station_previous=&product=&product_previous=&subproduct=&subproduct_previous=&date="+ dateDownload +"&date_previous=")
            except:
                try:
                    driver.get("http://cndreporting.logistics.corp/matrix_ofs/report_drilldown.php?metric=unprinted&category1=Total&category2=Total&cluster=NORTAM&cluster_previous=NORTAM&country=92&country_previous=92&region=&region_previous=&station=&station_previous=&product=&product_previous=&subproduct=&subproduct_previous=&date="+ dateDownload +"&date_previous=")
                except:
                    try:
                        driver.get("http://cndreporting.logistics.corp/matrix_ofs/report_drilldown.php?metric=unprinted&category1=Total&category2=Total&cluster=NORTAM&cluster_previous=NORTAM&country=92&country_previous=92&region=&region_previous=&station=&station_previous=&product=&product_previous=&subproduct=&subproduct_previous=&date="+ dateDownload +"&date_previous=")
                    except:
                        # sys.exit("Unable to load report; Please do check system must be in awake Mode")
                        sys.exit()
                        # msgbox('\n\n\n************************************\nUnable to load report; Please do check system must be in awake Mode\n\n************************************')
        if os.path.exists("C:\\Users\\"+sysUser+"\\Desktop\\Invstat\\Unprinted Archive\\"+"Unprinted.csv"):
            os.remove(cwd+"Unprinted Archive\\"+"Unprinted.csv")
        time.sleep(2)
        driver.find_element_by_id('excelDownload').click()
        shipCount = driver.find_element_by_xpath('/html/body/div/span/b').text
        # Showing records 1 to 1000 of 7511 records
        shipCount = re.findall("of\s(\d+)\srecords", shipCount)
        print("Unprinted Report Downloaded - Shipment Count: ",int(shipCount[0]))

        while not os.path.exists(cwd+"\\Unprinted Archive\\"+"Unprinted.csv"):
            time.sleep(5)
        os.rename(cwd + "\\Unprinted Archive\\"+"Unprinted.csv",cwd + "\\Unprinted Archive\\"+"Unprinted "+ nameDate +".csv")
        driver.quit()
        wb = xw.Book(cwd+"\\Unprinted Archive\\Unprinted "+nameDate+".csv")
    # Delete blank rows
    if flag2 == 1:
        wb = xw.Book(latest_file2)
        print("Using old Unprinted Report: ", latest_file2)
    sht = wb.sheets[0]
    for i in range(1,10):
        if not sht.cells(i,1).value == None:
            break
    if sht.range('A1').value == None:
        sht.range('1:'+str(i-1)).api.Delete()
    # Filter Custono (H) -> 0
    sht.range('A1').api.AutoFilter(Field=8, Criteria1="0", Operator=7)

    sht2 = sh
    sht2.range('A4:C100').clear()

    column = 'N'
    sheetObj = sht
    strrng = sheetObj.range("B2", sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up')).api.SpecialCells(12).Cells(1,1).Row
    endrng = sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up').row
    sheetObj.range(column+str(strrng)+":"+column+str(endrng)).api.Copy()
    time.sleep(1)
    sht2.activate()
    sht2.range('A4').select()
    sht2.range('A4').api.PasteSpecial(12)
    time.sleep(2)
    sht2.range('A2').select()
    sht2.range('A4').value = sht2.range('A4').value
    sht.range('A1').api.AutoFilter(Field=8)
    print("Missing Payer# Filtered")

    # Missing payer name and wrong customer number > 0
    sht.range('A1').api.AutoFilter(Field=10, Criteria1="", Operator=7)
    sht.range('A1').api.AutoFilter(Field=8, Criteria1=">0", Operator=7)
    column = 'N'
    sheetObj = sht
    sheetObj.activate()
    strrng = sheetObj.range("B2", sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up')).api.SpecialCells(12).Cells(1,1).Row
    endrng = sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up').row
    sheetObj.range(column+str(strrng)+":"+column+str(endrng)).api.Copy()

    time.sleep(1)
    sht2.activate()
    lrr = int(sht2.range("A4").api.CurrentRegion.Address.split("$")[-1])
    print(lrr)
    sht2.range('A'+str(lrr+1)).select()
    sht2.activate()
    sht2.range('A'+str(lrr+1)).api.PasteSpecial(12)
    time.sleep(2)
    sht2.range('A2').select()
    sht2.range('A4').value = sht2.range('A4').value
    # Clear cell
    sheetObj.range("H"+str(strrng)+":"+"H"+str(endrng)).clear()
    sht.range('A1').api.AutoFilter(Field=8)
    sht.range('A1').api.AutoFilter(Field=10)

    print("Missing Customer Name Updated where Customer number present")

    # Find missing payer numbers
    # Report Daybookf
    rowCounter = 4
    sht2 = sh
    hbNum = sht2.range("A4").value
    while not hbNum == None:
        con.send_keys('esc')
        con.send_keys('pf3')
        con.wait_for_screen_cursor(22,23)
        con.set_text(value='SEABOOK',row=22, col=23)
        con.send_keys('enter')
        con.wait_for_text('SEA BOOKING',2,32)
        con.wait_for_screen_cursor(2,13)
        time.sleep(1)
        con.set_text(value='L',row=2, col=13)
        con.wait_for_screen_cursor(3,11)
        time.sleep(1)
        con.set_text(value=hbNum,row=3, col=11)
        con.send_keys('enter')
        time.sleep(1)
        con.send_keys('reset')
        time.sleep(1)
        con.send_keys('pf3')
        payerNum = con.get_text(5,48,12)
        payerName = con.get_text(5,61,20)
        if not payerNum[-1] == ' ':
            sht2.range('B'+str(rowCounter)).value = payerNum.strip()
            sht2.range('C'+str(rowCounter)).value = payerName.strip()
        else:
            payerNum = con.get_text(6,48,12)
            payerName = con.get_text(6,61,20)
            if not payerNum[-1] == ' ':
                sht2.range('B'+str(rowCounter)).value = payerNum.strip()
                sht2.range('C'+str(rowCounter)).value = payerName.strip()
        con.send_keys('esc')
        con.send_keys('pf3')
        con.wait_for_screen_cursor(2,13)
        time.sleep(1)
        con.set_text(value='E',row=2, col=13)
        rowCounter = rowCounter + 1
        hbNum = sht2.range("A"+str(rowCounter)).value
        con.wait_for_screen_cursor(22,23)
        payerNum = ''
        payerName = ''

    time.sleep(5)



    # Update missing payer and number

    lrU = sht.cells(sht.range('A:A').rows.count, 1).end('up').row
    bok = sht.range('N2:N'+str(lrU)).value
    # Updating missing booking Number based on blank and '***NOT LOADED***'
    sht_re = temp.sheets[1]
    reg = sht_re.range('A2:A'+sht_re.range("A1").api.CurrentRegion.Address.split("$")[-1]).value
    reg_bok = sht_re.range('B2:B'+sht_re.range("A1").api.CurrentRegion.Address.split("$")[-1]).value
    regno = sht.range('D2:D'+str(lrU)).value
    regBokDict = vDict(reg, reg_bok) #dict(zip(reg, reg_bok))
    updateBok = []
    for i, j in zip(bok, regno):
        if (i == None) or (i == '') or (i == "***NOT LOADED***"):
            if j in regBokDict:
                updateBok.append(regBokDict[j])
            else:
                updateBok.append(i)
        else:
            updateBok.append(i)
    sht.range('N2:N'+str(lrU)).options(transpose=True).value = updateBok

    # Update missing payer and number
    missBook = sht2.range('A4:A100').value
    missNum = sht2.range('B4:B100').value
    missName = sht2.range('C4:C100').value

    dictNum = vDict(missBook, missNum)#dict(zip(missBook, missNum))
    dictName = vDict(missBook, missName)#dict(zip(missBook, missName))

    numU = sht.range('H2:H'+str(lrU)).value
    nameU = sht.range('J2:J'+str(lrU)).value
    dictUnum = {}
    dictUnum = vDict(bok, numU)#dict(zip(bok, numU))
    dictUname = vDict(bok, nameU)#dict(zip(bok, nameU))

    updateUnum = []
    for i, j in zip(bok,numU):
        if (i in dictNum) and (not i == None):
            if j == None or j == 0 or j == '':
                updateUnum.append(dictNum[i])
            else:
                updateUnum.append(j)
        else:
            updateUnum.append(j)

    updateUname = []
    for i, j in zip(bok, nameU):
        if (i in dictName) and (not i == None):
            updateUname.append(dictName[i])
        else:
            updateUname.append(j)

    sht.range('H2:H'+str(lrU)).options(transpose=True).value = updateUnum
    sht.range('J2:J'+str(lrU)).options(transpose=True).value = updateUname
    print("Update missing payer and number")
    # Run Invsta Macro
    # region = sht.range('A1').api.CurrentRegion.Address
    # rawData = sht.range('$A$2:'+region.split(':')[1]).value
    if os.path.exists(cwd+"\\Invstat123 Archive\\"+"INVSTAT_FULL "+datetime.datetime.now().strftime("%d-%b-%Y")+".xlsb"):
        os.rename(cwd+"\\Invstat123 Archive\\"+"INVSTAT_FULL "+datetime.datetime.now().strftime("%d-%b-%Y")+".xlsb", cwd+"\\Invstat123 Archive\\"+"INVSTAT_FULL "+nameDate+".xlsb")

    sht.activate()
    sht.range("A1").api.Select()
    sht.range("A1").api.CurrentRegion.Copy()
    mb = xw.Book(cwd+"\\Macro Files\\"+"InvstatMacro.xlsm")
    msht = mb.sheets('INVSTAT')
    msht.activate()
    msht.range('A2').api.Select()
    msht.api.PasteSpecial(12)
    time.sleep(6)
    # msht.range('A2').value = rawData
    msht.range("2:2").api.Delete()
    time.sleep(2)
    mb.macro('InvstatMidMacro.invstatMidMacro').run()
    mb.close()
    wb.close()
    print("Invstat Macro : Completed")
    # Scope, group name and hawb# updation
    wb = xw.Book(cwd+"\\Invstat123 Archive\\"+"INVSTAT_FULL "+datetime.datetime.now().strftime("%d-%b-%Y")+".xlsb")
    sht = wb.sheets('INVSTAT')
    time.sleep(4)

    sht.range('A1').value = '#'

    lc = sht.cells(1, sht.range('1:1').columns.count).end('left').column
    lr = sht.cells(sht.range('A:A').rows.count, 1).end('up').row
    OneIDconca = sht.range("BU2:BU"+str(lr)).value
    OneIDonly = sht.range("H2:H"+str(lr)).value
    obj = wc.Dispatch('excel.application')
    obj.AskToUpdateLinks = False
    obj.DisplayAlerts = False
    scopeName = sht2.range('G4').value
    scopeSheet = sht2.range('I4').value
    print(cwd,'\\HCL Scope Archive\\',scopeName)
    excel = obj.Workbooks.Open(cwd +"\\HCL Scope Archive\\"+scopeName)
    scopeFile = xw.Book(cwd+"\\HCL Scope Archive\\"+scopeName)

    shtSc = scopeFile.sheets(scopeSheet)

    lrS = shtSc.cells(shtSc.range('A:A').rows.count, 1).end('up').row
    oneIDconcaSc = shtSc.range('A2:A'+str(lrS)).value
    temOneId = []
    for x in oneIDconcaSc:
        if 'str' in str(type(x)):
            temOneId.append(x.upper())
        else:
            temOneId.append(x)

    oneIDconcaSc =  temOneId
    temOneId = None
    oneIDonlySc = shtSc.range('B2:B'+str(lrS)).value
    OwnerSc = shtSc.range('E2:E'+str(lrS)).value
    GroupSc = shtSc.range('AB2:AB'+str(lrS)).value
    dictScope = {}
    groupScope = {}
    for i, j in zip(oneIDconcaSc, OwnerSc):
        if i in dictScope:
            pass
        else:
            if 'str' in str(type(i)):
                dictScope[i.upper()] = j
            else:
                dictScope[i] = j

    for i, j in zip(oneIDonlySc, GroupSc):
        if i in groupScope:
            pass
        else:
            if 'str' in str(type(i)):
                groupScope[i.upper()] = j
            else:
                groupScope[i] = j

    scopeArray = []
    for i in OneIDconca:
        if i in dictScope:
            scopeArray.append(dictScope[i])
        else:
            scopeArray.append('Need to verify')

    groupArray = []
    for i in OneIDonly:
        if i in groupScope:
            groupArray.append(groupScope[i])
        else:
            groupArray.append('')

    sht.range('BW2').options(transpose=True).value = scopeArray
    time.sleep(1)
    scope2Array = []
    for i in scopeArray:
        if i.title() == 'Chennai':
            scope2Array.append('HCL')
        elif i.title() == 'Pune':
            scope2Array.append('HCL')
        elif i.title() == '':
            scope2Array.append('Need to verify')
        else:
            scope2Array.append(i.title())

    sht.range('BY2').options(transpose=True).value = groupArray
    sht.range('BX2').options(transpose=True).value = scope2Array
    scopeFile.close()

    # excel.Close()

    groupName = sht.range('BY2:BY'+str(lr)).value
    customerName = sht.range('J2:J'+str(lr)).value
    updateGroup = []
    for i,j in zip(groupName, customerName):
        if not i == None or i == "":
            updateGroup.append(i)
        else:
            updateGroup.append(j)

    sht.range('BY2:BY'+str(lr)).options(transpose=True).value = updateGroup
    print("Scope and Group Name Updated")
    bkhawb = xw.Book(cwd+"\\Daybookf Archive\\"+"BookingwithHawbCompleteArchive.xlsb")
    # sht1 = bkhawb.sheets('Other')
    # sht2 = bkhawb.sheets('Mar,Apr&May')
    sht1 = bkhawb.sheets[0]
    sht2 = bkhawb.sheets[1]
# Question

    def create_sheets_array(workBook_obj, col):
        sht_count = workBook_obj.sheets.count
        listArray = []
        for s in range(0, sht_count):
            asht = workBook_obj.sheets[s]
            region = asht.range('A1').api.CurrentRegion.Address.split('$')[-1]
            if not region == '1':
                data1 = asht.range(col + '2:' + col + region).value
                listArray.extend(data1)
        return(listArray)


    def lookup_array(lookupArray, ValueDict):
        result_array = []
        for i in lookupArray:
            if i in ValueDict:
                result_array.append(ValueDict[i])
            else:
                result_array.append('')
        return(result_array)

    bok_num = create_sheets_array(bkhawb, "A")
    hawb_num = create_sheets_array(bkhawb, "D")
    bok_hawb_dict = vDict(bok_num, hawb_num)
    bok_target = sht.range("N2:N"+str(lr)).value
    updatehb = lookup_array(bok_target, bok_hawb_dict)

    matdrd = create_sheets_array(bkhawb, "I")
    matdrd_dict = vDict(bok_num, matdrd)
    matd = lookup_array(bok_target, matdrd_dict)

    sht.range("BZ2:BZ"+str(lr)).options(transpose=True).value = updatehb
    sht.range("CA2:CA"+str(lr)).options(transpose=True).value = matd
    print("Hawb# Updated")
    shtP = wb.sheets('Summary')
    shtP.activate()
    # Pivote hiding
    wb.api.RefreshAll()
    time.sleep(2)
    wb.sheets[0].api.PivotTables("PivotTable2").PivotFields("Scope").ShowDetail = False
    wb.save()
    time.sleep(3)
    bkhawb.close()
    temp.save()
    temp.close()
    # wb.close()
    # Macro end time
    try:
        con.close()
    except:
        pass


    try:
        excel.Close()
    except:
        pass
    et = datetime.datetime.now()
    ttm = (et - st).total_seconds()
    tt = 'Invstat123-Ceva Dashboard: Macro ran successfully.\nTime taken to Complete: ' + str(ttm) +" Seconds"
    print(tt)

    count = 1 # Number of transactions count
    initialaht = 2400 # Time in seconds for completetion or per transaction
    process = "CEVA"
    strFileFullName = "Invstat123 Reporting Macro" # Macro  name
    modulename = "IMS-Invstat123 Full Report Automation" # Department-Project_name
    insert_timestamp(st_new, initialaht, process, strFileFullName, modulename, count)
    time.sleep(20)
    # msgbox(tt)

except:
    logging.basicConfig(format='%(asctime)s %(message)s',datefmt='%m/%d/%Y %I:%M:%S %p',level=logging.DEBUG, filename= cwd+'\\Exception.txt')
    logging.exception('***********************************************' * 2)
    time.sleep(20)
    # msg = "Error Ocurred: Please check exception in "+ cwd+'\\Exception.txt'
    try:
        con.close()
    except:
        pass
    # msgbox(msg)
