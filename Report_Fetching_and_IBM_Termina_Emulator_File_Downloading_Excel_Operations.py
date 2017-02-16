from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import urllib.request
import zipfile
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
import logging

cwd = os.getcwd()
sysUser = os.getlogin()
print(sysUser)
cwd = r'C:\Users\vijaykumar.mane\Desktop\Invstat'
dateDownload = datetime.datetime.now().strftime("%Y-%m-%d")
nameDate = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M")
nameDate = '2017-02-15-12-34'

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
    temp = xw.Book(cwd+'\\Macro Files\\InputFile.xlsm')
    sh = temp.sheets('Sheet1')

    # Ceva Dashboard Credentials
    username = sh.range('G8').value
    password = sh.range('G9').value
    time.sleep(1)
    # Generate file ID from IBM Terminal Emulator

    userOFS = sh.range('G5').value
    passOFS = sh.range('G6').value
    print(userOFS,' ', passOFS)
    sessionName = 'A'
    profileName = 'OFS-live.WS'
    try:

        con.wait_for_text('Select Records',1,34)
        con = iseries(uid=userOFS, pwd=passOFS, session=sessionName, profile=profileName)
        con.connect()
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
        con.wait_for_screen(20,7)
        con.set_text(value='1',row=20, col=7)
        con.send_keys('enter')
        # Select region
        con.wait_for_text('USA',6,11)
        con.wait_for_screen(20,7)
        con.set_text(value='1',row=20, col=7)
        con.send_keys('enter')
        time.sleep(2)
        # Report Daybookf

        con.wait_for_text('Forwarding',4,20)
        con.set_text(value='DAYBOOKF',row=22, col=23)
        con.send_keys('enter')
        # Modify query
        end_date = datetime.datetime.now().strftime("%Y%m%d")
        start_date = (datetime.datetime.now() + datetime.timedelta(-30)).strftime("%Y%m%d")
        con.set_text(value=start_date+' '+end_date ,row=7, col=35)
        time.sleep(1)
        con.set_text(value="'L' 'G' 'N' 'B' 'E' 'K'" ,row=12, col=35)
        con.send_keys('enter')
        time.sleep(5)
        # Loop while to get fileID
        while 'Query running' in con.get_rect_text(24,2,24,60):
            time.sleep(2)
            print(con.get_rect_text(24,2,24,60))
        # Get file ID
        con.wait_for_text('Forwarding',4,20)
        fileID = con.get_rect_text(24,2,24,60)
        fileID = re.findall('DAYBOOKF\s+in\s+(\w+)\s+was',fileID)
        if len(fileID) > 0:
            downloadId = fileID[0]
        else:
            msgbox('Query Failed.\n Please run again')

    except:
        msgbox('Error Ocurred, please confirm\n1.Please close IBM emulator if open\n2.Create IBM Emulator Profile as "OFS-live.WS" and run again')
        sys.exit()
    # OFS Close

    print(downloadId)
    if downloadId is None:
        msgbox('Daybookf report not Generated Please run it again')
        sys.exist()

    # Download file from IBM iSeries
    path1 = cwd+'\\Daybookf Archive\\'
    ip='10.235.108.20'
    uid= userOFS
    pwd= passOFS
    path=path1+"Daybookf "+nameDate+'.xls'
    fname= downloadId +'/'+'DAYBOOKF('+uid.upper()+')'
    # iseries.iSeries_download(ip, uid, pwd, fname, path)

    print('Daybookf Downloaded')
    time.sleep(5)
    df = xw.Book(path)
    wbt = xw.Book()
    tsht = wbt.sheets[0]
    flag = None
    sCount = df.sheets.count - 1
    while sCount >= 0:
        shtt = df.sheets[sCount]
        region2 = tsht.range('A1').api.CurrentRegion.Address
        if flag == None:
            shtt.range('$A$1').api.CurrentRegion.Copy()
            flag = 1
            tsht.api.Activate()
            tsht.range('A1').api.Select()
            tsht.api.Paste()
        else:
            shtt.range("1:1").api.Delete()
            shtt.range('$A$1').api.CurrentRegion.Copy()
            d = str(int(region2.split('$')[-1]) + 1)
            tsht.activate()
            tsht.range('A'+d).api.Select()
            tsht.api.Paste()
        sCount -= 1
    tsht.range('A1').api.Select()
    dfsht = wbt.sheets[0]
    dfa = xw.Book(cwd+'\\Daybookf Archive\\'+'BookingwithHawbCompleteArchive.xlsb')

    dfasht1 = dfa.sheets[0]
    dfasht2 = dfa.sheets[1]

    # Last row function
    def lr(shtO):
        lrN = shtO.cells(shtO.range('A:A').rows.count, 1).end('up').row
        return(lrN)

    lrDf = lr(dfsht)

    today = datetime.datetime.now()
    DD = datetime.timedelta(days=30)
    earlier = today - DD
    earlier_str = earlier.strftime("%Y%m%d")
    dfasht1.range('A1').api.AutoFilter(Field=7, Criteria1='>='+ earlier_str, Operator=7)
    time.sleep(2)
    column = 'G'
    sheetObj = dfasht1
    strrng = sheetObj.range("B2", sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up')).api.SpecialCells(12).Cells(1,1).Row
    endrng = sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up').row
    sheetObj.range(str(strrng+1)+":"+str(endrng)).api.Delete()
    time.sleep(5)
    dfasht1.range('A1').api.AutoFilter(Field=7)
    time.sleep(5)

    lrDfasht1 = lr(dfasht1)
    # Create Date
    createDate_df = dfsht.range('A2:A'+str(lrDf)).value
    dfasht1.range('G'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df

    # Create Date
    createDate_df = dfsht.range('F2:F'+str(lrDf)).value
    dfasht1.range('A'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
    # CSORNADR1
    createDate_df = dfsht.range('X2:X'+str(lrDf)).value
    dfasht1.range('B'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df

    # CSORNO
    createDate_df1 = dfsht.range('Y2:Y'+str(lrDf)).value
    # HBLNO
    createDate_df2 = dfsht.range('BZ2:BZ'+str(lrDf)).value
    # HOUSENO
    createDate_df3 = dfsht.range('Y2:Y'+str(lrDf)).value
    createDate_df4 = []
    # Merging of columns
    print(len(createDate_df1),' ',len(createDate_df1),' ',len(createDate_df1),' ')
    for i,j,k in zip(createDate_df1, createDate_df2, createDate_df3):
        if i is None:
            if j is not None:
                createDate_df4.append(j)
            else:
                createDate_df4.append(k)
        else:
            createDate_df4.append(i)
    dfasht1.range('C'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df4
    # HAWBNO
    createDate_df = dfsht.range('AA2:AA'+str(lrDf)).value
    dfasht1.range('D'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
    # AIRSTATUS
    createDate_df = dfsht.range('BU2:BU'+str(lrDf)).value
    dfasht1.range('E'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
    # FWDSTATUS
    createDate_df = dfsht.range('BV2:BV'+str(lrDf)).value
    dfasht1.range('F'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
    # MATRCDATE
    createDate_df = dfsht.range('L2:L'+str(lrDf)).value
    dfasht1.range('H'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
    # MATRDDATE
    createDate_df = dfsht.range('P2:P'+str(lrDf)).value
    dfasht1.range('I'+str(lrDfasht1+1)).options(transpose=True).value = createDate_df
    # Save Archive Daybookf
    dfa.save()
    dfa.close()
    df.close()
    wbt.close()
    # File Download path

    chromeOptions = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : cwd+"\\Unprinted Archive\\"}
    chromeOptions.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(
        "C:\\Users\\" + sysUser + "\\Downloads\\chromedriver.exe", chrome_options=chromeOptions)
    driver.maximize_window()
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
            driver.get("http://cndreporting.logistics.corp/matrix_ofs/report_drilldown.php?metric=unprinted&category1=Total&category2=Total&cluster=NORTAM&cluster_previous=NORTAM&country=92&country_previous=92&region=&region_previous=&station=&station_previous=&product=&product_previous=&subproduct=&subproduct_previous=&date="+ dateDownload +"&date_previous=")
            msgbox('\n\n\n************************************\nUnable to load report; Please do check system must be in awake Mode\n\n************************************')
    if os.path.exists("C:\\Users\\"+sysUser+"\\Desktop\\Invstat\\Unprinted Archive\\"+"Unprinted.csv"):
        os.remove(cwd+"Unprinted Archive\\"+"Unprinted.csv")
    time.sleep(2)
    driver.find_element_by_id('excelDownload').click()
    shipCount = driver.find_element_by_xpath('/html/body/div/span/b').text
    # Showing records 1 to 1000 of 7511 records
    shipCount = re.findall("of\s(\d+)\srecords", shipCount)
    #print("Unprinted Report Downloaded - Shipment Count: ",int(shipCount[0]))

    while not os.path.exists(cwd+"\\Unprinted Archive\\"+"Unprinted.csv"):
        time.sleep(5)
    os.rename(cwd + "\\Unprinted Archive\\"+"Unprinted.csv",cwd + "\\Unprinted Archive\\"+"Unprinted "+ nameDate +".csv")
    driver.quit()

    # Delete blank rows
    wb = xw.Book(cwd+"\\Unprinted Archive\\Unprinted "+nameDate+".csv")
    sht = wb.sheets[0]
    for i in range(1,10):
        if sht.cells(i,1).value is not None:
            break
    if sht.range('A1').value is None:
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
    sht2.range('A4').api.PasteSpecial(12)
    sht2.activate()
    sht2.range('A4').select()
    time.sleep(2)
    sht2.range('A4').value = sht2.range('A4').value
    print("Missing Payer# Filtered")

    # Find missing payer numbers
    # Report Daybookf
    rowCounter = 4
    sht2 = sh
    hbNum = sht2.range("A4").value
    while hbNum is not None:
        con.send_keys('esc')
        con.send_keys('pf3')
        con.wait_for_screen(22,23)
        con.set_text(value='SEABOOK',row=22, col=23)
        con.send_keys('enter')
        con.wait_for_text('SEA BOOKING',2,32)
        con.wait_for_screen(2,13)
        time.sleep(1)
        con.set_text(value='L',row=2, col=13)
        con.wait_for_screen(3,11)
        con.set_text(value=hbNum,row=3, col=11)
        con.send_keys('enter')
        time.sleep(1)
        con.send_keys('reset')
        time.sleep(1)
        con.send_keys('pf3')
        payerNum = con.get_text(5,48,12)
        payerName = con.get_text(5,61,20)
        if payerNum[-1] is not ' ':
            sht2.range('B'+str(rowCounter)).value = payerNum.strip()
            sht2.range('C'+str(rowCounter)).value = payerName.strip()
        else:
            payerNum = con.get_text(6,48,12)
            payerName = con.get_text(6,61,20)
            if payerNum[-1] is not ' ':
                sht2.range('B'+str(rowCounter)).value = payerNum.strip()
                sht2.range('C'+str(rowCounter)).value = payerName.strip()
        con.send_keys('esc')
        con.send_keys('pf3')
        con.wait_for_screen(2,13)
        con.set_text(value='E',row=2, col=13)
        rowCounter = rowCounter + 1
        hbNum = sht2.range("A"+str(rowCounter)).value
        con.wait_for_screen(22,23)
        payerNum = ''
        payerName = ''

    time.sleep(5)
    missBook = sht2.range('A4:A100').value
    missNum = sht2.range('A4:A100').value
    missName = sht2.range('A4:A100').value
    dictNum = {}
    for i, j in zip(missBook, missNum):
        if i is not None:
            dictNum[i] = j

    dictName = {}
    for i, j in zip(missBook, missName):
        if i is not None:
            dictNum[i] = j

    sht.range('A1').api.AutoFilter(Field=8)

    lrU = sht.cells(sht.range('A:A').rows.count, 1).end('up').row
    bok = sht.range('N2:N'+str(lrU)).value
    numU = sht.range('H2:H'+str(lrU)).value
    nameU = sht.range('J2:J'+str(lrU)).value

    dictUnum = {}
    for i,j in zip(bok, numU):
        dictNum[i] = j
    dictUname = {}
    for i, j in zip(bok, nameU):
        dictUname[i] = j

    updateUnum = []
    for i in dictUnum:
        if i in dictNum:
            updateUnum.append(dictNum[i])
        else:
            updateUnum.append(i)

    updateUname = []
    for i in dictUname:
        if i in dictName:
            updateUnum.append(dictName[i])
        else:
            updateUnum.append(i)

    sht.range('H2:H'+str(lrU)).options(transpose=True).value = updateUnum
    sht.range('J2:J'+str(lrU)).options(transpose=True).value = updateUname
    # Run Invsta Macro
    region = sht.range('A1').api.CurrentRegion.Address
    rawData = sht.range('$A$2:'+region.split(':')[1]).value

    mb = xw.Book(cwd+"\\Macro Files\\"+"InvstatMacro.xlsm")
    msht = mb.sheets('INVSTAT')
    msht.range('A2').value = rawData
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
    scopeName = sht2.range('G4').value
    scopeSheet = sht2.range('I4').value
    print('Scope Name',cwd ,scopeName)
    excel = obj.Workbooks.Open(cwd +"\\HCL Scope Archive\\"+scopeName)
    scopeFile = xw.Book(cwd+"\\HCL Scope Archive\\"+scopeName)

    shtSc = scopeFile.sheets(scopeSheet)

    lrS = shtSc.cells(shtSc.range('A:A').rows.count, 1).end('up').row
    oneIDconcaSc = shtSc.range('A2:A'+str(lrS)).value
    oneIDonlySc = shtSc.range('A2:A'+str(lrS)).value
    OwnerSc = shtSc.range('E2:E'+str(lrS)).value
    GroupSc = shtSc.range('AB2:AB'+str(lrS)).value
    dictScope = {}
    groupScope = {}
    for i, j in zip(oneIDconcaSc, OwnerSc):
        dictScope[i] = j

    for i, j in zip(oneIDonlySc, GroupSc):
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

    scope2Array = []
    for i in scopeArray:
        if i == 'Chennai':
            scope2Array.append('HCL')
        elif i == 'Pune':
            scope2Array.append('HCL')
        elif i is '':
            scope2Array.append('Need to verify')
        else:
            scope2Array.append(i)

    sht.range('BY2').options(transpose=True).value = groupArray
    sht.range('BX2').options(transpose=True).value = scope2Array
    scopeFile.close()
    # excel.Close()

    groupName = sht.range('BY2:BY'+str(lr)).value
    customerName = sht.range('J2:J'+str(lr)).value
    updateGroup = []
    for i,j in zip(groupName, customerName):
        if i is not None:
            updateGroup.append(i)
        else:
            updateGroup.append(j)

    sht.range('BY2:BY'+str(lr)).options(transpose=True).value = updateGroup
    print("Scope and Group Name Updated")
    bkhawb = xw.Book(path1+"BookingwithHawbCompleteArchive.xlsb")
    sht1 = bkhawb.sheets('Other')
    sht2 = bkhawb.sheets('Mar,Apr&May')

    def vFunction(vSheet, vCol, tSheet, tCol1, tCol2, dCol):
        vLr = vSheet.cells(vSheet.range('A:A').rows.count, 1).end('up').row
        tLr = tSheet.cells(tSheet.range('A:A').rows.count, 1).end('up').row
        vArray = vSheet.range(vCol+"2:"+vCol+str(vLr)).value
        vtArray = tSheet.range(tCol1+"2:"+tCol1+str(tLr)).value
        ttArray = tSheet.range(tCol2+"2:"+tCol2+str(tLr)).value
        tDict = {}
        for i, j in zip(vtArray, ttArray):
            tDict[i] = j
        vlookArray = []
        for i in vArray:
            if i in tDict:
                vlookArray.append(tDict[i])
            else:
                vlookArray.append('')
        vSheet.range(dCol+"2:"+dCol+str(vLr)).options(transpose=True).value = vlookArray

    vFunction(sht,'N',sht1,'A','D','BZ')
    vFunction(sht,'N',sht2,'A','D','CA')

    hb1 = sht.range('BZ2:BZ'+str(lr)).value
    hb2 = sht.range('CA2:CA'+str(lr)).value
    updatehb = []

    for i,j in zip(hb1, hb2):
        if i is None:
            updatehb.append(j)
        else:
            updatehb.append(i)

    sht.range("BZ2:BZ"+str(lr)).options(transpose=True).value = updatehb
    sht.range("CA2:CA"+str(lr)).options(transpose=True).value = None
    print("Hawb# Updated")
    shtP = wb.sheets('Summary')
    shtP.activate()
    # Pivote hiding
    shtP.api.PivotTables("PivotTable2").PivotFields("Scope").ShowDetail = False
    wb.api.RefreshAll()
    wb.save()
    time.sleep(3)
    bkhawb.close()
    temp.close()
    msgbox('Invstat123-Ceva Dashboard: Macro ran successfully')

except:
    logging.basicConfig(format='%(asctime)s %(message)s',datefmt='%m/%d/%Y %I:%M:%S %p',level=logging.DEBUG, filename= cwd+'\\Exception.txt')
    logging.exception('***********************************************' * 2)
