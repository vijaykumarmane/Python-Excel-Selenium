import urllib.request
from selenium import webdriver
import stat
import time
import pandas as pd
import requests
import zipfile
import sys
import shutil
import re
import datetime
import glob
import os
import xlwings as xw
from ibm import iseries
from Kaizen_timestamp import insert_timestamp

sysUser = os.getlogin()
cwd = os.getcwd()
cwd = r"C:\Users\vijaykumar.mane\Desktop\TPB Report"

st = time.gmtime()
input_wb = xw.Book(cwd+"\\Input Files\\"+"InputFileTPB.xlsx")
in_sht = input_wb.sheets[0]
# driver = webdriver.Firefox()
username = in_sht.range('B3').value
password = in_sht.range('B4').value
signonwindow = in_sht.range('B8').value
dateDownload = datetime.datetime.now().strftime("%Y-%m-%d")
nameDate = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
bookingWithHawb_path = in_sht.range('B7').value


def vDict(keyValue, valueValue):
    tempDict = {}
    for i, j in zip(keyValue, valueValue):
        if not i in tempDict:
            tempDict[i] = j
    return(tempDict)

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

# Age of file
def file_age_in_seconds(pathname):
    return time.time() - os.stat(pathname)[stat.ST_MTIME]

list_of_files = glob.glob(cwd + '\\Raw Data Archive\\*.xls')
if len(list_of_files) == 0:
    flag = 0
else:
    latest_file2 = max(list_of_files, key=os.path.getctime)
    age3 = file_age_in_seconds(latest_file2)
    if age3 < 10800:
        flag = 1
        new_file = latest_file2
        print("Old TPB Raw File to Use: ", latest_file2)
    else:
        flag = 0
    # File Download path


if flag == 0:
    chromeOptions = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : cwd + "\\Raw Data Archive"}
    chromeOptions.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(
        "C:\\Users\\" + sysUser + "\\Downloads\\chromedriver.exe", chrome_options=chromeOptions)
    # driver.maximize_window()
    driver.get("http://querybook.live.invoize.com/Default.aspx")
    driver.find_element_by_id('LoginMain_UserName').send_keys(username)
    driver.find_element_by_id('LoginMain_Password').send_keys(password)
    driver.find_element_by_id('LoginMain_LoginButton').click()
    driver.get("http://querybook.live.invoize.com/Edit.aspx?qid=714")
    driver.find_element_by_id('ExportToExcel').click()
    time.sleep(45)
    list_of_files = glob.glob(cwd + '\\Raw Data Archive\\*.crdownload')
    while len(list_of_files) > 0:
        list_of_files = list_of_files = glob.glob(cwd + '\\Raw Data Archive\\*.crdownload')
        time.sleep(10)
    time.sleep(5)
    driver.quit()
    list_of_files = glob.glob(cwd + '\\Raw Data Archive\\*.xls')
    new_file = max(list_of_files, key=os.path.getctime)
    print("New TPB Raw Report: ", new_file)
# # Test
# new_file = r'C:\Users\vijaykumar.mane\Desktop\TPBReport\Raw Data Archive\TPB Report_17-Feb-2017_160253.xls'

# Modify Raw Report
wbr = xw.Book(new_file)
shtr = wbr.sheets[0]
# Insert Column to Right
shtr.range('P:P').api.Insert(Shift=-4161)
shtr.range('P:P').api.Insert(Shift=-4161)
shtr.range('P1').value = 'Conversion Rate'
shtr.range('Q1').value = 'Amount in USD'
shtr.range('W:W').api.Insert(Shift=-4161)
shtr.range('W:W').api.Insert(Shift=-4161)
shtr.range('W1').value = 'Conca'
shtr.range('X1').value = 'HBL Count'
print('Columns inserted')
# Updating Currency conversion rates
cwbook = xw.Book(cwd +'\\Conversion Rate Archive\\'+"Conversion rates.xlsx")
csh = cwbook.sheets[0]
c_lr = (csh.range('A1').api.CurrentRegion.Address).split('$')[-1]
code = csh.range("A2:A"+c_lr).value
rate = csh.range("B2:B"+c_lr).value
c_table = vDict(code, rate)#dict(zip(code, rate))
r_lr = (shtr.range('A1').api.CurrentRegion.Address).split('$')[-1]
v_col = shtr.range('N2:N'+r_lr).value
print('Conversion Rate Updated')
cwbook.close()

vl_value = []
for i in v_col:
    if i in c_table:
        vl_value.append(c_table[i])
    elif not i ==  None:
        url = ('https://currency-api.appspot.com/api/%s/USD.json') % (i.upper())
        r = requests.get(url)
        vl_value.append(r.json()['rate'])
        print('Missing Currency in conversion rate sheet: ', i.upper(), "--> USD", r.json()['rate'], "Updated")
    else:
        vl_value.append('')
shtr.range('P2:P'+r_lr).options(transpose=True).value = vl_value
amount = shtr.range('O2:O'+r_lr).value
total = []
for i, j in zip(vl_value,amount):
    if not i == None and not j == None:
        try:
            total.append(i*j)
        except:
            total.append("")
    else:
        total.append("")
shtr.range('Q2:Q'+r_lr).options(transpose=True).value = total
hawb = shtr.range('V2:V'+r_lr).value
booking = shtr.range('R2:R'+r_lr).value
conca = []
for i, j, k in zip(booking, hawb, v_col):
    if type(j) == float:
        j = str(int(j))
    conca.append(str(i)+str(j)+str(k))
shtr.range('W2:W'+r_lr).options(transpose=True).value = conca
dupeDict = {}
dupeStatus = []
for i in conca:
    if not i in dupeDict:
        dupeDict[i] = 1
        dupeStatus.append('1')
    else:
        dupeStatus.append('0')
shtr.range('X2').options(transpose=True).value = dupeStatus
shtr.range('BH1').value = 'Hawb Present in Invstat'
print('Total USD, Conca and HB Count Updated')

userOFS = in_sht.range('B5').value
passOFS = in_sht.range('B6').value
print(userOFS,' ', passOFS)
sessionName = in_sht.range('F3').value #'B'
profileName = in_sht.range('F4').value #'OFS-live.WS'

# Date range creating Function


def date_range_array(startDate, endDate, partitions, formatDate):
    dates_array = []
    dates_array.append(startDate.strftime(formatDate))
    delta = (startDate - endDate).days//partitions
    while startDate <= endDate:
        startDate = startDate - datetime.timedelta(delta)
        if startDate > endDate:
            break
        dates_array.append(startDate.strftime(formatDate))
        startDate = startDate - datetime.timedelta(-1)
        dates_array.append(startDate.strftime(formatDate))
    if startDate > endDate:
        dates_array.append(endDate.strftime(formatDate))
    return(dates_array)

yer = int(in_sht.range('B10').value)
mon = int(in_sht.range('C10').value)
da = int(in_sht.range('D10').value)
par = int(in_sht.range('H10').value)
dateRange = date_range_array(datetime.date(yer,mon,da), datetime.datetime.now().date(), par, "%Y%m%d")
LoopCounter = int(len(dateRange)/2)

list_of_files2 = glob.glob(cwd+"\\Full Invstat Archive\\*.xlsb")
latest_invstat = max(list_of_files2, key=os.path.getctime)
age_inv = file_age_in_seconds(latest_invstat)
if age_inv < 10800:
    flag3 = 0
else:
    flag3 = 1

print(flag3)

if flag3 == 1:
    try:
        shutil.rmtree(cwd+'\\Full Invstat Archive\\Archive')
        os.makedirs(cwd+'\\Full Invstat Archive\\Archive')
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
    except:
        # msgbox('Error Ocurred, please confirm\n1.Please close IBM emulator if open\n2.Create IBM Emulator Profile as "OFS-live.WS" and run again')
        sys.exit()

    stdcomplete1 = datetime.datetime.now()
    stdc = datetime.datetime.now()
    # Report Daybookf
    for i in range(0, LoopCounter):
        # con.wait_for_text('Forwarding',4,20)
        time.sleep(1)
        con.set_text(value='INVSTAT123',row=22, col=23)
        con.send_keys('enter')
        time.sleep(1)

        if 'unmonitored by' in con.get_rect_text(24,2,24,60):
            # msgbox("Function check. QRY5080 unmonitored by BASUXCMD at statement 0000024800, ins\nPlease check might be other using same Credentials to run query.")
            sys.exit()

        # Modify query
        con.wait_for_text('Select Records',1,34)
        start_date = dateRange[i*2]
        end_date = dateRange[(i*2)+1]
        print('Report Running for date range: ',start_date," ",end_date)
        fileTimeStamp = []
        nameDate = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        fileTimeStamp.append(nameDate)
        time.sleep(1)
        con.set_text(value=start_date+' '+end_date ,row=7, col=35)
        time.sleep(1)
        con.set_text(value="RANGE" ,row=8, col=28)
        time.sleep(1)
        con.set_text(value="1 8      " ,row=8, col=35)
        con.send_keys('enter')
        time.sleep(5)

        # Loop while to get fileID
        while 'Query running' in con.get_rect_text(24,2,24,60) and not 'unmonitored' in con.get_rect_text(24,2,24,60):
            if 'unmonitored by' in con.get_rect_text(24,2,24,60):
                sys.exit()
            time.sleep(2)
            # print(i, '/',par,con.get_rect_text(24,2,24,60))
        # Get file ID
        con.wait_for_text('Forwarding',4,20)
        fileID = con.get_rect_text(24,2,24,60)
        fileID = re.findall('INVSTAT123\s+in\s+(\w+)\s+was',fileID)
        if len(fileID) > 0:
            downloadId = fileID[0]
        else:
            # msgbox('Query Failed.\n Please run again')
            pass

        print(i+1,"/",par," -->  ",downloadId)
        if downloadId == None:
            # msgbox('Daybookf report not Generated Please run it again')
            sys.exist()

        # Download file from IBM iSeries
        path1 = cwd+'\\Full Invstat Archive\\'
        ip='10.235.108.20'
        uid= userOFS
        pwd= passOFS
        path=path1+"Archive\\"+"Part Range Invstat "+nameDate+'.xls'
        fname= downloadId +'/'+'INVSTAT123('+uid.upper()+')'
        std = datetime.datetime.now()
        iseries.iSeries_download(ip, uid, pwd, fname, path)
        etd = datetime.datetime.now()
        ttstd = (etd - std).total_seconds()
        print(i+1,"/",par,' Full Invstat Report Downloaded in : ',str(ttstd) ,'\n',path)
        # time.sleep(5)
    time.sleep(60)
    etdc = datetime.datetime.now()
    ttstdc = (etdc - stdc).total_seconds()
    print("*******Complete time taken to download Invstat: Seconds : ",ttstdc)
    listFiles = glob.glob(cwd +'\\Full Invstat Archive\\Archive\\*')
    print(listFiles)
    print(len(listFiles)," : Total Files\n",'\n'.join(listFiles))


    def deleteFilter(sheet1):
        # for k in range(0, counter):
        # sheetObj = BooktoConsolidate_in_Wb_Obj.sheets[k]
        sheetObj = sheet1
        region = sheetObj.range("A1").api.CurrentRegion.Address.split("$")[-1]
        if region == "1":
            return
        sheetObj.range('A1').api.AutoFilter(Field=9, Criteria1="40", Operator=7)
        strrng = sheetObj.range("B2", sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up')).api.SpecialCells(12).Cells(1,1).Row
        endrng = sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up').row
        sheetObj.range(str(strrng)+":"+str(endrng)).api.Delete()
        time.sleep(2)
        sheetObj.range('A1').api.AutoFilter(Field=9)

        sheetObj.range('A1').api.AutoFilter(Field=10, Criteria1="0", Operator=7)
        strrng = sheetObj.range("B2", sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up')).api.SpecialCells(12).Cells(1,1).Row
        endrng = sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up').row
        sheetObj.range(str(strrng)+":"+str(endrng)).api.Delete()
        time.sleep(2)
        sheetObj.range('A1').api.AutoFilter(Field=10)

        sheetObj.range('A1').api.AutoFilter(Field=10, Criteria1="2", Operator=7)
        strrng = sheetObj.range("B2", sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up')).api.SpecialCells(12).Cells(1,1).Row
        endrng = sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up').row
        sheetObj.range(str(strrng)+":"+str(endrng)).api.Delete()
        time.sleep(4)
        sheetObj.range('A1').api.AutoFilter(Field=10)

        sheetObj.range('A1').api.AutoFilter(Field=10, Criteria1="6", Operator=7)
        strrng = sheetObj.range("B2", sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up')).api.SpecialCells(12).Cells(1,1).Row
        endrng = sheetObj.cells(sheetObj.range('B:B').rows.count, "B").end('up').row
        sheetObj.range(str(strrng)+":"+str(endrng)).api.Delete()
        time.sleep(3)
        sheetObj.range('A1').api.AutoFilter(Field=10)

    stdco = datetime.datetime.now()

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
            # print("Extracting sheet# : ",i)
            df1 = df[(df['NOTETYPE'] != 40) & (df['NOTECLASS'] != 2) & (df['NOTECLASS'] != 0) & (df['NOTECLASS'] != 6)]
            df = None
            # full_sheet.range('A1').api.CurrentRegion.Copy()
            current_position_data = int(temp_sht.range('A1').api.CurrentRegion.Address.split('$')[-1])
            while not current_position_data <= 983000:
                k += 1
                temp_sht = temp_wb.sheets[k]
                current_position_data = int(temp_sht.range('A1').api.CurrentRegion.Address.split('$')[-1])
            if current_position_data == 1:
                current_position_data -= 1
                temp_sht.range('A'+str(current_position_data+1)).options(index=False).value = df1
            else:
                temp_sht.range('A'+str(current_position_data+1)).options(index=False).value = df1
                temp_sht.range(str(current_position_data+1)+":"+str(current_position_data+1)).api.Delete()
                # deleteFilter(temp_sht)
            df = None
        full_inv.close()
        full_inv = None
        time.sleep(2)


    nameDate = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M")
    tempBook = xw.Book(cwd + "\\Input Files\\" + 'Template.xlsb')
    destinationPath = path1 + "Full INVSTAT123 Consoliated "+ nameDate +".xlsb"
    # destinationPath = r'C:\Users\vijaykumar.mane\Desktop\TPBReport\Full INVSTAT123 Consoliated 2017-02-17-19-36.xlsb'
    tempBook.save(destinationPath)
    tempBook.close()
    BooktoConsolidate_in_Wb_Obj  = xw.Book(destinationPath)

    cou = 1
    for i in listFiles:
        consolidate_excelbook(i, BooktoConsolidate_in_Wb_Obj)
        os.remove(i)
        # print(cou,"  Consolidation done for ", i)
        cou += 1

    etdco = datetime.datetime.now()
    ttstdco = (etdco - stdco).total_seconds()
    # print("*******Complete Consolidation Time: Seconds ",ttstdco)
    # counter = BooktoConsolidate_in_Wb_Obj.sheets.count

    BooktoConsolidate_in_Wb_Obj.save()
    print('Complete Consolidation Done')

    endcomplete1 = datetime.datetime.now()
    ttstdComplete = (endcomplete1 - stdcomplete1).total_seconds()
    msg = "Full Invstat Complete time taken: " + str(ttstdComplete)
    print(msg)
    print("Full Invstat Complete time taken: ", ttstdComplete)

if flag3 == 0:
    BooktoConsolidate_in_Wb_Obj = xw.Book(latest_invstat)

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
# # Test
# BooktoConsolidate_in_Wb_Obj = xw.Book(r'C:\Users\vijaykumar.mane\Desktop\TPBReport\Full INVSTAT123 Consoliated 2017-02-17-19-36.xlsb')


bookingArray_full_inv = create_sheets_array(BooktoConsolidate_in_Wb_Obj, 'E')

BooktoConsolidate_in_Wb_Obj.close()
# Copy booking with hawb
bkhawb_Age = file_age_in_seconds(bookingWithHawb_path)
print("Age of Booking with Hawb in Minute:", bkhawb_Age//60)

shutil.copyfile(bookingWithHawb_path, cwd+"//Input Files//tempBookBKH.xlsb")
bookingWithHawb_path = cwd+"//Input Files//tempBookBKH.xlsb"

bkh = xw.Book(bookingWithHawb_path)
bookingArray_bkh = create_sheets_array(bkh, 'A')
bookingArray_hbl = create_sheets_array(bkh, 'D')
table_Dict = vDict(bookingArray_bkh, bookingArray_hbl)#dict(zip(bookingArray_bkh, bookingArray_hbl))
hbl_table = []
bk_table = []
bkh.close()

for b in bookingArray_full_inv:
    if b in table_Dict:
        hbl_table.append(table_Dict[b])
        bk_table.append(b)
    else:
        pass

hbl_table1 = vDict(hbl_table, hbl_table)
b_table = vDict(hbl_table,bk_table)
hawb = shtr.range('V2:V'+r_lr).value
hawb_status = []
bk_array = []
for m in hawb:
    if (m in hbl_table1) and (not m == None) and (not m == ''):
        hawb_status.append(hbl_table1[m])
        bk_array.append(b_table[m])
    else:
        hawb_status.append('No')
        bk_array.append("NA")

shtr.range('BH2:BH'+r_lr).options(transpose=True).value = hawb_status
shtr.range('BI2:BI'+r_lr).options(transpose=True).value = bk_array

print("Hawb Updated")

time.sleep(2)
tpb_temp = xw.Book(cwd+'\\Input Files\\'+"TPB Report_Templatev1.xlsb")
tpb_sht = tpb_temp.sheets("TPB Report")
shtr.activate()
shtr.range('A1').api.CurrentRegion.Copy()
tpb_sht.activate()
tpb_sht.range('A2').api.Select()
tpb_sht.api.PasteSpecial(12)
time.sleep(5)
tpb_sht.range('2:2').api.Delete()
tpb_sht.range('A2').api.Select()
time.sleep(2)
tpb_sht = tpb_temp.sheets('Summary')
tpb_sht.activate()
tpb_temp.api.RefreshAll()
time.sleep(2)
wbr.close()
tpb_temp.save(cwd+"\\TPB Archive\\"+'TPB Report '+ nameDate +".xlsb")
input_wb.close()
if flag3 == 1:
    con.close()

count = 1 # Number of transactions count
initialaht = 2400 # Time in seconds for completetion or per transaction
process = "CEVA"
strFileFullName = "TPB Reporting Macro" # Macro  name
modulename = "IMS-TPB Report Automation" # Department-Project_name
insert_timestamp(st, initialaht, process, strFileFullName, modulename, count)

print("TPB Report Created successfully")
time.sleep(20)
