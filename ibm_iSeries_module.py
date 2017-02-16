from comtypes.client import CreateObject

class iseries(object):

    def __init__(self, uid, pwd, session, profile):
        self.uid = uid
        self.pwd = pwd
        self.session = session
        self.profile = profile

    def obj_cnmgr(self):
        conMgr = CreateObject("PCOMM.autECLConnMgr")
        return (conMgr)

    def obj_cnses(self):
        SessObj = CreateObject("PCOMM.autECLSession")
        return (SessObj)

    def obj_cnlist(self):
        conList = CreateObject("PCOMM.autECLConnList")
        return (conList)

    def obj_cntest(self):
        conTest = CreateObject("PCOMM.autECLOIA")
        return (conTest)

    def obj_cnobj(self):
        conObj = CreateObject("PCOMM.autECLPS")
        return (conObj)

    def connect(self):
        conMgr = CreateObject("PCOMM.autECLConnMgr")	
        # start a new connection
        conMgr.StartConnection ("profile=" + self.profile +" connname=" + self.session)

    def test_window(self):
        obj = CreateObject("PCOMM.autECLOIA")
        obj.SetConnectionByName (self.session)
        while 1:
            obj.WaitForInputReady(2500)
            if not (obj.Started == 'False'):
                break
        test = obj.Started
        if test == True:
            print('Emulator window Started')
        else:
            print('Emulator window Not - Started')

    def start_communication(self):
        PSObj = CreateObject("PCOMM.autECLPS")
        conList = CreateObject("PCOMM.autECLConnList")
        conList.Refresh
        PSObj.SetConnectionByHandle (conList.ConnInfo(1).Handle)
        PSObj.startcommunication

    def test_comm(self):
        obj = CreateObject("PCOMM.autECLOIA")
        obj.SetConnectionByName (self.session)
        test = obj.CommStarted
        if test == True:
            print('Communication Started')
        else:
            print('Communication Not - Started')

    def app_wait(self, timeout=''):
        self.timeout = timeout
        obj = CreateObject("PCOMM.autECLOIA")
        obj.SetConnectionByName (self.session)
        obj.WaitForAppAvailable(self.timeout)

    def inp_wait(self, timeout=''):
        self.timeout = timeout
        obj = CreateObject("PCOMM.autECLOIA")
        obj.SetConnectionByName (self.session)
        obj.WaitForInputReady(self.timeout)

    def system_check(self):
        # Object Creation
        conTest = CreateObject("PCOMM.autECLOIA")
        conList = CreateObject("PCOMM.autECLConnList")
        conTest.SetConnectionByName (self.session)
        conTest.WaitForInputReady (7000)
        conList.Refresh
        if conTest.InputInhibited == 0:
            print('Not Inhibited ')
        if conTest.InputInhibited == 1:
            print('System Wait ')
        if conTest.InputInhibited == 2:
            print('Communication Check ')
        if conTest.InputInhibited == 3:
            print('Program Check ')
        if conTest.InputInhibited == 4:
            print('Machine Check ')
        if conTest.InputInhibited == 5:
            print('Other Inhibit ')

    def set_text(self, value, row='', col=''):
        self.value = value
        self.row = row
        self.col = col
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        if self.row == '' or self.col == '':
            conObj.SetText (self.value)
        else:
            conObj.SetText (self.value, self.row, self.col)

    def send_keys(self, key):
        self.key = key
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        conObj.SendKeys ("[" + self.key + "]")

    def get_text(self, row='', col='', length=''):
        self.row = row
        self.col = col
        self.length = length
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        if self.row == '' and self.col == '':
            return (conObj.GetText())
        else:
            return (conObj.GetText (self.row, self.col, self.length))

    def get_rect_text(self, srow, scol, erow, ecol):
        self.srow = srow
        self.scol = scol
        self.erow = erow
        self.ecol = ecol
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        return (conObj.GetTextRect (self.srow, self.scol, self.erow, self.ecol))

    def set_cursor(self, row, col):
        self.row = row
        self.col = col
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        conObj.SetCursorPos(self.row, self.col)

    def search_text(self, text, direction, row, col):
        self.text = text
        self.direction = direction
        self.row = row
        self.col = col
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        return conObj.SearchText (self.text, self.direction, self.row, self.col)

    def wait_for_text(self, value, row='', col='', timeout=''):
        self.value = value
        self.row = row
        self.col = col
        self.timeout = timeout
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        conObj.WaitForString (self.value, self.row, self.col, self.timeout)

    def wait_for_rect_text(self, value, srow, scol, erow, ecol, timeout=''):
        self.srow = srow
        self.scol = scol
        self.erow = erow
        self.ecol = ecol
        self.timeout = timeout
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        conObj.WaitForStringInRect (self.srow, self.scol, self.erow, self.ecol, self.timeout)

    def wait_while_text(self, value, row='', col='', timeout=''):
        self.value = value
        self.row = row
        self.col = col
        self.timeout = timeout
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        conObj.WaitWhileString (self.value, self.row, self.col, self.timeout)

    def wait_while_rect_text(self, value, srow, scol, erow, ecol, timeout=''):
        self.srow = srow
        self.scol = scol
        self.erow = erow
        self.ecol = ecol
        self.timeout = timeout
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        conObj.WaitWhileStringInRect (self.srow, self.scol, self.erow, self.ecol, self.timeout)

    def wait_for_screen(self, row='', col='', timeout=''):
        self.row = row
        self.col = col
        self.timeout = timeout
        scr = CreateObject("PCOMM.autECLScreenDesc")
        conObj = CreateObject("PCOMM.autECLPS")
        conObj.SetConnectionByName (self.session)
        scr.AddCursorPos (self.row, self.col)
        if conObj.WaitForScreen (scr, self.timeout):
            print ("screen reached")
        else:
            print ("screen not reached")		

    def iSeries_download(ip, uid, pwd, fname, path):
        dlr = CreateObject('cwbx.DatabaseDownloadRequest')
        dlr.system = CreateObject('cwbx.AS400System')
        dlr.system.Define (ip)
        dlr.system.UserId = uid
        dlr.system.Password = pwd
        dlr.system.Signon()
        dlr.AS400File.Name = fname
        dlr.pcFile.FileType = 16
        dlr.pcFile.Name = path
        dlr.download()
