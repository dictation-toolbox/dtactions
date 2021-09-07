import win32com.client
import time
import os
import os.path
from pythoncom import com_error
from dtactions.unimacro.actionclasses.actionbases import AllActions
from dtactions.unimacro import unimacroutils as natqh


class ExcelActions(AllActions):
    """attach to excel and perform necessary actions
    """
    appList = [] # here goes, for all instances, the excel application object
    positions = {} # dict with book/sheet/(col, row) (tuple)
    rows = {} # dict with book/sheet/row  (str)
    columns = {} # dict with book/sheet/col (str)
    def __init__(self, progInfo):
        AllActions.__init__(self, progInfo)
        self.connect() # sets self.app
        if not self.app:
            return
        self.update(progInfo)

    def reset(self, progInfo=None):
        """reset progInfo and prev variables. Leave self.app
        """
        AllActions.reset(self, progInfo)
        self.prevBook = self.prevSheet = self.prevPosition = None

    def update(self, progInfo):
        print('update of ExcelActions, %s'% self.topHandle)
        if self.topHandle != progInfo[3]:
            print('reset, topHandle: %s, progInfo: %s'% (self.topHandle, repr(progInfo)))
            self.reset(progInfo)
            if self.app:
                print('wrong app for excel, window handle invalid, app is connected to %s, forground window has %s\nPlease close conflicting Excel instance!'% (self.topHandle, progInfo[3]))
                self.disconnect() # sets self.app to None
            else:
                print('try to reconnect to %s'% self.topHandle)
                self.connect()
        if self.app:
            self.checkForChanges(progInfo)
        else:
            print('no valid excel instance available for hndle: %s'% self.topHandle)
        
    def isInForeground(self, app=None, title=None, progInfo=None):
        """return True if app is in foreground
        """
        thisApp = app or self.app
        title = title or self.topTitle
        title = title.lower()
        
        if not thisApp:
            return False
        if thisApp.Workbooks.Count == 0:
            print('Got excel app with no workbooks')
            return False
        try:
            name = thisApp.ActiveWorkbook.Name
        except AttributeError:
            print('workbook is not saved yet.')
            return True
        #print 'activeworkbook: %s'% name
        if title.find(name.lower()) >= 0:
            return True
        print('============\nexcel, isInForeground: cannot find name \n"%s" \nin title: \n"%s"\nProbably you have more excel instances open. Unimacro only can connect to this one instance.\n'% (name, title))
        return False        
        
    def checkForChanges(self, progInfo=None):
        """return 1 if book, sheet or position has changed since previous call
        """
        progInfo = progInfo or unimacroutils.getProgInfo()
        changed = 0
        if not self.app:
            self.progInfo = progInfo
            self.connect()
            if not self.app:
                return
        
        try:
            visible = self.app.Visible
        except:
            self.app = None
            self.connect()
            if not self.app:
                return
        
        title = self.topTitle = progInfo[1]

        if not self.isInForeground(app=self.app, title=title, progInfo=progInfo):
            print('not in foreground, excel')
            if self.prevBook:
                changed = 8 # from foreground into background
                self.prevBook = self.prevSheet = self.prevPosition = None
            self.book = self.sheet = self.Position = None
            self.currentRow = self.currentLine = None
            self.currentColumn = None
            self.progInfo = progInfo
            return


        self.book = self.app.ActiveWorkbook
        
        if self.book is None:
            print('ExcelAction, no active workbook')
            self.bookName = ''
            self.sheetName = ''
            self.sheet = None
        else:
            self.bookName = self.book.Name
            self.sheet = self.app.ActiveSheet
            self.sheetName = self.sheet.Name
        if self.prevBook != self.bookName:
            self.prevBook = self.bookName
            changed += 4
        if self.prevSheet != self.sheetName:
            self.prevSheet = self.sheetName
            changed += 2
        if changed:
            # print 'excel-actions: update book and/or sheet variables'
            self.positions.setdefault(self.bookName,{})
            self.positions[self.bookName].setdefault(self.sheetName, [])
            self.columns.setdefault(self.bookName,{})
            self.columns[self.bookName].setdefault(self.sheetName, [])
            self.rows.setdefault(self.bookName,{})
            self.rows[self.bookName].setdefault(self.sheetName, [])
            self.currentPositions = self.positions[self.bookName][self.sheetName]
            self.currentColumns = self.columns[self.bookName][self.sheetName]
            self.currentRows = self.rows[self.bookName][self.sheetName]
        if self.sheet:
            #print 'excel-actions: %s, update current position'% self.sheet

            cr = self.savePosition()
            self.currentPosition = cr
            self.currentRow = self.currentLine = cr[1]
            self.currentColumn = cr[0]
            #print 'currentColumn: %s, currentRow: %s, currentLine: %s'% (self.currentColumn, self.currentRow, self.currentLine)
            if cr != self.prevPosition:
                self.prevPosition = cr
                changed += 1
            #else:
            #    print 'same position: %s'% repr(cr)
        else:
            print('excel-actions, no self.sheet.')
        #print 'return code checkForChanges: %s'% changed
        return changed
    
    #connect to programs:
    def connect(self):
        """connect to excel set self.app
        
        set self.app to None if no excel instance available
        also reset the relevant variables...
        return value irrelevant
        """
        print('Connecting to to excel...')
        self.prevBook = self.prevSheet = self.prevPosition = None
        if self.prog != 'excel':
            print("excel-actions, should only be called when excel is the foreground window, not: %s"% self.prog)
            self.app = None
            return
        title = self.topTitle
        try:
            self.app = win32com.client.GetActiveObject(Class="Excel.Application")
        except com_error:
            print('Cannot attach to Excel.')
            self.app = None
            return
        if self.isInForeground():
            return  # OK
        else:
            print('ExcelActions: cannot attach to Excel, because instance is not in the foreground.')
            print('Probably there are more excel instances active.')
            print('Please close excel instances and try again')
            self.disconnect()
            return


    def recentMatchesTitle(self, recentFile, windowTitle):
        """check the recent file with the current window title
        """
        p, name = os.path.split(recentFile)
        if windowTitle.find(name) >= 0:
            return True
        

    def disconnect(self):
        """disconnect from excel and set self.app to None
        
        reset the relevant variables
        by removing the connection (not powerful enough)
        """
        # if self.app:
            # try:
            #     self.app.Quit()
            # except com_error:
            #     pass
        self.reset()
        self.app = None

    def savePosition(self):
        """save current position in positions dict
        
        returns the current (col, row) tuple
        """
        cr = self.getCurrentPosition()
        c, r = cr
        self.pushToListIfDifferent(self.currentPositions, cr)
        self.pushToListIfDifferent(self.currentRows, r)
        self.pushToListIfDifferent(self.currentColumns, c)
        
        return cr
    
    def getBookFromTitle(self, title, app=None):
        """return the book name corresponding with title (being the window title)
        """
        bookList = self.getBooksList(app=app)
        if bookList:
            for b in bookList:
                if title.find(b) >= 0:
                    print('found book in instance: %s'% b)
                    return b
        print('not found title in instance, bookList: %s'% bookList)

    def getBookNameFromTitle(self, title):
        """return the book name corresponding with title (being the window title)
        """
        excelTexts = ['Microsoft Excel -', '- Microsoft Excel']
        print('getBookNameFromTitle: %s'% title)
        if title.find('['):
            title = title.split('[')[0].strip()
        for e in excelTexts:
            if title.find(e) >= 0:
                title = title.replace(e, '').strip()
        print('result: %s'% title)
        return title
            
        
    
    def getBooksList(self,app=None):
        """get list of strings, the names of the open workbooks
        """
        app = app or self.app
        if app:
            books = []
            for i in range(app.Workbooks.Count):
                b = app.Workbooks(i+1)
                books.append(b.Name)
            return books
        return []

    def getCurrentLineNumber(self, handle=None):
        if not self.app: return
        lineStr = self.getCurrentPosition()[1]
        return int(lineStr)
    
    
    def getSheetsList(self, book=None):
        """get list of strings, the names of the open workbooks
        """
        if book is None:
            book = self.book
        elif isinstance(book, str):
            book = self.app.Workbooks[book]
        
        if book:
            sheets = []
            for i in range(book.Sheets.Count):
                s = self.app.Worksheets(i+1)
                sheets.append(s.Name) # assume this is Unicode...
            return sheets

    
    def selectSheet(self, sheet):
        """select the sheet by name of number
        """
        self.app.Sheets(sheet).Activate()
    
    def getCurrentPosition(self):
        """return row and col of activecell
        
        as a side effect remember the (changed position)
        """
        if not (self.app and self.book and self.sheet):
            return None, None
        ac = self.app.ActiveCell
        comingFrom = ac.Address
        #print 'activecell: %s'% comingFrom
        cr = [value.lower() for value in comingFrom.split("$") if value] # assume value unicode
        if len(cr) == 2:
            #print 'currentposition, lencr: %s, cr: %s'% (len(cr), cr)
            return tuple(cr)
        else:
            print('excel-actions, no currentposition, lencr: %s, cr: %s'% (len(cr), cr))
            return None, None

    def pushToListIfDifferent(self, List, value):
        """add to list (in place) of value differs from last value in list
        
        for positions, value is (c,r) tuple
        """
        if not value:
            return
        if List and List[-1] == value:
            return
        List.append(value)
    
    def getPreviousRow(self):
        """return the previous row number
        """
        cr = self.getCurrentPosition()
        c, r = cr
        while 1:
            newr = self.popFromList(self.currentRows)
            if newr is None:
                return
            if r != newr:
                return newr

    def getPreviousColumn(self):
        """return the previous col letter
        """
        cr = self.getCurrentPosition()
        c, r = cr
        while 1:
            newc = self.popFromList(self.currentColumns)
            if newc is None:
                return
            if c != newc:
                return newc
    
    def popFromList(self, List):
        """pop from list a different value than currentValue
        
        and return None if List is exhausted
        """
        if List:
            value = List.pop()
            return value


    # functions that do an action from the action.py module, in case of excel:
    # one parameter should be given    
    def metaaction_gotoline(self, rowNum):
        """overrule for gotoline meta-action
        
        goto line in the current column
        """
        rowNum = str(rowNum)
        if not self.app:  return
        cPrev, rPrev = self.getCurrentPosition()
        if rowNum == rPrev:
            print('row already selected')
            return 1
        try:
            range = cPrev + rowNum
            #print 'current range: %s, %s'% (rPrev, cPrev)
            sheet = self.app.ActiveSheet
            #print 'app: %s, sheet: %s (%s), range: %s'% (app, sheet, sheet.Name, range)
            sheet.Range(range).Select()
            return 1
        except:
            print('something wrong in excel-actions, metaaction_gotoline.')
            return
            
    def metaaction_selectline(self, dummy=None):
        """select the current line
        """
        if not self.app: return
        self.app.ActiveCell.EntireRow.Select()
        return 1

    def metaaction_remove(self, dummy=None):
        """remove the selection, assume rows or columns are selecte
        """
        if not self.app: return
        self.app.Selection.Delete()

        return 1
        
    def metaaction_insert(self, dummy=None):
        """insert the number of lines that are selected
        """
        if not self.app: return
        self.app.Selection.Insert()

        return 1

    metaaction_lineinsert = metaaction_insert

    def metaaction_pasteinsert(self, dummy=None):
        """insert the number of lines that are selected
        """
        if not self.app: return
        self.app.Selection.Insert()
        self.app.CutCopyMode = False

        return 1

    def metaaction_selectpreviousline(self, dummy=None):
        """select the previous line with respect to the activecell
        """
        if not self.app: return
        wantedLine = int(self.currentRow) - 1
        self.metaaction_gotoline(wantedLine)
        self.app.ActiveCell.EntireRow.Select()
        return 1

    def metaaction_movetotopofselection(self, dummy=None):
        """select first line of a selected range
        """
        if not self.app: return
        self.app.Selection.Rows(1).Select()
        return 1

    def metaaction_movetobottomofselection(self, dummy=None):
        """select first line of a selected range
        """
        if not self.app: return
        nRows = self.app.Selection.Rows.Count
        self.app.Selection.Rows(nRows).Select()
        return 1
    def metaaction_movecopypaste(self, dummy=None):
        """insert the clipboard after a movecopy action.
        """
        if not self.app: return
        self.app.ActiveCell.EntireRow.Insert()
        self.app.CutCopyMode = False
        return 1
        
    def metaaction_lineback(self, dummy=None):
        """goes back to the previous row
        """
        if not self.app: return
        #self.app.ActiveCell.EntireRow.Select()
        prevRow = self.getPreviousRow()
        print('prevRow: %s'% prevRow)
        if prevRow:                
            self.gotoRow(prevRow)
        return 1
        
if __name__ == '__main__':
    progInfo = ('excel', 'Microsoft Excel - Map1', 'top', 921280)
    excel = ExcelActions(progInfo)
    if excel.app:
        #excel.app.Visible = True
        print('activeCell: %s'% excel.app.ActiveCell)
        print('books: %s'% excel.app.Workbooks.Count)
        print('foreground: %s'%excel.isInForeground())
        print('now click on excel please')
        print(excel.app.hndle)
        excel.metaaction_gotoline(345)
        time.sleep(5)
    else:
        print('no excel.app: %s'% excel.app)
            
