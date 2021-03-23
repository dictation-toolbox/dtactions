# actions classes for specific programs, base classes first:
import unimacro.messagefunctions as mess
import unimacro.natlinkutilsqh as natqh as natqh

class AllActions:
    def __init__(self, progInfo):
        self.reset(progInfo=progInfo)

    # init same as reset:        
    def reset(self, progInfo=None):
        if progInfo is None:
            progInfo = natqh.getProgInfo()
        self.prog, self.topTitle, self.topchild, self.className, self.topHandle = progInfo
        self.progInfo = None
        
    def update(self, newProgInfo=None):
        if newProgInfo is None:
            newProgInfo = natqh.getProgInfo()
        if newProgInfo == self.progInfo:
            return
        # print('allactions: new prog info, overload for your specific program: %s'% self.prog)
        self.progInfo = newProgInfo
        self.prog, self.topTitle, self.topchild, self.className, self.topHandle = newProgInfo
        
    def getCurrentLineNumber(self):
        pass
    
# base actions for applications that are connected through Windows Messages Functions
class MessageActions(AllActions):
    def __init__(self, progInfo):
        AllActions.__init__(self, progInfo=progInfo)
        self.classname = None   #  Scintilla can be handled quite generically with setting this variable
        self.prevTabName = None
        self.update(progInfo)

    def update(self, progInfo):
        """if prog info changes, then probably only title changes so notice and do nothing
        """
        if progInfo == self.progInfo:
            return 1 # OK
        self.ctrl = self.handle = self.getInnerHandle(self.topHandle)
        if not self.handle:
            if progInfo and progInfo[2] == 'top':
                print('no handle found for (top) edit control for program: %s'% self.prog)
            return
        self.progInfo = progInfo
        print('updated program info: %s, edit control handle: %s'% (repr(self.progInfo), self.handle)) 
        return self.handle  # None if no valid handle        

    def getCurrentLineNumber(self, handle=None):
        handle = handle or self.handle
        if not handle: return
        linenumber = mess.getLineNumber(handle, classname=self.classname)  # charpos = -1
        return linenumber

    def getNumberOfLines(self, handle=None):
        nl = mess.getNumberOfLines(handle, classname=self.classname)  
        return nl
    
    def isVisible(self, handle=None):
        handle = handle or self.handle
        if not handle: return
        return mess.isVisible(handle, classname=self.classname)
    
    def getEditText(self, handle=None):
        handle = handle or self.handle
        if not handle: return
        tList = mess.getEditText(self.ctrl)
        return ''.join(tList)
    
    getWindowText = getEditText
        
    def replaceSelection(self, output, handle=None):
        """overwrite selection with output
        """
        handle = handle or self.handle
        if not handle: return
        mess.replaceEditText(self.ctrl, output)

    def clearBoth(self, handle=None):
        """clear the target window, and subsequently the dictobj
        """
        handle = handle or self.handle
        if not handle: return
        mess.setEditText(handle, "")
        self.updateState()
        
    def getSelection(self, handle=None):
        """get the selection of the edit control
        
        return a 2 tuple
        """
        handle = handle or self.handle
        if not handle: return
        return mess.getSelection(handle)

    def setSelection(self, start, end, handle=None):
        """change the selection of the edit control
        """
        handle = handle or self.handle
        if not handle: return
        mess.setSelection(handle, start, end)
    
    def getVisibleRegion(self, handle=None):
        """Utility subroutine which calculates the visible region of the edit

         control and returns the start and end of the current visible region.
         Never got this working
        """
        handle = handle or self.handle
        if not handle: return
        return None, None
    
        ## to be investigated
        buf = self.getWindowText()
        return 0, len(buf)
        
        top,bottom,left,right = handle.GetClientRect()
        firstLine = handle.GetFirstVisibleLine()
        visStart = handle.LineIndex(firstLine)

        lineCount = handle.GetLineCount()
        lastLine = lineCount
        for line in range(firstLine+1,lineCount):
            charInLine = handle.LineIndex(line)
            left,top = handle.GetCharPos(charInLine)
            if top >= bottom:
                break
            lastLine = line

        visEnd = handle.LineIndex(lastLine+1)
        if visEnd == -1:
            visEnd = len(handle.GetWindowText())
        return visStart,visEnd
    
    
        