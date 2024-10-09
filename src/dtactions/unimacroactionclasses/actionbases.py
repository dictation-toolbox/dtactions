"""actions classes for specific programs, base classes first

This module provides the different classes to base a program specific actionsclass on

"""
#pylint:disable=C0116, W0201
from dtactions import messagefunctions as mess
from dtactions.unimacro import unimacroutils

class AllActions:
    """base class for all actions that run in actionclasses
    """
    def __init__(self, progInfo):
        self.reset(progInfo=progInfo)
        
    # init same as reset:        
    def reset(self, progInfo=None):
        if progInfo is None:
            progInfo = unimacroutils.getProgInfo()
        self.progInfo = progInfo
        
    def update(self, newProgInfo=None):
        if newProgInfo is None:
            newProgInfo = unimacroutils.getProgInfo()
        if newProgInfo == self.progInfo:
            return
        # print('allactions: new prog info, overload for your specific program: %s'% self.prog)
        self.progInfo = newProgInfo
        
    def getCurrentLineNumber(self):
        pass
    
# base actions for applications that are connected through Windows Messages Functions
class MessageActions(AllActions):
    """class based upon windows messages, to get information about its status
    """
    def __init__(self, progInfo):
        AllActions.__init__(self, progInfo=progInfo)
        self.classname = None   #  Scintilla can be handled quite generically with setting this variable
        self.prevTabName = None
        self.progInfo = None
        self.update(progInfo)

    def getInnerHandle(self, topHndle):
        """should be overridden by a specific Messages window class
        """
        raise NotImplementedError

    def update(self, newProgInfo=None):
        """if prog info changes, then probably only title changes so notice and do nothing
        """
        if newProgInfo == self.progInfo:
            return 1 # OK
        if not newProgInfo:
            print('actionbases, update, no new ProgInfo')
            return 0
        self.progInfo = newProgInfo
        self.innerHndle = self.getInnerHandle(self.progInfo.hndle)
        if not self.innerHndle:
            if newProgInfo and newProgInfo.toporchild == 'top':
                print(f'no handle found for (top) edit control for program: {self.progInfo.prog}')
            return 0
        print(f'updated program info: {self.progInfo}, edit control (inner) handle: {self.innerHndle}')
        return self.innerHndle # None if no valid handle        

    def getCurrentLineNumber(self, handle=None):
        self.update()
        innerHndle = handle or self.innerHndle
        if not innerHndle:
            return None
        # progpath, prog, title, topchild, classname, hndle = self.progInfo
        linenumber = mess.getLineNumber(innerHndle, classname=self.progInfo.classname)  # charpos = -1
        # print(f'linenumber from MessageActions: {linenumber}')
        return linenumber

    def getNumberOfLines(self, handle=None):
        nl = mess.getNumberOfLines(handle, classname=self.classname)  
        return nl
    
    def isVisible(self, handle=None):
        handle = handle or self.innerHndle
        if not handle:
            return None
        return mess.isVisible(handle, classname=self.classname)
    
    def getEditText(self, handle=None):
        handle = handle or self.innerHndle
        if not handle:
            return None
        tList = mess.getEditText(handle)
        return ''.join(tList)
    
    getWindowText = getEditText
        
    def replaceSelection(self, output, handle=None):
        """overwrite selection with output
        """
        handle = handle or self.innerHndle
        if not handle:
            return
        mess.replaceEditText(handle, output)

    def clearBoth(self, handle=None):
        """clear the target window, and subsequently the dictobj
        """
        handle = handle or self.innerHndle
        if not handle:
            return
        mess.setEditText(handle, "")
        newProgInfo = unimacroutils.getProgInfo()
        self.update(newProgInfo)
        
    def getSelection(self, handle=None):
        """get the selection of the edit control
        
        return a 2 tuple
        """
        handle = handle or self.innerHndle
        if not handle:
            return None
        return mess.getSelection(handle)

    def setSelection(self, start, end, handle=None):
        """change the selection of the edit control
        """
        handle = handle or self.innerHndle
        if not handle:
            return
        mess.setSelection(handle, start, end)
    
    def getVisibleRegion(self, handle=None):
        """Utility subroutine which calculates the visible region of the edit

         control and returns the start and end of the current visible region.
         Never got this working
        """
        handle = handle or self.innerHndle
        if not handle:
            return None
        return None, None
        
        # TODOQH    
        # ## to be investigated
        # buf = self.getWindowText()
        # return 0, len(buf)
        
        # top,bottom,left,right = handle.GetClientRect()
        # firstLine = handle.GetFirstVisibleLine()
        # visStart = handle.LineIndex(firstLine)
        # 
        # lineCount = handle.GetLineCount()
        # lastLine = lineCount
        # for line in range(firstLine+1,lineCount):
        #     charInLine = handle.LineIndex(line)
        #     left,top = handle.GetCharPos(charInLine)
        #     if top >= bottom:
        #         break
        #     lastLine = line
        # 
        # visEnd = handle.LineIndex(lastLine+1)
        # if visEnd == -1:
        #     visEnd = len(handle.GetWindowText())
        # return visStart,visEnd
    
    
        