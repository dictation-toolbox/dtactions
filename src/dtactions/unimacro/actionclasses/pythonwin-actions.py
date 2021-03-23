"""actions from application pythonwin
"""
import win32gui
import win32con
import unimacro.messagefunctions
import as
import mf
from .actionbases import MessageActions
import pprint
class PythonwinActions(MessageActions):
    handlesDict = {}
    def __init__(self, progInfo):
        MessageActions.__init__(self, progInfo)
        
    def getInnerHandle(self, handle):
        
        controls = mf.findControls(handle, wantedClass="Scintilla")
        print 'Scintilla controls: %s'% controls
        for c in controls:
            ln = self.getCurrentLineNumber(c)
            numberLines = self.getNumberOfLines(c)
            visible1 = mf.isVisible(c)
            print 'c: %s, linenumber: %s, nl: %s, visible: %s'% (c, ln, numberLines, visible1)
                
if __name__ == '__main__':

    pass