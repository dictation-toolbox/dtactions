"""actions from application win32pad

see http://www.gena01.com/win32pad/, the only select and say notepad like window with line numbers
"""
import unimacro.messagefunctions
import as
import mf
from .actionbases import MessageActions
import pprint

class Win32padActions(MessageActions):
    def __init__(self, progInfo):
        MessageActions.__init__(self, progInfo)
        
    def getInnerHandle(self, handle):
        findValues = mf.findControls(handle, wantedClass="RichEdit20A")
        if findValues:
            return findValues[0]
        # else probably a child window
    
    #def metaaction_gotoline(self, n):
    #    """go to a specified line directly, no calling of dialog
    #    """
        
    
if __name__ == '__main__':
    # search all win32pad instances and dump the controls:
    tw = mf.findTopWindows(wantedText="win32pad")
    for t in tw:
        print 'top: %s'% t
        pprint.pprint( mf.dumpWindow(t) )
        
        