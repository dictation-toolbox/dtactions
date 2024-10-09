"""actions from application win32pad

see http://www.gena01.com/win32pad/, the only select and say notepad like window with line numbers
"""
#pylint:disable=R0903
import pprint
import dtactions.messagefunctions as mf
from .actionbases import MessageActions


class Win32padActions(MessageActions):
    def __init__(self, progInfo):
        MessageActions.__init__(self, progInfo)
        
    def getInnerHandle(self, topHndle):
        findValues = mf.findControls(topHndle, wantedClass="RichEdit20A")
        if findValues:
            print(f'win32pad, topHndle: {topHndle}, innerHandle (take first: {findValues}')
            return findValues[0]
        # else probably a child window
        return None
    #def metaaction_gotoline(self, n):
    #    """go to a specified line directly, no calling of dialog
    #    """
        
    
if __name__ == '__main__':
    # search all win32pad instances and dump the controls:
    tw = mf.findTopWindows(wantedText="win32pad")
    for th in tw:
        print(f'top: {th}')
        pprint.pprint( mf.dumpWindow(th) )
        
        