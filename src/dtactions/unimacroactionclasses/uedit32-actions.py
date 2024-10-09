"""actions from application uedit32, UltraEdit

(from IDM Computer Solutions, www.ultraedit.com)
"""
#pylint:disable=R0903
import win32gui
import dtactions.messagefunctions as mf
from .actionbases import MessageActions

class Uedit32Actions(MessageActions):
    def __init__(self, progInfo):
        MessageActions.__init__(self, progInfo)
        
    def getInnerHandle(self, topHndle):
        """procedure to get inner handle for EditControl,
        
        via a control that holds the filename, which is extracted from the window title
        """
        title = win32gui.GetWindowText(topHndle)
        #print 'title: %s'% title
        tabName = title.split(']')[0]
        tabName = tabName.split('[')[-1]
        self.prevTabName = tabName
        #print 'tabName: %s'% tabName
        controls = mf.findControls(topHndle, wantedText = tabName)
        if not controls:
            #print 'uedit32Actions, no control with title: "%s" found'% tabName
            return None
        if len(controls) > 1:
            print(f'uedit32Actions, strange, found more controls, matching the title: "{tabName}", take first')
        c = controls[0]
        innerControl = mf.findControls(c, wantedClass="EditControl")[0]
        return innerControl
    
if __name__ == '__main__':

    # test, fill in handle of ultraedit (uedit32) (with "give window info" from grammar _general)
    focus =199852
    print(f'focus; {focus}')
    ProgInfo = ('uedit32', 'babbababa', 'top', focus)
    ue = Uedit32Actions( ProgInfo )
