"""actions from application uedit64, UltraEdit

not fully tested, all this, Quintijn Hoogenboom, 2021

(from IDM Computer Solutions, www.ultraedit.com)
"""
import win32gui
import ctypes
import win32api
from dtactions import messagefunctions as mf
from actionbases import MessageActions
import pprint

class Uedit32Actions(MessageActions):
    def __init__(self, progInfo):
        MessageActions.__init__(self, progInfo)
        
    def getInnerHandle(self, handle):
        """procedure to get inner handle for EditControl,
        
        via a control that holds the filename, which is extracted from the window title
        """
        title = win32gui.GetWindowText(handle)
        #print 'title: %s'% title
        tabName = title.split(']')[0]
        tabName = tabName.split('[')[-1]
        self.prevTabName = tabName
        #print 'tabName: %s'% tabName
        controls = mf.findControls(handle, wantedText = tabName)
        if not controls:
            #print 'uedit32Actions, no control with title: "%s" found'% tabName
            return
        if len(controls) > 1:
            print('uedit64Actions, strange, found more controls, matching the title: "%s", take first'% tabName)
        c = controls[0]
        innerControl = mf.findControls(c, wantedClass="EditControl")[0]
        return innerControl
    
    
    
    
if __name__ == '__main__':
    import sys
    sys.path.append(r'C:\DT\Unimacro\src\unimacro')   # testing...

    # test, fill in handle of ultraedit (uedit32) (with "give window info" from grammar _general)
    focus =199852
    print('focus; %s'% focus)
    progInfo = ('uedit32', 'babbababa', 'top', focus)
    ue = Uedit32Actions( progInfo )
