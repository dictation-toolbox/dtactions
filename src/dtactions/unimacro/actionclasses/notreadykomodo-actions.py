"""actions from application komodo

see www.activestate.com. Does not work (yet)
"""
import win32gui
import ctypes
import win32api
from dtactions import messagefunctions as mf
from .actionbases import MessageActions
import pprint

class KomodoActions(MessageActions):
    def __init__(self, progInfo):
        MessageActions.__init__(self, progInfo)
        self.classname = "Scintilla"
        
    def getInnerHandle(self, handle):
        # cannot get the correct "inner handle"
        return
        nextHndle = handle 
        user32 = ctypes.windll.user32
        #
        controls = mf.findControls(handle, wantedClass="Scintilla")
        print('Scintilla controls: %s'% controls)
        for c in controls:
            ln = self.getCurrentLineNumber(c)
            numberLines = self.getNumberOfLines(c)
            visible1 = self.isVisible(c)
            info = win32gui.GetWindowPlacement(c) 
            print('c: %s, linenumber: %s, nl: %s, info: %s'% (c, ln, numberLines, repr(info)))
            parent = c
            while 1:
                parent = win32gui.GetParent(parent)
                clName = win32gui.GetClassName(parent)
                visible = self.isVisible(parent)
                info = win32gui.GetWindowPlacement(parent) 
                print('parent: %s, class: %s, visible: %s, info: %s'% (parent, clName, visible, repr(info)))
                if parent == handle:
                    print('at top')
                    break
        

# several tries for getting introspection of the Komodo controls failed sofar (QH, October 2013) 
import ctypes
import win32con
def get_windows(startWith=None):
    '''Returns windows in z-order (top first)'''
    user32 = ctypes.windll.user32
    lst = []
    top = user32.GetTopWindow(startWith)
    if not top:
        return lst
    lst.append(top)
    next = top
    print('top: %s'% next)
    while True:
        
        next = user32.GetTopWindow(next)
        print('next: %s'% next)
        if not next:
            break
    #    lst.append(next)
    #return lst        

if __name__ == '__main__':

    focus =66284
    
    print('focus; %s'% focus)
    progInfo = ('komodo', 'babbababa', 'top', focus)
    ka = KomodoActions( progInfo )
    get_windows()
    #tw = mf.findTopWindows(wantedText="komodo")
    #for t in tw:
    #    print 't:%s'% t
    #    lst = get_windows(t)
    #    for l in lst:
    #        visible = win32gui.SendMessage(l, win32con.WS_VISIBLE)
    #        print 'l: %s, visible: %s'% (l, visible)
    ##    #pr
    ##    topchild = user32.GetTopWindow(t)
    ##    nextchild = user32.GetTopWindow(topchild)
    ##    print 'tophandle: %s, topchild: %s, nextchild: %s'% (t, topchild, nextchild)
    ##    
    #    pprint.pprint( mf.dumpWindow(t) )
    #    print '----------------------------------------------------------'        
    #    