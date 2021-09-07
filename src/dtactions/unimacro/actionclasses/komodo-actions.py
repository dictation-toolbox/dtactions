"""actions from application komodo

see www.activestate.com. Does not work yet for s&s or line numbers modulo 100
now special (QH) metaactions for very special tasks on selections of a file only
"""
#pylint:disable=C0115, C0116. R0201
import time
# from dtactions import messagefunctions as mf
from dtactions.unimacro.actionclasses.actionbases import AllActions
# from dtactions.unimacro.unimacroactions import doAction as action
from dtactions.unimacro.unimacroactions import doKeystroke as keystroke
from dtactions.sendkeys import sendkeys
import natlinkclipboard


# class KomodoActions(MessageActions):
class KomodoActions(AllActions):
    def __init__(self, progInfo):
        AllActions.__init__(self, progInfo)
        self.classname = "Komodo"
        
    def getInnerHandle(self, handle):
        # cannot get the correct "inner handle"
        handle = None
        return handle
        # nextHndle = handle 
        # user32 = ctypes.windll.user32
        # #
        # controls = mf.findControls(handle, wantedClass="Scintilla")
        # print('Scintilla controls: %s'% controls)
        # for c in controls:
        #     ln = self.getCurrentLineNumber(c)
        #     numberLines = self.getNumberOfLines(c)
        #     visible1 = self.isVisible(c)
        #     info = win32gui.GetWindowPlacement(c) 
        #     print('c: %s, linenumber: %s, nl: %s, info: %s'% (c, ln, numberLines, repr(info)))
        #     parent = c
        #     while 1:
        #         parent = win32gui.GetParent(parent)
        #         clName = win32gui.GetClassName(parent)
        #         visible = self.isVisible(parent)
        #         info = win32gui.GetWindowPlacement(parent) 
        #         print('parent: %s, class: %s, visible: %s, info: %s'% (parent, clName, visible, repr(info)))
        #         if parent == handle:
        #             print('at top')
        #             break
    
    def getCurrentLineNumber(self):
        debug = 0
        t1 = time.time()
        if self.toporchild == "child":
            return 0
        cb = natlinkclipboard.Clipboard(save_clear=True, debug=debug)  # clear "debug" to get rid of timing line
        shortcutkey = "{alt+shift+c}"
        keystroke(shortcutkey)
        
        # # now collect the clipboard, at most waiting 10 intervals of 0.1 second.
        result = cb.get_text(waiting_interval=0.01, waiting_iterations=10)    # should be the current line number
        # print(f'result from clipboard: {result}')
        t2 = time.time()
        lapse = t2 - t1
        if debug:
            print(f'time in getCurrentLineNumber: {lapse:.3f}')
        try:
            return int(result)
        except (ValueError, TypeError):
            return 0
        
    def metaaction_makeunicodestrings(self, dummy=None):
        """for doctest testing, put u in front of every string
        """
        print('metaaction_makeunicodestrings, for Komodo')
        sendkeys("{ctrl+c}")
        t = unimacroutils.getClipboard()
        print('in: %s'% t)
        t = replaceStringToUnicode(t)
        unimacroutils.setClipboard(t)
        natlinkutils.playString("{ctrl+v}{down}")

def _test():
    """do doctests, for changing function
    """
    import doctest
    doctest.testmod()
    
