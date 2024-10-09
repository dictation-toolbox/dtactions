"""actions from application komodo

see www.activestate.com. Does not work yet for s&s or line numbers modulo 100
now special (QH) metaactions for very special tasks on selections of a file only
"""
#pylint:disable=C0115, C0116
import time

import natlink
# from dtactions import messagefunctions as mf
from dtactions.unimacro.actionclasses.actionbases import AllActions
# from dtactions.unimacro.unimacroactions import doAction as action
from dtactions.sendkeys import sendkeys as keystroke

# class KomodoActions(MessageActions):
class KomodoActions(AllActions):
    def __init__(self, progInfo):
        AllActions.__init__(self, progInfo)
        self.classname = "Komodo"
        
    def getInnerHandle(self, topHndle):
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
        if self.progInfo.toporchild == "child":
            return 0
        shortcutkey = "{alt+shift+c}"
        keystroke(shortcutkey)
        time.sleep(0.1)
        result = natlink.getClipboard()
        
        try:
            return int(result)
        except (ValueError, TypeError):
            print('getCurrentLineNumber does not work correct for Komodo:')
            print('Say "Edit lines" for editing the options of the grammar _lines,')
            print('and change option "line numbers modulo hundred" in section [general] to F')
            print(f'Komodo needs a macro to be implemented and bound to the shortcut keys {shortcutkey}.')
            print('For implementing this macro, see: https://qh.antenna.nl/unimacro/grammars/globalgrammars/lines/implementationdetails.html')
            return 0
        
    # def metaaction_makeunicodestrings(self, dummy=None):
    #     """for doctest testing, put u in front of every string
    #     """
    #     print('metaaction_makeunicodestrings, for Komodo')
    #     sendkeys("{ctrl+c}")
    #     t = unimacroutils.getClipboard()
    #     print(f'in: {t}')
    #     t = replaceStringToUnicode(t)
    #     unimacroutils.setClipboard(t)
    #     sendkeys("{ctrl+v}{down}")

def _test():
    """do doctests, for changing function
    """
    import doctest
    doctest.testmod()
    
