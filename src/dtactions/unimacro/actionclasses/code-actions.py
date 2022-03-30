"""actions from application code (Visual Studio)

configuring instructions for Visual Studio at the bottom of this file

getting the current line number!
"""
import time
from dtactions.unimacro.actionclasses.actionbases import AllActions
from dtactions.unimacro.unimacroactions import doAction as action
from dtactions.unimacro.unimacroactions import doAction as action
from dtactions import natlinkclipboard

class  CodeActions(AllActions):
    def __init__(self, progInfo):
        AllActions.__init__(self, progInfo)
        
    def getCurrentLineNumber(self, handle=None):
        debug - 0
        t1 = time.time()
        if self.toporchild == "child":
            return 0
        cb = natlinkclipboard.Clipboard(save_clear=True, debug=debug)  # clear "debug" to get rid of timing line
        # via the command palette:
        # action("{shift+ctrl+p}; copy current line to clipboard; {enter};")
          
        # better via a shortcut key, goto file, preferences, keyboard shortcuts
        # type copy current line number to clipboard
        # press ctrl+alt+c    (choose different keystroke and change in next line!!)
        shortcutkey = "{ctrl+alt+c}"
        keystroke(shortcutkey)
        
        # now collect the clipboard, at most waiting 10 intervals of 0.1 second.
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
        
    
if __name__ == '__main__':
    pass        
        
# Note, the program name of Visual Studio Code is "code.exe", so this module is called code.py
#
# Please install the VS Code extension called 'copy-current-line-number'
# that copies the current editor line number to the clipboard.
#
# Next go to the File, Preferences, Keyboard Shortcuts.
# lookup your extension (copy current line number), and attach ctrl+alt+c to this extension.

# After you configure the grammar _lines.py, like:
# [general]
# line numbers modulo hundred = T
# and toggle the microphone, possibly restart Dragon,
# the feature should work.
#
# each line number below 100, will jump to the nearest line with respect to the current line.
        
        