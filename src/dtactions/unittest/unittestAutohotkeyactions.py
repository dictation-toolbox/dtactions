"""
UnittestAutohotkeyactions.py

This module tests the actions, that are performed by autohotkey (autohotkeyactions)

Quintijn Hoogenboom, 2021
"""
import unittest
import time
from pathlib import Path
from profilehooks import profile
from dtactions import autohotkeyactions
from dtactions.sendkeys import sendkeys

try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print('\n'.join(['If trying to test this in a git cloned package,',
          'please run this module after "build_package" and "flit install --symlink"',
          'otherwise, this is an unexpected error', 'please report']))
    raise

thisDir = getThisDir(__file__)
dtactionsDir = thisDir.parent

dataDir = Path.home()/".autohotkey"
checkDirectory(dataDir, create=True)

logFileName = dataDir/"testresult.txt"
print(f'output will be logged in {logFileName}')
print('start UnittestAutohotkeyactions', file=open(logFileName, 'w'))

class UnittestAutohotkeyactions(unittest.TestCase):
    """test actions of the module "autohotkeyactions"
    """
    def setUp(self):
        self.notepadHndles = []        
    def tearDown(self):
        for hndle in self.notepadHndles:
            if autohotkeyactions.SetForegroundWindow(hndle) == 0:
                autohotkeyactions.killWindow()


    def testKillWindow(self):
        """test the autohotkey version of killWindow
        """
        thisHndle = autohotkeyactions.GetForegroundWindow()
        # thisProgInfo = autohotkeyactions.getProgInfo()
        
        ## empty notepad window close again:    
        notepadInfo = autohotkeyactions.ahkBringup("notepad")
        notepadHndle = notepadInfo.hndle
        self.assertTrue(notepadHndle > 0, "notepad should have a valid window hndle, not {notepadHndle}")
        
        ## create a child window
        sendkeys("{ctrl+o}")
        time.sleep(0.5)
        childInfo = autohotkeyactions.getProgInfo()
        childHndle = childInfo[-1]
        
        print(f'notepad, notepadHndle: {notepadHndle}, child: {childHndle}\n{childInfo}')

        result = autohotkeyactions.SetForegroundWindow(thisHndle)        
        ## failed experiment: cannot find back the child window, when you get the top window in the foreground:
        # self.assertTrue(result == 0, f'calling window should be in the foreground again {thisHndle}')
        # result = autohotkeyactions.SetForegroundWindow(notepadHndle)        
        # self.assertTrue(result == 0, f'calling notepadHndle seems to succeed, but child window is in front\n\tnotepad: {notepadInfo}\n\tchildwindow: {childHndle}')
        # 
        # sendkeys("{alt+tab}")
        # time.sleep(0.5)
        # sendkeys("{alt+tab}")
        # time.sleep(0.5)
        # 
        # result = autohotkeyactions.GetForegroundWindow()
        
        # when the hndle of the open child window is given, this works all right
        # the child window is closed, with {esc}, and then the normal killWindow procedure follows.
        result = autohotkeyactions.killWindow(childHndle)
        self.assertTrue(result is True, f'result of killing notepad should be 0, not {result}')

        ## now with text in the window:
        notepadHndle = autohotkeyactions.ahkBringup("notepad")[-1]
        self.assertTrue(notepadHndle > 0, "notepad should have a valid window hndle, not {notepadHndle}")

        ## print a line of text:
        sendkeys("the quick brown fox...")

        result = autohotkeyactions.SetForegroundWindow(thisHndle)
        self.assertTrue(result == 0, f'calling window should be in the foreground again {thisHndle}')

        result = autohotkeyactions.killWindow(notepadHndle)
        self.assertTrue(result is True, f'result of killing notepad should be 0, not {result}')
        
        result = autohotkeyactions.SetForegroundWindow(thisHndle)        
        self.assertTrue(result == 0, f'calling window should be in the foreground again {thisHndle}')

        ####### close with wrong key::::::::::::::::::::::::::::
        ## empty notepad window close again:    
        notepadHndle = autohotkeyactions.ahkBringup("notepad")[-1]
        self.assertTrue(notepadHndle > 0, "notepad should have a valid window hndle, not {notepadHndle}")
        
        result = autohotkeyactions.SetForegroundWindow(thisHndle)
        self.assertTrue(result == 0, f'calling window should be in the foreground again {thisHndle}')

        result = autohotkeyactions.killWindow(notepadHndle, key_close="{alt+f5}")   # should be {alt+f4}
        self.assertFalse(result is True, f'result of killing notepad should NOT be True\n\t{result}')

        result = autohotkeyactions.killWindow(notepadHndle)   # should be {alt+f4}
        self.assertTrue(result is True, f'result of killing notepad should be True, not\n\t{result}')
        
        result = autohotkeyactions.SetForegroundWindow(thisHndle)        
        self.assertTrue(result == 0, f'calling window should be in the foreground again {thisHndle}')
        
        ## now with text in the window:
        notepadHndle = autohotkeyactions.ahkBringup("notepad")[-1]
        self.assertTrue(notepadHndle > 0, "notepad should have a valid window hndle, not {notepadHndle}")

        ## print a line of text:
        sendkeys("the quick brown fox...")

        result = autohotkeyactions.SetForegroundWindow(thisHndle)
        self.assertTrue(result == 0, f'calling window should be in the foreground again {thisHndle}')

        result = autohotkeyactions.killWindow(notepadHndle, key_close_dialog="{alt+m}")
        self.assertFalse(result is True, f'result of killing notepad should be not True\n\t{result}')

        ## now do it good:
        result = autohotkeyactions.killWindow()
        self.assertTrue(result is True, f'result of killing notepad should be True, not:\n\t{result}')
        
        result = autohotkeyactions.SetForegroundWindow(thisHndle)        
        self.assertTrue(result == 0, f'calling window should be in the foreground again {thisHndle}')




def log(text):
    """print text and log to logFile
    """
    print(text)
    print(text, file=open(logFileName, "a"))

def run():
    """run the unittest procedure"""
    print('starting UnittestAutohotkeyactions')

    suite = unittest.makeSuite(UnittestAutohotkeyactions, 'test')
    unittest.TextTestRunner().run(suite)

if __name__ == "__main__":
    print(f'run the tests, result will be in {logFileName}')
    run()
