"""
UnittestAutohotkeyactions.py

This module tests the actions, that are performed by autohotkey (autohotkeyactions)

Quintijn Hoogenboom, 2021
"""
import unittest
import time
from pathlib import Path
from profilehooks import profile
from dtactions import autohotkeyactions as ahk
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

class UnittestAutohotkeyactions(unittest.TestCase):
    """test actions of the module "autohotkeyactions"
    """
    def setUp(self):
        self.notepadHndles = []        
    def tearDown(self):
        for hndle in self.notepadHndles:
            if ahk.SetForegroundWindow(hndle) == 0:
                ahk.killWindow()

    @profile(filename=dataDir/'profileinfo.txt')
    def testSimple(self):
        """only testing GetForegroundWindow and getProgInfo (or GetProgInfo)
        
        Also used for timing the getProgInfo routine
        
        """
        thisHndle = ahk.GetForegroundWindow()
        thisProgInfo = ahk.getProgInfo()
        hndleFromGetProgInfo = thisProgInfo.hndle   # or thisProgInfo[-1]
        mess = '\n'.join(['',
            'GetForegroundWindow and getProgInfo should return the same hndle but:',
            f'getForegroundWindow: {thisHndle}',
            f'from getProgInfo: {hndleFromGetProgInfo}', '='*40])
        self.assertEqual(thisHndle, hndleFromGetProgInfo, mess)
        
        notepadHndle = ahk.ahkBringup("notepad")[-1]
        
        for _ in range(5):
            ahk.SetForegroundWindow(notepadHndle)
            sendkeys('xxx')
            ahk.SetForegroundWindow(thisHndle)

        time.sleep(1)
        ahk.killWindow(notepadHndle)
        ahk.SetForegroundWindow(notepadHndle)
        ahk.SetForegroundWindow(thisHndle)
        ahk.SetForegroundWindow(notepadHndle)
        result = ahk.SetForegroundWindow(123)
        mess = '\n'.join(['',
            'SetForegroundWindow should not return hndle of non existing window (123):',
            f'expected: {notepadHndle}',
            f'got {result}', '='*40])
        
        self.assertNotEqual(result, 123, mess)


    def testKillWindow(self):
        """test the autohotkey version of killWindow
        """
        thisHndle = ahk.GetForegroundWindow()
        # thisProgInfo = ahk.getProgInfo()
        
        ## empty notepad window close again:    
        notepadInfo = ahk.ahkBringup("notepad")
        notepadHndle = notepadInfo.hndle
        self.assertTrue(notepadHndle > 0, "notepad should have a valid window hndle, not {notepadHndle}")
        
        ## create a child window
        sendkeys("{ctrl+o}")
        time.sleep(0.5)
        childInfo = ahk.getProgInfo()
        childHndle = childInfo[-1]
        
        # print(f'notepad, notepadHndle: {notepadHndle}, child: {childHndle}\n{childInfo}')

        result = ahk.SetForegroundWindow(thisHndle)        
        ## failed experiment: cannot find back the child window, when you get the top window in the foreground:
        # self.assertTrue(result is True, f'calling window should be in the foreground again {thisHndle}')
        # result = ahk.SetForegroundWindow(notepadHndle)        
        # self.assertTrue(result is True, f'calling notepadHndle seems to succeed, but child window is in front\n\tnotepad: {notepadInfo}\n\tchildwindow: {childHndle}')
        # 
        # sendkeys("{alt+tab}")
        # time.sleep(0.5)
        # sendkeys("{alt+tab}")
        # time.sleep(0.5)
        # 
        # result = ahk.GetForegroundWindow()
        
        # when the hndle of the open child window is given, this works all right
        # the child window is closed, with {esc}, and then the normal killWindow procedure follows.
        result = ahk.killWindow(childHndle)
        self.assertTrue(result is True, f'result of killing notepad should be 0, not {result}')

        ## now with text in the window:
        notepadHndle = ahk.ahkBringup("notepad")[-1]
        self.assertTrue(notepadHndle > 0, "notepad should have a valid window hndle, not {notepadHndle}")

        ## print a line of text:
        sendkeys("the quick brown fox...")

        result = ahk.SetForegroundWindow(thisHndle)
        self.assertTrue(result is True, f'calling window should be in the foreground again {thisHndle}')

        result = ahk.killWindow(notepadHndle)
        self.assertTrue(result is True, f'result of killing notepad should be 0, not {result}')
        
        result = ahk.SetForegroundWindow(thisHndle)        
        self.assertTrue(result is True, f'calling window should be in the foreground again {thisHndle}')

        ####### close with wrong key::::::::::::::::::::::::::::
        ## empty notepad window close again:    
        notepadHndle = ahk.ahkBringup("notepad")[-1]
        self.assertTrue(notepadHndle > 0, "notepad should have a valid window hndle, not {notepadHndle}")
        
        result = ahk.SetForegroundWindow(thisHndle)
        self.assertTrue(result is True, f'calling window should be in the foreground again {thisHndle}')

        result = ahk.killWindow(notepadHndle, key_close="{alt+f5}")   # should be {alt+f4}
        self.assertFalse(result is True, f'result of killing notepad should NOT be True\n\t{result}')

        result = ahk.killWindow(notepadHndle)   # should be {alt+f4}
        self.assertTrue(result is True, f'result of killing notepad should be True, not\n\t{result}')
        
        result = ahk.SetForegroundWindow(thisHndle)        
        self.assertTrue(result is True, f'calling window should be in the foreground again {thisHndle}')
        
        ## now with text in the window:
        notepadHndle = ahk.ahkBringup("notepad")[-1]
        self.assertTrue(notepadHndle > 0, "notepad should have a valid window hndle, not {notepadHndle}")

        ## print a line of text:
        sendkeys("the quick brown fox...")

        result = ahk.SetForegroundWindow(thisHndle)
        self.assertTrue(result is True, f'calling window should be in the foreground again {thisHndle}')

        result = ahk.killWindow(notepadHndle, key_close_dialog="{alt+m}")
        self.assertFalse(result is True, f'result of killing notepad should be not True\n\t{result}')

        ## now do it good:
        result = ahk.killWindow()
        self.assertTrue(result is True, f'result of killing notepad should be True, not:\n\t{result}')
        
        result = ahk.SetForegroundWindow(thisHndle)        
        self.assertTrue(result is True, f'calling window should be in the foreground again {thisHndle}')

def run():
    """run the unittest procedure"""
    print('starting UnittestAutohotkeyactions')

    suite = unittest.makeSuite(UnittestAutohotkeyactions, 'test')
    unittest.TextTestRunner().run(suite)

if __name__ == "__main__":
    run()
