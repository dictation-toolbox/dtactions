"""
UnittestAutohotkeyactions.py

This module tests the actions, that are performed by autohotkey (autohotkeyactions)

Quintijn Hoogenboom, 2021
"""
import unittest
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
        pass        
    def tearDown(self):
        pass


    @profile(filename=dataDir/'profileinfo.txt')
    def testSimple(self):
        """only testing GetForegroundWindow and getProgInfo (or GetProgInfo)
        
        """
        thisHndle = autohotkeyactions.GetForegroundWindow()
        thisProgInfo = autohotkeyactions.getProgInfo()
        hndleFromGetProgInfo = thisProgInfo.hndle   # or thisProgInfo[-1]
        mess = '\n'.join(['',
            'GetForegroundWindow and getProgInfo should return the same hndle but:',
            f'getForegroundWindow: {thisHndle}',
            f'from getProgInfo: {hndleFromGetProgInfo}', '='*40])
        self.assertEqual(thisHndle, hndleFromGetProgInfo, mess)
        
        notepadHndle = autohotkeyactions.ahkBringup("notepad")[-1]
        
        for _ in range(5):
            autohotkeyactions.SetForegroundWindow(notepadHndle)
            autohotkeyactions.SetForegroundWindow(thisHndle)

        killWindow(notepadHndle, key_close="{alt+f4}", key_close_dialog="n")
        autohotkeyactions.SetForegroundWindow(notepadHndle)
        autohotkeyactions.SetForegroundWindow(thisHndle)
        autohotkeyactions.SetForegroundWindow(notepadHndle)
        result = autohotkeyactions.SetForegroundWindow(123)
        mess = '\n'.join(['',
            'SetForegroundWindow should not return hndle of non existing window (123):',
            f'expected: {notepadHndle}',
            f'got {result}', '='*40])
        
        self.assertNotEqual(result, 123, mess)
            
def killWindow(hndle, key_close=None, key_close_dialog=None):
    """kill the app with hndle,
    
    like the unimacro shorthand command KW (killwindow)
    """
    result = autohotkeyactions.SetForegroundWindow(hndle)
    if result:
        print(f'window {hndle} not any more available')
        return
    key_close = key_close or "{alt+f4}"
    key_close_dialog = key_close_dialog or "n"
    progInfo = autohotkeyactions.getProgInfo()
    foregroundProg = progInfo.prog
    if progInfo.hndle != hndle:
        print(f'invalid window {progInfo.hndle} in the foreground, want {hndle}')
        return
    if progInfo.toporchild == 'child':
        print(f'child window in the foreground, expected top {hndle}')
        return
    sendkeys(key_close)
    progInfo = autohotkeyactions.getProgInfo()
    if progInfo.prog != foregroundProg:
        return
    if progInfo.toporchild == 'child':
        sendkeys(key_close_dialog)

    progInfo = autohotkeyactions.getProgInfo()
    if progInfo.toporchild == 'child':
        print('killWindow, failed to close child dialog')
        return
     
    


def log(text):
    print(text, file=open(logFileName, "a"))

def run():
    print('starting UnittestAutohotkeyactions')

    suite = unittest.makeSuite(UnittestAutohotkeyactions, 'test')
    unittest.TextTestRunner().run(suite)

if __name__ == "__main__":
    print(f'run the tests, result will be in {logFileName}')
    run()
