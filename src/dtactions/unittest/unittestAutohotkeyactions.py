"""
UnittestAutohotkeyactions.py

This module tests the actions, that are performed by autohotkey (autohotkeyactions)

Quintijn Hoogenboom, 2021
"""
import unittest
import time
from pathlib import Path
# from profilehooks import profile
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

    # @profile(filename=dataDir/'profileinfo.txt')
    def tttestSimple(self):
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
        
        ahk.SetForegroundWindow(thisHndle)
        ahk.SetForegroundWindow(notepadHndle)
        ahk.SetForegroundWindow(thisHndle)
        
        
        ## create a child window
        # sendkeys("{ctrl+o}")
        # time.sleep(0.5)
        # childInfo = ahk.getProgInfo()
        # childHndle = childInfo[-1]
        # 
        # print(f'notepad, notepadHndle: {notepadHndle}, child: {childHndle}\n{childInfo}')

        # result = ahk.SetForegroundWindow(thisHndle)        
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
        # result = ahk.killWindow(childHndle)
        # self.assertTrue(result is True, f'result of killing notepad should be 0, not {result}')

        ## now with text in the window:
        notepadHndle = ahk.ahkBringup("notepad")[-1]
        self.assertTrue(notepadHndle > 0, "notepad should have a valid window hndle, not {notepadHndle}")

        ## print a line of text:
        sendkeys("the quick brown fox...")

        result = ahk.SetForegroundWindow(thisHndle)
        self.assertTrue(result is True, f'calling window should be in the foreground again {thisHndle}')

        result = ahk.killWindow(notepadHndle)
        self.assertTrue(result is True, f'result of killing notepad should be True, not {result}')
        
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

    def tttestAhkBringupNotePad(self):
        """test easy bringup of Notepad

        second call tries to activate to the first openend instance        
        """
        ### test notepad:
        thisHndle = ahk.GetForegroundWindow()
        notepadInfo = ahk.ahkBringup("notepad")
        notepadHndle = notepadInfo.hndle
        self.assertGreater(notepadHndle, 0, "notepad should have a valid window hndle, not {notepadHndle}")
        self.assertEqual(notepadInfo.prog, 'notepad', 'notepadInfo, prog does not match')
        notepadTitle = notepadInfo.title
        searchTitle = notepadTitle.split()[0]   ## only Naamloos
        
        # go back, then pick up window with same title:
        ahk.SetForegroundWindow(thisHndle)
        notepadInfo2 = ahk.ahkBringup("notepad", title=searchTitle)
        notepadHndle2 = notepadInfo2.hndle
        self.assertGreater(notepadHndle2, 0, "notepad 2 should have a valid window hndle, not {notepadHndle2}")
        self.assertEqual(notepadHndle2, notepadHndle, "notepad 2 should be same window as notepad")
        self.assertTupleEqual(notepadInfo2, notepadInfo2, 'notepadInfo2 should equal notepadInfo')
        
        ahk.killWindow(notepadHndle)
        ahk.killWindow(notepadHndle2)
        nowThisHndle = ahk.GetForegroundWindow()
        self.assertEqual(nowThisHndle, thisHndle, 'test should end up in same window as it started')

        ## open a filepath, the errormessagefromahk.txt file (1 empty line)
        filepath = Path.home()/'.autohotkey'/'errormessagefromahk.txt'
        self.assertTrue(filepath.is_file(), f'file {filepath} should exist')
        notepadInfo3 = ahk.ahkBringup("notepad", filepath=filepath)
        notepadHndle3 = notepadInfo3.hndle
        self.assertGreater(notepadHndle3, 0, "notepad 3 should have a valid window hndle, not {notepadHndle2}")
        self.assertEqual(notepadInfo3.title, str(filepath), f'filepath should be in window title: {filepath}')



    def tttestAhkBringupThunderbirdAppTitle(self):
        """test starting or bringing to the foreground Thunderbird
        
        Note the default waitForStart is increased from 1 to 3 seconds, as starting Thunderbird takes
        a bit of time.
        """
        ### test Thunderbird, app and title, activate if already open...
        thisHndle = ahk.GetForegroundWindow()
        ## if thunderbird is not open, and you remove waitForStart, you probably get the fail message below,
        ## as it takes more than a second to start thunderbird...
        thunderbirdInfo = ahk.ahkBringup("thunderbird", title="Mozilla Thunderbird", waitForStart=5)
        if isinstance(thunderbirdInfo, str):
            self.fail(f'starting/activating thunderbird failed, message: {thunderbirdInfo}')
        thunderbirdHndle = thunderbirdInfo.hndle
        
        ahk.SetForegroundWindow(thisHndle)
      
        self.assertGreater(thunderbirdHndle, 0, "thunderbird should have a valid window hndle, not {thunderbirdHndle}")
        self.assertEqual(thunderbirdInfo.prog, 'thunderbird', 'thunderbirdInfo, prog does not match')

        # leave thunderbird open after testing
        # ahk.killWindow(thunderbirdHndle)

        ahk.SetForegroundWindow(thisHndle)
        nowThisHndle = ahk.GetForegroundWindow()
        self.assertEqual(nowThisHndle, thisHndle, 'test should end up in same window as it started')
        
        ## switch to thunderbird twice, and then back:
        thunderbirdInfoTwo = ahk.ahkBringup("thunderbird", title="Mozilla Thunderbird", waitForStart=5)
        if isinstance(thunderbirdInfoTwo, str):
            self.fail(f'starting/activating thunderbird failed, message: {thunderbirdInfoTwo}')
        thunderbirdInfoThree = ahk.ahkBringup("thunderbird", title="Mozilla Thunderbird", waitForStart=5)
        if isinstance(thunderbirdInfoThree, str):
            self.fail(f'starting/activating thunderbird failed, message: {thunderbirdInfoThree}')

        ## back to calling window:        
        ahk.SetForegroundWindow(thisHndle)
        nowThisHndle = ahk.GetForegroundWindow()
        self.assertEqual(nowThisHndle, thisHndle, 'test should end up in same window as it started')

        

    def tttestAhkBringupWindword(self):
        """test starting or bringing to the foreground Windword
        
        """
        ### test winword, app and title, activate if already open...
        thisHndle = ahk.GetForegroundWindow()
        ## if winword is not open, and you remove waitForStart, you probably get the fail message below,
        ## as it takes more than a second to start winword...
        winwordInfo = ahk.ahkBringup("winword", title="Word", waitForStart=10)
        if isinstance(winwordInfo, str):
            self.fail(f'starting winword failed, message: "{winwordInfo}"')
        winwordHndle = winwordInfo.hndle
        
        ahk.SetForegroundWindow(thisHndle)
      
        self.assertGreater(winwordHndle, 0, "winword should have a valid window hndle, not {winwordHndle}")
        self.assertEqual(winwordInfo.prog, 'WINWORD', 'winwordInfo, prog does not match')

        # leave winword open after testing
        ahk.SetForegroundWindow(thisHndle)
        nowThisHndle = ahk.GetForegroundWindow()
        self.assertEqual(nowThisHndle, thisHndle, 'test should end up in same window as it started')

def run():
    """run the unittest procedure"""
    print('starting UnittestAutohotkeyactions')

    suite = unittest.makeSuite(UnittestAutohotkeyactions, 'test')
    unittest.TextTestRunner().run(suite)

if __name__ == "__main__":
    run()
