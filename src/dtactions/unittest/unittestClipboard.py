"""
Python Macro Language for Dragon NaturallySpeaking
  (c) Copyright 1999 by Joel Gould
  Portions (c) Copyright 1999 by Dragon Systems, Inc.

unittestClipboard.py

This module tests the clipboard module, natlinkclipboard
now in the dtactions repository of the dictation toolbox
as developed by Christo Butcher clipboard.py for Dragonfly and
enhanced by Quintijn Hoogenboom, 2019, 2021
"""
# pylint: disable=C0115, R0902, C0116, W0201, R0915
import unittest
import time
from pathlib import Path
import win32gui
from dtactions from dtactions import natlinkclipboard
from dtactions import autohotkeyactions as ahk
from dtactions.unimacro import unimacroutils
from dtactions.sendkeys import sendkeys
try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print('Run this module after "build_package" and "flit install --symlink"\n')
    raise

thisDir = getThisDir(__file__)
dtactionsDir = thisDir.parent

dataDir = Path.home()/".dtactions"
checkDirectory(dataDir)

logFileName = dataDir/"testresult.txt"
print(f'output will be logged in {logFileName}')
print('start UnittestActions', file=open(logFileName, 'w'))

testFilesDir = thisDir/'test_clipboardfiles'
checkDirectory(testFilesDir, create=True)

class TestError(Exception):
    pass

#---------------------------------------------------------------------------
# These tests should be run after we call natConnect
# no reopen user at each test anymore..
# no default open window (open window will be the calling program)
# default .ini files pop up when you first run this test. just ignore them.
# the recording of print presents problems.
# All should go to testresult.txt in this same directory
class UnittestClipboard(unittest.TestCase):
    """testing the clipboard functions of dtactions.natlinkclipboard
    """
    def setUp(self):
        # self.connect()  # switched off
        self.thisHndle = ahk.GetForegroundWindow()
        self.org_text = "xyz"*3
        natlinkclipboard.Clipboard.set_system_text(self.org_text)
        # self.setupWindows()
        self.tempFileHndles = []
        self.killActions = {}   # put a teardown action here if {alt+f4} is not OK
        self.modInfos = {}      # the module info of each hndle (tuple of length 3)
        # possible handles in different tests...
        self.docx0Hndle = self.docx1Hndle = self.docx2Hndle = None
        self.text0Hndle = self.text1Hndle = self.text2Hndle = None
        self.expl0Hndle = None
        self.thundHndle = None   # thunderbird new mail message
        self.frescHndle = None   # Frescobaldi (music type setting)

    def tearDown(self):
        for hndle in self.tempFileHndles:
            if not win32gui.IsWindow(hndle):
                print("???window does not exist: %s"% hndle)
                continue
            # if hndle == self.thisHndle:
            #     # print('window hndle %s may not match "thisHndle": %s'% (hndle, self.thisHndle))
            #     continue
            # print('close window with hndle: %s'% hndle)

            # should not be necessary:
            # curHndle = ahk.SetForegroundWindow(hndle)

            if hndle in self.killActions:
                # unimacroutils.SetForegroundWindow(self.thisHndle)
                # place to break in debug mode
                # unimacroutils.SetForegroundWindow(curHndle)
                ahk.killWindow(self.killActions[hndle])
        _result = ahk.SetForegroundWindow(self.thisHndle)
        # self.disconnect()  # disabled, natConnect
        notClosedHndles = []
        for hndle in self.tempFileHndles:
            if hndle and win32gui.IsWindow(hndle):
                notClosedHndles.append(hndle)

        if self.frescHndle:
            if self.frescHndle in notClosedHndles:
                notClosedHndles.remove(self.frescHndle)
            else:
                raise TestError('Fresobaldi should not be closed after the test')
        if notClosedHndles:
            raise TestError('Hndles not closed, but should: %s'% notClosedHndles)

    # def connect(self):
    #     # start with 1 for thread safety when run from pythonwin:
    #     # natlink.natConnect(natconnectOption)
    #     # sys.stdout = open(logFileName, 'a')
    #     # sys.stderr = open(logFileName, 'a')
    #     self.isConnected = False
    # 
    # def disconnect(self):
    #     # natlink.natDisconnect()
    #     self.isConnected = False

    def getTextViaClipboard(self):
        """get text of window and selection or cursor via copy paste actions

        pass waitTime for testing response time.

        waiting_iterations is default 10, but changed to 3 if an empty clipboard is expected

        return: text, selStart, selEnd, cursorPos

        """
        #pylint:disable=R0201
        cb = natlinkclipboard.Clipboard(save_clear=True, debug=1)
        # waitTime = 0.001 defaults now...

        # unimacroutils.SetForegroundWindow(self.thisHndle)
        # unimacroutils.SetForegroundWindow(self.testHndle)

        ## first try if there is a selection:
        sendkeys("{ctrl+c}")
        text = cb.get_text()
        return text

    def getThisDirTextViaClipboard(self, waitTime):
        """get text of window and selection or cursor via copy paste actions

        pass waitTime for testing response time.

        waiting_iterations is default 10, but changed to 3 if an empty clipboard is expected

        return: text, selStart, selEnd, cursorPos

        """
        # pylint: disable=R0912, R0915
        cb = natlinkclipboard.Clipboard(save_clear=True, debug=1)
        waitTime = 0.001

        # unimacroutils.SetForegroundWindow(self.thisHndle)
        # unimacroutils.SetForegroundWindow(self.testHndle)

        ## first try if there is a selection:
        sendkeys("{ctrl+c}")
        selection = cb.get_text()
        if selection:
            print('selection: %s'% selection)
            sendkeys("{left}")
            cb.clear_clipboard()


        # now try text left of selection:
        sendkeys("{ctrl+shift+home}{ctrl+c}")

        # unimacroutils.SetForegroundWindow(self.thisHndle)
        # unimacroutils.SetForegroundWindow(self.testHndle)


        before = cb.get_text(waitTime)
        if before:
            sendkeys("{right}")
            cb.clear_clipboard()
        else:
            print("at start of document")

        # unimacroutils.SetForegroundWindow(self.thisHndle)
        # unimacroutils.SetForegroundWindow(self.testHndle)

        text = before + selection
        selStart = len(before)
        selEnd = len(text)

        # skip to end of selection:
        if selection:
            sendkeys("{right %s}"% len(selection))

        # unimacroutils.SetForegroundWindow(self.thisHndle)
        # unimacroutils.SetForegroundWindow(self.testHndle)

        if text:
            # now select to end:
            sendkeys("{shift+ctrl+end}{ctrl+c}")
            after = cb.get_text(waitTime)
            print("after: %s"% repr(after))

            if not after:
                if self.testHndle == self.thundHndle:
                    ## after an empty selection, go one to the left!!
                    sendkeys("{left}")
                after = ""
            else:
                # undo selection
                sendkeys("{left}")
            if after.endswith('\0'):
                after = after.replace('\0', '')
                print("removed null char at end of selection (1): %s"% after)
                # after = afterplusone[1:]
        else:
            # presumably at start of buffer
            sendkeys("{shift+ctrl+end}{ctrl+c}")
            after = cb.get_text()
            if after:
                # undo selection
                sendkeys("{left}")
            else:
                pass
            if after.endswith('\0'):
                after = after.replace('\0', '')
                print("removed null char at end of selection (2): %s"% after)
                # after = afterplusone[1:]
                cb.clear_clipboard()

        if selection:
            lensel = len(selection)
            # select from left to right;

            # unimacroutils.SetForegroundWindow(self.thisHndle)
            # unimacroutils.SetForegroundWindow(self.testHndle)

            sendkeys("{left %s}{shift+right %s}"% (lensel, lensel))

        text = text + after
        cursorPos = selEnd

        # unimacroutils.SetForegroundWindow(self.thisHndle)
        # unimacroutils.SetForegroundWindow(self.testHndle)


        return text, selStart, selEnd, cursorPos




    def setupWindows(self):
        """make several windows to which can be switched, and which can be copied from

        TODOQH: make bringup of a folder and copy paste of the folder info work... (if somebody wants this)
        """
        self.thisHndle = ahk.GetForegroundWindow()

        dirWindows = "C:\\windows"
        result = ahk.autohotkeyBringup(app="explore", filepath=dirWindows)
        if not result:
            print('no result for %s'% dirWindows)
            return
        hndle = result.hndle
        self.tempFileHndles.append(hndle)
        self.expl0Hndle = hndle
        unimacroutils.SetForegroundWindow(self.thisHndle)

    def setupDocxFile(self):
        """open Word file and do the testing on it...
        """
        docxFile2 = "natlink.docx"
        docxPath2 = testFilesDir/docxFile2
        if not docxPath2.is_file():
            raise OSError('file does not exist: %s'% docxPath2)
        result = ahk.autohotkeyBringup(app=None, filepath=docxPath2)
        if result:
            hndle = result.hndle
        else:
            raise TestError("autohotkeyBringup of Word gives no result: %s"% result)
        self.tempFileHndles.append(hndle)
        unimacroutils.SetForegroundWindow(self.thisHndle)
        print('testempty.docx (docx2Hndle): %s'% result)
        self.tempFileHndles.append(result)
        self.docx2Hndle = result

    def setupTextFiles(self):
        """make some text files for clipboard testing
        """
        textFile0 = "testempty.txt"
        textPath0 = testFilesDir/textFile0
        open(textPath0, 'w')
        result = ahk.autohotkeyBringup(app=None, filepath=textPath0)
        self.assertTrue(isinstance(result, tuple), 'result of autohotkeyBringup should be a tuple (namedtuple)')

        hndle = result.hndle
        unimacroutils.SetForegroundWindow(self.thisHndle)
        print('testempty (text0Hndle): %s'% hndle)
        self.tempFileHndles.append(hndle)
        self.text0Hndle = hndle

        textFile1 = "testsmall.txt"
        textPath1 = testFilesDir/textFile1
        self.text1Txt = "small abacadabra\n"*2
        open(textPath1, 'w').write(self.text1Txt)
        result = ahk.autohotkeyBringup(app=None, filepath=textPath1)
        hndle = result.hndle
        self.tempFileHndles.append(hndle)
        unimacroutils.SetForegroundWindow(self.thisHndle)
        print('testsmall (text1Hndle): %s'% hndle)
        self.text1Hndle = hndle
        self.tempFileHndles.append(hndle)

        textFile2 = "testlarge.txt"
        textPath2 = testFilesDir/textFile2
        self.text2Txt = "large abacadabra\n"*1000
        open(textPath2, 'w').write(self.text2Txt)
        result = ahk.autohotkeyBringup(app=None, filepath=textPath2)
        hndle = result.hndle
        self.tempFileHndles.append(hndle)
        unimacroutils.SetForegroundWindow(self.thisHndle)
        print('testlarge (text2Hndle): %s'% hndle)
        self.text2Hndle = hndle
        self.killActions[self.text0Hndle] = f'KW({self.text0Hndle})'
        self.killActions[self.text1Hndle] = f'KW({self.text1Hndle})'
        self.killActions[self.text2Hndle] = f'KW({self.text2Hndle})'


        self.tempFileHndles.append(hndle)

    def setupThunderbirdNewWindow(self):
        """start an empty mail message in Thunderbird

        at shutdown do not kill this window
        """
        ## this extraaow, and goes to the message pane:
        extra = '\n'.join(['Send, ^n',
                'SetTitleMatchMode, 2',
                'WinWaitNotActive, ##title##',
                'Sleep, 100',
                'Send, {tab 2}'])     # note SetTitleMatchMode and Sleep, apparently necessary

        result = ahk.ahkBringup(app="thunderbird.exe", title="Mozilla Thunderbird", extra=extra)
        print(f'setupThunderbirdNewWindow, after ahkBringup: {result}')
        ahk.do_ahk_script(extra)
        hndle = ahk.GetForegroundWindow()
        print(f'GetForegroundWindow: {hndle}')
        self.tempFileHndles.append(hndle)
        self.thundHndle = hndle
        self.killActions[hndle] = "KW({ctrl+w}, w)"   # kill the window, with letter w if windows has text....
        self.modInfos[hndle] = result
        sendkeys("Thunderbird new window abacadabra")

    def setupFrescobaldiNewPane(self):
        """start an empty pane in Frescobaldi (front end for Lilypond music type setter)

        at shutdown do not kill this window
        """
        ## this extra makes an empty window, and goes to the message pane:
        extra = '\n'.join(["Sleep, 100", 'Send, ^n', "Sleep, 100"])

        result = ahk.ahkBringup(app=r"C:\Program Files (x86)\Frescobaldi\frescobaldi.exe", title="Frescobaldi",
                                extra=extra)
        hndle = result.hndle
        sendkeys("Frescobaldi abacadabra")
        self.tempFileHndles.append(hndle)
        self.frescHndle = hndle
        self.killActions[hndle] = "KW({ctrl+w}, {right}{enter})"
        self.modInfos[hndle] = result

    def testEmpty(self):
        """only testing the setup and teardown

        implicitly testing actions (via autohotkey, AHK) like starting a process and switching to a window handle

        (this presented a lot of trouble, getting the correct files open with the correct window handles)

        Now with AHK on, this runs more smooth than ever. With or without Dragon running.
        """
        print("doing an empty test, just to test setUp and tearDown")
        self.setupTextFiles()
        # self.setupDocxFile()
        # self.setupWindows()
        # self.setupThunderbirdNewWindow()
        # self.setupFrescobaldiNewPane()
        unimacroutils.SetForegroundWindow(self.thisHndle)
        # print("Frescobaldi handle: ", self.frescHndle)
        # unimacroutils.SetForegroundWindow(self.thisHndle)

    def tttestCopyClipboardGetSetText(self):
        """use the clipboard for getting and the window text

        Run with one target at a time


        """
        print("testCopyClipboardSwitching")
        # self.setupDocxFile()
        self.setupThunderbirdNewWindow()
        self.testHndle = self.thundHndle
        # self.setupFrescobaldiNewPane()
        # self.testHndle = self.frescHndle
        # self.setupTextFiles()


        for i in range(1, 21, 10):
            waitTime = 0.001/i

            t0 = time.time()

            ## start with empty window
            sendkeys("{ctrl+a}{del}")

            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            expText = ""
            expPosTuple = (0,0,0)  # startSel, endSel, cursorPos
            self.assertTrue(expText == text, "text not as expected")
            self.assertTrue(expPosTuple == (startSel, endSel, cursorPos), "positions not as expected")

            sendkeys("Hello world")

            # these pair of lines provide the possibility to debug the tests
            # put the breakpoint at the second line...
            # unimacroutils.SetForegroundWindow(self.thisHndle)
            # unimacroutils.SetForegroundWindow(self.testHndle)


            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            expText = "Hello world"
            expPosTuple = (11, 11, 12)  # startSel, endSel, cursorPos
            self.assertTupleEqual(text, expText, "text not as expected")
            self.assertTupleEqual((startSel, endSel, cursorPos), expPosTuple, "positions not as expected")

            # now select word world
            sendkeys("{shift+ctrl+left}")

            # unimacroutils.SetForegroundWindow(self.thisHndle)
            # unimacroutils.SetForegroundWindow(self.testHndle)


            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            expText = "Hello world"
            expPosTuple = (6, 11, 11)  # startSel, endSel, cursorPos
            self.assertEqual(expText, text, "text not as expected")
            self.assertEqual(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            ## and again:
            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            self.assertEqual(expText, text, "text not as expected")
            self.assertEqual(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            # unimacroutils.SetForegroundWindow(self.thisHndle)
            # unimacroutils.SetForegroundWindow(self.testHndle)

            # now replace the word world, and make a second line
            sendkeys('SCLIP(WORLD{enter}How are you going?{enter}In this Corona world?)')
            expText = "Hello WORLD\nHow are you going?\nIn this Corona world?"
            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            expPosTuple = (52, 52, 52)  # startSel, endSel, cursorPos
            self.assertEqual(expText, text, "text not as expected")
            self.assertEqual(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            ## and again:
            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            self.assertEqual(expText, text, "text not as expected")
            self.assertEqual(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            # select all
            sendkeys("{ctrl+home}{shift+ctrl+end}")
            expText = "Hello WORLD\nHow are you going?\nIn this Corona world?"
            # if self.testHndle == self.thundHndle:
            #     expText += "\n"
            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            expPosTuple = (0, len(expText), len(expText))  # startSel, endSel, cursorPos
            self.assertEqual(expText, text, "text not as expected")
            self.assertEqual(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            # and again:
            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            expPosTuple = (0, len(expText), len(expText))  # startSel, endSel, cursorPos
            self.assertEqual(expText, text, "text not as expected")
            self.assertEqual(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            # select "are you"
            sendkeys("{ctrl+home}{down}{right 4}{shift+right 7}")
            expText = "Hello WORLD\nHow are you going?\nIn this Corona world?"
            # not here!!
            # if self.testHndle == self.thundHndle:
            #     expText += "\n"
            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            expPosTuple = (16, 23, 23) # startSel, endSel, cursorPos
            self.assertEqual(expText, text, "text not as expected")
            self.assertEqual(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            ## and again:
            text, startSel, endSel, cursorPos = self.getTextViaClipboard()
            expPosTuple = (16, 23, 23) # startSel, endSel, cursorPos
            self.assertEqual(expText, text, "text not as expected")
            self.assertEqual(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")


            t1 = time.time()
            elapsed = t1 - t0
            print("time %.6f for round with waitTime: %.6f "% (elapsed, waitTime))


    def tttestCopyClipboardSwitching(self):
        """test the copying of the clipboard

        with empty, small and larger .txt file.

        .docx does not run smooth


        """
        print("testCopyClipboardSwitching")
        # self.setupDocxFile()
        # self.setupWindows()
        # self.setupThunderbirdNewWindow()
        self.setupFrescobaldiNewPane()
        self.setupTextFiles()

        # cb = natlinkclipboard.Clipboard(save_clear=True, debug=2)   # debug flag can be set or omitted (or set to 0)

        ## longer test:
        # for i in range(1, 10, 5):
        #     waitTime = 0.001/i
        #     print('---- round: %s, waitTime %.4f '% (i, waitTime))
        # 
        #     ## first clear the clipboard:
        #     cb.clear_clipboard()
        #     got = cb.get_text(waitTime)
        #     print('after clear_clipboard, text: %s (cb: %s)'% (got, cb))
        #     self.assertEqual("", got, "should have no text now, waitTime: %s (times 4)"% waitTime)
        # 
        #     print("now testempty.txt -----------------")
        #     # empty file:
        #     unimacroutils.SetForegroundWindow(self.text0Hndle)
        #     sendkeys("{ctrl+a}{ctrl+c}")
        #     got = cb.get_text(waitTime)
        #     print('after select all copy of empty file, no change should have happened:')
        #     # print('got: %s, (cb: %s)'% (got, cb))
        #     self.assertEqual("", got, "should have no text now, waitTime: %s"% waitTime)
        # 
        #     unimacroutils.SetForegroundWindow(self.thisHndle)
        # 
        #     print("now testsmall.txt -----------------")
        #     unimacroutils.SetForegroundWindow(self.text1Hndle)
        #     sendkeys("{ctrl+a}{ctrl+c}")
        #     got = cb.get_text(waitTime)
        #     print('after select all copy of testsmall.txt')
        #     # print('got: %s, (cb: %s)'% (got, cb))
        #     exp = self.text1Txt
        #     self.assertEqual(exp, got, "testsmall should have two lines of text, waitTime: %s"% waitTime)
        # 
        # 
        #     # test large.txt
        #     print("now testlarge.txt -----------------")
        #     unimacroutils.SetForegroundWindow(self.text2Hndle)
        #     sendkeys("{ctrl+a}{ctrl+c}")
        #     cb.get_text(waiting_interval=waitTime)
        #     got = cb.get_text()
        #     if got:
        #         lengot = len(got)
        #     else:
        #         print('testlarge, no result from get_text: %s'% cb)
        #     exp = len(self.text2Txt)
        #     self.assertEqual(exp, lengot, "should have long text now, waitTime: %s"% waitTime)
        #     # empty for the next round:
        #     # cb.set_text("")
        #     time.sleep(waitTime)
        # 
        #     unimacroutils.SetForegroundWindow(self.thisHndle)
        # 
        #     # test natlink.docx
        #     print("now natlink.docx (%s)-----------------"% self.docx2Hndle)
        #     if not self.docx2Hndle:
        #         print("word document not available: %s"% (self.docx2Hndle))
        #         continue
        #     unimacroutils.SetForegroundWindow(self.docx2Hndle)
        #     sendkeys("{ctrl+a}{ctrl+c}")
        #     cb.get_text(waiting_interval=waitTime)
        #     got = cb.get_text()
        #     if got:
        #         lengot = len(got)
        #     else:
        #         print('testlarge, no result from get_text: %s'% cb)
        #     exp = len(self.text2Txt)
        #     self.assertEqual(exp, lengot, "should have long text now, waitTime: %s"% waitTime)
        #     # empty for the next round:
        #     # cb.set_text("")
        #     time.sleep(waitTime)
        # 
        # del cb
        # got_org_text = natlinkclipboard.Clipboard.get_system_text()
        # self.assertEqual(self.org_text, got_org_text, "restored text from clipboard not as expected")


    def tttestSetForegroundWindow(self):
        """test switching the different windows, including this

        This is a side test for getting the testing system at work...

        and should work without problems with the AutoHotkey script...

        About 0.15 seconds to get a window in the foreground...
        """
        # self.setupDocxFile()
        self.setupTextFiles()            # self.text0Hndle, ..., self.text2Hndle
        self.setupWindows()
        self.setupFrescobaldiNewPane()   # self.frescHndle
        self.setupThunderbirdNewWindow() # self.thundHndle

        thisHndle = unimacroutils.GetForegroundWindow()
        print("thisHndle: %s"% thisHndle)
        unknownHndle = 5678659
        hndles = [thisHndle, self.docx2Hndle, self.docx1Hndle, self.text0Hndle, self.text1Hndle, self.text2Hndle,
                  self.frescHndle, self.thundHndle, unknownHndle]
        t0 = time.time()
        rounds = 10
        for i in range(1, rounds+1):
            print('start round %s'% i)
            for h in hndles:
                if not h:
                    continue
                result = unimacroutils.SetForegroundWindow(h)
                if result:
                    self.assertNotEqual(unknownHndle, h, "hndle should not be unknown (fake) hndle")
                else:
                    self.assertEqual(unknownHndle, h, "hndle should be one of the valid hndles")

            time.sleep(0.1)
        t1 = time.time()
        deliberate = rounds*0.1
        total = t1 - t0
        nettime = t1 - t0 - deliberate
        nswitches = rounds*len(hndles)
        timeperswitch = nettime/nswitches
        print('nswitches: %s, nettime: %.4f, time per switch: %.4f'% ( nswitches, nettime, timeperswitch))
        print('total: %s, deliberate: %s'% (total, deliberate))


    def log(self, t):
        #pylint:disable=R0201
        # if self.isConnected:
        #     natlink.displayText(t, 0)
        # else:
        
        print(t)
        print(t, file=open(logFileName, "a"))

def run():
    print('starting UnittestClipboard')
    # trick: if you only want one or two tests to perform, change
    # the test names to her example def test....
    # and change the word 'test' into 'tttest'...
    # do not forget to change back and do all the tests when you are done.
    # each test does a natConnect and a natDisconnect...
    # sys.stdout = open(logFileName, 'a')
    # sys.stderr = open(logFileName, 'a')

    suite = unittest.makeSuite(UnittestClipboard, 'test')
    unittest.TextTestRunner().run(suite)

if __name__ == "__main__":
    print("run, result will be in %s"% logFileName)
    run()
