#
# Python Macro Language for Dragon NaturallySpeaking
#   (c) Copyright 1999 by Joel Gould
#   Portions (c) Copyright 1999 by Dragon Systems, Inc.
#
# unittestClipboard.py
#
# This module tests the clipboard module, natlinkclipboard
# now in the dtactions repository of the dictation toolbox
# as developed by Christo Butcher clipboard.py for Dragonfly and
# enhanced by Quintijn Hoogenboom, 2019, 2021

#
import sys
import unittest
import types
import os
import os.path
import time
import win32gui
from pathqh import path

thisDir = path('.')
unimacroFolder = (thisDir/'..'/'..'/'unimacro'/'src'/'unimacro').normpath()
if os.path.isdir(unimacroFolder):
    if not unimacroFolder in sys.path:
        sys.path.append(unimacroFolder)
else:
    raise IOError(f'Invalid unimacro folder: {unimacroFolder}')

import TestCaseWithHelpers
import natlink
import natlinkclipboard



import actions
from actions import doAction as action
import natlinkutilsqh
import natlinkutils

class TestError(Exception):
    pass

natconnectOption = 0 # no threading has most chances to pass...




logFileName = thisDir/"testresult.txt"
print('printing will go to %s'% logFileName)
print('start unittestClipboard', file=open(logFileName, 'w'))

testFilesDir = thisDir/'test_clipboardfiles'
if not testFilesDir.isdir():
    testFilesDir.mkdir()

#---------------------------------------------------------------------------
# These tests should be run after we call natConnect
# no reopen user at each test anymore..
# no default open window (open window will be the calling program)
# default .ini files pop up when you first run this test. just ignore them.
# the recording of print presents problems.
# All should go to testresult.txt in this same directory
class UnittestClipboard(TestCaseWithHelpers.TestCaseWithHelpers):
    def setUp(self):
        self.connect()  # switched off

        self.thisHndle = win32gui.GetForegroundWindow()
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
            if hndle == self.thisHndle:
                # print('window hndle %s may not match "thisHndle": %s'% (hndle, self.thisHndle))
                continue
            # print('close window with hndle: %s'% hndle)
            natlinkutilsqh.SetForegroundWindow(hndle)
            curHndle = natlinkutilsqh.GetForegroundWindow()

            if hndle == curHndle:
                if hndle in self.killActions:
                    # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
                    # place to break in debug mode
                    # natlinkutilsqh.SetForegroundWindow(curHndle)
                    action(self.killActions[hndle], modInfo=self.modInfos[hndle])
                else:
                    natlinkutils.playString("{alt+f4}")
        natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        self.disconnect()  # disabled, natConnect
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
            raise TestError('Hndles not close, but should: %s'% notClosedHndles)

    def connect(self):
        # start with 1 for thread safety when run from pythonwin:
        # natlink.natConnect(natconnectOption)
        # sys.stdout = open(logFileName, 'a')
        # sys.stderr = open(logFileName, 'a')
        self.isConnected = False
        pass
    
    def disconnect(self):
        # natlink.natDisconnect()
        self.isConnected = False
        pass

    def getTextViaClipboard(self, waitTime):
        """get text of window and selection or cursor via copy paste actions
        
        pass waitTime for testing response time.
        
        waiting_iterations is default 10, but changed to 3 if an empty clipboard is expected
        
        return: text, selStart, selEnd, cursorPos
               
        """
        cb = natlinkclipboard.Clipboard(save_clear=True, debug=1)
        waitTime = 0.001
      
        # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        # natlinkutilsqh.SetForegroundWindow(self.testHndle)
        
        ## first try if there is a selection:
        action("{ctrl+c}")
        selection = cb.get_text(waitTime, waiting_iterations=3)
        if selection:
            print('selection: %s'% selection)
            action("{left}")
            cb.clear_clipboard()

           
        # now try text left of selection:
        action("{ctrl+shift+home}{ctrl+c}")
     
        # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        # natlinkutilsqh.SetForegroundWindow(self.testHndle)
     
     
        before = cb.get_text(waitTime)
        if before:
            action("{right}")
            cb.clear_clipboard()
        else:
            print("at start of document")

        # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        # natlinkutilsqh.SetForegroundWindow(self.testHndle)

        text = before + selection
        selStart = len(before)
        selEnd = len(text)
    
        # skip to end of selection:       
        if selection:
            action("{right %s}"% len(selection))

        # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        # natlinkutilsqh.SetForegroundWindow(self.testHndle)
        
        if text:
            # now select to end:
            action("{shift+ctrl+end}{ctrl+c}")
            after = cb.get_text(waitTime)
            print("after: %s"% repr(after))
            
            if not after:
                if self.testHndle == self.thundHndle:
                    ## after an empty selection, go one to the left!!
                    action("{left}")
                    pass
                after = ""
            else:
                # undo selection
                action("{left}")
            if after.endswith('\0'):
                after = after.replace('\0', '')
                print("removed null char at end of selection (1): %s"% after)
                # after = afterplusone[1:]
        else:
            # presumably at start of buffer
            action("{shift+ctrl+end}{ctrl+c}")
            after = cb.get_text(waitTime, waiting_iterations=3)
            if after:
                # undo selection 
                action("{left}")
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
            
            # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
            # natlinkutilsqh.SetForegroundWindow(self.testHndle)
            
            action("{left %s}{shift+right %s}"% (lensel, lensel))

        text = text + after
        cursorPos = selEnd            

        # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        # natlinkutilsqh.SetForegroundWindow(self.testHndle)


        return text, selStart, selEnd, cursorPos
        

        

    def setupWindows(self):
        """make several windows to which can be switched, and which can be copied from
        
        TODOQH: make bringup of a folder and copy paste of the folder info work... (if somebody wants this)
        """
        return
        self.thisHndle = natlinkutilsqh.GetForegroundWindow()

        dirWindows = "C:\\windows"
        result = actions.AutoHotkeyBringUp(app="explore", filepath=dirWindows)
        if not result:
            print('no result for %s'% dirWindows)
            return
        pPath, wTitle, hndle = result
        self.tempFileHndles.append(hndle)
        self.expl0Hndle = hndle
        natlinkutilsqh.SetForegroundWindow(self.thisHndle)

    def setupDocxFile(self):
        """open Word file and do the testing on it...
        """
        docxFile2 = "natlink.docx"
        docxPath2 = os.path.normpath( os.path.join(testFilesDir, docxFile2))
        if not os.path.isfile(docxPath2):
            raise IOError('file does not exist: %s'% docxPath2)
        result = actions.AutoHotkeyBringUp(app=None, filepath=docxPath2)
        if result:
            pPath, wTitle, hndle = result
        else:
            raise TestError("AutoHotkeyBringUp of Word gives no result: %s"% result)
        self.tempFileHndles.append(hndle)
        natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        print('testempty.docx (docx2Hndle): %s'% result)
        self.tempFileHndles.append(result)
        self.docx2Hndle = result        
        
    def setupTextFiles(self):
        """make some text files for clipboard testing
        """
        textFile0 = "testempty.txt"
        textPath0 = os.path.join(testFilesDir, textFile0)
        open(textPath0, 'w')
        result = actions.AutoHotkeyBringUp(app=None, filepath=textPath0)
        pPath, wTitle, hndle = result
        natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        print('testempty (text0Hndle): %s'% hndle)
        self.tempFileHndles.append(hndle)
        self.text0Hndle = hndle

        textFile1 = "testsmall.txt"
        textPath1 = os.path.join(testFilesDir, textFile1)
        self.text1Txt = "small abacadabra\n"*2
        open(textPath1, 'w').write(self.text1Txt)
        result = actions.AutoHotkeyBringUp(app=None, filepath=textPath1)
        pPath, wTitle, hndle = result
        self.tempFileHndles.append(hndle)
        natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        print('testsmall (text1Hndle): %s'% hndle)
        self.text1Hndle = hndle
        self.tempFileHndles.append(hndle)

        textFile2 = "testlarge.txt"
        textPath2 = os.path.join(testFilesDir, textFile2)
        self.text2Txt = "large abacadabra\n"*1000
        open(textPath2, 'w').write(self.text2Txt)
        result = actions.AutoHotkeyBringUp(app=None, filepath=textPath2)
        pPath, wTitle, hndle = result
        self.tempFileHndles.append(hndle)
        natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        print('testlarge (text2Hndle): %s'% hndle)
        self.text2Hndle = hndle
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

        result = actions.AutoHotkeyBringUp(app="thunderbird.exe", title="Mozilla Thunderbird", extra=extra)
        pPath, wTitle, hndle = result
        self.tempFileHndles.append(hndle)
        self.thundHndle = hndle
        self.killActions[hndle] = "KW({ctrl+w}, w)"   # kill the window, with letter w if windows has text....
        self.modInfos[hndle] = result
        action("Thunderbird new window abacadabra")

    def setupFrescobaldiNewPane(self):
        """start an empty pane in Frescobaldi (front end for Lilypond music type setter)
        
        at shutdown do not kill this window
        """
        ## this extra makes an empty window, and goes to the message pane:
        extra = '\n'.join(["Sleep, 100", 'Send, ^n', "Sleep, 100"])

        result = actions.AutoHotkeyBringUp(app=r"C:\Program Files (x86)\Frescobaldi\frescobaldi.exe", title="Frescobaldi", extra=extra)
        pPath, wTitle, hndle = result
        action("Frescobaldi abacadabra")
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
        natlinkutilsqh.SetForegroundWindow(self.thisHndle)
        # print("Frescobaldi handle: ", self.frescHndle)
        # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
                
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
            action("{ctrl+a}{del}")
            
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime=waitTime)
            expText = ""
            expPosTuple = (0,0,0)  # startSel, endSel, cursorPos
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")
            
            action("Hello world")

            # these pair of lines provide the possibility to debug the tests
            # put the breakpoint at the second line...
            # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
            # natlinkutilsqh.SetForegroundWindow(self.testHndle)
    
    
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime)
            expText = "Hello world"
            expPosTuple = (11, 11, 11)  # startSel, endSel, cursorPos
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")
            pass

    
            # now select word world
            action("{shift+ctrl+left}")
    
            # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
            # natlinkutilsqh.SetForegroundWindow(self.testHndle)

    
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime)
            expText = "Hello world"
            expPosTuple = (6, 11, 11)  # startSel, endSel, cursorPos
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            ## and again:
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime)
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")
    
            # natlinkutilsqh.SetForegroundWindow(self.thisHndle)
            # natlinkutilsqh.SetForegroundWindow(self.testHndle)
    
            # now replace the word world, and make a second line
            action('SCLIP(WORLD{enter}How are you going?{enter}In this Corona world?)')
            expText = "Hello WORLD\nHow are you going?\nIn this Corona world?"
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime)
            expPosTuple = (52, 52, 52)  # startSel, endSel, cursorPos
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")
            pass

            ## and again:    
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime)
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")
    
            # select all
            action("{ctrl+home}{shift+ctrl+end}")
            expText = "Hello WORLD\nHow are you going?\nIn this Corona world?"
            # if self.testHndle == self.thundHndle:
            #     expText += "\n"
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime)
            expPosTuple = (0, len(expText), len(expText))  # startSel, endSel, cursorPos
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            # and again:
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime)
            expPosTuple = (0, len(expText), len(expText))  # startSel, endSel, cursorPos
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")

            # select "are you"
            action("{ctrl+home}{down}{right 4}{shift+right 7}")
            expText = "Hello WORLD\nHow are you going?\nIn this Corona world?"
            # not here!!
            # if self.testHndle == self.thundHndle:
            #     expText += "\n"
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime)
            expPosTuple = (16, 23, 23) # startSel, endSel, cursorPos
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")
    
            ## and again:
            text, startSel, endSel, cursorPos = self.getTextViaClipboard(waitTime)
            expPosTuple = (16, 23, 23) # startSel, endSel, cursorPos
            self.assert_equal(expText, text, "text not as expected")
            self.assert_equal(expPosTuple, (startSel, endSel, cursorPos), "positions not as expected")


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

        cb = natlinkclipboard.Clipboard(save_clear=True, debug=2)   # debug flag can be set or omitted (or set to 0)
        return

        ## longer test:
        for i in range(1, 10, 5):
            waitTime = 0.001/i
            print('---- round: %s, waitTime %.4f '% (i, waitTime))
            
            ## first clear the clipboard:
            cb.clear_clipboard()
            got = cb.get_text(waitTime)
            print('after clear_clipboard, text: %s (cb: %s)'% (got, cb))
            self.assert_equal("", got, "should have no text now, waitTime: %s (times 4)"% waitTime)

            print("now testempty.txt -----------------")
            # empty file: 
            natlinkutilsqh.SetForegroundWindow(self.text0Hndle)
            natlinkutils.playString("{ctrl+a}{ctrl+c}")
            got = cb.get_text(waitTime)
            print('after select all copy of empty file, no change should have happened:')
            # print('got: %s, (cb: %s)'% (got, cb))
            self.assert_equal("", got, "should have no text now, waitTime: %s"% waitTime)
            
            natlinkutilsqh.SetForegroundWindow(self.thisHndle)

            print("now testsmall.txt -----------------")
            natlinkutilsqh.SetForegroundWindow(self.text1Hndle)
            natlinkutils.playString("{ctrl+a}{ctrl+c}")
            got = cb.get_text(waitTime)
            print('after select all copy of testsmall.txt')
            # print('got: %s, (cb: %s)'% (got, cb))
            exp = self.text1Txt
            self.assert_equal(exp, got, "testsmall should have two lines of text, waitTime: %s"% waitTime)
            
            
            # test large.txt
            print("now testlarge.txt -----------------")
            natlinkutilsqh.SetForegroundWindow(self.text2Hndle)
            natlinkutils.playString("{ctrl+a}{ctrl+c}")
            cb.get_text(waiting_interval=waitTime)
            got = cb.get_text()
            if got:
                lengot = len(got)
            else:
                print('testlarge, no result from get_text: %s'% cb)
            exp = len(self.text2Txt)
            self.assert_equal(exp, lengot, "should have long text now, waitTime: %s"% waitTime)
            # empty for the next round:
            # cb.set_text("")
            time.sleep(waitTime)

            natlinkutilsqh.SetForegroundWindow(self.thisHndle)

            # test natlink.docx
            print("now natlink.docx (%s)-----------------"% self.docx2Hndle)
            if not self.docx2Hndle:
                print("word document not available: %s"% (self.docx2Hndle))
                continue
            natlinkutilsqh.SetForegroundWindow(self.docx2Hndle)
            natlinkutils.playString("{ctrl+a}{ctrl+c}")
            cb.get_text(waiting_interval=waitTime)
            got = cb.get_text()
            if got:
                lengot = len(got)
            else:
                print('testlarge, no result from get_text: %s'% cb)
            exp = len(self.text2Txt)
            self.assert_equal(exp, lengot, "should have long text now, waitTime: %s"% waitTime)
            # empty for the next round:
            # cb.set_text("")
            time.sleep(waitTime)
            
        del cb
        got_org_text = natlinkclipboard.Clipboard.get_system_text()
        self.assert_equal(self.org_text, got_org_text, "restored text from clipboard not as expected")

    
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
        
        thisHndle = natlinkutilsqh.GetForegroundWindow()
        print("thisHndle: %s"% thisHndle)
        unknownHndle = 5678659
        hndles = [thisHndle, self.docx2Hndle, self.docx1Hndle, self.text0Hndle, self.text1Hndle, self.text2Hndle,
                  self.frescHndle, self.thundHndle, unknownHndle]
        t0 = time.time()
        rounds = 10
        for i in range(1, rounds+1):
            print('start round %s'% i)
            for h in hndles:
                if not h: continue
                result = natlinkutilsqh.SetForegroundWindow(h)
                if result:
                    self.assert_not_equal(unknownHndle, h, "hndle should not be unknown (fake) hndle")
                else:
                    self.assert_equal(unknownHndle, h, "hndle should be one of the valid hndles")

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
        if self.isConnected:
            natlink.displayText(t, 0)
        else:
            print(t)
        print(t, file=open(logFileName, "a"))

def run():
    print('starting UnittestClipboard') 
    # trick: if you only want one or two tests to perform, change
    # the test names to her example def test....
    # and change the word 'test' into 'tttest'...
    # do not forget to change back and do all the tests when you are done.
    # each test does a natConnect and a natDisconnect...
    sys.stdout = open(logFileName, 'a')
    sys.stderr = open(logFileName, 'a')
    
    suite = unittest.makeSuite(UnittestClipboard, 'test')
    result = unittest.TextTestRunner().run(suite)
    
if __name__ == "__main__":
    print("run, result will be in %s"% logFileName)
    run()
