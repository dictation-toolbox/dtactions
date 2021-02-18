#
# Python Macro Language for Dragon NaturallySpeaking
#   (c) Copyright 1999 by Joel Gould
#   Portions (c) Copyright 1999 by Dragon Systems, Inc.
#
# unittestMessagefunction.py
# this script tests the message functions, needed for communication with
# a (non active) window.
# Developed by Quintijn Hoogenboom, for Jason Koller, application for global
# dictation caption and putting the results into a fixed target window.
#
# below fill in parameters for finding correct window:
#  tester: Name of the test you are performing...
#  application (testApp)
#  window title (testCaption) (case insensitive, may be part of, or None)
#  className of app (testClass) (case sensitive)
#  W["testcloseapp"] (1: close at each test, 0: leave open after each test)
#      (text after each line, '' in Notepad,
#                                      '\r' in WordPad,
#                                      '\r\n' in aligen ????)
#               (now ignore, getting only the text without \r and \n characters)
#  linesep is important for synchronising the select and say functions.
#                 '\n' Notepad,
#                 '\r\n' for WordPad, aligen (???)

# increase or decrease for more visible or faster testing:
visibleTime = 0.5 # set higher, if you want to see what happens, 1, 2...

import sys
import os
import os.path
import win32clipboard
from pathqh import path

#need this here (hard coded, sorry) for it can be run without NatSpeak being on
extraPaths = [path(r"C:\natlinkGIT3\unimacro")]
for extraPath in extraPaths:
    extraPath.isdir()
    extraNorm = extraPath.normpath()
    if extraNorm not in sys.path:
        sys.path.append(extraNorm)
# little trick to keep testers apart (do not bother)
try:
    import testermod
    tester = testermod.tester
except ImportError:
    tester = "aligen"

import windowparameters

W = windowparameters.PROGS[tester]

import sys
import unittest
import os
import os.path
import time
import traceback        # for printing exceptions
from struct import pack
import natlink
import win32gui
import win32ui
from messagefunctions import *
import messagefunctions as mess
import TestCaseWithHelpers
class TestError(Exception):pass

#---------------------------------------------------------------------------
class UnittestMessagefunctions(TestCaseWithHelpers.TestCaseWithHelpers):
    def setUp(self):
        self.__class__.ctrl = None
        self.__class__.app = None
        for k,v in list(W.items()):
            setattr(self, k, v)

        self.getEditControl()
        if self.ctrl is None:
            raise TestError("could not get edit handle")
        #content = mess.getEditText(self.ctrl)
        #if len(content) > 1:
        #    print 'content of window:\n|%s|'% repr(content)
    
    def tearDown(self):
        time.sleep(visibleTime)
        if W["testcloseapp"]:
            self.clearAppWindow()
        else:
            if self.app and self.ctrl:
                setEditText(self.ctrl, "", classname=self.classname)
            time.sleep(0.2)



    #def assert_selection_through_clipboard_contents(self, expected, text=""):
    #    """tests the contents of the clipboard"""
    #    text = text or "clipboard contents not as expected"
    #    meHndle = win32gui.GetForegroundWindow()
    #    natqh.SetForegroundWindow(self.app)
    #    time.sleep(0.05)
    #    activateMenuItem(self.app, W["commandcut"])
    #    time.sleep(0.05)
    #    t = getClipboard()
    #    natqh.SetForegroundWindow(meHndle)
    #    
    #    text = text + '\nExpected: %s\nGot: %s\n'% (expected, t)
    #    self.assert_(t == expected, text)


    def wait(self, t=None):
        time.sleep(t or 0.1)

    #def getAppWindow(self):
    #    appWindows = findTopWindows(wantedClass=W["windowclass"], wantedText=W["windowcaption"])
    #    if len(appWindows) > 1:
    #        print 'warning, more appWindows active! %s'% appWindows
    #        self.__class__.app = appWindows[0]
    #    elif appWindows:
    #        self.__class__.app = appWindows[0]
    #    else:
    #        if not W["shouldstartauto"]:
    #            raise TestError("application (class: %s, caption: %s) should be active before you start the tests"%
    #                            (W["windowclass"], W["windowcaption"]))
    #        meHndle = win32gui.GetForegroundWindow()
    #        result = os.startfile(W["apppath"])
    #        if W["testcloseapp"]:
    #            sleepTime = 0.1
    #        else:
    #            sleepTime = 0.5
    #        for iTry in range(20):
    #            time.sleep(sleepTime)
    #            appWindows = findTopWindows(wantedClass=W["windowclass"], wantedText=W["windowcaption"])
    #            if appWindows: break
    #        else:
    #            print 'starting %s failed, dumping apps'% testApp 
    #            pprint.pprint(dumpTopWindows())
    #            return
    #
    #        if meHndle:
    #            try:
    #                natqh.SetForegroundWindow(meHndle)
    #            except:
    #                pass
    #        self.__class__.app = appWindows[0]
    def getAppWindow(self):
        appWindow = findTopWindow(wantedClass=W["windowclass"], wantedText=W["windowcaption"])
        if appWindow:
            isVisible = win32gui.IsWindowVisible(appWindow)
            #print 'is visible? %s'% isVisible
            self.__class__.app = appWindow
            return
        if not W["shouldstartauto"]:
            raise TestError("application (class: %s, caption: %s) should be active before you start the tests"%
                            (W["windowclass"], W["windowcaption"]))
        meHndle = win32gui.GetForegroundWindow()
        result = os.startfile(W["apppath"])
        if W["testcloseapp"]:
            sleepTime = 0.1
        else:
            sleepTime = 0.5
        for iTry in range(20):
            time.sleep(sleepTime)
            appWindow = findTopWindow(wantedClass=W["windowclass"], wantedText=W["windowcaption"])
            if appWindow: break
        else:
            print('starting %s failed, dumping apps'% testApp) 
            pprint.pprint(dumpTopWindows())
            return

        if meHndle:
            try:
                natqh.SetForegroundWindow(meHndle)
            except:
                pass
        self.__class__.app = appWindow
            
    def getEditControl(self):
        # get application, if not found, set ctrl to None
        self.getAppWindow()
        if not self.app:
            self.__class__.ctrl = None
            return
        if tester == 'pythonwin':
            wTitle = win32gui.GetWindowText(self.app)
            filename = wTitle.split('[')[-1][:-1]  # remove [ and ], leaving only the filenam
            def selectionFunction(hndle, gotTitle, gotClass):
                """special for selecting the Afx class with the same title as the complete window title bar
                being the filename in question
                
                special for pythonwin and only for the FIRST search action for child windows.
                After the correct Afx window has been identified, the Scintilla child window is the correct one.
                """
                if gotTitle == filename:
                    #print 'got afx with title: %s'% gotTitle
                    return True
        else:
            selectionFunction = None
        currentSel = None
        if self.ctrl:
            currentSel = mess.getSelection(self.ctrl)
        if currentSel and currentSel != (0,0):
            return
        wC, wT = W["editcontrol"], W["edittext"]
        choiceControl = 0
        if type(wT) in [type(None), bytes, int]: wT =  [wT]
        if type(wC) in [type(None), bytes, int]: wC =  [wC]
        ctrl = self.app            
        for wtext, wclass in zip(wT, wC):
            ctrls =  mess.findControls(ctrl,wantedText=wtext, wantedClass=wclass,selectionFunction=selectionFunction)
            if selectionFunction:
                selectionFunction = None # execute only for first findControls action pythonwin
                if tester == 'pythonwin':
                    choiceControl = -1

            if len(ctrls):
                ctrl = ctrls[choiceControl]
                #print 'editHndle set to: %s'% editHndle
                if len(ctrls) > 1:
                    for hndle in ctrls:
                        id = win32gui.GetDlgCtrlID(hndle)
                        if id == 0xd6:
                            ctrl = hndle
                            break
            else:
                pprint.pprint(mess.dumpWindow(self.app))
                raise ValueError("could not find the editHndle of the control: %s in application: %s"%
                              (self.editcontrol, self.apppath))
                
        self.__class__.ctrl = ctrl
        self.__class__.classname = win32gui.GetClassName(ctrl)

    def clearAppWindow(self):
        if self.app:
            time.sleep(visibleTime)
            setEditText(self.ctrl, '', classname=self.classname)
            quitProgram(self.app)
            self.app = None
            time.sleep(0.1)

    def dumpWindows(self):
        """dump non-standard window names"""
        dumpTopWindows(all=None)
    #---------------------------------------------------------------------------
    def test_get_controls_of_application(self):
        self.assertTrue(self.app > 0, "app should be there now")
        if self.ctrl:
            print('edit area: %s'% self.ctrl)
        else:
            print('all controls of app %s------'% testApp)
            print('all controls of application:-----------------')
            pprint.pprint(dumpWindow(self.app))
        
    def test_existence_of_application(self):
        self.assertTrue(self.app > 0, "app (handle to applicatione) should be there now")
        #print 'all top windows:------------'
        #pprint.pprint(dumpTopWindows())
        # try closing and then through exception again:
        # only if you choose to close app after each test:
        # this raises an error in tearDown, because the app is already closed then...
        # maybe sort out later.
        #if W["testcloseapp"]:
        #    quitProgram(self.app)
        #pass
        self.assertTrue(self.ctrl > 0, "ctrl (handle to edit control) should be there now")
        print('content of window at start:\n|%s|'% mess.getEditText(self.ctrl))
        
    def test_set_text_in_application(self):
        """setting and getting text in application
        
        setting can be done as a string or as a list of strings.
        In the latter case newlines are inserted.
        """
        sep = self.linesep
        self.assertTrue(self.app > 0, "app should be there now")
        gotText = getEditText(self.ctrl)
        self.assert_equal([sep], gotText, 'text of edit box should be empty at start, adjust W["linesep"] or W["aftertext"] in windowparameters!')

        # setting "hello there"
        setText = "hello there"
        setEditText(self.ctrl, setText)
        gotText = getEditText(self.ctrl)
        expected = [setText+sep]
        # so aftertext is pasted after the text!
        self.assert_equal(expected, gotText, "text of edit box should be the same as the set text")

        # settting "in Amsterdam"
        setText = "in Amsterdam"
        setEditText(self.ctrl, setText)
        gotText = getEditText(self.ctrl)
        expected = [setText+sep]
        self.assert_equal(expected, gotText, "text of edit box should be the same as the set text")
        
        # appending ", hello there."
        setText = ", hello there."
        setEditText(self.ctrl, setText,append=True)
        gotText = getEditText(self.ctrl)
        expected = ["in Amsterdam, hello there."+sep]
        self.assert_equal(expected, gotText, "text of edit box should be the same as the set text")
        
        # appending a new line
        setText = ["", "", "How are you doing?", ""]
        setEditText(self.ctrl, setText,append=True)
        gotText = getEditText(self.ctrl)
        expected = ['in Amsterdam, hello there.'+sep, sep,
                    'How are you doing?'+sep, sep]
        self.assert_equal(expected, gotText, "text of edit box should be the same as the set text")

        # clear all and append as one string with \r
        self.clearAppWindow()
        sep2 = sep + sep
        setText = sep2.join(["first", "third line"])
        setEditText(self.ctrl, setText)
        gotText = getEditText(self.ctrl)
        expected = ["first"+sep, sep, "third line"+sep]
        self.assert_equal(expected, gotText, "text of edit box should be the same as the set text (sent as one string")
        
    def test_selection_set_text_in_application(self):
        """setting and getting text in application in a selected range
        
        """
        sep = self.linesep
        self.assertTrue(self.app > 0, "app should be there now")
        # setting "hello there"
        setText = "hello there"
        setEditText(self.ctrl, setText)
        gotText = getEditText(self.ctrl)
        expected = [setText+sep]
        self.assert_equal(expected, gotText, "text of edit box should contain initial text now")
        # get selection (at end)
        
        # set selection:
        lo, hi = getSelection(self.ctrl)
        expected = len(setText), len(setText)
        self.assert_equal( expected, (lo,hi), "selection returned differs from trying to set")
        
        setSelection(self.ctrl, 1,4)  # "ell" 
        lo, hi = getSelection(self.ctrl)
        self.assert_equal( (1,4), (lo,hi), "selection returned differs from trying to set")

        # append a second paragraph:
        setText = ["ELLLLLL"]
        replaceEditText(self.ctrl, setText)
        gotText = getEditText(self.ctrl)
        
        expected = ['hELLLLLLo there'+sep]
        self.assert_equal(expected, gotText, "text of edit box after insert in selection is wrong")

        lo, hi = getSelection(self.ctrl)
        expectedRange = (8,8)  #  1 + 7,  start of previous + length of insertion
        self.assert_equal(expectedRange, (lo, hi), "selection after insertion not as expected")
        #clearClipboard()
        activateMenuItem(self.app, W["commandselectall"])
        lo, hi = getSelection(self.ctrl)
        expectedSel = (0, len(expected[0]))
        self.assert_equal(expectedSel, (lo, hi), "selection after select all not as expected")


    def test_scrolling_text_in_application(self):
        """see what happens if text scrolls in a control"""
        sep = self.linesep
        self.assertTrue(self.app > 0, "app should be there now")
        # setting text:
        longerText = "This text is longer than 80 characters, so should scroll in most controls. I therefore like to see what happens with this text."
        secondPara = "This goes on the next line. Is also considerable of length, so we can find the differences between scrolled lines and real new lines."
        setEditText(self.ctrl, [longerText, secondPara] )
        gotText = getEditText(self.ctrl)
        self.assertTrue(len(gotText) > 1, "text should scroll, make window more narrow")
        
        expected = ''.join([longerText+sep, secondPara+sep])
        got = ''.join(gotText)
        self.assert_equal(expected, got, "text of edit box should contain initial text now")
        # get selection (at end)

    def test_getting_more_info_from_application(self):
        """try to get text length and other things
        
        this is quite a messy test, used also for timing purposes.
        Need a U: (virtual) drive for testing (with hotspot) and use the small print_hotspot program to
        print the results
        """
        sep = self.linesep
        aft = self.aftertext
        self.assertTrue(self.app > 0, "app should be there now")
        # setting text:
        longerText = "This text is longer than 80 characters, so should scroll in most controls. I therefore like to see what happens with this text. I hope to see the scrolling taking place, without updating the line number. I type on and on in the hope that even a very wide control cannot handle this amount of text in one line."
                     
        total = []
        for i in range(100):
            total.append("This is a sligthly longer line %s."% (i+1,))
        total.append(longerText)
        setEditText(self.ctrl, sep.join(total) )
        gotText = getEditText(self.ctrl)
        self.assertTrue(len(gotText) > 1, "text should scroll, make window more narrow")
        
        expected = sep.join(total) + aft
        got = ''.join(gotText)
        self.assert_equal(expected, got, "text of edit box should contain initial text now")
        
        numLines = getNumberOfLines(self.ctrl)
        self.assert_equal(101, numLines, "number of lines in control does not match")
        # line number of call is one less than on the screen (0 based, while screen is 1 based)
        line = getTextLine(self.ctrl, 1)
        expected = 'This is a sligthly longer line 2.\r'
        self.assert_equal(expected, line, "text of line 2 (internal 1) does not match expected")

        line = getTextLine(self.ctrl, 20)
        expected = 'This is a sligthly longer line 21.\r'
        self.assert_equal(expected, line, "text of line 21 (internal 20) does not match expected")


        line = getTextLine(self.ctrl, 0)
        expected = 'This is a sligthly longer line 1.\r'
        self.assert_equal(expected, line, "text of line 1 (internal 0) does not match expected")


        line = getTextLine(self.ctrl, 123456)
        expected = ''
        self.assert_equal(expected, line, "text of line VERY_LARGE_NUMBER does not match expected")

        lastLine = getTextLine(self.ctrl, numLines-1)
        expected = 'like to see what happens with this text.\r'
        self.assert_equal(expected, lastLine, "text of last line does not match expected")

        setSelection(self.ctrl, 10, 50)
        ln = getLineNumber(self.ctrl) # of cursor position
        setSelection(self.ctrl, 500, 500)
        ln = getLineNumber(self.ctrl)
        setSelection(self.ctrl, 5000, 5001)
        ln = getLineNumber(self.ctrl)
        setSelection(self.ctrl, 5000, 4000)
        ln = getLineNumber(self.ctrl)
        setSelection(self.ctrl, 5000, 400000)
        ln = getLineNumber(self.ctrl)


        #getStartEndOfLine(self.ctrl, ln)
        info = getNumberOfLines(self.ctrl)
        print('info: %s'% repr(info))
        for i in range(1):
            rawLength = getRawTextLength(self.ctrl)
            gotTextAll, bufLen = getWindowTextAll(self.ctrl, rawLength=rawLength)
        print('rawLength: %s'% rawLength)
        # get selection (at end)
        gotText = ''.join(getEditText(self.ctrl))
        print('length of buffer: %s'% len(gotText))
        #print 'buffer (old way): %s'% repr(gotText)


        gotTextAll, bufLen = getWindowTextAll(self.ctrl)
        setEditText(self.ctrl, gotTextAll)
        gotTextAll2, bufLen2 = getWindowTextAll(self.ctrl)
        print('second All== firstAll? %s'% (gotTextAll == gotTextAll2))
        print('gotText (%s): %s'% (len(gotText), repr(gotText[:100])))
        print('gotText2 (%s): %s'% (len(gotTextAll2), repr(gotTextAll2[:100])))
        setEditText(self.ctrl, gotTextAll2)
        gotText2 = ''.join(getEditText(self.ctrl))
        setEditText(self.ctrl, gotText2)
        gotText3 = ''.join(getEditText(self.ctrl))
        print('gotText2 (%s): %s'% (len(gotTextAll2), repr(gotTextAll2[:100])))
        print('Org: 1 == 2? %s'% (gotText == gotText2))
        print('Org: 2 == 3? %s'% (gotText3 == gotText2))

        #for i in range(5000):
        #    total.append("This is another still sligthly longer line %s."% (i+1,))
        #setEditText(self.ctrl, sep.join(total) )
        #gotTextAll, bufLen = getWindowTextAll(self.ctrl)
        #print 'buflen after more lines: %s'% bufLen
        
        
    def test_selection_set_text_in_application(self):
        """setting and getting text in application in a selected range
        
        """
        sep = self.linesep
        self.assertTrue(self.app > 0, "app should be there now")
        # setting "hello there"
        setText = "hello there"
        setEditText(self.ctrl, setText, classname=self.classname)
        gotText = getEditText(self.ctrl, self.classname)
        expected = [setText+sep]
        self.assert_equal(expected, gotText, "text of edit box should contain initial text now")
        
        # get selection (at end)
        
        # set selection:
        lo, hi = getSelection(self.ctrl, classname=self.classname)
        expected = len(setText), len(setText)
        self.assert_equal( expected, (lo,hi), "selection returned differs from trying to set")
        
        setSelection(self.ctrl, 1,4, classname=self.classname)  # "ell" 
        lo, hi = getSelection(self.ctrl)
        self.assert_equal( (1,4), (lo,hi), "selection returned differs from trying to set")

        # append a second paragraph:
        setText = ["ELLLLLL"]
        replaceEditText(self.ctrl, setText, classname=self.classname)
        gotText = getEditText(self.ctrl)
        
        expected = ['hELLLLLLo there'+sep]
        self.assert_equal(expected, gotText, "text of edit box after insert in selection is wrong")

        lo, hi = getSelection(self.ctrl, classname=self.classname)
        expected = (8,8)  #  1 + 7,  start of previous + length of insertion
      
        self.assert_equal(expected, (lo, hi), "selection after insertion not as expected")

        setText = [" at the bottom"]
        appendEditText(self.ctrl, setText, classname=self.classname)
        gotText = getEditText(self.ctrl)
        
        expected = ['hELLLLLLo there at the bottom'+sep]
        self.assert_equal(expected, gotText, "text of edit box after insert in selection is wrong")

        lo, hi = getSelection(self.ctrl, classname=self.classname)
        expected = (29, 29)  
        self.assert_equal(expected, (lo, hi), "selection after insertion not as expected")

        setText = [" longer text, because there should be a scroll now, depending on the settings of wordpad, notepad or aligen."]
        appendEditText(self.ctrl, setText, classname=self.classname)
        gotText = getEditText(self.ctrl)
        expected = ['hELLLLLLo there at the bottom'+sep]
        lo, hi = getSelection(self.ctrl, classname=self.classname)
        expectedSelection = (137, 137)  
        self.assert_equal(expectedSelection, (lo, hi), "selection after insertion not as expected")
        expected = ['hELLLLLLo there at the bottom longer ', 'text, because there should be a scroll ', 'now, depending on the settings of ', 'wordpad, notepad or aligen.']
        self.assert_equal(''.join(expected)+sep, ''.join(gotText), "longer text of edit box after insert in selection is wrong (Normal for notepad in different sizes)")


    def test_visible_line_in_application(self):
        """getting the first visible line in an application
        
        """
        sep = self.linesep
        
        self.assertTrue(self.app > 0, "app should be there now")
        # setting "hello there"
        L = []
        for i in range(100):
            L.append(str(i) + " " + "A"*50 + " " + "B"*50)
            
        setEditText(self.ctrl, L)
        lineNo = getFirstVisibleLine(self.ctrl)
        expectedMin = 100
        self.assertTrue(lineNo > expectedMin, "lineNo should be at least %s, is now: %s"%
                     (expectedMin, lineNo))
        print('first visible line in app: %s'% lineNo)
        gotText = getEditText(self.ctrl, visible=True)
        pass

    def test_menu_functions(self):
        """test select all, copy and cut functions
        
        """
        sep = self.linesep
        editcontrol = self.editcontrol
        
        self.assertTrue(self.app > 0, "app should be there now")
        # first selectall with empty window:
        gotText = getEditText(self.ctrl)
        expected = [sep]
        self.assert_equal(expected, gotText, "empty edit box should contain only aftertext now")
        time.sleep(0.1)
        lo, hi = getSelection(self.ctrl)
        expSel = (0, 0)  # or (0, len(sep)) now it seems the selection is WITHOUT the trailing \r or \r\n characters
        self.assertTrue(expSel == (lo,hi), "selection initially empty screen not as expected\nExpected: %s\nActual: %s"%
                     (repr(expSel), repr((lo, hi))))
        #clearClipboard()
        activateMenuItem(self.app, W["commandselectall"])
        time.sleep(0.1)
        lo, hi = getSelection(self.ctrl)
        if sep == '\r\n':
            expSel = (0,  len(sep))  # aligen
        else:                        # differs from
            expSel = (0, 0)          # wordpad
        self.assertTrue(expSel == (lo,hi), "selection empty screen after command Select All not as expected\nExpected: %s\nActual: %s"%
                     (repr(expSel), repr((lo, hi))))
        
        #now put some text in and check:
        setText = "text for selection"
        setEditText(self.ctrl, setText)
        gotText = getEditText(self.ctrl)
        expected = [setText+sep]
        self.assert_equal(expected, gotText, "text of edit box should contain initial text now")
        #clearClipboard()
        activateMenuItem(self.app, W["commandselectall"])
        time.sleep(0.1)
        lo, hi = getSelection(self.ctrl)
        # for one line the selection is WITHOUT the trailing \r character!! (wordpad)
        if sep == '\r\n':
            expSel = (0,len(setText+sep))   # again aligen
        elif editcontrol == "RichEdit20A":
            # DragonPad, new aligen
            expSel = (0,len(setText+sep))     # DragonPad  (sep is \r)
        else:
            expSel = (0,len(setText))       # wordpad
        self.assertTrue(expSel == (lo,hi), "selection after select all not as expected, check W['commandselectall']\nExpected: %s\nActual: %s"%
                     (repr(expSel), repr((lo, hi))))

        setText = ["", "more", "", "", "lines", "", "", "of", "", "", "text.", ""]
        setEditText(self.ctrl, setText)
        gotText = getEditText(self.ctrl)
        print('gotText: %s'% repr(gotText))
        expected = [s+sep for s in setText]
        self.assert_equal(expected, gotText, "text of edit box should contain more lines of text now")
        #clearClipboard()
        activateMenuItem(self.app, W["commandselectall"])
        time.sleep(0.1)
        lo, hi = getSelection(self.ctrl)
        hiExp= sum(len(s) for s in expected) 
        expSel = (0,hiExp)
        self.assertTrue(expSel == (lo,hi), "selection after select all not as expected (trailing empty line), check W['commandselectall']\nExpected: %s\nActual: %s"%
                     (repr(expSel), repr((lo, hi))))

        setText = ["", "more", "", "", "lines"]
        setEditText(self.ctrl, setText)
        gotText = getEditText(self.ctrl)
        expected = [s+sep for s in setText]
        self.assert_equal(expected, gotText, "text of edit box (no trailing empty line) should contain initial text now")
        #clearClipboard()
        activateMenuItem(self.app, W["commandselectall"])
        time.sleep(0.1)
        lo, hi = getSelection(self.ctrl)
        hiExp= sum(len(s) for s in expected) 
        # now the selection is INCLUDING the trailing \r character!!
        expSel = (0,hiExp)
        self.assertTrue(expSel == (lo,hi), "selection after select all not as expected (no trailing em, check W['commandselectall']\nExpected: %s\nActual: %s"%
                     (repr(expSel), repr((lo, hi))))

        


    def test_sending_keystrokes(self):
        sep = self.linesep
        #sendKey(self.ctrl, sep) # fails for aligen (\r\n)
        sendKey(self.ctrl, "aA")
        time.sleep(0.1)
        sendKey(self.ctrl, "{left}")
        time.sleep(0.1)
        sendKey(self.ctrl, "zZ")
        time.sleep(0.1)
        gotText = getEditText(self.ctrl)
        expected = ["azZA" + sep]
        self.assertTrue(expected == gotText, "sendking keystrokes give unexpected tesult\nExpected: %s\nActual: %s"%
                     (repr(expected), repr(gotText)))

        # trying delete key:
        sendKey(self.ctrl, "{delete}")
        gotText = getEditText(self.ctrl)
        expected = ["azZ" + sep]
        self.assertTrue(expected == gotText, "sendking keystrokes (DELETE) give unexpected tesult\nExpected: %s\nActual: %s"%
                     (repr(expected), repr(gotText)))
        # trying backspace (back) key:
        sendKey(self.ctrl, "{back}")
        gotText = getEditText(self.ctrl)
        expected = ["az" + sep]
        self.assertTrue(expected == gotText, "sendking keystrokes (BACK, backspace) give unexpected tesult\nExpected: %s\nActual: %s"%
                     (repr(expected), repr(gotText)))

        # extra little trick:
        sendKey(self.ctrl, "{backspace}")
        gotText = getEditText(self.ctrl)
        expected = ["a" + sep]
        self.assertTrue(expected == gotText, "sendking keystrokes (BACK, backspace) give unexpected tesult\nExpected: %s\nActual: %s"%
                     (repr(expected), repr(gotText)))

        # now try to delete selection:
        sendKey(self.ctrl, "abcdefgh")
        time.sleep(0.1)
        activateMenuItem(self.app, W["commandselectall"])
        time.sleep(0.1)
        sendKey(self.ctrl, '{del}')
        time.sleep(0.1)
        gotText = getEditText(self.ctrl)
        expected = [sep]
        self.assertTrue(expected == gotText, "sendking keystrokes (delete) should have emptied the window\nExpected: %s\nActual: %s"%
                     (repr(expected), repr(gotText)))

def try_win32pad():
    """needs win32pad being open and with a file in, explores the line number facilities
    """
    appWindows = findTopWindows(wantedClass="win32padClass")
    print('# appWindows: %s'% len(appWindows))

    if len(appWindows) != 1:
        print('no win32pad found, or more instances: %s'% len(appWindows))
        return
    
    appWindow = appWindows[0]
    title = win32gui.GetWindowText(appWindow)
    print('appWindow win32pad: %s (%s)'% (appWindow, title))
    pprint.pprint(dumpWindow(appWindow))


    
           
def try_32770():
    """dialog for open etc from office etc
    
    
    only print details if Adres: of Address: is found...
    """
    appWindows = findTopWindows(wantedClass="#32770")
    print('# appWindows: %s'% len(appWindows))
    for appWindow in appWindows:
        title = win32gui.GetWindowText(appWindow)
        print('appWindow 32770: %s (%s)'% (appWindow, title))
    #    
    #    
    #for appWindow in appWindows:
    ##testControl = 723600
    ##if not testControl in appWindows:
    ##    print 'testup new test for your purposes'
    ##    return
    #pprint.pprint(dumpWindow(testControl))
    
        controls = findControls(appWindow, selectionFunction=selFuncExplorerAddress)
        if controls:
            print('controls: %s'% controls)
            hndle = controls[0]
            text = win32gui.GetWindowText(hndle)
            folder = text.split(": ", 1)[1]
            if os.path.isdir(folder):
                print('ok: %s'% folder)
            else:
                print('no folder: %s'% folder)
            pprint.pprint(dumpWindow(appWindow))
        #else:
    pprint.pprint(dumpWindow(appWindow))
            

    
    #controls = findControls(appWindow, selectionFunction=selFuncExplorerAddress)
    #if controls:
    #    hndle = controls[0]
    #    text = win32gui.GetWindowText(hndle)
    #    folder = text.split(": ", 1)[1]
    #    print folder
    #    if os.path.isdir(folder):
    #        print 'ok: %s'% folder
    #


        
        
def run_hotshot():
    import hotshot, hotshot.stats
    filePath = r'U:\messagesfunctionstest.prof'
    prof = hotshot.Profile(filePath)
    prof.runcall(run)
    prof.close()
    # printing with print_hotshot in miscqh...       
        

def run():
    print('starting unittestMessagefunctions')
    unittest.main()
    

if __name__ == "__main__":
    run()
    #run_hotshot()
    #try_explorer()  #open explorer window and run
    #try_32770()  # probeert alle 32770 windows en geeft van een open dialog window het interne pad
    
    #try_win32pad()