#
# unimacroutils.py (in dtactions/unimacro)
# was module natlinkutilsqh.py in unimacro
#
#  Quintijn Hoogenboom
#  February 2000/.../March 2021/July 2021
#
#  
"""a set of utility functions in dtactions/unimacro

Previous this was natlinkutilsqh.py in unimacro

"""
#pylint:disable=C0302, C0116, C0321, R0912, R0913, R0914, R0915, W0613, C0209, W0602, W0621, R1715, W0702
#pylint:disable=E1101, R1735
import time
import re
import os
import os.path
import sys
import collections
from pathlib import Path
import win32gui
import win32clipboard
import win32con
# import pywintypes

from dtactions import monitorfunctions  # elaborated version QH
from dtactions import autohotkeyactions

import natlink
from natlinkcore import natlinkutils
from natlinkcore import natlinkstatus
status = natlinkstatus.NatlinkStatus()

DEBUG = 0
# errors for (mainly) the command grammar:

class NatlinkCommandError(Exception):
    """NatlinkCommandError"""
class NatlinkCommandTimeOut(Exception):
    """NatlinkCommandTimeOut"""
class UnimacroError(Exception):
    """UnimacroError"""

# default Waiting Times:
defaultWaitingTime = 0.1        # 100 milliseconds
visibleWaitFactor = 3                 # default for Wait and W
longWaitFactor = 10                   # times 3 for visible wait
shortWaitFactor = 0.3                 # times 10 for long wait
                                # times 0.3 for short wait
debugMode = 0     #-1

pendingExecScripts = []

ProgInfo = collections.namedtuple('ProgInfo', 'progpath prog title toporchild classname hndle'.split(' '))



# class CaseInsensitiveSet(set):
#     """makes the comparison for adding new elements case insensitive
#     Adding takes the capitalisation of the first item added
#     """
#     def __hash__(self):
#         return hash(self.lower())
#     def __eq__(self, other):
#         return self.lower() == other.lower()

def getLanguage():
    """get current language, 3 characters
    """
    return status.get_language()

def getUserLanguage():
    """get current language, long name

    """
    raise NotImplementedError('getUserLanguage (the long language name) is no longer available, use natlinkstatus function "get_language"')

# utility functions----------------------------------------
## matchWindow from natlinkutils:
##def matchWindow(moduleInfo, modName, wndText):
##    if len(moduleInfo)<3 or not moduleInfo[0]: return None
##    curName = getBaseName(moduleInfo[0].lower() )
##    if curName != modName: return None
##    if -1 == moduleInfo[1].find(wndText): return None
##    return moduleInfo[2]

# A utility function which determines whether modInfo matches a specified
# module name.
# If modInfo is not given, it is got here.
# 
# Returns module name (program name!) on match and None on mismatch.
# This is a variant on matchWindow (from Joel), but now you can also check only the module, 
# and not the (sub) window of this module.
# Also you can specify a window title or a list/tuple of window titles.
# If you specify a window, it can be checked with exact title or exact case.
# If nothing is specified lower case strings are compared, and only part
# of the window title has to be given.
def getCurrentModuleSafe(nWait=5, waitingTime=0.01):
    """in case natlink.getCurrentModule returns None, try again a few times...
    default 5 x 0.01 = 0.05 (50 milliseconds)
    
    returns the tuple: 
    (full file name of foreground program, window title, window handle)
    """
    modInfo = natlink.getCurrentModule()
    if modInfo:
        return modInfo
    for i in range(nWait):
        time.sleep(waitingTime)
        modInfo = natlink.getCurrentModule()
        waited = (i+1)*waitingTime
        if modInfo:
            print(f'getCurrentModuleSafe, found modInfo after {waited} seconds')
            return modInfo

    print(f'getCurrentModuleSafe, could not find modInfo after {waited} seconds')
    return None

def matchModule(modName, wantedTitle=None, progInfo=None, titleExact=0, caseExact=0):
    """return modName if title matches, otherwise None
    """
    if not progInfo:
        progInfo = getProgInfo()
    if not progInfo.hndle:
        return None
    modName = progInfo.prog
##    print 'modName: %s, basename module: %s'% (modName, getBaseName(modInfo[0]))
    if not wantedTitle:
        return modName
   
    winTitle = progInfo.title
    if isinstance(wantedTitle, str):    
        if not caseExact:
            wantedTitle = wantedTitle.lower()
            winTitle = winTitle.lower()
        if titleExact:
            if winTitle == wantedTitle:
                return modName
        else:
            if winTitle.find(wantedTitle) >= 0:
                return modName
    elif isinstance(wantedTitle, (list, tuple)):
        for t in wantedTitle:
            if matchModule(modName, t, progInfo=progInfo, titleExact=titleExact, caseExact=caseExact):
                return modName
    return None

# Matches if only the window title ok. If modInfo is not given, it is extracted here.
#
# The wantedTitle can also be a list or tuple, in that case the routine is called
# recursive.
# titleExact and caseExact can be given.  If not strings are converted to lower case
# and a part of the window title may be given.

def matchTitle(wantedTitle, modInfo=None, titleExact=0, caseExact=0):
    """return progName if title matches, otherwise None
    """
    #pylint:disable=R0911    
    if modInfo is None:
        modInfo = natlink.getCurrentModule()
    if not (modInfo[0] and len(modInfo)) == 3:
        return None
    winTitle = modInfo[1]
    progName = getProgName(modInfo)
    
    if isinstance(wantedTitle, str):    
        if not caseExact:
            wantedTitle = wantedTitle.lower()
            winTitle = winTitle.lower()
        if titleExact:
            if winTitle == wantedTitle:
                return progName
            return False
        if winTitle.find(wantedTitle) >= 0:
            return progName
        return False
    if isinstance(wantedTitle, (list, tuple)):
        for t in wantedTitle:
            if matchTitle(t, modInfo, titleExact, caseExact):
                return progName
        return False
    print("unexpected place for matchTitle (unimacroutils)")
    return False

# Get basename of file:
def getBaseNameLower(name):
    """old fashioned way of getting the filename
    """
    return os.path.splitext(os.path.split(name)[1])[0].lower()


def getProgName(modInfo=None):
    """return the name of program (in the foreground)
    """
    if not modInfo:
        modInfo = natlink.getCurrentModule()
    return modInfo[0]

# ProgInfo = collections.namedtuple('ProgInfo', 'progpath prog title toporchild classname hndle'.split(' '))

def getProgInfo(modInfo=None):
    """returns program info as namedtuple (progpath, prog, title, toporchild, classname, hndle)


    now length 6, including the classname, but also a named tuple!!

    prog always lowercase
    
    title now with mixed (lower and upper case) letters.

    toporchild 'top' or 'child', or '' if no valid window
    """
    #pylint:disable=W0702
    if modInfo is None:
        try:
            modInfo = natlink.getCurrentModule()
        # print("modInfo through natlink: %s"% repr(modInfo))
        except natlink.NatError:
            progInfo = autohotkeyactions.getProgInfo()
            return progInfo
    
    _hndle = modInfo[2]
    if not _hndle:
        ## assume desktop, no foreground window, treat as top...
        return ProgInfo("", "", "", "top", "", 0)
    progpath = modInfo[0]
    prog = Path(modInfo[0].lower()).stem  
    title = modInfo[1]
    if isTopWindow(modInfo[2]):
        toporchild = 'top'
    else:
        toporchild = 'child'
         
    HNDLE = int(modInfo[2])
         
    classname = win32gui.GetClassName(HNDLE)

    return ProgInfo(progpath, prog, title, toporchild, classname, HNDLE)
            
def getClassName(modInfo=None):
    """returns the class name of the foreground window
    take modInfo (tuple or int (the handle), or get it here)
    """
    if isinstance(modInfo, int):
        _hndle = modInfo
    elif isinstance(modInfo, tuple):
        _hndle = modInfo[2]
    else:
        _hndle = natlink.getCurrentModule()[2]
    if _hndle:
        return win32gui.GetClassName(_hndle)
    return None

def matchWindow(criteria, modInfo=None, progInfo=None):
    """looks for a matching window based on the dictionary of criteria

    criteria can either be a key with None as value, or a key
    with a part of window title as value (string), or a key with
    a list of parts of window titles as value.

    special these are:
    'all' (everything matches, value is ignored, so normally None)
    'none'  (nothing matches)
    'empty' (matches when no valid progInfo is found)

    progInfo is a tuple: (prog, title, topchild, hndle),
    prog being the lower case name of the programme
    title being the lower case converted title
    topchild being 'top' if top window, 'child' if child window,
                        if no valid module info

    progInfo may be omitted as well as modInfo.
        For best performance progInfo should be given,
        otherwise giving mtake unimacro couldodInfo is also faster than omitting it.

    """
    #pylint:disable=R0911, R0912
    #print 'matchwindow(qh): %s'% criteria
    if not isinstance(criteria, dict):
        print(f'type criteria in matchWindow function should be dict, not: {criteria}')
        return None
    
    if 'all' in criteria:
        return 1 
    if 'none' in criteria:
        return None

    _progpath, prog, title, topchild, _classname, _hndle = progInfo or getProgInfo(modInfo)
    if 'empty' in criteria and prog == '':
        return 1
    if 'top' in criteria and topchild != 'top':
        return None
    if 'child' in criteria and topchild != 'child':
        return None
    if prog in criteria:
        pot = criteria[prog]  # part of title
        if not pot:
            return 1   # no title given, so all titles match
        if isinstance(pot, str):
            return title.find(pot) >= 0   # one part of title given, check this
        if isinstance(pot, list):
            for t in pot:
                if title.find(t) >= 0:    # more possibilities for part of title
                    return 1
    elif 'top' in criteria or 'child' in criteria:
        return 1
    return None

## bringup functionality (moved from natlinkutilsbj to here):
pendingBringUps = []
pendingExecScripts = []

def doPendingBringUps():
    """called from _control"""
    #pylint:disable=W0603, W0702
    global pendingBringUps
    if not pendingBringUps:
        return
    for p in pendingBringUps:
        try:
            #print 'try to do pending bringup: %s'% p
            natlink.execScript(p)
        except:
            print('delayed bringup does not work: %s'% p)
    pendingBringUps = []
    
    
def doPendingExecScripts():
    """do pending execScripts, if there are any (from gotBegin or initialisation phase)
    """
    #pylint:disable=W0603, W0702    
    global pendingExecScripts
    if not pendingExecScripts:
        return
    for p in pendingExecScripts:
        try:
            #print 'try to do pending execScript: %s'% p
            natlink.execScript(p)
        except:
            print('does not work (in doPendingExecScripts): %s'% p)
    pendingExecScripts = []    

def ExecScript(script, callingFrom=None):
    """dall execScript, and put in pending list if fails (from gotBegin or initialisation phase)
    """
    #pylint:disable=W0603
    global pendingExecScripts
    if callingFrom and (getattr(callingFrom, 'status', '') == 'new' or
                        getattr(callingFrom, 'inGotBegin', 0)):
        print(f'pending execScript for {callingFrom.getName()}')
        pendingExecScripts.append(script)
        return None
                            
    try:
        natlink.execScript(script)
    except natlink.NatError as t:
        print('error in execScript: %s'% script)
        print('message: %s'% t)
        print('-------')
        return None
    return 1

def AppBringUp(App, Exec=None, Args=None, windowStyle=None, directory=None, callingFrom=None):
    """central BringUp function, can maintain old bringups
    BJ, extended QH.
    can be called from actions (UnimacroBringUp, or (do_)BRINGUP.

    """
    #pylint:disable=W0603, R0913, R0912
    global pendingBringUps
    app = str(App).lower()
    if Args:
        if isinstance(Args, list):
            args=' '.join(Args)
        else:
            args = Args
    else:
        args = None

##    if (GetOS()!='windows_nt'):
##        app=App.lower()
##    else:
##        app=App
    if not Exec:
        if args:
            cmdline = 'AppBringUp "%s", "%s"'% (app, args)
        else:
            raise UnimacroError("AppBringup should have at least Exec or Args (now: %s, %s)"%
                                (Exec, Args))
    else:
        ExecString = Exec
        # if os.path.isfile(Exec):
        #     if AppBringUpDoExec83:
        #         (fullName,ext)=os.path.splitext(Exec)
        #         if ext=='':
        #             Exec83=os.path.dirname(Exec)+'\\'+os.path.basename(Exec)
        #         else:
        #             Exec83=Exec
        #         ExecString = Exec83
        
        if args:
            # komodo, notepad (??)
            ExecString += " %s"% args
        cmdline = 'AppBringUp "%s","%s"'% (app, ExecString)

    if windowStyle:
        cmdline += ',%s'% windowStyle
    else:
        cmdline += ","
    if directory:
        cmdline += ', "%s"'% directory

    cmdline = cmdline.strip(", ")        
    #print "AppBringUp: %s"% cmdline

    if callingFrom and (getattr(callingFrom, 'status', '') == 'new' or
                        getattr(callingFrom, 'inGotBegin', 0)):
        #print 'pending AppBringUp for %s, %s'% (callingFrom.getName(), script)
        pendingBringUps.append(cmdline)
        return None
    try:
        ## this is a tricky thing, execScript only recognises str, not unicode!
        ## as default, ascii or cp1252 or latin-1 is taken, the windows defaults.
        _result = natlink.execScript(cmdline)
    except natlink.NatError as _t:
        #print 'wait for bringup until later: %s'% cmdline
        pendingBringUps.append(cmdline)
        return None
##    print 'ready withAppBringUp'
    return 1

# word formatting parameters:
wf_WordWasAddedByTheUser = 0x00000001
wf_InternalUseOnly1 = 0x00000002
wf_InternalUseOnly2 = 0x00000004
wf_WordCanNotBeDeleted = 0x00000008
wf_NormallyCapitalizeTheNextWord = 0x00000010
wf_AlwaysCapitalizeTheNextWord = 0x00000020
wf_UppercaseTheNextWord = 0x00000040
wf_LowercaseTheNextWord = 0x00000080
wf_NoSpaceFollowingThisWord = 0x00000100
wf_TwoSpacesFollowingThisWord = 0x00000200
wf_NoSpacesBetweenWordsWithThisFlagSet = 0x00000400
wf_TurnCapitalizationModeOn = 0x00000800
wf_TurnUppercaseModeOn = 0x00001000
wf_TurnLowercaseModeOn = 0x00002000
wf_TurnOffSpacingBetweenWords = 0x00004000
wf_RestoreNormalSpacing = 0x00008000
wf_InternalUseOnly3 = 0x00010000
wf_SuppressAfterAWordWhichEndsInAPeriod = 0x00020000
wf_DoNotApplyFormattingToThisWord = 0x00040000
wf_DoNotResetTheSpacingState = 0x00080000
wf_DoNotResetTheCapitalizationState = 0x00100000
wf_NoSpacePreceedingThisWord = 0x00200000
wf_RestoreNormalCapitalization = 0x00400000
wf_FollowThisWordWithOneNewLineCharacters = 0x00800000
wf_FollowThisWordWithTwoNewLineCharacters = 0x01000000
wf_DoNotCapitalizeThisWordInATitle = 0x02000000
wf_InternalUseOnly = 0x04000000
wf_AddAnExtraSpaceFollowingThisWord = 0x08000000
wf_InternalUseOnly4 = 0x10000000
wf_InternalUseOnly5 = 0x20000000
wf_WordWasAddedByTheVocabularyBuilder = 0x40000000

##wf_AddedInVersion8 = 0x20000000
##wf_DroppedInVersion9 = 0x60000000 ^ 0xffffffff
##print 'dropped: %x'% wf_DroppedInVersion9
#  List of word formatting properties as can be handled in
#    getWordInfo and setWordInfo
wordFormatting = {
    'WordWasAddedByTheUser': 0x00000001,
    'InternalUseOnly1': 0x00000002,
    'InternalUseOnly2': 0x00000004,
    'WordCanNotBeDeleted': 0x00000008,
    'NormallyCapitalizeTheNextWord': 0x00000010,
    'AlwaysCapitalizeTheNextWord': 0x00000020,
    'UppercaseTheNextWord': 0x00000040,
    'LowercaseTheNextWord': 0x00000080,
    'NoSpaceFollowingThisWord': 0x00000100,
    'TwoSpacesFollowingThisWord': 0x00000200,
    'NoSpacesBetweenWordsWithThisFlagSet': 0x00000400,
    'TurnCapitalizationModeOn': 0x00000800,
    'TurnUppercaseModeOn': 0x00001000,
    'TurnLowercaseModeOn': 0x00002000,
    'TurnOffSpacingBetweenWords': 0x00004000,
    'RestoreNormalSpacing': 0x00008000,
    'InternalUseOnly3': 0x00010000,
    'SuppressAfterAWordWhichEndsInAPeriod': 0x00020000,
    'DoNotApplyFormattingToThisWord': 0x00040000,
    'DoNotResetTheSpacingState': 0x00080000,
    'DoNotResetTheCapitalizationState': 0x00100000,
    'NoSpacePreceedingThisWord': 0x00200000,
    'RestoreNormalCapitalization': 0x00400000,
    'FollowThisWordWithOneNewLineCharacters': 0x00800000,
    'FollowThisWordWithTwoNewLineCharacters': 0x01000000,
    'DoNotCapitalizeThisWordInATitle': 0x02000000,
    'InternalUseOnly': 0x04000000,
    'AddAnExtraSpaceFollowingThisWord': 0x08000000,
    'InternalUseOnly4': 0x10000000,
    'InternalUseOnly5': 0x20000000,
    'WordWasAddedByTheVocabularyBuilder': 0x40000000,
                  
    }

def ListOfProperties(props):
    """return a list of word properties
    """
    l = []
    keyList = list(wordFormatting.keys())
    keyList.sort()
    for k in keyList:
        if props & wordFormatting[k]:
            l.append(k)
    return l
            

def makeWordProperties(listOfProps):
    """return the number that is the adding of the props

    in version 8 if word was added by user, add special to it
    in version 9 drop 0x60000000 from any props number
    """
    props = sys.maxsize
    for l in listOfProps:
        props += wordFormatting[l]
    props = sys.maxsize & props

# add to string 't' the count 'n'
#    eg   '{PgUp}', '3' --> '{PgUp 3}',
#         'abc', 3    --> 'abccc' 
#         'abc', '16'    --> 'ab{c 16}' 
def doCount(t,n):
    """add to string 't' the count 'n'

    eg   '{PgUp}', '3' --> '{PgUp 3}',
         'abc', 3    --> 'abccc' 
         'abc', '16'    --> 'ab{c 16}'
    """
    if len(t) == 0:
        return t
    if isinstance(n, str):
        n = int(n)
    if n <= 1:
        return t
    # do the work with braces:a}
    if hasBraces(t):
        if len(t) > 2 and not hasBraces(t[1:-1]):
            # balancing braces  { }:
            return "%s %s}"% (t[:-1], n)
        return t * n
    # only last char repeated:
    return t[:-1] + t[-1] * n

def hasBraces(t):
    """return True if t has braces
    """
    if not (t and isinstance(t, str)):
        return None
    first = '{' in t or '}' in t
    if first:
        if '{' in t[1:]:
            return None  # more than one character, treat as no braces
    return first

def doModifier(t,mod):
    """add Shift, Ctrl or Alt to the string 
       'a' wordt '{Ctrl+Shift+a}'
       '{Left} wordt '{Ctrl+Left}'
    """
    if len(t) == 0:
        return t
    if t.lower().find(mod.lower()) >= 0:
        return t # duplicate modifier?
    if t[0] == '{' and t[-1] == '}':
        return '{%s%s'% (mod, t[1:])
    return '{%s%s}'% (mod, t)

def doCaps(t):
    """make t uppercase if caps > 0
    
    strange old function
    """
    #pylint:disable=C0200
    for i in  range(len(t)):
        if t[i].islower():
            t = t[:i] + t[i].capitalize() + t[i+1:]
    return t
        
# mouseState to be altered when a downstate is asked for. Must be released before other actions
mouseState = 0 # button that is down (left = 1, right=2)
mouseModifier = 0 # could also be 'ctrl' or 'shift'(to do!)
# mouseStartPosition set by RM (rememberMouse()), returned to by CancelMouse()
mouseStartPosition = () # x, y

# mouse actions, called by MP and RMP from _commands.py)
# mouse = 'left' is most common (and default) input. If mouse = 0, '' of None
#   no clicking is done, if the mouse state = 0
#   if the mouseState = 1 (left) or 2 (right) dragging will continue.
#
# nClick = 1 is the normal input, if nClick = -1, the button is pushed down,
# if nClick = 0, the button is released

#
# for going to Joels routine, double work:
buttons = dict(noclick=0, left=1, right=2, middle=4)
joelsButtons = ['noclick', 'left', 'right', 'middle']
##print 'buttons: ', buttons

mouseDown = ['', natlinkutils.wm_lbuttondown, natlinkutils.wm_rbuttondown, natlinkutils.wm_mbuttondown]
mouseUp = ['', natlinkutils.wm_lbuttonup, natlinkutils.wm_rbuttonup, natlinkutils.wm_mbuttonup]

def doMouse(absorrel, screenorwindow, xpos, ypos, mouse='left', nClick=1, modifier = 0):
    """complicated, but complete mouse routine
    
    - absorrel: 0 = absolute positions (in pixels), 1 = relative positions (between -1.0 and 1.0)
    - screenorwindow: where to take the position from:
                0 = complete screen
                1 = active window
                2 = relative to the current position
                ## new 2017, relative to current monitor
                3 = relative to current monitor
                4 = relative to current monitor work area (excluding eg task bar and Dragon bar)
                5 = relative to the client area of or window (eg the body of an e-mail)
    - xpos, ypos: relative or absolute position horizontal and vertical
    - mouse: the button to be clicked: 'left', 'right', 1, 2, or 0 if no button is clicked
    - nClick: 1 or 2 OR -1 if the button has to be pushed down. 0 if the button must be released.


    -combined values for mouse for vocola MP and RMP calls:
    leftdouble, rightdouble, middledouble, leftrelease, rightrelease, middlerelease,
    leftup, rightup, middleup (same as release), and click (identical with left) and
    -also doubleclickleft or doubleleft etc.
    - noclick or move
    -abbreviations: down for leftdown, up for leftup release for leftrelease.
    
    Pushing down buttons down must be remembered, and reversed before 
    other actions are performed.  
    Therefore the mouse state is remembered in the variable mouseState.
    
    2024: because of Dragon 16, with playEvents not working any more, this function restricts to
    movements and clicking, not dragging any more.
    
    
    """
     #pylint:disable=W0603
    global mouseState
    hndle = natlink.getCurrentModule()[2]
    rect = None
    if hndle:
        rect = win32gui.GetWindowRect(hndle)
    xold,yold = natlink.getCursorPos()
    # screenorwindow corrensponds with relativeTo in DNS
    # only 0 (whole screen), 1 (relative to window), 2 (relative to cursorPos)
    # are considered here 5, inside client area
    if screenorwindow not in [0,1,2,3,4,5]:  # only within current window
        print("doMouse, only screenorwindow 0 ... 5 valid:", screenorwindow)
        return
    
    if screenorwindow == 0:
        _width, _height, xMin, yMin, xMax, yMax = monitorfunctions.getScreenRectData()
    elif screenorwindow == 1:
        if hndle:
            rect = win32gui.GetWindowRect(hndle)
        else:
            print("doMouse, no valid foreground window")
            return
        _width, _height, xMin, yMin, xMax, yMax = getRectData(rect)
    elif screenorwindow == 3:
        # get current monitor
        rect = monitorfunctions.get_current_monitor_rect( (xold, yold) )
        _width, _height, xMin, yMin, xMax, yMax = getRectData(rect)
    elif screenorwindow == 4:
        # get current monitor work area (skipping task bar and Dragon bar)
        rect = monitorfunctions.get_current_monitor_rect_work( (xold, yold) )
        _width, _height, xMin, yMin, xMax, yMax = getRectData(rect)
        # print 'w: %s, h: %s, xMin: %s, yMin: %s, xMax: %s, yMax: %s'% (
        #     width, height, xMin, yMin, xMax, yMax)
    elif screenorwindow == 5:
        if hndle:
            ## TODOQH???
            _clxold, _clyold = win32gui.ScreenToClient(hndle, (xold, yold) )
            rect = win32gui.GetClientRect(hndle)
        else:
            print("doMouse, no valid foreground window")
            return
        # active window
        _width, _height, xMin, yMin, xMax, yMax = getRectData(rect)

    if screenorwindow == 2:  # relative to current position
        xnew = xold + xpos
        ynew = yold + ypos
        xp, yp = monitorfunctions.get_closest_position( (xnew, ynew) )
    elif screenorwindow in (0, 1, 3, 4, 5): 
        if absorrel == 1: # relative:
            xp, yp = relToCoord(xpos, xMin, xMax), relToCoord(ypos, yMin, yMax)
        else:
            xp, yp = xpos, ypos   # get rid of: getMouseAbsolute(ypos, yMin, yMax)
        xp, yp = checkMousePosition(xp,xMin,xMax), checkMousePosition(yp,yMin,yMax)
        if screenorwindow == 5:
            clPos = (xp, yp)
            xp, yp = win32gui.ClientToScreen(hndle, clPos)
        # print 'before get_closest_position: (%s, %s)'% (xp, yp)
        xp,yp =  monitorfunctions.get_closest_position( (xp, yp) )
        # print 'after  get_closest_position: (%s, %s)'% (xp, yp)
    
    if debugMode == -1:
        print('Mouse to: %s, %s' % (xp, yp))
    nclick = 1
    onlyMove = 0
    if mouse and isinstance(mouse, str):
        # special variables for vocola combined calls:
        if mouse in ("noclick", "0"):
            nclick = 0
            mouse = ""
        elif mouse == 'move':
            onlyMove = 1
            nclick = 0
            mouse = ""
        else:
            # get the click and button from text:
            mouse, nclick = catchClick(mouse)
            if not nclick is None:
                nClick = nclick
            if not mouse:
                mouse = 'left'
    if mouse:
        if isinstance(mouse, str):
            if mouse not in buttons:
                print('doMouse warning: invalid value for mouse: %s (taking left button)'% mouse)
            btn = buttons.get(mouse, 1)
        else:
            btn = mouse
    else:
        #print 'no click, mouseState: %s'% mouseState
        btn = 0
        nclick = 0
    #nClick = nClick or nclick   # take things from "mouse" if nClick  not used
    print('btn: %s, nClick: %s, current mouseState: %s'% (btn, nClick, mouseState))
    if onlyMove:
        print('onlyMove to %s, %x'% (xp, yp))
        set_mouse_position(xp, yp)   # absolute  
    elif not mouseState:  # ongecompliceerd
        if nClick > 0:
            natlink.playEvents([(natlinkutils.wm_mousemove, xp, yp)])
            if btn:
                if debugMode == -1:
                    print('ButtonClick %s, %s' % (btn, nClick))
                else:
    ##                Wait()  # before clicking!
                    buttonClick(btn, nClick)
##                for i in range(nClick):
##                    print 'click'
##                    natlink.playEvents([(mouseDown[btn], xp, yp)])
##                    natlink.playEvents([(mouseUp[btn], xp, yp)])
##                    Wait(0.01)
                    
        elif nClick == 0:
            natlink.playEvents([(natlinkutils.wm_mousemove, xp, yp)])
        elif nClick == -1:
            if btn:
                if (xold,yold) != (xp, yp):
                    print('mousedown at old: %s, %s (%s)'% (xold, yold, repr(mouseDown[btn])))
                    natlink.playEvents([(mouseDown[btn], xold, yold)])

    ##            Wait()  # before clicking!
                print('mousedown at new: %s, %s (%s)'% (xp, yp, repr(mouseDown[btn])))
                natlink.playEvents([(mouseDown[btn], xp, yp)])
                mouseState = btn
    elif btn:  # muis was omlaag!
        #print 'btn: %s, wanted mouseState: %s'% (btn, mouseState)
        if mouseState != btn or nClick > 0: # change button
            if (xold,yold) != (xp, yp):
                natlink.playEvents([(mouseDown[mouseState], xold, yold)])
                natlink.playEvents([(mouseDown[mouseState], xp, yp)])

##            Wait()  # before clicking!
            natlink.playEvents([(mouseUp[mouseState], xp, yp)])
            mouseState = 0
            doMouse(0, 0, xp, yp, btn, nClick)
        # end or continue dragging state:
        if (xold,yold) != (xp, yp):
            natlink.playEvents([(mouseDown[btn], xold, yold)])
            natlink.playEvents([(mouseDown[btn], xp, yp)])
        if nClick == -1:
            pass
        else:
            natlink.playEvents([(mouseUp[btn], xp, yp)])
            mouseState = 0
    else:
        # no btn, but mouseState, simply move:
        natlink.playEvents([(natlinkutils.wm_mousemove, xp, yp)])  

def catchClick(mouse):
    """return reduced mouse command and nClick

    if mouse starts or ends with "doubleclick" or "double", nClick=2
    if mouse starts or ends with "click" , nClick=2
    if mouse starts or ends with "down", nClick=-1
    if mouse starts or endswith "up" or "release", nClick=0
    """    
    for clicker, result in [("doubleclick", 2), ("double", 2),
                            ("click", 1), ("down", -1),
                            ("up", 0), ("release", 0)]:
        if mouse.startswith(clicker) or mouse.endswith(clicker):
            mouse = mouse.replace(clicker, "")
            return mouse, result
    # fall through:
    return mouse, None 

##buttons = makedict(left=1, right=2, middle=4)
##joelsButtons = ['', 'left', 'right', 'middle']
def buttonClick(button='left', nclick=1, modifiers=None):
    """do a natspeak buttonclick, but release mouse first if necessary
    """
     # make button numeric:
    buttons = {'left':1, 'right':2, 'middle':4}
    if button in buttons:
        button = buttons[button]
    if button not in [1, 2, 4]:
        raise UnimacroError('buttonClick invalid button: %s'% button)
    if nclick not in [1,2]:
        raise UnimacroError('buttonClick invalid number of clicks: %s'% nclick)
    if mouseState:
        releaseMouse()
        
    # natlinkutils.buttonClick(button, nclick, modifiers)
    if button == 1 and nclick == 1:
        script = 'ButtonClick'
    else:
        script = f'ButtonClick {button},{nclick}'
    # print(f'buttonClick via execScript: "{script}"')
    natlink.execScript(script)


    #     script = f'ButtonClick {button},{nclick}'
    # # print(f'buttonClick via execScript: "{script}"')
    # natlink.execScript(script)

def set_mouse_position(xp, yp):
    """set to absolute position
    
    when Dragon is running, with playEvents for version <= 15,
    via execScript if version >= 16
    
    when Dragon is NOT running, try via Autohotkey (to be done)
    """
    xp, yp = int(xp), int(yp)
    if natlink.isNatSpeakRunning():
        if natlink.getDNSVersion() <= 15:
            natlink.playEvents(  [(natlinkutils.wm_mousemove, xp, yp)])
        else:
            script = f'SetMousePosition 1, {xp},{yp}'
            print(f'buttonClick via execScript: "{script}"')
            natlink.execScript(script)  
    else:
        print('attempt set_mouse_position via autohotkey')

def releaseMouse():
    """restores the default mouseState
    """
     #pylint:disable=W0603
    global mouseState
    if mouseState:
        (xp,yp) = natlink.getCursorPos()
        print('releasing mouse at %s, %s (%s)'% (xp, yp, repr(mouseUp[mouseState])))
        natlink.playEvents([(mouseUp[mouseState], xp, yp)])
        mouseState = 0
        Wait()

def mousePushDown(mouse='left'):
    """only pushes the mouse down

    """
    if mouseState:
        endMouse()
##    print 'pushing down with %s'% mouse
    doMouse(0, 2, 0, 0, mouse, -1)  # abs, rel to window, x, y, click
            
def endMouse():
    """replaced by releaseMouse"""
    releaseMouse()
        
def rememberMouse():
     #pylint:disable=W0603    
    global mouseStartPosition
    mouseStartPosition = natlink.getCursorPos()

def cancelMouse():        
     #pylint:disable=W0603    
    global mouseStartPosition
    if not mouseStartPosition:
        endMouse()
        print('cancelMouse, no mouseStartPosition')
        return
    if mouseState:
        doMouse(0, 0, mouseStartPosition[0], mouseStartPosition[1],
                mouseState, 0)
    else:
        doMouse(0, 0, mouseStartPosition[0], mouseStartPosition[1],
                0, 0)
    mouseStartPosition = ()
        
def checkMousePosition(pos, Min, Max):
    #print 'checking:', pos, Min, Max
    if pos >= Max: return Max - 1
    if pos < Min: return Min
    return int(pos)

def getRectData(rect):
    """get from a rect the width, height, xMin, xMax, yMin, yMax
    
    example: see unittestMouse.py (in unimacro_test directory)
    """
    xMin, yMin, xMax, yMax = tuple(rect)
    width, height = rect[2] - rect[0], rect[3] - rect[1]
    return width, height, xMin, yMin, xMax, yMax 
    
def coordToRel(x, xMin, xMax, side=0):
    """calculate relative coordinate, must be between 0 and 1
    side == 0: calculate from xMin (result positive) (default)
    side == 1: calculate from xMax (result negative)

    (can be used for x and y coordinates)
    
    example: see unittestMouse.py (in unimacro_test directory)
    """
    if not xMin <= x < xMax:
        return None # None means no result
    width = xMax - xMin
    if side == 0:
        return float(x-xMin)/width
    return float(x-xMax)/width


def relToCoord(relValue, xMin, xMax):
    """calculate coordinate, with relative value given
    
    if relValue >= 0 calculate from xMin
    if relValue < 0 calculate from xMax

    (can be used for x and y coordinates)
    
    example: see unittestMouse.py (in unimacro_test directory)
    """
    width = xMax - xMin
    if relValue >= 0:
        return int(xMin + float(relValue) * width + 0.5)
    return int(xMax + float(relValue) * width + 0.5)
   
#def getMouseRelative(rel, range, Min, Max):
#    """newer function, range is calculated inside function
#    """
#    return relToCoord(rel, Min, Max)

# def getMouseAbsolute(Pos, Min, Max):
#     """strange function was meant for positive Min and Max
#     
#     obsolete QH 2015
#     
#     with multiple screens it sucks
#     """
# 
#     if Pos < Min:
#         return Min
#     elif Pos > Max:
#         return Max
#     return Pos

cornerDict = {0:"top left", 1:"top right", 2:"bottom left", 3:"bottom right"}
whichDict = {0: "whole screen", 1:"active window", 3:"active monitor", 5:"client area"}
absorrelDict = {0: "absolute", 1:"relative"}

def getMousePositionActionString(absorrel, which, position):
    """return the proper action line (MP or RMP)
    parameter as in getMousePosition below
    
    if invalid, print lines in Messages window and return ""
    
    """
    mousePos = getMousePosition(absorrel, which, position)
    if mousePos is None:
        print(("current mouse position is invalid for a Unimacro Shorthand Command with parameters:\n"
              "absorrel: %s (%s), which: %s (%s), corner position: %s (%s)"% (absorrel, absorrelDict[absorrel],
                                             which, whichDict[which],
                                             position, cornerDict[position])))
        return ""
    x, y = mousePos
    if absorrel:
        return "RMP(%s, %.6s, %.6s)"% (which, x, y)
    return "MP(%s, %s, %s)"% (which, x, y)

def printMousePosition(absorrel, printAll = 0):
    result = getMousePositions(absorrel, printAll=printAll)
    print(result)
    
def getMousePositions(absorrel, printAll = 0):
    """get output for printing the mouse positions
    
    this function can be invoked by PMP and PRMP and PALLMP.
    
    these positions are recognised by the Unimacro Shorthand Commands MP and RMP
    """
    L = []
    if printAll:
        L.append('-'*80)
        cornerRange = list(range(4))
    else:
        cornerRange = list(range(1))
    if absorrel:  # 1: relative:
        L.append('RELATIVE MOUSE POSITIONS:')
    else:
        L.append('ABSOLUTE MOUSE POSITIONS:')
    for whichnum, which in whichDict.items():
        L.append('---related to %s:'% which.upper())
        for cornerPos in cornerRange:
            L.append("%s: %s"% (cornerDict[cornerPos], getMousePositionActionString(absorrel, whichnum, cornerPos)))
        if printAll:
            L.append('-'*20)
    return '\n'.join(L)

def getMousePosition(absorrel=0, which=0, position=0):
    """get the parameters for doMouse
    
    absorrel: 0 abs, 1 rel
    which: 1: active window, 3: active monitor, 5: client area, 0: whole screen
    position: 0: topleft,  1: topright, 2: bottomleft, 3: bottomright
    
    result:
    None if mouseposition is invalid for the choice in question (outside boundaries probably)
    
    otherwise a 2 tuple:
    x, y 
    """
    #pylint:disable=R0912, W0612 
    x, y = currentPos = natlink.getCursorPos()
    _hndle = natlink.getCurrentModule()[2]
    print(f'which: {which}')
    if which == 0:
        # complete screen
        width, height, xMin, yMin, xMax, yMax = monitorfunctions.getScreenRectData()
        if absorrel == 0:  # 0 absolute, just return position
            # screen absolute, only position:
            return x, y
    elif which == 1:
        # active window
        rect = win32gui.GetWindowRect(_hndle)
        width, height, xMin, yMin, xMax, yMax = getRectData(rect)
    elif which == 3:
        # active monitor
        rect = monitorfunctions.get_current_monitor_rect( currentPos )
        width, height, xMin, yMin, xMax, yMax = getRectData(rect)
        ##
    elif which == 5:
        # client area coordinates:
        x, y = win32gui.ScreenToClient(_hndle, currentPos)
        rect = win32gui.GetClientRect(_hndle)
        width, height, xMin, yMin, xMax, yMax = getRectData(rect)
    else:
        raise ValueError(f'getMousePosition, variable "which" should be in (0, 1, 3, 5), not: {which}')

    # now test for boundaries and abs or rel:
    if x < xMin or x > xMax or y < yMin or y > yMax:
        print('mouse position outside active window')
        return None
    if absorrel:  # 1: relative:
        #print 'RELATIVE MOUSE POSITIONS:'
        #print '---related to ACTIVE WINDOW:'
        if position == 0: # top left
            x, y = coordToRel(x, xMin, xMax, 0), coordToRel(y, yMin, yMax, 0)
        if position == 1: # top right
            x, y = coordToRel(x, xMin, xMax, 1), coordToRel(y, yMin, yMax, 0)
        if position == 2: # bottom left
            x, y = coordToRel(x, xMin, xMax, 0), coordToRel(y, yMin, yMax, 1)
        if position == 3: # bottom right
            x, y = coordToRel(x, xMin, xMax, 1), coordToRel(y, yMin, yMax, 1)
        return x, y
                    
    # else:        # absolute
    #print 'ABSOLUTE MOUSE POSITIONS:'
    #print '---related to ACTIVE WINDOW:'    
    if position == 0: # top left
        x, y = x-xMin, y-yMin
    if position == 1: # top right
        x, y = x-xMax, y-yMin
    if position == 2: # bottom left
        x, y = x-xMin, y-yMax
    if position == 3: # bottom right
        x, y = x-xMax, y-yMax
    return x, y
            

# gives true if the window is a "TOP" window (with a button on the
# windows task bar. False otherwise (it is a child window then)
def isTopWindow(hndle):
    """returns 1 if window is top,

    behaviour changed in python 2.3.4, parent = 0, no exception any more
    TODO QH: is this correct?
    """
    #pylint:disable=W0702
    try:
        parent = win32gui.GetParent(hndle)
    except:
        return 1
    else:
        return parent == 0
        

# remember the current window, needed before waitForTitleChange can be done
# modInfo can be left away
Hndle = None
windowTitle = ""
def clearWindowHandle():
    """set to 0"""
    #pylint:disable=W0603    
    global Hndle
    Hndle = None

# This is called if the user clicks on the tray icon.  We simply cancel
# movement in all cases.
waitingCanceled = 0
iconState = 0
iconDirectory = os.path.join(status.getUnimacroDirectory(), 'icons')

# this sets the icontray for several waiting situations:::
def setTrayIcon(state=None, toolTip=None, comingFrom=None):
    """activate the trayIcon depending on the state
    
    switched off, reactivate later!!
    
    -If comingFrom is passed it can be either a instance (so the self of the calling grammar)
    or a function/method.
    -If comingFrom is an instance, a method "onTrayIcon" is taken (if possible) from the instance.
    In those cases this function (onTrayIcon) is called if the user clicks on the trayIcon.
    
    """
    #pylint:disable=W0101, R1710, W0603
    global iconState
    return None
    if state is None:
        natlink.setTrayIcon()
        return None # reset!
    #print 'natlinkqh setTrayIcon: %s'% state
    if state == 'waiting':
        toolTip = toolTip or 'unimacro is waiting'
        iconName = os.path.join(iconDirectory, 'waiting')
    else:
        iconName = state

    iconState = (iconState + 1)%2
##        print 'iconName: %s'% iconName
    if iconName[1] == ':':
        if not iconName.endswith('.ico'):
            # absolute path, attach extension:
            iconName = iconName + '.ico'
    elif iconName in ['right', 'left', 'up', 'down']:
        ## only if the 2 is not in the calling function yet...
        if iconState:
            iconName += '2'
    
    # if isinstance(comingFrom, instance):
    #     func = getattr(comingFrom, 'onTrayIcon', None)
    #     if func:
    #         toolTip += ' (grammar: %s)'% comingFrom.getName()
    # elif isinstance(comingFrom, func):
    #     func = comingFrom
    # elif comingFrom:
    #     func = None
    #     #print 'unimacroutils.setTrayIcon, comingFrom not of correct type (%s): %s'% (comingFrom, type(comingFrom))
    # else:
    #     func = None
    func = None    
    if func is None:
        toolTip += " (cannot be canceled)"
        try:
            natlink.setTrayIcon(iconName,toolTip)
        except natlink.NatError:
            natlink.setTrayIcon()
    else:
        try:
            natlink.setTrayIcon(iconName,toolTip,func)
        except natlink.NatError:
            print('cannot set tray icon "%s" (comingFrom: %s, func: %s), try to clear'% (iconName, comingFrom, func))
            natlink.setTrayIcon()
    return None
#         
# 
def clearTrayIcon():
    """clearTrayIcon, reactivate later
    """
    #pylint:disable=W0603
    global waitingCanceled, iconState
    waitingCanceled = 0
    iconState = 0
    # natlink.setTrayIcon()
# 

def rememberWindow(modInfo=None, progInfo=None, comingFrom=None):
    """remember the current window
    """
    #pylint:disable=W0603, W0613
    global Hndle, waitingCanceled, windowTitle
    # if not hndle true, then raised in error, because
    # rememberWindow is called before and not finished correct
    waitingCanceled = 0
    modInfo = modInfo or getCurrentModuleSafe()
    Hndle = modInfo[2]
    windowTitle = modInfo[1]
##    print 'set window to %s'% Hndle
    if not Hndle:
        print('warning, no window to remember: %s'% Hndle)
    # print 'rememberWindow, %s, %s'% (Hndle, windowTitle)
    return Hndle   

def waitForWindowTitle(titleName, nWait=10, waitingTime=0.1, comingFrom=None):
    """wait for window with title to come into the foreground

        waiting for a specific window title (word from...)
        all lowercase...
        timing out if nWait times is waited (eg 50 times), waitingTime suggested as 0.05 seconds
        (2.5 second total then!)

    """
##    print 'WWT, comingFrom: %s'% comingFrom
    titleName = titleName.lower()
    for _i in range(nWait):
##        print 'interrupted? %s'% comingFrom.interrupted
        if comingFrom and comingFrom.interrupted:
            # clearTrayIcon()
            print('waiting canceled (no TrayIcon handling')
            return None
        
        currentTitleName = natlink.getCurrentModule()[1].lower()
        # if empty (no window active) or matching:
        #print 'checking, currentTitle: %s, wantedTitle: %s'% (currentTitleName, titleName)
        if currentTitleName:
            if isinstance(titleName, str):
                if currentTitleName.find(titleName) >= 0:
                    # clearTrayIcon()
                    #print 'string test, found'
                    return 1
            else:
                for t in titleName:
                    if currentTitleName.find(t) >= 0:
                        # clearTrayIcon()
                        #print 'list test, found on  %s'% t
                        #found
                        return 1
        #print 'waiting, currentTitle: %s, wantedTitle: %s'% (currentTitleName, titleName)
        # setTrayIcon('waiting', comingFrom=comingFrom)
        Wait(waitingTime, comingFrom=comingFrom)
    # clearTrayIcon()
    print('Waiting for window title "%s" lasts too long, failed\nGot title: %s' % (titleName, currentTitleName))
    return None


def waitForNewWindow(nWait=10, waitingTime=0.1, comingFrom=None, debug=None, progInfo=None):
    """wait for a new window to appear in the foreground
    
    rememberWindow must run before. nWait and waitingTime as suggested above
    """
    #pylint:disable=R1710, W0613, W0612
    if Hndle is None:
        raise NatlinkCommandError("waitForNewWindow, no valid old hndle, do a rememberWindow() first")
    for i in range(nWait):
        if waitingCanceled:
            # clearTrayIcon()
            print('waiting canceled')
            return None

        modInfo = getCurrentModuleSafe()
        if not modInfo:
            print('waitForNewWindow failed, no modInfo')
            # clearTrayIcon()
            return None
        
        stepsToBeStable = max(3, i) # if it took longer to bring window in front, test more steps
        progN, titleN, hndleN = modInfo
        if hndleN != Hndle:
            # new window, wait for stable situation
            succes = 0
            for j in range(stepsToBeStable*3):
                # changed! reset wait a little and OK:
                Wait(waitingTime)
                newModInfo = getCurrentModuleSafe()
                if modInfo == newModInfo:
                    succes += 1
                    if succes >= stepsToBeStable:
                        break
                else:
                    succes = 0
                    modInfo = newModInfo
            else:
                print("waitForNewWindow: Found new window, but modInfo was not stable more than %s times"% stepsToBeStable)
                clearTrayIcon()
                return None
            if debug and j > stepsToBeStable + 1:
                extra = j - stepsToBeStable - 1
                print('title stable times after %s extra steps'% extra)
            clearTrayIcon()
            return 1
        setTrayIcon('waiting')
        Wait(waitingTime)
    clearTrayIcon()
    print("waiting for new window lasts too long, fail")
    return None

def waitForNewWindowTitle(nWait=10, waitingTime=0.1, comingFrom=None, progInfo=None):
    """wait for a new window with a matching title (???)
    """
    #pylint:disable=R1710, W0613
    if Hndle is None:
        raise NatlinkCommandError("waitForNewWindow, no valid old hndle, do a rememberWindow() first")
    for _i in range(nWait):
        if waitingCanceled:
            clearTrayIcon()
            print('waiting canceled')
            return None

        modInfo = natlink.getCurrentModule()
        if modInfo[1] and modInfo[1] != windowTitle:
            # changed! reset wait a little and OK:
            Wait(waitingTime)
            clearTrayIcon()
            return 1
        setTrayIcon('waiting')
        Wait(waitingTime)
    clearTrayIcon()
    print("waiting for new window title lasts too long, fail")
    return None
    
# return to window that was remembered by rememberWindow.
# nWait and waitingTime as suggested above.
def returnToWindow(nWait=5, waitingTime=0.05, winHandle=None, winTitle=None, **kw):
    """return to previous remembered window
    
    mostly do not specify winHandle and winTitle, as it is set as global
    variables hndle and windowTitle
    
    """
    winHandle = winHandle or Hndle
    winTitle = winTitle or windowTitle
    if not winHandle:
        print('returnToWindow, no window handle to return to, do nothing')
        return None
    if not win32gui.IsWindow(winHandle):
        print('returnToWindow, not a valid window: %s (%s)'% (winHandle, winTitle))
        return None
    # go:
    # print 'returning to window: %s (%s)'% (winHandle, winTitle)
    return SetForegroundWindow(winHandle, waitingTime=waitingTime, nWait=nWait)
    

# for returning a string, int or float:
restring = re.compile(r'[\'"].*[\'"]$')
reint = re.compile(r'[-]?[0-9]+$')
refloat = re.compile(r'[-]?[0-9]*[.][0-9]+$')

def arg2IntFloatOrString(arg):
    """utility function
    """
    #pylint:disable=C0321
    if reint.match(arg): return  int(arg)
    if refloat.match(arg): return  float(arg)
    return arg


def setDebugMode(n):
    """set the debug mode for commands that come from the "_commands" grammar
    """
    #pylint:disable=W0603    
    global debugMode
    debugMode = n

    
##def convertToNatSpeakArgString(input):
##    if not input:
##        return []
##    out = ' '
##    for i in input:
##        out = out + `i` + ', '
##    return out[:-2]


#  Wait some time, if no time is given,
#  wait 0.1 seconds = 100 milliseconds
def Wait(tt=None, comingFrom=None):
    """wait in milliseconds (version 5) or seconds (version 7)
    
    assume maximum time is 10 seconds, so if t > 10, milliseconds were given.
    
    between five and 10 a warning is given!
    
    no input: defaultWaitingTime is taken
    """
    t = tt or defaultWaitingTime
    if t > 10:
        print('warning, changed waiting time to seconds: %s'%t)
        t = t/1000.0
    elif t >= 5:
        print('warning, long waiting time: %s'% t)
        
    if debugMode == -1:
        print("Wait %s" % t)
    elif debugMode:
        t = t*debugMode
    if comingFrom is None:
        time.sleep(t)
    else:
        comingFrom.Wait(t)


def visibleWait():
    """wait less than defaultWaitingTime:
    """
    Wait(defaultWaitingTime*visibleWaitFactor)
    

def shortWait():
    """wait less than defaultWaitingTime:
    """
    Wait(defaultWaitingTime*shortWaitFactor)
    
def longWait():
    """wait 10 times the defaultWaitingTime:
    """
    Wait(defaultWaitingTime*longWaitFactor)

recharspace = re.compile('^[a-zA-Z- ]+$')
def addWordIfNecessary(w):
    """ see if word is already there, if not add.

    """    
    w = w.strip()
    if not w:
        return
    if not recharspace.match(w):
        print('invalid character in word to add: %s'% w)
        return
        
    isInVoc = natlink.getWordInfo(w,1) is not None
    isInActiveVoc = natlink.getWordInfo(w,0) is not None
    if isInActiveVoc:
        return
    try:
        if isInVoc:    # from backup vocabulary:
            print('make backup word active:', w)
            natlink.addWord(w,0)
            add2logfile(w, 'activated words.txt')
        else:
            print('adding word ', w)
            natlink.addWord(w)
            add2logfile(w, 'new words.txt')
            
    except natlink.InvalidWord:
        print('not added to vocabulary, invalid word: %s'% w)

def deleteWordIfNecessary(w):
    """delete word from vocabulary -- if needed
    """
    if not w:
        return
    isInActiveVoc = natlink.getWordInfo(w,0) is not None
    if isInActiveVoc:
        natlink.deleteWord(w)

def debugPrint(t):
    """print to debug file (is this working???)
    """
    if not DEBUG:
        return
    print(t)
    
def GetForegroundWindow():
    """return the handle of the current foreground window
    """
    return win32gui.GetForegroundWindow()
    

def SetForegroundWindow(h, waitingTime=0.1, nWait=3, debug=None):
    """gets the window in front
    
    Autohotkey is used if active!!
    
    When the switch is not made within 3 steps (of default waiting time),
    win+b (giving the system tray) is sent, and then the waiting cycle is done again a few times.
    
    tested in PyTest/unittestClipboard.py: testSwitchingWindows
    """
    #pylint:disable=R0912, C0321
    if not h:
        raise UnimacroError("no valid handle given for set foreground window: %s"% h)
    if autohotkeyactions.ahk_is_active():
        script = f'WinActivate, ahk_id {h}'
        ## TODOQH
        autohotkeyactions.do_ahk_script(script)
        # print(f'result SetForegroundWindow: {result}, wanted hndle: {h}')
        curHndle = win32gui.GetForegroundWindow()
        if curHndle == h:
            # print("autohotkey WinActivate succeeded: %s, wait 0.3 more seconds"% h)
            # time.sleep(0.3)
            return 1
        print(f'autohotkey did not get window with hndle {h} in the foreground')
        return None

    curHndle = win32gui.GetForegroundWindow()
    if curHndle == h:
        if debug: print('got it in one shot!! %s'% h)
        return 1
    
    if not win32gui.IsWindow(h):
        print('SetForegroundWindow: not a valid window hndle: %s'% h)
        return None
        
    for doKeystroke in [""]:  #  "{win+b}"]: ####, "{win+m}"]:
        result = _setForegroundWindow(h, doKeystroke, waitingTime=waitingTime, nWait=nWait, debug=debug)
        if result:
            if doKeystroke:
                if debug: print('SetForegroundWindow to %s success, after keystroke: %s'% (h, doKeystroke))
            else:
                if debug: print('SetForegroundWindow to %s success'% h)
            return result
        # if doKeystroke:
        #     if debug: 'return to calling window: %s'% curHndle
        #     win32gui.SetForegroundWindow(curHndle)
    return None          
    
def _setForegroundWindow(hndle, doKeystroke=None, waitingTime=0.1, nWait=3, debug=None):
    """try to switch to hndle
    """
    #pylint:disable=W0621, C0321, R0912
    if doKeystroke:
        if debug: print('try to get %s in foreground with keystroke: %s'% (hndle, doKeystroke))
        ## TODOQH make keystrokes here
        # natlinkutils.playString(doKeystroke)
    if win32gui.IsIconic(hndle):
        if debug: print('window %s is iconic, try to restore...')
        monitorfunctions.restore_window(hndle)
        Wait()
        for i in range(nWait):
            if win32gui.IsIconic(hndle):
                if debug and i:
                    print('window is still iconic, wait longer %s'% i)
                Wait(waitingTime)
            else:
                break
        else:
            # if debug: print('_setForegroundWindow, %s is still "Iconic"'% h)
            return None        
        
    win32gui.SetForegroundWindow(hndle)
  
    # except pywintypes.error as details:
    #     if details[0] in [0, 183]:
    #         pass
    #         # print 'could not bring to foreground: %s'% hndle
    #     else:
    #         raise
    for i in range(3):
        newH = getCurrentModuleSafe()[2]
        if newH == hndle:
            Wait()  #extra for safety
            return 1
        if debug and i:
            print("_setForegroundWindow, waiting %s (for %s, current: %s)"% (i, hndle, newH))
    if debug: print('_setForegroundWindow, with keystroke %s no result to set foregroundwindow to %s'% (doKeystroke, hndle))
    return None

titleHandles = {}
# switch to window with text in the title:
def switchToWindowWithTitle(wantedTitle, caseExact=0, titleExact=0):
    """keep cache of title handles for faster execution

    """
    #pylint:disable=W0603, C0321
    global titleHandles    
    extension = ''
    if titleExact: extension += "T"
    if caseExact: extension += "C"
    functionName = "lookForWindowText" + extension
    try:
        lookForFunction = globals()[functionName]
    except KeyError:
        print("invalid function in switchToWindowWithTitle: %s" % functionName)
        return None
    if len(titleHandles) > 100: # in case too many different switches
        print('clearing switchWindow title handles')
        titleHandles.clear()
    tryHandle = titleHandles.get(wantedTitle, 0)
    if tryHandle and not lookForFunction(tryHandle, wantedTitle):
        pass
    else:
        win32gui.EnumWindows(lookForFunction, wantedTitle)
        titleHandles[wantedTitle] = natlink.getCurrentModule()[2]
    Wait(0.1) # safety
    return 1
        
# look for window with part of text, case exact:
def lookForWindowTextC(hwnd, text):
    #pylint:disable=C0116
    if win32gui.GetWindowText(hwnd).find(text) >= 0:
        SetForegroundWindow(hwnd)
        return None
    return 1
# look for window Messages, default case, part of text, all to lower case:
def lookForWindowText(hwnd, text):
    #pylint:disable=C0116
    if win32gui.GetWindowText(hwnd).lower().find(text.lower()) >= 0:
        SetForegroundWindow(hwnd)
        return None
    return 1
# look for window with exact match:
def lookForWindowTextTC(hwnd, text):
    #pylint:disable=C0116
    if win32gui.GetWindowText(hwnd).strip() == text.strip():
        SetForegroundWindow(hwnd)
        return None
    return 1
# look for window Messages, whole title, but convert to lower case:
def lookForWindowTextT(hwnd, text):
    #pylint:disable=C0116
    if win32gui.GetWindowText(hwnd).lower().strip() == text.lower().strip():
        SetForegroundWindow(hwnd)
        return None
    return 1

def returnFromMessagesWindow():
    #pylint:disable=C0116    
    if matchModule('natspeak', 'Messages from Python Macros'):
        natlink.playString("{Alt+Tab}", natlinkutils.hook_f_systemkeys)

            
# Returns the date on a file or 0 if the file does not exist        
def getFileDate(fileName):
    #pylint:disable=C0116, C0321    
    try: return os.path.getmtime(fileName)
    except OSError: return 0        # file not found


def printListorString(arg):
    #pylint:disable=C0116    
    if isinstance(arg, str):
        print(arg)
    elif isinstance(arg, (list, tuple)):
        for l in arg:
            print(l)

      
#QH13062003  clipboard helper functions--------------------------------------

previousClipboardText = []

def saveClipboard():
    """Saves and clears the clipboard, and puts the text content in a global variable

    The global variable "previousClipboardText"
    is used to restore to the clipboard in the function "restoreClipboard"
    
    No input parameters, no result, the global variable is set

    """
    #pylint:disable=W0702    
    # global previousClipboardText
    t = getClipboard()
    previousClipboardText.append(t)
    for _i in range(10):
        try:
            win32clipboard.OpenClipboard()
            break
        except:
            time.sleep(0.1)
            continue
        else:
            print("could not open, save and empty the clipboard")
            return
    try:
        win32clipboard.EmptyClipboard()
    finally:
        win32clipboard.CloseClipboard()
    #print 'previousClipboardText: %s'% repr(previousClipboardText)

def clearClipboard():
    """clears the clipboard

    No input parameters, no result,

    """
    #pylint:disable=W0702
    # t0 = time.time()
    for _i in range(3):
        try:
            win32clipboard.OpenClipboard()
        except:
            print('error opening the clipboard')
            shortWait()
        else:
            break
    try:
        win32clipboard.EmptyClipboard()
    finally:
        win32clipboard.CloseClipboard()
    # print 'clearClipboard: %.4f'% (time.time() - t0,)

format_unicode = win32con.CF_UNICODETEXT

def restoreClipboard():
    """Restores the previously saved clipboard text into the clipboard

    No input, no result. The global variable is emptied.

    """
    #pylint:disable=W0702       
    # global previousClipboardText
    if previousClipboardText:
        t = previousClipboardText.pop()
    else:
        print('No "previousClipboardText" available, empty clipboard...')
        t = None
        return
    try:
        for _i in range(10):
            try:
                win32clipboard.OpenClipboard()
                break
            except:
                time.sleep(0.1)
                continue
            else:
                print("could not restore clipboard")
                return
        win32clipboard.EmptyClipboard()
        if t:
            win32clipboard.SetClipboardData(format_unicode, t)
    finally:
        win32clipboard.CloseClipboard()

def getClipboard():
    """get clipboard through natlink, and strips off backslash r   

    """
    #pylint:disable=W0702       
    # t0 = time.time()
    t = None
    for _i in range(3):
        try:
            t = natlink.getClipboard()
            if not t is None:
                # print ' at try', i
                break
        except:
            print('getClipboard, but apparently empty')
            shortWait()
        else:
            break

    if t:
        # print 'got clipboard: "%s"'% repr(t)
        t = t.replace('\r', '')
        return t    
    print('getClipboard, got clipboard, but empty')
    return ''



def setClipboard(t, format=1):
    """set clipboard with text
    format = win32con.CF_UNICODETEXT (13): as unicode

    """
    #pylint:disable=W0622
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(format, t)
    except:
        print(f'exception in unimacroutils/setClipboard of "{t}"')
    finally:
        win32clipboard.CloseClipboard()

def checkLists(one, two):
    """returns to lists, only in first, only in second

    if lists are equal 2 empty lists are returned
    """
    onlyone = []
    onlytwo = []
    for o in one:
        if not o in two:
            onlyone.append(o)
    for t in two:
        if not t in one:
            onlytwo.append(t)
        return onlyone, onlytwo

def cleanString(s):
    """converts a string with leading and trailing and
    intermittent whitespace into a string that is stripped
    and has only single spaces between words, newlines are removed.

>>> cleanString('foo  bar')
'foo bar'
>>> cleanString('\\n foo \\n\\n  bar ')
'foo bar'
    """
    return ' '.join([x.strip() for x in s.split()])



def cleanParagraphs(t):
    """make long paragraphs from selection, cleaning newlines and white space

    multiple empty lines are ignored, so spacing is 2 newlines at most.

    returns the new string    
    """
    T = t.split('\n\n')
    OUT = []
    hadEmptyLine = 0
    for p in T:
        if p.strip():
            OUT.append(cleanString(p))  
            hadEmptyLine = 0
        else:
            if not hadEmptyLine:
                OUT.append('')
            hadEmptyLine = 1
    return '\n\n'.join(OUT)
            

def stripSpokenForm(t):
    """strip the spoken from from a word from the Dragon Vocabulary
    """
    
    p = t.find('\\')
    if p == 0:
##        print 'backslash first: |%s|'% t
        if t[:4] in ['\\spa']:
            return ' '
        if t == '\\\\backslash':
            return '\\'
    elif p > 0:
        return t.split('\\', 1)[0]
    return t

# # get screen parameters:
# monitorfunctions.monitor_info()
# screenRect = monitorfunctions.VIRTUAL_SCREEN[:]
# screenWidth, screenHeight, screenXMin, screenYMin, screenXMax, screenYMax = getRectData(screenRect)
# logFolder = None
# try:
#     uud = getUnimacroUserDirectory()
#     if uud:
#         logFolder = os.path.join(uud, getLanguage() + "_log", getUser())
#         utilsqh.createFolderIfNotExistent(logFolder)
#         # print 'natlinkutilsqh, logfolder: %s'% logFolder
# except natlink.NatError:
#     pass
#     # print 'natlinkutilsqh, no logFolder active'


def add2logfile(word, filename):
    """add word to the filename in logFolder
    
    Disabled for now
    """
    #pylint:disable=E0602, W0702
    logFile = False
    if not logFile:
        return # silent
    try:
        with open(os.path.join(logFolder, filename), 'a', encoding='utf-8') as fp:
            fp.write(word + '\n')
        print('written to %s: %s' % (os.path.join(logFolder), filename))
    except:
        pass
    
def _test():
    #pylint:disable=C0415
    import doctest
    return  doctest.testmod()

if __name__ == "__main__":
    _test()
    
    natlink.natConnect()
    _progInfo = getProgInfo()
    print("progInfo: %s"% repr(_progInfo))
    natlink.natDisconnect()
    