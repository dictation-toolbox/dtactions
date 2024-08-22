# Module     : winGuiAuto.py
# Synopsis   : Windows GUI automation utilities
# Programmer : Simon Brunning - simon@brunningonline.net
# Date       : 25 June 2003
# Version    : 1.0 
# Copyright  : Released to the public domain. Provided as-is, with no warranty.
# Notes      : Requires Python 2.3, win32all and ctypes 



# Modifications by Tim Couper - tim@tizmoi.net
# 22 Jul 2004
# findControls: deduplicates the list to be returned 
# findControl: handles win32gui.error from initial call to findControls
# getMenuInfo: improved algorithm for calculating no. of items in a menu
# activateMenuItem: improved algorithm for calculating no. of items in a menu
#           
# GLOBALLY corrected spelling: seperator -> separator
#                            : descendent -> descendant
# added findMenuItem
# adapted for NatSpeak/NatLink interaction, Quintijn Hoogenboomhallo
# october 2009. (starting with watsup utilities)
# pylint: disable=C0302, C0115
'''"Windows GUI automation utilities"

The standard pattern of usage of winGuiAuto, originally written by TimCouper,
is in three stages; identify a
window, identify a control in the window, trigger an action on the control.

The end result is always that you wish to have an effect upon some Windows GUI
control.

The first stage is to identify the window within which the control can be
found. To do this, you can use either :code:`findTopWindow` or :code:`findTopWindows`. 

Usually, specifying caption text, the window's class, or both, will be
sufficient to identify the required window. (Note that the window's class is a
Windows API concept)

Testing can be done interactively at the bottom of the file. You need the hndle of the application in order to
run most of the test functions

Helper class and functions
==========================

'''

import array
import ctypes
import types
import os
import struct
import time
import pprint
import re
## python 3 only:
from urllib.parse import unquote

import pywintypes
import win32api
import win32con
import winxpgui as win32gui

GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW

# for SendKeys:
braceExact = re.compile (r'[{][^}]+[}]$')
hasBraces = re.compile (r'([{].+?[}])')
BracesExtractKey = re.compile (r'^[{]((alt|ctrl|shift)[+])*(?P<k>[^ ]+?)( [0-9]+)?[}]$', re.I)

## Scintilla constants:
# from dragonfly.actions import scintillacon

def findTopWindow(wantedText=None,
                  wantedClass=None):
    '''Find the hwnd of a top level window.
    
    You can identify windows using caption (part of) or class,

    Arguments:
      wantedText     Text which the required window's caption must
                     *contain*. (Case insensitive match.)
      wantedClass    Class to which the required window must belong.
    Returns:
      handle if found (the first match)
      None if not found
    '''
    if not (wantedText or wantedClass):
        raise ValueError("findTopWindow, wantedText and/or wantedClass must be passed, not: %s and %s"%
                         (wantedText, wantedClass))
    topWindow = findTopWindows(wantedText, wantedClass, checkImmediate=1)
    return topWindow # checking already done!

# for checking on the fly:
wText = None
wClass = None
checkI = None
selFunc = None

def findTopWindows(wantedText=None, wantedClass=None, selectionFunction=None, checkImmediate=None):
    '''Find the hwnd of top level windows.
    
    You can identify windows using captions, classes, a custom selection
    function, or any combination of these. (Multiple selection criteria are
    ANDed. If this isn't what's wanted, use a selection function.)

    Arguments:
    wantedText          Text which required windows' captions must contain.
    wantedClass         Class to which required windows must belong.
    selectionFunction   Window selection function. Reference to a function
                        should be passed here. The function should take hwnd as
                        an argument, and should return True when passed the
                        hwnd of a desired window.

    Returns:            A list containing the window handles of all top level
                        windows matching the supplied selection criteria.

    Usage example:      optDialogs = findTopWindows(wantedText="Options")
    
    checkImmediate: if true, check the options inside the _windowEnumerationHandler, and
                    quit if criterium has been met (only for wantedText and wantedClass)
    
    '''
    # pylint: disable=W0603, R0912
    global wText, wClass, checkI
    results = []
    topWindows = []
    if checkImmediate:
        checkI, wText, wClass, _selFunc = 1, _normaliseText(wantedText), wantedClass, selectionFunction
    else:
        checkI, wText, wClass, _selFunc = None, None, None, None
        
    for i in range(10):
        try:
            win32gui.EnumWindows(_windowEnumerationHandler, topWindows)
        except win32gui.error:
            print('got error in EnumWindows, try again: %s'% i)
        else:
            break
    if checkI:
        checkI = wText = wClass = None
        if topWindows:
            # return 1 window handle:
            return topWindows[0][0]
        return 0
        
    # checking the results:
    if wantedText:
        wantedText = _normaliseText(wantedText)
    for result in topWindows:
        if isinstance(result, tuple) and len(result) == 3:
            hwnd, windowText, windowClass = result
        else:
            print('invalid entry')
            continue
        if wantedText and \
           _normaliseText(windowText).find(wantedText) == -1:
            continue
        if wantedClass and not windowClass.find(wantedClass) == 0:
            continue
        if selectionFunction and not selectionFunction(hwnd):
            continue
        results.append(hwnd)
    return results
    
def getForegroundWindow():
    """get hndle of the current foreground window
    """
    hndle = win32gui.GetForegroundWindow()
    return hndle
    
def dumpTopWindows(doAll=None):
    '''TODO'''
    defaultTitles = ['', 'Default IME', 'Engine Window']
    defaultClasses = ['MSCTFIME UI']
    topWindows = []
    try:
        win32gui.EnumWindows(_windowEnumerationHandler, topWindows)
    except win32gui.error:
        print('got error enumerating windows')
        
    if doAll:
        return topWindows
    selectedWindows = [(hndle,title,clname) for (hndle,title,clname) in topWindows
        if title not in defaultTitles and clname not in defaultClasses]
    return selectedWindows
    
def dumpWindow(hwnd):
    '''Dump all controls from a window into a nested list
    
    Useful during development, allowing to you discover the structure of the
    contents of a window, showing the text and class of all contained controls.
    
    Think of it as a poor man's Spy++. ;-)

    Arguments:      The window handle of the top level window to dump.

    Returns         A nested list of controls. Each entry consists of the
                    control's hwnd, its text, its class, and its sub-controls,
                    if any.

    Usage example:  replaceDialog = findTopWindow(wantedText='Replace')
                    pprint.pprint(dumpWindow(replaceDialog))
    '''
    windows = []
    try:
        win32gui.EnumChildWindows(hwnd, _windowEnumerationHandler, windows)
    except win32gui.error:
        # No child windows
        return windows # empty here
    for window in windows:
        childHwnd, _windowText, _windowClass = window
        window_content = dumpWindow(childHwnd)
        if window_content:
            windows.append(window_content)
       
    def dedup(thelist):
        '''De-duplicate deeply nested windows list.'''
        def listContainsSublists(thelist):
            return bool([sublist
                         for sublist in thelist
                         if isinstance(sublist, list)])
        found=[]
        def dodedup(thelist):
            todel = []
            for index, thing in enumerate(thelist):
                if isinstance(thing, list) and listContainsSublists(thing):
                    dodedup(thing)
                else:
                    if thing in found:
                        todel.append(index)
                    else:
                        found.append(thing)
            todel.reverse()
            for todel in todel:
                del thelist[todel]
        dodedup(thelist)
    dedup(windows)
    
    return windows

def findControl(topHwnd,
                wantedText=None,
                wantedClass=None,
                selectionFunction=None,
                maxWait=1,
                retryInterval=0.1):   ###, checkImmediate=None):
    '''Find a control.
    
    You can identify a control within a top level window, using caption, class,
    a custom selection function, or any combination of these. (Multiple
    selection criteria are ANDed. If this isn't what's wanted, use a selection
    function.)
    
    If no control matching the specified selection criteria is found
    immediately, further attempts will be made. The retry interval and maximum
    time to wait for a matching control can be specified.

    Arguments:
    topHwnd             The window handle of the top level window in which the
                        required controls reside.
    wantedText          Text which the required control's captions must contain.
    wantedClass         Class to which the required control must belong.
    selectionFunction   Control selection function. Reference to a function
                        should be passed here. The function should take
                        hwnd, controlText, controlClass as its arguments
                        and should return True when passed the
                        hwnd of the desired control.
    maxWait             The maximum time to wait for a matching control, in
                        seconds.
    retryInterval       How frequently to look for a matching control, in
                        seconds

    Returns:            The window handle of the first control matching the
                        supplied selection criteria.
                    
    Raises:
    WinGuiAutoError     When no control or multiple controls found.

    Usage example:      optDialog = findTopWindow(wantedText="Options")
                        okButton = findControl(optDialog,
                                               wantedClass="Button",
                                               wantedText="OK")
    '''
    # pylint: disable=R0913
    controls = findControls(topHwnd,
                            wantedText=wantedText,
                            wantedClass=wantedClass,
                            selectionFunction=selectionFunction)
    # check for None returned:  Tim 6 Jul 2004
    if controls is None:
        raise WinGuiAutoError("EnumChildWindows failed with win32gui.error "  +
                              repr(topHwnd) +
                              ", wantedText=" +
                              repr(wantedText) +
                              ", wantedClass=" +
                              repr(wantedClass) +
                              ", selectionFunction=" +
                              repr(selectionFunction) +
                              ", maxWait=" +
                              repr(maxWait) +
                              ", retryInterval=" +
                              str(retryInterval) 
                              )
    
    if len(controls) > 1:
        raise WinGuiAutoError("Multiple controls found for topHwnd=" +
                              repr(topHwnd) +
                              ", wantedText=" +
                              repr(wantedText) +
                              ", wantedClass=" +
                              repr(wantedClass) +
                              ", selectionFunction=" +
                              repr(selectionFunction) +
                              ", maxWait=" +
                              repr(maxWait) +
                              ", retryInterval=" +
                              str(retryInterval))
    elif controls:
        return controls[0]
    if (maxWait-retryInterval) >= 0:
        time.sleep(retryInterval)
        try:
            return findControl(topHwnd=topHwnd,
                               wantedText=wantedText,
                               wantedClass=wantedClass,
                               selectionFunction=selectionFunction,
                               maxWait=maxWait-retryInterval,
                               retryInterval=retryInterval)
        except WinGuiAutoError:
            raise WinGuiAutoError("No control found for topHwnd=" +
                                  repr(topHwnd) +
                                  ", wantedText=" +
                                  repr(wantedText) +
                                  ", wantedClass=" +
                                  repr(wantedClass) +
                                  ", selectionFunction=" +
                                  repr(selectionFunction) +
                                  ", maxWait=" +
                                  repr(maxWait) +
                                  ", retryInterval=" +
                                  str(retryInterval))
        
    raise WinGuiAutoError("No control found for topHwnd=" +
                              repr(topHwnd) +
                              ", wantedText=" +
                              repr(wantedText) +
                              ", wantedClass=" +
                              repr(wantedClass) +
                              ", selectionFunction=" +
                              repr(selectionFunction) +
                              ", maxWait=" +
                              repr(maxWait) +
                              ", retryInterval=" +
                              str(retryInterval))

def findControls(topHwnd,
                 wantedText=None,
                 wantedClass=None,
                 selectionFunction=None):
    '''Find controls.
    
    You can identify controls using captions, classes, a custom selection
    function, or any combination of these. (Multiple selection criteria are
    ANDed. If this isn't what's wanted, use a selection function.)

    Arguments:
    topHwnd             The window handle of the top level window in which the
                        required controls reside.
    wantedText          Text which the required controls' captions must contain.
    wantedClass         Class to which the required controls must belong.
    selectionFunction   Control selection function. Reference to a function
                        should be passed here. The function should take hwnd as
                        an argument, and should return True when passed the
                        hwnd of a desired control.

    Returns:            The window handles of the controls matching the
                        supplied selection criteria.    

    Usage example:      optDialog = findTopWindow(wantedText="Options")
                        def findButtons(hwnd, windowText, windowClass):
                            return windowClass == "Button"
                        buttons = findControl(optDialog, wantedText="Button")
    '''
    if wantedText:
        wantedText = _normaliseText(wantedText)
    def searchChildWindows(currentHwnd):
        results = []
        childWindows = []
        try:
            win32gui.EnumChildWindows(currentHwnd,
                                      _windowEnumerationHandler,
                                      childWindows)
        except win32gui.error:
            # This seems to mean that the control *cannot* have child windows,
            # i.e. is not a container.
            return
        for childHwnd, windowText, windowClass in childWindows:
            descendantMatchingHwnds = searchChildWindows(childHwnd)
            if descendantMatchingHwnds:
                results += descendantMatchingHwnds

            if wantedText:
                if not windowText:
                    continue
                #print('wanted: %s, window: %s'% (wantedText, windowText)
                if _normaliseText(windowText).find(wantedText) == -1:
                    continue
            if wantedClass and \
               not windowClass.find(wantedClass) == 0:
                continue

            if selectionFunction and \
               not selectionFunction(childHwnd, windowText, windowClass):
                continue
            results.append(childHwnd)
        return results

    # deduplicate the returned windows:  Tim 6 Jul 2004
    #return searchChildWindows(topHwnd)
    
    hlist=searchChildWindows(topHwnd)
    
    if hlist:
        ## remove duplicates
        return list(set(hlist))
    else:
        return hlist
    
def findAdditionalControls(controls,
                 wantedText=None,
                 wantedClass=None,
                 selectionFunction=None):
    '''Find controls, from a list of already found controls
    
    You can identify controls using captions, classes, a custom selection
    function, or any combination of these. (Multiple selection criteria are
    ANDed. If this isn't what's wanted, use a selection function.)

    Arguments:
    topHwnd             The window handle of the top level window in which the
                        required controls reside.
    wantedText          Text which the required controls' captions must contain.
    wantedClass         Class to which the required controls must belong.
    selectionFunction   Control selection function. Reference to a function
                        should be passed here. The function should take hwnd as
                        an argument, and should return True when passed the
                        hwnd of a desired control.

    Returns:            The window handles of the controls matching the
                        supplied selection criteria.    

    Usage example:      optDialog = findTopWindow(wantedText="Options")
                        def findButtons(hwnd, windowText, windowClass):
                            return windowClass == "Button"
                        buttons = findControl(optDialog, wantedText="Button")
    '''
    if wantedText:
        wantedText = _normaliseText(wantedText)
    def searchChildWindows(currentHwnd):
        results = []
        childWindows = []
        try:
            win32gui.EnumChildWindows(currentHwnd,
                                      _windowEnumerationHandler,
                                      childWindows)
        except win32gui.error:
            # This seems to mean that the control *cannot* have child windows,
            # i.e. is not a container.
            return
        for childHwnd, windowText, windowClass in childWindows:
            descendantMatchingHwnds = searchChildWindows(childHwnd)
            if descendantMatchingHwnds:
                results += descendantMatchingHwnds

            if wantedText:
                if not windowText:
                    continue
                #print('wanted: %s, window: %s'% (wantedText, windowText)
                if _normaliseText(windowText).find(wantedText) == -1:
                    continue
            if wantedClass and \
               not windowClass.find(wantedClass) == 0:
                continue

            if selectionFunction and \
               not selectionFunction(childHwnd, windowText, windowClass):
                continue
            results.append(childHwnd)
        return results
    

    # deduplicate the returned windows:  Tim 6 Jul 2004
    #return searchChildWindows(topHwnd)
    resultsList = []
    for control in controls:
        hlist=searchChildWindows(control)
    
        if hlist:
            # deduplicate the list:
            hdict={}
            for h in hlist:
                hdict[h]=''
            resultsList.append(hdict.keys())

            
def selFuncExplorerAddress(_hndle, text, Class):
    """needed for quick access to explorer CabinetWClass controls
    
    Adres (Dutch) or Address (English) and other languages ???
    """
    if Class == 'ToolbarWindow32':
        if text and text.startswith('Adres: ') or text.startswith('Address: '):
            return True
    return None

## zie ook _folders:

def getwindowtext(hwnd):
    """unicode version of getwindowstext,
    need ctypes, see top of file
    """
    length = GetWindowTextLength(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    GetWindowText(hwnd, buff, length + 1)
    return buff.value
  
def getFolderFromCabinetWClass(hndle):
    """extract folder from a CabinetWClass window
    
    (Unimacro grammar _folders)
    """
    controls = findControls(hndle, selectionFunction=selFuncExplorerAddress)
    if controls:
        hndle = controls[0]
        text = getwindowtext(hndle)
        ## result is probably u"Address: C:\Documents" or so
        folder = extractFolderFromWindowText(text)
        return folder
    return None

def getFolderFromDialog(hndle, className):
    """dialog for open etc from office etc
    if #32770, but can be other dialogs, like the messages window itself!
    
    """
    if className == '#32770':
        controls = list(findControls(hndle, selectionFunction=selFuncExplorerAddress))
        if controls:
            hndle = controls[0]
            text = getwindowtext(hndle)
            folder = extractFolderFromWindowText(text)
            return folder
    return None
        
def extractFolderFromWindowText(text):
    """get folder info from CabinetWClass or #32770 window:
    
    if "location:" is found, get that info, urlparse unquote(d)
    if otherwise : is found (Adress: or Adres: (enx or nld)) return the info after that :
    
    no check of valid directory is done
    """
    if text.find("location:") >= 0:
        # ms-search:
        folder = text.split("location:", 1)[1]
        folder = unquote(folder)
    elif text.find(":") >= 0:
        # Adress: or Adres:
        folder = text.split(": ", 1)[1]
    else:
        # no activefolder info:
        return None
    return folder
    
def getTopMenu(hWnd):
    '''Get a window's main, top level menu.
    
    Arguments:
    hWnd            The window handle of the top level window for which the top
                    level menu is required.

    Returns:        The menu handle of the window's main, top level menu.

    Usage example:  hMenu = getTopMenu(hWnd)
    '''
    return ctypes.windll.user32.GetMenu(ctypes.c_long(hWnd))

def activateMenuItem(hWnd, menuItemPath):
    '''Activate a menu item
    
    Arguments:
    hWnd                The window handle of the top level window whose menu you 
                        wish to activate.
    menuItemPath        The path to the required menu item. This should be a
                        sequence specifying the path through the menu to the
                        required item. Each item in this path can be specified
                        either as an index, or as a menu name.
                    
    Raises:
    WinGuiAutoError     When the requested menu option isn't found.

    Usage example:      activateMenuItem(notepadWindow, ('file', 'open'))
    
                        Which is exactly equivalent to...
                    
                        activateMenuItem(notepadWindow, (0, 1))
    '''
    # pylint: disable=W0104
    # By Axel Kowald (kowald@molgen.mpg.de)
    # Modified by S Brunning to accept strings in addition to indicies.

    # Top level menu    
    hMenu = getTopMenu(hWnd)
    
    # Get top level menu's item count. Is there a better way to do this?
    # Yes .. Tim Couper 22 Jul 2004
    
    hMenuItemCount=win32gui.GetMenuItemCount(hMenu)
    
##    for hMenuItemCount in xrange(256):
##        try:
##            _getMenuInfo(hMenu, hMenuItemCount)
##        except WinGuiAutoError:
##            break
##    hMenuItemCount -= 1
    
    # Walk down submenus
    for submenu in menuItemPath[:-1]:
        try: # submenu is an index
            0 + submenu
            submenuInfo = _getMenuInfo(hMenu, submenu)
            hMenu, hMenuItemCount = submenuInfo.submenu, submenuInfo.itemCount
        except TypeError: # Hopefully, submenu is a menu name
            try:
                _dump, hMenu, hMenuItemCount = _findNamedSubmenu(hMenu,
                                                                hMenuItemCount,
                                                                submenu)
            except WinGuiAutoError:
                raise WinGuiAutoError("Menu path " +
                                      repr(menuItemPath) +
                                      " cannot be found.")
           
    # Get required menu item's ID. (the one at the end).
    menuItem = menuItemPath[-1]
    try: # menuItem is an index
        0 + menuItem
        menuItemID = ctypes.windll.user32.GetMenuItemID(hMenu,
                                                        menuItem)
    except TypeError: # Hopefully, menuItem is a menu name
        try:
            subMenuIndex, _dump1, _dump2 = _findNamedSubmenu(hMenu,
                                        hMenuItemCount,
                                        menuItem)
        except WinGuiAutoError:
            raise WinGuiAutoError("Menu path " +
                                  repr(menuItemPath) +
                                  " cannot be found.")
        menuItemID = ctypes.windll.user32.GetMenuItemID(hMenu, subMenuIndex)

    # Activate    
    win32gui.PostMessage(hWnd, win32con.WM_COMMAND, menuItemID, 0)
    time.sleep(0.05)
##def findMenuItems(hWnd,wantedText):
##    """Finds menu items whose captions contain the text"""
##    hMenu = getTopMenu(hWnd)
##    hMenuItemCount=win32gui.GetMenuItemCount(hMenu)
##    
##    for topItem in xrange(hMenuItemCount):
##            
    
def getMenuInfo(hWnd, menuItemPath):
    '''TODO'''
    # pylint: disable=W0104
    # Top level menu    
    hMenu = getTopMenu(hWnd)
        
    # Get top level menu's item count. Is there a better way to do this?
    # Yes .. Tim Couper 22 Jul 2004
    hMenuItemCount=win32gui.GetMenuItemCount(hMenu)
    
##    for hMenuItemCount in xrange(256):
##        try:
##            _getMenuInfo(hMenu, hMenuItemCount)
##        except WinGuiAutoError:
##            break
##    hMenuItemCount -= 1
    submenuInfo=None 
    
    # Walk down submenus
    for submenu in menuItemPath:
        try: # submenu is an index
            0 + submenu
            submenuInfo = _getMenuInfo(hMenu, submenu)
            hMenu, hMenuItemCount = submenuInfo.submenu, submenuInfo.itemCount
        except TypeError: # Hopefully, submenu is a menu name
            try:
                submenuIndex, new_hMenu, hMenuItemCount = _findNamedSubmenu(hMenu,
                                                                hMenuItemCount,
                                                                submenu)
                submenuInfo = _getMenuInfo(hMenu, submenuIndex)
                hMenu = new_hMenu
            except WinGuiAutoError:
                raise WinGuiAutoError("Menu path " +
                                      repr(menuItemPath) +
                                      " cannot be found.")
    if submenuInfo is None:
        raise WinGuiAutoError("Menu path " +
                              repr(menuItemPath) +
                              " cannot be found. (Null menu path?)")
        
    
    return submenuInfo
    
def _getMenuInfo(hMenu, uIDItem):
    '''Get various info about a menu item.
    
    Arguments:
    hMenu               The menu in which the item is to be found.
    uIDItem             The item's index

    Returns:            Menu item information object. This object is basically
                        a 'bunch'
                        (see http://aspn.activestate.com/ASPN/Cookbook/Python/Recipe/52308).
                        It will have useful attributes: name, itemCount,
                        submenu, isChecked, isDisabled, isGreyed, and
                        isSeparator
                    
    Raises:
    WinGuiAutoError     When the requested menu option isn't found.       

    Usage example:      submenuInfo = _getMenuInfo(hMenu, submenu)
                        hMenu, hMenuItemCount = submenuInfo.submenu, submenuInfo.itemCount
    '''
    # An object to hold the menu info
    class MenuInfo(Bunch):
        pass
    menuInfo = MenuInfo()

    # Menu state    
    menuState = ctypes.windll.user32.GetMenuState(hMenu,
                                                  uIDItem,
                                                  win32con.MF_BYPOSITION)
    if menuState == -1:
        raise WinGuiAutoError("No such menu item, hMenu=" +
                              str(hMenu) +
                              " uIDItem=" +
                              str(uIDItem))
    
    # Menu name
    menuName = ctypes.c_buffer("\000" * 32)
    ctypes.windll.user32.GetMenuStringA(ctypes.c_int(hMenu),
                                        ctypes.c_int(uIDItem),
                                        menuName, ctypes.c_int(len(menuName)),
                                        win32con.MF_BYPOSITION)
    menuInfo.name = menuName.value

    # Sub menu info
    menuInfo.itemCount = menuState >> 8
    if bool(menuState & win32con.MF_POPUP):
        menuInfo.submenu = ctypes.windll.user32.GetSubMenu(hMenu, uIDItem)
    else:
        menuInfo.submenu = None
        
    menuInfo.isChecked = bool(menuState & win32con.MF_CHECKED)
    menuInfo.isDisabled = bool(menuState & win32con.MF_DISABLED)
    menuInfo.isGreyed = bool(menuState & win32con.MF_GRAYED)
    menuInfo.isSeparator = bool(menuState & win32con.MF_SEPARATOR)
    # ... there are more, but these are the ones I'm interested in thus far.
    
    return menuInfo

def clickButton(hwnd):
    '''Simulates a single mouse click on a button

    Arguments:
    hwnd    Window handle of the required button.

    Usage example:  okButton = findControl(fontDialog,
                                           wantedClass="Button",
                                           wantedText="OK")
                    clickButton(okButton)
    '''
    _sendNotifyMessage(hwnd, win32con.BN_CLICKED)

def clickStatic(hwnd):
    '''Simulates a single mouse click on a static

    Arguments:
    hwnd    Window handle of the required static.

    Usage example:  TODO
    '''
    _sendNotifyMessage(hwnd, win32con.STN_CLICKED)

def doubleClickStatic(hwnd):
    '''Simulates a double mouse click on a static

    Arguments:
    hwnd    Window handle of the required static.

    Usage example:  TODO
    '''
    _sendNotifyMessage(hwnd, win32con.STN_DBLCLK)

def getComboboxItems(hwnd):
    '''Returns the items in a combo box control.

    Arguments:
    hwnd            Window handle for the combo box.

    Returns:        Combo box items.

    Usage example:  fontCombo = findControl(fontDialog, wantedClass="ComboBox")
                    fontComboItems = getComboboxItems(fontCombo)
    '''
    
    return _getMultipleWindowValues(hwnd,
                                    getCountMessage=win32con.CB_GETCOUNT,
                                    getValueMessage=win32con.CB_GETLBTEXT)

def selectComboboxItem(hwnd, item):
    '''Selects a specified item in a Combo box control.

    Arguments:
    hwnd            Window handle of the required combo box.
    item            The reqired item. Either an index, of the text of the
                    required item.

    Usage example:  fontComboItems = getComboboxItems(fontCombo)
                    selectComboboxItem(fontCombo,
                                       random.choice(fontComboItems))
    '''
    try: # item is an index Use this to select
        0 + item
        win32gui.SendMessage(hwnd, win32con.CB_SETCURSEL, item, 0)
        _sendNotifyMessage(hwnd, win32con.CBN_SELCHANGE)
    except TypeError: # Item is a string - find the index, and use that
        items = getComboboxItems(hwnd)
        itemIndex = items.index(item)
        selectComboboxItem(hwnd, itemIndex)

def getListboxItems(hwnd):
    '''Returns the items in a list box control.

    Arguments:
    hwnd            Window handle for the list box.

    Returns:        List box items.

    Usage example:  docType = findControl(newDialog, wantedClass="ListBox")
                    typeListBox = getListboxItems(docType)
    '''
    
    return _getMultipleWindowValues(hwnd,
                                    getCountMessage=win32con.LB_GETCOUNT,
                                    getValueMessage=win32con.LB_GETTEXT)

def selectListboxItem(hwnd, item):
    '''Selects a specified item in a list box control.

    Arguments:
    hwnd            Window handle of the required list box.
    item            The reqired item. Either an index, of the text of the
                    required item.

    Usage example:  docType = findControl(newDialog, wantedClass="ListBox")
                    typeListBox = getListboxItems(docType)
                    
                    # Select a type at random
                    selectListboxItem(docType,
                                      random.randint(0, len(typeListBox)-1))
    '''
    try: # item is an index Use this to select
        0 + item
        win32gui.SendMessage(hwnd, win32con.LB_SETCURSEL, item, 0)
        _sendNotifyMessage(hwnd, win32con.LBN_SELCHANGE)
    except TypeError: # Item is a string - find the index, and use that
        items = getListboxItems(hwnd)
        itemIndex = items.index(item)
        selectListboxItem(hwnd, itemIndex)

#EM_GETFIRSTVISIBLELINE
def getFirstVisibleLine(hwnd, classname=None):
    
    if classname and classname == 'Scintilla':
        getfirstvisibleline = scintillacon.SCI_GETFIRSTVISIBLELINE
    else:
        getfirstvisibleline = win32con.EM_GETFIRSTVISIBLELINE
    
    return win32gui.SendMessage(hwnd, getfirstvisibleline, 0, 0)
                                    
def getEditText(hwnd, visible=False, classname=None):
    '''Returns the text in an edit control.

    Arguments:
    hwnd            Window handle for the edit control.

    Returns         Edit control text lines. (as a list of lines!)

    Usage example:  pprint.pprint(getEditText(editArea))
    
    Note: this is the slow version, use getWindowTextAll for faster access, but
    the end of line characters must be corrected depending on the control.
    
    '''
    if classname and classname == 'Scintilla':
        getline, getlinecount = scintillacon.SCI_GETLINE, scintillacon.SCI_GETLINECOUNT
    else:
        getline, getlinecount = win32con.EM_GETLINE, win32con.EM_GETLINECOUNT
        
    if visible:
        firstLine = getFirstVisibleLine(hwnd)
    else:
        firstLine = 0
    List =  _getMultipleWindowValues(hwnd,
                                    getCountMessage=win32con.EM_GETLINECOUNT,
                                    getValueMessage=win32con.EM_GETLINE,
                                    first=firstLine)
    return List # [l.rstrip('\n\r') for l in List]
    
def setEditText(hwnd, text, append=False, classname=None):
    '''Set an edit control's text.

::

    Arguments:
    
    hwnd            The edit control's hwnd.
    text            The text to send to the control. This can be a single
                    string, or a sequence of strings. If the latter, each will
                    be become a a separate line in the control.
    append          Should the new text be appended to the existing text?
                    Defaults to False, meaning that any existing text will be
                    replaced. If True, the new text will be appended to the end
                    of the existing text.
                    Note that the first line of the new text will be directly
                    appended to the end of the last line of the existing text.
                    If appending lines of text, you may wish to pass in an
                    empty string as the 1st element of the 'text' argument.
    
    Usage example:  print("Enter various bits of text."
                    setEditText(editArea, "Hello, again!")
                    setEditText(editArea, "You still there?")
                    setEditText(editArea, ["Here come", "two lines!"])
                    
                    print("Add some..."
                    setEditText(editArea, ["", "And a 3rd one!"], append=True)
    '''

    if type(text) in (types.StringType,):
        pass
    else:
        text = '\r\n'.join(text)
    #
    ## Ensure that text is a list        
    #
    #try:
    #    text + ''
    #    text = [text]
    #except TypeError:
    #    pass

    # Set the current selection range, depending on append flag
    # -1 (append) means clear selection and put at cursor position
    if classname and classname == "Scintilla":
        setsel, replacesel = scintillacon.SCI_SETSEL, scintillacon.SCI_REPLACESEL
    else:
        setsel, replacesel = win32con.EM_SETSEL, win32con.EM_REPLACESEL
    
    if append:
        win32gui.SendMessage(hwnd, setsel, -1, 0)
    else:
        win32gui.SendMessage(hwnd, setsel, 0, -1)
                             
    # Send the text
    win32gui.SendMessage(hwnd, replacesel, True, text)



def replaceEditText(hwnd, text, classname=None):
    '''replace selection of Edit text with text
    
    Arguments:
    hwnd            The edit control's hwnd.
    text            The text to send to the control. This can be a single
                    string, or a sequence of strings. If the latter, each will
                    be become a a separate line in the control.
    '''
    if classname and classname == "Scintilla":
        setsel, replacesel = scintillacon.SCI_SETSEL, scintillacon.SCI_REPLACESEL
    else:
        setsel, replacesel = win32con.EM_SETSEL, win32con.EM_REPLACESEL
    
    if type(text) in (types.StringType,):
        pass
    else:
        text = '\n'.join(text)
    ## Ensure that text is a list        
    #try:
    #    text + ''
    #    text = [text]
    #except TypeError:
    #    pass

    # assume the selection is set (or just end of buffer)
    # Send the text
    win32gui.SendMessage(hwnd, replacesel, True, text)

def appendEditText(hwnd, text, place="end", classname=None):
    '''replace selection of Edit text with text
    
    Arguments:
    hwnd            The edit control's hwnd.
    text            The text to send to the control. This can be a single
                    string, or a sequence of strings. If the latter, each will
                    be become a a separate line in the control.
    place:          "end" at end of buffer, can also be "start"
    '''
    if classname and classname == "Scintilla":
        replacesel = scintillacon.SCI_SETSEL, scintillacon.SCI_REPLACESEL
    else:
        replacesel = win32con.EM_REPLACESEL
    
    if type(text) in (types.StringType,):
        pass
    else:
        text = '\n'.join(text)

    if place == "end":
        current = os.linesep.join(getEditText(hwnd))
        length = len(current)
        setSelection(hwnd, length, length)
    elif place == "start":
        setSelection(hwnd, 0,0)
    else:
        raise ValueError("appendEditText, invalid place parameter: %s"% place)
        
    # Send the text
    win32gui.SendMessage(hwnd, replacesel, True, text)

def getSelection(hndle, classname=None):
    """returns the start and end position of the selection.
    """
    if classname and classname == "Scintilla":
        getsel= scintillacon.SCI_GETSEL
    else:
        getsel = win32con.EM_GETSEL
    
    selInfo = win32gui.SendMessage(hndle, getsel, 0, 0)
    st = win32gui.LOWORD(selInfo)
    en = win32gui.HIWORD(selInfo)
    return st, en

def isVisible(hndle, classname=None):
    """check if window/control is visible
    """
    if classname == 'Scintilla':
        testnum = scintillacon.SCI_GETFIRSTVISIBLELINE
    else:
        return # not know how to do...
        
    v = win32gui.SendMessage(hndle, testnum)
    return v


def setSelection(hndle, lo, hi, classname=None):

    if classname and classname == "Scintilla":
        setsel = scintillacon.SCI_SETSEL
    else:
        setsel = win32con.EM_SETSEL

    win32gui.SendMessage(hndle, setsel, lo, hi)

def getLineNumber(hndle, charPos=-1, classname=None):
    """get the number of the line containing char or the selection (1 based!)
    """
    num = win32gui.SendMessage(hndle, win32con.EM_LINEFROMCHAR, charPos, 0)
    return num + 1

def getNumberOfLines(hndle, classname=None):
    """retrieves the number of lines in the buffer
    
    the total buflen fails, goes better with getRawTextLength (but is not accurate)
    """
    if classname and classname == "Scintilla":
        getlinecount = scintillacon.SCI_GETLINECOUNT 
    else:
        getlinecount = win32con.EM_GETLINECOUNT
    
    numLines = win32gui.SendMessage(hndle, getlinecount, 0, 0)
    return numLines

# length and buffer for receiving 1 line of text. If text length exceeds the
# LINE_BUFFER_LENGTH, it is doubled.
## 'c' naar 'b'  TODOQH
LINE_BUFFER_LENGTH = 256
LINE_TEXT_BUFFER = array.array('b', b' ' * (LINE_BUFFER_LENGTH))
bufferlength  = struct.pack('i', LINE_BUFFER_LENGTH) # This is a C style int.
LINE_TEXT_BUFFER = array.array('b', bufferlength + b' ' * LINE_BUFFER_LENGTH)

def getTextLine(hndle, lineNum, classname=None):
    """get a line of text from buffer, with line number lineNum
    
    Note: the lines are dependent on the width of the control, if it scrolls,
    the number of lines increases.
    Note: lineNum is 0 based!!!
    """
    global LINE_BUFFER_LENGTH, LINE_TEXT_BUFFER
    if classname and classname == "Scintilla":
        getline= scintillacon.SCI_GETLINE
    else:
        getline = win32con.EM_GETLINE
    
    
    
    
    lenReceived =  win32gui.SendMessage(hndle, getline,
                                        lineNum, LINE_TEXT_BUFFER)
    if (LINE_BUFFER_LENGTH - lenReceived) < 10:
        LINE_BUFFER_LENGTH *= 2
        bufferlength  = struct.pack('i', LINE_BUFFER_LENGTH) # This is a C style int.
        LINE_TEXT_BUFFER = array.array('b', bufferlength + b' ' * LINE_BUFFER_LENGTH)
        print('line too small, double to %s'% LINE_BUFFER_LENGTH)
        return getTextLine(hndle, lineNum)

    return LINE_TEXT_BUFFER.tostring()[:lenReceived]

def getRawTextLength(hndle, classname=None):
    """get the one stroke text length, this is with uncorrected \r\n characters
    
    can be used to check if buffer has changed in length...
    (this call takes nearly as much time as the getWindowTextAll call)
    """
    if classname and classname == "Scintilla":
        gettextlength = scintillacon.SCI_GETTEXTLENGTH
    else:
        gettextlength = win32con.WM_GETTEXTLENGTH
    
    return win32gui.SendMessage(hndle, gettextlength, 0, 0)
#

offsetLetter = 0  # ???           
keySynonyms = dict(backspace="back",
                   enter="return")
keySynonyms['del'] = "delete"


## functions with sendkeys seem not to work. Want the hndle of a control, but this is not doable.
## the _sendkeys via ctypes seems to work better. 
def sendKey(hndle, key):
    """sendKey function from messagesfunctions, not in active use
    """
    # pylint: disable=C0321
    kdown = win32con.WM_KEYDOWN
    char = win32con.WM_CHAR
    kup = win32con.WM_KEYUP
    sm = win32gui.SendMessage
    if isinstance(key, str):
        ctrl = alt = shift = 0
        kCode = None
        if key.startswith("{") and key.endswith("}"):
            key = key.strip("{}")
            keys = key.split("+")
            for k in keys:
                if k == 'ctrl': ctrl = 1
                elif k == 'shift': shift = 1
                elif k == 'alt': alt = 1
                else:
                    k = keySynonyms.get(k, k)
                    kCode = getattr(win32con, "VK_%s"%k.upper(),None)
                    if kCode is None:
                        if len(k) == 1:
                            kCode = ord(k.upper())
                        else:
                            raise ValueError("invalid key: {%s"% key[1:])
            if ctrl:  sm(hndle, kdown, win32con.VK_CONTROL, 0)
            if shift: sm(hndle, kdown, win32con.VK_SHIFT, 0)
            if alt:   sm(hndle, kdown, win32con.VK_MENU, 0)
        
            if kCode:
                sm(hndle, kdown, kCode, 0)
                sm(hndle, kup, kCode, 0)
            if alt:   sm(hndle, kup, win32con.VK_MENU, 0)
            if shift: sm(hndle, kup, win32con.VK_SHIFT, 0)
            if ctrl:  sm(hndle, kup, win32con.VK_CONTROL, 0)
            
        else:
            for k in key:
                if ord(k) >= 32:
                    sm(hndle, char, ord(k), 0)  # from space on
                else:
                    sm(hndle, kdown, ord(k), 0) # eg \r and \n
                    sm(hndle, kup, ord(k), 0)
                    
    else:
        raise ValueError("SendKey expects a string, not: %s"% key)

def waitForForegroundWindow(className=None, nWait=50, waitingTime=0.5):
    """wait some time for foreground window to get in front"""
    if isinstance(className, str):
        className = [className]
    elif className is None:
        print('waitForForegroundWindow: no classname given')
        return None
    
    for _ in range(nWait):
        hndle = getForegroundWindow()
        if hndle:
            currentClass = win32gui.GetClassName(hndle)
            if currentClass in className:
                return hndle
        time.sleep(waitingTime)
    return None

def SendKeys(hndle, keysString):
    """send the keys to the window with hndle
    
    This function could be useful when you want to send keystrokes to
    other windows. But it fails, and should be repaired in that case.
    
    The :code:`SendMessage` trick, with windows messages, *does not work*,
    
    The :code:`SendInput` mechanism (as used in Dragonfly,
    and so in the `sendkeys.py` module of this module) is favourable.
 
    Example:
    ::
      hndle = getForegroundWindow()
      SendKeys(hndle, "{shift+a}")   ## fails
      SendKeys(hndle, "abc")         ## works
    
    """
    keysList = splitKeysString(keysString)
    for key in keysList:
        sendKey(hndle,key)
    
def splitKeysString(keysString):
    """split in single keys, either with braces around or single keystrokes
    """
    L = filter(None, hasBraces.split(keysString))
    M = []
    for item in L:
        if braceExact.match(item):
            M.append(item)
        else:
            M.extend(list(item))
    return M
    

def quitProgram(hwnd):
    """send a WM_QUIT message
    """
    #win32gui.DestroyWindow(hwnd)
    win32gui.PostMessage(hwnd,
                         win32con.WM_DESTROY,
                         0,0)

    #_sendNotifyMessage(hwnd, win32con.WM_ACTIVATE)
    #oldHndle = win32gui.GetActiveWindow()
    #win32gui.SetActiveWindow(hwnd)
    #time.sleep(1)
    #
    #win32gui.PostQuitMessage(hwnd)
    #time.sleep(1)
    #win32gui.SetActiveWindow(oldHndle)
    #

# buffer and length for one stroke get window text:
BUFFER_LENGTH = 1024
TEXT_BUFFER = array.array('b', b' ' * (BUFFER_LENGTH))
def getWindowTextAll(hwnd, rawLength=None):
    """get the buffer of a edit control in one stroke
    
    return the buffer and the raw length (uncorrected for \r\n and after text character)
    """
    # pylint: disable=W0603
    global TEXT_BUFFER, BUFFER_LENGTH
    if rawLength is None:
        rawLength = BUFFER_LENGTH 
    if rawLength > BUFFER_LENGTH:
        BUFFER_LENGTH = rawLength + 1000   # take a little extra!
        TEXT_BUFFER = array.array('b', b' ' * BUFFER_LENGTH)
    valueLength = win32gui.SendMessage(hwnd, win32con.WM_GETTEXT,
                                            BUFFER_LENGTH, TEXT_BUFFER)
    if not valueLength:
        return '', 0
    
    if (BUFFER_LENGTH - valueLength) <= 10:
        print('probably too small buffer for this control: %s'% valueLength)
        rawLength = getRawTextLength(hwnd) + 1000
        return getWindowTextAll(hwnd, rawLength=rawLength)
        
    result = TEXT_BUFFER.tostring()[:valueLength]
    return result, valueLength

def _getMultipleWindowValues(hwnd, getCountMessage, getValueMessage, first=0):
    '''
    
    A common pattern in the Win32 API is that in order to retrieve a
    series of values, you use one message to get a count of available
    items, and another to retrieve them. This internal utility function
    performs the common processing for this pattern.

    Arguments:
    hwnd                Window handle for the window for which items should be
                        retrieved.
    getCountMessage     Item count message.
    getValueMessage     Value retrieval message.

    Returns:            Retrieved items.
    '''
    result = []
    
    MAX_VALUE_LENGTH = 256
    bufferlength  = struct.pack('i', MAX_VALUE_LENGTH) # This is a C style int.
    valuecount = win32gui.SendMessage(hwnd, getCountMessage, 0, 0)
    for itemIndex in range(first, valuecount):
        valuebuffer = array.array('c',
                                  bufferlength +
                                  ' ' * (MAX_VALUE_LENGTH - len(bufferlength)))
        valueLength = win32gui.SendMessage(hwnd,
                                           getValueMessage,
                                           itemIndex,
                                           valuebuffer)
        result.append(valuebuffer.tostring()[:valueLength])
    return result

def _windowEnumerationHandler(hwnd, resultList):
    '''win32gui.EnumWindows() callback.
    
    Pass to win32gui.EnumWindows() or win32gui.EnumChildWindows() to
    generate a list of window handle, window text, window class tuples.
    stop enumeration, if checkI(mmediate) is set and contents match 
    '''
    if hwnd == 0:
        print('_windowEnumerationHandler, hwnd == 0, skip')
        return None
    # try:
    _testHndle = win32gui.GetWindow(hwnd, 0)
    # except pywintypes.error:
    #     if details[0] == 1400:
    #         print(f'caught 1400 error, skip this window: {hwnd}'% hwnd)
    #         return None
    #     raise
    if checkI:
        text = className = None
        if wText:
            text = win32gui.GetWindowText(hwnd)
            if wText and _normaliseText(text).find(wText) == -1:
                return
        if wClass:
            className = win32gui.GetClassName(hwnd)
            if wClass and className.find(wClass) != 0:
                return
        if not wText:
            text = win32gui.GetWindowText(hwnd)
        if not wClass:
            className = win32gui.GetClassName(hwnd)
        # got it, append and stop:
        resultList.append(  (hwnd, text, className) )
        return True
    # pass on window title and class name:
    text = win32gui.GetWindowText(hwnd)
    className = win32gui.GetClassName(hwnd)
    #visible = win32gui.SendMessage(hwnd, 2491)
    resultList.append(  (hwnd, text, className) )   
    return None
    
def _buildWinLong(high, low):
    '''Build a windows long parameter from high and low words.
    
    See http://support.microsoft.com/support/kb/articles/q189/1/70.asp
    '''
    # return ((high << 16) | low)
    return int(struct.unpack('>L',
                             struct.pack('>2H',
                                         high,
                                         low)) [0])
        
def _sendNotifyMessage(hwnd, notifyMessage):
    '''Send a notify message to a control.'''
    win32gui.SendMessage(win32gui.GetParent(hwnd),
                         win32con.WM_COMMAND,
                         _buildWinLong(notifyMessage,
                                       win32api.GetWindowLong(hwnd,
                                                              win32con.GWL_ID)),
                         hwnd)

def _normaliseText(controlText):
    '''Remove '&' characters, but keep the case
    
    leave None or '' unchanged
    
    Useful for matching control text.'''
    if controlText and controlText.find("&"):
        return controlText.replace('&', '')
    return controlText

def _findNamedSubmenu(hMenu, hMenuItemCount, submenuName):
    '''Find the index number of a menu's submenu with a specific name.'''
    for submenuIndex in range(hMenuItemCount):
        submenuInfo = _getMenuInfo(hMenu, submenuIndex)
        # submenuInfoName = submenuInfo.name
        if _normaliseText(submenuInfo.name).startswith(_normaliseText(submenuName)):
            return submenuIndex, submenuInfo.submenu, submenuInfo.itemCount
    raise WinGuiAutoError("No submenu found for hMenu=" +
                          repr(hMenu) +
                          ", hMenuItemCount=" +
                          repr(hMenuItemCount) +
                          ", submenuName=" +
                          repr(submenuName))

def _dedup(thelist):
    '''De-duplicate deeply nested list.'''
    found=[]
    def dodedup(thelist):
        for index, thing in enumerate(thelist):
            if isinstance(thing, list):
                dodedup(thing)
            else:
                if thing in found:
                    del thelist[index]
                else:
                    found.append(thing)
                
                              
class Bunch:
    # pylint: disable=R0903
    def __init__(self, **kwds):
        self.__dict__.update(kwds)
        
    def __str__(self):
        state = ["%s=%r" % (attribute, value)
                 for (attribute, value)
                 in self.__dict__.items()]
        return '\n'.join(state)
    
class WinGuiAutoError(Exception):
    pass

def self_test():
    #print("Let's see what top windows we have at the 'mo:")
    #pprint.pprint(dumpTopWindows())
    #x = input('->')
    print("Open and locate Notepad")
    os.startfile('notepad')
    notepadWindow = findTopWindow(wantedClass='Notepad')
    _x = input('->')
    print("Locate the 'find' edit box")
    findValue = findControls(notepadWindow, wantedClass="Edit")[0]
    _x = input('->')                               
    print("Enter some text - and wait long enough for it to be seen")
    setEditText(findValue, "Hello, mate!")
    time.sleep(.5)
    _x = input('->')
    print("Locate notepad's edit area, and enter various bits of text.")
    editArea = findControl(notepadWindow,wantedClass="Edit")
    setEditText(editArea, "Hello, again!")
    time.sleep(.5)
    setEditText(editArea, "You still there?")
    time.sleep(.5)
    setEditText(editArea, ["Here come", "two lines!"])
    time.sleep(.5)
    _x = input('->')
    
    print("Add some...")
    setEditText(editArea, ["", "And a 3rd one!"], append=True)
    time.sleep(.5)
    
    print("See what's there now:")
    pprint.pprint(getEditText(editArea))
    _x = input('->')
    
    print("Exit notepad")
    activateMenuItem(notepadWindow, ('bestand', 'afsluiten'))  # was 'file', 'exit'
    time.sleep(.5)
    
    print("Don't save.")
    saveDialog = findTopWindow(selectionFunction=lambda hwnd:
                                                 win32gui.GetWindowText(hwnd)=='Notepad')
    noButton = findControl(saveDialog,wantedClass="Button", wantedText="nee")  # was 'no'
    clickButton(noButton)
    _x = input('->')

def test_with_dragonpad1(setText=1):
    dragonpadWindow = findTopWindow(wantedClass='TalkpadClass')
    print('dragonpadWindow: %s'% dragonpadWindow)
    findValues = findControls(dragonpadWindow, wantedClass="RichEdit20A")
    if not findValues:
        return None
    findValue = findValues[0]
    print('findValue: %s'% findValue)
    if setText:
        setEditText(findValue, "Hello, mate!")
    return findValue

def test_with_notepad1(setText=1):
    notepadWindow = findTopWindow(wantedClass='Notepad')
    print('notepadWindow: %s'% notepadWindow)
    if not notepadWindow:
        return None
    findValue = findControls(notepadWindow, wantedClass="Edit")[0]
    print('findValue: %s'% findValue)
    if setText:
        setEditText(findValue, "Hello, mate!")
    return findValue
    # do "switch to notepad" and dictate "this is a test"
    # result:
    # Hello, mate! Hello, this is another test

def test_with_wordpad1():
    wordpadWindow = findTopWindow(wantedClass='WordPadClass')
    print('wordpadWindow: %s'% wordpadWindow)
    findValue = findControls(wordpadWindow, wantedClass="RICHEDIT50W")[0]
    print('findValue: %s'% findValue)
    if not findValue:
        return None
    if setText:
        setEditText(findValue, "Hello, mate!")
    return findValue
    # do "switch to wordpad" and dictate "this is a test"
    # result:
    # Hello, mate! Hello, this is a test

def test_with_excel():
    excelWindows = findTopWindows(wantedClass='XLMAIN')
    for xl in excelWindows:
        print("xl: %s"% xl)
        contents = dumpWindow(xl)
        pprint.pprint(contents)
        
    return excelWindows

def test_with_pythonwin(topHwnd):
    """pass top handle, use selection function from windowsparameter.py
    """
    import windowparameters
    selFunc = windowparameters.getPythonwinEditControl
    controls = findControls(topHwnd, selectionFunction=selFunc)
    print('controls pythonwin: %s'% controls)

    contents = dumpWindow(topHwnd)
    pprint.pprint(contents)
    return controls

def test_with_komodo_messages():
    """also TCP model could be set up, this is the try via scintilla window and messages
    
    discussion on: http://community.activestate.com/node/10359
    
    """
    komodoWindows = findTopWindows(wantedClass='MozillaWindowClass')
    for ko in komodoWindows:
        contents = dumpWindow(ko)
        if contents:
            print("ko: %s"% ko)
            pprint.pprint(contents)
        
    # do "switch to wordpad" and dictate "this is a test"
    # result:
    # Hello, mate! Hello, this is a test

def doHotshotTest(editControl, nTests=6):
    """do some testing on control.
    
    With each higher value of nTests the amount of text is doubled (in the last set/get calls)
    """
    import hotshot, hotshot.stats
    filePath = r'D:\messagesfunctions.prof'
    prof = hotshot.Profile(filePath)
    prof.runcall(hotshot_test, editControl, nTests)
    prof.close()
    stats = hotshot.stats.load(filePath)
    stats.strip_dirs()
    stats.sort_stats('time', 'calls')
    stats.print_stats(20)

def doHotshotTestField(editControl):
    """do some testing on control.
    
    With each higher value of nTests the amount of text is doubled (in the last set/get calls)
    """
    import hotshot, hotshot.stats
    filePath = r'D:\messagesfunctions.prof'
    prof = hotshot.Profile(filePath)
    prof.runcall(hotshot_test_selection, editControl, 21, 23)
    prof.close()
    stats = hotshot.stats.load(filePath)
    stats.strip_dirs()
    stats.sort_stats('time', 'calls')
    stats.print_stats(20)

def hotshot_test_selection(editArea, selStart, selEnd):
    """do some actions in hotshot profile mode, when a field is selected
    """
    setSelection(editArea, selStart, selEnd)
    time.sleep(1)
    _t = getEditText(editArea)
    time.sleep(1)
    print('selection at: %s, %s'% (selStart, selEnd))

def doHotshotTestOutsideField(editControl):
    """do some testing on control.
    
    With each higher value of nTests the amount of text is doubled (in the last set/get calls)
    """
    import hotshot, hotshot.stats
    filePath = r'D:\messagesfunctionsoutside.prof'
    prof = hotshot.Profile(filePath)
    prof.runcall(hotshot_test_selection, editControl, 10, 20)
    prof.close()
    stats = hotshot.stats.load(filePath)
    stats.strip_dirs()
    stats.sort_stats('time', 'calls')
    stats.print_stats(20)

    
def hotshot_test(editArea, nTests):
    """do some actions in hotshot profile mode
    """
    setEditText(editArea, "")
    setEditText(editArea, "Hello, again!")
    setEditText(editArea, "You still there?")
    #setEditText(editArea, ["Here come", "two lines! The second one is not too short!"])
    #setEditText(editArea, ["", "And a 3rd one! This one is considerable longer, because there must be something to test. I hope this testing gives some reliable result.", ""], append=True)
    setEditText(editArea, ["Here come", "two lines! The second one is not too short!"])
    setEditText(editArea, ["", "", "And a 3rd one! This one is considerable longer, because there must be something to test. I hope this testing gives some reliable result.", "", ""], append=True)


    for _i in range(nTests):
        t = getEditText(editArea)
        setEditText(editArea, t, append=True)
    result = getEditText(editArea)
    lines = len(result)
    lenresult = len(''.join(result))
    print('***number of tests: %s, lines: %s, characters: %s  ***'% (nTests, lines, lenresult))

def findPythonwinControl(topHwnd):
    """pass the top handle"""
    windowTitle = win32gui.GetWindowText(topHwnd)
    wantedText = windowTitle.split("[")[-1][:-1]
    result = []
    controls1 = findControls(topHwnd,wantedText=wantedText)
    if controls1:
        result.append(findAdditionalControls(controls1, wantedClass="Scintilla"))
    print('pythonwin: %s'% result)
    


def findDragonBar():
    db = findTopWindow(wantedClass='DgnBarMainWindowCls')
    
def testExplorerViews(hndle):
    # PostMessage, 0x111, 28713,0,, ahk_pid %Win_pid% medium icons, AHK
    # PostMessage = ctypes.windll.user32.PostMessageA
    win32gui.PostMessage(hndle, 28713, 0, 0)
    win32gui.PostMessage(hndle, 28716, 0, 0)
    
        
if __name__ == '__main__':
    #findDragonBar()
    
    #self_test()
    #editArea = test_with_aligen1(setText=0)
    #editArea = test_with_dragonpad1(setText=0)
    #editArea = test_with_notepad1()
    #editArea = test_with_wordpad1()
    #if not editArea:
    #    print('did not find the correct control, ensure your application is on!!'
    #else:
    #    #doHotshotTest(editArea, 6)
    #    doHotshotTestField(editArea)
    #    doHotshotTestOutsideField(editArea)
    
    #getFolderFromCabinetWClass(1180166)
    #test_with_excel()
    #test_with_komodo_messages()
    #findPythonwinControl(133488)  # uses selectionFunction getPythonwinEditControl
    #test_with_pythonwin(133488)
    #pprint.pprint(dumpXplorer2())
    
    # view settings explorer (this one fails!)
    # testExplorerViews(986418)
    # g = getFolderFromCabinetWClass(1837622)
    # g = getFolderFromDialog(394720, '#32770')
    # print(g, type(g))
    # hndle = getForegroundWindow()
    # SendKeys(hndle, "{shift+a}")   ## fails
    # SendKeys(hndle, "abc")         ## works
    pass