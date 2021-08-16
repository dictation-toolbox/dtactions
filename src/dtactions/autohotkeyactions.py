"""unimacro natlink macro wrapper/extensions
(c) copyright 2003 Quintijn Hoogenboom (quintijn@users.sourceforge.net)

 autohotkeyactions.py 
 make it independent from Natlink (Dragon), but it is windows specific

 written by: Quintijn Hoogenboom (QH softwaretraining & advies)
 December 2013/.../March 2021, adding AutoHotkey support

This module contains actions via AutoHotkey

(see autohotkey.com)

1. the ahkexe and ahkscriptfolder are checked for at the bottom of the import procedure.
   Previous scripts are copied from the directory dtactions/samples/autohotkey

2. ahk_is_active returns whether AutoHotkey is running on your system. 

3. call_ahk_script_path calls the script you have defined.

4. If you have the script in a string, call call_ahk_script.
    The script is copied to tempscript.ahk in the ahkscriptfolder and executed with call_ahk_script_path
    
5. When scripts want information back into python, write this information into a file eg "INFOfromAHK.txt".
   Scripts that use this feature, should have a function in this module, in order to read the data from this file.
   See getProgInfo and getCurrentModule below.

timing results:
  1    0.000    0.000    0.401    0.401 autohotkeyactions.py:456(killWindow)
  4    0.000    0.000    0.261    0.065 autohotkeyactions.py:91(getProgInfo)
  1    0.000    0.000    0.234    0.234 autohotkeyactions.py:209(ahkBringup)
  1    0.000    0.000    0.052    0.052 autohotkeyactions.py:323(GetForegroundWindow)
  
"""
import subprocess
import shutil
import filecmp
from pathlib import Path
import collections
import os.path  # only for getting the ahk exe path...
import time
from textwrap import dedent
from dtactions.sendkeys import sendkeys
## get thisDir and other basics...
try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print('Run this module after "build_package" and "flit install --symlink"\n')
    raise


dtactions = thisDir = getThisDir(__file__)
sampleAhkDirectory = dtactions/'samples'/'autohotkey'
checkDirectory(sampleAhkDirectory)

ahkexe = None
ahkscriptfolder = None

def ahk_is_active():
    """return True if autohotkey is active
    """
    return bool(ahkexe) and bool(ahkscriptfolder)

def do_ahk_script(script, filename=None):
    """try autohotkey integration
    """
    filename = filename or 'tempscript.ahk'
    if not ahk_is_active():
        print('ahk is not active, cannot run script')
        return
    #print 'AHK with script: %s'% script
    scriptPath = ahkscriptfolder/filename
    if isinstance(script, (list, tuple)):
        script = '\n'.join(script)         
    if not script.endswith('\n'):
        script += '\n'
    
    with open(scriptPath, 'w') as fp:
        fp.write(script)
    call_ahk_script_path(scriptPath)

def call_ahk_script_path(scriptPath):
    """call the specified ahk script  
    
    you can call .ahk scripts or .exe compiled files,
    but .exe files seem not to give better performance...
    
    use the global variable ahkexe as executable
    
    """
    scriptPath = str(scriptPath)
    if scriptPath.endswith('.ahk'):
        result = subprocess.call([ahkexe, scriptPath, ""])
    elif scriptPath.endswith('exe'):
        result = subprocess.call([scriptPath, "", ""])
    else:
        raise ValueError(f'autohotkeyactions, call_ahk_script_path: path should end with ".ahk", or ".exe"\n    path: {scriptPath}')
    if result:
        print('non-zero result of call_ahk_script_path "%s": %s'% (scriptPath, result))

ProgInfo = collections.namedtuple('ProgInfo', 'progpath prog title toporchild classname hndle'.split(' '))

def getProgInfo():
    """get the prog info, like natlink.getCurrentModule enhanced with toporchild and classname
    
    returns program info as namedtuple (progpath, prog, title, toporchild, classname, hndle)

    So the length of progInfo is 6!

    toporchild 'top' or 'child', or '' if no valid window
          
    """
    ErrorFile = clearErrorMessagesFile()
    ProgInfoFile = clearProgInfoFile()
    
    script = getProgInfoScript(info_file=ProgInfoFile)
    do_ahk_script(script)
    
    return getProgInfoResult(info_file=ProgInfoFile)

GetProgInfo = getProgInfo

def getProgInfoScript(info_file):
    """return the ahk script for getting the progInfo
    """
    script = """
; put prog info of current window in file.
;(progpath, prog, title, toporchild, classname, hndle)

WinGet pPath, ProcessPath, A
WinGetTitle, Title, A
WinGetClass, Class, A
WinGet wHndle, ID, A
wHndle := wHndle + 0
toporchild := "top"
if DllCall("GetParent", UInt, WinExist("A")) {
    toporchild := "child"    
}

FileDelete, ##proginfofile##
FileAppend, %pPath%`n, ##proginfofile##
FileAppend, %Title%`n, ##proginfofile##
FileAppend, %toporchild%`n, ##proginfofile##
FileAppend, %Class%`n, ##proginfofile##
FileAppend, %wHndle%, ##proginfofile##
"""
    script = script.replace('##proginfofile##', str(info_file))
    return script

def getProgInfoResult(info_file):
    """extract the contents of the info_file, and return the progInfo
    """
    with open(info_file, 'r') as fp:
        progInfo = fp.read().split('\n')
        
    # note ahk returns 5 lines, but ProgInfo has 6 items.   
    if len(progInfo) == 5:
        # make hndle decimal number:
        progpath, title, toporchild, classname, hndle = progInfo
        if hndle:
            hndle = int(hndle)
        else:
            hndle = 0
        prog = Path(progpath).stem
        return ProgInfo(progpath, prog, title, toporchild, classname, hndle)
    raise ValueError(f'ahk script getProgInfo did not return correct result:\n    length {len(progInfo)}\n    text: {repr(progInfo)}')


#def call_ahk_script_text(scriptText):
#    """call the specified ahk script as a text string
#    
#    use the com module, does not work (yet)
#    
#    """
#    script = scriptText
#    if not script.strip().endswith("ExitApp"):
#        script += "\nExitApp"
#    print 'do script via dll: %s'% script
#    ahk = win32com.client.Dispatch("AutoHotkey.Script")
#        
#    ahk.ahktextdll(script)
 
def copySampleAhkScripts(fromFolder, toFolder):
    """copy (in new Autohotkey directory) the sample script files
    """
    if not fromFolder.is_dir():
        print(f'No sample_ahk dir found (should be in dtactions/sample): {fromFolder}')
        return
    for f in fromFolder.glob("*.ahk"):
        inputFile = f
        inputName = f.name
        # stem = f.stem
        outputFile = toFolder/inputName
        if not outputFile.is_file():
            print(f'---copy AutoHotkey script "{inputName}" from\nSamples directory to "{toFolder}"----')
            shutil.copyfile(inputFile, outputFile)
        elif getFileDate(inputFile) > getFileDate(outputFile):
            if not filecmp.cmp(inputFile, outputFile): # inputFile (samples), newer and changed
            #     for i in range(1000):
            #         outputStem = f'{stem}_{i:03d}'
            #         newOutputFile = outputFile.with_name(newOutputName)
            #         if not newOutputFile.is_file():
            #             break
            #     else: 
            #         raise OSError('no unused newOutputFile available last: {newOutputFile}')
            #     print(f'AutoHotkey script "{inputName}" has been changed in "sample_ahk"\n   copy to "{toFolder}"\n   keep backup in {newOutputFile}')
            #     shutil.copyfile(outputFile, newOutputFile)
                shutil.copyfile(inputFile, outputFile)
        
def ahkBringup(app, filepath=None, title=None, extra=None, waitForStart=1):
    """start a program, folder, file, with AutoHotkey, or bring to foreground
    
    This functions is related to UnimacroBringup, which works with AppBringup from Dragon,
    but sometimes fails.
    
    Besides, this function can also work without Dragon being on, not relying on natlink.
    So better fit for debugging purposes.
    
    app:               can be without .exe (so notepad or notepad.exe), or the complete path
                       of the app. Advised when you want to activate files in a specific app,
                       that is not the default app for the filetype.
    filepath:          optional, the file to open.
    title:             activate title if the app already exists
    extra:             do extra actions after finding the correct app in the foreground,
                       such as opening a new document
    waitForStart:      (default: 3), wait longer for a new process to start (in seconds)
    waitForActivate:   obsolete, not needed in practice
    
    currently appBringup with app and title seems to be pretty stable, tested with thunderbird.
    other variants need testing
    
    returns: progInfo if ahkBringup succeeded
    returns: error message if ahkBringup failed

    error messages (str):
        1. when in python code, return direct
        2. when raising an error in the ahk script, put text in errormessagefromahk.txt
        
    Uses two text files for getting the result:
        1. errormessagefromahk.txt: retrieves error messages as reported from the ahk script
        2. proginfofromahk.txt: gives the progInfo of the window that has been brought up
    """
    #pylint:disable=R0913, R0911, R0912
    if not ahk_is_active():
        return 'cannot run ahkBringup, autohotkey is not active'
    ErrorFile = clearErrorMessagesFile()
    ProgInfoFile = clearProgInfoFile()
    
    title  = title or ''
    filepath = filepath or ''
    extra = extra or ''
    # if app.endswith('.exe'):
    #     basename = app[:4]
    # else:
    #     basename = app
    # app = app + '.exe'

    
    ## treat mode = open or edit, finding a app in actions.ini:
    if ((app and app.lower() == "winword") and
        (filepath and (str(filepath).endswith(".docx") or str(filepath).endswith('.doc')))):
        if Path(filepath).is_file():
            script = _get_run_winword_script(filepath)
        else:
            return f'ahkBringup winword with filepath: "{filepath}": does not exist'

    else:
        # get function and vars
        Func, Vars = _get_run_function_vars(app, filepath, title)
        if Func is None:
            return f'ahkBringup, invalid combination of input variables\n    app: "{app}"\n    filepath: "{filepath}"\n    title: "{title}"'
        script = Func(*Vars)

    if isinstance(script, str):
        script = [script]
    # script.append("WinWait, ahk_pid %NewPID%") ## this is done in the Func above
    # now grab all
    # why this one???
    # script.append("WinGet, pPath, ProcessPath, ahk_pid %NewPID%")
    if extra:
        script.append(extra)
    # always:    
    script.append(getProgInfoScript(info_file=ProgInfoFile))
    script = '\n'.join(script)
    ## do the replacements:
    script = script.replace('##waitForStart##', str(waitForStart))
    script = script.replace('##proginfofile##', str(ProgInfoFile))
    script = script.replace('##ErrorFile##', str(ErrorFile))

    ## do the script!!
    do_ahk_script(script)
    message = open(ErrorFile, 'r').read().strip()
    if message:
        return message  

    # see if there is an error message:
    errors = readErrorMessagesFile()
    if errors:
        return errors

    progInfo = getProgInfo()
    ## collect the progInfo:
    if len(progInfo) != 6:
        return 'progInfo returns tuple of invalid length (should be 6): {len(progInfo)}\n===={progInfo})'
    if not progInfo.hndle:
        return 'autohotkeyactions.appBringup, no valid progInfo returned, hndle is empty'
    if progInfo and isinstance(progInfo, str):
        ## error message:
        return progInfo
    ## OK!:
    return progInfo

autohotkeyBringup = ahkBringup

## here the internal functions for appBringup:
def _get_run_winword_script(filepath):
    """construct script than opens an existing word document
    
    filepath is the Word document to open.

    The retrieval of the progInfo is done in the appBringup function    
    """
    script = '''\
        Word := ComObjCreate("Word.Application")
        Word.Visible := True
        Word.Documents.Open("##filepath##")
        Word.Visible := 1
        Word.Activate
        if ErrorLevel {
            FileAppend, ahkBringup: could not start winword with document ##filepath##te, ##ErrorFile##
            return
        }
    '''
    script = script.replace('##filepath##', filepath)
    script = dedent(script)
    return script

def _get_run_function_vars(app, filepath, title):
    """return _app_filepath or _app_title or _app, dependent on the variables being True
 
    This is the script that is tested with thunderbird in testAhkBringupThunderbirdAppTitle

    input: values of app, filepath and title
    output: callable function or None, tuple of input variables
    
>>> _get_run_function_vars('notepad', '', '')               #doctest: +ELLIPSIS
(<function _get_run_app_script at ...>, ('notepad',))
>>> _get_run_function_vars('notepad', '', 'window title')   #doctest: +ELLIPSIS
(<function _get_run_app_title_script at ...>, ('notepad', 'window title'))

# wrong call:
>>> _get_run_function_vars('', '', '')   #doctest: +ELLIPSIS
(None, ())
    
    """
    Parts = []
    Vars = []
    if app:
        Vars.append(app)
        Parts.append('_app')
    if filepath:
        Parts.append('_filepath')
        Vars.append(filepath)
    if title:
        Parts.append('_title')
        Vars.append(title)
    func_name = f'_get_run{"".join(Parts)}_script'
    if func_name in globals():
        func = globals()[func_name]
    else:
        func = None
    return func, tuple(Vars)

def _get_run_app_title_script(app, title):
    """constructs start of script switching to window title if present otherwise open app
    
    """
    script = '''\
        SetTitleMatchMode, 2
        ; _get_run_app_title_script, app: "##app##", title: "##title##" 
        wantedHndle:= WinExist(##title##)
        WinGet, activeHndle, ID, A
        
        if (wantedHndle) {
            if (activeHndle != wantedHndle) {
                if WinExist(##title##)
                    WinActivate
            }
        }
        else {
            Run, ##app##,,,NewPID
            WinWait ahk_pid %NewPID%,,##waitForStart##
            if ErrorLevel {
                FileAppend, Could not activate app ##app##, ##ErrorFile##
                return
                    }
        }
        ; end of _get_run_app_title_script part of script
        '''
    script = script.replace('##app##', f'"{app}"')
    script = script.replace('##title##', f'"{title}"')
    script = dedent(script)
    return script

def _get_run_app_filepath_script(app, filepath):
    """constructs start of script switching running app with a filepath
    """
    script = [f'Run, "{app} {filepath}",,, NewPID']
    errorlines = '''
        if ErrorLevel {
            FileAppend, Could not activate app: "##app##" with filepath: "##filepath##", ##ErrorFile##
            return
            }
        '''
    errorlines = errorlines.replace('##app##', app)
    errorlines = errorlines.replace('##filepath##', filepath)
    errorlines = dedent(errorlines)
        
    script.append(errorlines)
    return script

def _get_run_filepath_script(filepath):
    """constructs start of script switching running a filepath
    """
    script = [f'Run, "{filepath}",,, NewPID']
    errorlines = '''
        if ErrorLevel {
            FileAppend, Could not start filepath: "##filepath##", ##ErrorFile##
            return
            }
        '''
    errorlines = errorlines.replace('##filepath##', filepath)
    errorlines = dedent(errorlines)
    script.append(errorlines)
    return script

def _get_run_app_script(app):
    """constructs start of script switching running an app
    """
    script = [f'Run, "{app}",,, NewPID']
    errorlines = '''
        if ErrorLevel {
            FileAppend, Could not start app: "##app##", ##ErrorFile##
            return
            }
        '''
    errorlines = errorlines.replace('##app##', app)
    errorlines = dedent(errorlines)
    script.append(errorlines)
    return script

def GetForegroundWindow():
    """return the hndle of the ForegroundWindow
    
    OK: return value, int, the hndle of the Foreground window
    Error: return value is str: error message explaining the failure
    """
    HndleFile = ahkscriptfolder/"foregroundhndlefromahk.txt"
    
    script = '''\
        ; put window hndle of current window in file.
        WinGet pPath, ProcessPath, A
        WinGet wHndle, ID, A
        wHndle := wHndle + 0
        FileDelete, ##hndlefile##
        FileAppend, %wHndle%, ##hndlefile##
    '''
    script = script.replace('##hndlefile##', str(HndleFile))
    do_ahk_script(script, filename="getforegroundwindow.ahk")

    with open(HndleFile, 'r') as fp:
        gotHndle = fp.read().strip()
    try:
        if gotHndle:
            hndleInt = int(gotHndle)
            return hndleInt
        mess = 'autohotkeyactions, GetForegroundWindow did not return a window hndle'
        return mess
    except ValueError:
        if gotHndle:
            mess = f'autohotkeyactions, GetForegroundWindow: did not get correct hndle window: {gotHndle}'
            # print(mess)
            return mess
        mess = 'autohotkeyactions, GetForegroundWindow: did not get correct hndle window: (empty)'
        # print(mess)
        return mess

def SetForegroundWindow(hndle):
    """bring window with hndle into the foreground
     
    return True: success
    other return: message why it went wrong
    
    """
    ErrorFile = clearErrorMessagesFile()
    ProgInfoFile = clearProgInfoFile()

    script = '''\
        WinActivate, ahk_id ##hndle##
        WinWait, ahk_pid ##hndle##,,1
        if ErrorLevel {
            FileAppend, Could not get window with hndle ##hndle## in the foreground, ##ErrorFile##
            return
        }

        WinGet wHndle, ID, A
        wHndle := wHndle + 0
        FileDelete, ##proginfofile##
        FileAppend, %wHndle%, ##proginfofile##
        '''
    script = script.replace('##hndle##', str(hndle))
    script = script.replace('##proginfofile##', str(ProgInfoFile))
    script = script.replace('##ErrorFile##', str(ErrorFile))
    do_ahk_script(dedent(script))
    
    if Path(ErrorFile).is_file():
        mess = f'Error with SetForegroundWindow to {hndle}'
        return mess

    if not Path(ProgInfoFile).is_file():
        mess = f'Error with SetForegroundWindow to {hndle}, InfoFile cannot be found'
        return mess
        
    winHndle = open(ProgInfoFile, 'r').read().strip()
    if winHndle:
        try:
            winHndle = int(winHndle)
        except ValueError:
            mess = f'ahk script getIntoForeground to "{hndle}" did not return correct winHndle: {winHndle}'
            return mess
            
        if winHndle == hndle:
            return True  # ok, then return True
        mess = f'could not switch to wanted hndle {hndle} (got {winHndle})'
        return mess
    raise ValueError('ahk script getIntoForeground did not return anything')

def getFileDate(modName):
    """return the last modified date/time of file, 0 if file does not exist
    """
    try:
        modTime = modName.stat().st_mtime
        return modTime
    except OSError:
        return 0        # file not found

def GetAhkExe():
    """try to get executable of autohotkey.exe, if not there, empty string is put in ahkexe
    
    return None, raises OSError if ahkExe is not found
    """
    # pylint: disable=W0603  # (using the global statement)
    global ahkexe
    # no succes, go on with program files:
    pf = os.path.expandvars("%PROGRAMFILES%")
    if pf.find('(x86)')>0:
        # 64 bit:
        pf = os.path.expandvars("%PROGRAMW6432%") 
    if pf and Path(pf).is_dir():
        pf = Path(pf)
    elif pf:
        raise OSError('cannot find (old style) program files directory: (empty)')
    else:
        raise OSError(f'cannot find (old style) program files directory: {pf}')
    
    ahk = pf/"autohotkey"/"autohotkey.exe"
    if ahk.is_file():
        ahkexe = ahk
        #print 'AutoHotkey found, %s'% ahkexe
    else:
        ahkexe = ""
        print(f'AutoHotkey not found on this computer {ahkexe}')
    
def GetAhkScriptFolder():
    """try to get AutoHotkey folder as subdirectory of HOME
    
    On Windows mostly C:\\Users\\Username\\.autohotkey
    
    create if non-existent.
    """
    # pylint: disable=W0603  # (using the global statement)
    global ahkscriptfolder

    if ahkexe is None:
        GetAhkExe()
    if not ahkexe:
        ahkscriptfolder = ""
        return ahkscriptfolder

    if not ahkscriptfolder is None:
        ## for repeated use:
        return ahkscriptfolder

    scriptfolder = Path.home()/".autohotkey"
    checkDirectory(scriptfolder, create=True)
    ahkscriptfolder = scriptfolder
    copySampleAhkScripts(sampleAhkDirectory, ahkscriptfolder)
    return scriptfolder

def clearProgInfoFile():
    """remove the previous ProgInfo file if present
    
    return the path of the ProgInfoFile
    """
    ProgInfoFile = ahkscriptfolder/"proginfofromahk.txt"
    if ProgInfoFile.is_file():
        ProgInfoFile.unlink()
    return ProgInfoFile

def clearErrorMessagesFile():
    """make an empty ErrorMessagesFile
    
    return the path of the ErrorMessagesFile
    """
    ErrorFile = ahkscriptfolder/"errormessagefromahk.txt"
    with open(ErrorFile, 'w') as f:
        f.write('\n')
    return ErrorFile

def readErrorMessagesFile():
    """get the error messages if any
    """
    ErrorFile = ahkscriptfolder/"errormessagefromahk.txt"
    with open(ErrorFile, 'r') as f:
        mess = f.read()
    if mess.strip():
        return f'autohotkeyactions.ahkBringup failed:\n===={mess}'
    return ''



def killWindow(hndle=None, key_close=None, key_close_dialog=None, silent=True):
    """kill the app with hndle,
    
    like the unimacro shorthand command KW (killwindow)
        
    input:
    hndle: int, windows handle of the app to be closed, normally leave away,
                the foreground hndle is taken
    key_close: keystrokes to be performed to close a top window, default: `{alt+f4}`
    key_close_dialog: keystrokes to close the dialog if a dialog (child window) is in
                      the foreground. Default  `{alt+n}`
    silent: if messages are printed (True by default)
    
    returns: True if success
             message with progInfo (if possible) if fail
    
    tested with unittestAutohotkeyactions.py...
    """
    # pylint: disable=R0911, R0912
    ErrorFile = clearErrorMessagesFile()
    ProgInfoFile = clearProgInfoFile()
   
    if hndle:
        result = SetForegroundWindow(hndle)
        if result is not True:
            mess = f'window {hndle} not any more available'
            if not silent:
                print(mess)
            return mess
    progInfo = getProgInfo()
    if hndle:
        if hndle != progInfo.hndle:
            mess = f'hndle {hndle} does not match hndle of foreground window\n{progInfo}'
            return mess
    else:
        hndle = progInfo.hndle
    
    key_close = key_close or "{alt+f4}"
    key_close_dialog = key_close_dialog or "{alt+n}"

    if progInfo.toporchild == 'child':
        if not silent:
            print('child window in the foreground, first close, then proceed')
        sendkeys("{escape}")
        time.sleep(0.5)
        foregroundProgInfo = getProgInfo()
        if progInfo.prog != foregroundProgInfo.prog:
            mess = f'after esc key, there is now another program in the foreground: {progInfo.prog}'
            return mess
    
    ## close the window:
    sendkeys(key_close)
    newProgInfo = getProgInfo()
    if newProgInfo.prog != progInfo.prog:
        # OK!
        return True
    
    ## no success, child window must be closed:
    if newProgInfo.toporchild == 'top':
        mess = f'if file needs saving, a child window should be on top now, not {newProgInfo}'
        return mess
    
    ## try to send the no save close dialog key now:
    sendkeys(key_close_dialog)

    progInfo = getProgInfo()
    if progInfo.toporchild == 'child':
        mess = f'killWindow, failed to close child dialog\n\t{progInfo}'
        return mess
    
    if hndle == progInfo.hndle:
        mess = f'killing window failed for {hndle}\n\t{progInfo}'
        return mess
    return True


## initialise ahkexe and ahkscriptfolder:
# GetAhkExe() # is done in next line:
if ahkscriptfolder is None:
    GetAhkScriptFolder()

def test():
    # pylint: disable=C0115, C0116, C0415
    import doctest
    doctest.testmod()


if __name__ ==  "__main__":
    test()
    print(f'ahk_is_active: {ahk_is_active()}')
    Result = getProgInfo()
    print(f'\nresult of getProgModInfo: (start)\n{repr(Result)}')
    Result = ahkBringup("notepad")
    print(f'\nresult of ahkBringup("notepad"):\n{repr(Result)}')
    if Result.hndle:
        killWindow(Result.hndle)
    
    


