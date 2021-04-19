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
    
5. When scripts want information back into python, write this information into "INFOfromAHK.txt".
   Scripts that use this feature, should have a function in this module, in order to read the data from this file.
   See getProgInfo and getCurrentModule below.
   
"""
import subprocess
import shutil
import filecmp
from pathlib import Path
import collections
import os.path  # only for getting the ahk exe path...
import time

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

def do_ahk_script(script):
    """try autohotkey integration
    """
    if not ahk_is_active():
        print('ahk is not active, cannot run script')
        return
    #print 'AHK with script: %s'% script
    scriptPath = ahkscriptfolder/'tempscript.ahk'
    with open(scriptPath, 'w') as fp:
        fp.write(script+'\n')
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
        return 

ProgInfo = collections.namedtuple('ProgInfo', 'progpath prog title toporchild classname hndle'.split(' '))

def getProgInfo():
    """get the prog info, like natlink.getCurrentModule enhanced with toporchild and classname
    
    returns program info as namedtuple (progpath, prog, title, toporchild, classname, hndle)

    So the length of progInfo is 6!

    toporchild 'top' or 'child', or '' if no valid window
          
    """
    WinInfoFile = ahkscriptfolder/"proginfofromahk.txt"
    if WinInfoFile.is_file():
        WinInfoFile.unlink()
    
    script = """; put prog info of current window in file.
;(prog, title, toporchild, classname, hndle)

WinGet pPath, ProcessPath, A
WinGetTitle, Title, A
WinGetClass, Class, A
WinGet wHndle, ID, A
wHndle := wHndle + 0
toporchild := "top"
if DllCall("GetParent", UInt, WinExist("A")) {
    toporchild := "child"    
}
WinGetClass, Class, A
WinGetTitle, Title, A

FileDelete, ##INFOfile##
FileAppend, %pPath%`n, ##INFOfile##
FileAppend, %Title%`n, ##INFOfile##
FileAppend, %toporchild%`n, ##INFOfile##
FileAppend, %Class%`n, ##INFOfile##
FileAppend, %wHndle%, ##INFOfile##
"""
    script = script.replace('##INFOfile##', str(WinInfoFile))
    do_ahk_script(script)

    with open(WinInfoFile, 'r') as fp:
        progInfo = fp.read().split('\n')
        
    # note ahk returns 5 lines, but ProgInfo has 6 items.   
        
    if len(progInfo) == 5:
        # make hndle decimal number:
        pPath, wTitle, toporchild, classname, hndle = progInfo
        hndle = int(hndle)
        prog = Path(pPath).stem
        return ProgInfo(pPath, prog, wTitle, toporchild, classname, hndle)
    raise ValueError(f'ahk script getProgInfo did not return correct result:\n    length {len(progInfo)}\n    text: {repr(progInfo)}')

GetProgInfo = getProgInfo


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
          
def GetRunWinwordScript(filepath, HNDLEfile):
    """construct script than opens a word document
    
    filepath is the Word document to open.
    HNDLEfile is a complete path to a file that will hold the windows handle after the script has run
    The script can retrieve the handle from this file (for the moment see actions.UnimacroBringup)
    
    """
    script = f'''Word := ComObjCreate("Word.Application")
Word.Visible := True
Word.Documents.Open("{filepath}")
Word.Visible := 1
Word.Activate
WinGet, hWnd, ID, A
FileDelete, {HNDLEfile}
FileAppend, %%hWnd%%, {HNDLEfile}
'''% (filepath, HNDLEfile, HNDLEfile)
    return script

def ahkBringup(app, filepath=None, title=None):
    """start a program, folder, file, with AutoHotkey
    
    This functions is related to UnimacroBringup, which works with AppBringup from Dragon,
    but sometimes fails.
    
    Besides, this function can also work without Dragon being on, not relying on natlink.
    So better fit for debugging purposes.
    
    """
    # pylint: disable=R1710
    if not ahk_is_active():
        print('cannot run ahkBringup, autohotkey is not active')
    WinInfoFile = str(ahkscriptfolder/"WININFOfromAHK.txt")
    
    ## treat mode = open or edit, finding a app in actions.ini:
    if ((app and app.lower() == "winword") or
        (filepath and (str(filepath).endswith(".docx") or str(filepath).endswith('.doc')))):
        script = GetRunWinwordScript(filepath, WinInfoFile)
        do_ahk_script(script)

    elif app and title:
        ## start eg thunderbird.exe this way
        ## can also give additional startup instructions in extra
        extra = extra or ""
        script = '''SetTitleMatchMode, 2
Process, Exist, ##basename##
if !ErrorLevel = 0
{
IfWinNotActive, ##title##,
WinActivate, ##title##, 
WinWaitActive, ##title##,,1
if ErrorLevel {
    return
}
}
else
{
Run, ##app##
WinWait, ##title##,,5
if ErrorLevel {
    MsgBox, AutoHotkey, WinWait for running ##basename## timed out
    return
}
}
##extra##
WinGet pPath, ProcessPath, A
WinGetTitle, Title, A
WinGet wHndle, ID, A
FileDelete, ##INFOfile##
FileAppend, %pPath%`n, ##INFOfile##
FileAppend, %Title%`n, ##INFOfile##
FileAppend, %wHndle%, ##INFOfile##

'''
        basename = app.name
        script = script.replace('##extra##', extra)
        script = script.replace('##app##', app)
        script = script.replace('##basename##', basename)
        script = script.replace('##title##', title)
        script = script.replace('##INFOfile##', WinInfoFile)
        do_ahk_script(script)
            
    else:
        ## other programs:
        if app and filepath:
            script = ["Run, %s, %s,,NewPID"% (app, filepath)]
        elif filepath:
            script = ["Run, %s,,, NewPID"% filepath]
        elif app:
            script = ["Run, %s,,, NewPID"% app]
    
        script.append("WinWait, ahk_pid %NewPID%")
    
        script.append("WinGet, pPath, ProcessPath, ahk_pid %NewPID%")
        script.append("WinGetTitle, Title, A")  ##, ID, ahk_pid %NewPID%")
        script.append("WinGet, wHndle, ID, ahk_pid %NewPID%")
        script.append("""wHndle := wHndle + 0
toporchild := "top"
if DllCall("GetParent", UInt, WinExist("A")) {
    toporchild := "child"    
}
WinGetClass, Class, A
WinGetTitle, Title, A
""")
        script.append('FileDelete, ' + WinInfoFile)
        script.append('FileAppend, %pPath%`n, ' + WinInfoFile)
        script.append('FileAppend, %Title%`n, ' + WinInfoFile)
        script.append('FileAppend, %toporchild%`n, ' + WinInfoFile)
        script.append('FileAppend, %Class%`n, ' + WinInfoFile)
        script.append('FileAppend, %wHndle%, ' + WinInfoFile)
        script = '\n'.join(script)

        do_ahk_script(script)

    ## collect the wHndle:
        progInfo = open(WinInfoFile, 'r').read().split('\n')
        if len(progInfo) == 5:
            pPath, wTitle, toporchild, classname, hndle = progInfo
            hndle = int(hndle)
            prog = Path(pPath).stem
            return ProgInfo(pPath, prog, wTitle, toporchild, classname, hndle)

        print(f'autohotkeyactions, return of getting proInfo in appBringup: "{repr(progInfo)}" (length: {len(progInfo)}')
        return

autohotkeyBringup = ahkBringup

def GetForegroundWindow():
    """return the hndle of the ForegroundWindow
    
    return value: int: the hndle of the Foreground window
    return value: str: error message, function failed
    """
    WinInfoFile = ahkscriptfolder/"foregroundhndlefromahk.txt"
    if WinInfoFile.is_file():
        WinInfoFile.unlink()
    
    script = """; put window hndle of current window in file.

WinGet pPath, ProcessPath, A
WinGet wHndle, ID, A
wHndle := wHndle + 0
FileDelete, ##INFOfile##
FileAppend, %wHndle%, ##INFOfile##
"""
    script = script.replace('##INFOfile##', str(WinInfoFile))
    do_ahk_script(script)

    with open(WinInfoFile, 'r') as fp:
        gotHndle = fp.read().strip()
    try:
        hndleInt = int(gotHndle)
        return hndleInt
    except ValueError:
        # pylint: disable=R1705  # (unnecessary else)
        if gotHndle:
            mess = f'autohotkeyactions, GetForegroundWindow: did not get correct hndle window: {gotHndle}'
            print(mess)
            return mess
        else:
            mess = 'autohotkeyactions, GetForegroundWindow: did not get correct hndle window: (empty)'
            print(mess)
            return mess

def SetForegroundWindow(hndle):
    """bring window with hndle into the foreground
     
    return value 0: success
    other return value, mostly the hndle of the active window: failure
    
    """
    InfoFile = ahkscriptfolder/"INFOfromAHK.txt"
    if InfoFile.is_file():
        InfoFile.unlink()

    script = f'''WinActivate, ahk_id {hndle}
WinGet wHndle, ID, A
wHndle := wHndle + 0
FileDelete, ##INFOfile##
FileAppend, %wHndle%, ##INFOfile##
'''
    script = script.replace('##INFOfile##', str(InfoFile))
    do_ahk_script(script)
    winHndle = open(InfoFile, 'r').read().strip()
    if winHndle:
        try:
            winHndle = int(winHndle)
        except ValueError:
            print(f'ahk script getIntoForeground to "{hndle}" did not return correct winHndle: {winHndle}')
            return winHndle
            
        if winHndle == hndle:
            return 0  # ok, then return None
        print(f'could not switch to wanted hndle {hndle} (got {winHndle})')
        return winHndle
    raise ValueError('ahk script getIntoForeground did not return anything')

def getFileDate(modName):
    """return the last modified date/time of file
    """
    try:
        modTime = modName.stat().st_mtime
        return modTime
    except OSError:
        return 0        # file not found

def GetAhkExe():
    """try to get executable of autohotkey.exe, if not there, empty string is put in ahkexe
    
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

def killWindow(hndle, key_close=None, key_close_dialog=None):
    """kill the app with hndle,
    
    like the unimacro shorthand command KW (killwindow)
        
    input:
    hndle: int, windows handle of the app to be closed
    key_close: keystrokes to be performed to close a top window, default: `{alt+f4}`
    key_close_dialog: keystrokes to close the dialog if a dialog (child window) is in
                      the foreground. Default  `{esc}`
    
    tested with unittestAutohotkeyactions.py...
    """
    result = SetForegroundWindow(hndle)
    if result:
        print(f'window {hndle} not any more available')
        return
    key_close = key_close or "{alt+f4}"
    key_close_dialog = key_close_dialog or "{alt+n}"
    progInfo = getProgInfo()
    foregroundProg = progInfo.prog
    if progInfo.hndle != hndle:
        print(f'invalid window {progInfo.hndle} in the foreground, want {hndle}')
        return
    if progInfo.toporchild == 'child':
        print(f'child window in the foreground, expected top {hndle}')
        return
    sendkeys(key_close)
    progInfo = getProgInfo()
    if progInfo.prog != foregroundProg:
        return
    if progInfo.toporchild == 'child':
        sendkeys(key_close_dialog)

    progInfo = getProgInfo()
    if progInfo.toporchild == 'child':
        print('killWindow, failed to close child dialog')




## initialise ahkexe and ahkscriptfolder:
# GetAhkExe() # is done in next line:
if ahkscriptfolder is None:
    GetAhkScriptFolder()




if __name__ ==  "__main__":

    from sendkeys import sendkeys
    print(f'ahk_is_active: {ahk_is_active()}')
    Result = getProgInfo()
    print(f'\nresult of getProgModInfo: (start)\n{repr(Result)}')
    Result = ahkBringup("notepad")
    print(f'\nresult of ahkBringup("notepad"):\n{repr(Result)}')
    Result = getProgInfo()
    print(f'\nresult of getProgModInfo (should be notepad):\n{repr(Result)}')
    time.sleep(1)
    # quite notepad again...
    sendkeys("{alt+f4}")
    Result = getProgInfo()
    print(f'\nresult of getProgModInfo:\n{repr(Result)}')
    
    
    


