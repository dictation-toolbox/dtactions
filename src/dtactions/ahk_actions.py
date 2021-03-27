# unimacro natlink macro wrapper/extensions
# (c) copyright 2003 Quintijn Hoogenboom (quintijn@users.sourceforge.net)
#
# autohotkeyactions.py 
#  written by: Quintijn Hoogenboom (QH softwaretraining & advies)
#  December 2013/.../March 2021, adding AutoHotkey support
#
"""This module contains actions via AutoHotkey

1. the exe and scriptfolder are collected at first call and stored in
   variables ahkexe and ahkscriptfolder

2. ahk_is_active returns whether AutoHotkey is running on your system

3. call_ahk_script_path calls the script you have defined. If you have the script in
   a string, the string is copied to tempscript.ahk and executed.
   
4. if %hndle% is in your script, this string is substituted by the handle of the foreground window and the copy
   of the script goes to tempscript.ahk and is then executed.
   
5. At first call the sample_ahk directory of Unimacro is copied into the  AutoHotkey in your Documents folder,
   this folder is created if it does not exist yet. Of course only files that do not exist in the AutoHotkey directory
   are copied this way (copySampleAhkScripts function).
   
GetAhkExe gets the correct ahkexe, or "" if AutoHotkey is not on your computer
GetAhkScriptFolder get the correct scriptfolder ((AutoHotkey in your Documents folder).autohotkey in your home directory)
    and copies scripts as described in 5. above
     
"""
import glob
import sys
import subprocess
import stat
import shutil
from pathlib import Path
from natlinkcore import natlinkcorefunctions
from natlinkcore import natlinkstatus
import win32gui

## get thisDir and other basics...
try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print(f'Run this module after "build_package" and "flit install --symlink"\n')
    raise

dtactions = thisDir = getThisDir(__file__)
##### get actions.ini from baseDirectory or SampleDirectory into userDirectory:
sampleAhkDirectory = dtactions/'samples'/'autohotkey'
checkDirectory(sampleAhkDirectory)

ahkexe = None
ahkscriptfolder = None

def ahk_is_active():
    return bool(ahkexe) and bool(ahkscriptfolder)

def do_ahk_script(script, hndle=None):
    """try autohotkey integration
    """
    global ahkscriptfolder
    if not ahk_is_active():
        print('ahk is not active, cannot run script')
        return

    if ahkscriptfolder is None:
        GetAhkScriptFolder()
    if not ahkscriptfolder:
        raise OSError("no folder for AutoHotkey scripts found")
   
    #print 'AHK with script: %s'% script
    if script.endswith(".ahk"): 
        scriptPath = ahkscriptfolder/script
        if scriptPath.is_file():
            scriptText = open(scriptPath, 'r').read()
            if scriptText.find(r'%hndle%') >= 0:
                #print 'take scriptText for replacing %%hndle%%: %s'% scriptText
                script = scriptText
            else:
                # just run ahk script:
                result = call_ahk_script_path(scriptPath)
                if result:
                    return 'AHK error: %s'% result
                else:
                    return 1
        else:
            return "action AHK, not an existing script file: %s (%s)"% (script, scriptPath)
    if script.find(r"%hndle%") >= 0:
        if hndle is None:
            hndle = win32gui.GetForegroundWindow()
        script = script.replace("%hndle%", "%s"% hndle)
        #print 'substituted script: %s'% script
    #print 'AHK with script: %s'% script
    scriptPath = ahkscriptfolder/'tempscript.ahk'
    open(scriptPath, 'w').write(script+'\n')
    result = call_ahk_script_path(scriptPath)
    if result:
        return 'AHK error: %s'% result
    else:
        return 1

def getModInfo():
    """get the module info, like natlink.getCurrentModule
    """
    scriptFolder = GetAhkScriptFolder()
    WinInfoFile = os.path.join(scriptFolder, "WININFOfromAHK.txt")
    script = """; put module info of current window in file. 

WinGet pPath, ProcessPath, A
WinGetTitle, Title, A
WinGet wHndle, ID, A
FileDelete, ##WININFOfile##
FileAppend, %pPath%`n, ##WININFOfile##
FileAppend, %Title%`n, ##WININFOfile##
FileAppend, %wHndle%, ##WININFOfile##
"""
    script = script.replace('##WININFOfile##', WinInfoFile)
    result = do_ahk_script(script)

    if result == 1:
        winInfo = open(WinInfoFile, 'r').read().split('\n')
        if len(winInfo) == 3:
            # make hndle decimal number:
            pPath, wTitle, hndle = winInfo
            hndle = int(hndle, 16)
            # print('extracted pPath: %s, wTitle: %s and hndle: %s'% (pPath, wTitle, hndle))
            return pPath, wTitle, hndle
    raise ValueError('ahk script getModInfo did not return correct result: %s'% result)
    

def call_ahk_script_path(scriptPath):
    """call the specified ahk script
    
    use the global variable ahkexe as executable
    
    """
    result = subprocess.call([ahkexe, scriptPath, ""])
    if result:
        print('non-zero result of call_ahk_script_path "%s": %s'% (scriptPath, result))
        return 

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
# 
def copySampleAhkScripts(fromFolder, toFolder):
    """copy (in new Autohotkey directory) the sample script files
    """
    if not fromFolder.is_dir():
        print(f'No sample_ahk dir found (should be in dtactions/sample): {fromFolder}')
        return
    for f in fromFolder.glob("*.ahk"):
        inputFile = f
        inputName = f.name
        stem = f.stem
        outputFile = toFolder/inputName
        if not outputFile.is_file():
            print('---copy AutoHotkey script "{inputName}" from\nSamples directory to "{toFolder}"----'% (filename, fromFolder, toFolder))
            shutil.copyfile(inputFile, outputFile)
        elif getFileDate(inputFile) > getFileDate(outputFile):
            if compare_f2f(inputFile, outputFile):
                for i in range(1000):
                    outputName = f'{stem_{i:03d}}'
                    newOutputFile = outputFile(with_name(newOutputName))
                    if not newOutputFile.is_file():
                        break
                else:
                    raise OSError('no unused newOutputFile available last: {newOutputFile}')
                print(f'AutoHotkey script "{inputName}" has been changed in "sample_ahk"\n   copy to "{toFolder}"\n   keep backup in {newOutputFile}')
                shutil.copyfile(outputFile, newOutputFile)
                shutil.copyfile(inputFile, outputFile)
          
def GetRunWinwordScript(filepath, HNDLEfile):
    """construct script than opens a word document
    
    filepath is the Word document to open.
    HNDLEfile is a complete path to a file that will hold the windows handle after the script has run
    The Bringup script can retrieve the handle from this file (for the moment see actions.UnimacroBringup)
    
    """
    script = '''Word := ComObjCreate("Word.Application")
Word.Visible := True
Word.Documents.Open("%s")
Word.Visible := 1
Word.Activate
WinGet, hWnd, ID, A
FileDelete, %s
FileAppend, %%hWnd%%, %s
'''% (filepath, HNDLEfile, HNDLEfile)
    return script

def ahkBringup(app, filepath=None, title=None, extra=None, modInfo=None, progInfo=None):
    """start a program, folder, file, with AutoHotkey
    
    This functions is related to UnimacroBringup, which works with AppBringup from Dragon,
    but sometimes fails.
    
    Besides, this function can also work without Dragon being on, not relying on natlink.
    So better fit for debugging purposes.
    
    """
    if not ahk_is_active():
        print(f'cannot run ahkBringup, autohotkey is not active')
    WinInfoFile = ahkScriptFolder/"WININFOfromAHK.txt"
    
    ## treat mode = open or edit, finding a app in actions.ini:
    if ((app and app.lower() == "winword") or
        (filepath and (filepath.endswith(".docx") or filepath.endswith('.doc')))):
        script = autohotkeyactions.GetRunWinwordScript(filepath, WinInfoFile)
        result = autohotkeyactions.do_ahk_script(script)

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
FileDelete, ##WININFOfile##
FileAppend, %pPath%`n, ##WININFOfile##
FileAppend, %Title%`n, ##WININFOfile##
FileAppend, %wHndle%, ##WININFOfile##

'''
        basename = os.path.basename(app)
        script = script.replace('##extra##', extra)
        script = script.replace('##app##', app)
        script = script.replace('##basename##', basename)
        script = script.replace('##title##', title)
        script = script.replace('##WININFOfile##', WinInfoFile)
        result = autohotkeyactions.do_ahk_script(script)
            
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
        script.append('FileDelete, ' + WinInfoFile)
        script.append('FileAppend, %pPath%`n, ' + WinInfoFile)
        script.append('FileAppend, %Title%`n, ' + WinInfoFile)
        script.append('FileAppend, %wHndle%, ' + WinInfoFile)
        script = '\n'.join(script)

        result = autohotkeyactions.do_ahk_script(script)

    ## collect the wHndle:
    if result == 1:
        winInfo = open(WinInfoFile, 'r').read().split('\n')
        if len(winInfo) == 3:
            # make hndle decimal number:
            pPath, wTitle, hndle = winInfo
            hndle = int(hndle, 16)
            print('extracted pPath: %s, wTitle: %s and hndle: %s'% (pPath, wTitle, hndle))
            return pPath, wTitle, hndle
        else:
            if natlink.isNatSpeakRunning():
                mess = "Result of ahk_script should be a 3 item list (pPath, wTitle, hndle), not: %s"% repr(winInfo)
                do_MSG(str(mess))
            print(str())
            return 0
    else:
        if natlink.isNatSpeakRunning():
            do_MSG(str(result))
        print(str(result))
        return 0



def getFileDate(modName):
    try: return os.stat(modName)[stat.ST_MTIME]
    except OSError: return 0        # file not found

def compare_f2f(f1, f2):
    """Helper to compare two files, return 0 if they are equal."""

    BUFSIZE = 8192
    fp1 = open(f1)
    try:
        fp2 = open(f2)
        try:
            while 1:
                b1 = fp1.read(BUFSIZE)
                b2 = fp2.read(BUFSIZE)
                if not b1 and not b2: return 0
                c = b1 != b2
                if c:
                    return c
        finally:
            fp2.close()
    finally:
        fp1.close()

def GetAhkExe():
    """try to get executable of autohotkey.exe, if not there, empty string is put in ahkexe
    
    """
    global ahkexe
    # no succes, go on with program files:
    pf = natlinkcorefunctions.getExtendedEnv("PROGRAMFILES")
    if pf.find('(x86)')>0:
        # 64 bit:
        pf = natlinkcorefunctions.getExtendedEnv("PROGRAMW6432")  # the old pf directory
    if pf and Path(pf).is_dir():
        pf = Path(pf)
    elif pf:
        raise OSError(f'cannot find (old style) program files directory: (empty)')
    else:
        raise OSError(f'cannot find (old style) program files directory: {pf}')
    
    ahk = pf/"autohotkey"/"autohotkey.exe"
    if ahk.is_file():
        ahkexe = ahk
        #print 'AutoHotkey found, %s'% ahkexe
    else:
        ahkexe = ""
        #print 'AutoHotkey not found on this computer (%s)'% ahkexe
    
def GetAhkScriptFolder():
    """try to get AutoHotkey folder as subdirectory of HOME
    
    On Windows mostly C:\\Users\\Username\\.autohotkey
    
    create if non-existent.
    """
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

## initialise ahkexe and ahkscriptfolder:
GetAhkExe()
GetAhkScriptFolder()

if __name__ ==  "__main__":
    print(f'ahk_is_active: {ahk_is_active()}')
    result = getModInfo()
    print('result of getModInfo: ', repr(result))
    result = ahkBringup("notepad")