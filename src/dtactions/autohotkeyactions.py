# (unimacro - natlink macro wrapper/extensions)
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
import os
import sys
import subprocess
import stat
import shutil
from natlinkcore import natlinkcorefunctions
from natlinkcore import natlinkstatus
import win32gui

## get thisDir and other basics...
try:
    from dtactions.__init__ import findInSitePackages
except ModuleNotFoundError:
    findInSitePackages = None

def getThisDir():
    """get directory of this, if possible in site-packages
    
    Check for symlink and presence in site-packages directory
    """
    thisFile = __file__
    thisDir = os.path.split(thisFile)[0]
    if findInSitePackages:
        thisDir = findInSitePackages(thisDir)
    return thisDir

dtactions = thisDir = getThisDir()
##### get actions.ini from baseDirectory or SampleDirectory into userDirectory:
sampleAhkDirectory = os.path.join(dtactions, 'samples')

ahkexe = None
ahkscriptfolder = None

def ahk_is_active():
    if ahkexe is None:
        GetAhkExe()
    return ahkexe

def do_ahk_script(script, hndle=None):
    """try autohotkey integration
    """
    global ahkscriptfolder
    if ahkexe is None:
        GetAhkExe() 
    if not ahkexe:
        raise ValueError("cannot run AHK action, autohotkey.exe not found")
    if ahkscriptfolder is None:
        GetAhkScriptFolder()
    if not ahkscriptfolder:
        raise IOError("no folder for AutoHotkey scripts found")
   
    #print 'AHK with script: %s'% script
    if script.endswith(".ahk"): 
        scriptPath = os.path.join(ahkscriptfolder, script)
        if os.path.isfile(scriptPath):
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
    scriptPath = os.path.join(ahkscriptfolder, 'tempscript.ahk')
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
    
def GetAhkExe():
    """try to get executable of autohotkey.exe, if not there, empty string is put in ahkexe
    
    """
    global ahkexe
    # no succes, go on with program files:
    pf = natlinkcorefunctions.getExtendedEnv("PROGRAMFILES")
    if pf.find('(x86)')>0:
        # 64 bit:
        pf = natlinkcorefunctions.getExtendedEnv("PROGRAMW6432")  # the old pf directory
    if not os.path.isdir(pf):
        raise IOError("cannot find (old style) program files directory: %s"% pf)
    
    
    ahk = os.path.join(pf, "autohotkey", "autohotkey.exe")
    if os.path.isfile(ahk):
        ahkexe = ahk
        #print 'AutoHotkey found, %s'% ahkexe
    else:
        ahkexe = ""
        #print 'AutoHotkey not found on this computer (%s)'% ahkexe


    
def GetAhkScriptFolder():
    """try to get AutoHotkey folder as subdirectory of PERSONAL
    
    create if non-existent.
    
    """
    global ahkscriptfolder

    if not ahkscriptfolder is None:
        ## for repeated use:
        return ahkscriptfolder

    scriptfolder = os.path.expanduser("~\\.autohotkey")
    if not os.path.isdir(scriptfolder):
        os.mkdir(scriptfolder)
    if not os.path.isdir(scriptfolder):
        raise IOError(f'scriptFolder {scriptfolder} does not exist and cannot be created')
    ahkscriptfolder = scriptfolder
    copySampleAhkScripts(sampleAhkDirectory, ahkscriptfolder)
    return scriptfolder

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
    if not os.path.isdir(fromFolder):
        print('No sample_ahk dir found (should be in Unimacro directory): "%s"'% fromFolder)
        return
    globString = f'{fromFolder}\\*.ahk'
    for f in glob.glob(globString):
        inputFile = f
        dirPart, filename = os.path.split(f)
        outputFile = os.path.join(toFolder, filename)
        if not os.path.isfile(outputFile):
            print('---copy AutoHotkey script "%s" from\nSamples directory "%s"\nTo  "%s"\n----'% (filename, fromFolder, toFolder))
            shutil.copyfile(inputFile, outputFile)
        elif getFileDate(inputFile) > getFileDate(outputFile):
            if compare_f2f(inputFile, outputFile):
                oldCopy = outputFile + 'old'
                if os.path.isfile(oldCopy):
                    print('AutoHotkey script "%s" has been changed in "sample_ahk", copy to "%s"'% (filename, toFolder))
                else:
                    print('AutoHotkey script "%s" has been changed in "sample_ahk", copy to "%s"\n(keep backup in %s)'% (filename, toFolder, oldCopy))
                    shutil.copyfile(outputFile, oldCopy)
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


if __name__ ==  "__main__":
    result = getModInfo()
    print('result of getModInfo: ', repr(result))