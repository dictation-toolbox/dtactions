#
# Python Macro Language for Dragon NaturallySpeaking
#   (c) Copyright 1999 by Joel Gould
#   Portions (c) Copyright 1999 by Dragon Systems, Inc.
#
#pylint:disable=E0611, E0401, R0911, R0912, W0702, C0116
#pylint:disable=C0209
"""extenvvars.py

Keep track of environment variables, including added "fake" variables like DROPBOX, "This PC", etc.

 Quintijn Hoogenboom, January 2008/September 2015/November 2022 (python3)


"""
import os
import re
import copy
from win32com.shell import shell, shellcon
try:
    import natlink
    natlinkAvailable = True
except ImportError:
    natlinkAvailable = False
    
if natlinkAvailable:
    if not natlink.isNatSpeakRunning():
        natlinkAvailable = False
        
if natlinkAvailable:
    from natlinkcore import natlinkstatus
    status = natlinkstatus.NatlinkStatus()
    status_dict_full = copy.copy(status.getNatlinkStatusDict())
    
    status_dict = {key.upper():value for (key, value) in status_dict_full.items() if os.path.isdir(value)}
    del status_dict_full
else:
    status = None
    status_dict = None
# for extended environment variables:
reEnv = re.compile('(%[A-Z_]+%)', re.I)


def getFolderFromLibraryName(fName):
    """from windows library names extract the real folder
    
    the CabinetWClass and #32770 controls return "Documents", "Dropbox", "Desktop" etc.
    try to resolve these throug the extended env variables.
    """
    if fName.startswith("Docum"):  # Documents in Dutch and English
        return getExtendedEnv("PERSONAL")
    if fName in ["Muziek", "Music"]:
        return getExtendedEnv("MYMUSIC")
    if fName in ["Pictures", "Afbeeldingen"]:
        return getExtendedEnv("MYPICTURES")
    if fName in ["Videos", "Video's"]:
        return getExtendedEnv("MYVIDEO")
    if fName in ['OneDrive']:
        return getExtendedEnv("OneDrive")
    if fName in ['Desktop', "Bureaublad"]:
        return getExtendedEnv("DESKTOP")
    if fName in ['Quick access', 'Snelle toegang']:
        templatesfolder = getExtendedEnv('TEMPLATES')
        if os.path.isdir(templatesfolder):
            QuickAccess = os.path.normpath(os.path.join(templatesfolder, "..", "Libraries"))
            if os.path.isdir(QuickAccess):
                return QuickAccess
    if fName == 'Dropbox':
        return getDropboxFolder()
    if fName in ['Download', 'Downloads']:
        personal = getExtendedEnv('PERSONAL')
        userDir = os.path.normpath(os.path.join(personal, '..'))
        if os.path.isdir(userDir):
            tryDir = os.path.normpath(os.path.join(userDir, fName))
            if os.path.isdir(tryDir):
                return tryDir
    usersHome = os.path.normpath(os.path.join(r"C:\Users", fName))
    if os.path.isdir(usersHome):
        return usersHome
    if fName in ["This PC", "Deze pc"]:
        return "\\"
    print('cannot find folder for Library name: %s'% fName)
    return ""


def getDropboxFolder(containsFolder=None):
    """get the dropbox folder, or the subfolder which is specified.
    
    Searching is done in all 'C:\\Users' folders, and in the root of "C:"
    (See DirsToTry)
    
    raises OSError if more folders are found (should not happen, I think)
    if containsFolder is not passed, the dropbox main folder is returned
    if containsFolder is passed, this folder is returned if it is found in the dropbox folder
    
    otherwise None is returned.
    """
    results = []
    root = 'C:\\Users'
    dirsToTry = [ os.path.join(root, s) for s in os.listdir(root) if os.path.isdir(os.path.join(root,s)) ]
    dirsToTry.append('C:\\')
    for root in dirsToTry:
        if not os.path.isdir(root):
            continue
        try:
            subs = os.listdir(root)
        except WindowsError:
            continue
        if 'Dropbox' in subs:
            subAbs = os.path.join(root, 'Dropbox')
            subsub = os.listdir(subAbs)
            if not ('.dropbox' in subsub and os.path.isfile(os.path.join(subAbs,'.dropbox'))):
                continue
            if containsFolder:
                result = matchesStart(subsub, containsFolder, caseSensitive=False)
                if result:
                    results.append(os.path.join(subAbs,result))
            else:
                results.append(subAbs)
    if not results:
        return ''
    if len(results) > 1:
        raise OSError('getDropboxFolder, more dropbox folders found: %s')
    return results[0]                 

def matchesStart(listOfDirs, checkDir, caseSensitive):
    """return result from list if checkDir matches, mostly case insensitive
    """
    if not caseSensitive:
        checkDir = checkDir.lower()
    for l in listOfDirs:
        if not caseSensitive:
            ll = l.lower()
        else:
            ll = l
        if ll.startswith(checkDir):
            return l
    return False   
        
            

def getExtendedEnv(var):
    """get from environ or windows CSLID

    HOME is environ['HOME'] or CSLID_PERSONAL
    ~ is HOME

    DROPBOX added via getDropboxFolder is this module (QH, December 1, 2018)

    Also the directories that are configured inside Natlink can be retrieved, via natlinkstatus.py
    Natlink_SettingsDir, DragonflyUserDirectory, VocolaDirectory, AhkUserDir
    
    Note: these settings are case sensitive! You can leave out Dir or Directory.
    
    """
##    var = var.strip()
    var = var.strip("% ").upper()
    if status_dict:
        result = getDirectoryFromNatlinkstatus(var)
        if result:
            return result
  
    
    if var == "~":
        var = 'HOME'

    if var in os.environ:
        return os.environ[var]

    if var == 'DROPBOX':
        result = getDropboxFolder()
        if result:
            return result
        raise ValueError('getExtendedEnv, cannot find path for "DROPBOX"')

    if var == 'NOTEPAD':
        windowsDir = getExtendedEnv("WINDOWS")
        notepadPath = os.path.join(windowsDir, 'notepad.exe')
        if os.path.isfile(notepadPath):
            return notepadPath
        raise ValueError('getExtendedEnv, cannot find path for "NOTEPAD"')

    # try to get from CSIDL system call:
    if var == 'HOME':
        var2 = 'PERSONAL'
    else:
        var2 = var
        
    try:
        CSIDL_variable =  'CSIDL_%s'% var2
        shellnumber = getattr(shellcon,CSIDL_variable, -1)
    except:
        print('getExtendedEnv, cannot find in environ or CSIDL: "%s"'% var2)
        return ''
    if shellnumber < 0:
        # on some systems have SYSTEMROOT instead of SYSTEM:
        if var == 'SYSTEM':
            return getExtendedEnv('SYSTEMROOT')
        return ''
        # raise ValueError('getExtendedEnv, cannot find in environ or CSIDL: "%s"'% var2)
    try:
        result = shell.SHGetFolderPath (0, shellnumber, 0, 0)
    except:
        return ''

    
    result = str(result)
    result = os.path.normpath(result)
    if result and os.path.isdir(result):
    # on some systems apparently:
        return result
    if result:
        print(f'getExtendedEnv: no valid path found for "{var}": "{result}"')
    else:
        print(f'getExtendedEnv: no path found for "{var}"')
    return None
        
def getDirectoryFromNatlinkstatus(envvar):
    """see if directory can can be retrieved from envvar
    """
    # if natlink not available:
    if not natlinkAvailable:
        # print(f'natlink not available for get "{envvar}"')
        return None

    # try if function in natlinkstatus:
    if not status_dict:
        return None
    
    for extra in ('', 'DIRECTORY', 'DIR'):
        var2 = envvar + extra
        result = status_dict.get(var2, "")
        if result:
            return result
    return None
       
def expandEnvVariableAtStart(filepath): 
    """try to substitute environment variable into a path name

    """
    filepath = filepath.strip()

    if filepath.startswith('~'):
        folderpart = getExtendedEnv('~')
        filepart = filepath[1:]
        filepart = filepart.strip('/\\ ')
        return os.path.normpath(os.path.join(folderpart, filepart))
    if reEnv.match(filepath):
        envVar = reEnv.match(filepath).group(1)
        # get the envVar...
        try:
            folderpart = getExtendedEnv(envVar)
        except ValueError:
            print('invalid (extended) environment variable: %s'% envVar)
        else:
            # OK, found:
            filepart = filepath[len(envVar)+1:]
            filepart = filepart.strip('/\\ ')
            return os.path.normpath(os.path.join(folderpart, filepart))
    # no match
    return filepath
    
def expandEnvVariables(filepath): 
    """try to substitute environment variable into a path name,

    ~ only at the start,

    %XXX% can be anywhere in the string.

    """
    filepath = filepath.strip()
    
    if filepath.startswith('~'):
        folderpart = getExtendedEnv('~')
        filepart = filepath[1:]
        filepart = filepart.strip('/\\ ')
        filepath = os.path.normpath(os.path.join(folderpart, filepart))
    
    if reEnv.search(filepath):
        List = reEnv.split(filepath)
        #print 'parts: %s'% List
        List2 = []
        for part in List:
            if not part:
                continue
            if part == "~" or (part.startswith("%") and part.endswith("%")):
                try:
                    folderpart = getExtendedEnv(part)
                except ValueError:
                    folderpart = part
                List2.append(folderpart)
            else:
                List2.append(part)
        filepath = ''.join(List2)
        return os.path.normpath(filepath)
    # no match
    return filepath


if __name__ == "__main__":

    print('testing       expandEnvVariableAtStart')
    print('also see expandEnvVar in natlinkstatus!!')
    for p in ("~", "%home%", "D:\\natlink\\unimacro", "~/unimacroqh",
              "%HOME%/personal",
              "%WINDOWS%\\folder\\strange testfolder"):
        expanded = expandEnvVariableAtStart(p)
        print('expandEnvVariablesAtStart: %s: %s'% (p, expanded))
    print('testing       expandEnvVariables')  
    for p in ("%NATLINK%\\unimacro", "%DROPBOX%/QuintijnHerold/jachthutten", "D:\\%username%", "%UNIMACROUSER%",
              "%HOME%/personal", "%HOME%", "%personal%"
              "%WINDOWS%\\folder\\strange testfolder"):
        expanded = expandEnvVariables(p)
        print('expandEnvVariables: %s: %s'% (p, expanded))

    np = getExtendedEnv("NOTEPAD")
    print(np)
    for lName in ['Snelle toegang', 'Quick access', 'Documenten', 'Documents', 'Muziek', 'Afbeeldingen', 'Dropbox', 'OneDrive', 'Desktop', 'Bureaublad']:
        f = getFolderFromLibraryName(lName)
        print('folder from library name %s: %s'% (lName, f))
