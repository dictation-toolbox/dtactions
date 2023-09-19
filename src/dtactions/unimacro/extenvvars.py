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
from win32com.shell import shell, shellcon
from natlinkcore import natlinkstatus
status = natlinkstatus.NatlinkStatus()
# for extended environment variables:
reEnv = re.compile('(%[A-Z_]+%)', re.I)

# keep track of found env variables, fill, if you wish, with
# getAllFolderEnvironmentVariables.
# substitute back with substituteEnvVariableAtStart.
# and substite forward with expandEnvVariableAtStart
# in all cases a private envDict can be user, or the global dict recentEnv
#
# to collect all env variables, call getAllFolderEnvironmentVariables, see below
recentEnv = {}

def addToRecentEnv(name, value):
    """to be filled for NATLINK variables from natlinkstatus
    """
    recentEnv[name] = value

def deleteFromRecentEnv(name):
    """to possibly delete from recentEnv, from natlinkstatus
    """
    if name in recentEnv:
        del recentEnv[name]

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
        
            

def getExtendedEnv(var, envDict=None, displayMessage=1):
    """get from environ or windows CSLID

    HOME is environ['HOME'] or CSLID_PERSONAL
    ~ is HOME

    DROPBOX added via getDropboxFolder is this module (QH, December 1, 2018)

    Also the directories that are configured inside Natlink can be retrieved, via natlinkstatus.py
    Natlink_SettingsDir, DragonflyUserDirectory, VocolaDirectory, AhkUserDir
    
    Note: these settings are case sensitive! You can leave out Dir or Directory.
    
    As envDict for recent results either a private (passed in) dict is taken, or
    the global recentEnv.

    This is merely for "caching results"

    """
    if envDict is None:
        myEnvDict = recentEnv
    else:
        myEnvDict = envDict
##    var = var.strip()
    var = var.strip("% ")

    result = getDirectoryFromNatlinkstatus(var)
    if result:
        return result
    var = var.upper()
    
    if var == "~":
        var = 'HOME'

    if var in myEnvDict:
        return myEnvDict[var]

    if var in os.environ:
        myEnvDict[var] = os.environ[var]
        return myEnvDict[var]

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
            return getExtendedEnv('SYSTEMROOT', envDict=envDict)
        return ''
        # raise ValueError('getExtendedEnv, cannot find in environ or CSIDL: "%s"'% var2)
    try:
        result = shell.SHGetFolderPath (0, shellnumber, 0, 0)
    except:
        if displayMessage:
            print('getExtendedEnv, cannot find in environ or CSIDL: "%s"'% var2)
        return ''

    
    result = str(result)
    result = os.path.normpath(result)
    myEnvDict[var] = result
    # on some systems apparently:
    if var == 'SYSTEMROOT':

        myEnvDict['SYSTEM'] = result
    return result

def getDirectoryFromNatlinkstatus(envvar):
    """see if directory can can be retrieved from envvar
    """
    # try if function in natlinkstatus:
    for extra in ('', 'Directory', 'Dir'):
        var2 = envvar + extra
        if var2 in status.__dict__:
            funcName = f'get{var2}'
            func = getattr(status, funcName)
            result = func()
            if result:
                return result
    return None



def clearRecentEnv():
    """for testing, clears above global dictionary
    """
    recentEnv.clear()

def getAllFolderEnvironmentVariables(fillRecentEnv=None):
    """return, as a dict, all the environ AND all CSLID variables that result into a folder
    
    Now also implemented:  Also include NATLINK, UNIMACRO, VOICECODE, DRAGONFLY, VOCOLAUSERDIR, UNIMACROUSERDIR
    This is done by calling from natlinkstatus, see there and example in natlinkmain.

    Optionally put them in recentEnv, if you specify fillRecentEnv to 1 (True)

    """
    #pylint:disable=W0603
    D = {}

    for k in dir(shellcon):
        if k.startswith("CSIDL_"):
            kStripped = k[6:]
            try:
                v = getExtendedEnv(kStripped, displayMessage=None)
            except ValueError:
                continue
            if len(v) > 2 and os.path.isdir(v):
                D[kStripped] = v
            elif v == '.':
                D[kStripped] = os.getcwd
    # os.environ overrules CSIDL:
    for k in os.environ:
        v = os.environ[k]
        if os.path.isdir(v):
            v = os.path.normpath(v)
            if k in D and D[k] != v:
                print('warning, CSIDL also exists for key: %s, take os.environ value: %s'% (k, v))
            D[k] = v
    if isinstance(fillRecentEnv, dict):
        recentEnv.update(D)
    return D

#def setInRecentEnv(key, value):
#    if key in recentEnv:
#        if recentEnv[key] == value:
#            print 'already set (the same): %s, %s'% (key, value)
#        else:
#            print 'already set (but different): %s, %s'% (key, value)
#        return
#    print 'setting in recentEnv: %s to %s'% (key, value)
#    recentEnv[key] = value

def substituteEnvVariableAtStart(filepath, envDict=None): 
    r"""try to substitute back one of the (preused) environment variables back

    into the start of a filename

    if ~ (HOME) is D:\My documents,
    the path "D:\My documents\folder\file.txt" should return "~\folder\file.txt"

    pass in a dict of possible environment variables, which can be taken from recent calls, or
    from  envDict = getAllFolderEnvironmentVariables().

    Alternatively you can call getAllFolderEnvironmentVariables once, and use the recentEnv
    of this module! getAllFolderEnvironmentVariables(fillRecentEnv)

    If you do not pass such a dict, recentEnv is taken, but this recentEnv holds only what has been
    asked for in the session, so no complete list!

    """
    if envDict is None:
        envDict = recentEnv
    Keys = list(envDict.keys())
    # sort, longest result first, shortest keyname second:
    decorated = [(-len(envDict[k]), len(k), k) for k in Keys]
    decorated.sort()
    Keys = [k for (dummy1,dummy2, k) in decorated]
    for k in Keys:
        val = envDict[k]
        if filepath.lower().startswith(val.lower()):
            if k in ("HOME", "PERSONAL"):
                k = "~"
            else:
                k = "%" + k + "%"
            filepart = filepath[len(val):]
            filepart = filepart.strip('/\\ ')
            return os.path.join(k, filepart)
    # no hit, return original:
    return filepath
       
def expandEnvVariableAtStart(filepath, envDict=None): 
    """try to substitute environment variable into a path name

    """
    filepath = filepath.strip()

    if filepath.startswith('~'):
        folderpart = getExtendedEnv('~', envDict)
        filepart = filepath[1:]
        filepart = filepart.strip('/\\ ')
        return os.path.normpath(os.path.join(folderpart, filepart))
    if reEnv.match(filepath):
        envVar = reEnv.match(filepath).group(1)
        # get the envVar...
        try:
            folderpart = getExtendedEnv(envVar, envDict)
        except ValueError:
            print('invalid (extended) environment variable: %s'% envVar)
        else:
            # OK, found:
            filepart = filepath[len(envVar)+1:]
            filepart = filepart.strip('/\\ ')
            return os.path.normpath(os.path.join(folderpart, filepart))
    # no match
    return filepath
    
def expandEnvVariables(filepath, envDict=None): 
    """try to substitute environment variable into a path name,

    ~ only at the start,

    %XXX% can be anywhere in the string.

    """
    filepath = filepath.strip()
    
    if filepath.startswith('~'):
        folderpart = getExtendedEnv('~', envDict)
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
                    folderpart = getExtendedEnv(part, envDict)
                except ValueError:
                    folderpart = part
                List2.append(folderpart)
            else:
                List2.append(part)
        filepath = ''.join(List2)
        return os.path.normpath(filepath)
    # no match
    return filepath

def printAllEnvVariables():
    for k in sorted(recentEnv.keys()):
        print("%s\t%s"% (k, recentEnv[k]))

if __name__ == "__main__":
    Vars = getAllFolderEnvironmentVariables()
    for kk in sorted(Vars):
        print('%s: %s'% (kk, Vars[kk]))
        if not os.path.isdir(Vars[kk]):
            print('----- not a directory: %s (%s)'% (Vars[kk], kk))

    print('testing       expandEnvVariableAtStart')
    print('also see expandEnvVar in natlinkstatus!!')
    for p in ("D:\\natlink\\unimacro", "~/unimacroqh",
              "%HOME%/personal",
              "%WINDOWS%\\folder\\strange testfolder"):
        expanded = expandEnvVariableAtStart(p)
        print('expandEnvVariablesAtStart: %s: %s'% (p, expanded))
    print('testing       expandEnvVariables')  
    for p in ("%DROPBOX%/QuintijnHerold/jachthutten", "D:\\%username%", "%NATLINK%\\unimacro", "%UNIMACROUSER%",
              "%HOME%/personal", "%HOME%", "%personal%"
              "%WINDOWS%\\folder\\strange testfolder"):
        expanded = expandEnvVariables(p)
        print('expandEnvVariables: %s: %s'% (p, expanded))

    # testIniSection = NatlinkstatusInifileSection()
    # print testIniSection.keys()
    # testIniSection.set("test", "een test")
    # testval = testIniSection.get("test")
    # print 'testval: %s'% testval
    # testIniSection.delete("test")
    # testval = testIniSection.get("test")
    # print 'testval: %s'% testval
    print('recentEnv: %s'% len(recentEnv))
    np = getExtendedEnv("NOTEPAD")
    print(np)
    for lName in ['Snelle toegang', 'Quick access', 'Documenten', 'Documents', 'Muziek', 'Afbeeldingen', 'Dropbox', 'OneDrive', 'Desktop', 'Bureaublad']:
        f = getFolderFromLibraryName(lName)
        print('folder from library name %s: %s'% (lName, f))
