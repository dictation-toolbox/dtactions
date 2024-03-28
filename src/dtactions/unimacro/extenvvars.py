#
# Python Macro Language for Dragon NaturallySpeaking
#   (c) Copyright 1999 by Joel Gould
#   Portions (c) Copyright 1999 by Dragon Systems, Inc.
#
#pylint:disable=E0611, E0401, R0911, R0912, W0702, C0116
#pylint:disable=C0209
"""extenvvars.py

Keep track of environment variables, including added "fake" variables like DROPBOX

 Quintijn Hoogenboom, January 2008/September 2015/November 2022 (python3)/March 24


"""
import os
from os.path import normpath, isfile, isdir, join
import copy
import re
from pathlib import Path
from win32com.shell import shell, shellcon
import platformdirs
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

class ExtEnvVars:
    """gives and "remembers" environment variables, with extensions
    
    several sorts of env variables are handled:
    
    Environment variables, like %xxxx%:
    - os.environ
    - all from natlinkstatus.py (if available) (like %natlink% or %unimacro%)


    Libraries, that are returned from a getFolderFromActiveWindow (_folders grammar of unimacro):
        - several from "Libraries" (like Music, Documents), via platformdirs
        - others from "CSIDL_variables" directly
    """
    
    def __init__(self):
        self.recentEnv = {}
        self.homeDir = self.getHome()        # starting at least reventEnv dict.

    def addToRecentEnv(self, name, value):
        """to be filled for NATLINK variables from natlinkstatus and other env variables, sort of caching mechanism
        
        empty values can be cached...
        """
        if not value:
            self.recentEnv[name] = None
            return
        value = str(value)
        if not isdir(value):
            self.recentEnv[name] = None
            raise ValueError('extenvvars, addToRecentEnv: not a valid path for for name: "%s": "%s"'% (name, value))
        self.recentEnv[name] = value

    def deleteFromRecentEnv(self, name):
        """to possibly delete from recentEnv, from natlinkstatus
        """
        if name in self.recentEnv:
            del self.recentEnv[name]

    def getFolderFromLibraryName(self, name):
        """from windows library names extract the real folder
        
        the CabinetWClass and #32770 controls return "Documents", "Dropbox", "Desktop" etc.
        try to resolve these throug the extended env variables,
        - with platformdirs
        - with more special trick
        """
        if name in self.recentEnv:
            return self.recentEnv[name]
        
        if name == "~":
            home_dir = Path().home()
            self.addToRecentEnv("~", home_dir)
            return self.recentEnv["~"]
        
        
        ## these are NOT cached:
        if name.startswith("Docum"):  # Documents in Dutch and English
            return platformdirs.windows.get_win_folder("CSIDL_PERSONAL")
        if name in ["Muziek", "Music"]:
            return platformdirs.windows.get_win_folder("CSIDL_MYMUSIC")
        if name in ["Pictures", "Afbeeldingen"]:
            return platformdirs.windows.get_win_folder("CSIDL_MYPICTURES")
        if name in ["Videos", "Video's"]:
            return platformdirs.windows.get_win_folder("CSIDL_MYVIDEO")
        if name in ['Desktop', "Bureaublad"]:
            return platformdirs.windows.get_win_folder("CSIDL_DESKTOPDIRECTORY")
        if name in ['Download', 'Downloads']:
            return platformdirs.windows.get_win_folder("CSIDL_DOWNLOADS")
        if name in ['APPDATA']:
            return platformdirs.windows.get_win_folder("CSIDL_APPDATA")
        if name in ['COMMON_APPDATA']:
            return platformdirs.windows.get_win_folder("CSIDL_COMMON_APPDATA")
        if name in ['LOCAL_APPDATA']:
            return platformdirs.windows.get_win_folder("CSIDL_LOCAL_APPDATA")
        
        ## extra cases:
        # if name in ['Quick access', 'Snelle toegang']:
        #     templatesfolder =self. getExtendedEnv('TEMPLATES')
        #     if isdir(templatesfolder):
        #         QuickAccess = normpath(join(templatesfolder, "..", "Libraries"))
        #         if isdir(QuickAccess):
        #             return QuickAccess
        if name.upper() == 'DROPBOX':
            return self.getDropboxFolder()
        if name.upper() in ['ONEDRIVE']:
            return self.getExtendedEnv("OneDrive")
        
        if name in ["This PC", "Deze pc"]:
            return "\\"
        print('getFolderFromLibraryName, cannot find folder for Library name: %s'% name)
        return ""

    def getHome(self):
        """get the home directory, with pathlib.Path().home()
        
        If env variable HOME is there, and points to another directory, give a warning. 
        
        result is cached in recentEnv dict.
        
        """
        if "~" in self.recentEnv:
            return self.recentEnv["~"]

        Home = str(Path().home())
        if Home and isdir(Home):
            self.addToRecentEnv("~", Home)   # cache as a str
        HomeFromEnv = os.environ.get("HOME", "")
        if HomeFromEnv and isdir(HomeFromEnv):
            self.addToRecentEnv("HOME", HomeFromEnv)
            if HomeFromEnv != Home:
                print(f'Warning: "~" and "HOME" point to different directories\n\t"~": "{Home}"\n\t"HOME": "{HomeFromEnv}"')
        else:
            self.addToRecentEnv("HOME", Home)
        return self.recentEnv["~"]
            

    def getDropboxFolder(self):
        """get the dropbox folder, or the subfolder which is specified.
        
        searched is in home directory of the user...
        
        raises OSError if more folders are found (should not happen, I think)
        if containsFolder is not passed, the dropbox main folder is returned
        if containsFolder is passed, this folder is returned if it is found in the dropbox folder
        
        otherwise None is returned.
        """
        if "DROPBOX" in self.recentEnv:
            return self.recentEnv["DROPBOX"]

        dirsToTry = [ "%DROPBOX%", "~", 'C:\\']    #### for s in os.listdir(root) if isdir(join(root,s)) ]
        for root in dirsToTry:
            result = self.getDropboxFolder2(root)
            if result:
                self.addToRecentEnv("DROPBOX", result)
                return self.recentEnv["DROPBOX"]
        return False 
            
    def getDropboxFolder2(self, root):
        """helper function
        """
        
        if root.startswith("%"):
            root = root.strip("% ").upper()
            root = os.environ.get(root, "")
        
        if root and not isdir(root):
            return False
        
        if root.lower().endswith("dropbox"):
            return root
        next_dir = join(root, "dropbox")
        if isdir(next_dir):
            return next_dir
        return False
    
    def matchesStart(self, listOfDirs, checkDir, caseSensitive):
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
        
            

    def getExtendedEnv(self, var, noCache=False, displayMessage=True):
        r"""get from os.environ or windows CSLID
    
        HOME is os.environ['HOME'] or CSLID_PERSONAL
        ~ is HOME (this is doubtful, str(pathlib.Path.home()) gives 'C:\Users\User', also when HOME points to another directory
    
        DROPBOX added via self.getDropboxFolder is this module (QH, December 1, 2018)
    
        Also the directories that are configured inside Natlink can be retrieved, via natlinkstatus.py
        Natlink_SettingsDir, DragonflyUserDirectory, VocolaDirectory, AhkUserDir:
        
        Note: these settings are case sensitive! You can leave out Dir or Directory.
    
        """
        var = var.strip("% ")
    
        if var in  self.recentEnv:
            return self.recentEnv[var]
    
        # Note the Natlink variables must conform to "Unimacro" etc. Best don't append "Dir" or "Directory":
        # It is tried to append ("Dir" or "Directory" (or nothing) to it in the call
        # Don't capitalise these variables!!)
        result = self.getDirectoryFromNatlinkstatus(var)
        if result:
            self.recentEnv[var] = result
            return result
      
        
        
        if var == "~":
            var = 'HOME'
        
        if var in os.environ:
            result = os.environ[var]
            self.addToRecentEnv(var, result)
            return self.recentEnv[var]
    
        if var == 'DROPBOX':
            result = self.getDropboxFolder()
            if result:
                self.addToRecentEnv(var, result)
                return result
            raise ValueError('getExtendedEnv, cannot find path for "DROPBOX"')
    
        if var == 'NOTEPAD':
            windowsDir =self. getExtendedEnv("WINDOWS")
            notepadPath = join(windowsDir, 'notepad.exe')
            if isfile(notepadPath):
                return notepadPath
            raise ValueError('getExtendedEnv, cannot find path for "NOTEPAD"')
    
        # try to get from CSIDL system call:
        try:
            CSIDL_variable =  'CSIDL_%s'% var
            shellnumber = getattr(shellcon,CSIDL_variable, -1)
        except:
            print('getExtendedEnv, cannot find in os.environ or CSIDL: "%s"'% var)
            return ''
        if shellnumber < 0:
            # on some systems have SYSTEMROOT instead of SYSTEM:
            if var == 'SYSTEM':
                return self.getExtendedEnv('SYSTEMROOT')
            # raise ValueError('getExtendedEnv, cannot find in os.environ or CSIDL: "%s"'% var2)
        try:
            result = shell.SHGetFolderPath (0, shellnumber, 0, 0)
        except:
            if displayMessage:
                print('getExtendedEnv, cannot find in os.environ or CSIDL: "%s"'% var)
            return ''
        if noCache:
            return result
        self.addToRecentEnv(var, result)
        return self.recentEnv[var]

    def getDirectoryFromNatlinkstatus(self, envvar):
        """see if directory can can be retrieved from envvar
        """
        # if natlink not available:
        if not natlinkAvailable:
            # print(f'natlink not available for get "{envvar}"')
            return None
    
        # try if function in natlinkstatus:
        if not status:
            return None
        for extra in ('', 'Directory', 'Dir'):
            var2 = envvar + extra
            if var2 in status.__dict__:
                funcName = f'get{var2}'
                func = getattr(status, funcName)
                result = func()
                if result:
                    return result
        return None
    
    
    
    def clearRecentEnv(self):
        """for testing, clears above global dictionary
        """
        self.recentEnv.clear()
    
    def getAllFolderEnvironmentVariables(self, displayMessage=False):
        """return, as a dict, all the os.environ AND all CSLID variables that result into a folder
        
        Now also implemented:  Also include NATLINK, UNIMACRO, VOICECODE, DRAGONFLY, VOCOLAUSERDIR, UNIMACROUSERDIR
        This is done by calling from natlinkstatus, see there and example in natlinkmain.
    
        Optionally put them in recentEnv, if you specify fillRecentEnv to 1 (True)
    
        """
        D = {}
    
        for k in dir(shellcon):
            if k.startswith("CSIDL_"):
                kStripped = k[6:]
                try:
                    v =self.getExtendedEnv(kStripped, noCache=True, displayMessage=displayMessage)    ## displayMessage=False)
                except ValueError:
                    continue
                if len(str(v)) > 2 and isdir(v):
                    D[kStripped] = v
        # os.environ overrules CSIDL:
        for k, v in os.environ.items():
            if isdir(v):
                v = normpath(v)
                if k in D and D[k] != v:
                    print('warning, CSIDL also exists for key: %s, take os.environ value: %s'% (k, v))
                D[k] = v
        return D
    def substituteEnvVariableAtStart(self, filepath, envDict=None): 
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
            envDict = self.recentEnv
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
                return join(k, filepart)
        # no hit, return original:
        return filepath
           
    def expandEnvVariableAtStart(self, filepath, envDict=None): 
        """try to substitute environment variable into a path name
    
        """
        filepath = filepath.strip()
    
        if filepath.startswith('~'):
            folderpart =self. getExtendedEnv('~', envDict)
            filepart = filepath[1:]
            filepart = filepart.strip('/\\ ')
            return normpath(join(folderpart, filepart))
        if reEnv.match(filepath):
            envVar = reEnv.match(filepath).group(1)
            # get the envVar...
            try:
                folderpart =self. getExtendedEnv(envVar, envDict)
            except ValueError:
                print('invalid (extended) environment variable: %s'% envVar)
            else:
                # OK, found:
                filepart = filepath[len(envVar)+1:]
                filepart = filepart.strip('/\\ ')
                return normpath(join(folderpart, filepart))
        # no match
        return filepath
        
    def expandEnvVariables(self, filepath, envDict=None): 
        """try to substitute environment variable into a path name,
    
        ~ only at the start,
    
        %XXX% can be anywhere in the string.
    
        """
        filepath = filepath.strip()
        
        if filepath.startswith('~'):
            folderpart =self. getExtendedEnv('~', envDict)
            filepart = filepath[1:]
            filepart = filepart.strip('/\\ ')
            filepath = normpath(join(folderpart, filepart))
        
        if reEnv.search(filepath):
            List = reEnv.split(filepath)
            #print 'parts: %s'% List
            List2 = []
            for part in List:
                if not part:
                    continue
                if part == "~" or (part.startswith("%") and part.endswith("%")):
                    try:
                        folderpart = self.getExtendedEnv(part, envDict)
                    except ValueError:
                        folderpart = part
                    List2.append(folderpart)
                else:
                    List2.append(part)
            List2 = [str(item) for item in List2]
            filepath = ''.join(List2)
            return normpath(filepath)
        # no match
        return filepath
    
    
    
    def printAllEnvVariables(self):
        for k in sorted(self.recentEnv.keys()):
            print("%s\t%s"% (k, self.recentEnv[k]))
    
if __name__ == "__main__":
    env = ExtEnvVars()
    Vars = env.getAllFolderEnvironmentVariables()
    print('recemtEnv: ')
    print(env.recentEnv)
    print('='*80)
    

    print('testing       expandEnvVariableAtStart')
    print('also see expandEnvVar in natlinkstatus!!')
    for p in ("~", "%home%", "D:\\natlink\\unimacro", "~/unimacroqh",
              "%HOME%/personal",
              "%WINDOWS%\\folder\\strange testfolder"):
        expanded = env.expandEnvVariableAtStart(p)
        print('expandEnvVariablesAtStart: %s: %s'% (p, expanded))
    print('testing       expandEnvVariables')  
    for p in ("%NATLINK%\\unimacro", "%DROPBOX%/QuintijnHerold/jachthutten", "D:\\%username%", "%UNIMACROUSER%",
              "%HOME%/personal", "%HOME%", "%personal%"
              "%WINDOWS%\\folder\\strange testfolder"):
        expanded = env.expandEnvVariables(p)
        print('expandEnvVariables: %s: %s'% (p, expanded))

