"""dtactions __init__

including utility functions,
    - getThisDir, for getting the calling directory of module (in site-packages, also when this is a valid symlink),
    - findInSitePackages, supporting getThisDir
    - checkDirectory, check existence of directory, with create=True, do create if directory does not exist yet.
    
Note: -as user, having pipped the package, the scripts run from the site-packages directory,
          no editing in source files is meant to be done
      -as developer, you need to clone the package, then `build_package` and,
       after a `pip uninstall dtactions` do: `flit install --symlink`.
       When you edit a file, either via the site-packages (symlinked) directory or via the cloned directory,
       the changes will come in the cloned file, so can be committed again.
       See more instructions in the file README.md in the source directory of the package.

Start with the following lines near the top of your python file:

```
try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print(f'Run this module after "build_package" and "flit install --symlink"\n')
    raise
   
thisDir = getThisDir(__file__) # optional: ... , warnings=True)
```

Also retrieve the DtactionsDirectory and DtactionsUserDirectory from here:

```
from Dtactionscore.__init__ import getDtactionsDirectory, getDtactionsUserDirectory

print(f'DtactionsDirectory: {getDtactionsDirectory()})
print(f'DtactionsUserDirectory: {getDtactionsUserDirectory()})
```

NOTE: these two directories can also be obtained via Dtactionsstatus:
```
import natlinkstatus
status = natlinkstatus.DtactionsStatus()
status.getDtactionsDirectory()
status.getDtactionsUserDirectory()
```

"""

## version to be updated when a new release is sent to pypi:
__version__ = '1.5.3'     # august 2022
#__version__ = '1.5.0'    # work in progress, released wxpython for python > 3.8, dragonfly is dependency
#             '1.4.2'     # work in progress, sendkeys reasonably ok, unimacro actions and clipboard not yet...
# __version__ = '1.3.5'   # working on path details, with HOME or DICTATIONTOOLBOXUSER
# __version__ = '1.3.3'   # setting the user directory
                          # adding getDtactionsDirectory and getDtactionsUserDirectory
##----------------------
import sys
import os
from pathlib import Path, WindowsPath

def getDtactionsDirectory():
    """return the root directory of dtactions
    """
    return getThisDir(__file__)

def getDtactionsUserDirectory():
    """get the NatlinkUserDirectory
    
    Here are config files, especially .natlink
    
    By default the users home directory is taken. This directory can be overridden by setting
    the environment variable DICTATIONTOOLBOXUSER to an existing directory.
    Restart then the calling program
    """
    # default dtHome:
    dictation_toolbox_user = os.getenv("DICTATIONTOOLBOXUSER")
    if dictation_toolbox_user:
        dtHome = Path(dictation_toolbox_user)
        if not dtHome.is_dir():
            dtHome = WindowsPath.home()
            print(f'dtactions.getDtactionsUserDirectory: environment variable DICTATIONTOOLBOXUSER does not point to a valid directory: "{dictation_toolbox_user}", take "{dtHome}"')
    else:
        dtHome = WindowsPath.home()

    dtactions_ini_folder = dtHome / ".dtactions"
    if not dtactions_ini_folder.is_dir():
        dtactions_ini_folder.mkdir()   #make it if it doesn't exist
    return str(dtactions_ini_folder)

def getThisDir(fileOfModule, warnings=False):
    """get directory of calling module, if possible in site-packages
    
    call at top of module with `getThisDir(__file__)`
    
    If you want to get warnings (each one only once, pass `warnings = True`)
    
    More above and in the explanation of findInSitePackages.
    """
    thisFile = Path(fileOfModule)
    thisDir = thisFile.parent
    thisDir = findInSitePackages(thisDir, warnings=warnings)
    return thisDir

def findInSitePackages(directory, warnings):
    """get corresponding directory in site-packages 
    
    For users, just having pipped this package, the "directory" is returned, assuming it is in
    the site-packages.
    
    For developers, the directory is either
    --in a clone from github.
        The corresponding directory in site-packages should be a symlink,
        otherwise there was no "flit install --symlink" yet.
    --a directory in the site-packages. This directory should be a symlink to a cloned dir.
    
    The site-packages directory is returned, but the actual files accessed are in the cloned directory.
    
    To get this "GOOD" situation, you perform the steps as pointed out above (or in the README.md file)

    When problems arise, set warnings=True, in the call, preferably when calling getThisDir in the calling module.
    """
    dirstr = str(directory)
    if dirstr.find('\\src\\') < 0:
        if warnings:
            warning(f'directory {dirstr} not connected to a github clone directory, changes will not persist across updates...')
        return directory

    commonpart = dirstr.rsplit('\\src\\', maxsplit=1)[-1]
    spDir = Path(sys.prefix, 'Lib', 'site-packages', commonpart)
    if spDir.is_dir():
        spResolve = spDir.resolve()
        if spResolve == spDir:
            if warnings:
                warning(f'corresponding site-packages directory is not a symlink: {spDir}.\nPlease use "flit install --symlink" when you want to test this package')
        elif spResolve == directory:
            # print(f'directory is symlink: {spDir} and resolves to {directory} all right')
            ## return the symbolic link in the site-packages directory, that resolves to directory!!
            return spDir
        else:
            if warnings:
                warning(f'directory is symlink: {spDir} but does NOT resolve to {directory}, but to {spResolve}')
    else:
        if warnings:
            warning('findInSitePackages, not a valid directory in site-packages, no "flit install --symlink" yet: {spDir}')
    return directory        

def checkDirectory(directory, create=None):
    """check existence of directory path
    
    create == False, None (default): raise OSError if directory is not there

    create == True: create if not existent yet... 
              raise OSError if something strange happens...
    returns None
    """
    if not isinstance(directory, Path):
        directory = Path(directory)
    if directory.is_dir():
        return
    if create is False:
        raise OSError(f'Cannot find directory {directory}, but it should be there.')
    if directory.exists():
        raise OSError(f'path exists, but is not a directory: {directory}')
    directory.mkdir(parents=True)
    if directory.is_dir():
        print('created directory: {directory}')
    else:
        raise OSError(f'did not manage to create directory: {directory}')

warningTexts = []
def warning(text):
    """print warning only once, if warnings is set!
    
    warnings can be set in the calling functions above...
    """
    textForeward = text.replace("\\", "/")
    if textForeward in warningTexts:
        return
    warningTexts.append(textForeward)
    print(text)
    

if __name__ == "__main__":
    print(f'dtactions_user_directory: "{getDtactionsUserDirectory()}"')
    
