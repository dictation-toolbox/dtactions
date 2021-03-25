"""dtactions __init__

including utility functions,
    - getThisDir, for getting the calling directory of module (in site-packages, also when this is a valid symlink),
    - findInSitePackages, supporting getThisDir
    -checkDirectory, check existence of directory, with create=True, do create if directory does not exist yet.
    
Note: -as user, having pipped the package, the scripts run from the site-packages directory,
       no editing is meant to be done
      -as developer, you have to clone the package, then `build_package` and,
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

thisDir = getThisDir(__file__)
```
"""

__version__ = '1.1.0'  # Quintijn to test

import sys
from pathlib import Path

def getThisDir(fileOfModule):
    """get directory of calling module, if possible in site-packages
    
    call at top of module with "getThisDir(__file__)
    
    Check for symlink and presence in site-packages directory (in this case work is done on this repository)
    """
    thisFile = Path(fileOfModule)
    thisDir = thisFile.parent
    thisDir = findInSitePackages(thisDir)
    return thisDir

def findInSitePackages(cloneDir):
    """get corresponding directory in site-packages 
    
    This directory should be a symlink, otherwise there was no "flit install --symlink" yet.
    
    GOOD: When the package is "flit installed --symlink", so you can work in your clone and
    see the results happen in the site-packages directory. Only for developers
    
    If not found, return the input directory (cloneDir)
    If not "coupled" return the input directory, but issue a warning
    """
    cloneDirStr = str(cloneDir)
    if cloneDirStr.find('\\src\\') < 0:
        return cloneDir
        # raise IOErrorprint(f'This function should only be called when "\\src\\" is in the path')
    commonpart = cloneDirStr.split('\\src\\')[-1]
    spDir = Path(sys.prefix, 'Lib', 'site-packages', commonpart)
    if spDir.is_dir():
        spResolve = spDir.resolve()
        if spResolve == spDir:
            print(f'corresponding site-packages directory is not a symlink: {spDir}.\nPlease use "flit install --symlink" when you want to test this package')
        elif spResolve == cloneDir:
            # print(f'directory is symlink: {spDir} and resolves to {cloneDir} all right')
            ## return the symbolic link in the site-packages directory, that resolves to cloneDir!!
            return spDir
        else:
            print(f'directory is symlink: {spDir} but does NOT resolve to {cloneDir}, but to {spResolve}')
    else:
        print('findInSitePackages, not a valid directory in site-packages, no "flit install --symlink" yet: {spDir}')
    return cloneDir        

def checkDirectory(newDir, create=None):
    """check existence of directory path
    
    create if not existent yet... if create == True
    if create == False, raise an error if directory is not there
    raise IOError if something strange happens...
    """
    if not isinstance(newDir, Path):
        newDir = Path(newDir)
    if newDir.is_dir():
        return
    elif create is False:
        raise OSError(f'Cannot find directory {newDir}, but it should be there.')
    if newDir.exists():
        raise OSError(f'path exists, but is not a directory: {newDir}')
    newDir.mkdir(parents=True)
    if newDir.is_dir():
        print('created directory: {newDir}')
    else:
        raise OSError(f'did not manage to create directory: {newDir}')
                      