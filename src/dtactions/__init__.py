"""dtactions"""

__version__ = '1.1.0'  # Quintijn to test

# print(f' starting __init__ of dtactions version: {__version__}')

import sys
import os
import os.path

def getThisDir(fileOfModule):
    """get directory of calling module, if possible in site-packages
    
    call at top of module with "getThisDir(__file__)
    
    Check for symlink and presence in site-packages directory (in this case work is done on this repository)
    """
    thisFile = fileOfModule
    thisDir = os.path.split(thisFile)[0]
    if findInSitePackages:
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
    cloneDir = str(cloneDir)
    if cloneDir.find('\\src\\') < 0:
        return cloneDir
        # raise IOErrorprint(f'This function should only be called when "\\src\\" is in the path')
    commonpart = cloneDir.split('\\src\\')[-1]
    spDir = os.path.join(sys.prefix, 'Lib', 'site-packages', commonpart)
    if os.path.isdir(spDir):
        spResolve = os.path.realpath(spDir)
        if spResolve == spDir:
            print(f'corresponding site-packages directory is not a symlink: {spDir}.\nPlease use "flit install --symlink" when you want to test this package')
        elif spResolve == cloneDir:
            # print(f'directory is symlink: {spDir} and resolves to {cloneDir} all right')
            return spDir
        else:
            print(f'directory is symlink: {spDir} but does NOT resolve to {cloneDir}, but to {spResolve}')
    else:
        print('findInSitePackages, not a valid directory in site-packages, no "flit install --symlink" yet: {spDir}')
    return cloneDir        

def checkDirectory(newDir, create=True):
    """check existence of directory path
    
    create if not existent yet... if create == True
    if create == False, raise an error if directory is not there
    raise IOError if something strange happens...
    """
    if os.path.isdir(newDir):
        return
    elif create is False:
        raise OSError(f'Cannot find directory {newDir}, but it should be there.')
    if os.path.exists(newDir):
        raise OSError(f'path exists, but is not a directory: {newDir}')
    os.makedirs(newDir)
    if os.path.isdir(newDir):
        print('created directory: {newDir}')
    else:
        raise OSError(f'did not manage to create directory: {newDir}')
                      