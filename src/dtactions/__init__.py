"""dtactions __init__

including utility functions,
    - getThisDir, for getting the calling directory of module (in site-packages, also when this is a valid symlink),
    - findInSitePackages, supporting getThisDir
    - checkDirectory, check existence of directory, with create=True, do create if directory does not exist yet.
    
Note: -as user, having pipped the package, the scripts run from the site-packages directory,
          no editing in source files is meant to be done
      -as developer, you need to clone the package, then `build_package` and,
       after a `pip uninstall dtactions` do: `pip install -e .` from the root of this project.

Start with the following lines near the top of your python file:

```
try:
    from dtactions import getThisDir, checkDirectory
except ModuleNotFoundError:
    print(f'Run this module after "build_package" and "flit install --symlink"\n')
    raise
   
thisDir = getThisDir(__file__) # optional: ... , warnings=True)
```

Also retrieve the DtactionsDirectory and DtactionsUserDirectory from here:

```
from dtactions import getDtactionsDirectory, getDtactionsUserDirectory

print(f'DtactionsDirectory: {getDtactionsDirectory()})
print(f'DtactionsUserDirectory: {getDtactionsUserDirectory()})
```

NOTE: these two directories can also be obtained via Dtactionsstatus:
```
from natlinkcore import natlinkstatus
status = natlinkstatus.DtactionsStatus()
status.getDtactionsDirectory()
status.getDtactionsUserDirectory()
```

"""
## version to be updated when a new release is sent to pypi:
__version__ = '1.6.3'     
##----------------------
import os
from pathlib import Path, WindowsPath

def getDtactionsDirectory():
    """return the root directory of dtactions
    """
    return str(Path(__file__).parent)

def getDtactionsUserDirectory() -> str:
    """get the UserDirectory for Dtactions
    
    Here are config files, especially "natlink.ini".
    
    It seems convenient to use as default the ".natlink" directory, as a sub directory of
    the users home directory. (which can be overriden by the NATLINK_SETTINGSDIR
    environment variable).
    
    This directory can be overridden by setting the environment variable DTACTIONS_USERDIR
    to an existing directory.
    """
    # default dtHome:
    dta_user_dir = os.getenv("DTACTIONS_USERDIR")
    if dta_user_dir:
        if Path(dta_user_dir).is_dir():
            return str(dta_user_dir)
        print(f'WARNING, dtactions.getDtactionsUserDirectory: environment variable DTACTIONS_USERDIR does not point to a valid directory: "{dta_user_dir}"')

    dta_home = WindowsPath.home()

    dtactions_ini_path = dta_home/".dtactions"
    if not dtactions_ini_path.is_dir():
        dtactions_ini_path.mkdir()   #make it if it doesn't exist
    if not dtactions_ini_path.is_dir():
        raise IOError(F'dtactions.__init__: dtactions_ini_path cannot be created "{dtactions_ini_path}"' )
    return str(dtactions_ini_path)


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
    print(f'dtactions_directory: "{getDtactionsDirectory()}"')
    
