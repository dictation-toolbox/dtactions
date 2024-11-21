"""dtactions __init__

    
Note: -as user, having pipped the package, the scripts run from the site-packages directory,
          no editing in source files is meant to be done
      -as developer, you need to clone the package, 
       after a `pip uninstall dtactions` do: `pip install -e .` from the root of this project.


You can retrieve the DtactionsDirectory and DtactionsUserDirectory from here:

```
from dtactions import getDtactionsDirectory, getDtactionsUserDirectory
    print(f'dtactions_directory: "{getDtactionsDirectory()}"')
    print(f'dtactions_user_directory: "{getDtactionsUserDirectory()}"')
 
```

Or with the "Path" versions:

```
from dtactions import getDtactionsPath, getDtactionsUserPath
    print(f'dtactions_user_path: "{getDtactionsUserPath()}", (type: "{Type}"')
    Type = type(getDtactionsPath())
    print(f'dtactions_path: "{getDtactionsPath()}", (type: "{Type}")')
```


"""


##----------------------
import os
from pathlib import Path, WindowsPath
# 
# def get_home_path() -> Path:
#     """get home path, can be tweaked by pytest testing
#     """
#     return WindowsPath.home()

def getDtactionsDirectory() -> str:
    """return the root directory of dtactions
    """
    return str(Path(__file__).parent)

def getDtactionsPath() -> Path:
    """return the root directory of dtactions as Path instance
    """
    return Path(__file__).parent

def getDtactionsUserDirectory() -> str:
    """get the UserDirectory for Dtactions

    This is by default your-home-directory/.dtactions (so decoupled from the Natlink user directory)

    You can override this by setting environment variable `DTACTIONS_USERDIR`
    to a valid directory of your choice

    """
    # default dtHome:
    dta_user_dir = os.getenv("DTACTIONS_USERDIR")
    if dta_user_dir:
        if Path(dta_user_dir).is_dir():
            return str(dta_user_dir)
        print(f'WARNING, dtactions.getDtactionsUserDirectory: environment variable DTACTIONS_USERDIR does not point to a valid directory: "{dta_user_dir}"')

    home_path = WindowsPath.home()

    dtactions_ini_path = home_path/".dtactions"
    if not dtactions_ini_path.is_dir():
        dtactions_ini_path.mkdir()   #make it if it doesn't exist
    if not dtactions_ini_path.is_dir():
        raise IOError(f'dtactions.__init__: dtactions_ini_path cannot be created "{dtactions_ini_path}"' )
    return str(dtactions_ini_path)

def getDtactionsUserPath() -> Path:
    """the "Path" version of getDtactionsUserDirectory
    """
    return Path(getDtactionsUserDirectory())


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
    print(f'dtactions_directory: "{getDtactionsDirectory()}"')
    print(f'dtactions_user_directory: "{getDtactionsUserDirectory()}"')
 
    Type = type(getDtactionsPath())
    print(f'dtactions_path: "{getDtactionsPath()}", (type: "{Type}")')
 
    Type = type(getDtactionsUserPath())
    print(f'dtactions_user_path: "{getDtactionsUserPath()}", (type: "{Type}"')
    
