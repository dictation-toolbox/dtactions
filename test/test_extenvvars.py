"""
This module tests the extenvvars module, getting "extended environment variables"

The module is in dtactions/unimacro.

Quintijn Hoogenboom, Februari 2023
"""
from os.path import isdir
from pathlib import Path
import pytest
from dtactions.unimacro import extenvvars

thisDir = Path(__file__).parent
dtactionsDir = thisDir.parent

def testStartOfInstance():
    """instance of ExtEnvVars starts with persistent dict of recentEnv entries, and adds at startup "~"

    """
    ext1 = extenvvars.ExtEnvVars()
    assert len(ext1.recentEnv) == 2
    assert isdir(ext1.homeDir)


def test_getNatlinkEnvVariables():
    """this tests if the natlink folders can be got with this function
    
    If "Dir" or "Directory" is missing, try to add:
        eg "AhkUser" is expanded to "AhkUserDir"
        (Ahk can be missing, if not installed)
        
        and "Unimacro" is expanded to "UnimacroDirectory"
    
    """
    ext = extenvvars.ExtEnvVars()    
    result = ext.getExtendedEnv("%Unimacro%")    
    assert len(result)
    assert Path(result).is_dir()
    # result = ext.getExtendedEnv("%VocolaUserDirectory%")    
    # assert Path(result).is_dir()
    # assert len(result)
    
    # AHK maybe not installed:
    result1 = ext.getExtendedEnv("%AhkUser%")    
    if result1:
        assert Path(result1).is_dir()
    result2 = ext.getExtendedEnv("%AhkUserDir%")
    if result2:
        assert Path(result2).is_dir()

def test_otherEnvVariables():
    """this tests if "other" env variables are correctly reported
    
    """
    ext = extenvvars.ExtEnvVars()    
    result = ext.getExtendedEnv("%HOME%")    
    assert len(result) > 0
    assert isdir(result)
    
    result = ext.getExtendedEnv("%SYSTEM%")    
    assert len(result) > 0
    assert isdir(result)
    
def test_windows_Library_dirs():
    """most of them with platformdirs.windows, some special treatment.
    
    These are directory names, that are "shorthand", like Music, Desktop.
    
    All go to getFolderFromLibraryName, and names are currently from English or Dutch windows system. 
    """
    ## Music
    ext = extenvvars.ExtEnvVars()
    result = ext.getFolderFromLibraryName('Muziek')    
    assert len(result) 
    assert Path(result).is_dir()
    result2 = ext.getFolderFromLibraryName('Music')    
    assert result2 == result

    ## Desktop
    result = ext.getFolderFromLibraryName('Desktop')    
    assert len(result) 
    assert Path(result).is_dir()
    result2 = ext.getFolderFromLibraryName('Bureaublad')    
    assert result2 == result

    ## Downloads
    result = ext.getFolderFromLibraryName('Downloads')    
    assert len(result) 
    assert Path(result).is_dir()

    ## Documents (same as PERSONAL!)
    result = ext.getFolderFromLibraryName('Documents')    
    assert len(result) 
    assert Path(result).is_dir()
    result2 = ext.getFolderFromLibraryName('Documenten')  ## Dutch    
    assert result2 == result

    ## Pictures
    result = ext.getFolderFromLibraryName('Pictures')    
    assert len(result) 
    assert Path(result).is_dir()
    result2 = ext.getFolderFromLibraryName('Afbeeldingen')    
    assert result2 == result

    # ## Quick access
    # result = ext.getFolderFromLibraryName('Quick access')    
    # assert len(result) 
    # assert Path(result).is_dir()
    # result2 = ext.getFolderFromLibraryName('Snelle toegang')    
    # assert result2 == result
    
    ## special cases:
    home_folder = ext.getFolderFromLibraryName('~')    
    documents_folder = ext.getFolderFromLibraryName('Documents')
    personal_folder = ext.getExtendedEnv('PERSONAL')
    assert documents_folder == personal_folder
    assert documents_folder.startswith(home_folder)
    appdata_folder = ext.getFolderFromLibraryName('APPDATA')
    assert appdata_folder == str(Path(home_folder)/'AppData'/'Roaming')
    local_appdata_folder = ext.getFolderFromLibraryName('LOCAL_APPDATA')
    assert local_appdata_folder.startswith(home_folder)
    common_appdata_folder = ext.getFolderFromLibraryName('COMMON_APPDATA')
    assert Path(common_appdata_folder).is_dir()

  # for lName in ['Snelle toegang', 'Quick access', 'Dropbox', 'OneDrive'']:
    #     f = getFolderFromLibraryName(lName)
  
if __name__ == "__main__":
    pytest.main(['test_extenvvars.py'])
