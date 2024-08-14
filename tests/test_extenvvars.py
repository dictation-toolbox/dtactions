"""
This module tests the extenvvars module, getting "extended environment variables"

The module is in dtactions/unimacro.


Quintijn Hoogenboom, Februari 2023/August 2024
"""
#pylint:disable = W0621, E1101
from os.path import isdir
from pathlib import Path
import pytest
from dtactions.unimacro import extenvvars
import natlink
from natlinkcore import natlinkstatus
status = natlinkstatus.NatlinkStatus()
thisDir = Path(__file__).parent
dtactionsDir = thisDir.parent


@pytest.fixture()
def envvars():
    return extenvvars.ExtEnvVars()


def testStartOfInstance(envvars):
    """instance of ExtEnvVars starts with persistent dict of recentEnv entries, and adds at startup "~"

    """
    assert len(envvars.recentEnv) == 2
    assert '~' in envvars.recentEnv
    assert 'HOME' in envvars.recentEnv
    ## the two values are equal...
    assert len(set(envvars.recentEnv.values())) == 1    
    assert isdir(envvars.homeDir)



@pytest.mark.parametrize("var_name",
                ["natlink", "natlinkcore", "NATLINKCORE", "Natlinkcore", "dtactions"]
                         )
def test_getNatlinkEnvVariables(envvars, var_name):
    """this tests if the natlink folders can be got with this function
    
    If "Dir" or "Directory" is missing, try to add:
        eg "AhkUser" is expanded to "AhkuserDir" (equivalent to "AhkUserDir")
        and "Unimacro" is expanded to "UnimacroDirectory"
        
    The variables are case insensitive and are converted to eg "DtactionsDirectory"
    
    """
    result = envvars.getExtendedEnv(var_name)
    if natlink.isNatSpeakRunning():
        assert len(result)
        assert Path(result).is_dir()
    else:
        # variables are not available when natspeak is not running
        assert not result


@pytest.mark.parametrize("var_name",
                ["unimacrouser", "unimacrodata", "unimacrogrammars"]
                        )
def test_getUnimacroEnvironmentVariables(envvars, var_name):
    """testing the unimacro variables, Unimacro can be on or off
    
    When Unimacro is not enabled (but possibly installed), the results should be fals
        
    """
    result = envvars.getExtendedEnv(var_name)
    if natlink.isNatSpeakRunning() and status.unimacroIsEnabled():
        assert len(result)
        assert Path(result).is_dir()
    else:
        # variables are not available when unimacro is not available
        assert not result

# @pytest.mark.parametrize("var_name",
#                 ["unimacro"]
#                         )
# def test_getUnimacroEnvironmentVariablesAlwaysAvailable(envvars, var_name):
#     """testing the unimacro variables, Unimacro can be on or off
#     
#     When Unimacro is off, but Unimacro has been enabled, the result can be a valid directory.
#     
#     When Dragon is not running...
#
#     Too confusing when the unimacro directory is shown. When it is not installed, it is not...
#     
#     """
#     result = envvars.getExtendedEnv(var_name)
#     if natlink.isNatSpeakRunning():
#         assert len(result)
#         assert Path(result).is_dir()
#     else:
#         assert not result
        
        
        
def test_getAllFolderEnvironmentVariables(envvars):
    """get the dict that this function gets, and test a few values
    """
    D = envvars.getAllFolderEnvironmentVariables()
    assert isinstance(D, dict )
    assert D['COMMON_APPDATA'] == r'C:\ProgramData'
    assert D['COMMON_STARTMENU'] == r'C:\ProgramData\Microsoft\Windows\Start Menu'
    

@pytest.mark.parametrize("var_name",
                ["%HOME%", "%SYSTEM%"]
                        )
def test_otherEnvVariables(envvars, var_name):
    """this tests if "other" env variables are correctly reported
    
    Note the precise capitalisation, and English and Dutch names are recognised here.
    (When other languages are wanted, add them in extenvvars.py around line 110)
    
    """
    
    result = envvars.getExtendedEnv(var_name)
    assert len(result) > 0
    assert isdir(result)

    
@pytest.mark.parametrize("first_name, alternative_name",
                [ ("Music", "Muziek"),
                  ("Desktop", "Bureaublad"),
                  ("Downloads", ""),
                  ("Documents", "Documenten"),
                  ("Pictures", "Afbeeldingen"),
                  ("~", "HOME"),
                  ("Dropbox", ""),
                  ("OneDrive", ""),
                  ]
                        )
def test_windows_Library_dirs(envvars, first_name, alternative_name):
    """most of them with platformdirs.windows, some special treatment.
    
    These are directory names, that are "shorthand", like Music, Desktop.
    
    All go to getFolderFromLibraryName, and names are currently from English or Dutch windows system. 
    """
    result1 = envvars.getFolderFromLibraryName(first_name)    
    assert len(result1) 
    assert Path(result1).is_dir()
    if alternative_name:
        result2 = envvars.getFolderFromLibraryName(alternative_name)    
        assert result2
        assert result2 == result1

    ## Quick access
    # result = envvars.getFolderFromLibraryName('Quick access')    
    # assert len(result) 
    # assert Path(result).is_dir()
    # result2 = envvars.getFolderFromLibraryName('Snelle toegang')    
    # assert result2 == result
    
    ## special cases:
    home_folder = envvars.getFolderFromLibraryName('~')    
    documents_folder = envvars.getFolderFromLibraryName('Documents')
    personal_folder = envvars.getExtendedEnv('PERSONAL')
    assert documents_folder == personal_folder
    assert documents_folder.startswith(home_folder)
    appdata_folder = envvars.getFolderFromLibraryName('APPDATA')
    assert appdata_folder == str(Path(home_folder)/'AppData'/'Roaming')
    local_appdata_folder = envvars.getFolderFromLibraryName('LOCAL_APPDATA')
    assert local_appdata_folder.startswith(home_folder)
    common_appdata_folder = envvars.getFolderFromLibraryName('COMMON_APPDATA')
    assert Path(common_appdata_folder).is_dir()

  # for lName in ['Snelle toegang', 'Quick access', 'Dropbox', 'OneDrive'']:
    #     f = getFolderFromLibraryName(lName)
  
def main():
    pytest.main(["test_extenvvars.py"])
    
if __name__ == "__main__":
    main()
    