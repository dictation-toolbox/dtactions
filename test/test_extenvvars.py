"""
This module tests the extenvvars module, getting "extended environment variables"

The module is in dtactions/unimacro.

Quintijn Hoogenboom, Februari 2023
"""
from pathlib import Path
import pytest
from dtactions.unimacro import extenvvars

thisDir = Path(__file__).parent
dtactionsDir = thisDir.parent

def test_getNatlinkEnvVariables():
    """this tests if the natlink folders can be got with this function
    
    If "Dir" or "Directory" is missing, try to add:
        eg "AhkUser" is expanded to "AhkUserDir"
        and "Unimacro" is expanded to "UnimacroDirectory"
    
    """
    result = extenvvars.getExtendedEnv("%Unimacro%")    
    assert len(result)
    assert Path(result).is_dir()
    result = extenvvars.getExtendedEnv("%VocolaUserDirectory%")    
    assert Path(result).is_dir()
    assert len(result) 
    result1 = extenvvars.getExtendedEnv("%AhkUser%")    
    assert Path(result1).is_dir()
    assert len(result1) 
    result2 = extenvvars.getExtendedEnv("%AhkUserDir%")    
    assert Path(result2).is_dir()
    assert len(result2)
    assert result1 == result2

def test_otherEnvVariables():
    """this tests if "other" env variables are correctly reported
    
    """
    result = extenvvars.getExtendedEnv("%HOME%")    
    assert len(result)
    assert Path(result).is_dir()
    
    result = extenvvars.getExtendedEnv("%SYSTEM%")    
    assert len(result) 
    assert Path(result).is_dir()
    
    
if __name__ == "__main__":
    pytest.main(['test_extenvvars.py'])
