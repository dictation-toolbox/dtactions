"""
This module tests the inivars module (the "unimacro inivars") with pytest

Quintijn Hoogenboom, 2021/2022 
"""
import time
from pathlib import Path
import pytest

# from dtactions import natlinkclipboard
# from dtactions.unimacro import unimacroinivars as inivars  # old style
from dtactions.unimacro import inivars

try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print('Run this module after "build_package" and "flit install --symlink"\n')
    raise

thisDir = getThisDir(__file__)
dtactionsDir = thisDir.parent

testDir = Path(thisDir)/'test_inivarsfiles'

def test_inifile():
    """testing a breaking inifile
    """
    testFile = '_brackets.ini'
    testPath = testDir/testFile
    ini = inivars.IniVars(testPath)
    assert isinstance(ini, inivars.IniVars)
    
    
    
if __name__ == "__main__":
    pytest.main(['test_inivars.py'])
