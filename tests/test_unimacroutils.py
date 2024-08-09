"""
This module tests the unimacro utils module 

Quintijn Hoogenboom, 2022
"""
import time
from pathlib import Path
import pytest
import win32gui
# from dtactions import natlinkclipboard
# from dtactions.unimacro import unimacroinivars as inivars  # old style
from dtactions.unimacro import unimacroutils


thisDir = Path(__file__).parent
dtactionsDir = thisDir.parent


def test_classname():
    """testing getting the classname from a window handle
    """
    hndle = win32gui.GetForegroundWindow()
    assert isinstance(hndle, int)
    assert hndle > 0
    classname = win32gui.GetClassName(hndle)
    assert isinstance(classname, str)
    assert len(classname) > 0
    
def test_getProgInfo():
    """testing the getProgInfo function
    """
    progInfo = unimacroutils.getProgInfo()
    print(f'progInfo: {progInfo}')
    assert len(progInfo) == 6
    assert isinstance(progInfo.hndle, int)
    
    
    
if __name__ == "__main__":
    pytest.main(['test_unimacroutils.py'])
