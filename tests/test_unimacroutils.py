"""
This module tests the unimacro utils module

Note: the important functions `getProgInfo` and `getModInfo` are using  autohotkey (if installed) when
Dragon is not running, and the call to `natlink.getCurrentModule()` throws an error.

Quintijn Hoogenboom, 2022/2024
"""
import sys
import time
from pathlib import Path
import pytest
import win32gui
# from dtactions import natlinkclipboard
# from dtactions.unimacro import unimacroinivars as inivars  # old style
import natlink
from dtactions.unimacro import unimacroutils
from dtactions import autohotkeyactions


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
    if natlink.isNatSpeakRunning():
        progInfoAhk = autohotkeyactions.getProgInfo()
        assert progInfo == progInfoAhk
    else:
        print('test_getProgInfo: cannot test equal result of autohotkeyactions.getProgInfo, and via natlink.getModInfo, because Dragon is not running')

def test_mouse_move(nat_conn):
    """test moving the mouse, also with Dragon16
    """
    xp, yp = 500, 200
    unimacroutils.set_mouse_position(xp, yp)
    xpn, ypn = unimacroutils.getMousePosition()
    xpnl, ypnl = natlink.getCursorPos()
    assert (xpnl, ypnl) == (xpn, ypn)
      
    assert (xp, yp) == (xpn, ypn)
   
def test_getMousePositions(nat_conn):
    """test the function which prints the mouse positions
    """
    xp, yp = 500, 200
    unimacroutils.set_mouse_position(xp, yp)
    absolute = 0
    result = unimacroutils.getMousePositions(absolute)
    assert result == 'abc'

def test_myoutput(capsys, nat_conn): 
    print("hello")
    sys.stderr.write("world\n")
    captured = capsys.readouterr()
    assert captured.out == "hello\n"
    assert captured.err == "world\n"
    print("next")
    captured = capsys.readouterr()
    assert captured.out == "next\n"
    
if __name__ == "__main__":
    pytest.main(['test_unimacroutils.py::test_getMousePositions'])
