"""
This module tests the actions module (the "unimacro actions") with pytest

Quintijn Hoogenboom, 2021/2022 
"""
import time
from pathlib import Path
import pytest

# from dtactions import natlinkclipboard
# from dtactions import unimacroactions as actions  # old style
from dtactions import unimacroactions as actions
from dtactions import unimacroutils

try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print('Run this module after "build_package" and "flit install --symlink"\n')
    raise

thisDir = getThisDir(__file__)
dtactionsDir = thisDir.parent

dataDirDtactions = Path.home()/".dtactions"
dataDir = dataDirDtactions/'unimacro'
checkDirectory(dataDir) 

def test_simple():
    """only testing an empty action and T (true) and F (false), but with 1 as positive test result...Hello world
    """
    # this is when we go to class:
    # act = actions.Action()

    # now old module:Hello world
    act = actions
    result = act.doAction('T')
    assert result in (1, True)
    result = act.doAction('F')
    assert result is None
    
def xxx_test_keystroke():
    """test simple keystrokes and recover with clipboard
    """
    act = actions
    keystroke = act.doKeystroke
    keystroke("Hello world")
    keystroke('{shift+left 11}{ctrl+x}')
    time.sleep(0.01)
    result = unimacroutils.getClipboard()
    #Hello world
    assert result == 'Hello world'
#
def test_save_return_clipboard():
    """test if the previous clipboard is collected afterwards
    """
    act = actions
    action = act.doAction
    keystroke = act.doKeystroke
    keystroke("{shift+up 3}{ctrl+c}")
    time.sleep(0.1)
    keystroke("{down 3}")
    start_text = unimacroutils.getClipboard()
    action("CLIPSAVE")

    keystroke("Hello world")
    keystroke('{shift+left 11}{ctrl+x}')
    time.sleep(0.1)
    result = unimacroutils.getClipboard()
    assert result == "Hello world"
    action("CLIPRESTORE")
    time.sleep(0.1)
    end_text = unimacroutils.getClipboard()
    assert start_text == end_text
    
if __name__ == "__main__":
    pytest.main(['test_actions.py'])
    #
