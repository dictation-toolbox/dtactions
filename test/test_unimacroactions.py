"""
This module tests the unimacroactions module 

Quintijn Hoogenboom, December 2022
"""
from pathlib import Path
import pytest
import win32gui
# from dtactions import natlinkclipboard
# from dtactions.unimacro import unimacroinivars as inivars  # old style
from dtactions.unimacro.unimacroactions import *

thisDir = Path(__file__).parent
dtactionsDir = thisDir.parent


def test_matchProgTitleWithDict():
    """this tests if a prog and title matches with definitions in a Dict
    
    the Dict is taken from unimacroactions.ini
    section [general], key "child behaves like top" or "top behaves like child"
    """
    child_behaves_like_top = {"natspeak": ["dragon-balk", "dragonbar"]}
    
    # no match:
    assert matchProgTitleWithDict('prog', 'title', child_behaves_like_top, matchPart=None) is False

    assert matchProgTitleWithDict('natspeak', 'Dragonbar', child_behaves_like_top) is True
    assert matchProgTitleWithDict('natspeak', 'ragonbar', child_behaves_like_top) is True
    assert matchProgTitleWithDict('natspeak', 'Dragonbar', child_behaves_like_top, matchPart=True) is True
    assert matchProgTitleWithDict('natspeak', 'ragonbar', child_behaves_like_top, matchPart=True) is False


# test SCLIP via unimacro/unimacroactions.py direct run.
    
    
if __name__ == "__main__":
    pytest.main(['test_unimacroutils.py'])
