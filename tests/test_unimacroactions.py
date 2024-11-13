"""
This module tests the unimacroactions module 

Quintijn Hoogenboom, December 2022
"""
from pathlib import Path
import pytest
# from dtactions import natlinkclipboard
# from dtactions import unimacroinivars as inivars  # old style
from dtactions.unimacroactions import *

thisDir = Path(__file__).parent
dtactionsDir = thisDir.parent

def test_dtactions_inifile():
    pass


def test_matchProgTitleWithDict():
    """this tests if a prog and title matches with definitions in a Dict
    
    the Dict is taken from unimacroactions.ini
    section [general], key "child behaves like top" or "top behaves like child"
    """
    # the definition may contain part of the wanted title, but... only in matchPart is True...
    child_behaves_like_top = {"natspeak": ["dragon-balk", "ragonbar"]}
    
    # no match:
    assert matchProgTitleWithDict('prog', 'title', child_behaves_like_top, matchPart=None) is False

    # title must match exact (but case insensitive)
    assert matchProgTitleWithDict('natspeak', 'Dragon-balk', child_behaves_like_top) is True
    assert matchProgTitleWithDict('natspeak', 'Dragonbar', child_behaves_like_top) is False
    assert matchProgTitleWithDict('natspeak', 'Dragonbar', child_behaves_like_top, matchPart=True) is True
    # also good:
    assert matchProgTitleWithDict('natspeak', 'Dragon-balk', child_behaves_like_top, matchPart=True) is True

def test_dtactions_has_unimacroactions_inifile(dtactions_setup):
    """see if a copy of a sample ini file for unimacroactions is copied properly
    """
    dtactions_userdir = dtactions_setup[0]
    assert Path(dtactions_userdir).is_dir()
    pass
# test SCLIP via unimacro/unimacroactions.py direct run.
    
    
if __name__ == "__main__":
    pytest.main(['test_unimacroactions.py'])
