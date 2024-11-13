r"""
This module tests the basic setup of the dtactions project.

The directory in which user (ini) files can be put, and other things,
should be in ~\.dtactions, or in a directory that is defined by
env variable DTACCTIONS_USERDIR.

Quintijn Hoogenboom, December 2022/October 2024
"""
#pylint:disable = C0415
from pathlib import Path
import pytest

thisDir = Path(__file__).parent
dtactionsDir = thisDir.parent

def test_dtactions_inifile():
    pass


def test_dtactions_userdir_env_var(dtactions_setup_with_env_var):
    """test if env var DTACTIONS_USERDIR gives correct dtactions_user_dir
    """
    dtactions_userdir = dtactions_setup_with_env_var
    print(f'test_dtactions_userdir_env_var, directory: "{dtactions_userdir}"')
    assert Path(dtactions_userdir).is_dir()
    import dtactions
    actual_dta_userdir = dtactions.getDtactionsUserDirectory()
    assert actual_dta_userdir == dtactions_userdir
    

def test_dtactions_test_copy_of_unimacroactions_ini_file(dtactions_setup_with_env_var):
    """test if the unimacroactions.ini file is copied into the dtactions user dir
    
    """
    dtactions_userdir = dtactions_setup_with_env_var
    assert Path(dtactions_userdir).is_dir()
    import dtactions
    
    actual_dta_userdir = dtactions.getDtactionsUserDirectory()
    assert actual_dta_userdir == dtactions_userdir
    
    from dtactions import unimacroactions
    from dtactions import inivars
    ua_file = Path(actual_dta_userdir)/'unimacroactions.ini'
    ini = inivars.IniVars(ua_file)
    assert ini
    
    

def test_dtactions_userdir_default(dtactions_setup_default):
    r"""see if a copy of a sample ini file for dtactions is copied to
    ~\.dtactions when no or invalid env variable DTACTIONS_USERDIR is given
    
    TODO: see question in conftest.py (QH to Doug)
    """
    dtactions_userdir = dtactions_setup_default
    import dtactions
    assert Path(dtactions_userdir).is_dir()
    actual_dta_userdir = dtactions.getDtactionsUserDirectory()
    assert actual_dta_userdir == dtactions_userdir
    

    
    
if __name__ == "__main__":
    pytest.main(['test_dtactions_base.py'])
