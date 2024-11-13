"""common fixtures for the dtactions project
"""
#pylint:disable = E1101
# from string import Template
from shutil import copy as file_copy
from pathlib import Path, WindowsPath
import pathlib
from functools import cache
import importlib
import pytest
import natlink
import dtactions
from natlinkcore import natlinkstatus
thisDir = Path(__file__).parent

@pytest.fixture(scope="module")
def nl_status():
    status = natlinkstatus.NatlinkStatus()
    return status

@pytest.fixture(scope="module")
def nat_conn():
    yield natlink.natConnect()
    natlink.natDisconnect()


@cache          #only needs to be called once
def dtactions_source_dir() ->Path:
    return Path(importlib.util.find_spec("dtactions").submodule_search_locations[0])

@cache 
def dtactions_sample_ini_dir() ->Path:
    return dtactions_source_dir()/"sample_ini"

 
def copy_unimacroactions_ini():
    inifile_name = "unimacroactions.ini"
    dtactions_user_dir = Path(dtactions.getDtactionsUserDirectory())
    source_file=dtactions_sample_ini_dir()/"unimacroactions.ini"
    target_file=dtactions_user_dir/inifile_name
    return file_copy(source_file,target_file)
 
#a function to build the fixture to copy a grammar ini file to dtactions user directory.  see sample in test_brackets.py
#the fixture itself will return the same thing to the tests as the dtactions_setup fixture.

# def make_copy_unimacroactions_ini_fixture(samplegrammar_ini_file : str):
#     @pytest.fixture()
#     def unimacroactions_ini_fixture(dtactions_setup):
#         copy_unimacroactions_ini(dtactions_setup[0])
#         return dtactions_setup
#     return unimacroactions_ini_fixture




@pytest.fixture()
def dtactions_setup_with_env_var(tmpdir):
  
    tmp_test_root = tmpdir

    dtactions_userdir=str(tmp_test_root.mkdir("dtactions_user_directory"))
    print(f'dtactions_setup, dtactions_userdir: "{dtactions_userdir}"')
    pytest.MonkeyPatch().setenv("DTACTIONS_USERDIR", dtactions_userdir)
    return dtactions_userdir

@pytest.fixture()
def dtactions_setup_default(tmpdir):
    """try to send home to one of the newly created folders. MonkeyPatch throws an error
    
    TODO QH
    """
    
    
    
    def fake_home():
        return dta_user_default
    
    tmp_test_root = tmpdir

    dta_user_default = str(tmp_test_root.mkdir("dta_user_default"))
    print(f'dtactions_setup, dta_user_default: "{dta_user_default}"')
    pytest.MonkeyPatch().setenv("DTACTIONS_USERDIR","")
    pytest.MonkeyPatch.setattr(WindowsPath.home, fake_home)
    return dta_user_default  


    
    


def test_foo(dtactions_setup):
    pass
 
