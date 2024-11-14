"""common fixtures for the dtactions project
"""
#pylint:disable = E1101
# from string import Template
from shutil import copy as file_copy
from pathlib import Path
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
def dtactions_source_path() -> Path:
    return Path(importlib.util.find_spec("dtactions").submodule_search_locations[0])

@cache 
def dtactions_sample_ini_path() -> Path:
    return dtactions_source_path()/"sample_ini"

 
def copy_unimacroactions_ini():
    inifile_name = "unimacroactions.ini"
    dtactions_user_path = dtactions.getDtactionsUserPath()
    source_file=dtactions_sample_ini_path()/"unimacroactions.ini"
    target_file=dtactions_user_path/inifile_name
    return file_copy(source_file,target_file)
 

@pytest.fixture()
def dtactions_setup_with_env_var(tmp_path_factory):
    dta_user_path = tmp_path_factory.mktemp("dtactions_user_dir")
    print(f'dtactions_setup, dta_userdir: "{dta_user_path}"')
    pytest.MonkeyPatch().setenv("DTACTIONS_USERDIR", str(dta_user_path))
    return dta_user_path

@pytest.fixture()
def dtactions_setup_default(tmp_path_factory):
    """try to send home to one of the newly created folders. MonkeyPatch throws an error
    
    TODO QH
    """
    dta_test_home_path = tmp_path_factory.mktemp('user_home_dir')
    def fake_home():
        return dta_test_home_path
    
    print(f'dtactions_setup, default home path: "{dta_test_home_path}"')
    pytest.MonkeyPatch().setenv("DTACTIONS_USERDIR","")
    pytest.MonkeyPatch.setattr(dtactions, 'get_home_path', fake_home)
    return dta_test_home_path


    
    


def test_foo(dtactions_setup):
    pass
 
