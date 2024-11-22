"""common fixtures for the dtactions project
"""
#pylint:disable = E1101
# from string import Template
from shutil import copy as file_copy
from pathlib import Path, WindowsPath
from functools import cache
import importlib
import pytest
import natlink
import dtactions
from natlinkcore import natlinkstatus
thisPath = Path(__file__).parent

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
def dtactions_setup_with_env_var(tmp_path, monkeypatch):
    """dta_user_path points to a tmp_path
    """
    dta_user_path = tmp_path   
    print(f'dtactions_setup, dta_userdir: "{dta_user_path}"')
    monkeypatch.setenv("DTACTIONS_USERDIR", str(dta_user_path))
    return dta_user_path

@pytest.fixture()
def dtactions_setup_default(tmp_path, monkeypatch):
    """return path of fake_home / '.dtactions'
    
    """
    dta_test_home_path = tmp_path
    def fake_home():
        return dta_test_home_path
    
    print(f'dtactions_setup, default home path: "{dta_test_home_path}"')
    monkeypatch.setenv("DTACTIONS_USERDIR","")
    monkeypatch.setattr(WindowsPath, 'home', fake_home)
    dta_userdir = dta_test_home_path/'.dtactions'
    dta_userdir.mkdir()
    return dta_userdir

@pytest.fixture()
def test_files_path():
    """return path of test_files (used in test_unimacroactions.py)
    
    """
    return thisPath / 'test_files'


    
    


def test_foo(dtactions_setup):
    pass
 
