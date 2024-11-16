"""
This module tests the unimacroactions module 

Quintijn Hoogenboom, December 2022/November 2024
"""
#pylint:disable = C0415, W0611, W0107
import pytest
# from dtactions import natlinkclipboard
# from dtactions import unimacroinivars as inivars  # old style
import dtactions

def test_path_and_inifile_default(dtactions_setup_default):
    """test if the unimacroactions.ini file is copied into the dtactions user dir
    and check the validity of that ini file
    
    """
    from dtactions import unimacroactions
    from dtactions import inivars
    
    dta_user_path = dtactions_setup_default
    assert dta_user_path.is_dir()
    
    actual_dta_user_path = dtactions.getDtactionsUserPath()
    assert actual_dta_user_path == dta_user_path

    assert (dta_user_path/'unimacroactions.ini').is_file()
    
    ua_file = dta_user_path/'unimacroactions.ini'
    assert ua_file.is_file()
    ini = inivars.IniVars(ua_file)
    assert ini


# 
# def test_path_and_inifile_env_var(dtactions_setup_with_env_var):
#     """test if the unimacroactions.ini file is copied into the dtactions user dir
#     and check the validity of that ini file
#     This one cannot co-exist with "test_path_and_inifile_default" above...
#
#     Not relevant for unimacroactions testing though.
#     
#     """
#     from dtactions import unimacroactions
#     from dtactions import inivars
# 
# 
#     dta_user_path = dtactions_setup_with_env_var
#     assert dta_user_path.is_dir()
#     
#     actual_dta_user_path = dtactions.getDtactionsUserPath()
#     assert actual_dta_user_path == dta_user_path
# 
#     assert (dta_user_path/'unimacroactions.ini').is_file()
#     
#     ua_file = dta_user_path/'unimacroactions.ini'
#     ini = inivars.IniVars(ua_file)
#     assert ini



def test_matchProgTitleWithDict(dtactions_setup_default):
    """this tests if a prog and title matches with definitions in a Dict
    
    the Dict is taken from unimacroactions.ini
    section [general], key "child behaves like top" or "top behaves like child"
    """
    from dtactions import unimacroactions as ua
    from dtactions import inivars
    
    # the definition may contain part of the wanted title, but... only in matchPart is True...
    child_behaves_like_top = {"natspeak": ["dragon-balk", "ragonbar"]}
    
    # no match:
    assert ua.matchProgTitleWithDict('prog', 'title', child_behaves_like_top, matchPart=None) is False

    # title must match exact (but case insensitive)
    assert ua.matchProgTitleWithDict('natspeak', 'Dragon-balk', child_behaves_like_top) is True
    assert ua.matchProgTitleWithDict('natspeak', 'Dragonbar', child_behaves_like_top) is False
    assert ua.matchProgTitleWithDict('natspeak', 'Dragonbar', child_behaves_like_top, matchPart=True) is True
    # also good:
    assert ua.matchProgTitleWithDict('natspeak', 'Dragon-balk', child_behaves_like_top, matchPart=True) is True


def test_bringup(dtactions_setup_default, tmp_path, nat_conn):
    """see if bringup works also with wrong input in unimacroactions.ini
    
    """
    from dtactions import unimacroactions as ua
    from dtactions import inivars
    file_to_bringup = tmp_path/'test file.txt'
    with open(file_to_bringup, 'w', encoding='utf-8') as fp:
        fp.write('open file testing')
        fp.write('\n')
    # Notepad appears with the file, close manually...
    result = ua.UnimacroBringUp(app='open', filepath=file_to_bringup)
    assert result == 1

    
def test_bringup_wrong_file_setting(dtactions_setup_default, tmp_path, nat_conn):
    """see if bringup works also with wrong input in unimacroactions.ini
    
    """
    from dtactions import inivars
    
    # tweak ini instance: (see basic test above)
    dta_user_path = dtactions_setup_default
    ua_file = dta_user_path/'unimacroactions.ini'
    ini = inivars.IniVars(ua_file)
    ini.set('bringup edit', 'py', 'komodo')
    ini.set('bringup komodo', 'path', 'C:/path/to/invalid/komodo.exe')
    ini.write()

    # only now import unimacroactions!!
    from dtactions import unimacroactions as ua
    
    file_to_bringup = tmp_path/'test.py'
    with open(file_to_bringup, 'w', encoding='utf-8') as fp:
        fp.write('#test edit python file')
        fp.write('\n')
    # Notepad appears with the file, close manually...
    result = ua.UnimacroBringUp(app='edit', filepath=file_to_bringup)
    assert result == 1

def test_do_alert(tmp_path, nat_conn):
    """see if bringup works also with wrong input in unimacroactions.ini
    
    """
    from dtactions import unimacroactions as ua
    result = ua.do_ALERT()
    assert result == 1
    
    

# test SCLIP via unimacro/unimacroactions.py direct run.
    
    
if __name__ == "__main__":
    pytest.main(['test_unimacroactions.py'])
