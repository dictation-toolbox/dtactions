"""
This module tests the possibilities of AppBringUp, the Dragon (dvc) AppBringUp.

Quintijn Hoogenboom, November 2024
"""
#pylint:disable = C0415, W0611, R0914
import time
from shutil import copy as file_copy
import pytest
# from dtactions import natlinkclipboard
# from dtactions import unimacroinivars as inivars  # old style
import dtactions
from dtactions import unimacroutils as uu


def test_bringup(dtactions_setup_default, tmp_path, nat_conn):
    """see if basic bringup works
    
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
    prog_info = uu.getProgInfo()    
    assert prog_info.title.find('test file.txt') >= 0
    ua.doAction('{alt+f4}')

    

def test_bringup_and_switch_two_files_notepad(dtactions_setup_default, tmp_path, nat_conn):
    """see if bringup works with default notepad, and switching between the two...
    
    """
    from dtactions import unimacroactions as ua
    from dtactions import inivars
    file_one = tmp_path/'test file one.txt'
    with open(file_one, 'w', encoding='utf-8') as fp:
        fp.write('open file testing on')
        fp.write('\n')
    file_two = tmp_path/'testfil\xe9_two.txt'
    with open(file_two, 'w', encoding='utf-8') as fp:
        fp.write('open file testing two with Montr\xe9al.')
        fp.write('\n')
    # Notepad appears with the file, close manually...
    result = ua.UnimacroBringUp(app='open', filepath=file_one)
    time.sleep(0.3)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find('test file one') >= 0

    result = ua.UnimacroBringUp(app='open', filepath=file_two)
    time.sleep(0.3)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find('_two.txt') >= 0

    result = ua.UnimacroBringUp(app='open', filepath=file_one)
    time.sleep(0.3)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find('test file one') >= 0
    # close the file:
    ua.doAction('{alt+f4}')
    

    result = ua.UnimacroBringUp(app='open', filepath=file_two)
    assert result == 1
    time.sleep(0.3)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find('_two.txt') >= 0
    # close the file:
    ua.doAction('{alt+f4}')


@pytest.mark.skip(reason="switching back to file 1 does not work, try/test manually")
def test_bringup_and_switch_two_files_excel(dtactions_setup_default, tmp_path, test_files_path, nat_conn):
    """see if bringup works with excel files, and switching between the two...
    
    two files appear, but not switching back again. QH 22-11-2024
    
    """
    from dtactions import unimacroactions as ua
    from dtactions import inivars
    file_one = file_two = None
    name_one, name_two = 'test excel 1.xlsx', 'test excel 2.xlsx'
    for name in name_one, name_two:
        source_file = test_files_path/name
        target_file= tmp_path/name
        file_copy(source_file,target_file)
        if file_one is None:
            file_one = target_file
        else:
            file_two = target_file
    assert file_one.is_file()
    assert file_two.is_file()

    result = ua.UnimacroBringUp(app='open', filepath=file_one)
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_one) >= 0

    result = ua.UnimacroBringUp(app='open', filepath=file_two)
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_two) >= 0

    result = ua.UnimacroBringUp(app='open', filepath=file_one)
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_one) >= 0
    # close the file:
    ua.doAction('{ctrl+f4}')
    

    result = ua.UnimacroBringUp(app='open', filepath=file_two)
    assert result == 1
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_two) >= 0
    # close the file:
    ua.doAction('{ctrl+f4}')

# @pytest.mark.skip(reason="switching back to file 1 does not work, try/test manually")
def test_bringup_and_switch_two_files_winword(dtactions_setup_default, tmp_path, test_files_path, nat_conn):
    """see if bringup works with word docx files, and switching between the two...
    
    same problem as with excel... switching back does not work... (QH 22-11-2024)

    test manually by inserting function (after ::) in bottom line of file, and removing the
    @pytest.mark.skip above.
    
    """
    from dtactions import unimacroactions as ua
    from dtactions import inivars
    file_one = file_two = None
    name_one, name_two = 'test word 1.docx', 'test word 2.docx'
    for name in name_one, name_two:
        source_file = test_files_path/name
        target_file= tmp_path/name
        file_copy(source_file,target_file)
        if file_one is None:
            file_one = target_file
        else:
            file_two = target_file
    assert file_one.is_file()
    assert file_two.is_file()

    result = ua.UnimacroBringUp(app='open', filepath=file_one)
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_one) >= 0

    result = ua.UnimacroBringUp(app='open', filepath=file_two)
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_two) >= 0

    # ua.doAction("WINKEY b")    # remove focus, gives no improvement.
    # ua.doAction("W")

    result = ua.UnimacroBringUp(app='open', filepath=file_one)
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_one) >= 0
    # close the file:
    ua.doAction('{ctrl+f4}')
    
    ua.doAction("WINKEY b")
    ua.doAction("W")

    result = ua.UnimacroBringUp(app='open', filepath=file_two)
    assert result == 1
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_two) >= 0
    # close the file:
    ua.doAction('{ctrl+f4}')


@pytest.mark.skip(reason="trying vscode TODO: Doug??")
def test_bringup_and_switch_two_files_vscode(dtactions_setup_default, tmp_path, test_files_path, nat_conn):
    """see if bringup works with several files, with visual studio as parameters
    
    """
    from dtactions import inivars
    
    dta_user_path = dtactions_setup_default
    ua_file = dta_user_path/'unimacroactions.ini'
    ini = inivars.IniVars(ua_file)
    # ini.set('bringup edit', 'py', 'vscode')
    # ini.set('bringup edit', 'txt', 'vscode')
    # ini.set('bringup vscode', 'path', r'C:\Users\Gebruiker\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code')
    ini.set('bringup edit', 'py', 'komodo')
    ini.set('bringup edit', 'txt', 'komodo')
    ini.set('bringup komodo', 'path', r'C:\Program Files (x86)\ActiveState Komodo IDE 12\komodo.exe')
    ini.write()

    # only now import unimacroactions!!
    from dtactions import unimacroactions as ua    
    
    file_one = file_two = None
    name_one, name_two = 'ide test txt.txt', 'ide test python.py'
    for name in name_one, name_two:
        source_file = test_files_path/name
        target_file= tmp_path/name
        file_copy(source_file,target_file)
        if file_one is None:
            file_one = target_file
        else:
            file_two = target_file
    assert file_one.is_file()
    assert file_two.is_file()

    result = ua.UnimacroBringUp(app='edit', filepath=file_one)
    time.sleep(5)  # long time starting (another!) Komodo
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_one) >= 0

    result = ua.UnimacroBringUp(app='edit', filepath=file_two)
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_two) >= 0

    # ua.doAction("WINKEY b")    # remove focus, gives no improvement.
    # ua.doAction("W")

    result = ua.UnimacroBringUp(app='edit', filepath=file_one)
    time.sleep(2)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_one) >= 0
    # close the file:
    ua.doAction('{ctrl+f4}')
    
    # ua.doAction("WINKEY b")
    # ua.doAction("W")

    result = ua.UnimacroBringUp(app='edit', filepath=file_two)
    assert result == 1
    time.sleep(5)    # 2 seconds not enough??
    prog_info = uu.getProgInfo()
    assert prog_info.title.find(name_two) >= 0
    # close the file:
    ua.doAction('{ctrl+f4}')

    
def test_bringup_wrong_file_setting(dtactions_setup_default, tmp_path, nat_conn):
    """see if bringup works also with wrong input in unimacroactions.ini
    
    the file should then be brought up in notepad...
    
    """
    from dtactions import inivars

    # tweak ini instance: (see basic test above)
    dta_user_path = dtactions_setup_default
    ua_file = dta_user_path/'unimacroactions.ini'
    ini = inivars.IniVars(ua_file)
    ini.set('bringup edit', 'py', 'komodo')
    ini.set('bringup komodo', 'path', r'C:\invalid\path\to\komodo')
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
    time.sleep(0.7)
    prog_info = uu.getProgInfo()
    assert prog_info.title.find('test.py') >= 0
    ua.doAction('{alt+f4}')

if __name__ == "__main__":
    # pytest.main(['test_appbringup.py::test_bringup_and_switch_two_files_vscode'])
    pytest.main(['test_appbringup.py'])
