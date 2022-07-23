# (unimacro - natlink macro wrapper/extensions)
# (c) copyright 2003 Quintijn Hoogenboom (quintijn@users.sourceforge.net)
#                    Ben Staniford (ben_staniford@users.sourceforge.net)
#                    Bart Jan van Os (bjvo@users.sourceforge.net)
#
#
# unimacroactions.py (former actions.py)
#  written by: Quintijn Hoogenboom (QH softwaretraining & advies)
#  June 2003/restructure March 2021/June 2022
#
#pylint:disable=C0302, C0209, C0321
#pylint:disable=E1101  
#pylint:disable=W0603
"""This module contains actions that can be called from natlink grammars.

The central functions are "doAction" and "doKeystroke".

Extensive use is made from the ini file "actions.ini".

Editing actions, debugging actions and showing actions is performed through
special functions inside this module, but calling from another file (with
natlink grammar) is needed to activate these functions by voice.

"""
import re
import sys
import types
import shutil
import copy
from pathlib import Path
# import win32api
# import win32gui
# import win32con
# import win32com.client
import html.entities
import time
import datetime
import natlink
from dtactions import monitorfunctions
from dtactions import autohotkeyactions
from dtactions import sendkeys
from dtactions.unimacro import unimacroutils
from dtactions.unimacro import inivars

external_actions_modules = {}  # the modules, None if not available (for prog)
external_action_instances = {} # the instances, None if not available (for hndle)
     
class ActionError(Exception):
    """ActionError"""
class KeystrokeError(Exception):
    """KeystrokeError"""
class UnimacroError(Exception):
    """UnimacroError"""
try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print('Run this module after "build_package" and "flit install --symlink"\n')
    raise

dtactionsDir = thisDir = getThisDir(__file__)
##### get actions.ini from baseDirectory or SampleDirectory into userDirectory:
sampleDirectory = dtactionsDir.parent
sampleDirectory = sampleDirectory/'samples'/'unimacro'
checkDirectory(sampleDirectory, create=False)
inifilename = 'unimacroactions.ini'
 
sampleInifile = sampleDirectory/inifilename
if not sampleInifile.is_file():
    raise OSError(f'\nNo ini file "{inifilename}" found in {sampleDirectory}\nCHECK YOUR CONFIGURATION!\n')

userDirectory = Path.home()/".dtactions"/"unimacro"
checkDirectory(userDirectory)

inifile = userDirectory/inifilename
if not inifile.is_file():
    shutil.copy(sampleInifile, inifile)
print('actions in dtactions not working yet, Unimacro and Vocola should revert to actions.py in Unimacro')
print(f'dtaction/unimacro/actions, inifile: {inifile}')

whatFile = userDirectory/(__name__ + '.txt')
debugSock = None
debugFile = userDirectory/'dtactions_debug.txt'
samples = []

# ========================================
metaActions = re.compile(r'(<<[^>]+>>)')
metaAction = re.compile(r'<<([^>]+)>>$')
metaNumber = re.compile(r' ([0-9]+)$')
metaNumberBack = re.compile(r'\bn\b')
actionIsList = re.compile('[,;\n]', re.M)
# for keystroke matching and splitting:
braceExact = re.compile (r'[{][^}]+[}]$')
hasBraces = re.compile (r'([{].+?[}])')
BracesExtractKey = re.compile (r'^[{]((alt|ctrl|shift)[+])*(?P<k>[^ ]+?)( [0-9]+)?[}]$', re.I)
        
# debugging and checking for changes (whenever actions are edited)
# are controlled to these global variables.  Whenever actions are edited
# through the function "editActions", checking for changes is performed
# until natlink is reloaded (or aNatSpeak is restarted)
debug = 0
checkForChanges = 0
iniFileDate = 0

# for Dragon, dvc commands:
natspeakCommands = ['ActiveControlPick', 'ActiveMenuPick', 'AppBringUp', 'AppSwapWith', 'Beep', 'ButtonClick',
 'ClearDesktop', 'ControlPick', 'DdeExecute', 'DdePoke', 'DllCall', 'DragToPoint', 'GoToSleep', 
 'HeardWord', 'HtmlHelp', 'MenuCancel', 'MenuPick', 'MouseGrid', 'MsgBoxConfirm', 'PlaySound',
 'RememberPoint', 'RunScriptFile', 'SendKeys', 'SendSystemKeys', 'SetCharFormat', 'SetMicrophone',
 'SetMousePosition', 'SetNaturalText', 'ShellExecute', 'TTSPlayString', 'Wait', 'WakeUp', 'WinHelp',
 'SetRecognitionMode']

class Action:
    """central class, performing a variety of actions
    
    """
    #pylint:disable=R0902, R0904
    def __init__(self, action=None, pauseBA=None, pauseBK=None, 
                 progInfo=None, modInfo=None, sectionList=None, comment='', comingFrom=None):
        #pylint:disable=R0913
        self.action = action
        self.ini = inivars.IniVars(inifile)
        self.iniFileDate = unimacroutils.getFileDate(inifile)

        self.pauseBA = pauseBA or self.ini.get('default', 'pause between actions')
        self.pauseBK = pauseBK or self.ini.get('default', 'pause between keystrokes')
        self.progInfo = progInfo
        self.modInfo = modInfo
        if progInfo:
            self.progInfo = progInfo
        else:
            self.progInfo = unimacroutils.getProgInfo(modInfo=modInfo)
            
        self.sectionList = sectionList
        self.comment = comment
        self.comingFrom = comingFrom
        self.completeAction = None
        self.completeKeystrokes = ""
        self.topchildDict = {}
        self.childtopDict = {}
        self.pendingMessage = ''
    
    def __call__(self, action):
        return self.doAction(action)

    def doCheckForChanges(self):
        """open or refresh inifile
        """
        # global  ini, iniFileDate, topchildDict, childtopDict
        prevIni = copy.deepcopy(self.ini)
        prevDate = self.iniFileDate
        newDate = unimacroutils.getFileDate(inifile)  ## inifile is global var
        if newDate > self.iniFileDate:
            D(1, '----------reloading ini file')
            try:
                self.topchildDict.clear()
                self.childtopDict.clear()
                self.ini = inivars.IniVars(inifile)
                self.iniFileDate = unimacroutils.getFileDate(inifile)
            except inivars.IniError:
                msg = 'repair actions ini file: \n\n' + str(sys.exc_info()[1])
                #win32api.ShellExecute(0, "open", inifile, None , "", 1)
                time.sleep(0.5)
                Message(msg)
                self.iniFileDate = 0
                self.ini = copy.deepcopy(prevIni)
                self.iniFileDate = prevDate

    def doAction(self, action):
        """central method, do an action
        """
        #pylint:disable=W0603
        global checkForChanges
        #pylint:disable=R0911
        # global pendingMessage, checkForChanges
        self.completeAction = action
        D(4, f'doAction: {self.completeAction}')
    
        if not self.ini:
            checkForChanges = 1
            if self.pendingMessage:
                mess = self.pendingMessage
                self.pendingMessage = None
                Message(mess, alert=1)
                D(1, 'no valid inifile for actions')
                return None
        if checkForChanges:
            D(5, 'checking for changes')
            self.doCheckForChanges() # resetting the ini file if changes were made
        if not self.ini:
            D(1, 'no valid inifile for actions')
            return None

        actionsList = list(self.ini.getIniList(action))
        # if action.find('USC') >= 0:
        #     bList = []
        #     for a in actionsList:
        #         if a.find('USC') >= 0:
        #             bList.extend(a.split('USC'))
        #         else:
        #             bList.append(a)
        #     actionsList= [t.strip() for t in bList if t]
            
        D(2, f'action: {action}, actionsList: {actionsList}')
        if not actionsList:
            return None
        for a in actionsList:
            result = self.do_part_of_action(a)
            if result:
                return result  # error message??
            if self.comingFrom and self.comingFrom.interrupted:
                return "interrupted"
            if self.pauseBA and a != actionsList[-1]:
                self.do_W(self.pauseBA)
        # no errors, return None
        return None
    
    def do_part_of_action(self, action):
        """do a "partial action", keystroke, USC, etc.
        
        This looks like an atomic actions, but these can also be splitted again,
        and recursively call itself.
        
        these are separated by ";", and in between modInfo and  pauseBA are checked
        
        return a message if an error occurs
        return None if all OK
        """
        #pylint:disable=R0914, R0911, R0912, R0915
        if metaAction.match(action):  # exactly a meta action, <<....>>
            a = metaAction.match(action).group(1)
            aNew = self.getMetaAction(a, self.sectionList, self.progInfo)
            if isinstance(aNew, tuple):
                # found function
                func, number = aNew
                func(number)
                return 1
            D(5, 'doing meta action: <<%s>>: %s'% (a, aNew))
            if aNew:
                aNewList = list(inivars.getIniList(aNew))
                res = 0
                for aa in aNewList:
                    if aa:
                        D(3, '\tdoing part of meta action: <<%s>>: %s'% (a, aa))
                        res = self.do_part_of_action(aa)
                        if not res:
                            return None
                return res
            if aNew == '':
                D(1, 'empty action')
                return 1
            D(3, 'error in meta action: "%s"'% a)
            partCom = '<<%s>>'%a
            t = '_actions, no valid meta action: "%s"'% partCom
            if partCom != self.completeAction:
                t += '\ncomplete command: "%s"'% self.completeAction
            raise ActionError(t)
    
        # try qh command:
        if action.find('(') > 0 and action.strip().endswith(')'):
            com, rest = [t.strip() for t in action.split('(', 1)]
            rest = rest.strip()[:-1]
        elif action.find(' ') > 0:
            com, rest = [t.strip() for t in action.split(' ',1)]
        else:
            com, rest = action.strip(), ''
    
        if 'do_'+com in globals():
            funcName = 'do_'+com
            args = convertToPythonArgs(rest)
            kw = {}
            kw['progInfo'] = self.progInfo
            kw['comingFrom'] = self.comingFrom
            func = globals()[funcName]
            if not isinstance(func, types.FunctionType):
                raise UnimacroError('appears to be not a function: %s (%s)'% (funcName, func))
            D(5, 'doing USC command: |%s|, with args: %s and kw: %s'% (com, repr(args), kw))
            if debug > 1: self.do_W(debug*0.2)
            if args:
                result = func(*args, **kw)
            else:
                result = func(**kw)
                    
            D(5, 'did it, result: %s'% result)
            if debug > 1: self.do_W(debug*0.2)
            # skip pause between actions
            #self.do_W(pauseBA)
            return result
    
        # try meta actions inside:
        if metaActions.search(action): # action contains meta actions
            As = [t.strip() for t in metaActions.split(action)]
            D(5, 'meta actions: %s'% As)
            for a in As:
                if not a: continue
                res = self.do_part_of_action(a)
                if not res:
                    return None
                self.do_W(self.pauseBA)
                # skip pause between actions
                #self.do_W(pauseBA)
            return 1
    
        # try natspeak command:
    
        if com in natspeakCommands:
            rest = self.convertToDvcArgs(rest)
            C = com + ' ' + rest
            D(1, 'do dvc command: |%s|'% C)
            if debug > 1: self.do_W(debug*0.2)
            natlink.execScript(com + ' ' + rest)
            if debug > 1: self.do_W(debug*0.2)
            # skip pause between actions
            #self.do_W(pauseBA)
            return 1
    
        # all the rest:
        D(5, 'do string: |%s|'% action)
        if debug > 1: self.do_W(debug*0.2)
        if action:
            self.doKeystroke(action)
            if debug > 1: self.do_W(debug*0.2)
            # skip pause between actions
            #self.do_W(pauseBA)
            # if topLevel: # first (nonrecursive) call,
            #     print('end of complete action')
            return 1
        if debug: print('empty keystrokes')
        ## all else failed...
        return None     
        
    def doKeystroke(self, keystrokes, hardKeys=None, pauseBK=None):
        """do keystrokes to the foreground window.
        
        According to hardkey or pause between keystrokes, most of the time not needed,
        the keystrokes are splitted, and sent one by one.
        """
        #print 'doing keystroke: {%s'% keystrokes[1:]
        #pylint:disable=W0603
        global checkForChanges
        if not keystrokes:
            return None

        if hardKeys is None and pauseBK is None:
            # for debugging, remains when instance stays open until a fresh doKeystroke call occurs
            self.completeKeystrokes = keystrokes
            # setting default, come here only if non recursive call:
            hardKeys = ['none']
            pauseBK = 0
    
            if not self.ini:
                checkForChanges = 1
            if checkForChanges:
                D(5, 'checking for changes')
                self.doCheckForChanges() # resetting the ini file if changes were made# 
            if not self.ini:
                D(1, 'no valid inifile for keystrokes')
                hardKeys = ['none']
                pauseBK = 0
            else:    
                if self.sectionList is None:
                    self.sectionList = self.getSectionList(self.progInfo)
                pauseBK = int(self.setting('pause between keystrokes', '0',
                                              sectionList=self.sectionList))
                hardKeys = self.getHardkeySettings()
                hardKeys = ['none']
        
        if pauseBK:
            l = hasBraces.split(keystrokes)
            for k in l:
                if not k:
                    continue
                if braceExact.match(keystrokes):
                    self.doKeystroke(k, hardKeys=hardKeys, pauseBK = 0)
                else:
                    for K in k:
                        self.doKeystroke(K, hardKeys=hardKeys, pauseBK = 0)
                D(5, 'pausing: %s msec after keystroke: |%s|'% (pauseBK, k))
                # skip pausing between keystrokes
                #self.do_W(pauseBK)
        elif braceExact.match(keystrokes):
            # exactly 1 {key}:
            D(5, 'exact keystrokes, hardKeys[0]: %s'% hardKeys[0])
            if hardKeys[0] == 'none':
                sendkeys.sendkeys(keystrokes)  # the fastest way
                return None
            if hardKeys[0] == 'all':
                ### must be implemented later
                sendkeys.sendkeys(keystrokes)
                return None
            m = BracesExtractKey.match(keystrokes)
            if m:
                # the key part is known and valid word
                # eg tab, down etc
                keyPart = m.group(0).lower()
                for mod in 'shift+', 'ctrl+', 'alt+':
                    if keyPart.find(mod)>0: keyPart = keyPart.replace(mod, '')
                if keyPart in hardKeys:
                    D(3, 'doing "hard": |%s|'% keystrokes)
                    ### TODOQH, later hard keys
                    sendkeys.sendkeys(keystrokes)
                    return None
                D(3, 'doing "soft" (%s): |%s|'% (keystrokes, hardKeys))
                sendkeys.sendkeys(keystrokes)
                return None
            D(3, 'doing "soft" (%s): |%s|'% (keystrokes, hardKeys))
            sendkeys.sendkeys(keystrokes)
            return None
            
        # now proceed with more complex keystroke possibilities:
        if self.hardKeys[0]  == 'all':
            ## TODOQH, hard keys
            # natlinkutils.playString(keystrokes, natlinkutils.hook_f_systemkeys)
            sendkeys.sendkeys(keystrokes)
            return None
        if self.hardKeys[0]  == 'none':
            sendkeys.sendkeys(keystrokes)
            return None
        if hasBraces.search(keystrokes):
            keystrokeList = hasBraces.split(keystrokes)
            for k in keystrokeList:
                D(5, 'part of keystrokes: |%s|' % k)
                if not k:
                    continue
                #print 'recursing? : %s (%s)'% (k, keystrokeList)
                sendkeys.sendkeys(keystrokes)
                # self.doKeystroke(k, hardKeys=hardKeys, pauseBK = 0)
        else:
            D(5, 'no braces keystrokes: |%s|' % keystrokes)
            sendkeys.sendkeys(keystrokes)
        return None
    
    def getMetaAction(self, ma, sectionList=None, progInfo=None):
        """return the action that is found in the ini file.
        
        """
        m = metaNumber.search(ma)
        if m:
            number = m.group(1)
            A = ma.replace(number, 'n')
            actionName = A.replace(number, '')
            actionName = actionName.replace(' ', '')
        else:
            A = ma
            number = 0
            actionName = A.replace(' ', '')
        # try via actions_prog module:
        ext_instance = self.get_instance_from_progInfo(progInfo)
        if ext_instance:
            prog = progInfo.prog
            funcName = 'metaaction_%s'% actionName
            func = getattr(ext_instance,funcName, None)
            if func:
                D(1, 'action by function from prog %s: |%s|(%s), arg: %s'% (prog, actionName, funcName, number))
                result = func, number
                if result: return result
                # otherwise go on with "normal" meta actions...
    
        # no result in actions_prog module, continue normal way:
        D(5, 'search for action: |%s|, sectionList: %s' %
                        (A, sectionList))
        
        aNew = self.setting(A, default=None, sectionList=sectionList)  #self.getFromIni
        if aNew is None:
            print('action: not found, meta action for %s: |%s|, searched in sectionList: %s' % \
                  (A, aNew, sectionList))
            return False
        if m:
            aNew = metaNumberBack.sub(number, aNew)
        if debug:
            section = self.ini.getMatchingSection(sectionList, A)
            D(1, '<<%s>> from [%s]: %s'% (A, section, aNew)) 
        return aNew        
        
    # last one undocumented, only for version 7
    
    def getSectionList(self, progInfo=None):
        if not progInfo:
            progInfo = unimacroutils.getProgInfo()
        prog, title, topchild, _classname, _hndle = progInfo
        D(5, 'search for prog: %s and title: %s' % (prog, title))
        D(5, 'type prog: %s, type title: %s'% (type(prog), type(title)))
        L = self.ini.getSectionsWithPrefix(prog, title)
        L2 = self.ini.getSectionsWithPrefix(prog, topchild)  # catch program top or program child
        for item in L2:
            if not item in L:
                L.append(item)
        L.extend(self.ini.getSectionsWithPrefix('default', topchild)) # catcg default top or default child
        D(5, 'section list with progInfo: %s:\n===== %s' % (progInfo, L))
                        
        return L
    
    def convertToDvcArgs(self, text):
        text = text.strip()
        if not text:
            return ''
        L = text.split(',')
        L = list(map(_convertToDvcArg, L))
        return ', '.join(L)
    
            
    
    def getFromIni(self, keyword, default='',
                    sectionList=None, progInfo=None):
        if not self.ini:
            return ''
        if sectionList is None:
            if self.progInfo is None:
                self.progInfo = unimacroutils.getProgInfo()
            _progpath, prog, title, _topchild, _classname, _hndle = self.progInfo
            sectionList = self.ini.getSectionsWithPrefix(prog, title) + \
                          self.ini.getSectionsWithPrefix('default', title)
            D(5, 'getFromIni, sectionList: |%s|' % sectionList)
        value = self.ini.get(sectionList, keyword, default)
        D(5, 'got from setting/getFromIni: %s (keyword: %s'% (value, keyword))
        return value
    
    setting = getFromIni

    def getHardkeySettings(self):
        """get from ini the keystrokes that should be done as "hardKeys"
        
        hardKeys are (with Natlink) implemented as SendSystemKeys
        """
        if not self.sectionList:
            raise ValueError('action, getHardkeySettings, sectionList should be filled')
        hardKeys = self.setting('keystrokes with systemkeys', 'none', sectionList=self.sectionList)
        D(5, f'hardKeys setting: {self.hardKeys}')
    
        hardKeys = actionIsList.split(hardKeys)
        if hardKeys:
            hardKeys = [k.strip() for k in hardKeys]
            D(5, 'hardKeys as list: |%s|'% hardKeys)
            D(5, 'new keystokes: |%s|, hardKeys: %s'%
                            (self.completeKeystrokes, hardKeys))
        return hardKeys

    
    def get_external_module(self, prog):
        """try to from natlinkcore import actions_prog and put in external_actions_modules
        
        if module not there, put None in this external_actions_modules dict
        """
        #print 'ask for program: %s'% prog
        if prog in external_actions_modules:
            return external_actions_modules[prog]
        try:
            modname = '%s-actions'% str(prog)
            _temp = __import__('actionclasses', fromlist=[modname])
            mod = getattr(_temp, modname)
            external_actions_modules[prog] = mod
            print('get_external_module, found actions module: %s'% modname)
            return mod
        except AttributeError:
            external_actions_modules[prog] = None
            # print 'get_external_module, no module found for: %s'% prog
        return None
     
    def get_instance_from_progInfo(self, progInfo=None):
        """return the correct instances for progInfo
        """
        progInfo = progInfo or self.progInfo
        prog, hndle = progInfo.prog, progInfo.hndle
        if hndle in external_action_instances:
            instance = external_action_instances[hndle]
            instance.update(progInfo)
            return instance
        
        mod = self.get_external_module(prog)
        if not mod:
            # print 'no external module instance: %s'% progInfo.prog
            return None   # no module, no instance
        classRef = getattr(mod, '%sActions'% prog.capitalize())
        if classRef:
            instance = classRef(progInfo)
            instance.update(progInfo)
            print('new instance for actions for prog: %s, hndle: %s'% (prog, hndle))
        else:
            instance = None
        external_action_instances[hndle] = instance
    
        return instance
    
    
    
        
    
    def showActions(self, progInfo=None, lineLen=60, sort=1, comingFrom=None, name=None):
        
        def T(language, Dict):
            if language in Dict:
                return Dict[language]
            return Dict['enx']
        
        progInfo = progInfo or self.progInfo
        if not self.progInfo:
            progInfo = unimacroutils.getProgInfo()
            
        prog, title, topchild, classname, hndle = progInfo
        language = unimacroutils.getLanguage()
        
        sectionList = self.getSectionList(progInfo)
    
        l = [f'actions for program: {prog}\n\twindow title: {title} (topchild: {topchild}, classname: {classname}, hndle: {hndle}',
             '\tsection list: {sectionList}',
             '']
        l.append(self.ini.formatKeysOrderedFromSections(sectionList,
                                    lineLen=lineLen, sort=sort))
        
        l.append(T(language, dict(enx= '\n"edit actions" to edit the actions',
                                  nld='\n"bewerk acties" om de acties te bewerken')))
        l.append(T(language, dict(enx='"show actions" gives this list',
                                nld= '"toon acties" geeft deze lijst')))
        l.append(T(language, dict(enx='"trace actions (on|off|0|1|...|10)" sets/clears trace mode',
                                  nld='"trees acties (aan|uit|0|1|...|10)" zet trace modus aan of uit')))
        l.append(T(language, dict(enx='consult grammar "control" for the exact commands',
                                  nld='raadpleeg grammatica "controle" voor de precieze commando\'s')))
        
        with open(whatFile, 'w', encoding='utf-8') as f:
            f.write('\n'.join(l))
        
        if comingFrom:
            name=name or ""
            comingFrom.openFileDefault(whatFile, name=name)
        else:
            print(f'show file: {whatFile}')
            print('showing actions by ShellExecute crashed NatSpeak, please open by hand above file')
        #win32api.ShellExecute(0, "open", whatFile, None , "", 1)
    
    def editActions(self, comingFrom=None, name=None):
        global checkForChanges, iniFileDate, topchildDict, childtopDict
        iniFileDate = unimacroutils.getFileDate(inifile)
        checkForChanges = 1
        topchildDict = None
        childtopDict = None
        if comingFrom:
            name=name or ""
            comingFrom.openFileDefault(inifile, name=name)
        else:
            print('inifile: ', inifile)
            print('editing actions by ShellExecute crashed NatSpeak, please edit by hand above file')
        #win32api.ShellExecute(0, "open", inifile, None , "", 1)
    
    def setPosition(self, name, pos, prog=None):
        """set in inifile, section 'positions'
        use prog for program specific positions
        """
        global checkForChanges, iniFileDate
        iniFileDate = unimacroutils.getFileDate(inifile)
        checkForChanges = 1
        if prog:
            section = 'positions %s'% prog
        else:
            section = 'positions'
        self.ini.set(section, name, pos)
        unimacroutils.Wait(0.1)
        self.ini.write()
    
    def getPosition(name, prog=None):
        """get from inifile, section 'positions'
        for application specific gets, use prog as programSpecific
        
        """
        if prog:
            section = 'positions %s'% prog
        else:
            section = 'positions'
        return ini.getInt(section, name) or 0
    # -----------------------------------------------------------
    
    def do_TEST(self, *args, **kw):
        # x, y = unimacroutils.testmonitorinfo(args[0], args[1])
        # print 'in do_test: ', x, y
        #pylint:disable=C0415
        import ctypes
        tup = natlink.getCurrentModule()
        hndle = tup[2]
        print('foreground hndle: %s, type: %s'% (hndle, type(hndle)))
        curprog = ctypes.windll.user32.GetForegroundWindow()
        print('curprog: %s'% curprog)
        # result = ctypes.windll.user32.GetGUIThreadInfo(hndle, info)
        # 
        # print('result: %s'% result)
        # print('info: %s'% repr(info))
        # print('dir info: %s'% dir(info))
        kbLayout = ctypes.windll.user32.GetKeyboardLayout(0)
        print('kbLayout: %s'% kbLayout)
    
    
    def do_AHK(self, script, **kw):
        """try autohotkey integration
        """
        result = autohotkeyactions.do_ahk_script(script)
        if isinstance(result, str):
            do_MSG(result)
        return result
    
    # HeardWord (recognition mimic), convert arguments to list:
    def do_HW(self, *args, **kw):
        words = list(args)
        origWords = copy.copy(words)
    ##    print 'origWords for HW(recognitionMimic): |%s| (length: %s)'% (origWords, len(origWords))
        
        if len(words)  == 1:
            if words[0].find(',') > 0:
                words = words[0].split(',')
            else:
                words = words[0].split()
                        
        try:
            D(3, 'words for HW(recognitionMimic): |%s|'% words)
            natlink.recognitionMimic(words)
        except:
            if words != origWords:
                try:
                    D(3, 'words for HW(recognitionMimic): |%s|'% origWords)
                    natlink.recognitionMimic(origWords)
                except:
                    print("action HW probably has invalid words (splitted: %s) or (glued together: %s)" % \
                                     (words, origWords))
                    return
            else:
                print("Cannot recognise words, HW probably has invalid words: %s" % words)
                return
        return 1
    
    def do_MP(self, scrorwind, x, y, mouse='left', nClick=1, **kw):
        # Mouse position and click.
        # first parameter 0, absolute
        # new 2017: 3 = relative to active monitor
        # and 4 - relative to work area of active monitor (not considering task bar or dragon bar)
        if not scrorwind in [0,1,2,3,4,5]:
            raise UnimacroError('Mouse action not supported with relativeTo: %s' % scrorwind)
        # first parameter 1: relative:
        unimacroutils.doMouse(0,scrorwind,x,y,mouse,nClick)  # abs, rel to window, x, y, click
        return 1
    
    def do_CLICK(self, mouse='left', nClick=1, **kw):
        scrorwind = 2
        x, y, = 0, 0
        # click at current position:
        unimacroutils.doMouse(0,scrorwind,x,y,mouse,nClick) 
        return 1
    
    
    def do_ENDMOUSE(self, **kw):
        unimacroutils.endMouse()
        return 1
        
    def do_CANCELMOUSE(self, **kw):
        unimacroutils.cancelMouse()
        return 1
    
    def do_MDOWN(self, button='left', **kw):
        unimacroutils.mousePushDown(button)
        return 1
    
    def do_RM(self, **kw):
        unimacroutils.rememberMouse()
        return 1
     
    def do_MOUSEISMOVING(self, **kw):
        xold, yold = natlink.getCursorPos()
        time.sleep(0.05)
        for i in range(2):
            x, y = natlink.getCursorPos()
            time.sleep(0.05)
            if not (x == xold and y == yold):
                return 1
        return 0
        
    
    def do_WAITMOUSEMOVE(self, **kw):
        """wait for the mouse start moving
        
        cancel after 2 seconds
        """
        xold,yold = natlink.getCursorPos()
        for i in range(40):
            time.sleep(0.05)
            x, y = natlink.getCursorPos()
            if x != xold or y != yold:
                return 1
        else:
            print('no mouse move detected, cancel action')
            return 0  # no result
    
    def do_WAITMOUSESTOP(self, **kw):
        """wait for the mouse stops moving
        
        you get 2 seconds for stopping the movement.
        If you move more than 200 in x or y direction, the function is cancelled (with error)
        
        return value: 1 if succes
                      None is failure
        
        """
        natlink.setMicState('off')
        xstart, ystart = xold,yold = natlink.getCursorPos()
    
        while 1:
            time.sleep(0.05)
            x, y = natlink.getCursorPos()
            if x == xold and y == yold:
                for j in range(10):
                    # check if it remains so (half a second)
                    time.sleep(0.05)
                    x, y = natlink.getCursorPos()
                    if x != xold or y != yold:
                        break
                else:
                    print('mouse stopped moving')
                    natlink.setMicState('on')
                    return 1
            if (x-xold)*(x-xold) + (y-yold)*(y-yold) > 100*100:
                print("user canceled WAITMOUSESTOP canceled by moving mouse more than 100 pixels")
                return
            xold, yold = x, y
            if natlink.getMicState() == "on":
                print('user canceled WAITMOUSESTOP by switching on microphone')
                return
    
    def do_CHECKMOUSESTEADY(self, **kw):
        """returns 1 if the mouse is steady
        
        for 2.5 seconds
        """
        xold, yold = natlink.getCursorPos()
        for i in range(50):
            time.sleep(0.05)
            x, y = natlink.getCursorPos()
            if x == xold and y == yold:
                for j in range(10):
                    # check if it remains so (half a second
                    time.sleep(0.05)
                    x, y = natlink.getCursorPos()
                    if x != xold or y != yold:
                        break
                else:
                    print('mouse stopped moving')
                    natlink.setMicState('on')
                    return 1
    
    
    
    def do_CLICKIFSTEADY(self, mouse='left', nClick=1, **kw):
        """clicks only if mouse is steady,
        
        beeps if it was moving before, if it is steady just click
        returns 0 if it keeps moving
        """
        unimacroutils.doMouse(0,scrorwind,x,y,mouse,nClick) 
        return 1
    
    
    def do_RMP(self, scrorwind, x, y, mouse='left', nClick=1, **kw):
        # relative mouse position and click
        # new 2017: 3 = relative to active monitor
        if not scrorwind in [0,1,3,5]:
            raise UnimacroError('Mouse action not supported with relativeTo: %s' % scrorwind)
        # first parameter 1: relative:
        unimacroutils.doMouse(1,scrorwind,x,y,mouse,nClick)  # relative, rel to window, x, y,click
        return 1
        
    def do_PRMP(self, all=0, **kw):
        # print relative mouse position
        unimacroutils.printMousePosition(1,all)  # relative
        return 1
        
    
    def do_PMP(self, all=0, **kw):
        # print absolute mouse position
        unimacroutils.printMousePosition(0,all)  # absolute
        return 1
    
    def do_PALLMP(self, **kw):
        # print all mouse positions
        unimacroutils.printMousePosition(0,1)  # absolute
        unimacroutils.printMousePosition(1,1)  # relative
        
    
    # Get the NatSpeak main menu:
    def do_NSM(self, **kw):
        modInfo = natlink.getCurrentModule()
        prog = unimacroutils.getProgName(modInfo)
        if prog == 'natspeak':
            if modInfo[1].find('DragonPad') >= 0:
                natlinkutils.playString('{alt+n}')
                return 1
        natlink.recognitionMimic(["NaturallySpeaking"])
        return unimacroutils.waitForWindowTitle(['DragonBar', 'Dragon-balk', 'Voicebar'],10,0.1)
    
    
    # shorthand for sendsystemkeys:
    def do_SSK(self, s, **kw):
        natlinkutils.playString(s, natlinkutils.hook_f_systemkeys)
        return 1
    
    def do_S(self, s, **kw):
        """do a simple keystroke
        temporarily with {shift} in front because of bugworkaround (Frank)
        """
        doKeystroke(s)
        return 1
    
    
    def do_ALTNUM(self, s, **kw):
        """send keystrokes with alt and numkey inserted
        
        werkt niet!!
        """
        if type(s) == int:
            s = str(s)
        keydown = natlinkutils.wm_syskeydown
        keyup = natlinkutils.wm_syskeyup
        
        altdown = (keydown, natlinkutils.vk_menu, 1)
        altup = (keyup, natlinkutils.vk_menu, 1)
        numkeyzero = win32con.VK_NUMPAD0
    
        sequence = [altdown]
        for code in s:
            if not code in '0123456789':
                print('code should be a digit, not: %s (%s)'% (code, s))
                return
            i = int(s)
            keycode = numkeyzero + int(code)
            sequence.append( (keydown, keycode, 1))
            sequence.append( (keyup, keycode, 1))
        sequence.append(altup)
        natlink.playEvents(sequence)
        
    def do_SCLIP(self, *s, **kw):
        """send keystrokes through the clipboard
        """
        unimacroutils.saveClipboard()
        unimacroutils.Wait()    
        ## actions should be able to catch , in string, now , seems to be separator for
        ## function parameters.
        ## assume , = ", "
        ## also assume {enter} = newline. No {enter 2}  etc recognised
        # print 's: %s type: %s'% (s, type(s))
        total = ', '.join(str(part) for part in s)
        #if len(s) > 1:
        #    for i, t in enumerate(s):
        #        print "SCLIP:", i, t
        total = total.replace("{enter}", "\n")
        # print 'SCLIP: %s type: %s'% (total, type(total))
        unimacroutils.setClipboard(total, format=13)
        unimacroutils.Wait()
        #print 'send through clipboard: %s'% total5N
        doAction("<<paste>>")
        unimacroutils.Wait()
        unimacroutils.restoreClipboard() 
    
    
    def do_RW(self, **kw):
        unimacroutils.rememberWindow()
        return 1
    
    def do_CW(self, **kw):
        """obsolete..."""
        unimacroutils.clearWindowHandle()
        return 1
    
    def do_RTW(self, **kw):
        unimacroutils.returnToWindow()
        return 1
    
    def do_SELECTWORD(self, count=1, direction=None, **kw):
        """select the word under the cursor"""
        print('try to select %s word(s) under cursor (direction: %s)'% (count, direction))
        unimacroutils.saveClipboard()
        if not direction in ['left', 'right']:
            # try if at end of word:
            doKeystroke("{extright}{shift+extleft}{ctrl+c}{extright}{extleft}")
            t = unimacroutils.getClipboard()
            if t in utilsqh.letters:
                direction = 'right'
                print('make direction right')
            else:
                direction = 'left'
                print('make direction left')
        if direction == 'left':
            doKeystroke("{extleft}{ctrl+extright}{shift+ctrl+extleft %s}"% count)
        elif direction == 'right':
            doKeystroke("{extright}{ctrl+extleft}{shift+ctrl+extright %s}"% count)
    
        unimacroutils.Wait()
        doAction("<<copy>>")
        unimacroutils.visibleWait()
        t = unimacroutils.getClipboard()
        if len(t) == 1 and t in utilsqh.letters:
            pass
        elif direction == 'left':
            while t and t.startswith(' '):
                doKeystroke("{shift+extright}")
                t = t[1:]
        else:
            while t and t.endswith(' '):
                doKeystroke("{shift+extleft}")
                t = t[:-1]
        unimacroutils.restoreClipboard()
        #print 'SELECTWORD, selected word: |%s| (leave on clipboard: %s'% (repr(t), repr(unimacroutils.getClipboard()))
        return t
    
    
    # wait 20 (default) times standard waitingtime (default 0.05 seconds)
    # (so default = 1 second)
    # for a new window title, if no title is given
    # wait for a change
    def do_WTC(self, nWait=20, waitingTime=0.05, **kw):
        """wait for a change in window title, on succes return 1
        
        after found, check also if title is stable
        """
        return unimacroutils.waitForNewWindow(nWait, waitingTime, **kw)
    ##    unimacroutils.ForceGotBegin()
    
    # wait for Window Title
    def do_WWT(self, titleName, nWait=20, waitingTime=0.05, **kw):
        """wait for specified window title, on succes return 1
        """
        return unimacroutils.waitForWindowTitle(titleName, nWait, waitingTime, **kw)
      
    # waiting function:
    def do_W(self, t=None, **kw):
        t = t or 0.1
        D(7, 'waiting: %s'%t)
        if debug and t > 2: D(1, 'waiting: %s'%t)
        unimacroutils.Wait(t)
        return 1
            
    do_WAIT = do_W
    # Long Wait:
    def do_LW(self, **kw):
        unimacroutils.longWait()
        return 1
    do_LONGWAIT = do_LW
    # Visible Wait:
    def do_VW(self, **kw):
        unimacroutils.visibleWait()
        return 1
    do_VISIBLEWAIT = do_VW
    
    # Short Wait:
    def do_SW(self, **kw):
        unimacroutils.shortWait()
        return 1
    do_SHORTWAIT = do_SW
    
    def do_KW(self, action1=None, action2=None, progInfo=None, comingFrom=None):
        """kill window
    
        """
        if action1:
            if action2:
    ##        print 'KW with: %s'% action1
                killWindow(action1, action2, progInfo=progInfo, comingFrom=comingFrom)
            else:
                killWindow(action1, progInfo=progInfo, comingFrom=comingFrom)
        else:
    ##        print 'KW without action 1'
            killWindow(progInfo=progInfo, comingFrom=comingFrom)
        return 1
    
    ## def do_RS():
    ##     """reformat selection, cleaning newlines
    
    ##     """
    ##     unimacroutils.saveClipboard()
    ##     doKeystroke('{ctrl+c}')
    ##     t = unimacroutils.getClipboard()
    ##     unimacroutils.restoreClipboard()
    ##     if t:
    ##         T = unimacroutils.cleanParagraphs(t)
    ##         doKeystroke(T)
    ##     else:
    ##         print 'reformat selection (RS) requires a selection first'
    
    def do_DATE(self, Format=None, Action=None, **kw):
        """give today's date
    
        format maybe adapted, default = "%d/%m"    
        action maybe adapted, default = "print",
            alternatives are "speak"
        """
        formatStrip = 0
        if Format is None or Format == 0:
            if Action in ['speak']:
                formatStrip = 0
                Format = "%B %d"
            else:
                formatStrip = 1
                Format = "%m/%d"
            
        cdate = datetime.date.today()
        fdate = cdate.strftime(Format)
        if Action in [None, 0, "print"]:
            if formatStrip:
                fdate = fdate.lstrip('0')
                fdate = fdate.replace('/0', '/')
                print('stripped fdate: |%s|'% fdate)
            doKeystroke(fdate)
        elif Action in ['speak']:
            command = 'TTSPlayString "%s"'% fdate
            natlink.execScript(command)
        else:
            print('invalid Action for DATE: %s'% Action)
        return 1
    
    def do_TIME(self, Format=None, Action=None, **kw):
        """give current time
    
        Format maybe adapted, default = "%H:%M"    
        Action maybe adapted, default = "print",
            alternatives are "speak"
        """
        ctime = datetime.datetime.now().time()
        formatStrip = 0
        if Format is None or Format == 0:
            if Action in ['speak']:
                formatStrip = 0
                Format = "%H:%M"
            else:
                formatStrip = 1
                Format = "%H:%M"
            
        ftime = ctime.strftime(Format)
        if Action in [None, "print"]:
            doKeystroke(ftime)
        elif Action in ['speak']:
            command = 'TTSPlayString "%s"'% ftime
            natlink.execScript(command)
        else:
            print('invalid Action for DATE: %s'% Action)
            
        return 1
    
    Date = do_DATE
    
    def do_SPEAK(self, t, **kw):
        """speak text through TTSPlayString
        """
        command = 'TTSPlayString "%s"'% t
        natlink.execScript(command)
    
    def do_PRINT(self, t, **kw):
        """print text to Messages of Python Macros window
        """
        print(t)
    
    def do_PRINTALLENVVARIABLES(self, **kw):
        """print all environment variables to the messages window
        """
        print('-'*40)
        print('These can be used for "expansion", as eg %NATLINKDIRECTORY% in the grammar _folders and other places:')
        natlinkcorefunctions.printAllEnvVariables()
        print('-'*40)
    
    def do_PRINTNATLINKENVVARIABLES(self, **kw):
        """print all environment variables to the messages window
        """
        natlinkEnvVariables = natlinkstatus.AddNatlinkEnvironmentVariables()
        print('-'*40)
        print('These can be used for "expansion", as eg %NATLINKDIRECTORY% in the grammar _folders and other places:')
        for k in sorted(natlinkEnvVariables.keys()):
            print("%s\t%s"% (k, natlinkEnvVariables[k]))
        print('-'*40)
        
    def do_T(self, **kw):
        """return true only"""
        return 1
    
    def do_F(self, **kw):
        """return false only (empty action will do as well)"""
        return 
    
    def do_A(self, n, **kw):
        """print ascii code
    
        A 208 prints the ETH character
        
    
        """
        doKeystroke(chr(n))
        return 1
    
    def do_U(self, n, **kw):
        """print unicode code

        U Delta
        U 00cb  (Euml)
        U Alpha or U 0391 (same) (numbers are the hex code, like in the word insert symbol dialog.
        """
        if type(n) != str:
            Code = html.entities.name2codepoint.get(n, n)
            if type(Code) == int:
                # is in defs
                pass
            else:
                # some 4 letter hexcode:
                # print 'stringlike code: %s, convert to hex number'% Code
                exec("Code = 0x%s"% Code)
        elif type(n) == int:
            # intended as hex code, but by unimacro converted into int. 
            # convert to back to hex:
            # print 'numlike code, convert to hex: %s'% n
            exec("Code = 0x%.4d"% n)
        else:
            raise ValueError("action U, invalid type of %s (%s)"% (n, type(n)))
        # lookup in above file, if found we have an integer code
        # print 'U in: %s, Code: %s(type: %s)'% (n, Code, type(Code))
        if Code <256:
            # print 'do direct, ascii: %s, %s'% (Code, chr(Code))
            natlinkutils.playString(chr(Code))
            return
        u = chr(Code)
        # output through the clipboard with special code:
        unimacroutils.saveClipboard()
        #win32con.CF_UNICODETEXT = 13
        unimacroutils.setClipboard(u, format=13)
        natlinkutils.playString('{ctrl+v}')
        unimacroutils.restoreClipboard()    
        return 1
                    
    
    
    def do_MSG(self, *args, **kw):
        """message on screen"""
        t = ', '.join(args)
        return Message(t, **kw)
    
    def do_DOCUMENT(self, number=None, **kw):
        """switch to document (program specific) with number"""
    ##    print 'action: goto task: %s'% number
        prog, title, topchild, classname, hndle = unimacroutils.getProgInfo()
        if not prog:
            print('action DOCUMENT, no program in foreground: %s (%S)'% (prog, title))
            return
        if number:
            try:
                count = int(number)
            except ValueError:
                print('action DOCUMENT, invalid number: %s'% number)
                return
            if not count:
                print('action DOCUMENT, invalid number: %s'% number)
                return
    
            section = 'positions %s'% prog
            if not ini.get(section):
                print('no mouse positions defined for DOCUMENT action in program: %s'% prog)
                return
    
            mouseX1 = ini.getInt(section, 'mousex1')
            mouseY1 = ini.getInt(section, 'mousey1')
            mouseXdiff = ini.getInt(section, 'mousexdiff')
            mouseYdiff = ini.getInt(section, 'mouseydiff')
            mx = mouseX1 + (count-1)*mouseXdiff
            my = mouseY1 + (count-1)*mouseYdiff
    ##        print 'mx, my:', mx, my
            #print 'task to %s, %s'% (mx, my)
            unimacroutils.doMouse(0, 0, mx, my)
    ##        unimacroutils.shortWait()
    ##        unimacroutils.buttonClick()
        else:
            print('call action DOCUMENT with a number!')
            return
        return 1
    
        
    def do_TASK(self, number=None, **kw):
        """switch to task with number"""
    ##    print 'action: goto task: %s'% number
        prog, title, topchild, classname, hndle = kw['progInfo']
        if prog == 'explorer' and not title:
            doKeystroke('{esc}')
            unimacroutils.shortWait()
        if number:
            try:
                count = int(number)
            except ValueError:
                print('action TASK, invalid number: %s'% number)
                return
            if not count:
                print('action TASK, invalid number: %s'% number)
                return
    
            # extra to remove focus:        
            do_WINKEY("b")
            do_SW()
    
            mouseX1 = ini.getInt('positions', 'mousex1')
            mouseY1 = ini.getInt('positions', 'mousey1')
            mouseXdiff = ini.getInt('positions', 'mousexdiff')
            mouseYdiff = ini.getInt('positions', 'mouseydiff')
            mx = mouseX1 + (count-1)*mouseXdiff
            my = mouseY1 + (count-1)*mouseYdiff
            # print 'mx, my:', mx, my
            #print 'task to %s, %s'% (mx, my)
            unimacroutils.doMouse(0, 0, mx, my)
            # unimacroutils.longWait()
    ##        unimacroutils.shortWait()
    ##        unimacroutils.buttonClick()
        else:
            print('call action TASK with a number!')
        return 1
    
    def do_TOCLOCK(self, click=None, **kw):
        """position mouse on clock, which gives taskbar menu
        """
        x = ini.get('positions', 'clockx')
        y = ini.get('positions', 'clocky')
        try:
            x = int(x) or 0
            y = int(y) or 0
        except ValueError:
            x = y = 0
        if x and y:
            unimacroutils.doMouse(0,0,x,y,click)
            unimacroutils.Wait()
        else:
            print('invalid mouse position for clock, do "task position clock" from grammar _general')
        return 1
     
    def do_CLIPSAVE(self, **kw):
        """saves and empties the clipboard"""
        unimacroutils.saveClipboard()
        return 1
    
    def do_CLIPRESTORE(self, **kw):
        """saves and empties the clipboard"""
        unimacroutils.restoreClipboard()
        return 1
    
    def do_CLIPISNOTEMPTY(self, **kw):
        """returns 1 if clipboard is not empty
    
        should be done after a CLIPEMPTY
        restores the clipboard if 0
        """
        t = unimacroutils.getClipboard()
        if t:
            return 1
        D(1, 'empty clipboard found, restore and return')
        unimacroutils.restoreClipboard()
        
    def do_GETCLIPBOARD(self, **kw):
        """returns the contents of the clipboars"""
        return unimacroutils.getClipboard()
       
    def do_COPYNAME(self, **kw):
        """returns the name of a file or folder if windows explorer or #32770
        """
        print('abacadabra')
        return 'abacadabra'
    
        
    def do_IFWT(self, title, action, **kw):
        """unimacro shorthand command IfWindowTitleDoAction
        insert a standard wait in order to let the previous action be performed...
        """
        print('do_IFWT, %s, %s'% (title, action))
        self.do_W()
        return IfWindowTitleDoAction(title, action)
    
    def IfWindowTitleDoAction(self, title, action, **kw):
        """do an action only if the title matches the window title
        """
        if unimacroutils.matchTitle(title):
            # print 'title: %s, yes, action: %s'% (title, action)
            doAction(action)
        # else:
            # print 'title: %s, does not match'% title
        return 1
        
    def killWindow(self, action1='<<windowclose>>', action2='<<killletter>>', modInfo=None, progInfo=None, comingFrom=None):
        """Closes a window and asks automatically for confirmation
    
        The default action 1 is "{alt+f4}",
    
        The default action 2 is <<killletter>>, that can be changed to
        <<saveletter>>, in order to make this a "save and close window" 
        command
     
        """
        if not progInfo:
            progInfo = unimacroutils.getProgInfo(modInfo=modInfo)
        
        prog, title, topchild, classname, hndle = progInfo
            
        progNew = prog
        prevHandle = hndle
        doAction(action1, progInfo=progInfo, comingFrom=comingFrom)
        unimacroutils.shortWait()
        count = 0
        while count < 20:
            count += 1
            try:
                modInfo = natlink.getCurrentModule()
                print("modinfo through natlink: %s"% repr(modInfo))
            except:
                modInfo = autohotkeyactions.getModInfo()
                print("modinfo through AHK: %s"% repr(modInfo))
            progNew = unimacroutils.getProgName(modInfo)
                
            if progNew != prog: break
            hndle = modInfo[2]
            if hndle != prevHandle:
    
                if not unimacroutils.isTopWindow(hndle):
                    # child:
                    print('do action2: %s'% action2)
                    doAction(action2)
                elif topWindowBehavesLikeChild(modInfo):
                    print('topWindowBehavesLikeChild action2: %s'% action2)
                    doAction(action2)
                break
            
            unimacroutils.shortWait()
        else:
            # no break occurred, false return:
            return 0 
        return 1
   
    def topWindowBehavesLikeChild(self, modInfo):
        """return the result of the ini file dict
        
        cache the contents in topchildDict
        """
        #pylint:disable=W0702       
        global topchildDict
        if topchildDict is None:
            topchildDict = ini.getDict('general', 'top behaves like child')
            #print 'topchildDict: %s'% topchildDict
        if topchildDict == {}:
            return
        prog, title, topchild, classname, hndle = getProgInfo(modInfo)
        result = matchProgTitleWithDict(prog, title, topchildDict, matchPart=1)
        if result: return result
        className = win32gui.GetClassName(hndle)
        # print('className: %s'% className)
        if className in ["MozillaDialogClass"]:
            return True
                
    def childWindowBehavesLikeTop(self, modInfo):
        """return the result of the ini file dict
        
        input: modInfo (module info: (progpath, windowTitle, hndle) )
        cache the contents in childtopDict
        """
        global childtopDict
        if childtopDict is None:
            childtopDict = ini.getDict('general', 'child behaves like top')
            #print 'childtopDict: %s'% childtopDict
        if childtopDict == {}:
            return
        prog, title, topchild, _classname, _hndle = getProgInfo(modInfo)
        return matchProgTitleWithDict(prog, title, childtopDict, matchPart=1)

    
    
    def matchProgTitleWithDict(self, prog, title, Dict, matchPart=None):
        """see if prog is in dict, if so, check title with value(s)
        
        """
        if not Dict: return
        if prog in Dict:
            titles = Dict[prog]
            if type(titles) == str:
                titles = [titles]
            titles = [t.lower() for t in titles]
            title = title.strip().lower()
            if matchPart:
                for t in titles:
                    if title.find(t) >= 0:
                        return 1
                return
            # matchPart = False (default)
            if title in titles:
                return 1
    
    
    def do_ALERT(self, alert=1, **kw):
        micState = natlink.getMicState()
        if micState in ['on', 'sleeping']:
            natlink.setMicState('off')
        if alert:
            try:
                nAlert = int(alert)
            except ValueError:
                nAlert = 1
            for i in range(nAlert):
                natlink.execScript('PlaySound "'+baseDirectory+'\\ding.wav"')
        unimacroutils.Wait(0.1)
        if micState != 'off':
            natlink.setMicState(micState)
        return 1
    
    Alert = do_ALERT
    
    def do_WINKEY(self, letter=None, **kw):
        """call the winkeys.dll with one letter
        
        if a number is taken, this number is converted into a string"""
        winkey = win32con.VK_LWIN         # 91
        keyup = win32con.KEYEVENTF_KEYUP  # 2
        win32api.keybd_event(winkey, 0, 0, 0)  # key down
        try:
            if not letter is None:
                do_SSK(letter)
        finally:
            win32api.keybd_event(winkey, 0, keyup, 0)  # key up
        return 1
    
        #
        #dllFile = os.path.join(unimacroutils.getOriginalUnimacroDirectory(), "dlls", "DNSKeys.dll")
        #letter = str(letter)
        #if letter.lower() == "{tab}":
        #    letter = '\t'
        #if len(letter) != 1:
        #    print 'cannot do WINKEY with key unequal length 1: |%s|'% letter
        #    return
        #natlink.execScript('DllCall "%s","WinKey","%s"'% (dllFile,letter))
    
    Winkey = do_WINKEY        
    
    def do_TASKTOSCREEN(self, screennumber, winHndle=None, **kw):
        """call the monitorfunctions to put task to screen 0, 1, ... depending on the number of screens"""
        if winHndle is None:
            winHndle = win32gui.GetForegroundWindow()
        mon = monitorfunctions.get_nearest_monitor_window(winHndle)
        allMonitors = copy.copy(monitorfunctions.MONITOR_HNDLES)
        try:
            wantedMonitor = allMonitors[screennumber]
        except IndexError:
            print('wanted monitor: %s, allMonitors: %s'% (screennumber, allMonitors))
            return
        if wantedMonitor == mon:
            print('already on monitor %s (%s)'% (screennumber, wantedMonitor))
            return_WTC
        resize = monitorfunctions.window_can_be_resized(winHndle)
        monitorfunctions.move_to_monitor(winHndle, wantedMonitor, mon, resize)
        return 1
    
    def do_TASKOD(self, winHndle=None, **kw):
        """call the monitorfunctions to put task in other display"""
        if winHndle is None:
            winHndle = win32gui.GetForegroundWindow()
        mon = monitorfunctions.get_nearest_monitor_window(winHndle)
        otherMons = monitorfunctions.get_other_monitors(mon)
        if not otherMons:
            print('only one monitor found')
            return
        otherMon = otherMons[0]
        if len(otherMons) > 1:
            print('more than 1 other monitors found (%s), take first: %s'% \
                       (otherMons, otherMon))
        resize = monitorfunctions.window_can_be_resized(winHndle)
        monitorfunctions.move_to_monitor(winHndle, otherMon, mon, resize)
        return 1
    
    def do_TASKMAX(self, winHndle=None, **kw):
        """call the monitorfunctions maximize task"""
        if winHndle is None:
            winHndle = win32gui.GetForegroundWindow()
        monitorfunctions.maximize_window(winHndle)
        return 1
    def do_TASKMIN(self, winHndle=None, **kw):
        """call the monitorfunctions to minimize task"""
        if winHndle is None:
            winHndle = win32gui.GetForegroundWindow()
        monitorfunctions.minimize_window(winHndle)
        return 1
    def do_TASKRESTORE(self, winHndle=None, keepinside=1, **kw):
        """call the monitorfunctions to restore task, keep inside monitor by default"""
        if winHndle is None:
            winHndle = win32gui.GetForegroundWindow()
        monitorfunctions.restore_window(winHndle, keepinside=keepinside)
        return 1
    
    TaskOtherDisplay = do_TASKOD        
    
    
    # do emacs command:
    def do_EMACS(self, *args, **kw):
        """do emacs command in minibuffer"""
        doAction("{alt+x}")
        self.do_W()
        for a in args:
          doAction(a)
        self.do_W()
        doAction("{enter}")
        self.do_W()
        return 1
    
    
    # put quotes around text and escape quotes inside text.
    def checkTextInMessage(self, t):
        """change " \n and \r so execScript can hndle it
    
        especially for the test in a MessageBoxConfirm
        """
        if t.find('"') >= 0:
            t = '""'.join(t.split('"'))
        if t.find('\n') >= 0:
            t = t.replace('\n', '"+chr$(10)+"')
        if t.find('\r') >= 0:
            t = t.replace('\r', '"+chr$(13)+"')
        return t
    
    def stripTextInMessage(self, t):
        """strip " \n and \r, so execScript does not raise an error
    
        especially for the title in a MessageBoxConfirm
        """
        if t.find('"') >= 0:
            t = t.replace('"', '')
        if t.find('\n') >= 0:
            t = t.replace('\n', ' ')
        if t.find('\r') >= 0:
            t = t.replace('\r', ' ')
        return t
    
    
    # the icons in the msgboxconfirm:
    MsgboxConfirmIconDict = dict(critical=16, query=32, warning=48, information=64)
    
    def Message(self, t, title=None, icon=64, alert=None, switchOnMic=None, progInfo=None, comingFrom=None):
        """put message on screen
    
        from grammar, call through self.DisplayMessage, only in some circumstances
        this function is called
    
        but can also be called directly of from shorthand command MSG
        
        switchOnMic can be True, in which case the mic is temporarily switched on
        while the message box is on.
        """
        if icon in MsgboxConfirmIconDict:
            icon = MsgboxConfirmIconDict[icon]
        if icon not in list(MsgboxConfirmIconDict.values()):
            raise ValueError("Unimacro actions Message, invalid value for icon: %s"% icon)
    
        if type(t) != str:
            t = '\n'.join(t)
    
        tt = checkTextInMessage(t)
        title = title or "Unimacro Message"
        title = stripTextInMessage(title)
        if alert:
            do_ALERT(alert)
        micState = natlink.getMicState()
        #print "Message, initial mic state: %s"% micState
        
        if switchOnMic and micState in ['sleeping', 'off']:
                natlink.setMicState('on')
    
        try:
            natlink.execScript('MsgBoxConfirm "%s", %s, "%s"'% (tt, icon, title))
        except SyntaxError:
            print('execScript SyntaxError\n' \
                  'tt: %s\n' \
                  'icon: %s\n' \
                  'title: %s\n'% (tt, icon, title))
        newMicState = natlink.getMicState()
        if switchOnMic and micState != newMicState:
            natlink.setMicState(micState)
        
    do_MESSAGE = do_MSG
    
    def YesNo(self, t, title=None, icon=32, alert=None, defaultToSecondButton=0, progInfo=None, comingFrom=None):
        """put message on screen, ask for yes or no
    
        if yes return True    
    
        but can also be called directly of from shorthand command YESNO
    
        the microphone is switched on if not so already. After the test (execScript)
        it should be "on" when No was answered (return False)
        and "sleeping" when Yes was answered (return True)
        
        if the microphone is "off" after the execScript, the user probably switched it
        off, and a special 0 is returned.
        
        Note: the execScript is an old dvc script, with limited possibilities!
        """
        if icon in MsgboxConfirmIconDict:
            icon = MsgboxConfirmIconDict[icon]
        if icon not in list(MsgboxConfirmIconDict.values()):
            raise ValueError("Unimacro actions Message, invalid value for icon: %s"% icon)
    
        if type(t) != str:
            t = '\n'.join(t)
        tt = checkTextInMessage(t)
        title = title or "Unimacro Question"
        title = stripTextInMessage(title)
    
        if icon in MsgboxConfirmIconDict:
            icon = MsgboxConfirmIconDict[icon]
        if icon not in list(MsgboxConfirmIconDict.values()):
            raise ValueError("Unimacro actions YesNo, invalid value for icon: %s"% icon)
        icon += 4 # yes no button
        if defaultToSecondButton:
            icon += 256
    
        micState = natlink.getMicState()
        #print "YesNo, initial mic state: %s"% micState
    
        if alert:
            do_ALERT(alert)
    
        if micState != 'on':
            unimacroutils.Wait(0.05)
            natlink.setMicState('on')
        newMicState = natlink.getMicState()
        unimacroutils.Wait(0.1)
        ttt = tt
        for i in range(3):
            try:
                cmd = 'MsgBoxConfirm "%s", %s, "%s"\n\nSetMicrophone 1\nGoToSleep'% (ttt, icon, title)
                #print 'cmd: %s'% cmd
                natlink.execScript(cmd)
            except natlink.SyntaxError:
                print('execScript SyntaxError\n' \
                      'tt: %s\n' \
                      'icon: %s\n' \
                      'title: %s\n'% (tt, icon, title))
            unimacroutils.Wait(0.1)
            newMicState = natlink.getMicState()
            result = (newMicState == 'sleeping')
            if newMicState != 'off': break   # ok, either on or sleeping
            # try again (maximum 3 times)
            unimacroutils.Wait(0.05)
            natlink.setMicState('on')
            ttt = checkTextInMessage("Please do not switch off the microphone\nwhile (re)answering the question:\n\n")+tt
        else:
            raise UserWarning("microphone should not be switched off while answering the YesNo question\n(and you got 3 chances to answer correct)")
        if micState != newMicState:
            natlink.setMicState(micState)
            unimacroutils.Wait(0.05)
        return result
    
    
    do_YESNO = YesNo
    
    # ##QH13062003  cursor positioning and other switch things for voicecoder--------------
    # def putCursor(self, ):
    #     """just put cursor text at cursor or after selection"""
    #     doAction("{end}%s"% cursorText)
    #     
    # 
    # def findCursor():
    #     """find the previous entered cursor text"""
    #     doAction('<<startsearch>>; "%s"; VW; <<searchgo>>'% cursorText)
    #     prog, title, topchild, classname, hndle = unimacroutils.getProgInfo()
    #     if prog == 'emacs':
    #         doAction("{shift+left %s}"% len(cursorText))
    #     doAction("CLIPSAVE; <<cut>>")
    #     t = unimacroutils.getClipboard()
    #     if t == cursorText:
    #         doAction("CLIPRESTORE")
    #         return 1
    #     else:
    #         print('invalid clipboard: %s'% t)
    #         doAction("<<paste>>; CLIPRESTORE")
    # 
    
    bringups = {}
    # special:
    # voicecodeApp = 'emacs'
    
    def UnimacroBringUp(self, app, filepath=None, title=None, extra=None, modInfo=None, progInfo=None, comingFrom=None):
        """get a running copy of app in the foreground
    
        the full path can be set in section [bringup app], key path
        the appname can also be set in this section, key appname
        
        if filepath == None:
            if appname is in the foreground, no switching needs to be done.
            if app is called for, waiting is until app is in the foreground.
        
        if filepath points to a valid file, this is added to appArgs.
        
        special cases for app:
            open: look in [bringup open] section for application extension, if not found,
                    just open (with app set to None)
            edit: look in [bringup edit] section for application extension, if not found,
                  take default edit program from edit section.
                  
        if class or title is given, first attaching to a running instance is tried.
        """
        global bringups
        if checkForChanges:
            doCheckForChanges() # resetting the ini file if changes were made
    
        # intermediate app and special treatment:
        # for voicecode (which you can call with BRINGUP voicecode) you need
        # voicecodeApp ('emacs') and (optional, but in this case) function voicecodeBringUp
        #
        if filepath:
            app2 = None
            while app in ['open', 'edit']:
                if app == app2: break   #open = open...
                dummy, ext = os.path.splitext(filepath)
                section = 'bringup %s'% app
                ext = ext.strip('.')
                app2 = ini.get(section, ext)
                if app2:
                    app = app2
                else:
                    app = None
        else:
            D(1, 'starting UnimacroBringUp for app: %s'% app)
            specialApp = app + 'App'
            specialFunc = app + 'BringUp'
            if specialApp in globals() or specialFunc in globals():
                if specialApp in globals():
                    appS = globals()[specialApp]
            ##        print 'do special: %s'% appS
                    if not UnimacroBringUp(appS):
                        D(1, 'could not bringup: %s'% appS)
                        return
                if specialFunc in globals():
                    func = globals()[specialFunc]
            ##                print 'func: %s (%s)'% (specialFunc, func)
                    if type(func) != types.FunctionType:
                        raise UnimacroError("UnimacroBringUp, not a proper function: %s (%s)"% (specialFunc, func))
                    return func()
            
        # now the "normal" cases:
        # get possibly different name and path from inifile:
    
        if app:
            appName = ini.get("bringup %s"% app, "name") or app
            appPath = ini.get("bringup %s"% app, "path") or None
    
            if appPath:
                #print 'appPath: %s'% appPath
                if os.path.isfile(appPath):
                    appPath = os.path.normpath(appPath)
                else:
                    appPath2 = natlinkcorefunctions.expandEnvVariableAtStart(appPath)
                    if os.path.isfile(appPath2):
                        appPath = os.path.normpath(appPath2)
                    elif appPath.lower().startswith("%programfiles%"):
                        appPathVariant = appPath.lower().replace("%programfiles", "%PROGRAMW6432%")
                        appPath2 = natlinkcorefunctions.expandEnvVariableAtStart(appPath)
                        if os.path.isfile(appPath2):
                            appPath = os.path.normpath(appPath2)
                        elif os.path.isdir(appPath2):
                            appPath = os.path.join(appPath2, appName)
                            if os.path.isfile(appPath):
                                appPath = os.path.normpath(appPath)
                            else:
                                raise OSError('invalid path for PROGRAMFILES to PROGRAMW6432, app %s: %s (expanded: %s)'% (app, appPath, appPath2))
                        else:
                            raise OSError('invalid path for app with PROGRAMFILES %s: %s (expanded: %s)'% (app, appPath, appPath2))
                    else:
                        raise OSError('invalid path for  app %s: %s (expanded: %s)'% (app, appPath, appPath2))
            else:
                appPath = appName or app
            appArgs = ini.get("bringup %s"% app, "args") or None
        else:
            appPath = None
            appArgs = None
            appName = None
        if appName:
            if filepath:
                appName = appName + " " + filepath
        else:
            appName = filepath
    
        if filepath:
            filepath = '""'+filepath+'""'
            if appArgs:
                #if filepath.find(" ") > 0:
                    # insert DOUBLE DOUBLE QUOTES for vba line recognition
                filepath = '""'+filepath+'""'
                appArgs = appArgs + " " + filepath
            else:
                appArgs = filepath
        # for future:
        appWindowStyle = ini.get("bringup %s"% app, "style") or None
        appDirectory = ini.get("bringup %s"% app, "directory") or None
        
        if not filepath:
            # for attaching to a running instance:
            appTitle = ini.get("bringup %s"% app, "title") or None
            appClass = ini.get("bringup %s"% app, "class") or None
            ## TODOQH
            prog, title, topchild, classname, hndle = unimacroutils.getProgInfo()
            progFull, titleFull, hndle = natlink.getCurrentModule()
        
            if self.windowCorrespondsToApp(app, appName, prog, title):
                D(1, 'already in this app: %s'% app)
                if app not in bringups:
                    bringups[app] = (prog, title, hndle)
                return 1
        
            if app in bringups:
                # try to simply switchto previous:
                try:
                    do_RW()
                    hndle = bringups[app][2]
                    D(1, 'hndle to switch to: %s'% hndle)
                    if not unimacroutils.SetForegroundWindow(hndle):
                        print('could not bring to foreground: %s, exit action'% hndle)
                        
                    if do_WTC():
                        prog, title, topchild, classname, hndle = unimacroutils.getProgInfo()
                        if prog == appName:
                            return 1
                except:
                    D(1, 'error in switching to previous app: %s'% app)
                    D(2, 'bringups: %s'% bringups)
                    D(1, 'delete %s from bringups'% app)
                    del bringups[app]
                    
        result = unimacroutils.AppBringUp(appName, appPath, appArgs, appWindowStyle, appDirectory)

        if extra:
            doAction(extra)
            
        return result
    
    
    
    ## was:
    
    def getAppPath(self, app):
        """use from openFileDefault, get path of application or None
        """
        appPath = ini.get("bringup %s"% app, "path") or None
        return appPath
    
    def getAppName(self, app):
        """use from openFileDefault, get name of application or None (mainly for edit)
        """
        appName = ini.get("bringup %s"% app, "name") or None
        return appName
    
    def getAppForEditExt(self, ext):
        """use from openFileDefault, when choosing "edit", check with extension
        """
        section = "bringup edit"
        ext = ext.strip('.')
        appName = ini.get(section, ext) or ini.get(section, 'name')
        return appName
        
    
    def dragonpadBringUp(self):
        i = 0
        natlink.recognitionMimic(["Start", "DragonPad"])
        sleepTime = 0.3
        waitSteps = 10
        while i < waitSteps:
            i += 1
            prog, title, topchild, classname, hndle = unimacroutils.getProgInfo()
            if self.windowCorrespondsToApp('dragonpad', 'natspeak', prog, title):
                break
            self.do_W(sleepTime)
            if i > waitSteps/2:
                print('try to check for DP: %s (%s)'% (prog, i))
        else:
            print('could not bringup DragonPad after %s seconds'% sleepTime*waitSteps)
            return
        return 1
            
    def messagesBringUp(self):
        """switch to the messages from python macros window"""
        if autohotkeyactions.ahk_is_active():
            do_AHK("showmessageswindow.ahk")
            return 1
        
        if unimacroutils.switchToWindowWithTitle('Messages from python macros'):
            return 1
        if not unimacroutils.switchToWindowWithTitle('Messages from python macros'):
            raise UnimacroError("cannot bring messages window to front")
        return 1
    
    def windowCorrespondsToApp(self, app, appName, actualProg, actualTitle):
        """program, plus possibly title corresponds, so in correct window
        
    
        """
        if app == 'dragonpad':
            return (appName == actualProg and actualTitle.startswith("DragonPad"))
        else:
            return appName == actualProg
        
    
    def clearBringups(self):
        global bringups
        bringups.clear()
    
    do_CLEARBRINGUPS = clearBringups
    do_BRINGUP = UnimacroBringUp      
    
def doAction(action):
    """make an instance of Action class and perform action
    """
    _action = Action()
    _action.doAction(action)

def doKeystroke(keystrokes):
    """make an instance of Action class and do keystrokes
    """
    _action = Action()
    _action.doKeystroke(keystrokes)

def Message(mess, title=None, icon=64, alert=None, switchOnMic=None, progInfo=None, comingFrom=None):
    """make an instance of Action class and send a message
    """
    _action = Action()
    _action.Message(mess, title, icon, alert, switchOnMic, progInfo, comingFrom)


hasDoubleQuotes = re.compile(r'^".*"$')
hasSingleQuotes = re.compile(r"^'.*'$")
hasDoubleQuote = re.compile(r'"')
def _convertToDvcArg(t):
    #pylint:disable=R0911, C0321
    t = t.strip()
    if not t: return ''
    D(1, 'convertToDvcArg: |%s|'%t)

    # if input string is a number, return string directly
    try:
        _i = int(t)
        return t
    except ValueError:
        pass
    try:
        _f = float(t)
        return t
    except ValueError:
        pass

    # now proceeding with strings:    
    if hasDoubleQuotes.match(t):
        return t
    if hasSingleQuotes.match(t):
        if len(t) > 2:
            return '"' + t[1:-1] + '"'
        return ""
    if t.find('"') > 0:
        t = t.replace('"', '""')
    return '"%s"'%t

def _convertToPythonArg(t):
    t = t.strip()
    if not t:
        return ''
    D(1, 'convertToPythonArg: |%s|'%t)

    # if input string is a number, return string directly
    try:
        i = int(t)
        if t == '0':
            return 0
        if t.startswith('0'):
            print('warning convertToPythonArg, could be int, but assume string: %s'% t)
            return '%s'% t
        return i
    except ValueError:
        pass
    try:
        f = float(t)
        if t.find(".") >= 0:
            return f
        else:
            print('warning convertToPythonArg, can be float, but assume string: %s'% t)
            return '%s'% t
    except ValueError:
        pass

    # now proceeding with strings:    
    if hasDoubleQuotes.match(t):
        return t[1:-1]
    if hasSingleQuotes.match(t):
        return t[1:-1]
    return t

def convertToPythonArgs(text):
    """convert to numbers and strings,

    IF argument is enclosed in " " or ' ' it is kept as a string.

    """    
    text = text.strip()
    if not text:
        return None  # None
    L = text.split(',')
    L = [_convertToPythonArg(l) for l in L]
    return tuple(L)

def writeDebug(s):
    """write line to debug file
    """
    if debugSock:
        debugSock.write(s+'\n')
        debugSock.flush()
        print(f'_actions: {s}')
    else:
        print(f'_actions debug: {s}')
       

def D(level, message):
    if debug >= level:
        print(message)

if __name__ == '__main__':
    Act = Action()
    Result = Act('T')
    print(f'result of do_T: {result}')
    
    