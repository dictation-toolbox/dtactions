# (unimacro - natlink macro wrapper/extensions)
# (c) copyright 2003 Quintijn Hoogenboom (quintijn@users.sourceforge.net)
#                    Ben Staniford (ben_staniford@users.sourceforge.net)
#                    Bart Jan van Os (bjvo@users.sourceforge.net)
#
#
# unimacroactions.py (former actions.py)
#  written by: Quintijn Hoogenboom (QH softwaretraining & advies)
#  June 2003/restructure March 2021
#

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
import subprocess  # for calling a ahk script

from natlinkcore import inivars
from dtactions import monitorfunctions
from dtactions import messagefunctions
from dtactions import autohotkeyactions
from dtactions import sendkeys
# import natlinkcore.natlinkutils as natut
# import unimacro.natlinkutilsqh as unimacroutils
# from natlinkcore import natlink
# from natlinkcore import natlinkcorefunctions # extended environment variables....
# from natlinkcore import natlinkstatus
from natlinkcore import utilsqh
from dtactions.unimacro import unimacroutils
external_actions_modules = {}  # the modules, None if not available (for prog)
external_action_instances = {} # the instances, None if not available (for hndle)
     
class ActionError(Exception): pass
class KeystrokeError(Exception): pass
class UnimacroError(Exception): pass
# pendingMessage = ''
try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print(f'Run this module after "build_package" and "flit install --symlink"\n')
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

print(f'inifile: {inifile}')

whatFile = userDirectory/(__name__ + '.txt')
debugSock = None
debugFile = userDirectory/'dtactions_debug.txt'
samples = []

try:  #
    ini = inivars.IniVars(inifile)
except inivars.IniError:
    
    print('Error in unimacroactions inifile: %s'% inifile)
    m = str(sys.exc_info()[1])
    print('message: %s'% m)
    print('unimacroactions cannot work')
    pendingMessage = 'Please repair action.ini file\n\n' + m
    ini = None

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
    def __init__(self, action=None, pauseBA=None, pauseBK=None, 
                 progInfo=None, modInfo=None, sectionList=None, comment='', comingFrom=None,
                 debug=None):
        self.action = action
        self.pauseBA = pauseBA
        self.pauseBK = pauseBK
        self.progInfo = progInfo
        self.modInfo = modInfo
        self.sectionList = sectionList
        self.comment = comment
        self.comingFrom = comingFrom
        self.debug = debug or 0

    def __call__(self, action):
        return self.doAction(action)

    def doAction(self, action):
        global pendingMessage, checkForChanges
        topLevel = 0
        if self.comingFrom and self.comingFrom.interrupted:
            print('command was interrupted')
            return
        self.completeAction = action
        if self.debug > 4: D(f'doAction: {self.completeAction}')
    
        if not ini:
            checkForChanges = 1
            if pendingMessage:
                m = pendingMessage
                pendingMessage = ''
                Message(m, alert=1)
                D('no valid inifile for actions')
                return
        if checkForChanges:
            if debug > 5: D('checking for changes')
            doCheckForChanges() # resetting the ini file if changes were made
        if not ini:
            D('no valid inifile for actions')
            return
        # if self.modInfo is None:
        #     try:
        #         modInfo = natlink.getCurrentModule()
        #         # print("modInfo through natlink: %s"% repr(modInfo))
        #     except:
        #         modInfo = autohotkeyactions.getModInfo()
        #         # print("modInfo through AHK: %s"% repr(modInfo))
        #    
        # progNew = unimacroutils.getProgName(modInfo)    
        # 
        # if progInfo is None:
        #     progInfo = unimacroutils.getProgInfo(modInfo=modInfo)
        # #D('new progInfo: %s'% repr(progInfo))
        # prog, title, topchild, classname, hndle = progInfo
        # if sectionList == None:
        #     sectionList = getSectionList(progInfo)
        # if pauseBA == None:
        #     pauseBA = float(setting('pause between actions', '0', sectionList=sectionList))
        # if pauseBK == None:
        #     pauseBK = float(setting('pause between keystrokes', '0', sectionList=sectionList))
        # 
        # if debug: D('------new action: %(action)s, '% locals())
        # if debug > 3 and comment:
        #     D('extra info: %(comment)s'% locals())
        # if debug > 3: D('\n\tprogInfo: %(progInfo)s, '
        #                 'pause between actions: %(pauseBA)s' % locals())
        # if debug > 5: D('sectionList at start action: %s'% sectionList)
            
        # now get a list of the separate actions, separated by a ";":
        actionsList = list(inivars.getIniList(action))
        # if action.find('USC') >= 0:
        #     bList = []
        #     for a in actionsList:
        #         if a.find('USC') >= 0:
        #             bList.extend(a.split('USC'))
        #         else:
        #             bList.append(a)
        #     actionsList= [t.strip() for t in bList if t]
            
        if debug > 2: D(f'action: {action}, actionsList: {actionsList}')
        if not actionsList:
            return
        for a in actionsList:
            result = self.doPartialAction(a)
            if result:
                return result  # error message??
            if comingFrom and comingFrom.interrupted:
                return "interrupted"
            if self.pauseBA and a != actionsList[-1]:
                self.do_W(self.pauseBA)
        # no errors, return None
        
    def doPartialAction(self, action):
        """do a "partial action", keystroke, USC, etc.
        
        these are separated by ";", and in between modInfo and  pauseBA are checked
        
        return a message if an error occurs
        return None if all OK
        """
        
        if metaAction.match(action):  # exactly a meta action, <<....>>
            a = metaAction.match(action).group(1)
            aNew = getMetaAction(a, sectionList, progInfo)
            if type(aNew) == tuple:
                # found function
                func, number = aNew
                if type(func) in (types.FunctionType, types.UnboundMethodType):
                    func(number)
                    return 1
                else:
                    print('Error, not a valid action "%s" for "%s" in program "%s"'% (func, action, prog))
                    return
            if debug > 5: D('doing meta action: <<%s>>: %s'% (a, aNew))
            if aNew:
                aNewList = list(inivars.getIniList(aNew))
                res = 0
                for aa in aNewList:
                    if aa:
                        if debug > 3: D('\tdoing part of meta action: <<%s>>: %s'% (a, aa))
                        res = doAction(aa, completeAction=completeAction,
                                 pauseBA=pauseBA, pauseBK=pauseBK,
                                 progInfo=progInfo, sectionList=sectionList,
                                 comment=comment, comingFrom=comingFrom)
                        if not res: return
                return res
            elif aNew == '':
                if debug > 1: D('empty action')
                return 1
            else:
                if debug > 3: D('error in meta action: "%s"'% a)
                partCom = '<<%s>>'%a
                t = '_actions, no valid meta action: "%s"'% partCom
                if partCom != completeAction:
                    t += '\ncomplete command: "%s"'% completeAction
                raise ActionError(t)
    
        # try qh command:
        if action.find('(') > 0 and action.strip().endswith(')'):
            com, rest = tuple([t.strip() for t in action.split('(', 1)])
            rest = rest.strip()[:-1]
        elif action.find(' ') > 0:
            com, rest = tuple([t.strip() for t in action.split(' ',1)])
        else:
            com, rest = action.strip(), ''
    
        if 'do_'+com in globals():
            funcName = 'do_'+com
            args = convertToPythonArgs(rest)
            kw = {}
            kw['progInfo'] = progInfo
            kw['comingFrom'] = comingFrom
            func = globals()[funcName]
            if not type(func) == types.FunctionType:
                raise UnimacroError('appears to be not a function: %s (%s)'% (funcName, func))
            if debug > 5: D('doing USC command: |%s|, with args: %s and kw: %s'% (com, repr(args), kw))
            if debug > 1: do_W(debug*0.2)
            if args:
                result = func(*args, **kw)
            else:
                result = func(**kw)
                    
            if debug > 5: print('did it, result: %s'% result)
            if debug > 1: do_W(debug*0.2)
            # skip pause between actions
            #do_W(pauseBA)
            return result
    
        # try meta actions inside:
        if metaActions.search(action): # action contains meta actions
            As = [t.strip() for t in metaActions.split(action)]
            if debug > 5: D('meta actions: %s'% As)
            for a in As:
                if not a: continue
                res = doAction(a, completeAction,
                         pauseBA=pauseBA, pauseBK=pauseBK, 
                         progInfo=progInfo, sectionList=sectionList,
                         comment=comment, comingFrom=comingFrom)
                if not res: return
                do_W(pauseBA)
                # skip pause between actions
                #do_W(pauseBA)
            return 1
    
        # try natspeak command:
    
        if com in natspeakCommands:
            rest = convertToDvcArgs(rest)
            C = com + ' ' + rest
            if debug: D('do dvc command: |%s|'% C)
            if debug > 1: do_W(debug*0.2)
            natlink.execScript(com + ' ' + rest)
            if debug > 5: print('did it')
            if debug > 1: do_W(debug*0.2)
            # skip pause between actions
            #do_W(pauseBA)
            return 1
    
        # all the rest:
        if debug > 5: D('do string: |%s|'% action)
        if debug > 1: do_W(debug*0.2)
        if action:
            self.doKeystroke(action)
            if debug > 5: print('did it') 
            if debug > 1: do_W(debug*0.2)
            # skip pause between actions
            #do_W(pauseBA)
            # if topLevel: # first (nonrecursive) call,
            #     print('end of complete action')
            return 1
        else:
            if debug:
                print('empty keystrokes')
             #     
        
    def doKeystroke(self, keystrokes):
        global pendingMessage, checkForChanges
        #print 'doing keystroke: {%s'% keystrokes[1:]
        if not keystrokes:
            return
    
        if not ini:
            checkForChanges = 1
            if pendingMessage:
                m = pendingMessage
                pendingMessage = ''
                Message(m, alert=1)
                D('no valid inifile for actions')
                return
        if checkForChanges:
            if debug > 5: D('checking for changes')
            self.doCheckForChanges() # resetting the ini file if changes were made# 
        if not ini:
            D('no valid inifile for keystrokes')
            self.hardKeys = ['none']
            self.pauseBK = 0
        # if type(hardKeys) != str:
        #     hardKeys = [hardKeys]
        # elif hardKeys == 1:
        #     hardKeys = ['all']
        # elif hardKeys == 0:
        #     hardKeys = ['none']
        else:    
            if self.sectionList == None:
                self.sectionList = self.getSectionList(self.progInfo)
            self.pauseBK = int(self.setting('pause between keystrokes', '0',
                                          sectionList=self.sectionList))
            self.hardKeys = self.getHardkeySettings()
            self.hardKeys = ['none']
            
        if self.pauseBK:
           l = hasBraces.split(keystrokes)
           for k in l:
               if not k:
                   continue
               elif braceExact.match(keystrokes):
                   doKeystroke(k, hardKeys=hardKeys, pauseBK = 0)
               else:
                   for K in k:
                       doKeystroke(K, hardKeys=hardKeys, pauseBK = 0)
               if debug > 5: D('pausing: %s msec after keystroke: |%s|'%
                               (pauseBK, k))
               # skip pausing between keystrokes
               #do_W(pauseBK)
        elif braceExact.match(keystrokes):
            # exactly 1 {key}:
            if debug > 5: D('exact keystrokes, hardKeys[0]: %s'% hardKeys[0])
            if hardKeys[0] == 'none':
                sendkeys.sendkeys(keystrokes)  # the fastest way
                return
            elif hardKeys[0] == 'all':
                ### must be implemented later
                sendkeys.sendkeys(keystrokes)
                return
            m = BracesExtractKey.match(keystrokes)
            if m:
                # the key part is known and valid word
                # eg tab, down etc
                keyPart = m.group(0).lower()
                for mod in 'shift+', 'ctrl+', 'alt+':
                    if keyPart.find(mod)>0: keyPart = keyPart.replace(mod, '')
                if keyPart in hardKeys:
                    if debug > 3: D('doing "hard": |%s|'% keystrokes)
                    ### TODOQH, later hard keys
                    sendkeys.sendkeys(keystrokes)
                    return
                else:
                    if debug > 3: D('doing "soft" (%s): |%s|'% (keystrokes, hardKeys))
                    sendkeys.sendkeys(keystrokes)
                    return
            else:
                if debug > 3: D('doing "soft" (%s): |%s|'% (keystrokes, hardKeys))
                sendkeys.sendkeys(keystrokes)
                return
            
        # now proceed with more complex keystroke possibilities:
        if self.hardKeys[0]  == 'all':
            ## TODOQH, hard keys
            # natut.playString(keystrokes, natut.hook_f_systemkeys)
            sendkeys.sendkeys(keystrokes)
            return
        elif self.hardKeys[0]  == 'none':
            sendkeys.sendkeys(keystrokes)
            return
        if hasBraces.search(keystrokes):
            keystrokeList = hasBraces.split(keystrokes)
            for k in keystrokeList:
                if debug > 5: D('part of keystrokes: |%s|' % k)
                if not k: continue
                #print 'recursing? : %s (%s)'% (k, keystrokeList)
                sendkeys.sendkeys(keystrokes)
                # self.doKeystroke(k, hardKeys=hardKeys, pauseBK = 0)
        else:
            if debug > 5: D('no braces keystrokes: |%s|' % keystrokes)
            sendkeys.sendkeys(keystrokes)
            ##T
    def getMetaAction(self, ma, sectionList=None, progInfo=None):
        """return the action that is found in the ini file.
        
        if 
        """
        m = metaNumber.search(ma)
        if m:
            number = m.group(1)
            A = ma.replace(number, 'n')
            actionName = a.replace(number, '')
            actionName = actionName.replace(' ', '')
        else:
            A = ma
            number = 0
            actionName = a.replace(' ', '')
        # try via actions_prog module:
        ext_instance = self.get_instance_from_progInfo(progInfo)
        if ext_instance:
            prog = progInfo[0]
            funcName = 'metaaction_%s'% actionName
            func = getattr(ext_instance,funcName, None)
            if func:
                if debug > 1:
                    D('action by function from prog %s: |%s|(%s), arg: %s'% (prog, actionName, funcName, number))
                result = func, number
                if result: return result
                # otherwise go on with "normal" meta actions...
    
        # no result in actions_prog module, continue normal way:
        if debug > 5: D('search for action: |%s|, sectionList: %s' %
                        (A, sectionList))
        
        aNew = setting(A, default=None, sectionList=sectionList)
        if aNew == None:
            print('action: not found, meta action for %s: |%s|, searched in sectionList: %s' % \
                  (a, aNew, sectionList))
            return 
        if m:
            aNew = metaNumberBack.sub(number, aNew)
        if debug:
            section = ini.getMatchingSection(sectionList, A)
            D('<<%s>> from [%s]: %s'% (A, section, aNew)) 
        return aNew        
        
    # last one undocumented, only for version 7
    
    def getSectionList(self, progInfo=None):
        if not progInfo:
            progInfo = unimacroutils.getProgInfo()
        prog, title, topchild, classname, hndle = progInfo
        if debug > 5:
            D('search for prog: %s and title: %s' % (prog, title))
            D('type prog: %s, type title: %s'% (type(prog), type(title)))
        L = ini.getSectionsWithPrefix(prog, title)
        L2 = ini.getSectionsWithPrefix(prog, topchild)  # catch program top or program child
        for item in L2:
            if not item in L:
                L.append(item)
        L.extend(ini.getSectionsWithPrefix('default', topchild)) # catcg default top or default child
        if debug > 5: D('section list with progInfo: %s:\n===== %s' %
                        (progInfo, L))
                        
        return L
    
    def convertToDvcArgs(self, text):
        text = text.strip()
        if not text:
            return ''
        L = text.split(',')
        L = list(map(_convertToDvcArg, L))
        return ', '.join(L)
    
    hasDoubleQuotes = re.compile(r'^".*"$')
    hasSingleQuotes = re.compile(r"^'.*'$")
    hasDoubleQuote = re.compile(r'"')
    def _convertToDvcArg(self, t):
        t = t.strip()
        if not t: return ''
        if debug > 1: D('convertToDvcArg: |%s|'%t)
    
        # if input string is a number, return string directly
        try:
            i = int(t)
            return t
        except ValueError:
            pass
        try:
            f = float(t)
            return t
        except ValueError:
            pass
    
        # now proceeding with strings:    
        if hasDoubleQuotes.match(t):
            return t
        elif hasSingleQuotes.match(t):
            if len(t) > 2:
                return '"' + t[1:-1] + '"'
            else:
                return ""
        if t.find('"') > 0:
            t = t.replace('"', '""')
        return '"%s"'%t
    
    def convertToPythonArgs(self, text):
        """convert to numbers and strings,
    
        IF argument is enclosed in " " or ' ' it is kept as a string.
    
        """    
        text = text.strip()
        if not text:
            return    # None
        L = text.split(',')
        L = [self._convertToPythonArg[l] for l in L]
        return tuple(L)
    
    def _convertToPythonArg(self, t):
        t = t.strip()
        if not t: return ''
        if debug > 1: D('convertToPythonArg: |%s|'%t)
    
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
        elif hasSingleQuotes.match(t):
            return t[1:-1]
        else:
            return t
    ##    elif hasDoubleQuote.search(t):
    ##        return "'%s'"% t
    ##    else:
    ##        return '"%s"'% t
            
    
    def getFromIni(self, keyword, default='',
                    sectionList=None, progInfo=None):
        if not ini:
            return ''
        if sectionList == None:
            if progInfo == None: progInfo = unimacroutils.getProgInfo()
            prog, title, topchild, classname, hndle = progInfo
            sectionList = ini.getSectionsWithPrefix(prog, title) + \
                          ini.getSectionsWithPrefix('default', title)
            if debug > 5: D('getFromIni, sectionList: |%s|' % sectionList)
        value = ini.get(sectionList, keyword, default)
        if debug > 5: D('got from setting/getFromIni: %s (keyword: %s'% (value, keyword))
        return value
    
    setting = getFromIni

    def getHardkeySettings(self):
        """get from ini the keystrokes that should be done as "hardKeys"
        
        hardKeys are (with Natlink) implemented as SendSystemKeys
        """
        if not self.sectionList:
            raise ValueError(f'action, getHardkeySettings, sectionList should be filled')
        hardKeys = self.setting('keystrokes with systemkeys', 'none', sectionList=self.sectionList)
        if debug > 5: D(f'hardKeys setting: {self.hardKeys}')
    
        hardKeys = actionIsList.split(hardKeys)
        if hardKeys:
            hardKeys = [k.strip() for k in hardKeys]
            if debug > 5: D('hardKeys as list: |%s|'% hardKeys)
    
            if debug > 5: D('new keystokes: |%s|, hardKeys: %s, pauseBK: %s'%
                            (keystrokes, hardKeys, pauseBK))
    if debug > 5: D('doKeystroke, pauseBK: %s, hardKeys: %s'% (pauseBK, hardKeys))


    
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
            import traceback
            external_actions_modules[prog] = None
            # print 'get_external_module, no module found for: %s'% prog
     
    def get_instance_from_progInfo(self, progInfo=None):
        """return the correct instances for progInfo
        """
        progInfo = progInfo or self.progInfo
        prog, title, topchild, classname, hndle = progInfo
        prog = str(prog)
        if hndle in external_action_instances:
            instance = external_action_instances[hndle]
            instance.update(progInfo)
            return instance
        
        mod = self.get_external_module(prog)
        if not mod:
            # print 'no external module instance: %s'% progInfo[0]
            return # no module, no instance
        classRef = getattr(mod, '%sActions'% prog.capitalize())
        if classRef:
            instance = classRef(progInfo)
            instance.update(progInfo)
            print('new instance for actions for prog: %s, hndle: %s'% (prog, hndle))
        else:
            instance = None
        external_action_instances[hndle] = instance
    
        return instance
    
    def doCheckForChanges(self, previousIni=None):
        global  ini, iniFileDate, topchildDict, childtopDict
        newDate = unimacroutils.getFileDate(inifile)
        if newDate > iniFileDate:
            D('----------reloading ini file')
            try:
                topchildDict = None
                childtopDict = None
                ini = inivars.IniVars(inifile)
            except inivars.IniError:
                msg = 'repair actions ini file: \n\n' + str(sys.exc_info()[1])
                #win32api.ShellExecute(0, "open", inifile, None , "", 1)
                time.sleep(0.5)
                Message(msg)
                iniFileDate = newDate
                ini = previousIni
            iniFileDate = newDate
    
    
    def writeDebug(self, s):
        if debugSock:
            debugSock.write(s+'\n')
            debugSock.flush()
            print(f'_actions: {s}')
        else:
            print(f'_actions debug: {s}')
           
    D =writeDebug
    def debugActions(self, n, openMode='w'):
        global debug, debugSock
        debug = n
        print('setting debug actions: %s'% debug)
        if debugSock:
            debugSock.close()
            debugSock = None
        if n:
            debugSock = open(debugFile, openMode)
            
            
    
    
    def debugActionsShow(self):
        print(f'opening debugFile: {debugFile} automatically is disabled.')
        #win32api.ShellExecute(0, "open", debugFile, None , "", 1)
        
    
    def showActions(self, progInfo=None, lineLen=60, sort=1, comingFrom=None, name=None):
        progInfo = progInfo or self.progInfo
        if not self.progInfo:
            progInfo = unimacroutils.getProgInfo()
            
        prog, title, topchild, classname, hndle = progInfo
        language = unimacroutils.getLanguage()
        
        sectionList = self.getSectionList(progInfo)
    
        l = [f'actions for program: {prog}\n\twindow title: {title} (topchild: {topchild}, classname: {classname}, hndle: {hndle}',
             '\tsection list: {sectionList}',
             '']
        l.append(ini.formatKeysOrderedFromSections(sectionList,
                                    lineLen=lineLen, sort=sort))
        
        T = getTranslation  # local function here, because no knowledge of the grammars is here
        l.append(T(language, dict(enx= '\n"edit actions" to edit the actions',
                                  nld='\n"bewerk acties" om de acties te bewerken')))
        l.append(T(language, dict(enx='"show actions" gives this list',
                                nld= '"toon acties" geeft deze lijst')))
        l.append(T(language, dict(enx='"trace actions (on|off|0|1|...|10)" sets/clears trace mode',
                                  nld='"trees acties (aan|uit|0|1|...|10)" zet trace modus aan of uit')))
        l.append(T(language, dict(enx='consult grammar "control" for the exact commands',
                                  nld='raadpleeg grammatica "controle" voor de precieze commando\'s')))
        
        open(whatFile, 'w').write('\n'.join(l))
        
        if comingFrom:
            name=name or ""
            comingFrom.openFileDefault(whatFile, name=name)
        else:
            print(f'show file: {whatFile}')
            print('showing actions by ShellExecute crashed NatSpeak, please open by hand above file')
        #win32api.ShellExecute(0, "open", whatFile, None , "", 1)
    
    def getTranslation(self, language, Dict):
        """get with self.language as key the text from dict, if invalid, return 'enx' text
        """
        if language in Dict:
            return Dict[language]
        else:
            return Dict['enx']
    
    
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
    
    def setPosition(name, pos, prog=None):
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
        ini.set(section, name, pos)
        unimacroutils.Wait(0.1)
        ini.write()
    
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
        if type(result) == str:
            do_MSG(result)
            return
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
            if debug > 3:
                D('words for HW(recognitionMimic): |%s|'% words)
            natlink.recognitionMimic(words)
        except:
            if words != origWords:
                try:
                    if debug > 3:
                        D('words for HW(recognitionMimic): |%s|'% origWords)
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
                natut.playString('{alt+n}')
                return 1
        natlink.recognitionMimic(["NaturallySpeaking"])
        return unimacroutils.waitForWindowTitle(['DragonBar', 'Dragon-balk', 'Voicebar'],10,0.1)
    
    
    # shorthand for sendsystemkeys:
    def do_SSK(self, s, **kw):
        natut.playString(s, natut.hook_f_systemkeys)
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
        keydown = natut.wm_syskeydown
        keyup = natut.wm_syskeyup
        
        altdown = (keydown, natut.vk_menu, 1)
        altup = (keyup, natut.vk_menu, 1)
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
        if debug > 7: D('waiting: %s'%t)
        elif debug and t > 2: D('waiting: %s'%t)
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
            natut.playString(chr(Code))
            return
        u = chr(Code)
        # output through the clipboard with special code:
        unimacroutils.saveClipboard()
        #win32con.CF_UNICODETEXT = 13
        unimacroutils.setClipboard(u, format=13)
        natut.playString('{ctrl+v}')
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
        D('empty clipboard found, restore and return')
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
        do_W()
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
    
    topchildDict = None
    childtopDict = None
    
    def topWindowBehavesLikeChild(self, modInfo):
        """return the result of the ini file dict
        
        cache the contents in topchildDict
        """
        global topchildDict
        if topchildDict is None:
            topchildDict = ini.getDict('general', 'top behaves like child')
            #print 'topchildDict: %s'% topchildDict
        if topchildDict == {}:
            return
        prog, title, topchild, classname, hndle = unimacroutils.getProgInfo(modInfo)
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
        prog, title, topchild, classname, hndle = unimacroutils.getProgInfo(modInfo)
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
        do_W()
        for a in args:
          doAction(a)
        do_W()
        doAction("{enter}")
        do_W()
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
            if debug: D('starting UnimacroBringUp for app: %s'% app)
            specialApp = app + 'App'
            specialFunc = app + 'BringUp'
            if specialApp in globals() or specialFunc in globals():
                if specialApp in globals():
                    appS = globals()[specialApp]
            ##        print 'do special: %s'% appS
                    if not UnimacroBringUp(appS):
                        D('could not bringup: %s'% appS)
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
                                raise IOError('invalid path for PROGRAMFILES to PROGRAMW6432, app %s: %s (expanded: %s)'% (app, appPath, appPath2))
                        else:
                            raise IOError('invalid path for app with PROGRAMFILES %s: %s (expanded: %s)'% (app, appPath, appPath2))
                    else:
                        raise IOError('invalid path for  app %s: %s (expanded: %s)'% (app, appPath, appPath2))
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
                if debug > 1: D('already in this app: %s'% app)
                if app not in bringups:
                    bringups[app] = (prog, title, hndle)
                return 1
        
            if app in bringups:
                # try to simply switchto previous:
                try:
                    do_RW()
                    hndle = bringups[app][2]
                    if debug: D('hndle to switch to: %s'% hndle)
                    if not unimacroutils.SetForegroundWindow(hndle):
                        print('could not bring to foreground: %s, exit action'% hndle)
                        
                    if do_WTC():
                        prog, title, topchild, classname, hndle = unimacroutils.getProgInfo()
                        if prog == appName:
                            return 1
                except:
                    if debug: D('error in switching to previous app: %s'% app)
                    if debug > 2: D('bringups: %s'% bringups)
                    if debug: D('delete %s from bringups'% app)
                    del bringups[app]
                    
        result = unimacroutils.AppBringUp(appName, appPath, appArgs, appWindowStyle, appDirectory)

        if extra:
            doAction(extra)
            
        return result
    
    def AutoHotkeyBringUp(self, app, filepath=None, title=None, extra=None, modInfo=None, progInfo=None):
        """start a program, folder, file, with AutoHotkey
        
        This functions is related to UnimacroBringup, which works with AppBringup from Dragon,
        but sometimes fails.
        
        Besides, this function can also work without Dragon being on, not relying on natlink.
        So better fit for debugging purposes.
        
        """
        scriptFolder = autohotkeyactions.GetAhkScriptFolder()
        if not os.path.isdir(scriptFolder):
            raise IOError('no scriptfolder for AHK: %s'%s)
        WinInfoFile = os.path.join(scriptFolder, "WININFOfromAHK.txt")
        
        ## treat mode = open or edit, finding a app in actions.ini:
        if ((app and app.lower() == "winword") or
            (filepath and (filepath.endswith(".docx") or filepath.endswith('.doc')))):
            script = autohotkeyactions.GetRunWinwordScript(filepath, WinInfoFile)
            result = autohotkeyactions.do_ahk_script(script)
    
        elif app and title:
            ## start eg thunderbird.exe this way
            ## can also give additional startup instructions in extra
            extra = extra or ""
            script = '''SetTitleMatchMode, 2
    Process, Exist, ##basename##
    if !ErrorLevel = 0
    {
    IfWinNotActive, ##title##,
    WinActivate, ##title##, 
    WinWaitActive, ##title##,,1
    if ErrorLevel {
        return
    }
    }
    else
    {
    Run, ##app##
    WinWait, ##title##,,5
    if ErrorLevel {
        MsgBox, AutoHotkey, WinWait for running ##basename## timed out
        return
    }
    }
    ##extra##
    WinGet pPath, ProcessPath, A
    WinGetTitle, Title, A
    WinGet wHndle, ID, A
    FileDelete, ##WININFOfile##
    FileAppend, %pPath%`n, ##WININFOfile##
    FileAppend, %Title%`n, ##WININFOfile##
    FileAppend, %wHndle%, ##WININFOfile##
    
    '''
            basename = os.path.basename(app)
            script = script.replace('##extra##', extra)
            script = script.replace('##app##', app)
            script = script.replace('##basename##', basename)
            script = script.replace('##title##', title)
            script = script.replace('##WININFOfile##', WinInfoFile)
            result = autohotkeyactions.do_ahk_script(script)
                
        else:
            ## other programs:
            if app and filepath:
                script = ["Run, %s, %s,,NewPID"% (app, filepath)]
            elif filepath:
                script = ["Run, %s,,, NewPID"% filepath]
            elif app:
                script = ["Run, %s,,, NewPID"% app]
        
            script.append("WinWait, ahk_pid %NewPID%")
        
            script.append("WinGet, pPath, ProcessPath, ahk_pid %NewPID%")
            script.append("WinGetTitle, Title, A")  ##, ID, ahk_pid %NewPID%")
            script.append("WinGet, wHndle, ID, ahk_pid %NewPID%")
            script.append('FileDelete, ' + WinInfoFile)
            script.append('FileAppend, %pPath%`n, ' + WinInfoFile)
            script.append('FileAppend, %Title%`n, ' + WinInfoFile)
            script.append('FileAppend, %wHndle%, ' + WinInfoFile)
            script = '\n'.join(script)
    
            result = autohotkeyactions.do_ahk_script(script)
    
        ## collect the wHndle:
        if result == 1:
            winInfo = open(WinInfoFile, 'r').read().split('\n')
            if len(winInfo) == 3:
                # make hndle decimal number:
                pPath, wTitle, hndle = winInfo
                hndle = int(hndle, 16)
                print('extracted pPath: %s, wTitle: %s and hndle: %s'% (pPath, wTitle, hndle))
                return pPath, wTitle, hndle
            else:
                if natlink.isNatSpeakRunning():
                    mess = "Result of ahk_script should be a 3 item list (pPath, wTitle, hndle), not: %s"% repr(winInfo)
                    do_MSG(str(mess))
                print(str())
                return 0
        else:
            if natlink.isNatSpeakRunning():
                do_MSG(str(result))
            print(str(result))
            return 0
    
    
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
            do_W(sleepTime)
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
    
    

if debug:
    try:
        debugSock = open(debugFile, 'w')
    except IOError:
        print('_actions, IOError, cannot write debug statements to: %s'% debugFile)
else:
    try:
        debugFile.unlink()
    except OSError:
        pass

if __name__ == '__main__':
    act = Action()
    result = act('T')
    print(f'result of do_T: {result}')
    
    