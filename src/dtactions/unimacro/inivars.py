"""contains class IniVars, that does inifiles

"""
#pylint:disable=C0116, C0302, R0911, R0912, R0914, R0915, R0901, R0902, R0904, R1702, C0209
import os
import os.path
import re
import copy
import locale
from collections import UserDict
from pathlib import Path

from dtactions.unimacro import utilsqh

from natlinkcore import readwritefile

locale.setlocale(locale.LC_ALL, '')

lineNum = 0
fileName = ''
class IniError(Exception):
    """Return the line number in the reading section, and the filename if there"""
    def __init__(self, value):
        self.value = value
        self.lineNum = lineNum
        self.fileName = fileName
        super().__init__(value)
        
    def __str__(self):
        s = ['Inivars error ']
        if fileName:
            s.append('in file %s, '%fileName)
        if lineNum:
            s.append('on line %s'%lineNum)
        s.append(': ')
        s.append(self.value)
        return ''.join(s)


DEBUG = 0
# doctest at the bottom
#reAllSections = re.compile(r'^\s*[([^]]+)\]\s*$', re.M)
#reAllKeys = re.compile(r'(^[^=\n]+)[=]', re.M)
reValidSection = re.compile(r'\[\b([- \.\w]+)]\s*$')
reFindKeyValue = re.compile(r'\b([- \.\w]+)\s*=(.*)$')
reValidKey = re.compile(r'(\w[\w \.\-]*)$')
reQuotes = re.compile(r'[\'"]')

reListValueSplit = re.compile(r'[\n;]', re.M)
reWhiteSpace = re.compile(r'\s+')

reDoubleQuotes = re.compile(r'^"([^"]*)"$', re.M)
reSingleQuotes = re.compile(r"^'([^']*)'$", re.M)

def quoteSpecial(t, extraProtect = None):
    """add quotes to string, to protect starting quotes, spaces

    extraProtect is possibly a string/list/tuple of characters that
    also need to be protected with quotes
    
>>> quoteSpecial("abc")
'abc'
>>> quoteSpecial(" abc  ")
'" abc  "'
>>> quoteSpecial("'abc' ")
'"\\'abc\\' "'

with newlines in text:
>>> quoteSpecial("'abc'\\n ")
'"\\'abc\\'\\n "'
>>> quoteSpecial(' "ab"')
'\\' "ab"\\''

but singular quotes inside a string are kept as they were:

>>> quoteSpecial('a" b "c')    
'a" b "c'

with both quotes:
>>> quoteSpecial('a" \\'b "c')    
'a" \\'b "c'

>>> quoteSpecial("'a\\" 'b \\"c'")    
'"\\'a&quot; \\'b &quot;c\\'"'

>>> quoteSpecial('" \\'b "c\\' xyz')    
'"&quot; \\'b &quot;c\\' xyz"'

>>> quoteSpecial('ab"')
'ab"'
>>> quoteSpecial("ab' '")
"ab' '"
>>> quoteSpecial("a' b 'c")    
"a' b 'c"

now for the list possibility:

>>> quoteSpecial("a;bc", ";\\n")
'"a;bc"'
>>> quoteSpecial("ab\xe9\\nc", ";\\n")
'"ab\\xe9\\nc"'

and for the dict possibility:

>>> quoteSpecial("abc,", ",")
'"abc,"'
>>> quoteSpecial("a,bc", ",")
'"a,bc"'
>>> quoteSpecial("a,bc", "'|;")
'a,bc'
    
    """
    if t is None:
        return ''
    if t.strip() != t:
        # leading or trailing spaces:
        if t.find('"') >= 0:
            if t.find("'") >= 0:
                t = t.replace('"', '&quot;')
                #raise IniError("string may not contain single AND double quotes: |%s|"% t)
            return "'%s'" % t
        return '"%s"'% t
    if t.find('"') == 0:
        if t.find("'") >= 0:
            t = t.replace('"', '&quot;')
            #print 'Inifile warning: text contains single AND double quotes:'
            return '"%s"'% t
        return "'%s'"% t
    if t.find("'") == 0:
        if t.find('"') >= 0:
            t = t.replace('"', '&quot;')
            #raise IniError("string may start with single quote AND contain double quotes: |%s|"% t)
        return '"%s"'% t
    if extraProtect:
        for c in extraProtect:
            if t.find(c) >= 0:
                if t.find('"') >= 0:
                    if t.find("'") >= 0:
                        raise IniError("string contains character to protect |%s| AND single and double quotes: |%s|"% (c, t))
                    return "'%s'"% t
                return '"%s"'% t
    return t

def quoteSpecialDict(t):
    """quote special with additional protection of comma

    """
    return quoteSpecial(t, ',')

def quoteSpecialList(t):
    """quote special with additional protection of semicolon, newline

    """
    return quoteSpecial(t, ';\n')

def stripSpecial(t):
    """strips text string, BUT if quotes or single quotes leaves rest

    new lines are preserved, BUT all lines are stripped    

>>> stripSpecial('""')
''
>>> stripSpecial('abc')
'abc'
>>> stripSpecial(" x ")
'x'
>>> stripSpecial("' a\xe9 '")
' a\\xe9 '
>>> stripSpecial('" a\xe9  "')
' a\\xe9  '

    """
    
    #if t.find('\n') >= 0:
    #    return '\n'.join(map(stripSpecial, t.split('\n')))
    #
    t = t.strip()
    if not t:
        return ''
    if reDoubleQuotes.match(t):
        r = reDoubleQuotes.match(t)
        inside = r.group(1)
        if inside.find('"') == -1:
            return inside
        raise IniError('invalid double quotes in string: |%s|'% t)
    if reSingleQuotes.match(t):
        r = reSingleQuotes.match(t)
        inside = r.group(1)
        if inside.find("'") == -1:
            return inside
        raise IniError('invalid single quotes in string: |%s|'% t)
    if t.find("'") == 0 or t.find('"') == 0:
        raise IniError('starting quote without matching end quote in string: |%s|'% t)
    return t

def getIniList(t, sep=(";", "\n")):
    """gets a list from inifile with quotes entries

    inside this function a state is maintained, for keeping track
    of the action to be taken if some character occurs:

    state = 0: start of string, or start after separator
    state = 1: normal string, including spaces
    state = 2: inside a quoted string, that everything go except
    the quote that is maintained in hadQuote
    state = 3: after a quoted string, waits for a separator or
    the end of the string. Raises error if anything else than a
    space his met

    >>> list(getIniList("a; c"))
    ['a', 'c']
    >>> list(getIniList("a;c"))
    ['a', 'c']
    >>> list(getIniList("a; c;"))
    ['a', 'c', '']
    >>> list(getIniList("'a\\"b'; c"))
    ['a"b', 'c']
    >>> list(getIniList('"a "; c'))
    ['a ', 'c']
    >>> list(getIniList("a ; ' c '"))
    ['a', ' c ']
    >>> list(getIniList("';a '; ' c '"))
    [';a ', ' c ']
    >>> list(getIniList("'a '\\n' c '"))
    ['a ', ' c ']

    """
    i =  0
    l = len(t)
    state = 0
    hadQuote = ''
##    print '---------------------------length: %s'% l
    for j in range(l):
        c = t[j]
##        print 'do c: |%s|, index: %s'% (c, j)
        if c in ("'", '"'):
            if state == 0:
                # starting quoted state
                state = 2
                hadQuote = c
                i = j + 1
                continue
            if state == 1:
                # quoted inside a string, let it pass
                continue
            if state == 2:
                if c == hadQuote:
                    yield t[i:j]
                    state = 3
                    continue
            if state == 3:
                raise IniError('invalid character |%s| after quoted string: |%s|'%
                               (c, t))
        if c in sep:
            if state == 0:
                # yielding empty string
                yield ''
            if state == 1:
                # end of normal string
                yield t[i:j].strip()
            if state == 2:
                # ignore separator when in quoted string
                continue
            # reset the state now:
            state = 0
            i = j + 1
            continue

        if c  == ' ':
            continue
        if state == 3:
            raise IniError('invalid character |%s| after quoted string: |%s|'%
                           (c, t))
        if state == 0:
            i = j
            state = 1
            continue

    if state == 0:
        # empty string at end
        yield ''
    if state == 1:
##        print 'yielding normal last part: |%s|, i: %s,length: %s'% (t[i:],i,l)
        yield t[i:].strip()
    if state == 2:
        raise IniError('no end of quoted string found: |%s|'% t)


def getIniDict(t):
    """gets a dict from inifile with generator function

    provides this generator with one list item (getDict should call getIniList first)
    each time a tuple pair (key, value) is returned,
    with value being a string or a list of strings, is separated by comma's

    If more keys are provided (separated by comma's),  these result
    in different yield statements    

    >>> list(getIniDict("a"))
    [('a', None)]
    >>> list(getIniDict("apricot:value of a"))
    [('apricot', 'value of a')]
    >>> list(getIniDict("a: ' '"))
    [('a', ' ')]
    >>> list(getIniDict('a: "with, comma"'))
    [('a', 'with, comma')]
    >>> list(getIniDict("a: 'with, comma'"))
    [('a', 'with, comma')]
    >>> list(getIniDict('a: more, "intricate, with, comma", example'))
    [('a', ['more', 'intricate, with, comma', 'example'])]
    >>> list(getIniDict("a: c, d"))
    [('a', ['c', 'd'])]
    >>> list(getIniDict("a, b: c"))
    [('a', 'c'), ('b', 'c')]
    >>> list(getIniDict("a,b : c, d"))
    [('a', ['c', 'd']), ('b', ['c', 'd'])]
    """
    if t.find('\n') >= 0:
        raise IniError('getIniDict must be called through getIniList, so newline chars are not possible: |%s|'%
                       t)
    if t.find(':') >= 0:
        Keys, Values = [item.strip() for item in t.split(':', 1)]
    else:
        Keys = t.strip()
        Values = ''

    if not Values:
        Values = None
    else:
        Values = list(getIniList(Values, ","))
        if len(Values) == 1:
            Values = Values[0]

    if not Keys:
        return

    if Keys.find(',') > 0:
        Keys = [k.strip() for k in Keys.split(',')]
        for k in Keys:
            if not reValidKey.match(k):
                raise IniError('invalid character in key |%s| of dictionary entry: |%s|'%
                               (k, t))
            yield k, Values
    else:
        if not reValidKey.match(Keys):
            raise IniError('invalid character in key |%s| of dictionary entry: |%s|'%
                           (Keys, t))
        yield Keys, Values
        
            
# obsolete python3        
# def lensort(a, b):
#     """sorts two strings, longest first, if length is equal, normal sort
# 
#     >>> lensort('a', 'b')
#     -1
#     >>> lensort('aaa', 'a')
#     -1
#     >>> lensort('zzz', 'zzzzz')
#     1
# 
#     """    
#     la = len(a)
#     lb = len(b)
#     if la == lb:
#         return cmp(a, b)
#     else:
#         return -cmp(la, lb)

class IniSection(UserDict):
    """represents a section of an inivars instance"""

    def __init__(self, parent):
        """init with ignore case if given in inivars

        """
        self._parent = parent
        self._returnStrings = None
        self._SKIgnorecase = parent._SKIgnorecase
        UserDict.__init__(self)

    def __getitem__(self, key):
        '''double underscore means internal value'''
        if not isinstance(key, str):
            raise TypeError('inivars, __getitem__ of IniSection expects str for key "%s", not: %s'% (key, type(key)))

        if key.startswith == '__':
            return self.__dict__[key]
        key = key.strip()
        if reWhiteSpace.search(key):
            key = reWhiteSpace.sub(' ', key)
        if self._SKIgnorecase:
            key = key.lower()
        try:
            value = UserDict.__getitem__(self, key)
            return value
        except KeyError:
            return "" 


    def __setitem__(self, key, value):
        key = key.strip()
        if reWhiteSpace.search(key):
            key = reWhiteSpace.sub(' ', key)
        if not key:
            raise IniError('__setitem__, invalid key |%s|(false), with value |%s| in inifile: %s'% \
                  (key, value, self._parent._file))
        
        if type(key) in (str, bytes):
            if key.startswith == '__':
                self.__dict__[key] = value
            else:
                if self._SKIgnorecase:
                    key = key.lower()
                UserDict.__setitem__(self, key, value)
        else:
            raise TypeError('inivars, IniSection __setitem__ expects str for key "%s", not: %s'% (key, type(key)))


    def __delitem__(self, key):
        key = key.strip()
        if reWhiteSpace.search(key):
            key = reWhiteSpace.sub(' ', key)
        if not key:
            raise IniError('__delitem__, invalid key |%s|(false),  inifile: %s'% \
                  (key, self._parent._file))
        
        if type(key) in (str, bytes):
            if key.startswith == '__':
                del self.__dict__[key]
            else:
                if self._SKIgnorecase:
                    key = key.lower()
                try:
                    UserDict.__delitem__(self, key)
                except KeyError:
                    pass
        else:
            raise TypeError('inivars, __delitem__ of IniSection expects str for key "%s", not: %s'% (key, type(key)))

    def __iter__(self):
        return iter(self.data)

            
        
class IniVars(UserDict):
    """do inivars from an .ini file, or possibly (extending this class)
    with other file types.

    File doesn't have to exist before.
    version 8:  all go to unicode, and subclass UserDict (november 2018)

    version 7:  quoting is allowed when spaces or special characters
                " or ' or ; in lists or  ; or , in dicts are found.

    version 6: change format of getDict:
               if key has value it is followed by a colon, more keys can
               be defined in one stroke

    version 6:
        options at start: sectionsLC = 1: convert all section names to lowercase (default 0)
                          keysLC = 0: convert all keys to lowercase (default 0)
        gettting of keys, with optional parameters
        

    version 5:
    In this version also a getDict, getList (was already), getInt
    getFloat and getTuple are defined.  When setting a dict, list, etc.
    inside the instance this type is conserved, when writing back
    into the inifile the appropriate formatting is done.

    version 4:    
    In this version 4 empty sections and keys are accepted, and set as
    [empty section] or

    [section]
    empty key =

    If the set command contains lists, elements are cycled through,
    If the sets command contains empty values these empty values are set like
    the example above.

    When an empty section or key is asked for [] or '' is returned.  You can see
    no difference if a section doesn't exist or is empty, as long has no default values
    are provided in the get function.

    When you get a value from a key through a list of sections, as soon as the key
    has been found (also when value is empty), this value is returned.

    second version:
    getList and setList are included

    third version, new methods:
    getSectionPostfixesWithPrefix:
                    sets a private dict _sectionPostfixesWP, see below
    getFromSectionsWithPrefix(prefix, text, key)
                    text is a string (eg a window title) in which the postfix
                    of the section name is found.
    getSectionsWithPrefix(prefix, longerText)
                    gets a list of section names, that contain prefix and postfix
                    matches a part of the longerText


    get is extended with a list of sections:
                    get(['a', 'b'], 'key'): the list of sections searched,
                    until a nonempty value is found.
                    get(['a', 'b']): returns all the keys in their respective sections,
                    without making duplicates

    getKeysOrderedFromSections gives a dictionary with the keys
                    of several sections ordered

    formatKeysOrderedFromSections gives a long string of the keys
                    of several sections ordered


    getMatchingSection gives from a section list the name of thefirst section
                    that contains the key. Returns "section not found" if not found.
                    
                
                    

>>> import os
>>> try: os.remove('empty.ini')
... except: pass
>>> ini = IniVars('empty.ini')
>>> ini.write()
>>> os.path.isfile('empty.ini')
1
>>> try: os.remove('simple.ini')
... except: pass
>>> ini = IniVars('simple.ini')
>>> ini.set('s', 'k', 'v')
>>> ini.set('s','k2','v2')
>>> ini.write()
>>> ini2 = IniVars('simple.ini')
>>> ini2.get()
['s']
>>> ini2.get('s')
['k', 'k2']
>>> ini2.get('s', 'k')
'v'
>>> ini == ini2
1
>>> ini.getFilename()
'simple.ini'
>>> ini2 = IniVars('simple.ini', returnStrings=True)
Warning inivars: returnStrings is Obsolete, inivars only returns Unicode
>>> ini2.get('s', 'k')
'v'
>>> ini2.get('s')
['k', 'k2']
    """
    _SKIgnorecase = None
    
    def __init__(self, File, **kw):
        """init from valid files, raise error if invalid file

        """
        UserDict.__init__(self)
        # add str function in case file is a path instance:
        self._repairErrors = kw.get('repairErrors', None)
        self._name = os.path.basename(File)
        self._file = File
        self._ext = self.getExtension(File)
        self._changed = 0
        self._maxLength = 60
        self._SKIgnorecase = kw.get('SKIgnorecase', None)
        self._codingscheme = None
        self._bom = None
        self._rawtext = ""
        self._returnStrings = False
        # instance to read (and write back) inifile...
        self._rwfile = readwritefile.ReadWriteFile()
        self._sectionPostfixesWP = {}   # ???
        if kw and 'returnStrings' in kw:
            value = kw['returnStrings']
            if value:
                print("Warning inivars: returnStrings is Obsolete, inivars only returns Unicode")
        if not self._ext:
            raise IniError('file has no extension: %s'% self._file)

        # start with new file:
        if not os.path.isfile(self._file):
            return
        readFuncName = '_read' + self._ext
        try:
            readFunc = getattr(self, readFuncName)
        except AttributeError as exc:
            raise IniError('file has invalid extension: %s'% self._file) from exc
        readFunc(self._file)

    def __bool__(self):
        """always true!"""
        return True

    def getFilename(self):
        return self._file
    def getName(self):
        return self._name

    def __contains__(self, key):
        if not isinstance(key, str):
            raise TypeError('inivars, __contains__ of IniVars expects str for key "%s", not: %s'% (key, type(key)))

        if key.startswith == '__':
            return key in self.__dict__
        return UserDict.__contains__(self, key)



    def __getitem__(self, key):
        if not isinstance(key, str):
            raise TypeError('inivars, __getitem__ of IniVars expects str for key "%s", not: %s'% (key, type(key)))

        if key.startswith('__'):
            return self.__dict__[key]
        key = key.strip()
        if reWhiteSpace.search(key):
            key = reWhiteSpace.sub(' ', key)
        if self._SKIgnorecase:
            key = key.lower()
        try:
            value = UserDict.__getitem__(self, key)
        except KeyError:
            value = None
        return value

    def __setitem__(self, key, value):
        
        # if type(key) == bytes:
        #     key = str(key)
        
        key = key.strip()
        if reWhiteSpace.search(key):
            key = reWhiteSpace.sub(' ', key)
        if not key:
            raise IniError('__setitem__, invalid key |%s|(false), with value |%s| in inifile: %s'% \
                  (key, value, self._file))
        
        if type(key) in (str, bytes):
            if key.startswith == '__':
                self.__dict__[key] = value
            else:
                if self._SKIgnorecase:
                    key = key.lower()
                UserDict.__setitem__(self, key, value)
        else:
            raise TypeError('inivars, __setitem__ of IniVars expects str for key "%s", not: %s'% (key, type(key)))

    def __delitem__(self, key):
        key = key.strip()
        if reWhiteSpace.search(key):
            key = reWhiteSpace.sub(' ', key)
        if not key:
            raise IniError('__delitem__, invalid key |%s|(false) in inifile: %s'% (key, self._file))
        
        if isinstance(key, str):
            if key.startswith == '__':
                del self.__dict__[key]
            else:
                if self._SKIgnorecase:
                    key = key.lower()
                try:
                    UserDict.__delitem__(self, key)
                except KeyError:
                    pass
        else:
            raise TypeError('inivars, __delitem__ of IniVars expects str for key "%s", not: %s'% (key, type(key)))
            
    def __iter__(self):
        return iter(self.data)
##    def _readPy(self, file):
##
##        # prepare file for importing:
##        d, f = os.path.split(file)
##        if not d in sys.path:
##            sys.path.append(d)
##        r, t = os.path.splitext(f)
##        mod = __import__(r)
##        d = {}o
##        for k,v in mod.__dict__.items():
##            if k[0:2] != '__' and type(v) == types.DictType:
##                d[k] = v
##        if DEBUG:  'read from py:', d
##        return d
#     def returnStringOrUnicode(self, returnValue):
#         """if option returnStrings is set to True, only return strings
#         
#         This function affects only Strings (binary or text)
#         other types are passed unchanged.
# >>> try: os.remove('returnunicode.ini')
# ... except: pass
# >>> ini = IniVars("returnunicode.ini")
# >>> ini.set("section","string","value")
# >>> ini.get("section")
# ['string']
# >>> ini.get("section", 'string')
# 'value'
# >>> ini.set("section","number", 12)
# >>> ini.get("section", 'number')
# 12
# >>> ini.set("section","none", None)
# >>> ini.get("section", 'none')
# >>> ini.get("section")
# ['none', 'string', 'number']
# >>> ini.get()
# ['section']
# 
# 
#  ## now returnString=True
# >>> try: os.remove('returnbinary.ini')
# ... except: pass
# >>> ini = IniVars("returnbinary.ini", returnStrings=True)
# >>> ini.set("section","string","value")
# >>> ini.get("section")
# ['string']
# >>> ini.get("section", 'string')
# 'value'
# >>> ini.set("section","number", 12)
# >>> ini.get("section", 'number')
# 12
# >>> ini.set("section","stringnone", None)
# >>> ini.get("section", 'stringnone')
# >>> ini.get("section")
# ['stringnone', 'string', 'number']
# >>> ini.get()
# ['section']
#         """
#         if type(returnValue) in (bytes, str):
#             if self._returnStrings and type(returnValue) == str:
#                 return utilsqh.convertToBinary(returnValue)
#             elif (not self._returnStrings) and type(returnValue) == bytes:
#                 return utilsqh.convertToUnicode(returnValue)
#             else:
#                 return returnValue
#         else:
#             return returnValue
                                          
        
    def _readIni(self, File):
        #pylint:disable=W0603
        global lineNum, fileName
        lineNum = 0
        fileName = File
        section = None
        sectionName = ''
        key = None
        keyName = ''
        sectionNameLines = {}
        self._rawtext = self._rwfile.readAnything(File)

        # rwfile.writeAnything(output_path, output_string)

        rawList = self._rawtext.split('\n')
        for line in rawList:
            line = line.rstrip()
            
            lineNum += 1
            m = reValidSection.match(line)
            if self._repairErrors and not m:
                # with repairErrors option, characters in front of a valid section:
                n = reValidSection.search(line)
                if n and not m:
                    print('ignore data in front of section: "%s"\n\t(please correct later in file "%s", line %s)'% (line,
                                                                fileName, lineNum))                                                 
                    m = n
                elif not sectionName:
                    if not line:
                        continue
                    print('no valid section found yet, skip line: "%s"\n\t(please correct later in file "%s", line %s)'% (line,
                                                                fileName, lineNum))                                                 
                    continue
            if m:
                sectionName = m.group(1).strip()
                if sectionName in self:
                    if self._repairErrors:
                        print('Warning: duplicate section "%s" on line %s and on line %s, take latter one\n\t(please correct later in file "%s")'% (
                            sectionName, sectionNameLines[sectionName], lineNum, fileName))
                        del self[sectionName]
                    else:
                        raise IniError('Duplicate section "%s" on line %s and on line %s\n\t(please correct in file "%s")'% (
                            sectionName, sectionNameLines[sectionName], lineNum, fileName))
                self[sectionName] = IniSection(parent = self)
                section = self[sectionName]
                key = None
                sectionNameLines[sectionName] = lineNum
                continue
            m = reFindKeyValue.match(line)
            if m:
                keyName = m.group(1).strip()
                if section is None:
                    raise IniError('no section defined yet')
                if keyName in section:
                    if self._repairErrors:
                        print('Warning: duplicate keyname "%s" in section %s on line %s, take latter one\n\t(please correct later in file "%s")'% (
                             keyName, sectionName, lineNum, fileName))
                        del section[keyName]
                    else:
                        raise IniError('Duplicate keyname "%s" in section %s on line %s\n\t(please correct in file "%s")'% (
                            keyName, sectionName, lineNum, fileName))
                    
                section[keyName] = [m.group(2)] 
                key = section[keyName]
                continue
            
            if key:
                # append to list of lines of key: stripping spaces etc.
                key.append(line.strip())
            elif line.strip():
                if section is None:
                    raise IniError('no key or section found yet')
                if key is None:
                    if self._repairErrors:
                        print('Warning: no key found in section "%s" on line %s, ignore\n\t(please correct later in file "%s")'% (
                             sectionName, lineNum, fileName))
                        continue
                    raise IniError('No key found in section "%s" on line %s\n\t(please correct in file "%s")'% (
                        sectionName, lineNum, fileName))

        
        for s in self:
            section = self[s]
            for k in section:
                section[k] = listToString(section[k])

    def writeIfChanged(self, File=None):
        if self._changed:
            self.write(File=File)
            self._changed = 0
            
    def write(self, File=None):
        if not File:
            File = self._file
            ext = self._ext
        else:
            ext = self.getExtension(File)
        if ext == 'Ini':
            self._writeIni(File)
        else:
            raise IniError('invalid extension for writing to File: %s'% File)

    def _writeIni(self, File):
        """writes to file of type ini"""
        L = []
        sections = self.get()
        sections.sort()
        hasTrailingNewline = 1   # no newline for section
        for s in sections:
            hadTrailingNewline = hasTrailingNewline
            hasTrailingNewline = 0
            if  not hadTrailingNewline:
                L.append('')    # prevent extra newlines at top or after multiline key
            L.append('[%s]'% s)
            keys = self.get(s)
            #  key char has '-', for sitegen:
            keyhashyphen = 0
            for k in keys:
                if k.find('-') > 0:
                    keyhashyphen = 1
                    break
            if keyhashyphen:
               # special case for sitegen:
                keys = sortHyphenKeys(keys)
            else:
                keys.sort()

            for k in keys:
                hadTrailingNewline = hasTrailingNewline
                hasTrailingNewline = 0
                v = self[s][k]
                if isinstance(v, int):
                    L.append('%s = %s' % (k, v))
                elif isinstance(v, float):
                    L.append('%s = %s' % (k, v))
                elif isinstance(v, bool):
                    L.append('%s = %s' % (k, str(v)))
                elif not v:
                    # print 'k: %s(%s)'% (k, type(k))
                    L.append('%s =' % k)
                        
                elif isinstance(v, (list, tuple)):
                    valueList = list(map(quoteSpecialList, v))
                    startString = '%s = '% k
                    length = len(startString)
                    listToWrite = []
                    for v in valueList:
                        if listToWrite and length + len(v) + 2 > self._maxLength:
                            L.append('%s%s' % (startString, '; '.join(listToWrite)))
                            listToWrite = [v]
                            startString = ' '*len(startString)
                            length = len(startString) + len(v)
                        else:
                            listToWrite.append(v)
                            length += len(v) + 2
                    if length > 72:
                        hasTrailingNewline = 1
                    L.append('%s%s' % (startString, '; '.join(listToWrite)))
                    L.append('')
                elif isinstance(v, dict):
                    inverse = {}
                    for K, V in list(v.items()):
                        if isinstance(V, str):
                            vv = quoteSpecialDict(V)                            
                        elif isinstance(V, (list, tuple)):
                            vv = ', '.join(map(quoteSpecialDict, V))
                        elif V is None:
                            vv = None
                        if vv in inverse:
                            inverse[vv].append(K)
                        else:
                            inverse[vv] = [K]
                    startString = '%s = '% k
                    length = len(startString)
                    if not inverse:
                        L.append('%s' % startString)
                    else:
                        if None in inverse:
                            L.append('%s%s' % (startString, ', '.join(inverse[None])))
                            del inverse[None]
                            startString = ' '*len(startString)
                        if inverse:
                            for K, V in list(inverse.items()):
##                                print 'writing back value: |%s|, startString: |%s|, keys: |%s|'% \
##                                      (K, startString, V)
                                L.append('%s%s: %s' % (startString, ', '.join(V), K))
                                startString = ' '*len(startString)
                    hasTrailingNewline = 1
                    L.append('')
                else:
                    if not isinstance(v, str):
                        v = str(v)
                    v = v.strip()
                    if v.find('\n') >= 0:
                        hasTrailingNewline = 1
                        V = v.split('\n')
                        if not hadTrailingNewline:
                            L.append('')    # 1 extra newline
                        L.append('%s =' % k)
                        spacing = ' '*4
                        for li in V:
                            if li:
                                L.append('%s%s' % (spacing, li))
                            else:
                                L.append('')
                        L.append('')
                    elif len(k) + len(v) > 72:
                        if not hasTrailingNewline:
                            L.append('')
                        L.append('%s = %s' % (k, v))
                        L.append('')
                        hasTrailingNewline = 1
                    else:
                        L.append('%s = %s' % (k, v))

            if not hadTrailingNewline:
                L.append('')
##        if path(self._file).isfile():
##            old = open(self._file).read()
##        else:
##            old = ""
        new = '\n'.join(L)
##        if old == new:

##            pass
####            print 'no changes'
##        else:
####            self.saveOldInifile()
        self._rwfile.writeAnything(File, new)

    def saveOldInifile(self):
        """make copy -1, ..., -9 for previous versions"""
        orgfile = self._file
        for i in range(1,10):
            newfile = Path('%s-%s'% (orgfile, i))
            if not newfile.is_file():
                break
        else:
            newfile.unlink()

    def get(self, s=None, k=None, value="", stripping=stripSpecial):
        """get sections, keys or values, in version 3 extended
        with s also being a list, examples of this far below.

        with version 6, strips more smart, see function stripSpecial below
        when setting a string, this smart quoting must also be performed.

        >>> import os
        >>> try: os.remove('get.ini')
        ... except: pass
        >>> ini = IniVars("get.ini")
        >>> ini.get()
        []
        >>> ini.get("s")
        []
        >>> ini.get()
        []
        >>> ini.get("s", "k")
        ''
        >>> ini.get("s", "k", "v")
        'v'

        spaces in section or key name are stripped and made single:
        
        >>> ini.set(" s  t ", "a  k ", 'strange')
        >>> ini.get("s t", 'a k')
        'strange'
        >>> ini.get("s  t ", 'a      k ')
        'strange'
        
        >>> 
        """
        ## override general class method get of UserDict class:
        #pylint:disable=W0237
        if isinstance(s, (list, tuple)):
            if k:
                k = k.strip()
                if reWhiteSpace.search(k):
                    k = reWhiteSpace.sub(' ', k)
                # look for the section with this key
                for S in s:

                    v = self.get(S, k, None, stripping=stripping)
                    if v is not None:
                        return v
                return value
            # get list of all keys with these sections:
            L = []
            for S in s:
                l = self.get(S)
                for _k in l:
                    if k not in L:
                        L.append(_k)
                return L
            # not found, return default:
            if k:
                return value
            # asking for list of possible keys:
            return L
        if s:
            s = s.strip()
            if reWhiteSpace.search(s):
                s = reWhiteSpace.sub(' ', s)

            if self.hasSection(s):
                # s exists, process key requests
                if k:
                    k = k.strip()
                    if reWhiteSpace.search(k):
                        k = reWhiteSpace.sub(' ', k)

                    # request value from s,k
                    if self.hasKey(s, k):
                        # k exists, return value
                        v = self[s][k]
                        if stripping and isinstance(v, str):
                            return stripping(v)
                        return v
                    # no key, return default value
                    return value
                # no key given, request a list of keys
                KeysList = list(self[s].keys())
                return KeysList
            if k:
                return value
            # s doesn't exist, return empty list of keys
            return []
        # no section, request section list
        Keys = list(self.keys())
        return Keys

    def set(self, s, k=None, v=None):
        """set section, key to value


        section can be a list, as can be the key.  If the value
        is not given, empty sections or keys are made.
        >>> try: os.remove('set.ini')
        ... except: pass
        >>> ini = IniVars("set.ini")
        >>> ini.set("section","key","value")
        >>> ini.get("section")
        ['key']
        >>> ini.get("section","key")
        'value'
        >>> ini.set('empty section')
        >>> ini.get()
        ['section', 'empty section']
        >>> ini.set(['empty 1', 'empty 2'])
        >>> ini.get('empty 1')
        []
        >>> ini.set(['not empty 1'], ['empty key 1', 'empty key 2'])
        >>> ini.get('not empty 1', 'empty key 1')
        >>> ini.set(['not empty 1'], ['key 1', 'key 2'], ' value ')
        >>> ini.get('not empty 1', 'key 1')
        ' value '
        >>> ini.set('quotes', 'double', '" a "')
        >>> ini.get('quotes', 'double')
        '" a "'
        >>> ini.set('quotes', 'single', "'  a  '")
        >>> ini.get('quotes', 'single')
        "'  a  '"
        >>> ini.close()
        >>> ini = IniVars("set.ini")
        >>> L = ini.get()
        >>> L.sort()
        >>> L
        ['empty 1', 'empty 2', 'empty section', 'not empty 1', 'quotes', 'section']
        >>> ini.get('not empty 1', 'key 1')
        ' value '
        >>> ini.get('not empty 1', 'empty key 1')
        ''
        >>> ini.get('empty 1')
        []
        >>> ini.get('quotes', 'single')
        "'  a  '"
        >>> ini.get('quotes', 'double')
        '" a "'
        """
        # self._changed = 1
        if isinstance(s, (list, tuple)):
            for S in s:
                self.set(S, k, v)
            return
        # now s in string:
        if not isinstance(s, str):
            s = str(s)

        # now make new section if not existing before:
        s = s.strip()
        if reWhiteSpace.search(s):
            s = reWhiteSpace.sub(' ', s).strip()

        if not reValidSection.match('['+s+']'):
            raise IniError("Invalid section name to set to: %s"% s)

        if not self.hasSection(s):
            self[s] = IniSection(parent = self)

        # no key: empty section:
        if k is None:
            if v is not None:
                raise IniError('setting empty section with nonempty value: %s'% v)
            return
            
        if isinstance(k, (list, tuple)):
            for K in k:
                self.set(s, K, v)
            return

        if not isinstance(k, str):
            k = str(k)
        if not isinstance(k, str):
            raise IniError('key must be list, tuple or string, not: %s (type: %s)'% (k, type(k)))
            
        # finally set s, k, to v

        k = k.strip()
        if reWhiteSpace.search(k):
            k = reWhiteSpace.sub(' ', k)
        if reQuotes.search(k):
            k = reQuotes.sub('', k)
        if not reValidKey.match(k):
            raise IniError('key contains invalid character(s): |%s|'% repr(k))
        
        # print "s: %s(%s), k: %s(%s), v: %s(%s)"% (s, type(s), k, type(k), v, type(v))
        if isinstance(v, str):
            v = quoteSpecial(v)
        prevV = None
        try:
            prevV = self[s][k]
        except KeyError:
            pass
        if v != prevV:
            self._changed = 1    
            self[s][k] = v

    def delete(self, s=None, k=None):
        """delete sections, keys or all

        >>> import os
        >>> try: os.remove('delete.ini')
        ... except: pass
        >>> ini = IniVars("delete.ini")
        >>> ini.set("section","key","value")
        >>> ini.get("section","key")
        'value'
        >>> ini.delete("section","key")
        >>> ini.get("section","key")
        ''
        >>> ini.get("section")
        []
        >>> ini.set("section","key","value")
        >>> ini.get("section","key")
        'value'
        >>> ini.delete()
        >>> ini.get("section")
        []
        >>> ini.get("section", "key")
        ''
        >>> ini.close()
        >>>
        """
        if s:
            s = s.strip()
        if s and reWhiteSpace.search(s):
            s = reWhiteSpace.sub(' ', s)

        if s:
            if self.hasSection(s):
                # s exists
                if k:
                    k = k.strip()
                if k:
                    if reWhiteSpace.search(k):
                        k = reWhiteSpace.sub(' ', k)

                    # delete given key
                    if self.hasKey(s, k):
                        # k exists, delete it
                        self._changed = 1
                        del self[s][k]
                    if not self[s]:
                        self._changed = 1
                        del self[s]
                else:
                    # no key given, delete whole section
                    self._changed = 1
                    del self[s]
        else:
            # no section, delete all sections and keys
            if self.toDict():
                self._changed = 1
                self.clear()

    def hasSection(self, s):
        if reWhiteSpace.search(s):
            s = reWhiteSpace.sub(' ', s).strip()
        if self._SKIgnorecase:
            s = s.lower()
        return s in self

    def hasKey(self, s, k):
        k = k.strip()
        if reWhiteSpace.search(k):
            k = reWhiteSpace.sub(' ', k)
        s = s.strip()
        if reWhiteSpace.search(s):
            s = reWhiteSpace.sub(' ', s)
        if self._SKIgnorecase:
            s = s.lower()
            k = k.lower()

        return self.hasSection(s) and k in self[s]

    def hasValue(self, s, k):
        k = k.strip()
        if reWhiteSpace.search(k):
            k = reWhiteSpace.sub(' ', k)
        s = s.strip()
        if reWhiteSpace.search(s):
            s = reWhiteSpace.sub(' ', s)
        if self._SKIgnorecase:
            s = s.lower()
            k = k.lower()
            
        return self.hasKey(s, k) and self[s][k]

    def close(self):
        """close the instance, write data back if changes were made

        """
        if self._changed:
            self.write()

    def getExtension(self, File):
        """capitalize text part of extension,
        .ini -> Ini
        .py -> Py 
        empty string -> ''
        """
        if File:
            extension = os.path.splitext(os.path.split(File)[1])[1]
            extension = extension[1:].capitalize()
            return extension
        return ''
    
##    def __call__(self, file):
##        """calling this class constructs a new instance,
##        with file as parameter
##        
##        result: instance
##        """
##        return IniVars(file)

    def getTuple(self, section, key):
        return tuple(self.getList(section, key))

    def getList(self, section, key, default=None):
        """get a value and convert into a list
        see example above and:
        
        >>> import os
        >>> try: os.remove('getlist.ini')
        ... except: pass
        >>> ini = IniVars('getlist.ini')
        >>> ini.set('s', 'empty')
        >>> ini.getList('s', 'empty')
        []
        >>> ini.set('s', 'one', 'value')
        >>> ini.getList('s', 'one')
        ['value']
        >>> ini.set('s', 'two', 'value; value')
        >>> ini.getList('s', 'two')
        ['value', 'value']
        >>> ini.set('s', 'three', 'value 1\\nvalue 2\\nvalue 3 and more')
        >>> ini.getList('s', 'three')
        ['value 1', 'value 2', 'value 3 and more']
        >>> 
        """
        if not self:
            return []
        try:
            value = self[section][key]
        except (KeyError, TypeError) as exc:
            if default is None:
                return []
            if isinstance(default, list):
                return list(default)
            raise IniError('invalid type for default of getList: |%s|%`default`') from exc
        if not value:
            return []
        if isinstance(value, list):
            return list(value)
        L = list(getIniList(value))
        return L

    def getDict(self, section, key, default=None):
        """get a value and convert into a dict

        first the value is split it into items like in the getList function,
        after that each item is examined:
        1. only one word: this is the key, value = None
        2. more words, look for the ":", before the ":" more keys can be defined
           at once, after the ":" is the value, two possibilities:
            a. no comma inside: value is a string
            b. comma is inside: value is a list, splitted by the comma

        >>> import os
        >>> try: os.remove('getdict.ini')
        ... except: pass
        >>> ini = IniVars('getdict.ini')
        >>> ini.set('s', 'empty')
        >>> ini.getDict('s', 'empty')
        {}
        
        default if base empty
        >>> ini.getDict('s', 'empty', {'s': None})
        {'s': None}

        setting and getting a dict:

        >>> d = {'a': 'b'}
        >>> e = {'empty': None}
        >>> ini.set('s', 'dict', d)
        >>> ini.set('s', 'empty value', e)
        >>> ini.getDict('s', 'dict')
        {'a': 'b'}
        >>> ini.getDict('s', 'empty value')
        {'empty': None}
        >>> ini.close()
        >>> ini = IniVars("getdict.ini")
        >>> ini.get('s', 'empty value')
        'empty'
        >>> ini.get('s', 'dict')
        'a: b'
        >>> ini.getDict('s', 'empty value')
        {'empty': None}
        >>> ini.getDict('s', 'dict')
        {'a': 'b'}
        
        setting and getting through direct string values:
        
        >>> ini.set('s', 'one', 'value')
        >>> ini.getDict('s', 'one')
        {'value': None}
        >>> ini.set('s', 'two', 'key: value 1\\n key2 : value 2')
        >>> ini.getDict('s', 'two')
        {'key': 'value 1', 'key2': 'value 2'}
        >>> ini.set('s', 'three', 'keyempty\\nkey1: value 1\\nkey2, key3: value1, value2')
        >>> ini.getDict('s', 'three')
        {'keyempty': None, 'key1': 'value 1', 'key2': ['value1', 'value2'], 'key3': ['value1', 'value2']}

        also the value of the value can be a list:
        
        >>> ini.set('s', 'list', 'key: a, b, c')
        >>> ini.getDict('s', 'list')
        {'key': ['a', 'b', 'c']}
        >>> ini.close()
        
        """
        try:
            value = self[section][key]
        except (TypeError, KeyError) as exc:
            if default is None:
                return {}
            if isinstance(default, dict):
                return dict(default)
            raise IniError('invalid type for default of getDict: |%s|%`default`') from exc
        if not value:
            if default is None:
                return {}
            if isinstance(default, dict):
                return dict(default)
            raise IniError('invalid type for default of getDict: |%s|%`default`')            
            
        if isinstance(value, dict):
            return copy.deepcopy(value)
        
        D = {}
        for listvalue in getIniList(value):
            for k, v in getIniDict(listvalue):
                if k in D:
                    raise IniError('duplicate key |%s| in getDict: %s'%
                                   (k, value))
                D[k] = v
            
        return D
            
    def getInt(self, section, key, default=0):
        """get a value and convert into a int


        >>> import os
        >>> try: os.remove('getint.ini')
        ... except: pass
        >>> ini = IniVars('getint.ini')
        >>> ini.set('s', 'three', 3)
        >>> ini.getInt('s', 'three')
        3
        >>> ini.getInt('s', 'unknown')
        0
        >>> ini.getInt('s', 'unknowndefault', 11)
        11
        
        """
        try:
            i = self[section][key]
        except (KeyError, TypeError):
            return default

        if isinstance(i, int):
            return i
        if isinstance(i, str):
            if i:
                try:
                    return int(i)
                except ValueError as exc:
                    raise IniError('ini method getInt, value not a valid integer: %s (section: %s, key: %s)'% (section, key, i)) from exc
            else:
                return default
        raise IniError('invalid type for getInt (probably intermediate set without write: %s)(section: %s, key: %s'%
                       (repr(i), section, key))

    def getBool(self, section, key, default=False):
        """get a value and convert into boolean

t, T, true, True, Waar, waar, w, W, 1 -->> True
empty, o, f, F, False, false, Onwaar, o, none -->> False
        
        """
        try:
            i = self[section][key]
        except (KeyError, TypeError):
            return default
        if not i:
            return False
        i = str(i)
        if i.lower()[0] in ['t', 'w', '1']:
            return True
        if i.lower()[0] in ['f', 'o', '0']:
            return False
        raise IniError('inivars, getBool, unexpected value: "%s" (section: %s, key: %s)'%
                         (i, section, key))

    def getFloat(self, section, key, default=0.0):
        """get a value and convert into a float


        >>> import os
        >>> try: os.remove('getfloat.ini')
        ... except: pass
        >>> ini = IniVars('getfloat.ini')
        >>> ini.set('s', 'three', '3')
        >>> ini.getFloat('s', 'three')
        3.0
        >>> ini.getFloat('s', 'unknown')
        0.0
        
        """
        try:
            i = self[section][key]
        except KeyError:
            return default
        if isinstance(i, float):
            return i
        if isinstance(i, str):
            if i:
                try:
                    return float(i)
                except ValueError as exc:
                    if ',' in i:
                        j = i.replace(',', '.')
                        try:
                            return float(j)
                        except ValueError as exc2:
                            raise IniError('ini method getFloat, value not a valid floating number (comma replaced to dot): %s (section: %s, key: %s)'%
                                       (section, key, i)) from exc2
                    else:
                        raise IniError('ini method getFloat, value not a valid floating number: %s (section: %s, key: %s)'%
                                   (section, key, i)) from exc
            else:
                return 0.0
        raise IniError('invalid type for getFloat (probably intermediate set without write: %s)(section: %s, key: %s'%
                       (repr(i), section, key))
    
        
                

    def getSectionPostfixesWithPrefix(self, prefix):
        """get all postfixes of sections, sorted, with a given prefix.

        A list is returned (and also set in the private variable _sectionPostfixesWP).
        So repeated calls can be handled quicker, BUT no refreshing is done!  A static
        ini instance is supposed.

        Prefix and postfix are separated by one space, and the list that is returned
        can be empty and is sorted by longest postfix first.

        So sections [p], [p f], [p ff] give as result (in _sectionPostfixesWP['p']):
        ['ff', 'f', '']
                

        >>> ini = IniVars('simple.ini')
        >>> ini.set('pref', 'key', '1')
        >>> ini.set('pref', 'm', '1')
        >>> ini.set('pref f', 'key', '2')
        >>> ini.set('pref f', 'l', '2')
        >>> ini.set('pref foo', 'key', '4')
        >>> ini.set('pref faa', 'key', '3')
        >>> ini.set('pref eggs', 'key', '5')
        >>> ini.set('pref eggs', 'k', '6')
        >>> ini.getSectionPostfixesWithPrefix('pref')
        ['eggs', 'faa', 'foo', 'f', '']
        >>> ini.getSectionPostfixesWithPrefix('pr')
        []

        
        Now go and search back, longest match first!
        >>> ini.getFromSectionsWithPrefix('pr', 'foo bar eggs', 'key')
        ''
 
        >>> ini.getFromSectionsWithPrefix('pref', 'foo bar eggs', 'key')
        '5'
        >>> ini.getFromSectionsWithPrefix('pref', 'egg foo bar', 'key')
        '4'
        >>> ini.getFromSectionsWithPrefix('pref', '', 'key')
        '1'
        >>> ini.getFromSectionsWithPrefix('pref', 'withfooandeggs', 'key')
        '5'
        >>> ini.getFromSectionsWithPrefix('pref', 'completely different', 'key')
        '2'
        >>> ini.getFromSectionsWithPrefix('pref', 'completely else', 'key')
        '1'
        >>> ini.getFromSectionsWithPrefix('pr', 'foo bar eggs', 'key')
        ''
        >>> ini.getFromSectionsWithPrefix('pref', 'foo bar eggs', 'l')
        '2'
        >>> ini.getFromSectionsWithPrefix('pref', 'foo bar eggs', 'm')
        '1'

        Get a list of sections that meeting the requirements:

        >>> ini.getSectionsWithPrefix('pref')
        ['pref eggs', 'pref faa', 'pref foo', 'pref f', 'pref']
        >>> ini.getSectionsWithPrefix('pref', 'abcd')
        ['pref']
        >>> ini.getSectionsWithPrefix('pref', 'foo')
        ['pref foo', 'pref f', 'pref']
        >>> ini.getSectionsWithPrefix('pref', 'egg')
        ['pref']
        >>> ini.getSectionsWithPrefix('pr', 'egg')
        []

        for exact finding the longerText can be a list or a tuple:
        >>> ini.getSectionsWithPrefix('pref', [])
        []
        >>> ini.getSectionsWithPrefix('pref', ['foo','eggs', ''])
        ['pref foo', 'pref eggs', 'pref']
        
        These lists can be used with the next get call,
        so if the section is a list,
        the entries of a list are searched until
        a non empty value is found!

        >>> L = ini.getSectionsWithPrefix('pref')
        >>> ini.get(L, 'key')
        '5'
        >>> L = ini.getSectionsWithPrefix('pref', 'this foo and another thing')
        >>> ini.get(L, 'key')
        '4'
        >>> ini.get([], 'key')
        ''
        >>> ini.get(['pref'], 'k')
        ''
        
        We can also extract a list of all possible keys,
        and also ordered in a dictionary, leaving out doubles.
        >>> L = ini.getSectionsWithPrefix('pref') # with all sections with prefix pref
        >>> ini.get(L)
        ['key', 'k', 'l', 'm']
        >>> ini.getKeysOrderedFromSections(L)
        {'pref eggs': ['key', 'k'], 'pref faa': [], 'pref foo': [], 'pref f': ['l'], 'pref': ['m']}
        >>> L = ini.getSectionsWithPrefix('pref', 'this foo and another thing') # with selection
        >>> L
        ['pref foo', 'pref f', 'pref']
        >>> ini.get(L)
        ['key', 'l', 'm']
        >>> ini.getKeysOrderedFromSections(L)
        {'pref foo': ['key'], 'pref f': ['l'], 'pref': ['m']}

        And format this dictionary into a long string:

        >>> ini.formatKeysOrderedFromSections(L)
        '[pref foo]\\nkey\\n\\n[pref f]\\nl\\n\\n[pref]\\nm\\n'

        >>> ini.formatKeysOrderedFromSections(L, giveLength=True)
        '[pref foo] (1)\\nkey\\n\\n[pref f] (1)\\nl\\n\\n[pref] (1)\\nm\\n'


        ## if you specify giveLength as int, only sections with more items than giveLength
        ## have the number added.
        >>> ini.formatKeysOrderedFromSections(L, giveLength=10)
        '[pref foo]\\nkey\\n\\n[pref f]\\nl\\n\\n[pref]\\nm\\n'
                
        For debugging purposes the section that matches a key
        from a section list can be retrieved:

        >>> ini.getMatchingSection(L, 'key')
        'pref foo'
        >>> ini.getMatchingSection(L, 'l')
        'pref f'
        >>> ini.getMatchingSection(L, 'm')
        'pref'
        >>> ini.getMatchingSection(L, 'no valid key')
        'section not found'


        """
        if prefix in self._sectionPostfixesWP:
            return self._sectionPostfixesWP[prefix]
        # call with new prefix, construct new entry:        
        l = []
        for s in self.get():
            if s.find(prefix) == 0:
                L = [item.strip() for item in s.split(' ', 1)]
                if L[0] == prefix: # proceed:
                    if len(L) == 2:
                        if ' '.join(L) != s:
                            raise IniError('getting section postfixes with prefix, exactly 1 space'
                                           'required bewtween prefix and postfix: %s'% s)
                        l.append(L[1]) # combined section key "prefix rest"
                    else:
                        l.append('')   # section key identical prefix, adding ''
        if l:
            l.sort(key=len)
            l.reverse()
        self._sectionPostfixesWP[prefix] = l
##        print 'new entry of _sectionPostfixesWP,%s: %s'% (prefix, l)
        return self._sectionPostfixesWP[prefix]

    def getFromSectionsWithPrefix(self, prefix, longerText, key):
        """get from all possible sections, as soon as longerText is found in postfix,
        the value is returned  longerText is required, and will often be a window title.

        examples see above

        """
        postfixes = self.getSectionPostfixesWithPrefix(prefix)
        
        for postfix  in postfixes:
            if postfix and longerText.find(postfix) >= 0:
                v = self.get(prefix + ' ' + postfix, key)
                if v:
                    return v
            elif not postfix:   # empty string, the section is [prefix] only
                v = self.get(prefix, key)
                if v:
                    return v
        return ''


    def getSectionsWithPrefix(self, prefix, longerText=None):
        """get all possible sections, as soon as (part of) longerText matches postfix,
        the value is returned. longerText will often be a window title.
        If longerText is not given, a list of all sections with this prefix is returned.

        examples and testing see above        

        """
        postfixes = self.getSectionPostfixesWithPrefix(prefix)
        #print 'postfixes: %s'% postfixes
        if longerText is None:
            L = [prefix + ' ' + postfix for postfix in postfixes]
        elif isinstance(longerText, str):
            # print 'longerText: %s (%s)'% (longerText, type(longerText))
            # for postfix in postfixes:
            #     print '---postfix: %s (%s)'% (postfix, type(postfix))
            L = [prefix + ' ' + postfix for postfix in postfixes
                                    if longerText.find(postfix) >= 0]
        elif isinstance(longerText, (list, tuple)):
            L = [prefix + ' ' + postfix for postfix in longerText if postfix in postfixes ]
        else:
            return [prefix]
        return [l.strip() for l in L]

    def getKeysOrderedFromSections(self, sectionList):
        """get all possible keys from a list of sections 

        A dictionary is returned, with keys being the section keys,
        and as values a list of keys that are taken from this section.

        examples and testing see above        

        """
        allKeys = []
        D = {}
        for s in sectionList:
            keys = self.get(s)
            D[s] = []
            for k in keys:
                if k not in allKeys:
                    D[s].append(k)
                    allKeys.append(k)
        return D

    def ifOnlyNumbersInValues(self, sectionName):
        """if the values are only numbers, return first and last key
        otherwise return None, see formatReverseNumbersDict
        """
        Keys = self.get(sectionName)
        if not Keys:
            return None
        D = {}
        for k in Keys:
            v = self.get(sectionName, k)
            if not v:
                return None
            try:
                v = int(v)
            except (ValueError, TypeError):
                return None
            # now we now v is an integer number
            D.setdefault(v, []).append(k)
        # when here, a numbers dict is found
        items = D
        return formatReverseNumbersDict(dict(items))

    def formatKeysOrderedFromSections(self, sectionList, lineLen=60, sort=1, giveLength=None):
        """formats in a long string all possible keys from a list of sections 

        The dictionary of the function "getKeysOrderedFromSections" is used,
        and the formatting is done with a generator function.

        examples and testing see above        

        """
        D = self.getKeysOrderedFromSections(sectionList)
        L = []
        for k in sectionList:
            if giveLength:
                lenValues = len(D[k])
                if isinstance(giveLength, int) and lenValues < giveLength:
                    L.append('[%s]'%k)
                else:
                    L.append('[%s] (%s)'% (k, len(D[k])))
            else:
                L.append('[%s]'%k)
            L.append(utilsqh.formatListColumns(D[k], lineLen=lineLen, sort=sort))
            L.append('')
        return '\n'.join(L)
    

    def getMatchingSection(self, sectionList, key):
        """gives the section that has the key

        """
        for s in sectionList:
            if key in self.get(s):
                return s
        return 'section not found'

    def getKeysWithPrefix(self, section, prefix, glue=None, includePrefix=None):
        """get keys within a section with a fixed prefix
        
        glue can be "-"
        if possible, the list is sorted numeric, if mixed, the numbers (until 99999) go first
        if includePrefix, prefix itself is inserted at start of resulting list
        (see unittest, testKeysWithPrefix)
        
        >>> ini = IniVars('simple.ini')
        >>> ini.set('example', 'key', '1')
        >>> ini.set('example', 'key-4', '2')
        >>> ini.set('example', 'key-12', '3')

        >>> ini.getKeysWithPrefix('example', 'key', glue='-')
        ['key-4', 'key-12']
        >>> ini.getKeysWithPrefix('example', 'key', glue='-')
        ['key-4', 'key-12']
        >>> ini.set('mixed', 'foo', '1')
        >>> ini.set('mixed', 'foo 4', '2')
        >>> ini.set('mixed', 'foo 12', '3')
        >>> ini.set('mixed', 'foo text', '4')
        >>> ini.getKeysWithPrefix('mixed', 'foo', includePrefix=1)
        ['foo', 'foo 4', 'foo 12', 'foo text']

        # if no glue prefix takes anything and no numeric sort, so mostly you will prefer a glue character
        # like ' ' or '-'/
        >>> ini.getKeysWithPrefix('mixed', 'fo')
        ['foo', 'foo 12', 'foo 4', 'foo text']

        # mismatch, if glue prefix must exactly match first word:
        >>> ini.getKeysWithPrefix('mixed', 'fo', glue=' ')
        []


        
        """
        if not prefix:
            return []
        lenprefix = len(prefix)
        rawlist = [k for k in self.get(section) if k.startswith(prefix)]
        if not rawlist:
            return []
        if prefix in rawlist:
            gotPrefix = 1
            rawlist.remove(prefix)
        else:
            gotPrefix = 0
            
        if not rawlist:
            if includePrefix and gotPrefix:
                return [prefix]
            return []

        dec = []
        for k in rawlist:
            if glue:
                try:
                    start, num = [item.strip() for item in k.rsplit(glue, 1)]
                except ValueError:
                    continue
            else:
                start, num = prefix, k[lenprefix:].strip()
            try:
                num = int(num)
            except ValueError:
                num = 99999
            if start == prefix:    
                dec.append( (num, start, k) )
        if not dec:
            return []
        dec.sort()
        resultList = [k for (num, start, k) in dec]
        if includePrefix and gotPrefix:
            resultList.insert(0, prefix)
        return resultList
        
    def toDict(self, section=None):
        """for testing, return contents as a pure dict"""
        if section:
            if section in self:
                return dict(self[section])
            return {}
        # whole inifile:
        D = {}
        for (k,v) in self.items():
            D[k] = dict(v)
        return D            

    def fromDict(self, Dict, section=None):
        """enter ini data from a dict"""
        if section:
            for k, v in Dict.items():
                self.set(section, k, v)
            return
        # complete file, double dict
        for _section, kv in Dict.items():
            if kv and isinstance(kv, dict):
                for k, v in kv.items():
                    self.set(_section, k, v)
            else:
                raise TypeError('inivars, fromDict, invalid value for section "%s", should be a dict: %s'%
                                (section, repr(kv)))

def sortHyphenKeys(keys):
    """sort keys, first by trailing number (-nnn), second by trunk name (after xx-)

    for use in siteGen, if no hyphens occur in any key, or not of form en-key or key-nnn or en-key-nnn
    (en = language code, en, fr, ...)
    (nnn = a numbenaar mijnr)
    nothing happens

>>> sortHyphenKeys(['second', 'numbered-2', 'numbered-10'])
['second', 'numbered-2', 'numbered-10']

>>> sortHyphenKeys(['a', 'bbb-ccc', 'x', 'bbb'])
['a', 'bbb', 'bbb-ccc', 'x']

>>> sortHyphenKeys(['second', 'no sort'])
['no sort', 'second']
>>> sortHyphenKeys(['single', 's', 'double', 'en-double'])
['s', 'single', 'double', 'en-double']

>>> sortHyphenKeys(['single', 's', 'double', 'en-double', 'double-2', 'en-double-2', 'double-10', 'en-double-10'])
['s', 'single', 'double', 'en-double', 'double-2', 'en-double-2', 'double-10', 'en-double-10']


>>> sortHyphenKeys(['triple', 'en-triple', 'fr-triple', 'en-triple-3', 'fr-triple-3', 'triple-13', 'fr-triple-13', 'single', 's', 'double', 'en-double', 'double-2', 'en-double-2', 'double-10', 'en-double-10'])
['s', 'single', 'double', 'en-double', 'triple', 'en-triple', 'fr-triple', 'double-2', 'en-double-2', 'triple-3', 'en-triple-3', 'fr-triple-3', 'double-10', 'en-double-10', 'triple-13', 'fr-triple-13']


    """
        
    D  = {}
    for k in keys:
        parts = k.split('-')
        try:
            index = int(parts[ - 1])
        except ValueError:
            D.setdefault(0, []).append(k)
        else:
            D.setdefault(index, []).append(k)
    mainKeys = list(D.keys())
    mainKeys.sort()
    sortedKeys = []
    for mainKey in mainKeys:
        sortedKeys.extend(sortLanguageKeys(D[mainKey]))
    return sortedKeys

def reverseDictWithDuplicates(D):
    """for testing, values are lists always
    """
    reverseD = {}
    for k,v  in list(D.items()):
        reverseD.setdefault(v, []).append(k)
    return reverseD

def formatReverseNumbersDict(D):
    """format as efficient as possible

    keys: the spoken words possibly with equivalents
    value: a list of numbers, to be ordered rising order
    
    (reverse with reverseDictWithDuplicates)
>>> D = dict([(1, ['one']), (2, ['two']), (3, ['three']), (4, ['four']), (5, ['five']), (6, ['six']), (7, ['seven']), (8, ['eight']), (9, ['nine'])])
>>> formatReverseNumbersDict(D)
'one ... nine'

    one ... twenty or one, two, twice, three ... ten
    """
    keys = list(D.keys())
    keys.sort()
    # print 'items: %s'% items
    it = utilsqh.peek_ahead(keys)
    kPrev = None
    L = []
    increment = 1
    for k in it:
        preview = it.preview
        if preview != it.sentinel:
            knext = preview
        if preview == it.sentinel or len(D[knext]) > 1 or knext != k + increment:
            # close current item
            if kPrev is None:
                if L:
                    L.append(', ')
                L.append(', '.join(D[k]))
            elif k - kPrev != increment:
                L.append(' ... %s'% ', '.join(D[k]))
                increment = k - kPrev
            else:
                # next item only:
                L.append(', %s'% ', '.join(D[k]))
            kPrev = None
        else:
            # there is a next, consisting of one text item
            if kPrev is not None:
                continue
            if len(D[k]) > 1:
                if L:
                    L.append(', ')
                L.append(', '.join(D[k]))
            else:
                # one item of longer possibly ... separated sequence
                if L:
                    L.append(', ')
                L.append('%s'% D[k][0])
                kPrev = k
    # for line in L:
    #     print 'line: %s (%s_)'% (line, type(line))
    return ''.join(L)

    ###???
    #minNum, maxNum = min(nums), max(nums)
    ##print 'min: %s, max: %s, nums: %s'% (minNum, maxNum, nums)
    #hadPrev = 0
    #L = []
    #for i in range(minNum, maxNum+1):
    #    v = D.get(i, None)
    #    if v == None:
    #        hadPrev = 0
    #    if len(v) > 1:
    #        L.append(', '.join(v))
    #    if hadPrev:
    #        i
    #        
    #if nums == range(minNum, maxNum+1):
    #    return (D[minNum], D[maxNum])


def sortLanguageKeys(keys):
    """sort single keys first, then language keys, with neutral on top
    
>>> sortLanguageKeys(['triple', 'en-triple', 'fr-triple', 'single', 's', 'double', 'en-double'])
['s', 'single', 'double', 'en-double', 'triple', 'en-triple', 'fr-triple']
    """
    _i = 1
    langKeys = [k for k in keys if len(k) > 2 and k[2] == '-']
    trunkKeys = {k[3:] for k in langKeys}
    trunkKeys = list(trunkKeys)
    nonLangKeys = [k for k in keys if k not in langKeys]
    singleKeys = [k for k in nonLangKeys if k not in trunkKeys]
    trunkKeys.sort()
    singleKeys.sort()
    total = singleKeys[:]
    for trunk in trunkKeys:
        # get trunk + accompanying language keys:
        accompany = [k for k in langKeys if k.find(trunk) == 3]
        accompany.sort()
        total.append(trunk)
        for k in accompany:
            total.append(k)
            langKeys.remove(k)
    if langKeys:
        langKeys.sort()
        total.extend(langKeys)
    return total

def listToString(valueList):
    """convert list of items (from _readIni) into a string

    utility for readIni function.    

    first strip off empty lines at beginning or end
    second if 1 line: return stripped state
    third: in more lines: return join of lines, only right strip now

>>> listToString([''])
''
>>> listToString(['   '])
''
>>> listToString(['', ' abc ', '', 'def    ', ''])
' abc\\n\\ndef'
>>> listToString(['', '  stripped ', '', '', '  ', '\\t \\t'])
'stripped'

    """
    if len(valueList) == 0:
        raise IniError('empty keylist')
    while len(valueList) > 0 and not valueList[0].strip():
        valueList.pop(0)
    while  len(valueList) > 0 and not valueList[-1].strip():
        valueList.pop()
    if len(valueList) == 1:
        return valueList[0].strip()
    return '\n'.join([v.rstrip() for v in valueList])

# patterns for runPythonCode:

def testReSearch(reExpression, texts):
    res = []
    for t in texts:
        if reExpression.search(t):
            res.append(1)
        else:
            res.append(0)
    return res

def testReSub(reExpression, texts, replacement):
    res = []
    for t in texts:
        if reExpression.search(t):
            res.append(reExpression.sub(replacement, t))
        else:
            res.append(t)
    return res

def testReMatch(reExpression, texts):
    res = []
    for t in texts:
        if reExpression.match(t):
            res.append(1)
        else:
            res.append(0)
    return res
def readExplicit(ini,mode=None ):
    """Read all sections and keys, and set them back in the ini instance

    Mode can also be list or dict.
    """
    for s in ini:
        for k in ini.get(s):
            if mode == "list":
                v = ini.getList(s, k)
            elif mode == "dict":
                v = ini.getDict(s, k)
            elif mode is None:
                v = ini.get(s, k)
            else:
                raise IniError('invalid mode for test function readExplicit: %s'% mode)
            ini.set(s, k, v)
        
retest = '''

>>> testReMatch(reValidSection,  ['[valid section]', '[a3]', '[x]','[3x]', '[x - y]', '[X. Y.]'])
[1, 1, 1, 1, 1, 1]


>>> testReMatch(reValidSection,  [' [invalid section]',  '[-x3]', '[.xyz]'])
[0, 0, 0]

>>> testReMatch(reValidKey,  ['valid key ',  '3x', 'A. B. C.'])
[1, 1, 1]


>>> testReMatch(reValidKey,  [' invalid key]',  'x[s]', '-x'])
[0, 0, 0]

>>> testReMatch(reFindKeyValue,  ['valid key = ',  '3x =', 'A. B. C. ='])
[1, 1, 1]


>>> testReMatch(reFindKeyValue,  [' invalid key=]',  'x[s]=', '-x='])
[0, 0, 0]


>>> testReSearch(reWhiteSpace,  [' abc', '  a', 'a  b'])
[1, 1, 1]



White Space:

>>> testReSearch(reWhiteSpace,  [' abc', '  a', 'a  b'])
[1, 1, 1]

>>> testReSub(reWhiteSpace, ['abc', ' a b', ' a b   c'], '')
['abc', 'ab', 'abc']

>>> testReSearch(reWhiteSpace,  ['abc'])
[0]


'''




goodfilesTest = '''
>>> r = 'projects/miscqh/testinivars'
>>> root = utilsqh.getRoot('c:/'+r, 'd:/'+r)
>>> test = root/'test'
>>> utilsqh.makeEmptyFolder(test)

basic testing:

>>> L = ['basic', 'multiplelines']
>>> for n in L:
...     ini = IniVars(root/(n+'.ini'))
...     readExplicit(ini)
...     ini.write(root/'test'/(n+'.ini'))
...     ini2 = IniVars(root/'test'/(n+'.ini'))
...     readExplicit(ini2)
...     print 'testFile: %s, result: %s'% (n, ini == ini2)
testFile: basic, result: True
testFile: multiplelines, result: True

list:

>>> L = ['list']
>>> for n in L:
...     ini = IniVars(root/(n+'.ini'))
...     readExplicit(ini, mode = 'list')
...     ini.write(root/'test'/(n+'.ini'))
...     ini2 = IniVars(root/'test'/(n+'.ini'))
...     readExplicit(ini2, mode = 'list')
...     print 'testFile, list: %s, result: %s'% (n, ini == ini2)
testFile, list: list, result: True



dict:

>>> L = ['dict']
>>> for n in L:
...     ini = IniVars(root/(n+'.ini'))
...     readExplicit(ini, mode = 'dict')
...     ini.write(root/'test'/(n+'.ini'))
...     ini2 = IniVars(root/'test'/(n+'.ini'))
...     readExplicit(ini2, mode = 'dict')
...     print 'testFile, dict: %s, result: %s'% (n, ini == ini2)
testFile, dict: dict, result: True


whatSKIgnorecase:

>>> L = ['uckey', 'ucsection']
>>> for n in L:
...     ini = IniVars(root/(n+'.ini'), SKIgnorecase=1)
...     readExplicit(ini)
...     ini.set('New Section', 'New Key', 'New Value')
...     ini.write(root/'test'/(n+'.ini'))
...     ini2 = IniVars(root/'test'/(n+'.ini'))
...     readExplicit(ini2)
...     print 'ini2, sections, keys lowercase:%s'% ini2
...     print 'testFile, SKIgnorecase: %s, result: %s'% (n, ini == ini2)
ini2, sections, keys lowercase:{'section': {'key with capitals': 'v'}, 'new section': {'new key': 'New Value'}}
testFile, SKIgnorecase: uckey, result: True
ini2, sections, keys lowercase:{'section': {'k': 'v'}, 'new section': {'new key': 'New Value'}}
testFile, SKIgnorecase: ucsection, result: True


spaces:

>>> L = ['spacing']
>>> for n in L:
...     ini = IniVars(root/(n+'.ini'))
...     readExplicit(ini)
...     ini.write(root/'test'/(n+'.ini'))
...     ini2 = IniVars(root/'test'/(n+'.ini'))
...     readExplicit(ini2)
...     print 'ini2: look at the spacing of: %s'% n
...     strini2 = str(ini2)
...     strini2 = strini2.replace('\\\\n', '|')
...     print 'spacing (|): %s'% strini2
...     print 'result (identical?): %s'% (ini == ini2,)
ini2: look at the spacing of: spacing
spacing (|): {'spacing section': {'one line': 'value', 'multiple lines': 'first line|second line||third line is new paragaph|fourth line has no leading spaces'}}
result (identical?): True

'''

#
#__test__ = {'regular expressions': retest,
#            'good files': goodfilesTest}

def test():
    #pylint:disable=C0415
    import doctest
    return doctest.testmod()


if __name__ == "__main__":
    test()

