"""utility functions from Quintijn, used in unimacro and in local programs.
   in python3 also the path module as subclass of the standard path class
"""
#pylint:disable=C0116, C0302,  W0613, R0912, R0914, R0915, R0911, W0702
import unicodedata
import os
import re
import collections
from pathlib import Path
    
def unifyaccentedchars(to_translate):
    """change acuted characters with combining code to single characters
    
The combining variant (s3) is converted into the one character shorter variant s1, apart from the capitals.
This is done by the NFC variant of the normalize function
    
>>> s1 = "cafcomb\u00e9"   # single char e acute no change
>>> s2 = unifyaccentedchars(s1)
>>> s2
'cafcomb\u00e9'

## this one changes to s1:
>>> s3 = "CafCombe\u0301"   # combining char e acute   0301 is the combining code
>>> s4 = unifyaccentedchars(s3)   # combining char e acute
>>> s4
'CafComb\u00e9'
>>> len(s1), len(s2), len(s3), len(s4)
(8, 8, 9, 8)
>>> s1 == s3.lower()
False
>>> s2 == s4.lower()
True

    (from Fluent Python)
    """
    norm_txt = unicodedata.normalize('NFC', to_translate)
    return norm_txt

def normalizeaccentedchars(to_translate):
    """change acutechars to ascii 
    
>>> s1 = "cafnon\u00e9"   # single char e acute
>>> normalizeaccentedchars(s1)  
'cafnone'
>>> s2 = "CafCombe\u0301"   # combining char e acute   0301 is the combining code
>>> normalizeaccentedchars(s2)   # combining char e acute
'CafCombe'
>>> len(s1), len(s2)
(7, 9)

    (from Fluent Python)
    """
    norm_txt = unicodedata.normalize('NFD', to_translate)
    shaved = ''.join(c for c in norm_txt if not unicodedata.combining(c))
    return shaved

def doubleaccentedchars(to_translate):
    """change acutechars to ascii, but double them e acute = ee
    (from Fluent Python, adaptation QH)
    

>>> s1 = "double caf\\u00e9"   # single char e acute
>>> doubleaccentedchars(s1)
'double cafee'
>>> s2 = "Double Cafe\\u0301"   # combining char e acute
>>> doubleaccentedchars(s2)
'Double Cafee'

>>> doubleaccentedchars("enqu\N{LATIN SMALL LETTER E}\N{COMBINING CIRCUMFLEX ACCENT}te")
'enqueete'

    """
    norm_txt = unicodedata.normalize('NFD', to_translate) ## haal char en accent uit elkaar
    shaved = []
    last = ""
    for c in norm_txt:
        comb = unicodedata.combining(c)
        if comb:
            if comb == 230:
                # accent aegu, accent grave, accent circonflex, decide in favour of accent aegu, double char
                # print('doubleaccentedchars, combining value %s, double char: %s (%s)'% (comb, last, to_translate))
                shaved.append(last)
            elif comb == 202:
                if last == "c":
                    # print('c cedilla, change to "s" (%s)'% to_translate)
                    shaved.pop()
                    shaved.append("s")
                else:
                    # print("c cedilla, but NO C, ignore (%s)"% to_translate)
                    pass
            else:
                print(f'(yet) unknown combining char {comb} in "{to_translate}", ignore')
            last = ""
        else:
            shaved.append(c)
            last = c
    return ''.join(shaved)

class peek_ahead:
    """ An iterator that supports a peek operation.
    
    Improved for python3 after QH's example: Adapted for python 3 by Paulo Roma.
    
    this is a merge of example 19.18 of python cookbook part 2, peek ahead more steps
    and the simpler example 16.7, which peeks ahead one step and stores it in
    the self.preview variable.
    
    Adapted so the peek function never raises an error, but gives the
    self.sentinel value in order to identify the exhaustion of the iter object.
    
    Example usage (Paulo):
    
    >>> p = peek_ahead(range(4))
    >>> p.peek()
    0
    >>> p.next(1)
    [0]
    >>> p.isFirst()
    True
    >>> p.preview
    1
    >>> p.isFirst()
    True
    >>> p.peek(3)
    [1, 2, 3]
    >>> p.next(2)
    [1, 2]
    >>> p.peek(2) #doctest: +ELLIPSIS
    [3, <object object at ...>]
    >>> p.peek(1)
    [3]
    >>> p.next(2)
    Traceback (most recent call last):
    StopIteration
    >>> p.next()
    3
    >>> p.isLast()
    True
    >>> p.next()
    Traceback (most recent call last):
    StopIteration
    >>> p.next(0)
    []
    >>> p.peek()  #doctest: +ELLIPSIS
    <object object at ...>
    >>> p.preview #doctest: +ELLIPSIS
    <object object at ...>
    >>> p.isLast()  # after the iter process p.isLast remains True
    True
    
    ### my old unittests, QH:
    
        >>> p = peek_ahead(range(4))
    >>> p.peek()
    0
    >>> p.next(1)
    [0]
    >>> p.isFirst()
    True
    >>> p.preview
    1
    >>> p.isFirst()
    True
    >>> p.peek(3)
    [1, 2, 3]
    >>> p.next(2)
    [1, 2]
    >>> p.peek(2) #doctest: +ELLIPSIS
    [3, <object object at ...>]
    >>> p.peek(4) #doctest: +ELLIPSIS
    [3, <object object at ...>, <object object at ...>, <object object at ...>]
    >>> p.peek(1)
    [3]
    >>> p.next(2)
    Traceback (most recent call last):
    StopIteration
    >>> p.next()
    3
    >>> p.isLast()
    True
    >>> p.next()
    Traceback (most recent call last):
    StopIteration
    >>> p.next(0)
    []
    >>> p.peek()  #doctest: +ELLIPSIS
    <object object at ...>
    >>> p.preview #doctest: +ELLIPSIS
    <object object at ...>
    >>> p.isLast()  # after the iter process p.isLast remains True
    True

    From example 16.7 from python cookbook 2.

    The preview can be inspected through it.preview

    ignoring duplicates:
    >>> it = peek_ahead('122345567')
    >>> for i in it:
    ...     if it.preview == i:
    ...         continue
    ...     print(i, end=" ")
    1 2 3 4 5 6 7 

    getting duplicates together:
    >>> it = peek_ahead('abbcdddde')
    >>> for i in it:
    ...     if it.preview == i:
    ...         dup = 1
    ...         while 1:
    ...             i = it.next()
    ...             dup += 1
    ...             if i != it.preview:
    ...                 print(i*dup, end=" ")
    ...                 break
    ...     else:
    ...         print(i, end=" ")
    ...
    a bb c dddd e 
    
    """
    ## schildwacht (guard)
    sentinel = object()
    def __init__(self, iterable):
        ## iterator
        self._iterable = iter(iterable)
        try:
           ## next method hold for speed
            self._nit = self._iterable.next
        except AttributeError:
            self._nit = self._iterable.__next__
        ## deque object initialized left-to-right (using append())
        self._cache = collections.deque()
        ## initialize the first preview already
        self._fillcache(1)
        ## peek at leftmost item
        self.preview = self._cache[0]
        ## keeping the count allows checking isFirst and isLast status
        self.count = -1

    def __iter__(self):
        """return an iterator
        """
        return self

    def _fillcache(self, n):
        """fill _cache of items to come, with one extra for the preview variable
        """
        if n is None:
            n = 1
        while len(self._cache) < n+1:
            try:
                Next = self._nit()
            except StopIteration:
                # store sentinel, to identify end of iter:
                Next = self.sentinel
            self._cache.append(Next)

    def __next__(self, n=None):
        """gives next item of the iter, or a list of n items
        
        raises StopIteration if the iter is exhausted (self.sentinel is found),
        but in case of n > 1 keeps the iter alive for a smaller "next" calls
        """
        self._fillcache(n)
        if n is None:
            result = self._cache.popleft()
            if result == self.sentinel:
                # find sentinel, so end of iter:
                self.preview = self._cache[0]
                raise StopIteration
            self.count += 1
        else:
            result = [self._cache.popleft() for i in range(n)]
            if result and result[-1] == self.sentinel:
                # recache for future use:
                self._cache.clear()
                self._cache.extend(result)
                self.preview = self._cache[0]
                raise StopIteration
            self.count += n
        self.preview = self._cache[0]
        return result

    def next(self,n=None):
        """python2 compatibility
        """
        #pylint:disable=C2801
        return self.__next__(n)
    
    def isFirst(self):
        """returns true if iter is at first position
        """
        return self.count == 0

    def isLast(self):
        """returns true if iter is at last position or after StopIteration
        """
        return self.preview == self.sentinel

    def hasNext(self):
        """returns true if iter is not at last position
        """
        return not self.isLast()
        
    def peek(self, n=None):
        """gives next item, without exhausting the iter, or a list of 0 or more next items
        
        with n == None, you can also use the self.preview variable, which is the first item
        to come.
        """
        self._fillcache(n)
        if n is None:
            result = self._cache[0]
        else:
            result = [self._cache[i] for i in range(n)]
        return result
    # old name:
    # peekable = peek_ahead

## TODOQH
class peek_ahead_stripped(peek_ahead):
    """ Iterator that strips lines of text, and returns (leftSpaces,strippedLine)

    sentinel is just False, such that peeking ahead can check for truth input

    >>> lines = ['line1', '', ' one space ahead','', '   three spaces ahead, 1 empty line before']
    >>> list(peek_ahead_stripped(lines))
    [(0, 'line1'), (0, ''), (1, 'one space ahead'), (0, ''), (3, 'three spaces ahead, 1 empty line before')]

    example of testing look ahead

    >>> lines = ['line1 ', '', 'line2 (last)']
    >>> it = peek_ahead_stripped(lines)
    >>> for spaces, text in it:
    ...     print('current line: |', text, '|', end=' ')
    ...     if it.preview is it.sentinel:
    ...         print(', cannot preview, end of peek_ahead_stripped')
    ...     elif it.preview[1]:
    ...         print(', non empty preview: |', it.preview[1], '|')
    ...     else:
    ...         print(', empty preview')
    current line: | line1 | , empty preview
    current line: |  | , non empty preview: | line2 (last) |
    current line: | line2 (last) | , cannot preview, end of peek_ahead_stripped

    """
    def _fillcache(self, n):
        """fill _cache of items to come, special treatment for this stripped subclass
        """
        if n is None:
            n = 1
        while len(self._cache) < n+1:
            try:
                line = self._nit()
                line = line.rstrip()
                Next = (len(line) - len(line.lstrip()), line.lstrip())
            except StopIteration:
                # store sentinel, to identify end of iter:
                Next = self.sentinel
            self._cache.append(Next)
     
# helper string functions:
def replaceExt(fileName, ext):
    """change extension of file

>>> replaceExt("a.psd", ".jpg")
'a.jpg'
>>> replaceExt("a/b/c/d.psd", "jpg")
'a/b/c/d.jpg'
    """
    ext = addToStart(ext, ".")
    fileName = str(fileName)
    a, _extOld = os.path.splitext(fileName)
    return a + ext

def getExt(fileName):
    """return the extension of a file

>>> getExt(u"a.psd")
'.psd'
>>> getExt("a/b/c/d.psd")
'.psd'
>>> getExt("abcd")
''
>>> getExt("a/b/xyz")
''
    """
    _a, ext = os.path.splitext(fileName)
    return ext

def removeFromStart(text, toRemove, ignoreCase=None):
    """returns the text with "toRemove" stripped from the start if it matches
>>> removeFromStart('abcd', 'a')
'bcd'
>>> removeFromStart('abcd', 'not')
'abcd'

working of ignoreCase:

>>> removeFromStart('ABCD', 'a')
'ABCD'
>>> removeFromStart('ABCD', 'ab', ignoreCase=1)
'CD'
>>> removeFromStart('abcd', 'ABC', ignoreCase=1)
'd'

    """
    if ignoreCase:
        text2 = text.lower()
        toRemove = toRemove.lower()
    else:
        text2 = text
    if text2.startswith(toRemove):
        return text[len(toRemove):]
    return text


def removeFromEnd(text, toRemove, ignoreCase=None):
    """returns the text with "toRemove" stripped from the end if it matches

>>> removeFromEnd('a.jpg', '.jpg')
'a'
>>> removeFromEnd('b.jpg', '.gif')
'b.jpg'

working of ignoreCase:

>>> removeFromEnd('C.JPG', '.jpg')
'C.JPG'
>>> removeFromEnd('D.JPG', '.jpg', ignoreCase=1)
'D'
>>> removeFromEnd('d.jpg', '.JPG', ignoreCase=1)
'd'

    """
    if ignoreCase:
        text2 = text.lower()
        toRemove = toRemove.lower()
    else:
        text2 = text
    if text2.endswith(toRemove):
        return text[:-len(toRemove)]
    return text

def addToStart(text, toAdd, ignoreCase=None):
    """returns text with "toAdd" added at the start if it was not already there

    if ignoreCase:
        return the start of the string with the case as in "toAdd"

>>> addToStart('a-text', 'a-')
'a-text'
>>> addToStart('text', 'b-')
'b-text'
>>> addToStart('B-text', 'b-')
'b-B-text'

working of ignoreCase:

>>> addToStart('C-Text', 'c-', ignoreCase=1)
'c-Text'
>>> addToStart('d-Text', 'D-', ignoreCase=1)
'D-Text'

    """
    if ignoreCase:
        text2 = text.lower()
        toAdd2 = toAdd.lower()
    else:
        text2 = text
        toAdd2 = toAdd
    if text2.startswith(toAdd2):
        return toAdd + text[len(toAdd):]
    return toAdd + text

def addToEnd(text, toAdd, ignoreCase=None):
    """returns text with "toAdd" added at the end if it was not already there

    if ignoreCase:
        return the end of the string with the case as in "toAdd"

>>> addToEnd('a.jpg', '.jpg')
'a.jpg'
>>> addToEnd('b', '.jpg')
'b.jpg'

working of ignoreCase:

>>> addToEnd('Cd.JPG', '.jpg', ignoreCase=1)
'Cd.jpg'
>>> addToEnd('Ef.jpg', '.JPG', ignoreCase=1)
'Ef.JPG'

    """
    if ignoreCase:
        text2 = text.lower()
        toAdd2 = toAdd.lower()
    else:
        text2 = text
        toAdd2 = toAdd
    if text2.endswith(toAdd2):
        return text[:-len(toAdd)] + toAdd
    return text + toAdd

def firstLetterCapitalize(t):
    """capitalize only the first letter of the string
    """
    if t:
        return t[0].upper() + t[1:]
    return ""

def extToLower(fileName):
    """leave name part intact, but change extension to lowercase
>>> extToLower("aBc.jpg")
'aBc.jpg'
>>> extToLower("ABC.JPG")
'ABC.jpg'
>>> extToLower("D:/a/B/ABC.JPG")
'D:/a/B/ABC.jpg'



    """
    f, ext = os.path.splitext(fileName)
    return f + ext.lower()


def appendBeforeExt(text, toAppend):
    """append text just before the extension of the filename part

>>> appendBeforeExt("short.html", '__t')
'short__t.html'
>>> appendBeforeExt("http://a/b/c/d/long.html", '__b')
'http://a/b/c/d/long__b.html'
    """
    base, ext = os.path.splitext(text)
    return base + toAppend + ext

Classes = ('ExploreWClass', 'CabinetWClass')


## to pathqh:
# # def toUnixName(t, glueChar="", lowercase=1, canHaveExtension=1, canHaveFolders=1, mayBeEmpty=False):
    
def sendkeys_escape(Str):
    """escape with {} keys that have a special meaning in sendkeys
    + ^ % ~ { } [ ]

>>> sendkeys_escape('abcd')
'abcd'
>>> sendkeys_escape('+bcd')
'{+}bcd'
>>> sendkeys_escape('+a^b%c~d{f}g[h]i')
'{+}a{^}b{%}c{~}d{{}f{}}g{[}h{]}i'
>>> sendkeys_escape('+^%~{}[]')
'{+}{^}{%}{~}{{}{}}{[}{]}'

    """
    ## make str, for python 2 for the time being:
    return ''.join(map(_sendkeys_escape, Str))

def _sendkeys_escape(s):
    """Escape one character in the set or return if different"""
    if s in ('+', '^', '%', '~', '{' , '}' , '[' , ']' ) :
        result = '{%s}' % s
    else:
        result = s
    return str(result)

def getSublists(L, maxLen, sepLen):
    """generator function, that gives pieces of the list, up to
    the maximum length, accounting for the separator length
    """
    if not L:
        yield L
        return
    listPart = [L[0]]
    lenPart = len(L[0])
    for w in L[1:]:
        lw = len(w)
        if lw + lenPart > maxLen:
            yield listPart
            listPart = [w]
            lenPart = lw
        else:
            lenPart += lw + sepLen
            listPart.append(w)
    yield listPart

def getWordsUntilLength(t, maxLength):
    """take words until maxLength is reached
>>> getWordsUntilLength('this is a test', 60)
'this is a test'
>>> getWordsUntilLength('this is a test', 7)
'this is'
>>> getWordsUntilLength('this is a test', 2)
'this'
    
    
    """
    t = t.replace(',', '')
    t = t.replace('.', '')
    T = t.split()
    while T:
        t = ' '.join(T)
        if len(t) <= maxLength:
            return t
        T.pop()
    return t

def splitLongString(S, maxLen=70, prefix='', prefixOnlyFirstLine=0):
    """Splits a (long) string into newline separated parts,

    a list of strings is returned.

    possibly with a fixed prefix, or a prefix for the first line only.
    Possibly items inside the line are separated by a given separator

    maxLen = maximum line length, can be exceeded is a very long word is there
    prefix = text that is inserted in front of each line, default ''
    prefixOnlyFirstLine = 1: following lines as blank prefix, default 0
    >>> splitLongString('foo', 80)
    ['foo']
    >>> splitLongString(' foo   bar and another set of  words  ', 80)
    ['foo bar and another set of words']
    >>> splitLongString(' foo   bar and another set of  words  ', 20,
    ... prefix='    # ')
    ['    # foo bar and', '    # another set of', '    # words']
    >>> splitLongString(' foo   bar and another set of  words  ', 20,
    ... prefix='entry = ', prefixOnlyFirstLine=1)
    ['entry = foo bar and', '        another set', '        of words']
    """
    assert isinstance(S, str)
    L = [t.strip() for t in S.split()]
    lOut = []
    for part in getSublists(L, maxLen=maxLen-len(prefix), sepLen=1):
        lOut.append(prefix + ' '.join(part))
        if prefixOnlyFirstLine:
            prefix = ' '*len(prefix)
    return lOut



def cleanString(s):
    """converts a string with leading and trailing and
    intermittent whitespace into a string that is stripped
    and has only single spaces between words
>>> cleanString('foo bar')
'foo bar'
>>> cleanString('foo  bar')
'foo bar'
>>> cleanString('\\n foo \\n\\n  bar ')
'foo bar'
>>> cleanString('')
''

    """
    return ' '.join([x.strip() for x in s.split()])

def formatListColumns(List, lineLen = 70, sort = 0):
    """formats a list in columns

    Uses a generator function "splitList", that gives a sequence of
    sub lists of length n.

    The items are separated by at least two spaces, if the list
    can be placed on one line, the list is comma separated

>>> formatListColumns([''])
''
>>> formatListColumns(['a','b'])
'a, b'
>>> formatListColumns(['foo', 'bar', 'longer entry'], lineLen=5)
'foo\\nbar\\nlonger entry'
>>> formatListColumns(['foo', 'bar', 'longer entry'], lineLen=5, sort=1)
'bar\\nfoo\\nlonger entry'
>>> print(formatListColumns(['afoo', 'bar', 'clonger', 'dmore', 'else', 'ftest'], lineLen=20, sort=1))
afoo     dmore
bar      else
clonger  ftest
>>> print(formatListColumns(['foo', 'bar', 'longer entry'], lineLen=20))
foo  longer entry
bar

    """
    if sort:
        List = sorted(List, key=str.casefold)
    s = ', '.join(List)

    # short list, simply join with comma space:
    if len(s) <= lineLen:
        return s

    maxLen = max(list(map(len, List)))

    # too long elements in list, return "\n" separated string:
    if maxLen > lineLen:
        return '\n'.join(List)


    nRow = len(s)//lineLen + 1
    lenList = len(List)

    # try for successive number of rows:
    while nRow < lenList//2 + 2:
        lines = []
        for i in range(nRow):
            lines.append([])
        maxLenTotal = 0
        for parts in splitList(List, nRow):
            maxLenParts = max(list(map(len, parts))) + 2
            maxLenTotal += maxLenParts
            for i, part in enumerate(parts):
                lines[i].append(part.ljust(maxLenParts))
        if maxLenTotal > lineLen:
            nRow += 1
        else:
            # return '\n'.join(map(string.strip, map(string.join, lines)))
            return '\n'.join([''.join(t).strip() for t in lines])
    # unexpected long list:
    return '\n'.join(List)

hasDoubleQuotes = re.compile(r'^".*"$')
hasSingleQuotes = re.compile(r"^'.*'$")

def convertToPythonArgs(text):
    """convert to numbers and strings

    IF argument is enclosed in " " or ' ' it is kept as a string.
    """    
    text = text.strip()
    if not text:
        return None
    L = text.split(',')
    L = list(map(_convertToPythonArg, L))
    return tuple(L)

def convertToPythonArgsKwargs(text):
    """convert to numbers and strings,

    IF argument is enclosed in " " or ' ' it is kept as a string.

    also do kwargs now...
>>> convertToPythonArgsKwargs('')
((), {})
>>> convertToPythonArgsKwargs('hello')
(('hello',), {})
>>> convertToPythonArgsKwargs('width=50')
((), {'width': 50})
>>> convertToPythonArgsKwargs('"hello", width=50')
(('hello',), {'width': 50})

    

    """
    L = []
    K = {}
    text = text.strip()
    if not text:
        return tuple(L), K 
    textList = text.split(',')
    for arg in textList:
        if arg.find("=") > 0:
            kw, arg2 = [t.strip() for t in arg.split("=", 1)]
            if kw.find(" ") == -1:
                arg2 = _convertToPythonArg(arg2)
                K[kw] = arg2
            else:
                L.append(arg.strip())
        else:
            L.append(_convertToPythonArg(arg))
    return tuple(L), K

def _convertToPythonArg(t):
    t = t.strip()
    if not t:
        return ''

    # if input string is a number, return string directly
    try:
        i = int(t)
        if t == '0':
            return 0
        if t.startswith('0'):
            print(f'warning convertToPythonArg, could be int, but assume string: {t}')
            return str(t)
        return i
    except ValueError:
        pass
    try:
        f = float(t)
        if t.find(".") >= 0:
            return f
        print(f'warning convertToPythonArg, can be float, but assume string: {t}')
        return str(t)
    except ValueError:
        pass

    # now proceeding with strings:    
    if hasDoubleQuotes.match(t):
        return t[1:-1]
    if hasSingleQuotes.match(t):
        return t[1:-1]
    return t

def splitList(L, n):
    """generator function that splits a list in sublists of length n

    """
    O = []
    for l in L:
        O.append(l)
        if len(O) == n:
            yield O
            O = []
    if O:
        yield O

def makeReadable(t):
    """squeeze text for readability
    
    helper for print lines...
    """
    t = t.strip()
    t = t.replace('\n', '\\\\')
    if len(t) > 100:
        return t[:50] + ' ... ' + t[-50:]
    return t

## functions for generating alternative paths in virtual drives
## uses reAltenativePaths, defined in the top of this module
## put in utilsqh.py! used in sitegen AND in _folders.py grammar of Unimacro:
# for alternatives in virtual drive definitions:
reAltenativePaths = re.compile(r"(\([^|()]+?(\|[^|()]+?)+\))")

def generate_alternatives(s):
    """generates altenatives if (xxx|yyy) is found, otherwise just yields s
    Helper for cross_loop_alternatives
    """
    m = reAltenativePaths.match(s)
    if m:
        alternatives = s[1:-1].split("|")
        for item in alternatives:
            yield item
    else:
        yield s
        
def cross_loop_alternatives(*sequences):
    """helper function for loop_through_alternative_paths
    """
    if sequences:
        for x in generate_alternatives(sequences[0]):
            for y in cross_loop_alternatives(*sequences[1:]):
                yield (x,) + y
    else:
        yield ()

def loop_through_alternative_paths(pathdefinition):
    """can hold alternatives (a|b)
    
>>> list(loop_through_alternative_paths("(C|D):/xxxx/yyyy"))
['C:/xxxx/yyyy', 'D:/xxxx/yyyy']
>>> list(loop_through_alternative_paths("(C:|D:|E:)/Do(kum|cum)ent(s|en)"))
['C:/Dokuments', 'C:/Dokumenten', 'C:/Documents', 'C:/Documenten', 'D:/Dokuments', 'D:/Dokumenten', 'D:/Documents', 'D:/Documenten', 'E:/Dokuments', 'E:/Dokumenten', 'E:/Documents', 'E:/Documenten']

    So "(C|D):/natlink" first yields "C:/natlink" and then "D:/natlink".
    More alternatives in one item are possible, see second example.
    """
    m = reAltenativePaths.search(pathdefinition)
    if m:
        result = reAltenativePaths.split(pathdefinition)
        result = [x for x in result if x and not x.startswith("|")]
        for pathdef in cross_loop_alternatives(*result):
            yield ''.join(pathdef)
    else:
        # no alternatives, simply yield the pathdefinition:
        yield pathdefinition

def getValidPath(variablePathDefinition):
    r"""check the different alternatives of the definition
    
# >>> homefolder = getValidPath(r'(A|B|D|C|D):\Users\(Dell|Gebruiker)\.Natlink')
# >>> homefolder
# WindowsPath('C:/Users/Gebruiker/.Natlink')
# >>> str(homefolder)
# 'C:\\Users\\Gebruiker\\.Natlink'

    return the first valid path, None if not found
    """
    for p in loop_through_alternative_paths(variablePathDefinition):
        if os.path.exists(p):
            return Path(p)
    return None

def _test():
    #pylint:disable=C0415
    import doctest
    return doctest.testmod()

if __name__ == "__main__":


    def revTrue(t):
        return "y"
    def revFalse(t):
        return "n"
    def revAbort(t):
        return
    _test()
