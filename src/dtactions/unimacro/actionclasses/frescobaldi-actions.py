"""actions from application frescobaldi, the editor for lilypond music note printing

now works with copy and so getting the wanted contents of text around the cursor.

"""
from dtactions.unimacro.actionclasses.actionbases import AllActions
import pprint
from natlinkcore import natlinkutils
from dtactions.unimacro import unimacroutils as natqh
import time

class FrescobaldiActions(AllActions):
    def __init__(self, progInfo):
        AllActions.__init__(self, progInfo)
        
    def hasSelection(self):
        """returns the text if text is selected, otherwise None
        """
        unimacroutils.saveClipboard()
        self.playString("{ctrl+c}")
        t = unimacroutils.getClipboard() or None
        unimacroutils.restoreClipboard()
        return t

    def playString(self, t):
        """send through to natlinkutils.self.playString
        """
        natlinkutils.playString(t)

    def getClipboard(self, sleep=0.05):
        """wait little longer if no result
        """
        for i in range(5):
            time.sleep(sleep)
            clip = unimacroutils.getClipboard()
            if clip:
                return clip
        print('got nothing on clipboard')


    def getPrevNext(self, n=1, sleep=0.05):
        """return character to the left and to the right of the cursor
        assume no selection active.
        normally return cursor in same position
        """
        unimacroutils.saveClipboard()
        self.playString("{left %s}"% n)
        self.playString("{shift+right %s}"% (n*2,))
        time.sleep(sleep)
        self.playString("{ctrl+c}")
        result = self.getClipboard(sleep=sleep)
        self.playString("{left}{right %s}"% n)
        time.sleep(sleep)
        unimacroutils.restoreClipboard()
        if not result:
            print('nothing on clipboard')
            return
       
        if len(result) == n*2:
            return result[:n], result[n:]
        elif result == '\n':
            print('getPrevNext, assume at end of file...')
            # assume at end of file, could also be begin of file, but too rare too handle
            self.playString("{right}")
            return result, result
        else:
            print('getPrevNext, len not 2: %s, (%s)'% (len(result), repr(result)))
            return "", result
        
        
    def getNextLine(self, n=1):
        """get line following the cursor

        n > 1, take more lines down for a larger range
        """
        self.playString("{shift+down %s}{ctrl+c}"% n)
        unimacroutils.Wait()        
        result = unimacroutils.getClipboard()
        nup = result.count('\n')
        if nup:
            self.playString("{shift+up %s}"% nup)
        return result
        
    def getPrevLine(self, n=1):
        """get line preceding the cursor
        
        more lines possible, if n > 1
        """
        self.playString("{shift+up %s}{ctrl+c}"% n)
        unimacroutils.Wait()        
        result = unimacroutils.getClipboard()
        ndown = result.count('\n')
        if ndown:
            self.playString("{shift+down %s}"% ndown)
        return result
    
    def getWindowText(self):
        pass
    
        
                   
    
if __name__ == '__main__':
    # search all frescobaldi instances and dump the controls:
    # does not give a useful result
    from dtactions import messagefunctions as mf
    # trying to hack into frescobaldi, no luck (yet)
    tw = mf.findTopWindows(wantedText="frescobaldi")
    for t in tw:
        dw = mf.dumpWindow(t)
        if dw:
            print('top: %s, length dump: %s'% (t, len(dw)))
            for subwindow in dw:
                print('sub: %s'% subwindow)
                hndle = subwindow[0]
                subhndle = subwindow[-1][0]
                try:
                    numl = mf.getNumberOfLines(hndle, classname='Scintilla')
                    print('hndle: %s, subhndle: %s, numlines: %s'% (hndle, subhndle, numl))
                except TypeError:
                    print('wrong hndle: %s'% subhndle)
        else:
            print('try topwindow')
            hndle = t
            try:
                numl = mf.getNumberOfLines(hndle, classname='Scintilla')
                print('top hndle: %s, numlines: %s'% (hndle, numl))
            except TypeError:
                print('wrong hndle: %s'% subhndle)
            
            print('top: %s, mp dumpwindow'% t)
        # pprint.pprint( mf.dumpWindow(t) )
    ## get correct window handle with "give window info" (_general grammar)
    progInfo = ('frescobaldi', 'naamloos', 'top', 396616)
    time.sleep(3)
    fa = FrescobaldiActions( progInfo )
    fa.getPrevNext()
    for j in range(1,11):
        for i in range(10):
            result = fa.getPrevNext(j)
            if result:
                print(result, len(result))
            else:
                print('no result')
            fa.playString("{left 3}")
        