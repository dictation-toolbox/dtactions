"""actions from application pythonwin
"""
#pylint:disable=R0903
import dtactions.messagefunctions as mf
from .actionbases import MessageActions
class PythonwinActions(MessageActions):
    handlesDict = {}
    def __init__(self, progInfo):
        MessageActions.__init__(self, progInfo)
        
    def getInnerHandle(self, topHndle):
        
        controls = mf.findControls(topHndle, wantedClass="Scintilla")
        print(f'Scintilla controls: {controls}')
        for c in controls:
            ln = self.getCurrentLineNumber(c)
            numberLines = self.getNumberOfLines(c)
            visible1 = mf.isVisible(c)
            print(f'control: {c}, linenumber: {ln}, nl: {numberLines}, visible: {visible1}')
                
if __name__ == '__main__':
    pass
