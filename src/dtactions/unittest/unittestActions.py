#
# Python Macro Language for Dragon NaturallySpeaking
#   (c) Copyright 1999 by Joel Gould
#   Portions (c) Copyright 1999 by Dragon Systems, Inc.
#
# unittestActions.py
#
# This module tests the actions module (the "unimacro actions")
# Quintijn Hoogenboom, 2021
#
import sys
import unittest
import os
import os.path
import time

try:
    from dtactions.__init__ import getThisDir, checkDirectory
except ModuleNotFoundError:
    print(f'Run this module after "build_package" and "flit install --symlink"\n')
    raise

thisDir = getThisDir(__file__)
dtactionsDir = os.path.normpath(os.path.join(thisDir, '..'))

# import TestCaseWithHelpers
import unittest
# import natlink
from dtactions import natlinkclipboard
from dtactions.unimacro import actions
# from dtactions.unimacro.actions import 

class TestError(Exception):
    pass

dataDirDtactions = os.path.expanduser("~\\.dtactions")
dataDir = os.path.join(dataDirDtactions, 'unimacro')
checkDirectory(dataDir)

logFileName = os.path.join(dataDir, "testresult.txt")  #
print(f'output will be logged in {logFileName}')
print('start UnittestActions', file=open(logFileName, 'w'))

class UnittestActions(unittest.TestCase):
    def setUp(self):
        pass        
    def tearDown(self):
        pass

    def testSimple(self):
        """only testing an empty action
        
        """
        testName = "testSimple"
        print(f'do {testName}')
        act = actions.Action()
        result = act.doAction('T')
        print(f'result action T: {result}')
        result = act.doAction('F')
        print(f'result action F: {result}')
        result = act.doKeystroke("Hello world")
        print(f'result doKeystroke: {result}')
                
    def log(self, t):
        print(t, file=open(logFileName, "a"))

def run():
    print('starting UnittestActions') 
    sys.stdout = open(logFileName, 'a')
    sys.stderr = open(logFileName, 'a')
    
    suite = unittest.makeSuite(UnittestActions, 'test')
    result = unittest.TextTestRunner().run(suite)
    
if __name__ == "__main__":
    print(f'run the tests, result will be in {logFileName}')
    run()
