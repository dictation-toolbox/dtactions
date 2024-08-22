# dtactions

dtactions is an OpenSource extension module for the speech recognition program Dragon.
It is meant to perform actions that are common to other packages like Dragonfly, Unimacro and Vocola.

This document describes how to install dtactions for end users and for developers.

## Status
Sucessfully upgraded to Python 3.

## Instructions for End Users

If you would like to install dtactions for use, but not as a developer, here are the instructions:

Install Python and Natlink and the packages you would like to use (Dragonfly, Caster, Unimacro, Vocola) as described in the Natlink repository README.
   

1. Install dtactions
   
   It will also pull any prerequisites from the [Python Packaging Index](https://pypi.org/).

   In a `cmd` or `powershell` window (possibly in elevated mode) run:
   - `py -m pip install dtactions`
   

## Instructions for Developers

If you are working on dtactions the most convenient setup is an [editable install](https://peps.python.org/pep-0660/).  Your local git repository can be anywhere convenient. 

Uninstall the packages you wish to develop. i.e pip if you want to work on `dtactions`:
  `py -m pip uninstall dtactions` and answer yes to all the questions about removing files from your python scripts folder.

Run `py -m pip install -e .[dev,test]`  from the dtactions project root.  You can skip `[dev,test]` after the first install.


### Unit testing
Run pytest to run the tests, written in a combinatin of [unittest](https://docs.python.org/3/library/unittest.html) 
and [pytest](https://docs.pytest.org/).  If adding a test, pytest seems to be a lot more convenient and powerful.

Most tests go in `tests`;  tests that require a natlink install go in `natlink_test` as not every package dependent on natlink.  

You can run `py -m pip install dtactions[test]` or `py -m pip install dtactions[natlink_test]` if you don't have the prequisites like pytest.  

You can run pytest from project root folder to run the tests that don't depend on natlink being installed.  For the natlink-dependent tests, run 
`py -m pytest natlink_test`.  

## Notes About Packaging for Developers

The package is specified in `pyproject.toml`. 

Too build the package locally, 

`python -m build`.  If this fails because you don't have build installed, run `pip install -e .[dev]`.    

To publish to [dtactions in the Python Packaging Index](https://pypi.org/project/dtactions/), increase the version number in `__init__.py` and [draft a new release](https://github.com/dictation-toolbox/dtactions/releases).


