# dtactions

dtactions is an OpenSource extension module for the speech recognition program Dragon.
It is meant to perform actions that are common to other packages like Dragonfly, Unimacro and Vocola.

This document describes how to instlall dtactions for end users and for developers.

## Status

dtactions code has been updated from Python 2 to Python 3. It is experimental at this moment.

The packages are ccurrently published in the [Test Python Packaging Index](https://test.pypi.org/) rather than
the [Python Packaging Index](https://pypi.org/). The pip commands are a bit more complicated for this.

## Instructions for End Users

If you would like to install dtactions for use, but not as a developer, here are the instructions:

Install Python and Natlink and the packages you would like to use (Dragonfly, Caster, Unimacro, Vocola) as described in the Natlink repository README.
   

1. Install dtactions
   It will also pull any prequisites from the [Python Packaging Index](https://pypi.org/).

   - `py -m pip install  dtactions`

   This will install the packages in your Python site-packages area. It will also add the following commands, which should be
   in your path now in your commmand prompt:


## Instructions for Developers

Your local git repository can be anywhere conveninent. 

Uninstall the packages you wish to develop. i.e pip if you want to work on dtactions:
  `py -m pip uninstall dtactions` and answer yes to all the questions about removing files from your python scripts folder.

Run `py -m pip install -e .`  from the dtactions project root.  


### Unit testing
Run pytest to run the tests, written in a combinatin of [unittest](https://docs.python.org/3/library/unittest.html) 
and [pytest](https://docs.pytest.org/).  IF adding a test, pytest seems to be a lot more convenient and powerful.

Most tests go in test;  tests that require a natlink install go in natlink_test as not every package dependent on natlink.  

You can run `py -m pip install dtactions[test]` or `py -m pip install dtactions[natlink_test]` if you don't have the prequisites like pytest.  

You can run pytest from project root folder to run the tests that don't depend on natlink being installed.  For the natlink-dependent tests, run 
`py -m pytest natlink_test`.  

## Notes About Packaging for Developers

The package is specified in `pyproject.toml` and built with [flit](https://pypi.org/project/flit/). The build_package command
(a batch file in the root folder of dtactions) builds a source distribution.

`py -m flit build` builds the package.  `py -m flit publish` publishes to [Python Packaging Index](https://pypi.org/).


 
Version numbers of the packages must be increased before your publish to [Test Python Packaging Index](https://test.pypi.org/)
or . These are specified in **init**.py in `src/dtactions`. Don't bother changing the
version numbers unless you are publishing.

If you are going to publish to a package index, you will need a .pypirc in your home directory. If you don't have one,
it is suggested you start with pypirc_template as the file format is rather finicky.
