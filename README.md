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

   - `pip install --no-cache --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple dtactions`

   This will install the packages in your Python site-packages area. It will also add the following commands, which should be
   in your path now in your commmand prompt:

   - natlinkconfigfunctions
   - natlinkstatus
   - startnatlinkconfig

## Instructions for Developers

Your local git repository can be anywhere conveninent. It no longer needs to be in a specific location relative to other
[dictation-toolbox](https://github.com/dictation-toolbox) packages.

- Install as per the instructions for end users, to get any python prequisites in.
- Install [flit](https://pypi.org/project/flit/) `pip install flit`. This is a python package build tool that is required for developer workflow.
- Uninstall the packages you wish to develop. i.e pip if you want to work on dtactions:
  `pip uninstall dtactions` and answer yes to all the questions about removing files from your python scripts folder.
- Build the Python packages. In the root folder of your dtactions repository, run `build_package` in your shell. This creates the package.  
  At this step, if you have any untracked files
  in your git repository, you will have to correct them with a `git add` command or adding those files to .gitignore.
- The cool part: `flit install --symlink'. This will install dtactions into site-packages by symolically linking
  site-packages/dtactions to the src/dtactions folder of your git repository. You can edit the files in site-packages/dtactions or
  in your git repository area as you prefer - they are the same files, not copies.

Oddly, when you follow this workflow and register dtactions by running startnatlinkcofig or natlinkconfigfunctions, even though the
python paths those commands pickup, you will find that the natlinkcorepath will be in our git repository.

## Notes About Packaging for Developers

The package is specified in `pyproject.toml` and built with [flit](https://pypi.org/project/flit/). The build_package command
(a batch file in the root folder of dtactions) builds a source distribution.

Several scripts are specfied in pyproject.toml in the scripts section. Scripts are automatically generated
and placed in the python distribution "Scripts" folder. Those scripts are then available in the system path for
users to run. Note the `flit install --symlink` will install scripts as batchfiles; `pip install dtactions` will install
scripts as .exe files.

Version numbers of the packages must be increased before your publish to [Test Python Packaging Index](https://test.pypi.org/)
or [Python Packaging Index](https://pypi.org/). These are specified in **init**.py in `src/dtactions`. Don't bother changing the
version numbers unless you are publishing.

This command will publish to [Test Python Packaging Index](https://test.pypi.org/): `publish_package_testpypi`.
This will publish to [Python Packaging Index](https://pypi.org/): `publish_package_pypy`.

If you are going to publish to a package index, you will need a .pypirc in your home directory. If you don't have one,
it is suggested you start with pypirc_template as the file format is rather finicky.
