# dtactions

Common OS action and related code from dictation-toolbox projects

## More Information
 Please refer to the README file in the project repository [https://github.com/dictation-toolbox/dtactions](https://github.com/dictation-toolbox/dtactions) or [https://dtactions/readthedocs.io/](https://dtactions/readthedocs.io/)

## Installing from PyPi
You can install from [The Python Package Index (PyPI)](https://pypi.org/) with 

`pip install dtactions`

## More information
See the readme.md of the repository 'natlinkcore'

## Building the Python Package Locally (flit)

The build happens through a powershell script.  You don't have to know much powershell.  

The package is built with [Flit](https://flit.pypa.io/).  The package will be produced in
dist/dtactions-x.y.z-py3-none-any.whl.  To install it `pip install dist/dtactions-x.y.z-py3-none-any.whl` replacing x.y with the version numbers.

Normally if you are developing dtactions, you will with instead to install with `pip install -e .`, which will
let you make and test changes without reinstalling dtactions with pip.

**Note the flit install --symlink or --pth-file options can be problematic so just use pip.**

 
To start a powershell from the command prompt, type `powershell`.

To build the package:


`flit build`   from powershell or command prompt, which will run the the tests in dtactions/test, then build the the package.


When you have multiple versions of python on your computer, you may need the full path to flit, eg

`\C:\Python310-32\Scripts\flit build`

To publish the package to [The Python Package Index (PyPI)](https://pypi.org/)

`\C:\Python310-32\Scripts\flit publish`

## Publishing checklist
Before you bump the version number in __init__.py and publish:
- Check the pyroject.toml file for package dependancies.  Do you need a specfic or newer version of some package?
Then add or update the version # requirement in dtactions.  
- don't publish if the tests are failing. But regrettably, currently the testing is not working. 
