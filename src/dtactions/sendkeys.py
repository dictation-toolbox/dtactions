### this file calls the vocola Keys extension, http://vocola.net/unofficial/keys.html
### this software is released under the MIT license.
"""Sending keystrokes to the foreground window.

  Usage:

At top your module insert:

:code:`from dtactions.sendkeys import sendkeys`

and then in a function:
    
    :code:`sendkeys("keystrokes")`

This module now adopts the Dragonfly :code:`action_key` module,
so `"{alt+w}"` is (in the function) converted to `"a-w"` etc.

Optionally, you can also use sendsystemkeys, which is implmented via Dragon SendSystemKeys (via natlink.execScript)

(Quintijn Hoogenboom, 2021-04-04)
"""

import natlink
from dtactions.vocola_sendkeys import ext_keys

def sendkeys(keys):
    """sends keystrokes via the vocola Keys extension 
   
:code:`"{shift+right 4}"`

Tested at bottom of this file interactively...
    """
    ext_keys.send_input(keys)
    
def sendsystemkeys(keys):
    """sends keystrokes "the hard way" via Dragon's SendSystemKeys
   
Tested at bottom of this file interactively...
    """
    if not keys:
        return
    natlink.execScript(f'SendSystemKeys("{keys}")')
    
    
    
if __name__ == "__main__":
    # sendkeys("{a 3}") #aaa
    # sendkeys("x y z ")
    # sendkeys("test, test, met komma.{home}")
    sendkeys("{ctrl+end}{up 2}{home}{shift+end}{del}this is wrong{shift+left 5}right")
    # ##
"""

"""
