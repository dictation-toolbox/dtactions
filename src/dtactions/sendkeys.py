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

(Quintijn Hoogenboom, 2021-04-04)
"""

from dtactions.vocola_sendkeys import ext_keys

def sendkeys(keys):
    """sends keystrokes via the vocola Keys extension
   
Keystrokes following Unimacro/Vocola convention are translated into Dragonfly notation:

:code:`"{shift+right 4}"`

Tested at bottom of this file interactively...
    """
    ext_keys.send_input(keys)
    
if __name__ == "__main__":
    sendkeys("{a 3}") #aaa
    sendkeys("x y z ")
    sendkeys("test, test, met komma.{home}")
    sendkeys("{ctrl+extend}{up 2}test, test,{ctrl+home}{ctrl+end}{up 2} nogmaals.")
    # ##
    ##
"""
aaax nogmaals.test, test, y z test, test, met komma.

""" 