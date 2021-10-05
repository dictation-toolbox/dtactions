# This file is part of Dragonfly and Unimacro
# (c) Copyright 2007, 2008 by Christo Butcher
# (c) Copyright 2002, by Quintijn Hoogenboom
# Licensed under the LGPL.
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
import re
from dragonfly.actions import action_key, action_text

chord_pattern = re.compile(r'(\{.*?\})')
split_text_from_key = re.compile(r'(\w+)(.*$)')

# more synonyms can be defined here if necessary:
synonym_keys = dict(esc="escape")
def sendkeys(keys):
    """sends keystrokes via dragonfly class :code:`action_key.Keys`
   
Keystrokes following Unimacro/Vocola convention are translated into Dragonfly notation:

:code:`"{shift+right 4}"` is translated to :code:`"s-right:4"`.
   
Multiple keystrokes, pauses before the individual keys and a pause after all the keys can be given in Dragonfly notation:

:code:`"{down/25:12/100}"` presses the `down` key 12 times,
with a pause before each key of 25/100 seconds
and a pause of 1 second (100/100 seconds) after the 12 down presses


Extra notation (Quintijn):

:code:`"!"` at the end of a chord (the keystrokes definition between the braced, :code:`{` and :code:`}`,
triggers the `use_hardware=True` event,
comparable with the SendSystemKeys mechanism of Dragon.

Tested at bottom of this file interactively...
    """
    m = chord_pattern.search(keys)
    if m:
        matches = chord_pattern.split(keys)
        # print(matches)
        for part in matches:
            if not part:
                continue
            if part.startswith('{'):
                part = part[1:-1].lower()  # strip { and } and make lowercase
                use_hardware = False
               
                if part.endswith("!"):
                    use_hardware = True
                    part = part.rstrip("!")
                    part = part.rstrip()
                parts = part.split("+")
                if len(parts) > 1:
                    rest = parts[-1]
                    modifiers = [p[0] for p in parts[:-1]]
                    modifiers = ''.join(modifiers) + "-"
                else:
                    # no modifiers:
                    modifiers = ""
                    rest = part
                m_key = split_text_from_key.match(rest)
                if m_key:
                    key, attr = m_key.groups()
                else:
                    raise ValueError('sendkeys, found no key to press in {part}, total: {keys}')
                key = synonym_keys.get(key, key)

                if attr.find(' ') == 0:
                    attr = attr.replace(' ', ':', 1)

                key = modifiers + key + attr
                # print(f'sendkeys, key: {key}')        
                action_key.Key(key, use_hardware=use_hardware).execute()
            else:
                # part_keys = ','.join(t for t in part).replace(" ", "space")
                # action_key.Key(part_keys).execute()
                action_text.Text(part).execute()
    else:
        # all_keys = ','.join(t for t in keys).replace(" ", "space")
        # action_key.Key(all_keys).execute()
        action_text.Text(keys).execute()
    
if __name__ == "__main__":
    # sendkeys("{a 3}") #aaa
    
    # try holding down a key:
    # (first selects slow 4 lines, then fast, then goes back to start)
    # sendkeys("{shift:down}{down/25:4}{shift:up/100}")
    # sendkeys("{shift:down}{down:4/100}{shift:up}")
    # sendkeys("{Up:8}{down 8}")  #
    
    # leaves empty:
    # t3 = 'abc{shift+left/50}{del/100}def{shift+left 2/50}ghi{left 4/100}{shift+end/100}{del/100}{backspace 2}'  
    # sendkeys(t3)   # leave

    # When running next 3 lines several times, starting with cursor on the last line (after #),
    # you will see abcde appear several times:
    # sendkeys("a b c d e")  
    # sendkeys("x y z {home}")  
    # sendkeys("{ctrl+a/100}") #
    # sendkeys("{ctrl+end!}{up/50}{end}")
    # sendkeys("x y z ")
    sendkeys("test, test, met komma.")
    sendkeys("{ctrl+end}test, test,{ctrl+home}{ctrl+end} met komma.")
    ##
    ##
