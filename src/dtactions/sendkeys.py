# This file is part of Dragonfly and Unimacro
# (c) Copyright 2007, 2008 by Christo Butcher
# (c) Copyright 2002, by Quintijn Hoogenboom
# Licensed under the LGPL.
"""shortcut to the sendkeys mechanism of Natlink/Vocola

usage: :code:`from dtactions.sendkeys import sendkeys`

and then: :code"`sendkeys("keystrokes")`

This module now adopts the dragonfly action_key and action_text modules,
so {alt+w} must be converted to "a-w" etc.

Some keys require a synonym, for example "esc", which is "escape" in action_key.py

(Quintijn Hoogenboom, 2021-04-04)
"""
import re
import time
from dragonfly.actions import action_key

chord_pattern = re.compile(r'(\{.*?\})')
text_of_key = re.compile(r'\w+')

synonym_keys = dict(esc="escape")
def sendkeys(keys):
    """send keystrokes via dragonfly action_key.Keys() function
   
   Keystrokes following Unimacro/Vocola convention are translated into Dragonfly notation:
   `{shift+right 4}` is translated to `s-right:4`
   
   Pause before the keys and after the keys can be given in Dragonfly notation:
   `{down/25:10/100}` goes 10 times down, with pause before next key of 25/100 seconds
                      and 1 second after the 10 down presses

    extra notation (Quintijn):   
    "!" at the end of a chord, triggers use_hardware=True
    """
    m = chord_pattern.search(keys)
    if m:
        matches = chord_pattern.split(keys)
        # print(matches)
        for part in matches:
            if not part:
                continue
            if part.startswith('{'):
                part = part[1:-1]  # strip { and }
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
                m_key = text_of_key.match(rest)
                if m_key:
                    key = m_key.group()
                else:
                    raise ValueError('sendkeys, found no key to press in {part}, total: {keys}')
                key = synonym_keys.get(key, key)

                if key.find(' ') > 0:
                    key = key.replace(' ', ':')

                key = modifiers + key
                    
                action_key.Key(key, use_hardware=use_hardware).execute()
            else:
                part_keys = ','.join(t for t in part)
                action_key.Key(part_keys).execute()
                # action_text.Text(part).execute()
    else:
        all_keys = ','.join(t for t in keys)
        action_key.Key(all_keys).execute()
        # action_text.Text(keys).execute()
    
if __name__ == "__main__":
    # t1 = 'a{ctrl+o}hallo{escape}k'
    # sendkeys(t1)
    # t2 = 'abc{left}def{left 2}ghi'
    # sendkeys(t2)
    
    # only ab should remain:
    t3 = 'abc{shift+left/100}{del/200}def{shift+left 2/100}ghi{left 4/100}{shift+end/100}{del}'
    sendkeys(t3)
    
    sendkeys("abcde")
    sendkeys("{ctrl+a}")
    time.sleep(1)
    sendkeys("{ctrl+end}")
    #
    