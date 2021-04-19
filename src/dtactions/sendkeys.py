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
from dragonfly.actions import action_key

chord_pattern = re.compile(r'(\{.*?\})')
synonym_keys = dict(esc="escape")

def sendkeys(keys):
    """send keystrokes via dragonfly action_key.Keys() function
   
    "!" at the end of a chord, triggers use_hardware=True
    """
    m = chord_pattern.search(keys)
    if m:
        matches = chord_pattern.split(keys)
        print(matches)
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
                    key = parts[-1]
                    if key in synonym_keys:
                        key = synonym_keys[key]
                    modifiers = [p[0] for p in parts[:-1]]
                    key = ''.join(modifiers) + "-" + key
                else:
                    key = part
                if key.find(' ') > 0:
                    key = key.replace(' ', ':')
                    
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
    t3 = '{shift+left!}def{shift+left 2}ghi{left 4}{shift+end}{del}'
    sendkeys(t3)
    
    sendkeys("abcde")
    # abcde
    