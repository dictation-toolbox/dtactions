# This file is part of Dragonfly.
# (c) Copyright 2007, 2008 by Christo Butcher
# Licensed under the LGPL.
"""shortcut to the sendkeys mechanism of Natlink/Vocola

usage: import sendkeys

and then: sendkeys.sendkeys("keystrokes")

This module now adopts the dragonfly action_key and action_text modules,
so {alt+w} must be converted to "a-w" etc.

Some keys require a synonym, for example "esc", which is "escape" in action_key.py

(Quintijn Hoogenboom, 2021-04-04)
"""
import re
from dragonfly.actions import action_key, action_text

chord_pattern = re.compile(r'(\{.*?\})')
synonym_keys = dict(esc="escape")

def sendkeys(keys):
    """go via dragonfly action_key
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
                parts = part.split("+")
                if len(parts) > 1:
                    # pylint: disable=R1715
                    key = parts[-1]
                    if key in synonym_keys:
                        key = synonym_keys[key]
                    modifiers = [p[0] for p in parts[:-1]]
                    key = ''.join(modifiers) + "-" + key
                else:
                    key = part
                if key.find(' ') > 0:
                    key = key.replace(' ', ':')
                    
                action_key.Key(key).execute()
            else:
                action_text.Text(part).execute()
    else:
        action_text.Text(keys).execute()
    
if __name__ == "__main__":
    # t1 = 'a{ctrl+o}hallo{escape}k'
    # sendkeys(t1)
    # t2 = 'abc{left}def{left 2}ghi'
    # sendkeys(t2)
    
    # only ab should remain:
    t3 = 'abc{shift+left}def{shift+left 2}ghi{left 4}{shift+end}{del}'
    sendkeys(t3)
    


