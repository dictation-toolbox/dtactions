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
split_text_from_key = re.compile(r'(\w+)(.*$)')

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
    
    tested at bottom of this file interactively...
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
                part_keys = ','.join(t for t in part)
                action_key.Key(part_keys).execute()
                # action_text.Text(part).execute()
    else:
        all_keys = ','.join(t for t in keys)
        action_key.Key(all_keys).execute()
        # action_text.Text(keys).execute()
    
if __name__ == "__main__":
    pass
    # sendkeys("{a 3}") #aaa
    
    # try holding down a key:
    # (first selects slow 4 lines, then fast, then goes back to start)
    # sendkeys("{shift:down}{down/25:4}{shift:up/100}")
    # sendkeys("{shift:down}{down:4/100}{shift:up}")
    # sendkeys("{up 8}")
    
    # leaves empty:
    # t3 = 'abc{shift+left/50}{del/100}def{shift+left 2/50}ghi{left 4/100}{shift+end/100}{del/100}{backspace 2}'  
    # sendkeys(t3)   # leave

    # When running next 3 lines several times, starting with cursor on the last line (after #),
    # you will see abcde appear several times:
    # sendkeys("abcde")  
    # sendkeys("{ctrl+a/100}") #
    # sendkeys("{ctrl+end!}{up/50}{end}")
    # 
