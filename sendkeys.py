"""shortcut to the sendkeys mechanism of Natlink/Vocola

usage: import sendkeys

and then: sendkeys.sendkeys("keystrokes")

Note: the notation with +, % and ^ is not valid any more, change
      + to shift
      ^ to ctrl
      % to alt
      
      eg: +{home} becomes {shift+home}
          ^p becomes {ctrl+p}
      etc. 

(Quintijn Hoogenboom, 2020-12-14)
"""
import ExtendedSendDragonKeys
import SendInput


def sendkeys(keys):
    """send keystrokes to foreground window
    """
    if not keys:
        return
    SendInput.send_input(
        ExtendedSendDragonKeys.senddragonkeys_to_events(keys))
    pass

if __name__ == "__main__":
    sendkeys("abc")