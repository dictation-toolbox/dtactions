
.. _RefSendkeys:

sendkeys
========

The :code:`sendkeys` function sends keystrokes to the foreground window.

The format that is used in Unimacro and Vocola is translated into Dragonfly
format, and the :code:`dragonfly.actions.action_key.Key` class performs the
actions.

This replaces the :code:`natlink.playString` function of Natlink
and the :code:`SendInput` of Vocola.

sendkeys module
---------------
.. automodule:: dtactions.sendkeys
   :members: