
.. _RefSendkeys:

sendkeys
========

The :code:`sendkeys` function sends keystrokes to the foreground window.

In fact, this function is used, when you call the :code:`natlink.playString` function of Natlink!


So you can either do:

::

    from dtactions.sendkeys import sendkeys
    (...)
    sendkeys('hello world')

or:

::

    from natlink import playString
    (...)
    playString('hello world again')



