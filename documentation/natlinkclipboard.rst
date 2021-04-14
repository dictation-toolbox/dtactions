
.. _RefNatlinkClipboard:

natlinkclipboard
==============================================================================

The :code:`natlinkclipboard` module offers easy access to and manipulation of
the Windows clipboard.  The :class:`Clipboard` class forms the core of this
module.  Each instance of this class is a container with a structure similar
to the system clipboard, mapping content formats to content data.


Clipboard class
------------------------------------------------------------------------------

.. autoclass:: dtactions.natlinkclipboard.Clipboard
   :members:


Utility functions
------------------------------------------------------------------------------

.. autofunction:: dtactions.natlinkclipboard.OpenClipboardCautious
