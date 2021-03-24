;for Unimacro, changes %hndle% into the current foreground hndle.
;this script is preferably called from a string, as is used in natlinkutilsqh.SetForegroundWindow
;

WinActivate, ahk_id  %hndle%
