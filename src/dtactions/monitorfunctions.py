"""gives the monitor information of all the monitors 

and provide functions for moving a (top level) window to a position on
the current or on a different monitor.

The start of this module was taken from O'Reilly, and has been enhanced for use
with NatLink speech recognition commands, see http://qh.antenna.nl and
http://qh.antenna.nl/unimacro. Quintijn Hoogenboom, february 2010.


The basic information is collected in
-MONITOR_INFO, a dictionary with keys the handles of the monitor
    These keys are converted to int, and for 2 monitors probably 65537 and
    65539, and put in global variable MONITOR_HNDLES (a list)

    Each item of MONITOR_INFO is again a dictionary with for example my second monitor info:
    {'Device': '\\\\.\\DISPLAY2',
     'Flags': 0,
     'Monitor': (1680, 0, 2704, 768),
     'Work': (1787, 0, 2704, 768),
     'offsetx': 107,
     'offsety': 0}
    Thus holding the Monitor info and the Work info. As a extra offsetx and offsety are calculated,
    which give the width/height of task bar and possibly other "bars". In this example I have the task bar
    vertically placed at the left side of this second monitor, and it has a width of 107 pixels.

-MONITOR_HNDLES: a list of the available monitors (the handles are int, see above)
-VIRTUAL_SCREEN: a 4 tuple giving the (left, top, right, bottom) of the complete (virtual) screen
-BORDERX, BORDERY: the border width of windows. With this a window can be size a little

MA, WA, RA
Each area (MA: Monitor, WA: Work, RA: restore area of a window) is a 4 length tuple
giving (left, top, right, bottom).

Biggest puzzle of the calculations is changing the restore_area (from GetWindowPlacement) to the Work area of
a new monitor. Important is the offsetx and offsety (difference between Monitor coordinates and Work area
coordinates). The RA (restore_area) is relative to the Work area of the monitor it is on, so the offsetx and offsety
must be subtracted from the calculated coordinates. (something like that)

This module provides functions for getting:

--all the info: monitor_info(force=None): giving the above mentioned data

--which is the nearest monitor, using API functions:
    get_nearest_monitor_window(winHndle)
    get_nearest_monitor_point(point)
(could also make get_nearest_monitor_rect, but was not needed here)

--further info:
    get_other_monitors(mon): give a list of the other monitor hndles (after collecting the current monitor)
    
-- for individual points (used by natlinkutilsqh.py in Unimacro (NatLink, speech recognition)
    is_inside_monitor(point): returns True if the point is inside one of the monitors (monitor area)
    get_closest_position(point) returns a point that is closest to an outside point on one of the monitors
    
-- for user calls:
    maximize_window(winHndle): just maximize
    minimize_window(winHndle): just minimize

    move_to_monitor(winHndle, newMonitor, oldMonitor, resize): move to another monitor
        preserving position of restore_area as much as possible.
        resize: 0 if window is (assumed to be) fixed in size, can be found with:
    window_can_be_resized(winHndle):
        return 1 if a window can be resized (like Komodo etc). Not eg calc.
        
    restore_window(winHndle, monitor, ...): placing in various spots and widths/heights
        see at definition for parameters
    
-- helper functions:
    
"""

import win32api
import math
import win32gui
import win32con
import time
import pprint
import types
import math
import copy
from dtactions import messagefunctions # only for taskbar position (left, bottom etc)

MONITOR_INFO = None
MONITOR_HNDLES = None
BORDERX = BORDERY = None
VIRTUAL_SCREEN = None
NMON = None

def monitor_info():
    """collecting all the essential information
    """
    global MONITOR_INFO
    global MONITOR_HNDLES
    global BORDERX, BORDERY
    global VIRTUAL_SCREEN
    global NMON  # number of monitors
    NMON = win32api.GetSystemMetrics(win32con.SM_CMONITORS)    # 80
    if NMON < 1:
        raise ValueError("monitor_info: system should have at least one monitor, strange result: %s"% NMON)
    MONITOR_INFO = {}
    ALL_MONITOR_INFO = [item for item in win32api.EnumDisplayMonitors(None, None)]
    for hndle, dummy, monitorRect in ALL_MONITOR_INFO:
        hndle = int(hndle)
        MONITOR_INFO[hndle] = win32api.GetMonitorInfo(hndle)
        m = MONITOR_INFO[hndle]
        m['offsetx'] = m['Work'][0] - m['Monitor'][0]
        m['offsety'] = m['Work'][1] - m['Monitor'][1]
        
    MONITOR_HNDLES = list(MONITOR_INFO.keys())
    BORDERX = win32api.GetSystemMetrics(win32con.SM_CXBORDER)  # 5
    BORDERY = win32api.GetSystemMetrics(win32con.SM_CYBORDER)  # 6
    VIRTUAL_SCREEN = []
    VIRTUAL_SCREEN.append(win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN))  # 76
    VIRTUAL_SCREEN.append(win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN))  # 77
    VIRTUAL_SCREEN.append(VIRTUAL_SCREEN[0] + win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN))  # 78
    VIRTUAL_SCREEN.append(VIRTUAL_SCREEN[1] + win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN))  # 79

def getScreenRectData():
    """return width, height, xmin, ymin, xmax, ymax for complete screen
    """
    monitor_info()
    vs = VIRTUAL_SCREEN
    return vs[2]-vs[0], vs[3]-vs[1], vs[0], vs[1], vs[2], vs[3]

def fake_monitor_info_for_testing(nmon, virtual_screen):
    """test if changed monitor data come through in calling program
    """
    global NMON, VIRTUAL_SCREEN
    NMON = nmon
    VIRTUAL_SCREEN = virtual_screen
    print('fake_monitor_info_for_testing, set NMON to %s and VIRTUAL_SCREEN to %s'% (nmon, virtual_screen))
    
    
###########################################
# three wrapper functions around api calls:
#    -MonitorFromPoint
#    -MonitorFromRect
#    -MonitorFromWindow
# using constant: MONITOR_DEFAULTTONEAREST



def get_nearest_monitor_window(winHndle):
    """give monitor number of the monitor which is nearest to the window
    
    input the handle of the window, most often got by:
        winHndle = win32gui.GetForegroundWindow()
    output: the monitor handle as an integer (65537, 65539 often on a 2 monitor configuration)
    """
    mon = win32api.MonitorFromWindow(winHndle, win32con.MONITOR_DEFAULTTONEAREST)
    return int(mon)

def get_nearest_monitor_point( point ):
    """give monitor number of the monitor which is nearest to the point
    point is a tuple (x, y) of coordinates
    """
    mon = win32api.MonitorFromPoint(point , win32con.MONITOR_DEFAULTTONEAREST)
    return int(mon)

def get_other_monitors(mon):
    """give list of other monitors
    """
    mon = int(mon)
    if MONITOR_HNDLES is None:
        monitor_info()
    if MONITOR_HNDLES is None:
        raise ValueError("no monitor handles found")
    #print 'mon: %s, int(mon): %s'% (mon, int(mon))
    return [hndle for hndle in MONITOR_HNDLES if hndle != mon]
    #other = [hndle for ]

def get_current_monitor_rect( pos):
    """get rectangle positions of the foreground monitor,
    pass (xpos, ypos) of current mouse position
    """
    monitor_info()
    mon = get_nearest_monitor_point( pos )
    moninfo = MONITOR_INFO[mon]
    rect = copy.copy(moninfo['Monitor'])
    return rect
def get_current_monitor_rect_work( pos):
    """get rectangle positions of the foreground monitor,
    pass (xpos, ypos) of current mouse position
    """
    monitor_info()
    mon = get_nearest_monitor_point( pos )
    moninfo = MONITOR_INFO[mon]
    rect = copy.copy(moninfo['Work'])
    return rect

def get_monitor_rect(monitorIndex):
    """get rectangle positions of indexed monitor
    """
    monitor_info()
    try:
        moninfo = MONITOR_INFO[MONITOR_HNDLES[monitorIndex]]
    except IndexError:
        print('monitorfunctions.get_monitor_rect: monitorindex not available %s'% monitorIndex)
        return
    rect = copy.copy(moninfo['Monitor'])
    return rect

def get_monitor_rect_work(monitorIndex):
    """get rectangle positions of indexed monitor, the Work area
    
    so ignoring Dragon bar or taskbar.
    """
    monitor_info()
    try:
        moninfo = MONITOR_INFO[MONITOR_HNDLES[monitorIndex]]
    except IndexError:
        print('monitorfunctions.get_monitor_rect: monitorindex not available %s'% monitorIndex)
        return
    rect = copy.copy(moninfo['Work'])
    return rect

def window_can_be_resized(windowHndle):
    """returns 1 if the window can be resized (has a maximize button)
    """
    wstyle = win32api.GetWindowLong(windowHndle, win32con.GWL_STYLE)
    canBeResized = (wstyle & win32con.WS_MAXIMIZEBOX == win32con.WS_MAXIMIZEBOX)
    #print 'wstyle: %x, maximizebox: %x, '
    return canBeResized

def get_taskbar_position():
    """return left, top, right or bottom
    """
    monitor_info()
    m = MONITOR_INFO
    hApp = messagefunctions.findTopWindow(wantedClass='Shell_TrayWnd')
    if not hApp:
        print('no taskbar (system tray) found')
        return
    info = list( win32gui.GetWindowPlacement(hApp) )
    RA = list(info[4])
    mon = get_nearest_monitor_window(hApp)
    work = list(MONITOR_INFO[mon]['Work'])
    #print 'RA:   %s'% RA
    #print 'work: %s'% work
    # bottom 3
    # right 2
    # top 1
    # left 0
    if RA[2] <= work[0] + 2*BORDERX:
        return 'left'  # right pos of RA == left pos of work area
    elif RA[1] >= work[3] - 2*BORDERY: 
        return 'bottom' # top of RA == bottom of work area
    elif RA[0] >= work[2] - 2*BORDERX: 
        return 'right' # left of RA == right of work area (account for Dragon bar)
    elif RA[3] <= work[1] + 2*BORDERY:
        return 'top'

#======================== on same monitor: ==========================================

def restore_window(winHndle, monitor=None, xwidth=None, ywidth=None,
                        xpos=None, ypos=None, keepinside=None):
    """move window inside same monitor, restore format
    
    (for moving an amount of pixels or across the monitor border use move_window
     for resizing an amount of pixels, to an edge or across the monitor border use
    stretch_window or shrink_window)
    
    if no parameters are passed, the window is restored, but inside its
    working area if it fits. If it does not fit, the left top is shown.

    winHndle: hndle of the program
    monitor: hndle of the monitor (from which the monitor_info is got)
    xwidth, ywidth: passed to width in _get_new_coordinates_same_monitor, see there
    xpos, ypos: passed to position in _get_new_coordinates_same_monitor, see there
    keepinside: should normally be the same as the window_can_be_resized value,
        default taken from this function
    
    """
    monitor_info()
    if not monitor:
        monitor = get_nearest_monitor_window(winHndle)
    toRestore = win32con.SW_RESTORE
    info = list( win32gui.GetWindowPlacement(winHndle) )
    MI = MONITOR_INFO[monitor]
    RA = list(info[4])
    maximized =  ( info[1] == win32con.SW_SHOWMAXIMIZED)
    if keepinside is None:
        keepinside = not window_can_be_resized(winHndle)

    # for non resizable windows:
    if not window_can_be_resized(winHndle):
        print('no resize window, setting xwidth and ywidth to 0(width was: %s, %s'%\
                (xwidth, ywidth))
        ## center in correct way:
        #if xwidth:
        #    xpos = 0.5
        #if not ywidth:
        #    ypos = 0.5
        # set width to non resize:
        xwidth = ywidth = 0
    #print 'maximized: %s'% maximized
    #print 'previous  RA: %s'% RA
    newRA = _change_restore_area(RA, monitor_info=MI,
                                 xpos=xpos, ypos=ypos,
                                 xwidth=xwidth, ywidth=ywidth,
                                 keepinside=keepinside)
    info[1] = toRestore
    info[4] = tuple(newRA)
    win32gui.SetWindowPlacement(winHndle, tuple(info) )
    
def move_window(winHndle, direction, amount, units='pixels',
                keepinside=None, keepinsideall=1, monitor=None):
    """moves a window, can go across monitor borders and out of the total area
    
        winHndle: hndle of the program
        monitor: hdnle of the monitor (is collected if not passed)
        direction: a string, left, right, up, down or
            a direction in degrees (0 = up, 90 = right, 180 = down, 270 = left)
        amount: number of pixels (for moving to left edge, use restore_window)
        keepinside: keep inside the work area of the monitor
        keepinsideall: keep inside the virtual area of all monitors.
        
    """
    monitor_info()
    if not monitor:
        monitor = get_nearest_monitor_window(winHndle)
    toRestore = win32con.SW_RESTORE
    info = list( win32gui.GetWindowPlacement(winHndle) )
    MI = MONITOR_INFO[monitor]
    RA = list(info[4])
    resize = 0
    newRA = _move_resize_restore_area(RA, resize, direction, amount, units=units,
                               keepinside=keepinside, keepinsideall=keepinsideall,
                               monitor_info=MI)
    info[1] = toRestore
    info[4] = tuple(newRA)
    win32gui.SetWindowPlacement(winHndle, tuple(info) )

def stretch_window(winHndle, direction, amount, units='pixels',
                keepinside=None, keepinsideall=1, monitor=None):
    """resize a window, making it larger, can go across monitor borders and out of the total area
    
        winHndle: hndle of the program
        monitor: hdnle of the monitor (is collected if not passed)
        direction: a string, left, right, up, down or
            a direction in degrees (0 = up, 90 = right, 180 = down, 270 = left)
        amount: number of pixels (for moving to left edge, use restore_window)
        keepinside: keep inside the work area of the monitor
        keepinsideall: keep inside the virtual area of all monitors.
        
    """
    monitor_info()
    if not monitor:
        monitor = get_nearest_monitor_window(winHndle)
    toRestore = win32con.SW_RESTORE
    info = list( win32gui.GetWindowPlacement(winHndle) )
    MI = MONITOR_INFO[monitor]
    RA = list(info[4])
    resize = 1
    newRA = _move_resize_restore_area(RA, resize, direction, amount, units=units,
                               keepinside=keepinside, keepinsideall=keepinsideall,
                               monitor_info=MI)
    info[1] = toRestore
    info[4] = tuple(newRA)
    win32gui.SetWindowPlacement(winHndle, tuple(info) )

def shrink_window(winHndle, direction, amount, units='pixels',
                keepinside=None, keepinsideall=1, monitor=None):
    """resize a window, making it smaller, can go across monitor borders and out of the total area

    
        winHndle: hndle of the program
        monitor: hdnle of the monitor (is collected if not passed)
        direction: a string, left, right, up, down or
            a direction in degrees (0 = up, 90 = right, 180 = down, 270 = left)
        amount: number of pixels (for moving to left edge, use restore_window)
        when shrink relative, take size of the window rather than distance to edge or corner

        keepinside: keep inside the work area of the monitor
        keepinsideall: keep inside the virtual area of all monitors.
        
    """
    monitor_info()
    if not monitor:
        monitor = get_nearest_monitor_window(winHndle)
    toRestore = win32con.SW_RESTORE
    info = list( win32gui.GetWindowPlacement(winHndle) )
    MI = MONITOR_INFO[monitor]
    RA = list(info[4])
    resize = -1
    newRA = _move_resize_restore_area(RA, resize, direction, amount, units=units,
                               keepinside=keepinside, keepinsideall=keepinsideall,
                               monitor_info=MI)
    info[1] = toRestore
    info[4] = tuple(newRA)
    win32gui.SetWindowPlacement(winHndle, tuple(info) )
    

def maximize_window(winHndle):
    """maximize to the monitor (it is on at the moment)"""
    monitor_info()
    toMaximize = win32con.SW_SHOWMAXIMIZED
    info = list( win32gui.GetWindowPlacement(winHndle) )
    info[1] = toMaximize
    win32gui.SetWindowPlacement(winHndle, tuple(info) )

def minimize_window(winHndle):
    """minimize the window"""
    monitor_info()
    toMinimize = win32con.SW_SHOWMINIMIZED
    info = list( win32gui.GetWindowPlacement(winHndle) )
    info[1] = toMinimize
    win32gui.SetWindowPlacement(winHndle, tuple(info) )


def _move_resize_restore_area(RestoreArea, resize, direction, amount, units,
                              keepinside, keepinsideall, monitor_info, min_size= (100, 70) ): 
    """change the coordinates of the RA according to direction and amount
    
    the amount in combination with direction provides numerous, confusing possibilities
    
    RA: the restore area
    resize: 0 = move, 1 = resize in the target direction
                      -1 = resize away from target direction (making smaller)
    direction: left|right|up|down|lefttop|righttop|leftbottom|rightbottom or
               a number of degrees (0=up, 90=right, 180=down, 270=left)
    units: relative (amount between 0 and 1)|pixels|percent (of screen size)
    amount: 0 < amount <= 1: for units == relative
            >= 1 (int): an absolute number of pixels|percent of screen size
    keepinside: whether to check for the window being inside after the move/resize
    keepinsideall: whether to check for the window being inside the virtual screen (all monitors)
    monitor_info: the info (dict) of the current monitor
    
    """
    RA = RestoreArea[:]
    raWidth = RA[2] - RA[0]
    raHeight = RA[3] - RA[1]
    WA = monitor_info['Work']
    MA = monitor_info['Monitor']
    #offsetx = monitor_info['offsetx']
    #offsety = monitor_info['offsety']
    
    # norm to 0 oriented
    RA[0] -= MA[0]
    RA[1] -= MA[1]
    RA[2] -= MA[0]
    RA[3] -= MA[1]
    WAWidth = WA[2] - WA[0]
    WAHeight = WA[3] - WA[1]
    boundingbox = [0, 0, WAWidth, WAHeight]
    alpha, distance = None, None
    if direction == 'lefttop':
        alpha, distance = _get_angle_distance_side_corners(RA, boundingbox, 0, 0)
    elif direction == 'righttop':
        alpha, distance = _get_angle_distance_side_corners(RA, boundingbox, 1, 0)
    elif direction == 'leftbottom':
        alpha, distance = _get_angle_distance_side_corners(RA, boundingbox, 0, 1)
    elif direction == 'rightbottom':
        alpha, distance = _get_angle_distance_side_corners(RA, boundingbox, 1, 1)
    elif direction == 'right':
        alpha, distance = 90, boundingbox[2] - RA[2]
    elif direction in ('down','bottom'):
        alpha, distance = 180, boundingbox[3] - RA[3]
    elif direction == 'left':
        alpha, distance = 270, RA[0] - boundingbox[0]
    elif direction in ('top', 'up'):
        alpha, distance = 0, RA[1] - boundingbox[1]
    elif type(direction
              ) in (float, int):
        alpha = direction

    # setting reverse (for making smaller)
    if resize == -1:
        reverse = resize
    else:
        reverse = 1
    
    if units == 'pixels':
        if alpha is None:
            raise ValueError('alpha unknown, do not know where to move to')
        deltax, deltay = _get_deltax_deltay_from_angle_distance(alpha, amount)
        amountx =  _round_float(deltax * reverse)
        amounty =  _round_float(deltay * reverse)

    elif units == 'relative':
        if alpha is None:
            raise ValueError('alpha unknown, do not know where to move to')
        distance = _get_distance_in_direction(RA, boundingbox, alpha)  #if distance is None:
        #    nearest_corner, distance = 
        deltax, deltay = _get_deltax_deltay_from_angle_distance(alpha, distance)
        amountx = _round_float(deltax * reverse * amount)
        amounty = _round_float(deltay * reverse * amount)
    else:
        raise ValueError("_move_resize_restore_area: units should be 'relative' or 'pixels', not '%s'"% units)
    side_corner = None
    if direction == 'up':
        side_corner = 'top'
    elif direction == 'down':
        side_corner = 'bottom'
    elif type(side_corner) == bytes:
        side_corner = direction
        
        
    RA = _adjust_coordinates_side_corners(RA, amountx, amounty, resize, side_corner=side_corner, min_size=min_size)

    if keepinside:
        fixed_width = not resize
        RA[0], RA[2] = keepinside_restore_area(RA[0], RA[2], WAWidth, fixed_width=fixed_width, margin=1)
        RA[1], RA[3] = keepinside_restore_area(RA[1], RA[3], WAHeight, fixed_width=fixed_width, margin=1)
        
    # correct for WA again:
    RA[0] += MA[0]
    RA[1] += MA[1]
    RA[2] += MA[0]
    RA[3] += MA[1]

    if keepinsideall:
        fixed_width = not resize # if it is a move
        RA[0], RA[2] = keepinside_all_screens(RA[0], RA[2], VIRTUAL_SCREEN[0], VIRTUAL_SCREEN[2], fixed_width)
        RA[1], RA[3] = keepinside_all_screens(RA[1], RA[3], VIRTUAL_SCREEN[0], VIRTUAL_SCREEN[2], fixed_width)
        
    return RA

def _get_distance_in_direction(RestoreArea, boundingbox, angle):
    """calculate the distance from window (restore area) to the bounding box
    
    alpha is direction to appropriate cornerpoint
    return distance  (0 if direction fails, distance to cornerpoint maximum)
    """
    RA = RestoreArea[:]
    angle = angle % 360
    if 0 <= angle < 90:
        alpha, distance = _get_angle_distance_side_corners(RA, boundingbox, 1, 0)
        if not 0 <= alpha < 90: return 0 # not same direction
        angle, alpha = 90 - angle, 90 - alpha # for easier goniometry
    elif 90 <= angle < 180:
        alpha, distance = _get_angle_distance_side_corners(RA, boundingbox, 1, 1)
        if not 90 <= alpha < 180: return 0
        angle, alpha = angle - 90, alpha -90
    elif 180 <= angle < 270:
        alpha, distance = _get_angle_distance_side_corners(RA, boundingbox, 0, 1)
        if not 180 <= alpha < 270: return 0
        angle, alpha = 270 - angle, 270 - alpha
    elif 270 <= angle <= 360:
        alpha, distance = _get_angle_distance_side_corners(RA, boundingbox, 0, 0)
        if not 270 <= alpha  < 360: return 0 # not same direction
        angle, alpha = angle - 270, alpha - 270 # for easier goniometry
    if angle == alpha:
        return distance
    elif angle < alpha:
        return distance * math.cos(math.radians(alpha))/math.cos(math.radians(angle))
    else:
        return distance * math.sin(math.radians(alpha))/math.sin(math.radians(angle))


def _round_float(f):
    """round to integer
>>> _round_float(0)
0
>>> _round_float(0.4)
0
>>> _round_float(0.5)
1
>>> _round_float(-0.5)
-1
>>> _round_float(-1.9)
-2
>>> _round_float(1.9999999999999999)
2
    """
    if f > 0:
        return int(f + 0.5)
    elif f < 0:
        return int(f - 0.5)
    else:
        return 0

def _get_angle_distance_side_corners(RA, boundingbox, xindex, yindex):
    """return angle and distance of RA point and WA point
       give RA and WA and:
        xindex: 0 left, 1 right
        yindex: 0 top, 1 bottom
        
    return (alpha (in degrees), distance (in pixels))
    """
    bb = boundingbox ##(should be [0, 0, widhtofWA, heightofWA])
    if xindex == 0 and yindex == 0:
        return _get_angle_distance( RA[0], RA[1], bb[0], bb[1])
    elif xindex == 0 and yindex == 1:
        return _get_angle_distance( RA[0], RA[3], bb[0], bb[3])
    elif xindex == 1 and yindex == 0:
        return _get_angle_distance( RA[2], RA[1], bb[2], bb[1])
    elif xindex == 1 and yindex == 1:
        return _get_angle_distance( RA[2], RA[3], bb[2], bb[3])
    else:
        raise ValueError("_get_angle_distance_side_corners: invalid parameters xindex (%s), yindex (%s) (should be 0 or 1)"%
                         (xindex, yindex))
        
def _get_angle_distance(px, py, qx, qy):
    """calculate angle and distance of two points
    px, py: (pixel) coordinates of first point (corner of window)
    qx, qy: (pixel) coordinates of second point (corner of monitor)
    return: (angle, dist)
            angle (in degrees) up = 0, right = 90 etc, viewed from p to q
            dist: pixels () (rounded to int)
(see unittestMonitorfunctions for more tests)
#>>> _get_angle_distance(20, 100, 0, 0)# fourth quadrant
#(348.69, 101.98)
#>>> _get_angle_distance(0, 100, 0, 0)# point up
#(0, 100.0)
#>>> _get_angle_distance(300, 600, 1000, 600)# point right
#(90, 700.0)

    """
    dx, dy = qx-px, qy-py
    dist = math.sqrt(dx*dx + dy*dy)
    if dx == 0:
        if dy == 0:
            return (0, 0)
        elif dy > 0:
            return (180, dist)
        else:
            return (0, dist)
    elif dy == 0:
        if dx > 0:
            return (90, dist)
        else:
            # dx < 0
            return (270, dist)
    else:
        alpha = math.degrees(math.atan(math.fabs(1.0*dy/dx)))
        if dx > 0 and dy < 0: # first quadrant
            alpha = 90 - alpha
        elif dx > 0 and dy > 0:
            # second quadrant
            alpha = 90 + alpha
        elif dx < 0 and dy > 0:
            # third quadrant
            alpha = 270 - alpha
        elif dx < 0 and dy < 0:
            # fourth quadrant
            alpha = 270 + alpha
        else:
            raise ValueError("impossible to come here")
        return alpha, dist
  
def _get_deltax_deltay_from_angle_distance(angle, distance):
    """calculates "back" the px - py and qx - qy from _get_angle_distance

#>>> angle, distance = _get_angle_distance(20, 100, 0, 0)# fourth quadrant
#>>> angle, distance
#(348.69, 101.98)
#>>> _get_deltax_deltay_from_angle_distance(angle, distance)
#(-20, 99.99999)
#
#>>> angle, distance = _get_angle_distance(300, 600, 1000, 600)# point right
#>>> angle, distance
#(90, 700.0)
#>>> _get_deltax_deltay_from_angle_distance(angle, distance)
#(700.0, 0)
test with unittestMonitorfunctions)
    
    """
    alpha = angle
    deltax = distance * math.sin(math.radians(alpha)) 
    deltay = - distance * math.cos(math.radians(alpha)) 
    return deltax, deltay
    
def _adjust_coordinates_side_corners(RestoreArea, amountx, amounty, resize, side_corner=None, min_size=(100,70)):
    """adjust RestoreArea in the wanted direction
       give RA
        amountx and amounty (can be negative)
                eg amountx = -5 means go left 5, resize == 1 doing on left side (larger)
                                                 resize == -1 doing on right side (smaller)
                                                (resize false: does not matter where)
        resize (include the other point)
                resize = 0(None): move, also change the opposite point
                resize = 1: make larger (see note below)
                resize = -1: make smaller (only relevant if side_corner is NOT given)
        side_corner in lefttop, righttop leftbottom rightbottom or
                       left right up down or center or None
                (note: if side_corner is given, a resize (if given) is always into the
                    direction with respect to side_corner, so smaller/larger is not controlled by
                    resize)
            (assume amount points in correct direction)
       resize: 0 = cannot resize (move)
       min_size (in case of resize, minimum size of width/height)
        
    returns adjusted RA
    
(see unittestMonitorfunctions for tests)
    """
    RA = RestoreArea[:]
    validvaluesside_corner = ('rightbottom', 'righttop', 'leftbottom', 'lefttop',
                              'left', 'right', 'top', 'bottom',
                              'leftcenter', 'rightcenter', 'center',
                              'centertop', 'centerbottom')
    if side_corner is None:
        if resize == -1:
            reverse = resize
        else:
            reverse = 1
            
        if amountx == 0:
            if amounty == 0:
                return RA  # no changes to be expected
            elif amounty * reverse > 0:
                side_corner = 'bottom'
            else:
                # amounty < 0:
                side_corner = 'top'
        elif amountx * reverse > 0:
            if amounty == 0:
                side_corner = 'right'
            elif amounty * reverse > 0:
                side_corner = 'rightbottom'
            else:
                # amounty < 0:
                side_corner = 'righttop'
        else:
            # amountx * reverse negative
            if amounty == 0:
                side_corner = 'left'
            elif amounty * reverse > 0:
                side_corner = 'leftbottom'
            else:
                # amounty < 0:
                side_corner = 'lefttop'


    if side_corner not in validvaluesside_corner:
        raise ValueError("side_corner ('%s') should be one of the valid values: %s"%
                         (side_corner, validvaluesside_corner))
    if resize and resize not in (-1, 1, True):
        raise ValueError("resize ('%s') should be false, or True, 1 (larger) or -1 (smaller)"% resize)
        
    if amountx:
        if side_corner in ('lefttop', 'leftbottom', 'left'):
            RA[0] += amountx
            if resize:
                # allow for minimum size of result:
                RA[0] = min(RA[0], RA[2]-min_size[0])
            else:
                RA[2] += amountx
        elif side_corner in ('centertop', 'centerbottom', 'center'):
            # resize not relevant
            left = int(amountx/2)
            right = amountx - left
            RA[0] -= left
            RA[2] += right
        else:           
            RA[2] += amountx
            if resize:
                RA[2] = max(RA[0]+min_size[0], RA[2])
            else:
                RA[0] += amountx

    if amounty:
        if side_corner in ('lefttop', 'righttop', 'top'):
            # doing y upwards:
            RA[1] += amounty
            if resize:
                RA[1] = min(RA[1], RA[3]-min_size[1])
            else:
                RA[3] += amounty
        elif side_corner in ('leftcenter', 'rightcenter', 'center'):
            # resize not relevant
            up = int(amounty/2)
            down = amounty - right
            RA[1] -= up
            RA[3] += down
        else:
            # doing y downwards
            RA[3] += amounty
            if resize:
                RA[3] = max(RA[1]+min_size[1], RA[3])
            else:
                RA[1] += amounty

    return RA

        
def _change_restore_area(RA, monitor_info, xwidth, ywidth,
                            xpos, ypos, keepinside): 
    """change the placing or the RA in the same monitor
     for parameters, see restore_window above and
     _get_new_coordinates_same_monitor below
    
    """
    raWidth = RA[2] - RA[0]
    raHeight = RA[3] - RA[1]
    WA = monitor_info['Work']
    MA = monitor_info['Monitor']
    #offsetx = monitor_info['offsetx']
    #offsety = monitor_info['offsety']
    
    # norm to 0 oriented
    RA[0] -= MA[0]
    RA[1] -= MA[1]
    RA[2] -= MA[0]
    RA[3] -= MA[1]
    WAWidth = WA[2] - WA[0]
    WAHeight = WA[3] - WA[1]
    
    RA[0], RA[2] = _get_new_coordinates_same_monitor(begin=RA[0], end=RA[2], size=WAWidth,
                                                    width=xwidth, positioning=xpos)
    RA[1], RA[3] = _get_new_coordinates_same_monitor(begin=RA[1], end=RA[3], size=WAHeight,
                                                    width=ywidth, positioning=ypos)

    fixed_width = not keepinside
    RA[0], RA[2] = keepinside_restore_area(RA[0], RA[2], WAWidth, fixed_width=fixed_width, margin=1)
    RA[1], RA[3] = keepinside_restore_area(RA[1], RA[3], WAHeight, fixed_width=fixed_width, margin=1)

    # correct for WA again:
    RA[0] += MA[0]
    RA[1] += MA[1]
    RA[2] += MA[0]
    RA[3] += MA[1]
    
    return RA

def _get_new_coordinates_same_monitor(begin, end, size, width, positioning, minWidth=10):
    """sqeeze or make larger  the coordinates of the window in either direction
    
    think in width here:
    begin, end, actual coordinates with respect to left or top of
        actual monitor 
    size: allowed size of the work area (width of height)
    width: if false (None or 0) the same width is taken
           if 1: size is taken
           if < 1 (and > 0, assumed) this part of the size is taken
    
    positioning: give the wanted position on the monitor:
        None: do nothing, leave begin
        'left', 'up': move to begin
        'right', 'down': move to end
        'center': center on the size
        'relative': take the same spacing left and right as before the width change
        a number between 0 and 1 (inclusive): calculate the left and right spacing
            according to this number

    """
    oldwidth = end - begin
    if not width:
        width = end - begin
    elif width == 1:
        # take size of monitor, positioning parameter irrelevant:
        return 0, size
    elif 0 < width <= 1:
        width = int(size * width)  # 0.5 takes half the screen
    else:
        raise ValueError('get_new_coordinates_same_monitor, illogical value for width: %s'% width)
    width = max(width, minWidth)

    if positioning is None:
        return begin, begin + width
    if type(positioning) == bytes:
        
        if positioning in ('left', 'up'):
            positioning = 0
        elif positioning in ('right', 'down'):
            positioning = 1
        elif positioning in ('center',):
            positioning = 0.5
        elif positioning in ('relative',):
            # relative to old spacing
            oldspacing = size - oldwidth
            newspacing = max(size - width, 0)
            if oldspacing > 0:
                positioning = 1.0 * begin / oldspacing
            else:
                positioning = 0.5
        else:
            raise ValueError("_get_new_coordinates_same_monitor: invalid positioning option: %s"% positioning)
    if positioning < 0 or positioning > 1:
        raise ValueError("get_new_coordinates_same_monitor: positioning number, should be between 0 and 1 (inclusive), not: %s"% positioning)
    spacing = max((size - width), 0)
    left = int(spacing * positioning)
    return left, left + width

def keepinside_all_screens(left, right, virtL, virtR, fixed_width, border=1):
    """keep inside the virtual screen, allow for border
    
    
    (horizontal and vertical called separate
    """
    if left < virtL - border:
        if fixed_width:
            right += virtL - left - 1
            if right > virtR + border:
                right = virtR + border
        left = virtL - border
    if right > virtR + border:
        if fixed_width:
            left -= right - virtR - 1
            if left < virtL - border:
                left = virtL - border
        right = virtR + border
    return left, right

def keepinside_restore_area(left, right, size, fixed_width=1, margin=1):
    """adjust (restore)
    """
    span = right - left
    if -margin < left < right < size + 1:
        return left, right
    if fixed_width:
        if left < -margin or span > size + margin*2:
            left = -margin
            right = span - margin
        elif right > size + margin:
            right = size + margin
            left = right - span
    else:
        # width may alter, adjust to possibilities of monitor
        if left < -margin or span > size + margin*2:
            left = -margin
            right = min(left+span, size+margin)
        elif right > size + margin:
            right = size + margin
            left = max(right-span, -margin)
        
    return left, right

def keepinside_restore_area(left, right, size, fixed_width=1, margin=1):
    """adjust (restore)
    """
    span = right - left
    if -margin < left < right < size + 1:
        return left, right
    if fixed_width:
        if left < -margin or span > size + margin*2:
            left = -margin
            right = span - margin
        elif right > size + margin:
            right = size + margin
            left = right - span
    else:
        # width may alter, adjust to possibilities of monitor
        if left < -margin or span > size + margin*2:
            left = -margin
            right = min(left+span, size+margin)
        elif right > size + margin:
            right = size + margin
            left = max(right-span, -margin)
        
    return left, right
            

#===================== more monitors: ===============================================
def move_to_monitor(hndle, monitor, currentMon, resize=0):
    """move window with hndle to monitor with hndle monitor
    """
    monitor_info()
    
    toMin = win32con.SW_SHOWMINIMIZED
    toMax = win32con.SW_SHOWMAXIMIZED
    toRestore = win32con.SW_RESTORE
    info = list( win32gui.GetWindowPlacement(hndle) )
    MInew = MONITOR_INFO[monitor]
    MIcur = MONITOR_INFO[currentMon]
    RA = list(info[4])
    maximized =  ( info[1] == win32con.SW_SHOWMAXIMIZED)
    newRA = correct_restore_area(RA, monitor_info_new=MInew,
                                 monitor_info_old=MIcur,
                                 resize=resize,
                                 keepinside=resize)  
    info[1] = toRestore
    if maximized:
        offsetx = MInew['offsetx']
        offsety = MInew['offsety']
        work = MInew['Work']

        # set RA temp to place where maximized window will come as well:
        info[4] = (work[0]-offsetx-BORDERX, work[1]-offsety-BORDERY, work[2]-offsetx+BORDERX, work[3]-offsety+BORDERY)
        info[1] = toRestore
        # set to restore in new maximized coordinates:
        win32gui.SetWindowPlacement(hndle, tuple(info) )
        info[1] = toMax
        info[4] = tuple(newRA)
        # set to max, with new calculated restore parameters:
        win32gui.SetWindowPlacement(hndle, tuple(info) )
    else:
        info[4] = tuple(newRA)
        win32gui.SetWindowPlacement(hndle, tuple(info) )
 
def correct_restore_area(RA, monitor_info_new, monitor_info_old,
                         resize=None, keepinside=None):
    """place RA inside the newWA, preserving relative position as much as possible (other monitor)
    
    input: restore_area (RA) in 'Work' coordinates
           info of monitor_info_new and monitor_info_old
           resize: resize relative to previous if true, keep size if false.
           keepinside: keep the window inside the work area of the monitor (normally when resize = true)
    
    """
    #print 'old RA: %s'% RA
    raWidth = RA[2] - RA[0]
    raHeight = RA[3] - RA[1]
    oldWA = monitor_info_old['Work']
    newWA = monitor_info_new['Work']
    oldMA = monitor_info_old['Monitor']
    newMA = monitor_info_new['Monitor']
    #offsetx_new = monitor_info_new['offsetx']
    #offsety_new = monitor_info_new['offsety']
    #offsetx_old = monitor_info_old['offsetx']
    #offsety_old = monitor_info_old['offsety']
    
    # norm to 0 oriented
    RA[0] -= oldMA[0]
    RA[1] -= oldMA[1]
    RA[2] -= oldMA[0]
    RA[3] -= oldMA[1]
    #print 'old calc area: %s'% RA
    
    newWAWidth = newWA[2] - newWA[0]
    newWAHeight = newWA[3] - newWA[1]
    oldWAWidth = oldWA[2] - oldWA[0]
    oldWAHeight = oldWA[3] - oldWA[1]
    
    if resize:
        RA[0], RA[2] = _get_new_coordinates_resize_other_monitor(begin=RA[0], end=RA[2], oldSize=oldWAWidth,
                                           newSize=newWAWidth)
        RA[1], RA[3] = _get_new_coordinates_resize_other_monitor(begin=RA[1], end=RA[3], oldSize=oldWAHeight,
                                           newSize=newWAHeight)
    else:
        # no resize, keep width, height:
        RA[0], RA[2] = _get_new_coordinates_fixed_other_monitor(begin=RA[0], end=RA[2], oldSize=oldWAWidth,
                                           newSize=newWAWidth)
        RA[1], RA[3] = _get_new_coordinates_fixed_other_monitor(begin=RA[1], end=RA[3], oldSize=oldWAHeight,
                                           newSize=newWAHeight)

    fixed_width = not resize
    RA[0], RA[2] = keepinside_restore_area(RA[0], RA[2], newWAWidth, fixed_width=fixed_width, margin=1)
    RA[1], RA[3] = keepinside_restore_area(RA[1], RA[3], newWAHeight, fixed_width=fixed_width, margin=1)

    # correct for new WA:
    RA[0] += newMA[0]
    RA[1] += newMA[1]
    RA[2] += newMA[0]
    RA[3] += newMA[1]

    #print 'new RA: %s'% RA
    return RA


def _get_new_coordinates_resize_other_monitor(begin, end, oldSize, newSize):
    """sqeeze or make larger the coordinates of the window in either direction
    """
    newBegin = int(begin * newSize / oldSize + 0.5)
    newEnd =  int(end * newSize / oldSize + 0.5)
    return newBegin, newEnd

def _get_new_coordinates_fixed_other_monitor(begin, end, oldSize, newSize):
    """get the coordinates for the window in the new monitor, position at same relative place
    """
    width = end - begin
    # relative to old spacing
    oldspacing = oldSize - width
    newspacing = max(newSize - width, 0)
    if oldspacing > 0:
        newBegin = int(begin * newspacing / oldspacing + 0.5)
        newEnd =  newBegin + width
    else:
        # position center:
        spacing = int((newSize - width)/2)
        newBegin, newEnd = spacing, spacing + width
    return newBegin, newEnd


def is_inside_monitor( point ):
    """give the number of the monitor in which the position lies (starting with 1)

    if outside, return None
    """
    xpos, ypos = point
    monitor_info()
    mon = get_nearest_monitor_point( point )
    if not mon in MONITOR_INFO:
        monitor_info()
    monitor_area = MONITOR_INFO[mon]['Monitor']
    left, top, right, bottom = monitor_area
    if left <= xpos < right and top <= ypos < bottom:
        return mon
    
        
def get_closest_position( point ):
    """return the closest valid mouse position on any monitor
    """
    monitor_info()
    xpos, ypos = point
    if is_inside_monitor( point ):
        return  (xpos, ypos)
    mon = get_nearest_monitor_point( point )
    monitor_area = MONITOR_INFO[mon]['Monitor']
    newPoint = _get_nearest_inside_point( point, monitor_area)
    return newPoint

def _get_nearest_inside_point(point, monitor_area):
    """assume xpos, ypos OUTSIDE monitor_area, get shortest connection
    
    return distance, and xnew, ynew
    """
    xpos, ypos = point
    left, top, right, bottom = monitor_area
    if xpos < left:
        xnew = left
    elif xpos >= right:
        xnew = right - 1
    else:
        xnew = xpos
    if ypos < top:
        ynew = top
    elif ypos >= bottom:
        ynew = bottom - 1
    else:
        ynew = ypos
    #dist = math.sqrt( (xpos-xnew)*(xpos-xnew) + (ypos-ynew)*(ypos-ynew) )
    return (xnew, ynew)

def _get_restore_area(windowHndle):
    """for debugging purposes, get the restore area of window
    """
    info = list( win32gui.GetWindowPlacement(windowHndle) )
    win32gui.SetWindowPlacement(windowHndle, tuple(info) )
    return list(info[4])
            
def _set_restore_area(hndle, left, top, right, bottom):
    """just for testing difficult cases"""
    
    info = list( win32gui.GetWindowPlacement(hndle) )
    info[4] = (left, top, right, bottom)
    win32gui.SetWindowPlacement(hndle, tuple(info) )

#######################            
#various test functions, demo things
#switch on or off at bottom of file:
#for detailed testing of inner functions, see unittestMonitorfunctions.py
#
def test_individual_points():
    """testing of points being inside or near some monitor
    """
    print('for two monitors (my case) some additional calls:')
    print('is inside (one of the ) monitors:')
    print('0, 0, always: ', is_inside_monitor( (0,0) ))
    print('800, 500, mostly inside first: ', is_inside_monitor( (800,500) ))   #first
    print('2000, 600 often inside second:', is_inside_monitor( (2000,600) ))
    print('-10, -1000 outside: ', is_inside_monitor( (-10, -1000) ))
    print('10000, 5000 outside: ', is_inside_monitor( (10000, 5000) ))
    print('guessing nearest position if outside:')
    print('point inside: ', get_closest_position( (0,0) ))
    print('point inside: ', get_closest_position( (500,500) ))
    print('left of monitor 1 (-10, 200): ', get_closest_position( (-10, 200) ))
    print('down of monitor 2 (1800, 1000): ', get_closest_position( (1800, 1000) ))
    print('big guess of (10000, -5000)', get_closest_position( (10000, -5000) ))

def test_get_nearest_monitors():
    """test getting the nearest monitor, wrappers around API functions
    """

    mon = get_nearest_monitor_point( (500, 500) )
    print('nearest monitor point: %s'% mon)
    winHndle = win32gui.GetForegroundWindow()
    print('window hndle: %s'% winHndle)
    mon = get_nearest_monitor_window(winHndle)
    print('nearest monitor: %s'% mon)
    canBeResized = window_can_be_resized(winHndle)
    print('window can be resized: %s'% canBeResized)
    

def test_basic_values():
    """testing the system metrics and monitor info of multiple monitors
    """
    monitor_info()
    print('---number of monitors: %s'% NMON)
    print('---total virtual screen: %s'% VIRTUAL_SCREEN)
    print('---border of windows: %s, %s'%  (BORDERX, BORDERY))
    print('---MONITOR_HNDLES: %s'% MONITOR_HNDLES)
    for mon in MONITOR_HNDLES:
        print('----monitor: %s:'% mon)
        pprint.pprint(MONITOR_INFO[mon])
    print('----')   

defaultSleepTime = 1
def wait(mult=1):
    """wait multiple of default wait time
    """
    time.sleep(defaultSleepTime*mult)

def test_mouse_to_other_monitor():
    """try to position the mouse in another monitor
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    canBeResized = window_can_be_resized(winHndle)
    mon = get_nearest_monitor_window(winHndle)
    others = get_other_monitors(mon)
    moninfo = MONITOR_INFO[others[0]]
    mousefour = moninfo['Monitor']
    x, y = mousefour[0] + int((mousefour[2] - mousefour[0])/2.0), mousefour[1] + int((mousefour[3] - mousefour[1])/2.0)
    print('middle of other monitor:', x, y)
    
    
def test_move_to_other_monitor():
    """try moving to other monitor, keeping position as close as possible
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    canBeResized = window_can_be_resized(winHndle)
    mon = get_nearest_monitor_window(winHndle)
    others = get_other_monitors(mon)
    print('other monitors: %s'% others)
    restore_window(winHndle)
    wait()
    otherMon = others[0]
    move_to_monitor(winHndle, otherMon, mon, resize=canBeResized)
    wait()
    move_to_monitor(winHndle, mon, otherMon, resize=canBeResized)
    wait()
    move_to_monitor(winHndle, otherMon, mon, resize=0)
    wait()
    move_to_monitor(winHndle, mon, otherMon, resize=0)
    wait()
    maximize_window(winHndle)
    move_to_monitor(winHndle, otherMon, mon, resize=canBeResized)
    wait()
    move_to_monitor(winHndle, mon, otherMon, resize=canBeResized)
 


def test_move_around_monitor_fixed_size():
    """test different parameters of xwidth and ywidth for positioning a window
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='left', ypos='top')
    wait()
    restore_window(winHndle, mon, xpos='right', ypos='top')
    wait()
    restore_window(winHndle, mon, xpos='right', ypos='bottom')
    wait()
    restore_window(winHndle, mon, xpos='left', ypos='bottom')
    wait()
    restore_window(winHndle, mon, xpos=0.2, ypos=0.2) # near left, top
    wait()
    restore_window(winHndle, mon, xpos=0.8, ypos=0.2) # near right top etc. 
    wait()
    restore_window(winHndle, mon, xpos=0.8, ypos=0.8)
    wait()
    restore_window(winHndle, mon, xpos=0.2, ypos=0.8)
    wait()
    restore_window(winHndle, mon, xpos=0, ypos=0)  # identical left, top
    wait()
    restore_window(winHndle, mon, xpos=1, ypos=1)  # identical right, bottom
    wait()
    restore_window(winHndle, mon, xpos=0.5, ypos=0.5) # identical center, center
    wait()
    
def test_center_different_sizes():
    """test different parameters of xwidth and ywidth for different sizes of window
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.4, ywidth=0.4, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.3, ywidth=0.3, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.3, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.7, ywidth=0.3, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.7, ywidth=0.5, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.7, ywidth=0.7, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.9, ywidth=0.9, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.995, ywidth=0.995, keepinside=1)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()


def test_default_restore():
    """test restore after having set the RA to ridiculous values
    
    mon can be left away.
    keepinside makes the window fit to the monitor always.
    note the window is kept on the same monitor, though the _set_restore_area could
    suspect different.
    
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    _set_restore_area(winHndle, -1700, -1000, 1000, 5000)
    mon = get_nearest_monitor_window(winHndle)
    print('monitor: %s'% mon)
    restore_window(winHndle, mon)
    wait()
    _set_restore_area(winHndle, -1700, -1000, 1000, 5000)
    restore_window(winHndle, mon, keepinside=1)
    wait()

    _set_restore_area(winHndle, 500, -1000, 4700, 5000)
    mon = get_nearest_monitor_window(winHndle)
    print('monitor: %s'% mon)
    restore_window(winHndle)
    wait()
    _set_restore_area(winHndle, 500, -1000, 4700, 5000)
    restore_window(winHndle, keepinside=1)
    wait()
    maximize_window(winHndle)

def test_move_far_away():
    """test moving across borders of monitor, testing keepinside_all
    
    going to corners can give unexpected results, jumping to another monitor...
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)
    canBeResized = window_can_be_resized(winHndle)
    # you can also test from the other monitor by uncommenting next 3 lines:
    #others = get_other_monitors(mon)
    #move_to_monitor(winHndle, others[0], mon, resize=canBeResized)
    #mon = get_nearest_monitor_window(winHndle)

    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'left', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'right', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'top', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'down', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'lefttop', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'rightbottom', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'leftbottom', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'righttop', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    
def test_stretch_far_away():
    """test stretching across borders of monitor, testing keepinside_all
    
    stretching to corners can give unexpected results, jumping to another monitor...
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)
    canBeResized = window_can_be_resized(winHndle)
    # you can also test from the other monitor by uncommenting next 3 lines:
    #others = get_other_monitors(mon)
    #move_to_monitor(winHndle, others[0], mon, resize=canBeResized)
    #mon = get_nearest_monitor_window(winHndle)

    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'left', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'right', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'top', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'down', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'lefttop', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'rightbottom', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'leftbottom', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'righttop', 5000)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()




def test_move_same_window_sides_corners():
    """test moving around the monitor
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)

    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'left', 50)
    wait()
    move_window(winHndle, 'up', 50)
    wait()
    move_window(winHndle, 'right', 50)
    wait()
    move_window(winHndle, 'down', 50)
    wait()

    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'left', 0.9, units='relative')
    wait()
    move_window(winHndle, 'up', 0.9, units='relative')
    wait()
    move_window(winHndle, 'right', 0.9, units='relative')
    wait()
    move_window(winHndle, 'down', 0.9, units='relative')
    wait()

    # to corners:
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'lefttop', 1, units='relative')
    wait()
    move_window(winHndle, 'righttop', 1, units='relative')
    wait()
    move_window(winHndle, 'rightbottom', 1, units='relative')
    wait()
    move_window(winHndle, 'leftbottom', 1, units='relative')
    wait()

    # towards corners:
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, 'lefttop', 0.5, units='relative')
    wait()
    move_window(winHndle, 'righttop', 0.5, units='relative')
    wait()
    move_window(winHndle, 'rightbottom', 0.5, units='relative')
    wait()
    move_window(winHndle, 'leftbottom', 0.5, units='relative')
    wait()
    maximize_window(winHndle)

def test_move_same_window_degrees():
    """test moving around the monitor with an  angle
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)


    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    for i in range(75):
        dist = 50 + i
        angle = i*20 % 360
        move_window(winHndle,angle, dist)
        wait(0.1)
    wait(2)
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    for i in range(75):
        dist = 50 + i*3
        angle = i*20 % 360
        move_window(winHndle,angle, dist, keepinside=1)
        wait(0.1)

def test_move_relative_degrees():
    """test moving around the monitor with an  angle relative to the current distance of the nearest corner
    
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)
    
    # when outside:
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    move_window(winHndle, amount=0.9, units='relative', direction='up')
    wait()
    move_window(winHndle, amount=50, direction='up')
    wait()
    move_window(winHndle, 45, 0.5, units='relative')
    
    return
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    for i in range(20):
        dist = 0.8
        angle = i*20 % 360
        move_window(winHndle,angle, dist, units='relative')
        wait()
        restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
        wait(0.5)
    
    # range with greatest jumps (probably)   
    for i in range(20):
        dist = 1
        angle = i + 48
        move_window(winHndle,angle, dist, units='relative')
        wait(0.3)
        restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
        wait(0.1)


    #wait(2)
    #restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    #for i in range(75):
    #    dist = 50 + i*3
    #    angle = i*20 % 360
    #    move_window(winHndle,angle, dist, keepinside=1)
    #    wait(0.1)
def test_restore_window_half_width():
    """test moving around edges, corners, half width, height
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)

    restore_window(winHndle, xpos='left', ypos=None, xwidth=0.5, ywidth=1, keepinside=1)
    wait()
    restore_window(winHndle, xpos=None, ypos='top', xwidth=1, ywidth=0.5, keepinside=1)
    wait()
    restore_window(winHndle, xpos='right', ypos=None, xwidth=0.5, ywidth=1, keepinside=1)
    wait()
    restore_window(winHndle, xpos=None, ypos='down', xwidth=1, ywidth=0.5, keepinside=1)
    wait()
    
    # corners
    restore_window(winHndle, xpos='left', ypos='top', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    restore_window(winHndle, xpos='right', ypos='top', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    restore_window(winHndle, xpos='right', ypos='bottom', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    restore_window(winHndle, xpos='left', ypos='down', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    
    
    maximize_window(winHndle)

def test_restore_window_third_width():
    """test moving around edges, corners, half width, height
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)

    restore_window(winHndle, xpos='left', ypos=None, xwidth=0.33, ywidth=1, keepinside=1)
    wait()
    restore_window(winHndle, xpos=None, ypos='top', xwidth=1, ywidth=0.33, keepinside=1)
    wait()
    restore_window(winHndle, xpos='right', ypos=None, xwidth=0.33, ywidth=1, keepinside=1)
    wait()
    restore_window(winHndle, xpos=None, ypos='down', xwidth=1, ywidth=0.33, keepinside=1)
    wait()
    maximize_window(winHndle)
  

def test_stretch_same_window_sides_corners():
    """test moving around the monitor
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)

    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'left', 50)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'up', 50)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'right', 50)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'down', 50)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    # now stretch in all directions, adding to the size:
    stretch_window(winHndle, 'left', 0.9, units='relative')
    wait()
    stretch_window(winHndle, 'up', 0.9, units='relative')
    wait()
    stretch_window(winHndle, 'right', 0.9, units='relative')
    wait()
    stretch_window(winHndle, 'down', 0.9, units='relative')
    wait()

    # to corners:
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'lefttop', 1, units='relative')
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'righttop', 1, units='relative')
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'rightbottom', 1, units='relative')
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'leftbottom', 1, units='relative')
    wait()

    # towards corners:
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'lefttop', 0.5, units='relative')
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'righttop', 0.5, units='relative')
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'rightbottom', 0.5, units='relative')
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.5, ywidth=0.5, keepinside=1)
    wait()
    stretch_window(winHndle, 'leftbottom', 0.5, units='relative')
    wait()
    maximize_window(winHndle)

def test_shrink_same_window_sides_corners():
    """test moving around the monitor
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    mon = get_nearest_monitor_window(winHndle)

    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    shrink_window(winHndle, 'left', 50)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    shrink_window(winHndle, 'up', 50)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    shrink_window(winHndle, 'right', 50)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    shrink_window(winHndle, 'down', 50)
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    # now shrink in all directions, adding to the size:
    shrink_window(winHndle, 'left', 0.5, units='relative')
    wait()
    shrink_window(winHndle, 'up', 0.5, units='relative')
    wait()
    shrink_window(winHndle, 'right', 0.5, units='relative')
    wait()
    shrink_window(winHndle, 'down', 0.5, units='relative')
    wait()
    # two times around:
    shrink_window(winHndle, 'left', 0.5, units='relative')
    wait()
    shrink_window(winHndle, 'up', 0.5, units='relative')
    wait()
    shrink_window(winHndle, 'right', 0.5, units='relative')
    wait()
    shrink_window(winHndle, 'down', 0.5, units='relative')
    wait()

    # away from corners, same distance as the corner distance:
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    shrink_window(winHndle, 'lefttop', 1, units='relative')
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    shrink_window(winHndle, 'righttop', 1, units='relative')
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    shrink_window(winHndle, 'rightbottom', 1, units='relative')
    wait()
    restore_window(winHndle, mon, xpos='center', ypos='center', xwidth=0.8, ywidth=0.8, keepinside=1)
    wait()
    shrink_window(winHndle, 'leftbottom', 1, units='relative')
    wait()

    maximize_window(winHndle)


def test_minimize_maximize_restore():
    """testing minimize/maximize actions of window on one screen
    """
    monitor_info()
    winHndle = win32gui.GetForegroundWindow()
    maximize_window(winHndle)
    wait()
    minimize_window(winHndle)
    wait()
    maximize_window(winHndle)
 
def test_taskbar_position():
    print('taskbar located: %s'% get_taskbar_position())
    
def do_doctest():
    import doctest
    doctest.testmod()
   
if __name__ == "__main__":
    winHndle = win32gui.GetForegroundWindow()
    #print 'ra: %s'%  _get_restore_area(winHndle)   
    test_basic_values()
    test_taskbar_position()
    test_mouse_to_other_monitor()
    #restore_window(winHndle)
    #test_get_nearest_monitors()
    #test_individual_points()
    
    #test_move_to_other_monitor()
    #test_move_around_monitor_fixed_size()
    #test_center_different_sizes()
    #test_default_restore()
    #test_minimize_maximize_restore()
    #test_move_same_window_sides_corners()
    #test_stretch_same_window_sides_corners()
    #test_shrink_same_window_sides_corners()
    #test_move_same_window_degrees()
    #test_move_relative_degrees()
    #test_move_far_away()
    #test_stretch_far_away()
    #test_restore_window_half_width()
    #test_restore_window_third_width()
    #do_doctest()  # most testing moved to unittestMonitorfunctions.py...