Attribute VB_Name = "MMain"
'
' Description:    Contains the API calls to handle the Windows timer
'                 and perform the resizing
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit
Option Compare Text

''''''''''''''''''''''''''''''''''''''''''''''''''
' Module-level Declarations Follow
''''''''''''''''''''''''''''''''''''''''''''''''''
'The Excel application we're given at startup
Public gvbeapp As VBIDE.VBE

'The Windows timer ID
Dim miTimerID As Long
Dim mdStart As Double

''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows API Declarations, Constants and Types Follow
''''''''''''''''''''''''''''''''''''''''''''''''''

'Create an API-based OnTime event, so we can do stuff after the menu item has finished
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

'Cancel an OnTime event
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

'Functions used in FindOurWindow
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long

'API functions to find and move the windows
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'API functions to change the label style
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)

Private Const SS_LEFTNOWORDWRAP = &HC&

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Creates a Windows Timer procedure, to fire every 0.2 seconds.
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 30 Apr 08   Stephen Bullen  Created
'
Public Sub SetTimerProc()
    If miTimerID > 0 Then KillTimer 0, miTimerID
    
    miTimerID = SetTimer(0, 0, 200, AddressOf TimerCallback)
    mdStart = Now()
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Called by the Windows Timer routine.
'             Looks for the 'Tools > References' dialog and changes the text boxes
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 30 Apr 08   Stephen Bullen  Created
'
Public Sub TimerCallback(ByVal hWnd As Long, ByVal lngMsg As Long, ByVal lngID As Long, ByVal lngTime As Long)

    Dim lhWnd As Long
    Dim hWndDesktop As Long
    Dim lLabel As Long
    Dim uLabelPos As RECT
    Dim lNudge As Long
    Dim lStyle As Long
    
    'Auto-kill after 10 seconds
    If Now() - mdStart > TimeValue("00:00:10") Then
        KillTimer 0, miTimerID
        miTimerID = 0
        Exit Sub
    End If
    
    'Find the Tools > References dialog
    lhWnd = FindOurWindow("#32770", "References - " & gvbeapp.ActiveVBProject.Name)
    
    If lhWnd > 0 Then
    
        'Found it, so kill the Timer to avoid re-entry
        KillTimer 0, miTimerID
        miTimerID = 0

        'Find the "Location" label
        lLabel = FindWindowEx(lhWnd, 0&, "Static", "Location:")
        
        'Get its position and size (in client coordinates)
        GetWindowPosition lhWnd, lLabel, uLabelPos
        
        'How big is a label?
        lNudge = uLabelPos.Bottom - uLabelPos.Top
        
        'And move it up a bit
        MoveWindow lLabel, uLabelPos.Left, uLabelPos.Top - lNudge / 2, uLabelPos.Right - uLabelPos.Left, lNudge, True

        'The next label is the path, so find that and resize it
        lLabel = FindWindowEx(lhWnd, lLabel, "Static", vbNullString)
        
        'Get its position and size
        GetWindowPosition lhWnd, lLabel, uLabelPos
        
        'And move it, setting it to double-height
        MoveWindow lLabel, uLabelPos.Left, uLabelPos.Top - lNudge / 2, uLabelPos.Right - uLabelPos.Left, lNudge * 2, True
        
        'And allow it to wrap with a vertical scrollbar
        lStyle = GetWindowLong(lLabel, GWL_STYLE)
        lStyle = lStyle And Not SS_LEFTNOWORDWRAP
        SetWindowLong lLabel, GWL_STYLE, lStyle
        
        'The next is "Language", so move down a bit
        lLabel = FindWindowEx(lhWnd, lLabel, "Static", vbNullString)
        
        'Get its position and size
        GetWindowPosition lhWnd, lLabel, uLabelPos
        
        'And move it
        MoveWindow lLabel, uLabelPos.Left, uLabelPos.Top + lNudge / 2, uLabelPos.Right - uLabelPos.Left, lNudge, True
        
        'And lastly, the "Language result"
        lLabel = FindWindowEx(lhWnd, lLabel, "Static", vbNullString)
        
        'Get its position and size
        GetWindowPosition lhWnd, lLabel, uLabelPos
        
        'And move it
        MoveWindow lLabel, uLabelPos.Left, uLabelPos.Top + lNudge / 2, uLabelPos.Right - uLabelPos.Left, lNudge, True
    End If
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Gets the position of a label on a dialog, in screen coordinates
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 30 Apr 08   Stephen Bullen  Created
'
Private Sub GetWindowPosition(lhWnd As Long, lLabel As Long, uLabelPos As RECT)

    Dim uPt As POINTAPI
    
    'Get the label's position and size
    GetWindowRect lLabel, uLabelPos

    'Copy left/top to a POINT structure
    uPt.x = uLabelPos.Left
    uPt.y = uLabelPos.Top
    
    'Convert to client coords
    ScreenToClient lhWnd, uPt
    
    'And update the RECT
    uLabelPos.Left = uPt.x
    uLabelPos.Top = uPt.y

    'Do same to right/bottom
    uPt.x = uLabelPos.Right
    uPt.y = uLabelPos.Bottom
    
    'Convert to client coords
    ScreenToClient lhWnd, uPt
    
    'And update the RECT
    uLabelPos.Right = uPt.x
    uLabelPos.Bottom = uPt.y

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Finds a top-level window of the given class and
'           caption that belongs to this instance of Excel,
'           by matching the process IDs
'
' Arguments:    sClass      The window class name to look for
'               sCaption    The window caption to look for
'
' Returns:      Long        The handle of Excel's main window
'
' Date          Developer       Action
' --------------------------------------------------------------
' 30 Apr 08   Stephen Bullen  Created
'
Function FindOurWindow(Optional ByVal sClass As String = vbNullString, Optional ByVal sCaption As String = vbNullString)

  Dim hWndDesktop As Long
  Dim hWnd As Long
  Dim hProcThis As Long
  Dim hProcWindow As Long

  'All top-level windows are children of the desktop,
  'so get that handle first
  hWndDesktop = GetDesktopWindow

  'Get the ID of this instance of Excel, to match
  hProcThis = GetCurrentProcessId

  Do
    'Find the next child window of the desktop that
    'matches the given window class and/or caption.
    'The first time in, hWnd will be zero, so we'll get
    'the first matching window.  Each call will pass the
    'handle of the window we found the last time, thereby
    'getting the next one (if any)
    hWnd = FindWindowEx(hWndDesktop, hWnd, sClass, sCaption)

    'Get the ID of the process that owns the window we found
    GetWindowThreadProcessId hWnd, hProcWindow

    'Loop until the window's process matches this process,
    'or we didn't find the window
  Loop Until hProcWindow = hProcThis Or hWnd = 0

  'Return the handle we found
  FindOurWindow = hWnd

End Function

