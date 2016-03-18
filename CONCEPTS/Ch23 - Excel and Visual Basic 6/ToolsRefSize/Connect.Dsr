VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   8775
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   8190
   _ExtentX        =   14446
   _ExtentY        =   15478
   _Version        =   393216
   Description     =   "COM Addin to enlarge the size of the labels on the bottom of the Tools > References dialog."
   DisplayName     =   "Tools Reference Resizer"
   AppName         =   "Visual Basic for Applications IDE"
   AppVer          =   "6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0"
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Description:    This module contains the startup and shutdown code.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit
Option Compare Text

'A variable to hook the Tools > References menu item
Dim WithEvents btnToolsRefs As Office.CommandBarButton
Attribute btnToolsRefs.VB_VarHelpID = -1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   This routine is run every time the application is opened.
'             It handles initialization of the gvbeapp and sets up the menu event hook
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 30 Apr 08   Stephen Bullen  Created
'
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    'Store away the Excel application
    Set gvbeapp = Application

    'Hook the builtin "Tools > References" button
    Set btnToolsRefs = gvbeapp.CommandBars.FindControl(Id:=942)

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Handle the user clicking the Tools > References button.
'             Starts a windows timer that waits for the dialog to show up
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 30 Apr 08   Stephen Bullen  Created
'
Private Sub btnToolsRefs_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

    'Someone clicked Tools > References, so start a timer proc to wait for it to open
    SetTimerProc

End Sub


