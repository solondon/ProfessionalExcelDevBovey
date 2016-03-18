VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} dsrConnect 
   ClientHeight    =   7980
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   9180
   _ExtentX        =   16193
   _ExtentY        =   14076
   _Version        =   393216
   Description     =   "A VB6 COM Addin that adds a command bar to Excel containing a button for each of the Paste Special options."
   DisplayName     =   "Paste Special Bar - VB6"
   AppName         =   "Microsoft Excel"
   AppVer          =   "Microsoft Excel 9.0"
   LoadName        =   "Startup"
   LoadBehavior    =   3
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel"
End
Attribute VB_Name = "dsrConnect"
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
' Ch23 - COM Addins use an Addin Designer to handle the connection to Excel, which
'        uses the OnConnection event instead of Auto_Open or Workbook_Open and the
'        OnDisconnection event instead of Auto_Close or Workbook_Close.
'
Option Explicit

' **************************************************************************
' Module Constant Declarations Follow
' **************************************************************************
Private Const msMODULE As String = "dsrConnect"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   This routine is run every time the application is opened.
'             It handles initialization of the gxlapp.
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 30 Apr 08   Rob Bovey       Created
' 06 Jun 08   Rob Bovey       Moved code from Auto_Open in MOpenClose to here
'
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    Const sSOURCE As String = "AddinInstance_OnConnection"
    
    Dim bOpenFailed As Boolean
    
    On Error GoTo ErrorHandler
    
    Set gxlApp = Application
    
    ' Initialize global variables.
    InitGlobals

    ' Build the custom commandbars
    If Not bBuildCommandBars() Then Err.Raise glHANDLED_ERROR
    
    'Set the commandbars position
    SetCommandBarPosition gxlApp.CommandBars(gsMENU_NAME)
    
    ' Instantiate the control event handler class variable.
    Set gclsControlEvents = New CControlEvents
    
ErrorExit:

    ' Reset critical application properties.
    If bOpenFailed Then ShutdownApplication
    Exit Sub
    
ErrorHandler:
    If Err.Number <> glHANDLED_ERROR Then Err.Description = Err.Description & " (" & sSOURCE & ")"
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        bOpenFailed = True
        Resume ErrorExit
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   This routine runs automatically every time the application
'             workbook is closed. If application shutdown code is not already
'             running (i.e. via a call from the Exit menu), this procedure
'             calls the application shutdown procedure.
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 30 Apr 08   Rob Bovey       Created
' 06 Jun 08   Rob Bovey       Moved code from Auto_Close in MOpenClose to here
'
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    ' Call standard shutdown code if it isn't already running.
    If Not gbShutdownInProgress Then
        ShutdownApplication
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   This routine shuts down the gxlapp.
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 30 Apr 08   Rob Bovey       Created
' 06 Jun 08   Rob Bovey       Deleted line to close the workbook
'
Public Sub ShutdownApplication()

    On Error Resume Next
    
    ' This flag prevents this routine from being called a second time
    ' by Auto_Close if has already been called by another procedure.
    gbShutdownInProgress = True

    Set gclsControlEvents = Nothing
    
    'Store the commandbar position in the registry
    StoreCommandBarPosition gxlApp.CommandBars(gsMENU_NAME)
    
    ResetCommandBars
    
End Sub

