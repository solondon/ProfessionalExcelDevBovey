VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHelloWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit

' ************************************************************
' Class Constant Declarations Follow
' ************************************************************
' SetWindowLongA API constant.
Private Const GWL_HWNDPARENT As Long = -8


' ************************************************************
' Class Variable Declarations Follow
' ************************************************************
' Object reference to the calling Excel Application.
Private mxlApp As Excel.Application
' Window handle of the calling Excel Application.
Private mlXLhWnd As Long


' ************************************************************
' Class DLL Declaractions Follow
' ************************************************************
Private Declare Function FindWindowA Lib "user32" _
                        (ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As Long
Private Declare Function SetWindowLongA Lib "user32" _
                        (ByVal hWnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long


' ************************************************************
' Class Property Procedures Follow
' ************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This property procedure is used to pass a reference
'           to the calling Excel Application object into the DLL.
'
' Arguments:    xlApp           A reference to the calling Excel
'                               Application object.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Property Set ExcelApp(ByRef xlApp As Excel.Application)
    Set mxlApp = xlApp
    ' Get the window handle of the Excel Application object
    ' as soon as it is passed to us.
    If Val(Application.Version) > 9 Then
        ' Excel XP and higher expose the window handle directly
        mlXLhWnd = Application.hWnd
    Else
        ' Excel 2000 requres an API call for the window handle.
        mlXLhWnd = FindWindowA(vbNullString, mxlApp.Caption)
    End If
End Property


' ************************************************************
' Class Event Procedures Follow
' ************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This event procedure destroys the object variable
'           holding the reference to the calling Excel
'           Application object when the class is destroyed.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Private Sub Class_Terminate()
    Set mxlApp = Nothing
End Sub


' ************************************************************
' Class Method Procedures Follow
' ************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This method procedure displays a message box in the
'           calling Excel application.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Sub ShowMessage()
    MsgBox "Hello World!"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This method procedure enters a text string into the
'           active cell of the calling Excel application.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Sub WriteMessage()
    mxlApp.ActiveCell.Value = "Hello World!"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This method procedure displays a VB6 form as a child
'           window of the calling Excel Application object.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Sub ShowVB6Form()
    Dim frmHelloWorld As FHelloWorld
    Set frmHelloWorld = New FHelloWorld
    Load frmHelloWorld
    ' Parent the Form window to the Excel Application window.
    SetWindowLongA frmHelloWorld.hWnd, GWL_HWNDPARENT, mlXLhWnd
    frmHelloWorld.Show vbModal
    Unload frmHelloWorld
    Set frmHelloWorld = Nothing
End Sub
