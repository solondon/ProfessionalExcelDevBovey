VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDialogHandler"
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
    mlXLhWnd = FindWindowA(vbNullString, mxlApp.Caption)
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
' Comments: This method procedure is used to display the VB6
'           form as a child window of the calling Excel
'           Application object. It also returns the option and
'           value selections made in the form by the user to
'           the calling procedure.
'
' Arguments:    sOption         Returned by this procedure.
'                               The name of the option selected
'                               by the user.
'               dValue          Returned by this procedure.
'                               The value selected by the user.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Sub ShowVB6Form(ByRef sOption As String, _
                                        ByRef dValue As Double)
    Dim frmCtrlArrays As FControlArrays
    Set frmCtrlArrays = New FControlArrays
    Load frmCtrlArrays
    ' Parent the Form window to the Excel Application window.
    SetWindowLongA frmCtrlArrays.hWnd, GWL_HWNDPARENT, mlXLhWnd
    frmCtrlArrays.Show vbModal
    sOption = frmCtrlArrays.OptionSelected
    dValue = frmCtrlArrays.ListSelection
    Unload frmCtrlArrays
    Set frmCtrlArrays = Nothing
End Sub


