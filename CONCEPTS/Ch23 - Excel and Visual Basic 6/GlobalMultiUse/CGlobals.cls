VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit

' ************************************************************
' Class Type Declarations Follow
' ************************************************************
Public Type MY_UDT
    lVar1 As Long
    sVar2 As String
    vVar3 As Variant
End Type


' ************************************************************
' Class Enum Declarations Follow
' ************************************************************
Public Enum mpDemo
    mpDemoEnum1 = 1
    mpDemoEnum2
    mpDemoEnum3
End Enum


' ************************************************************
' Class Method Procedures Follow
' ************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This method procedure determines if the specified
'           worksheet exists in the specified workbook.
'
' Arguments:    wkbBook         The workbook in which to look
'                               for the specified worksheet name.
'               szSheetName     The name of the worksheet to
'                               look for.
'
' Returns:      Boolean         True if the specified worksheet
'                               was located in the specified
'                               workbook, False otherwise.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Function bSheetExists(ByRef wkbBook As Excel.Workbook, _
                        ByRef szSheetName As String) As Boolean
    On Error Resume Next
    bSheetExists = Not (wkbBook.Sheets(szSheetName) Is Nothing)
End Function

