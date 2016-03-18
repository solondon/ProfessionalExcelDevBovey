Attribute VB_Name = "MSystemCode"
'
' Description:  Contains support routines developed specifically for this application.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Determines whether Word is available for automation
'           on the computer.
'
' Arguments:    None
'
' Returns:      Boolean         True if Word is available or
'                               False if it is not.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Function bWordAvailable() As Boolean
    Dim wdApp As Object
    ' Attempt to start an instance of Word.
    On Error Resume Next
        Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    ' Return the result of the test.
    If Not wdApp Is Nothing Then
        ' If we started Word we need to close it.
        wdApp.Quit
        Set wdApp = Nothing
        bWordAvailable = True
    Else
        bWordAvailable = False
    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Determines whether Outlook is available for
'           automation on the computer.
'
' Arguments:    None
'
' Returns:      Boolean         True if Outlook is available or
'                               False if it is not.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Function bOutlookAvailable() As Boolean
    Dim bWasRunning As Boolean
    Dim olApp As Object
    On Error Resume Next
        ' Attempt to get a reference to a currently open
        ' instance of Outlook.
        Set olApp = GetObject(, "Outlook.Application")
        If olApp Is Nothing Then
            ' If this fails, attempt to start a new instance.
            Set olApp = CreateObject("Outlook.Application")
        Else
            ' Otherwise flag that Outlook was already running
            ' so that we don't try to close it.
            bWasRunning = True
        End If
    On Error GoTo 0
    ' Return the result of the test.
    If Not olApp Is Nothing Then
        ' If we started Outlook we need to close it.
        If Not bWasRunning Then olApp.Quit
        Set olApp = Nothing
        bOutlookAvailable = True
    Else
        bOutlookAvailable = False
    End If
End Function

