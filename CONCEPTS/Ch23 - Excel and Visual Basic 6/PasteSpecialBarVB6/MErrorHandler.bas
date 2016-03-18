Attribute VB_Name = "MErrorHandler"
'
' Description:    This module contains the central error
'                 handler and related constant declarations.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
' Ch23 - The only changes required to modify this for use in a COM Addin were
'        to use App.Title and App.Path (for writing the error log) instead of
'        ThisWorkbook
'
Option Explicit
Option Private Module

Public Const gbDEBUG_MODE As Boolean = False
Public Const glHANDLED_ERROR As Long = 9999
Public Const glUSER_CANCEL As Long = 18

Private Const msSILENT_ERROR As String = "UserCancel"
Private Const msFILE_ERROR_LOG As String = "Error.log"


Public Function bCentralErrorHandler( _
            ByVal sModule As String, _
            ByVal sProc As String, _
            Optional ByVal sFile As String, _
            Optional ByVal bEntryPoint As Boolean) As Boolean

    Static sErrMsg As String
    
    Dim iFile As Integer
    Dim lErrNum As Long
    Dim sFullSource As String
    Dim sPath As String
    Dim sLogText As String
    
    ' Grab the error info before it's cleared by
    ' On Error Resume Next below.
    lErrNum = Err.Number
    ' If this is a user cancel, set the silent error flag
    ' message. This will cause the error to be ignored.
    If lErrNum = glUSER_CANCEL Then sErrMsg = msSILENT_ERROR
    ' If this is the originating error, the static error
    ' message variable will be empty. In that case, store
    ' the originating error message in the static variable.
    If Len(sErrMsg) = 0 Then sErrMsg = Err.Description

    ' We cannot allow errors in the central error handler.
    On Error Resume Next
    
    ' Load the default filename if required.
    'Ch23 - use the dll's name instead of the workbook's
    If Len(sFile) = 0 Then sFile = App.EXEName
    
    ' Get the application directory.
    'Ch23 - use the dll's path instead of the workbook's
    sPath = App.Path
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    ' Construct the fully-qualified error source name.
    sFullSource = "[" & sFile & "]" & sModule & "." & sProc

    ' Create the error text to be logged.
    sLogText = "  " & sFullSource & ", Error " & _
                        CStr(lErrNum) & ": " & sErrMsg
    
    ' Open the log file, write out the error information and
    ' close the log file.
    iFile = FreeFile()
    Open sPath & msFILE_ERROR_LOG For Append As #iFile
    Print #iFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
    If bEntryPoint Then Print #iFile,
    Close #iFile
        
    ' Do not display or debug silent errors.
    If sErrMsg <> msSILENT_ERROR Then
    
        ' Show the error message when we reach the entry point
        ' procedure or immediately if we are in debug mode.
        If bEntryPoint Or gbDEBUG_MODE Then
            gxlApp.ScreenUpdating = True
            MsgBox sErrMsg, vbCritical, gsAPP_TITLE
            ' Clear the static error message variable once
            ' we've reached the entry point so that we're ready
            ' to handle the next error.
            sErrMsg = vbNullString
        End If
        
        ' The return value is the debug mode status.
        bCentralErrorHandler = gbDEBUG_MODE
        
    Else
        ' If this is a silent error, clear the static error
        ' message variable when we reach the entry point.
        If bEntryPoint Then sErrMsg = vbNullString
        bCentralErrorHandler = False
    End If
    
End Function
