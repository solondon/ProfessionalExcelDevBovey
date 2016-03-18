VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Class to perform centralised performance monitoring
'
'   To add to another VB Project, copy this class and the Globals module to that project.
'   The class requires a reference to the Microsoft Scripting Runtime
'
'   Version Date        Author          Comment
'   0.0.1   01-10-2003  Stephen Bullen  Initial Version
'
Option Compare Binary
Option Explicit

'UDT to store the timing information for a procedure, used for the monitoring of individual procedures
Type ProcMonitor
    ProcID As String                                  'The unique procedure ID, usually "ProjectName.ModuleName.ProcedureName"
    Start As Currency                                 'The start time (set to zero once processed)
    Duration As Currency                              'The duration for that procedure
End Type

'UDT to store the timing results
Type ProcResults
    ProcID As String
    Count As Long
    TotalDuration As Currency
    MaxDuration As Currency
End Type

'API calls to access the Windows high-performance timers
Private Declare Function QueryFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long
Private Declare Function QueryCounter Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long

'The overall start and end times of the monitoring pass
Dim mcStart As Currency, mcEnd As Currency
Attribute mcEnd.VB_VarUserMemId = 1073938436


'Called by each routine to notify the performance monitor when it starts
'Routine logs the routine ID (Project.Module.Procedure) and start time in an array
Public Sub PerfMonProcStart(sID As String)

    If Not pbMonitoring Then Exit Sub

    'Next item in the array, making sure we have enough items
    plProcIdx = plProcIdx + 1
    If plProcIdx > UBound(pauProcMonitor) Then ReDim Preserve pauProcMonitor(1 To plProcIdx + 100)

    'Log the proc start info
    With pauProcMonitor(plProcIdx)
        .ProcID = sID
        QueryCounter .Start
        .Duration = 0
    End With

End Sub


'Called by each routine to notify the performance monitor when it ends
'Routine scans back through the array, adding up the durations of each lower-level routine
'until it finds the proc-start entry for the given ID.
'It then works out the duration spent within the routine (= Stop time - Start time - lower-level durations)
'It saves this information to a ProcResults UDT stored in a Dictionary object, logging the ID, number of calls, total
'duration and the maximum duration of a single call.
Public Sub PerfMonProcEnd(sID As String)

    Dim cDur As Currency, uProcResults As ProcResults
    Dim cEnd As Currency

    If Not pbMonitoring Then Exit Sub

    'Loop back to find the start entry for this proc, adding all lower-level durations
    For plProcIdx = plProcIdx To 1 Step -1

        'Start=0 means it's a processed sub-routine, so total their durations
        If pauProcMonitor(plProcIdx).Start = 0 Then
            cDur = cDur + pauProcMonitor(plProcIdx).Duration

        ElseIf pauProcMonitor(plProcIdx).ProcID = sID Then
            Exit For
        End If
    Next

    'If we found the start...
    If plProcIdx > 0 Then

        'Get the end time
        QueryCounter cEnd

        'Work out the duration spent in this routine
        With pauProcMonitor(plProcIdx)

            'Remember how long this one took in total
            .Duration = cEnd - .Start
            .Start = 0

            'And how much time was actually spent in this routine
            cDur = .Duration - cDur
        End With

        'If we've been stored before...
        If pdictResults.Exists(sID) Then

            '... get the stored info and add to it
            uProcResults = pdictResults(sID)
            uProcResults.TotalDuration = uProcResults.TotalDuration + cDur
            uProcResults.Count = uProcResults.Count + 1
            If cDur > uProcResults.MaxDuration Then uProcResults.MaxDuration = cDur
        Else
            'Otherwise create a new info record
            uProcResults.ProcID = sID
            uProcResults.Count = 1
            uProcResults.TotalDuration = cDur
            uProcResults.MaxDuration = cDur
        End If

        'Store it in the Dictionary
        pdictResults(sID) = uProcResults
    End If

End Sub


'Start the monitoring
Public Sub PerfMonStartMonitoring()

    On Error Resume Next

    'We're monitoring!
    pbMonitoring = True

    'Initialise the results dictionary, monitoring array and array index
    Set pdictResults = New Dictionary
    ReDim pauProcMonitor(1 To 100)
    plProcIdx = 0

    'Get the start time of this monitoring pass
    QueryCounter mcStart

End Sub


'End the monitoring
Public Sub PerfMonStopMonitoring(Optional sFileName As String)

    Dim cFreq As Currency
    Dim vaResults As Variant
    Dim uProcResult As ProcResults
    Dim i As Integer
    Dim sResult As String
    Dim iFile As Long
    Dim sErrMessage As String

    'We've stopped monitoring!
    pbMonitoring = False

    'Record the end time
    QueryCounter mcEnd

    'Retrieve the clock speed
    QueryFrequency cFreq

    'Get the performance data
    vaResults = pdictResults.Items

    'Convert the UDT array to a string for writing

    'Write the total logging duration to the file as a sanity check
    sResult = "Tot Dur" & vbTab & Format$((mcEnd - mcStart) / cFreq, "#,##0.00000") & vbCr & vbCr

    'Write header record
    sResult = sResult & Replace$("Project.Module.Proc.Count.Total.Avg.Max", ".", vbTab) & vbCr

    'Write all the detail records
    For i = LBound(vaResults) To UBound(vaResults)
        uProcResult = vaResults(i)
        With uProcResult
            '.ProcID         = The procedure ID in the form DLL.Module.Proc
            '.Count          = The number of times the routine was called
            '.TotalDuration  = The total time spent inside the routine
            '.MaxDuration    = The maximum time spent inside the routine for a single call

            sResult = sResult & Replace$(.ProcID, ".", vbTab) & vbTab & _
                      Format$(.Count, "#,##0") & vbTab & _
                      Format$(.TotalDuration / cFreq, "#,##0.00000") & vbTab & _
                      Format$(.TotalDuration / cFreq / .Count, "#,##0.00000") & vbTab & _
                      Format$(.MaxDuration / cFreq, "#,##0.00000") & vbCr
        End With
    Next

    If Len(sFileName) = 0 Then
        'No file supplied, so write to the clipboard
        Clipboard.Clear
        Clipboard.SetText sResult
    Else
        'Writing to a file
        On Error GoTo Err_FileWriteError

ptrWriteFile:
        'Get the next file number
        iFile = FreeFile

        'Create the file
        Open sFileName For Output As #iFile
        Print #iFile, sResult
        Close #iFile
    End If

    Exit Sub

Err_FileWriteError:

    Select Case Err.Number
    Case 52, 70                                       'Permission Denied
        sErrMessage = "The file '" & sFileName & "' is locked and can't be written to." & vbCr & _
                      "Click Abort to stop writing the performance monitoring results." & vbCr & _
                      "Click Retry to try to write to the file again." & vbCr & _
                      "Click Ignore to write the results to the clipboard."

    Case Else                                         'Any other error
        sErrMessage = "An error occurred while attempting to create the performance monitoring" & vbCr & _
                      "results file '" & sFileName & "'." & vbCr & _
                      "The error is " & Err.Number & ": " & Err.Description & vbCr & _
                      "Click Abort to stop writing the performance monitoring results." & vbCr & _
                      "Click Retry to try to write to the file again." & vbCr & _
                      "Click Ignore to write the results to the clipboard."
    End Select

    Select Case MsgBox(sErrMessage, vbAbortRetryIgnore + vbDefaultButton2 + vbCritical, "Perfomance Monitor")
    Case vbAbort
        'Do nothing - we'll just fall out of the routine

    Case vbRetry
        'Try again (e.g. after the user has closed the file)
        Resume ptrWriteFile

    Case vbIgnore
        'Write to the clipboard
        Clipboard.Clear
        Clipboard.SetText sResult
    End Select

    Exit Sub
End Sub


