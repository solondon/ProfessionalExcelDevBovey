VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPerfMon"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Class to support performance monitoring of combined Excel/VB6
'   development.  Instead of adding a reference to the PerfMonitor
'   dll in the Excel project, include this class in the VB6 project.
'   Excel will then call here, which passes the calls on to VB's
'   instance of the PerfMonitor dll, so both Excel and VB6 routines
'   are listed in the same monitoring results table.
'
'   Version Date        Author          Comment
'   0.0.1   09-02-2004  Stephen Bullen  Initial Version
'
Option Explicit

'Called to indicate the start of a procedure
Public Sub PerfMonProcStart(sProc As String)
    PerfMonitor.PerfMonProcStart sProc
End Sub

'Called to indicate the end of a procedure
Public Sub PerfMonProcEnd(sProc As String)
    PerfMonitor.PerfMonProcEnd sProc
End Sub

'Called to indicate the start of the monitoring
Public Sub PerfMonStartMonitoring()
    PerfMonitor.PerfMonStartMonitoring
End Sub

'Called to indicate the end of the monitoring
Public Sub PerfMonStopMonitoring(Optional sFile As String)
    PerfMonitor.PerfMonStopMonitoring sFile
End Sub
