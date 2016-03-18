Attribute VB_Name = "Globals"
'
'   Module to hold global variables used by the performance monitoring
'
'   Version Date        Author          Comment
'   0.0.1   01-10-2003  Stephen Bullen  Initial Version
'
Option Compare Binary
Option Explicit

'Whether we're switched on
Public pbMonitoring As Boolean

'Dictionary of UDTs for storing results
Public pdictResults As Dictionary

'Array of log information, tracing the call stack
Public pauProcMonitor() As ProcMonitor

'Pointer to the last used element in the array
Public plProcIdx As Long


