
Welcome to PerfMon
------------------

PerfMon is a set of addins for the Office and VB6 IDEs, to allow in-line monitoring of VB/VBA code execution.  The files in this folder are:

 PerfMon ReadMe.txt 	This file!
 PerfMonitor.dll	An ActiveX dll to monitor VB/VBA routines.
 PerfMonOffice.dll	An addin to the Office VBE to add/remove code to call to the PerfMonitor dll
 PerfMonVB6.dll		An addin to the VB6 IDE to add/remove code to call to the PerfMonitor dll
 CPerfMon.cls		A class to allow PerfMon to work in a combined Excel/VB6 solution

To use PerfMon:

 1. Use Regsrvr32 to register the dlls
 2. This will add a PerfMon menu to the Tools > Addins menu and code window right-click menu
 3. Click Tools > Addins > PerfMon > Add PerfMon Calls
 4. Select which routines to add the calls to and click OK.  This will add a line of code to the top and bottom of every routine that calls into the PerfMonitor dll, notifying it of the start and end of the routine.
 5. Find the routine to monitor and add calls to PerfMonStartMonitoring and PerfMonStopMonitoring, giving an optional filename.  Once done, the routine may look like the following:

Sub LongRoutine()

    PerfMonStartMonitoring
    PerfMonProcStart "Project.Module.LongRoutine"

    'Do stuff

    PerfMonProcEnd "Project.Module.LongRoutine"
    PerfMonStopMonitoring "C:\LongRoutine.txt"

End Sub

If you do not supply a file name to the PerfMonStopMonitoring call, the results will be copied to the clipboard.

 6. Run the routine
 7. Import the tab-delimited text file into a new Excel workbook for analysis.


To analyse a long routine, you can break it down by adding dummy start and end calls, such as:

Sub LongRoutine()

    PerfMonStartMonitoring
    PerfMonProcStart "Project.Module.LongRoutine"


    PerfMonProcStart "Project.Module.LongRoutine1"
    'Do stuff
    PerfMonProcEnd "Project.Module.LongRoutine1"


    PerfMonProcStart "Project.Module.LongRoutine2"
    'Do stuff
    PerfMonProcEnd "Project.Module.LongRoutine2"


    PerfMonProcEnd "Project.Module.LongRoutine"
    PerfMonStopMonitoring "C:\LongRoutine.txt"

End Sub


If you are doing combined Excel/VB6 development, the CPerfMon class should be added to the VB6 project and the Excel project should *not* reference the PerfMonitor dll.  All the calls that the addin adds to the Excel code will then call the procedures in the CPerfMon class, which in turn routes them to VB6's instance of the PerfMonitor dll.  This means that all the routines in both the Excel project and the VB project will be shown in the same results list.

