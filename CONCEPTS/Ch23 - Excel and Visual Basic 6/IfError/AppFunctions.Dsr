VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} AppFunctions 
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10665
   _ExtentX        =   18812
   _ExtentY        =   16272
   _Version        =   393216
   AppName         =   "Microsoft Excel"
   AppVer          =   "Microsoft Excel 10.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel"
End
Attribute VB_Name = "AppFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Description:  Designer class for Excel to use to connect to our Automation Addin,
'               allowing us to use the Excel object model from within our functions
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit

'Reference to the Excel application
Dim mxlApp As Excel.Application


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Called when the automation addin is loaded
'
' Date          Developer       Action
' --------------------------------------------------------------
' 08 Jun 08     Rob Bovey  		Created
'
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    'Store away a reference to the Application object
    Set mxlApp = Application
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Volatile function to return VB's Timer value
'
' Date          Developer       Action
' --------------------------------------------------------------
' 08 Jun 08     Rob Bovey  		Created
'
Public Function VBTimer() As Double
    mxlApp.Volatile True
    VBTimer = Timer
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Function to count how many items in Source lie between Min and Max
'
' Date          Developer       Action
' --------------------------------------------------------------
' 08 Jun 08     Rob Bovey  		Created
'
Public Function CountBetween(ByRef Source As Range, ByVal Min As Double, ByVal Max As Double) As Double
    
    'If we get an error, return zero
    On Error GoTo ErrHandler
    
    'Count the items bigger than Min
    CountBetween = mxlApp.WorksheetFunction.CountIf(Source, ">" & Min)
    
    'Subtract the items bigger than Max, giving the number of items in between
    CountBetween = CountBetween - mxlApp.WorksheetFunction.CountIf(Source, ">=" & Max)

Exit Function

ErrHandler:
    CountBetween = 0
End Function
