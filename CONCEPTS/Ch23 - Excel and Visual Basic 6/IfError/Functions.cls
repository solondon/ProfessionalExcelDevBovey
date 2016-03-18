VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Functions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Description:  Contains VB6 functions designed to be used as
'               custom Excel worksheet functions, through
'               Automation Addins
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:     This function provides a short-cut replacement
'               for the common worksheet function construct:
'               =IF(ISERROR(<some_function>),0,<some_function>)
'
' Arguments:    ToEvaluate      A value, expression or cell
'                               reference to be evaluated.
'               Default         A value, expression or cell
'                               reference to be returned if the
'                               ToEvaluate argument evaluates to
'                               an error condition.
'
' Returns:      Variant         vToEvaluate will be returned if
'                               it does not evaluate to an
'                               error, otherwise the vDefault
'                               argument will be returned.
'
' Date          Developer       Action
' --------------------------------------------------------------
' 30 Apr 08     Rob Bovey       Created
' 08 Jun 08     Rob Bovey  		Copied to VB6 as an Automation Addin
'
Public Function IFERROR(ByRef ToEvaluate As Variant, _
                        ByRef Default As Variant) As Variant
    If IsError(ToEvaluate) Then
        IFERROR = Default
    Else
        IFERROR = ToEvaluate
    End If
End Function



