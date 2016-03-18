Attribute VB_Name = "MGlobals"
'
' Description:    This module holds declarations for global constants,
'                 variables, type structures, and DLLs.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
' Ch23 - Added the gxlApp variable to hold a reference to Excel.
'        Renamed the toolbar constants, so it will work alongside the Excel
'        Addin from Chapter 8
'
Option Explicit
Option Private Module

' **************************************************************************
' Global Constant Declarations Follow
' **************************************************************************
' Application identification constant.
Public Const gsAPP_TITLE As String = "Paste Special Bar - VB6"

Public Const gsMENU_NAME As String = "Paste Special VB6"
Public Const gsMENU_TAG As String = "pxlPasteSpecialVB6"
Public Const gsMENU_PS_ALL As String = "All"
Public Const gsMENU_PS_FORMULAS As String = "Formulas"
Public Const gsMENU_PS_VALUES As String = "Values"
Public Const gsMENU_PS_FORMATS As String = "Formatting"
Public Const gsMENU_PS_COMMENTS As String = "Comments"
Public Const gsMENU_PS_VALIDATION As String = "Validation"
Public Const gsMENU_PS_COLWIDTHS As String = "ColWidths"

'Registry locations for storing the commandbar position settings
Public Const gsREG_APP As String = "Professional Excel Development\Paste Special Bar VB6"
Public Const gsREG_SETTINGS As String = "CommandBarSettings"

' **************************************************************************
' Global Variable Declarations Follow
' **************************************************************************
Public gbShutdownInProgress As Boolean
Public gclsControlEvents As CControlEvents

' Ch23
Public gxlApp As Excel.Application


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   This routine initializes application global variables.
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 04/30/08    Rob Bovey       Created
'
Public Sub InitGlobals()
    gbShutdownInProgress = False
End Sub

