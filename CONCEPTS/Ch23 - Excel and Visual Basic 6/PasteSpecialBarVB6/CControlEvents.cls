VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CControlEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Description:  Class module to hook the Paste Special buttons.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
' Ch23 - The only changes required to make this class work within a
'        COM Addin is to prefix 'CommandBars' and 'Selection' by our
'        gxlApp variable
'
Option Explicit

' **************************************************************
' Class Variable Declarations Follow
' **************************************************************
Private WithEvents mctlPasteSpecial As office.CommandBarButton
Attribute mctlPasteSpecial.VB_VarHelpID = -1


' **************************************************************
' Class Event Procedures Follow
' **************************************************************
Private Sub Class_Initialize()
    ' Find and hook one of our custom buttons.
    ' The Click event will fire when *any* of the controls with
    ' the same ID and Tag are clicked, as well as when the
    ' built-in control whose ID we're using is clicked.
    ' We've given all our controls the same ID and Tag, so
    ' we're handling the click events for all our controls
    ' using a single hook and event handler.
    Set mctlPasteSpecial = gxlApp.CommandBars.FindControl(Tag:=gsMENU_TAG)
End Sub


Private Sub Class_Terminate()
    Set mctlPasteSpecial = Nothing
End Sub


' **************************************************************
' mctlPasteSpecial Event Procedures Follow
' **************************************************************
Private Sub mctlPasteSpecial_Click(ByVal Ctrl As office.CommandBarButton, CancelDefault As Boolean)

    Dim uPasteType As XlPasteType

    ' This is called for all instances of the built-in
    ' Paste Special > Values button as well as our custom
    ' Paste Special buttons, so check if it's one of ours.
    ' If the button is not one of ours, we'll do nothing
    ' and Excel will perform its normal action for that
    ' button.
    If Ctrl.Tag = gsMENU_TAG Then
    
        ' It is one of ours, so set the appropriate paste type.
        Select Case Ctrl.Parameter
            Case gsMENU_PS_ALL
                uPasteType = xlPasteAll
            Case gsMENU_PS_FORMULAS
                uPasteType = xlPasteFormulas
            Case gsMENU_PS_VALUES
                uPasteType = xlPasteValues
            Case gsMENU_PS_FORMATS
                uPasteType = xlPasteFormats
            Case gsMENU_PS_COMMENTS
                uPasteType = xlPasteComments
            Case gsMENU_PS_VALIDATION
                uPasteType = 6  ' xlPasteValidation in 2002+
            Case gsMENU_PS_COLWIDTHS
                uPasteType = 8  ' xlPasteColumnWidths in 2002+
        End Select
    
        ' If the paste special doesn't succeed, fail silently.
        On Error Resume Next
            gxlApp.Selection.PasteSpecial uPasteType
        On Error GoTo 0
        
        ' We handled the event, so cancel its default behavior.
        CancelDefault = True
        
    End If
    
End Sub
