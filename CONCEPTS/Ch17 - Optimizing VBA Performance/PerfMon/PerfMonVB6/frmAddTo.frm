VERSION 5.00
Begin VB.Form frmAddTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Remove PerfMon Calls"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAddRefs 
      Caption         =   "Add &References to PerfMonitor library"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.OptionButton optSelProc 
      Caption         =   "&Selected Procedure (<ProcName>)"
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4050
   End
   Begin VB.OptionButton optSelModule 
      Caption         =   "Selected &Module (<ModName>)"
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4050
   End
   Begin VB.OptionButton optSelProject 
      Caption         =   "Selected &Project (<ProjName>)"
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4050
   End
   Begin VB.OptionButton optAllProjects 
      Caption         =   "&All Open Projects"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   -1  'True
      Width           =   4050
   End
   Begin VB.Label lblAddTo 
      Caption         =   "Add PerfMon calls to:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   4080
   End
End
Attribute VB_Name = "frmAddTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   Form to allow the user to specify the scope of the PerfMon calls
'   and whether to add/remove a reference to the PerfMonitor dll
'
'   Version Date        Author          Comment
'   0.0.1   09-02-2004  Stephen Bullen  Initial Version
'
Option Explicit
Option Compare Binary

'Whether OK or Cancel was pressed
Dim mbOK As Boolean

'Whether adding or removing calls
Dim mlAddRem As pmAddRemove

'Allow the calling code to specify whether adding or removing calls
'Check the IDE to see if there is a project/module selected
'Set the captions on the labels accordingly
Public Sub Initialise(oVBE As VBIDE.VBE, lAddRem As pmAddRemove)

    Dim lLine As Long
    Dim lType As vbext_ProcKind
    Dim sProc As String
    Dim lScope As pmScope

    mlAddRem = lAddRem

    'Set the correct captions, depending on whether we're adding or removing
    If lAddRem = pmAddRemoveAdd Then
        Me.Caption = "Add PerfMon Calls"
        Me.lblAddTo.Caption = "Add PerfMon calls to:"
        lScope = Val(GetSetting("PerfMon", "VB6", "AddScope", "0"))
        Me.chkAddRefs.Value = 1
        Me.chkAddRefs.Caption = "Add &References to PerfMonitor library"
    Else
        Me.Caption = "Remove PerfMon Calls"
        Me.lblAddTo.Caption = "Remove PerfMon calls from:"
        lScope = Val(GetSetting("PerfMon", "VB6", "RemScope", "0"))
        Me.chkAddRefs.Value = 0
        Me.chkAddRefs.Caption = "Remove &References to PerfMonitor library"
    End If

    'Default the labels to no active project, module or procedure
    Me.optSelProject.Enabled = False
    Me.optSelProject.Caption = "Selected &Project"
    Me.optSelModule.Enabled = False
    Me.optSelModule.Caption = "Selected &Module"
    Me.optSelProc.Enabled = False
    Me.optSelProc.Caption = "&Selected Procedure"

    'Is there an unprotected project selected?
    If Not oVBE.ActiveVBProject Is Nothing Then

        'Yes, so change the option button to show the project name and enable it
        Me.optSelProject.Enabled = True
        Me.optSelProject.Caption = "Selected &Project (" & oVBE.ActiveVBProject.Name & ")"

        'Is there a selected VB component, and does it have a code module
        If Not oVBE.SelectedVBComponent Is Nothing Then
            If HasCodeModule(oVBE.SelectedVBComponent) Then

                'Yes, so change the option button to show the module name and enable it
                Me.optSelModule.Enabled = True
                Me.optSelModule.Caption = "Selected &Module (" & oVBE.SelectedVBComponent.Name & ")"

                'Is there a code pane showing?
                If Not oVBE.ActiveCodePane Is Nothing Then

                    'Yes, so get the procedure under the cursor
                    oVBE.ActiveCodePane.GetSelection lLine, 0, 0, 0
                    sProc = oVBE.ActiveCodePane.CodeModule.ProcOfLine(lLine, lType)
                End If

                'Did we get a procedure?
                If sProc <> "" Then

                    'Yes, so add its type and enable that option
                    Select Case lType
                    Case vbext_pk_Get: sProc = sProc & " [Get]"
                    Case vbext_pk_Let: sProc = sProc & " [Let]"
                    Case vbext_pk_Set: sProc = sProc & " [Set]"
                    End Select

                    Me.optSelProc.Enabled = True
                    Me.optSelProc.Caption = "&Selected Procedure (" & sProc & ")"
                End If
            End If
        End If
    End If

    'Set the initial scope, depending on the registry value and whether the option is enabled.
    Me.optAllProjects.Value = True
    Me.optAllProjects.Value = lScope = pmScopeAllProjects
    Me.optSelProject.Value = lScope = pmScopeSelProject
    Me.optSelModule.Value = lScope = pmScopeSelModule And Me.optSelModule.Enabled
    Me.optSelProc.Value = lScope = pmScopeSelProc And Me.optSelProc.Enabled

End Sub

'Handle clicking the OK button
Private Sub cmdOK_Click()
    mbOK = True
    Me.Hide

    If mlAddRem = pmAddRemoveAdd Then
        SaveSetting "PerfMon", "VB6", "AddScope", Scope
    Else
        SaveSetting "PerfMon", "VB6", "RemScope", Scope
    End If

End Sub

'Handle clicking the Cancel button
Private Sub cmdCancel_Click()
    mbOK = False
    Me.Hide
End Sub

'Closing the form with the [x] is the same as cancelling
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        cmdCancel_Click
        Cancel = True
    End If
End Sub

'Whether the user OK'd the form or not
Public Property Get OK() As Boolean
    OK = mbOK
End Property

'The selected scope
Public Property Get Scope() As pmScope

    Select Case True
    Case Me.optAllProjects.Value: Scope = pmScopeAllProjects
    Case Me.optSelProject.Value: Scope = pmScopeSelProject
    Case Me.optSelModule.Value: Scope = pmScopeSelModule
    Case Me.optSelProc.Value: Scope = pmScopeSelProc
    End Select

End Property

'Whether to add/remove the reference to the PerfMonitor dll
Public Property Get AddRemoveRefs() As Boolean
    AddRemoveRefs = Me.chkAddRefs.Value
End Property

'Helper function to check if a VBComponent has a code module
Private Function HasCodeModule(oVBC As VBComponent) As Boolean

    Dim oCM As CodeModule
    On Error Resume Next
    Set oCM = oVBC.CodeModule
    HasCodeModule = Err = 0

End Function
