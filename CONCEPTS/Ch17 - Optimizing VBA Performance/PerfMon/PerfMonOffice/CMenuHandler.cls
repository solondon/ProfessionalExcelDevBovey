VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenuHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   Class to handle the menu click events and perform the adding/removing of the calls
'   We use the VBE.Events.CommandBarEvents instead of CommandBarButton because the VBE Events
'   method works with both VB6 and the Office IDE, while the CommandBarButton method only works
'   with the Office IDE.
'
'   Version Date        Author          Comment
'   0.0.1   09-02-2004  Stephen Bullen  Initial Version
'
Option Explicit
Option Compare Binary

Dim moVBE As VBIDE.VBE
Dim WithEvents oMenu As CommandBarEvents
Attribute oMenu.VB_VarHelpID = -1

'Store references to the VBE and the button we're handling events for
Public Sub Initialise(oVBE As VBIDE.VBE, oButton As CommandBarButton)
    Set moVBE = oVBE
    Set oMenu = oVBE.Events.CommandBarEvents(oButton)
End Sub


'Called when the menu item is clicked
Private Sub oMenu_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Select Case CommandBarControl.Parameter
    Case "About"
        frmAbout.Show vbModal
    Case "AddCode"
        AddRemove pmAddRemoveAdd
    Case "RemoveCode"
        AddRemove pmAddRemoveRemove
    End Select

    handled = True
End Sub


'Common routine to add or remove PerfMon calls
Private Sub AddRemove(lAddRem As pmAddRemove)

    Dim oVBP As VBProject, oVBC As VBComponent
    Dim lLine As Long, lProcKind As vbext_ProcKind
    Dim sProc As String
    Dim oAddTo As frmAddTo
    Dim lScope As pmScope
    Dim lEndLine As Long
    Dim oRef As VBIDE.Reference
    Dim sCurrProc As String, lCurrKind As vbext_ProcKind
    Dim sProcID As String

    'The GUID of the PerfMonitor dll, so we can add/remove references to it
    Const cPERFMON_GUID As String = "{54115356-8232-4792-9D61-A3EB3B84ECF0}"

    On Error GoTo ErrHandler

    'Initialise and show the 'Add To' form, to get the scope of where to add/remove the calls
    Set oAddTo = New frmAddTo
    oAddTo.Initialise moVBE, lAddRem

    oAddTo.Show vbModal

    'If Cancelled, quit the routine
    If Not oAddTo.OK Then Exit Sub

    'Get the form's scope
    lScope = oAddTo.Scope

    'If we only want to process the current procedure, work out which one it is
    If lScope = pmScopeSelProc Then
        moVBE.ActiveCodePane.GetSelection lLine, 0, 0, 0
        sCurrProc = moVBE.SelectedVBComponent.CodeModule.ProcOfLine(lLine, lCurrKind)
    End If

    'Initialise and show the progress form
    If lAddRem = pmAddRemoveAdd Then
        frmProgress.Action = "Adding PerfMon calls to procedure:"
    Else
        frmProgress.Action = "Removing PerfMon calls from procedure:"
    End If

    frmProgress.Procedure = ""
    frmProgress.Show vbModeless

    'Loop through all the projects
    For Each oVBP In moVBE.VBProjects

        'If doing all projects, or this is the selected project...
        If oVBP.Protection = vbext_pp_none And (lScope = pmScopeAllProjects Or oVBP Is moVBE.ActiveVBProject) Then

            'Loop through all VB Components in the project
            For Each oVBC In oVBP.VBComponents

                'Does it have a code module?
                If HasCodeModule(oVBC) Then

                    'If doing all projects or all modules, or this is the selected module...
                    If lScope < pmScopeSelModule Or oVBC Is moVBE.SelectedVBComponent Then

                        With oVBC.CodeModule

                            'Skip the declaration lines
                            lLine = .CountOfDeclarationLines + 1

                            'Loop through the module, using Do...Loop as the module changes size!
                            Do Until lLine >= .CountOfLines

                                'Get the procedure name and type at this line
                                sProc = .ProcOfLine(lLine, lProcKind)

                                'If doing all projects, all modules, all procs or this is the selected proc...
                                If lScope < pmScopeSelProc Or (sProc = sCurrProc And lProcKind = lCurrKind) Then

                                    'Did we get a procedure?
                                    If sProc <> "" Then

                                        'Get its end line
                                        lEndLine = lLine + .ProcCountLines(sProc, lProcKind) - 1

                                        'Work out its unique ID
                                        sProcID = oVBP.Name & "." & oVBC.Name & "." & sProc

                                        Select Case lProcKind
                                        Case vbext_pk_Set: sProcID = sProcID & "_Set"
                                        Case vbext_pk_Let: sProcID = sProcID & "_Let"
                                        Case vbext_pk_Get: sProcID = sProcID & "_Get"
                                        End Select

                                        'Update the progress form
                                        frmProgress.Procedure = sProcID

                                        'Add/Rremove the code from this procedure
                                        AddRemCode oVBC, sProc, lProcKind, lLine, lEndLine, lAddRem = pmAddRemoveAdd, sProcID
                                    End If
                                End If

                                'Skip the rest of this procedure
                                lLine = lLine + .ProcCountLines(sProc, lProcKind)
                            Loop
                        End With                      'Code module
                    End If                            'Do this module?
                End If                                'HasCodeModule?
            Next                                      'VBComponent

            'Add/remove the reference to the PerfMonitor Dll
            If oAddTo.AddRemoveRefs Then
                If lAddRem = pmAddRemoveAdd Then
                    On Error Resume Next
                    oVBP.References.AddFromGuid cPERFMON_GUID, 1, 0
                    On Error GoTo ErrHandler
                Else
                    For Each oRef In oVBP.References
                        If oRef.Guid = cPERFMON_GUID Then
                            oVBP.References.Remove oRef
                            Exit For
                        End If
                    Next
                End If
            End If
        End If                                        'Do this project?
    Next                                              'VBProject

    'Unload the progress form
    Unload frmProgress

    Exit Sub

ErrHandler:
    Stop
    Resume

End Sub


'Add/remove PerfMon calls for a given procedure
Private Sub AddRemCode(oVBC As VBComponent, ByVal sProc As String, ByVal lProcKind As vbext_ProcKind, _
                       ByVal lStart As Long, ByVal lEnd As Long, ByVal bDoAdd As Boolean, ByVal sProcID As String)

    Dim lLine As Long, lBodyLine As Long
    Dim sLine As String
    Dim i As Long, j As Long
    Dim v As Variant
    Dim oCM As CodeModule

    On Error GoTo ErrHandler

    'Get the code module
    Set oCM = oVBC.CodeModule

    '**********************
    'Remove existing calls
    '**********************
    
    lLine = lStart
    Do
        'Get the line of code
        sLine = oCM.Lines(lLine, 1)

        'Does it have a PerfMon call in?
        If InStr(1, sLine, "PerfMon") > 0 Then
            For Each v In Array("PerfMonProcStart", "PerfMonProcEnd", "PerfMonStartMonitoring", "PerfMonStopMonitoring")
                i = InStr(1, sLine, v)
        
                If i > 0 Then
                    'Yes, so find the end of it
                    For j = i + Len(v) To Len(sLine)
                        If Mid$(sLine, j, 1) = """" Then
                            j = InStr(j + 1, sLine, """")
                        ElseIf Mid$(sLine, j, 1) = "'" Or Mid$(sLine, j, 1) = ":" Then
                            Exit For
                        End If
                    Next
                    
                    If Mid$(sLine, j, 2) = ": " Then j = j + 2
        
                    'Strip it out
                    sLine = Left$(sLine, i - 1) & Mid$(sLine, j)
        
                    'If that was the only thing there...
                    If Trim$(sLine) = "" Then
                        'Delete the line and adjust our pointers
                        oCM.DeleteLines lLine, 1
                        lLine = lLine - 1
                        lEnd = lEnd - 1
                    Else
                        'Replace the line with the PerfMon call removed
                        oCM.ReplaceLine lLine, sLine
                    End If
                End If
            Next
        End If

        lLine = lLine + 1
    Loop While lLine <= lEnd

    '**********************
    'Add new calls
    '**********************
    
    If bDoAdd Then
        lLine = lStart
        Do
            'Get the line of code
            sLine = Trim$(oCM.Lines(lLine, 1))

            'Ignore comments and blank lines
            If Len(sLine) = 0 Or Left$(sLine, 1) = "'" Or Left$(sLine, 4) = "Rem " Then GoTo ptrNextLine

            'Get the line number of the start of the procedure (i.e. after procedure header comments)
            lBodyLine = oCM.ProcBodyLine(sProc, lProcKind)
            
            'If we're before the procedure code, ...
            If lLine <= lBodyLine Then
            
                'We want to add a call to PerfMonProcStart as the first line in the procedure
                
                'Skip over continued procedure declaration lines
                Do Until Right$(oCM.Lines(lBodyLine, 1), 2) <> " _"
                    lBodyLine = lBodyLine + 1
                Loop

                'Check if the next line is the end of the procedure, and if so, skip to the end.
                sLine = oCM.Lines(lBodyLine, 1)
                For Each v In Array("End Sub", "End Function", "End Property")
                    If InStr(1, sLine, v) > 0 Then GoTo ptrNextLine
                Next

                'Insert a line for the PerfMonProcStart call
                oCM.InsertLines lBodyLine + 1, "PerfMonProcStart """ & sProcID & """"
                
                'Adjust our pointers
                lLine = lBodyLine + 1
                lEnd = lEnd + 1
            Else
                'We're inside the procedure, so look for exit points
                For Each v In Array("Exit Sub", "Exit Function", "Exit Property")
                    i = InStr(1, sLine, v)
                    If i > 0 Then
                        'Found one, so insert a call to PerfMonProcEnd within the line
                        sLine = Left$(sLine, i - 1) & "PerfMonProcEnd """ & sProcID & """:" & Mid$(sLine, i)
                        oCM.ReplaceLine lLine, sLine
                        Exit For
                    End If
                Next
                
                'As well as exit points, look for the end of the routine.
                For Each v In Array("End Sub", "End Function", "End Property")
                    i = InStr(1, sLine, v)
                    If i > 0 Then
                        'Insert a line before the end of the routine
                        oCM.InsertLines lLine, "PerfMonProcEnd """ & sProcID & """"
                        lLine = lLine + 1
                        lEnd = lEnd + 1
                        Exit For
                    End If
                Next
            End If
ptrNextLine:
            lLine = lLine + 1
        Loop While lLine <= lEnd
    End If                                            'bDoAdd

    Exit Sub

ErrHandler:
    Stop
    Resume

End Sub

'Helper routine to check if a VBComponent has a code module.
'While they all do in the Office VBE, they might not in VB6.
Private Function HasCodeModule(oVBC As VBComponent) As Boolean

    Dim oCM As CodeModule
    On Error Resume Next
    Set oCM = oVBC.CodeModule
    HasCodeModule = Err = 0

End Function


