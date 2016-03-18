Attribute VB_Name = "MStandardCode"
'
' Description:  Contains standard code library routines that are
'               used without modification in many different
'               projects.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
' Ch23 - Removed the ResetAppProperties routine, as it doesn't apply to COM Addins
'
Option Explicit
Option Private Module

' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const mszMODULE As String = "MStandardCode"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Set a commandbar's position and visibilty, reading the values
'             from the registry
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 04/30/08    Rob Bovey       Created
'
Public Sub SetCommandBarPosition(ByRef cbBar As office.CommandBar)

    If Not cbBar Is Nothing Then
        'Set the bar's position, by reading them from the registry.  This ensures the bar
        'appears in the same place, with the same visibility, as when Excel was closed.
        With cbBar
            'Are there any registry settings?
            If GetSetting(gsREG_APP, gsREG_SETTINGS, "Visible", "NotSet") = "NotSet" Then
                'No, so just make it visible and use the defaults
                .Visible = True
            Else
                'Yes, so set the bar's visibility and docking status
                .Visible = GetSetting(gsREG_APP, gsREG_SETTINGS, "Visible", "Y") = "Y"
                .Position = CLng(GetSetting(gsREG_APP, gsREG_SETTINGS, "Position"))
                
                If .Position = msoBarFloating Then
                    'If floating, we set all four position values
                    .Left = CLng(GetSetting(gsREG_APP, gsREG_SETTINGS, "Left"))
                    .Top = CLng(GetSetting(gsREG_APP, gsREG_SETTINGS, "Top"))
                    .Width = CLng(GetSetting(gsREG_APP, gsREG_SETTINGS, "Width"))
                    .Height = CLng(GetSetting(gsREG_APP, gsREG_SETTINGS, "Height"))
                Else
                    'Not floating, so set which row and where on that row/column
                    .RowIndex = CLng(GetSetting(gsREG_APP, gsREG_SETTINGS, "RowIndex"))
                    
                    If .Position = msoBarTop Or .Position = msoBarBottom Then
                        .Left = CLng(GetSetting(gsREG_APP, gsREG_SETTINGS, "Left"))
                    Else
                        .Top = CLng(GetSetting(gsREG_APP, gsREG_SETTINGS, "Top"))
                    End If
                End If
            End If
        End With
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Save a commandbar's position and visibilty, so we can recreate
'             it with the same settings
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 04/30/08    Rob Bovey       Created
'
Public Sub StoreCommandBarPosition(ByRef cbBar As office.CommandBar)
    
    'If we found it, write its positional settings to the registry
    If Not cbBar Is Nothing Then
        With cbBar
            SaveSetting gsREG_APP, gsREG_SETTINGS, "Visible", IIf(.Visible, "Y", "N")
            SaveSetting gsREG_APP, gsREG_SETTINGS, "Position", CStr(.Position)
            SaveSetting gsREG_APP, gsREG_SETTINGS, "Left", CStr(.Left)
            SaveSetting gsREG_APP, gsREG_SETTINGS, "Top", CStr(.Top)
            SaveSetting gsREG_APP, gsREG_SETTINGS, "Width", CStr(.Width)
            SaveSetting gsREG_APP, gsREG_SETTINGS, "Height", CStr(.Height)
            SaveSetting gsREG_APP, gsREG_SETTINGS, "RowIndex", CStr(.RowIndex)
        End With
    End If
    
End Sub






