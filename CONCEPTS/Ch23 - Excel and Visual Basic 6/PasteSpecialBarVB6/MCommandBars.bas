Attribute VB_Name = "MCommandBars"
'
' Description:  This module builds the custom CommandBar.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
' Ch23 - The addin workbook from Chapter 8 used the table-driven command bar builder.
'        In this COM Addin, it is easier to create our command bar directly.
'
Option Explicit
Option Private Module

' ****************************************************************************
' Module Constant Declarations Follow
' ****************************************************************************
Private Const msMODULE As String = "MCommandBars"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Creates the Paste Special command bar, item by item
'
' Returns:    Boolean         True on success, False on error.
'
' Date        Developer       Action
' ----------------------------------------------------------------------------
' 06 Jun 08   Stephen Bullen  Created
'
Public Function bBuildCommandBars() As Boolean
    
    Const sSOURCE As String = "bBuildCommandBars"

    Dim cbBar As CommandBar
    
    On Error GoTo ErrorHandler

    'Delete the commandbar in case it was left behind after a crash
    ResetCommandBars
    
    'Create our commandbar, intially floating
    Set cbBar = gxlApp.CommandBars.Add(gsMENU_NAME, msoBarFloating, False, True)
    
    'Add the buttons to the command bar
    AddButton cbBar, "Paste All", gsMENU_TAG, gsMENU_PS_ALL, "picAll"
    AddButton cbBar, "Paste Formulas", gsMENU_TAG, gsMENU_PS_FORMULAS, "picFormulas"
    AddButton cbBar, "Paste Values", "", gsMENU_PS_VALUES, ""
    AddButton cbBar, "Paste Formatting", gsMENU_TAG, gsMENU_PS_FORMATS, 369
    AddButton cbBar, "Paste Comments", gsMENU_TAG, gsMENU_PS_COMMENTS, "picComments"
    AddButton cbBar, "Paste Validation", gsMENU_TAG, gsMENU_PS_VALIDATION, "picValidation"
    AddButton cbBar, "Paste Column Widths", gsMENU_TAG, gsMENU_PS_COLWIDTHS, "picWidths"
    
    bBuildCommandBars = True
    
ErrorExit:

    Exit Function
    
ErrorHandler:
    If Err.Number <> glHANDLED_ERROR Then Err.Description = Err.Description & " (" & sSOURCE & ")"
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Add a button to the Paste Special command bar.  Every button is
'             given the built-in ID of 370, being that of Paste Values
'
' Arguments:  cbBar         The command bar to add the button to
'             sTip          The button's tooltip text
'             sTag          The button's Tag, gsMENU_TAG for all except Paste Values itself
'             sParameter    The button's parameter, used to identify the action
'             vPicture      The button's image, either a face ID or the ID of an image resource
'
' Date        Developer       Action
' ----------------------------------------------------------------------------
' 06 Jun 08   Stephen Bullen  Created
'
Private Sub AddButton(ByRef cbBar As CommandBar, ByVal sTip As String, ByVal sTag As String, _
                      ByVal sParameter As String, ByVal vPicture As Variant)

    Const PASTE_VALUES_ID = 370
    
    Dim btnButton As CommandBarButton

    'Add the button
    Set btnButton = cbBar.Controls.Add(msoControlButton, PASTE_VALUES_ID, sParameter, , True)
    With btnButton
    
        'Set the tooltip and tag if we were given one
        If Len(sTip) > 0 Then .ToolTipText = sTip
        If Len(sTag) > 0 Then .Tag = sTag
        
        'Work out what to do with the picture
        If vPicture <> "" Then

            'If it's a number, assume it's a face ID
            If IsNumeric(vPicture) Then
                .FaceId = Val(vPicture)
            Else
                If Val(gxlApp.Version) >= 10 Then
                    'If Excel 2002+, we set the picture and mask,
                    'loading the bitmaps from the resource file
                    .Picture = LoadResPicture(vPicture, vbResBitmap)
                    .Mask = LoadResPicture(vPicture & "Mask", vbResBitmap)
                Else
                    'Excel 2000, so copy the picture to the clipboard,
                    'paste it to the button and clear the clipboard.
                    'N.B. Uses the code in the MCopyTransparent module, to
                    '     make the bitmap transparent
                    CopyBitmapAsButtonFace LoadResPicture(vPicture, vbResBitmap), RGB(236, 233, 216)
                    .PasteFace
                    
                    Clipboard.SetText ""
                    Clipboard.Clear
                End If
            End If
        End If
    End With

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Delete our commandbar when shutting down, in case we're uninstalled
'
' Date        Developer       Action
' --------------------------------------------------------------------------
' 06 Jun 08   Stephen Bullen  Created
'
Public Sub ResetCommandBars()
    On Error Resume Next
    gxlApp.CommandBars(gsMENU_NAME).Delete
End Sub


