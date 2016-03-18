Imports Extensibility
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core

#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the $SAFEOBJNAME$Setup project, 
' right click the project in the Solution Explorer, then choose install.
#End Region

<GuidAttribute("B7B5D1ED-8528-4C11-B70A-EE1BD7E4D1BD"), _
 ProgIdAttribute("FirstAddin.Connect")> _
Public Class Connect

#Region "Implements"

    Implements IDTExtensibility2

#End Region

#Region "Module-wide variables"

    'The variable for Excel Application's object.
    Private mxlApp As Excel.Application = Nothing

    'The hooking command bar control event.
    Private WithEvents mcmdFirstAddinButton As  _
                       Office.CommandBarButton = Nothing

    'The command bar variable.
    Private mcbrFirstAddin As Office.CommandBar = Nothing

    'Constants string variables.
    Private Const msTITLE As String = "Our First Add-in"
    Private Const msTAG As String = "OFA"
    Private Const msBUTTON1 As String = "Button 1"
    Private Const msBUTTON2 As String = "Button 2"
    Private Const msBUTTON3 As String = "Button 3"

#End Region

#Region "Connection and Disconnection"

    Public Sub OnConnection(ByVal application As Object, _
                            ByVal connectMode As  _
                            ext_ConnectMode, _
                            ByVal addInInst As Object, _
                            ByRef custom As System.Array) _
                            Implements IDTExtensibility2.OnConnection

        'Customized error message.
        Const sERROR_MESSAGE As String = _
              "An unexpected error has occured."

        Try

            'Instantiate the Excel Application's variable.
            mxlApp = CType(application, Excel.Application)

            'Make sure that the command bar does not exists.
            Delete_Commandbar()

            'Set up the command bar.
            Create_New_Commandbar()

            'Find and hook one of our customs buttons which will hook
            'all our customs buttons to the Click event.
            mcmdFirstAddinButton = _
                      CType(mxlApp.CommandBars.FindControl(Tag:=msTAG),  _
                                              Office.CommandBarButton)

            'The following lines may be necessary to add in order to support
            'Windows XP visual style.
            System.Windows.Forms.Application.EnableVisualStyles()
            System.Windows.Forms.Application.DoEvents()

        Catch GeneralEx As Exception

            'Show the customized message.
            MessageBox.Show(text:=sERROR_MESSAGE, _
                            caption:=msTITLE, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

        End Try

    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As ext_DisconnectMode, _
                               ByRef custom As System.Array) _
                               Implements _
                               IDTExtensibility2.OnDisconnection
        Try

            Delete_Commandbar()

            Release_All_COMObjects(mcmdFirstAddinButton)
            Release_All_COMObjects(mcbrFirstAddin)
            Release_All_COMObjects(mxlApp)

        Catch GeneralEx As Exception

            'An error message during the debug process should be added."

        End Try

    End Sub

#End Region

#Region "Command bar work"

    Private Sub Create_New_Commandbar()

        'The button control variable.
        Dim ctlFirstAddin As Office.CommandBarButton = Nothing

        'Create the temporarily commandbar.
        mcbrFirstAddin = CType(mxlApp.CommandBars.Add(Name:=msTITLE, _
                               Position:=Office.MsoBarPosition.msoBarTop, _
                               Temporary:=True), Office.CommandBar)
        With mcbrFirstAddin

            'Add a new button.
            ctlFirstAddin = CType(.Controls.Add( _
                                  Office.MsoControlType.msoControlButton),  _
                                  Office.CommandBarButton)
            'Configure the created new button.
            With ctlFirstAddin
                .Caption = msBUTTON1
                .FaceId = 71
                .Parameter = msBUTTON1
                .Style = Office.MsoButtonStyle.msoButtonIconAndCaption
                .Tag = msTAG
                .TooltipText = msBUTTON1
            End With

            'Add an additional new button.
            ctlFirstAddin = CType(.Controls.Add( _
                                  Office.MsoControlType.msoControlButton),  _
                                  Office.CommandBarButton)
            'Configure the created new button.
            With ctlFirstAddin
                .BeginGroup = True
                .Caption = msBUTTON2
                .FaceId = 72
                .Parameter = msBUTTON2
                .Style = Office.MsoButtonStyle.msoButtonIconAndCaption
                .Tag = msTAG
                .TooltipText = msBUTTON2
            End With

            'Add an additional new button.
            ctlFirstAddin = CType(.Controls.Add( _
                                  Office.MsoControlType.msoControlButton),  _
                                  Office.CommandBarButton)
            'Configure the created new button.
            With ctlFirstAddin
                .BeginGroup = True
                .FaceId = 73
                .Caption = msBUTTON3
                .Parameter = msBUTTON3
                .Style = Office.MsoButtonStyle.msoButtonIconAndCaption
                .Tag = msTAG
                .TooltipText = msBUTTON3
            End With
            'Position of the command bar.
            .Position = Office.MsoBarPosition.msoBarTop
            'Make the command bar visible.
            .Visible = True
        End With

        'Set the WithEvents to hook the created buttons, all controls that
        'have the same Tag property will fire the mcmdPetrasButton_Click
        'event.
        mcmdFirstAddinButton = ctlFirstAddin

        'Release of the button control object.
        If (ctlFirstAddin IsNot Nothing) Then ctlFirstAddin = Nothing

    End Sub

    Private Sub cmdFirstAddinButton_Click( _
                        ByVal Ctrl As Office.CommandBarButton, _
                        ByRef CancelDefault As Boolean) _
                        Handles mcmdFirstAddinButton.Click

        'Customized error message.
        Const sERROR_MESSAGE As String = _
              "Cannot execute the wanted action."

        Dim sTEXT As String = "You clicked on "

        Try
            'Make sure it is one of ours controls.
            If Ctrl.Tag = msTAG Then

                'Run the appropriate message.
                Select Case Ctrl.Parameter
                    Case msBUTTON1
                        MessageBox.Show(text:=sTEXT & msBUTTON1, _
                                        caption:=msTITLE)
                    Case msBUTTON2
                        MessageBox.Show(text:=sTEXT & msBUTTON2, _
                                        caption:=msTITLE)
                    Case msBUTTON3
                        MessageBox.Show(text:=sTEXT & msBUTTON3, _
                                        caption:=msTITLE)
                End Select

            End If

            'We handled the event, so cancel its default behavior.
            CancelDefault = True

        Catch Generalex As Exception

            'Show the customized message.
            MessageBox.Show(text:=sERROR_MESSAGE, _
                            caption:=msTITLE, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

        End Try

    End Sub

    Private Sub Delete_Commandbar()

        'The check button control variable.
        Dim ctlCheck As Office.CommandBarButton = Nothing

        'If the command bar exists then get the first control 
        'from the command bar. 
        ctlCheck = CType(mxlApp.CommandBars.FindControl(Tag:=msTAG),  _
                                           Office.CommandBarButton)

        'If the command bar exists then delete it.
        If (ctlCheck IsNot Nothing) Then
            mxlApp.CommandBars(msTITLE).Delete()
        End If

    End Sub

#End Region

#Region "Release all objects"

    Private Sub Release_All_COMObjects(ByVal oxlObject As Object)

        Try
            Marshal.ReleaseComObject(oxlObject)
            oxlObject = Nothing
        Catch ex As Exception
            oxlObject = Nothing
        End Try

    End Sub

#End Region

#Region "Necessary Procedures"

    Public Sub OnBeginShutdown(ByRef custom As System.Array) _
                               Implements IDTExtensibility2.OnBeginShutdown
        'This procedure is required.
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) _
                              Implements IDTExtensibility2.OnAddInsUpdate
        'This procedure is required.
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) _
                                 Implements IDTExtensibility2.OnStartupComplete
        'This procedure is required.
    End Sub

#End Region

End Class
