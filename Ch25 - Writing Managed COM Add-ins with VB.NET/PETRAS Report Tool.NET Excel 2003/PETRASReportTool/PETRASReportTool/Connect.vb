'
' Description:  This main class for connecting and disconneting the
'               add-in, to build and tear down the custom menu in Excel.     
'
' Authors:      Dennis Wallentin, www.excelkb.com
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Imported namespaces
'
Imports Extensibility
Imports System.Runtime.InteropServices
Imports Office = Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel


#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the $SAFEOBJNAME$Setup project, 
' right click the project in the Solution Explorer, then choose install.
#End Region

<GuidAttribute("9D5657E0-301C-40F6-8D72-B4F9F003C2A2"), _
 ProgIdAttribute("PETRASReportTool.Connect")> _
Public Class Connect

#Region "Implements"

    Implements IDTExtensibility2

#End Region

#Region "Module-wide variables"

    'The hooking command bar control event.
    Private WithEvents mcmdPetrasButton As  _
                       Office.CommandBarButton = Nothing

    'The tag variable for our control.
    Private Const msTAG As String = "PETRASNET"

    'Parameter variables for the two buttons on the
    'custom menu.
    Private Const msABOUT As String = "&About"
    Private Const msREPORT As String = "&Report"

#End Region

#Region "Connection and Disconnection"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function connect the add-in to Excel.
    '           
    ' Arguments:    
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/11/08      Dennis Wallentin    Ch25    Initial version

    Public Sub OnConnection(ByVal application As Object, _
                            ByVal connectMode As ext_ConnectMode, _
                            ByVal addInInst As Object, _
                            ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

        'This is used upon checking if right version of Excel is installed or not.
        Const sMESSAGEWRONGVERSION As String = "Version 2002 and later of Excel must" + vbNewLine + _
                                                 "be installed in order to proceed."

        'Customized error message.
        Const sERROR_MESSAGE As String = _
              "An unexpected error has occured."

        'Create and instantiate a new instance of the class.
        Dim cMethods As New CCommonMethods()

        Try
            'Get a reference to Excel.
            swXLApp = (CType(application, Excel.Application))

            'Check to see that Excel 2002 or later is installed on the computer.
            Dim shInstalled As Short = cMethods.shCheck_Excel_Version_Installed

            If shInstalled = xlVersion.WrongVersion Then

                'Customized message that the wrong Excel version is installed.
                MessageBox.Show(text:=sMESSAGEWRONGVERSION, _
                                caption:=swsCaption, _
                                buttons:=MessageBoxButtons.OK, _
                                icon:=MessageBoxIcon.Stop)

                Exit Try

            End If

            'Create the custom menu.
            Create_Tool_Menu()

            'The following lines may be necessary to add in order to support
            'Windows XP visual style.
            System.Windows.Forms.Application.EnableVisualStyles()
            System.Windows.Forms.Application.DoEvents()

        Catch Generalex As Exception

            'Show the customized message.
            MessageBox.Show(text:=sERROR_MESSAGE, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

        Finally

            'Prepare the object for GC.
            If Not IsNothing(Expression:=cMethods) Then cMethods = Nothing

        End Try

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function disconnect the add-in from Excel.
    '           
    ' Arguments:    
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/11/08      Dennis Wallentin    Ch25    Initial version

    Public Sub OnDisconnection(ByVal RemoveMode As ext_DisconnectMode, _
                               ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection

        'Delete the custom menu.
        Delete_Tool_Menu()

        'Prepare the objects for GC. 
        If Not IsNothing(Expression:=mcmdPetrasButton) Then mcmdPetrasButton = Nothing
        swXLApp = Nothing

    End Sub

#End Region

#Region "Build and tear down the custom menu."

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function build the custom menu.
    '           
    ' Arguments:    None
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/12/08      Dennis Wallentin    Ch25    Initial version

    Friend Sub Create_Tool_Menu()

        Const sPETRASREPORT As String = "&Petras Report Tool.NET"
        Const sREPORTTOOLTIP As String = "Create a report for PETRAS."
        Const sABOUTTOOLTIP As String = "About PETRAS Report Tool.NET"

        'The commandbar variable.
        Dim cbrPetras As Office.CommandBar = Nothing

        'The Popup control variable.
        Dim ctlPetras As Office.CommandBarPopup = Nothing

        'The button control variable.
        Dim ctlReport As Office.CommandBarButton = Nothing

        'Grab the worksheet menu commandbar. 
        cbrPetras = CType(swXLApp.CommandBars(1), Office.CommandBar)

        'Can we find out control?
        ctlPetras = CType(cbrPetras.FindControl(Tag:=msTAG),  _
                           Microsoft.Office.Core.CommandBarPopup)

        'If the custom menu does not exist create it.
        If ctlPetras Is Nothing Then
            With cbrPetras

                'Add the popup control to the worksheet menu.
                ctlPetras = CType(.Controls.Add( _
                                   Office.MsoControlType.msoControlPopup),  _
                                   Office.CommandBarPopup)

                With ctlPetras

                    'Configure the created popup control.
                    .Caption = sPETRASREPORT
                    .Tag = msTAG

                    'Add a new button control.
                    ctlReport = CType(.Controls.Add( _
                                       Office.MsoControlType.msoControlButton),  _
                                       Office.CommandBarButton)

                    'Configure the added button control.
                    With ctlReport
                        .Caption = msREPORT
                        .FaceId = 162
                        .Parameter = msREPORT
                        .Style = Office.MsoButtonStyle.msoButtonIconAndCaption
                        .Tag = msTAG
                        .TooltipText = sREPORTTOOLTIP
                    End With

                    'Add a new button control.
                    ctlReport = CType(.Controls.Add( _
                                      Office.MsoControlType.msoControlButton),  _
                                      Office.CommandBarButton)

                    'Configure the added button control.
                    With ctlReport
                        .BeginGroup = True
                        .Caption = msABOUT
                        .FaceId = 487
                        .Parameter = msABOUT
                        .Style = Office.MsoButtonStyle.msoButtonIconAndCaption
                        .Tag = msTAG
                        .TooltipText = sABOUTTOOLTIP
                    End With
                End With

            End With

            'Set the WithEvents to hook the created buttons, all controls that
            'have the same Tag property will fire the mcmdPetrasButton_Click
            'event.
            mcmdPetrasButton = ctlReport

        Else

            'Re-hooking the WithEvents to the custom buttons. 
            mcmdPetrasButton = CType(swXLApp.CommandBars.FindControls(Tag:=msTAG).Item(2),  _
                                     Microsoft.Office.Core.CommandBarButton)

        End If

        'Prepare the objects for GC.
        If (ctlReport IsNot Nothing) Then ctlReport = Nothing
        If (ctlPetras IsNot Nothing) Then ctlPetras = Nothing
        If (cbrPetras IsNot Nothing) Then cbrPetras = Nothing

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function tear down the custom menu.
    '           
    ' Arguments:    None
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/12/08      Dennis Wallentin    Ch25    Initial version

    Friend Sub Delete_Tool_Menu()

        Dim octlItem As Office.CommandBarControl = Nothing

        'Find and delete our controls.
        For Each octlItem In swXLApp.CommandBars.FindControls(Tag:=msTAG)
            octlItem.Delete()
        Next

        'Prepare object for GC.
        If (octlItem IsNot Nothing) Then octlItem = Nothing

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments:     This subroutine handle all buttons click event 
    '               in our custom menu.
    ' Arguments:    None
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/12/08      Dennis Wallentin    Ch25    Initial version

    Private Sub mcmdPetrasButton_Click(ByVal Ctrl As Office.CommandBarButton, _
                                       ByRef CancelDefault As Boolean) Handles mcmdPetrasButton.Click

        'Create and instantiate a new instance of the class.
        Dim cCMethods As New CCommonMethods

        Select Case Ctrl.Parameter

            'User selected to show the Report form.
            Case msREPORT : cCMethods.Load_Form(sForm:=msREPORT)

                'User select to show the About form.
            Case msABOUT : cCMethods.Load_Form(sForm:=msABOUT)

        End Select

        'Prepare object for GC.
        If (cCMethods IsNot Nothing) Then cCMethods = Nothing

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
