'
' Description:  This class module contains all the code to interact
'               with the Windows Form's UI and user selections.
'
' Authors:      Dennis Wallentin, www.excelkb.com
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Imported namespaces
'
'To use the native .NET Messagebox object.
Imports System.Windows.Forms
'To use the StringBuilder.
Imports System.Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Class frmMain

    'The default date for the control dtpStartDate.
    Private mdtStartDate As Date = DateTime.Now.AddMonths(-2)

    'The default date for the control dtpEndDate and at the
    'same time the latest possible date to include for the
    'control.
    Private mdtEndMaxDate As Date = Today

    'A variable that indicates if the initially loading of
    'client's names has been done or not.
    Private mbLoadClientsList As Boolean = False

    'A variable that is used as a flag if we have 
    'successfully created a connection to the database.
    Private mbIsConnected As Boolean = Nothing

    'A datatable variable to be used for populating the 
    'cboClients control with clients names.
    Dim mdtClients As DataTable = Nothing

    'A datatable variable to be used for populating the 
    'datagridview control with data and to export data.
    Private mdtTable As DataTable = Nothing

    'Variables shared by several procedures.
    Private msSelectedClient As String = Nothing
    Private msSelectedProject As String
    Private msSelectedStartDate As String = Nothing
    Private msSelectedEndDate As String = Nothing

    'The variable for the connecting Windows Form.
    Dim mfrmConnecting As frmConnecting

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure load the main Windows Form.
    '           It populates the client's name combobox and manipulate
    '           other controls.
    '           
    '
    ' Arguments:    sender              System.Object
    '               e                   System.EventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version
    ' 04/23/08      Dennis Wallentin    Ch24    Added tooltip
    ' 05/13/08      Dennis Wallentin    Ch24    Added Background-
    '                                           worker component                                            

    Private Sub Form1_Load(ByVal sender As System.Object, _
                           ByVal e As System.EventArgs) _
                           Handles MyBase.Load

        'The help file in use.
        Const sHELPNAMESPACE As String = "PETRAS_Report_Tool.chm"

        'Constants values for all the controls's tooltips.
        Const sDGVTIP As String = "Shows a preview of the report."
        Const sCBOCLIENT As String = "Select the client the report will target."
        Const sCBOPROJECT As String = "Select the project the report will target."
        Const sSTARTDATE As String = "Select the start date for the report to cover from."
        Const sENDDATE As String = "Select the end date for the report to include."
        Const sCONSULTANTS As String = "Check this control if the report should " + vbNewLine + _
                                       "include the consultants."
        Const sACTIVITIES As String = "Check this control if the report should " + vbNewLine + _
                                      "include the activities."
        Const sCLEARSETTINGS As String = "Resets all controls to their default settings."
        Const sCREATEREPORT As String = "Creates the report based on the above selections."
        Const sLINK As String = "Open the web browser and browse to XL-Dennis homepage."
        Const sEXPORTEXCEL As String = "Export the created report to an Excel workbook."
        Const sEXPORTXML As String = "Save the created report to a XML file."
        Const sCLOSE As String = "Closes the program."


        'The start date cannot be older then 13 months back for
        'the dtpStartDate control.
        Dim dtStartMinDate As Date = DateTime.Now.AddMonths(-13)

        'The start date cannot be newer then one month backwards
        'from now for the dtpStartDate control.
        Dim dtStartMaxDate As Date = DateTime.Now.AddMonths(-1)

        'The end date cannot be earlier then 29 days from now
        'for the dtpEndDate control.
        Dim dtEndMinDate As Date = DateTime.Now.AddDays(-29)

        'Manipulation of some of the Windows Form's main properties.
        With Me
            .CancelButton = cmdClose
            .FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            .Icon = My.Resources.PetrasIcon
            .MaximizeBox = False
            .MinimizeBox = False
            .Text = swsCaption
        End With

        'If one of the controls cboClients or cboProjects has focus
        'when the user clicks on the Clear Settings's button then the
        'validation event of the control is fired and shows the error
        'provider. To prevent it the property CausesValidation of the 
        'command is set to false.  
        Me.cmdClearSettings.CausesValidation = False

        'Initially settings for the two combobox controls.
        With Me
            .cboClients.DropDownStyle = ComboBoxStyle.DropDownList
            .cboProjects.DropDownStyle = ComboBoxStyle.DropDownList
        End With

        'Initially settings for the two checkbox controls.
        With Me
            .chkActivities.Checked = True
            .chkConsultants.Checked = True
        End With

        'Manipulate the two DateTime Picker Controls.
        With Me.dtpStartDate
            .Format = DateTimePickerFormat.Short
            .Value = mdtStartDate
            .MinDate = dtStartMinDate
            .MaxDate = dtStartMaxDate
        End With

        With Me.dtpEndDate
            .Format = DateTimePickerFormat.Short
            .Value = mdtEndMaxDate
            .MinDate = dtEndMinDate
            .MaxDate = mdtEndMaxDate
        End With

        'The two export buttons controls should only be enabled when the
        'user have created a report.
        Me.Enable_CmdButtons(bEnable:=False)

        'Some settings for the DataGridView control.
        With Me.dgvReport
            .AllowUserToOrderColumns = True
            .AllowUserToResizeColumns = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
        End With

        'Add tooltips to the controls.
        With Me.ToolTip1
            .SetToolTip(control:=Me.dgvReport, caption:=sDGVTIP)
            .SetToolTip(control:=Me.cboClients, caption:=sCBOCLIENT)
            .SetToolTip(control:=Me.cboProjects, caption:=sCBOPROJECT)
            .SetToolTip(control:=Me.dtpStartDate, caption:=sSTARTDATE)
            .SetToolTip(control:=Me.dtpEndDate, caption:=sENDDATE)
            .SetToolTip(control:=Me.chkConsultants, caption:=sCONSULTANTS)
            .SetToolTip(control:=Me.chkActivities, caption:=sACTIVITIES)
            .SetToolTip(control:=Me.cmdClearSettings, caption:=sCLEARSETTINGS)
            .SetToolTip(control:=Me.cmdCreateReport, caption:=sCREATEREPORT)
            .SetToolTip(control:=Me.llblBook, caption:=sLINK)
            .SetToolTip(control:=Me.cmdExportExcel, caption:=sEXPORTEXCEL)
            .SetToolTip(control:=Me.cmdExportXML, caption:=sEXPORTXML)
            .SetToolTip(control:=Me.cmdClose, caption:=sCLOSE)
        End With

        'To invoke help use the F1-button when the Windows Form has been loaded. 

        'Setting the helpfile to the HelpProvider component.
        Me.HelpProvider1.HelpNamespace = swsPath + sHELPNAMESPACE

        'To add a simple Form-based help it is much easier to do it at design-time
        'then at runtime by using code. 

        'For this application we add the following:
        'Property: Helpkeyword on HelpProvider1
        'Value: "About.htm"
        'Property: HelpNavigator on HelpProvider
        'Value: Topic

        'Settings for the BackgroundWorker component.
        With Me.BackgroundWorker1
            'Makes it possible to cancel the operation.
            .WorkerSupportsCancellation = True
            'Start the background execution.
            .RunWorkerAsync()
        End With

        'Change the cursor while waiting to BackgroundWorker component
        'has been finished.
        Me.Cursor = Cursors.WaitCursor

    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure tries to etablish a connection to
    '           the database when the Windows Form is initially 
    '           loaded.
    '
    ' Arguments:    sender              Object
    '               e                   System.ComponentModel.DoWorkEventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 05/13/08      Dennis Wallentin    Ch24    Initial version

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, _
                                         ByVal e As System.ComponentModel.DoWorkEventArgs) _
                                         Handles BackgroundWorker1.DoWork

        'Instantiate a new instance of the connecting Windows Form.
        mfrmConnecting = New frmConnecting

        'Position the Windows Form and display it.
        With mfrmConnecting
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
        End With

        'Can we connect to the database?
        If MDataReports.bConnect_Database() = False Then

            'OK, we cannot establish a connection to the database
            'so we cancel the background operation.
            Me.BackgroundWorker1.CancelAsync()

            'Let us tell it for the other backgroundWorker event -
            'RunWorkerCompleted.
            mbIsConnected = False

        Else

            'Let us tell it for the other backgroundWorker event -
            'RunWorkerCompleted.
            mbIsConnected = True

        End If

        'Close the connecting Windows Form.
        mfrmConnecting.Close()

        'Releases all resources the variable has consumed from
        'the memory.
        mfrmConnecting.Dispose()

        'Release the reference the variable holds and prepare it
        'to be collected by the Garbage Collector (GC) when it
        'next time comes around.
        mfrmConnecting = Nothing

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure is executed when the backgroundWorker1
    '           DoWork events is done.
    '
    ' Arguments:    sender              Object
    '               e                   System.ComponentModel.
    '                                   RunWorkerCompletedEventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 05/13/08      Dennis Wallentin    Ch24    Initial version

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, _
                                                     ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) _
                                                     Handles BackgroundWorker1.RunWorkerCompleted


        'If we have managed to connect to the database then we can continue.
        If mbIsConnected Then

            'Populate the datatable with clients names.
            mdtClients = MDataReports.dtGet_Clients

            'If not the returned value is nothing we continue.
            If Not IsNothing(Expression:=mdtClients) Then

                'Populate the client's combobox with the clients's names.
                With Me.cboClients
                    .DataSource = mdtClients
                    .DisplayMember = "ClientName"
                    .SelectedIndex = -1
                End With

                'We are done with loading the clients names so we set the flag
                'to true.
                mbLoadClientsList = True

            End If

        End If

        'Restore the cursor.
        Me.Cursor = Cursors.Default

    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure load the list of projects's names based
    '           on user's selection of client name.    
    '
    ' Arguments:    sender              System.Object
    '               e                   System.EventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version

    Private Sub cboClients_SelectedIndexChanged(ByVal sender As System.Object, _
                                                ByVal e As System.EventArgs) _
                                                Handles cboClients.SelectedIndexChanged

        'Check to see if the clients's list has been loaded or not.
        If mbLoadClientsList = True Then

            'Check if any selection of client has been made.
            If Me.cboClients.SelectedIndex >= 0 Then

                'Retrieve the selected client's name.
                Dim drvRow As DataRowView = CType(Me.cboClients.SelectedItem, DataRowView)

                'Populate the list for available projects.
                With Me.cboProjects
                    .DataSource = MDataReports.dtGetProjects_Client( _
                                                iClient:=Convert.ToInt32(drvRow("ClientID")))
                    .DisplayMember = "ProjectName"
                    .SelectedIndex = -1
                End With

            End If

        End If

    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure populate the datagridview with data.               
    '
    ' Arguments:    sender              System.Object
    '               e                   System.EventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/23/08      Dennis Wallentin    Ch24    Initial version
    ' 04/24/08      Dennis Wallentin    Ch24    Added additional options
    ' 04/27/08      Dennis Wallentin    Ch24    Added additional options

    Private Sub cmdCreateReport_Click(ByVal sender As Object, _
                                      ByVal e As System.EventArgs) _
                                      Handles cmdCreateReport.Click

        'Customized messages.
        Const sMESSAGENOSELECTION As String = "You must select both a client " + vbNewLine + _
                                              "name and a project name."
        Const sMESSAGENORECORDS As String = "No records were found that match the criterias."


        'If both a client and a project is selected then we can continue.
        If Me.cboClients.SelectedIndex <> -1 And Me.cboProjects.SelectedIndex <> -1 Then

            'Get the selected project.
            Dim drvSelectedProject As DataRowView = CType(Me.cboProjects.SelectedItem, DataRowView)

            'Get the selected project ID.
            Dim iProjectID As Integer = Convert.ToInt32(value:=drvSelectedProject("ProjectID"))

            'Variable for the datagridview control in use.
            Dim dgvColumn As DataGridViewColumn = Nothing

            'Prepare the datagridview for new data.
            With Me.dgvReport
                .DataSource = Nothing
                .Refresh()
            End With

            'Check the status for the two checkbox controls which control to include or
            'not the activities's field and the consultants's field.
            If Me.chkConsultants.Checked = False And Me.chkActivities.Checked = True Then

                'Include only the activities.
                mdtTable = MDataReports.dtActivities(lnProjID:=iProjectID, _
                                                     dtStart:=Me.dtpStartDate.Value.Date, _
                                                     dtEnd:=Me.dtpEndDate.Value.Date)
            Else

                'Include both the activities and consultants.
                If Me.chkConsultants.Checked = True And Me.chkActivities.Checked = True Then
                    mdtTable = MDataReports.dtActivities_Consultants(lnProjId:=iProjectID, _
                                                                     dtStart:=Me.dtpStartDate.Value.Date, _
                                                                     dtEnd:=Me.dtpEndDate.Value.Date)
                Else

                    'Include only the consultants.
                    If Me.chkConsultants.Checked = True And Me.chkActivities.Checked = False Then
                        mdtTable = MDataReports.dtConsultants(lnProjId:=iProjectID, _
                                                              dtStart:=Me.dtpStartDate.Value.Date, _
                                                              dtEnd:=Me.dtpEndDate.Value.Date)
                    Else

                        'Exclude both the activities and consultants.
                        mdtTable = MDataReports.dtSummary(lnProjId:=iProjectID, _
                                                          dtStart:=Me.dtpStartDate.Value.Date, _
                                                          dtEnd:=Me.dtpEndDate.Value.Date)
                    End If
                End If
            End If

            'Make sure that we get some data.
            If Not IsNothing(mdtTable) Then

                'Populate the datagridview and hide some columns.
                With Me.dgvReport
                    .DataSource = mdtTable
                    'The following columns are included when exporting to XML-files but they
                    'are not relevant to preview.
                    .Columns("ClientID").Visible = False
                    .Columns("ClientName").Visible = False
                    .Columns("ProjectID").Visible = False
                    .Columns("ProjectName").Visible = False
                    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                End With

                'Split the colum names before adding them to the datagridview.
                For Each dgvColumn In dgvReport.Columns
                    dgvColumn.HeaderCell.Value = sSplit_Column_Names(sColumnName:=dgvColumn.HeaderText)
                Next

                'Make the two export buttons availably.
                Me.Enable_CmdButtons(bEnable:=True)


            Else

                'Show the customized message.
                MessageBox.Show(text:=sMESSAGENORECORDS, _
                        caption:=swsCaption, _
                        buttons:=MessageBoxButtons.OK, _
                        icon:=MessageBoxIcon.Warning)

                'Disable the two export controls.                
                Me.Enable_CmdButtons(bEnable:=False)

            End If

        Else

            'Show customized message.
            MessageBox.Show(text:=sMESSAGENOSELECTION, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Warning)

            'Set focus on the control that has no selection
            'made.
            If Me.cboClients.SelectedIndex = -1 Then
                Me.cboClients.Focus()
            Else
                Me.cboProjects.Focus()
            End If
        End If


    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure collect required data to export to
    '           Excel and invoke the export.
    '
    ' Arguments:    sender              System.Object
    '               e                   System.EventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/30/08      Dennis Wallentin    Ch24    Initial version
    ' 05/01/08      Dennis Wallentin    Ch24    Check Excel is installed
    ' 05/02/08      Dennis Wallentin    Ch24    Check if template exists or not
    ' 05/08/08      Dennis Wallentin    Ch24    Added the enumerations
    '

    Private Sub cmdExportExcel_Click(ByVal sender As System.Object, _
                                     ByVal e As System.EventArgs) _
                                     Handles cmdExportExcel.Click

        'Customized messages.
        Const sMESSAGENOEXCEL As String = "It appears that Excel is not " + vbNewLine + _
                                          "installed on this computer."

        Const sMESSAGEWRONGVERSION As String = "Version 10 and later of Excel must" + vbNewLine + _
                                               "be installed in order to proceed."

        'Retrieve the required values for creating the report.
        With Me
            msSelectedClient = .cboClients.Text
            msSelectedProject = .cboProjects.Text
            msSelectedStartDate = .dtpStartDate.Value.Date.ToString("MMddyyyy")
            msSelectedEndDate = .dtpEndDate.Value.Date.ToString("MMddyyyy")
        End With

        'Clear the solutions wide variable.
        swshSelectedTemplate = Nothing

        'Check to see that Excel 2002 or later is installed on the computer.
        Dim shInstalled As Short = MExportExcel.shCheck_Excel_Version_Installed

        Select Case shInstalled
            Case xlVersion.NoVersion

                'Customized message that Excel is not installed.
                MessageBox.Show(text:=sMESSAGENOEXCEL, _
                                caption:=swsCaption, _
                                buttons:=MessageBoxButtons.OK, _
                                icon:=MessageBoxIcon.Stop)

            Case xlVersion.WrongVersion

                'Customized message that the wrong Excel version is installed.
                MessageBox.Show(text:=sMESSAGEWRONGVERSION, _
                                caption:=swsCaption, _
                                buttons:=MessageBoxButtons.OK, _
                                icon:=MessageBoxIcon.Stop)

            Case xlVersion.RightVersion

                'The selected options controls which template to use.
                If Me.chkActivities.Checked And Me.chkConsultants.Checked Then

                    swshSelectedTemplate = xltTemplate.xltActivitiesConsultants

                ElseIf Me.chkActivities.Checked And Not Me.chkConsultants.Checked Then

                    swshSelectedTemplate = xltTemplate.xlActivities

                ElseIf Not Me.chkActivities.Checked And Me.chkConsultants.Checked Then

                    swshSelectedTemplate = xltTemplate.xltConsultants

                Else

                    swshSelectedTemplate = xltTemplate.xltSummary

                End If

                'Check to see that the selected Excel template file exists or not.                
                If MExportExcel.bFile_Exist() Then

                    'Call the export to Excel function.
                    If MExportExcel.bExport_Excel(dtTable:=mdtTable, _
                                                 sClient:=msSelectedClient, _
                                                 sProject:=msSelectedProject, _
                                                 sStartDate:=msSelectedStartDate, _
                                                 sEndDate:=msSelectedEndDate) Then

                        'Call sub procedure which sets all controls to their default values.
                        Me.Restore_Settings()
                    End If


                Else

                    MessageBox.Show(text:="Please make sure that the Excel template file " + vbNewLine + _
                                    swsPath + "\" + swTemplateListArr(swshSelectedTemplate) + vbNewLine + _
                                    "actually exists.", _
                                    caption:=swsCaption, _
                                    buttons:=MessageBoxButtons.OK, _
                                    icon:=MessageBoxIcon.Stop)



                End If

        End Select

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure collect required data to export to
    '           an XML file.
    '
    ' Arguments:    sender              System.Object
    '               e                   System.EventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version

    Private Sub cmdExportXML_Click(ByVal sender As System.Object, _
                                   ByVal e As System.EventArgs) _
                                   Handles cmdExportXML.Click

        'Customized message.
        Const sMESSAGESAVEDFILES As String = "Files have successfully been saved."

        'Retrieve the required values for creating the report.
        With Me
            msSelectedClient = .cboClients.Text
            msSelectedProject = .cboProjects.Text
            msSelectedStartDate = .dtpStartDate.Value.Date.ToString("MMddyyyy")
            msSelectedEndDate = .dtpEndDate.Value.Date.ToString("MMddyyyy")
        End With

        'Call the export to a XML file function.
        If MExportXML.bSave_DataTable_XML( _
                            dtTable:=mdtTable, sClient:=msSelectedClient, sProject:=msSelectedProject, _
                            sStartDate:=msSelectedStartDate, sEndDate:=msSelectedEndDate) Then

            'Show the customized message.
            MessageBox.Show(text:=sMESSAGESAVEDFILES, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Information)

            'Set all controls to their default values.
            Me.Restore_Settings()
        End If

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure calls the a function to restore all
    '           controls to their default values.
    '
    ' Arguments:    sender              System.Object
    '               e                   System.EventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version

    Private Sub cmdClearSettings_Click(ByVal sender As System.Object, _
                                       ByVal e As System.EventArgs) _
                                       Handles cmdClearSettings.Click

        'Call the sub procedure to reset the controls to their default 
        'settings.
        Restore_Settings()

    End Sub


#Region "General procedures for the Windows Form."

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure controls the error provider for both the
    '           client's control and the project's control.
    '
    ' Arguments:    sender              System.Object
    '               e                   System.EventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version

    Private Sub Client_Project_Validating(ByVal sender As Object, _
                                          ByVal e As System.ComponentModel.CancelEventArgs) _
                                          Handles cboClients.Validating, _
                                          cboProjects.Validating

        'This event procedure shows how we can work with several controls in a 'control array'.
        'We use an associated name for the controls we target, in this case Client_Project_
        'Validating.
        'We hook the controls's validating events to this subroutine.

        'Customized message for the clients's control.
        Const sMESSAGECLIENTERROR As String = "You need to select a client."

        'Customized message for the projects's control.
        Const sMESSAGEPROJECTERROR As String = "You need to select a project."

        'The variable that holds the control in use.
        Dim Ctrl As Control = CType(sender, Control)

        If Ctrl.Text = "" Then

            Select Case Ctrl.Name

                Case "cboClients"

                    'Show the error provider with the message for the cboClients control.
                    Me.ErrorProvider1.SetError(control:=Ctrl, value:=sMESSAGECLIENTERROR)

                Case Else

                    'Show the error provider with the message for the cboProjects control.
                    Me.ErrorProvider1.SetError(control:=Ctrl, value:=sMESSAGEPROJECTERROR)

            End Select

        Else

            'Restore the error provider to its default value.

            Me.ErrorProvider1.SetError(control:=Ctrl, value:="")

        End If

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure change the enable status for the
    '           two export command buttons.
    '
    ' Arguments:    bEnable             Boolean 
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version

    Private Sub Enable_CmdButtons(ByVal bEnable As Boolean)

        'An array which includes the button control we target.
        Dim cmdControlsArr As Button() = {cmdExportExcel, cmdExportXML}

        Dim cmdButton As Button = Nothing

        'Iteration to either make the control enabled or disabled.
        For Each cmdButton In cmdControlsArr

            If bEnable Then

                cmdButton.Enabled = True

            Else

                cmdButton.Enabled = False

            End If

        Next

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function split the column names into separate words
    '           by inserting a space before each bew capital letter.
    '
    ' Arguments:    SColumnName         String
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version

    Private Function sSplit_Column_Names(ByVal sColumnName As String) As String

        'Variable which will hold one character at the time.
        Dim chCharacter As Char = Nothing

        'Variable that keep tracks of the number of character each string
        'expression contain.
        Dim iCharacterCounter As Integer = Nothing

        'The StringBuilder class provides the Append method, that inserts a new
        'string to an existing string. It allows us to split and rebuild the 
        'column names in a smooth way.
        Dim sbStringBuilder As New System.Text.StringBuilder

        'Add the first character.
        sbStringBuilder.Append(value:=sColumnName(index:=0))

        'Iterate through each string expression. Since we already have added
        'the first character and the last character will not be examine the
        'length of each string expression are reduced with 2.
        For iCharacterCounter = 1 To sColumnName.Length - 2

            chCharacter = sColumnName(iCharacterCounter)

            'Check if the character is a space or not.
            If chCharacter = " " Then

                'If the character is a space we just add it.
                sbStringBuilder.Append(value:=chCharacter)
                iCharacterCounter += 1
                sbStringBuilder.Append(value:=Char.ToUpper(sColumnName(iCharacterCounter)))

            Else

                'If the character is in uppercase we first add a space and then
                'we add the character to the string.
                If Char.IsUpper(chCharacter) Then

                    sbStringBuilder.Append(value:=" ")

                End If

                'Add the character to the string.
                sbStringBuilder.Append(value:=chCharacter)

            End If

        Next

        'Finally we add the last character.
        sbStringBuilder.Append(value:=sColumnName(index:=sColumnName.Length - 1))

        'OK, we are done so back to the calling procedure.
        Return sbStringBuilder.ToString

    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure restore the controls to their default
    '           values.
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version


    Private Sub Restore_Settings()

        'Restore the controls to their default values.
        With Me
            .dtpStartDate.Value = mdtStartDate
            .dtpEndDate.Value = mdtEndMaxDate
            .cboClients.SelectedIndex = -1
            .cboProjects.SelectedIndex = -1
            .chkConsultants.Checked = True
            .chkActivities.Checked = True
            .ErrorProvider1.SetError(control:=Me.cboClients, value:="")
            .ErrorProvider1.SetError(control:=Me.cboProjects, value:="")

            Me.Enable_CmdButtons(bEnable:=False)

            With .dgvReport
                'By setting the source to nothing the datagridview becomes empty.
                .DataSource = Nothing
                .Refresh()
            End With

        End With


        If mdtTable IsNot Nothing Then

            'Releases all resources the variable has consumed from
            'the memory.
            mdtTable.Dispose()

            'Release the reference the variable holds and prepare it
            'to be collected by the Garbage Collector (GC) when it
            'comes around.

            mdtTable = Nothing
        End If

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure close the Windows Form and release
    '           resources.
    '
    ' Arguments:    sender              System.Object
    '               e                   System.Windows.Forms.
    '                                   LinkLabelLinkClickedEventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 05/11/08      Dennis Wallentin    Ch24    Initial version

    Private Sub llblBook_LinkClicked(ByVal sender As System.Object, _
                                     ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) _
                                     Handles llblBook.LinkClicked

        'Call the procedure that starts up the browser and show the homepage.
        MCommonFunctions.Visit_Link()

    End Sub

#End Region

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This procedure close the Windows Form and release
    '           resources.
    '
    ' Arguments:    sender              System.Object
    '               e                   System.EventArgs
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version

    Private Sub cmdClose_Click(ByVal sender As System.Object, _
                               ByVal e As System.EventArgs) _
                               Handles cmdClose.Click


        If Not IsNothing(Expression:=mdtClients) Then

            'Releases all resources the variable has consumed from
            'the memory.
            mdtClients.Dispose()

            'Release the reference the variable holds and prepare it
            'to be collected by the Garbage Collector (GC) when it
            'comes around.
            mdtClients = Nothing

        End If


        If Not IsNothing(Expression:=mdtTable) Then

            mdtTable.Dispose()
            mdtTable = Nothing

        End If

        If Not IsNothing(Expression:=mfrmConnecting) Then

            mfrmConnecting.Dispose()
            mfrmConnecting = Nothing

        End If

        Me.Close()

    End Sub


End Class