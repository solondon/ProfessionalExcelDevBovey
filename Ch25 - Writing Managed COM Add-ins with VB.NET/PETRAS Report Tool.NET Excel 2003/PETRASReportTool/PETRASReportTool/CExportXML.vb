'
' Description:  This class contains the main function to export to
'               XML files.  
'
' Authors:      Dennis Wallentin, www.excelkb.com
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Class CExportXML

#Region "Export data to XML file."

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function allow users to save selected project's
    '           data into a XML file and at the same time it automatically
    '           createst a XDS file with the same file name as the XML
    '           file.
    '           
    ' Arguments:    dtTable     The in-memory datatable we get all data from.
    '               sClient     The selected client's name.
    '               sProject    The selected project's name.
    '               sStartDate  The selected start date for the report.
    '               sEndDate    The selected end date for the report.
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/12/08      Dennis Wallentin    Ch25    Initial version

    Friend Function bSave_DataTable_XML(ByVal dtTable As DataTable, _
                                        ByVal sClient As String, _
                                        ByVal sProject As String, _
                                        ByVal sStartDate As String, _
                                        ByVal sEndDate As String) As Boolean

        'Constant customized string expressions.
        Const sREPORTPREFIX As String = "PETRAS "

        Const sMESSAGENOTSAVEDXML As String = "An unexpected error occurred when trying " + vbNewLine + _
                                           "to create and save the files."

        Const sMESSAGENOTSAVEDGENERAL As String = "An unexpected error occurred in the application."

        'Create and instantiate a new instance of the class.
        Dim cCMethods As New CCommonMethods()

        'Creates the time period.
        Dim sPeriod As String = sStartDate + "-" + sEndDate

        'Creates the suggested file name to be saved.
        Dim sFileName As String = sREPORTPREFIX + cCMethods.sCreate_Name(sClientName:=sClient, _
                                                              sProjectName:=sProject) + " " + sPeriod

        'The boolean variable to return the output from the function.
        Dim bSaved As Boolean = Nothing

        'The variable for the main object.
        Dim frmSaveFile As SaveFileDialog = Nothing

        Try

            'Instantiate a new instance of the main object.
            frmSaveFile = New SaveFileDialog

            'Show the save file dialog.
            With frmSaveFile
                .Filter = "XML File|*.xml"
                .Title = "Save report to XML file"
                .FileName = sFileName
            End With

            'If user has not canceled we proceed.
            If frmSaveFile.ShowDialog = Windows.Forms.DialogResult.OK Then

                'In case the user has customized the file name we need to check
                'it.
                If sFileName <> frmSaveFile.FileName Then
                    sFileName = frmSaveFile.FileName
                End If

                'Write the data to the XML file.
                dtTable.WriteXml(fileName:=sFileName)

                'Create the Schema file for the XML file. 
                dtTable.WriteXmlSchema(fileName:=Strings.Left(sFileName, Len(sFileName) - 4) & ".xsd")

                bSaved = True

            End If


        Catch XMLexc As Xml.XmlException

            'Show customized message.
            MessageBox.Show(text:=sMESSAGENOTSAVEDXML, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

            'Things didn't worked out as we expected so we set the boolean 
            'value to false.
            bSaved = False


        Catch Generalexc As Exception

            'Show customized message.
            MessageBox.Show(text:=sMESSAGENOTSAVEDGENERAL, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

            'Things didn't worked out as we expected so we set the boolean 
            'value to false.
            bSaved = False

        Finally

            'Release all resources consumed by the variable from the
            'memory.
            frmSaveFile.Dispose()

            'Prepare the object for GC.
            If (frmSaveFile IsNot Nothing) Then frmSaveFile = Nothing
            If (cCMethods IsNot Nothing) Then cCMethods = Nothing

        End Try

        'Inform the calling procedure about the outcome.
        Return bSaved

    End Function

#End Region

End Class
