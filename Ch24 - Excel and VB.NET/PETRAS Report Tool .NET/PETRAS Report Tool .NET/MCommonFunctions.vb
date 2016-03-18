'
' Description:  This module contains the common functions to the
'               project.
'
' Authors:      Dennis Wallentin, www.excelkb.com
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Imported namespaces

Module MCommonFunctions

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function creates a project specific name which
    '           is used to add a relevant name to worksheets and as
    '           a suggestion for XML-files names.
    '
    ' Arguments:    sClientName     The selected client's name.
    '               sProjectName    The selected project's name
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 05/02/08      Dennis Wallentin    Ch24    Initial version

    Friend Function sCreate_Name(ByVal sClientName As String, _
                                 ByVal sProjectName As String) As String

        'Extract the first three charachters from the client's name as well as
        'from te project's name.
        Dim sName As String = Strings.Left(str:=sClientName, Length:=3) + " " + _
                              Strings.Left(str:=sProjectName, Length:=3)

        Return sName

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function loads the homepage in the browser.
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 05/11/08      Dennis Wallentin    Ch24    Initial version

    Friend Sub Visit_Link()

        Try
            'Initiate and load the homepage in the browser.
            Process.Start(fileName:="http://www.excelkb.com")

        Catch Generalexc As Exception

            MessageBox.Show(text:=Generalexc.ToString(), caption:=My.Application.Info.Title.ToString())

        End Try
    End Sub

End Module
