'
' Description:  This class contains the common methods to
'               the project.
'
' Authors:      Dennis Wallentin, www.excelkb.com
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Imported namespaces
'
'To check that files exist or not.
Imports System.IO
'To work with the assembly itself.
Imports System.Reflection
'To read the Windows Registry subkey.
Imports Microsoft.Win32
'To use regular expressions.
Imports System.Text.RegularExpressions

Public Class CCommonMethods

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This subroutine loads either the Report Main Form or 
    '           the About Form.
    '
    ' Arguments:    sForm
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/15/08      Dennis Wallentin    Ch25    Initial version

    Friend Sub Load_Form(ByVal sForm As String)

        'Customized error message.
        Const sERROR_MESSAGE As String = _
              "An unexpected error has occured."

        'Instantiate a new instance of the NativeWindow class. 
        Dim appWindow As New NativeWindow

        'A Windows Form variable.
        Dim frm As Form = Nothing

        Try
            'Which form to load?
            If sForm = "&Report" Then
                frm = New frmMain
            Else
                frm = New frmAbout
            End If

            'Assign a handle to Excel's Main Window.
            appWindow.AssignHandle(Process.GetCurrentProcess().MainWindowHandle)
            'Show the Windows Form and let Excel's Main Window be the owner of the form.
            If frm.ShowDialog(appWindow) = DialogResult.OK Then Exit Try

        Catch Generalex As Exception

            'Show the customized message.
            MessageBox.Show(text:=sERROR_MESSAGE, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

        Finally

            If (appWindow IsNot Nothing) Then

                'Release the handle.
                appWindow.ReleaseHandle()

                appWindow = Nothing

            End If

            If (frm IsNot Nothing) Then

                'Dispose the Windows Form class.
                frm.Dispose()
                frm = Nothing

            End If

        End Try

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function loads the homepage in the browser.
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/12/08      Dennis Wallentin    Ch25    Initial version

    Friend Sub Visit_Link()

        Try
            'Initiate and load the homepage in the browser.
            Process.Start(fileName:="http://www.excelkb.com")

        Catch Generalexc As Exception

            MessageBox.Show(text:=Generalexc.ToString(), caption:=My.Application.Info.Title.ToString())

        End Try
    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function checks if the selected template exists or
    '           not.
    '           
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 17/09/08      Dennis Wallentin    Ch25    Initial version

    Friend Function Get_Path() As String

        'This will get us the full path and the name of the assembly.
        Dim sFullpath As String = Assembly.GetExecutingAssembly().GetName().CodeBase.ToString()

        'Clean the string variable.
        If sFullpath.StartsWith("file") Then
            sFullpath = sFullpath.Substring(startIndex:=8)
        End If

        'Get the pathway.
        Dim sPath As String = Path.GetDirectoryName(sFullpath) + "\"

        Return sPath

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function checks if Excel is available or not.
    '           If available then it checks which version is installed.                   
    '           It can return one of the following values:
    '           NoVersion (0):      Excel is not installed.
    '           WrongVersion (1):   Wrong Excel version is available.
    '           RightVersion(2):    Right Excel version is installed, 
    '                               Excel 2002 and higher.                                '               
    '           Nothing:            An error has occurred.
    '
    ' Arguments:    None
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 05/01/08      Dennis Wallentin    Ch24    Initial version
    ' 09/13/08      Dennis Wallentin    Ch25    Added

    Friend Function shCheck_Excel_Version_Installed() As Short

        Const sERROR_MESSAGE As String = "An unexpected error has occured " + _
                                         "when trying to read the registry."

        'The subkey we are interested in is located in the HKEY_CLASSES_ROOT
        'Class.
        'The subkey's value looks like the following: Excel.Application.10
        Const sXL_SUBKEY As String = "\Excel.Application\CurVer"

        Dim rkVersionkey As RegistryKey = Nothing
        Dim sVersion As String = String.Empty
        Dim sXLVersion As String = String.Empty

        'The regular expression which is interpretated as:
        'Look for integer values in the intervall 8-9
        'in the end of the retrieved subkey's string value.
        Dim sRegExpr As String = "[8-9]$"

        Dim shStatus As Short = Nothing

        Try
            'Open the subkey.
            rkVersionkey = Registry.ClassesRoot.OpenSubKey(name:=sXL_SUBKEY, _
                                                           writable:=False)

            'If we cannot open the subkey then Excel is not available.
            If rkVersionkey Is Nothing Then
                shStatus = xlVersion.NoVersion
            End If

            'Excel is installed and we can retrieve the wanted information.
            sXLVersion = CStr(rkVersionkey.GetValue(name:=sVersion))

            'Compare the retrieved value with our defined regular expression.
            If Regex.IsMatch(input:=sXLVersion, pattern:=sRegExpr) Then
                'Excel 97 or Excel 2000 is installed.
                shStatus = xlVersion.WrongVersion
            Else
                'Excel 2002 or later is available.
                shStatus = xlVersion.RightVersion
            End If

        Catch Generalexc As Exception

            'Show the customized message.
            MessageBox.Show(text:=sERROR_MESSAGE, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

            'Things didn't worked out as we expected so we set the 
            'return variable to nothing.
            shStatus = Nothing

        Finally

            If rkVersionkey IsNot Nothing Then

                'We need to close the opened subkey.
                rkVersionkey.Close()

                'Prepare the object for GC.
                rkVersionkey = Nothing

            End If

        End Try

        'Inform the calling procedure about the outcome.
        Return shStatus

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function checks if the selected template exists or
    '           not.
    '           
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/16/08      Dennis Wallentin    Ch25    Initial version

    Friend Function bFile_Exist() As Boolean

        'Check if the file exists or not.
        If File.Exists(path:=Me.Get_Path() + swTemplateListArr(swshSelectedTemplate).ToString()) Then
            bFile_Exist = True
        Else
            bFile_Exist = False
        End If

    End Function

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
    ' 09/12/08      Dennis Wallentin    Ch25    Initial version

    Friend Function sCreate_Name(ByVal sClientName As String, _
                                 ByVal sProjectName As String) As String

        'Extract the first three charachters from the client's name as well as
        'from te project's name.
        Dim sName As String = Strings.Left(str:=sClientName, Length:=3) + " " + _
                              Strings.Left(str:=sProjectName, Length:=3)

        Return sName

    End Function

End Class
