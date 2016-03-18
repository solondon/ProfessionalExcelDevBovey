'
' Description:  This module contains all entry points into the application.
'
' Authors:      Dennis Wallentin, www.excelkb.com
'
'
' To create a startup module with a Main procedure the easiest way is to do
' the following.
' 1. Create a Windows Form project.
' 2. Add a module and add a Main procedure in it.
' 2. Open the project's properties page.
' 3. Select the Application tab.
' 4. Uncheck the option Enable application framework.
' 5. Select the Statup object Sub Main in the Application tab.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Imported namespaces
'
'To use the native .NET Messagebox object and access the 
'Application's objects easier.
Imports System.Windows.Forms
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Module MStartUp


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This is the startup procedure for the report tool.
    '           It calls the connection function which tries to create
    '           a connection to the database. If so then it loads and
    '           shows the Windows Form. If not then it shows a
    '           customized message to the user and then closes.
    '
    ' Arguments:    None
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 04/22/08      Dennis Wallentin    Ch24    Initial version
    ' 04/23/08      Dennis Wallentin    Ch24    Added swdtClients
    '

    Sub Main()

        'Enable Windows XP's style.
        Application.EnableVisualStyles()

        'Declare and instantiate the Windows Form.
        Dim frm As New frmMain

        'Set the position of the main Windows Form.
        frm.StartPosition = FormStartPosition.CenterScreen

        'Show the main Windows Form.
        Application.Run(mainForm:=frm)

        'The following lines are executed when the Windows Form
        'is closed.

        'Releases all resources the variable has consumed from
        'the memory.
        frm.Dispose()

        'Release the reference the variable holds and prepare it
        'to be collected by the Garbage Collector (GC) when it
        'comes around.
        frm = Nothing

    End Sub

End Module
