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
Imports System.IO
Imports System.Reflection
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
    Implements Office.IRibbonExtensibility

#End Region

#Region "Module-wide variables"


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
        Const sMESSAGEWRONGVERSION As String = "Version 2007 and later of Excel must" + vbNewLine + _
                                                 "be installed in order to proceed."

        'Customized error message.
        Const sERROR_MESSAGE As String = _
              "An unexpected error has occured."

        'Create and instantiate a new instance of the class.
        Dim cMethods As New CCommonMethods()

        Try
            'Get a reference to Excel.
            swXLApp = (CType(application, Excel.Application))

            'Check to see that Excel 2007 or later is installed on the computer.
            Dim shInstalled As Short = cMethods.shCheck_Excel_Version_Installed

            If shInstalled = xlVersion.WrongVersion Then

                'Customized message that the wrong Excel version is installed.
                MessageBox.Show(text:=sMESSAGEWRONGVERSION, _
                                caption:=swsCaption, _
                                buttons:=MessageBoxButtons.OK, _
                                icon:=MessageBoxIcon.Stop)

            End If

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

        'Prepare the object for GC. 
        swXLApp = Nothing

    End Sub

#End Region

#Region "Ribbon UI Handling"

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function read the RibbonCustom.xml file
    '           and is required by the IRibbonExtensibility interface.
    '
    ' Arguments:    ribbonID
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/22/08      Dennis Wallentin    Ch25    Initial version

    Public Function GetCustomUI(ByVal ribbonID As String) As String _
                           Implements Office.IRibbonExtensibility.GetCustomUI

        'The resource we want to retrieve the XML markup from.
        Const sResourceName As String = "RibbonCustom.xml"

        'Variable for iterating the collection of resources.
        Dim sName As String = Nothing

        'Set a reference to this assembly during runtime.
        Dim asm As Assembly = Assembly.GetExecutingAssembly()

        'Get the collection of resource names in this assembly.
        Dim ResourceNames() As String = asm.GetManifestResourceNames()

        'Iterate through the collection until it finds the resource
        'named RibbonCustom.xml.
        For Each sName In ResourceNames

            If sName.EndsWith(sResourceName) Then
                'Create an instance of the StremReader object that 
                'reads the embedded file containing the XML markup.
                Dim srResourceReader As StreamReader = _
                New StreamReader(asm.GetManifestResourceStream(sName))

                'Reads the content of the resource file.
                Dim sResource As String = srResourceReader.ReadToEnd()

                'Close the StreamReader.
                srResourceReader.Close()

                'Returns the XML to the Ribbon UI.
                Return sResource

            End If

        Next

        Return Nothing

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments:     This subroutine handle all buttons click event 
    '               in our custom menu.
    ' Arguments:    None
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/22/08      Dennis Wallentin    Ch25    Initial version

    Public Sub Reports_Click(ByVal control As Office.IRibbonControl)

        'Variables for the two buttons on the custom menu.
        Const sABOUT As String = "&About"
        Const sREPORT As String = "&Report"


        'Create and instantiate a new instance of the class.
        Dim cCMethods As New CCommonMethods

        Select Case control.Id

            'User selected to show the Report form.
            Case "rxbtnReport" : cCMethods.Load_Form(sForm:=sREPORT)

                'User select to show the About form.
            Case "rxbtnAbout" : cCMethods.Load_Form(sForm:=sABOUT)

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
