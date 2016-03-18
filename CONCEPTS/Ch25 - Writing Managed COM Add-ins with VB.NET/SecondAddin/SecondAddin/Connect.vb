imports Extensibility
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Reflection
Imports System.IO

#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the $SAFEOBJNAME$Setup project, 
' right click the project in the Solution Explorer, then choose install.
#End Region

<GuidAttribute("D4AB6A6F-37B0-4FDF-8E12-D273DFE70EF8"), ProgIdAttribute("SecondAddin.Connect")> _
Public Class Connect

#Region "Implements"

    Implements IDTExtensibility2

    Implements IRibbonExtensibility

#End Region

#Region "Module-level variables."

    'The variable for Excel Application's object.
    Private mxlApp As Excel.Application = Nothing

    'Constants string variables.
    Private Const msTITLE As String = "Ribbon Handling"

    Private Const msBUTTON1 As String = "Time Report"
    Private Const msBUTTON2 As String = "Chart Report"
    Private Const msBUTTON3 As String = "Data Report"

    'Customized error message.
    Private Const msERROR_MESSAGE As String = _
                  "Cannot execute the wanted action."

    'Customized click message.
    Dim msTEXT As String = "You clicked on "

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

            Release_All_COMObjects(mxlApp)

        Catch GeneralEx As Exception

            'An error message during the debug process should be added."

        End Try

    End Sub

#End Region

#Region "Ribbon UI"

    Public Function GetCustomUI(ByVal ribbonID As String) As String _
                           Implements IRibbonExtensibility.GetCustomUI

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

                'Returns the XML to the Ribbon user interface.
                Return sResource

            End If

        Next

        Return Nothing

    End Function

    Public Sub PED_Reports_Click(ByVal control As IRibbonControl)

        Try
            'Run the appropriate message.
            Select Case control.Id
                Case "rxbtnTime"
                    MessageBox.Show(text:=msTEXT & msBUTTON1, _
                                    caption:=msTITLE)
                Case "rxbtnChart"
                    MessageBox.Show(text:=msTEXT & msBUTTON2, _
                                    caption:=msTITLE)
                Case "rxbtnData"
                    MessageBox.Show(text:=msTEXT & msBUTTON3, _
                                        caption:=msTITLE)

            End Select

        Catch Generalex As Exception

            'Show the customized message.
            MessageBox.Show(text:=msERROR_MESSAGE, _
                            caption:=msTITLE, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)
        End Try
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

#Region "Necessary Event Procedures"

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
