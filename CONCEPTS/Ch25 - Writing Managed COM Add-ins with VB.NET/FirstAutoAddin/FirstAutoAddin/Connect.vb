Imports Extensibility
Imports System.Runtime.InteropServices
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

<GuidAttribute("29EA6584-871E-438A-ADA8-8EBA0A7BCC5A"), _
 ProgIdAttribute("FirstAutoAddin.Connect"), _
 ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Connect

#Region "Implements"

    Implements IDTExtensibility2

#End Region

#Region "Modul-wide variables."

    'The variable for Excel Application's object.
    Dim mxlApp As Excel.Application

    'Constants string variable.
    Private Const msTITLE As String = "Automation Add-in"

#End Region

#Region "User Defined Functions."

    Public Function IFERROR(ByVal ToEvaluate As Object, _
                            ByVal UseDefault As Object) As Object

        Dim objOutput As Object = Nothing

        'This line will never be executed as .NET
        'exclude CVErr values.
        If IsError(ToEvaluate) Then
            objOutput = UseDefault
        Else
            objOutput = ToEvaluate
        End If

        Return objOutput

    End Function

    Public Function VBTIMER() As Double

        Dim dTime As Double = Nothing

        mxlApp.Volatile(True)

        dTime = Microsoft.VisualBasic.Timer

        Return dTime

    End Function

    Public Function COUNTBETWEEN(ByRef Source As Excel.Range, _
                                 ByVal Min As Double, _
                                 ByRef Max As Double) As Double

        Dim dCountBetween As Double = Nothing

        Try

            dCountBetween = mxlApp.WorksheetFunction.CountIf( _
                                              Source, ">" & Min)

            dCountBetween = dCountBetween - mxlApp.WorksheetFunction. _
                                            CountIf(Source, ">=" & Max)

        Catch GeneralEx As Exception

            dCountBetween = 0

        End Try

        Return dCountBetween

    End Function

#End Region

#Region "Connect and Disconnect"

    Public Sub OnConnection(ByVal application As Object, _
                            ByVal connectMode As ext_ConnectMode, _
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
                               Implements IDTExtensibility2.OnDisconnection
        'Cleaning up.
        Marshal.ReleaseComObject(mxlApp)
        mxlApp = Nothing
    End Sub

#End Region

#Region "Necessary Procedures."
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



