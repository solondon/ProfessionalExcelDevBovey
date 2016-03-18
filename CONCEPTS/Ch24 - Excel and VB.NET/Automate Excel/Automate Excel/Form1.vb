'Namespace alias for Excel.
Imports Excel = Microsoft.Office.Interop.Excel

'To release COM objects and to Catch COM errors.
Imports System.Runtime.InteropServices


Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, _
                          ByVal e As System.EventArgs) _
                          Handles Button1.Click

        'Excel COM objects to be used.
        Dim xlApp As Excel.Application = Nothing
        Dim xlWkbNew As Excel.Workbook = Nothing
        Dim xlWksMain As Excel.Worksheet = Nothing
        Dim xlRngData As Excel.Range = Nothing
        Dim sData() As String = {"Hello", "World", "!"}


        Try
            'Instantiate a new Excel session.
            xlApp = New Excel.Application

            'Add a new workbook.
            xlWkbNew = xlApp.Workbooks.Add

            'Reference the first worksheet in the workbook.
            xlWksMain = CType(xlWkbNew.Worksheets(Index:=1),  _
                              Excel.Worksheet)

            'Reference to the range which we will write some data to.
            xlRngData = CType(xlWksMain.Range("A1:C1"),  _
                              Excel.Range)

            'Write the data to the range.
            xlRngData.Value = sData

            'Save the workbook.
            xlWkbNew.SaveAs(Filename:="c:\Test\New.xls")

            'Make Excel visible for the user.
            With xlApp
                .UserControl = True
                .Visible = True
            End With


        Catch COMex As COMException

        Catch ex As Exception

            '....
        Finally

            'Calling the Garbish Collector twice. 
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

            'Releasing the Excel objects.
            ReleaseCOMObject(xlRngData)
            ReleaseCOMObject(xlWksMain)
            ReleaseCOMObject(xlWkbNew)
            ReleaseCOMObject(xlApp)

        End Try

    End Sub

    Private Sub ReleaseCOMObject(ByVal oxlObject As Object)
        Try
            Marshal.ReleaseComObject(oxlObject)
            oxlObject = Nothing
        Catch ex As Exception
            oxlObject = Nothing
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class
