Imports System.Windows.Forms

Public Class CTPControl

    Private Sub Cctp_Load(ByVal sender As System.Object, _
                          ByVal e As System.EventArgs) _
                          Handles MyBase.Load

        'Manipulate the two DateTime Picker Controls.
        With Me.dtpStartDate
            .Format = DateTimePickerFormat.Short
            .Value = DateTime.Now.AddDays(-10)
            .MinDate = DateTime.Now.AddMonths(-10)
            .MaxDate = DateTime.Now.AddMonths(-1)
        End With

        With Me.dtpEndDate
            .Format = DateTimePickerFormat.Short
            .Value = Today
            .MinDate = DateTime.Now.AddDays(-10)
            .MaxDate = Today
        End With

        'Only hard coded for the example.
        With Me
            .cboClient.Text = "SC Forest"
            .cboProject.Text = "ADD102 - X200"
        End With

    End Sub

    Private Sub cmdCreate_Report_Click(ByVal sender As System.Object, _
                                       ByVal e As System.EventArgs) _
                                       Handles cmdCreate_Report.Click

        'Write data from the Report CTP to the active worksheet.
        If (Globals.ThisAddIn.Application.ActiveSheet IsNot Nothing) Then
            Dim wsTarget As Excel.Worksheet = _
            CType(Globals.ThisAddIn.Application.ActiveSheet, Excel.Worksheet)

            With Me
                wsTarget.Range("D7").Value = .cboClient.Text
                wsTarget.Range("D8").Value = .cboProject.Text
                wsTarget.Range("D9").Value = CDate(.dtpStartDate.Text)
                wsTarget.Range("D10").Value = CDate(.dtpEndDate.Text)
            End With

            wsTarget = Nothing

        End If


    End Sub

    Private Sub cmdClose_Pane_Click(ByVal sender As System.Object, _
                                    ByVal e As System.EventArgs)

        'Close the Report pane.
        Globals.ThisAddIn.CtpTaskPane.Visible = False
    End Sub

End Class
