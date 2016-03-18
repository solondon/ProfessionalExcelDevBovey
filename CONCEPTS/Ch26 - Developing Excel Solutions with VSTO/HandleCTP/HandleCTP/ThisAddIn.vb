Imports Tools = Microsoft.Office.Tools

Public Class ThisAddIn

    'Declare an instance.
    Private CTP As CTPControl

    'To use the VisibleChanged event of the CTP
    Private WithEvents CTPVisible As Tools.CustomTaskPane

    Private m_ctpTaskPane As Tools.CustomTaskPane

    'Make the CTP available in the solution.
    Public ReadOnly Property CtpTaskPane() As Tools.CustomTaskPane
        Get
            Return CTPVisible
        End Get

    End Property

    Private Sub ThisAddIn_Startup(ByVal sender As Object, _
                                  ByVal e As System.EventArgs) _
                                  Handles Me.Startup

        'Instantiate a new instance of the class.
        CTP = New CTPControl()
        'Add the CTP to the collection of CTPs.
        CTPVisible = Me.CustomTaskPanes.Add(CTP, "Report")

        CTP = Nothing


    End Sub

    Private Sub CTPVisible_VisibleChanged(ByVal sender As Object, _
                                          ByVal e As System.EventArgs) _
                                          Handles CTPVisible.VisibleChanged

        'Synchronize the ToggleButton state (to not pressed)
        'when users close the CTP via the Close button (x).
        Globals.Ribbons.Report.ToggleButton.Checked = _
                                       CTPVisible.Visible


    End Sub
End Class

