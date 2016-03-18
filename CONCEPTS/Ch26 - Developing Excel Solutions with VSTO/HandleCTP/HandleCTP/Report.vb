Imports Tools = Microsoft.Office.Tools

Public Class Report

    Private Sub ToggleButton_Click(ByVal sender As System.Object, _
                                   ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
                                   Handles ToggleButton.Click

        'Show or hide the CTP.
        Globals.ThisAddIn.CtpTaskPane.Visible = _
        CType(sender, Tools.Ribbon.RibbonToggleButton).Checked

    End Sub
End Class
