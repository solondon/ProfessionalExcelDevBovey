Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Populate the listbox with the retrieved data by using a DataReader.
        Me.ListBox1.DataSource = MData.Retrieve_Data_With_DataReader

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Populate the listbox with the retrieved data by using a DataSet.
        'When using a DataTable as the source we need to explicit use the
        'DisplayMember property of the ListBox control to tell which
        'column to be displayed.
        With Me.ListBox2
            .DataSource = MData.Retrieve_Data_With_DataSet
            Me.ListBox2.DisplayMember = "Company"
        End With

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Close the Windows Form.
        Me.Close()
    End Sub
End Class
