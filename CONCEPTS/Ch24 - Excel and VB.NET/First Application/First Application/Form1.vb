'To get access to the messagebox object.
Imports System.Windows.Forms

Public Class Form1

    Private Sub Form1_Load(ByVal sender As Object, _
                           ByVal e As System.EventArgs) _
                           Handles Me.Load

        'The array with names.
        Dim sArrNames() As String = {"Rob Bovey", "Stephen Bullen", _
                                     "John Green", "Dennis Wallentin"}
        With Me

            'The caption of the Form.
            .Text = "First Application"

            'The captions of the label and button controls.
            .Label1.Text = "Select the name:"
            .Button1.Text = "&Show value"
            .Button2.Text = "&Close"

            'Populate the combobox control with the list of names.
            With .ComboBox1
                .DataSource = sArrNames
                .SelectedIndex = -1
            End With

        End With

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) _
                              Handles Button1.Click

        'Make sure that a name has been selected.
        If Me.ComboBox1.SelectedIndex <> -1 Then

            'Show the selected value.
            MessageBox.Show( _
                    text:=Me.ComboBox1.SelectedValue.ToString(), _
                    caption:="First Application")

        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) _
                              Handles Button2.Click
        Me.Close()
    End Sub
End Class
