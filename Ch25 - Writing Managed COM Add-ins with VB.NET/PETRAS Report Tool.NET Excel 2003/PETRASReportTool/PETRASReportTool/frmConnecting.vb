'
' Description:  This class module contains code that manipulate
'               the Windows Form's UI. It is used to inform
'               users that the tool tries to access the database.
'
' Authors:      Dennis Wallentin, www.excelkb.com
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Class frmConnecting

    Private Sub frmAccess_Load(ByVal sender As System.Object, _
                               ByVal e As System.EventArgs) Handles MyBase.Load

        Const sCAPTION As String = "Connecting"
        Const sMESSAGE As String = "Please wait while trying to connect to the database."

        'Manipulating some of the Windows Form's main properties.
        With Me
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            .Icon = My.Resources.PetrasIcon
            .Text = sCAPTION
            'Put and keep the Windows Form on top. 
            .TopMost = True
        End With

        With Me.lblConnecting
            .Width = 35
            .Text = sMESSAGE
            .Update()
        End With
    End Sub

End Class