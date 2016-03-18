<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConnecting
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblConnecting = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblConnecting
        '
        Me.lblConnecting.AutoSize = True
        Me.lblConnecting.Location = New System.Drawing.Point(12, 9)
        Me.lblConnecting.Name = "lblConnecting"
        Me.lblConnecting.Size = New System.Drawing.Size(71, 13)
        Me.lblConnecting.TabIndex = 0
        Me.lblConnecting.Text = "lblConnecting"
        '
        'frmConnecting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(272, 33)
        Me.Controls.Add(Me.lblConnecting)
        Me.Name = "frmConnecting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmAccess"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblConnecting As System.Windows.Forms.Label
End Class
