<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CTPControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboClient = New System.Windows.Forms.ComboBox
        Me.cboProject = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.chkActivities = New System.Windows.Forms.CheckBox
        Me.chkConsultants = New System.Windows.Forms.CheckBox
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.cmdCreate_Report = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Client:"
        '
        'cboClient
        '
        Me.cboClient.FormattingEnabled = True
        Me.cboClient.Location = New System.Drawing.Point(12, 32)
        Me.cboClient.Name = "cboClient"
        Me.cboClient.Size = New System.Drawing.Size(124, 21)
        Me.cboClient.TabIndex = 0
        '
        'cboProject
        '
        Me.cboProject.FormattingEnabled = True
        Me.cboProject.Location = New System.Drawing.Point(12, 78)
        Me.cboProject.Name = "cboProject"
        Me.cboProject.Size = New System.Drawing.Size(124, 21)
        Me.cboProject.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Project:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 159)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "End date:"
        '
        'chkActivities
        '
        Me.chkActivities.AutoSize = True
        Me.chkActivities.Location = New System.Drawing.Point(12, 210)
        Me.chkActivities.Name = "chkActivities"
        Me.chkActivities.Size = New System.Drawing.Size(68, 17)
        Me.chkActivities.TabIndex = 4
        Me.chkActivities.Text = "Activities"
        Me.chkActivities.UseVisualStyleBackColor = True
        '
        'chkConsultants
        '
        Me.chkConsultants.AutoSize = True
        Me.chkConsultants.Location = New System.Drawing.Point(12, 233)
        Me.chkConsultants.Name = "chkConsultants"
        Me.chkConsultants.Size = New System.Drawing.Size(81, 17)
        Me.chkConsultants.TabIndex = 5
        Me.chkConsultants.Text = "Consultants"
        Me.chkConsultants.UseVisualStyleBackColor = True
        '
        'dtpStartDate
        '
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStartDate.Location = New System.Drawing.Point(12, 127)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(124, 20)
        Me.dtpStartDate.TabIndex = 2
        '
        'dtpEndDate
        '
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpEndDate.Location = New System.Drawing.Point(12, 175)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(124, 20)
        Me.dtpEndDate.TabIndex = 3
        '
        'cmdCreate_Report
        '
        Me.cmdCreate_Report.Location = New System.Drawing.Point(12, 272)
        Me.cmdCreate_Report.Name = "cmdCreate_Report"
        Me.cmdCreate_Report.Size = New System.Drawing.Size(107, 25)
        Me.cmdCreate_Report.TabIndex = 6
        Me.cmdCreate_Report.Text = "Create &Report"
        Me.cmdCreate_Report.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(9, 111)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Start date:"
        '
        'CTPControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cmdCreate_Report)
        Me.Controls.Add(Me.dtpEndDate)
        Me.Controls.Add(Me.dtpStartDate)
        Me.Controls.Add(Me.chkConsultants)
        Me.Controls.Add(Me.chkActivities)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboProject)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboClient)
        Me.Controls.Add(Me.Label1)
        Me.Name = "CTPControl"
        Me.Size = New System.Drawing.Size(146, 382)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboClient As System.Windows.Forms.ComboBox
    Friend WithEvents cboProject As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents chkActivities As System.Windows.Forms.CheckBox
    Friend WithEvents chkConsultants As System.Windows.Forms.CheckBox
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmdCreate_Report As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
