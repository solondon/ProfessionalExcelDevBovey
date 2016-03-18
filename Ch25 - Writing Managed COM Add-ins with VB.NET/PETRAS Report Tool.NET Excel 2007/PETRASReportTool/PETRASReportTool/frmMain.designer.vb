<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container
        Me.gbxReportSettings = New System.Windows.Forms.GroupBox
        Me.gbxClientProject = New System.Windows.Forms.GroupBox
        Me.cboProjects = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboClients = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmdClearSettings = New System.Windows.Forms.Button
        Me.cmdCreateReport = New System.Windows.Forms.Button
        Me.gbxConsultantsActivities = New System.Windows.Forms.GroupBox
        Me.chkActivities = New System.Windows.Forms.CheckBox
        Me.chkConsultants = New System.Windows.Forms.CheckBox
        Me.gbxPeriod = New System.Windows.Forms.GroupBox
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdExportXML = New System.Windows.Forms.Button
        Me.cmdExportExcel = New System.Windows.Forms.Button
        Me.llblBook = New System.Windows.Forms.LinkLabel
        Me.Label5 = New System.Windows.Forms.Label
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.dgvReport = New System.Windows.Forms.DataGridView
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.lblConnection = New System.Windows.Forms.Label
        Me.gbxReportSettings.SuspendLayout()
        Me.gbxClientProject.SuspendLayout()
        Me.gbxConsultantsActivities.SuspendLayout()
        Me.gbxPeriod.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvReport, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbxReportSettings
        '
        Me.gbxReportSettings.Controls.Add(Me.gbxClientProject)
        Me.gbxReportSettings.Controls.Add(Me.cmdClearSettings)
        Me.gbxReportSettings.Controls.Add(Me.cmdCreateReport)
        Me.gbxReportSettings.Controls.Add(Me.gbxConsultantsActivities)
        Me.gbxReportSettings.Controls.Add(Me.gbxPeriod)
        Me.gbxReportSettings.Location = New System.Drawing.Point(448, 27)
        Me.gbxReportSettings.Name = "gbxReportSettings"
        Me.gbxReportSettings.Size = New System.Drawing.Size(228, 300)
        Me.gbxReportSettings.TabIndex = 1
        Me.gbxReportSettings.TabStop = False
        Me.gbxReportSettings.Text = "Report settings:"
        '
        'gbxClientProject
        '
        Me.gbxClientProject.Controls.Add(Me.cboProjects)
        Me.gbxClientProject.Controls.Add(Me.Label2)
        Me.gbxClientProject.Controls.Add(Me.cboClients)
        Me.gbxClientProject.Controls.Add(Me.Label1)
        Me.gbxClientProject.Location = New System.Drawing.Point(5, 16)
        Me.gbxClientProject.Name = "gbxClientProject"
        Me.gbxClientProject.Size = New System.Drawing.Size(216, 114)
        Me.gbxClientProject.TabIndex = 0
        Me.gbxClientProject.TabStop = False
        '
        'cboProjects
        '
        Me.cboProjects.FormattingEnabled = True
        Me.cboProjects.Location = New System.Drawing.Point(9, 77)
        Me.cboProjects.Name = "cboProjects"
        Me.cboProjects.Size = New System.Drawing.Size(136, 21)
        Me.cboProjects.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Project:"
        '
        'cboClients
        '
        Me.cboClients.FormattingEnabled = True
        Me.cboClients.Location = New System.Drawing.Point(9, 34)
        Me.cboClients.Name = "cboClients"
        Me.cboClients.Size = New System.Drawing.Size(136, 21)
        Me.cboClients.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Client:"
        '
        'cmdClearSettings
        '
        Me.cmdClearSettings.Location = New System.Drawing.Point(49, 264)
        Me.cmdClearSettings.Name = "cmdClearSettings"
        Me.cmdClearSettings.Size = New System.Drawing.Size(83, 26)
        Me.cmdClearSettings.TabIndex = 3
        Me.cmdClearSettings.Text = "C&lear Settings"
        Me.cmdClearSettings.UseVisualStyleBackColor = True
        '
        'cmdCreateReport
        '
        Me.cmdCreateReport.Location = New System.Drawing.Point(138, 264)
        Me.cmdCreateReport.Name = "cmdCreateReport"
        Me.cmdCreateReport.Size = New System.Drawing.Size(83, 26)
        Me.cmdCreateReport.TabIndex = 4
        Me.cmdCreateReport.Text = "Create &Report"
        Me.cmdCreateReport.UseVisualStyleBackColor = True
        '
        'gbxConsultantsActivities
        '
        Me.gbxConsultantsActivities.Controls.Add(Me.chkActivities)
        Me.gbxConsultantsActivities.Controls.Add(Me.chkConsultants)
        Me.gbxConsultantsActivities.Location = New System.Drawing.Point(6, 202)
        Me.gbxConsultantsActivities.Name = "gbxConsultantsActivities"
        Me.gbxConsultantsActivities.Size = New System.Drawing.Size(216, 49)
        Me.gbxConsultantsActivities.TabIndex = 2
        Me.gbxConsultantsActivities.TabStop = False
        Me.gbxConsultantsActivities.Text = "Show fields:"
        '
        'chkActivities
        '
        Me.chkActivities.AutoSize = True
        Me.chkActivities.Location = New System.Drawing.Point(8, 19)
        Me.chkActivities.Name = "chkActivities"
        Me.chkActivities.Size = New System.Drawing.Size(68, 17)
        Me.chkActivities.TabIndex = 0
        Me.chkActivities.Text = "&Activities"
        Me.chkActivities.UseVisualStyleBackColor = True
        '
        'chkConsultants
        '
        Me.chkConsultants.AutoSize = True
        Me.chkConsultants.Location = New System.Drawing.Point(111, 19)
        Me.chkConsultants.Name = "chkConsultants"
        Me.chkConsultants.Size = New System.Drawing.Size(84, 17)
        Me.chkConsultants.TabIndex = 1
        Me.chkConsultants.Text = "C&onsultants:"
        Me.chkConsultants.UseVisualStyleBackColor = True
        '
        'gbxPeriod
        '
        Me.gbxPeriod.Controls.Add(Me.dtpEndDate)
        Me.gbxPeriod.Controls.Add(Me.Label4)
        Me.gbxPeriod.Controls.Add(Me.dtpStartDate)
        Me.gbxPeriod.Controls.Add(Me.Label3)
        Me.gbxPeriod.Location = New System.Drawing.Point(6, 136)
        Me.gbxPeriod.Name = "gbxPeriod"
        Me.gbxPeriod.Size = New System.Drawing.Size(216, 60)
        Me.gbxPeriod.TabIndex = 1
        Me.gbxPeriod.TabStop = False
        Me.gbxPeriod.Text = "Period:"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.Location = New System.Drawing.Point(111, 29)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(96, 20)
        Me.dtpEndDate.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(108, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "End date:"
        '
        'dtpStartDate
        '
        Me.dtpStartDate.Location = New System.Drawing.Point(8, 29)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(96, 20)
        Me.dtpStartDate.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Start date:"
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.Location = New System.Drawing.Point(587, 342)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(83, 26)
        Me.cmdClose.TabIndex = 5
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cmdExportXML
        '
        Me.cmdExportXML.Location = New System.Drawing.Point(497, 342)
        Me.cmdExportXML.Name = "cmdExportXML"
        Me.cmdExportXML.Size = New System.Drawing.Size(83, 26)
        Me.cmdExportXML.TabIndex = 4
        Me.cmdExportXML.Text = "Export &XML"
        Me.cmdExportXML.UseVisualStyleBackColor = True
        '
        'cmdExportExcel
        '
        Me.cmdExportExcel.Location = New System.Drawing.Point(408, 342)
        Me.cmdExportExcel.Name = "cmdExportExcel"
        Me.cmdExportExcel.Size = New System.Drawing.Size(83, 26)
        Me.cmdExportExcel.TabIndex = 3
        Me.cmdExportExcel.Text = "&Export Excel"
        Me.cmdExportExcel.UseVisualStyleBackColor = True
        '
        'llblBook
        '
        Me.llblBook.AutoSize = True
        Me.llblBook.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.llblBook.Location = New System.Drawing.Point(13, 355)
        Me.llblBook.Name = "llblBook"
        Me.llblBook.Size = New System.Drawing.Size(79, 12)
        Me.llblBook.TabIndex = 2
        Me.llblBook.TabStop = True
        Me.llblBook.Text = "© 2009 XL-Dennis"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(78, 13)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Preview report:"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'HelpProvider1
        '
        Me.HelpProvider1.HelpNamespace = ""
        '
        'dgvReport
        '
        Me.dgvReport.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvReport.Location = New System.Drawing.Point(16, 32)
        Me.dgvReport.Name = "dgvReport"
        Me.dgvReport.Size = New System.Drawing.Size(419, 294)
        Me.dgvReport.TabIndex = 0
        '
        'BackgroundWorker1
        '
        '
        'lblConnection
        '
        Me.lblConnection.AutoSize = True
        Me.lblConnection.Location = New System.Drawing.Point(148, 349)
        Me.lblConnection.Name = "lblConnection"
        Me.lblConnection.Size = New System.Drawing.Size(0, 13)
        Me.lblConnection.TabIndex = 6
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(688, 384)
        Me.Controls.Add(Me.lblConnection)
        Me.Controls.Add(Me.dgvReport)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.llblBook)
        Me.Controls.Add(Me.cmdExportExcel)
        Me.Controls.Add(Me.cmdExportXML)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.gbxReportSettings)
        Me.HelpProvider1.SetHelpKeyword(Me, "About.htm")
        Me.HelpProvider1.SetHelpNavigator(Me, System.Windows.Forms.HelpNavigator.Topic)
        Me.Name = "frmMain"
        Me.HelpProvider1.SetShowHelp(Me, True)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Form1"
        Me.gbxReportSettings.ResumeLayout(False)
        Me.gbxClientProject.ResumeLayout(False)
        Me.gbxClientProject.PerformLayout()
        Me.gbxConsultantsActivities.ResumeLayout(False)
        Me.gbxConsultantsActivities.PerformLayout()
        Me.gbxPeriod.ResumeLayout(False)
        Me.gbxPeriod.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvReport, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbxReportSettings As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboClients As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents gbxPeriod As System.Windows.Forms.GroupBox
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents gbxConsultantsActivities As System.Windows.Forms.GroupBox
    Friend WithEvents chkConsultants As System.Windows.Forms.CheckBox
    Friend WithEvents chkActivities As System.Windows.Forms.CheckBox
    Friend WithEvents cmdCreateReport As System.Windows.Forms.Button
    Friend WithEvents cmdClearSettings As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdExportXML As System.Windows.Forms.Button
    Friend WithEvents cmdExportExcel As System.Windows.Forms.Button
    Friend WithEvents llblBook As System.Windows.Forms.LinkLabel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboProjects As System.Windows.Forms.ComboBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents gbxClientProject As System.Windows.Forms.GroupBox
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents dgvReport As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents lblConnection As System.Windows.Forms.Label

End Class
