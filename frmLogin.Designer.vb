<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLogin
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents txtServerName As System.Windows.Forms.TextBox
	Public WithEvents cmdOpen As System.Windows.Forms.Button
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents txtDatabaseName As System.Windows.Forms.TextBox
	Public WithEvents lblServerName As System.Windows.Forms.Label
	Public WithEvents lblDatabaseName As System.Windows.Forms.Label
	Public WithEvents lblVersion As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtServerName = New System.Windows.Forms.TextBox()
        Me.cmdOpen = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.txtDatabaseName = New System.Windows.Forms.TextBox()
        Me.lblServerName = New System.Windows.Forms.Label()
        Me.lblDatabaseName = New System.Windows.Forms.Label()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtDSNName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDSNName_Synergy = New System.Windows.Forms.TextBox()
        Me.txtServerName_Synergy = New System.Windows.Forms.TextBox()
        Me.txtDatabaseName_Synergy = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtServerName
        '
        Me.txtServerName.AcceptsReturn = True
        Me.txtServerName.BackColor = System.Drawing.SystemColors.Window
        Me.txtServerName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtServerName.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServerName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtServerName.Location = New System.Drawing.Point(128, 73)
        Me.txtServerName.MaxLength = 0
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.ReadOnly = True
        Me.txtServerName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtServerName.Size = New System.Drawing.Size(97, 22)
        Me.txtServerName.TabIndex = 3
        '
        'cmdOpen
        '
        Me.cmdOpen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOpen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOpen.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOpen.Location = New System.Drawing.Point(19, 145)
        Me.cmdOpen.Name = "cmdOpen"
        Me.cmdOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOpen.Size = New System.Drawing.Size(89, 25)
        Me.cmdOpen.TabIndex = 2
        Me.cmdOpen.Text = "&Open"
        Me.cmdOpen.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(128, 145)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(97, 25)
        Me.cmdExit.TabIndex = 1
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'txtDatabaseName
        '
        Me.txtDatabaseName.AcceptsReturn = True
        Me.txtDatabaseName.BackColor = System.Drawing.SystemColors.Window
        Me.txtDatabaseName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDatabaseName.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDatabaseName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDatabaseName.Location = New System.Drawing.Point(128, 102)
        Me.txtDatabaseName.MaxLength = 0
        Me.txtDatabaseName.Name = "txtDatabaseName"
        Me.txtDatabaseName.ReadOnly = True
        Me.txtDatabaseName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDatabaseName.Size = New System.Drawing.Size(97, 22)
        Me.txtDatabaseName.TabIndex = 0
        '
        'lblServerName
        '
        Me.lblServerName.BackColor = System.Drawing.SystemColors.Control
        Me.lblServerName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblServerName.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblServerName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblServerName.Location = New System.Drawing.Point(16, 73)
        Me.lblServerName.Name = "lblServerName"
        Me.lblServerName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblServerName.Size = New System.Drawing.Size(106, 25)
        Me.lblServerName.TabIndex = 6
        Me.lblServerName.Text = "Server Name"
        Me.lblServerName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDatabaseName
        '
        Me.lblDatabaseName.BackColor = System.Drawing.SystemColors.Control
        Me.lblDatabaseName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDatabaseName.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDatabaseName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDatabaseName.Location = New System.Drawing.Point(16, 102)
        Me.lblDatabaseName.Name = "lblDatabaseName"
        Me.lblDatabaseName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDatabaseName.Size = New System.Drawing.Size(106, 25)
        Me.lblDatabaseName.TabIndex = 5
        Me.lblDatabaseName.Text = "DatabaseName"
        Me.lblDatabaseName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblVersion
        '
        Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblVersion.Location = New System.Drawing.Point(1, 177)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(192, 25)
        Me.lblVersion.TabIndex = 4
        Me.lblVersion.Text = "Version 2018.03.18"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(181, 174)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(41, 20)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = False
        Me.Button1.Visible = False
        '
        'txtDSNName
        '
        Me.txtDSNName.AcceptsReturn = True
        Me.txtDSNName.BackColor = System.Drawing.SystemColors.Window
        Me.txtDSNName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDSNName.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDSNName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDSNName.Location = New System.Drawing.Point(128, 44)
        Me.txtDSNName.MaxLength = 0
        Me.txtDSNName.Name = "txtDSNName"
        Me.txtDSNName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDSNName.Size = New System.Drawing.Size(97, 22)
        Me.txtDSNName.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(106, 25)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "DSN Name"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(128, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(97, 25)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Macola"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(246, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(106, 25)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Synergy"
        '
        'txtDSNName_Synergy
        '
        Me.txtDSNName_Synergy.AcceptsReturn = True
        Me.txtDSNName_Synergy.BackColor = System.Drawing.SystemColors.Window
        Me.txtDSNName_Synergy.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDSNName_Synergy.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDSNName_Synergy.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDSNName_Synergy.Location = New System.Drawing.Point(240, 41)
        Me.txtDSNName_Synergy.MaxLength = 0
        Me.txtDSNName_Synergy.Name = "txtDSNName_Synergy"
        Me.txtDSNName_Synergy.ReadOnly = True
        Me.txtDSNName_Synergy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDSNName_Synergy.Size = New System.Drawing.Size(97, 22)
        Me.txtDSNName_Synergy.TabIndex = 13
        '
        'txtServerName_Synergy
        '
        Me.txtServerName_Synergy.AcceptsReturn = True
        Me.txtServerName_Synergy.BackColor = System.Drawing.SystemColors.Window
        Me.txtServerName_Synergy.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtServerName_Synergy.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServerName_Synergy.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtServerName_Synergy.Location = New System.Drawing.Point(240, 69)
        Me.txtServerName_Synergy.MaxLength = 0
        Me.txtServerName_Synergy.Name = "txtServerName_Synergy"
        Me.txtServerName_Synergy.ReadOnly = True
        Me.txtServerName_Synergy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtServerName_Synergy.Size = New System.Drawing.Size(97, 22)
        Me.txtServerName_Synergy.TabIndex = 12
        '
        'txtDatabaseName_Synergy
        '
        Me.txtDatabaseName_Synergy.AcceptsReturn = True
        Me.txtDatabaseName_Synergy.BackColor = System.Drawing.SystemColors.Window
        Me.txtDatabaseName_Synergy.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDatabaseName_Synergy.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDatabaseName_Synergy.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDatabaseName_Synergy.Location = New System.Drawing.Point(240, 100)
        Me.txtDatabaseName_Synergy.MaxLength = 0
        Me.txtDatabaseName_Synergy.Name = "txtDatabaseName_Synergy"
        Me.txtDatabaseName_Synergy.ReadOnly = True
        Me.txtDatabaseName_Synergy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDatabaseName_Synergy.Size = New System.Drawing.Size(97, 22)
        Me.txtDatabaseName_Synergy.TabIndex = 11
        '
        'frmLogin
        '
        Me.AcceptButton = Me.cmdOpen
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Pink
        Me.CancelButton = Me.cmdExit
        Me.ClientSize = New System.Drawing.Size(234, 195)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDSNName_Synergy)
        Me.Controls.Add(Me.txtServerName_Synergy)
        Me.Controls.Add(Me.txtDatabaseName_Synergy)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtDSNName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtServerName)
        Me.Controls.Add(Me.cmdOpen)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.txtDatabaseName)
        Me.Controls.Add(Me.lblServerName)
        Me.Controls.Add(Me.lblDatabaseName)
        Me.Controls.Add(Me.lblVersion)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLogin"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Orders - Login"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Public WithEvents txtDSNName As TextBox
    Public WithEvents Label1 As Label
    Public WithEvents Label2 As Label
    Public WithEvents Label3 As Label
    Public WithEvents txtDSNName_Synergy As TextBox
    Public WithEvents txtServerName_Synergy As TextBox
    Public WithEvents txtDatabaseName_Synergy As TextBox
#End Region
End Class