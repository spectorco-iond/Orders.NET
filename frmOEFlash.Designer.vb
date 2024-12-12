<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOEFlash
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
    Public WithEvents txtCusNo As System.Windows.Forms.TextBox
	Public WithEvents txtCusName As System.Windows.Forms.TextBox
	Public WithEvents lblCusNo As System.Windows.Forms.Label
	Public WithEvents lblCusName As System.Windows.Forms.Label
	Public WithEvents fraCustomerInfo As System.Windows.Forms.GroupBox
	Public WithEvents chkUnlock As System.Windows.Forms.CheckBox
	Public WithEvents cmdDeleteComment As System.Windows.Forms.Button
	Public WithEvents cmdInsertComment As System.Windows.Forms.Button
	Public WithEvents fraEdit As System.Windows.Forms.GroupBox
	Public WithEvents cmdFullScreen As System.Windows.Forms.Button
	Public WithEvents cmdClose As System.Windows.Forms.Button
    Public WithEvents fraComments As System.Windows.Forms.GroupBox
    Public WithEvents lblVersion As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdClose = New System.Windows.Forms.Button
        Me.fraCustomerInfo = New System.Windows.Forms.GroupBox
        Me.txtClassID = New System.Windows.Forms.TextBox
        Me.lblCusName = New System.Windows.Forms.Label
        Me.txtEQP = New System.Windows.Forms.TextBox
        Me.txtCusNo = New System.Windows.Forms.TextBox
        Me.txtCusName = New System.Windows.Forms.TextBox
        Me.lblCusNo = New System.Windows.Forms.Label
        Me.fraEdit = New System.Windows.Forms.GroupBox
        Me.cmdDown = New System.Windows.Forms.Button
        Me.cmdUp = New System.Windows.Forms.Button
        Me.chkInsertAtEnd = New System.Windows.Forms.CheckBox
        Me.chkUnlock = New System.Windows.Forms.CheckBox
        Me.cmdDeleteComment = New System.Windows.Forms.Button
        Me.cmdInsertComment = New System.Windows.Forms.Button
        Me.cmdFullScreen = New System.Windows.Forms.Button
        Me.fraComments = New System.Windows.Forms.GroupBox
        Me.dgvComments = New System.Windows.Forms.DataGridView
        Me.lblVersion = New System.Windows.Forms.Label
        Me.fraCustomerInfo.SuspendLayout()
        Me.fraEdit.SuspendLayout()
        Me.fraComments.SuspendLayout()
        CType(Me.dgvComments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(661, 56)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(89, 33)
        Me.cmdClose.TabIndex = 5
        Me.cmdClose.Text = "&Close"
        Me.ToolTip1.SetToolTip(Me.cmdClose, "Return to Header Screen")
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'fraCustomerInfo
        '
        Me.fraCustomerInfo.BackColor = System.Drawing.SystemColors.Control
        Me.fraCustomerInfo.Controls.Add(Me.txtClassID)
        Me.fraCustomerInfo.Controls.Add(Me.lblCusName)
        Me.fraCustomerInfo.Controls.Add(Me.txtEQP)
        Me.fraCustomerInfo.Controls.Add(Me.txtCusNo)
        Me.fraCustomerInfo.Controls.Add(Me.txtCusName)
        Me.fraCustomerInfo.Controls.Add(Me.lblCusNo)
        Me.fraCustomerInfo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraCustomerInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCustomerInfo.Location = New System.Drawing.Point(8, 8)
        Me.fraCustomerInfo.Name = "fraCustomerInfo"
        Me.fraCustomerInfo.Padding = New System.Windows.Forms.Padding(0)
        Me.fraCustomerInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCustomerInfo.Size = New System.Drawing.Size(289, 81)
        Me.fraCustomerInfo.TabIndex = 0
        Me.fraCustomerInfo.TabStop = False
        Me.fraCustomerInfo.Text = "Customer Info"
        '
        'txtClassID
        '
        Me.txtClassID.AcceptsReturn = True
        Me.txtClassID.BackColor = System.Drawing.SystemColors.Control
        Me.txtClassID.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtClassID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtClassID.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtClassID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtClassID.Location = New System.Drawing.Point(11, 60)
        Me.txtClassID.MaxLength = 0
        Me.txtClassID.Multiline = True
        Me.txtClassID.Name = "txtClassID"
        Me.txtClassID.ReadOnly = True
        Me.txtClassID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtClassID.Size = New System.Drawing.Size(95, 16)
        Me.txtClassID.TabIndex = 6
        '
        'lblCusName
        '
        Me.lblCusName.BackColor = System.Drawing.SystemColors.Control
        Me.lblCusName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCusName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCusName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCusName.Location = New System.Drawing.Point(8, 38)
        Me.lblCusName.Name = "lblCusName"
        Me.lblCusName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCusName.Size = New System.Drawing.Size(97, 17)
        Me.lblCusName.TabIndex = 3
        Me.lblCusName.Text = "Customer Name"
        '
        'txtEQP
        '
        Me.txtEQP.AcceptsReturn = True
        Me.txtEQP.BackColor = System.Drawing.SystemColors.Control
        Me.txtEQP.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtEQP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEQP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEQP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEQP.Location = New System.Drawing.Point(112, 60)
        Me.txtEQP.MaxLength = 0
        Me.txtEQP.Multiline = True
        Me.txtEQP.Name = "txtEQP"
        Me.txtEQP.ReadOnly = True
        Me.txtEQP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEQP.Size = New System.Drawing.Size(161, 16)
        Me.txtEQP.TabIndex = 5
        '
        'txtCusNo
        '
        Me.txtCusNo.AcceptsReturn = True
        Me.txtCusNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCusNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCusNo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCusNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCusNo.Location = New System.Drawing.Point(112, 14)
        Me.txtCusNo.MaxLength = 0
        Me.txtCusNo.Name = "txtCusNo"
        Me.txtCusNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCusNo.Size = New System.Drawing.Size(65, 20)
        Me.txtCusNo.TabIndex = 1
        '
        'txtCusName
        '
        Me.txtCusName.AcceptsReturn = True
        Me.txtCusName.BackColor = System.Drawing.SystemColors.Control
        Me.txtCusName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtCusName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCusName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCusName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCusName.Location = New System.Drawing.Point(112, 38)
        Me.txtCusName.MaxLength = 0
        Me.txtCusName.Multiline = True
        Me.txtCusName.Name = "txtCusName"
        Me.txtCusName.ReadOnly = True
        Me.txtCusName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCusName.Size = New System.Drawing.Size(161, 17)
        Me.txtCusName.TabIndex = 4
        '
        'lblCusNo
        '
        Me.lblCusNo.BackColor = System.Drawing.SystemColors.Control
        Me.lblCusNo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCusNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCusNo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCusNo.Location = New System.Drawing.Point(8, 18)
        Me.lblCusNo.Name = "lblCusNo"
        Me.lblCusNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCusNo.Size = New System.Drawing.Size(81, 17)
        Me.lblCusNo.TabIndex = 0
        Me.lblCusNo.Text = "Customer No"
        '
        'fraEdit
        '
        Me.fraEdit.BackColor = System.Drawing.SystemColors.Control
        Me.fraEdit.Controls.Add(Me.cmdDown)
        Me.fraEdit.Controls.Add(Me.cmdUp)
        Me.fraEdit.Controls.Add(Me.chkInsertAtEnd)
        Me.fraEdit.Controls.Add(Me.chkUnlock)
        Me.fraEdit.Controls.Add(Me.cmdDeleteComment)
        Me.fraEdit.Controls.Add(Me.cmdInsertComment)
        Me.fraEdit.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraEdit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraEdit.Location = New System.Drawing.Point(304, 8)
        Me.fraEdit.Name = "fraEdit"
        Me.fraEdit.Padding = New System.Windows.Forms.Padding(0)
        Me.fraEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraEdit.Size = New System.Drawing.Size(351, 81)
        Me.fraEdit.TabIndex = 3
        Me.fraEdit.TabStop = False
        Me.fraEdit.Text = "Comment Editing"
        '
        'cmdDown
        '
        Me.cmdDown.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDown.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDown.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDown.Location = New System.Drawing.Point(280, 48)
        Me.cmdDown.Name = "cmdDown"
        Me.cmdDown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDown.Size = New System.Drawing.Size(60, 25)
        Me.cmdDown.TabIndex = 5
        Me.cmdDown.Text = "Down"
        Me.cmdDown.UseVisualStyleBackColor = False
        '
        'cmdUp
        '
        Me.cmdUp.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUp.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUp.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUp.Location = New System.Drawing.Point(280, 16)
        Me.cmdUp.Name = "cmdUp"
        Me.cmdUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUp.Size = New System.Drawing.Size(60, 25)
        Me.cmdUp.TabIndex = 4
        Me.cmdUp.Text = "Up"
        Me.cmdUp.UseVisualStyleBackColor = False
        '
        'chkInsertAtEnd
        '
        Me.chkInsertAtEnd.BackColor = System.Drawing.SystemColors.Control
        Me.chkInsertAtEnd.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkInsertAtEnd.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkInsertAtEnd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkInsertAtEnd.Location = New System.Drawing.Point(16, 38)
        Me.chkInsertAtEnd.Name = "chkInsertAtEnd"
        Me.chkInsertAtEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkInsertAtEnd.Size = New System.Drawing.Size(152, 25)
        Me.chkInsertAtEnd.TabIndex = 3
        Me.chkInsertAtEnd.Text = "Insert at the end"
        Me.chkInsertAtEnd.UseVisualStyleBackColor = False
        '
        'chkUnlock
        '
        Me.chkUnlock.BackColor = System.Drawing.SystemColors.Control
        Me.chkUnlock.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkUnlock.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUnlock.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkUnlock.Location = New System.Drawing.Point(16, 16)
        Me.chkUnlock.Name = "chkUnlock"
        Me.chkUnlock.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkUnlock.Size = New System.Drawing.Size(152, 25)
        Me.chkUnlock.TabIndex = 0
        Me.chkUnlock.Text = "Click to Unlock for Editing"
        Me.chkUnlock.UseVisualStyleBackColor = False
        '
        'cmdDeleteComment
        '
        Me.cmdDeleteComment.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDeleteComment.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDeleteComment.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeleteComment.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDeleteComment.Location = New System.Drawing.Point(174, 48)
        Me.cmdDeleteComment.Name = "cmdDeleteComment"
        Me.cmdDeleteComment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDeleteComment.Size = New System.Drawing.Size(100, 25)
        Me.cmdDeleteComment.TabIndex = 2
        Me.cmdDeleteComment.Text = "Delete Comment"
        Me.cmdDeleteComment.UseVisualStyleBackColor = False
        '
        'cmdInsertComment
        '
        Me.cmdInsertComment.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInsertComment.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdInsertComment.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdInsertComment.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInsertComment.Location = New System.Drawing.Point(174, 16)
        Me.cmdInsertComment.Name = "cmdInsertComment"
        Me.cmdInsertComment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdInsertComment.Size = New System.Drawing.Size(100, 25)
        Me.cmdInsertComment.TabIndex = 1
        Me.cmdInsertComment.Text = "Insert Comment"
        Me.cmdInsertComment.UseVisualStyleBackColor = False
        '
        'cmdFullScreen
        '
        Me.cmdFullScreen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdFullScreen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdFullScreen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFullScreen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdFullScreen.Location = New System.Drawing.Point(640, 16)
        Me.cmdFullScreen.Name = "cmdFullScreen"
        Me.cmdFullScreen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdFullScreen.Size = New System.Drawing.Size(0, 0)
        Me.cmdFullScreen.TabIndex = 4
        Me.cmdFullScreen.Text = "Small Screen"
        Me.cmdFullScreen.UseVisualStyleBackColor = False
        '
        'fraComments
        '
        Me.fraComments.BackColor = System.Drawing.SystemColors.Control
        Me.fraComments.Controls.Add(Me.dgvComments)
        Me.fraComments.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraComments.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraComments.Location = New System.Drawing.Point(8, 96)
        Me.fraComments.Name = "fraComments"
        Me.fraComments.Padding = New System.Windows.Forms.Padding(0)
        Me.fraComments.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraComments.Size = New System.Drawing.Size(745, 345)
        Me.fraComments.TabIndex = 1
        Me.fraComments.TabStop = False
        Me.fraComments.Text = "Customer Comments"
        '
        'dgvComments
        '
        Me.dgvComments.AllowUserToAddRows = False
        Me.dgvComments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvComments.Location = New System.Drawing.Point(11, 22)
        Me.dgvComments.Name = "dgvComments"
        Me.dgvComments.Size = New System.Drawing.Size(722, 309)
        Me.dgvComments.TabIndex = 2
        '
        'lblVersion
        '
        Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblVersion.Location = New System.Drawing.Point(0, 448)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(89, 17)
        Me.lblVersion.TabIndex = 2
        Me.lblVersion.Text = "lblVersion"
        '
        'frmOEFlash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(762, 456)
        Me.Controls.Add(Me.fraCustomerInfo)
        Me.Controls.Add(Me.fraEdit)
        Me.Controls.Add(Me.cmdFullScreen)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.fraComments)
        Me.Controls.Add(Me.lblVersion)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(770, 600)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(770, 280)
        Me.Name = "frmOEFlash"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "OE Flash"
        Me.fraCustomerInfo.ResumeLayout(False)
        Me.fraCustomerInfo.PerformLayout()
        Me.fraEdit.ResumeLayout(False)
        Me.fraComments.ResumeLayout(False)
        CType(Me.dgvComments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support"
    Friend WithEvents dgvComments As System.Windows.Forms.DataGridView
    Public WithEvents chkInsertAtEnd As System.Windows.Forms.CheckBox
    Public WithEvents cmdDown As System.Windows.Forms.Button
    Public WithEvents cmdUp As System.Windows.Forms.Button
    Public WithEvents txtEQP As System.Windows.Forms.TextBox
    Public WithEvents txtClassID As System.Windows.Forms.TextBox
#End Region
End Class