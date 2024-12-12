<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmShipLinks
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
        Me.dgvShipLinks = New System.Windows.Forms.DataGridView()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdAddShip_Link = New System.Windows.Forms.Button()
        Me.cmdDeleteShip_Link = New System.Windows.Forms.Button()
        Me.lblShip_Link_No = New System.Windows.Forms.Label()
        Me.txtShip_Link_No = New Orders.xTextBox()
        CType(Me.dgvShipLinks, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvShipLinks
        '
        Me.dgvShipLinks.AllowUserToAddRows = False
        Me.dgvShipLinks.AllowUserToDeleteRows = False
        Me.dgvShipLinks.AllowUserToResizeRows = False
        Me.dgvShipLinks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvShipLinks.Location = New System.Drawing.Point(12, 58)
        Me.dgvShipLinks.Name = "dgvShipLinks"
        Me.dgvShipLinks.Size = New System.Drawing.Size(223, 281)
        Me.dgvShipLinks.TabIndex = 1
        Me.dgvShipLinks.TabStop = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(151, 345)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(85, 24)
        Me.cmdClose.TabIndex = 10
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'cmdAddShip_Link
        '
        Me.cmdAddShip_Link.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddShip_Link.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddShip_Link.Enabled = False
        Me.cmdAddShip_Link.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddShip_Link.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddShip_Link.Location = New System.Drawing.Point(150, 8)
        Me.cmdAddShip_Link.Name = "cmdAddShip_Link"
        Me.cmdAddShip_Link.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddShip_Link.Size = New System.Drawing.Size(86, 24)
        Me.cmdAddShip_Link.TabIndex = 11
        Me.cmdAddShip_Link.Text = "Add Link"
        Me.cmdAddShip_Link.UseVisualStyleBackColor = False
        Me.cmdAddShip_Link.Visible = False
        '
        'cmdDeleteShip_Link
        '
        Me.cmdDeleteShip_Link.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDeleteShip_Link.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDeleteShip_Link.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeleteShip_Link.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDeleteShip_Link.Location = New System.Drawing.Point(12, 345)
        Me.cmdDeleteShip_Link.Name = "cmdDeleteShip_Link"
        Me.cmdDeleteShip_Link.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDeleteShip_Link.Size = New System.Drawing.Size(85, 24)
        Me.cmdDeleteShip_Link.TabIndex = 12
        Me.cmdDeleteShip_Link.Text = "Remove"
        Me.cmdDeleteShip_Link.UseVisualStyleBackColor = False
        Me.cmdDeleteShip_Link.Visible = False
        '
        'lblShip_Link_No
        '
        Me.lblShip_Link_No.BackColor = System.Drawing.SystemColors.Control
        Me.lblShip_Link_No.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblShip_Link_No.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblShip_Link_No.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShip_Link_No.Location = New System.Drawing.Point(12, 12)
        Me.lblShip_Link_No.Name = "lblShip_Link_No"
        Me.lblShip_Link_No.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblShip_Link_No.Size = New System.Drawing.Size(132, 17)
        Me.lblShip_Link_No.TabIndex = 13
        Me.lblShip_Link_No.Text = "Link P/O Number"
        '
        'txtShip_Link_No
        '
        Me.txtShip_Link_No.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_Link_No.DataLength = CType(30, Long)
        Me.txtShip_Link_No.DataType = Orders.CDataType.dtString
        Me.txtShip_Link_No.DateValue = New Date(CType(0, Long))
        Me.txtShip_Link_No.DecimalDigits = CType(0, Long)
        Me.txtShip_Link_No.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShip_Link_No.Location = New System.Drawing.Point(12, 32)
        Me.txtShip_Link_No.Mask = Nothing
        Me.txtShip_Link_No.Name = "txtShip_Link_No"
        Me.txtShip_Link_No.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_Link_No.OldValue = ""
        Me.txtShip_Link_No.ReadOnly = True
        Me.txtShip_Link_No.Size = New System.Drawing.Size(224, 22)
        Me.txtShip_Link_No.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_Link_No.StringValue = Nothing
        Me.txtShip_Link_No.TabIndex = 14
        Me.txtShip_Link_No.TextDataID = Nothing
        '
        'frmShipLinks
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(248, 381)
        Me.ControlBox = False
        Me.Controls.Add(Me.txtShip_Link_No)
        Me.Controls.Add(Me.lblShip_Link_No)
        Me.Controls.Add(Me.cmdDeleteShip_Link)
        Me.Controls.Add(Me.cmdAddShip_Link)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.dgvShipLinks)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmShipLinks"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Shipping Links"
        CType(Me.dgvShipLinks, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvShipLinks As System.Windows.Forms.DataGridView
    Public WithEvents cmdClose As System.Windows.Forms.Button
    Public WithEvents cmdAddShip_Link As System.Windows.Forms.Button
    Public WithEvents cmdDeleteShip_Link As System.Windows.Forms.Button
    Friend WithEvents txtShip_Link_No As Orders.xTextBox
    Public WithEvents lblShip_Link_No As System.Windows.Forms.Label
End Class
