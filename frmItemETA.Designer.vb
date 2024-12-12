<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmItemETA
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvLocDetails = New System.Windows.Forms.DataGridView
        Me.lblItem_No = New System.Windows.Forms.Label
        Me.cmdClose = New System.Windows.Forms.Button
        Me.txtItem_No = New Orders.xTextBox
        CType(Me.dgvLocDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvLocDetails
        '
        Me.dgvLocDetails.AllowUserToAddRows = False
        Me.dgvLocDetails.AllowUserToDeleteRows = False
        Me.dgvLocDetails.AllowUserToResizeRows = False
        Me.dgvLocDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvLocDetails.Location = New System.Drawing.Point(12, 35)
        Me.dgvLocDetails.Name = "dgvLocDetails"
        Me.dgvLocDetails.ReadOnly = True
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgvLocDetails.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvLocDetails.Size = New System.Drawing.Size(524, 300)
        Me.dgvLocDetails.TabIndex = 129
        Me.dgvLocDetails.TabStop = False
        '
        'lblItem_No
        '
        Me.lblItem_No.BackColor = System.Drawing.SystemColors.Control
        Me.lblItem_No.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblItem_No.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblItem_No.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblItem_No.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblItem_No.Location = New System.Drawing.Point(9, 10)
        Me.lblItem_No.Name = "lblItem_No"
        Me.lblItem_No.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblItem_No.Size = New System.Drawing.Size(185, 20)
        Me.lblItem_No.TabIndex = 127
        Me.lblItem_No.Text = "Item No"
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdClose.Location = New System.Drawing.Point(461, 341)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(75, 24)
        Me.cmdClose.TabIndex = 130
        Me.cmdClose.TabStop = False
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'txtItem_No
        '
        Me.txtItem_No.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtItem_No.DataLength = CType(25, Long)
        Me.txtItem_No.DataType = Orders.CDataType.dtString
        Me.txtItem_No.DateValue = New Date(CType(0, Long))
        Me.txtItem_No.DecimalDigits = CType(0, Long)
        Me.txtItem_No.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItem_No.Location = New System.Drawing.Point(200, 7)
        Me.txtItem_No.Mask = Nothing
        Me.txtItem_No.Name = "txtItem_No"
        Me.txtItem_No.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtItem_No.OldValue = Nothing
        Me.txtItem_No.ReadOnly = True
        Me.txtItem_No.Size = New System.Drawing.Size(336, 22)
        Me.txtItem_No.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtItem_No.StringValue = Nothing
        Me.txtItem_No.TabIndex = 128
        Me.txtItem_No.TextDataID = Nothing
        '
        'frmItemETA
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(548, 372)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.dgvLocDetails)
        Me.Controls.Add(Me.txtItem_No)
        Me.Controls.Add(Me.lblItem_No)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmItemETA"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmItemETA"
        CType(Me.dgvLocDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvLocDetails As System.Windows.Forms.DataGridView
    Friend WithEvents txtItem_No As Orders.xTextBox
    Friend WithEvents lblItem_No As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
End Class
