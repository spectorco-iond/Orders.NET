<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRepeatImprint
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.cmdCopyImprintData = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.dgvExtraFields = New Orders.XDataGridView
        CType(Me.dgvExtraFields, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdCopyImprintData
        '
        Me.cmdCopyImprintData.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopyImprintData.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCopyImprintData.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopyImprintData.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopyImprintData.Location = New System.Drawing.Point(486, 526)
        Me.cmdCopyImprintData.Name = "cmdCopyImprintData"
        Me.cmdCopyImprintData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCopyImprintData.Size = New System.Drawing.Size(121, 33)
        Me.cmdCopyImprintData.TabIndex = 4
        Me.cmdCopyImprintData.Text = "Copy Imprint Data"
        Me.cmdCopyImprintData.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(613, 526)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(89, 33)
        Me.cmdClose.TabIndex = 3
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'dgvExtraFields
        '
        Me.dgvExtraFields.AllowUserToAddRows = False
        Me.dgvExtraFields.AllowUserToDeleteRows = False
        Me.dgvExtraFields.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvExtraFields.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvExtraFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvExtraFields.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvExtraFields.Location = New System.Drawing.Point(12, 12)
        Me.dgvExtraFields.Name = "dgvExtraFields"
        Me.dgvExtraFields.ReadOnly = True
        Me.dgvExtraFields.Size = New System.Drawing.Size(690, 508)
        Me.dgvExtraFields.TabIndex = 5
        '
        'frmRepeatImprint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(714, 571)
        Me.Controls.Add(Me.dgvExtraFields)
        Me.Controls.Add(Me.cmdCopyImprintData)
        Me.Controls.Add(Me.cmdClose)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRepeatImprint"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Repeat Imprint Data"
        CType(Me.dgvExtraFields, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents cmdCopyImprintData As System.Windows.Forms.Button
    Public WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents dgvExtraFields As Orders.XDataGridView
End Class
