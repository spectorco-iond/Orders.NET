<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucOEFlashCharges
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.chkAllDates = New System.Windows.Forms.CheckBox
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.cmdChargeAdd = New System.Windows.Forms.Button
        Me.dgvCharges = New System.Windows.Forms.DataGridView
        CType(Me.dgvCharges, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkAllDates
        '
        Me.chkAllDates.AutoSize = True
        Me.chkAllDates.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAllDates.Location = New System.Drawing.Point(517, 3)
        Me.chkAllDates.Name = "chkAllDates"
        Me.chkAllDates.Size = New System.Drawing.Size(80, 20)
        Me.chkAllDates.TabIndex = 16
        Me.chkAllDates.Text = "All Dates"
        Me.chkAllDates.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(669, 0)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(60, 25)
        Me.cmdDelete.TabIndex = 15
        Me.cmdDelete.Text = "Delete"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'cmdChargeAdd
        '
        Me.cmdChargeAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdChargeAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdChargeAdd.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdChargeAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdChargeAdd.Location = New System.Drawing.Point(603, 0)
        Me.cmdChargeAdd.Name = "cmdChargeAdd"
        Me.cmdChargeAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdChargeAdd.Size = New System.Drawing.Size(60, 25)
        Me.cmdChargeAdd.TabIndex = 14
        Me.cmdChargeAdd.Text = "Add"
        Me.cmdChargeAdd.UseVisualStyleBackColor = False
        '
        'dgvCharges
        '
        Me.dgvCharges.AllowUserToAddRows = False
        Me.dgvCharges.AllowUserToDeleteRows = False
        Me.dgvCharges.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgvCharges.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvCharges.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvCharges.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCharges.Location = New System.Drawing.Point(0, 31)
        Me.dgvCharges.MultiSelect = False
        Me.dgvCharges.Name = "dgvCharges"
        Me.dgvCharges.RowHeadersWidth = 20
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvCharges.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvCharges.Size = New System.Drawing.Size(729, 324)
        Me.dgvCharges.TabIndex = 13
        Me.dgvCharges.TabStop = False
        '
        'ucOEFlashCharges
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.Controls.Add(Me.chkAllDates)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdChargeAdd)
        Me.Controls.Add(Me.dgvCharges)
        Me.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ucOEFlashCharges"
        Me.Size = New System.Drawing.Size(735, 359)
        CType(Me.dgvCharges, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkAllDates As System.Windows.Forms.CheckBox
    Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents cmdChargeAdd As System.Windows.Forms.Button
    Friend WithEvents dgvCharges As System.Windows.Forms.DataGridView

End Class
