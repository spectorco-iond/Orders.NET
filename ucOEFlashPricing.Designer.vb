<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucOEFlashPricing
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.chkAllDates = New System.Windows.Forms.CheckBox
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.cmdChargeAdd = New System.Windows.Forms.Button
        Me.dgvPricing = New System.Windows.Forms.DataGridView
        CType(Me.dgvPricing, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkAllDates
        '
        Me.chkAllDates.AutoSize = True
        Me.chkAllDates.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAllDates.Location = New System.Drawing.Point(517, 3)
        Me.chkAllDates.Name = "chkAllDates"
        Me.chkAllDates.Size = New System.Drawing.Size(80, 20)
        Me.chkAllDates.TabIndex = 20
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
        Me.cmdDelete.TabIndex = 19
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
        Me.cmdChargeAdd.TabIndex = 18
        Me.cmdChargeAdd.Text = "Add"
        Me.cmdChargeAdd.UseVisualStyleBackColor = False
        '
        'dgvPricing
        '
        Me.dgvPricing.AllowUserToAddRows = False
        Me.dgvPricing.AllowUserToDeleteRows = False
        Me.dgvPricing.AllowUserToResizeRows = False
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgvPricing.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle6
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvPricing.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvPricing.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvPricing.DefaultCellStyle = DataGridViewCellStyle8
        Me.dgvPricing.Location = New System.Drawing.Point(0, 31)
        Me.dgvPricing.MultiSelect = False
        Me.dgvPricing.Name = "dgvPricing"
        Me.dgvPricing.RowHeadersWidth = 20
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvPricing.RowsDefaultCellStyle = DataGridViewCellStyle9
        Me.dgvPricing.Size = New System.Drawing.Size(729, 324)
        Me.dgvPricing.TabIndex = 17
        Me.dgvPricing.TabStop = False
        '
        'ucOEFlashPricing
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.Controls.Add(Me.chkAllDates)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdChargeAdd)
        Me.Controls.Add(Me.dgvPricing)
        Me.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ucOEFlashPricing"
        Me.Size = New System.Drawing.Size(732, 358)
        CType(Me.dgvPricing, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkAllDates As System.Windows.Forms.CheckBox
    Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents cmdChargeAdd As System.Windows.Forms.Button
    Friend WithEvents dgvPricing As System.Windows.Forms.DataGridView

End Class
