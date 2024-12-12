<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProductLineEntry
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
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdApplyToAll = New System.Windows.Forms.Button
        Me.cmdResetAll = New System.Windows.Forms.Button
        Me.dgvOrderLines = New System.Windows.Forms.DataGridView
        CType(Me.dgvOrderLines, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(714, 447)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(100, 24)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.Text = "Add to order"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(820, 447)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(100, 24)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdApplyToAll
        '
        Me.cmdApplyToAll.Location = New System.Drawing.Point(12, 447)
        Me.cmdApplyToAll.Name = "cmdApplyToAll"
        Me.cmdApplyToAll.Size = New System.Drawing.Size(100, 24)
        Me.cmdApplyToAll.TabIndex = 3
        Me.cmdApplyToAll.Text = "Apply to all"
        Me.cmdApplyToAll.UseVisualStyleBackColor = True
        '
        'cmdResetAll
        '
        Me.cmdResetAll.Location = New System.Drawing.Point(116, 447)
        Me.cmdResetAll.Name = "cmdResetAll"
        Me.cmdResetAll.Size = New System.Drawing.Size(100, 24)
        Me.cmdResetAll.TabIndex = 4
        Me.cmdResetAll.Text = "Reset All"
        Me.cmdResetAll.UseVisualStyleBackColor = True
        '
        'dgvOrderLines
        '
        Me.dgvOrderLines.AllowUserToAddRows = False
        Me.dgvOrderLines.AllowUserToDeleteRows = False
        Me.dgvOrderLines.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.dgvOrderLines.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOrderLines.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke
        Me.dgvOrderLines.Location = New System.Drawing.Point(12, 12)
        Me.dgvOrderLines.Name = "dgvOrderLines"
        Me.dgvOrderLines.Size = New System.Drawing.Size(908, 429)
        Me.dgvOrderLines.TabIndex = 8
        '
        'frmProductLineEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(932, 479)
        Me.Controls.Add(Me.dgvOrderLines)
        Me.Controls.Add(Me.cmdResetAll)
        Me.Controls.Add(Me.cmdApplyToAll)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(500, 200)
        Me.Name = "frmProductLineEntry"
        Me.Text = "frmProductLineEntry"
        CType(Me.dgvOrderLines, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdApplyToAll As System.Windows.Forms.Button
    Friend WithEvents cmdResetAll As System.Windows.Forms.Button
    Friend WithEvents dgvOrderLines As System.Windows.Forms.DataGridView
End Class
