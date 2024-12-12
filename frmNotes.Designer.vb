<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNotes
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
        Me.cmdClose = New System.Windows.Forms.Button
        Me.dgvNotes = New System.Windows.Forms.DataGridView
        Me.panButtons = New System.Windows.Forms.Panel
        Me.cmdNew = New System.Windows.Forms.Button
        Me.cmdDelete = New System.Windows.Forms.Button
        CType(Me.dgvNotes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(561, -1)
        Me.cmdClose.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(81, 26)
        Me.cmdClose.TabIndex = 0
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'dgvNotes
        '
        Me.dgvNotes.AllowUserToAddRows = False
        Me.dgvNotes.AllowUserToDeleteRows = False
        Me.dgvNotes.AllowUserToResizeRows = False
        Me.dgvNotes.BackgroundColor = System.Drawing.SystemColors.Control
        Me.dgvNotes.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvNotes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvNotes.Location = New System.Drawing.Point(12, 15)
        Me.dgvNotes.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvNotes.MultiSelect = False
        Me.dgvNotes.Name = "dgvNotes"
        Me.dgvNotes.ReadOnly = True
        Me.dgvNotes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvNotes.Size = New System.Drawing.Size(642, 204)
        Me.dgvNotes.TabIndex = 0
        '
        'panButtons
        '
        Me.panButtons.Controls.Add(Me.cmdNew)
        Me.panButtons.Controls.Add(Me.cmdDelete)
        Me.panButtons.Controls.Add(Me.cmdClose)
        Me.panButtons.Location = New System.Drawing.Point(12, 228)
        Me.panButtons.Name = "panButtons"
        Me.panButtons.Size = New System.Drawing.Size(642, 26)
        Me.panButtons.TabIndex = 1
        '
        'cmdNew
        '
        Me.cmdNew.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNew.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNew.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNew.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNew.Location = New System.Drawing.Point(383, -1)
        Me.cmdNew.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdNew.Name = "cmdNew"
        Me.cmdNew.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNew.Size = New System.Drawing.Size(81, 26)
        Me.cmdNew.TabIndex = 1
        Me.cmdNew.Text = "New"
        Me.cmdNew.UseVisualStyleBackColor = False
        '
        'cmdDelete
        '
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(472, -1)
        Me.cmdDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(81, 26)
        Me.cmdDelete.TabIndex = 2
        Me.cmdDelete.Text = "Delete"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'frmNotes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(667, 266)
        Me.Controls.Add(Me.panButtons)
        Me.Controls.Add(Me.dgvNotes)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(675, 300)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(675, 300)
        Me.Name = "frmNotes"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Notes"
        CType(Me.dgvNotes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents dgvNotes As System.Windows.Forms.DataGridView
    Friend WithEvents panButtons As System.Windows.Forms.Panel
    Public WithEvents cmdNew As System.Windows.Forms.Button
    Public WithEvents cmdDelete As System.Windows.Forms.Button
End Class
