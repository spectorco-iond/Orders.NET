<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucDocument2
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
        Me.lblPreview = New System.Windows.Forms.Label
        Me.wbShowFile = New System.Windows.Forms.WebBrowser
        Me.lblFiles = New System.Windows.Forms.Label
        Me.lblType = New System.Windows.Forms.Label
        Me.lbTypes = New System.Windows.Forms.ListBox
        Me.cmdCopyDoc = New System.Windows.Forms.Button
        Me.cmdDeleteDoc = New System.Windows.Forms.Button
        Me.cmdRenameDoc = New System.Windows.Forms.Button
        Me.cmdProtectDoc = New System.Windows.Forms.Button
        Me.cmdPreviewDoc = New System.Windows.Forms.Button
        Me.cmdOpenDoc = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.XDataGridView1 = New Orders.XDataGridView
        CType(Me.XDataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblPreview
        '
        Me.lblPreview.AutoSize = True
        Me.lblPreview.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPreview.Location = New System.Drawing.Point(4, 279)
        Me.lblPreview.Name = "lblPreview"
        Me.lblPreview.Size = New System.Drawing.Size(52, 16)
        Me.lblPreview.TabIndex = 42
        Me.lblPreview.Text = "Preview"
        '
        'wbShowFile
        '
        Me.wbShowFile.Location = New System.Drawing.Point(7, 300)
        Me.wbShowFile.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.wbShowFile.MinimumSize = New System.Drawing.Size(27, 31)
        Me.wbShowFile.Name = "wbShowFile"
        Me.wbShowFile.Size = New System.Drawing.Size(950, 439)
        Me.wbShowFile.TabIndex = 39
        '
        'lblFiles
        '
        Me.lblFiles.AutoSize = True
        Me.lblFiles.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFiles.Location = New System.Drawing.Point(448, 5)
        Me.lblFiles.Name = "lblFiles"
        Me.lblFiles.Size = New System.Drawing.Size(36, 16)
        Me.lblFiles.TabIndex = 37
        Me.lblFiles.Text = "Files"
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblType.Location = New System.Drawing.Point(3, 5)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(36, 16)
        Me.lblType.TabIndex = 36
        Me.lblType.Text = "Type"
        '
        'lbTypes
        '
        Me.lbTypes.AllowDrop = True
        Me.lbTypes.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbTypes.FormattingEnabled = True
        Me.lbTypes.ItemHeight = 16
        Me.lbTypes.Location = New System.Drawing.Point(7, 30)
        Me.lbTypes.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.lbTypes.Name = "lbTypes"
        Me.lbTypes.Size = New System.Drawing.Size(244, 244)
        Me.lbTypes.TabIndex = 35
        '
        'cmdCopyDoc
        '
        Me.cmdCopyDoc.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopyDoc.Location = New System.Drawing.Point(619, 250)
        Me.cmdCopyDoc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCopyDoc.Name = "cmdCopyDoc"
        Me.cmdCopyDoc.Size = New System.Drawing.Size(80, 24)
        Me.cmdCopyDoc.TabIndex = 52
        Me.cmdCopyDoc.Text = "Copy"
        Me.cmdCopyDoc.UseVisualStyleBackColor = True
        '
        'cmdDeleteDoc
        '
        Me.cmdDeleteDoc.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeleteDoc.Location = New System.Drawing.Point(705, 250)
        Me.cmdDeleteDoc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdDeleteDoc.Name = "cmdDeleteDoc"
        Me.cmdDeleteDoc.Size = New System.Drawing.Size(80, 24)
        Me.cmdDeleteDoc.TabIndex = 53
        Me.cmdDeleteDoc.Text = "Delete"
        Me.cmdDeleteDoc.UseVisualStyleBackColor = True
        '
        'cmdRenameDoc
        '
        Me.cmdRenameDoc.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRenameDoc.Location = New System.Drawing.Point(791, 250)
        Me.cmdRenameDoc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdRenameDoc.Name = "cmdRenameDoc"
        Me.cmdRenameDoc.Size = New System.Drawing.Size(80, 24)
        Me.cmdRenameDoc.TabIndex = 54
        Me.cmdRenameDoc.Text = "Rename"
        Me.cmdRenameDoc.UseVisualStyleBackColor = True
        '
        'cmdProtectDoc
        '
        Me.cmdProtectDoc.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdProtectDoc.Location = New System.Drawing.Point(877, 250)
        Me.cmdProtectDoc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdProtectDoc.Name = "cmdProtectDoc"
        Me.cmdProtectDoc.Size = New System.Drawing.Size(80, 24)
        Me.cmdProtectDoc.TabIndex = 55
        Me.cmdProtectDoc.Text = "Protect"
        Me.cmdProtectDoc.UseVisualStyleBackColor = True
        '
        'cmdPreviewDoc
        '
        Me.cmdPreviewDoc.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPreviewDoc.Location = New System.Drawing.Point(447, 250)
        Me.cmdPreviewDoc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdPreviewDoc.Name = "cmdPreviewDoc"
        Me.cmdPreviewDoc.Size = New System.Drawing.Size(80, 24)
        Me.cmdPreviewDoc.TabIndex = 56
        Me.cmdPreviewDoc.Text = "Preview"
        Me.cmdPreviewDoc.UseVisualStyleBackColor = True
        '
        'cmdOpenDoc
        '
        Me.cmdOpenDoc.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenDoc.Location = New System.Drawing.Point(533, 250)
        Me.cmdOpenDoc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOpenDoc.Name = "cmdOpenDoc"
        Me.cmdOpenDoc.Size = New System.Drawing.Size(80, 24)
        Me.cmdOpenDoc.TabIndex = 57
        Me.cmdOpenDoc.Text = "Open"
        Me.cmdOpenDoc.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(446, 23)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(511, 215)
        Me.Button1.TabIndex = 59
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'XDataGridView1
        '
        Me.XDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.XDataGridView1.Location = New System.Drawing.Point(7, 30)
        Me.XDataGridView1.Name = "XDataGridView1"
        Me.XDataGridView1.Size = New System.Drawing.Size(244, 244)
        Me.XDataGridView1.TabIndex = 60
        '
        'ucDocument2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.XDataGridView1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cmdOpenDoc)
        Me.Controls.Add(Me.cmdPreviewDoc)
        Me.Controls.Add(Me.cmdProtectDoc)
        Me.Controls.Add(Me.cmdRenameDoc)
        Me.Controls.Add(Me.cmdDeleteDoc)
        Me.Controls.Add(Me.cmdCopyDoc)
        Me.Controls.Add(Me.lblPreview)
        Me.Controls.Add(Me.wbShowFile)
        Me.Controls.Add(Me.lblFiles)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.lbTypes)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "ucDocument2"
        Me.Size = New System.Drawing.Size(960, 761)
        CType(Me.XDataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblPreview As System.Windows.Forms.Label
    Friend WithEvents wbShowFile As System.Windows.Forms.WebBrowser
    Friend WithEvents lblFiles As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents lbTypes As System.Windows.Forms.ListBox
    Friend WithEvents cmdCopyDoc As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteDoc As System.Windows.Forms.Button
    Friend WithEvents cmdRenameDoc As System.Windows.Forms.Button
    Friend WithEvents cmdProtectDoc As System.Windows.Forms.Button
    Friend WithEvents cmdPreviewDoc As System.Windows.Forms.Button
    Friend WithEvents cmdOpenDoc As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents XDataGridView1 As Orders.XDataGridView

End Class
