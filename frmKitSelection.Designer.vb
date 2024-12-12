<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmKitSelection
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtItem_No = New System.Windows.Forms.TextBox
        Me.txtDescription_1 = New System.Windows.Forms.TextBox
        Me.txtDescription_2 = New System.Windows.Forms.TextBox
        Me.dgvItems = New System.Windows.Forms.DataGridView
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdNew = New System.Windows.Forms.Button
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.txtUnit_Price = New Orders.xTextBox
        CType(Me.dgvItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Kit Item"
        '
        'txtItem_No
        '
        Me.txtItem_No.Location = New System.Drawing.Point(105, 12)
        Me.txtItem_No.Name = "txtItem_No"
        Me.txtItem_No.Size = New System.Drawing.Size(252, 22)
        Me.txtItem_No.TabIndex = 2
        Me.txtItem_No.TabStop = False
        '
        'txtDescription_1
        '
        Me.txtDescription_1.Location = New System.Drawing.Point(105, 40)
        Me.txtDescription_1.Name = "txtDescription_1"
        Me.txtDescription_1.Size = New System.Drawing.Size(252, 22)
        Me.txtDescription_1.TabIndex = 3
        Me.txtDescription_1.TabStop = False
        '
        'txtDescription_2
        '
        Me.txtDescription_2.Location = New System.Drawing.Point(105, 68)
        Me.txtDescription_2.Name = "txtDescription_2"
        Me.txtDescription_2.Size = New System.Drawing.Size(252, 22)
        Me.txtDescription_2.TabIndex = 4
        Me.txtDescription_2.TabStop = False
        '
        'dgvItems
        '
        Me.dgvItems.AllowUserToAddRows = False
        Me.dgvItems.AllowUserToDeleteRows = False
        Me.dgvItems.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.dgvItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItems.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke
        Me.dgvItems.Location = New System.Drawing.Point(15, 105)
        Me.dgvItems.MultiSelect = False
        Me.dgvItems.Name = "dgvItems"
        Me.dgvItems.ReadOnly = True
        Me.dgvItems.RowHeadersWidth = 20
        Me.dgvItems.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItems.Size = New System.Drawing.Size(836, 408)
        Me.dgvItems.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(549, 533)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(97, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Unit Price Total"
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(789, 568)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(62, 27)
        Me.cmdClose.TabIndex = 12
        Me.cmdClose.TabStop = False
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Enabled = False
        Me.cmdSave.Location = New System.Drawing.Point(585, 568)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(62, 27)
        Me.cmdSave.TabIndex = 13
        Me.cmdSave.TabStop = False
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        Me.cmdSave.Visible = False
        '
        'cmdNew
        '
        Me.cmdNew.Enabled = False
        Me.cmdNew.Location = New System.Drawing.Point(653, 568)
        Me.cmdNew.Name = "cmdNew"
        Me.cmdNew.Size = New System.Drawing.Size(62, 27)
        Me.cmdNew.TabIndex = 14
        Me.cmdNew.TabStop = False
        Me.cmdNew.Text = "New"
        Me.cmdNew.UseVisualStyleBackColor = True
        Me.cmdNew.Visible = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Enabled = False
        Me.cmdDelete.Location = New System.Drawing.Point(721, 568)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(62, 27)
        Me.cmdDelete.TabIndex = 15
        Me.cmdDelete.TabStop = False
        Me.cmdDelete.Text = "Delete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        Me.cmdDelete.Visible = False
        '
        'txtUnit_Price
        '
        Me.txtUnit_Price.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtUnit_Price.DataLength = CType(0, Long)
        Me.txtUnit_Price.DataType = Orders.CDataType.dtNumWithDecimals
        Me.txtUnit_Price.DateValue = New Date(CType(0, Long))
        Me.txtUnit_Price.DecimalDigits = CType(6, Long)
        Me.txtUnit_Price.Location = New System.Drawing.Point(681, 530)
        Me.txtUnit_Price.Mask = Nothing
        Me.txtUnit_Price.Name = "txtUnit_Price"
        Me.txtUnit_Price.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtUnit_Price.OldValue = ""
        Me.txtUnit_Price.Size = New System.Drawing.Size(170, 22)
        Me.txtUnit_Price.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtUnit_Price.StringValue = Nothing
        Me.txtUnit_Price.TabIndex = 16
        Me.txtUnit_Price.TabStop = False
        Me.txtUnit_Price.TextDataID = Nothing
        '
        'frmKitSelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(863, 607)
        Me.Controls.Add(Me.txtUnit_Price)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdNew)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dgvItems)
        Me.Controls.Add(Me.txtDescription_2)
        Me.Controls.Add(Me.txtDescription_1)
        Me.Controls.Add(Me.txtItem_No)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmKitSelection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Kit Selection"
        CType(Me.dgvItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtItem_No As System.Windows.Forms.TextBox
    Friend WithEvents txtDescription_1 As System.Windows.Forms.TextBox
    Friend WithEvents txtDescription_2 As System.Windows.Forms.TextBox
    Friend WithEvents dgvItems As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdNew As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents txtUnit_Price As Orders.xTextBox
End Class
