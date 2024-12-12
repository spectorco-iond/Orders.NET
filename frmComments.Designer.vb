<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmComments
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.panButtons = New System.Windows.Forms.Panel
        Me.cmdDeleteAll = New System.Windows.Forms.Button
        Me.cmdAutoComments = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdNew = New System.Windows.Forms.Button
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.dgvComments = New System.Windows.Forms.DataGridView
        Me.lblOEI_Ord_No = New System.Windows.Forms.Label
        Me.lblCopyFromPO = New System.Windows.Forms.Label
        Me.cmdCopyFromPO = New System.Windows.Forms.Button
        Me.cmdMoveUp = New System.Windows.Forms.Button
        Me.cmdMoveDown = New System.Windows.Forms.Button
        Me.txtCopyFromPO = New Orders.xTextBox
        Me.txtNew_Comment = New Orders.xTextBox
        Me.txtOEI_Ord_No = New Orders.xTextBox
        Me.panButtons.SuspendLayout()
        CType(Me.dgvComments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'panButtons
        '
        Me.panButtons.Controls.Add(Me.cmdDeleteAll)
        Me.panButtons.Controls.Add(Me.cmdAutoComments)
        Me.panButtons.Controls.Add(Me.cmdSave)
        Me.panButtons.Controls.Add(Me.cmdNew)
        Me.panButtons.Controls.Add(Me.cmdDelete)
        Me.panButtons.Controls.Add(Me.cmdClose)
        Me.panButtons.Location = New System.Drawing.Point(13, 286)
        Me.panButtons.Name = "panButtons"
        Me.panButtons.Size = New System.Drawing.Size(386, 26)
        Me.panButtons.TabIndex = 3
        '
        'cmdDeleteAll
        '
        Me.cmdDeleteAll.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDeleteAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDeleteAll.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeleteAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDeleteAll.Location = New System.Drawing.Point(185, -1)
        Me.cmdDeleteAll.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdDeleteAll.Name = "cmdDeleteAll"
        Me.cmdDeleteAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDeleteAll.Size = New System.Drawing.Size(72, 26)
        Me.cmdDeleteAll.TabIndex = 5
        Me.cmdDeleteAll.Text = "Delete All"
        Me.cmdDeleteAll.UseVisualStyleBackColor = False
        '
        'cmdAutoComments
        '
        Me.cmdAutoComments.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAutoComments.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAutoComments.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAutoComments.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAutoComments.Location = New System.Drawing.Point(261, -1)
        Me.cmdAutoComments.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdAutoComments.Name = "cmdAutoComments"
        Me.cmdAutoComments.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAutoComments.Size = New System.Drawing.Size(58, 26)
        Me.cmdAutoComments.TabIndex = 3
        Me.cmdAutoComments.Text = "&Auto"
        Me.cmdAutoComments.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Location = New System.Drawing.Point(-1, -1)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(58, 26)
        Me.cmdSave.TabIndex = 0
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdNew
        '
        Me.cmdNew.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNew.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNew.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNew.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNew.Location = New System.Drawing.Point(61, -1)
        Me.cmdNew.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdNew.Name = "cmdNew"
        Me.cmdNew.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNew.Size = New System.Drawing.Size(58, 26)
        Me.cmdNew.TabIndex = 1
        Me.cmdNew.Text = "&New"
        Me.cmdNew.UseVisualStyleBackColor = False
        '
        'cmdDelete
        '
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(123, -1)
        Me.cmdDelete.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(58, 26)
        Me.cmdDelete.TabIndex = 2
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(324, -1)
        Me.cmdClose.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(62, 26)
        Me.cmdClose.TabIndex = 4
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'dgvComments
        '
        Me.dgvComments.AllowUserToAddRows = False
        Me.dgvComments.AllowUserToDeleteRows = False
        Me.dgvComments.AllowUserToResizeRows = False
        Me.dgvComments.BackgroundColor = System.Drawing.SystemColors.Control
        Me.dgvComments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvComments.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvComments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvComments.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvComments.Location = New System.Drawing.Point(13, 93)
        Me.dgvComments.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvComments.MultiSelect = False
        Me.dgvComments.Name = "dgvComments"
        Me.dgvComments.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvComments.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvComments.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvComments.Size = New System.Drawing.Size(386, 184)
        Me.dgvComments.TabIndex = 3
        '
        'lblOEI_Ord_No
        '
        Me.lblOEI_Ord_No.BackColor = System.Drawing.SystemColors.Control
        Me.lblOEI_Ord_No.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOEI_Ord_No.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblOEI_Ord_No.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOEI_Ord_No.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblOEI_Ord_No.Location = New System.Drawing.Point(10, 15)
        Me.lblOEI_Ord_No.Name = "lblOEI_Ord_No"
        Me.lblOEI_Ord_No.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOEI_Ord_No.Size = New System.Drawing.Size(77, 20)
        Me.lblOEI_Ord_No.TabIndex = 0
        Me.lblOEI_Ord_No.Text = "Order"
        '
        'lblCopyFromPO
        '
        Me.lblCopyFromPO.BackColor = System.Drawing.SystemColors.Control
        Me.lblCopyFromPO.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCopyFromPO.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblCopyFromPO.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCopyFromPO.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblCopyFromPO.Location = New System.Drawing.Point(10, 41)
        Me.lblCopyFromPO.Name = "lblCopyFromPO"
        Me.lblCopyFromPO.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCopyFromPO.Size = New System.Drawing.Size(104, 20)
        Me.lblCopyFromPO.TabIndex = 4
        Me.lblCopyFromPO.Text = "Copy From P.O."
        '
        'cmdCopyFromPO
        '
        Me.cmdCopyFromPO.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopyFromPO.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCopyFromPO.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopyFromPO.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopyFromPO.Location = New System.Drawing.Point(337, 35)
        Me.cmdCopyFromPO.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdCopyFromPO.Name = "cmdCopyFromPO"
        Me.cmdCopyFromPO.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCopyFromPO.Size = New System.Drawing.Size(62, 26)
        Me.cmdCopyFromPO.TabIndex = 6
        Me.cmdCopyFromPO.TabStop = False
        Me.cmdCopyFromPO.Text = "Copy"
        Me.cmdCopyFromPO.UseVisualStyleBackColor = False
        '
        'cmdMoveUp
        '
        Me.cmdMoveUp.BackColor = System.Drawing.SystemColors.Control
        Me.cmdMoveUp.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMoveUp.Font = New System.Drawing.Font("Wingdings 3", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.cmdMoveUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMoveUp.Location = New System.Drawing.Point(336, 64)
        Me.cmdMoveUp.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdMoveUp.Name = "cmdMoveUp"
        Me.cmdMoveUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMoveUp.Size = New System.Drawing.Size(29, 26)
        Me.cmdMoveUp.TabIndex = 7
        Me.cmdMoveUp.TabStop = False
        Me.cmdMoveUp.Text = "h"
        Me.cmdMoveUp.UseVisualStyleBackColor = False
        '
        'cmdMoveDown
        '
        Me.cmdMoveDown.BackColor = System.Drawing.SystemColors.Control
        Me.cmdMoveDown.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMoveDown.Font = New System.Drawing.Font("Wingdings 3", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.cmdMoveDown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMoveDown.Location = New System.Drawing.Point(370, 64)
        Me.cmdMoveDown.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdMoveDown.Name = "cmdMoveDown"
        Me.cmdMoveDown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMoveDown.Size = New System.Drawing.Size(29, 26)
        Me.cmdMoveDown.TabIndex = 8
        Me.cmdMoveDown.TabStop = False
        Me.cmdMoveDown.Text = "i"
        Me.cmdMoveDown.UseVisualStyleBackColor = False
        '
        'txtCopyFromPO
        '
        Me.txtCopyFromPO.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtCopyFromPO.DataLength = CType(25, Long)
        Me.txtCopyFromPO.DataType = Orders.CDataType.dtString
        Me.txtCopyFromPO.DateValue = New Date(CType(0, Long))
        Me.txtCopyFromPO.DecimalDigits = CType(0, Long)
        Me.txtCopyFromPO.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCopyFromPO.Location = New System.Drawing.Point(112, 38)
        Me.txtCopyFromPO.Mask = Nothing
        Me.txtCopyFromPO.Name = "txtCopyFromPO"
        Me.txtCopyFromPO.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtCopyFromPO.OldValue = Nothing
        Me.txtCopyFromPO.Size = New System.Drawing.Size(218, 22)
        Me.txtCopyFromPO.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtCopyFromPO.StringValue = Nothing
        Me.txtCopyFromPO.TabIndex = 5
        Me.txtCopyFromPO.TextDataID = Nothing
        '
        'txtNew_Comment
        '
        Me.txtNew_Comment.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtNew_Comment.DataLength = CType(40, Long)
        Me.txtNew_Comment.DataType = Orders.CDataType.dtString
        Me.txtNew_Comment.DateValue = New Date(CType(0, Long))
        Me.txtNew_Comment.DecimalDigits = CType(0, Long)
        Me.txtNew_Comment.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNew_Comment.Location = New System.Drawing.Point(11, 64)
        Me.txtNew_Comment.Mask = Nothing
        Me.txtNew_Comment.Name = "txtNew_Comment"
        Me.txtNew_Comment.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtNew_Comment.OldValue = Nothing
        Me.txtNew_Comment.Size = New System.Drawing.Size(319, 22)
        Me.txtNew_Comment.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtNew_Comment.StringValue = Nothing
        Me.txtNew_Comment.TabIndex = 2
        Me.txtNew_Comment.TextDataID = Nothing
        '
        'txtOEI_Ord_No
        '
        Me.txtOEI_Ord_No.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtOEI_Ord_No.DataLength = CType(25, Long)
        Me.txtOEI_Ord_No.DataType = Orders.CDataType.dtString
        Me.txtOEI_Ord_No.DateValue = New Date(CType(0, Long))
        Me.txtOEI_Ord_No.DecimalDigits = CType(0, Long)
        Me.txtOEI_Ord_No.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOEI_Ord_No.Location = New System.Drawing.Point(112, 12)
        Me.txtOEI_Ord_No.Mask = Nothing
        Me.txtOEI_Ord_No.Name = "txtOEI_Ord_No"
        Me.txtOEI_Ord_No.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtOEI_Ord_No.OldValue = Nothing
        Me.txtOEI_Ord_No.ReadOnly = True
        Me.txtOEI_Ord_No.Size = New System.Drawing.Size(218, 22)
        Me.txtOEI_Ord_No.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtOEI_Ord_No.StringValue = Nothing
        Me.txtOEI_Ord_No.TabIndex = 1
        Me.txtOEI_Ord_No.TextDataID = Nothing
        '
        'frmComments
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(412, 326)
        Me.Controls.Add(Me.cmdMoveDown)
        Me.Controls.Add(Me.cmdMoveUp)
        Me.Controls.Add(Me.cmdCopyFromPO)
        Me.Controls.Add(Me.txtCopyFromPO)
        Me.Controls.Add(Me.lblCopyFromPO)
        Me.Controls.Add(Me.txtNew_Comment)
        Me.Controls.Add(Me.txtOEI_Ord_No)
        Me.Controls.Add(Me.lblOEI_Ord_No)
        Me.Controls.Add(Me.panButtons)
        Me.Controls.Add(Me.dgvComments)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(420, 540)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(420, 360)
        Me.Name = "frmComments"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Comments"
        Me.panButtons.ResumeLayout(False)
        CType(Me.dgvComments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents panButtons As System.Windows.Forms.Panel
    Public WithEvents cmdNew As System.Windows.Forms.Button
    Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents dgvComments As System.Windows.Forms.DataGridView
    Friend WithEvents txtOEI_Ord_No As Orders.xTextBox
    Friend WithEvents lblOEI_Ord_No As System.Windows.Forms.Label
    Friend WithEvents txtNew_Comment As Orders.xTextBox
    Public WithEvents cmdSave As System.Windows.Forms.Button
    Public WithEvents cmdAutoComments As System.Windows.Forms.Button
    Friend WithEvents txtCopyFromPO As Orders.xTextBox
    Friend WithEvents lblCopyFromPO As System.Windows.Forms.Label
    Public WithEvents cmdCopyFromPO As System.Windows.Forms.Button
    Public WithEvents cmdMoveUp As System.Windows.Forms.Button
    Public WithEvents cmdMoveDown As System.Windows.Forms.Button
    Public WithEvents cmdDeleteAll As System.Windows.Forms.Button
End Class
