<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMixMatchSelector
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
        Me.cmdSubmit = New System.Windows.Forms.Button()
        Me.dgvComponents = New System.Windows.Forms.DataGridView()
        Me.lblQty = New System.Windows.Forms.Label()
        Me.txtQty = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.item_pic = New System.Windows.Forms.DataGridViewImageColumn()
        Me.cmp_group = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.item_no = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.loc = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvComponents, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdSubmit
        '
        Me.cmdSubmit.Location = New System.Drawing.Point(446, 319)
        Me.cmdSubmit.Name = "cmdSubmit"
        Me.cmdSubmit.Size = New System.Drawing.Size(75, 23)
        Me.cmdSubmit.TabIndex = 0
        Me.cmdSubmit.Text = "Submit"
        Me.cmdSubmit.UseVisualStyleBackColor = True
        '
        'dgvComponents
        '
        Me.dgvComponents.AllowUserToAddRows = False
        Me.dgvComponents.AllowUserToDeleteRows = False
        Me.dgvComponents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvComponents.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.item_pic, Me.cmp_group, Me.item_no, Me.loc})
        Me.dgvComponents.Location = New System.Drawing.Point(35, 12)
        Me.dgvComponents.Name = "dgvComponents"
        Me.dgvComponents.Size = New System.Drawing.Size(486, 286)
        Me.dgvComponents.TabIndex = 1
        '
        'lblQty
        '
        Me.lblQty.AutoSize = True
        Me.lblQty.Location = New System.Drawing.Point(32, 319)
        Me.lblQty.Name = "lblQty"
        Me.lblQty.Size = New System.Drawing.Size(46, 13)
        Me.lblQty.TabIndex = 2
        Me.lblQty.Text = "Quantity"
        '
        'txtQty
        '
        Me.txtQty.Location = New System.Drawing.Point(84, 316)
        Me.txtQty.Name = "txtQty"
        Me.txtQty.Size = New System.Drawing.Size(99, 20)
        Me.txtQty.TabIndex = 3
        Me.txtQty.Text = "1"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(351, 319)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'item_pic
        '
        Me.item_pic.HeaderText = "Image"
        Me.item_pic.Name = "item_pic"
        Me.item_pic.ReadOnly = True
        '
        'cmp_group
        '
        Me.cmp_group.HeaderText = "Component Group"
        Me.cmp_group.Name = "cmp_group"
        Me.cmp_group.ReadOnly = True
        Me.cmp_group.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.cmp_group.Width = 130
        '
        'item_no
        '
        Me.item_no.HeaderText = "Item No"
        Me.item_no.Name = "item_no"
        Me.item_no.Width = 130
        '
        'loc
        '
        Me.loc.HeaderText = "Loc"
        Me.loc.Name = "loc"
        Me.loc.ReadOnly = True
        Me.loc.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.loc.Width = 50
        '
        'frmMixMatchSelector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(588, 352)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.txtQty)
        Me.Controls.Add(Me.lblQty)
        Me.Controls.Add(Me.dgvComponents)
        Me.Controls.Add(Me.cmdSubmit)
        Me.Name = "frmMixMatchSelector"
        Me.Text = "Mix Match Selector"
        CType(Me.dgvComponents, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdSubmit As System.Windows.Forms.Button
    Friend WithEvents dgvComponents As System.Windows.Forms.DataGridView
    Friend WithEvents lblQty As System.Windows.Forms.Label
    Friend WithEvents txtQty As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents item_pic As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents cmp_group As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents item_no As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents loc As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
