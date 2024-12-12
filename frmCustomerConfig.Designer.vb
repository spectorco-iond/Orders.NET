<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCustomerConfig
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
        Dim DataGridViewCellStyle21 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle22 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle23 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle24 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle25 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle26 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle27 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle28 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle29 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle30 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.tcOEFlash = New System.Windows.Forms.TabControl
        Me.OETab_General = New System.Windows.Forms.TabPage
        Me.dgvGeneral = New System.Windows.Forms.DataGridView
        Me.chkOEFlash = New System.Windows.Forms.CheckBox
        Me.dgvOEFlash = New System.Windows.Forms.DataGridView
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.OETab_Charges = New System.Windows.Forms.TabPage
        Me.OETab_Pricing = New System.Windows.Forms.TabPage
        Me.OETab_Shipping = New System.Windows.Forms.TabPage
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.dgvShipping = New System.Windows.Forms.DataGridView
        Me.OETab_Correspondance = New System.Windows.Forms.TabPage
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.dgvCorrespondance = New System.Windows.Forms.DataGridView
        Me.fraCustomerInfo = New System.Windows.Forms.GroupBox
        Me.txtClassID = New System.Windows.Forms.TextBox
        Me.txtEQP = New System.Windows.Forms.TextBox
        Me.txtCusNo = New System.Windows.Forms.TextBox
        Me.txtCusName = New System.Windows.Forms.TextBox
        Me.lblCusNo = New System.Windows.Forms.Label
        Me.Uc_OEFlashGeneral = New Orders.ucOEFlashCustomer
        Me.uc_OEFlashCharges = New Orders.ucOEFlashCharges
        Me.Uc_OEFlashShipping = New Orders.ucOEFlashShipping
        Me.Uc_OEFlashCorrespondance = New Orders.ucOEFlashCorrespondance
        Me.tcOEFlash.SuspendLayout()
        Me.OETab_General.SuspendLayout()
        CType(Me.dgvGeneral, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvOEFlash, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OETab_Charges.SuspendLayout()
        Me.OETab_Shipping.SuspendLayout()
        CType(Me.dgvShipping, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OETab_Correspondance.SuspendLayout()
        CType(Me.dgvCorrespondance, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraCustomerInfo.SuspendLayout()
        Me.SuspendLayout()
        '
        'tcOEFlash
        '
        Me.tcOEFlash.Controls.Add(Me.OETab_General)
        Me.tcOEFlash.Controls.Add(Me.OETab_Charges)
        Me.tcOEFlash.Controls.Add(Me.OETab_Pricing)
        Me.tcOEFlash.Controls.Add(Me.OETab_Shipping)
        Me.tcOEFlash.Controls.Add(Me.OETab_Correspondance)
        Me.tcOEFlash.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tcOEFlash.Location = New System.Drawing.Point(12, 92)
        Me.tcOEFlash.Name = "tcOEFlash"
        Me.tcOEFlash.SelectedIndex = 0
        Me.tcOEFlash.Size = New System.Drawing.Size(747, 398)
        Me.tcOEFlash.TabIndex = 12
        '
        'OETab_General
        '
        Me.OETab_General.Controls.Add(Me.Uc_OEFlashGeneral)
        Me.OETab_General.Controls.Add(Me.dgvGeneral)
        Me.OETab_General.Controls.Add(Me.chkOEFlash)
        Me.OETab_General.Controls.Add(Me.dgvOEFlash)
        Me.OETab_General.Controls.Add(Me.Button7)
        Me.OETab_General.Controls.Add(Me.Button8)
        Me.OETab_General.Location = New System.Drawing.Point(4, 25)
        Me.OETab_General.Name = "OETab_General"
        Me.OETab_General.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_General.Size = New System.Drawing.Size(739, 369)
        Me.OETab_General.TabIndex = 0
        Me.OETab_General.Text = "General"
        Me.OETab_General.UseVisualStyleBackColor = True
        '
        'dgvGeneral
        '
        Me.dgvGeneral.AllowUserToAddRows = False
        Me.dgvGeneral.AllowUserToDeleteRows = False
        Me.dgvGeneral.AllowUserToResizeRows = False
        DataGridViewCellStyle21.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvGeneral.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle21
        DataGridViewCellStyle22.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle22.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle22.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle22.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle22.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle22.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle22.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGeneral.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle22
        Me.dgvGeneral.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGeneral.Location = New System.Drawing.Point(3, 37)
        Me.dgvGeneral.Name = "dgvGeneral"
        Me.dgvGeneral.RowHeadersWidth = 20
        DataGridViewCellStyle23.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvGeneral.RowsDefaultCellStyle = DataGridViewCellStyle23
        Me.dgvGeneral.Size = New System.Drawing.Size(729, 324)
        Me.dgvGeneral.TabIndex = 1
        Me.dgvGeneral.TabStop = False
        '
        'chkOEFlash
        '
        Me.chkOEFlash.AutoSize = True
        Me.chkOEFlash.Location = New System.Drawing.Point(519, 9)
        Me.chkOEFlash.Name = "chkOEFlash"
        Me.chkOEFlash.Size = New System.Drawing.Size(81, 20)
        Me.chkOEFlash.TabIndex = 11
        Me.chkOEFlash.Text = "Old Style"
        Me.chkOEFlash.UseVisualStyleBackColor = True
        '
        'dgvOEFlash
        '
        Me.dgvOEFlash.AllowUserToAddRows = False
        Me.dgvOEFlash.AllowUserToDeleteRows = False
        Me.dgvOEFlash.AllowUserToResizeRows = False
        DataGridViewCellStyle24.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvOEFlash.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle24
        DataGridViewCellStyle25.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle25.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle25.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle25.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle25.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle25.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle25.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvOEFlash.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle25
        Me.dgvOEFlash.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOEFlash.Location = New System.Drawing.Point(3, 37)
        Me.dgvOEFlash.Name = "dgvOEFlash"
        Me.dgvOEFlash.RowHeadersWidth = 20
        DataGridViewCellStyle26.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvOEFlash.RowsDefaultCellStyle = DataGridViewCellStyle26
        Me.dgvOEFlash.Size = New System.Drawing.Size(729, 324)
        Me.dgvOEFlash.TabIndex = 10
        Me.dgvOEFlash.TabStop = False
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.SystemColors.Control
        Me.Button7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button7.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button7.Location = New System.Drawing.Point(672, 6)
        Me.Button7.Name = "Button7"
        Me.Button7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button7.Size = New System.Drawing.Size(60, 25)
        Me.Button7.TabIndex = 9
        Me.Button7.Text = "Delete"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'Button8
        '
        Me.Button8.BackColor = System.Drawing.SystemColors.Control
        Me.Button8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button8.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button8.Location = New System.Drawing.Point(606, 6)
        Me.Button8.Name = "Button8"
        Me.Button8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button8.Size = New System.Drawing.Size(60, 25)
        Me.Button8.TabIndex = 8
        Me.Button8.Text = "Add"
        Me.Button8.UseVisualStyleBackColor = False
        '
        'OETab_Charges
        '
        Me.OETab_Charges.Controls.Add(Me.uc_OEFlashCharges)
        Me.OETab_Charges.Location = New System.Drawing.Point(4, 25)
        Me.OETab_Charges.Name = "OETab_Charges"
        Me.OETab_Charges.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_Charges.Size = New System.Drawing.Size(739, 369)
        Me.OETab_Charges.TabIndex = 1
        Me.OETab_Charges.Text = "Charges"
        Me.OETab_Charges.UseVisualStyleBackColor = True
        '
        'OETab_Pricing
        '
        Me.OETab_Pricing.Location = New System.Drawing.Point(4, 25)
        Me.OETab_Pricing.Name = "OETab_Pricing"
        Me.OETab_Pricing.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_Pricing.Size = New System.Drawing.Size(739, 369)
        Me.OETab_Pricing.TabIndex = 2
        Me.OETab_Pricing.Text = "Pricing"
        Me.OETab_Pricing.UseVisualStyleBackColor = True
        '
        'OETab_Shipping
        '
        Me.OETab_Shipping.Controls.Add(Me.Uc_OEFlashShipping)
        Me.OETab_Shipping.Controls.Add(Me.Button3)
        Me.OETab_Shipping.Controls.Add(Me.Button4)
        Me.OETab_Shipping.Controls.Add(Me.dgvShipping)
        Me.OETab_Shipping.Location = New System.Drawing.Point(4, 25)
        Me.OETab_Shipping.Name = "OETab_Shipping"
        Me.OETab_Shipping.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_Shipping.Size = New System.Drawing.Size(739, 369)
        Me.OETab_Shipping.TabIndex = 3
        Me.OETab_Shipping.Text = "Shipping"
        Me.OETab_Shipping.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.SystemColors.Control
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button3.Location = New System.Drawing.Point(672, 6)
        Me.Button3.Name = "Button3"
        Me.Button3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button3.Size = New System.Drawing.Size(60, 25)
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "Delete"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.SystemColors.Control
        Me.Button4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button4.Location = New System.Drawing.Point(606, 6)
        Me.Button4.Name = "Button4"
        Me.Button4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button4.Size = New System.Drawing.Size(60, 25)
        Me.Button4.TabIndex = 8
        Me.Button4.Text = "Add"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'dgvShipping
        '
        Me.dgvShipping.AllowUserToAddRows = False
        Me.dgvShipping.AllowUserToDeleteRows = False
        Me.dgvShipping.AllowUserToResizeRows = False
        DataGridViewCellStyle27.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvShipping.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle27
        Me.dgvShipping.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvShipping.Location = New System.Drawing.Point(3, 37)
        Me.dgvShipping.Name = "dgvShipping"
        Me.dgvShipping.RowHeadersWidth = 20
        DataGridViewCellStyle28.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvShipping.RowsDefaultCellStyle = DataGridViewCellStyle28
        Me.dgvShipping.Size = New System.Drawing.Size(729, 324)
        Me.dgvShipping.TabIndex = 1
        Me.dgvShipping.TabStop = False
        '
        'OETab_Correspondance
        '
        Me.OETab_Correspondance.Controls.Add(Me.Uc_OEFlashCorrespondance)
        Me.OETab_Correspondance.Controls.Add(Me.Button5)
        Me.OETab_Correspondance.Controls.Add(Me.Button6)
        Me.OETab_Correspondance.Controls.Add(Me.dgvCorrespondance)
        Me.OETab_Correspondance.Location = New System.Drawing.Point(4, 25)
        Me.OETab_Correspondance.Name = "OETab_Correspondance"
        Me.OETab_Correspondance.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_Correspondance.Size = New System.Drawing.Size(739, 369)
        Me.OETab_Correspondance.TabIndex = 4
        Me.OETab_Correspondance.Text = "Correspondence"
        Me.OETab_Correspondance.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.SystemColors.Control
        Me.Button5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button5.Location = New System.Drawing.Point(672, 6)
        Me.Button5.Name = "Button5"
        Me.Button5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button5.Size = New System.Drawing.Size(60, 25)
        Me.Button5.TabIndex = 9
        Me.Button5.Text = "Delete"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.SystemColors.Control
        Me.Button6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button6.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button6.Location = New System.Drawing.Point(606, 6)
        Me.Button6.Name = "Button6"
        Me.Button6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button6.Size = New System.Drawing.Size(60, 25)
        Me.Button6.TabIndex = 8
        Me.Button6.Text = "Add"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'dgvCorrespondance
        '
        Me.dgvCorrespondance.AllowUserToAddRows = False
        Me.dgvCorrespondance.AllowUserToDeleteRows = False
        Me.dgvCorrespondance.AllowUserToResizeRows = False
        DataGridViewCellStyle29.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvCorrespondance.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle29
        Me.dgvCorrespondance.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCorrespondance.Location = New System.Drawing.Point(3, 37)
        Me.dgvCorrespondance.Name = "dgvCorrespondance"
        Me.dgvCorrespondance.RowHeadersWidth = 20
        DataGridViewCellStyle30.Font = New System.Drawing.Font("Arial Narrow", 9.0!)
        Me.dgvCorrespondance.RowsDefaultCellStyle = DataGridViewCellStyle30
        Me.dgvCorrespondance.Size = New System.Drawing.Size(729, 324)
        Me.dgvCorrespondance.TabIndex = 1
        Me.dgvCorrespondance.TabStop = False
        '
        'fraCustomerInfo
        '
        Me.fraCustomerInfo.BackColor = System.Drawing.SystemColors.Control
        Me.fraCustomerInfo.Controls.Add(Me.txtClassID)
        Me.fraCustomerInfo.Controls.Add(Me.txtEQP)
        Me.fraCustomerInfo.Controls.Add(Me.txtCusNo)
        Me.fraCustomerInfo.Controls.Add(Me.txtCusName)
        Me.fraCustomerInfo.Controls.Add(Me.lblCusNo)
        Me.fraCustomerInfo.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraCustomerInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCustomerInfo.Location = New System.Drawing.Point(12, 12)
        Me.fraCustomerInfo.Name = "fraCustomerInfo"
        Me.fraCustomerInfo.Padding = New System.Windows.Forms.Padding(0)
        Me.fraCustomerInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCustomerInfo.Size = New System.Drawing.Size(422, 74)
        Me.fraCustomerInfo.TabIndex = 14
        Me.fraCustomerInfo.TabStop = False
        Me.fraCustomerInfo.Text = "Customer Info"
        '
        'txtClassID
        '
        Me.txtClassID.AcceptsReturn = True
        Me.txtClassID.BackColor = System.Drawing.SystemColors.Control
        Me.txtClassID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtClassID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtClassID.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtClassID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtClassID.Location = New System.Drawing.Point(11, 42)
        Me.txtClassID.MaxLength = 0
        Me.txtClassID.Multiline = True
        Me.txtClassID.Name = "txtClassID"
        Me.txtClassID.ReadOnly = True
        Me.txtClassID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtClassID.Size = New System.Drawing.Size(95, 21)
        Me.txtClassID.TabIndex = 6
        '
        'txtEQP
        '
        Me.txtEQP.AcceptsReturn = True
        Me.txtEQP.BackColor = System.Drawing.SystemColors.Control
        Me.txtEQP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEQP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEQP.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEQP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEQP.Location = New System.Drawing.Point(112, 42)
        Me.txtEQP.MaxLength = 0
        Me.txtEQP.Multiline = True
        Me.txtEQP.Name = "txtEQP"
        Me.txtEQP.ReadOnly = True
        Me.txtEQP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEQP.Size = New System.Drawing.Size(301, 21)
        Me.txtEQP.TabIndex = 5
        '
        'txtCusNo
        '
        Me.txtCusNo.AcceptsReturn = True
        Me.txtCusNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCusNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCusNo.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCusNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCusNo.Location = New System.Drawing.Point(112, 14)
        Me.txtCusNo.MaxLength = 0
        Me.txtCusNo.Name = "txtCusNo"
        Me.txtCusNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCusNo.Size = New System.Drawing.Size(65, 22)
        Me.txtCusNo.TabIndex = 1
        '
        'txtCusName
        '
        Me.txtCusName.AcceptsReturn = True
        Me.txtCusName.BackColor = System.Drawing.SystemColors.Control
        Me.txtCusName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCusName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCusName.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCusName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCusName.Location = New System.Drawing.Point(183, 14)
        Me.txtCusName.MaxLength = 0
        Me.txtCusName.Multiline = True
        Me.txtCusName.Name = "txtCusName"
        Me.txtCusName.ReadOnly = True
        Me.txtCusName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCusName.Size = New System.Drawing.Size(230, 22)
        Me.txtCusName.TabIndex = 4
        '
        'lblCusNo
        '
        Me.lblCusNo.BackColor = System.Drawing.SystemColors.Control
        Me.lblCusNo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCusNo.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCusNo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCusNo.Location = New System.Drawing.Point(8, 18)
        Me.lblCusNo.Name = "lblCusNo"
        Me.lblCusNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCusNo.Size = New System.Drawing.Size(81, 17)
        Me.lblCusNo.TabIndex = 0
        Me.lblCusNo.Text = "Customer No"
        '
        'Uc_OEFlashGeneral
        '
        Me.Uc_OEFlashGeneral.Location = New System.Drawing.Point(7, 9)
        Me.Uc_OEFlashGeneral.Name = "Uc_OEFlashGeneral"
        Me.Uc_OEFlashGeneral.Size = New System.Drawing.Size(78, 71)
        Me.Uc_OEFlashGeneral.TabIndex = 12
        '
        'uc_OEFlashCharges
        '
        Me.uc_OEFlashCharges.Cus_No = ""
        Me.uc_OEFlashCharges.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uc_OEFlashCharges.Location = New System.Drawing.Point(1, 7)
        Me.uc_OEFlashCharges.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.uc_OEFlashCharges.Name = "uc_OEFlashCharges"
        Me.uc_OEFlashCharges.Size = New System.Drawing.Size(742, 358)
        Me.uc_OEFlashCharges.TabIndex = 0
        '
        'Uc_OEFlashShipping
        '
        Me.Uc_OEFlashShipping.Location = New System.Drawing.Point(19, 15)
        Me.Uc_OEFlashShipping.Name = "Uc_OEFlashShipping"
        Me.Uc_OEFlashShipping.Size = New System.Drawing.Size(150, 150)
        Me.Uc_OEFlashShipping.TabIndex = 10
        '
        'Uc_OEFlashCorrespondance
        '
        Me.Uc_OEFlashCorrespondance.Location = New System.Drawing.Point(15, 16)
        Me.Uc_OEFlashCorrespondance.Name = "Uc_OEFlashCorrespondance"
        Me.Uc_OEFlashCorrespondance.Size = New System.Drawing.Size(150, 150)
        Me.Uc_OEFlashCorrespondance.TabIndex = 10
        '
        'frmCustomerConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(771, 502)
        Me.Controls.Add(Me.fraCustomerInfo)
        Me.Controls.Add(Me.tcOEFlash)
        Me.Name = "frmCustomerConfig"
        Me.Text = "frmCustomerConfig"
        Me.tcOEFlash.ResumeLayout(False)
        Me.OETab_General.ResumeLayout(False)
        Me.OETab_General.PerformLayout()
        CType(Me.dgvGeneral, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvOEFlash, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OETab_Charges.ResumeLayout(False)
        Me.OETab_Shipping.ResumeLayout(False)
        CType(Me.dgvShipping, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OETab_Correspondance.ResumeLayout(False)
        CType(Me.dgvCorrespondance, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraCustomerInfo.ResumeLayout(False)
        Me.fraCustomerInfo.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tcOEFlash As System.Windows.Forms.TabControl
    Friend WithEvents OETab_General As System.Windows.Forms.TabPage
    Public WithEvents Button7 As System.Windows.Forms.Button
    Public WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents dgvGeneral As System.Windows.Forms.DataGridView
    Friend WithEvents OETab_Charges As System.Windows.Forms.TabPage
    Friend WithEvents OETab_Pricing As System.Windows.Forms.TabPage
    Friend WithEvents OETab_Shipping As System.Windows.Forms.TabPage
    Public WithEvents Button3 As System.Windows.Forms.Button
    Public WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents dgvShipping As System.Windows.Forms.DataGridView
    Friend WithEvents OETab_Correspondance As System.Windows.Forms.TabPage
    Public WithEvents Button5 As System.Windows.Forms.Button
    Public WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents dgvCorrespondance As System.Windows.Forms.DataGridView
    Public WithEvents fraCustomerInfo As System.Windows.Forms.GroupBox
    Public WithEvents txtClassID As System.Windows.Forms.TextBox
    Public WithEvents txtEQP As System.Windows.Forms.TextBox
    Public WithEvents txtCusNo As System.Windows.Forms.TextBox
    Public WithEvents txtCusName As System.Windows.Forms.TextBox
    Public WithEvents lblCusNo As System.Windows.Forms.Label
    Friend WithEvents dgvOEFlash As System.Windows.Forms.DataGridView
    Friend WithEvents chkOEFlash As System.Windows.Forms.CheckBox
    Friend WithEvents uc_OEFlashCharges As Orders.ucOEFlashCharges
    Friend WithEvents Uc_OEFlashGeneral As Orders.ucOEFlashCustomer
    Friend WithEvents Uc_OEFlashShipping As Orders.ucOEFlashShipping
    Friend WithEvents Uc_OEFlashCorrespondance As Orders.ucOEFlashCorrespondance
End Class
