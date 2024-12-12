<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucOEFlash
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
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.fraCustomerInfo = New System.Windows.Forms.GroupBox
        Me.txtClassID = New System.Windows.Forms.TextBox
        Me.txtEQP = New System.Windows.Forms.TextBox
        Me.txtCusNo = New System.Windows.Forms.TextBox
        Me.txtCusName = New System.Windows.Forms.TextBox
        Me.lblCusNo = New System.Windows.Forms.Label
        Me.cmdFullScreen = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblVersion = New System.Windows.Forms.Label
        Me.tcOEFlash = New System.Windows.Forms.TabControl
        Me.OETab_General = New System.Windows.Forms.TabPage
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.dgvGeneral = New System.Windows.Forms.DataGridView
        Me.OETab_2 = New System.Windows.Forms.TabPage
        Me.cmdChargeDelete = New System.Windows.Forms.Button
        Me.cmdChargeAdd = New System.Windows.Forms.Button
        Me.dgvCharges = New System.Windows.Forms.DataGridView
        Me.OETab_3 = New System.Windows.Forms.TabPage
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.dgvPricing = New System.Windows.Forms.DataGridView
        Me.OETab_4 = New System.Windows.Forms.TabPage
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.dgvShipping = New System.Windows.Forms.DataGridView
        Me.OETab_5 = New System.Windows.Forms.TabPage
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.dgvCorrespondance = New System.Windows.Forms.DataGridView
        Me.fraCustomerInfo.SuspendLayout()
        Me.tcOEFlash.SuspendLayout()
        Me.OETab_General.SuspendLayout()
        CType(Me.dgvGeneral, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OETab_2.SuspendLayout()
        CType(Me.dgvCharges, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OETab_3.SuspendLayout()
        CType(Me.dgvPricing, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OETab_4.SuspendLayout()
        CType(Me.dgvShipping, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OETab_5.SuspendLayout()
        CType(Me.dgvCorrespondance, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.fraCustomerInfo.Location = New System.Drawing.Point(3, 3)
        Me.fraCustomerInfo.Name = "fraCustomerInfo"
        Me.fraCustomerInfo.Padding = New System.Windows.Forms.Padding(0)
        Me.fraCustomerInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCustomerInfo.Size = New System.Drawing.Size(422, 70)
        Me.fraCustomerInfo.TabIndex = 6
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
        'cmdFullScreen
        '
        Me.cmdFullScreen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdFullScreen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdFullScreen.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFullScreen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdFullScreen.Location = New System.Drawing.Point(635, 11)
        Me.cmdFullScreen.Name = "cmdFullScreen"
        Me.cmdFullScreen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdFullScreen.Size = New System.Drawing.Size(0, 0)
        Me.cmdFullScreen.TabIndex = 10
        Me.cmdFullScreen.Text = "Small Screen"
        Me.cmdFullScreen.UseVisualStyleBackColor = False
        '
        'lblVersion
        '
        Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblVersion.Location = New System.Drawing.Point(-5, 443)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(89, 17)
        Me.lblVersion.TabIndex = 8
        Me.lblVersion.Text = "lblVersion"
        '
        'tcOEFlash
        '
        Me.tcOEFlash.Controls.Add(Me.OETab_General)
        Me.tcOEFlash.Controls.Add(Me.OETab_2)
        Me.tcOEFlash.Controls.Add(Me.OETab_3)
        Me.tcOEFlash.Controls.Add(Me.OETab_4)
        Me.tcOEFlash.Controls.Add(Me.OETab_5)
        Me.tcOEFlash.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tcOEFlash.Location = New System.Drawing.Point(3, 79)
        Me.tcOEFlash.Name = "tcOEFlash"
        Me.tcOEFlash.SelectedIndex = 0
        Me.tcOEFlash.Size = New System.Drawing.Size(747, 397)
        Me.tcOEFlash.TabIndex = 11
        '
        'OETab_General
        '
        Me.OETab_General.Controls.Add(Me.Button7)
        Me.OETab_General.Controls.Add(Me.Button8)
        Me.OETab_General.Controls.Add(Me.dgvGeneral)
        Me.OETab_General.Location = New System.Drawing.Point(4, 25)
        Me.OETab_General.Name = "OETab_General"
        Me.OETab_General.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_General.Size = New System.Drawing.Size(739, 368)
        Me.OETab_General.TabIndex = 0
        Me.OETab_General.Text = "General"
        Me.OETab_General.UseVisualStyleBackColor = True
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
        'dgvGeneral
        '
        Me.dgvGeneral.AllowUserToAddRows = False
        Me.dgvGeneral.AllowUserToDeleteRows = False
        Me.dgvGeneral.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgvGeneral.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvGeneral.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGeneral.Location = New System.Drawing.Point(3, 37)
        Me.dgvGeneral.Name = "dgvGeneral"
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgvGeneral.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvGeneral.Size = New System.Drawing.Size(729, 324)
        Me.dgvGeneral.TabIndex = 1
        Me.dgvGeneral.TabStop = False
        '
        'OETab_2
        '
        Me.OETab_2.Controls.Add(Me.cmdChargeDelete)
        Me.OETab_2.Controls.Add(Me.cmdChargeAdd)
        Me.OETab_2.Controls.Add(Me.dgvCharges)
        Me.OETab_2.Location = New System.Drawing.Point(4, 25)
        Me.OETab_2.Name = "OETab_2"
        Me.OETab_2.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_2.Size = New System.Drawing.Size(739, 368)
        Me.OETab_2.TabIndex = 1
        Me.OETab_2.Text = "Charges"
        Me.OETab_2.UseVisualStyleBackColor = True
        '
        'cmdChargeDelete
        '
        Me.cmdChargeDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdChargeDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdChargeDelete.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdChargeDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdChargeDelete.Location = New System.Drawing.Point(672, 6)
        Me.cmdChargeDelete.Name = "cmdChargeDelete"
        Me.cmdChargeDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdChargeDelete.Size = New System.Drawing.Size(60, 25)
        Me.cmdChargeDelete.TabIndex = 7
        Me.cmdChargeDelete.Text = "Delete"
        Me.cmdChargeDelete.UseVisualStyleBackColor = False
        '
        'cmdChargeAdd
        '
        Me.cmdChargeAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdChargeAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdChargeAdd.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdChargeAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdChargeAdd.Location = New System.Drawing.Point(606, 6)
        Me.cmdChargeAdd.Name = "cmdChargeAdd"
        Me.cmdChargeAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdChargeAdd.Size = New System.Drawing.Size(60, 25)
        Me.cmdChargeAdd.TabIndex = 6
        Me.cmdChargeAdd.Text = "Add"
        Me.cmdChargeAdd.UseVisualStyleBackColor = False
        '
        'dgvCharges
        '
        Me.dgvCharges.AllowUserToAddRows = False
        Me.dgvCharges.AllowUserToDeleteRows = False
        Me.dgvCharges.AllowUserToResizeRows = False
        Me.dgvCharges.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCharges.Location = New System.Drawing.Point(3, 37)
        Me.dgvCharges.Name = "dgvCharges"
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgvCharges.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvCharges.Size = New System.Drawing.Size(729, 324)
        Me.dgvCharges.TabIndex = 1
        Me.dgvCharges.TabStop = False
        '
        'OETab_3
        '
        Me.OETab_3.Controls.Add(Me.Button1)
        Me.OETab_3.Controls.Add(Me.Button2)
        Me.OETab_3.Controls.Add(Me.dgvPricing)
        Me.OETab_3.Location = New System.Drawing.Point(4, 25)
        Me.OETab_3.Name = "OETab_3"
        Me.OETab_3.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_3.Size = New System.Drawing.Size(739, 368)
        Me.OETab_3.TabIndex = 2
        Me.OETab_3.Text = "Pricing"
        Me.OETab_3.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.Control
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button1.Location = New System.Drawing.Point(672, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button1.Size = New System.Drawing.Size(60, 25)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Delete"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.SystemColors.Control
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button2.Location = New System.Drawing.Point(606, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button2.Size = New System.Drawing.Size(60, 25)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "Add"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'dgvPricing
        '
        Me.dgvPricing.AllowUserToAddRows = False
        Me.dgvPricing.AllowUserToDeleteRows = False
        Me.dgvPricing.AllowUserToResizeRows = False
        Me.dgvPricing.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPricing.Location = New System.Drawing.Point(3, 37)
        Me.dgvPricing.Name = "dgvPricing"
        Me.dgvPricing.Size = New System.Drawing.Size(729, 324)
        Me.dgvPricing.TabIndex = 1
        Me.dgvPricing.TabStop = False
        '
        'OETab_4
        '
        Me.OETab_4.Controls.Add(Me.Button3)
        Me.OETab_4.Controls.Add(Me.Button4)
        Me.OETab_4.Controls.Add(Me.dgvShipping)
        Me.OETab_4.Location = New System.Drawing.Point(4, 25)
        Me.OETab_4.Name = "OETab_4"
        Me.OETab_4.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_4.Size = New System.Drawing.Size(739, 368)
        Me.OETab_4.TabIndex = 3
        Me.OETab_4.Text = "Shipping"
        Me.OETab_4.UseVisualStyleBackColor = True
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
        Me.dgvShipping.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvShipping.Location = New System.Drawing.Point(3, 37)
        Me.dgvShipping.Name = "dgvShipping"
        Me.dgvShipping.Size = New System.Drawing.Size(729, 324)
        Me.dgvShipping.TabIndex = 1
        Me.dgvShipping.TabStop = False
        '
        'OETab_5
        '
        Me.OETab_5.Controls.Add(Me.Button5)
        Me.OETab_5.Controls.Add(Me.Button6)
        Me.OETab_5.Controls.Add(Me.dgvCorrespondance)
        Me.OETab_5.Location = New System.Drawing.Point(4, 25)
        Me.OETab_5.Name = "OETab_5"
        Me.OETab_5.Padding = New System.Windows.Forms.Padding(3)
        Me.OETab_5.Size = New System.Drawing.Size(739, 368)
        Me.OETab_5.TabIndex = 4
        Me.OETab_5.Text = "Correspondence"
        Me.OETab_5.UseVisualStyleBackColor = True
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
        Me.dgvCorrespondance.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCorrespondance.Location = New System.Drawing.Point(3, 37)
        Me.dgvCorrespondance.Name = "dgvCorrespondance"
        Me.dgvCorrespondance.Size = New System.Drawing.Size(729, 324)
        Me.dgvCorrespondance.TabIndex = 1
        Me.dgvCorrespondance.TabStop = False
        '
        'ucOEFlash
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.Controls.Add(Me.tcOEFlash)
        Me.Controls.Add(Me.fraCustomerInfo)
        Me.Controls.Add(Me.cmdFullScreen)
        Me.Controls.Add(Me.lblVersion)
        Me.Name = "ucOEFlash"
        Me.Size = New System.Drawing.Size(753, 482)
        Me.fraCustomerInfo.ResumeLayout(False)
        Me.fraCustomerInfo.PerformLayout()
        Me.tcOEFlash.ResumeLayout(False)
        Me.OETab_General.ResumeLayout(False)
        CType(Me.dgvGeneral, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OETab_2.ResumeLayout(False)
        CType(Me.dgvCharges, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OETab_3.ResumeLayout(False)
        CType(Me.dgvPricing, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OETab_4.ResumeLayout(False)
        CType(Me.dgvShipping, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OETab_5.ResumeLayout(False)
        CType(Me.dgvCorrespondance, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents fraCustomerInfo As System.Windows.Forms.GroupBox
    Public WithEvents txtClassID As System.Windows.Forms.TextBox
    Public WithEvents txtEQP As System.Windows.Forms.TextBox
    Public WithEvents txtCusNo As System.Windows.Forms.TextBox
    Public WithEvents txtCusName As System.Windows.Forms.TextBox
    Public WithEvents lblCusNo As System.Windows.Forms.Label
    Public WithEvents cmdFullScreen As System.Windows.Forms.Button
    Public WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents tcOEFlash As System.Windows.Forms.TabControl
    Friend WithEvents OETab_General As System.Windows.Forms.TabPage
    Friend WithEvents OETab_2 As System.Windows.Forms.TabPage
    Friend WithEvents OETab_3 As System.Windows.Forms.TabPage
    Friend WithEvents OETab_4 As System.Windows.Forms.TabPage
    Friend WithEvents OETab_5 As System.Windows.Forms.TabPage
    Friend WithEvents dgvGeneral As System.Windows.Forms.DataGridView
    Friend WithEvents dgvCharges As System.Windows.Forms.DataGridView
    Friend WithEvents dgvPricing As System.Windows.Forms.DataGridView
    Friend WithEvents dgvShipping As System.Windows.Forms.DataGridView
    Public WithEvents cmdChargeDelete As System.Windows.Forms.Button
    Public WithEvents cmdChargeAdd As System.Windows.Forms.Button
    Public WithEvents Button1 As System.Windows.Forms.Button
    Public WithEvents Button2 As System.Windows.Forms.Button
    Public WithEvents Button3 As System.Windows.Forms.Button
    Public WithEvents Button4 As System.Windows.Forms.Button
    Public WithEvents Button5 As System.Windows.Forms.Button
    Public WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents dgvCorrespondance As System.Windows.Forms.DataGridView
    Public WithEvents Button7 As System.Windows.Forms.Button
    Public WithEvents Button8 As System.Windows.Forms.Button

End Class
