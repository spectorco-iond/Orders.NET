<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucContacts
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
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdCopyContacts = New System.Windows.Forms.Button()
        Me.cmdRemoveContact = New System.Windows.Forms.Button()
        Me.fraCustomer = New System.Windows.Forms.GroupBox()
        Me.txtCusNo = New System.Windows.Forms.TextBox()
        Me.txtCusName = New System.Windows.Forms.TextBox()
        Me.dgvContacts = New System.Windows.Forms.DataGridView()
        Me.optType = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me._optType_5 = New System.Windows.Forms.RadioButton()
        Me._optType_0 = New System.Windows.Forms.RadioButton()
        Me._optType_2 = New System.Windows.Forms.RadioButton()
        Me._optType_3 = New System.Windows.Forms.RadioButton()
        Me._optType_4 = New System.Windows.Forms.RadioButton()
        Me._optType_1 = New System.Windows.Forms.RadioButton()
        Me.cmdInsertNewContact = New System.Windows.Forms.Button()
        Me.fraContacts = New System.Windows.Forms.GroupBox()
        Me.fraMethod = New System.Windows.Forms.GroupBox()
        Me.optAllMethod = New System.Windows.Forms.RadioButton()
        Me.optFax = New System.Windows.Forms.RadioButton()
        Me.optEmail = New System.Windows.Forms.RadioButton()
        Me.fraType = New System.Windows.Forms.GroupBox()
        Me.optAllType = New System.Windows.Forms.RadioButton()
        Me.optPrimary = New System.Windows.Forms.RadioButton()
        Me.optSecondary = New System.Windows.Forms.RadioButton()
        Me.fraContactTypes = New System.Windows.Forms.GroupBox()
        Me.txt90DayRemind = New Orders.xTextBox()
        Me.txtPreShipping = New Orders.xTextBox()
        Me.txtShipping = New Orders.xTextBox()
        Me.txtOrderAck = New Orders.xTextBox()
        Me.txtPaperproof = New Orders.xTextBox()
        Me.txtOEContact = New Orders.xTextBox()
        Me.grpAddContact = New System.Windows.Forms.GroupBox()
        Me.cmdClear = New System.Windows.Forms.Button()
        Me.txtID = New Orders.xTextBox()
        Me.lblID = New System.Windows.Forms.Label()
        Me.txtExt = New Orders.xTextBox()
        Me.txtTel = New Orders.xTextBox()
        Me.txtEmail = New Orders.xTextBox()
        Me.txtFax = New Orders.xTextBox()
        Me.txtLastName = New Orders.xTextBox()
        Me.txtFirstName = New Orders.xTextBox()
        Me.lblTel = New System.Windows.Forms.Label()
        Me.lblFax = New System.Windows.Forms.Label()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.cboPred = New System.Windows.Forms.ComboBox()
        Me.cboLang = New System.Windows.Forms.ComboBox()
        Me.lblPred = New System.Windows.Forms.Label()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.lblLang = New System.Windows.Forms.Label()
        Me.fraCustomer.SuspendLayout()
        CType(Me.dgvContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraContacts.SuspendLayout()
        Me.fraMethod.SuspendLayout()
        Me.fraType.SuspendLayout()
        Me.fraContactTypes.SuspendLayout()
        Me.grpAddContact.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCopyContacts
        '
        Me.cmdCopyContacts.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopyContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCopyContacts.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdCopyContacts.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopyContacts.Location = New System.Drawing.Point(8, 190)
        Me.cmdCopyContacts.Name = "cmdCopyContacts"
        Me.cmdCopyContacts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCopyContacts.Size = New System.Drawing.Size(130, 24)
        Me.cmdCopyContacts.TabIndex = 12
        Me.cmdCopyContacts.Text = "Copy to all Types"
        Me.cmdCopyContacts.UseVisualStyleBackColor = False
        '
        'cmdRemoveContact
        '
        Me.cmdRemoveContact.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRemoveContact.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRemoveContact.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdRemoveContact.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRemoveContact.Location = New System.Drawing.Point(148, 190)
        Me.cmdRemoveContact.Name = "cmdRemoveContact"
        Me.cmdRemoveContact.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRemoveContact.Size = New System.Drawing.Size(130, 24)
        Me.cmdRemoveContact.TabIndex = 13
        Me.cmdRemoveContact.Text = "Remove Contact"
        Me.cmdRemoveContact.UseVisualStyleBackColor = False
        '
        'fraCustomer
        '
        Me.fraCustomer.BackColor = System.Drawing.SystemColors.Control
        Me.fraCustomer.Controls.Add(Me.txtCusNo)
        Me.fraCustomer.Controls.Add(Me.txtCusName)
        Me.fraCustomer.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.fraCustomer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCustomer.Location = New System.Drawing.Point(3, 3)
        Me.fraCustomer.Name = "fraCustomer"
        Me.fraCustomer.Padding = New System.Windows.Forms.Padding(0)
        Me.fraCustomer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCustomer.Size = New System.Drawing.Size(291, 88)
        Me.fraCustomer.TabIndex = 0
        Me.fraCustomer.TabStop = False
        Me.fraCustomer.Text = "Customer Information"
        '
        'txtCusNo
        '
        Me.txtCusNo.AcceptsReturn = True
        Me.txtCusNo.BackColor = System.Drawing.SystemColors.Control
        Me.txtCusNo.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtCusNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCusNo.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.txtCusNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCusNo.Location = New System.Drawing.Point(8, 24)
        Me.txtCusNo.MaxLength = 20
        Me.txtCusNo.Name = "txtCusNo"
        Me.txtCusNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCusNo.Size = New System.Drawing.Size(270, 15)
        Me.txtCusNo.TabIndex = 0
        Me.txtCusNo.TabStop = False
        '
        'txtCusName
        '
        Me.txtCusName.AcceptsReturn = True
        Me.txtCusName.BackColor = System.Drawing.SystemColors.Control
        Me.txtCusName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtCusName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCusName.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.txtCusName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCusName.Location = New System.Drawing.Point(8, 45)
        Me.txtCusName.MaxLength = 20
        Me.txtCusName.Multiline = True
        Me.txtCusName.Name = "txtCusName"
        Me.txtCusName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCusName.Size = New System.Drawing.Size(270, 32)
        Me.txtCusName.TabIndex = 1
        Me.txtCusName.TabStop = False
        '
        'dgvContacts
        '
        Me.dgvContacts.AllowUserToAddRows = False
        Me.dgvContacts.AllowUserToDeleteRows = False
        Me.dgvContacts.AllowUserToResizeRows = False
        Me.dgvContacts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvContacts.Location = New System.Drawing.Point(8, 18)
        Me.dgvContacts.Name = "dgvContacts"
        Me.dgvContacts.Size = New System.Drawing.Size(515, 244)
        Me.dgvContacts.TabIndex = 0
        Me.dgvContacts.TabStop = False
        '
        '_optType_5
        '
        Me._optType_5.BackColor = System.Drawing.SystemColors.Control
        Me._optType_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._optType_5.Font = New System.Drawing.Font("Arial", 9.75!)
        Me._optType_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optType.SetIndex(Me._optType_5, CType(5, Short))
        Me._optType_5.Location = New System.Drawing.Point(8, 157)
        Me._optType_5.Name = "_optType_5"
        Me._optType_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optType_5.Size = New System.Drawing.Size(129, 21)
        Me._optType_5.TabIndex = 10
        Me._optType_5.Text = "90-Day Reminder"
        Me._optType_5.UseVisualStyleBackColor = False
        '
        '_optType_0
        '
        Me._optType_0.BackColor = System.Drawing.SystemColors.Control
        Me._optType_0.Checked = True
        Me._optType_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optType_0.Font = New System.Drawing.Font("Arial", 9.75!)
        Me._optType_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optType.SetIndex(Me._optType_0, CType(0, Short))
        Me._optType_0.Location = New System.Drawing.Point(8, 29)
        Me._optType_0.Name = "_optType_0"
        Me._optType_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optType_0.Size = New System.Drawing.Size(129, 17)
        Me._optType_0.TabIndex = 0
        Me._optType_0.TabStop = True
        Me._optType_0.Text = "OE Contact (Main)"
        Me._optType_0.UseVisualStyleBackColor = False
        '
        '_optType_2
        '
        Me._optType_2.BackColor = System.Drawing.SystemColors.Control
        Me._optType_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optType_2.Font = New System.Drawing.Font("Arial", 9.75!)
        Me._optType_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optType.SetIndex(Me._optType_2, CType(2, Short))
        Me._optType_2.Location = New System.Drawing.Point(8, 133)
        Me._optType_2.Name = "_optType_2"
        Me._optType_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optType_2.Size = New System.Drawing.Size(130, 20)
        Me._optType_2.TabIndex = 8
        Me._optType_2.Text = "Pre-Shipping Adv."
        Me._optType_2.UseVisualStyleBackColor = False
        '
        '_optType_3
        '
        Me._optType_3.BackColor = System.Drawing.SystemColors.Control
        Me._optType_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._optType_3.Font = New System.Drawing.Font("Arial", 9.75!)
        Me._optType_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optType.SetIndex(Me._optType_3, CType(3, Short))
        Me._optType_3.Location = New System.Drawing.Point(8, 109)
        Me._optType_3.Name = "_optType_3"
        Me._optType_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optType_3.Size = New System.Drawing.Size(129, 20)
        Me._optType_3.TabIndex = 6
        Me._optType_3.Text = "Shipping Advice"
        Me._optType_3.UseVisualStyleBackColor = False
        '
        '_optType_4
        '
        Me._optType_4.BackColor = System.Drawing.SystemColors.Control
        Me._optType_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._optType_4.Font = New System.Drawing.Font("Arial", 9.75!)
        Me._optType_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optType.SetIndex(Me._optType_4, CType(4, Short))
        Me._optType_4.Location = New System.Drawing.Point(8, 85)
        Me._optType_4.Name = "_optType_4"
        Me._optType_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optType_4.Size = New System.Drawing.Size(130, 20)
        Me._optType_4.TabIndex = 4
        Me._optType_4.Text = "Order Ack"
        Me._optType_4.UseVisualStyleBackColor = False
        '
        '_optType_1
        '
        Me._optType_1.BackColor = System.Drawing.SystemColors.Control
        Me._optType_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optType_1.Font = New System.Drawing.Font("Arial", 9.75!)
        Me._optType_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optType.SetIndex(Me._optType_1, CType(1, Short))
        Me._optType_1.Location = New System.Drawing.Point(8, 61)
        Me._optType_1.Name = "_optType_1"
        Me._optType_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optType_1.Size = New System.Drawing.Size(129, 20)
        Me._optType_1.TabIndex = 2
        Me._optType_1.Text = "Paperproof"
        Me._optType_1.UseVisualStyleBackColor = False
        '
        'cmdInsertNewContact
        '
        Me.cmdInsertNewContact.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInsertNewContact.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdInsertNewContact.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdInsertNewContact.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInsertNewContact.Location = New System.Drawing.Point(599, 46)
        Me.cmdInsertNewContact.Name = "cmdInsertNewContact"
        Me.cmdInsertNewContact.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdInsertNewContact.Size = New System.Drawing.Size(107, 24)
        Me.cmdInsertNewContact.TabIndex = 15
        Me.cmdInsertNewContact.Text = "Save"
        Me.cmdInsertNewContact.UseVisualStyleBackColor = False
        '
        'fraContacts
        '
        Me.fraContacts.BackColor = System.Drawing.SystemColors.Control
        Me.fraContacts.Controls.Add(Me.fraMethod)
        Me.fraContacts.Controls.Add(Me.fraType)
        Me.fraContacts.Controls.Add(Me.dgvContacts)
        Me.fraContacts.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.fraContacts.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraContacts.Location = New System.Drawing.Point(300, 3)
        Me.fraContacts.Name = "fraContacts"
        Me.fraContacts.Padding = New System.Windows.Forms.Padding(0)
        Me.fraContacts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraContacts.Size = New System.Drawing.Size(533, 323)
        Me.fraContacts.TabIndex = 2
        Me.fraContacts.TabStop = False
        Me.fraContacts.Text = "Contact list"
        '
        'fraMethod
        '
        Me.fraMethod.BackColor = System.Drawing.SystemColors.Control
        Me.fraMethod.Controls.Add(Me.optAllMethod)
        Me.fraMethod.Controls.Add(Me.optFax)
        Me.fraMethod.Controls.Add(Me.optEmail)
        Me.fraMethod.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.fraMethod.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMethod.Location = New System.Drawing.Point(220, 268)
        Me.fraMethod.Name = "fraMethod"
        Me.fraMethod.Padding = New System.Windows.Forms.Padding(0)
        Me.fraMethod.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMethod.Size = New System.Drawing.Size(206, 46)
        Me.fraMethod.TabIndex = 2
        Me.fraMethod.TabStop = False
        Me.fraMethod.Text = "Contact Method"
        '
        'optAllMethod
        '
        Me.optAllMethod.BackColor = System.Drawing.SystemColors.Control
        Me.optAllMethod.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAllMethod.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.optAllMethod.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAllMethod.Location = New System.Drawing.Point(148, 17)
        Me.optAllMethod.Name = "optAllMethod"
        Me.optAllMethod.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAllMethod.Size = New System.Drawing.Size(54, 23)
        Me.optAllMethod.TabIndex = 2
        Me.optAllMethod.TabStop = True
        Me.optAllMethod.Text = "All"
        Me.optAllMethod.UseVisualStyleBackColor = False
        Me.optAllMethod.Visible = False
        '
        'optFax
        '
        Me.optFax.BackColor = System.Drawing.SystemColors.Control
        Me.optFax.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFax.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.optFax.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFax.Location = New System.Drawing.Point(8, 18)
        Me.optFax.Name = "optFax"
        Me.optFax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFax.Size = New System.Drawing.Size(54, 23)
        Me.optFax.TabIndex = 0
        Me.optFax.TabStop = True
        Me.optFax.Text = " Fax"
        Me.optFax.UseVisualStyleBackColor = False
        '
        'optEmail
        '
        Me.optEmail.BackColor = System.Drawing.SystemColors.Control
        Me.optEmail.Cursor = System.Windows.Forms.Cursors.Default
        Me.optEmail.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.optEmail.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optEmail.Location = New System.Drawing.Point(68, 18)
        Me.optEmail.Name = "optEmail"
        Me.optEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optEmail.Size = New System.Drawing.Size(75, 23)
        Me.optEmail.TabIndex = 1
        Me.optEmail.TabStop = True
        Me.optEmail.Text = " Email"
        Me.optEmail.UseVisualStyleBackColor = False
        '
        'fraType
        '
        Me.fraType.BackColor = System.Drawing.SystemColors.Control
        Me.fraType.Controls.Add(Me.optAllType)
        Me.fraType.Controls.Add(Me.optPrimary)
        Me.fraType.Controls.Add(Me.optSecondary)
        Me.fraType.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.fraType.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraType.Location = New System.Drawing.Point(8, 268)
        Me.fraType.Name = "fraType"
        Me.fraType.Padding = New System.Windows.Forms.Padding(0)
        Me.fraType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraType.Size = New System.Drawing.Size(206, 46)
        Me.fraType.TabIndex = 1
        Me.fraType.TabStop = False
        Me.fraType.Text = "Contact Type"
        '
        'optAllType
        '
        Me.optAllType.BackColor = System.Drawing.SystemColors.Control
        Me.optAllType.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAllType.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.optAllType.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAllType.Location = New System.Drawing.Point(149, 18)
        Me.optAllType.Name = "optAllType"
        Me.optAllType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAllType.Size = New System.Drawing.Size(54, 23)
        Me.optAllType.TabIndex = 2
        Me.optAllType.Text = "All"
        Me.optAllType.UseVisualStyleBackColor = False
        Me.optAllType.Visible = False
        '
        'optPrimary
        '
        Me.optPrimary.BackColor = System.Drawing.SystemColors.Control
        Me.optPrimary.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPrimary.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.optPrimary.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPrimary.Location = New System.Drawing.Point(8, 18)
        Me.optPrimary.Name = "optPrimary"
        Me.optPrimary.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPrimary.Size = New System.Drawing.Size(73, 23)
        Me.optPrimary.TabIndex = 0
        Me.optPrimary.Text = "Primary"
        Me.optPrimary.UseVisualStyleBackColor = False
        '
        'optSecondary
        '
        Me.optSecondary.BackColor = System.Drawing.SystemColors.Control
        Me.optSecondary.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSecondary.Enabled = False
        Me.optSecondary.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.optSecondary.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSecondary.Location = New System.Drawing.Point(87, 18)
        Me.optSecondary.Name = "optSecondary"
        Me.optSecondary.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSecondary.Size = New System.Drawing.Size(55, 23)
        Me.optSecondary.TabIndex = 1
        Me.optSecondary.Text = "Sec."
        Me.optSecondary.UseVisualStyleBackColor = False
        Me.optSecondary.Visible = False
        '
        'fraContactTypes
        '
        Me.fraContactTypes.BackColor = System.Drawing.SystemColors.Control
        Me.fraContactTypes.Controls.Add(Me.txt90DayRemind)
        Me.fraContactTypes.Controls.Add(Me.txtPreShipping)
        Me.fraContactTypes.Controls.Add(Me.txtShipping)
        Me.fraContactTypes.Controls.Add(Me.txtOrderAck)
        Me.fraContactTypes.Controls.Add(Me.txtPaperproof)
        Me.fraContactTypes.Controls.Add(Me.txtOEContact)
        Me.fraContactTypes.Controls.Add(Me.cmdCopyContacts)
        Me.fraContactTypes.Controls.Add(Me.cmdRemoveContact)
        Me.fraContactTypes.Controls.Add(Me._optType_0)
        Me.fraContactTypes.Controls.Add(Me._optType_5)
        Me.fraContactTypes.Controls.Add(Me._optType_2)
        Me.fraContactTypes.Controls.Add(Me._optType_3)
        Me.fraContactTypes.Controls.Add(Me._optType_4)
        Me.fraContactTypes.Controls.Add(Me._optType_1)
        Me.fraContactTypes.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.fraContactTypes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraContactTypes.Location = New System.Drawing.Point(3, 97)
        Me.fraContactTypes.Name = "fraContactTypes"
        Me.fraContactTypes.Padding = New System.Windows.Forms.Padding(0)
        Me.fraContactTypes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraContactTypes.Size = New System.Drawing.Size(291, 229)
        Me.fraContactTypes.TabIndex = 1
        Me.fraContactTypes.TabStop = False
        Me.fraContactTypes.Text = "Contact Types"
        '
        'txt90DayRemind
        '
        Me.txt90DayRemind.BackColor = System.Drawing.SystemColors.Window
        Me.txt90DayRemind.CharacterInput = Orders.Cinput.NumericOnly
        Me.txt90DayRemind.DataLength = CType(0, Long)
        Me.txt90DayRemind.DataType = Orders.CDataType.dtString
        Me.txt90DayRemind.DateValue = New Date(CType(0, Long))
        Me.txt90DayRemind.DecimalDigits = CType(0, Long)
        Me.txt90DayRemind.Location = New System.Drawing.Point(143, 155)
        Me.txt90DayRemind.Mask = Nothing
        Me.txt90DayRemind.Name = "txt90DayRemind"
        Me.txt90DayRemind.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txt90DayRemind.OldValue = ""
        Me.txt90DayRemind.ReadOnly = True
        Me.txt90DayRemind.Size = New System.Drawing.Size(140, 22)
        Me.txt90DayRemind.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txt90DayRemind.StringValue = Nothing
        Me.txt90DayRemind.TabIndex = 11
        Me.txt90DayRemind.TextDataID = Nothing
        '
        'txtPreShipping
        '
        Me.txtPreShipping.BackColor = System.Drawing.SystemColors.Window
        Me.txtPreShipping.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtPreShipping.DataLength = CType(0, Long)
        Me.txtPreShipping.DataType = Orders.CDataType.dtString
        Me.txtPreShipping.DateValue = New Date(CType(0, Long))
        Me.txtPreShipping.DecimalDigits = CType(0, Long)
        Me.txtPreShipping.Location = New System.Drawing.Point(143, 131)
        Me.txtPreShipping.Mask = Nothing
        Me.txtPreShipping.Name = "txtPreShipping"
        Me.txtPreShipping.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtPreShipping.OldValue = ""
        Me.txtPreShipping.ReadOnly = True
        Me.txtPreShipping.Size = New System.Drawing.Size(140, 22)
        Me.txtPreShipping.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtPreShipping.StringValue = Nothing
        Me.txtPreShipping.TabIndex = 9
        Me.txtPreShipping.TextDataID = Nothing
        '
        'txtShipping
        '
        Me.txtShipping.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipping.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShipping.DataLength = CType(0, Long)
        Me.txtShipping.DataType = Orders.CDataType.dtString
        Me.txtShipping.DateValue = New Date(CType(0, Long))
        Me.txtShipping.DecimalDigits = CType(0, Long)
        Me.txtShipping.Location = New System.Drawing.Point(143, 107)
        Me.txtShipping.Mask = Nothing
        Me.txtShipping.Name = "txtShipping"
        Me.txtShipping.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShipping.OldValue = ""
        Me.txtShipping.ReadOnly = True
        Me.txtShipping.Size = New System.Drawing.Size(140, 22)
        Me.txtShipping.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShipping.StringValue = Nothing
        Me.txtShipping.TabIndex = 7
        Me.txtShipping.TextDataID = Nothing
        '
        'txtOrderAck
        '
        Me.txtOrderAck.BackColor = System.Drawing.SystemColors.Window
        Me.txtOrderAck.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtOrderAck.DataLength = CType(0, Long)
        Me.txtOrderAck.DataType = Orders.CDataType.dtString
        Me.txtOrderAck.DateValue = New Date(CType(0, Long))
        Me.txtOrderAck.DecimalDigits = CType(0, Long)
        Me.txtOrderAck.Location = New System.Drawing.Point(143, 83)
        Me.txtOrderAck.Mask = Nothing
        Me.txtOrderAck.Name = "txtOrderAck"
        Me.txtOrderAck.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtOrderAck.OldValue = ""
        Me.txtOrderAck.ReadOnly = True
        Me.txtOrderAck.Size = New System.Drawing.Size(140, 22)
        Me.txtOrderAck.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtOrderAck.StringValue = Nothing
        Me.txtOrderAck.TabIndex = 5
        Me.txtOrderAck.TextDataID = Nothing
        '
        'txtPaperproof
        '
        Me.txtPaperproof.BackColor = System.Drawing.SystemColors.Window
        Me.txtPaperproof.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtPaperproof.DataLength = CType(0, Long)
        Me.txtPaperproof.DataType = Orders.CDataType.dtString
        Me.txtPaperproof.DateValue = New Date(CType(0, Long))
        Me.txtPaperproof.DecimalDigits = CType(0, Long)
        Me.txtPaperproof.Location = New System.Drawing.Point(143, 59)
        Me.txtPaperproof.Mask = Nothing
        Me.txtPaperproof.Name = "txtPaperproof"
        Me.txtPaperproof.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtPaperproof.OldValue = ""
        Me.txtPaperproof.ReadOnly = True
        Me.txtPaperproof.Size = New System.Drawing.Size(140, 22)
        Me.txtPaperproof.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtPaperproof.StringValue = Nothing
        Me.txtPaperproof.TabIndex = 3
        Me.txtPaperproof.TextDataID = Nothing
        '
        'txtOEContact
        '
        Me.txtOEContact.BackColor = System.Drawing.SystemColors.Window
        Me.txtOEContact.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtOEContact.DataLength = CType(0, Long)
        Me.txtOEContact.DataType = Orders.CDataType.dtString
        Me.txtOEContact.DateValue = New Date(CType(0, Long))
        Me.txtOEContact.DecimalDigits = CType(0, Long)
        Me.txtOEContact.Location = New System.Drawing.Point(143, 27)
        Me.txtOEContact.Mask = Nothing
        Me.txtOEContact.Name = "txtOEContact"
        Me.txtOEContact.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtOEContact.OldValue = ""
        Me.txtOEContact.ReadOnly = True
        Me.txtOEContact.Size = New System.Drawing.Size(140, 22)
        Me.txtOEContact.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtOEContact.StringValue = Nothing
        Me.txtOEContact.TabIndex = 1
        Me.txtOEContact.TextDataID = Nothing
        '
        'grpAddContact
        '
        Me.grpAddContact.BackColor = System.Drawing.SystemColors.Control
        Me.grpAddContact.Controls.Add(Me.cmdClear)
        Me.grpAddContact.Controls.Add(Me.txtID)
        Me.grpAddContact.Controls.Add(Me.lblID)
        Me.grpAddContact.Controls.Add(Me.txtExt)
        Me.grpAddContact.Controls.Add(Me.txtTel)
        Me.grpAddContact.Controls.Add(Me.txtEmail)
        Me.grpAddContact.Controls.Add(Me.txtFax)
        Me.grpAddContact.Controls.Add(Me.txtLastName)
        Me.grpAddContact.Controls.Add(Me.txtFirstName)
        Me.grpAddContact.Controls.Add(Me.lblTel)
        Me.grpAddContact.Controls.Add(Me.lblFax)
        Me.grpAddContact.Controls.Add(Me.lblEmail)
        Me.grpAddContact.Controls.Add(Me.cboPred)
        Me.grpAddContact.Controls.Add(Me.cboLang)
        Me.grpAddContact.Controls.Add(Me.lblPred)
        Me.grpAddContact.Controls.Add(Me.lblFirstName)
        Me.grpAddContact.Controls.Add(Me.lblLastName)
        Me.grpAddContact.Controls.Add(Me.lblLang)
        Me.grpAddContact.Controls.Add(Me.cmdInsertNewContact)
        Me.grpAddContact.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.grpAddContact.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpAddContact.Location = New System.Drawing.Point(3, 332)
        Me.grpAddContact.Name = "grpAddContact"
        Me.grpAddContact.Padding = New System.Windows.Forms.Padding(0)
        Me.grpAddContact.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.grpAddContact.Size = New System.Drawing.Size(723, 136)
        Me.grpAddContact.TabIndex = 3
        Me.grpAddContact.TabStop = False
        Me.grpAddContact.Text = "Adding a contact"
        '
        'cmdClear
        '
        Me.cmdClear.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClear.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClear.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdClear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClear.Location = New System.Drawing.Point(599, 19)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClear.Size = New System.Drawing.Size(107, 26)
        Me.cmdClear.TabIndex = 18
        Me.cmdClear.Text = "Clear"
        Me.cmdClear.UseVisualStyleBackColor = False
        '
        'txtID
        '
        Me.txtID.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtID.DataLength = CType(128, Long)
        Me.txtID.DataType = Orders.CDataType.dtString
        Me.txtID.DateValue = New Date(CType(0, Long))
        Me.txtID.DecimalDigits = CType(0, Long)
        Me.txtID.Location = New System.Drawing.Point(429, 99)
        Me.txtID.Mask = Nothing
        Me.txtID.Name = "txtID"
        Me.txtID.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtID.OldValue = ""
        Me.txtID.Size = New System.Drawing.Size(277, 22)
        Me.txtID.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtID.StringValue = Nothing
        Me.txtID.TabIndex = 17
        Me.txtID.TextDataID = Nothing
        Me.txtID.Visible = False
        '
        'lblID
        '
        Me.lblID.BackColor = System.Drawing.SystemColors.Control
        Me.lblID.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblID.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblID.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblID.Location = New System.Drawing.Point(309, 102)
        Me.lblID.Name = "lblID"
        Me.lblID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblID.Size = New System.Drawing.Size(105, 17)
        Me.lblID.TabIndex = 16
        Me.lblID.Text = "ID"
        Me.lblID.Visible = False
        '
        'txtExt
        '
        Me.txtExt.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtExt.DataLength = CType(4, Long)
        Me.txtExt.DataType = Orders.CDataType.dtString
        Me.txtExt.DateValue = New Date(CType(0, Long))
        Me.txtExt.DecimalDigits = CType(0, Long)
        Me.txtExt.Location = New System.Drawing.Point(537, 21)
        Me.txtExt.Mask = Nothing
        Me.txtExt.Name = "txtExt"
        Me.txtExt.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtExt.OldValue = ""
        Me.txtExt.Size = New System.Drawing.Size(42, 22)
        Me.txtExt.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtExt.StringValue = Nothing
        Me.txtExt.TabIndex = 10
        Me.txtExt.TextDataID = Nothing
        '
        'txtTel
        '
        Me.txtTel.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtTel.DataLength = CType(25, Long)
        Me.txtTel.DataType = Orders.CDataType.dtString
        Me.txtTel.DateValue = New Date(CType(0, Long))
        Me.txtTel.DecimalDigits = CType(0, Long)
        Me.txtTel.Location = New System.Drawing.Point(429, 21)
        Me.txtTel.Mask = Nothing
        Me.txtTel.Name = "txtTel"
        Me.txtTel.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtTel.OldValue = ""
        Me.txtTel.Size = New System.Drawing.Size(105, 22)
        Me.txtTel.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtTel.StringValue = Nothing
        Me.txtTel.TabIndex = 9
        Me.txtTel.TextDataID = Nothing
        '
        'txtEmail
        '
        Me.txtEmail.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtEmail.DataLength = CType(128, Long)
        Me.txtEmail.DataType = Orders.CDataType.dtString
        Me.txtEmail.DateValue = New Date(CType(0, Long))
        Me.txtEmail.DecimalDigits = CType(0, Long)
        Me.txtEmail.Location = New System.Drawing.Point(429, 73)
        Me.txtEmail.Mask = Nothing
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtEmail.OldValue = ""
        Me.txtEmail.Size = New System.Drawing.Size(277, 22)
        Me.txtEmail.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtEmail.StringValue = Nothing
        Me.txtEmail.TabIndex = 14
        Me.txtEmail.TextDataID = Nothing
        '
        'txtFax
        '
        Me.txtFax.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtFax.DataLength = CType(25, Long)
        Me.txtFax.DataType = Orders.CDataType.dtString
        Me.txtFax.DateValue = New Date(CType(0, Long))
        Me.txtFax.DecimalDigits = CType(0, Long)
        Me.txtFax.Location = New System.Drawing.Point(429, 47)
        Me.txtFax.Mask = Nothing
        Me.txtFax.Name = "txtFax"
        Me.txtFax.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtFax.OldValue = ""
        Me.txtFax.Size = New System.Drawing.Size(150, 22)
        Me.txtFax.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtFax.StringValue = Nothing
        Me.txtFax.TabIndex = 12
        Me.txtFax.TextDataID = Nothing
        '
        'txtLastName
        '
        Me.txtLastName.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtLastName.DataLength = CType(50, Long)
        Me.txtLastName.DataType = Orders.CDataType.dtString
        Me.txtLastName.DateValue = New Date(CType(0, Long))
        Me.txtLastName.DecimalDigits = CType(0, Long)
        Me.txtLastName.Location = New System.Drawing.Point(133, 73)
        Me.txtLastName.Mask = Nothing
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtLastName.OldValue = ""
        Me.txtLastName.Size = New System.Drawing.Size(150, 22)
        Me.txtLastName.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtLastName.StringValue = Nothing
        Me.txtLastName.TabIndex = 5
        Me.txtLastName.TextDataID = Nothing
        '
        'txtFirstName
        '
        Me.txtFirstName.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtFirstName.DataLength = CType(30, Long)
        Me.txtFirstName.DataType = Orders.CDataType.dtString
        Me.txtFirstName.DateValue = New Date(CType(0, Long))
        Me.txtFirstName.DecimalDigits = CType(0, Long)
        Me.txtFirstName.Location = New System.Drawing.Point(133, 47)
        Me.txtFirstName.Mask = Nothing
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtFirstName.OldValue = ""
        Me.txtFirstName.Size = New System.Drawing.Size(150, 22)
        Me.txtFirstName.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtFirstName.StringValue = Nothing
        Me.txtFirstName.TabIndex = 3
        Me.txtFirstName.TextDataID = Nothing
        '
        'lblTel
        '
        Me.lblTel.BackColor = System.Drawing.SystemColors.Control
        Me.lblTel.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTel.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblTel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTel.Location = New System.Drawing.Point(309, 24)
        Me.lblTel.Name = "lblTel"
        Me.lblTel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTel.Size = New System.Drawing.Size(125, 17)
        Me.lblTel.TabIndex = 8
        Me.lblTel.Text = "Telephone (+ Ext.)"
        '
        'lblFax
        '
        Me.lblFax.BackColor = System.Drawing.SystemColors.Control
        Me.lblFax.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFax.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblFax.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFax.Location = New System.Drawing.Point(310, 50)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFax.Size = New System.Drawing.Size(105, 17)
        Me.lblFax.TabIndex = 11
        Me.lblFax.Text = "Fax"
        '
        'lblEmail
        '
        Me.lblEmail.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmail.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEmail.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblEmail.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEmail.Location = New System.Drawing.Point(309, 76)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEmail.Size = New System.Drawing.Size(105, 17)
        Me.lblEmail.TabIndex = 13
        Me.lblEmail.Text = "Email"
        '
        'cboPred
        '
        Me.cboPred.BackColor = System.Drawing.SystemColors.Window
        Me.cboPred.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPred.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cboPred.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPred.Location = New System.Drawing.Point(133, 19)
        Me.cboPred.Name = "cboPred"
        Me.cboPred.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPred.Size = New System.Drawing.Size(73, 24)
        Me.cboPred.TabIndex = 1
        '
        'cboLang
        '
        Me.cboLang.BackColor = System.Drawing.SystemColors.Window
        Me.cboLang.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLang.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cboLang.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLang.Location = New System.Drawing.Point(133, 99)
        Me.cboLang.Name = "cboLang"
        Me.cboLang.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLang.Size = New System.Drawing.Size(73, 24)
        Me.cboLang.TabIndex = 7
        '
        'lblPred
        '
        Me.lblPred.BackColor = System.Drawing.SystemColors.Control
        Me.lblPred.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPred.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblPred.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPred.Location = New System.Drawing.Point(13, 24)
        Me.lblPred.Name = "lblPred"
        Me.lblPred.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPred.Size = New System.Drawing.Size(114, 17)
        Me.lblPred.TabIndex = 0
        Me.lblPred.Text = "Pred"
        '
        'lblFirstName
        '
        Me.lblFirstName.BackColor = System.Drawing.SystemColors.Control
        Me.lblFirstName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFirstName.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblFirstName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFirstName.Location = New System.Drawing.Point(13, 50)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFirstName.Size = New System.Drawing.Size(114, 17)
        Me.lblFirstName.TabIndex = 2
        Me.lblFirstName.Text = "First Name"
        '
        'lblLastName
        '
        Me.lblLastName.BackColor = System.Drawing.SystemColors.Control
        Me.lblLastName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLastName.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblLastName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLastName.Location = New System.Drawing.Point(13, 76)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLastName.Size = New System.Drawing.Size(114, 25)
        Me.lblLastName.TabIndex = 4
        Me.lblLastName.Text = "Last Name"
        '
        'lblLang
        '
        Me.lblLang.BackColor = System.Drawing.SystemColors.Control
        Me.lblLang.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLang.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblLang.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLang.Location = New System.Drawing.Point(13, 102)
        Me.lblLang.Name = "lblLang"
        Me.lblLang.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLang.Size = New System.Drawing.Size(124, 17)
        Me.lblLang.TabIndex = 6
        Me.lblLang.Text = "Preferred Language"
        '
        'ucContacts
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.Controls.Add(Me.grpAddContact)
        Me.Controls.Add(Me.fraCustomer)
        Me.Controls.Add(Me.fraContacts)
        Me.Controls.Add(Me.fraContactTypes)
        Me.Name = "ucContacts"
        Me.Size = New System.Drawing.Size(980, 555)
        Me.fraCustomer.ResumeLayout(False)
        Me.fraCustomer.PerformLayout()
        CType(Me.dgvContacts, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraContacts.ResumeLayout(False)
        Me.fraMethod.ResumeLayout(False)
        Me.fraType.ResumeLayout(False)
        Me.fraContactTypes.ResumeLayout(False)
        Me.fraContactTypes.PerformLayout()
        Me.grpAddContact.ResumeLayout(False)
        Me.grpAddContact.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdCopyContacts As System.Windows.Forms.Button
    Public WithEvents cmdRemoveContact As System.Windows.Forms.Button
    Public WithEvents fraCustomer As System.Windows.Forms.GroupBox
    Public WithEvents txtCusNo As System.Windows.Forms.TextBox
    Public WithEvents txtCusName As System.Windows.Forms.TextBox
    Friend WithEvents dgvContacts As System.Windows.Forms.DataGridView
    Public WithEvents optType As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents _optType_5 As System.Windows.Forms.RadioButton
    Public WithEvents _optType_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optType_2 As System.Windows.Forms.RadioButton
    Public WithEvents _optType_3 As System.Windows.Forms.RadioButton
    Public WithEvents _optType_4 As System.Windows.Forms.RadioButton
    Public WithEvents _optType_1 As System.Windows.Forms.RadioButton
    Public WithEvents cmdInsertNewContact As System.Windows.Forms.Button
    Public WithEvents fraContacts As System.Windows.Forms.GroupBox
    Public WithEvents fraContactTypes As System.Windows.Forms.GroupBox
    Public WithEvents grpAddContact As System.Windows.Forms.GroupBox
    Public WithEvents fraMethod As System.Windows.Forms.GroupBox
    Public WithEvents optFax As System.Windows.Forms.RadioButton
    Public WithEvents optEmail As System.Windows.Forms.RadioButton
    Public WithEvents fraType As System.Windows.Forms.GroupBox
    Public WithEvents optPrimary As System.Windows.Forms.RadioButton
    Public WithEvents optSecondary As System.Windows.Forms.RadioButton
    Public WithEvents lblTel As System.Windows.Forms.Label
    Public WithEvents lblFax As System.Windows.Forms.Label
    Public WithEvents lblEmail As System.Windows.Forms.Label
    Public WithEvents cboPred As System.Windows.Forms.ComboBox
    Public WithEvents cboLang As System.Windows.Forms.ComboBox
    Public WithEvents lblPred As System.Windows.Forms.Label
    Public WithEvents lblFirstName As System.Windows.Forms.Label
    Public WithEvents lblLastName As System.Windows.Forms.Label
    Public WithEvents lblLang As System.Windows.Forms.Label
    Friend WithEvents txtExt As Orders.xTextBox
    Friend WithEvents txtTel As Orders.xTextBox
    Friend WithEvents txtEmail As Orders.xTextBox
    Friend WithEvents txtFax As Orders.xTextBox
    Friend WithEvents txtLastName As Orders.xTextBox
    Friend WithEvents txtFirstName As Orders.xTextBox
    Public WithEvents optAllMethod As System.Windows.Forms.RadioButton
    Public WithEvents optAllType As System.Windows.Forms.RadioButton
    Friend WithEvents txt90DayRemind As Orders.xTextBox
    Friend WithEvents txtPreShipping As Orders.xTextBox
    Friend WithEvents txtShipping As Orders.xTextBox
    Friend WithEvents txtOrderAck As Orders.xTextBox
    Friend WithEvents txtPaperproof As Orders.xTextBox
    Friend WithEvents txtOEContact As Orders.xTextBox
    Friend WithEvents txtID As Orders.xTextBox
    Public WithEvents lblID As System.Windows.Forms.Label
    Public WithEvents cmdClear As System.Windows.Forms.Button

End Class
