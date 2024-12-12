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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProductLineEntry))
        Me.cmdAddToOrder = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdApplyToAll = New System.Windows.Forms.Button()
        Me.cmdResetAll = New System.Windows.Forms.Button()
        Me.dgvItems = New System.Windows.Forms.DataGridView()
        Me.cmdNotes = New System.Windows.Forms.Button()
        Me.cmdSearch = New System.Windows.Forms.Button()
        Me.lblRefill = New System.Windows.Forms.Label()
        Me.txtRefill = New System.Windows.Forms.TextBox()
        Me.cboRefill = New System.Windows.Forms.ComboBox()
        Me.lblPackaging = New System.Windows.Forms.Label()
        Me.txtPackaging = New System.Windows.Forms.TextBox()
        Me.cboPackaging = New System.Windows.Forms.ComboBox()
        Me.txtImprint_Color = New System.Windows.Forms.TextBox()
        Me.lblImprint_Color = New System.Windows.Forms.Label()
        Me.cboImprint_Color = New System.Windows.Forms.ComboBox()
        Me.lblImprint_Location = New System.Windows.Forms.Label()
        Me.txtImprint_Location = New System.Windows.Forms.TextBox()
        Me.lblSpecialComments = New System.Windows.Forms.Label()
        Me.txtSpecial_Comments = New System.Windows.Forms.TextBox()
        Me.txtComments = New System.Windows.Forms.TextBox()
        Me.lblComments = New System.Windows.Forms.Label()
        Me.txtNum_Impr_11 = New System.Windows.Forms.TextBox()
        Me.lblNum_Impr_1 = New System.Windows.Forms.Label()
        Me.cboImprint_Location = New System.Windows.Forms.ComboBox()
        Me.txtImprint = New System.Windows.Forms.TextBox()
        Me.panImprint = New System.Windows.Forms.Panel()
        Me.txtTipInPageTxt = New System.Windows.Forms.TextBox()
        Me.lblTipInPageTxt = New System.Windows.Forms.Label()
        Me.cboTipInPageTxt = New System.Windows.Forms.ComboBox()
        Me.txtRoundedCorners = New System.Windows.Forms.TextBox()
        Me.txtThreadColorScribl = New System.Windows.Forms.TextBox()
        Me.txtLamination = New System.Windows.Forms.TextBox()
        Me.lblRoundedCorners = New System.Windows.Forms.Label()
        Me.cboRoundedCorners = New System.Windows.Forms.ComboBox()
        Me.lblThreadColorScribl = New System.Windows.Forms.Label()
        Me.cbothreadColorScribl = New System.Windows.Forms.ComboBox()
        Me.lblLamination_Scribl = New System.Windows.Forms.Label()
        Me.cboLamination = New System.Windows.Forms.ComboBox()
        Me.lblEnd_User = New System.Windows.Forms.Label()
        Me.txtEnd_User = New System.Windows.Forms.TextBox()
        Me.txtNum_Impr_3 = New Orders.xTextBox()
        Me.lblNum_Impr_3 = New System.Windows.Forms.Label()
        Me.txtNum_Impr_1 = New Orders.xTextBox()
        Me.txtNum_Impr_2 = New Orders.xTextBox()
        Me.txtNum_Impr_12 = New System.Windows.Forms.TextBox()
        Me.lblNum_Impr_2 = New System.Windows.Forms.Label()
        Me.lblIndustry = New System.Windows.Forms.Label()
        Me.cboIndustry = New System.Windows.Forms.ComboBox()
        Me.lblImprint = New System.Windows.Forms.Label()
        Me.cmdCopy = New System.Windows.Forms.Button()
        Me.txtImprint_Item_No = New System.Windows.Forms.TextBox()
        Me.lblImprint_Item_No = New System.Windows.Forms.Label()
        Me.UcImprint1 = New Orders.ucImprint()
        Me.cmdUpdateSelected = New System.Windows.Forms.Button()
        Me.cmdUpdateAllItem = New System.Windows.Forms.Button()
        Me.cmdUpdateAll = New System.Windows.Forms.Button()
        Me.gbRepeatData = New System.Windows.Forms.GroupBox()
        Me.cmdGetRepeatData = New System.Windows.Forms.Button()
        Me.txtRepeat_Ord_No = New System.Windows.Forms.TextBox()
        Me.lblRepeat_Ord_No = New System.Windows.Forms.Label()
        Me.cmdClear_Imprint = New System.Windows.Forms.Button()
        Me.txtIndustry = New System.Windows.Forms.TextBox()
        Me.sbOrder = New System.Windows.Forms.StatusStrip()
        Me._sbOrder_Panel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblFoil_Color = New System.Windows.Forms.Label()
        Me.cboFoil_Color = New System.Windows.Forms.ComboBox()
        Me.txtFoil_Color = New System.Windows.Forms.TextBox()
        CType(Me.dgvItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panImprint.SuspendLayout()
        Me.gbRepeatData.SuspendLayout()
        Me.sbOrder.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdAddToOrder
        '
        Me.cmdAddToOrder.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddToOrder.Location = New System.Drawing.Point(797, 667)
        Me.cmdAddToOrder.Name = "cmdAddToOrder"
        Me.cmdAddToOrder.Size = New System.Drawing.Size(100, 24)
        Me.cmdAddToOrder.TabIndex = 1
        Me.cmdAddToOrder.Text = "&Add to order"
        Me.cmdAddToOrder.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.Location = New System.Drawing.Point(903, 667)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(100, 24)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdApplyToAll
        '
        Me.cmdApplyToAll.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdApplyToAll.Location = New System.Drawing.Point(12, 667)
        Me.cmdApplyToAll.Name = "cmdApplyToAll"
        Me.cmdApplyToAll.Size = New System.Drawing.Size(100, 24)
        Me.cmdApplyToAll.TabIndex = 3
        Me.cmdApplyToAll.Text = "Apply to all"
        Me.cmdApplyToAll.UseVisualStyleBackColor = True
        '
        'cmdResetAll
        '
        Me.cmdResetAll.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetAll.Location = New System.Drawing.Point(114, 667)
        Me.cmdResetAll.Name = "cmdResetAll"
        Me.cmdResetAll.Size = New System.Drawing.Size(100, 24)
        Me.cmdResetAll.TabIndex = 4
        Me.cmdResetAll.Text = "Reset All"
        Me.cmdResetAll.UseVisualStyleBackColor = True
        Me.cmdResetAll.Visible = False
        '
        'dgvItems
        '
        Me.dgvItems.AllowUserToAddRows = False
        Me.dgvItems.AllowUserToDeleteRows = False
        Me.dgvItems.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvItems.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvItems.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvItems.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke
        Me.dgvItems.GridColor = System.Drawing.SystemColors.Control
        Me.dgvItems.Location = New System.Drawing.Point(12, 12)
        Me.dgvItems.MultiSelect = False
        Me.dgvItems.Name = "dgvItems"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvItems.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvItems.RowHeadersWidth = 20
        Me.dgvItems.Size = New System.Drawing.Size(990, 438)
        Me.dgvItems.TabIndex = 0
        '
        'cmdNotes
        '
        Me.cmdNotes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNotes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNotes.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdNotes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNotes.Image = CType(resources.GetObject("cmdNotes.Image"), System.Drawing.Image)
        Me.cmdNotes.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdNotes.Location = New System.Drawing.Point(600, 450)
        Me.cmdNotes.Name = "cmdNotes"
        Me.cmdNotes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNotes.Size = New System.Drawing.Size(0, 0)
        Me.cmdNotes.TabIndex = 7
        Me.cmdNotes.TabStop = False
        Me.cmdNotes.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdNotes.UseVisualStyleBackColor = False
        '
        'cmdSearch
        '
        Me.cmdSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSearch.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSearch.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSearch.Image = CType(resources.GetObject("cmdSearch.Image"), System.Drawing.Image)
        Me.cmdSearch.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdSearch.Location = New System.Drawing.Point(610, 450)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSearch.Size = New System.Drawing.Size(0, 0)
        Me.cmdSearch.TabIndex = 8
        Me.cmdSearch.TabStop = False
        Me.cmdSearch.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdSearch.UseVisualStyleBackColor = False
        '
        'lblRefill
        '
        Me.lblRefill.AutoSize = True
        Me.lblRefill.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRefill.Location = New System.Drawing.Point(671, 113)
        Me.lblRefill.Name = "lblRefill"
        Me.lblRefill.Size = New System.Drawing.Size(35, 15)
        Me.lblRefill.TabIndex = 21
        Me.lblRefill.Text = "Refill"
        '
        'txtRefill
        '
        Me.txtRefill.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRefill.Location = New System.Drawing.Point(671, 140)
        Me.txtRefill.Name = "txtRefill"
        Me.txtRefill.Size = New System.Drawing.Size(170, 21)
        Me.txtRefill.TabIndex = 23
        '
        'cboRefill
        '
        Me.cboRefill.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRefill.FormattingEnabled = True
        Me.cboRefill.Location = New System.Drawing.Point(712, 110)
        Me.cboRefill.Name = "cboRefill"
        Me.cboRefill.Size = New System.Drawing.Size(129, 23)
        Me.cboRefill.TabIndex = 22
        '
        'lblPackaging
        '
        Me.lblPackaging.AutoSize = True
        Me.lblPackaging.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPackaging.Location = New System.Drawing.Point(495, 113)
        Me.lblPackaging.Name = "lblPackaging"
        Me.lblPackaging.Size = New System.Drawing.Size(37, 15)
        Me.lblPackaging.TabIndex = 18
        Me.lblPackaging.Text = "Pack."
        '
        'txtPackaging
        '
        Me.txtPackaging.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPackaging.Location = New System.Drawing.Point(495, 140)
        Me.txtPackaging.Name = "txtPackaging"
        Me.txtPackaging.Size = New System.Drawing.Size(170, 21)
        Me.txtPackaging.TabIndex = 20
        '
        'cboPackaging
        '
        Me.cboPackaging.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPackaging.FormattingEnabled = True
        Me.cboPackaging.Location = New System.Drawing.Point(538, 110)
        Me.cboPackaging.Name = "cboPackaging"
        Me.cboPackaging.Size = New System.Drawing.Size(127, 23)
        Me.cboPackaging.TabIndex = 19
        '
        'txtImprint_Color
        '
        Me.txtImprint_Color.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtImprint_Color.Location = New System.Drawing.Point(319, 140)
        Me.txtImprint_Color.Name = "txtImprint_Color"
        Me.txtImprint_Color.Size = New System.Drawing.Size(170, 21)
        Me.txtImprint_Color.TabIndex = 17
        '
        'lblImprint_Color
        '
        Me.lblImprint_Color.AutoSize = True
        Me.lblImprint_Color.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblImprint_Color.Location = New System.Drawing.Point(316, 114)
        Me.lblImprint_Color.Name = "lblImprint_Color"
        Me.lblImprint_Color.Size = New System.Drawing.Size(37, 15)
        Me.lblImprint_Color.TabIndex = 15
        Me.lblImprint_Color.Text = "Color"
        '
        'cboImprint_Color
        '
        Me.cboImprint_Color.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboImprint_Color.FormattingEnabled = True
        Me.cboImprint_Color.Location = New System.Drawing.Point(359, 110)
        Me.cboImprint_Color.Name = "cboImprint_Color"
        Me.cboImprint_Color.Size = New System.Drawing.Size(130, 23)
        Me.cboImprint_Color.TabIndex = 16
        '
        'lblImprint_Location
        '
        Me.lblImprint_Location.AutoSize = True
        Me.lblImprint_Location.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblImprint_Location.Location = New System.Drawing.Point(143, 114)
        Me.lblImprint_Location.Name = "lblImprint_Location"
        Me.lblImprint_Location.Size = New System.Drawing.Size(30, 15)
        Me.lblImprint_Location.TabIndex = 12
        Me.lblImprint_Location.Text = "Loc."
        '
        'txtImprint_Location
        '
        Me.txtImprint_Location.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtImprint_Location.Location = New System.Drawing.Point(143, 140)
        Me.txtImprint_Location.Name = "txtImprint_Location"
        Me.txtImprint_Location.Size = New System.Drawing.Size(170, 21)
        Me.txtImprint_Location.TabIndex = 14
        '
        'lblSpecialComments
        '
        Me.lblSpecialComments.AutoSize = True
        Me.lblSpecialComments.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSpecialComments.Location = New System.Drawing.Point(748, 6)
        Me.lblSpecialComments.Name = "lblSpecialComments"
        Me.lblSpecialComments.Size = New System.Drawing.Size(110, 15)
        Me.lblSpecialComments.TabIndex = 10
        Me.lblSpecialComments.Text = "Special comments"
        '
        'txtSpecial_Comments
        '
        Me.txtSpecial_Comments.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpecial_Comments.Location = New System.Drawing.Point(751, 25)
        Me.txtSpecial_Comments.Name = "txtSpecial_Comments"
        Me.txtSpecial_Comments.Size = New System.Drawing.Size(228, 21)
        Me.txtSpecial_Comments.TabIndex = 11
        '
        'txtComments
        '
        Me.txtComments.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtComments.Location = New System.Drawing.Point(517, 25)
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(228, 21)
        Me.txtComments.TabIndex = 9
        '
        'lblComments
        '
        Me.lblComments.AutoSize = True
        Me.lblComments.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblComments.Location = New System.Drawing.Point(514, 6)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(69, 15)
        Me.lblComments.TabIndex = 8
        Me.lblComments.Text = "Comments"
        '
        'txtNum_Impr_11
        '
        Me.txtNum_Impr_11.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNum_Impr_11.Location = New System.Drawing.Point(319, 25)
        Me.txtNum_Impr_11.Name = "txtNum_Impr_11"
        Me.txtNum_Impr_11.Size = New System.Drawing.Size(60, 21)
        Me.txtNum_Impr_11.TabIndex = 5
        Me.txtNum_Impr_11.Visible = False
        '
        'lblNum_Impr_1
        '
        Me.lblNum_Impr_1.AutoSize = True
        Me.lblNum_Impr_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNum_Impr_1.Location = New System.Drawing.Point(316, 6)
        Me.lblNum_Impr_1.Name = "lblNum_Impr_1"
        Me.lblNum_Impr_1.Size = New System.Drawing.Size(68, 15)
        Me.lblNum_Impr_1.TabIndex = 4
        Me.lblNum_Impr_1.Text = "Num Imp 1"
        '
        'cboImprint_Location
        '
        Me.cboImprint_Location.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboImprint_Location.FormattingEnabled = True
        Me.cboImprint_Location.Location = New System.Drawing.Point(179, 110)
        Me.cboImprint_Location.Name = "cboImprint_Location"
        Me.cboImprint_Location.Size = New System.Drawing.Size(134, 23)
        Me.cboImprint_Location.TabIndex = 13
        '
        'txtImprint
        '
        Me.txtImprint.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtImprint.Location = New System.Drawing.Point(143, 25)
        Me.txtImprint.Name = "txtImprint"
        Me.txtImprint.Size = New System.Drawing.Size(170, 21)
        Me.txtImprint.TabIndex = 3
        '
        'panImprint
        '
        Me.panImprint.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.panImprint.Controls.Add(Me.txtFoil_Color)
        Me.panImprint.Controls.Add(Me.lblFoil_Color)
        Me.panImprint.Controls.Add(Me.cboFoil_Color)
        Me.panImprint.Controls.Add(Me.txtTipInPageTxt)
        Me.panImprint.Controls.Add(Me.lblTipInPageTxt)
        Me.panImprint.Controls.Add(Me.cboTipInPageTxt)
        Me.panImprint.Controls.Add(Me.txtRoundedCorners)
        Me.panImprint.Controls.Add(Me.txtThreadColorScribl)
        Me.panImprint.Controls.Add(Me.txtLamination)
        Me.panImprint.Controls.Add(Me.lblRoundedCorners)
        Me.panImprint.Controls.Add(Me.cboRoundedCorners)
        Me.panImprint.Controls.Add(Me.lblThreadColorScribl)
        Me.panImprint.Controls.Add(Me.cbothreadColorScribl)
        Me.panImprint.Controls.Add(Me.lblLamination_Scribl)
        Me.panImprint.Controls.Add(Me.cboLamination)
        Me.panImprint.Controls.Add(Me.lblEnd_User)
        Me.panImprint.Controls.Add(Me.txtEnd_User)
        Me.panImprint.Controls.Add(Me.txtNum_Impr_3)
        Me.panImprint.Controls.Add(Me.lblNum_Impr_3)
        Me.panImprint.Controls.Add(Me.txtNum_Impr_1)
        Me.panImprint.Controls.Add(Me.txtNum_Impr_2)
        Me.panImprint.Controls.Add(Me.txtNum_Impr_12)
        Me.panImprint.Controls.Add(Me.lblNum_Impr_2)
        Me.panImprint.Controls.Add(Me.lblIndustry)
        Me.panImprint.Controls.Add(Me.cboIndustry)
        Me.panImprint.Controls.Add(Me.lblRefill)
        Me.panImprint.Controls.Add(Me.txtRefill)
        Me.panImprint.Controls.Add(Me.cboRefill)
        Me.panImprint.Controls.Add(Me.lblPackaging)
        Me.panImprint.Controls.Add(Me.txtPackaging)
        Me.panImprint.Controls.Add(Me.cboPackaging)
        Me.panImprint.Controls.Add(Me.txtImprint_Color)
        Me.panImprint.Controls.Add(Me.lblImprint_Color)
        Me.panImprint.Controls.Add(Me.cboImprint_Color)
        Me.panImprint.Controls.Add(Me.lblImprint_Location)
        Me.panImprint.Controls.Add(Me.txtImprint_Location)
        Me.panImprint.Controls.Add(Me.lblSpecialComments)
        Me.panImprint.Controls.Add(Me.txtSpecial_Comments)
        Me.panImprint.Controls.Add(Me.txtComments)
        Me.panImprint.Controls.Add(Me.lblComments)
        Me.panImprint.Controls.Add(Me.txtNum_Impr_11)
        Me.panImprint.Controls.Add(Me.lblNum_Impr_1)
        Me.panImprint.Controls.Add(Me.cboImprint_Location)
        Me.panImprint.Controls.Add(Me.txtImprint)
        Me.panImprint.Controls.Add(Me.lblImprint)
        Me.panImprint.Controls.Add(Me.cmdCopy)
        Me.panImprint.Controls.Add(Me.txtImprint_Item_No)
        Me.panImprint.Controls.Add(Me.lblImprint_Item_No)
        Me.panImprint.Controls.Add(Me.UcImprint1)
        Me.panImprint.Controls.Add(Me.cmdUpdateSelected)
        Me.panImprint.Controls.Add(Me.cmdUpdateAllItem)
        Me.panImprint.Controls.Add(Me.cmdUpdateAll)
        Me.panImprint.Controls.Add(Me.gbRepeatData)
        Me.panImprint.Controls.Add(Me.cmdClear_Imprint)
        Me.panImprint.Controls.Add(Me.txtIndustry)
        Me.panImprint.Location = New System.Drawing.Point(12, 456)
        Me.panImprint.Name = "panImprint"
        Me.panImprint.Size = New System.Drawing.Size(991, 205)
        Me.panImprint.TabIndex = 6
        '
        'txtTipInPageTxt
        '
        Me.txtTipInPageTxt.Location = New System.Drawing.Point(847, 78)
        Me.txtTipInPageTxt.Name = "txtTipInPageTxt"
        Me.txtTipInPageTxt.Size = New System.Drawing.Size(132, 20)
        Me.txtTipInPageTxt.TabIndex = 188
        '
        'lblTipInPageTxt
        '
        Me.lblTipInPageTxt.AutoSize = True
        Me.lblTipInPageTxt.Location = New System.Drawing.Point(789, 56)
        Me.lblTipInPageTxt.Name = "lblTipInPageTxt"
        Me.lblTipInPageTxt.Size = New System.Drawing.Size(62, 13)
        Me.lblTipInPageTxt.TabIndex = 186
        Me.lblTipInPageTxt.Text = "Tip In Page"
        '
        'cboTipInPageTxt
        '
        Me.cboTipInPageTxt.FormattingEnabled = True
        Me.cboTipInPageTxt.Location = New System.Drawing.Point(851, 52)
        Me.cboTipInPageTxt.Name = "cboTipInPageTxt"
        Me.cboTipInPageTxt.Size = New System.Drawing.Size(128, 21)
        Me.cboTipInPageTxt.TabIndex = 187
        '
        'txtRoundedCorners
        '
        Me.txtRoundedCorners.Location = New System.Drawing.Point(659, 78)
        Me.txtRoundedCorners.Name = "txtRoundedCorners"
        Me.txtRoundedCorners.Size = New System.Drawing.Size(128, 20)
        Me.txtRoundedCorners.TabIndex = 183
        '
        'txtThreadColorScribl
        '
        Me.txtThreadColorScribl.Location = New System.Drawing.Point(446, 78)
        Me.txtThreadColorScribl.Name = "txtThreadColorScribl"
        Me.txtThreadColorScribl.Size = New System.Drawing.Size(121, 20)
        Me.txtThreadColorScribl.TabIndex = 182
        '
        'txtLamination
        '
        Me.txtLamination.Location = New System.Drawing.Point(28, 79)
        Me.txtLamination.Name = "txtLamination"
        Me.txtLamination.Size = New System.Drawing.Size(129, 20)
        Me.txtLamination.TabIndex = 181
        '
        'lblRoundedCorners
        '
        Me.lblRoundedCorners.AutoSize = True
        Me.lblRoundedCorners.Location = New System.Drawing.Point(568, 56)
        Me.lblRoundedCorners.Name = "lblRoundedCorners"
        Me.lblRoundedCorners.Size = New System.Drawing.Size(81, 13)
        Me.lblRoundedCorners.TabIndex = 179
        Me.lblRoundedCorners.Text = "Round Corners."
        '
        'cboRoundedCorners
        '
        Me.cboRoundedCorners.FormattingEnabled = True
        Me.cboRoundedCorners.Location = New System.Drawing.Point(659, 53)
        Me.cboRoundedCorners.Name = "cboRoundedCorners"
        Me.cboRoundedCorners.Size = New System.Drawing.Size(128, 21)
        Me.cboRoundedCorners.TabIndex = 180
        '
        'lblThreadColorScribl
        '
        Me.lblThreadColorScribl.AutoSize = True
        Me.lblThreadColorScribl.Location = New System.Drawing.Point(371, 56)
        Me.lblThreadColorScribl.Name = "lblThreadColorScribl"
        Me.lblThreadColorScribl.Size = New System.Drawing.Size(71, 13)
        Me.lblThreadColorScribl.TabIndex = 177
        Me.lblThreadColorScribl.Text = "Thread Color."
        '
        'cbothreadColorScribl
        '
        Me.cbothreadColorScribl.FormattingEnabled = True
        Me.cbothreadColorScribl.Location = New System.Drawing.Point(446, 53)
        Me.cbothreadColorScribl.Name = "cbothreadColorScribl"
        Me.cbothreadColorScribl.Size = New System.Drawing.Size(121, 21)
        Me.cbothreadColorScribl.TabIndex = 178
        '
        'lblLamination_Scribl
        '
        Me.lblLamination_Scribl.AutoSize = True
        Me.lblLamination_Scribl.Location = New System.Drawing.Point(6, 56)
        Me.lblLamination_Scribl.Name = "lblLamination_Scribl"
        Me.lblLamination_Scribl.Size = New System.Drawing.Size(35, 13)
        Me.lblLamination_Scribl.TabIndex = 175
        Me.lblLamination_Scribl.Text = "Lamin"
        '
        'cboLamination
        '
        Me.cboLamination.FormattingEnabled = True
        Me.cboLamination.Location = New System.Drawing.Point(39, 53)
        Me.cboLamination.Name = "cboLamination"
        Me.cboLamination.Size = New System.Drawing.Size(134, 21)
        Me.cboLamination.TabIndex = 176
        '
        'lblEnd_User
        '
        Me.lblEnd_User.AutoSize = True
        Me.lblEnd_User.Location = New System.Drawing.Point(603, 9)
        Me.lblEnd_User.Name = "lblEnd_User"
        Me.lblEnd_User.Size = New System.Drawing.Size(51, 13)
        Me.lblEnd_User.TabIndex = 169
        Me.lblEnd_User.Text = "End User"
        Me.lblEnd_User.Visible = False
        '
        'txtEnd_User
        '
        Me.txtEnd_User.Location = New System.Drawing.Point(641, 6)
        Me.txtEnd_User.Name = "txtEnd_User"
        Me.txtEnd_User.Size = New System.Drawing.Size(39, 20)
        Me.txtEnd_User.TabIndex = 170
        Me.txtEnd_User.Visible = False
        '
        'txtNum_Impr_3
        '
        Me.txtNum_Impr_3.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtNum_Impr_3.DataLength = CType(0, Long)
        Me.txtNum_Impr_3.DataType = Orders.CDataType.dtNumWithoutDecimals
        Me.txtNum_Impr_3.DateValue = New Date(CType(0, Long))
        Me.txtNum_Impr_3.DecimalDigits = CType(0, Long)
        Me.txtNum_Impr_3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNum_Impr_3.Location = New System.Drawing.Point(451, 25)
        Me.txtNum_Impr_3.Mask = Nothing
        Me.txtNum_Impr_3.Name = "txtNum_Impr_3"
        Me.txtNum_Impr_3.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtNum_Impr_3.OldValue = ""
        Me.txtNum_Impr_3.Size = New System.Drawing.Size(60, 21)
        Me.txtNum_Impr_3.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtNum_Impr_3.StringValue = Nothing
        Me.txtNum_Impr_3.TabIndex = 168
        Me.txtNum_Impr_3.TextDataID = Nothing
        '
        'lblNum_Impr_3
        '
        Me.lblNum_Impr_3.AutoSize = True
        Me.lblNum_Impr_3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNum_Impr_3.Location = New System.Drawing.Point(448, 6)
        Me.lblNum_Impr_3.Name = "lblNum_Impr_3"
        Me.lblNum_Impr_3.Size = New System.Drawing.Size(68, 15)
        Me.lblNum_Impr_3.TabIndex = 167
        Me.lblNum_Impr_3.Text = "Num Imp 3"
        '
        'txtNum_Impr_1
        '
        Me.txtNum_Impr_1.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtNum_Impr_1.DataLength = CType(0, Long)
        Me.txtNum_Impr_1.DataType = Orders.CDataType.dtNumWithoutDecimals
        Me.txtNum_Impr_1.DateValue = New Date(CType(0, Long))
        Me.txtNum_Impr_1.DecimalDigits = CType(0, Long)
        Me.txtNum_Impr_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNum_Impr_1.Location = New System.Drawing.Point(319, 25)
        Me.txtNum_Impr_1.Mask = Nothing
        Me.txtNum_Impr_1.Name = "txtNum_Impr_1"
        Me.txtNum_Impr_1.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtNum_Impr_1.OldValue = ""
        Me.txtNum_Impr_1.Size = New System.Drawing.Size(60, 21)
        Me.txtNum_Impr_1.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtNum_Impr_1.StringValue = Nothing
        Me.txtNum_Impr_1.TabIndex = 5
        Me.txtNum_Impr_1.TextDataID = Nothing
        '
        'txtNum_Impr_2
        '
        Me.txtNum_Impr_2.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtNum_Impr_2.DataLength = CType(0, Long)
        Me.txtNum_Impr_2.DataType = Orders.CDataType.dtNumWithoutDecimals
        Me.txtNum_Impr_2.DateValue = New Date(CType(0, Long))
        Me.txtNum_Impr_2.DecimalDigits = CType(0, Long)
        Me.txtNum_Impr_2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNum_Impr_2.Location = New System.Drawing.Point(385, 25)
        Me.txtNum_Impr_2.Mask = Nothing
        Me.txtNum_Impr_2.Name = "txtNum_Impr_2"
        Me.txtNum_Impr_2.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtNum_Impr_2.OldValue = ""
        Me.txtNum_Impr_2.Size = New System.Drawing.Size(60, 21)
        Me.txtNum_Impr_2.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtNum_Impr_2.StringValue = Nothing
        Me.txtNum_Impr_2.TabIndex = 7
        Me.txtNum_Impr_2.TextDataID = Nothing
        '
        'txtNum_Impr_12
        '
        Me.txtNum_Impr_12.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNum_Impr_12.Location = New System.Drawing.Point(385, 25)
        Me.txtNum_Impr_12.Name = "txtNum_Impr_12"
        Me.txtNum_Impr_12.Size = New System.Drawing.Size(60, 21)
        Me.txtNum_Impr_12.TabIndex = 166
        Me.txtNum_Impr_12.Visible = False
        '
        'lblNum_Impr_2
        '
        Me.lblNum_Impr_2.AutoSize = True
        Me.lblNum_Impr_2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNum_Impr_2.Location = New System.Drawing.Point(382, 6)
        Me.lblNum_Impr_2.Name = "lblNum_Impr_2"
        Me.lblNum_Impr_2.Size = New System.Drawing.Size(68, 15)
        Me.lblNum_Impr_2.TabIndex = 6
        Me.lblNum_Impr_2.Text = "Num Imp 2"
        '
        'lblIndustry
        '
        Me.lblIndustry.AutoSize = True
        Me.lblIndustry.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIndustry.Location = New System.Drawing.Point(847, 113)
        Me.lblIndustry.Name = "lblIndustry"
        Me.lblIndustry.Size = New System.Drawing.Size(50, 15)
        Me.lblIndustry.TabIndex = 24
        Me.lblIndustry.Text = "Industry"
        '
        'cboIndustry
        '
        Me.cboIndustry.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboIndustry.FormattingEnabled = True
        Me.cboIndustry.Location = New System.Drawing.Point(847, 138)
        Me.cboIndustry.Name = "cboIndustry"
        Me.cboIndustry.Size = New System.Drawing.Size(132, 23)
        Me.cboIndustry.TabIndex = 25
        '
        'lblImprint
        '
        Me.lblImprint.AutoSize = True
        Me.lblImprint.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblImprint.Location = New System.Drawing.Point(143, 6)
        Me.lblImprint.Name = "lblImprint"
        Me.lblImprint.Size = New System.Drawing.Size(45, 15)
        Me.lblImprint.TabIndex = 2
        Me.lblImprint.Text = "Imprint"
        '
        'cmdCopy
        '
        Me.cmdCopy.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopy.Location = New System.Drawing.Point(568, 170)
        Me.cmdCopy.Name = "cmdCopy"
        Me.cmdCopy.Size = New System.Drawing.Size(120, 25)
        Me.cmdCopy.TabIndex = 29
        Me.cmdCopy.Text = "Copy"
        Me.cmdCopy.UseVisualStyleBackColor = True
        '
        'txtImprint_Item_No
        '
        Me.txtImprint_Item_No.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtImprint_Item_No.Location = New System.Drawing.Point(9, 25)
        Me.txtImprint_Item_No.Name = "txtImprint_Item_No"
        Me.txtImprint_Item_No.Size = New System.Drawing.Size(123, 22)
        Me.txtImprint_Item_No.TabIndex = 1
        '
        'lblImprint_Item_No
        '
        Me.lblImprint_Item_No.AutoSize = True
        Me.lblImprint_Item_No.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblImprint_Item_No.Location = New System.Drawing.Point(9, 6)
        Me.lblImprint_Item_No.Name = "lblImprint_Item_No"
        Me.lblImprint_Item_No.Size = New System.Drawing.Size(53, 16)
        Me.lblImprint_Item_No.TabIndex = 0
        Me.lblImprint_Item_No.Text = "Item No"
        '
        'UcImprint1
        '
        Me.UcImprint1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UcImprint1.Location = New System.Drawing.Point(889, 148)
        Me.UcImprint1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UcImprint1.Name = "UcImprint1"
        Me.UcImprint1.Size = New System.Drawing.Size(57, 46)
        Me.UcImprint1.TabIndex = 24
        Me.UcImprint1.Visible = False
        '
        'cmdUpdateSelected
        '
        Me.cmdUpdateSelected.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateSelected.Location = New System.Drawing.Point(694, 170)
        Me.cmdUpdateSelected.Name = "cmdUpdateSelected"
        Me.cmdUpdateSelected.Size = New System.Drawing.Size(120, 25)
        Me.cmdUpdateSelected.TabIndex = 30
        Me.cmdUpdateSelected.Text = "Paste"
        Me.cmdUpdateSelected.UseVisualStyleBackColor = True
        '
        'cmdUpdateAllItem
        '
        Me.cmdUpdateAllItem.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateAllItem.Location = New System.Drawing.Point(395, 170)
        Me.cmdUpdateAllItem.Name = "cmdUpdateAllItem"
        Me.cmdUpdateAllItem.Size = New System.Drawing.Size(167, 25)
        Me.cmdUpdateAllItem.TabIndex = 28
        Me.cmdUpdateAllItem.Text = "Update All 00EC1040"
        Me.cmdUpdateAllItem.UseVisualStyleBackColor = True
        '
        'cmdUpdateAll
        '
        Me.cmdUpdateAll.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateAll.Location = New System.Drawing.Point(269, 170)
        Me.cmdUpdateAll.Name = "cmdUpdateAll"
        Me.cmdUpdateAll.Size = New System.Drawing.Size(120, 25)
        Me.cmdUpdateAll.TabIndex = 27
        Me.cmdUpdateAll.Text = "Update All"
        Me.cmdUpdateAll.UseVisualStyleBackColor = True
        '
        'gbRepeatData
        '
        Me.gbRepeatData.Controls.Add(Me.cmdGetRepeatData)
        Me.gbRepeatData.Controls.Add(Me.txtRepeat_Ord_No)
        Me.gbRepeatData.Controls.Add(Me.lblRepeat_Ord_No)
        Me.gbRepeatData.Location = New System.Drawing.Point(3, 105)
        Me.gbRepeatData.Name = "gbRepeatData"
        Me.gbRepeatData.Size = New System.Drawing.Size(133, 90)
        Me.gbRepeatData.TabIndex = 31
        Me.gbRepeatData.TabStop = False
        '
        'cmdGetRepeatData
        '
        Me.cmdGetRepeatData.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetRepeatData.Location = New System.Drawing.Point(6, 59)
        Me.cmdGetRepeatData.Name = "cmdGetRepeatData"
        Me.cmdGetRepeatData.Size = New System.Drawing.Size(120, 25)
        Me.cmdGetRepeatData.TabIndex = 2
        Me.cmdGetRepeatData.Text = "Get Repeat Data"
        Me.cmdGetRepeatData.UseVisualStyleBackColor = True
        '
        'txtRepeat_Ord_No
        '
        Me.txtRepeat_Ord_No.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRepeat_Ord_No.Location = New System.Drawing.Point(6, 31)
        Me.txtRepeat_Ord_No.Name = "txtRepeat_Ord_No"
        Me.txtRepeat_Ord_No.Size = New System.Drawing.Size(120, 22)
        Me.txtRepeat_Ord_No.TabIndex = 1
        '
        'lblRepeat_Ord_No
        '
        Me.lblRepeat_Ord_No.AutoSize = True
        Me.lblRepeat_Ord_No.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRepeat_Ord_No.Location = New System.Drawing.Point(6, 12)
        Me.lblRepeat_Ord_No.Name = "lblRepeat_Ord_No"
        Me.lblRepeat_Ord_No.Size = New System.Drawing.Size(60, 16)
        Me.lblRepeat_Ord_No.TabIndex = 0
        Me.lblRepeat_Ord_No.Text = "Order No"
        '
        'cmdClear_Imprint
        '
        Me.cmdClear_Imprint.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClear_Imprint.Location = New System.Drawing.Point(143, 170)
        Me.cmdClear_Imprint.Name = "cmdClear_Imprint"
        Me.cmdClear_Imprint.Size = New System.Drawing.Size(120, 25)
        Me.cmdClear_Imprint.TabIndex = 26
        Me.cmdClear_Imprint.Text = "Clear"
        Me.cmdClear_Imprint.UseVisualStyleBackColor = True
        '
        'txtIndustry
        '
        Me.txtIndustry.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIndustry.Location = New System.Drawing.Point(847, 140)
        Me.txtIndustry.Name = "txtIndustry"
        Me.txtIndustry.Size = New System.Drawing.Size(132, 21)
        Me.txtIndustry.TabIndex = 163
        Me.txtIndustry.Visible = False
        '
        'sbOrder
        '
        Me.sbOrder.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sbOrder.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._sbOrder_Panel1})
        Me.sbOrder.Location = New System.Drawing.Point(0, 699)
        Me.sbOrder.Name = "sbOrder"
        Me.sbOrder.Size = New System.Drawing.Size(1004, 22)
        Me.sbOrder.TabIndex = 6
        '
        '_sbOrder_Panel1
        '
        Me._sbOrder_Panel1.AutoSize = False
        Me._sbOrder_Panel1.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._sbOrder_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._sbOrder_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me._sbOrder_Panel1.Name = "_sbOrder_Panel1"
        Me._sbOrder_Panel1.Size = New System.Drawing.Size(989, 22)
        Me._sbOrder_Panel1.Spring = True
        Me._sbOrder_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFoil_Color
        '
        Me.lblFoil_Color.AutoSize = True
        Me.lblFoil_Color.Location = New System.Drawing.Point(184, 56)
        Me.lblFoil_Color.Name = "lblFoil_Color"
        Me.lblFoil_Color.Size = New System.Drawing.Size(50, 13)
        Me.lblFoil_Color.TabIndex = 189
        Me.lblFoil_Color.Text = "Foil Color"
        '
        'cboFoil_Color
        '
        Me.cboFoil_Color.FormattingEnabled = True
        Me.cboFoil_Color.Location = New System.Drawing.Point(241, 53)
        Me.cboFoil_Color.Name = "cboFoil_Color"
        Me.cboFoil_Color.Size = New System.Drawing.Size(121, 21)
        Me.cboFoil_Color.TabIndex = 190
        '
        'txtFoil_Color
        '
        Me.txtFoil_Color.Location = New System.Drawing.Point(241, 79)
        Me.txtFoil_Color.Name = "txtFoil_Color"
        Me.txtFoil_Color.Size = New System.Drawing.Size(121, 20)
        Me.txtFoil_Color.TabIndex = 191
        '
        'frmProductLineEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1004, 721)
        Me.Controls.Add(Me.sbOrder)
        Me.Controls.Add(Me.panImprint)
        Me.Controls.Add(Me.cmdSearch)
        Me.Controls.Add(Me.cmdNotes)
        Me.Controls.Add(Me.dgvItems)
        Me.Controls.Add(Me.cmdResetAll)
        Me.Controls.Add(Me.cmdApplyToAll)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdAddToOrder)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(1020, 760)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1020, 760)
        Me.Name = "frmProductLineEntry"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmProductLineEntry"
        CType(Me.dgvItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panImprint.ResumeLayout(False)
        Me.panImprint.PerformLayout()
        Me.gbRepeatData.ResumeLayout(False)
        Me.gbRepeatData.PerformLayout()
        Me.sbOrder.ResumeLayout(False)
        Me.sbOrder.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdAddToOrder As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdApplyToAll As System.Windows.Forms.Button
    Friend WithEvents cmdResetAll As System.Windows.Forms.Button
    Friend WithEvents dgvItems As System.Windows.Forms.DataGridView
    Friend WithEvents cmdNotes As System.Windows.Forms.Button
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents lblRefill As System.Windows.Forms.Label
    Friend WithEvents txtRefill As System.Windows.Forms.TextBox
    Friend WithEvents cboRefill As System.Windows.Forms.ComboBox
    Friend WithEvents lblPackaging As System.Windows.Forms.Label
    Friend WithEvents txtPackaging As System.Windows.Forms.TextBox
    Friend WithEvents cboPackaging As System.Windows.Forms.ComboBox
    Friend WithEvents txtImprint_Color As System.Windows.Forms.TextBox
    Friend WithEvents lblImprint_Color As System.Windows.Forms.Label
    Friend WithEvents cboImprint_Color As System.Windows.Forms.ComboBox
    Friend WithEvents lblImprint_Location As System.Windows.Forms.Label
    Friend WithEvents txtImprint_Location As System.Windows.Forms.TextBox
    Friend WithEvents lblSpecialComments As System.Windows.Forms.Label
    Friend WithEvents txtSpecial_Comments As System.Windows.Forms.TextBox
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents txtNum_Impr_11 As System.Windows.Forms.TextBox
    Friend WithEvents lblNum_Impr_1 As System.Windows.Forms.Label
    Friend WithEvents cboImprint_Location As System.Windows.Forms.ComboBox
    Friend WithEvents txtImprint As System.Windows.Forms.TextBox
    Friend WithEvents panImprint As System.Windows.Forms.Panel
    Friend WithEvents lblImprint As System.Windows.Forms.Label
    Friend WithEvents cmdCopy As System.Windows.Forms.Button
    Friend WithEvents txtImprint_Item_No As System.Windows.Forms.TextBox
    Friend WithEvents lblImprint_Item_No As System.Windows.Forms.Label
    Friend WithEvents UcImprint1 As Orders.ucImprint
    Friend WithEvents cmdUpdateSelected As System.Windows.Forms.Button
    Friend WithEvents cmdUpdateAllItem As System.Windows.Forms.Button
    Friend WithEvents cmdUpdateAll As System.Windows.Forms.Button
    Friend WithEvents gbRepeatData As System.Windows.Forms.GroupBox
    Friend WithEvents cmdGetRepeatData As System.Windows.Forms.Button
    Friend WithEvents txtRepeat_Ord_No As System.Windows.Forms.TextBox
    Friend WithEvents lblRepeat_Ord_No As System.Windows.Forms.Label
    Friend WithEvents cmdClear_Imprint As System.Windows.Forms.Button
    Public WithEvents sbOrder As System.Windows.Forms.StatusStrip
    Public WithEvents _sbOrder_Panel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblIndustry As System.Windows.Forms.Label
    Friend WithEvents txtIndustry As System.Windows.Forms.TextBox
    Friend WithEvents cboIndustry As System.Windows.Forms.ComboBox
    Friend WithEvents txtNum_Impr_12 As System.Windows.Forms.TextBox
    Friend WithEvents lblNum_Impr_2 As System.Windows.Forms.Label
    Friend WithEvents txtNum_Impr_1 As Orders.xTextBox
    Friend WithEvents txtNum_Impr_2 As Orders.xTextBox
    Friend WithEvents txtNum_Impr_3 As Orders.xTextBox
    Friend WithEvents lblNum_Impr_3 As System.Windows.Forms.Label
    Friend WithEvents lblEnd_User As System.Windows.Forms.Label
    Friend WithEvents txtEnd_User As System.Windows.Forms.TextBox
    Friend WithEvents lblRoundedCorners As Label
    Friend WithEvents cboRoundedCorners As ComboBox
    Friend WithEvents lblThreadColorScribl As Label
    Friend WithEvents cbothreadColorScribl As ComboBox
    Friend WithEvents lblLamination_Scribl As Label
    Friend WithEvents cboLamination As ComboBox
    Friend WithEvents txtRoundedCorners As TextBox
    Friend WithEvents txtThreadColorScribl As TextBox
    Friend WithEvents txtLamination As TextBox
    Friend WithEvents txtTipInPageTxt As TextBox
    Friend WithEvents lblTipInPageTxt As Label
    Friend WithEvents cboTipInPageTxt As ComboBox
    Friend WithEvents txtFoil_Color As TextBox
    Friend WithEvents lblFoil_Color As Label
    Friend WithEvents cboFoil_Color As ComboBox
End Class
