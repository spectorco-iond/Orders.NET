<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSearch
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents optOr As System.Windows.Forms.RadioButton
	Public WithEvents optAnd As System.Windows.Forms.RadioButton
	Public WithEvents Frame2 As System.Windows.Forms.Panel
	Public WithEvents optStartsWith As System.Windows.Forms.RadioButton
	Public WithEvents optContains As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.Panel
    Public WithEvents Check1 As System.Windows.Forms.CheckBox
    Public WithEvents _txtElement_10 As System.Windows.Forms.TextBox
    Public WithEvents _txtElement_8 As System.Windows.Forms.TextBox
    Public WithEvents _txtElement_6 As System.Windows.Forms.TextBox
    Public WithEvents _txtElement_4 As System.Windows.Forms.TextBox
    Public WithEvents _txtElement_2 As System.Windows.Forms.TextBox
    Public WithEvents fraRightElements As System.Windows.Forms.Panel
    Public WithEvents Check2 As System.Windows.Forms.CheckBox
    Public WithEvents _lblElement_10 As System.Windows.Forms.Label
    Public WithEvents _lblElement_8 As System.Windows.Forms.Label
    Public WithEvents _lblElement_6 As System.Windows.Forms.Label
    Public WithEvents _lblElement_4 As System.Windows.Forms.Label
    Public WithEvents _lblElement_2 As System.Windows.Forms.Label
    Public WithEvents fraRightLabels As System.Windows.Forms.Panel
    Public WithEvents Check3 As System.Windows.Forms.CheckBox
    Public WithEvents _txtElement_9 As System.Windows.Forms.TextBox
    Public WithEvents _txtElement_7 As System.Windows.Forms.TextBox
    Public WithEvents _txtElement_5 As System.Windows.Forms.TextBox
    Public WithEvents _txtElement_3 As System.Windows.Forms.TextBox
    Public WithEvents _txtElement_1 As System.Windows.Forms.TextBox
    Public WithEvents fraLeftElements As System.Windows.Forms.Panel
    Public WithEvents cmdClose As System.Windows.Forms.Button
    Public WithEvents cmdSelect As System.Windows.Forms.Button
    Public WithEvents fraConfirmButtons As System.Windows.Forms.Panel
    Public WithEvents txtTopRows As System.Windows.Forms.TextBox
    Public WithEvents cmdNext As System.Windows.Forms.Button
    Public WithEvents cmdPrevious As System.Windows.Forms.Button
    Public WithEvents cmdClear As System.Windows.Forms.Button
    Public WithEvents cmdSearch As System.Windows.Forms.Button
    Public WithEvents Check4 As System.Windows.Forms.CheckBox
    Public WithEvents _lblElement_9 As System.Windows.Forms.Label
    Public WithEvents _lblElement_7 As System.Windows.Forms.Label
    Public WithEvents _lblElement_5 As System.Windows.Forms.Label
    Public WithEvents _lblElement_3 As System.Windows.Forms.Label
    Public WithEvents _lblElement_1 As System.Windows.Forms.Label
    Public WithEvents fraLeftLabels As System.Windows.Forms.Panel
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblTopRows As System.Windows.Forms.Label
    Public WithEvents lblElement As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents txtElement As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame2 = New System.Windows.Forms.Panel()
        Me.optOr = New System.Windows.Forms.RadioButton()
        Me.optAnd = New System.Windows.Forms.RadioButton()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.optStartsWith = New System.Windows.Forms.RadioButton()
        Me.optContains = New System.Windows.Forms.RadioButton()
        Me.fraRightElements = New System.Windows.Forms.Panel()
        Me.Check1 = New System.Windows.Forms.CheckBox()
        Me._txtElement_10 = New System.Windows.Forms.TextBox()
        Me._txtElement_8 = New System.Windows.Forms.TextBox()
        Me._txtElement_6 = New System.Windows.Forms.TextBox()
        Me._txtElement_4 = New System.Windows.Forms.TextBox()
        Me._txtElement_2 = New System.Windows.Forms.TextBox()
        Me._txtElement_0 = New System.Windows.Forms.TextBox()
        Me.fraRightLabels = New System.Windows.Forms.Panel()
        Me._lblElement_2 = New System.Windows.Forms.Label()
        Me._lblElement_10 = New System.Windows.Forms.Label()
        Me._lblElement_8 = New System.Windows.Forms.Label()
        Me._lblElement_6 = New System.Windows.Forms.Label()
        Me._lblElement_4 = New System.Windows.Forms.Label()
        Me.Check2 = New System.Windows.Forms.CheckBox()
        Me._lblElement_0 = New System.Windows.Forms.Label()
        Me.fraLeftElements = New System.Windows.Forms.Panel()
        Me.Check3 = New System.Windows.Forms.CheckBox()
        Me._txtElement_9 = New System.Windows.Forms.TextBox()
        Me._txtElement_7 = New System.Windows.Forms.TextBox()
        Me._txtElement_5 = New System.Windows.Forms.TextBox()
        Me._txtElement_3 = New System.Windows.Forms.TextBox()
        Me._txtElement_1 = New System.Windows.Forms.TextBox()
        Me.fraConfirmButtons = New System.Windows.Forms.Panel()
        Me.cmdOpen = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdSelect = New System.Windows.Forms.Button()
        Me.txtTopRows = New System.Windows.Forms.TextBox()
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.cmdPrevious = New System.Windows.Forms.Button()
        Me.cmdClear = New System.Windows.Forms.Button()
        Me.cmdSearch = New System.Windows.Forms.Button()
        Me.fraLeftLabels = New System.Windows.Forms.Panel()
        Me.Check4 = New System.Windows.Forms.CheckBox()
        Me._lblElement_9 = New System.Windows.Forms.Label()
        Me._lblElement_7 = New System.Windows.Forms.Label()
        Me._lblElement_5 = New System.Windows.Forms.Label()
        Me._lblElement_3 = New System.Windows.Forms.Label()
        Me._lblElement_1 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblTopRows = New System.Windows.Forms.Label()
        Me.lblElement = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.txtElement = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.dgvSearch = New System.Windows.Forms.DataGridView()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.fraRightElements.SuspendLayout()
        Me.fraRightLabels.SuspendLayout()
        Me.fraLeftElements.SuspendLayout()
        Me.fraConfirmButtons.SuspendLayout()
        Me.fraLeftLabels.SuspendLayout()
        CType(Me.lblElement, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtElement, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSearch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optOr)
        Me.Frame2.Controls.Add(Me.optAnd)
        Me.Frame2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(236, 32)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(125, 25)
        Me.Frame2.TabIndex = 8
        Me.Frame2.Text = "Frame1"
        '
        'optOr
        '
        Me.optOr.BackColor = System.Drawing.SystemColors.Control
        Me.optOr.Cursor = System.Windows.Forms.Cursors.Default
        Me.optOr.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optOr.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optOr.Location = New System.Drawing.Point(64, 0)
        Me.optOr.Name = "optOr"
        Me.optOr.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optOr.Size = New System.Drawing.Size(48, 25)
        Me.optOr.TabIndex = 1
        Me.optOr.TabStop = True
        Me.optOr.Text = "OR"
        Me.optOr.UseVisualStyleBackColor = False
        '
        'optAnd
        '
        Me.optAnd.BackColor = System.Drawing.SystemColors.Control
        Me.optAnd.Checked = True
        Me.optAnd.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAnd.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optAnd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAnd.Location = New System.Drawing.Point(0, 0)
        Me.optAnd.Name = "optAnd"
        Me.optAnd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAnd.Size = New System.Drawing.Size(58, 25)
        Me.optAnd.TabIndex = 0
        Me.optAnd.TabStop = True
        Me.optAnd.Text = "AND"
        Me.optAnd.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optStartsWith)
        Me.Frame1.Controls.Add(Me.optContains)
        Me.Frame1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(44, 32)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(197, 25)
        Me.Frame1.TabIndex = 7
        Me.Frame1.Text = "Frame1"
        '
        'optStartsWith
        '
        Me.optStartsWith.BackColor = System.Drawing.SystemColors.Control
        Me.optStartsWith.Checked = True
        Me.optStartsWith.Cursor = System.Windows.Forms.Cursors.Default
        Me.optStartsWith.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optStartsWith.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optStartsWith.Location = New System.Drawing.Point(12, 0)
        Me.optStartsWith.Name = "optStartsWith"
        Me.optStartsWith.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optStartsWith.Size = New System.Drawing.Size(94, 25)
        Me.optStartsWith.TabIndex = 0
        Me.optStartsWith.TabStop = True
        Me.optStartsWith.Text = "Starts with"
        Me.optStartsWith.UseVisualStyleBackColor = False
        '
        'optContains
        '
        Me.optContains.BackColor = System.Drawing.SystemColors.Control
        Me.optContains.Cursor = System.Windows.Forms.Cursors.Default
        Me.optContains.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optContains.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optContains.Location = New System.Drawing.Point(108, 0)
        Me.optContains.Name = "optContains"
        Me.optContains.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optContains.Size = New System.Drawing.Size(90, 25)
        Me.optContains.TabIndex = 1
        Me.optContains.Text = "Contains"
        Me.optContains.UseVisualStyleBackColor = False
        '
        'fraRightElements
        '
        Me.fraRightElements.BackColor = System.Drawing.SystemColors.Control
        Me.fraRightElements.Controls.Add(Me.Check1)
        Me.fraRightElements.Controls.Add(Me._txtElement_10)
        Me.fraRightElements.Controls.Add(Me._txtElement_8)
        Me.fraRightElements.Controls.Add(Me._txtElement_6)
        Me.fraRightElements.Controls.Add(Me._txtElement_4)
        Me.fraRightElements.Controls.Add(Me._txtElement_2)
        Me.fraRightElements.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraRightElements.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraRightElements.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraRightElements.Location = New System.Drawing.Point(352, 56)
        Me.fraRightElements.Name = "fraRightElements"
        Me.fraRightElements.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRightElements.Size = New System.Drawing.Size(267, 200)
        Me.fraRightElements.TabIndex = 12
        Me.fraRightElements.Text = "Frame2"
        '
        'Check1
        '
        Me.Check1.BackColor = System.Drawing.SystemColors.Control
        Me.Check1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Check1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check1.Location = New System.Drawing.Point(0, 184)
        Me.Check1.Name = "Check1"
        Me.Check1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check1.Size = New System.Drawing.Size(17, 17)
        Me.Check1.TabIndex = 34
        Me.Check1.Text = "Check1"
        Me.Check1.UseVisualStyleBackColor = False
        Me.Check1.Visible = False
        '
        '_txtElement_10
        '
        Me._txtElement_10.AcceptsReturn = True
        Me._txtElement_10.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_10.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_10.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_10, CType(10, Short))
        Me._txtElement_10.Location = New System.Drawing.Point(4, 116)
        Me._txtElement_10.MaxLength = 0
        Me._txtElement_10.Name = "_txtElement_10"
        Me._txtElement_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_10.Size = New System.Drawing.Size(133, 22)
        Me._txtElement_10.TabIndex = 27
        Me._txtElement_10.Visible = False
        '
        '_txtElement_8
        '
        Me._txtElement_8.AcceptsReturn = True
        Me._txtElement_8.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_8.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_8.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_8, CType(8, Short))
        Me._txtElement_8.Location = New System.Drawing.Point(4, 88)
        Me._txtElement_8.MaxLength = 0
        Me._txtElement_8.Name = "_txtElement_8"
        Me._txtElement_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_8.Size = New System.Drawing.Size(133, 22)
        Me._txtElement_8.TabIndex = 26
        Me._txtElement_8.Visible = False
        '
        '_txtElement_6
        '
        Me._txtElement_6.AcceptsReturn = True
        Me._txtElement_6.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_6.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_6.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_6, CType(6, Short))
        Me._txtElement_6.Location = New System.Drawing.Point(4, 60)
        Me._txtElement_6.MaxLength = 0
        Me._txtElement_6.Name = "_txtElement_6"
        Me._txtElement_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_6.Size = New System.Drawing.Size(133, 22)
        Me._txtElement_6.TabIndex = 25
        Me._txtElement_6.Visible = False
        '
        '_txtElement_4
        '
        Me._txtElement_4.AcceptsReturn = True
        Me._txtElement_4.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_4, CType(4, Short))
        Me._txtElement_4.Location = New System.Drawing.Point(4, 32)
        Me._txtElement_4.MaxLength = 10
        Me._txtElement_4.Name = "_txtElement_4"
        Me._txtElement_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_4.Size = New System.Drawing.Size(133, 22)
        Me._txtElement_4.TabIndex = 24
        Me._txtElement_4.Visible = False
        '
        '_txtElement_2
        '
        Me._txtElement_2.AcceptsReturn = True
        Me._txtElement_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_2, CType(2, Short))
        Me._txtElement_2.Location = New System.Drawing.Point(4, 4)
        Me._txtElement_2.MaxLength = 0
        Me._txtElement_2.Name = "_txtElement_2"
        Me._txtElement_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_2.Size = New System.Drawing.Size(133, 22)
        Me._txtElement_2.TabIndex = 23
        Me._txtElement_2.Visible = False
        '
        '_txtElement_0
        '
        Me._txtElement_0.AcceptsReturn = True
        Me._txtElement_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_0.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_0, CType(0, Short))
        Me._txtElement_0.Location = New System.Drawing.Point(432, 32)
        Me._txtElement_0.MaxLength = 0
        Me._txtElement_0.Name = "_txtElement_0"
        Me._txtElement_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_0.Size = New System.Drawing.Size(57, 22)
        Me._txtElement_0.TabIndex = 43
        Me._txtElement_0.Visible = False
        '
        'fraRightLabels
        '
        Me.fraRightLabels.BackColor = System.Drawing.SystemColors.Control
        Me.fraRightLabels.Controls.Add(Me._lblElement_2)
        Me.fraRightLabels.Controls.Add(Me._lblElement_10)
        Me.fraRightLabels.Controls.Add(Me._lblElement_8)
        Me.fraRightLabels.Controls.Add(Me._lblElement_6)
        Me.fraRightLabels.Controls.Add(Me._lblElement_4)
        Me.fraRightLabels.Controls.Add(Me.Check2)
        Me.fraRightLabels.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraRightLabels.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraRightLabels.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraRightLabels.Location = New System.Drawing.Point(252, 60)
        Me.fraRightLabels.Name = "fraRightLabels"
        Me.fraRightLabels.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRightLabels.Size = New System.Drawing.Size(267, 200)
        Me.fraRightLabels.TabIndex = 11
        Me.fraRightLabels.Text = "Frame1"
        '
        '_lblElement_2
        '
        Me._lblElement_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_2, CType(2, Short))
        Me._lblElement_2.Location = New System.Drawing.Point(4, 4)
        Me._lblElement_2.Name = "_lblElement_2"
        Me._lblElement_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_2.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_2.TabIndex = 17
        Me._lblElement_2.Text = "."
        Me._lblElement_2.Visible = False
        '
        '_lblElement_10
        '
        Me._lblElement_10.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_10.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_10, CType(10, Short))
        Me._lblElement_10.Location = New System.Drawing.Point(4, 116)
        Me._lblElement_10.Name = "_lblElement_10"
        Me._lblElement_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_10.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_10.TabIndex = 21
        Me._lblElement_10.Text = "."
        Me._lblElement_10.Visible = False
        '
        '_lblElement_8
        '
        Me._lblElement_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_8.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_8, CType(8, Short))
        Me._lblElement_8.Location = New System.Drawing.Point(4, 88)
        Me._lblElement_8.Name = "_lblElement_8"
        Me._lblElement_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_8.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_8.TabIndex = 20
        Me._lblElement_8.Text = "."
        Me._lblElement_8.Visible = False
        '
        '_lblElement_6
        '
        Me._lblElement_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_6.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_6, CType(6, Short))
        Me._lblElement_6.Location = New System.Drawing.Point(4, 60)
        Me._lblElement_6.Name = "_lblElement_6"
        Me._lblElement_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_6.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_6.TabIndex = 19
        Me._lblElement_6.Text = "."
        Me._lblElement_6.Visible = False
        '
        '_lblElement_4
        '
        Me._lblElement_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_4, CType(4, Short))
        Me._lblElement_4.Location = New System.Drawing.Point(4, 32)
        Me._lblElement_4.Name = "_lblElement_4"
        Me._lblElement_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_4.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_4.TabIndex = 18
        Me._lblElement_4.Text = "."
        Me._lblElement_4.Visible = False
        '
        'Check2
        '
        Me.Check2.BackColor = System.Drawing.SystemColors.Control
        Me.Check2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Check2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check2.Location = New System.Drawing.Point(0, 184)
        Me.Check2.Name = "Check2"
        Me.Check2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check2.Size = New System.Drawing.Size(17, 17)
        Me.Check2.TabIndex = 35
        Me.Check2.Text = "Check1"
        Me.Check2.UseVisualStyleBackColor = False
        Me.Check2.Visible = False
        '
        '_lblElement_0
        '
        Me._lblElement_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._lblElement_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_0.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_0, CType(0, Short))
        Me._lblElement_0.Location = New System.Drawing.Point(367, 32)
        Me._lblElement_0.Name = "_lblElement_0"
        Me._lblElement_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_0.Size = New System.Drawing.Size(59, 20)
        Me._lblElement_0.TabIndex = 44
        Me._lblElement_0.Text = "."
        Me._lblElement_0.Visible = False
        '
        'fraLeftElements
        '
        Me.fraLeftElements.BackColor = System.Drawing.SystemColors.Control
        Me.fraLeftElements.Controls.Add(Me.Check3)
        Me.fraLeftElements.Controls.Add(Me._txtElement_9)
        Me.fraLeftElements.Controls.Add(Me._txtElement_7)
        Me.fraLeftElements.Controls.Add(Me._txtElement_5)
        Me.fraLeftElements.Controls.Add(Me._txtElement_3)
        Me.fraLeftElements.Controls.Add(Me._txtElement_1)
        Me.fraLeftElements.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraLeftElements.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraLeftElements.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraLeftElements.Location = New System.Drawing.Point(104, 56)
        Me.fraLeftElements.Name = "fraLeftElements"
        Me.fraLeftElements.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraLeftElements.Size = New System.Drawing.Size(267, 200)
        Me.fraLeftElements.TabIndex = 10
        Me.fraLeftElements.Text = "Frame1"
        '
        'Check3
        '
        Me.Check3.BackColor = System.Drawing.SystemColors.Control
        Me.Check3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Check3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check3.Location = New System.Drawing.Point(0, 184)
        Me.Check3.Name = "Check3"
        Me.Check3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check3.Size = New System.Drawing.Size(17, 17)
        Me.Check3.TabIndex = 36
        Me.Check3.Text = "Check1"
        Me.Check3.UseVisualStyleBackColor = False
        Me.Check3.Visible = False
        '
        '_txtElement_9
        '
        Me._txtElement_9.AcceptsReturn = True
        Me._txtElement_9.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_9.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_9.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_9, CType(9, Short))
        Me._txtElement_9.Location = New System.Drawing.Point(4, 116)
        Me._txtElement_9.MaxLength = 0
        Me._txtElement_9.Name = "_txtElement_9"
        Me._txtElement_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_9.Size = New System.Drawing.Size(134, 22)
        Me._txtElement_9.TabIndex = 15
        Me._txtElement_9.Visible = False
        '
        '_txtElement_7
        '
        Me._txtElement_7.AcceptsReturn = True
        Me._txtElement_7.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_7.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_7.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_7, CType(7, Short))
        Me._txtElement_7.Location = New System.Drawing.Point(4, 88)
        Me._txtElement_7.MaxLength = 0
        Me._txtElement_7.Name = "_txtElement_7"
        Me._txtElement_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_7.Size = New System.Drawing.Size(134, 22)
        Me._txtElement_7.TabIndex = 3
        Me._txtElement_7.Visible = False
        '
        '_txtElement_5
        '
        Me._txtElement_5.AcceptsReturn = True
        Me._txtElement_5.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_5, CType(5, Short))
        Me._txtElement_5.Location = New System.Drawing.Point(4, 60)
        Me._txtElement_5.MaxLength = 0
        Me._txtElement_5.Name = "_txtElement_5"
        Me._txtElement_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_5.Size = New System.Drawing.Size(134, 22)
        Me._txtElement_5.TabIndex = 2
        Me._txtElement_5.Visible = False
        '
        '_txtElement_3
        '
        Me._txtElement_3.AcceptsReturn = True
        Me._txtElement_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_3, CType(3, Short))
        Me._txtElement_3.Location = New System.Drawing.Point(4, 32)
        Me._txtElement_3.MaxLength = 0
        Me._txtElement_3.Name = "_txtElement_3"
        Me._txtElement_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_3.Size = New System.Drawing.Size(134, 22)
        Me._txtElement_3.TabIndex = 1
        Me._txtElement_3.Visible = False
        '
        '_txtElement_1
        '
        Me._txtElement_1.AcceptsReturn = True
        Me._txtElement_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtElement_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtElement_1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtElement_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtElement.SetIndex(Me._txtElement_1, CType(1, Short))
        Me._txtElement_1.Location = New System.Drawing.Point(4, 4)
        Me._txtElement_1.MaxLength = 0
        Me._txtElement_1.Name = "_txtElement_1"
        Me._txtElement_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtElement_1.Size = New System.Drawing.Size(134, 22)
        Me._txtElement_1.TabIndex = 0
        Me._txtElement_1.Visible = False
        '
        'fraConfirmButtons
        '
        Me.fraConfirmButtons.BackColor = System.Drawing.SystemColors.Control
        Me.fraConfirmButtons.Controls.Add(Me.cmdOpen)
        Me.fraConfirmButtons.Controls.Add(Me.cmdClose)
        Me.fraConfirmButtons.Controls.Add(Me.cmdSelect)
        Me.fraConfirmButtons.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraConfirmButtons.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraConfirmButtons.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraConfirmButtons.Location = New System.Drawing.Point(205, 392)
        Me.fraConfirmButtons.Name = "fraConfirmButtons"
        Me.fraConfirmButtons.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraConfirmButtons.Size = New System.Drawing.Size(300, 45)
        Me.fraConfirmButtons.TabIndex = 14
        Me.fraConfirmButtons.Text = "Frame1"
        '
        'cmdOpen
        '
        Me.cmdOpen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOpen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOpen.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOpen.Location = New System.Drawing.Point(24, 8)
        Me.cmdOpen.Name = "cmdOpen"
        Me.cmdOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOpen.Size = New System.Drawing.Size(81, 24)
        Me.cmdOpen.TabIndex = 0
        Me.cmdOpen.Text = "Open"
        Me.cmdOpen.UseVisualStyleBackColor = False
        Me.cmdOpen.Visible = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(199, 8)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(85, 24)
        Me.cmdClose.TabIndex = 2
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'cmdSelect
        '
        Me.cmdSelect.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSelect.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSelect.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSelect.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSelect.Location = New System.Drawing.Point(111, 8)
        Me.cmdSelect.Name = "cmdSelect"
        Me.cmdSelect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSelect.Size = New System.Drawing.Size(81, 24)
        Me.cmdSelect.TabIndex = 1
        Me.cmdSelect.Text = "Select"
        Me.cmdSelect.UseVisualStyleBackColor = False
        '
        'txtTopRows
        '
        Me.txtTopRows.AcceptsReturn = True
        Me.txtTopRows.BackColor = System.Drawing.SystemColors.Window
        Me.txtTopRows.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTopRows.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTopRows.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTopRows.Location = New System.Drawing.Point(360, 4)
        Me.txtTopRows.MaxLength = 0
        Me.txtTopRows.Name = "txtTopRows"
        Me.txtTopRows.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTopRows.Size = New System.Drawing.Size(37, 22)
        Me.txtTopRows.TabIndex = 4
        Me.txtTopRows.Text = "50"
        '
        'cmdNext
        '
        Me.cmdNext.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNext.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNext.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNext.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNext.Location = New System.Drawing.Point(268, 4)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNext.Size = New System.Drawing.Size(80, 24)
        Me.cmdNext.TabIndex = 3
        Me.cmdNext.Text = "Next"
        Me.cmdNext.UseVisualStyleBackColor = False
        '
        'cmdPrevious
        '
        Me.cmdPrevious.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPrevious.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPrevious.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrevious.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPrevious.Location = New System.Drawing.Point(180, 4)
        Me.cmdPrevious.Name = "cmdPrevious"
        Me.cmdPrevious.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPrevious.Size = New System.Drawing.Size(80, 24)
        Me.cmdPrevious.TabIndex = 2
        Me.cmdPrevious.Text = "Previous"
        Me.cmdPrevious.UseVisualStyleBackColor = False
        '
        'cmdClear
        '
        Me.cmdClear.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClear.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClear.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClear.Location = New System.Drawing.Point(92, 4)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClear.Size = New System.Drawing.Size(80, 24)
        Me.cmdClear.TabIndex = 1
        Me.cmdClear.Text = "Clear"
        Me.cmdClear.UseVisualStyleBackColor = False
        '
        'cmdSearch
        '
        Me.cmdSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSearch.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSearch.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSearch.Location = New System.Drawing.Point(4, 4)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSearch.Size = New System.Drawing.Size(80, 24)
        Me.cmdSearch.TabIndex = 0
        Me.cmdSearch.Text = "Search"
        Me.cmdSearch.UseVisualStyleBackColor = False
        '
        'fraLeftLabels
        '
        Me.fraLeftLabels.BackColor = System.Drawing.SystemColors.Control
        Me.fraLeftLabels.Controls.Add(Me.Check4)
        Me.fraLeftLabels.Controls.Add(Me._lblElement_9)
        Me.fraLeftLabels.Controls.Add(Me._lblElement_7)
        Me.fraLeftLabels.Controls.Add(Me._lblElement_5)
        Me.fraLeftLabels.Controls.Add(Me._lblElement_3)
        Me.fraLeftLabels.Controls.Add(Me._lblElement_1)
        Me.fraLeftLabels.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraLeftLabels.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraLeftLabels.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraLeftLabels.Location = New System.Drawing.Point(4, 60)
        Me.fraLeftLabels.Name = "fraLeftLabels"
        Me.fraLeftLabels.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraLeftLabels.Size = New System.Drawing.Size(267, 200)
        Me.fraLeftLabels.TabIndex = 9
        Me.fraLeftLabels.Text = "Frame3"
        '
        'Check4
        '
        Me.Check4.BackColor = System.Drawing.SystemColors.Control
        Me.Check4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Check4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check4.Location = New System.Drawing.Point(0, 184)
        Me.Check4.Name = "Check4"
        Me.Check4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check4.Size = New System.Drawing.Size(17, 17)
        Me.Check4.TabIndex = 37
        Me.Check4.Text = "Check1"
        Me.Check4.UseVisualStyleBackColor = False
        Me.Check4.Visible = False
        '
        '_lblElement_9
        '
        Me._lblElement_9.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_9.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_9, CType(9, Short))
        Me._lblElement_9.Location = New System.Drawing.Point(4, 116)
        Me._lblElement_9.Name = "_lblElement_9"
        Me._lblElement_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_9.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_9.TabIndex = 4
        Me._lblElement_9.Text = "."
        Me._lblElement_9.Visible = False
        '
        '_lblElement_7
        '
        Me._lblElement_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_7.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_7, CType(7, Short))
        Me._lblElement_7.Location = New System.Drawing.Point(4, 88)
        Me._lblElement_7.Name = "_lblElement_7"
        Me._lblElement_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_7.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_7.TabIndex = 3
        Me._lblElement_7.Text = "."
        Me._lblElement_7.Visible = False
        '
        '_lblElement_5
        '
        Me._lblElement_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_5, CType(5, Short))
        Me._lblElement_5.Location = New System.Drawing.Point(4, 60)
        Me._lblElement_5.Name = "_lblElement_5"
        Me._lblElement_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_5.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_5.TabIndex = 2
        Me._lblElement_5.Text = "."
        Me._lblElement_5.Visible = False
        '
        '_lblElement_3
        '
        Me._lblElement_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_3, CType(3, Short))
        Me._lblElement_3.Location = New System.Drawing.Point(4, 32)
        Me._lblElement_3.Name = "_lblElement_3"
        Me._lblElement_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_3.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_3.TabIndex = 1
        Me._lblElement_3.Text = "."
        Me._lblElement_3.Visible = False
        '
        '_lblElement_1
        '
        Me._lblElement_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblElement_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblElement_1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblElement_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblElement.SetIndex(Me._lblElement_1, CType(1, Short))
        Me._lblElement_1.Location = New System.Drawing.Point(4, 4)
        Me._lblElement_1.Name = "_lblElement_1"
        Me._lblElement_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblElement_1.Size = New System.Drawing.Size(200, 20)
        Me._lblElement_1.TabIndex = 0
        Me._lblElement_1.Text = "."
        Me._lblElement_1.Visible = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 21)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Filter"
        '
        'lblTopRows
        '
        Me.lblTopRows.BackColor = System.Drawing.SystemColors.Control
        Me.lblTopRows.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTopRows.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTopRows.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTopRows.Location = New System.Drawing.Point(396, 8)
        Me.lblTopRows.Name = "lblTopRows"
        Me.lblTopRows.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTopRows.Size = New System.Drawing.Size(49, 20)
        Me.lblTopRows.TabIndex = 5
        Me.lblTopRows.Text = "Rows"
        Me.lblTopRows.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtElement
        '
        '
        'dgvSearch
        '
        Me.dgvSearch.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSearch.Location = New System.Drawing.Point(8, 200)
        Me.dgvSearch.Name = "dgvSearch"
        Me.dgvSearch.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSearch.Size = New System.Drawing.Size(481, 194)
        Me.dgvSearch.TabIndex = 13
        '
        'frmSearch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(497, 432)
        Me.Controls.Add(Me.dgvSearch)
        Me.Controls.Add(Me._lblElement_0)
        Me.Controls.Add(Me._txtElement_0)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.fraRightElements)
        Me.Controls.Add(Me.txtTopRows)
        Me.Controls.Add(Me.cmdNext)
        Me.Controls.Add(Me.cmdPrevious)
        Me.Controls.Add(Me.cmdClear)
        Me.Controls.Add(Me.cmdSearch)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblTopRows)
        Me.Controls.Add(Me.fraConfirmButtons)
        Me.Controls.Add(Me.fraRightLabels)
        Me.Controls.Add(Me.fraLeftElements)
        Me.Controls.Add(Me.fraLeftLabels)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSearch"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Search"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.fraRightElements.ResumeLayout(False)
        Me.fraRightElements.PerformLayout()
        Me.fraRightLabels.ResumeLayout(False)
        Me.fraLeftElements.ResumeLayout(False)
        Me.fraLeftElements.PerformLayout()
        Me.fraConfirmButtons.ResumeLayout(False)
        Me.fraLeftLabels.ResumeLayout(False)
        CType(Me.lblElement, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtElement, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSearch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents _txtElement_0 As System.Windows.Forms.TextBox
    Public WithEvents _lblElement_0 As System.Windows.Forms.Label
    Friend WithEvents dgvSearch As System.Windows.Forms.DataGridView
    Public WithEvents cmdOpen As System.Windows.Forms.Button
#End Region
End Class