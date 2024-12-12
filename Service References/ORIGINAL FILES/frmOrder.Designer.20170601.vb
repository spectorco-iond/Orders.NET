<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOrder
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
    Public WithEvents _sstOrder_T04 As System.Windows.Forms.TabPage
    Public WithEvents _sstOrder_T02 As System.Windows.Forms.TabPage
    Public WithEvents _sstOrder_T05 As System.Windows.Forms.TabPage
    Public WithEvents _sstOrder_T06 As System.Windows.Forms.TabPage
    Public WithEvents _sstOrder_T07 As System.Windows.Forms.TabPage
    Public WithEvents _sstOrder_T08 As System.Windows.Forms.TabPage
    Public WithEvents _sbOrder_Panel1 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents sbOrder As System.Windows.Forms.StatusStrip
    Public WithEvents tbOrder As System.Windows.Forms.ToolStrip
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOrder))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.sstOrder = New System.Windows.Forms.TabControl()
        Me._sstOrder_T00 = New System.Windows.Forms.TabPage()
        Me.testSimulateThermo = New System.Windows.Forms.Button()
        Me.cmdETA_Calculator = New System.Windows.Forms.Button()
        Me.cmdReset_Cus_No = New System.Windows.Forms.Button()
        Me.cmdEDI_Alert_New_Inv = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.cmdMySqlTest = New System.Windows.Forms.Button()
        Me.btnThermo = New System.Windows.Forms.Button()
        Me.btnFileList = New System.Windows.Forms.Button()
        Me.btnEDI = New System.Windows.Forms.Button()
        Me.cmdOEI_Config = New System.Windows.Forms.Button()
        Me.cmdDirServices = New System.Windows.Forms.Button()
        Me.btnODBCConnect = New System.Windows.Forms.Button()
        Me.cmdEmail = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.UcOrder = New Orders.ucOrder()
        Me._sstOrder_T01 = New System.Windows.Forms.TabPage()
        Me.UcContacts1 = New Orders.ucContacts()
        Me._sstOrder_T02 = New System.Windows.Forms.TabPage()
        Me.cmdEDIEditSuppItem = New System.Windows.Forms.Button()
        Me.cmdDash = New System.Windows.Forms.Button()
        Me.cmdTraveler = New System.Windows.Forms.Button()
        Me.cmdPrice = New System.Windows.Forms.Button()
        Me.chkAutoCompleteReship = New System.Windows.Forms.CheckBox()
        Me.cmdQty = New System.Windows.Forms.Button()
        Me.cmdRequest_Dt = New System.Windows.Forms.Button()
        Me.cmdImprint = New System.Windows.Forms.Button()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdPrevCharges = New System.Windows.Forms.Button()
        Me.cboMoreOptions = New System.Windows.Forms.ComboBox()
        Me.chkLowRow = New System.Windows.Forms.CheckBox()
        Me.panImprint = New System.Windows.Forms.Panel()
        Me.btnDbgPrxAdjust = New System.Windows.Forms.Button()
        Me.lblEnd_User = New System.Windows.Forms.Label()
        Me.txtEnd_User = New System.Windows.Forms.TextBox()
        Me.txtNum_Impr_3 = New Orders.xTextBox()
        Me.lblNum_Impr_3 = New System.Windows.Forms.Label()
        Me.txtNum_Impr_1 = New Orders.xTextBox()
        Me.txtNum_Impr_2 = New Orders.xTextBox()
        Me.txtNum_Impr_12 = New System.Windows.Forms.TextBox()
        Me.lblNum_Impr_2 = New System.Windows.Forms.Label()
        Me.txtRepeat_From_ID = New System.Windows.Forms.TextBox()
        Me.cboIndustry = New System.Windows.Forms.ComboBox()
        Me.lblIndustry = New System.Windows.Forms.Label()
        Me.txtIndustry = New System.Windows.Forms.TextBox()
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
        Me.lblImprint = New System.Windows.Forms.Label()
        Me.cmdCopyDown = New System.Windows.Forms.Button()
        Me.txtImprint_Item_No = New System.Windows.Forms.TextBox()
        Me.lblImprint_Item_No = New System.Windows.Forms.Label()
        Me.UcImprint1 = New Orders.ucImprint()
        Me.cmdPaste = New System.Windows.Forms.Button()
        Me.cmdUpdateAllItem = New System.Windows.Forms.Button()
        Me.cmdUpdateAll = New System.Windows.Forms.Button()
        Me.gbRepeatData = New System.Windows.Forms.GroupBox()
        Me.cmdRepeatSearch = New System.Windows.Forms.Button()
        Me.cmdGetRepeatData = New System.Windows.Forms.Button()
        Me.txtRepeat_Ord_No = New System.Windows.Forms.TextBox()
        Me.lblRepeat_Ord_No = New System.Windows.Forms.Label()
        Me.cmdClear_Imprint = New System.Windows.Forms.Button()
        Me.dgvItems = New Orders.XDataGridView()
        Me.cmdNotes = New System.Windows.Forms.Button()
        Me.cmdSearch = New System.Windows.Forms.Button()
        Me.cmdRestart = New System.Windows.Forms.Button()
        Me.cmdCopyLine = New System.Windows.Forms.Button()
        Me.cmdDeleteLine = New System.Windows.Forms.Button()
        Me.cmdInsertLine = New System.Windows.Forms.Button()
        Me.XTextBox2 = New Orders.xTextBox()
        Me.UcOrderTotal1 = New Orders.ucOrderTotal()
        Me._sstOrder_T03 = New System.Windows.Forms.TabPage()
        Me.UcDocument1 = New Orders.ucDocument()
        Me._sstOrder_T04 = New System.Windows.Forms.TabPage()
        Me.UcHeader1 = New Orders.ucHeader()
        Me._sstOrder_T05 = New System.Windows.Forms.TabPage()
        Me.UcTaxes1 = New Orders.ucTaxes()
        Me._sstOrder_T06 = New System.Windows.Forms.TabPage()
        Me.UcSalesperson1 = New Orders.ucSalesperson()
        Me._sstOrder_T07 = New System.Windows.Forms.TabPage()
        Me.UcCreditInfo1 = New Orders.ucCreditInfo()
        Me._sstOrder_T08 = New System.Windows.Forms.TabPage()
        Me.cmdTest = New System.Windows.Forms.Button()
        Me.UcExtra1 = New Orders.ucExtra()
        Me._sstOrder_T09 = New System.Windows.Forms.TabPage()
        Me.wbItem = New System.Windows.Forms.WebBrowser()
        Me.dgvHistory = New Orders.XDataGridView()
        Me._sstOrder_T10 = New System.Windows.Forms.TabPage()
        Me.dgvReservedItems = New Orders.XDataGridView()
        Me.sbOrder = New System.Windows.Forms.StatusStrip()
        Me._sbOrder_Panel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tbOrder = New System.Windows.Forms.ToolStrip()
        Me.tsbNew = New System.Windows.Forms.ToolStripButton()
        Me.tsbExportToMacola = New System.Windows.Forms.ToolStripButton()
        Me.tsbXmlExport = New System.Windows.Forms.ToolStripButton()
        Me.lblWhite_Glove = New System.Windows.Forms.Label()
        Me.sstOrder.SuspendLayout()
        Me._sstOrder_T00.SuspendLayout()
        Me._sstOrder_T01.SuspendLayout()
        Me._sstOrder_T02.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panImprint.SuspendLayout()
        Me.gbRepeatData.SuspendLayout()
        CType(Me.dgvItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._sstOrder_T03.SuspendLayout()
        Me._sstOrder_T04.SuspendLayout()
        Me._sstOrder_T05.SuspendLayout()
        Me._sstOrder_T06.SuspendLayout()
        Me._sstOrder_T07.SuspendLayout()
        Me._sstOrder_T08.SuspendLayout()
        Me._sstOrder_T09.SuspendLayout()
        CType(Me.dgvHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._sstOrder_T10.SuspendLayout()
        CType(Me.dgvReservedItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.sbOrder.SuspendLayout()
        Me.tbOrder.SuspendLayout()
        Me.SuspendLayout()
        '
        'sstOrder
        '
        Me.sstOrder.Controls.Add(Me._sstOrder_T00)
        Me.sstOrder.Controls.Add(Me._sstOrder_T01)
        Me.sstOrder.Controls.Add(Me._sstOrder_T02)
        Me.sstOrder.Controls.Add(Me._sstOrder_T03)
        Me.sstOrder.Controls.Add(Me._sstOrder_T04)
        Me.sstOrder.Controls.Add(Me._sstOrder_T05)
        Me.sstOrder.Controls.Add(Me._sstOrder_T06)
        Me.sstOrder.Controls.Add(Me._sstOrder_T07)
        Me.sstOrder.Controls.Add(Me._sstOrder_T08)
        Me.sstOrder.Controls.Add(Me._sstOrder_T09)
        Me.sstOrder.Controls.Add(Me._sstOrder_T10)
        Me.sstOrder.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sstOrder.ItemSize = New System.Drawing.Size(42, 18)
        Me.sstOrder.Location = New System.Drawing.Point(12, 43)
        Me.sstOrder.Name = "sstOrder"
        Me.sstOrder.SelectedIndex = 0
        Me.sstOrder.Size = New System.Drawing.Size(1000, 600)
        Me.sstOrder.TabIndex = 1
        '
        '_sstOrder_T00
        '
        Me._sstOrder_T00.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._sstOrder_T00.Controls.Add(Me.testSimulateThermo)
        Me._sstOrder_T00.Controls.Add(Me.cmdETA_Calculator)
        Me._sstOrder_T00.Controls.Add(Me.cmdReset_Cus_No)
        Me._sstOrder_T00.Controls.Add(Me.cmdEDI_Alert_New_Inv)
        Me._sstOrder_T00.Controls.Add(Me.Button4)
        Me._sstOrder_T00.Controls.Add(Me.Button3)
        Me._sstOrder_T00.Controls.Add(Me.cmdMySqlTest)
        Me._sstOrder_T00.Controls.Add(Me.btnThermo)
        Me._sstOrder_T00.Controls.Add(Me.btnFileList)
        Me._sstOrder_T00.Controls.Add(Me.btnEDI)
        Me._sstOrder_T00.Controls.Add(Me.cmdOEI_Config)
        Me._sstOrder_T00.Controls.Add(Me.cmdDirServices)
        Me._sstOrder_T00.Controls.Add(Me.btnODBCConnect)
        Me._sstOrder_T00.Controls.Add(Me.cmdEmail)
        Me._sstOrder_T00.Controls.Add(Me.Button2)
        Me._sstOrder_T00.Controls.Add(Me.UcOrder)
        Me._sstOrder_T00.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T00.Name = "_sstOrder_T00"
        Me._sstOrder_T00.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T00.TabIndex = 8
        Me._sstOrder_T00.Text = "1-Header"
        Me._sstOrder_T00.UseVisualStyleBackColor = True
        '
        'testSimulateThermo
        '
        Me.testSimulateThermo.Location = New System.Drawing.Point(453, 323)
        Me.testSimulateThermo.Name = "testSimulateThermo"
        Me.testSimulateThermo.Size = New System.Drawing.Size(84, 24)
        Me.testSimulateThermo.TabIndex = 15
        Me.testSimulateThermo.Text = "testThermo"
        Me.testSimulateThermo.UseVisualStyleBackColor = True
        Me.testSimulateThermo.Visible = False
        '
        'cmdETA_Calculator
        '
        Me.cmdETA_Calculator.Location = New System.Drawing.Point(10, 396)
        Me.cmdETA_Calculator.Name = "cmdETA_Calculator"
        Me.cmdETA_Calculator.Size = New System.Drawing.Size(106, 25)
        Me.cmdETA_Calculator.TabIndex = 14
        Me.cmdETA_Calculator.Text = "ETA Calculator"
        Me.cmdETA_Calculator.UseVisualStyleBackColor = True
        Me.cmdETA_Calculator.Visible = False
        '
        'cmdReset_Cus_No
        '
        Me.cmdReset_Cus_No.Location = New System.Drawing.Point(11, 365)
        Me.cmdReset_Cus_No.Name = "cmdReset_Cus_No"
        Me.cmdReset_Cus_No.Size = New System.Drawing.Size(106, 25)
        Me.cmdReset_Cus_No.TabIndex = 13
        Me.cmdReset_Cus_No.Text = "Reset_Cus_No"
        Me.cmdReset_Cus_No.UseVisualStyleBackColor = True
        Me.cmdReset_Cus_No.Visible = False
        '
        'cmdEDI_Alert_New_Inv
        '
        Me.cmdEDI_Alert_New_Inv.Location = New System.Drawing.Point(11, 323)
        Me.cmdEDI_Alert_New_Inv.Name = "cmdEDI_Alert_New_Inv"
        Me.cmdEDI_Alert_New_Inv.Size = New System.Drawing.Size(108, 24)
        Me.cmdEDI_Alert_New_Inv.TabIndex = 12
        Me.cmdEDI_Alert_New_Inv.Text = "EDI Alert New Inv"
        Me.cmdEDI_Alert_New_Inv.UseVisualStyleBackColor = True
        Me.cmdEDI_Alert_New_Inv.Visible = False
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(8, 308)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 11
        Me.Button4.Text = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        Me.Button4.Visible = False
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(8, 278)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(111, 23)
        Me.Button3.TabIndex = 10
        Me.Button3.Text = "HV Service Test"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'cmdMySqlTest
        '
        Me.cmdMySqlTest.Location = New System.Drawing.Point(8, 249)
        Me.cmdMySqlTest.Name = "cmdMySqlTest"
        Me.cmdMySqlTest.Size = New System.Drawing.Size(75, 23)
        Me.cmdMySqlTest.TabIndex = 9
        Me.cmdMySqlTest.Text = "MySql"
        Me.cmdMySqlTest.UseVisualStyleBackColor = True
        Me.cmdMySqlTest.Visible = False
        '
        'btnThermo
        '
        Me.btnThermo.Location = New System.Drawing.Point(8, 220)
        Me.btnThermo.Name = "btnThermo"
        Me.btnThermo.Size = New System.Drawing.Size(75, 23)
        Me.btnThermo.TabIndex = 8
        Me.btnThermo.Text = "Thermo"
        Me.btnThermo.UseVisualStyleBackColor = True
        Me.btnThermo.Visible = False
        '
        'btnFileList
        '
        Me.btnFileList.Location = New System.Drawing.Point(8, 191)
        Me.btnFileList.Name = "btnFileList"
        Me.btnFileList.Size = New System.Drawing.Size(75, 23)
        Me.btnFileList.TabIndex = 7
        Me.btnFileList.Text = "FileList"
        Me.btnFileList.UseVisualStyleBackColor = True
        Me.btnFileList.Visible = False
        '
        'btnEDI
        '
        Me.btnEDI.Location = New System.Drawing.Point(8, 162)
        Me.btnEDI.Name = "btnEDI"
        Me.btnEDI.Size = New System.Drawing.Size(75, 23)
        Me.btnEDI.TabIndex = 6
        Me.btnEDI.Text = "EDI"
        Me.btnEDI.UseVisualStyleBackColor = True
        Me.btnEDI.Visible = False
        '
        'cmdOEI_Config
        '
        Me.cmdOEI_Config.Location = New System.Drawing.Point(8, 133)
        Me.cmdOEI_Config.Name = "cmdOEI_Config"
        Me.cmdOEI_Config.Size = New System.Drawing.Size(75, 23)
        Me.cmdOEI_Config.TabIndex = 5
        Me.cmdOEI_Config.Text = "OEI Config"
        Me.cmdOEI_Config.UseVisualStyleBackColor = True
        Me.cmdOEI_Config.Visible = False
        '
        'cmdDirServices
        '
        Me.cmdDirServices.Location = New System.Drawing.Point(8, 103)
        Me.cmdDirServices.Name = "cmdDirServices"
        Me.cmdDirServices.Size = New System.Drawing.Size(65, 24)
        Me.cmdDirServices.TabIndex = 4
        Me.cmdDirServices.Text = "Directory"
        Me.cmdDirServices.UseVisualStyleBackColor = True
        Me.cmdDirServices.Visible = False
        '
        'btnODBCConnect
        '
        Me.btnODBCConnect.Location = New System.Drawing.Point(8, 72)
        Me.btnODBCConnect.Name = "btnODBCConnect"
        Me.btnODBCConnect.Size = New System.Drawing.Size(65, 25)
        Me.btnODBCConnect.TabIndex = 3
        Me.btnODBCConnect.Text = "ODBC"
        Me.btnODBCConnect.UseVisualStyleBackColor = True
        Me.btnODBCConnect.Visible = False
        '
        'cmdEmail
        '
        Me.cmdEmail.Location = New System.Drawing.Point(8, 41)
        Me.cmdEmail.Name = "cmdEmail"
        Me.cmdEmail.Size = New System.Drawing.Size(65, 25)
        Me.cmdEmail.TabIndex = 2
        Me.cmdEmail.Text = "Email"
        Me.cmdEmail.UseVisualStyleBackColor = True
        Me.cmdEmail.Visible = False
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(8, 10)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(65, 24)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'UcOrder
        '
        Me.UcOrder.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange
        Me.UcOrder.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.UcOrder.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UcOrder.Location = New System.Drawing.Point(8, 10)
        Me.UcOrder.Name = "UcOrder"
        Me.UcOrder.Size = New System.Drawing.Size(980, 555)
        Me.UcOrder.TabIndex = 0
        '
        '_sstOrder_T01
        '
        Me._sstOrder_T01.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._sstOrder_T01.Controls.Add(Me.UcContacts1)
        Me._sstOrder_T01.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T01.Name = "_sstOrder_T01"
        Me._sstOrder_T01.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T01.TabIndex = 9
        Me._sstOrder_T01.Text = "2-Customer Contacts"
        Me._sstOrder_T01.UseVisualStyleBackColor = True
        '
        'UcContacts1
        '
        Me.UcContacts1.BackColor = System.Drawing.SystemColors.Control
        Me.UcContacts1.Location = New System.Drawing.Point(8, 10)
        Me.UcContacts1.Name = "UcContacts1"
        Me.UcContacts1.Size = New System.Drawing.Size(980, 555)
        Me.UcContacts1.TabIndex = 0
        '
        '_sstOrder_T02
        '
        Me._sstOrder_T02.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._sstOrder_T02.Controls.Add(Me.cmdEDIEditSuppItem)
        Me._sstOrder_T02.Controls.Add(Me.cmdDash)
        Me._sstOrder_T02.Controls.Add(Me.cmdTraveler)
        Me._sstOrder_T02.Controls.Add(Me.cmdPrice)
        Me._sstOrder_T02.Controls.Add(Me.chkAutoCompleteReship)
        Me._sstOrder_T02.Controls.Add(Me.cmdQty)
        Me._sstOrder_T02.Controls.Add(Me.cmdRequest_Dt)
        Me._sstOrder_T02.Controls.Add(Me.cmdImprint)
        Me._sstOrder_T02.Controls.Add(Me.DateTimePicker1)
        Me._sstOrder_T02.Controls.Add(Me.DataGridView1)
        Me._sstOrder_T02.Controls.Add(Me.Button1)
        Me._sstOrder_T02.Controls.Add(Me.cmdPrevCharges)
        Me._sstOrder_T02.Controls.Add(Me.cboMoreOptions)
        Me._sstOrder_T02.Controls.Add(Me.chkLowRow)
        Me._sstOrder_T02.Controls.Add(Me.panImprint)
        Me._sstOrder_T02.Controls.Add(Me.dgvItems)
        Me._sstOrder_T02.Controls.Add(Me.cmdNotes)
        Me._sstOrder_T02.Controls.Add(Me.cmdSearch)
        Me._sstOrder_T02.Controls.Add(Me.cmdRestart)
        Me._sstOrder_T02.Controls.Add(Me.cmdCopyLine)
        Me._sstOrder_T02.Controls.Add(Me.cmdDeleteLine)
        Me._sstOrder_T02.Controls.Add(Me.cmdInsertLine)
        Me._sstOrder_T02.Controls.Add(Me.XTextBox2)
        Me._sstOrder_T02.Controls.Add(Me.UcOrderTotal1)
        Me._sstOrder_T02.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T02.Name = "_sstOrder_T02"
        Me._sstOrder_T02.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T02.TabIndex = 2
        Me._sstOrder_T02.Text = "3-Lines"
        Me._sstOrder_T02.UseVisualStyleBackColor = True
        '
        'cmdEDIEditSuppItem
        '
        Me.cmdEDIEditSuppItem.Location = New System.Drawing.Point(297, 9)
        Me.cmdEDIEditSuppItem.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdEDIEditSuppItem.Name = "cmdEDIEditSuppItem"
        Me.cmdEDIEditSuppItem.Size = New System.Drawing.Size(65, 27)
        Me.cmdEDIEditSuppItem.TabIndex = 149
        Me.cmdEDIEditSuppItem.Text = "EDI "
        Me.cmdEDIEditSuppItem.UseVisualStyleBackColor = True
        '
        'cmdDash
        '
        Me.cmdDash.Location = New System.Drawing.Point(808, 54)
        Me.cmdDash.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdDash.Name = "cmdDash"
        Me.cmdDash.Size = New System.Drawing.Size(65, 27)
        Me.cmdDash.TabIndex = 17
        Me.cmdDash.Text = "Dash"
        Me.cmdDash.UseVisualStyleBackColor = True
        Me.cmdDash.Visible = False
        '
        'cmdTraveler
        '
        Me.cmdTraveler.Location = New System.Drawing.Point(853, 9)
        Me.cmdTraveler.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdTraveler.Name = "cmdTraveler"
        Me.cmdTraveler.Size = New System.Drawing.Size(65, 27)
        Me.cmdTraveler.TabIndex = 11
        Me.cmdTraveler.Text = "Traveler"
        Me.cmdTraveler.UseVisualStyleBackColor = True
        '
        'cmdPrice
        '
        Me.cmdPrice.Location = New System.Drawing.Point(642, 9)
        Me.cmdPrice.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdPrice.Name = "cmdPrice"
        Me.cmdPrice.Size = New System.Drawing.Size(70, 27)
        Me.cmdPrice.TabIndex = 8
        Me.cmdPrice.Text = "Unit Price"
        Me.cmdPrice.UseVisualStyleBackColor = True
        '
        'chkAutoCompleteReship
        '
        Me.chkAutoCompleteReship.AutoSize = True
        Me.chkAutoCompleteReship.Location = New System.Drawing.Point(422, 14)
        Me.chkAutoCompleteReship.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.chkAutoCompleteReship.Name = "chkAutoCompleteReship"
        Me.chkAutoCompleteReship.Size = New System.Drawing.Size(144, 19)
        Me.chkAutoCompleteReship.TabIndex = 5
        Me.chkAutoCompleteReship.Text = "Autocomplete Reship"
        Me.chkAutoCompleteReship.UseVisualStyleBackColor = True
        '
        'cmdQty
        '
        Me.cmdQty.Location = New System.Drawing.Point(570, 9)
        Me.cmdQty.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdQty.Name = "cmdQty"
        Me.cmdQty.Size = New System.Drawing.Size(70, 27)
        Me.cmdQty.TabIndex = 7
        Me.cmdQty.Text = "Qty Order"
        Me.cmdQty.UseVisualStyleBackColor = True
        '
        'cmdRequest_Dt
        '
        Me.cmdRequest_Dt.Location = New System.Drawing.Point(714, 9)
        Me.cmdRequest_Dt.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdRequest_Dt.Name = "cmdRequest_Dt"
        Me.cmdRequest_Dt.Size = New System.Drawing.Size(70, 27)
        Me.cmdRequest_Dt.TabIndex = 9
        Me.cmdRequest_Dt.Text = "Req Date"
        Me.cmdRequest_Dt.UseVisualStyleBackColor = True
        '
        'cmdImprint
        '
        Me.cmdImprint.Location = New System.Drawing.Point(786, 9)
        Me.cmdImprint.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdImprint.Name = "cmdImprint"
        Me.cmdImprint.Size = New System.Drawing.Size(65, 27)
        Me.cmdImprint.TabIndex = 10
        Me.cmdImprint.Text = "Imprint"
        Me.cmdImprint.UseVisualStyleBackColor = True
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(383, 55)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(90, 21)
        Me.DateTimePicker1.TabIndex = 13
        Me.DateTimePicker1.Visible = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke
        Me.DataGridView1.Location = New System.Drawing.Point(449, 12)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 20
        DataGridViewCellStyle1.NullValue = Nothing
        Me.DataGridView1.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.RowTemplate.Height = 18
        Me.DataGridView1.Size = New System.Drawing.Size(21, 24)
        Me.DataGridView1.TabIndex = 6
        Me.DataGridView1.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(721, 55)
        Me.Button1.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(85, 23)
        Me.Button1.TabIndex = 16
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'cmdPrevCharges
        '
        Me.cmdPrevCharges.Location = New System.Drawing.Point(920, 9)
        Me.cmdPrevCharges.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdPrevCharges.Name = "cmdPrevCharges"
        Me.cmdPrevCharges.Size = New System.Drawing.Size(65, 27)
        Me.cmdPrevCharges.TabIndex = 12
        Me.cmdPrevCharges.Text = "Charges"
        Me.cmdPrevCharges.UseVisualStyleBackColor = True
        '
        'cboMoreOptions
        '
        Me.cboMoreOptions.Font = New System.Drawing.Font("Arial", 10.5!)
        Me.cboMoreOptions.FormattingEnabled = True
        Me.cboMoreOptions.Items.AddRange(New Object() {"Import items from order", "Search item"})
        Me.cboMoreOptions.Location = New System.Drawing.Point(479, 56)
        Me.cboMoreOptions.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cboMoreOptions.Name = "cboMoreOptions"
        Me.cboMoreOptions.Size = New System.Drawing.Size(93, 24)
        Me.cboMoreOptions.TabIndex = 14
        Me.cboMoreOptions.Text = "More options..."
        Me.cboMoreOptions.Visible = False
        '
        'chkLowRow
        '
        Me.chkLowRow.AutoSize = True
        Me.chkLowRow.Location = New System.Drawing.Point(217, 14)
        Me.chkLowRow.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.chkLowRow.Name = "chkLowRow"
        Me.chkLowRow.Size = New System.Drawing.Size(77, 19)
        Me.chkLowRow.TabIndex = 4
        Me.chkLowRow.Text = "Low Row"
        Me.chkLowRow.UseVisualStyleBackColor = True
        '
        'panImprint
        '
        Me.panImprint.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.panImprint.Controls.Add(Me.btnDbgPrxAdjust)
        Me.panImprint.Controls.Add(Me.lblEnd_User)
        Me.panImprint.Controls.Add(Me.txtEnd_User)
        Me.panImprint.Controls.Add(Me.txtNum_Impr_3)
        Me.panImprint.Controls.Add(Me.lblNum_Impr_3)
        Me.panImprint.Controls.Add(Me.txtNum_Impr_1)
        Me.panImprint.Controls.Add(Me.txtNum_Impr_2)
        Me.panImprint.Controls.Add(Me.txtNum_Impr_12)
        Me.panImprint.Controls.Add(Me.lblNum_Impr_2)
        Me.panImprint.Controls.Add(Me.txtRepeat_From_ID)
        Me.panImprint.Controls.Add(Me.cboIndustry)
        Me.panImprint.Controls.Add(Me.lblIndustry)
        Me.panImprint.Controls.Add(Me.txtIndustry)
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
        Me.panImprint.Controls.Add(Me.cmdCopyDown)
        Me.panImprint.Controls.Add(Me.txtImprint_Item_No)
        Me.panImprint.Controls.Add(Me.lblImprint_Item_No)
        Me.panImprint.Controls.Add(Me.UcImprint1)
        Me.panImprint.Controls.Add(Me.cmdPaste)
        Me.panImprint.Controls.Add(Me.cmdUpdateAllItem)
        Me.panImprint.Controls.Add(Me.cmdUpdateAll)
        Me.panImprint.Controls.Add(Me.gbRepeatData)
        Me.panImprint.Controls.Add(Me.cmdClear_Imprint)
        Me.panImprint.Location = New System.Drawing.Point(8, 341)
        Me.panImprint.Name = "panImprint"
        Me.panImprint.Size = New System.Drawing.Size(977, 145)
        Me.panImprint.TabIndex = 19
        '
        'btnDbgPrxAdjust
        '
        Me.btnDbgPrxAdjust.Location = New System.Drawing.Point(880, 1)
        Me.btnDbgPrxAdjust.Name = "btnDbgPrxAdjust"
        Me.btnDbgPrxAdjust.Size = New System.Drawing.Size(87, 21)
        Me.btnDbgPrxAdjust.TabIndex = 150
        Me.btnDbgPrxAdjust.Text = "Price Adjust"
        Me.btnDbgPrxAdjust.UseVisualStyleBackColor = True
        '
        'lblEnd_User
        '
        Me.lblEnd_User.AutoSize = True
        Me.lblEnd_User.Location = New System.Drawing.Point(588, 7)
        Me.lblEnd_User.Name = "lblEnd_User"
        Me.lblEnd_User.Size = New System.Drawing.Size(59, 15)
        Me.lblEnd_User.TabIndex = 166
        Me.lblEnd_User.Text = "End User"
        Me.lblEnd_User.Visible = False
        '
        'txtEnd_User
        '
        Me.txtEnd_User.Location = New System.Drawing.Point(626, 4)
        Me.txtEnd_User.Name = "txtEnd_User"
        Me.txtEnd_User.Size = New System.Drawing.Size(39, 21)
        Me.txtEnd_User.TabIndex = 167
        Me.txtEnd_User.Visible = False
        '
        'txtNum_Impr_3
        '
        Me.txtNum_Impr_3.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtNum_Impr_3.DataLength = CType(0, Long)
        Me.txtNum_Impr_3.DataType = Orders.CDataType.dtNumWithoutDecimals
        Me.txtNum_Impr_3.DateValue = New Date(CType(0, Long))
        Me.txtNum_Impr_3.DecimalDigits = CType(0, Long)
        Me.txtNum_Impr_3.Location = New System.Drawing.Point(451, 25)
        Me.txtNum_Impr_3.Mask = Nothing
        Me.txtNum_Impr_3.Name = "txtNum_Impr_3"
        Me.txtNum_Impr_3.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtNum_Impr_3.OldValue = ""
        Me.txtNum_Impr_3.Size = New System.Drawing.Size(60, 21)
        Me.txtNum_Impr_3.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtNum_Impr_3.StringValue = Nothing
        Me.txtNum_Impr_3.TabIndex = 165
        Me.txtNum_Impr_3.TextDataID = Nothing
        '
        'lblNum_Impr_3
        '
        Me.lblNum_Impr_3.AutoSize = True
        Me.lblNum_Impr_3.Location = New System.Drawing.Point(448, 6)
        Me.lblNum_Impr_3.Name = "lblNum_Impr_3"
        Me.lblNum_Impr_3.Size = New System.Drawing.Size(68, 15)
        Me.lblNum_Impr_3.TabIndex = 164
        Me.lblNum_Impr_3.Text = "Num Imp 3"
        '
        'txtNum_Impr_1
        '
        Me.txtNum_Impr_1.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtNum_Impr_1.DataLength = CType(0, Long)
        Me.txtNum_Impr_1.DataType = Orders.CDataType.dtNumWithoutDecimals
        Me.txtNum_Impr_1.DateValue = New Date(CType(0, Long))
        Me.txtNum_Impr_1.DecimalDigits = CType(0, Long)
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
        Me.txtNum_Impr_12.Location = New System.Drawing.Point(390, 25)
        Me.txtNum_Impr_12.Name = "txtNum_Impr_12"
        Me.txtNum_Impr_12.Size = New System.Drawing.Size(55, 21)
        Me.txtNum_Impr_12.TabIndex = 163
        Me.txtNum_Impr_12.Visible = False
        '
        'lblNum_Impr_2
        '
        Me.lblNum_Impr_2.AutoSize = True
        Me.lblNum_Impr_2.Location = New System.Drawing.Point(382, 6)
        Me.lblNum_Impr_2.Name = "lblNum_Impr_2"
        Me.lblNum_Impr_2.Size = New System.Drawing.Size(68, 15)
        Me.lblNum_Impr_2.TabIndex = 6
        Me.lblNum_Impr_2.Text = "Num Imp 2"
        '
        'txtRepeat_From_ID
        '
        Me.txtRepeat_From_ID.Location = New System.Drawing.Point(820, 115)
        Me.txtRepeat_From_ID.Name = "txtRepeat_From_ID"
        Me.txtRepeat_From_ID.Size = New System.Drawing.Size(37, 21)
        Me.txtRepeat_From_ID.TabIndex = 162
        Me.txtRepeat_From_ID.Visible = False
        '
        'cboIndustry
        '
        Me.cboIndustry.FormattingEnabled = True
        Me.cboIndustry.Location = New System.Drawing.Point(847, 81)
        Me.cboIndustry.Name = "cboIndustry"
        Me.cboIndustry.Size = New System.Drawing.Size(123, 23)
        Me.cboIndustry.TabIndex = 25
        '
        'lblIndustry
        '
        Me.lblIndustry.AutoSize = True
        Me.lblIndustry.Location = New System.Drawing.Point(847, 57)
        Me.lblIndustry.Name = "lblIndustry"
        Me.lblIndustry.Size = New System.Drawing.Size(50, 15)
        Me.lblIndustry.TabIndex = 24
        Me.lblIndustry.Text = "Industry"
        '
        'txtIndustry
        '
        Me.txtIndustry.Location = New System.Drawing.Point(863, 83)
        Me.txtIndustry.Name = "txtIndustry"
        Me.txtIndustry.Size = New System.Drawing.Size(104, 21)
        Me.txtIndustry.TabIndex = 160
        Me.txtIndustry.Visible = False
        '
        'lblRefill
        '
        Me.lblRefill.AutoSize = True
        Me.lblRefill.Location = New System.Drawing.Point(668, 57)
        Me.lblRefill.Name = "lblRefill"
        Me.lblRefill.Size = New System.Drawing.Size(35, 15)
        Me.lblRefill.TabIndex = 21
        Me.lblRefill.Text = "Refill"
        '
        'txtRefill
        '
        Me.txtRefill.Location = New System.Drawing.Point(671, 83)
        Me.txtRefill.Name = "txtRefill"
        Me.txtRefill.Size = New System.Drawing.Size(170, 21)
        Me.txtRefill.TabIndex = 23
        '
        'cboRefill
        '
        Me.cboRefill.FormattingEnabled = True
        Me.cboRefill.Location = New System.Drawing.Point(709, 54)
        Me.cboRefill.Name = "cboRefill"
        Me.cboRefill.Size = New System.Drawing.Size(132, 23)
        Me.cboRefill.TabIndex = 22
        '
        'lblPackaging
        '
        Me.lblPackaging.AutoSize = True
        Me.lblPackaging.Location = New System.Drawing.Point(494, 57)
        Me.lblPackaging.Name = "lblPackaging"
        Me.lblPackaging.Size = New System.Drawing.Size(37, 15)
        Me.lblPackaging.TabIndex = 18
        Me.lblPackaging.Text = "Pack."
        '
        'txtPackaging
        '
        Me.txtPackaging.Location = New System.Drawing.Point(495, 83)
        Me.txtPackaging.Name = "txtPackaging"
        Me.txtPackaging.Size = New System.Drawing.Size(170, 21)
        Me.txtPackaging.TabIndex = 20
        '
        'cboPackaging
        '
        Me.cboPackaging.FormattingEnabled = True
        Me.cboPackaging.Location = New System.Drawing.Point(537, 54)
        Me.cboPackaging.Name = "cboPackaging"
        Me.cboPackaging.Size = New System.Drawing.Size(128, 23)
        Me.cboPackaging.TabIndex = 19
        '
        'txtImprint_Color
        '
        Me.txtImprint_Color.Location = New System.Drawing.Point(319, 83)
        Me.txtImprint_Color.Name = "txtImprint_Color"
        Me.txtImprint_Color.Size = New System.Drawing.Size(170, 21)
        Me.txtImprint_Color.TabIndex = 17
        '
        'lblImprint_Color
        '
        Me.lblImprint_Color.AutoSize = True
        Me.lblImprint_Color.Location = New System.Drawing.Point(315, 57)
        Me.lblImprint_Color.Name = "lblImprint_Color"
        Me.lblImprint_Color.Size = New System.Drawing.Size(37, 15)
        Me.lblImprint_Color.TabIndex = 15
        Me.lblImprint_Color.Text = "Color"
        '
        'cboImprint_Color
        '
        Me.cboImprint_Color.FormattingEnabled = True
        Me.cboImprint_Color.Location = New System.Drawing.Point(358, 54)
        Me.cboImprint_Color.Name = "cboImprint_Color"
        Me.cboImprint_Color.Size = New System.Drawing.Size(130, 23)
        Me.cboImprint_Color.TabIndex = 16
        '
        'lblImprint_Location
        '
        Me.lblImprint_Location.AutoSize = True
        Me.lblImprint_Location.Location = New System.Drawing.Point(143, 57)
        Me.lblImprint_Location.Name = "lblImprint_Location"
        Me.lblImprint_Location.Size = New System.Drawing.Size(30, 15)
        Me.lblImprint_Location.TabIndex = 12
        Me.lblImprint_Location.Text = "Loc."
        '
        'txtImprint_Location
        '
        Me.txtImprint_Location.Location = New System.Drawing.Point(143, 83)
        Me.txtImprint_Location.Name = "txtImprint_Location"
        Me.txtImprint_Location.Size = New System.Drawing.Size(170, 21)
        Me.txtImprint_Location.TabIndex = 14
        '
        'lblSpecialComments
        '
        Me.lblSpecialComments.AutoSize = True
        Me.lblSpecialComments.Location = New System.Drawing.Point(744, 6)
        Me.lblSpecialComments.Name = "lblSpecialComments"
        Me.lblSpecialComments.Size = New System.Drawing.Size(110, 15)
        Me.lblSpecialComments.TabIndex = 10
        Me.lblSpecialComments.Text = "Special comments"
        '
        'txtSpecial_Comments
        '
        Me.txtSpecial_Comments.Location = New System.Drawing.Point(747, 24)
        Me.txtSpecial_Comments.Name = "txtSpecial_Comments"
        Me.txtSpecial_Comments.Size = New System.Drawing.Size(224, 21)
        Me.txtSpecial_Comments.TabIndex = 11
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(517, 24)
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(224, 21)
        Me.txtComments.TabIndex = 9
        '
        'lblComments
        '
        Me.lblComments.AutoSize = True
        Me.lblComments.Location = New System.Drawing.Point(514, 7)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(69, 15)
        Me.lblComments.TabIndex = 8
        Me.lblComments.Text = "Comments"
        '
        'txtNum_Impr_11
        '
        Me.txtNum_Impr_11.Location = New System.Drawing.Point(319, 25)
        Me.txtNum_Impr_11.Name = "txtNum_Impr_11"
        Me.txtNum_Impr_11.Size = New System.Drawing.Size(60, 21)
        Me.txtNum_Impr_11.TabIndex = 3
        Me.txtNum_Impr_11.Visible = False
        '
        'lblNum_Impr_1
        '
        Me.lblNum_Impr_1.AutoSize = True
        Me.lblNum_Impr_1.Location = New System.Drawing.Point(315, 6)
        Me.lblNum_Impr_1.Name = "lblNum_Impr_1"
        Me.lblNum_Impr_1.Size = New System.Drawing.Size(68, 15)
        Me.lblNum_Impr_1.TabIndex = 4
        Me.lblNum_Impr_1.Text = "Num Imp 1"
        '
        'cboImprint_Location
        '
        Me.cboImprint_Location.FormattingEnabled = True
        Me.cboImprint_Location.Location = New System.Drawing.Point(179, 53)
        Me.cboImprint_Location.Name = "cboImprint_Location"
        Me.cboImprint_Location.Size = New System.Drawing.Size(134, 23)
        Me.cboImprint_Location.TabIndex = 13
        '
        'txtImprint
        '
        Me.txtImprint.Location = New System.Drawing.Point(143, 25)
        Me.txtImprint.Name = "txtImprint"
        Me.txtImprint.Size = New System.Drawing.Size(170, 21)
        Me.txtImprint.TabIndex = 3
        '
        'lblImprint
        '
        Me.lblImprint.AutoSize = True
        Me.lblImprint.Location = New System.Drawing.Point(143, 6)
        Me.lblImprint.Name = "lblImprint"
        Me.lblImprint.Size = New System.Drawing.Size(45, 15)
        Me.lblImprint.TabIndex = 2
        Me.lblImprint.Text = "Imprint"
        '
        'cmdCopyDown
        '
        Me.cmdCopyDown.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopyDown.Location = New System.Drawing.Point(568, 113)
        Me.cmdCopyDown.Name = "cmdCopyDown"
        Me.cmdCopyDown.Size = New System.Drawing.Size(120, 25)
        Me.cmdCopyDown.TabIndex = 29
        Me.cmdCopyDown.Text = "Copy"
        Me.cmdCopyDown.UseVisualStyleBackColor = True
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
        Me.UcImprint1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UcImprint1.Location = New System.Drawing.Point(863, 96)
        Me.UcImprint1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UcImprint1.Name = "UcImprint1"
        Me.UcImprint1.Size = New System.Drawing.Size(57, 46)
        Me.UcImprint1.TabIndex = 135
        Me.UcImprint1.Visible = False
        '
        'cmdPaste
        '
        Me.cmdPaste.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPaste.Location = New System.Drawing.Point(694, 113)
        Me.cmdPaste.Name = "cmdPaste"
        Me.cmdPaste.Size = New System.Drawing.Size(120, 25)
        Me.cmdPaste.TabIndex = 30
        Me.cmdPaste.Text = "Paste"
        Me.cmdPaste.UseVisualStyleBackColor = True
        '
        'cmdUpdateAllItem
        '
        Me.cmdUpdateAllItem.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateAllItem.Location = New System.Drawing.Point(395, 113)
        Me.cmdUpdateAllItem.Name = "cmdUpdateAllItem"
        Me.cmdUpdateAllItem.Size = New System.Drawing.Size(167, 25)
        Me.cmdUpdateAllItem.TabIndex = 28
        Me.cmdUpdateAllItem.Text = "Update All 00EC1040"
        Me.cmdUpdateAllItem.UseVisualStyleBackColor = True
        '
        'cmdUpdateAll
        '
        Me.cmdUpdateAll.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateAll.Location = New System.Drawing.Point(269, 113)
        Me.cmdUpdateAll.Name = "cmdUpdateAll"
        Me.cmdUpdateAll.Size = New System.Drawing.Size(120, 25)
        Me.cmdUpdateAll.TabIndex = 27
        Me.cmdUpdateAll.Text = "Update All"
        Me.cmdUpdateAll.UseVisualStyleBackColor = True
        '
        'gbRepeatData
        '
        Me.gbRepeatData.Controls.Add(Me.cmdRepeatSearch)
        Me.gbRepeatData.Controls.Add(Me.cmdGetRepeatData)
        Me.gbRepeatData.Controls.Add(Me.txtRepeat_Ord_No)
        Me.gbRepeatData.Controls.Add(Me.lblRepeat_Ord_No)
        Me.gbRepeatData.Location = New System.Drawing.Point(3, 48)
        Me.gbRepeatData.Name = "gbRepeatData"
        Me.gbRepeatData.Size = New System.Drawing.Size(133, 90)
        Me.gbRepeatData.TabIndex = 31
        Me.gbRepeatData.TabStop = False
        '
        'cmdRepeatSearch
        '
        Me.cmdRepeatSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRepeatSearch.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRepeatSearch.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdRepeatSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRepeatSearch.Image = CType(resources.GetObject("cmdRepeatSearch.Image"), System.Drawing.Image)
        Me.cmdRepeatSearch.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdRepeatSearch.Location = New System.Drawing.Point(102, 30)
        Me.cmdRepeatSearch.Name = "cmdRepeatSearch"
        Me.cmdRepeatSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRepeatSearch.Size = New System.Drawing.Size(24, 24)
        Me.cmdRepeatSearch.TabIndex = 2
        Me.cmdRepeatSearch.TabStop = False
        Me.cmdRepeatSearch.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdRepeatSearch.UseVisualStyleBackColor = False
        '
        'cmdGetRepeatData
        '
        Me.cmdGetRepeatData.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetRepeatData.Location = New System.Drawing.Point(6, 59)
        Me.cmdGetRepeatData.Name = "cmdGetRepeatData"
        Me.cmdGetRepeatData.Size = New System.Drawing.Size(120, 25)
        Me.cmdGetRepeatData.TabIndex = 3
        Me.cmdGetRepeatData.Text = "Get Repeat Data"
        Me.cmdGetRepeatData.UseVisualStyleBackColor = True
        '
        'txtRepeat_Ord_No
        '
        Me.txtRepeat_Ord_No.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRepeat_Ord_No.Location = New System.Drawing.Point(6, 31)
        Me.txtRepeat_Ord_No.Name = "txtRepeat_Ord_No"
        Me.txtRepeat_Ord_No.Size = New System.Drawing.Size(96, 22)
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
        Me.cmdClear_Imprint.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClear_Imprint.Location = New System.Drawing.Point(143, 113)
        Me.cmdClear_Imprint.Name = "cmdClear_Imprint"
        Me.cmdClear_Imprint.Size = New System.Drawing.Size(120, 25)
        Me.cmdClear_Imprint.TabIndex = 26
        Me.cmdClear_Imprint.Text = "Clear"
        Me.cmdClear_Imprint.UseVisualStyleBackColor = True
        '
        'dgvItems
        '
        Me.dgvItems.AllowUserToAddRows = False
        Me.dgvItems.AllowUserToDeleteRows = False
        Me.dgvItems.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.dgvItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItems.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke
        Me.dgvItems.Location = New System.Drawing.Point(8, 43)
        Me.dgvItems.MultiSelect = False
        Me.dgvItems.Name = "dgvItems"
        Me.dgvItems.RowHeadersWidth = 20
        Me.dgvItems.Size = New System.Drawing.Size(977, 292)
        Me.dgvItems.TabIndex = 18
        '
        'cmdNotes
        '
        Me.cmdNotes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNotes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNotes.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdNotes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNotes.Image = CType(resources.GetObject("cmdNotes.Image"), System.Drawing.Image)
        Me.cmdNotes.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdNotes.Location = New System.Drawing.Point(990, 7)
        Me.cmdNotes.Name = "cmdNotes"
        Me.cmdNotes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNotes.Size = New System.Drawing.Size(0, 0)
        Me.cmdNotes.TabIndex = 148
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
        Me.cmdSearch.Location = New System.Drawing.Point(970, 7)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSearch.Size = New System.Drawing.Size(0, 0)
        Me.cmdSearch.TabIndex = 147
        Me.cmdSearch.TabStop = False
        Me.cmdSearch.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdSearch.UseVisualStyleBackColor = False
        '
        'cmdRestart
        '
        Me.cmdRestart.Location = New System.Drawing.Point(141, 9)
        Me.cmdRestart.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdRestart.Name = "cmdRestart"
        Me.cmdRestart.Size = New System.Drawing.Size(65, 27)
        Me.cmdRestart.TabIndex = 3
        Me.cmdRestart.Text = "Restart"
        Me.cmdRestart.UseVisualStyleBackColor = True
        '
        'cmdCopyLine
        '
        Me.cmdCopyLine.Enabled = False
        Me.cmdCopyLine.Location = New System.Drawing.Point(296, 9)
        Me.cmdCopyLine.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdCopyLine.Name = "cmdCopyLine"
        Me.cmdCopyLine.Size = New System.Drawing.Size(65, 27)
        Me.cmdCopyLine.TabIndex = 2
        Me.cmdCopyLine.Text = "Copy"
        Me.cmdCopyLine.UseVisualStyleBackColor = True
        Me.cmdCopyLine.Visible = False
        '
        'cmdDeleteLine
        '
        Me.cmdDeleteLine.Location = New System.Drawing.Point(74, 9)
        Me.cmdDeleteLine.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdDeleteLine.Name = "cmdDeleteLine"
        Me.cmdDeleteLine.Size = New System.Drawing.Size(65, 27)
        Me.cmdDeleteLine.TabIndex = 1
        Me.cmdDeleteLine.Text = "&Delete"
        Me.cmdDeleteLine.UseVisualStyleBackColor = True
        '
        'cmdInsertLine
        '
        Me.cmdInsertLine.Location = New System.Drawing.Point(7, 9)
        Me.cmdInsertLine.Margin = New System.Windows.Forms.Padding(1, 3, 1, 3)
        Me.cmdInsertLine.Name = "cmdInsertLine"
        Me.cmdInsertLine.Size = New System.Drawing.Size(65, 27)
        Me.cmdInsertLine.TabIndex = 0
        Me.cmdInsertLine.Text = "&Add Line"
        Me.cmdInsertLine.UseVisualStyleBackColor = True
        '
        'XTextBox2
        '
        Me.XTextBox2.CharacterInput = Orders.Cinput.NumericOnly
        Me.XTextBox2.DataLength = CType(0, Long)
        Me.XTextBox2.DataType = Orders.CDataType.dtString
        Me.XTextBox2.DateValue = New Date(CType(0, Long))
        Me.XTextBox2.DecimalDigits = CType(0, Long)
        Me.XTextBox2.Font = New System.Drawing.Font("Arial", 10.5!)
        Me.XTextBox2.Location = New System.Drawing.Point(578, 56)
        Me.XTextBox2.Mask = Nothing
        Me.XTextBox2.Name = "XTextBox2"
        Me.XTextBox2.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.XTextBox2.OldValue = Nothing
        Me.XTextBox2.Size = New System.Drawing.Size(96, 24)
        Me.XTextBox2.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.XTextBox2.StringValue = Nothing
        Me.XTextBox2.TabIndex = 15
        Me.XTextBox2.TextDataID = Nothing
        Me.XTextBox2.Visible = False
        '
        'UcOrderTotal1
        '
        Me.UcOrderTotal1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.UcOrderTotal1.Curr_Trx_Rt = 0R
        Me.UcOrderTotal1.From_Curr_Cd_Desc = "US Dollars"
        Me.UcOrderTotal1.Location = New System.Drawing.Point(8, 492)
        Me.UcOrderTotal1.Name = "UcOrderTotal1"
        Me.UcOrderTotal1.Order_Amt = 0R
        Me.UcOrderTotal1.Order_Item_Count = 0
        Me.UcOrderTotal1.Order_Qty = 0
        Me.UcOrderTotal1.Ship_Amt = 0R
        Me.UcOrderTotal1.Ship_Qty = 0
        Me.UcOrderTotal1.Size = New System.Drawing.Size(977, 72)
        Me.UcOrderTotal1.TabIndex = 20
        Me.UcOrderTotal1.To_Curr_Cd_Desc = "Canadian Dollars"
        '
        '_sstOrder_T03
        '
        Me._sstOrder_T03.Controls.Add(Me.UcDocument1)
        Me._sstOrder_T03.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T03.Name = "_sstOrder_T03"
        Me._sstOrder_T03.Padding = New System.Windows.Forms.Padding(3)
        Me._sstOrder_T03.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T03.TabIndex = 13
        Me._sstOrder_T03.Text = "4-Documents"
        Me._sstOrder_T03.UseVisualStyleBackColor = True
        '
        'UcDocument1
        '
        Me.UcDocument1.Cus_No = Nothing
        Me.UcDocument1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UcDocument1.Item_Guid = Nothing
        Me.UcDocument1.Location = New System.Drawing.Point(8, 10)
        Me.UcDocument1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UcDocument1.Name = "UcDocument1"
        Me.UcDocument1.Ord_Guid = Nothing
        Me.UcDocument1.Size = New System.Drawing.Size(980, 555)
        Me.UcDocument1.TabIndex = 1
        '
        '_sstOrder_T04
        '
        Me._sstOrder_T04.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._sstOrder_T04.Controls.Add(Me.UcHeader1)
        Me._sstOrder_T04.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T04.Name = "_sstOrder_T04"
        Me._sstOrder_T04.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T04.TabIndex = 1
        Me._sstOrder_T04.Text = "5-Header (All)"
        Me._sstOrder_T04.UseVisualStyleBackColor = True
        '
        'UcHeader1
        '
        Me.UcHeader1.BackColor = System.Drawing.SystemColors.Control
        Me.UcHeader1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UcHeader1.Location = New System.Drawing.Point(8, 10)
        Me.UcHeader1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.UcHeader1.Name = "UcHeader1"
        Me.UcHeader1.Size = New System.Drawing.Size(980, 555)
        Me.UcHeader1.TabIndex = 0
        '
        '_sstOrder_T05
        '
        Me._sstOrder_T05.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._sstOrder_T05.Controls.Add(Me.UcTaxes1)
        Me._sstOrder_T05.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T05.Name = "_sstOrder_T05"
        Me._sstOrder_T05.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T05.TabIndex = 3
        Me._sstOrder_T05.Text = "6-Taxes"
        Me._sstOrder_T05.UseVisualStyleBackColor = True
        '
        'UcTaxes1
        '
        Me.UcTaxes1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UcTaxes1.Location = New System.Drawing.Point(8, 10)
        Me.UcTaxes1.Name = "UcTaxes1"
        Me.UcTaxes1.Size = New System.Drawing.Size(617, 402)
        Me.UcTaxes1.TabIndex = 0
        '
        '_sstOrder_T06
        '
        Me._sstOrder_T06.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._sstOrder_T06.Controls.Add(Me.UcSalesperson1)
        Me._sstOrder_T06.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T06.Name = "_sstOrder_T06"
        Me._sstOrder_T06.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T06.TabIndex = 4
        Me._sstOrder_T06.Text = "7-Salesperson"
        Me._sstOrder_T06.UseVisualStyleBackColor = True
        '
        'UcSalesperson1
        '
        Me.UcSalesperson1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UcSalesperson1.Location = New System.Drawing.Point(8, 10)
        Me.UcSalesperson1.Name = "UcSalesperson1"
        Me.UcSalesperson1.Size = New System.Drawing.Size(712, 261)
        Me.UcSalesperson1.TabIndex = 0
        '
        '_sstOrder_T07
        '
        Me._sstOrder_T07.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._sstOrder_T07.Controls.Add(Me.UcCreditInfo1)
        Me._sstOrder_T07.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T07.Name = "_sstOrder_T07"
        Me._sstOrder_T07.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T07.TabIndex = 5
        Me._sstOrder_T07.Text = "8-Credit Info"
        Me._sstOrder_T07.UseVisualStyleBackColor = True
        '
        'UcCreditInfo1
        '
        Me.UcCreditInfo1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UcCreditInfo1.Location = New System.Drawing.Point(8, 10)
        Me.UcCreditInfo1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.UcCreditInfo1.Name = "UcCreditInfo1"
        Me.UcCreditInfo1.Size = New System.Drawing.Size(655, 484)
        Me.UcCreditInfo1.TabIndex = 0
        '
        '_sstOrder_T08
        '
        Me._sstOrder_T08.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._sstOrder_T08.Controls.Add(Me.cmdTest)
        Me._sstOrder_T08.Controls.Add(Me.UcExtra1)
        Me._sstOrder_T08.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T08.Name = "_sstOrder_T08"
        Me._sstOrder_T08.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T08.TabIndex = 6
        Me._sstOrder_T08.Text = "9-Extra"
        Me._sstOrder_T08.UseVisualStyleBackColor = True
        '
        'cmdTest
        '
        Me.cmdTest.Location = New System.Drawing.Point(590, 20)
        Me.cmdTest.Name = "cmdTest"
        Me.cmdTest.Size = New System.Drawing.Size(46, 23)
        Me.cmdTest.TabIndex = 1
        Me.cmdTest.Text = "Test"
        Me.cmdTest.UseVisualStyleBackColor = True
        Me.cmdTest.Visible = False
        '
        'UcExtra1
        '
        Me.UcExtra1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.UcExtra1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UcExtra1.Location = New System.Drawing.Point(8, 10)
        Me.UcExtra1.Name = "UcExtra1"
        Me.UcExtra1.Size = New System.Drawing.Size(637, 423)
        Me.UcExtra1.TabIndex = 0
        '
        '_sstOrder_T09
        '
        Me._sstOrder_T09.Controls.Add(Me.wbItem)
        Me._sstOrder_T09.Controls.Add(Me.dgvHistory)
        Me._sstOrder_T09.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T09.Name = "_sstOrder_T09"
        Me._sstOrder_T09.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T09.TabIndex = 11
        Me._sstOrder_T09.Text = "10-History"
        Me._sstOrder_T09.UseVisualStyleBackColor = True
        '
        'wbItem
        '
        Me.wbItem.CausesValidation = False
        Me.wbItem.Location = New System.Drawing.Point(17, 19)
        Me.wbItem.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbItem.Name = "wbItem"
        Me.wbItem.ScriptErrorsSuppressed = True
        Me.wbItem.Size = New System.Drawing.Size(61, 69)
        Me.wbItem.TabIndex = 153
        Me.wbItem.TabStop = False
        Me.wbItem.Visible = False
        '
        'dgvHistory
        '
        Me.dgvHistory.AllowUserToAddRows = False
        Me.dgvHistory.AllowUserToDeleteRows = False
        Me.dgvHistory.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.dgvHistory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvHistory.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke
        Me.dgvHistory.Location = New System.Drawing.Point(8, 9)
        Me.dgvHistory.MultiSelect = False
        Me.dgvHistory.Name = "dgvHistory"
        Me.dgvHistory.ReadOnly = True
        Me.dgvHistory.RowHeadersWidth = 20
        Me.dgvHistory.Size = New System.Drawing.Size(977, 551)
        Me.dgvHistory.TabIndex = 154
        '
        '_sstOrder_T10
        '
        Me._sstOrder_T10.Controls.Add(Me.dgvReservedItems)
        Me._sstOrder_T10.Location = New System.Drawing.Point(4, 22)
        Me._sstOrder_T10.Name = "_sstOrder_T10"
        Me._sstOrder_T10.Padding = New System.Windows.Forms.Padding(3)
        Me._sstOrder_T10.Size = New System.Drawing.Size(992, 574)
        Me._sstOrder_T10.TabIndex = 12
        Me._sstOrder_T10.Text = "11-Reserved items"
        Me._sstOrder_T10.UseVisualStyleBackColor = True
        '
        'dgvReservedItems
        '
        Me.dgvReservedItems.AllowUserToAddRows = False
        Me.dgvReservedItems.AllowUserToDeleteRows = False
        Me.dgvReservedItems.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.dgvReservedItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvReservedItems.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke
        Me.dgvReservedItems.Location = New System.Drawing.Point(8, 9)
        Me.dgvReservedItems.MultiSelect = False
        Me.dgvReservedItems.Name = "dgvReservedItems"
        Me.dgvReservedItems.ReadOnly = True
        Me.dgvReservedItems.RowHeadersWidth = 20
        Me.dgvReservedItems.Size = New System.Drawing.Size(977, 551)
        Me.dgvReservedItems.TabIndex = 155
        '
        'sbOrder
        '
        Me.sbOrder.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sbOrder.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._sbOrder_Panel1})
        Me.sbOrder.Location = New System.Drawing.Point(0, 646)
        Me.sbOrder.Name = "sbOrder"
        Me.sbOrder.Size = New System.Drawing.Size(1014, 22)
        Me.sbOrder.TabIndex = 2
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
        Me._sbOrder_Panel1.Size = New System.Drawing.Size(999, 22)
        Me._sbOrder_Panel1.Spring = True
        Me._sbOrder_Panel1.Text = "B=Blanket  C=Credit Memo I=Invoice M=Master O=Order Q=Quote"
        Me._sbOrder_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbOrder
        '
        Me.tbOrder.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbNew, Me.tsbExportToMacola, Me.tsbXmlExport})
        Me.tbOrder.Location = New System.Drawing.Point(0, 0)
        Me.tbOrder.Name = "tbOrder"
        Me.tbOrder.Size = New System.Drawing.Size(1014, 25)
        Me.tbOrder.TabIndex = 0
        '
        'tsbNew
        '
        Me.tsbNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbNew.Image = CType(resources.GetObject("tsbNew.Image"), System.Drawing.Image)
        Me.tsbNew.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbNew.Name = "tsbNew"
        Me.tsbNew.Size = New System.Drawing.Size(35, 22)
        Me.tsbNew.Text = "New"
        '
        'tsbExportToMacola
        '
        Me.tsbExportToMacola.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbExportToMacola.Image = CType(resources.GetObject("tsbExportToMacola.Image"), System.Drawing.Image)
        Me.tsbExportToMacola.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbExportToMacola.Name = "tsbExportToMacola"
        Me.tsbExportToMacola.Size = New System.Drawing.Size(103, 22)
        Me.tsbExportToMacola.Text = "Export To Macola"
        '
        'tsbXmlExport
        '
        Me.tsbXmlExport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbXmlExport.Image = CType(resources.GetObject("tsbXmlExport.Image"), System.Drawing.Image)
        Me.tsbXmlExport.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbXmlExport.Name = "tsbXmlExport"
        Me.tsbXmlExport.Size = New System.Drawing.Size(71, 22)
        Me.tsbXmlExport.Text = "XML Export"
        '
        'lblWhite_Glove
        '
        Me.lblWhite_Glove.AutoSize = True
        Me.lblWhite_Glove.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblWhite_Glove.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWhite_Glove.ForeColor = System.Drawing.Color.White
        Me.lblWhite_Glove.Location = New System.Drawing.Point(420, 25)
        Me.lblWhite_Glove.Name = "lblWhite_Glove"
        Me.lblWhite_Glove.Size = New System.Drawing.Size(194, 19)
        Me.lblWhite_Glove.TabIndex = 3
        Me.lblWhite_Glove.Text = "       White Glove Alert       "
        Me.lblWhite_Glove.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblWhite_Glove.Visible = False
        '
        'frmOrder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1014, 668)
        Me.Controls.Add(Me.sstOrder)
        Me.Controls.Add(Me.sbOrder)
        Me.Controls.Add(Me.tbOrder)
        Me.Controls.Add(Me.lblWhite_Glove)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(8, 30)
        Me.MinimumSize = New System.Drawing.Size(1020, 700)
        Me.Name = "frmOrder"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Order"
        Me.sstOrder.ResumeLayout(False)
        Me._sstOrder_T00.ResumeLayout(False)
        Me._sstOrder_T01.ResumeLayout(False)
        Me._sstOrder_T02.ResumeLayout(False)
        Me._sstOrder_T02.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panImprint.ResumeLayout(False)
        Me.panImprint.PerformLayout()
        Me.gbRepeatData.ResumeLayout(False)
        Me.gbRepeatData.PerformLayout()
        CType(Me.dgvItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me._sstOrder_T03.ResumeLayout(False)
        Me._sstOrder_T04.ResumeLayout(False)
        Me._sstOrder_T05.ResumeLayout(False)
        Me._sstOrder_T06.ResumeLayout(False)
        Me._sstOrder_T07.ResumeLayout(False)
        Me._sstOrder_T08.ResumeLayout(False)
        Me._sstOrder_T09.ResumeLayout(False)
        CType(Me.dgvHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me._sstOrder_T10.ResumeLayout(False)
        CType(Me.dgvReservedItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.sbOrder.ResumeLayout(False)
        Me.sbOrder.PerformLayout()
        Me.tbOrder.ResumeLayout(False)
        Me.tbOrder.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UcTaxes1 As Orders.ucTaxes
    Friend WithEvents UcSalesperson1 As Orders.ucSalesperson
    Friend WithEvents UcCreditInfo1 As Orders.ucCreditInfo
    Friend WithEvents UcExtra1 As Orders.ucExtra
    Public WithEvents sstOrder As System.Windows.Forms.TabControl
    Friend WithEvents UcOrderTotal1 As Orders.ucOrderTotal
    Public WithEvents _sstOrder_T00 As System.Windows.Forms.TabPage
    Friend WithEvents UcOrder1 As Orders.ucOrder
    Friend WithEvents UcOrder As Orders.ucOrder
    Friend WithEvents tsbNew As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbExportToMacola As System.Windows.Forms.ToolStripButton
    Friend WithEvents UcHeader1 As Orders.ucHeader
    Friend WithEvents cmdCopyLine As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteLine As System.Windows.Forms.Button
    Friend WithEvents cmdRestart As System.Windows.Forms.Button
    Friend WithEvents cboMoreOptions As System.Windows.Forms.ComboBox
    Friend WithEvents XTextBox2 As Orders.xTextBox
    Friend WithEvents cmdPrevCharges As System.Windows.Forms.Button
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdNotes As System.Windows.Forms.Button
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents UcContacts1 As Orders.ucContacts
    Friend WithEvents UcImprint1 As Orders.ucImprint
    Friend WithEvents panImprint As System.Windows.Forms.Panel
    Friend WithEvents cmdPaste As System.Windows.Forms.Button
    Friend WithEvents cmdUpdateAllItem As System.Windows.Forms.Button
    Friend WithEvents cmdUpdateAll As System.Windows.Forms.Button
    Friend WithEvents gbRepeatData As System.Windows.Forms.GroupBox
    Friend WithEvents cmdGetRepeatData As System.Windows.Forms.Button
    Friend WithEvents txtRepeat_Ord_No As System.Windows.Forms.TextBox
    Friend WithEvents lblRepeat_Ord_No As System.Windows.Forms.Label
    Friend WithEvents cmdClear_Imprint As System.Windows.Forms.Button
    Friend WithEvents txtImprint_Item_No As System.Windows.Forms.TextBox
    Friend WithEvents lblImprint_Item_No As System.Windows.Forms.Label
    Friend WithEvents cmdCopyDown As System.Windows.Forms.Button
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
    Friend WithEvents lblImprint As System.Windows.Forms.Label
    Friend WithEvents cmdInsertLine As System.Windows.Forms.Button
    Friend WithEvents lblIndustry As System.Windows.Forms.Label
    Friend WithEvents txtIndustry As System.Windows.Forms.TextBox
    Friend WithEvents cboIndustry As System.Windows.Forms.ComboBox
    Friend WithEvents wbItem As System.Windows.Forms.WebBrowser
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents cmdDash As System.Windows.Forms.Button
    Friend WithEvents cmdRepeatSearch As System.Windows.Forms.Button
    Friend WithEvents cmdEmail As System.Windows.Forms.Button
    Friend WithEvents btnODBCConnect As System.Windows.Forms.Button
    Friend WithEvents chkLowRow As System.Windows.Forms.CheckBox
    Friend WithEvents cmdImprint As System.Windows.Forms.Button
    Friend WithEvents cmdTraveler As System.Windows.Forms.Button
    Friend WithEvents cmdQty As System.Windows.Forms.Button
    Friend WithEvents cmdRequest_Dt As System.Windows.Forms.Button
    Friend WithEvents cmdDirServices As System.Windows.Forms.Button
    Friend WithEvents chkAutoCompleteReship As System.Windows.Forms.CheckBox
    Public WithEvents _sstOrder_T01 As System.Windows.Forms.TabPage
    Public WithEvents _sstOrder_T09 As System.Windows.Forms.TabPage
    Friend WithEvents dgvHistory As Orders.XDataGridView
    Friend WithEvents cmdPrice As System.Windows.Forms.Button
    Friend WithEvents txtRepeat_From_ID As System.Windows.Forms.TextBox
    Friend WithEvents _sstOrder_T10 As System.Windows.Forms.TabPage
    Friend WithEvents dgvReservedItems As Orders.XDataGridView
    Friend WithEvents _sstOrder_T03 As System.Windows.Forms.TabPage
    Friend WithEvents UcDocument1 As Orders.ucDocument
    Friend WithEvents dgvItems As Orders.XDataGridView
    Friend WithEvents cmdOEI_Config As System.Windows.Forms.Button
    Friend WithEvents txtNum_Impr_12 As System.Windows.Forms.TextBox
    Friend WithEvents lblNum_Impr_2 As System.Windows.Forms.Label
    Friend WithEvents txtNum_Impr_1 As Orders.xTextBox
    Friend WithEvents txtNum_Impr_2 As Orders.xTextBox
    Friend WithEvents txtNum_Impr_3 As Orders.xTextBox
    Friend WithEvents lblNum_Impr_3 As System.Windows.Forms.Label
    Friend WithEvents btnEDI As System.Windows.Forms.Button
    Friend WithEvents btnFileList As System.Windows.Forms.Button
    Friend WithEvents btnThermo As System.Windows.Forms.Button
    Friend WithEvents cmdMySqlTest As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents cmdReset_Cus_No As System.Windows.Forms.Button
    Friend WithEvents cmdEDI_Alert_New_Inv As System.Windows.Forms.Button
    Friend WithEvents cmdEDIEditSuppItem As System.Windows.Forms.Button
    Friend WithEvents cmdETA_Calculator As System.Windows.Forms.Button
    Friend WithEvents lblWhite_Glove As System.Windows.Forms.Label
    Friend WithEvents lblEnd_User As System.Windows.Forms.Label
    Friend WithEvents txtEnd_User As System.Windows.Forms.TextBox
    Friend WithEvents tsbXmlExport As System.Windows.Forms.ToolStripButton
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    Friend WithEvents btnDbgPrxAdjust As System.Windows.Forms.Button
    Friend WithEvents testSimulateThermo As System.Windows.Forms.Button
#End Region
End Class