<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class ucOrder
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
    Friend WithEvents cboShipVia As System.Windows.Forms.ComboBox
    Friend WithEvents cboShipTo As System.Windows.Forms.ComboBox
    Friend WithEvents cboCustomerNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboShipToName As System.Windows.Forms.ComboBox
    Friend WithEvents cboBillToName As System.Windows.Forms.ComboBox
    Friend WithEvents cmdCopyBillTo As System.Windows.Forms.Button
    Friend WithEvents cmdShipToCountrySearch As System.Windows.Forms.Button
    Friend WithEvents cmdChangeBillTo As System.Windows.Forms.Button
    Friend WithEvents cmdBillToCountrySearch As System.Windows.Forms.Button
    Friend WithEvents cmdCommentsSearch As System.Windows.Forms.Button
    Friend WithEvents cmdDeptSearch As System.Windows.Forms.Button
    Friend WithEvents cmdDept_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdProfitCenterSearch As System.Windows.Forms.Button
    Friend WithEvents cmdTermsSearch As System.Windows.Forms.Button
    Friend WithEvents cmdShipViaSearch As System.Windows.Forms.Button
    Friend WithEvents cmdLocationSearch As System.Windows.Forms.Button
    Friend WithEvents cmdInstructionsSearch As System.Windows.Forms.Button
    Friend WithEvents cmdProjectSearch As System.Windows.Forms.Button
    Friend WithEvents cmdSalespersonSearch As System.Windows.Forms.Button
    Friend WithEvents CmdShipToSearch As System.Windows.Forms.Button
    Friend WithEvents cmdCustomerSearch As System.Windows.Forms.Button
    Friend WithEvents cmdDateTP As System.Windows.Forms.Button
    Friend WithEvents cmdOrderSearch As System.Windows.Forms.Button
    Friend WithEvents cboOrd_Type As System.Windows.Forms.ComboBox
    Friend WithEvents cmdCustomerConfig As System.Windows.Forms.Button
    Friend WithEvents cmdCopyOrderFromHistory As System.Windows.Forms.Button
    Friend WithEvents lblShipToCountryDesc As System.Windows.Forms.Label
    Friend WithEvents lblShipToCountry As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents lblShipToName As System.Windows.Forms.Label
    Friend WithEvents lblBillToCountryDesc As System.Windows.Forms.Label
    Friend WithEvents lblBillToCountry As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents lblBillToName As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents lblTermsDesc As System.Windows.Forms.Label
    Friend WithEvents lblShipViaDesc As System.Windows.Forms.Label
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblShipVia As System.Windows.Forms.Label
    Friend WithEvents lblTerms As System.Windows.Forms.Label
    Friend WithEvents lblDiscPct As System.Windows.Forms.Label
    Friend WithEvents lblProfitCenter As System.Windows.Forms.Label
    Friend WithEvents lblInstructions As System.Windows.Forms.Label
    Friend WithEvents lblProject As System.Windows.Forms.Label
    Friend WithEvents lblSalesPerson As System.Windows.Forms.Label
    Friend WithEvents lblPONumber As System.Windows.Forms.Label
    Friend WithEvents lblShipTo As System.Windows.Forms.Label
    Friend WithEvents lblCustomer As System.Windows.Forms.Label
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents lblOrder As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ucOrder))
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.LineShape1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.Line3 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.cboShipVia = New System.Windows.Forms.ComboBox()
        Me.cboShipTo = New System.Windows.Forms.ComboBox()
        Me.cboCustomerNo = New System.Windows.Forms.ComboBox()
        Me.cboShipToName = New System.Windows.Forms.ComboBox()
        Me.cboBillToName = New System.Windows.Forms.ComboBox()
        Me.cmdCopyBillTo = New System.Windows.Forms.Button()
        Me.cmdShipToCountrySearch = New System.Windows.Forms.Button()
        Me.cmdChangeBillTo = New System.Windows.Forms.Button()
        Me.cmdBillToCountrySearch = New System.Windows.Forms.Button()
        Me.cmdCommentsSearch = New System.Windows.Forms.Button()
        Me.cmdDeptSearch = New System.Windows.Forms.Button()
        Me.cmdDept_Notes = New System.Windows.Forms.Button()
        Me.cmdProfitCenterSearch = New System.Windows.Forms.Button()
        Me.cmdTermsSearch = New System.Windows.Forms.Button()
        Me.cmdShipViaSearch = New System.Windows.Forms.Button()
        Me.cmdLocationSearch = New System.Windows.Forms.Button()
        Me.cmdInstructionsSearch = New System.Windows.Forms.Button()
        Me.cmdProjectSearch = New System.Windows.Forms.Button()
        Me.cmdSalespersonSearch = New System.Windows.Forms.Button()
        Me.CmdShipToSearch = New System.Windows.Forms.Button()
        Me.cmdCustomerSearch = New System.Windows.Forms.Button()
        Me.cmdDateTP = New System.Windows.Forms.Button()
        Me.cmdOrderSearch = New System.Windows.Forms.Button()
        Me.cboOrd_Type = New System.Windows.Forms.ComboBox()
        Me.cmdCustomerConfig = New System.Windows.Forms.Button()
        Me.cmdCopyOrderFromHistory = New System.Windows.Forms.Button()
        Me.lblShipToCountryDesc = New System.Windows.Forms.Label()
        Me.lblShipToCountry = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.lblShipToName = New System.Windows.Forms.Label()
        Me.lblBillToCountryDesc = New System.Windows.Forms.Label()
        Me.lblBillToCountry = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.lblBillToName = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.lblTermsDesc = New System.Windows.Forms.Label()
        Me.lblShipViaDesc = New System.Windows.Forms.Label()
        Me.lblLocationDesc = New System.Windows.Forms.Label()
        Me.lblLocation = New System.Windows.Forms.Label()
        Me.lblShipVia = New System.Windows.Forms.Label()
        Me.lblTerms = New System.Windows.Forms.Label()
        Me.lblDiscPct = New System.Windows.Forms.Label()
        Me.lblProfitCenter = New System.Windows.Forms.Label()
        Me.lblInstructions = New System.Windows.Forms.Label()
        Me.lblProject = New System.Windows.Forms.Label()
        Me.lblSalesPerson = New System.Windows.Forms.Label()
        Me.lblPONumber = New System.Windows.Forms.Label()
        Me.lblShipTo = New System.Windows.Forms.Label()
        Me.lblCustomer = New System.Windows.Forms.Label()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.lblOrder = New System.Windows.Forms.Label()
        Me.lblType = New System.Windows.Forms.Label()
        Me.mcCalendar = New System.Windows.Forms.MonthCalendar()
        Me.lblCostCenter = New System.Windows.Forms.Label()
        Me.lblComments = New System.Windows.Forms.Label()
        Me.ttToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdMfg_Loc_Notes = New System.Windows.Forms.Button()
        Me.cmdOrd_No_Notes = New System.Windows.Forms.Button()
        Me.cmdCus_No_Notes = New System.Windows.Forms.Button()
        Me.cmdCus_Alt_Adr_Cd_Notes = New System.Windows.Forms.Button()
        Me.cmdSlspsn_No_Notes = New System.Windows.Forms.Button()
        Me.cmdJob_No_Notes = New System.Windows.Forms.Button()
        Me.cmdShip_Via_Cd_Notes = New System.Windows.Forms.Button()
        Me.cmdAr_Terms_Cd_Notes = New System.Windows.Forms.Button()
        Me.cmdProfit_Center_Notes = New System.Windows.Forms.Button()
        Me.cmdBill_To_Country_Notes = New System.Windows.Forms.Button()
        Me.cmdShip_To_Country_Notes = New System.Windows.Forms.Button()
        Me.cmdConvertQuote = New System.Windows.Forms.Button()
        Me.cmdPrintQuote = New System.Windows.Forms.Button()
        Me.cmdHistoryCopy = New System.Windows.Forms.Button()
        Me.cmdMasterCopy = New System.Windows.Forms.Button()
        Me.cmdLineUserFields = New System.Windows.Forms.Button()
        Me.cmdInvoice = New System.Windows.Forms.Button()
        Me.cmdPickTicket = New System.Windows.Forms.Button()
        Me.cmdAcknowledgement = New System.Windows.Forms.Button()
        Me.cmdComments = New System.Windows.Forms.Button()
        Me.cmdLocQty = New System.Windows.Forms.Button()
        Me.cmdItemNotes = New System.Windows.Forms.Button()
        Me.cmdItemInfo = New System.Windows.Forms.Button()
        Me.cmdCommissions = New System.Windows.Forms.Button()
        Me.cmdTax = New System.Windows.Forms.Button()
        Me.cmdEditCapturedBill = New System.Windows.Forms.Button()
        Me.cmdATP = New System.Windows.Forms.Button()
        Me.lblUser_Def_Fld_2 = New System.Windows.Forms.Label()
        Me.cboUser_Def_Fld_2 = New System.Windows.Forms.ComboBox()
        Me.lblUser_Def_Fld_3 = New System.Windows.Forms.Label()
        Me.lblUser_Def_Fld_1 = New System.Windows.Forms.Label()
        Me.lblUser_Def_Fld_4 = New System.Windows.Forms.Label()
        Me.lblContactEmail = New System.Windows.Forms.Label()
        Me.cmdUser_Def_Fld_4Search = New System.Windows.Forms.Button()
        Me.cmdOEFlash = New System.Windows.Forms.Button()
        Me.cmdOrderLink = New System.Windows.Forms.Button()
        Me.cmdCopyDate = New System.Windows.Forms.Button()
        Me.cmdUser_Def_Fld_5Search = New System.Windows.Forms.Button()
        Me.lblUser_Def_Fld_5 = New System.Windows.Forms.Label()
        Me.cmdShip_LinkEdit = New System.Windows.Forms.Button()
        Me.cmdShip_LinkSearch = New System.Windows.Forms.Button()
        Me.lblShip_Link = New System.Windows.Forms.Label()
        Me.cmdUser_Def_Fld_3Search = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.cmdShip_To_NameSearch = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.chkOrderAckSaveOnly = New System.Windows.Forms.CheckBox()
        Me.lblOptions = New System.Windows.Forms.Label()
        Me.chkRecover = New System.Windows.Forms.CheckBox()
        Me.chkRedo = New System.Windows.Forms.CheckBox()
        Me.chkExactRepeat = New System.Windows.Forms.CheckBox()
        Me.chkInvalidShipDateEmail = New System.Windows.Forms.CheckBox()
        Me.chkEmail2 = New System.Windows.Forms.CheckBox()
        Me.cboProgram = New System.Windows.Forms.ComboBox()
        Me.lblProgram = New System.Windows.Forms.Label()
        Me.cmdOrderRequest = New System.Windows.Forms.Button()
        Me.cmdSendEmail = New System.Windows.Forms.Button()
        Me.cmdEDI = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.cmdProg_Spector_CdSearch = New System.Windows.Forms.Button()
        Me.cmdCustomerDash = New System.Windows.Forms.Button()
        Me.lblEnd_User = New System.Windows.Forms.Label()
        Me.cmdEnd_UserSearch = New System.Windows.Forms.Button()
        Me.cmdEnd_User_Notes = New System.Windows.Forms.Button()
        Me.cmdInHandDate = New System.Windows.Forms.Button()
        Me.lblInHandsDate = New System.Windows.Forms.Label()
        Me.cmdApplyToSearch = New System.Windows.Forms.Button()
        Me.lblApplyTo = New System.Windows.Forms.Label()
        Me.cmdShipDateTP = New System.Windows.Forms.Button()
        Me.lblShipDate = New System.Windows.Forms.Label()
        Me.cmdXml = New System.Windows.Forms.Button()
        Me.cmdApplyToNote = New System.Windows.Forms.Button()
        Me.cbo_Extra_9 = New System.Windows.Forms.ComboBox()
        Me.lblExtra_9 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboWhyNotVegas = New System.Windows.Forms.ComboBox()
        Me.cboreSSP = New System.Windows.Forms.ComboBox()
        Me.lblSSp = New System.Windows.Forms.Label()
        Me.lblInvBatch = New System.Windows.Forms.Label()
        Me.xTxtInvBatch = New Orders.xTextBox()
        Me.txtCustGroup = New Orders.xTextBox()
        Me.txtOrd_Dt_Shipped = New Orders.xTextBox()
        Me.txtInHandsDate = New Orders.xTextBox()
        Me.txtApplyTo = New Orders.xTextBox()
        Me.txtDept = New Orders.xTextBox()
        Me.txtCus_Prog_ID = New Orders.xTextBox()
        Me.txtProg_Spector_Cd = New Orders.xTextBox()
        Me.txtShip_To_Zip = New Orders.xTextBox()
        Me.txtShip_To_State = New Orders.xTextBox()
        Me.txtShip_To_City = New Orders.xTextBox()
        Me.txtShip_To_Addr_4 = New Orders.xTextBox()
        Me.txtBill_To_Zip = New Orders.xTextBox()
        Me.txtBill_To_City = New Orders.xTextBox()
        Me.txtBill_To_State = New Orders.xTextBox()
        Me.txtBill_To_Addr_4 = New Orders.xTextBox()
        Me.txtShip_Link = New Orders.xTextBox()
        Me.txtUser_Def_Fld_5 = New Orders.xTextBox()
        Me.txtContactEmail = New Orders.xTextBox()
        Me.txtUser_Def_Fld_4 = New Orders.xTextBox()
        Me.txtUser_Def_Fld_3 = New Orders.xTextBox()
        Me.txtUser_Def_Fld_1 = New Orders.xTextBox()
        Me.txtShip_To_Country = New Orders.xTextBox()
        Me.txtBill_To_Country = New Orders.xTextBox()
        Me.txtShip_To_Addr_3 = New Orders.xTextBox()
        Me.txtShip_To_Addr_2 = New Orders.xTextBox()
        Me.txtShip_To_Addr_1 = New Orders.xTextBox()
        Me.txtShip_To_Name = New Orders.xTextBox()
        Me.txtBill_To_Addr_3 = New Orders.xTextBox()
        Me.txtBill_To_Addr_2 = New Orders.xTextBox()
        Me.txtBill_To_Addr_1 = New Orders.xTextBox()
        Me.txtBill_To_Name = New Orders.xTextBox()
        Me.txtShip_Instruction_2 = New Orders.xTextBox()
        Me.txtProfit_Center = New Orders.xTextBox()
        Me.txtAr_Terms_Cd = New Orders.xTextBox()
        Me.txtShip_Via_Cd = New Orders.xTextBox()
        Me.txtMfg_Loc = New Orders.xTextBox()
        Me.txtShip_Instruction_1 = New Orders.xTextBox()
        Me.txtJob_No = New Orders.xTextBox()
        Me.txtSlspsn_No = New Orders.xTextBox()
        Me.txtOe_Po_No = New Orders.xTextBox()
        Me.txtCus_Alt_Adr_Cd = New Orders.xTextBox()
        Me.txtCus_No = New Orders.xTextBox()
        Me.txtOrd_No = New Orders.xTextBox()
        Me.txtDiscount_Pct = New Orders.xTextBox()
        Me.txtOrd_Dt = New Orders.xTextBox()
        Me.txtEnd_User = New Orders.xTextBox()
        Me.lblVIP = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ShapeContainer1
        '
        resources.ApplyResources(Me.ShapeContainer1, "ShapeContainer1")
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape1, Me.Line3})
        Me.ShapeContainer1.TabStop = False
        '
        'LineShape1
        '
        resources.ApplyResources(Me.LineShape1, "LineShape1")
        Me.LineShape1.Name = "LineShape1"
        '
        'Line3
        '
        resources.ApplyResources(Me.Line3, "Line3")
        Me.Line3.Name = "Line3"
        '
        'cboShipVia
        '
        Me.cboShipVia.BackColor = System.Drawing.SystemColors.Window
        Me.cboShipVia.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cboShipVia, "cboShipVia")
        Me.cboShipVia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboShipVia.Name = "cboShipVia"
        '
        'cboShipTo
        '
        Me.cboShipTo.BackColor = System.Drawing.SystemColors.Window
        Me.cboShipTo.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cboShipTo, "cboShipTo")
        Me.cboShipTo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboShipTo.Name = "cboShipTo"
        '
        'cboCustomerNo
        '
        Me.cboCustomerNo.BackColor = System.Drawing.SystemColors.Window
        Me.cboCustomerNo.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cboCustomerNo, "cboCustomerNo")
        Me.cboCustomerNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCustomerNo.Name = "cboCustomerNo"
        '
        'cboShipToName
        '
        Me.cboShipToName.BackColor = System.Drawing.SystemColors.Window
        Me.cboShipToName.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cboShipToName, "cboShipToName")
        Me.cboShipToName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboShipToName.Name = "cboShipToName"
        '
        'cboBillToName
        '
        Me.cboBillToName.BackColor = System.Drawing.SystemColors.Window
        Me.cboBillToName.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cboBillToName, "cboBillToName")
        Me.cboBillToName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboBillToName.Name = "cboBillToName"
        '
        'cmdCopyBillTo
        '
        Me.cmdCopyBillTo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopyBillTo.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCopyBillTo, "cmdCopyBillTo")
        Me.cmdCopyBillTo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopyBillTo.Name = "cmdCopyBillTo"
        Me.cmdCopyBillTo.TabStop = False
        Me.cmdCopyBillTo.UseVisualStyleBackColor = False
        '
        'cmdShipToCountrySearch
        '
        Me.cmdShipToCountrySearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShipToCountrySearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdShipToCountrySearch, "cmdShipToCountrySearch")
        Me.cmdShipToCountrySearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShipToCountrySearch.Name = "cmdShipToCountrySearch"
        Me.cmdShipToCountrySearch.TabStop = False
        Me.cmdShipToCountrySearch.UseVisualStyleBackColor = False
        '
        'cmdChangeBillTo
        '
        Me.cmdChangeBillTo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdChangeBillTo.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdChangeBillTo, "cmdChangeBillTo")
        Me.cmdChangeBillTo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdChangeBillTo.Name = "cmdChangeBillTo"
        Me.cmdChangeBillTo.TabStop = False
        Me.cmdChangeBillTo.UseVisualStyleBackColor = False
        '
        'cmdBillToCountrySearch
        '
        Me.cmdBillToCountrySearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBillToCountrySearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdBillToCountrySearch, "cmdBillToCountrySearch")
        Me.cmdBillToCountrySearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBillToCountrySearch.Name = "cmdBillToCountrySearch"
        Me.cmdBillToCountrySearch.TabStop = False
        Me.cmdBillToCountrySearch.UseVisualStyleBackColor = False
        '
        'cmdCommentsSearch
        '
        Me.cmdCommentsSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCommentsSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCommentsSearch, "cmdCommentsSearch")
        Me.cmdCommentsSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCommentsSearch.Name = "cmdCommentsSearch"
        Me.cmdCommentsSearch.TabStop = False
        Me.cmdCommentsSearch.UseVisualStyleBackColor = False
        '
        'cmdDeptSearch
        '
        Me.cmdDeptSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDeptSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdDeptSearch, "cmdDeptSearch")
        Me.cmdDeptSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDeptSearch.Name = "cmdDeptSearch"
        Me.cmdDeptSearch.TabStop = False
        Me.cmdDeptSearch.UseVisualStyleBackColor = False
        '
        'cmdDept_Notes
        '
        Me.cmdDept_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDept_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdDept_Notes, "cmdDept_Notes")
        Me.cmdDept_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDept_Notes.Name = "cmdDept_Notes"
        Me.cmdDept_Notes.TabStop = False
        Me.cmdDept_Notes.UseVisualStyleBackColor = False
        '
        'cmdProfitCenterSearch
        '
        Me.cmdProfitCenterSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdProfitCenterSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdProfitCenterSearch, "cmdProfitCenterSearch")
        Me.cmdProfitCenterSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdProfitCenterSearch.Name = "cmdProfitCenterSearch"
        Me.cmdProfitCenterSearch.TabStop = False
        Me.cmdProfitCenterSearch.UseVisualStyleBackColor = False
        '
        'cmdTermsSearch
        '
        Me.cmdTermsSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdTermsSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdTermsSearch, "cmdTermsSearch")
        Me.cmdTermsSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdTermsSearch.Name = "cmdTermsSearch"
        Me.cmdTermsSearch.TabStop = False
        Me.cmdTermsSearch.UseVisualStyleBackColor = False
        '
        'cmdShipViaSearch
        '
        Me.cmdShipViaSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShipViaSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdShipViaSearch, "cmdShipViaSearch")
        Me.cmdShipViaSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShipViaSearch.Name = "cmdShipViaSearch"
        Me.cmdShipViaSearch.TabStop = False
        Me.cmdShipViaSearch.UseVisualStyleBackColor = False
        '
        'cmdLocationSearch
        '
        Me.cmdLocationSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdLocationSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdLocationSearch, "cmdLocationSearch")
        Me.cmdLocationSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLocationSearch.Name = "cmdLocationSearch"
        Me.cmdLocationSearch.TabStop = False
        Me.cmdLocationSearch.UseVisualStyleBackColor = False
        '
        'cmdInstructionsSearch
        '
        Me.cmdInstructionsSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInstructionsSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdInstructionsSearch, "cmdInstructionsSearch")
        Me.cmdInstructionsSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInstructionsSearch.Name = "cmdInstructionsSearch"
        Me.cmdInstructionsSearch.TabStop = False
        Me.cmdInstructionsSearch.UseVisualStyleBackColor = False
        '
        'cmdProjectSearch
        '
        Me.cmdProjectSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdProjectSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdProjectSearch, "cmdProjectSearch")
        Me.cmdProjectSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdProjectSearch.Name = "cmdProjectSearch"
        Me.cmdProjectSearch.TabStop = False
        Me.cmdProjectSearch.UseVisualStyleBackColor = False
        '
        'cmdSalespersonSearch
        '
        Me.cmdSalespersonSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSalespersonSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdSalespersonSearch, "cmdSalespersonSearch")
        Me.cmdSalespersonSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSalespersonSearch.Name = "cmdSalespersonSearch"
        Me.cmdSalespersonSearch.TabStop = False
        Me.cmdSalespersonSearch.UseVisualStyleBackColor = False
        '
        'CmdShipToSearch
        '
        Me.CmdShipToSearch.BackColor = System.Drawing.SystemColors.Control
        Me.CmdShipToSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.CmdShipToSearch, "CmdShipToSearch")
        Me.CmdShipToSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdShipToSearch.Name = "CmdShipToSearch"
        Me.CmdShipToSearch.TabStop = False
        Me.CmdShipToSearch.UseVisualStyleBackColor = False
        '
        'cmdCustomerSearch
        '
        Me.cmdCustomerSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCustomerSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCustomerSearch, "cmdCustomerSearch")
        Me.cmdCustomerSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCustomerSearch.Name = "cmdCustomerSearch"
        Me.cmdCustomerSearch.TabStop = False
        Me.cmdCustomerSearch.UseVisualStyleBackColor = False
        '
        'cmdDateTP
        '
        Me.cmdDateTP.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDateTP.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdDateTP, "cmdDateTP")
        Me.cmdDateTP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDateTP.Name = "cmdDateTP"
        Me.cmdDateTP.TabStop = False
        Me.cmdDateTP.UseVisualStyleBackColor = False
        '
        'cmdOrderSearch
        '
        Me.cmdOrderSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrderSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdOrderSearch, "cmdOrderSearch")
        Me.cmdOrderSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrderSearch.Name = "cmdOrderSearch"
        Me.cmdOrderSearch.TabStop = False
        Me.cmdOrderSearch.UseVisualStyleBackColor = False
        '
        'cboOrd_Type
        '
        Me.cboOrd_Type.BackColor = System.Drawing.SystemColors.Window
        Me.cboOrd_Type.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cboOrd_Type, "cboOrd_Type")
        Me.cboOrd_Type.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboOrd_Type.Items.AddRange(New Object() {resources.GetString("cboOrd_Type.Items"), resources.GetString("cboOrd_Type.Items1")})
        Me.cboOrd_Type.Name = "cboOrd_Type"
        '
        'cmdCustomerConfig
        '
        Me.cmdCustomerConfig.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCustomerConfig.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCustomerConfig, "cmdCustomerConfig")
        Me.cmdCustomerConfig.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCustomerConfig.Name = "cmdCustomerConfig"
        Me.cmdCustomerConfig.TabStop = False
        Me.cmdCustomerConfig.UseVisualStyleBackColor = False
        '
        'cmdCopyOrderFromHistory
        '
        Me.cmdCopyOrderFromHistory.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopyOrderFromHistory.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCopyOrderFromHistory, "cmdCopyOrderFromHistory")
        Me.cmdCopyOrderFromHistory.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopyOrderFromHistory.Name = "cmdCopyOrderFromHistory"
        Me.cmdCopyOrderFromHistory.TabStop = False
        Me.cmdCopyOrderFromHistory.UseVisualStyleBackColor = False
        '
        'lblShipToCountryDesc
        '
        Me.lblShipToCountryDesc.BackColor = System.Drawing.SystemColors.Control
        Me.lblShipToCountryDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblShipToCountryDesc.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblShipToCountryDesc, "lblShipToCountryDesc")
        Me.lblShipToCountryDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShipToCountryDesc.Name = "lblShipToCountryDesc"
        '
        'lblShipToCountry
        '
        Me.lblShipToCountry.BackColor = System.Drawing.SystemColors.Control
        Me.lblShipToCountry.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblShipToCountry, "lblShipToCountry")
        Me.lblShipToCountry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShipToCountry.Name = "lblShipToCountry"
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.SystemColors.Control
        Me.Label28.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label28, "Label28")
        Me.Label28.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label28.Name = "Label28"
        '
        'lblShipToName
        '
        Me.lblShipToName.BackColor = System.Drawing.SystemColors.Control
        Me.lblShipToName.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblShipToName, "lblShipToName")
        Me.lblShipToName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShipToName.Name = "lblShipToName"
        '
        'lblBillToCountryDesc
        '
        Me.lblBillToCountryDesc.BackColor = System.Drawing.SystemColors.Control
        Me.lblBillToCountryDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblBillToCountryDesc.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblBillToCountryDesc, "lblBillToCountryDesc")
        Me.lblBillToCountryDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBillToCountryDesc.Name = "lblBillToCountryDesc"
        '
        'lblBillToCountry
        '
        Me.lblBillToCountry.BackColor = System.Drawing.SystemColors.Control
        Me.lblBillToCountry.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblBillToCountry, "lblBillToCountry")
        Me.lblBillToCountry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBillToCountry.Name = "lblBillToCountry"
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.SystemColors.Control
        Me.Label24.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label24, "Label24")
        Me.Label24.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label24.Name = "Label24"
        '
        'lblBillToName
        '
        Me.lblBillToName.BackColor = System.Drawing.SystemColors.Control
        Me.lblBillToName.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblBillToName, "lblBillToName")
        Me.lblBillToName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBillToName.Name = "lblBillToName"
        '
        'Label21
        '
        Me.Label21.BackColor = System.Drawing.SystemColors.Control
        Me.Label21.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label21, "Label21")
        Me.Label21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label21.Name = "Label21"
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.SystemColors.Control
        Me.Label20.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label20, "Label20")
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Name = "Label20"
        '
        'lblTermsDesc
        '
        Me.lblTermsDesc.AutoEllipsis = True
        Me.lblTermsDesc.BackColor = System.Drawing.SystemColors.Control
        Me.lblTermsDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTermsDesc.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblTermsDesc, "lblTermsDesc")
        Me.lblTermsDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTermsDesc.Name = "lblTermsDesc"
        '
        'lblShipViaDesc
        '
        Me.lblShipViaDesc.AutoEllipsis = True
        Me.lblShipViaDesc.BackColor = System.Drawing.SystemColors.Control
        Me.lblShipViaDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblShipViaDesc.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblShipViaDesc, "lblShipViaDesc")
        Me.lblShipViaDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShipViaDesc.Name = "lblShipViaDesc"
        '
        'lblLocationDesc
        '
        Me.lblLocationDesc.AutoEllipsis = True
        Me.lblLocationDesc.BackColor = System.Drawing.SystemColors.Control
        Me.lblLocationDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLocationDesc.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblLocationDesc, "lblLocationDesc")
        Me.lblLocationDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLocationDesc.Name = "lblLocationDesc"
        '
        'lblLocation
        '
        Me.lblLocation.BackColor = System.Drawing.SystemColors.Control
        Me.lblLocation.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblLocation, "lblLocation")
        Me.lblLocation.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLocation.Name = "lblLocation"
        '
        'lblShipVia
        '
        Me.lblShipVia.BackColor = System.Drawing.SystemColors.Control
        Me.lblShipVia.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblShipVia, "lblShipVia")
        Me.lblShipVia.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShipVia.Name = "lblShipVia"
        '
        'lblTerms
        '
        Me.lblTerms.BackColor = System.Drawing.SystemColors.Control
        Me.lblTerms.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblTerms, "lblTerms")
        Me.lblTerms.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTerms.Name = "lblTerms"
        '
        'lblDiscPct
        '
        Me.lblDiscPct.BackColor = System.Drawing.SystemColors.Control
        Me.lblDiscPct.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblDiscPct, "lblDiscPct")
        Me.lblDiscPct.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDiscPct.Name = "lblDiscPct"
        '
        'lblProfitCenter
        '
        Me.lblProfitCenter.BackColor = System.Drawing.SystemColors.Control
        Me.lblProfitCenter.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblProfitCenter, "lblProfitCenter")
        Me.lblProfitCenter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblProfitCenter.Name = "lblProfitCenter"
        '
        'lblInstructions
        '
        Me.lblInstructions.BackColor = System.Drawing.SystemColors.Control
        Me.lblInstructions.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblInstructions, "lblInstructions")
        Me.lblInstructions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInstructions.Name = "lblInstructions"
        '
        'lblProject
        '
        Me.lblProject.BackColor = System.Drawing.SystemColors.Control
        Me.lblProject.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblProject, "lblProject")
        Me.lblProject.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblProject.Name = "lblProject"
        '
        'lblSalesPerson
        '
        Me.lblSalesPerson.BackColor = System.Drawing.SystemColors.Control
        Me.lblSalesPerson.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblSalesPerson, "lblSalesPerson")
        Me.lblSalesPerson.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSalesPerson.Name = "lblSalesPerson"
        '
        'lblPONumber
        '
        Me.lblPONumber.BackColor = System.Drawing.SystemColors.Control
        Me.lblPONumber.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblPONumber, "lblPONumber")
        Me.lblPONumber.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPONumber.Name = "lblPONumber"
        '
        'lblShipTo
        '
        Me.lblShipTo.BackColor = System.Drawing.SystemColors.Control
        Me.lblShipTo.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblShipTo, "lblShipTo")
        Me.lblShipTo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShipTo.Name = "lblShipTo"
        '
        'lblCustomer
        '
        Me.lblCustomer.BackColor = System.Drawing.SystemColors.Control
        Me.lblCustomer.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblCustomer, "lblCustomer")
        Me.lblCustomer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCustomer.Name = "lblCustomer"
        '
        'lblDate
        '
        Me.lblDate.BackColor = System.Drawing.SystemColors.Control
        Me.lblDate.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblDate, "lblDate")
        Me.lblDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDate.Name = "lblDate"
        '
        'lblOrder
        '
        Me.lblOrder.BackColor = System.Drawing.SystemColors.Control
        Me.lblOrder.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblOrder, "lblOrder")
        Me.lblOrder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrder.Name = "lblOrder"
        '
        'lblType
        '
        Me.lblType.BackColor = System.Drawing.SystemColors.Control
        Me.lblType.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblType, "lblType")
        Me.lblType.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblType.Name = "lblType"
        '
        'mcCalendar
        '
        Me.mcCalendar.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.mcCalendar, "mcCalendar")
        Me.mcCalendar.MaxSelectionCount = 1
        Me.mcCalendar.Name = "mcCalendar"
        '
        'lblCostCenter
        '
        Me.lblCostCenter.BackColor = System.Drawing.SystemColors.Control
        Me.lblCostCenter.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblCostCenter, "lblCostCenter")
        Me.lblCostCenter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCostCenter.Name = "lblCostCenter"
        '
        'lblComments
        '
        Me.lblComments.BackColor = System.Drawing.SystemColors.Control
        Me.lblComments.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblComments, "lblComments")
        Me.lblComments.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblComments.Name = "lblComments"
        '
        'cmdMfg_Loc_Notes
        '
        Me.cmdMfg_Loc_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdMfg_Loc_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdMfg_Loc_Notes, "cmdMfg_Loc_Notes")
        Me.cmdMfg_Loc_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMfg_Loc_Notes.Name = "cmdMfg_Loc_Notes"
        Me.cmdMfg_Loc_Notes.TabStop = False
        Me.cmdMfg_Loc_Notes.UseVisualStyleBackColor = False
        '
        'cmdOrd_No_Notes
        '
        Me.cmdOrd_No_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrd_No_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdOrd_No_Notes, "cmdOrd_No_Notes")
        Me.cmdOrd_No_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrd_No_Notes.Name = "cmdOrd_No_Notes"
        Me.cmdOrd_No_Notes.TabStop = False
        Me.cmdOrd_No_Notes.UseVisualStyleBackColor = False
        '
        'cmdCus_No_Notes
        '
        Me.cmdCus_No_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCus_No_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCus_No_Notes, "cmdCus_No_Notes")
        Me.cmdCus_No_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCus_No_Notes.Name = "cmdCus_No_Notes"
        Me.cmdCus_No_Notes.TabStop = False
        Me.cmdCus_No_Notes.UseVisualStyleBackColor = False
        '
        'cmdCus_Alt_Adr_Cd_Notes
        '
        Me.cmdCus_Alt_Adr_Cd_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCus_Alt_Adr_Cd_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCus_Alt_Adr_Cd_Notes, "cmdCus_Alt_Adr_Cd_Notes")
        Me.cmdCus_Alt_Adr_Cd_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCus_Alt_Adr_Cd_Notes.Name = "cmdCus_Alt_Adr_Cd_Notes"
        Me.cmdCus_Alt_Adr_Cd_Notes.TabStop = False
        Me.cmdCus_Alt_Adr_Cd_Notes.UseVisualStyleBackColor = False
        '
        'cmdSlspsn_No_Notes
        '
        Me.cmdSlspsn_No_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSlspsn_No_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdSlspsn_No_Notes, "cmdSlspsn_No_Notes")
        Me.cmdSlspsn_No_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSlspsn_No_Notes.Name = "cmdSlspsn_No_Notes"
        Me.cmdSlspsn_No_Notes.TabStop = False
        Me.cmdSlspsn_No_Notes.UseVisualStyleBackColor = False
        '
        'cmdJob_No_Notes
        '
        Me.cmdJob_No_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdJob_No_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdJob_No_Notes, "cmdJob_No_Notes")
        Me.cmdJob_No_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdJob_No_Notes.Name = "cmdJob_No_Notes"
        Me.cmdJob_No_Notes.TabStop = False
        Me.cmdJob_No_Notes.UseVisualStyleBackColor = False
        '
        'cmdShip_Via_Cd_Notes
        '
        Me.cmdShip_Via_Cd_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShip_Via_Cd_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdShip_Via_Cd_Notes, "cmdShip_Via_Cd_Notes")
        Me.cmdShip_Via_Cd_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShip_Via_Cd_Notes.Name = "cmdShip_Via_Cd_Notes"
        Me.cmdShip_Via_Cd_Notes.TabStop = False
        Me.cmdShip_Via_Cd_Notes.UseVisualStyleBackColor = False
        '
        'cmdAr_Terms_Cd_Notes
        '
        Me.cmdAr_Terms_Cd_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAr_Terms_Cd_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdAr_Terms_Cd_Notes, "cmdAr_Terms_Cd_Notes")
        Me.cmdAr_Terms_Cd_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAr_Terms_Cd_Notes.Name = "cmdAr_Terms_Cd_Notes"
        Me.cmdAr_Terms_Cd_Notes.TabStop = False
        Me.cmdAr_Terms_Cd_Notes.UseVisualStyleBackColor = False
        '
        'cmdProfit_Center_Notes
        '
        Me.cmdProfit_Center_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdProfit_Center_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdProfit_Center_Notes, "cmdProfit_Center_Notes")
        Me.cmdProfit_Center_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdProfit_Center_Notes.Name = "cmdProfit_Center_Notes"
        Me.cmdProfit_Center_Notes.TabStop = False
        Me.cmdProfit_Center_Notes.UseVisualStyleBackColor = False
        '
        'cmdBill_To_Country_Notes
        '
        Me.cmdBill_To_Country_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBill_To_Country_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdBill_To_Country_Notes, "cmdBill_To_Country_Notes")
        Me.cmdBill_To_Country_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBill_To_Country_Notes.Name = "cmdBill_To_Country_Notes"
        Me.cmdBill_To_Country_Notes.TabStop = False
        Me.cmdBill_To_Country_Notes.UseVisualStyleBackColor = False
        '
        'cmdShip_To_Country_Notes
        '
        Me.cmdShip_To_Country_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShip_To_Country_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdShip_To_Country_Notes, "cmdShip_To_Country_Notes")
        Me.cmdShip_To_Country_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShip_To_Country_Notes.Name = "cmdShip_To_Country_Notes"
        Me.cmdShip_To_Country_Notes.TabStop = False
        Me.cmdShip_To_Country_Notes.UseVisualStyleBackColor = False
        '
        'cmdConvertQuote
        '
        Me.cmdConvertQuote.BackColor = System.Drawing.SystemColors.Control
        Me.cmdConvertQuote.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdConvertQuote, "cmdConvertQuote")
        Me.cmdConvertQuote.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdConvertQuote.Name = "cmdConvertQuote"
        Me.cmdConvertQuote.TabStop = False
        Me.cmdConvertQuote.UseVisualStyleBackColor = False
        '
        'cmdPrintQuote
        '
        Me.cmdPrintQuote.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPrintQuote.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdPrintQuote, "cmdPrintQuote")
        Me.cmdPrintQuote.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPrintQuote.Name = "cmdPrintQuote"
        Me.cmdPrintQuote.TabStop = False
        Me.cmdPrintQuote.UseVisualStyleBackColor = False
        '
        'cmdHistoryCopy
        '
        Me.cmdHistoryCopy.BackColor = System.Drawing.SystemColors.Control
        Me.cmdHistoryCopy.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdHistoryCopy, "cmdHistoryCopy")
        Me.cmdHistoryCopy.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdHistoryCopy.Name = "cmdHistoryCopy"
        Me.cmdHistoryCopy.TabStop = False
        Me.cmdHistoryCopy.UseVisualStyleBackColor = False
        '
        'cmdMasterCopy
        '
        Me.cmdMasterCopy.BackColor = System.Drawing.SystemColors.Control
        Me.cmdMasterCopy.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdMasterCopy, "cmdMasterCopy")
        Me.cmdMasterCopy.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMasterCopy.Name = "cmdMasterCopy"
        Me.cmdMasterCopy.TabStop = False
        Me.cmdMasterCopy.UseVisualStyleBackColor = False
        '
        'cmdLineUserFields
        '
        Me.cmdLineUserFields.BackColor = System.Drawing.SystemColors.Control
        Me.cmdLineUserFields.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdLineUserFields, "cmdLineUserFields")
        Me.cmdLineUserFields.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLineUserFields.Name = "cmdLineUserFields"
        Me.cmdLineUserFields.TabStop = False
        Me.cmdLineUserFields.UseVisualStyleBackColor = False
        '
        'cmdInvoice
        '
        Me.cmdInvoice.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInvoice.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdInvoice, "cmdInvoice")
        Me.cmdInvoice.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInvoice.Name = "cmdInvoice"
        Me.cmdInvoice.TabStop = False
        Me.cmdInvoice.UseVisualStyleBackColor = False
        '
        'cmdPickTicket
        '
        Me.cmdPickTicket.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPickTicket.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdPickTicket, "cmdPickTicket")
        Me.cmdPickTicket.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPickTicket.Name = "cmdPickTicket"
        Me.cmdPickTicket.TabStop = False
        Me.cmdPickTicket.UseVisualStyleBackColor = False
        '
        'cmdAcknowledgement
        '
        Me.cmdAcknowledgement.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAcknowledgement.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdAcknowledgement, "cmdAcknowledgement")
        Me.cmdAcknowledgement.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAcknowledgement.Name = "cmdAcknowledgement"
        Me.cmdAcknowledgement.TabStop = False
        Me.cmdAcknowledgement.UseVisualStyleBackColor = False
        '
        'cmdComments
        '
        Me.cmdComments.BackColor = System.Drawing.SystemColors.Control
        Me.cmdComments.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdComments, "cmdComments")
        Me.cmdComments.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdComments.Name = "cmdComments"
        Me.cmdComments.TabStop = False
        Me.cmdComments.UseVisualStyleBackColor = False
        '
        'cmdLocQty
        '
        Me.cmdLocQty.BackColor = System.Drawing.SystemColors.Control
        Me.cmdLocQty.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdLocQty, "cmdLocQty")
        Me.cmdLocQty.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLocQty.Name = "cmdLocQty"
        Me.cmdLocQty.TabStop = False
        Me.cmdLocQty.UseVisualStyleBackColor = False
        '
        'cmdItemNotes
        '
        Me.cmdItemNotes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdItemNotes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdItemNotes, "cmdItemNotes")
        Me.cmdItemNotes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdItemNotes.Name = "cmdItemNotes"
        Me.cmdItemNotes.TabStop = False
        Me.cmdItemNotes.UseVisualStyleBackColor = False
        '
        'cmdItemInfo
        '
        Me.cmdItemInfo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdItemInfo.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdItemInfo, "cmdItemInfo")
        Me.cmdItemInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdItemInfo.Name = "cmdItemInfo"
        Me.cmdItemInfo.TabStop = False
        Me.cmdItemInfo.UseVisualStyleBackColor = False
        '
        'cmdCommissions
        '
        Me.cmdCommissions.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCommissions.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCommissions, "cmdCommissions")
        Me.cmdCommissions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCommissions.Name = "cmdCommissions"
        Me.cmdCommissions.TabStop = False
        Me.cmdCommissions.UseVisualStyleBackColor = False
        '
        'cmdTax
        '
        Me.cmdTax.BackColor = System.Drawing.SystemColors.Control
        Me.cmdTax.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdTax, "cmdTax")
        Me.cmdTax.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdTax.Name = "cmdTax"
        Me.cmdTax.TabStop = False
        Me.cmdTax.UseVisualStyleBackColor = False
        '
        'cmdEditCapturedBill
        '
        Me.cmdEditCapturedBill.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEditCapturedBill.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdEditCapturedBill, "cmdEditCapturedBill")
        Me.cmdEditCapturedBill.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEditCapturedBill.Name = "cmdEditCapturedBill"
        Me.cmdEditCapturedBill.TabStop = False
        Me.cmdEditCapturedBill.UseVisualStyleBackColor = False
        '
        'cmdATP
        '
        Me.cmdATP.BackColor = System.Drawing.SystemColors.Control
        Me.cmdATP.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdATP, "cmdATP")
        Me.cmdATP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdATP.Name = "cmdATP"
        Me.cmdATP.TabStop = False
        Me.cmdATP.UseVisualStyleBackColor = False
        '
        'lblUser_Def_Fld_2
        '
        Me.lblUser_Def_Fld_2.BackColor = System.Drawing.SystemColors.Control
        Me.lblUser_Def_Fld_2.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblUser_Def_Fld_2, "lblUser_Def_Fld_2")
        Me.lblUser_Def_Fld_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUser_Def_Fld_2.Name = "lblUser_Def_Fld_2"
        '
        'cboUser_Def_Fld_2
        '
        Me.cboUser_Def_Fld_2.BackColor = System.Drawing.SystemColors.Window
        Me.cboUser_Def_Fld_2.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboUser_Def_Fld_2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        resources.ApplyResources(Me.cboUser_Def_Fld_2, "cboUser_Def_Fld_2")
        Me.cboUser_Def_Fld_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboUser_Def_Fld_2.Items.AddRange(New Object() {resources.GetString("cboUser_Def_Fld_2.Items"), resources.GetString("cboUser_Def_Fld_2.Items1"), resources.GetString("cboUser_Def_Fld_2.Items2"), resources.GetString("cboUser_Def_Fld_2.Items3"), resources.GetString("cboUser_Def_Fld_2.Items4"), resources.GetString("cboUser_Def_Fld_2.Items5"), resources.GetString("cboUser_Def_Fld_2.Items6"), resources.GetString("cboUser_Def_Fld_2.Items7"), resources.GetString("cboUser_Def_Fld_2.Items8"), resources.GetString("cboUser_Def_Fld_2.Items9"), resources.GetString("cboUser_Def_Fld_2.Items10"), resources.GetString("cboUser_Def_Fld_2.Items11"), resources.GetString("cboUser_Def_Fld_2.Items12"), resources.GetString("cboUser_Def_Fld_2.Items13"), resources.GetString("cboUser_Def_Fld_2.Items14"), resources.GetString("cboUser_Def_Fld_2.Items15"), resources.GetString("cboUser_Def_Fld_2.Items16"), resources.GetString("cboUser_Def_Fld_2.Items17"), resources.GetString("cboUser_Def_Fld_2.Items18"), resources.GetString("cboUser_Def_Fld_2.Items19"), resources.GetString("cboUser_Def_Fld_2.Items20"), resources.GetString("cboUser_Def_Fld_2.Items21"), resources.GetString("cboUser_Def_Fld_2.Items22"), resources.GetString("cboUser_Def_Fld_2.Items23"), resources.GetString("cboUser_Def_Fld_2.Items24"), resources.GetString("cboUser_Def_Fld_2.Items25"), resources.GetString("cboUser_Def_Fld_2.Items26"), resources.GetString("cboUser_Def_Fld_2.Items27"), resources.GetString("cboUser_Def_Fld_2.Items28"), resources.GetString("cboUser_Def_Fld_2.Items29"), resources.GetString("cboUser_Def_Fld_2.Items30")})
        Me.cboUser_Def_Fld_2.Name = "cboUser_Def_Fld_2"
        '
        'lblUser_Def_Fld_3
        '
        Me.lblUser_Def_Fld_3.BackColor = System.Drawing.SystemColors.Control
        Me.lblUser_Def_Fld_3.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblUser_Def_Fld_3, "lblUser_Def_Fld_3")
        Me.lblUser_Def_Fld_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUser_Def_Fld_3.Name = "lblUser_Def_Fld_3"
        '
        'lblUser_Def_Fld_1
        '
        Me.lblUser_Def_Fld_1.BackColor = System.Drawing.SystemColors.Control
        Me.lblUser_Def_Fld_1.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblUser_Def_Fld_1, "lblUser_Def_Fld_1")
        Me.lblUser_Def_Fld_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUser_Def_Fld_1.Name = "lblUser_Def_Fld_1"
        '
        'lblUser_Def_Fld_4
        '
        Me.lblUser_Def_Fld_4.BackColor = System.Drawing.SystemColors.Control
        Me.lblUser_Def_Fld_4.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblUser_Def_Fld_4, "lblUser_Def_Fld_4")
        Me.lblUser_Def_Fld_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUser_Def_Fld_4.Name = "lblUser_Def_Fld_4"
        '
        'lblContactEmail
        '
        Me.lblContactEmail.BackColor = System.Drawing.SystemColors.Control
        Me.lblContactEmail.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblContactEmail, "lblContactEmail")
        Me.lblContactEmail.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblContactEmail.Name = "lblContactEmail"
        '
        'cmdUser_Def_Fld_4Search
        '
        Me.cmdUser_Def_Fld_4Search.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUser_Def_Fld_4Search.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdUser_Def_Fld_4Search, "cmdUser_Def_Fld_4Search")
        Me.cmdUser_Def_Fld_4Search.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUser_Def_Fld_4Search.Name = "cmdUser_Def_Fld_4Search"
        Me.cmdUser_Def_Fld_4Search.TabStop = False
        Me.cmdUser_Def_Fld_4Search.UseVisualStyleBackColor = False
        '
        'cmdOEFlash
        '
        Me.cmdOEFlash.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOEFlash.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdOEFlash, "cmdOEFlash")
        Me.cmdOEFlash.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOEFlash.Name = "cmdOEFlash"
        Me.cmdOEFlash.TabStop = False
        Me.cmdOEFlash.UseVisualStyleBackColor = False
        '
        'cmdOrderLink
        '
        Me.cmdOrderLink.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrderLink.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdOrderLink, "cmdOrderLink")
        Me.cmdOrderLink.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrderLink.Name = "cmdOrderLink"
        Me.cmdOrderLink.TabStop = False
        Me.cmdOrderLink.UseVisualStyleBackColor = False
        '
        'cmdCopyDate
        '
        Me.cmdCopyDate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopyDate.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCopyDate, "cmdCopyDate")
        Me.cmdCopyDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopyDate.Name = "cmdCopyDate"
        Me.cmdCopyDate.TabStop = False
        Me.cmdCopyDate.UseVisualStyleBackColor = False
        '
        'cmdUser_Def_Fld_5Search
        '
        Me.cmdUser_Def_Fld_5Search.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUser_Def_Fld_5Search.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdUser_Def_Fld_5Search, "cmdUser_Def_Fld_5Search")
        Me.cmdUser_Def_Fld_5Search.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUser_Def_Fld_5Search.Name = "cmdUser_Def_Fld_5Search"
        Me.cmdUser_Def_Fld_5Search.TabStop = False
        Me.cmdUser_Def_Fld_5Search.UseVisualStyleBackColor = False
        '
        'lblUser_Def_Fld_5
        '
        Me.lblUser_Def_Fld_5.BackColor = System.Drawing.SystemColors.Control
        Me.lblUser_Def_Fld_5.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblUser_Def_Fld_5, "lblUser_Def_Fld_5")
        Me.lblUser_Def_Fld_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUser_Def_Fld_5.Name = "lblUser_Def_Fld_5"
        '
        'cmdShip_LinkEdit
        '
        Me.cmdShip_LinkEdit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShip_LinkEdit.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdShip_LinkEdit, "cmdShip_LinkEdit")
        Me.cmdShip_LinkEdit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShip_LinkEdit.Name = "cmdShip_LinkEdit"
        Me.cmdShip_LinkEdit.TabStop = False
        Me.cmdShip_LinkEdit.UseVisualStyleBackColor = False
        '
        'cmdShip_LinkSearch
        '
        Me.cmdShip_LinkSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShip_LinkSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdShip_LinkSearch, "cmdShip_LinkSearch")
        Me.cmdShip_LinkSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShip_LinkSearch.Name = "cmdShip_LinkSearch"
        Me.cmdShip_LinkSearch.TabStop = False
        Me.cmdShip_LinkSearch.UseVisualStyleBackColor = False
        '
        'lblShip_Link
        '
        Me.lblShip_Link.BackColor = System.Drawing.SystemColors.Control
        Me.lblShip_Link.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblShip_Link, "lblShip_Link")
        Me.lblShip_Link.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShip_Link.Name = "lblShip_Link"
        '
        'cmdUser_Def_Fld_3Search
        '
        Me.cmdUser_Def_Fld_3Search.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUser_Def_Fld_3Search.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdUser_Def_Fld_3Search, "cmdUser_Def_Fld_3Search")
        Me.cmdUser_Def_Fld_3Search.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUser_Def_Fld_3Search.Name = "cmdUser_Def_Fld_3Search"
        Me.cmdUser_Def_Fld_3Search.TabStop = False
        Me.cmdUser_Def_Fld_3Search.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.Control
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Button1, "Button1")
        Me.Button1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button1.Name = "Button1"
        Me.Button1.TabStop = False
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        resources.ApplyResources(Me.Button2, "Button2")
        Me.Button2.Name = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        resources.ApplyResources(Me.Button3, "Button3")
        Me.Button3.Name = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        resources.ApplyResources(Me.Button4, "Button4")
        Me.Button4.Name = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        resources.ApplyResources(Me.Button5, "Button5")
        Me.Button5.Name = "Button5"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'cmdShip_To_NameSearch
        '
        Me.cmdShip_To_NameSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShip_To_NameSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdShip_To_NameSearch, "cmdShip_To_NameSearch")
        Me.cmdShip_To_NameSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShip_To_NameSearch.Name = "cmdShip_To_NameSearch"
        Me.cmdShip_To_NameSearch.TabStop = False
        Me.cmdShip_To_NameSearch.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Name = "Label1"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Name = "Label2"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Name = "Label3"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Name = "Label4"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Name = "Label5"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Name = "Label6"
        '
        'chkOrderAckSaveOnly
        '
        resources.ApplyResources(Me.chkOrderAckSaveOnly, "chkOrderAckSaveOnly")
        Me.chkOrderAckSaveOnly.Name = "chkOrderAckSaveOnly"
        Me.chkOrderAckSaveOnly.UseVisualStyleBackColor = True
        '
        'lblOptions
        '
        Me.lblOptions.BackColor = System.Drawing.SystemColors.Control
        Me.lblOptions.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblOptions, "lblOptions")
        Me.lblOptions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOptions.Name = "lblOptions"
        '
        'chkRecover
        '
        resources.ApplyResources(Me.chkRecover, "chkRecover")
        Me.chkRecover.Name = "chkRecover"
        Me.chkRecover.TabStop = False
        Me.chkRecover.UseVisualStyleBackColor = True
        '
        'chkRedo
        '
        resources.ApplyResources(Me.chkRedo, "chkRedo")
        Me.chkRedo.Name = "chkRedo"
        Me.chkRedo.TabStop = False
        Me.chkRedo.UseVisualStyleBackColor = True
        '
        'chkExactRepeat
        '
        resources.ApplyResources(Me.chkExactRepeat, "chkExactRepeat")
        Me.chkExactRepeat.Name = "chkExactRepeat"
        Me.chkExactRepeat.TabStop = False
        Me.chkExactRepeat.UseVisualStyleBackColor = True
        '
        'chkInvalidShipDateEmail
        '
        resources.ApplyResources(Me.chkInvalidShipDateEmail, "chkInvalidShipDateEmail")
        Me.chkInvalidShipDateEmail.Name = "chkInvalidShipDateEmail"
        Me.chkInvalidShipDateEmail.UseVisualStyleBackColor = True
        '
        'chkEmail2
        '
        resources.ApplyResources(Me.chkEmail2, "chkEmail2")
        Me.chkEmail2.Name = "chkEmail2"
        Me.chkEmail2.UseVisualStyleBackColor = True
        '
        'cboProgram
        '
        Me.cboProgram.BackColor = System.Drawing.SystemColors.Window
        Me.cboProgram.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboProgram.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        resources.ApplyResources(Me.cboProgram, "cboProgram")
        Me.cboProgram.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboProgram.Items.AddRange(New Object() {resources.GetString("cboProgram.Items"), resources.GetString("cboProgram.Items1"), resources.GetString("cboProgram.Items2"), resources.GetString("cboProgram.Items3"), resources.GetString("cboProgram.Items4"), resources.GetString("cboProgram.Items5"), resources.GetString("cboProgram.Items6"), resources.GetString("cboProgram.Items7"), resources.GetString("cboProgram.Items8"), resources.GetString("cboProgram.Items9"), resources.GetString("cboProgram.Items10"), resources.GetString("cboProgram.Items11"), resources.GetString("cboProgram.Items12"), resources.GetString("cboProgram.Items13"), resources.GetString("cboProgram.Items14"), resources.GetString("cboProgram.Items15"), resources.GetString("cboProgram.Items16"), resources.GetString("cboProgram.Items17"), resources.GetString("cboProgram.Items18"), resources.GetString("cboProgram.Items19"), resources.GetString("cboProgram.Items20"), resources.GetString("cboProgram.Items21"), resources.GetString("cboProgram.Items22")})
        Me.cboProgram.Name = "cboProgram"
        '
        'lblProgram
        '
        Me.lblProgram.BackColor = System.Drawing.SystemColors.Control
        Me.lblProgram.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblProgram, "lblProgram")
        Me.lblProgram.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblProgram.Name = "lblProgram"
        '
        'cmdOrderRequest
        '
        Me.cmdOrderRequest.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrderRequest.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdOrderRequest, "cmdOrderRequest")
        Me.cmdOrderRequest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrderRequest.Name = "cmdOrderRequest"
        Me.cmdOrderRequest.TabStop = False
        Me.cmdOrderRequest.UseVisualStyleBackColor = False
        '
        'cmdSendEmail
        '
        resources.ApplyResources(Me.cmdSendEmail, "cmdSendEmail")
        Me.cmdSendEmail.Name = "cmdSendEmail"
        Me.cmdSendEmail.UseVisualStyleBackColor = True
        '
        'cmdEDI
        '
        Me.cmdEDI.BackColor = System.Drawing.Color.Red
        resources.ApplyResources(Me.cmdEDI, "cmdEDI")
        Me.cmdEDI.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdEDI.Name = "cmdEDI"
        Me.cmdEDI.UseVisualStyleBackColor = False
        '
        'Button6
        '
        resources.ApplyResources(Me.Button6, "Button6")
        Me.Button6.Name = "Button6"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'cmdProg_Spector_CdSearch
        '
        Me.cmdProg_Spector_CdSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdProg_Spector_CdSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdProg_Spector_CdSearch, "cmdProg_Spector_CdSearch")
        Me.cmdProg_Spector_CdSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdProg_Spector_CdSearch.Name = "cmdProg_Spector_CdSearch"
        Me.cmdProg_Spector_CdSearch.TabStop = False
        Me.cmdProg_Spector_CdSearch.UseVisualStyleBackColor = False
        '
        'cmdCustomerDash
        '
        Me.cmdCustomerDash.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCustomerDash.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdCustomerDash, "cmdCustomerDash")
        Me.cmdCustomerDash.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCustomerDash.Name = "cmdCustomerDash"
        Me.cmdCustomerDash.TabStop = False
        Me.cmdCustomerDash.UseVisualStyleBackColor = False
        '
        'lblEnd_User
        '
        Me.lblEnd_User.BackColor = System.Drawing.SystemColors.Control
        Me.lblEnd_User.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblEnd_User, "lblEnd_User")
        Me.lblEnd_User.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEnd_User.Name = "lblEnd_User"
        '
        'cmdEnd_UserSearch
        '
        Me.cmdEnd_UserSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEnd_UserSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdEnd_UserSearch, "cmdEnd_UserSearch")
        Me.cmdEnd_UserSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnd_UserSearch.Name = "cmdEnd_UserSearch"
        Me.cmdEnd_UserSearch.TabStop = False
        Me.cmdEnd_UserSearch.UseVisualStyleBackColor = False
        '
        'cmdEnd_User_Notes
        '
        Me.cmdEnd_User_Notes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEnd_User_Notes.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdEnd_User_Notes, "cmdEnd_User_Notes")
        Me.cmdEnd_User_Notes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnd_User_Notes.Name = "cmdEnd_User_Notes"
        Me.cmdEnd_User_Notes.TabStop = False
        Me.cmdEnd_User_Notes.UseVisualStyleBackColor = False
        '
        'cmdInHandDate
        '
        Me.cmdInHandDate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInHandDate.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdInHandDate, "cmdInHandDate")
        Me.cmdInHandDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInHandDate.Name = "cmdInHandDate"
        Me.cmdInHandDate.TabStop = False
        Me.cmdInHandDate.UseVisualStyleBackColor = False
        '
        'lblInHandsDate
        '
        Me.lblInHandsDate.BackColor = System.Drawing.SystemColors.Control
        Me.lblInHandsDate.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblInHandsDate, "lblInHandsDate")
        Me.lblInHandsDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInHandsDate.Name = "lblInHandsDate"
        '
        'cmdApplyToSearch
        '
        Me.cmdApplyToSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdApplyToSearch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdApplyToSearch, "cmdApplyToSearch")
        Me.cmdApplyToSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdApplyToSearch.Name = "cmdApplyToSearch"
        Me.cmdApplyToSearch.TabStop = False
        Me.cmdApplyToSearch.UseVisualStyleBackColor = False
        '
        'lblApplyTo
        '
        Me.lblApplyTo.BackColor = System.Drawing.SystemColors.Control
        Me.lblApplyTo.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblApplyTo, "lblApplyTo")
        Me.lblApplyTo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblApplyTo.Name = "lblApplyTo"
        '
        'cmdShipDateTP
        '
        Me.cmdShipDateTP.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShipDateTP.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdShipDateTP, "cmdShipDateTP")
        Me.cmdShipDateTP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShipDateTP.Name = "cmdShipDateTP"
        Me.cmdShipDateTP.TabStop = False
        Me.cmdShipDateTP.UseVisualStyleBackColor = False
        '
        'lblShipDate
        '
        Me.lblShipDate.BackColor = System.Drawing.SystemColors.Control
        Me.lblShipDate.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblShipDate, "lblShipDate")
        Me.lblShipDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShipDate.Name = "lblShipDate"
        '
        'cmdXml
        '
        Me.cmdXml.BackColor = System.Drawing.SystemColors.Control
        resources.ApplyResources(Me.cmdXml, "cmdXml")
        Me.cmdXml.Name = "cmdXml"
        Me.cmdXml.UseVisualStyleBackColor = False
        '
        'cmdApplyToNote
        '
        Me.cmdApplyToNote.BackColor = System.Drawing.SystemColors.Control
        Me.cmdApplyToNote.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.cmdApplyToNote, "cmdApplyToNote")
        Me.cmdApplyToNote.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdApplyToNote.Name = "cmdApplyToNote"
        Me.cmdApplyToNote.TabStop = False
        Me.cmdApplyToNote.UseVisualStyleBackColor = False
        '
        'cbo_Extra_9
        '
        Me.cbo_Extra_9.BackColor = System.Drawing.SystemColors.Window
        Me.cbo_Extra_9.Cursor = System.Windows.Forms.Cursors.Default
        Me.cbo_Extra_9.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        resources.ApplyResources(Me.cbo_Extra_9, "cbo_Extra_9")
        Me.cbo_Extra_9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cbo_Extra_9.Name = "cbo_Extra_9"
        '
        'lblExtra_9
        '
        Me.lblExtra_9.BackColor = System.Drawing.SystemColors.Control
        Me.lblExtra_9.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblExtra_9, "lblExtra_9")
        Me.lblExtra_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblExtra_9.Name = "lblExtra_9"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Crimson
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.Label7, "Label7")
        Me.Label7.ForeColor = System.Drawing.SystemColors.Control
        Me.Label7.Name = "Label7"
        '
        'cboWhyNotVegas
        '
        Me.cboWhyNotVegas.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cboWhyNotVegas.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboWhyNotVegas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        resources.ApplyResources(Me.cboWhyNotVegas, "cboWhyNotVegas")
        Me.cboWhyNotVegas.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboWhyNotVegas.Name = "cboWhyNotVegas"
        '
        'cboreSSP
        '
        Me.cboreSSP.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cboreSSP.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboreSSP.DropDownHeight = 150
        Me.cboreSSP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        resources.ApplyResources(Me.cboreSSP, "cboreSSP")
        Me.cboreSSP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboreSSP.Name = "cboreSSP"
        '
        'lblSSp
        '
        Me.lblSSp.BackColor = System.Drawing.Color.Crimson
        Me.lblSSp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSSp.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblSSp, "lblSSp")
        Me.lblSSp.ForeColor = System.Drawing.Color.White
        Me.lblSSp.Name = "lblSSp"
        '
        'lblInvBatch
        '
        Me.lblInvBatch.BackColor = System.Drawing.SystemColors.Control
        Me.lblInvBatch.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.lblInvBatch, "lblInvBatch")
        Me.lblInvBatch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInvBatch.Name = "lblInvBatch"
        '
        'xTxtInvBatch
        '
        Me.xTxtInvBatch.CharacterInput = Orders.Cinput.NumericOnly
        Me.xTxtInvBatch.DataLength = CType(8, Long)
        Me.xTxtInvBatch.DataType = Orders.CDataType.dtString
        Me.xTxtInvBatch.DateValue = New Date(CType(0, Long))
        Me.xTxtInvBatch.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.xTxtInvBatch, "xTxtInvBatch")
        Me.xTxtInvBatch.Mask = Nothing
        Me.xTxtInvBatch.Name = "xTxtInvBatch"
        Me.xTxtInvBatch.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.xTxtInvBatch.OldValue = Nothing
        Me.xTxtInvBatch.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.xTxtInvBatch.StringValue = Nothing
        Me.xTxtInvBatch.TextDataID = Nothing
        '
        'txtCustGroup
        '
        Me.txtCustGroup.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtCustGroup.DataLength = CType(20, Long)
        Me.txtCustGroup.DataType = Orders.CDataType.dtString
        Me.txtCustGroup.DateValue = New Date(CType(0, Long))
        Me.txtCustGroup.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtCustGroup, "txtCustGroup")
        Me.txtCustGroup.Mask = Nothing
        Me.txtCustGroup.Name = "txtCustGroup"
        Me.txtCustGroup.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtCustGroup.OldValue = Nothing
        Me.txtCustGroup.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtCustGroup.StringValue = Nothing
        Me.txtCustGroup.TextDataID = Nothing
        '
        'txtOrd_Dt_Shipped
        '
        Me.txtOrd_Dt_Shipped.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtOrd_Dt_Shipped.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOrd_Dt_Shipped.DataLength = CType(10, Long)
        Me.txtOrd_Dt_Shipped.DataType = Orders.CDataType.dtString
        Me.txtOrd_Dt_Shipped.DateValue = New Date(CType(0, Long))
        Me.txtOrd_Dt_Shipped.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtOrd_Dt_Shipped, "txtOrd_Dt_Shipped")
        Me.txtOrd_Dt_Shipped.Mask = Nothing
        Me.txtOrd_Dt_Shipped.Name = "txtOrd_Dt_Shipped"
        Me.txtOrd_Dt_Shipped.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtOrd_Dt_Shipped.OldValue = Nothing
        Me.txtOrd_Dt_Shipped.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtOrd_Dt_Shipped.StringValue = Nothing
        Me.txtOrd_Dt_Shipped.TextDataID = Nothing
        '
        'txtInHandsDate
        '
        Me.txtInHandsDate.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtInHandsDate.DataLength = CType(0, Long)
        Me.txtInHandsDate.DataType = Orders.CDataType.dtString
        Me.txtInHandsDate.DateValue = New Date(CType(0, Long))
        Me.txtInHandsDate.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtInHandsDate, "txtInHandsDate")
        Me.txtInHandsDate.Mask = Nothing
        Me.txtInHandsDate.Name = "txtInHandsDate"
        Me.txtInHandsDate.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtInHandsDate.OldValue = ""
        Me.txtInHandsDate.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtInHandsDate.StringValue = Nothing
        Me.txtInHandsDate.TextDataID = Nothing
        '
        'txtApplyTo
        '
        Me.txtApplyTo.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtApplyTo.DataLength = CType(8, Long)
        Me.txtApplyTo.DataType = Orders.CDataType.dtString
        Me.txtApplyTo.DateValue = New Date(CType(0, Long))
        Me.txtApplyTo.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtApplyTo, "txtApplyTo")
        Me.txtApplyTo.Mask = Nothing
        Me.txtApplyTo.Name = "txtApplyTo"
        Me.txtApplyTo.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtApplyTo.OldValue = Nothing
        Me.txtApplyTo.SpacePadding = Orders.CSpacePadding.PaddingBefore
        Me.txtApplyTo.StringValue = Nothing
        Me.txtApplyTo.TextDataID = Nothing
        '
        'txtDept
        '
        Me.txtDept.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtDept.DataLength = CType(8, Long)
        Me.txtDept.DataType = Orders.CDataType.dtString
        Me.txtDept.DateValue = New Date(CType(0, Long))
        Me.txtDept.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtDept, "txtDept")
        Me.txtDept.Mask = Nothing
        Me.txtDept.Name = "txtDept"
        Me.txtDept.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtDept.OldValue = Nothing
        Me.txtDept.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtDept.StringValue = Nothing
        Me.txtDept.TextDataID = Nothing
        '
        'txtCus_Prog_ID
        '
        Me.txtCus_Prog_ID.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtCus_Prog_ID.DataLength = CType(20, Long)
        Me.txtCus_Prog_ID.DataType = Orders.CDataType.dtString
        Me.txtCus_Prog_ID.DateValue = New Date(CType(0, Long))
        Me.txtCus_Prog_ID.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtCus_Prog_ID, "txtCus_Prog_ID")
        Me.txtCus_Prog_ID.Mask = Nothing
        Me.txtCus_Prog_ID.Name = "txtCus_Prog_ID"
        Me.txtCus_Prog_ID.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtCus_Prog_ID.OldValue = Nothing
        Me.txtCus_Prog_ID.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtCus_Prog_ID.StringValue = Nothing
        Me.txtCus_Prog_ID.TextDataID = Nothing
        '
        'txtProg_Spector_Cd
        '
        Me.txtProg_Spector_Cd.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtProg_Spector_Cd.DataLength = CType(20, Long)
        Me.txtProg_Spector_Cd.DataType = Orders.CDataType.dtString
        Me.txtProg_Spector_Cd.DateValue = New Date(CType(0, Long))
        Me.txtProg_Spector_Cd.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtProg_Spector_Cd, "txtProg_Spector_Cd")
        Me.txtProg_Spector_Cd.Mask = Nothing
        Me.txtProg_Spector_Cd.Name = "txtProg_Spector_Cd"
        Me.txtProg_Spector_Cd.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtProg_Spector_Cd.OldValue = Nothing
        Me.txtProg_Spector_Cd.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtProg_Spector_Cd.StringValue = Nothing
        Me.txtProg_Spector_Cd.TextDataID = Nothing
        '
        'txtShip_To_Zip
        '
        Me.txtShip_To_Zip.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_To_Zip.DataLength = CType(20, Long)
        Me.txtShip_To_Zip.DataType = Orders.CDataType.dtString
        Me.txtShip_To_Zip.DateValue = New Date(CType(0, Long))
        Me.txtShip_To_Zip.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_To_Zip, "txtShip_To_Zip")
        Me.txtShip_To_Zip.Mask = Nothing
        Me.txtShip_To_Zip.Name = "txtShip_To_Zip"
        Me.txtShip_To_Zip.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_To_Zip.OldValue = Nothing
        Me.txtShip_To_Zip.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_To_Zip.StringValue = Nothing
        Me.txtShip_To_Zip.TextDataID = Nothing
        '
        'txtShip_To_State
        '
        Me.txtShip_To_State.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_To_State.DataLength = CType(3, Long)
        Me.txtShip_To_State.DataType = Orders.CDataType.dtString
        Me.txtShip_To_State.DateValue = New Date(CType(0, Long))
        Me.txtShip_To_State.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_To_State, "txtShip_To_State")
        Me.txtShip_To_State.Mask = Nothing
        Me.txtShip_To_State.Name = "txtShip_To_State"
        Me.txtShip_To_State.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_To_State.OldValue = Nothing
        Me.txtShip_To_State.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_To_State.StringValue = Nothing
        Me.txtShip_To_State.TextDataID = Nothing
        '
        'txtShip_To_City
        '
        Me.txtShip_To_City.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_To_City.DataLength = CType(30, Long)
        Me.txtShip_To_City.DataType = Orders.CDataType.dtString
        Me.txtShip_To_City.DateValue = New Date(CType(0, Long))
        Me.txtShip_To_City.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_To_City, "txtShip_To_City")
        Me.txtShip_To_City.Mask = Nothing
        Me.txtShip_To_City.Name = "txtShip_To_City"
        Me.txtShip_To_City.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_To_City.OldValue = Nothing
        Me.txtShip_To_City.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_To_City.StringValue = Nothing
        Me.txtShip_To_City.TextDataID = Nothing
        '
        'txtShip_To_Addr_4
        '
        Me.txtShip_To_Addr_4.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_To_Addr_4.DataLength = CType(40, Long)
        Me.txtShip_To_Addr_4.DataType = Orders.CDataType.dtString
        Me.txtShip_To_Addr_4.DateValue = New Date(CType(0, Long))
        Me.txtShip_To_Addr_4.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_To_Addr_4, "txtShip_To_Addr_4")
        Me.txtShip_To_Addr_4.Mask = Nothing
        Me.txtShip_To_Addr_4.Name = "txtShip_To_Addr_4"
        Me.txtShip_To_Addr_4.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_To_Addr_4.OldValue = Nothing
        Me.txtShip_To_Addr_4.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_To_Addr_4.StringValue = Nothing
        Me.txtShip_To_Addr_4.TextDataID = Nothing
        '
        'txtBill_To_Zip
        '
        Me.txtBill_To_Zip.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBill_To_Zip.DataLength = CType(20, Long)
        Me.txtBill_To_Zip.DataType = Orders.CDataType.dtString
        Me.txtBill_To_Zip.DateValue = New Date(CType(0, Long))
        Me.txtBill_To_Zip.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtBill_To_Zip, "txtBill_To_Zip")
        Me.txtBill_To_Zip.Mask = Nothing
        Me.txtBill_To_Zip.Name = "txtBill_To_Zip"
        Me.txtBill_To_Zip.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBill_To_Zip.OldValue = Nothing
        Me.txtBill_To_Zip.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBill_To_Zip.StringValue = Nothing
        Me.txtBill_To_Zip.TextDataID = Nothing
        '
        'txtBill_To_City
        '
        Me.txtBill_To_City.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBill_To_City.DataLength = CType(30, Long)
        Me.txtBill_To_City.DataType = Orders.CDataType.dtString
        Me.txtBill_To_City.DateValue = New Date(CType(0, Long))
        Me.txtBill_To_City.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtBill_To_City, "txtBill_To_City")
        Me.txtBill_To_City.Mask = Nothing
        Me.txtBill_To_City.Name = "txtBill_To_City"
        Me.txtBill_To_City.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBill_To_City.OldValue = Nothing
        Me.txtBill_To_City.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBill_To_City.StringValue = Nothing
        Me.txtBill_To_City.TextDataID = Nothing
        '
        'txtBill_To_State
        '
        Me.txtBill_To_State.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBill_To_State.DataLength = CType(3, Long)
        Me.txtBill_To_State.DataType = Orders.CDataType.dtString
        Me.txtBill_To_State.DateValue = New Date(CType(0, Long))
        Me.txtBill_To_State.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtBill_To_State, "txtBill_To_State")
        Me.txtBill_To_State.Mask = Nothing
        Me.txtBill_To_State.Name = "txtBill_To_State"
        Me.txtBill_To_State.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBill_To_State.OldValue = Nothing
        Me.txtBill_To_State.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBill_To_State.StringValue = Nothing
        Me.txtBill_To_State.TextDataID = Nothing
        '
        'txtBill_To_Addr_4
        '
        Me.txtBill_To_Addr_4.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBill_To_Addr_4.DataLength = CType(40, Long)
        Me.txtBill_To_Addr_4.DataType = Orders.CDataType.dtString
        Me.txtBill_To_Addr_4.DateValue = New Date(CType(0, Long))
        Me.txtBill_To_Addr_4.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtBill_To_Addr_4, "txtBill_To_Addr_4")
        Me.txtBill_To_Addr_4.Mask = Nothing
        Me.txtBill_To_Addr_4.Name = "txtBill_To_Addr_4"
        Me.txtBill_To_Addr_4.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBill_To_Addr_4.OldValue = Nothing
        Me.txtBill_To_Addr_4.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBill_To_Addr_4.StringValue = Nothing
        Me.txtBill_To_Addr_4.TextDataID = Nothing
        '
        'txtShip_Link
        '
        Me.txtShip_Link.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_Link.DataLength = CType(25, Long)
        Me.txtShip_Link.DataType = Orders.CDataType.dtString
        Me.txtShip_Link.DateValue = New Date(CType(0, Long))
        Me.txtShip_Link.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_Link, "txtShip_Link")
        Me.txtShip_Link.Mask = Nothing
        Me.txtShip_Link.Name = "txtShip_Link"
        Me.txtShip_Link.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_Link.OldValue = Nothing
        Me.txtShip_Link.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_Link.StringValue = Nothing
        Me.txtShip_Link.TextDataID = Nothing
        '
        'txtUser_Def_Fld_5
        '
        Me.txtUser_Def_Fld_5.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtUser_Def_Fld_5.DataLength = CType(30, Long)
        Me.txtUser_Def_Fld_5.DataType = Orders.CDataType.dtString
        Me.txtUser_Def_Fld_5.DateValue = New Date(CType(0, Long))
        Me.txtUser_Def_Fld_5.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtUser_Def_Fld_5, "txtUser_Def_Fld_5")
        Me.txtUser_Def_Fld_5.Mask = Nothing
        Me.txtUser_Def_Fld_5.Name = "txtUser_Def_Fld_5"
        Me.txtUser_Def_Fld_5.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtUser_Def_Fld_5.OldValue = Nothing
        Me.txtUser_Def_Fld_5.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtUser_Def_Fld_5.StringValue = Nothing
        Me.txtUser_Def_Fld_5.TextDataID = Nothing
        '
        'txtContactEmail
        '
        Me.txtContactEmail.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtContactEmail.DataLength = CType(128, Long)
        Me.txtContactEmail.DataType = Orders.CDataType.dtString
        Me.txtContactEmail.DateValue = New Date(CType(0, Long))
        Me.txtContactEmail.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtContactEmail, "txtContactEmail")
        Me.txtContactEmail.Mask = Nothing
        Me.txtContactEmail.Name = "txtContactEmail"
        Me.txtContactEmail.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtContactEmail.OldValue = Nothing
        Me.txtContactEmail.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtContactEmail.StringValue = Nothing
        Me.txtContactEmail.TextDataID = Nothing
        '
        'txtUser_Def_Fld_4
        '
        Me.txtUser_Def_Fld_4.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtUser_Def_Fld_4.DataLength = CType(20, Long)
        Me.txtUser_Def_Fld_4.DataType = Orders.CDataType.dtString
        Me.txtUser_Def_Fld_4.DateValue = New Date(CType(0, Long))
        Me.txtUser_Def_Fld_4.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtUser_Def_Fld_4, "txtUser_Def_Fld_4")
        Me.txtUser_Def_Fld_4.Mask = Nothing
        Me.txtUser_Def_Fld_4.Name = "txtUser_Def_Fld_4"
        Me.txtUser_Def_Fld_4.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtUser_Def_Fld_4.OldValue = Nothing
        Me.txtUser_Def_Fld_4.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtUser_Def_Fld_4.StringValue = Nothing
        Me.txtUser_Def_Fld_4.TextDataID = Nothing
        '
        'txtUser_Def_Fld_3
        '
        Me.txtUser_Def_Fld_3.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.txtUser_Def_Fld_3.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.txtUser_Def_Fld_3.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtUser_Def_Fld_3.DataLength = CType(30, Long)
        Me.txtUser_Def_Fld_3.DataType = Orders.CDataType.dtString
        Me.txtUser_Def_Fld_3.DateValue = New Date(CType(0, Long))
        Me.txtUser_Def_Fld_3.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtUser_Def_Fld_3, "txtUser_Def_Fld_3")
        Me.txtUser_Def_Fld_3.Mask = Nothing
        Me.txtUser_Def_Fld_3.Name = "txtUser_Def_Fld_3"
        Me.txtUser_Def_Fld_3.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtUser_Def_Fld_3.OldValue = Nothing
        Me.txtUser_Def_Fld_3.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtUser_Def_Fld_3.StringValue = Nothing
        Me.txtUser_Def_Fld_3.TextDataID = Nothing
        '
        'txtUser_Def_Fld_1
        '
        Me.txtUser_Def_Fld_1.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtUser_Def_Fld_1.DataLength = CType(30, Long)
        Me.txtUser_Def_Fld_1.DataType = Orders.CDataType.dtString
        Me.txtUser_Def_Fld_1.DateValue = New Date(CType(0, Long))
        Me.txtUser_Def_Fld_1.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtUser_Def_Fld_1, "txtUser_Def_Fld_1")
        Me.txtUser_Def_Fld_1.Mask = Nothing
        Me.txtUser_Def_Fld_1.Name = "txtUser_Def_Fld_1"
        Me.txtUser_Def_Fld_1.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtUser_Def_Fld_1.OldValue = Nothing
        Me.txtUser_Def_Fld_1.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtUser_Def_Fld_1.StringValue = Nothing
        Me.txtUser_Def_Fld_1.TextDataID = Nothing
        '
        'txtShip_To_Country
        '
        Me.txtShip_To_Country.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_To_Country.DataLength = CType(3, Long)
        Me.txtShip_To_Country.DataType = Orders.CDataType.dtString
        Me.txtShip_To_Country.DateValue = New Date(CType(0, Long))
        Me.txtShip_To_Country.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtShip_To_Country, "txtShip_To_Country")
        Me.txtShip_To_Country.Mask = Nothing
        Me.txtShip_To_Country.Name = "txtShip_To_Country"
        Me.txtShip_To_Country.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_To_Country.OldValue = Nothing
        Me.txtShip_To_Country.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_To_Country.StringValue = Nothing
        Me.txtShip_To_Country.TextDataID = Nothing
        '
        'txtBill_To_Country
        '
        Me.txtBill_To_Country.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBill_To_Country.DataLength = CType(3, Long)
        Me.txtBill_To_Country.DataType = Orders.CDataType.dtString
        Me.txtBill_To_Country.DateValue = New Date(CType(0, Long))
        Me.txtBill_To_Country.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtBill_To_Country, "txtBill_To_Country")
        Me.txtBill_To_Country.Mask = Nothing
        Me.txtBill_To_Country.Name = "txtBill_To_Country"
        Me.txtBill_To_Country.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBill_To_Country.OldValue = Nothing
        Me.txtBill_To_Country.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBill_To_Country.StringValue = Nothing
        Me.txtBill_To_Country.TextDataID = Nothing
        '
        'txtShip_To_Addr_3
        '
        Me.txtShip_To_Addr_3.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_To_Addr_3.DataLength = CType(40, Long)
        Me.txtShip_To_Addr_3.DataType = Orders.CDataType.dtString
        Me.txtShip_To_Addr_3.DateValue = New Date(CType(0, Long))
        Me.txtShip_To_Addr_3.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_To_Addr_3, "txtShip_To_Addr_3")
        Me.txtShip_To_Addr_3.Mask = Nothing
        Me.txtShip_To_Addr_3.Name = "txtShip_To_Addr_3"
        Me.txtShip_To_Addr_3.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_To_Addr_3.OldValue = Nothing
        Me.txtShip_To_Addr_3.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_To_Addr_3.StringValue = Nothing
        Me.txtShip_To_Addr_3.TextDataID = Nothing
        '
        'txtShip_To_Addr_2
        '
        Me.txtShip_To_Addr_2.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_To_Addr_2.DataLength = CType(40, Long)
        Me.txtShip_To_Addr_2.DataType = Orders.CDataType.dtString
        Me.txtShip_To_Addr_2.DateValue = New Date(CType(0, Long))
        Me.txtShip_To_Addr_2.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_To_Addr_2, "txtShip_To_Addr_2")
        Me.txtShip_To_Addr_2.Mask = Nothing
        Me.txtShip_To_Addr_2.Name = "txtShip_To_Addr_2"
        Me.txtShip_To_Addr_2.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_To_Addr_2.OldValue = Nothing
        Me.txtShip_To_Addr_2.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_To_Addr_2.StringValue = Nothing
        Me.txtShip_To_Addr_2.TextDataID = Nothing
        '
        'txtShip_To_Addr_1
        '
        Me.txtShip_To_Addr_1.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_To_Addr_1.DataLength = CType(40, Long)
        Me.txtShip_To_Addr_1.DataType = Orders.CDataType.dtString
        Me.txtShip_To_Addr_1.DateValue = New Date(CType(0, Long))
        Me.txtShip_To_Addr_1.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_To_Addr_1, "txtShip_To_Addr_1")
        Me.txtShip_To_Addr_1.Mask = Nothing
        Me.txtShip_To_Addr_1.Name = "txtShip_To_Addr_1"
        Me.txtShip_To_Addr_1.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_To_Addr_1.OldValue = Nothing
        Me.txtShip_To_Addr_1.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_To_Addr_1.StringValue = Nothing
        Me.txtShip_To_Addr_1.TextDataID = Nothing
        '
        'txtShip_To_Name
        '
        Me.txtShip_To_Name.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_To_Name.DataLength = CType(40, Long)
        Me.txtShip_To_Name.DataType = Orders.CDataType.dtString
        Me.txtShip_To_Name.DateValue = New Date(CType(0, Long))
        Me.txtShip_To_Name.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_To_Name, "txtShip_To_Name")
        Me.txtShip_To_Name.Mask = Nothing
        Me.txtShip_To_Name.Name = "txtShip_To_Name"
        Me.txtShip_To_Name.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_To_Name.OldValue = Nothing
        Me.txtShip_To_Name.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_To_Name.StringValue = Nothing
        Me.txtShip_To_Name.TextDataID = Nothing
        '
        'txtBill_To_Addr_3
        '
        Me.txtBill_To_Addr_3.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBill_To_Addr_3.DataLength = CType(40, Long)
        Me.txtBill_To_Addr_3.DataType = Orders.CDataType.dtString
        Me.txtBill_To_Addr_3.DateValue = New Date(CType(0, Long))
        Me.txtBill_To_Addr_3.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtBill_To_Addr_3, "txtBill_To_Addr_3")
        Me.txtBill_To_Addr_3.Mask = Nothing
        Me.txtBill_To_Addr_3.Name = "txtBill_To_Addr_3"
        Me.txtBill_To_Addr_3.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBill_To_Addr_3.OldValue = Nothing
        Me.txtBill_To_Addr_3.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBill_To_Addr_3.StringValue = Nothing
        Me.txtBill_To_Addr_3.TextDataID = Nothing
        '
        'txtBill_To_Addr_2
        '
        Me.txtBill_To_Addr_2.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBill_To_Addr_2.DataLength = CType(40, Long)
        Me.txtBill_To_Addr_2.DataType = Orders.CDataType.dtString
        Me.txtBill_To_Addr_2.DateValue = New Date(CType(0, Long))
        Me.txtBill_To_Addr_2.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtBill_To_Addr_2, "txtBill_To_Addr_2")
        Me.txtBill_To_Addr_2.Mask = Nothing
        Me.txtBill_To_Addr_2.Name = "txtBill_To_Addr_2"
        Me.txtBill_To_Addr_2.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBill_To_Addr_2.OldValue = Nothing
        Me.txtBill_To_Addr_2.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBill_To_Addr_2.StringValue = Nothing
        Me.txtBill_To_Addr_2.TextDataID = Nothing
        '
        'txtBill_To_Addr_1
        '
        Me.txtBill_To_Addr_1.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBill_To_Addr_1.DataLength = CType(40, Long)
        Me.txtBill_To_Addr_1.DataType = Orders.CDataType.dtString
        Me.txtBill_To_Addr_1.DateValue = New Date(CType(0, Long))
        Me.txtBill_To_Addr_1.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtBill_To_Addr_1, "txtBill_To_Addr_1")
        Me.txtBill_To_Addr_1.Mask = Nothing
        Me.txtBill_To_Addr_1.Name = "txtBill_To_Addr_1"
        Me.txtBill_To_Addr_1.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBill_To_Addr_1.OldValue = Nothing
        Me.txtBill_To_Addr_1.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBill_To_Addr_1.StringValue = Nothing
        Me.txtBill_To_Addr_1.TextDataID = Nothing
        '
        'txtBill_To_Name
        '
        Me.txtBill_To_Name.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBill_To_Name.DataLength = CType(40, Long)
        Me.txtBill_To_Name.DataType = Orders.CDataType.dtString
        Me.txtBill_To_Name.DateValue = New Date(CType(0, Long))
        Me.txtBill_To_Name.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtBill_To_Name, "txtBill_To_Name")
        Me.txtBill_To_Name.Mask = Nothing
        Me.txtBill_To_Name.Name = "txtBill_To_Name"
        Me.txtBill_To_Name.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBill_To_Name.OldValue = Nothing
        Me.txtBill_To_Name.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBill_To_Name.StringValue = Nothing
        Me.txtBill_To_Name.TextDataID = Nothing
        '
        'txtShip_Instruction_2
        '
        Me.txtShip_Instruction_2.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_Instruction_2.DataLength = CType(40, Long)
        Me.txtShip_Instruction_2.DataType = Orders.CDataType.dtString
        Me.txtShip_Instruction_2.DateValue = New Date(CType(0, Long))
        Me.txtShip_Instruction_2.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_Instruction_2, "txtShip_Instruction_2")
        Me.txtShip_Instruction_2.Mask = Nothing
        Me.txtShip_Instruction_2.Name = "txtShip_Instruction_2"
        Me.txtShip_Instruction_2.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_Instruction_2.OldValue = Nothing
        Me.txtShip_Instruction_2.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_Instruction_2.StringValue = Nothing
        Me.txtShip_Instruction_2.TextDataID = Nothing
        '
        'txtProfit_Center
        '
        Me.txtProfit_Center.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtProfit_Center.DataLength = CType(8, Long)
        Me.txtProfit_Center.DataType = Orders.CDataType.dtString
        Me.txtProfit_Center.DateValue = New Date(CType(0, Long))
        Me.txtProfit_Center.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtProfit_Center, "txtProfit_Center")
        Me.txtProfit_Center.Mask = Nothing
        Me.txtProfit_Center.Name = "txtProfit_Center"
        Me.txtProfit_Center.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtProfit_Center.OldValue = Nothing
        Me.txtProfit_Center.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtProfit_Center.StringValue = Nothing
        Me.txtProfit_Center.TextDataID = Nothing
        '
        'txtAr_Terms_Cd
        '
        Me.txtAr_Terms_Cd.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtAr_Terms_Cd.DataLength = CType(3, Long)
        Me.txtAr_Terms_Cd.DataType = Orders.CDataType.dtString
        Me.txtAr_Terms_Cd.DateValue = New Date(CType(0, Long))
        Me.txtAr_Terms_Cd.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtAr_Terms_Cd, "txtAr_Terms_Cd")
        Me.txtAr_Terms_Cd.Mask = Nothing
        Me.txtAr_Terms_Cd.Name = "txtAr_Terms_Cd"
        Me.txtAr_Terms_Cd.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtAr_Terms_Cd.OldValue = Nothing
        Me.txtAr_Terms_Cd.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtAr_Terms_Cd.StringValue = Nothing
        Me.txtAr_Terms_Cd.TextDataID = Nothing
        '
        'txtShip_Via_Cd
        '
        Me.txtShip_Via_Cd.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_Via_Cd.DataLength = CType(3, Long)
        Me.txtShip_Via_Cd.DataType = Orders.CDataType.dtString
        Me.txtShip_Via_Cd.DateValue = New Date(CType(0, Long))
        Me.txtShip_Via_Cd.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtShip_Via_Cd, "txtShip_Via_Cd")
        Me.txtShip_Via_Cd.Mask = Nothing
        Me.txtShip_Via_Cd.Name = "txtShip_Via_Cd"
        Me.txtShip_Via_Cd.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_Via_Cd.OldValue = Nothing
        Me.txtShip_Via_Cd.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_Via_Cd.StringValue = Nothing
        Me.txtShip_Via_Cd.TextDataID = Nothing
        '
        'txtMfg_Loc
        '
        Me.txtMfg_Loc.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtMfg_Loc.DataLength = CType(3, Long)
        Me.txtMfg_Loc.DataType = Orders.CDataType.dtString
        Me.txtMfg_Loc.DateValue = New Date(CType(0, Long))
        Me.txtMfg_Loc.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtMfg_Loc, "txtMfg_Loc")
        Me.txtMfg_Loc.Mask = Nothing
        Me.txtMfg_Loc.Name = "txtMfg_Loc"
        Me.txtMfg_Loc.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtMfg_Loc.OldValue = Nothing
        Me.txtMfg_Loc.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtMfg_Loc.StringValue = Nothing
        Me.txtMfg_Loc.TextDataID = Nothing
        '
        'txtShip_Instruction_1
        '
        Me.txtShip_Instruction_1.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtShip_Instruction_1.DataLength = CType(40, Long)
        Me.txtShip_Instruction_1.DataType = Orders.CDataType.dtString
        Me.txtShip_Instruction_1.DateValue = New Date(CType(0, Long))
        Me.txtShip_Instruction_1.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtShip_Instruction_1, "txtShip_Instruction_1")
        Me.txtShip_Instruction_1.Mask = Nothing
        Me.txtShip_Instruction_1.Name = "txtShip_Instruction_1"
        Me.txtShip_Instruction_1.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtShip_Instruction_1.OldValue = Nothing
        Me.txtShip_Instruction_1.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtShip_Instruction_1.StringValue = Nothing
        Me.txtShip_Instruction_1.TextDataID = Nothing
        '
        'txtJob_No
        '
        Me.txtJob_No.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtJob_No.DataLength = CType(20, Long)
        Me.txtJob_No.DataType = Orders.CDataType.dtString
        Me.txtJob_No.DateValue = New Date(CType(0, Long))
        Me.txtJob_No.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtJob_No, "txtJob_No")
        Me.txtJob_No.Mask = Nothing
        Me.txtJob_No.Name = "txtJob_No"
        Me.txtJob_No.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtJob_No.OldValue = Nothing
        Me.txtJob_No.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtJob_No.StringValue = Nothing
        Me.txtJob_No.TextDataID = Nothing
        '
        'txtSlspsn_No
        '
        Me.txtSlspsn_No.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtSlspsn_No.DataLength = CType(9, Long)
        Me.txtSlspsn_No.DataType = Orders.CDataType.dtString
        Me.txtSlspsn_No.DateValue = New Date(CType(0, Long))
        Me.txtSlspsn_No.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtSlspsn_No, "txtSlspsn_No")
        Me.txtSlspsn_No.Mask = Nothing
        Me.txtSlspsn_No.Name = "txtSlspsn_No"
        Me.txtSlspsn_No.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtSlspsn_No.OldValue = Nothing
        Me.txtSlspsn_No.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtSlspsn_No.StringValue = Nothing
        Me.txtSlspsn_No.TextDataID = Nothing
        '
        'txtOe_Po_No
        '
        Me.txtOe_Po_No.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtOe_Po_No.DataLength = CType(25, Long)
        Me.txtOe_Po_No.DataType = Orders.CDataType.dtString
        Me.txtOe_Po_No.DateValue = New Date(CType(0, Long))
        Me.txtOe_Po_No.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtOe_Po_No, "txtOe_Po_No")
        Me.txtOe_Po_No.Mask = Nothing
        Me.txtOe_Po_No.Name = "txtOe_Po_No"
        Me.txtOe_Po_No.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtOe_Po_No.OldValue = Nothing
        Me.txtOe_Po_No.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtOe_Po_No.StringValue = Nothing
        Me.txtOe_Po_No.TextDataID = Nothing
        '
        'txtCus_Alt_Adr_Cd
        '
        Me.txtCus_Alt_Adr_Cd.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtCus_Alt_Adr_Cd.DataLength = CType(15, Long)
        Me.txtCus_Alt_Adr_Cd.DataType = Orders.CDataType.dtString
        Me.txtCus_Alt_Adr_Cd.DateValue = New Date(CType(0, Long))
        Me.txtCus_Alt_Adr_Cd.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtCus_Alt_Adr_Cd, "txtCus_Alt_Adr_Cd")
        Me.txtCus_Alt_Adr_Cd.Mask = Nothing
        Me.txtCus_Alt_Adr_Cd.Name = "txtCus_Alt_Adr_Cd"
        Me.txtCus_Alt_Adr_Cd.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtCus_Alt_Adr_Cd.OldValue = Nothing
        Me.txtCus_Alt_Adr_Cd.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtCus_Alt_Adr_Cd.StringValue = Nothing
        Me.txtCus_Alt_Adr_Cd.TextDataID = Nothing
        '
        'txtCus_No
        '
        Me.txtCus_No.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtCus_No.DataLength = CType(20, Long)
        Me.txtCus_No.DataType = Orders.CDataType.dtString
        Me.txtCus_No.DateValue = New Date(CType(0, Long))
        Me.txtCus_No.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtCus_No, "txtCus_No")
        Me.txtCus_No.Mask = Nothing
        Me.txtCus_No.Name = "txtCus_No"
        Me.txtCus_No.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtCus_No.OldValue = Nothing
        Me.txtCus_No.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtCus_No.StringValue = Nothing
        Me.txtCus_No.TextDataID = Nothing
        '
        'txtOrd_No
        '
        Me.txtOrd_No.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtOrd_No.DataLength = CType(10, Long)
        Me.txtOrd_No.DataType = Orders.CDataType.dtString
        Me.txtOrd_No.DateValue = New Date(CType(0, Long))
        Me.txtOrd_No.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtOrd_No, "txtOrd_No")
        Me.txtOrd_No.Mask = Nothing
        Me.txtOrd_No.Name = "txtOrd_No"
        Me.txtOrd_No.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtOrd_No.OldValue = Nothing
        Me.txtOrd_No.SpacePadding = Orders.CSpacePadding.PaddingBefore
        Me.txtOrd_No.StringValue = Nothing
        Me.txtOrd_No.TextDataID = Nothing
        '
        'txtDiscount_Pct
        '
        Me.txtDiscount_Pct.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtDiscount_Pct.DataLength = CType(6, Long)
        Me.txtDiscount_Pct.DataType = Orders.CDataType.dtNumWithDecimals
        Me.txtDiscount_Pct.DateValue = New Date(CType(0, Long))
        Me.txtDiscount_Pct.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtDiscount_Pct, "txtDiscount_Pct")
        Me.txtDiscount_Pct.Mask = Nothing
        Me.txtDiscount_Pct.Name = "txtDiscount_Pct"
        Me.txtDiscount_Pct.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtDiscount_Pct.OldValue = Nothing
        Me.txtDiscount_Pct.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtDiscount_Pct.StringValue = Nothing
        Me.txtDiscount_Pct.TextDataID = Nothing
        '
        'txtOrd_Dt
        '
        Me.txtOrd_Dt.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtOrd_Dt.DataLength = CType(0, Long)
        Me.txtOrd_Dt.DataType = Orders.CDataType.dtDate
        Me.txtOrd_Dt.DateValue = New Date(CType(0, Long))
        Me.txtOrd_Dt.DecimalDigits = CType(0, Long)
        resources.ApplyResources(Me.txtOrd_Dt, "txtOrd_Dt")
        Me.txtOrd_Dt.Mask = Nothing
        Me.txtOrd_Dt.Name = "txtOrd_Dt"
        Me.txtOrd_Dt.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtOrd_Dt.OldValue = Nothing
        Me.txtOrd_Dt.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtOrd_Dt.StringValue = Nothing
        Me.txtOrd_Dt.TextDataID = Nothing
        '
        'txtEnd_User
        '
        Me.txtEnd_User.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtEnd_User.DataLength = CType(30, Long)
        Me.txtEnd_User.DataType = Orders.CDataType.dtString
        Me.txtEnd_User.DateValue = New Date(CType(0, Long))
        Me.txtEnd_User.DecimalDigits = CType(2, Long)
        resources.ApplyResources(Me.txtEnd_User, "txtEnd_User")
        Me.txtEnd_User.Mask = Nothing
        Me.txtEnd_User.Name = "txtEnd_User"
        Me.txtEnd_User.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtEnd_User.OldValue = Nothing
        Me.txtEnd_User.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtEnd_User.StringValue = Nothing
        Me.txtEnd_User.TextDataID = Nothing
        '
        'lblVIP
        '
        resources.ApplyResources(Me.lblVIP, "lblVIP")
        Me.lblVIP.BackColor = System.Drawing.Color.Gray
        Me.lblVIP.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblVIP.ForeColor = System.Drawing.Color.Gold
        Me.lblVIP.Name = "lblVIP"
        '
        'ucOrder
        '
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Controls.Add(Me.lblVIP)
        Me.Controls.Add(Me.xTxtInvBatch)
        Me.Controls.Add(Me.lblInvBatch)
        Me.Controls.Add(Me.lblSSp)
        Me.Controls.Add(Me.cboreSSP)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.cboWhyNotVegas)
        Me.Controls.Add(Me.lblExtra_9)
        Me.Controls.Add(Me.cbo_Extra_9)
        Me.Controls.Add(Me.txtCustGroup)
        Me.Controls.Add(Me.cmdApplyToNote)
        Me.Controls.Add(Me.cmdApplyToSearch)
        Me.Controls.Add(Me.txtOrd_Dt_Shipped)
        Me.Controls.Add(Me.lblApplyTo)
        Me.Controls.Add(Me.txtInHandsDate)
        Me.Controls.Add(Me.mcCalendar)
        Me.Controls.Add(Me.txtApplyTo)
        Me.Controls.Add(Me.cmdXml)
        Me.Controls.Add(Me.cmdInHandDate)
        Me.Controls.Add(Me.lblInHandsDate)
        Me.Controls.Add(Me.cmdShipDateTP)
        Me.Controls.Add(Me.lblShipDate)
        Me.Controls.Add(Me.txtDept)
        Me.Controls.Add(Me.lblCostCenter)
        Me.Controls.Add(Me.cmdDeptSearch)
        Me.Controls.Add(Me.cmdDept_Notes)
        Me.Controls.Add(Me.cmdCustomerDash)
        Me.Controls.Add(Me.txtCus_Prog_ID)
        Me.Controls.Add(Me.cmdProg_Spector_CdSearch)
        Me.Controls.Add(Me.txtProg_Spector_Cd)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.cmdEDI)
        Me.Controls.Add(Me.cmdSendEmail)
        Me.Controls.Add(Me.cmdOrderRequest)
        Me.Controls.Add(Me.lblProfitCenter)
        Me.Controls.Add(Me.cboProgram)
        Me.Controls.Add(Me.lblProgram)
        Me.Controls.Add(Me.chkEmail2)
        Me.Controls.Add(Me.chkInvalidShipDateEmail)
        Me.Controls.Add(Me.chkExactRepeat)
        Me.Controls.Add(Me.chkRedo)
        Me.Controls.Add(Me.chkRecover)
        Me.Controls.Add(Me.lblOptions)
        Me.Controls.Add(Me.chkOrderAckSaveOnly)
        Me.Controls.Add(Me.txtShip_To_Zip)
        Me.Controls.Add(Me.txtShip_To_State)
        Me.Controls.Add(Me.txtShip_To_City)
        Me.Controls.Add(Me.txtShip_To_Addr_4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtBill_To_Zip)
        Me.Controls.Add(Me.txtBill_To_City)
        Me.Controls.Add(Me.txtBill_To_State)
        Me.Controls.Add(Me.txtBill_To_Addr_4)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdShip_To_NameSearch)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.cmdEditCapturedBill)
        Me.Controls.Add(Me.cmdATP)
        Me.Controls.Add(Me.cmdInvoice)
        Me.Controls.Add(Me.cmdHistoryCopy)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cmdUser_Def_Fld_3Search)
        Me.Controls.Add(Me.cmdShip_LinkEdit)
        Me.Controls.Add(Me.txtShip_Link)
        Me.Controls.Add(Me.cmdShip_LinkSearch)
        Me.Controls.Add(Me.lblShip_Link)
        Me.Controls.Add(Me.cmdSalespersonSearch)
        Me.Controls.Add(Me.txtUser_Def_Fld_5)
        Me.Controls.Add(Me.cmdUser_Def_Fld_5Search)
        Me.Controls.Add(Me.lblUser_Def_Fld_5)
        Me.Controls.Add(Me.cmdCopyDate)
        Me.Controls.Add(Me.cmdOrderLink)
        Me.Controls.Add(Me.cmdOEFlash)
        Me.Controls.Add(Me.cmdUser_Def_Fld_4Search)
        Me.Controls.Add(Me.txtContactEmail)
        Me.Controls.Add(Me.lblContactEmail)
        Me.Controls.Add(Me.txtUser_Def_Fld_4)
        Me.Controls.Add(Me.lblUser_Def_Fld_4)
        Me.Controls.Add(Me.txtUser_Def_Fld_3)
        Me.Controls.Add(Me.txtUser_Def_Fld_1)
        Me.Controls.Add(Me.lblUser_Def_Fld_3)
        Me.Controls.Add(Me.lblUser_Def_Fld_1)
        Me.Controls.Add(Me.cboUser_Def_Fld_2)
        Me.Controls.Add(Me.lblUser_Def_Fld_2)
        Me.Controls.Add(Me.cmdCommissions)
        Me.Controls.Add(Me.cmdTax)
        Me.Controls.Add(Me.cmdComments)
        Me.Controls.Add(Me.cmdLocQty)
        Me.Controls.Add(Me.cmdItemNotes)
        Me.Controls.Add(Me.cmdItemInfo)
        Me.Controls.Add(Me.cmdLineUserFields)
        Me.Controls.Add(Me.cmdPickTicket)
        Me.Controls.Add(Me.cmdAcknowledgement)
        Me.Controls.Add(Me.cmdMasterCopy)
        Me.Controls.Add(Me.cmdPrintQuote)
        Me.Controls.Add(Me.cmdConvertQuote)
        Me.Controls.Add(Me.cmdShip_To_Country_Notes)
        Me.Controls.Add(Me.cmdBill_To_Country_Notes)
        Me.Controls.Add(Me.cmdProfit_Center_Notes)
        Me.Controls.Add(Me.cmdAr_Terms_Cd_Notes)
        Me.Controls.Add(Me.cmdShip_Via_Cd_Notes)
        Me.Controls.Add(Me.cmdJob_No_Notes)
        Me.Controls.Add(Me.cmdSlspsn_No_Notes)
        Me.Controls.Add(Me.cmdCus_Alt_Adr_Cd_Notes)
        Me.Controls.Add(Me.cmdCus_No_Notes)
        Me.Controls.Add(Me.cmdOrd_No_Notes)
        Me.Controls.Add(Me.cmdMfg_Loc_Notes)
        Me.Controls.Add(Me.lblComments)
        Me.Controls.Add(Me.txtShip_To_Country)
        Me.Controls.Add(Me.txtBill_To_Country)
        Me.Controls.Add(Me.txtShip_To_Addr_3)
        Me.Controls.Add(Me.txtShip_To_Addr_2)
        Me.Controls.Add(Me.txtShip_To_Addr_1)
        Me.Controls.Add(Me.txtShip_To_Name)
        Me.Controls.Add(Me.txtBill_To_Addr_3)
        Me.Controls.Add(Me.txtBill_To_Addr_2)
        Me.Controls.Add(Me.txtBill_To_Addr_1)
        Me.Controls.Add(Me.txtBill_To_Name)
        Me.Controls.Add(Me.txtShip_Instruction_2)
        Me.Controls.Add(Me.txtProfit_Center)
        Me.Controls.Add(Me.txtAr_Terms_Cd)
        Me.Controls.Add(Me.txtShip_Via_Cd)
        Me.Controls.Add(Me.txtMfg_Loc)
        Me.Controls.Add(Me.txtShip_Instruction_1)
        Me.Controls.Add(Me.txtJob_No)
        Me.Controls.Add(Me.txtSlspsn_No)
        Me.Controls.Add(Me.txtOe_Po_No)
        Me.Controls.Add(Me.txtCus_Alt_Adr_Cd)
        Me.Controls.Add(Me.txtCus_No)
        Me.Controls.Add(Me.txtOrd_No)
        Me.Controls.Add(Me.txtDiscount_Pct)
        Me.Controls.Add(Me.txtOrd_Dt)
        Me.Controls.Add(Me.cboShipVia)
        Me.Controls.Add(Me.cboShipTo)
        Me.Controls.Add(Me.cboCustomerNo)
        Me.Controls.Add(Me.cboShipToName)
        Me.Controls.Add(Me.cboBillToName)
        Me.Controls.Add(Me.cmdCopyBillTo)
        Me.Controls.Add(Me.cmdShipToCountrySearch)
        Me.Controls.Add(Me.cmdChangeBillTo)
        Me.Controls.Add(Me.cmdBillToCountrySearch)
        Me.Controls.Add(Me.cmdCommentsSearch)
        Me.Controls.Add(Me.cmdProfitCenterSearch)
        Me.Controls.Add(Me.cmdTermsSearch)
        Me.Controls.Add(Me.cmdShipViaSearch)
        Me.Controls.Add(Me.cmdLocationSearch)
        Me.Controls.Add(Me.cmdInstructionsSearch)
        Me.Controls.Add(Me.cmdProjectSearch)
        Me.Controls.Add(Me.CmdShipToSearch)
        Me.Controls.Add(Me.cmdCustomerSearch)
        Me.Controls.Add(Me.cmdDateTP)
        Me.Controls.Add(Me.cmdOrderSearch)
        Me.Controls.Add(Me.cboOrd_Type)
        Me.Controls.Add(Me.cmdCustomerConfig)
        Me.Controls.Add(Me.cmdCopyOrderFromHistory)
        Me.Controls.Add(Me.lblShipToCountryDesc)
        Me.Controls.Add(Me.lblShipToCountry)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.lblShipToName)
        Me.Controls.Add(Me.lblBillToCountryDesc)
        Me.Controls.Add(Me.lblBillToCountry)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.lblBillToName)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.lblTermsDesc)
        Me.Controls.Add(Me.lblShipViaDesc)
        Me.Controls.Add(Me.lblLocationDesc)
        Me.Controls.Add(Me.lblLocation)
        Me.Controls.Add(Me.lblShipVia)
        Me.Controls.Add(Me.lblTerms)
        Me.Controls.Add(Me.lblDiscPct)
        Me.Controls.Add(Me.lblInstructions)
        Me.Controls.Add(Me.lblProject)
        Me.Controls.Add(Me.lblSalesPerson)
        Me.Controls.Add(Me.lblPONumber)
        Me.Controls.Add(Me.lblShipTo)
        Me.Controls.Add(Me.lblCustomer)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.lblOrder)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.Controls.Add(Me.lblEnd_User)
        Me.Controls.Add(Me.cmdEnd_UserSearch)
        Me.Controls.Add(Me.cmdEnd_User_Notes)
        Me.Controls.Add(Me.txtEnd_User)
        resources.ApplyResources(Me, "$this")
        Me.Name = "ucOrder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblLocationDesc As System.Windows.Forms.Label
    Friend WithEvents mcCalendar As System.Windows.Forms.MonthCalendar
    Friend WithEvents lblCostCenter As System.Windows.Forms.Label
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents ttToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents txtOrd_Dt As Orders.xTextBox
    Friend WithEvents txtDiscount_Pct As Orders.xTextBox
    Friend WithEvents txtOrd_No As Orders.xTextBox
    Friend WithEvents txtCus_No As Orders.xTextBox
    Friend WithEvents txtCus_Alt_Adr_Cd As Orders.xTextBox
    Friend WithEvents txtOe_Po_No As Orders.xTextBox
    Friend WithEvents txtSlspsn_No As Orders.xTextBox
    Friend WithEvents txtJob_No As Orders.xTextBox
    Friend WithEvents txtShip_Instruction_1 As Orders.xTextBox
    Friend WithEvents txtMfg_Loc As Orders.xTextBox
    Friend WithEvents txtShip_Via_Cd As Orders.xTextBox
    Friend WithEvents txtAr_Terms_Cd As Orders.xTextBox
    Friend WithEvents txtProfit_Center As Orders.xTextBox
    Friend WithEvents txtDept As Orders.xTextBox
    Friend WithEvents txtShip_Instruction_2 As Orders.xTextBox
    Friend WithEvents txtBill_To_Name As Orders.xTextBox
    Friend WithEvents txtBill_To_Addr_1 As Orders.xTextBox
    Friend WithEvents txtBill_To_Addr_2 As Orders.xTextBox
    Friend WithEvents txtBill_To_Addr_3 As Orders.xTextBox
    Friend WithEvents txtBill_To_Addr_4 As Orders.xTextBox
    Friend WithEvents txtShip_To_Addr_4 As Orders.xTextBox
    Friend WithEvents txtShip_To_Addr_3 As Orders.xTextBox
    Friend WithEvents txtShip_To_Addr_2 As Orders.xTextBox
    Friend WithEvents txtShip_To_Addr_1 As Orders.xTextBox
    Friend WithEvents txtShip_To_Name As Orders.xTextBox
    Friend WithEvents txtBill_To_Country As Orders.xTextBox
    Friend WithEvents txtShip_To_Country As Orders.xTextBox
    Friend WithEvents cmdMfg_Loc_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdOrd_No_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdCus_No_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdCus_Alt_Adr_Cd_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdSlspsn_No_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdJob_No_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdShip_Via_Cd_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdAr_Terms_Cd_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdProfit_Center_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdBill_To_Country_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdShip_To_Country_Notes As System.Windows.Forms.Button
    Friend WithEvents cmdConvertQuote As System.Windows.Forms.Button
    Friend WithEvents cmdPrintQuote As System.Windows.Forms.Button
    Friend WithEvents cmdHistoryCopy As System.Windows.Forms.Button
    Friend WithEvents cmdMasterCopy As System.Windows.Forms.Button
    Friend WithEvents cmdLineUserFields As System.Windows.Forms.Button
    Friend WithEvents cmdInvoice As System.Windows.Forms.Button
    Friend WithEvents cmdPickTicket As System.Windows.Forms.Button
    Friend WithEvents cmdAcknowledgement As System.Windows.Forms.Button
    Friend WithEvents cmdComments As System.Windows.Forms.Button
    Friend WithEvents cmdLocQty As System.Windows.Forms.Button
    Friend WithEvents cmdItemNotes As System.Windows.Forms.Button
    Friend WithEvents cmdItemInfo As System.Windows.Forms.Button
    Friend WithEvents cmdCommissions As System.Windows.Forms.Button
    Friend WithEvents cmdTax As System.Windows.Forms.Button
    Friend WithEvents cmdEditCapturedBill As System.Windows.Forms.Button
    Friend WithEvents cmdATP As System.Windows.Forms.Button
    Friend WithEvents lblUser_Def_Fld_2 As System.Windows.Forms.Label
    Friend WithEvents cboUser_Def_Fld_2 As System.Windows.Forms.ComboBox
    Friend WithEvents txtUser_Def_Fld_3 As Orders.xTextBox
    Friend WithEvents txtUser_Def_Fld_1 As Orders.xTextBox
    Friend WithEvents lblUser_Def_Fld_3 As System.Windows.Forms.Label
    Friend WithEvents lblUser_Def_Fld_1 As System.Windows.Forms.Label
    Friend WithEvents txtUser_Def_Fld_4 As Orders.xTextBox
    Friend WithEvents lblUser_Def_Fld_4 As System.Windows.Forms.Label
    Friend WithEvents txtContactEmail As Orders.xTextBox
    Friend WithEvents lblContactEmail As System.Windows.Forms.Label
    Friend WithEvents cmdUser_Def_Fld_4Search As System.Windows.Forms.Button
    Friend WithEvents cmdOEFlash As System.Windows.Forms.Button
    Friend WithEvents cmdOrderLink As System.Windows.Forms.Button
    Friend WithEvents cmdCopyDate As System.Windows.Forms.Button
    Friend WithEvents txtUser_Def_Fld_5 As Orders.xTextBox
    Friend WithEvents cmdUser_Def_Fld_5Search As System.Windows.Forms.Button
    Friend WithEvents lblUser_Def_Fld_5 As System.Windows.Forms.Label
    Friend WithEvents cmdShip_LinkEdit As System.Windows.Forms.Button
    Friend WithEvents txtShip_Link As Orders.xTextBox
    Friend WithEvents cmdShip_LinkSearch As System.Windows.Forms.Button
    Friend WithEvents lblShip_Link As System.Windows.Forms.Label
    Friend WithEvents cmdUser_Def_Fld_3Search As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents cmdShip_To_NameSearch As System.Windows.Forms.Button
    Friend WithEvents txtBill_To_City As Orders.xTextBox
    Friend WithEvents txtBill_To_State As Orders.xTextBox
    Friend WithEvents txtBill_To_Zip As Orders.xTextBox
    Friend WithEvents txtShip_To_Zip As Orders.xTextBox
    Friend WithEvents txtShip_To_State As Orders.xTextBox
    Friend WithEvents txtShip_To_City As Orders.xTextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents chkOrderAckSaveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents lblOptions As System.Windows.Forms.Label
    Friend WithEvents chkRecover As System.Windows.Forms.CheckBox
    Friend WithEvents chkRedo As System.Windows.Forms.CheckBox
    Friend WithEvents chkExactRepeat As System.Windows.Forms.CheckBox
    Friend WithEvents chkInvalidShipDateEmail As System.Windows.Forms.CheckBox
    Friend WithEvents chkEmail2 As System.Windows.Forms.CheckBox
    Friend WithEvents cboProgram As System.Windows.Forms.ComboBox
    Friend WithEvents lblProgram As System.Windows.Forms.Label
    Friend WithEvents cmdOrderRequest As System.Windows.Forms.Button
    Friend WithEvents cmdSendEmail As System.Windows.Forms.Button
    Friend WithEvents cmdEDI As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents cmdProg_Spector_CdSearch As System.Windows.Forms.Button
    Friend WithEvents txtProg_Spector_Cd As Orders.xTextBox
    Friend WithEvents txtCus_Prog_ID As Orders.xTextBox
    Friend WithEvents cmdCustomerDash As System.Windows.Forms.Button
    Friend WithEvents txtEnd_User As Orders.xTextBox
    Friend WithEvents lblEnd_User As System.Windows.Forms.Label
    Friend WithEvents cmdEnd_UserSearch As System.Windows.Forms.Button
    Friend WithEvents cmdEnd_User_Notes As System.Windows.Forms.Button
    Friend WithEvents txtInHandsDate As Orders.xTextBox
    Friend WithEvents cmdInHandDate As System.Windows.Forms.Button
    Friend WithEvents lblInHandsDate As System.Windows.Forms.Label
    Friend WithEvents txtApplyTo As Orders.xTextBox
    Friend WithEvents cmdApplyToSearch As System.Windows.Forms.Button
    Friend WithEvents lblApplyTo As System.Windows.Forms.Label
    Friend WithEvents txtOrd_Dt_Shipped As Orders.xTextBox
    Friend WithEvents cmdShipDateTP As System.Windows.Forms.Button
    Friend WithEvents lblShipDate As System.Windows.Forms.Label
    Friend WithEvents cmdXml As System.Windows.Forms.Button
    Friend WithEvents cmdApplyToNote As System.Windows.Forms.Button
    Private WithEvents Line3 As PowerPacks.LineShape
    Private WithEvents ShapeContainer1 As PowerPacks.ShapeContainer
    Private WithEvents LineShape1 As PowerPacks.LineShape
    Friend WithEvents txtCustGroup As xTextBox
    Friend WithEvents cbo_Extra_9 As ComboBox
    Friend WithEvents lblExtra_9 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents cboWhyNotVegas As ComboBox
    Friend WithEvents cboreSSP As ComboBox
    Friend WithEvents lblSSp As Label
    Friend WithEvents xTxtInvBatch As xTextBox
    Friend WithEvents lblInvBatch As Label
    Friend WithEvents lblVIP As Label
    'Friend WithEvents txtOrd_Dt As Orders.xTextBox
    'Friend WithEvents txtDiscount_Pct As Orders.xTextBox

#End Region
End Class