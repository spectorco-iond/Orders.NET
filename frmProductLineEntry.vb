Imports System.Data.SqlClient
Imports System.IO

Public Class frmProductLineEntry

    Private blnNoMsgBox As Boolean = False

    ' Private m_oOrder As cOrder
    Private m_oOrdline As cOrdLine
    Private m_AllStarProgram As New cAllStarProgram

    Public Currency As String
    Public Customer_No As String
    Public Item_No As String
    Public Loc As String

    Private lInvalidate As Long = 0
    Private iRowPos As Integer = 0
    Private iColPos As Integer = 0
    Private pleEntryType As ProductLineEntryType = ProductLineEntryType.Search
    'Private _ProductStatus As Integer = 0
    Private blnLoading As Boolean

    Private oCopyFrom As cImprint
    Private intCopyRouteID As Integer 'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
    Private strCopyRoute As String ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
    Private intProductProof As Integer

    ' Will trace the items sequence of selection
    Private m_SelectSeq As Integer = 0

    Private m_Source As OEI_SourceEnum = OEI_SourceEnum.frmProductLineEntry

    Private overrideCellValidate As Boolean = False
    Private _25percent As Double = 25  'applied to 3ST and lower
    Private _50percent As Double = 50  'applied to 4ST and higher
    Private _100percent As Double = 100 'applied to all star levels if applicable

    Delegate Sub MyDelegate()
    Delegate Sub SetColumnIndex(ByVal i As Integer)
    Delegate Sub SetQty_To_Ship(ByVal pdblQty As Double)
    Delegate Sub AddToOrderDelegate()

    Public Enum ErrorType
        Reset = 0
        NoError = 0
        MustBeNumeric
        MustBeString
        MustBeDate
        IsEmpty
        MustBePositive
    End Enum

    Public Enum Columns

        FirstCol = 0
        Line_Seq_No
        Image
        Item_No
        Location
        Item_Desc_1_View
        Item_Desc_2_View
        Qty_Ordered
        Qty_To_Ship
        Qty_Lines
        Qty_Inventory
        Qty_Allocated
        Qty_On_Hand
        Qty_Prev_To_Ship
        Qty_Prev_BkOrd ' Add to properties
        Extra_9
        Extra_10 ' Add field here
        Unit_Price ' ?? not in Ordline ??
        Discount_Pct
        Calc_Price
        Request_Dt
        Promise_Dt
        Req_Ship_Dt
        Item_Desc_Charge
        Item_Desc_1
        Item_Desc_2
        Extra_1
        Imprint ' ---------------
        End_User
        Imprint_Location
        Imprint_Color
        Num_Impr_1
        Num_Impr_2
        Num_Impr_3

        '+++ 06.08.2020 ID------- Scribl project ---------
        Lamination
        '++ID 12.09.2021 Foil Color
        Foil_Color

        Thread_Color
        Rounded_Corners
        '++ID 08.06.2020 new field for Scribl Project ----
        Tip_In_PageTxt
        '----------------------------

        Packaging
        Refill
        Laser_Setup
        Industry
        Comments
        Special_Comments ' ---------------
        Printer_Comment
        Packer_Comment
        Printer_Instructions
        Packer_Instructions
        Repeat_From_ID
        RouteID
        Route
        ProductProof
        UOM
        Tax_Sched
        Revision_No
        BkOrd_Fg ' BackorderItem
        Recalc_SW ' Recalc
        Prod_Cat
        Kit_Feat_Fg
        Stocked_Fg
        OEI_W_Pixels
        OEI_H_Pixels
        OEI_Image_Rotate
        SelectSeq
        Item_Cd
        Loc_Status
        IsMixMatch
        Mix_Group
        LastCol
    End Enum

    'Public Enum _NewColumns
    '    _Image = 0
    '    _Item_No
    '    _Description
    '    _Description2
    '    _Qty_Ordered
    '    _Qty_To_Ship
    '    _Unit_Price  ' Add to properties
    '    _Qty_On_Hand
    '    _Qty_Prev_To_Ship ' Add to properties
    '    _Qty_Prev_BkOrd
    '    _TypeOfRefill
    '    _StandardRefill     ' check box
    '    _Request_Dt
    '    _Promise_Dt
    '    _LogoSetupCharge    ' Logo/imprint
    '    _AddLogoSetupCharge
    '    _LocSetupCharge
    '    _AddLocSetupCharge
    '    _ColorSetupCharge
    '    _AddColorSetupCharge
    '    _LowContrastNote
    '    _StandardPackaging  ' Checkbox
    '    _PackagingOptions   ' Combobox
    '    _PackagingQty       ' populates on qty
    '    _MiscCharge
    '    _ColorMatchCharge
    '    _PaperProof         ' blind, yes or no
    '    _VirtualProof       ' blind, yes or no
    '    _Route
    'End Enum

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub InitializeDataGridView()

        Try
            dgvItems.ColumnHeadersHeight = 15

            dgvItems.Columns.Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("Line_Seq_No", "Line", 0))
            dgvItems.Columns.Add(DGVImageColumn("Image", "Image", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_No", "Item", 125))
            dgvItems.Columns.Add(DGVTextBoxColumn("Location", "Loc", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Desc_1_View", "Description", 200))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Desc_2_View", "Item Description Line 2", 200))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_Ordered", "Qty Ordered", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_To_Ship", "Qty Shipped", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_Lines", "Qty Lines", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_Inventory", "Inventory", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_Allocated", "Qty Alloc", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_On_Hand", "Qty On Hand", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_Prev_To_Ship", "Preview Ship", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_Prev_BkOrd", "Preview Bk/Order", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Extra_9", "ETA Date", 80))
            dgvItems.Columns.Add(DGVTextBoxColumn("Extra_10", "ETA Qty", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Unit_Price", "Unit Price", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Discount_Pct", "Disc Pct", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Calc_Price", "Ext Price", 70))
            dgvItems.Columns.Add(DGVCalendarColumn("Request_Dt", "Requested", 100))
            dgvItems.Columns.Add(DGVCalendarColumn("Promise_Dt", "Promised", 100))
            dgvItems.Columns.Add(DGVCalendarColumn("Req_Ship_Dt", "Shipped", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Desc_Charge", "Charge Description", 300))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Desc_1", "Description", 200))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Desc_2", "Item Description Line 2", 200))

            dgvItems.Columns.Add(DGVCheckBoxColumn("Extra_1", "Exact Qty", 70))

            dgvItems.Columns.Add(DGVTextBoxColumn("Imprint", "Imprint", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("End_User", "End User", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Imprint_Location", "Imprint Location", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Imprint_Colour", "Imprint Colour", 75))
            dgvItems.Columns.Add(DGVTextBoxColumn("Num_Impr_1", "Num Impr 1", 75))
            dgvItems.Columns.Add(DGVTextBoxColumn("Num_Impr_2", "Num Impr2", 75))
            dgvItems.Columns.Add(DGVTextBoxColumn("Num_Impr_3", "Num Impr3", 75))


            ' +++ID 06.08.2020 --------------- column for scrable ----------------------
            Dim comboboxColumn As DataGridViewComboBoxColumn
            'dgvItems.Columns.Add(DGVTextBoxColumn("Industry", "Industry", 100))
            comboboxColumn = DGVComboBoxColumn("Lamination_Scribl", "Lamination", 120)
            With comboboxColumn
                .DataSource = m_oOrder.Get_Scribl_Fields("LAMINATION_SCRIBL")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With
            dgvItems.Columns.Add(comboboxColumn)
            '++ID 12.09.2021 Foil Color
            comboboxColumn = New DataGridViewComboBoxColumn
            'dgvItems.Columns.Add(DGVTextBoxColumn("Industry", "Industry", 100))
            comboboxColumn = DGVComboBoxColumn("Foil_Color", "Foil Color", 120)
            With comboboxColumn
                .DataSource = m_oOrder.Get_Scribl_Fields("Foil_Color")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With
            dgvItems.Columns.Add(comboboxColumn)

            '+++ID 06.08.2020 declare new combobox object before used only one declaration from Industry 
            comboboxColumn = New DataGridViewComboBoxColumn
            'dgvItems.Columns.Add(DGVTextBoxColumn("Industry", "Industry", 100))
            comboboxColumn = DGVComboBoxColumn("Thread_Color_Scribl", "Thread Color", 120)
            With comboboxColumn
                .DataSource = m_oOrder.Get_Scribl_Fields("Thread_Color_Scribl")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With
            dgvItems.Columns.Add(comboboxColumn)

            ' dgvItems.Columns.Add(DGVCheckBoxColumn("Rounded_Corners_Scribl", "Rounded corners ", 70))
            '+++ID 06.29.2020 changed properties from Checkbox to Combobox 
            comboboxColumn = New DataGridViewComboBoxColumn
            'dgvItems.Columns.Add(DGVTextBoxColumn("Industry", "Industry", 100))
            comboboxColumn = DGVComboBoxColumn("Rounded_Corners_Scribl", "Rounded Color", 120)
            With comboboxColumn
                .DataSource = m_oOrder.Get_Scribl_Fields("Rounded_Corners_Scribl")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With
            dgvItems.Columns.Add(comboboxColumn)

            'ID 07.28.2020 new checkbox-------------------------------------------------
            '   dgvItems.Columns.Add(DGVCheckBoxColumn("Tip_In_Page", "Tip In Page", 70))
            '---------------------------------------------------------------------------
            '++ID 08.06.2020 Tip_In_PageTxt
            '+++ID 08.05.2020 changed properties from Checkbox to Combobox 
            comboboxColumn = New DataGridViewComboBoxColumn
            'dgvItems.Columns.Add(DGVTextBoxColumn("Industry", "Industry", 100))
            comboboxColumn = DGVComboBoxColumn("Tip_In_PageTxt", "Tip In Page", 120)
            With comboboxColumn
                .DataSource = m_oOrder.Get_Scribl_Fields("Tip_In_PageTxt")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With
            dgvItems.Columns.Add(comboboxColumn)
            '-----------------------------------------------------------------------------------

            dgvItems.Columns.Add(DGVTextBoxColumn("Packaging", "Packaging", 75))
            dgvItems.Columns.Add(DGVTextBoxColumn("Refill", "Refill", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Laser_Setup", "Laser Setup", 100))

            comboboxColumn = New DataGridViewComboBoxColumn
            'dgvItems.Columns.Add(DGVTextBoxColumn("Industry", "Industry", 100))
            comboboxColumn = DGVComboBoxColumn("Industry", "Industry", 200)
            With comboboxColumn
                .DataSource = m_oOrder.GetCompleteCboIndustryList()
                .ValueMember = "Industry"
                .DisplayMember = "Industry"
            End With
            dgvItems.Columns.Add(comboboxColumn)

            'dgvItems.Columns.Add(DGVTextBoxColumn("Industry", "Industry", 100))

            dgvItems.Columns.Add(DGVTextBoxColumn("Comments", "Comments", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Special_Comments", "Spec. Comments", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Printer_Comment", "Printer Comment", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Packer_Comment", "Packer Comment", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Printer_Instructions", "Printer _Instructions", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Packer_Instructions", "Packer Instructions", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Repeat_From_ID", "Repeat_From_ID", 0))

            dgvItems.Columns.Add(DGVTextBoxColumn("RouteID", "RouteID", 0))

            comboboxColumn = DGVComboBoxColumn("Route", "Route", 200)
            With comboboxColumn
                .DataSource = m_oOrder.GetCompleteCboRouteList()
                .ValueMember = "Route"
                .DisplayMember = "Route"
            End With
            dgvItems.Columns.Add(comboboxColumn)

            dgvItems.Columns.Add(DGVCheckBoxColumn("ProductProof", "Product Proof", 70))

            dgvItems.Columns.Add(DGVTextBoxColumn("UOM", "UOM", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Tax_Sched", "Tax Schd", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Revision_No", "ECM Revision", 70))
            dgvItems.Columns.Add(DGVCheckBoxColumn("BkOrd_Fg", "Backorder Item", 70))
            dgvItems.Columns.Add(DGVCheckBoxColumn("Recalc_SW", "Recalc", 70))

            dgvItems.Columns.Add(DGVTextBoxColumn("Prod_Cat", "Prod_Cat Comments", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("Kit_Feat_Fg", "Kit_Feat_Fg", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("Stocked_Fg", "Stocked_Fg", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("OEI_W_Pixels", "OEI_W_Pixels", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("OEI_H_Pixels", "OEI_H_Pixels", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("Rotate", "Rotate", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("SelectSeq", "SelectSeq", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Cd", "Item_Cd", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Loc_Status", "Loc_Status", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("IsMixMatch", "IsMixMatch", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Mix_Group", "Mix_Group", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("LastCol", "LastCol", 0))

            dgvItems.Columns(Columns.FirstCol).Visible = False
            dgvItems.Columns(Columns.Line_Seq_No).Visible = False
            dgvItems.Columns(Columns.RouteID).Visible = False
            'dgvItems.Columns(Columns.Laser_Setup).Visible = False
            dgvItems.Columns(Columns.Prod_Cat).Visible = False
            dgvItems.Columns(Columns.Kit_Feat_Fg).Visible = False
            dgvItems.Columns(Columns.Stocked_Fg).Visible = False

            dgvItems.Columns(Columns.OEI_W_Pixels).Visible = False
            dgvItems.Columns(Columns.OEI_H_Pixels).Visible = False
            dgvItems.Columns(Columns.OEI_Image_Rotate).Visible = False
            dgvItems.Columns(Columns.SelectSeq).Visible = False
            dgvItems.Columns(Columns.Item_Cd).Visible = False
            dgvItems.Columns(Columns.Loc_Status).Visible = True
            dgvItems.Columns(Columns.Mix_Group).Visible = True

            dgvItems.Columns(Columns.Item_Desc_1_View).Visible = False
            dgvItems.Columns(Columns.Item_Desc_2_View).Visible = False
            dgvItems.Columns(Columns.Extra_1).Visible = False
            dgvItems.Columns(Columns.Extra_1).Visible = False
            'dgvItems.Columns(Columns.Extra_9).Visible = False
            'dgvItems.Columns(Columns.Extra_10).Visible = False

            dgvItems.Columns(Columns.Qty_Inventory).Visible = False
            dgvItems.Columns(Columns.Qty_Allocated).Visible = False
            dgvItems.Columns(Columns.Qty_Prev_To_Ship).Visible = False

            dgvItems.Columns(Columns.Printer_Comment).Visible = False
            'dgvItems.Columns(Columns.Printer_Instructions).Visible = False
            dgvItems.Columns(Columns.Packer_Comment).Visible = False
            'dgvItems.Columns(Columns.Packer_Instructions).Visible = False
            dgvItems.Columns(Columns.Repeat_From_ID).Visible = False

            dgvItems.Columns(Columns.IsMixMatch).Visible = False
            dgvItems.Columns(Columns.LastCol).Visible = False

            For lPos As Integer = 0 To dgvItems.Columns.Count - 1
                dgvItems.Columns(lPos).SortMode = DataGridViewColumnSortMode.NotSortable
            Next lPos

            If dgvItems.Columns(0).Width < 120 Then dgvItems.Columns(0).Width = 120

            For Each myRow As DataGridViewRow In dgvItems.Rows
                For lPos As Integer = 1 To myRow.Cells.Count - 1
                    myRow.Cells(lPos).Tag = myRow.Cells(lPos).Value
                Next lPos
            Next

            dgvItems.CausesValidation = True

            Dim oCellStyle As New DataGridViewCellStyle()
            oCellStyle.Format = "##,##0.000000"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvItems.Columns(Columns.Unit_Price).DefaultCellStyle = oCellStyle

            oCellStyle = New DataGridViewCellStyle()
            oCellStyle.Format = "##,##0.00"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvItems.Columns(Columns.Discount_Pct).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Calc_Price).DefaultCellStyle = oCellStyle

            dgvItems.Columns(Columns.Qty_Ordered).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_To_Ship).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Lines).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Inventory).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Allocated).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_On_Hand).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Prev_To_Ship).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Prev_BkOrd).DefaultCellStyle = oCellStyle

            dgvItems.Columns(Columns.FirstCol).Frozen = True
            dgvItems.Columns(Columns.Line_Seq_No).Frozen = True
            dgvItems.Columns(Columns.Image).Frozen = True
            dgvItems.Columns(Columns.Item_No).Frozen = True

            'oCellStyle = New DataGridViewCellStyle()
            'oCellStyle.Format = "d"
            'oCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            'dgvItems.Columns(Columns.Promise_Dt).DefaultCellStyle = oCellStyle
            'dgvItems.Columns(Columns.Request_Dt).DefaultCellStyle = oCellStyle
            'dgvItems.Columns(Columns.Req_Ship_Dt).DefaultCellStyle = oCellStyle

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub CreateTextBoxColumn(ByRef dgvColumn As DataGridViewColumn, ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0)

        dgvColumn.HeaderText = pstrHeaderText
        If plWidth <> 0 Then dgvColumn.Width = plWidth

    End Sub

    Private Function CreateImageColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, ByVal plWidth As Long) As DataGridViewImageColumn

        CreateImageColumn = New DataGridViewImageColumn
        CreateImageColumn.Name = pstrName
        CreateImageColumn.HeaderText = pstrHeaderText

    End Function

    Private Sub SetColumnIndexSub(ByVal columnIndex As Integer)

        dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(columnIndex)
        'dgvItems.BeginEdit(True)

    End Sub

    Private Sub SetQty_To_ShipSub(ByVal pdblQty As Double)

        dgvItems.CurrentRow.Cells(Columns.Qty_To_Ship).Value = pdblQty

    End Sub

    Public Sub LoadHumRes()

        Dim strSql As String

        Try
            strSql = "" & _
            "SELECT 	Image, FullName, 0 AS WIDTH, 0 AS HEIGHT " & _
            "FROM 		HumRes WITH (Nolock) " & _
            "ORDER BY   FullName "

            Dim db As New cDBA

            ''Dim myConnection As SqlConnection = New SqlConnection("Server=thor;Initial Catalog=100;Integrated Security=SSPI")
            'Dim myConnection As SqlConnection = New SqlConnection("Server=zeus;Initial Catalog=100;Integrated Security=SSPI")
            'Dim mySelectCommand As SqlCommand = New SqlCommand(strSql, myConnection)
            'Dim mySqlDataAdapter As SqlDataAdapter = New SqlDataAdapter(mySelectCommand)
            'Dim myDataSet As New DataSet()
            'mySqlDataAdapter.Fill(myDataSet)
            dgvItems.DataSource = Nothing
            'dgvItems.DataSource = myDataSet.Tables(0)
            dgvItems.DataSource = db.DataTable(strSql)

            Dim oImage As Image
            Dim data As Byte()
            Dim ms As MemoryStream

            For Each myRow As DataGridViewRow In dgvItems.Rows

                If Not (myRow.Cells(Columns.Image).Value.Equals(System.DBNull.Value)) Then
                    data = myRow.Cells(Columns.Image).Value ' (byte[]) dt.Rows[0]["IMAGE"];
                    ms = New MemoryStream(data)
                    oImage = Image.FromStream(ms)

                    myRow.Cells(2).Value = oImage.Width
                    myRow.Cells(3).Value = oImage.Height
                End If

            Next myRow

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub LoadProductGrid()

        Dim strSql As String = ""
        Dim intCountItems As Integer = 0
        Dim intCountCharges As Integer = 0

        If RTrim(Trim(Item_No)).Length < 4 And Not (Item_No = "44" Or Item_No = "66" Or Item_No = "88") Then Exit Sub

        'If (Mid(Item_No, 1, 2) <> "00") And (Not IsNumeric(Mid(Item_No, 1, 2))) Then Item_No = "00" & Item_No

        Try
            InitializeDataGridView()

            'Dim dtRequestDt As DateTime = IIf(FormatDateTime(m_oOrder.Ordhead.Shipping_Dt, DateFormat.ShortDate) = "01/01/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
            'Dim dtRequestDt As DateTime = IIf(m_oOrder.Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") = "01/01/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
            'Dim dtRequestDt As DateTime = IIf(m_oOrder.Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") = "01/01/0001", Now, m_oOrder.Ordhead.Shipping_Dt)
            'Dim RequestDt As String
            'RequestDt = dtRequestDt.Month.ToString.PadLeft(2, "0") & "/" & dtRequestDt.Day.ToString.PadLeft(2, "0") & "/" & dtRequestDt.Year.ToString.PadLeft(4, "0")

            Dim dtRequestDt As DateTime
            Dim RequestDt As String
            If m_oOrder.Ordhead.In_Hands_Date.Year = 1 Then
                dtRequestDt = m_oOrder.Ordhead.Ord_Dt_Shipped ' Now.Date ' DateAdd("d", 1, Now)
            Else
                dtRequestDt = m_oOrder.Ordhead.In_Hands_Date
            End If
            RequestDt = dtRequestDt.Month.ToString.PadLeft(2, "0") & "/" & dtRequestDt.Day.ToString.PadLeft(2, "0") & "/" & dtRequestDt.Year.ToString.PadLeft(4, "0")

            'Dim dtShippingDt As DateTime = IIf(m_oOrder.Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") = "01/01/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
            Dim dtShippingDt As DateTime = IIf(m_oOrder.Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") = "01/01/0001", Now, m_oOrder.Ordhead.Shipping_Dt)
            Dim ShippingDt As String
            ShippingDt = dtShippingDt.Month.ToString.PadLeft(2, "0") & "/" & dtShippingDt.Day.ToString.PadLeft(2, "0") & "/" & dtShippingDt.Year.ToString.PadLeft(4, "0")

            'Do not use FormatDateTime. If so, then rewrite the function to force MM/DD/YYYY And bypass Short date errors for SQL syntax.
            'RequestDt = FormatDateTime(IIf(FormatDateTime(m_oOrder.Ordhead.Shipping_Dt, DateFormat.ShortDate) = "1/1/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt), DateFormat.ShortDate)

            If Item_No = "CHARGESCALCULATION" Then
                'TODO: switch to finalized master database if wanted in the future
                'Call m_oOrder.Create_Preview_Item_Charges()
                'never load view OEI_CHARGES_PREVIEW doesn't exist
                strSql = "" &
                "SELECT 	'' as FirstCol, '' as Line_Seq_No, P.Picture AS Image, I.Item_No, '" & Loc & "' AS Location, " &
                "			ISNULL(I.Item_Desc_1, '') AS Item_Desc_1_View, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2_View, " &
                "           ISNULL(C.Qty_Ordered, 0) AS Qty_Ordered, ISNULL(C.Qty_Ordered, 0) AS Qty_To_Ship, CAST(1 AS FLOAT) AS Qty_Lines, " &
                "           ISNULL(L.Qty_On_Hand,0) " & IIf(Loc.Trim = "1", "- DBO.OEI_INV_BUFFER(I.Item_No, 'CAD')", "") & " AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) AS Qty_Allocated, " &
                "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0) " & IIf(Loc.Trim = "1", "- DBO.OEI_INV_BUFFER(I.Item_No, 'CAD')", "") & ") AS FLOAT) AS Qty_On_Hand, " &
                "           ISNULL(C.Qty_Ordered, 0) AS Qty_Prev_To_Ship, CAST(0 AS FLOAT) AS Qty_Prev_BkOrd, " &
                "           CAST('' AS CHAR(12)) AS Extra_9, CAST(0 AS FLOAT) AS Extra_10, " &
                "           DBO.OEI_Item_Price_20140101('" & Currency & "', '" & Customer_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &
                "           CAST(0 AS FLOAT) AS Discount_Pct, ISNULL(C.Unit_Price, 0) * ISNULL(C.Qty_Ordered, 0) AS Calc_Price, " &
                "           CAST('" & RequestDt & "' AS DATETIME) AS Request_DT, " &
                "           CAST('" & ShippingDt & "' AS DATETIME) AS Promise_Dt, " &
                "           CAST('" & ShippingDt & "' AS DATETIME) AS Req_Ship_Dt, " &
                "           ISNULL(C.Description, '') AS Item_Desc_Charge, ISNULL(I.Item_Desc_1, '') AS Item_Desc_1, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2, '' AS Extra_1, " &
                "           '' AS Imprint, '' AS End_User, '' AS Imprint_Location, '' AS Imprint_Colour, 0 AS Num_Impr_1, 0 AS Num_Impr_2, 0 AS Num_Impr_3, " &
                "           ISNULL(PK.Pack_CD, '') AS Packaging, '' AS Refill, '' AS Laser_Setup, '' AS Industry, '' AS Comments, '' AS Special_Comments, " &
                "           '' AS Printer_Comment, '' AS Packer_Comment, '' AS Printer_Instructions, '' AS Packer_Instructions, CAST(0 AS INT) AS Repeat_From_ID, " &
                "           0 AS RouteID, '' AS Route, CAST(0 AS BIT) AS ProductProof, ISNULL(I.UOM, '') AS UOM, '' AS Tax_Sched, '' AS Revision_No, CAST (0 AS BIT) AS BkOrd_Fg, " &
                "           CAST(1 AS BIT) AS Recalc_SW, " &
                "           ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, ISNULL(I.Stocked_Fg, '') AS Stocked_Fg, CAST(400 AS FLOAT) AS OEI_W_Pixels, CAST(56 AS FLOAT) AS OEI_H_Pixels, '' AS Rotate, 0 AS SelectSeq, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd, L.status as Loc_Status, 0 as Mix_Group, '' as LastCol " &
                "FROM 		IMITMIDX_SQL I WITH (nolock) " &
                "INNER JOIN	OEI_CHARGES_PREVIEW C WITH (Nolock) ON I.Item_No = C.Item_No " &
                "INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " &
                "LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " &
                "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                "LEFT JOIN  MDB_ITEM_MASTER MD WITH (Nolock) ON MD.ITEM_CD = I.user_def_fld_1 " &
                "LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MD.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
                "LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
                "WHERE 		C.Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND I.ACTIVITY_CD = 'A' " &
                "ORDER BY   C.ID " ' I.Item_No "







            Else

                '----------------------------------------------------------------------------------------
                'strSql = ";with validateitem as (
                '            select distinct item_no from [100].[dbo].[MWAV_INVENTORY_LEVELS_MIXMATCH_ion_New1.07.2019] where kititemno like '%" & Item_No & "%' )"


                '++ID 2.08.2019 modified the query added union for mix and match
                ' I put in comments until Pinky approve for update, start test 2.11.2019 until 02.18.2019
                '++ID 2.20.2019 removed the comments 
                '-----------------------------------------------------------------------------------------
                strSql = "Select * from ( "
                strSql &= "SELECT 	'' as FirstCol, '' as Line_Seq_No, P.Picture AS Image, I.Item_No,  ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'), '' ) As Location, " &
       "           ISNULL(I.Item_Desc_1, '') AS Item_Desc_1_View, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2_View, " &
       "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, CAST(0 AS FLOAT) AS Qty_Lines," &
       " " &      '   " CASE  WHEN ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'), '' )  = '1' THEN - DBO.OEI_INV_BUFFER(I.Item_No, 'CAD') ELSE  "" END AS Qty_Inventory, " &                          
       " Case  WHEN ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'),'')  in ('1','US1') " & '++ID 02.17.2023 was = 1 canged to  in ('1','US1')
       " THEN ISNULL(L.Qty_On_Hand,0)  - DBO.OEI_INV_BUFFER(I.Item_No, 'CAD') ELSE  ISNULL(L.Qty_On_Hand,0)   END AS Qty_Inventory,  " &
       " " &
        " " &
       " ISNULL(L.Qty_Allocated,0) AS Qty_Allocated, " &
       "  " & '" CASE  WHEN ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'), '' ) AS FLOAT) AS Qty_On_Hand, " &
       " " &
       " CAST((  CASE WHEN ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'),'')  in ('1','US1')  " & '++ID 02.17.2023 was = 1 canged to  in ('1','US1')
       " THEN  ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0) - DBO.OEI_INV_BUFFER(I.Item_No, 'CAD') ELSE " &
       " ISNULL(L.Qty_On_Hand, 0) - ISNULL(L.Qty_Allocated, 0) - ISNULL(PQV.Qty_Ordered, 0) " &
       "  End)  As FLOAT) As Qty_On_Hand,  " &
       " " &
       "  CAST(0 AS FLOAT) AS Qty_Prev_To_Ship, CAST(0 AS FLOAT) AS Qty_Prev_BkOrd, CAST('' AS CHAR(12)) AS Extra_9, CAST(0 AS FLOAT) AS Extra_10, " &
       " DBO.OEI_Item_Price_20140101('" & Currency & "', '" & Customer_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &
       " CAST(" & m_oOrder.Ordhead.Discount_Pct & " AS FLOAT) AS Discount_Pct, CAST(0 AS FLOAT) AS Calc_Price, " &
       " CAST('" & RequestDt & "' AS DATETIME) AS Request_DT, " &
       " CAST('" & ShippingDt & "' AS DATETIME) AS Promise_Dt, " &
       " CAST('" & ShippingDt & "' AS DATETIME) AS Req_Ship_Dt, " &
       "           " &
       "           '' AS Item_Desc_Charge, ISNULL(I.Item_Desc_1, '') AS Item_Desc_1, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2, '' AS Extra_1, " &
       "           '' AS Imprint, '' AS End_User, '' AS Imprint_Location, '' AS Imprint_Colour, 0 AS Num_Impr_1, 0 AS Num_Impr_2, 0 AS Num_Impr_3, " &
       "  '' AS Lamination_Scribl, '' AS Foil_Color, '' AS Thread_Color_Scribl, '' AS Rounded_corners_Scribl, '' As Tip_In_PageTxt,  " & '++ID 06.08.2020 scribl fields Tip_In_PageTxt was added 08.06.2020 ID
       "           ISNULL(PK.Pack_CD, '') AS Packaging, '' AS Refill, '' AS Laser_Setup, " &
       "           '" & IIf(m_oOrder.Ordhead.User_Def_Fld_2 = "SP", "Self Promo", "") & "' AS Industry, " &
       "           '' AS Comments, '' AS Special_Comments, " &
       "           '' AS Printer_Comment, '' AS Packer_Comment, '' AS Printer_Instructions, '' AS Packer_Instructions, CAST(0 AS INT) AS Repeat_From_ID, " &
       "           0 AS RouteID, '' AS Route, CAST(0 AS BIT) AS ProductProof, ISNULL(I.UOM, '') AS UOM, '' AS Tax_Sched, '' AS Revision_No, CASE ISNULL(I.BkOrd_FG, '') WHEN 'Y' THEN CAST(1 AS BIT) ELSE CAST (0 AS BIT) END AS BkOrd_Fg, " &
       "           CAST(1 AS BIT) AS Recalc_SW, " &
       "           ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, ISNULL(I.Stocked_Fg, '') AS Stocked_Fg, ISNULL(P.OEI_W_Pixels, 400) AS OEI_W_Pixels, ISNULL(P.OEI_H_Pixels, 400) AS OEI_H_Pixels, ISNULL(P.Rotate, '') AS Rotate, " &
       "           0 AS SelectSeq, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd, L.status as Loc_Status, ISNULL(DS.fldValue, 0) as IsMixMatch, 0 as Mix_Group, '' as LastCol "

                '         strSql = strSql & "    FROM 	   IMITMIDX_SQL I WITH (nolock) " &
                '"INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " &
                '"LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc AND PQV.Loc = '" & Loc & "' " &
                '"LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                '"LEFT JOIN  MDB_ITEM_MASTER MD WITH (Nolock) ON MD.ITEM_CD = I.user_def_fld_1 " &
                '"LEFT JOIN  MDB_ITEM_DYNAMIC_FLD_VALUE_STR DS WITH (Nolock) ON DS.ITEM_MASTER_ID = MD.ITEM_MASTER_ID AND DS.FLD_CD = 'MIX_AND_MATCH' " &
                '"LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MD.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
                '"LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
                '"WHERE 		I.Item_No LIKE '" & IIf(Item_No = "44" Or Item_No = "66" Or Item_No = "88", "", "%") & Item_No & "%' AND I.ACTIVITY_CD = 'A' "
                '         strSql = strSql & "AND I.pur_or_mfg <> 'M'  "    'Added July 13, 2017 T. Louzon
                '         '  strSql = strSql & "ORDER BY   I.Item_No "

                '++IS 02.17.2023 commented above and added below  for to filter by location 
                If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then
                    strSql &= "    FROM 	   IMITMIDX_SQL I WITH (nolock) " &
       "INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " &
       "LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc AND PQV.Loc = '" & Loc & "' " &
       "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
       "LEFT JOIN  MDB_ITEM_MASTER MD WITH (Nolock) ON MD.ITEM_CD = I.user_def_fld_1 " &
       "LEFT JOIN  MDB_ITEM_DYNAMIC_FLD_VALUE_STR DS WITH (Nolock) ON DS.ITEM_MASTER_ID = MD.ITEM_MASTER_ID AND DS.FLD_CD = 'MIX_AND_MATCH' " &
       "LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MD.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
       "LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
       "WHERE 		I.Item_No LIKE '" & IIf(Item_No = "44" Or Item_No = "66" Or Item_No = "88", "", "%") & Item_No & "%' AND I.ACTIVITY_CD = 'A' "
                    strSql = strSql & "AND I.pur_or_mfg <> 'M'  "    'Added July 13, 2017 T. Louzon
                    '  strSql = strSql & "ORDER BY   I.Item_No "
                Else
                    strSql &= " FROM [200].dbo.IMITMIDX_SQL I WITH (nolock) " &
       "INNER JOIN [200].dbo.OEI_Item_Loc_Qty_View_US L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " &
       "LEFT JOIN  [200].dbo.OEI_ORDLIN_PENDING_QTY_VIEW_US PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc AND PQV.Loc = '" & Loc & "' " &
       "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
       "LEFT JOIN  MDB_ITEM_MASTER MD WITH (Nolock) ON MD.ITEM_CD = I.user_def_fld_1 " &
       "LEFT JOIN  MDB_ITEM_DYNAMIC_FLD_VALUE_STR DS WITH (Nolock) ON DS.ITEM_MASTER_ID = MD.ITEM_MASTER_ID AND DS.FLD_CD = 'MIX_AND_MATCH' " &
       "LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MD.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
       "LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
       "WHERE 		I.Item_No LIKE '" & IIf(Item_No = "44" Or Item_No = "66" Or Item_No = "88", "", "%") & Item_No & "%' AND I.ACTIVITY_CD = 'A' "
                    strSql = strSql & "AND I.pur_or_mfg <> 'M'  "    'Added July 13, 2017 T. Louzon
                    '  strSql = strSql & "ORDER BY   I.Item_No "
                End If



                strSql &= " UNION ALL "

                strSql &= "SELECT 	'' as FirstCol, '' as Line_Seq_No, P.Picture AS Image, I.Item_No,  ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'), '' ) As Location, " &
       "           ISNULL(I.Item_Desc_1, '') AS Item_Desc_1_View, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2_View, " &
       "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, CAST(0 AS FLOAT) AS Qty_Lines," &
       " " &      '   " CASE  WHEN ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'), '' )  = '1' THEN - DBO.OEI_INV_BUFFER(I.Item_No, 'CAD') ELSE  "" END AS Qty_Inventory, " &                          
       " Case  WHEN ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'),'') in ('1','US1') " & '++ID 02.17.2023 was = 1 canged to  in ('1','US1')
       " THEN ISNULL(L.Qty_On_Hand,0)  - DBO.OEI_INV_BUFFER(I.Item_No, 'CAD') ELSE  ISNULL(L.Qty_On_Hand,0)   END AS Qty_Inventory,  " &
       " " &
       " ISNULL(L.Qty_Allocated,0) AS Qty_Allocated, " &
       "  " & '" CASE  WHEN ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'), '' ) AS FLOAT) AS Qty_On_Hand, " &
       " " &
       " CAST((  CASE WHEN ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(Loc) & "'),'') in ('1','US1') " & '++ID 02.17.2023 was = 1 canged to  in ('1','US1')
       " THEN  ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0) - DBO.OEI_INV_BUFFER(I.Item_No, 'CAD') ELSE " &
       " ISNULL(L.Qty_On_Hand, 0) - ISNULL(L.Qty_Allocated, 0) - ISNULL(PQV.Qty_Ordered, 0) " &
       "  End)  As FLOAT) As Qty_On_Hand,  " &
       " " &
       "  CAST(0 AS FLOAT) AS Qty_Prev_To_Ship, CAST(0 AS FLOAT) AS Qty_Prev_BkOrd, CAST('' AS CHAR(12)) AS Extra_9, CAST(0 AS FLOAT) AS Extra_10, " &
        " 0.00 as Unit_Price, " &   ' " DBO.OEI_Item_Price_20140101('" & Currency & "', '" & Customer_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &             
       " CAST(" & m_oOrder.Ordhead.Discount_Pct & " AS FLOAT) AS Discount_Pct, CAST(0 AS FLOAT) AS Calc_Price, " &
       " CAST('" & RequestDt & "' AS DATETIME) AS Request_DT, " &
       " CAST('" & ShippingDt & "' AS DATETIME) AS Promise_Dt, " &
       " CAST('" & ShippingDt & "' AS DATETIME) AS Req_Ship_Dt, " &
       "           " &
       "           '' AS Item_Desc_Charge, ISNULL(I.Item_Desc_1, '') AS Item_Desc_1, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2, '' AS Extra_1, " &
       "           '' AS Imprint, '' AS End_User, '' AS Imprint_Location, '' AS Imprint_Colour, 0 AS Num_Impr_1, 0 AS Num_Impr_2, 0 AS Num_Impr_3, " &
       "           '' AS Lamination_Scribl, '' AS Foil_Color, '' AS Thread_Color_Scribl, '' AS Rounded_Corners_Scribl, '' As Tip_In_PageTxt,  " & '++ID 06.08.2020 scribl fields Tip_In_PageTxt was added 08.06.2020 ID
       "           ISNULL(PK.Pack_CD, '') AS Packaging, '' AS Refill, '' AS Laser_Setup, " &
       "           '" & IIf(m_oOrder.Ordhead.User_Def_Fld_2 = "SP", "Self Promo", "") & "' AS Industry, " &
       "           '' AS Comments, '' AS Special_Comments, " &
       "           '' AS Printer_Comment, '' AS Packer_Comment, '' AS Printer_Instructions, '' AS Packer_Instructions, CAST(0 AS INT) AS Repeat_From_ID, " &
       "           0 AS RouteID, '' AS Route, CAST(0 AS BIT) AS ProductProof, ISNULL(I.UOM, '') AS UOM, '' AS Tax_Sched, '' AS Revision_No, CASE ISNULL(I.BkOrd_FG, '') WHEN 'Y' THEN CAST(1 AS BIT) ELSE CAST (0 AS BIT) END AS BkOrd_Fg, " &
       "           CAST(1 AS BIT) AS Recalc_SW, " &
       "           ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, ISNULL(I.Stocked_Fg, '') AS Stocked_Fg, ISNULL(P.OEI_W_Pixels, 400) AS OEI_W_Pixels, ISNULL(P.OEI_H_Pixels, 400) AS OEI_H_Pixels, ISNULL(P.Rotate, '') AS Rotate, " &
       "           0 AS SelectSeq, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd, L.status as Loc_Status, ISNULL(DS.fldValue, 0) as IsMixMatch, 0 as Mix_Group, '' as LastCol "


                '         strSql &= "   FROM 	IMITMIDX_SQL I WITH (nolock) " &
                '" INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " &
                '" LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc AND PQV.Loc = '" & Loc & "' " &
                '" LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                '" LEFT JOIN  MDB_ITEM_MASTER MD WITH (Nolock) ON MD.ITEM_CD = I.user_def_fld_1 " &
                '" LEFT JOIN  MDB_ITEM_DYNAMIC_FLD_VALUE_STR DS WITH (Nolock) ON DS.ITEM_MASTER_ID = MD.ITEM_MASTER_ID AND DS.FLD_CD = 'MIX_AND_MATCH' " &
                '" LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MD.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
                '" LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
                '" WHERE 			I.Item_No in (select distinct item_no from [100].[dbo].[MWAV_INVENTORY_LEVELS_MIXMATCH_ion_New1.07.2019] where kititemno like '%" & Item_No & "%') AND ISNULL(I.Kit_Feat_Fg, '') = '' AND I.ACTIVITY_CD = 'A' "
                '         strSql = strSql & "AND I.pur_or_mfg <> 'M'  "

                '++ID 02.17.2023 commented above for to add exception to filter sql by location 
                If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then
                    strSql &= "   FROM 	IMITMIDX_SQL I WITH (nolock) " &
           " INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " &
           " LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc AND PQV.Loc = '" & Loc & "' " &
           " LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
           " LEFT JOIN  MDB_ITEM_MASTER MD WITH (Nolock) ON MD.ITEM_CD = I.user_def_fld_1 " &
           " LEFT JOIN  MDB_ITEM_DYNAMIC_FLD_VALUE_STR DS WITH (Nolock) ON DS.ITEM_MASTER_ID = MD.ITEM_MASTER_ID AND DS.FLD_CD = 'MIX_AND_MATCH' " &
           " LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MD.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
           " LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
           " WHERE 			I.Item_No in (select distinct item_no from [100].[dbo].[MWAV_INVENTORY_LEVELS_MIXMATCH_ion_New1.07.2019] where kititemno like '%" & Item_No & "%') AND ISNULL(I.Kit_Feat_Fg, '') = '' AND I.ACTIVITY_CD = 'A' "
                    strSql &= " AND I.pur_or_mfg <> 'M'  "

                Else
                    strSql &= "   FROM 	[200].dbo.IMITMIDX_SQL I WITH (nolock) " &
           " INNER JOIN [200].dbo.OEI_Item_Loc_Qty_View_US L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " &
           " LEFT JOIN  [200].dbo.OEI_ORDLIN_PENDING_QTY_VIEW_US PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc AND PQV.Loc = '" & Loc & "' " &
           " LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
           " LEFT JOIN  MDB_ITEM_MASTER MD WITH (Nolock) ON MD.ITEM_CD = I.user_def_fld_1 " &
           " LEFT JOIN  MDB_ITEM_DYNAMIC_FLD_VALUE_STR DS WITH (Nolock) ON DS.ITEM_MASTER_ID = MD.ITEM_MASTER_ID AND DS.FLD_CD = 'MIX_AND_MATCH' " &
           " LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MD.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
           " LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
           " WHERE 			I.Item_No in (select distinct item_no from [100].[dbo].[MWAV_INVENTORY_LEVELS_MIXMATCH_ion_New1.07.2019] where kititemno like '%" & Item_No & "%') AND ISNULL(I.Kit_Feat_Fg, '') = '' AND I.ACTIVITY_CD = 'A' "
                    strSql &= " AND I.pur_or_mfg <> 'M'  "
                End If




                '  strSql = strSql & "ORDER BY   I.Item_No "
                strSql &= " ) as oei order by isnull(Kit_Feat_Fg,'') desc,ISNULL(IsMixMatch, 0) desc, Item_No "

                'after Pinky aprove need uncoment until here ----------------------------------------------------------|





            End If

            '"INNER JOIN IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '1 ' " & _

            ' We use the same entry window for substitutes and for item colors.
            ' Reference is passed prior 
            Select Case pleEntryType
                Case ProductLineEntryType.Search
                    ' use current strSql

                Case ProductLineEntryType.Substitutes
                    strSql = "" &
                    "SELECT 	P.Picture AS Image, I.Item_No, '" & Loc & "' AS Location, " &
                    "           ISNULL(I.Item_Desc_1, '') AS Item_Desc_1_View, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2_View, " &
                    "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, CAST(0 AS FLOAT) AS Qty_Lines, ISNULL(L.Qty_On_Hand,0) " & IIf(Loc.Trim = "1", "- DBO.OEI_INV_BUFFER(I.Item_No, 'CAD')", "") & " AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) AS Qty_Allocated, " &
                    "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0) " & IIf(Loc.Trim = "1", "- DBO.OEI_INV_BUFFER(I.Item_No, 'CAD')", "") & ") AS FLOAT) AS Qty_On_Hand, " &
                    "           CAST(0 AS FLOAT) AS Qty_Prev_To_Ship, CAST(0 AS FLOAT) AS Qty_Prev_BkOrd, CAST('' AS CHAR(12)) AS Extra_9, CAST(0 AS FLOAT) AS Extra_10, " &
                    "           DBO.OEI_Item_Price_20140101('" & Currency & "', '" & Customer_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &
                    "           CAST(" & m_oOrder.Ordhead.Discount_Pct & " AS FLOAT) AS Discount_Pct, CAST(0 AS FLOAT) AS Calc_Price, " &
                    "           CAST('" & RequestDt & "' AS DATETIME) AS Request_DT, " &
                    "           CAST('" & ShippingDt & "' AS DATETIME) AS Promise_Dt, " &
                    "           CAST('" & ShippingDt & "' AS DATETIME) AS Req_Ship_Dt, " &
                    "           ISNULL(I.Item_Desc_1, '') AS Item_Desc_1, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2, '' AS Extra_1, " &
                    "           ISNULL(I.UOM, '') AS UOM, '' AS Tax_Sched, '' AS Revision_No, CASE ISNULL(I.BkOrd_FG, '') WHEN 'Y' THEN CAST(1 AS BIT) ELSE CAST (0 AS BIT) END AS BkOrd_Fg, " &
                    "           CAST(1 AS BIT) AS Recalc_SW, 0 AS RouteID, '' AS Route, CAST(0 AS BIT) AS ProductProof, ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, ISNULL(I.Stocked_Fg, '') AS Stocked_Fg, ISNULL(P.OEI_W_Pixels, 400) AS OEI_W_Pixels, ISNULL(P.OEI_H_Pixels, 400) AS OEI_H_Pixels, ISNULL(P.Rotate, '') AS Rotate, 0 AS SelectSeq, " &
                    "           ISNULL(I.User_Def_Fld_1, '') AS Item_Cd, L.status as Loc_Status, 0 as Mix_Group, '' as LastCol " &
                    "FROM 		IMITMIDX_SQL I WITH (nolock) " &
                    "INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " &
                    "LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " &
                    "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                    "WHERE 		I.Item_No LIKE '%" & Item_No & "%' AND I.ACTIVITY_CD = 'A' " &
                    "ORDER BY   I.Item_No "
                    '"INNER JOIN IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '1 '" & _

                    '-----------------------------------------------------------------------
                    '++ID 02.17.2023
                    If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                        strSql = "" &
               "SELECT 	P.Picture AS Image, I.Item_No, '" & Loc & "' AS Location, " &
               "           ISNULL(I.Item_Desc_1, '') AS Item_Desc_1_View, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2_View, " &
               "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, CAST(0 AS FLOAT) AS Qty_Lines, ISNULL(L.Qty_On_Hand,0) " & IIf(Loc.Trim = "US1", "- DBO.OEI_INV_BUFFER(I.Item_No, 'CAD')", "") & " AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) AS Qty_Allocated, " &
               "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0) " & IIf(Loc.Trim = "US1", "- DBO.OEI_INV_BUFFER(I.Item_No, 'CAD')", "") & ") AS FLOAT) AS Qty_On_Hand, " &
               "           CAST(0 AS FLOAT) AS Qty_Prev_To_Ship, CAST(0 AS FLOAT) AS Qty_Prev_BkOrd, CAST('' AS CHAR(12)) AS Extra_9, CAST(0 AS FLOAT) AS Extra_10, " &
               "           DBO.OEI_Item_Price_20140101('" & Currency & "', '" & Customer_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &
               "           CAST(" & m_oOrder.Ordhead.Discount_Pct & " AS FLOAT) AS Discount_Pct, CAST(0 AS FLOAT) AS Calc_Price, " &
               "           CAST('" & RequestDt & "' AS DATETIME) AS Request_DT, " &
               "           CAST('" & ShippingDt & "' AS DATETIME) AS Promise_Dt, " &
               "           CAST('" & ShippingDt & "' AS DATETIME) AS Req_Ship_Dt, " &
               "           ISNULL(I.Item_Desc_1, '') AS Item_Desc_1, ISNULL(I.Item_Desc_2, '') AS Item_Desc_2, '' AS Extra_1, " &
               "           ISNULL(I.UOM, '') AS UOM, '' AS Tax_Sched, '' AS Revision_No, CASE ISNULL(I.BkOrd_FG, '') WHEN 'Y' THEN CAST(1 AS BIT) ELSE CAST (0 AS BIT) END AS BkOrd_Fg, " &
               "           CAST(1 AS BIT) AS Recalc_SW, 0 AS RouteID, '' AS Route, CAST(0 AS BIT) AS ProductProof, ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, ISNULL(I.Stocked_Fg, '') AS Stocked_Fg, ISNULL(P.OEI_W_Pixels, 400) AS OEI_W_Pixels, ISNULL(P.OEI_H_Pixels, 400) AS OEI_H_Pixels, ISNULL(P.Rotate, '') AS Rotate, 0 AS SelectSeq, " &
               "           ISNULL(I.User_Def_Fld_1, '') AS Item_Cd, L.status as Loc_Status, 0 as Mix_Group, '' as LastCol " &
               "FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) " &
               "INNER JOIN [200].dbo.OEI_Item_Loc_Qty_View_US L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " &
               "LEFT JOIN  [[200].dbo.OEI_ORDLIN_PENDING_QTY_VIEW_US PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " &
               "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
               "WHERE 		I.Item_No LIKE '%" & Item_No & "%' AND I.ACTIVITY_CD = 'A' " &
               "ORDER BY   I.Item_No "
                    End If


            End Select

            Dim chosenSKuSet As New ArrayList
            Dim db As New cDBA
            Dim genQty As Integer = 0
            Dim _dt As DataTable
            _dt = db.DataTable(strSql)


            dgvItems.DataSource = _dt ' db.DataTable(strSql)

            If m_oOrder.Ordhead.Mfg_Loc = "US1" And _dt.Rows.Count = 0 Then

                Dim OEI_PRE_READ_ORDER_EXCEPTION As cOEI_PRE_READ_ORDER_EXCEPTION = New cOEI_PRE_READ_ORDER_EXCEPTION(m_oOrder.Ordhead.Oe_Po_No, m_oOrder.Ordhead.OEI_Ord_No, Item_No, 2786)
                Dim OEI_PRE_READ_ORDER_EXCEPTION_DAL As cOEI_PRE_READ_ORDER_EXCEPTION_DAL = New cOEI_PRE_READ_ORDER_EXCEPTION_DAL()
                OEI_PRE_READ_ORDER_EXCEPTION_DAL.Save(OEI_PRE_READ_ORDER_EXCEPTION)


                MsgBox(" Item # " & Item_No & ", is not offered in Vegas. " & vbCrLf &
                       "  Please enter/switch to location 1 (Montreal). ")
            End If



            '++ID comments 2.08.2019 because changed the first logic.
            'before it was: the components of mix and match item need to have in item no same style carackterslike G975(00G975.02G975,10G975) now we don't want have same logic for mix.

            'If dgvItems.Rows(0).Cells(Columns.IsMixMatch).Value = 1 Then
            '    ' MsgBox("Is mix match")
            '    Dim oOrderMixSelectFilter As New frmMixMatchSelector()
            '    oOrderMixSelectFilter.mix_item_cd = dgvItems.Rows(0).Cells(Columns.Item_Cd).Value
            '    oOrderMixSelectFilter.item_loc = m_oOrder.Ordhead.Mfg_Loc

            '    oOrderMixSelectFilter.GenerateComponentGroups()
            '    oOrderMixSelectFilter.ShowDialog()
            '    ' Exit Sub

            '    chosenSKuSet = oOrderMixSelectFilter.getChosenSkuSet
            '    genQty = oOrderMixSelectFilter.getChosenInventory
            '    'oOrderMixSelectFilter.GenerateComponentGroups()

            '    'eliminate non chosen mix match components if the item selected is this type of item
            '    If Not chosenSKuSet Is Nothing AndAlso chosenSKuSet.Count > 0 Then
            '        For Each row As DataGridViewRow In dgvItems.Rows
            '            Dim isvalid As Boolean = False
            '            For Each item_no1 As String In chosenSKuSet
            '                'Dim curr_result_grid As DataGridView = dgvItems
            '                If item_no1 = row.Cells.Item(Columns.Item_No).Value.ToString Then
            '                    isvalid = True
            '                    Exit For
            '                End If
            '            Next
            '            If isvalid = False Then
            '                dgvItems.Rows.Item(row.Index).Visible = False
            '            Else
            '                dgvItems.Rows(row.Index).Selected = True
            '                'row.Cells(Columns.Qty_Ordered).Value = genQty
            '                row.Cells(Columns.Mix_Group).Value = determineNextGroupNumber()
            '                'TODO: fix trip validation bug (not working for all rows)
            '                updateTripValidation(row.Index, genQty)
            '            End If
            '        Next

            '        'TODO: set global inventory choice amount
            '    End If
            ' End If

            'Dim myConnection As SqlConnection = New SqlConnection("Server=thor;Initial Catalog=100;Integrated Security=SSPI")
            'Dim mySelectCommand As SqlCommand = New SqlCommand(strSql, myConnection)
            'Dim mySqlDataAdapter As SqlDataAdapter = New SqlDataAdapter(mySelectCommand)
            'Dim myDataSet As New DataSet()
            'mySqlDataAdapter.Fill(myDataSet)
            'dgvItems.DataSource = Nothing
            'dgvItems.DataSource = myDataSet.Tables(0)

            dgvItems.AllowUserToAddRows = False
            dgvItems.AllowUserToOrderColumns = False

            Dim oImage, oResizedImage As Image
            Dim data As Byte()
            Dim ms As MemoryStream
            Dim dblRatio As Double

            If dgvItems.Columns(Columns.Image).Width < 120 Then dgvItems.Columns(Columns.Image).Width = 120

            For Each myRow As DataGridViewRow In dgvItems.Rows

                Dim bmp As Bitmap

                If Not (myRow.Cells(Columns.Image).Value.Equals(System.DBNull.Value)) Then
                    data = myRow.Cells(Columns.Image).Value ' (byte[]) dt.Rows[0]["IMAGE"];
                    ms = New MemoryStream(data)
                    oImage = Image.FromStream(ms)

                    If myRow.Cells(Columns.OEI_W_Pixels).Value < 400 Then '"""""""""" <> 400

                        If myRow.Cells(Columns.OEI_W_Pixels).Value < myRow.Cells(Columns.OEI_H_Pixels).Value Then
                            If Not (myRow.Cells(Columns.OEI_Image_Rotate).Value = "N") Then
                                oImage.RotateFlip(RotateFlipType.Rotate270FlipX)
                            End If
                        End If

                        bmp = New Bitmap(120, 120)

                        Using g As Graphics = Graphics.FromImage(bmp)
                            g.DrawImage(oImage, 0, 0, bmp.Width, bmp.Height)
                        End Using
                        oResizedImage = bmp

                    Else

                        'dblRatio = oImage.Height / oImage.Width
                        dblRatio = myRow.Cells(Columns.OEI_H_Pixels).Value / myRow.Cells(Columns.OEI_W_Pixels).Value
                        If dblRatio > 2 Then
                            If Not (myRow.Cells(Columns.OEI_Image_Rotate).Value = "N") Then
                                oImage.RotateFlip(RotateFlipType.Rotate270FlipX)
                            End If
                            dblRatio = oImage.Width / oImage.Height
                        End If

                        If dblRatio >= 1 Then '"""""""""" > 1
                            If dblRatio > 5.5 Then dblRatio = 5.5
                            bmp = New Bitmap(120, CInt(120 / dblRatio))
                        Else
                            bmp = New Bitmap(CInt(120 * dblRatio), 120)
                        End If

                        Using g As Graphics = Graphics.FromImage(bmp)
                            g.DrawImage(oImage, 0, 0, bmp.Width, bmp.Height)
                        End Using
                        oResizedImage = bmp

                    End If

                    myRow.Cells(Columns.Image).Value = oResizedImage ' oImage

                    If myRow.Cells(Columns.OEI_W_Pixels).Value < 400 Then '"""""""""" <> 400
                        Dim dblHeight As Double
                        dblHeight = IIf(400 / myRow.Cells(Columns.OEI_W_Pixels).Value > 5.5, 5.5, 400 / myRow.Cells(Columns.OEI_W_Pixels).Value)
                        myRow.Height = 120 / dblHeight
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Height = 120 / dblHeight
                    Else
                        '    Call ResizeImage(dgvItems.CurrentRow.Index, g_oOrdline.Image)
                        myRow.Height = oResizedImage.Height ' IIf(oResizedImage.Height > 25, oResizedImage.Height + 5, 25)
                        dgvItems.Rows(myRow.Index).Height = oResizedImage.Height
                        'dgvItems.Columns(0).Width = IIf(dgvItems.Columns(0).Width > myRow.Cells(0).Size.Width, dgvItems.Columns(0).Width, myRow.Cells(0).Size.Width)
                    End If

                    'myRow.Height = oResizedImage.Height ' IIf(oResizedImage.Height > 25, oResizedImage.Height + 5, 25)
                    'dgvItems.Rows(myRow.Index).Height = oResizedImage.Height

                    dgvItems.Columns(0).Width = IIf(dgvItems.Columns(0).Width > myRow.Cells(0).Size.Width, dgvItems.Columns(0).Width, myRow.Cells(0).Size.Width)

                Else
                    bmp = New Bitmap(124, 24)
                    oResizedImage = bmp
                    myRow.Cells(Columns.Image).Value = oResizedImage ' oImage
                    dgvItems.Columns(0).Width = IIf(dgvItems.Columns(0).Width > myRow.Cells(0).Size.Width, dgvItems.Columns(0).Width, myRow.Cells(0).Size.Width)
                End If
                '++ID 2.08.2019 added MixAndMAtch
                If myRow.Cells(Columns.Kit_Feat_Fg).Value = "K" And myRow.Cells(Columns.IsMixMatch).Value <> 1 Then


                    strSql = "SELECT * FROM OEI_Kit_Item_Loc_Qty_View WHERE Kit_Item_No = '" & myRow.Cells(Columns.Item_No).Value & "' AND Loc = '" & Loc & "' "

                    '02.17.2023
                    If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                        strSql = "SELECT * FROM OEI_Kit_Item_Loc_Qty_View_US WHERE Kit_Item_No = '" & myRow.Cells(Columns.Item_No).Value & "' AND Loc = '" & Loc & "' "
                    End If

                    Dim dt As New DataTable

                    Dim dblInventory As Double = 99999999
                    Dim dblQty_Allocated As Double = 99999999
                    Dim dblQty_On_Hand As Double = 99999999

                    dt = db.DataTable(strSql)
                    If dt.Rows.Count <> 0 Then

                        For Each dtRow As DataRow In dt.Rows
                            Dim oPrice As cOEPriceFile = New cOEPriceFile(dtRow.Item("Item_No").ToString, m_oOrder.Ordhead.Curr_Cd)
                            Dim intInv_Buffer As Integer = IIf(dgvItems.CurrentRow.Cells(Columns.Location).Value = "1", oPrice.Get_Inventory_Buffer, 0)

                            If dblQty_On_Hand > (dtRow.Item("Qty_Available") - dtRow.Item("OEI_Qty_On_Ord") - intInv_Buffer) Then

                                dblInventory = dtRow.Item("Qty_On_Hand") - intInv_Buffer
                                dblQty_Allocated = (dtRow.Item("Qty_Allocated") + dtRow.Item("OEI_Qty_On_Ord"))
                                dblQty_On_Hand = (dtRow.Item("Qty_Available") - dtRow.Item("OEI_Qty_On_Ord") - intInv_Buffer)

                                'dblInventory = dtRow.Item("Qty_Available") - oPrice.Get_Inventory_Buffer
                                'dblQty_Allocated = (dtRow.Item("Qty_Allocated") + dtRow.Item("OEI_Qty_On_Ord"))
                                'dblQty_On_Hand = (dtRow.Item("Qty_On_Hand") - dtRow.Item("OEI_Qty_On_Ord") - oPrice.Get_Inventory_Buffer)

                            End If
                        Next
                    End If

                    myRow.Cells(Columns.Qty_Inventory).Value = dblInventory
                    myRow.Cells(Columns.Qty_Allocated).Value = dblQty_Allocated
                    myRow.Cells(Columns.Qty_On_Hand).Value = dblQty_On_Hand

                End If

                'If Mid(myRow.Cells(Columns.Item_No).Value, 1, 2) = "00" Then
                '    intCountItems += 1
                'Else
                '    intCountCharges += 1
                'End If

                If Mid(myRow.Cells(Columns.Item_No).Value, 1, 2) = "44" Or Mid(myRow.Cells(Columns.Item_No).Value, 1, 2) = "66" Or Mid(myRow.Cells(Columns.Item_No).Value, 1, 2) = "88" Then
                    intCountCharges += 1
                Else
                    intCountItems += 1
                End If

                Dim oItem As New cOrdLine()
                oItem.Set_Spec_Instructions(myRow.Cells(Columns.Item_No).Value, oItem.ImprintLine)

                If Not oItem.ImprintLine Is Nothing Then
                    myRow.Cells(Columns.Comments).Value = oItem.ImprintLine.Comments
                    myRow.Cells(Columns.Special_Comments).Value = oItem.ImprintLine.Special_Comments
                    myRow.Cells(Columns.Printer_Comment).Value = oItem.ImprintLine.Printer_Comment
                    myRow.Cells(Columns.Packer_Comment).Value = oItem.ImprintLine.Packer_Comment
                    myRow.Cells(Columns.Printer_Instructions).Value = oItem.ImprintLine.Printer_Instructions
                    myRow.Cells(Columns.Packer_Instructions).Value = oItem.ImprintLine.Packer_Instructions



                End If

            Next myRow

            'dgvItems.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.DisableResizing)
            dgvItems.RowHeadersVisible = False
            'dgvItems.RowHeadersWidth = 200

            If Item_No = "CHARGESCALCULATION" Then

                Call SetColumnsVisibleCharges(True)

            Else
                If intCountCharges > intCountItems Then

                    Call SetQtyColumnsVisible(False)

                End If
            End If

            'cboIndustry.DataSource = m_oOrder.GetCompleteCboIndustryList()

        Catch er As Exception
        End Try

    End Sub

    Private Function determineNextGroupNumber() As Integer
        Dim oLine As cOrdLine
        Dim currNumber As Integer = 0

        For Each oLine In m_oOrder.OrdLines
            If oLine.Mix_Group.ToString <> "" AndAlso oLine.Mix_Group > currNumber Then
                currNumber = oLine.Mix_Group
            End If
        Next

        Return currNumber + 1
    End Function

    Private Sub updateTripValidation(cCell As Integer, qty As Decimal)
        'enter qty ordered cell
        dgvItems.CurrentCell = dgvItems.Rows(cCell).Cells(Columns.Qty_Ordered)
        dgvItems.BeginEdit(True)
        dgvItems.EndEdit(True)

        dgvItems.Rows(cCell).Cells(Columns.Qty_Ordered).Value = qty

        'switch cells to force validation
        dgvItems.CurrentCell = dgvItems.Rows(cCell).Cells(Columns.Qty_To_Ship)
        dgvItems.BeginEdit(True)
        dgvItems.EndEdit(True)

        'enter qty ordered cell
        dgvItems.CurrentCell = dgvItems.Rows(cCell).Cells(Columns.Qty_Ordered)
        dgvItems.BeginEdit(True)
        dgvItems.EndEdit(True)
    End Sub

    Private Sub SetQtyColumnsVisible(ByVal pblnVisible As Boolean)

        ' Qty ordered & qty to ship must always be visible
        'dgvItems.Columns(Columns.Qty_Ordered).Visible = pblnVisible
        'dgvItems.Columns(Columns.Qty_To_Ship).Visible = pblnVisible

        dgvItems.Columns(Columns.Item_Desc_1_View).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Item_Desc_2_View).Visible = Not (pblnVisible)

        'dgvItems.Columns(Columns.Qty_Inventory).Visible = pblnVisible
        'dgvItems.Columns(Columns.Qty_Allocated).Visible = pblnVisible
        'dgvItems.Columns(Columns.Qty_On_Hand).Visible = pblnVisible

        'dgvItems.Columns(Columns.Qty_Prev_To_Ship).Visible = pblnVisible
        'dgvItems.Columns(Columns.Qty_Prev_BkOrd).Visible = pblnVisible

        'dgvItems.Columns(Columns.Extra_9).Visible = pblnVisible
        'dgvItems.Columns(Columns.Extra_10).Visible = pblnVisible

        dgvItems.Columns(Columns.Item_Desc_1).Visible = pblnVisible
        dgvItems.Columns(Columns.Item_Desc_2).Visible = pblnVisible

        dgvItems.Columns(Columns.Item_Desc_Charge).Visible = False

    End Sub

    Private Sub SetColumnsVisibleCharges(ByVal pblnVisible As Boolean)

        Debug.Print("")

        dgvItems.Columns(Columns.Image).Width = 0
        dgvItems.Columns(Columns.Item_Desc_1_View).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Item_Desc_2_View).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_To_Ship).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_Prev_To_Ship).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_Prev_BkOrd).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_Prev_BkOrd).Visible = Not (pblnVisible)

        dgvItems.Columns(Columns.Qty_Lines).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_Inventory).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_Allocated).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_On_Hand).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_Prev_To_Ship).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_Prev_BkOrd).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Qty_Prev_BkOrd).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Extra_9).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Extra_10).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Discount_Pct).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Request_Dt).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Promise_Dt).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Req_Ship_Dt).Visible = Not (pblnVisible)

        dgvItems.Columns(Columns.Extra_1).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Imprint).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.End_User).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Imprint_Location).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Imprint_Color).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Num_Impr_1).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Num_Impr_2).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Num_Impr_3).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Packaging).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Refill).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Laser_Setup).Visible = Not (pblnVisible)
        ''''dgvItems.Columns(Columns.Industry).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Comments).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Special_Comments).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Printer_Comment).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Packer_Comment).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Printer_Instructions).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Packer_Instructions).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Repeat_From_ID).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.RouteID).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Route).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.ProductProof).Visible = Not (pblnVisible)

        dgvItems.Columns(Columns.UOM).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Tax_Sched).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Revision_No).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.BkOrd_Fg).Visible = Not (pblnVisible)
        dgvItems.Columns(Columns.Recalc_SW).Visible = Not (pblnVisible)

    End Sub


#Region "Properties ##########################################"

#End Region

    Private Sub dgvItems_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItems.CellBeginEdit

        Select Case e.ColumnIndex

            Case Columns.FirstCol
                e.Cancel = True

            Case Columns.Line_Seq_No
                e.Cancel = True

            Case Columns.Image
                e.Cancel = True

            Case Columns.Item_No
                e.Cancel = True

            Case Columns.Location
                ' Let go

            Case Columns.Item_Desc_1_View
                ' Let go

            Case Columns.Item_Desc_2_View
                ' Let go

            Case Columns.Qty_Ordered
                ' Let go

            Case Columns.Qty_To_Ship
                ' Let go

            Case Columns.Qty_Lines
                If dgvItems.CurrentRow.Cells(Columns.Kit_Feat_Fg).Value.ToString = "K" Then
                    e.Cancel = True
                End If

            Case Columns.Qty_Inventory
                e.Cancel = True

            Case Columns.Qty_Allocated
                e.Cancel = True

            Case Columns.Qty_On_Hand
                e.Cancel = True

            Case Columns.Qty_Prev_To_Ship
                e.Cancel = True

            Case Columns.Qty_Prev_BkOrd
                e.Cancel = True

            Case Columns.Unit_Price
                ' Let go

            Case Columns.Calc_Price
                e.Cancel = True

            Case Columns.Request_Dt
                ' Let go

            Case Columns.Promise_Dt
                ' Let go

            Case Columns.Req_Ship_Dt
                ' Let go

            Case Columns.Item_Desc_1
                ' Let go

            Case Columns.Item_Desc_2
                ' Let go

            Case Columns.UOM
                e.Cancel = True

            Case Columns.Tax_Sched
                e.Cancel = True

            Case Columns.Revision_No
                e.Cancel = True

            Case Columns.BkOrd_Fg
                e.Cancel = True

            Case Columns.Recalc_SW
                e.Cancel = True

            Case Columns.RouteID
                e.Cancel = True

            Case Columns.Route
                ' Let go

            Case Columns.ProductProof
                If dgvItems.CurrentRow.Cells(Columns.Kit_Feat_Fg).Value.ToString = "K" Then
                    e.Cancel = True
                End If

            Case Columns.Imprint
                ' Let go

            Case Columns.End_User
                ' Let go

            Case Columns.Imprint_Location
                ' Let go

            Case Columns.Imprint_Color
                ' Let go

            Case Columns.Num_Impr_1
                ' Let go

            Case Columns.Num_Impr_2
                ' Let go

            Case Columns.Num_Impr_3
                ' Let go

            Case Columns.Packaging
                ' Let go

            Case Columns.Refill
                ' Let go

            Case Columns.Laser_Setup
                ' Let go

            Case Columns.Industry
                ' Let go

            Case Columns.Comments
                ' Let go

            Case Columns.Special_Comments
                ' Let go

            Case Columns.Printer_Comment
                ' Let Go

            Case Columns.Packer_Comment
                ' Let Go

            Case Columns.Printer_Instructions
                ' Let Go

            Case Columns.Packer_Instructions
                ' Let Go

            Case Columns.Repeat_From_ID
                ' Let Go

            Case Columns.Prod_Cat
                e.Cancel = True

            Case Columns.LastCol
                e.Cancel = True

            Case Columns.Lamination, Columns.Foil_Color, Columns.Thread_Color, Columns.Rounded_Corners, Columns.Tip_In_PageTxt
                ' Let Go
        End Select

        ' if qty ordered = 0 then we dont allow editing, must enter qty first.
        If dgvItems.CurrentCell.ColumnIndex > Columns.Qty_Ordered And RTrim(Trim(dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value)) = "0" Then
            e.Cancel = True
            Exit Sub
        End If

        If dgvItems.Rows(iRowPos).Cells(iColPos).Style.BackColor = Color.FromArgb(255, 192, 192) Then
            lInvalidate = ErrorType.NoError
            dgvItems.Rows(iRowPos).Cells(iColPos).Style.BackColor = Color.White
        End If

    End Sub


#Region "Route combo box #####################################"

    Private Sub AddComboBoxColumn(ByVal pstrColumnName As String, ByVal pintPos As Integer)

        Dim cboColumn As New DataGridViewComboBoxColumn

        With cboColumn
            .HeaderText = pstrColumnName
            .DataPropertyName = pstrColumnName
            .DropDownWidth = 200
            .Width = 200
            .MaxDropDownItems = 12
            .FlatStyle = FlatStyle.Standard
        End With

        dgvItems.Columns.Insert(pintPos, cboColumn)

    End Sub

    Private Sub SetRouteCombo(ByVal cboCell As DataGridViewComboBoxCell, ByVal pstrType As String, ByVal pstrValue As String)

        Try
            Dim item_no As String = "" 'justin add 20230301 for Vegas route
            item_no = LTrim(RTrim((dgvItems.Rows(cboCell.RowIndex).Cells(Columns.Item_No).Value.ToString)))

            With cboCell
                '.DataSource = RetrieveAlternativeTitles()

                '.ValueMember = "RouteId" ' ColumnName.TitleOfCourtesy.ToString()
                .ValueMember = "Route" ' ColumnName.TitleOfCourtesy.ToString()
                .DisplayMember = "Route"

                .DataSource = m_oOrder.ChangeCboRouteList(pstrType, dgvItems.Rows(cboCell.RowIndex).Cells(Columns.Prod_Cat).Value.ToString, pstrValue, item_no)


            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function PopulateComboBox(ByVal pstrSqlCommand As String) As DataTable

        Dim dt As New DataTable
        Dim db As New cDBA
        dt = db.DataTable(pstrSqlCommand)

        Return dt

        'Dim myConnection As SqlConnection = New SqlConnection("Server=thor;Initial Catalog=100;Integrated Security=SSPI")
        'Dim mySelectCommand As SqlCommand = New SqlCommand(pstrSqlCommand, myConnection)
        'Dim mySqlDataAdapter As SqlDataAdapter = New SqlDataAdapter(mySelectCommand)

        'Dim myDataSet As New DataSet()
        'mySqlDataAdapter.Fill(myDataSet)
        'Return myDataSet.Tables(0) ' dgvItems.DataSource = myDataSet.Tables(0)

    End Function

#End Region

    ' Returns the search button control associated with a textbox or a combobox
    Private Function GetNotesControlByColumnIndex() As Button

        GetNotesControlByColumnIndex = New Button

        Try
            Select Case dgvItems.CurrentCell.ColumnIndex

                Case Columns.Item_No, Columns.Location
                    GetNotesControlByColumnIndex = cmdNotes

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function


    ' Returns the search button control associated with a textbox or a combobox
    Private Function GetSearchControlByColumnIndex() As Button

        GetSearchControlByColumnIndex = New Button

        Try
            Select Case dgvItems.CurrentCell.ColumnIndex

                Case Columns.Item_No, Columns.Location
                    GetSearchControlByColumnIndex = cmdSearch

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' WILL NOT DO THIS BUTTON, ELSE WILL HAVE TO CODE EVERYTHING TO GET ORIGINAL DATA FIRST. TAG WILL NOT WORK HERE.
    Private Sub cmdResetAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResetAll.Click

        'If dgvItems.Rows.Count > 0 Then

        '    For Each oRow As DataGridViewRow In dgvItems.Rows
        '        oRow.Cells(Columns.Qty_To_Ship).Value = 0
        '    Next

        'End If

    End Sub

    Private Sub cmdApplyToAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApplyToAll.Click

        If dgvItems.CurrentRow.Cells(Columns.Qty_To_Ship).Style.BackColor = Color.FromArgb(255, 192, 192) Then
            Exit Sub
        End If

        Dim lRow As Long = dgvItems.CurrentCell.RowIndex
        Dim dblQtyOrder As Double = dgvItems.Rows(lRow).Cells(Columns.Qty_Ordered).Value
        Dim dblQtyOrdered As Double = dgvItems.Rows(lRow).Cells(Columns.Qty_To_Ship).Value
        Dim dblQtyLines As Double = dgvItems.Rows(lRow).Cells(Columns.Qty_Lines).Value
        Dim strItem As String = dgvItems.Rows(lRow).Cells(Columns.Item_No).Value

        If dgvItems.Rows.Count = 0 Then Exit Sub

        For Each oRow As DataGridViewRow In dgvItems.Rows

            If oRow.Cells(Columns.Item_No).Value <> strItem Then

                Dim oOrdline As cOrdLine = New cOrdLine(m_oOrder.Ordhead.Ord_GUID)
                SaveOrderLine(oRow, oOrdline)

                With dgvItems.Rows(dgvItems.CurrentRow.Index)

                    ' On copy, do not copy item prices, qty on hand, calc price, 
                    ' Also, calculate qty prev to ship, qty prev avail, qty prev BO
                    oOrdline.Qty_Ordered = dblQtyOrder
                    oRow.Cells(Columns.Qty_Ordered).Value = dblQtyOrder
                    oOrdline.Qty_To_Ship = dblQtyOrdered
                    oRow.Cells(Columns.Qty_To_Ship).Value = dblQtyOrdered
                    oOrdline.Qty_Lines = dblQtyLines
                    oRow.Cells(Columns.Qty_Lines).Value = dblQtyLines
                    'oOrdline.Qty_Lines = dblQtyLines
                    oRow.Cells(Columns.Qty_On_Hand).Value = oOrdline.Qty_On_Hand
                    oRow.Cells(Columns.Qty_Prev_To_Ship).Value = oOrdline.Qty_Prev_To_Ship
                    oRow.Cells(Columns.Qty_Prev_BkOrd).Value = oOrdline.Qty_Prev_Bkord
                    oOrdline.Unit_Price = .Cells(Columns.Unit_Price).Value
                    oRow.Cells(Columns.Unit_Price).Value = oOrdline.Unit_Price
                    oOrdline.Discount_Pct = .Cells(Columns.Discount_Pct).Value
                    oRow.Cells(Columns.Discount_Pct).Value = oOrdline.Discount_Pct
                    'oOrdline.Unit_Price
                    oRow.Cells(Columns.Calc_Price).Value = oOrdline.Calc_Price
                    oRow.Cells(Columns.Promise_Dt).Value = .Cells(Columns.Promise_Dt).Value
                    oRow.Cells(Columns.Request_Dt).Value = .Cells(Columns.Request_Dt).Value
                    oRow.Cells(Columns.Req_Ship_Dt).Value = .Cells(Columns.Req_Ship_Dt).Value

                    oRow.Cells(Columns.RouteID).Value = .Cells(Columns.RouteID).Value
                    oRow.Cells(Columns.Route).Value = .Cells(Columns.Route).Value
                    oRow.Cells(Columns.ProductProof).Value = .Cells(Columns.ProductProof).Value

                    oRow.Cells(Columns.Imprint).Value = .Cells(Columns.Imprint).Value
                    oRow.Cells(Columns.End_User).Value = .Cells(Columns.End_User).Value
                    oRow.Cells(Columns.Imprint_Location).Value = .Cells(Columns.Imprint_Location).Value
                    oRow.Cells(Columns.Imprint_Color).Value = .Cells(Columns.Imprint_Color).Value
                    oRow.Cells(Columns.Num_Impr_1).Value = .Cells(Columns.Num_Impr_1).Value
                    oRow.Cells(Columns.Num_Impr_2).Value = .Cells(Columns.Num_Impr_2).Value
                    oRow.Cells(Columns.Num_Impr_3).Value = .Cells(Columns.Num_Impr_3).Value
                    oRow.Cells(Columns.Packaging).Value = .Cells(Columns.Packaging).Value
                    oRow.Cells(Columns.Refill).Value = .Cells(Columns.Refill).Value
                    oRow.Cells(Columns.Laser_Setup).Value = .Cells(Columns.Laser_Setup).Value
                    oRow.Cells(Columns.Industry).Value = .Cells(Columns.Industry).Value
                    oRow.Cells(Columns.Comments).Value = .Cells(Columns.Comments).Value
                    oRow.Cells(Columns.Special_Comments).Value = .Cells(Columns.Special_Comments).Value
                    oRow.Cells(Columns.Printer_Comment).Value = .Cells(Columns.Printer_Comment).Value
                    oRow.Cells(Columns.Packer_Comment).Value = .Cells(Columns.Packer_Comment).Value
                    oRow.Cells(Columns.Printer_Instructions).Value = .Cells(Columns.Printer_Instructions).Value
                    oRow.Cells(Columns.Packer_Instructions).Value = .Cells(Columns.Packer_Instructions).Value
                    oRow.Cells(Columns.Repeat_From_ID).Value = .Cells(Columns.Repeat_From_ID).Value
                    oRow.Cells(Columns.Prod_Cat).Value = .Cells(Columns.Prod_Cat).Value

                    '06.08.2020 added for scrabl project -------------------
                    oRow.Cells(Columns.Lamination).Value = .Cells(Columns.Lamination).Value
                    '++ID 12.09.2021 Foil Color
                    oRow.Cells(Columns.Foil_Color).Value = .Cells(Columns.Foil_Color).Value

                    oRow.Cells(Columns.Thread_Color).Value = .Cells(Columns.Thread_Color).Value
                    oRow.Cells(Columns.Rounded_Corners).Value = .Cells(Columns.Rounded_Corners).Value
                    '-------------------------------------------------------
                End With

            End If

        Next

    End Sub

    Private Sub frmProductLineEntry_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        m_oOrder.Source = OEI_SourceEnum.frmOrder
        m_oOrder.Reload = True

    End Sub

    Private Sub frmProductLineEntry_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        m_oOrder.Source = OEI_SourceEnum.frmProductLineEntry

        ' LoadHumRes()
        LoadProductGrid()

        If dgvItems.Rows.Count > 0 Then dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Qty_Ordered)

        dgvItems.Focus()

    End Sub

    Private Sub frmProductLineEntry_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        '++ID 06.30.2020 commented

        'dgvItems.Height = Me.Height - 260
        'dgvItems.Width = Me.Width - 30

        'panImprint.Top = Me.Height - 240

        'cmdAddToOrder.Top = Me.Height - 90
        'cmdCancel.Top = cmdAddToOrder.Top
        'cmdApplyToAll.Top = cmdAddToOrder.Top
        'cmdResetAll.Top = cmdAddToOrder.Top

        'cmdAddToOrder.Left = Me.Width - 225
        'cmdCancel.Left = Me.Width - 120

    End Sub

    Private Sub cmdAddToOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddToOrder.Click
        frmOrder.ignoreImprintMsgCheck = True
        '_ProductStatus = Status.Insert

        m_Source = OEI_SourceEnum.frmPLEntrySelectedItem

        If Item_No = "CHARGESCALCULATION" Then
            Call SaveCharges()
        Else
            Call Save()
        End If

        ' _ProductStatus = Status.Insert
        'm_oOrder.Reload = True

        'reset message log
        cOrdLine.popMessageLog.Clear()

        Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        'reset message log
        cOrdLine.popMessageLog.Clear()
        '_ProductStatus = Status.Cancel
        'm_oOrder.Reload = True
        Close()

    End Sub

    ' Will not be used here
    Public Sub LoadOrder(ByRef pOrder As cOrder)

        m_oOrder = pOrder

    End Sub

    ' Will not be used here
    Public Sub NewOrder(ByRef pOrder As cOrder)

        m_oOrder = pOrder

    End Sub

    Public Sub Save()

        'dgvItems.AllowUserToAddRows = True

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            Dim intLinesCount As Integer = m_oOrder.OrdLines.Count

            m_oOrder.NewRowsToAdd = 0

            For iSeq As Integer = 1 To m_SelectSeq

                For Each oRow As DataGridViewRow In dgvItems.Rows

                    'Skip row if item is set to hold
                    If oRow.Cells(Columns.Loc_Status).Value = "H" Then Continue For

                    If oRow.Cells(Columns.SelectSeq).Value = iSeq Then

                        If oRow.Cells(Columns.Qty_Ordered).Value > 0 Then

                            If oRow.Cells(Columns.Qty_Lines).Value > 1 Then

                                Dim dblQtyShipped As Double = 0
                                For lPos As Integer = 1 To oRow.Cells(Columns.Qty_Lines).Value

                                    'Dim olItem As New cOrdLine(m_oOrder.Ordhead.Ord_GUID) ' m_oOrder)
                                    Dim olItem As New cOrdLine(m_oOrder.Ordhead.Ord_GUID, OEI_SourceEnum.frmPLEntrySelectedItem) ' m_oOrder)
                                    'm_oOrder.Source = OEI_SourceEnum.frmProductLineEntry

                                    SaveOrderLine(oRow, olItem)
                                    If Not olItem.ImprintLine.IsEmpty() Then olItem.ImprintLine.Save()

                                    olItem.Qty_Prev_To_Ship = 0
                                    olItem.Qty_Prev_Bkord = 0

                                    If (dblQtyShipped + olItem.Qty_To_Ship > olItem.Qty_On_Hand) Then
                                        If (olItem.Qty_On_Hand > 0) Then

                                            olItem.m_Qty_Prev_Bkord = dblQtyShipped + olItem.Qty_To_Ship - olItem.Qty_On_Hand

                                            If olItem.m_Qty_Prev_Bkord > olItem.Qty_Ordered Then
                                                olItem.m_Qty_Prev_Bkord = olItem.Qty_Ordered
                                            End If

                                            olItem.m_Qty_Prev_To_Ship = olItem.Qty_Ordered - olItem.Qty_Prev_Bkord
                                            olItem.m_Qty_To_Ship = olItem.m_Qty_Prev_To_Ship

                                        Else
                                            olItem.m_Qty_Prev_Bkord = olItem.Qty_Ordered
                                            olItem.m_Qty_Prev_To_Ship = 0
                                            olItem.m_Qty_To_Ship = 0

                                        End If

                                        dblQtyShipped += olItem.Qty_Ordered
                                        'olItem.Qty_Prev_To_Ship = olItem.Qty_On_Hand - dblQtyShipped
                                    Else
                                        olItem.Qty_Prev_Bkord = 0
                                        olItem.Qty_Prev_To_Ship = olItem.Qty_To_Ship
                                        dblQtyShipped += olItem.Qty_To_Ship
                                    End If

                                    'olItem.SaveToDB = True

                                    'olItem.Save()

                                    'olItem.SaveToDB = False

                                    'm_oOrder.AddOrderLine(olItem)

                                    olItem.Source = OEI_SourceEnum.frmPLEntrySelectedItem
                                    m_oOrder.AddExternalLine(olItem)

                                Next lPos

                            Else
                                Dim olItem As New cOrdLine(m_oOrder.Ordhead.Ord_GUID, OEI_SourceEnum.frmPLEntrySelectedItem) ' m_oOrder)
                                'm_oOrder.Source = OEI_SourceEnum.frmProductLineEntry

                                SaveOrderLine(oRow, olItem)
                                If Not olItem.ImprintLine.IsEmpty() Then olItem.ImprintLine.Save()

                                'olItem.SaveToDB = True

                                'olItem.Save()

                                'olItem.SaveToDB = False

                                'm_oOrder.AddOrderLine(olItem)

                                olItem.Source = OEI_SourceEnum.frmOrder
                                m_oOrder.AddExternalLine(olItem)

                            End If

                        End If

                    End If

                Next oRow

            Next iSeq

            m_oOrder.NewRowsToAdd = m_oOrder.OrdLines.Count - intLinesCount

            'dgvItems.AllowUserToAddRows = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SaveCharges()

        'dgvItems.AllowUserToAddRows = True

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            Dim intLinesCount As Integer = m_oOrder.OrdLines.Count

            m_oOrder.NewRowsToAdd = 0

            For Each oRow As DataGridViewRow In dgvItems.Rows

                If oRow.Cells(Columns.Qty_Ordered).Value > 0 Then

                    If oRow.Cells(Columns.Qty_Lines).Value > 1 Then

                        Dim dblQtyShipped As Double = 0
                        For lPos As Integer = 1 To oRow.Cells(Columns.Qty_Lines).Value

                            Dim olItem As New cOrdLine(m_oOrder.Ordhead.Ord_GUID, OEI_SourceEnum.frmPLEntrySelectedItem) ' m_oOrder)
                            'm_oOrder.Source = OEI_SourceEnum.frmProductLineEntry

                            SaveOrderLine(oRow, olItem)

                            olItem.Qty_Prev_To_Ship = 0
                            olItem.Qty_Prev_Bkord = 0

                            If (dblQtyShipped + olItem.Qty_To_Ship > olItem.Qty_On_Hand) Then
                                If (olItem.Qty_On_Hand > 0) Then

                                    olItem.m_Qty_Prev_Bkord = dblQtyShipped + olItem.Qty_To_Ship - olItem.Qty_On_Hand

                                    If olItem.m_Qty_Prev_Bkord > olItem.Qty_Ordered Then
                                        olItem.m_Qty_Prev_Bkord = olItem.Qty_Ordered
                                    End If

                                    olItem.m_Qty_Prev_To_Ship = olItem.Qty_Ordered - olItem.Qty_Prev_Bkord
                                    olItem.m_Qty_To_Ship = olItem.m_Qty_Prev_To_Ship

                                Else
                                    olItem.m_Qty_Prev_Bkord = olItem.Qty_Ordered
                                    olItem.m_Qty_Prev_To_Ship = 0
                                    olItem.m_Qty_To_Ship = 0

                                End If

                                dblQtyShipped += olItem.Qty_Ordered
                                'olItem.Qty_Prev_To_Ship = olItem.Qty_On_Hand - dblQtyShipped
                            Else
                                olItem.Qty_Prev_Bkord = 0
                                olItem.Qty_Prev_To_Ship = olItem.Qty_To_Ship
                                dblQtyShipped += olItem.Qty_To_Ship
                            End If

                            'olItem.SaveToDB = True

                            'olItem.Save()

                            'olItem.SaveToDB = False

                            'm_oOrder.AddOrderLine(olItem)

                            olItem.Source = OEI_SourceEnum.frmPLEntrySelectedItem
                            m_oOrder.AddExternalLine(olItem)

                        Next lPos

                    Else
                        Dim olItem As New cOrdLine(m_oOrder.Ordhead.Ord_GUID, OEI_SourceEnum.frmPLEntrySelectedItem) ' m_oOrder)
                        'm_oOrder.Source = OEI_SourceEnum.frmProductLineEntry

                        SaveOrderLine(oRow, olItem)

                        'olItem.SaveToDB = True

                        'olItem.Save()

                        'olItem.SaveToDB = False

                        'm_oOrder.AddOrderLine(olItem)

                        olItem.Source = OEI_SourceEnum.frmPLEntrySelectedItem
                        m_oOrder.AddExternalLine(olItem)

                    End If

                End If

            Next oRow

            m_oOrder.NewRowsToAdd = m_oOrder.OrdLines.Count - intLinesCount

            'dgvItems.AllowUserToAddRows = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '' Call this instance when you already have the Row in a DataGridViewRow
    'Private Function SaveOrderLine(ByRef oRow As DataGridViewRow) As cOrdLine

    '    Dim olItem As New cOrdLine(False) ' m_oOrder, True)
    '    olItem.Source = m_Source ' OEI_SourceEnum.frmProductLineEntry

    '    Try
    '        olItem.Loading = True

    '        olItem.m_Item_No = oRow.Cells(Columns.Item_No).Value
    '        olItem.m_Loc = oRow.Cells(Columns.Location).Value
    '        olItem.m_Qty_Ordered = oRow.Cells(Columns.Qty_Ordered).Value
    '        olItem.m_Qty_To_Ship = oRow.Cells(Columns.Qty_To_Ship).Value
    '        olItem.m_Qty_Lines = oRow.Cells(Columns.Qty_Lines).Value
    '        olItem.m_Qty_Inventory = oRow.Cells(Columns.Qty_Inventory).Value
    '        olItem.m_Qty_Allocated = oRow.Cells(Columns.Qty_Allocated).Value
    '        olItem.m_Qty_Prev_To_Ship = oRow.Cells(Columns.Qty_Prev_To_Ship).Value
    '        olItem.m_Qty_Prev_Bkord = oRow.Cells(Columns.Qty_Prev_BkOrd).Value
    '        olItem.m_Qty_On_Hand = oRow.Cells(Columns.Qty_On_Hand).Value
    '        olItem.m_Unit_Price = oRow.Cells(Columns.Unit_Price).Value
    '        olItem.m_Extra_9 = oRow.Cells(Columns.Extra_9).Value
    '        olItem.m_Extra_10 = oRow.Cells(Columns.Extra_10).Value
    '        olItem.m_Discount_Pct = oRow.Cells(Columns.Discount_Pct).Value
    '        olItem.m_Calc_Price = oRow.Cells(Columns.Calc_Price).Value
    '        olItem.m_Request_Dt = oRow.Cells(Columns.Request_Dt).Value
    '        olItem.m_Promise_Dt = oRow.Cells(Columns.Promise_Dt).Value
    '        olItem.m_Req_Ship_Dt = oRow.Cells(Columns.Req_Ship_Dt).Value
    '        olItem.m_Item_Desc_1 = oRow.Cells(Columns.Item_Desc_1).Value
    '        olItem.m_Item_Desc_2 = oRow.Cells(Columns.Item_Desc_2).Value
    '        olItem.m_Uom = oRow.Cells(Columns.UOM).Value
    '        olItem.m_Tax_Sched = oRow.Cells(Columns.Tax_Sched).Value
    '        olItem.m_Revision_No = oRow.Cells(Columns.Revision_No).Value
    '        olItem.m_Bkord_Fg = IIf(oRow.Cells(Columns.BkOrd_Fg).Value, "Y", "N")
    '        olItem.m_Recalc_Sw = IIf(oRow.Cells(Columns.Recalc_SW).Value, "Y", "N")
    '        olItem.m_End_Item_Cd = oRow.Cells(Columns.Kit_Feat_Fg).Value

    '        olItem.m_Route = oRow.Cells(Columns.Route).Value
    '        olItem.m_RouteID = oRow.Cells(Columns.RouteID).Value
    '        olItem.ProductProof_Bln = oRow.Cells(Columns.ProductProof).Value

    '        olItem.m_Prod_Cat = oRow.Cells(Columns.Prod_Cat).Value
    '        olItem.m_Stocked_Fg = oRow.Cells(Columns.Stocked_Fg).Value
    '        olItem.m_Item_Cd = Mid(oRow.Cells(Columns.Item_Cd).Value, 1, 10)

    '        ' add imprint columns
    '        olItem.Create_Imprint = True

    '        olItem.ImprintLine = New cImprint(olItem.Item_Guid)

    '        olItem.ImprintLine.Imprint = oRow.Cells(Columns.Imprint).Value
    '        olItem.ImprintLine.End_User = oRow.Cells(Columns.End_User).Value
    '        olItem.ImprintLine.Location = oRow.Cells(Columns.Imprint_Location).Value
    '        olItem.ImprintLine.Color = oRow.Cells(Columns.Imprint_Color).Value
    '        If IsNumeric(oRow.Cells(Columns.Num_Impr_1).Value) Then olItem.ImprintLine.Num_Impr_1 = oRow.Cells(Columns.Num_Impr_1).Value
    '        If IsNumeric(oRow.Cells(Columns.Num_Impr_2).Value) Then olItem.ImprintLine.Num_Impr_2 = oRow.Cells(Columns.Num_Impr_2).Value
    '        If IsNumeric(oRow.Cells(Columns.Num_Impr_3).Value) Then olItem.ImprintLine.Num_Impr_3 = oRow.Cells(Columns.Num_Impr_3).Value
    '        olItem.ImprintLine.Packaging = oRow.Cells(Columns.Packaging).Value
    '        olItem.ImprintLine.Refill = oRow.Cells(Columns.Refill).Value
    '        olItem.ImprintLine.Laser_Setup = oRow.Cells(Columns.Laser_Setup).Value
    '        olItem.ImprintLine.Industry = oRow.Cells(Columns.Industry).Value.ToString
    '        olItem.ImprintLine.Comments = oRow.Cells(Columns.Comments).Value
    '        olItem.ImprintLine.Special_Comments = oRow.Cells(Columns.Special_Comments).Value
    '        olItem.ImprintLine.Printer_Comment = oRow.Cells(Columns.Printer_Comment).Value
    '        olItem.ImprintLine.Packer_Comment = oRow.Cells(Columns.Packer_Comment).Value
    '        olItem.ImprintLine.Printer_Instructions = oRow.Cells(Columns.Printer_Instructions).Value
    '        olItem.ImprintLine.Packer_Instructions = oRow.Cells(Columns.Packer_Instructions).Value
    '        olItem.ImprintLine.Repeat_From_ID = oRow.Cells(Columns.Repeat_From_ID).Value

    '        olItem.m_OEI_W_Pixels = oRow.Cells(Columns.OEI_W_Pixels).Value
    '        olItem.m_OEI_H_Pixels = oRow.Cells(Columns.OEI_H_Pixels).Value

    '        olItem.Create_Traveler = True
    '        olItem.Create_Kit = True

    '        olItem.Loading = False

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    '    SaveOrderLine = olItem

    'End Function

    Private Sub SaveOrderLine(ByRef oRow As DataGridViewRow, ByRef oOrdline As cOrdLine)

        'Dim olItem As New cOrdLine(False) ' m_oOrder, True)

        Try
10:         With oOrdline
20:             .SaveToDB = False
30:             .Source = m_Source ' OEI_SourceEnum.frmProductLineEntry

40:             .Loading = True

50:             .m_Item_No = oRow.Cells(Columns.Item_No).Value
60:             .m_Loc = oRow.Cells(Columns.Location).Value
70:             .m_Qty_Ordered = oRow.Cells(Columns.Qty_Ordered).Value
80:             .m_Qty_To_Ship = oRow.Cells(Columns.Qty_To_Ship).Value
90:             .m_Qty_Lines = oRow.Cells(Columns.Qty_Lines).Value
100:            .m_Qty_Inventory = oRow.Cells(Columns.Qty_Inventory).Value
110:            .m_Qty_Allocated = oRow.Cells(Columns.Qty_Allocated).Value
120:            .m_Qty_Prev_To_Ship = oRow.Cells(Columns.Qty_Prev_To_Ship).Value
130:            .m_Qty_Prev_Bkord = oRow.Cells(Columns.Qty_Prev_BkOrd).Value
140:            .m_Qty_On_Hand = oRow.Cells(Columns.Qty_On_Hand).Value
150:            .m_Unit_Price = oRow.Cells(Columns.Unit_Price).Value

                'Added June 22, 2017 - T. Louzon
                .m_Unit_Price_BeforeSpecial = oRow.Cells(Columns.Unit_Price).Value

160:            .m_Extra_9 = oRow.Cells(Columns.Extra_9).Value
170:            .m_Extra_10 = oRow.Cells(Columns.Extra_10).Value
180:            .m_Discount_Pct = oRow.Cells(Columns.Discount_Pct).Value
190:            .m_Calc_Price = oRow.Cells(Columns.Calc_Price).Value
200:            .m_Request_Dt = oRow.Cells(Columns.Request_Dt).Value
210:            .m_Promise_Dt = oRow.Cells(Columns.Promise_Dt).Value
220:            .m_Req_Ship_Dt = oRow.Cells(Columns.Req_Ship_Dt).Value
230:            .m_Item_Desc_1 = oRow.Cells(Columns.Item_Desc_1).Value
240:            .m_Item_Desc_2 = oRow.Cells(Columns.Item_Desc_2).Value
250:            .m_Uom = oRow.Cells(Columns.UOM).Value
260:            .m_Tax_Sched = oRow.Cells(Columns.Tax_Sched).Value
270:            .m_Revision_No = oRow.Cells(Columns.Revision_No).Value
280:            .m_Bkord_Fg = IIf(oRow.Cells(Columns.BkOrd_Fg).Value, "Y", "N")
290:            .m_Recalc_Sw = IIf(oRow.Cells(Columns.Recalc_SW).Value, "Y", "N")
300:            .m_End_Item_Cd = oRow.Cells(Columns.Kit_Feat_Fg).Value
301:            .m_Item_Cd = oRow.Cells(Columns.Item_Cd).Value
302:            .m_Loc_Status = oRow.Cells(Columns.Loc_Status).Value
303:            .m_isMixMatch = oRow.Cells(Columns.IsMixMatch).Value
304:            .m_Mix_Group = oRow.Cells(Columns.Mix_Group).Value

310:            .m_Route = oRow.Cells(Columns.Route).Value
320:            .m_RouteID = oRow.Cells(Columns.RouteID).Value

                If (m_oOrder.Ordhead.Mfg_Loc.Trim = "3" Or m_oOrder.Ordhead.Mfg_Loc.Trim = "3N") And .m_RouteID = 0 Then
                    If .m_Item_No.ToUpper.Trim = "88RANDOMSAMP" Then
                        .m_Route = "Completed"
                        .m_RouteID = 39
                    ElseIf .m_Item_No.ToUpper.Trim = "88SPECSAMPL" Then
                        .m_Route = "" ' "Completed"
                        .m_RouteID = 0 ' 39
                    Else
                        .m_Route = "Random Samples"
                        .m_RouteID = 16
                    End If
                End If

330:            .ProductProof_Bln = oRow.Cells(Columns.ProductProof).Value

340:            .m_Prod_Cat = oRow.Cells(Columns.Prod_Cat).Value
350:            .m_Stocked_Fg = oRow.Cells(Columns.Stocked_Fg).Value
360:            .m_Item_Cd = Mid(oRow.Cells(Columns.Item_Cd).Value, 1, 10)
365:            .m_Mix_Group = oRow.Cells(Columns.Mix_Group).Value

                ' add imprint columns
370:            .Create_Imprint = True

380:            .ImprintLine = New cImprint(.Item_Guid)

                .ImprintLine.Imprint = oRow.Cells(Columns.Imprint).Value
                .ImprintLine.End_User = oRow.Cells(Columns.End_User).Value
                .ImprintLine.Location = oRow.Cells(Columns.Imprint_Location).Value
                .ImprintLine.Color = oRow.Cells(Columns.Imprint_Color).Value
                If IsNumeric(oRow.Cells(Columns.Num_Impr_1).Value) Then .ImprintLine.Num_Impr_1 = oRow.Cells(Columns.Num_Impr_1).Value
                If IsNumeric(oRow.Cells(Columns.Num_Impr_2).Value) Then .ImprintLine.Num_Impr_2 = oRow.Cells(Columns.Num_Impr_2).Value
                If IsNumeric(oRow.Cells(Columns.Num_Impr_3).Value) Then .ImprintLine.Num_Impr_3 = oRow.Cells(Columns.Num_Impr_3).Value
                .ImprintLine.Packaging = oRow.Cells(Columns.Packaging).Value
                .ImprintLine.Refill = oRow.Cells(Columns.Refill).Value
                .ImprintLine.Laser_Setup = oRow.Cells(Columns.Laser_Setup).Value
                .ImprintLine.Industry = oRow.Cells(Columns.Industry).Value.ToString
                .ImprintLine.Comments = oRow.Cells(Columns.Comments).Value
                .ImprintLine.Special_Comments = oRow.Cells(Columns.Special_Comments).Value
                .ImprintLine.Printer_Comment = oRow.Cells(Columns.Printer_Comment).Value
                .ImprintLine.Packer_Comment = oRow.Cells(Columns.Packer_Comment).Value
                .ImprintLine.Printer_Instructions = oRow.Cells(Columns.Printer_Instructions).Value
                .ImprintLine.Packer_Instructions = oRow.Cells(Columns.Packer_Instructions).Value
                .ImprintLine.Repeat_From_ID = oRow.Cells(Columns.Repeat_From_ID).Value

                '++ID 06.08.2020   added for scrabl project -----
                .ImprintLine.Lamination_Scribl = oRow.Cells(Columns.Lamination).Value
                '++ID 12.09.2021 Foil Color
                .ImprintLine.Foil_Color = oRow.Cells(Columns.Foil_Color).Value

                .ImprintLine.Thread_Color_Scribl = oRow.Cells(Columns.Thread_Color).Value
                .ImprintLine.Rounded_corners_Scribl = oRow.Cells(Columns.Rounded_Corners).Value
                '++ID 08.06.2020
                .ImprintLine.Tip_In_PageTxt = oRow.Cells(Columns.Tip_In_PageTxt).Value
                '------------------------------------------------

                '================================================================
                'All Star Program - December 6th 2017 - Clayton Solomon
                '================================================================

                '++ID 5.3.2018 return Discount
                Dim m_OrderType As String = ""
                Dim m_byRegion As String = ""
                Dim m_star As String = ""
                Dim m_Mfg_Loc As String = ""

                If String.IsNullOrEmpty(m_oOrder.Ordhead.User_Def_Fld_2) <> True Then m_OrderType = m_oOrder.Ordhead.User_Def_Fld_2

                If String.IsNullOrEmpty(m_oOrder.Ordhead.Customer.region) <> True Then m_byRegion = m_oOrder.Ordhead.Customer.region

                If String.IsNullOrEmpty(m_oOrder.Ordhead.Customer.ClassificationId) <> True Then m_star = m_oOrder.Ordhead.Customer.ClassificationId.ToUpper

                If String.IsNullOrEmpty(m_oOrder.Ordhead.Mfg_Loc) <> True Then m_Mfg_Loc = m_oOrder.Ordhead.Mfg_Loc

                Dim m_Calc_Price As Double = 0

                If String.IsNullOrEmpty(oRow.Cells(Columns.Calc_Price).Value.ToString) <> True Then m_Calc_Price = oRow.Cells(Columns.Calc_Price).Value

                Dim m_ItemNo As String = ""
                Dim m_Prod_Cat As String = ""

                If String.IsNullOrEmpty(oRow.Cells(Columns.Item_No).Value.ToString) <> True Then
                    m_ItemNo = oRow.Cells(Columns.Item_No).Value.ToString

                    If String.IsNullOrEmpty(Return_Prod_Cat(m_ItemNo)) <> True Then m_Prod_Cat = Return_Prod_Cat(m_ItemNo) ' come from insertt line in the function

                End If

                .m_Discount_Pct = Ret_Discount(m_OrderType, m_byRegion, m_star, m_ItemNo, m_Prod_Cat, m_Calc_Price, m_Mfg_Loc)
                '---------------------------- NEW -------------------------------

                'Select Case m_AllStarProgram.ClassificationId ' capture star level

                '    Case "1ST", "2ST", "3ST"
                '        If m_AllStarProgram.m_Region.Contains("USA") Then

                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "SS" Then
                '                'Pricing: EQP-25% (set-up, item cost and all other charges)
                '                .m_Discount_Pct = _25percent
                '            End If

                '            'omit RUNNING charges for SELF PROMO 
                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "SP" And Not (.m_Prod_Cat.ToUpper.Trim = "PRC" And .m_Calc_Price < 5) Then
                '                .m_Discount_Pct = _25percent
                '            End If

                '            'ASH and Tek RS greater than 20$ EQP-25% else just EQP
                '            'Everything else is just EQP
                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "RS" Then
                '                If m_oOrder.Ordhead.Mfg_Loc.Trim = "3" And (.m_Prod_Cat = "TEK" Or .m_Prod_Cat = "BAG") Then

                '                    If oRow.Cells(Columns.Calc_Price).Value > 20 Then
                '                        .m_Discount_Pct = _25percent
                '                    End If
                '                End If
                '            End If

                '        Else
                '            ' CA customer
                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "SP" Or
                '                m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "SB" Then ' SB is ss for canadian orders

                '                'If .m_Prod_Cat.ToUpper.Trim = "TEK" or .m_Prod_Cat.ToUpper.Trim = "BAG" or (.m_Item_Desc_1.ToUpper.Contains("DEBOSS") Then
                '                '    .m_Discount_Pct = _25percent
                '                'End If
                '                .m_Discount_Pct = _25percent
                '            End If

                '            'ASH and Tek RS greater than 20$ EQP-50% else just EQP
                '            'Everything else is just EQP
                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "RS" Then
                '                If m_oOrder.Ordhead.Mfg_Loc.Trim = "3" And (.m_Prod_Cat = "TEK" Or .m_Prod_Cat = "BAG") Then
                '                    If oRow.Cells(Columns.Calc_Price).Value > 20 And Not m_AllStarProgram.ClassificationId = "3ST" Then
                '                        .m_Discount_Pct = _50percent
                '                    End If
                '                End If
                '            End If
                '        End If

                '    Case "4ST", "5ST", "6ST", "7ST"

                '        If m_AllStarProgram.m_Region.Contains("USA") Then

                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "SS" Then
                '                'Pricing: EQP-50% (set-up, item cost and all other charges)
                '                .m_Discount_Pct = _50percent

                '            End If

                '            'omit RUNNING charges for SELF PROMO 
                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "SP" And Not (.m_Prod_Cat.ToUpper.Trim = "PRC" And .m_Calc_Price < 5) Then
                '                .m_Discount_Pct = _50percent
                '            End If

                '            'ASH and Tek RS greater than 20$ EQP-25% or EQP-50% else just EQP
                '            'Everything else is just EQP
                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "RS" Then
                '                If m_oOrder.Ordhead.Mfg_Loc.Trim = "3" And (.m_Prod_Cat = "TEK" Or .m_Prod_Cat = "BAG") Then
                '                    If oRow.Cells(Columns.Calc_Price).Value > 20 Then
                '                        .m_Discount_Pct = _50percent
                '                    End If
                '                End If
                '            End If

                '        Else
                '            ' CA customer
                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "SP" Then
                '                .m_Discount_Pct = _50percent
                '            ElseIf m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "SB" Then 'SB is SS for canaadian orders
                '                .m_Discount_Pct = _25percent
                '                'If .m_Prod_Cat.ToUpper.Trim = "TEK" or .m_Prod_Cat.ToUpper.Trim = "BAG" or (.m_Item_Desc_1.ToUpper.Contains("DEBOSS") Then
                '                '    .m_Discount_Pct = _25percent
                '                'End If
                '            End If

                '            'ASH and Tek RS greater than 20$ EQP-50% else just EQP
                '            'Everything else is just EQP
                '            If m_oOrder.Ordhead.User_Def_Fld_2.ToUpper.Trim = "RS" Then
                '                If m_oOrder.Ordhead.Mfg_Loc.Trim = "3" And (.m_Prod_Cat = "TEK" Or .m_Prod_Cat = "BAG") Then
                '                    If oRow.Cells(Columns.Calc_Price).Value > 20 Then
                '                        .m_Discount_Pct = _50percent ' 
                '                    End If

                '                End If
                '            End If
                '        End If

                'End Select

                'If Not .ImprintLine.IsEmpty() Then
                '    .ImprintLine.Save()
                'End If

                .m_OEI_W_Pixels = oRow.Cells(Columns.OEI_W_Pixels).Value
                .m_OEI_H_Pixels = oRow.Cells(Columns.OEI_H_Pixels).Value

                .Create_Traveler = True
                .Create_Kit = True

                .Loading = False

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & " Line: " & Err.Erl & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Function SaveMultiOrderLine(ByRef oRow As DataGridViewRow, ByVal pPos As Integer) As cOrdLine

    '    Dim olItem As New cOrdLine(False) ' m_oOrder, True)

    '    Try
    '        olItem.Loading = True

    '        olItem.m_Item_No = oRow.Cells(Columns.Item_No).Value
    '        olItem.m_Loc = oRow.Cells(Columns.Location).Value
    '        olItem.m_Qty_Ordered = oRow.Cells(Columns.Qty_Ordered).Value
    '        olItem.m_Qty_To_Ship = oRow.Cells(Columns.Qty_To_Ship).Value
    '        olItem.m_Qty_Inventory = oRow.Cells(Columns.Qty_Inventory).Value
    '        olItem.m_Qty_Allocated = oRow.Cells(Columns.Qty_Allocated).Value
    '        olItem.m_Qty_Prev_To_Ship = oRow.Cells(Columns.Qty_Prev_To_Ship).Value
    '        olItem.m_Qty_Prev_Bkord = oRow.Cells(Columns.Qty_Prev_BkOrd).Value
    '        olItem.m_Qty_On_Hand = oRow.Cells(Columns.Qty_On_Hand).Value
    '        olItem.m_Unit_Price = oRow.Cells(Columns.Unit_Price).Value
    '        olItem.m_Extra_9 = oRow.Cells(Columns.Extra_9).Value
    '        olItem.m_Extra_10 = oRow.Cells(Columns.Extra_10).Value
    '        olItem.m_Discount_Pct = oRow.Cells(Columns.Discount_Pct).Value
    '        olItem.m_Calc_Price = oRow.Cells(Columns.Calc_Price).Value
    '        olItem.m_Request_Dt = oRow.Cells(Columns.Request_Dt).Value
    '        olItem.m_Promise_Dt = oRow.Cells(Columns.Promise_Dt).Value
    '        olItem.m_Req_Ship_Dt = oRow.Cells(Columns.Req_Ship_Dt).Value
    '        olItem.m_Item_Desc_1 = oRow.Cells(Columns.Item_Desc_1).Value
    '        olItem.m_Item_Desc_2 = oRow.Cells(Columns.Item_Desc_2).Value
    '        olItem.m_Uom = oRow.Cells(Columns.UOM).Value
    '        olItem.m_Tax_Sched = oRow.Cells(Columns.Tax_Sched).Value
    '        olItem.m_Revision_No = oRow.Cells(Columns.Revision_No).Value
    '        olItem.m_Bkord_Fg = IIf(oRow.Cells(Columns.BkOrd_Fg).Value, "Y", "N")
    '        olItem.m_Recalc_Sw = IIf(oRow.Cells(Columns.Recalc_SW).Value, "Y", "N")
    '        olItem.m_End_Item_Cd = oRow.Cells(Columns.Kit_Feat_Fg).Value

    '        olItem.m_Route = oRow.Cells(Columns.Route).Value
    '        olItem.m_RouteID = oRow.Cells(Columns.RouteID).Value
    '        olItem.ProductProof_Bln = oRow.Cells(Columns.ProductProof).Value

    '        olItem.m_Prod_Cat = oRow.Cells(Columns.Prod_Cat).Value
    '        olItem.m_Stocked_Fg = oRow.Cells(Columns.Stocked_Fg).Value

    '        ' add imprint columns
    '        olItem.Create_Imprint = True

    '        olItem.ImprintLine = New cImprint(olItem.Item_Guid)

    '        olItem.ImprintLine.Imprint = oRow.Cells(Columns.Imprint).Value
    '        olItem.ImprintLine.End_User = oRow.Cells(Columns.End_User).Value
    '        olItem.ImprintLine.Location = oRow.Cells(Columns.Imprint_Location).Value
    '        olItem.ImprintLine.Color = oRow.Cells(Columns.Imprint_Color).Value
    '        If IsNumeric(oRow.Cells(Columns.Num_Imprints).Value) Then olItem.ImprintLine.Num = oRow.Cells(Columns.Num_Imprints).Value
    '        olItem.ImprintLine.Packaging = oRow.Cells(Columns.Packaging).Value
    '        olItem.ImprintLine.Refill = oRow.Cells(Columns.Refill).Value
    '        olItem.ImprintLine.Laser_Setup = oRow.Cells(Columns.Laser_Setup).Value
    '        olItem.ImprintLine.Industry = oRow.Cells(Columns.Industry).Value.ToString
    '        olItem.ImprintLine.Comments = oRow.Cells(Columns.Comments).Value
    '        olItem.ImprintLine.Special_Comments = oRow.Cells(Columns.Special_Comments).Value

    '        olItem.m_OEI_W_Pixels = oRow.Cells(Columns.OEI_W_Pixels).Value
    '        olItem.m_OEI_H_Pixels = oRow.Cells(Columns.OEI_H_Pixels).Value

    '        olItem.Create_Traveler = True
    '        olItem.Create_Kit = True

    '        olItem.Loading = False

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    '    SaveMultiOrderLine = olItem

    'End Function

    ' Call this instance when you only have the index of the row
    Private Sub SaveOrderLine(ByVal plRow As Long, ByRef oOrdline As cOrdLine)

        Dim oRow As DataGridViewRow = dgvItems.Rows(plRow)

        SaveOrderLine(oRow, oOrdline)

    End Sub

    Private Sub dgvItems_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.RowEnter

        'm_oOrdline = SaveOrderLine(e.RowIndex)
        m_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID)
        SaveOrderLine(dgvItems.Rows(e.RowIndex), m_oOrdline)

        Call RefreshImprint(e.RowIndex)

    End Sub

#Region "Public properties ################################################"

    Public Property EntryType() As ProductLineEntryType
        Get
            EntryType = pleEntryType
        End Get
        Set(ByVal value As ProductLineEntryType)
            pleEntryType = value
        End Set
    End Property

#End Region

    Private Sub dgvItems_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvItems.CellValidating

        'Dim strFormattedValue As String

        If dgvItems.CurrentCell.RowIndex = -1 Then Exit Sub ' Triggered when labels are changed

        'e.Cancel = False

        If dgvItems.CurrentCell.ColumnIndex <> Columns.Location And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Req_Ship_Dt And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Qty_Ordered And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Qty_To_Ship And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Qty_Lines And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Unit_Price And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Discount_Pct And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Route And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Request_Dt And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Promise_Dt And
           dgvItems.CurrentCell.ColumnIndex <> Columns.BkOrd_Fg And
           dgvItems.CurrentCell.ColumnIndex <> Columns.End_User And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Imprint And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Imprint_Location And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Imprint_Color And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Num_Impr_1 And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Num_Impr_2 And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Num_Impr_3 And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Lamination And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Foil_Color And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Thread_Color And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Rounded_Corners And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Tip_In_PageTxt And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Packaging And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Refill And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Laser_Setup And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Comments And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Route And
           dgvItems.CurrentCell.ColumnIndex <> Columns.RouteID And
           dgvItems.CurrentCell.ColumnIndex <> Columns.ProductProof And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Prod_Cat And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Industry And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Special_Comments And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Printer_Comment And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Packer_Comment And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Printer_Instructions And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Packer_Instructions And
           dgvItems.CurrentCell.ColumnIndex <> Columns.Repeat_From_ID Then Exit Sub

        iRowPos = dgvItems.CurrentCell.RowIndex
        iColPos = dgvItems.CurrentCell.ColumnIndex

        'MB++ Actuellement, il y a un trou dans cette procedure.
        ' -------------------------------------------------------
        ' Si la QTY BO vient d'une autre commande et qu'on traite une commande récupérée (ce qui
        ' ne devrait pas arriver parce qu'on ne traitera que de nouvelles commandes), alors lors
        ' d'un retrait de qté de commandes BO, il faut rajouter du code pour valider la QTÉ BO
        ' de la commande actuelle parce qu'on traite la QTÉ BO comme étant globale pour le moment.
        '--------------------------------------------------------

        If blnLoading Then Exit Sub

        If m_oOrdline Is Nothing Then

            'm_oOrdline = New cOrdLine(False) ' m_oOrder, True)
            'm_oOrdline.Source = OEI_SourceEnum.frmProductLineEntry
            m_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID)
            SaveOrderLine(dgvItems.CurrentRow, m_oOrdline)
            'm_oOrdline = SaveOrderLine(dgvItems.CurrentRow)

        End If

        Try
            Select Case e.ColumnIndex
                Case Columns.Item_No
                    Try
                        m_oOrdline.SaveToDB = False
                        m_oOrdline.Item_No = RTrim(Trim(e.FormattedValue)).ToString

                        'If m_oOrder.Reload Then
                        '    dgvItems.BeginInvoke(New MyDelegate(AddressOf LoadPreOrder))
                        'Else
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value = m_oOrdline.Item_No.ToUpper
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Location).Value = m_oOrdline.Loc
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_Desc_1).Value = m_oOrdline.Item_Desc_1
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_Desc_2).Value = m_oOrdline.Item_Desc_2
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Ordered).Value = m_oOrdline.Qty_Ordered
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_To_Ship).Value = m_oOrdline.Qty_To_Ship
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Lines).Value = m_oOrdline.Qty_Lines
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Inventory).Value = m_oOrdline.Qty_Inventory
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Allocated).Value = m_oOrdline.Qty_Allocated
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_To_Ship).Value = m_oOrdline.Qty_Prev_To_Ship
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_BkOrd).Value = m_oOrdline.Qty_Prev_Bkord
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_On_Hand).Value = m_oOrdline.Qty_On_Hand
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = m_oOrdline.Unit_Price
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Request_Dt).Value = m_oOrdline.Request_Dt
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Promise_Dt).Value = m_oOrdline.Promise_Dt
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Req_Ship_Dt).Value = m_oOrdline.Req_Ship_Dt
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.UOM).Value = m_oOrdline.Uom
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.BkOrd_Fg).Value = m_oOrdline.BkOrd_Fg_Bln
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Recalc_SW).Value = m_oOrdline.Recalc_Sw_Bln
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.OEI_W_Pixels).Value = m_oOrdline.m_OEI_W_Pixels
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.OEI_H_Pixels).Value = m_oOrdline.m_OEI_H_Pixels
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Mix_Group).Value = m_oOrdline.m_Mix_Group
                        'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_Cd).Value = m_oOrdline.Item_Cd


                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    End Try

                Case Columns.Location
                    Try
                        'If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Location_Must_Be_Numeric, True)

                        m_oOrdline.Loc = (e.FormattedValue)

                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Ordered).Value = m_oOrdline.Qty_Ordered
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_To_Ship).Value = m_oOrdline.Qty_To_Ship
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Inventory).Value = m_oOrdline.Qty_Inventory
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Allocated).Value = m_oOrdline.Qty_Allocated
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_To_Ship).Value = m_oOrdline.Qty_Prev_To_Ship
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_BkOrd).Value = m_oOrdline.Qty_Prev_Bkord
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_On_Hand).Value = m_oOrdline.Qty_On_Hand
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = m_oOrdline.Unit_Price
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Loc_Status).Value = m_oOrdline.m_Loc_Status
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Mix_Group).Value = m_oOrdline.m_Mix_Group
                        'Call SaveOrderLine()

                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    End Try

                Case Columns.Qty_Ordered

                    Try
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                        m_oOrdline.Qty_Ordered_Allows_Zero = CDbl(e.FormattedValue)

                        '++ID 8.8.2018 check how many Qty for setup 
                        'conditions for Check Setup Qty
                        '   "SELECT ITEM_NO, EXTRA_14 FROM   imitmidx_sql WHERE  item_no = '" & pstrItem_No.Trim & "' And Extra_14 = 27 "
                        If Mid(m_oOrdline.Item_No, 1, 2) = "44" Then
                            If Not m_oOrder.CheckSetupQty(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, CDbl(e.FormattedValue)) Then
                                Throw New OEException(OEError.Too_Many_Setups_Entered, True, False) 'Too_Many_Setups_Entered
                            End If
                        End If



                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    End Try

                    'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Ordered).Value = m_oOrdline.Qty_Ordered
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_On_Hand).Value = m_oOrdline.Qty_On_Hand
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_To_Ship).Value = m_oOrdline.Qty_To_Ship
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Lines).Value = m_oOrdline.Qty_Lines
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Inventory).Value = m_oOrdline.Qty_Inventory
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Allocated).Value = m_oOrdline.Qty_Allocated
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_To_Ship).Value = m_oOrdline.Qty_Prev_To_Ship
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_BkOrd).Value = m_oOrdline.Qty_Prev_Bkord
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_9).Value = m_oOrdline.Extra_9
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_10).Value = m_oOrdline.Extra_10
                     '++ID 5.3.2018 added discount 
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Discount_Pct).Value = m_oOrdline.Discount_Pct


                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = m_oOrdline.Unit_Price


                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price


                    '++ID 05.07.2018 ----------------------------------------

                    Dim m_OrderType As String = ""
                    Dim m_byRegion As String = ""
                    Dim m_star As String = ""
                    Dim m_Mfg_Loc As String = ""

                    If String.IsNullOrEmpty(m_oOrder.Ordhead.User_Def_Fld_2) <> True Then m_OrderType = m_oOrder.Ordhead.User_Def_Fld_2

                    If String.IsNullOrEmpty(m_oOrder.Ordhead.Customer.region) <> True Then m_byRegion = m_oOrder.Ordhead.Customer.region

                    If String.IsNullOrEmpty(m_oOrder.Ordhead.Customer.ClassificationId) <> True Then m_star = m_oOrder.Ordhead.Customer.ClassificationId.ToUpper

                    If String.IsNullOrEmpty(m_oOrder.Ordhead.Mfg_Loc) <> True Then m_Mfg_Loc = m_oOrder.Ordhead.Mfg_Loc

                    'Dim m_Prod_Cat As String = ""

                    'If String.IsNullOrEmpty(m_oOrdline.Item_No.ToUpper) <> True Then
                    '    If String.IsNullOrEmpty(Return_Prod_Cat(m_oOrdline.Item_No.ToUpper)) <> True Then m_Prod_Cat = Return_Prod_Cat(m_oOrdline.Item_No.ToUpper) ' come from insertt line in the function
                    'End If

                    '        m_oOrdline.Discount_Pct = Ret_Discount(m_OrderType, m_byRegion, m_star, m_oOrdline.Item_No.ToUpper, m_oOrdline.Prod_Cat, m_oOrdline.Calc_Price, m_Mfg_Loc)

                    m_oOrdline.Discount_Pct = Ret_Discount(m_OrderType, m_byRegion, m_star, m_oOrdline.Item_No.ToUpper, m_oOrdline.Prod_Cat, m_oOrdline.Calc_Price, m_oOrdline.Loc)
                    '---------------------------- NEW -------------------------------



                    ' m_oOrdline.Discount_Pct = 55

                    '   MsgBox("Calc_Price is : " & m_oOrdline.Calc_Price)
                    '
                    If dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Discount_Pct).Value <> m_oOrdline.Discount_Pct Then
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Discount_Pct).Value = m_oOrdline.Discount_Pct
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price
                    End If

                    If m_oOrdline.Qty_Ordered = 0 Then
                        ' We reset selection to unselected if qty is 0
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.SelectSeq).Value = 0
                    Else
                        ' If already selected, do nothing else add new SelectSeq
                        If dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.SelectSeq).Value = 0 Then
                            m_SelectSeq += 1
                            dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.SelectSeq).Value = m_SelectSeq
                        End If
                    End If

                    ''++ID 12.22.2021 missing prop65
                    '  If CDbl(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Ordered).Value) <> CDbl(e.FormattedValue) Then
                    'm_oOrdline.MissingProp65(m_oOrdline, m_oOrder, m_oOrdline.m_Imprint)
                    'm_oOrdline.MissingProp65(m_oOrdline, m_oOrder)

                    'message in packer_instruction no need to fill mesage
                    'If dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Packer_Instructions).Value.ToString.Contains("*APPLY PROP 65 LABELS") Then
                    'Else
                    ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Packer_Instructions).Value = m_oOrdline.ImprintLine.Packer_Instructions
                    'End If '*APPLY PROP 65 LABELS

                  '  End If

                Case Columns.Qty_To_Ship

                    Try
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format)
                        'm_oOrdline.Qty_To_Ship = CDbl(e.FormattedValue)

                        Dim oResult As New MsgBoxResult
                        If CDbl(e.FormattedValue) > m_oOrdline.Qty_Ordered Then
                            oResult = MsgBox(New OEExceptionMessage(OEError.Qty_To_Ship_GT_Qty_Ordered).Message, MsgBoxStyle.YesNo, "Error")
                        Else
                            oResult = MsgBoxResult.Yes
                        End If

                        If oResult = MsgBoxResult.Yes Then
                            m_oOrdline.Qty_To_Ship = CDbl(e.FormattedValue)
                        Else
                            m_oOrdline.Qty_To_Ship = m_oOrdline.Qty_Ordered
                            Dim oInvoke As New SetQty_To_Ship(AddressOf SetQty_To_ShipSub)
                            dgvItems.BeginInvoke(oInvoke, m_oOrdline.Qty_Ordered)
                        End If

                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    End Try

                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_To_Ship).Value = m_oOrdline.Qty_To_Ship
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_To_Ship).Value = m_oOrdline.Qty_Prev_To_Ship
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_BkOrd).Value = m_oOrdline.Qty_Prev_Bkord
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_9).Value = m_oOrdline.Extra_9
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_10).Value = m_oOrdline.Extra_10

                Case Columns.Qty_Lines

                    Try
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format)
                        m_oOrdline.Qty_Lines = CDbl(e.FormattedValue)

                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    End Try

                    'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Ordered).Value = m_oOrdline.Qty_Ordered
                    'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_On_Hand).Value = m_oOrdline.Qty_On_Hand
                    'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_To_Ship).Value = m_oOrdline.Qty_To_Ship
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Inventory).Value = m_oOrdline.Qty_Inventory
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Allocated).Value = m_oOrdline.Qty_Allocated
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_To_Ship).Value = m_oOrdline.Qty_Prev_To_Ship
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_BkOrd).Value = m_oOrdline.Qty_Prev_Bkord
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_9).Value = m_oOrdline.Extra_9
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_10).Value = m_oOrdline.Extra_10
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = m_oOrdline.Unit_Price
                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price

                Case Columns.Qty_Prev_To_Ship

                Case Columns.Unit_Price

                    Try
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format)
                        m_oOrdline.Unit_Price = e.FormattedValue

                        '++ID 5.3.2018 added discount 
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Discount_Pct).Value = m_oOrdline.Discount_Pct

                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = m_oOrdline.Unit_Price
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price

                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    End Try

                Case Columns.Discount_Pct

                    Try
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                        m_oOrdline.Discount_Pct = CDbl(e.FormattedValue)

                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Discount_Pct).Value = m_oOrdline.Discount_Pct
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price
                        'Call SaveOrderLine()

                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    End Try

                Case Columns.Calc_Price

                    MsgBox("I am in Calcualate price")
                Case Columns.Request_Dt

                    Try
                        If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format)
                        m_oOrdline.Request_Dt = CDate(e.FormattedValue)

                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    Finally
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Promise_Dt).Value = m_oOrdline.Promise_Dt
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Req_Ship_Dt).Value = m_oOrdline.Req_Ship_Dt

                    End Try

                Case Columns.Promise_Dt

                    Try
                        If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
                        m_oOrdline.Promise_Dt = CDate(e.FormattedValue)

                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Req_Ship_Dt).Value = m_oOrdline.Req_Ship_Dt

                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    End Try

                Case Columns.Req_Ship_Dt
                    Try
                        If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format)
                        m_oOrdline.Req_Ship_Dt = CDate(e.FormattedValue)

                    Catch oe_er As OEException
                        e.Cancel = oe_er.Cancel
                        If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                    End Try

                Case Columns.Item_Desc_1
                Case Columns.Item_Desc_2
                Case Columns.UOM
                Case Columns.Tax_Sched
                Case Columns.Revision_No
                    'Case Columns.Explode_Kit
                    'Case Columns.NoteText
                Case Columns.BkOrd_Fg
                    'Case Columns.Select_Cd
                Case Columns.Recalc_SW
                    'Case Columns.Filler_0004
                Case Columns.Route
                    If e.FormattedValue = "" Then Exit Sub

                    'm_oOrdline.SetRoute(e.FormattedValue)
                    'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value = m_oOrdline.RouteID

                    m_oOrdline.SetRoute(e.FormattedValue)
                    m_oOrdline.ImprintLine.SetNumImprints(m_oOrdline.RouteID, frmOrder.getCurrentLineItems())


                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value = m_oOrdline.RouteID
                    If m_oOrdline.ImprintLine.Num_Impr_1 = 0 Then
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_1).Value = DBNull.Value
                    Else
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_1).Value = m_oOrdline.ImprintLine.Num_Impr_1
                    End If

                    If m_oOrdline.ImprintLine.Num_Impr_2 = 0 Then
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_2).Value = DBNull.Value
                    Else
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_2).Value = m_oOrdline.ImprintLine.Num_Impr_2
                    End If

                    If m_oOrdline.ImprintLine.Num_Impr_3 = 0 Then
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_3).Value = DBNull.Value
                    Else
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_3).Value = m_oOrdline.ImprintLine.Num_Impr_3
                    End If
                    '++ID 12.22.2021 check if item missing prop65
                    'm_oOrdline.MissingProp65(m_oOrdline.Item_No, m_oOrder, m_oOrdline.m_Imprint)
                    'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Packer_Instructions).Value = m_oOrdline.ImprintLine.Packer_Instructions


                    'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_2).Value = g_oOrdline.ImprintLine.Num_Impr_2
                    'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_3).Value = g_oOrdline.ImprintLine.Num_Impr_3

                    'Call RefreshImprint(dgvItems.CurrentRow.Index)

                    'dgvItems.BeginInvoke(New Del_SetProductionAvailabilityColor(AddressOf g_oOrdline.SetProductionAvailabilityColor), dgvItems.CurrentRow.Cells(Columns.Item_No)) ' , dgvItems.CurrentCell.ColumnIndex)




                    '07.30.2024----------------------- Test for NO IMPRINT 
                    '      If g_oOrdline.RouteID = 17 Then
                    If ValidateNoImprintRoute(m_oOrdline.RouteID) <> False And Trim(m_oOrder.Ordhead.Mfg_Loc.ToUpper) <> "US1" Then
                        '17 NO IMPRINT , but need to check because we have a lot of No Imprint need to be discussed
                        Dim oprice As New cOEPriceFile
                        oprice.LoadPriceByType(m_oOrdline.Item_No, m_oOrder.Ordhead.Curr_Cd, "10")

                        'above function normally  give the price by type 10 but in case if price missing , type 6 retrun EQP price (column price 6 from Macola price table )
                        If oprice.Cd_Tp = 10 Then
                            m_oOrdline.Unit_Price = oprice.Prc_Or_Disc_1
                            '  m_oOrdline.Loc = "3N"
                        Else 'CODE TYPE NEED TO BE 6
                            m_oOrdline.Unit_Price = oprice.Prc_Or_Disc_6
                            ' m_oOrdline.Loc = "3N"

                            ' g_oOrdline.Set_Spec_Instructions(g_oOrdline.Item_No, g_oOrdline.RouteID, g_oOrdline.ImprintLine)
                        End If


                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Location).Value = m_oOrdline.Loc


                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = m_oOrdline.Unit_Price


                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Inventory).Value = m_oOrdline.Qty_Inventory
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Allocated).Value = m_oOrdline.Qty_Allocated

                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_On_Hand).Value = m_oOrdline.Qty_On_Hand

                        'Added June 21, 2017 - T. Louzon
                        'CHANGE BEFORE SPECIAL PRICE WHEN ITEM_NO CHANGES
                        '   dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price_BeforeSpecial).Value = g_oOrdline.Unit_Price

                    Else
                        g_oOrdline.Loc = m_oOrder.Ordhead.Mfg_Loc
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Location).Value = m_oOrdline.Loc

                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = m_oOrdline.Calc_Price
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = m_oOrdline.Unit_Price


                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Inventory).Value = m_oOrdline.Qty_Inventory
                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Allocated).Value = m_oOrdline.Qty_Allocated

                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_On_Hand).Value = m_oOrdline.Qty_On_Hand


                        ''brought function from the top of route exception. commented in the top line 1748
                        'g_oOrdline.Set_Spec_Instructions(g_oOrdline.Item_No, g_oOrdline.RouteID, g_oOrdline.ImprintLine)

                    End If






                Case Columns.ProductProof

                Case Columns.Imprint
                    txtImprint.Text = e.FormattedValue
                    If e.FormattedValue.trim <> String.Empty Then
                        Dim oUpdateFrom As cImprint
                        oUpdateFrom = GetImprintRow()
                        oUpdateFrom.Imprint = e.FormattedValue
                        If g_oOrdline.Set_Spec_Imprint_Instructions(m_oOrder.Ordhead.Cus_No, oUpdateFrom) Then
                            dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Printer_Instructions).Value = oUpdateFrom.Printer_Instructions
                            dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Packer_Instructions).Value = oUpdateFrom.Packer_Instructions
                        End If

                    End If

                Case Columns.End_User
                    txtEnd_User.Text = e.FormattedValue

                Case Columns.Imprint_Location
                    txtImprint_Location.Text = e.FormattedValue

                Case Columns.Imprint_Color
                    ' Do nothing for imprint data, it will be collected on save to frmOrder
                    txtImprint_Color.Text = e.FormattedValue

                Case Columns.Num_Impr_1
                    txtNum_Impr_1.Text = e.FormattedValue

                Case Columns.Num_Impr_2
                    txtNum_Impr_2.Text = e.FormattedValue

                Case Columns.Num_Impr_3
                    txtNum_Impr_3.Text = e.FormattedValue

                Case Columns.Packaging
                    txtPackaging.Text = e.FormattedValue

                Case Columns.Refill
                    txtRefill.Text = e.FormattedValue

                Case Columns.Laser_Setup
                        'txtRefill.Text = e.FormattedValue

                Case Columns.Industry
                    cboIndustry.Text = e.FormattedValue

 '++ID 07.28.2020 scribl project 
                Case Columns.Lamination
                    cboLamination.Text = e.FormattedValue
                Case Columns.Foil_Color '++ID 12.09.2021 Foil Color
                    cboFoil_Color.Text = e.FormattedValue
                Case Columns.Thread_Color
                    cbothreadColorScribl.Text = e.FormattedValue
                Case Columns.Rounded_Corners
                    cboRoundedCorners.Text = e.FormattedValue
                Case Columns.Tip_In_PageTxt
                    cboTipInPageTxt.Text = e.FormattedValue

 '-------------------------------------
                Case Columns.Comments
                    txtComments.Text = e.FormattedValue

                Case Columns.Special_Comments
                    txtSpecial_Comments.Text = e.FormattedValue

            End Select

        Catch er_oe As OEException
            MsgBox(er_oe.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'function used to validate as Route selected is one of  NO IMPRINT
    Private Function ValidateNoImprintRoute(ByVal _routeid As Int32) As Boolean
        ValidateNoImprintRoute = False
        Try
            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = " Select * FROM EXACT_TRAVELER_ROUTE  " _
                   & " Where RouteDescription Like '%no imprint%'  AND RouteId = " & _routeid & " And Active = 1 "


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                ValidateNoImprintRoute = True

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Function
    Private Sub dgvItems_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvItems.EditingControlShowing

        'Make sure we are in the first column.
        'If Me.dgvItems.CurrentCellAddress.X = 0 Then
        '    With e.Control
        If TypeOf e.Control Is TextBox Then

            Dim tb As TextBox = CType(e.Control, TextBox)
            'Remove the existing handler if there is one.
            RemoveHandler tb.KeyDown, AddressOf XdgvItems_KeyDown ' TextBox_TextChanged

            'Add a new handler.
            AddHandler tb.KeyDown, AddressOf XdgvItems_KeyDown ' TextBox_TextChanged
            '    End With
            'End If

        End If

    End Sub

    Private Sub XdgvItems_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) 'Handles dgvItems.KeyDown

        Dim txtElement As New Object

        Try
            'If TypeOf sender Is xTextBox Then

            '    txtElement = New xTextBox
            '    txtElement = DirectCast(sender, xTextBox)

            'ElseIf TypeOf sender Is ComboBox Then

            '    txtElement = New ComboBox
            '    txtElement = DirectCast(sender, ComboBox)

            'End If

            Select Case e.KeyValue

                ' Save changes ?
                Case Keys.Escape
                    'If txtElement.tabindex <txtCus_No.TabIndex Then
                    '    MsgBox("On supprime tout")
                    'Else
                    '    Dim oReponse As New MsgBoxResult
                    '    oReponse = MsgBox("Save Changes ?", MsgBoxStyle.YesNoCancel, "OE Interface")
                    '    If oReponse = MsgBoxResult.Yes Then
                    '        MsgBox("On save")
                    '    ElseIf oReponse = MsgBoxResult.No Then
                    '        MsgBox("on annule la commande")
                    '    Else
                    '        MsgBox("on fait rien")
                    '    End If
                    'End If

                    ' Search
            Case Keys.F7
                    Dim oButton As Button
                    oButton = GetSearchControlByColumnIndex() ' (sender.name)
                    If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

                    ' Notes
                Case Keys.F8
                    Dim oButton As Button
                    oButton = GetNotesControlByColumnIndex() ' (sender.name)
                    If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

                    ' Calendar
                    'Case Keys.F9
                    '    Dim oButton As Button
                    '    oButton = GetDateControlByControlName(sender.name)
                    '    If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#Region "Imprints #############################################################"

    Private blnImprintLoading As Boolean = False
    Private m_ComboFill As Boolean = False

    Private Sub SetColumnShortcuts()

        Select Case dgvItems.CurrentCell.ColumnIndex
            'Case Columns.Item_No
            '    _sbOrder_Panel1.Text = g_Search_Shortcut & g_Notes_Shortcut & g_Item_No_Shortcut

            'Case Columns.Location
            '    _sbOrder_Panel1.Text = g_Search_Shortcut & g_Notes_Shortcut

            Case Columns.Imprint
                _sbOrder_Panel1.Text = g_Imprint_Shortcut

            Case Columns.End_User
                _sbOrder_Panel1.Text = ""

            Case Columns.Imprint_Location
                _sbOrder_Panel1.Text = g_Imprint_Location_Shortcut

            Case Columns.Imprint_Color
                _sbOrder_Panel1.Text = g_Imprint_Color_Shortcut

            Case Columns.Packaging
                _sbOrder_Panel1.Text = g_Imprint_Packaging_Shortcut

            Case Else
                _sbOrder_Panel1.Text = " "

        End Select

    End Sub

    Private Sub cmdClear_Imprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear_Imprint.Click

        Call ImprintReset()

        'UcImprint1.Reset()

    End Sub

    Private Sub RefreshImprint(ByVal pintRow As Integer)

        Try
            If dgvItems.Rows.Count > 0 Then
                txtImprint_Item_No.Text = dgvItems.Rows(pintRow).Cells(Columns.Item_No).Value.ToString
            Else
                txtImprint_Item_No.Text = ""
            End If

            Call Enable_Imprint_Fields()

            If m_oOrdline.ImprintLine Is Nothing Then
                Call ImprintReset()
            Else
                Call ImprintFill(m_oOrdline.ImprintLine)
            End If
            'UcImprint1.ImprintFill(g_oOrdline.ImprintLine)
            'UcImprint1.txtImprint = dgvItems.Rows(pintRow).Cells(Columns.Item_No).Value

            If Trim(txtImprint_Item_No.Text.ToString).Length > 3 Then
                cmdUpdateAllItem.Enabled = True
                If Trim(txtImprint_Item_No.Text.ToString).Length > 11 Then
                    cmdUpdateAllItem.Text = "Update All " & Mid(txtImprint_Item_No.Text.ToString, 1, 8)
                Else
                    cmdUpdateAllItem.Text = "Update All " & Mid(txtImprint_Item_No.Text.ToString, 1, Trim(txtImprint_Item_No.Text.ToString).Length - 3)
                End If
            Else
                cmdUpdateAllItem.Enabled = False
                cmdUpdateAllItem.Text = "Update All xxxx"
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Enable_Imprint_Fields()

        Try
            Dim blnEnabled As Boolean = False

            blnEnabled = Not (Trim(txtImprint_Item_No.Text) = "")

            If blnEnabled = txtImprint.Enabled Then Exit Sub

            txtImprint_Item_No.Enabled = blnEnabled
            txtRepeat_Ord_No.Enabled = blnEnabled

            txtImprint.Enabled = blnEnabled
            txtEnd_User.Enabled = blnEnabled
            txtImprint_Location.Enabled = blnEnabled
            txtImprint_Color.Enabled = blnEnabled
            txtNum_Impr_1.Enabled = blnEnabled
            txtNum_Impr_2.Enabled = blnEnabled
            txtNum_Impr_3.Enabled = blnEnabled
            txtPackaging.Enabled = blnEnabled
            txtRefill.Enabled = blnEnabled
            txtIndustry.Enabled = blnEnabled
            txtComments.Enabled = blnEnabled
            txtSpecial_Comments.Enabled = blnEnabled

            '++ID 06.29.2020 added ------------------
            txtLamination.Enabled = blnEnabled
            txtThreadColorScribl.Enabled = blnEnabled
            txtRoundedCorners.Enabled = blnEnabled

            cboLamination.Enabled = blnEnabled
            cbothreadColorScribl.Enabled = blnEnabled
            cboRoundedCorners.Enabled = blnEnabled
            '----------------------------------------
            '++ID 08.06.2020 ------------------------
            cboTipInPageTxt.Enabled = blnEnabled
            txtTipInPageTxt.Enabled = blnEnabled
            '----------------------------------------

            cboImprint_Location.Enabled = blnEnabled
            cboImprint_Color.Enabled = blnEnabled
            cboPackaging.Enabled = blnEnabled
            cboRefill.Enabled = blnEnabled
            cboIndustry.Enabled = blnEnabled

            cmdClear_Imprint.Enabled = blnEnabled
            cmdGetRepeatData.Enabled = blnEnabled
            cmdUpdateAllItem.Enabled = blnEnabled
            cmdUpdateAll.Enabled = blnEnabled
            cmdUpdateSelected.Enabled = blnEnabled
            cmdCopy.Enabled = blnEnabled

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ImprintFill() ' ByRef pOrder As cOrder)

        Try
            Call ImprintComboFill()

            blnImprintLoading = True

            txtImprint.Text = m_oOrdline.ImprintLine.Imprint
            txtEnd_User.Text = m_oOrdline.ImprintLine.End_User
            txtImprint_Location.Text = m_oOrdline.ImprintLine.Location
            txtNum_Impr_1.Text = m_oOrdline.ImprintLine.Num_Impr_1
            txtNum_Impr_2.Text = m_oOrdline.ImprintLine.Num_Impr_2
            txtNum_Impr_3.Text = m_oOrdline.ImprintLine.Num_Impr_3
            txtImprint_Color.Text = m_oOrdline.ImprintLine.Color
            txtPackaging.Text = m_oOrdline.ImprintLine.Packaging
            txtRefill.Text = m_oOrdline.ImprintLine.Refill
            'txtLaserSetup.text = m_Imprint.LaserSetup
            cboIndustry.Text = m_oOrdline.ImprintLine.Industry
            txtComments.Text = m_oOrdline.ImprintLine.Comments
            txtSpecial_Comments.Text = m_oOrdline.ImprintLine.Special_Comments

            cboImprint_Color.Text = ""
            cboImprint_Location.Text = ""
            cboRefill.Text = ""
            cboPackaging.Text = ""
            'cboIndustry.Text = ""

            '++ID 06.29.2020 -------- oOrdline.ImprintLine come from Set_Spec_instruction , we can use if we have sec instruction for item -------------------
            txtLamination.Text = m_oOrdline.ImprintLine.Lamination_Scribl
            '++ID 12.09.2021 Foil Color
            txtFoil_Color.Text = m_oOrdline.ImprintLine.Foil_Color

            txtThreadColorScribl.Text = m_oOrdline.ImprintLine.Thread_Color_Scribl
            txtRoundedCorners.Text = m_oOrdline.ImprintLine.Rounded_corners_Scribl

            cboLamination.Text = ""
            cbothreadColorScribl.Text = ""
            cboRoundedCorners.Text = ""
            '-------------------------------------------
            '++ID 08.06.2020
            cboTipInPageTxt.Text = ""
            txtTipInPageTxt.Text = m_oOrdline.ImprintLine.Tip_In_PageTxt
            '-------------------------------------------

            blnImprintLoading = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ImprintFill(ByRef poImprint As cImprint)

        Try
            m_oOrdline.ImprintLine = poImprint
            Call ImprintFill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ImprintReset() ' ByRef pOrder As cOrder)

        Try
            If m_oOrdline.ImprintLine Is Nothing Then
                txtImprint_Item_No.Text = String.Empty
                txtImprint.Text = String.Empty
                txtEnd_User.Text = String.Empty
                txtImprint_Location.Text = String.Empty
                txtImprint_Color.Text = String.Empty
                txtNum_Impr_1.Text = String.Empty
                txtNum_Impr_2.Text = String.Empty
                txtNum_Impr_3.Text = String.Empty
                txtPackaging.Text = String.Empty
                txtRefill.Text = String.Empty
                txtIndustry.Text = String.Empty
                txtComments.Text = String.Empty
                txtSpecial_Comments.Text = String.Empty
                '++ID 07.28.2020 complited for lamination etc----------------------------------
                txtLamination.Text = String.Empty
                txtThreadColorScribl.Text = String.Empty
                txtRoundedCorners.Text = String.Empty

                '++ID 08.06.2020
                txtTipInPageTxt.Text = String.Empty

            Else
                m_oOrdline.ImprintLine.Reset()
                Call ImprintFill()
                With dgvItems.Rows(dgvItems.CurrentRow.Index)
                    .Cells(Columns.Imprint).Value = m_oOrdline.ImprintLine.Imprint
                    .Cells(Columns.End_User).Value = m_oOrdline.ImprintLine.End_User
                    .Cells(Columns.Imprint_Location).Value = m_oOrdline.ImprintLine.Location
                    .Cells(Columns.Imprint_Color).Value = m_oOrdline.ImprintLine.Color
                    .Cells(Columns.Num_Impr_1).Value = m_oOrdline.ImprintLine.Num_Impr_1
                    .Cells(Columns.Num_Impr_2).Value = m_oOrdline.ImprintLine.Num_Impr_2
                    .Cells(Columns.Num_Impr_3).Value = m_oOrdline.ImprintLine.Num_Impr_3
                    .Cells(Columns.Packaging).Value = m_oOrdline.ImprintLine.Packaging
                    .Cells(Columns.Refill).Value = m_oOrdline.ImprintLine.Refill
                    .Cells(Columns.Laser_Setup).Value = m_oOrdline.ImprintLine.Laser_Setup
                    .Cells(Columns.Industry).Value = m_oOrdline.ImprintLine.Industry
                    .Cells(Columns.Comments).Value = m_oOrdline.ImprintLine.Comments
                    .Cells(Columns.Special_Comments).Value = m_oOrdline.ImprintLine.Special_Comments
                    '.Cells(Columns.Printer_Comment).Value = m_oOrdline.ImprintLine.Printer_Comment
                    '.Cells(Columns.Packer_Comment).Value = m_oOrdline.ImprintLine.Packer_Comment
                    '.Cells(Columns.Printer_Instructions).Value = m_oOrdline.ImprintLine.Printer_Instructions
                    '.Cells(Columns.Packer_Instructions).Value = m_oOrdline.ImprintLine.Packer_Instructions
                    .Cells(Columns.Repeat_From_ID).Value = m_oOrdline.ImprintLine.Repeat_From_ID

                    '++ID 06.08.2020 scribl project ---------------------------------------------
                    .Cells(Columns.Lamination).Value = m_oOrdline.ImprintLine.Lamination_Scribl
                    '++ID 12.09.2021 Foil Color
                    .Cells(Columns.Foil_Color).Value = m_oOrdline.ImprintLine.Foil_Color

                    .Cells(Columns.Thread_Color).Value = m_oOrdline.ImprintLine.Thread_Color_Scribl
                    .Cells(Columns.Rounded_Corners).Value = m_oOrdline.ImprintLine.Rounded_corners_Scribl
                    '----------------------------------------------------------------------------
                    '++ID 07.28.2020
                    .Cells(Columns.Tip_In_PageTxt).Value = m_oOrdline.ImprintLine.Tip_In_PageTxt
                    '----------------------------------------------------------------------------

                End With
            End If
            txtRepeat_Ord_No.Text = String.Empty
            cboImprint_Color.Text = ""
            cboImprint_Location.Text = ""
            cboRefill.Text = ""
            cboPackaging.Text = ""
            cboIndustry.Text = ""

            '++ID 06.29.2020 -----------------------------------------
            cboLamination.Text = ""
            cbothreadColorScribl.Text = ""
            cboRoundedCorners.Text = ""

            '++ID 08.06.2020
            cboTipInPageTxt.Text = ""
            '---------------------------------------------------------


            With dgvItems.Rows(dgvItems.CurrentRow.Index)
                .Cells(Columns.RouteID).Value = 0
                .Cells(Columns.Route).Value = ""
                .Cells(Columns.ProductProof).Value = False
            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ImprintSave()

        Try
            ' We don't save OrdGUID and ItemGUID as they always stay the same here.
            m_oOrdline.ImprintLine.Imprint = txtImprint.Text
            m_oOrdline.ImprintLine.End_User = txtEnd_User.Text
            m_oOrdline.ImprintLine.Location = txtImprint_Location.Text
            If IsNumeric(txtNum_Impr_1.Text) Then
                m_oOrdline.ImprintLine.Num_Impr_1 = CInt(txtNum_Impr_1.Text)
            Else
                m_oOrdline.ImprintLine.Num_Impr_1 = 0
            End If
            If IsNumeric(txtNum_Impr_2.Text) Then
                m_oOrdline.ImprintLine.Num_Impr_2 = CInt(txtNum_Impr_2.Text)
            Else
                m_oOrdline.ImprintLine.Num_Impr_2 = 0
            End If
            If IsNumeric(txtNum_Impr_3.Text) Then
                m_oOrdline.ImprintLine.Num_Impr_3 = CInt(txtNum_Impr_3.Text)
            Else
                m_oOrdline.ImprintLine.Num_Impr_3 = 0
            End If
            m_oOrdline.ImprintLine.Color = txtImprint_Color.Text
            m_oOrdline.ImprintLine.Packaging = txtPackaging.Text
            m_oOrdline.ImprintLine.Refill = txtRefill.Text
            ' m_oOrdline.ImprintLine.LaserSetup = txtLaserSetup.text
            m_oOrdline.ImprintLine.Industry = txtIndustry.Text
            m_oOrdline.ImprintLine.Comments = txtComments.Text
            m_oOrdline.ImprintLine.Special_Comments = txtSpecial_Comments.Text




        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ImprintComboFill()

        Dim strSql As String
        Dim conn As New cDBA
        Dim dtCombo As New DataTable

        If m_ComboFill Then Exit Sub

        Try
            strSql = "Select * from EXACT_TRAVELER_EXTRA_FIELDS_xRef order by FieldValue Asc "
            dtCombo = conn.DataTable(strSql)

            If dtCombo.Rows.Count <> 0 Then

                Dim drComboElement As DataRow

                For Each drComboElement In dtCombo.Rows

                    With drComboElement

                        Select Case .Item("Field")
                            Case "ImprintLocation"
                                cboImprint_Location.Items.Add(.Item("FieldValue"))
                            Case "ImprintColour"
                                cboImprint_Color.Items.Add(.Item("FieldValue"))
                            Case "Packaging"
                                cboPackaging.Items.Add(.Item("FieldValue"))
                            Case "Refill"
                                cboRefill.Items.Add(.Item("FieldValue"))
                        End Select

                    End With

                Next

                m_ComboFill = True

            End If

            dtCombo = m_oOrder.GetCompleteCboIndustryList()

            If dtCombo.Rows.Count <> 0 Then

                Dim drComboElement As DataRow

                For Each drComboElement In dtCombo.Rows

                    With drComboElement

                        cboIndustry.Items.Add(.Item("Industry"))

                    End With

                Next

            End If

            '++ID 06.29.2020 ------------------------------------------------
            With cboLamination
                .DataSource = m_oOrder.Get_Scribl_Fields("LAMINATION_SCRIBL")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With
            '++ID 12.09.2021 Foil Color
            With cboFoil_Color
                .DataSource = m_oOrder.Get_Scribl_Fields("Foil_Color")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With

            With cbothreadColorScribl
                .DataSource = m_oOrder.Get_Scribl_Fields("Thread_Color_Scribl")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With

            With cboRoundedCorners
                .DataSource = m_oOrder.Get_Scribl_Fields("Rounded_Corners_Scribl")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With

            With cboTipInPageTxt
                .DataSource = m_oOrder.Get_Scribl_Fields("Tip_In_PageTxt")
                .ValueMember = "ENUM_VALUE"
                .DisplayMember = "ENUM_VALUE"
            End With
            '----------------------------------------------------------------


        Catch er As Exception
            MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public ReadOnly Property Order() As cOrder
    '    Get
    '        Order = m_oOrder
    '    End Get
    'End Property

    Private Sub cboImprintLocation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboImprint_Location.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            m_oOrdline.ImprintLine.AddLocation(cboImprint_Location.SelectedItem.ToString, txtImprint_Location.Text)
            txtImprint_Location.Text = m_oOrdline.ImprintLine.Location
            dgvItems.CurrentRow.Cells(Columns.Imprint_Location).Value = txtImprint_Location.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cboImprintColor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboImprint_Color.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            m_oOrdline.ImprintLine.AddColor(cboImprint_Color.SelectedItem.ToString, txtImprint_Color.Text)
            txtImprint_Color.Text = m_oOrdline.ImprintLine.Color
            dgvItems.CurrentRow.Cells(Columns.Imprint_Color).Value = txtImprint_Color.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '++ID 06.29.2020 added new event for new fields -------------------------------------------------
    Private Sub cboLamination_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLamination.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            If cboLamination.SelectedIndex <> 0 And cboLamination.SelectedIndex <> -1 Then

                m_oOrdline.ImprintLine.AddLamination(cboLamination.SelectedValue.ToString, txtLamination.Text)
                txtLamination.Text = m_oOrdline.ImprintLine.Lamination_Scribl
                dgvItems.CurrentRow.Cells(Columns.Lamination).Value = txtLamination.Text

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    '++ID 12.09.2021 Foil Color
    Private Sub cboFoil_Color_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFoil_Color.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            If cboFoil_Color.SelectedIndex <> 0 And cboFoil_Color.SelectedIndex <> -1 Then

                m_oOrdline.ImprintLine.AddFoil_Color(cboFoil_Color.SelectedValue.ToString, txtFoil_Color.Text)
                txtFoil_Color.Text = m_oOrdline.ImprintLine.Foil_Color
                dgvItems.CurrentRow.Cells(Columns.Foil_Color).Value = txtFoil_Color.Text

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub cbothreadColorScribl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbothreadColorScribl.SelectedIndexChanged
        Try
            If blnImprintLoading Then Exit Sub

            If cbothreadColorScribl.SelectedIndex <> 0 And cbothreadColorScribl.SelectedIndex <> -1 Then
                m_oOrdline.ImprintLine.AddThreadColor_Scribl(cbothreadColorScribl.SelectedValue.ToString, txtThreadColorScribl.Text)
                txtThreadColorScribl.Text = m_oOrdline.ImprintLine.Thread_Color_Scribl
                dgvItems.CurrentRow.Cells(Columns.Thread_Color).Value = txtThreadColorScribl.Text
            End If
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub


    Private Sub cboRoundedCorners_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboRoundedCorners.SelectedIndexChanged
        Try
            If blnImprintLoading Then Exit Sub
            If cboRoundedCorners.SelectedIndex <> 0 And cboRoundedCorners.SelectedIndex <> -1 Then
                m_oOrdline.ImprintLine.AddRoundedCorners(cboRoundedCorners.SelectedValue.ToString, txtRoundedCorners.Text)
                txtRoundedCorners.Text = m_oOrdline.ImprintLine.Rounded_corners_Scribl
                dgvItems.CurrentRow.Cells(Columns.Rounded_Corners).Value = txtRoundedCorners.Text
            End If


        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub
    '++ID 08.06.2020
    Private Sub cboTipInPageTxt_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTipInPageTxt.SelectedIndexChanged
        Try
            If blnImprintLoading Then Exit Sub
            If cboTipInPageTxt.SelectedIndex <> 0 And cboTipInPageTxt.SelectedIndex <> -1 Then
                m_oOrdline.ImprintLine.AddTip_In_PageTxt(cboTipInPageTxt.SelectedValue.ToString, txtTipInPageTxt.Text)
                txtTipInPageTxt.Text = m_oOrdline.ImprintLine.Tip_In_PageTxt
                dgvItems.CurrentRow.Cells(Columns.Tip_In_PageTxt).Value = txtTipInPageTxt.Text
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub txtTipInPageTxt_Leave(sender As Object, e As EventArgs) Handles txtTipInPageTxt.Leave
        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Tip_In_PageTxt = txtTipInPageTxt.Text
            dgvItems.CurrentRow.Cells(Columns.Tip_In_PageTxt).Value = txtTipInPageTxt.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub


    Private Sub txtLamination_Leave(sender As Object, e As EventArgs) Handles txtLamination.Leave
        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Lamination_Scribl = txtLamination.Text
            dgvItems.CurrentRow.Cells(Columns.Lamination).Value = txtLamination.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub
    '++ID 12.09.2021 Foil Color
    Private Sub txtFoil_Color_Leave(sender As Object, e As EventArgs) Handles txtFoil_Color.Leave
        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Foil_Color = txtFoil_Color.Text
            dgvItems.CurrentRow.Cells(Columns.Foil_Color).Value = txtFoil_Color.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub txtThreadColorScribl_Leave(sender As Object, e As EventArgs) Handles txtThreadColorScribl.Leave
        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Thread_Color_Scribl = txtThreadColorScribl.Text
            dgvItems.CurrentRow.Cells(Columns.Thread_Color).Value = txtThreadColorScribl.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub txtRoundedCorners_Leave(sender As Object, e As EventArgs) Handles txtRoundedCorners.Leave
        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Rounded_corners_Scribl = txtRoundedCorners.Text
            dgvItems.CurrentRow.Cells(Columns.Rounded_Corners).Value = txtRoundedCorners.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub
    '----------------------------------------------------------------------------------------------------

    '++ID 07.28.2020
    'Private Sub chkTipInPage_CheckStateChanged(sender As Object, e As EventArgs) Handles chkTipInPage.CheckStateChanged
    '    Try
    '        If blnImprintLoading Then Exit Sub

    '        If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

    '        m_oOrdline.ImprintLine.AddTip_In_Page(chkTipInPage.Checked)

    '        dgvItems.CurrentRow.Cells(Columns.Tip_In_Page).Value = chkTipInPage.Checked

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try
    'End Sub
    'Private Sub chkTipInPage_Leave(sender As Object, e As EventArgs) Handles chkTipInPage.Leave
    '    Try
    '        If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

    '        m_oOrdline.ImprintLine.Tip_In_Page = chkTipInPage.Checked
    '        dgvItems.CurrentRow.Cells(Columns.Tip_In_Page).Value = chkTipInPage.Checked
    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try
    'End Sub

    '-------------------------------------------------------------------------------------------


    Private Sub cboPackaging_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPackaging.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            m_oOrdline.ImprintLine.AddPackaging(cboPackaging.SelectedItem.ToString, txtPackaging.Text)
            txtPackaging.Text = m_oOrdline.ImprintLine.Packaging
            dgvItems.CurrentRow.Cells(Columns.Packaging).Value = txtPackaging.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cboRefill_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRefill.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            m_oOrdline.ImprintLine.AddRefill(cboRefill.SelectedItem.ToString, txtRefill.Text)
            txtRefill.Text = m_oOrdline.ImprintLine.Refill
            dgvItems.CurrentRow.Cells(Columns.Refill).Value = txtRefill.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtImprint_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImprint.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Imprint = txtImprint.Text
            dgvItems.CurrentRow.Cells(Columns.Imprint).Value = txtImprint.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtEnd_User_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEnd_User.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.End_User = txtEnd_User.Text
            dgvItems.CurrentRow.Cells(Columns.End_User).Value = txtEnd_User.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtImprintLocation_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImprint_Location.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Location = txtImprint_Location.Text
            dgvItems.CurrentRow.Cells(Columns.Imprint_Location).Value = txtImprint_Location.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtImprintColor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImprint_Color.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Color = txtImprint_Color.Text
            dgvItems.CurrentRow.Cells(Columns.Imprint_Color).Value = txtImprint_Color.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtNumImpr_1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNum_Impr_1.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Num_Impr_1 = txtNum_Impr_1.Text
            dgvItems.CurrentRow.Cells(Columns.Num_Impr_1).Value = txtNum_Impr_1.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtNumImpr_2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNum_Impr_2.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Num_Impr_2 = txtNum_Impr_2.Text
            dgvItems.CurrentRow.Cells(Columns.Num_Impr_2).Value = txtNum_Impr_2.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtNum_Impr_3_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNum_Impr_3.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Num_Impr_3 = txtNum_Impr_3.Text
            dgvItems.CurrentRow.Cells(Columns.Num_Impr_3).Value = txtNum_Impr_3.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtPackaging_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPackaging.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Packaging = txtPackaging.Text
            dgvItems.CurrentRow.Cells(Columns.Packaging).Value = txtPackaging.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtRefill_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRefill.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Refill = txtRefill.Text
            dgvItems.CurrentRow.Cells(Columns.Refill).Value = txtRefill.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtIndustry_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndustry.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Industry = txtIndustry.Text
            dgvItems.CurrentRow.Cells(Columns.Industry).Value = txtIndustry.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtComments_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtComments.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Comments = txtComments.Text
            dgvItems.CurrentRow.Cells(Columns.Comments).Value = txtComments.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtSpecialComments_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSpecial_Comments.Leave

        Try
            If m_oOrdline.ImprintLine Is Nothing Then Exit Sub

            m_oOrdline.ImprintLine.Special_Comments = txtSpecial_Comments.Text
            dgvItems.CurrentRow.Cells(Columns.Special_Comments).Value = txtSpecial_Comments.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Property Imprint(ByVal p_Imprint As cImprint)
    '    Get
    '        Imprint = m_Imprint
    '    End Get
    '    Set(ByVal value)
    '        m_Imprint = p_Imprint
    '    End Set
    'End Property

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Call Reset()

    End Sub

    Private Sub cmdUpdateAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateAll.Click

        Try
            If dgvItems.Rows.Count < 2 Then Exit Sub

            Dim intRouteID As Integer = dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            Dim strRoute As String = dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
            Dim intProductProof As Integer = IIf(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.ProductProof).Value, 1, 0)

            Dim oUpdateFrom As cImprint
            oUpdateFrom = GetImprintRow()

            For lPos As Integer = 0 To dgvItems.Rows.Count - 1

                If Not (dgvItems.Rows(lPos).Cells(Columns.Item_No).Value.ToString = oUpdateFrom.Item_No) Then

                    If dgvItems.Rows(lPos).Cells(Columns.Qty_Ordered).Value <> 0 Then

                        SetImprintRow(lPos, oUpdateFrom)

                        With dgvItems.Rows(lPos)
                            .Cells(Columns.RouteID).Value = intRouteID
                            .Cells(Columns.Route).Value = strRoute
                            .Cells(Columns.ProductProof).Value = (intProductProof = 1)
                        End With

                    End If

                End If

            Next lPos

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdUpdateAllItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateAllItem.Click

        Try
            If dgvItems.Rows.Count < 2 Then Exit Sub

            Dim intRouteID As Integer = dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            Dim strRoute As String = dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
            Dim intProductProof As Integer = IIf(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.ProductProof).Value, 1, 0)

            Dim oUpdateFrom As cImprint
            Dim strStart As String

            oUpdateFrom = GetImprintRow()

            If oUpdateFrom.Item_No.Length > 11 Then
                strStart = Mid(oUpdateFrom.Item_No, 1, 8)
            Else
                strStart = Mid(oUpdateFrom.Item_No, 1, Trim(oUpdateFrom.Item_No).Length - 3)
            End If

            For lPos As Integer = 0 To dgvItems.Rows.Count - 1

                If Not (dgvItems.Rows(lPos).Cells(Columns.Item_No).Value.ToString = oUpdateFrom.Item_No) Then

                    If dgvItems.Rows(lPos).Cells(Columns.Item_No).Value.ToString <> "" Then

                        If InStr(dgvItems.Rows(lPos).Cells(Columns.Item_No).Value.ToString, strStart) > 0 Then

                            SetImprintRow(lPos, oUpdateFrom)

                            With dgvItems.Rows(lPos)
                                .Cells(Columns.RouteID).Value = intRouteID
                                .Cells(Columns.Route).Value = strRoute
                                .Cells(Columns.ProductProof).Value = (intProductProof = 1)
                            End With

                        End If

                    End If

                End If

            Next lPos

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdCopyDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopy.Click

        Try
            ' Exit if only 1 row
            If dgvItems.Rows.Count < 2 Then Exit Sub

            ' Exit if on the last row (cannot copy down...)
            If dgvItems.CurrentRow.Index + 1 = dgvItems.Rows.Count Then Exit Sub

            ' Exit if on the next to last row and last row item_no is empty (cannot copy down...)
            If dgvItems.CurrentRow.Index + 2 = dgvItems.Rows.Count And dgvItems.Rows(dgvItems.CurrentRow.Index + 1).Cells(Columns.Item_No).Value.ToString = "" Then Exit Sub

            If dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString = "" Then
                intCopyRouteID = 0
            Else
                intCopyRouteID = CInt(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString)
            End If
            'intCopyRouteID = dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            strCopyRoute = dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
            intProductProof = IIf(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.ProductProof).Value, 1, 0)

            oCopyFrom = GetImprintRow()

            'SetImprintRow(dgvItems.CurrentRow.Index + 1, oUpdateFrom, intRouteID, strRoute)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        'Try
        '    ' Exit if only 1 row
        '    If dgvItems.Rows.Count < 2 Then Exit Sub

        '    ' Exit if on the last row (cannot copy down...)
        '    If dgvItems.CurrentRow.Index + 1 = dgvItems.Rows.Count Then Exit Sub

        '    ' Exit if on the next to last row and last row item_no is empty (cannot copy down...)
        '    If dgvItems.CurrentRow.Index + 2 = dgvItems.Rows.Count And dgvItems.Rows(dgvItems.CurrentRow.Index + 1).Cells(Columns.Item_No).Value.ToString = "" Then Exit Sub

        '    Dim oUpdateFrom As cImprint

        '    oUpdateFrom = GetImprintRow()

        '    SetImprintRow(dgvItems.CurrentRow.Index + 1, oUpdateFrom)

        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

    Private Sub cmdUpdateSelected_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateSelected.Click

        Try
            If oCopyFrom Is Nothing Then Exit Sub

            '' Exit if on the next to last row and last row item_no is empty (cannot copy down...)
            If dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString = "" Then Exit Sub

            'Dim intRouteID As Integer = g_oOrdline.RouteID 'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            'Dim strRoute As String = g_oOrdline.Route ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString

            'Dim oUpdateFrom As cImprint

            'oUpdateFrom = GetImprintRow()

            SetImprintRow(dgvItems.CurrentRow.Index, oCopyFrom) ', intCopyRouteID, strCopyRoute)

            dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value = intCopyRouteID
            dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value = strCopyRoute
            dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.ProductProof).Value = (intProductProof = 1)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function GetImprintRow() As cImprint

        GetImprintRow = New cImprint()
        GetImprintRow.Reset()

        Try

            If dgvItems.Rows.Count = 0 Then Exit Function
            If dgvItems.Rows(0).Cells(Columns.Item_No).Value.ToString = "" Then Exit Function

            With dgvItems.CurrentRow
                'GetImprintRow.Item_Guid = dgvItems.CurrentRow.Cells(Columns.Item_Guid).Value.ToString
                GetImprintRow.Item_No = .Cells(Columns.Item_No).Value.ToString
                GetImprintRow.Imprint = .Cells(Columns.Imprint).Value.ToString
                GetImprintRow.End_User = .Cells(Columns.End_User).Value.ToString
                GetImprintRow.Location = .Cells(Columns.Imprint_Location).Value.ToString
                GetImprintRow.Color = .Cells(Columns.Imprint_Color).Value.ToString
                If IsNumeric(.Cells(Columns.Num_Impr_1).Value.ToString) Then
                    GetImprintRow.Num_Impr_1 = .Cells(Columns.Num_Impr_1).Value
                Else
                    GetImprintRow.Num_Impr_1 = 0
                End If
                If IsNumeric(.Cells(Columns.Num_Impr_2).Value.ToString) Then
                    GetImprintRow.Num_Impr_2 = .Cells(Columns.Num_Impr_2).Value
                Else
                    GetImprintRow.Num_Impr_2 = 0
                End If
                If IsNumeric(.Cells(Columns.Num_Impr_3).Value.ToString) Then
                    GetImprintRow.Num_Impr_3 = .Cells(Columns.Num_Impr_3).Value
                Else
                    GetImprintRow.Num_Impr_3 = 0
                End If
                GetImprintRow.Packaging = .Cells(Columns.Packaging).Value.ToString
                GetImprintRow.Refill = .Cells(Columns.Refill).Value.ToString
                GetImprintRow.Laser_Setup = .Cells(Columns.Laser_Setup).Value.ToString
                GetImprintRow.Industry = .Cells(Columns.Industry).Value.ToString
                GetImprintRow.Comments = .Cells(Columns.Comments).Value.ToString
                GetImprintRow.Special_Comments = .Cells(Columns.Special_Comments).Value.ToString
                GetImprintRow.Printer_Comment = .Cells(Columns.Printer_Comment).Value.ToString
                GetImprintRow.Packer_Comment = .Cells(Columns.Packer_Comment).Value.ToString
                GetImprintRow.Printer_Instructions = .Cells(Columns.Printer_Instructions).Value.ToString
                GetImprintRow.Packer_Instructions = .Cells(Columns.Packer_Instructions).Value.ToString
                If IsNumeric(.Cells(Columns.Repeat_From_ID).Value.ToString) Then
                    GetImprintRow.Repeat_From_ID = .Cells(Columns.Repeat_From_ID).Value
                Else
                    GetImprintRow.Repeat_From_ID = 0
                End If

                '++ID 06.08.2020 scrabl project -------------------------------------
                GetImprintRow.Lamination_Scribl = .Cells(Columns.Lamination).Value.ToString
                '++ID 12.09.2021 Foil Color
                GetImprintRow.Foil_Color = .Cells(Columns.Foil_Color).Value.ToString

                GetImprintRow.Thread_Color_Scribl = .Cells(Columns.Thread_Color).Value.ToString
                GetImprintRow.Rounded_corners_Scribl = .Cells(Columns.Rounded_Corners).Value.ToString
                '--------------------------------------------------------------------
                '++ID 08.06.2020
                GetImprintRow.Tip_In_PageTxt = .Cells(Columns.Tip_In_PageTxt).Value.ToString
                '--------------------------------------------------------------------

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Private Sub SetImprintRow(ByVal pintRow As Integer, ByRef poSetFrom As cImprint)

        Try
            With dgvItems.Rows(pintRow)

                'If m_oOrder.OrdLines.Contains(.Cells(Columns.Item_Guid).Value.ToString) Then
                '    Dim oOrdline As cOrdLine
                '    oOrdline = m_oOrder.OrdLines(.Cells(Columns.Item_Guid).Value.ToString)
                '    If oOrdline.ImprintLine Is Nothing Then oOrdline.ImprintLine = New cImprint(m_oOrder.Ordhead.Ord_GUID, .Cells(Columns.Item_Guid).Value.ToString)

                '    oOrdline.ImprintLine.Imprint = poSetFrom.Imprint
                '    oOrdline.ImprintLine.Location = poSetFrom.Location
                '    oOrdline.ImprintLine.Color = poSetFrom.Color
                '    oOrdline.ImprintLine.Num = poSetFrom.Num
                '    oOrdline.ImprintLine.Packaging = poSetFrom.Packaging
                '    oOrdline.ImprintLine.Refill = poSetFrom.Refill
                '    oOrdline.ImprintLine.Comments = poSetFrom.Comments
                '    oOrdline.ImprintLine.Special_Comments = poSetFrom.Special_Comments

                'End If

                ' This will set the special instructions for each item. Must be put here because
                ' each item no of the copy is different, can't take from the source.
                g_oOrdline.Set_Spec_Instructions(.Cells(Columns.Item_No).Value, poSetFrom)

                .Cells(Columns.Imprint).Value = poSetFrom.Imprint
                .Cells(Columns.End_User).Value = poSetFrom.End_User
                .Cells(Columns.Imprint_Location).Value = poSetFrom.Location
                .Cells(Columns.Imprint_Color).Value = poSetFrom.Color
                .Cells(Columns.Num_Impr_1).Value = poSetFrom.Num_Impr_1
                .Cells(Columns.Num_Impr_2).Value = poSetFrom.Num_Impr_2
                .Cells(Columns.Num_Impr_3).Value = poSetFrom.Num_Impr_3
                .Cells(Columns.Packaging).Value = poSetFrom.Packaging
                .Cells(Columns.Refill).Value = poSetFrom.Refill
                .Cells(Columns.Laser_Setup).Value = poSetFrom.Laser_Setup
                .Cells(Columns.Industry).Value = poSetFrom.Industry
                .Cells(Columns.Comments).Value = poSetFrom.Comments
                .Cells(Columns.Special_Comments).Value = poSetFrom.Special_Comments
                .Cells(Columns.Printer_Comment).Value = poSetFrom.Printer_Comment
                .Cells(Columns.Packer_Comment).Value = poSetFrom.Packer_Comment
                .Cells(Columns.Printer_Instructions).Value = poSetFrom.Printer_Instructions
                .Cells(Columns.Packer_Instructions).Value = poSetFrom.Packer_Instructions
                .Cells(Columns.Repeat_From_ID).Value = poSetFrom.Repeat_From_ID

                '++ID 06.08.2020 scribl project ------------------------------------------
                .Cells(Columns.Lamination).Value = poSetFrom.Lamination_Scribl
                '++ID 12.09.2021 Foil Color
                .Cells(Columns.Foil_Color).Value = poSetFrom.Foil_Color

                .Cells(Columns.Thread_Color).Value = poSetFrom.Thread_Color_Scribl
                .Cells(Columns.Rounded_Corners).Value = poSetFrom.Rounded_corners_Scribl
                '-------------------------------------------------------------------------

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

    Private Sub dgvItems_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvItems.KeyDown
        ' If e.KeyCode = Keys.A And Keys.ControlKey 

        If e.KeyCode = Keys.Return Then
            If dgvItems.Rows.Count = 0 Then
                cmdCancel.PerformClick()
                Exit Sub
            Else
                'cmdAddToOrder.PerformClick()
                Dim oInvoke As New AddToOrderDelegate(AddressOf AddToOrder)
                dgvItems.BeginInvoke(oInvoke)
                Exit Sub
            End If

        End If

        If e.KeyCode = Keys.Escape Then
            cmdCancel.PerformClick()
            Exit Sub
        End If

        'If dgvItems.Rows.Count = 0 Then
        '    If e.KeyCode = Keys.Return Then
        '        cmdCancel.PerformClick()
        '        Exit Sub
        '    End If
        'Else
        'End If

        If dgvItems.CurrentCell Is Nothing Then Exit Sub

        Select Case dgvItems.CurrentCell.ColumnIndex

            Case Columns.Item_No

                If e.KeyCode = Keys.F8 Then

                    Call PerformGridNotesClick()

                End If

            Case Columns.Location

                If e.KeyCode = Keys.F7 Then

                    Call PerformGridSearchClick()

                ElseIf e.KeyCode = Keys.F8 Then

                    Call PerformGridNotesClick()

                End If

            Case Columns.Route

                If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                If e.Control Then

                    Dim strValue As String = dgvItems.CurrentCell.Value.ToString

                    Select Case e.KeyCode
                        Case Keys.A
                            Call SetRouteCombo(dgvItems.CurrentRow.Cells(Columns.Route), "All", strValue)

                        Case Keys.F
                            Call SetRouteCombo(dgvItems.CurrentRow.Cells(Columns.Route), "Flat", strValue)

                        Case Keys.S
                            Call SetRouteCombo(dgvItems.CurrentRow.Cells(Columns.Route), "Silk", strValue)

                        Case Keys.P
                            Call SetRouteCombo(dgvItems.CurrentRow.Cells(Columns.Route), "Pad", strValue)

                        Case Keys.T
                            Call SetRouteCombo(dgvItems.CurrentRow.Cells(Columns.Route), "Silk/Flat", strValue)

                        Case Keys.L
                            Call SetRouteCombo(dgvItems.CurrentRow.Cells(Columns.Route), "Laser/Silk", strValue)

                        Case Keys.D
                            Call SetRouteCombo(dgvItems.CurrentRow.Cells(Columns.Route), "Silk/Pad", strValue)

                        Case Keys.O
                            Call SetRouteCombo(dgvItems.CurrentRow.Cells(Columns.Route), "OS", strValue)

                        Case Keys.M
                            Call SetRouteCombo(dgvItems.CurrentRow.Cells(Columns.Route), "Misc", strValue)

                    End Select

                End If

            Case Columns.Imprint

                If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                If e.Control Then
                    Select Case e.KeyCode
                        Case Keys.N
                            dgvItems.CurrentCell.Value = "NO IMPRINT"
                        Case Keys.Space
                            dgvItems.CurrentCell.Value = " "

                    End Select
                End If

            Case Columns.End_User

                ' Ca serait cool ici de faire la liste des most used end user pour la compagnie
                ' Avec un shortcut Ctrl-1-2-3-4-5-6

                If e.Control Then
                    Select Case e.KeyCode
                        Case Keys.D1
                        Case Keys.D2
                        Case Keys.D3
                        Case Keys.D4
                        Case Keys.D5
                        Case Keys.D6
                    End Select
                End If

            Case Columns.Imprint_Location

                If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                If e.Control Then
                    Select Case e.KeyCode
                        Case Keys.C
                            dgvItems.CurrentCell.Value = "CAP"
                        Case Keys.B
                            dgvItems.CurrentCell.Value = "BARREL"
                        Case Keys.A
                            dgvItems.CurrentCell.Value = "ENLARGED AREA"
                        Case Keys.S
                            dgvItems.CurrentCell.Value = "STANDARD"
                        Case Keys.E
                            dgvItems.CurrentCell.Value = "ENLARGED"
                        Case Keys.F
                            dgvItems.CurrentCell.Value = "FRONT"
                        Case Keys.Space
                            dgvItems.CurrentCell.Value = " "
                    End Select
                End If

            Case Columns.Imprint_Color

                If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                If e.Control Then

                    Select Case e.KeyCode
                        Case Keys.B
                            dgvItems.CurrentCell.Value = "BLACK"
                        Case Keys.S
                            dgvItems.CurrentCell.Value = "SILVER"
                        Case Keys.L
                            dgvItems.CurrentCell.Value = "LASER"
                        Case Keys.W
                            dgvItems.CurrentCell.Value = "WHITE"
                        Case Keys.D
                            dgvItems.CurrentCell.Value = "DEBOSS"
                        Case Keys.X
                            dgvItems.CurrentCell.Value = "LASER/OXID"
                        Case Keys.Space
                            dgvItems.CurrentCell.Value = " "
                    End Select
                End If

            Case Columns.Packaging

                If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                If e.Control Then
                    Select Case e.KeyCode
                        Case Keys.B
                            dgvItems.CurrentCell.Value = "BULK"
                        Case Keys.C
                            dgvItems.CurrentCell.Value = "CELLO"
                        Case Keys.S
                            dgvItems.CurrentCell.Value = "STD"
                        Case Keys.P
                            dgvItems.CurrentCell.Value = "P100"
                        Case Keys.W
                            dgvItems.CurrentCell.Value = "SHRINK WRAP"
                        Case Keys.Space
                            dgvItems.CurrentCell.Value = " "
                    End Select

                End If

                'Case Else

                '    Select Case e.KeyCode
                '        Case Keys.Return
                '            Call ProductLineInsert()

                '        Case Keys.F5
                '            MsgBox("Open Notes")

                '        Case Keys.F7
                '            'MsgBox("Open search")

                '        Case Keys.F8
                '            'MsgBox("whatever...")

                '    End Select

        End Select

    End Sub

    Private Sub AddToOrder()

        cmdAddToOrder.PerformClick()

    End Sub

    Public Sub SetStatusBarPanelText(ByVal pintPanel As Integer, ByVal pstrText As String)

        Select Case pintPanel

            Case 1
                _sbOrder_Panel1.Text = pstrText

                'Case 2
                '    _sbOrder_Panel2.Text = pstrText

                'Case 3
                '    _sbOrder_Panel3.Text = pstrText

                'Case 4
                '    _sbOrder_Panel4.Text = pstrText

            Case Else

        End Select

    End Sub

    Private Sub dgvItems_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.CellEnter

        Call SetColumnShortcuts()

        If e.ColumnIndex = Columns.Route Then

            'Dim oitem As cOrdLine
            If dgvItems.Rows(e.RowIndex).Cells(Columns.Qty_Ordered).Value > dgvItems.Rows(e.RowIndex).Cells(Columns.Qty_On_Hand).Value Then
                '    oitem = m_oOrder.OrdLines(dgvItems.Rows(e.RowIndex).Cells(Columns.Kit_Item_Guid).Value.ToString)
                'Else
                '    oitem = m_oOrder.OrdLines(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_Guid).Value.ToString)
                'End If
                'If oitem.Qty_Prev_Bkord > 0 Then
                If dgvItems.CurrentRow.Cells(Columns.Location).Value <> "3N" Then
                    MsgBox("This item is no stock.")
                End If
            End If

            Call SetRouteCombo(dgvItems.Rows(e.RowIndex).Cells(Columns.Route), "All", dgvItems.Rows(e.RowIndex).Cells(Columns.Route).Value.ToString)

        End If

        If e.ColumnIndex = Columns.Unit_Price And dgvItems.CurrentRow.Cells(Columns.Kit_Feat_Fg).Value.ToString = "K" Then

            Dim oKit As New frmKitSelection
            Dim sKit As New cKit(m_oOrder.Ordhead.Ord_GUID)

            sKit.Item_No = dgvItems.CurrentRow.Cells(Columns.Item_No).Value.ToString
            sKit.Item_Desc_1 = dgvItems.CurrentRow.Cells(Columns.Item_Desc_1).Value.ToString
            sKit.Item_Desc_2 = dgvItems.CurrentRow.Cells(Columns.Item_Desc_2).Value.ToString
            sKit.Item_Price = dgvItems.CurrentRow.Cells(Columns.Unit_Price).Value.ToString
            sKit.Loc = dgvItems.CurrentRow.Cells(Columns.Location).Value.ToString

            sKit.Item_Price = dgvItems.CurrentRow.Cells(Columns.Unit_Price).Value.ToString
            sKit.load()
            oKit.Kit = sKit

            oKit.ShowDialog()
            oKit.Dispose()

        End If

    End Sub

    Private Sub PerformGridSearchClick()

        Try

            'If TypeOf eventSender Is Button Then

            '    btnElement = New Button
            '    btnElement = DirectCast(eventSender, Button)

            Dim oSearch As New cSearch(GetSearchElementByColumnIndex(dgvItems.CurrentCell.ColumnIndex))

            If oSearch.Form.FoundRow <> -1 Then
                dgvItems.CurrentCell.Value = oSearch.Form.FoundElementValue

            End If
            oSearch = Nothing

            'End If
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function GetSearchElementByColumnIndex(ByVal pintColumn As Integer) As String

        GetSearchElementByColumnIndex = ""

        Try
            Select Case pintColumn ' dgvItems.CurrentCell.ColumnIndex

                Case Columns.Location
                    GetSearchElementByColumnIndex = "IM-LOCATION"

                Case Else
                    GetSearchElementByColumnIndex = ""

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Private Sub PerformGridNotesClick()

        Try
            If dgvItems.CurrentCell.Equals(DBNull.Value) Then Exit Sub
            If dgvItems.CurrentCell.Value = "" Then Exit Sub

            Dim oNotes As New cNotes(GetNoteElementByColumnIndex(dgvItems.CurrentCell.ColumnIndex), dgvItems.CurrentCell.Value)

            oNotes = Nothing

            'End If
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function GetNoteElementByColumnIndex(ByVal pintColumn As Integer) As String

        GetNoteElementByColumnIndex = ""

        Try
            Select Case pintColumn ' dgvItems.CurrentCell.ColumnIndex

                Case Columns.Item_No
                    GetNoteElementByColumnIndex = "IM-ITEM"

                Case Columns.Location
                    GetNoteElementByColumnIndex = "IM-LOCATION"

                Case Else
                    GetNoteElementByColumnIndex = ""

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Private Sub cmdGetRepeatData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetRepeatData.Click

        If Trim(txtRepeat_Ord_No.Text) <> "" Then

            Dim oFrmRepeatImprint As New frmRepeatImprint(Trim(txtRepeat_Ord_No.Text))
            oFrmRepeatImprint.ShowDialog()

            g_oOrdline.Set_Spec_Instructions(txtImprint_Item_No.Text, oFrmRepeatImprint.Imprint)

            txtImprint.Text = oFrmRepeatImprint.Imprint.Imprint
            txtNum_Impr_1.Text = oFrmRepeatImprint.Imprint.Num_Impr_1
            txtNum_Impr_2.Text = oFrmRepeatImprint.Imprint.Num_Impr_2
            txtNum_Impr_3.Text = oFrmRepeatImprint.Imprint.Num_Impr_3
            txtComments.Text = oFrmRepeatImprint.Imprint.Comments
            txtSpecial_Comments.Text = oFrmRepeatImprint.Imprint.Special_Comments
            txtEnd_User.Text = oFrmRepeatImprint.Imprint.End_User
            txtImprint_Location.Text = oFrmRepeatImprint.Imprint.Location
            txtImprint_Color.Text = oFrmRepeatImprint.Imprint.Color
            txtPackaging.Text = oFrmRepeatImprint.Imprint.Packaging
            txtRefill.Text = oFrmRepeatImprint.Imprint.Refill
            txtIndustry.Text = oFrmRepeatImprint.Imprint.Industry

            With dgvItems.CurrentRow
                .Cells(Columns.Imprint).Value = oFrmRepeatImprint.Imprint.Imprint
                .Cells(Columns.Num_Impr_1).Value = oFrmRepeatImprint.Imprint.Num_Impr_1
                .Cells(Columns.Num_Impr_2).Value = oFrmRepeatImprint.Imprint.Num_Impr_2
                .Cells(Columns.Num_Impr_3).Value = oFrmRepeatImprint.Imprint.Num_Impr_3
                .Cells(Columns.Comments).Value = oFrmRepeatImprint.Imprint.Comments
                .Cells(Columns.Special_Comments).Value = oFrmRepeatImprint.Imprint.Special_Comments
                .Cells(Columns.End_User).Value = oFrmRepeatImprint.Imprint.End_User
                .Cells(Columns.Imprint_Location).Value = oFrmRepeatImprint.Imprint.Location
                .Cells(Columns.Imprint_Color).Value = oFrmRepeatImprint.Imprint.Color
                .Cells(Columns.Packaging).Value = oFrmRepeatImprint.Imprint.Packaging
                .Cells(Columns.Refill).Value = oFrmRepeatImprint.Imprint.Refill
                .Cells(Columns.Laser_Setup).Value = oFrmRepeatImprint.Imprint.Laser_Setup
                .Cells(Columns.Industry).Value = oFrmRepeatImprint.Imprint.Industry
                .Cells(Columns.Printer_Comment).Value = oFrmRepeatImprint.Imprint.Printer_Comment
                .Cells(Columns.Packer_Comment).Value = oFrmRepeatImprint.Imprint.Packer_Comment
                .Cells(Columns.Printer_Instructions).Value = oFrmRepeatImprint.Imprint.Printer_Instructions
                .Cells(Columns.Packer_Instructions).Value = oFrmRepeatImprint.Imprint.Packer_Instructions
                .Cells(Columns.Repeat_From_ID).Value = oFrmRepeatImprint.Imprint.Repeat_From_ID

                '++ID 06.08.2020 scribl project ------------------------------------------------------
                .Cells(Columns.Lamination).Value = oFrmRepeatImprint.Imprint.Lamination_Scribl
                '++ID 12.09.2021 Foil Color
                .Cells(Columns.Foil_Color).Value = oFrmRepeatImprint.Imprint.Foil_Color

                .Cells(Columns.Thread_Color).Value = oFrmRepeatImprint.Imprint.Thread_Color_Scribl
                .Cells(Columns.Rounded_Corners).Value = oFrmRepeatImprint.Imprint.Rounded_corners_Scribl
                '-------------------------------------------------------------------------------------

            End With

            Call ImprintSave()

            oFrmRepeatImprint.Dispose()

        End If

    End Sub

    Private Sub cboIndustry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboIndustry.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            m_oOrdline.ImprintLine.Industry = cboIndustry.SelectedItem.ToString ' txtIndustry.Text)
            txtIndustry.Text = m_oOrdline.ImprintLine.Industry
            dgvItems.CurrentRow.Cells(Columns.Industry).Value = txtIndustry.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvItems_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvItems.CellValueChanged
        Try



        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub
End Class

Public Enum ProductLineEntryType

    Search
    Substitutes

End Enum

