Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms.Panel
Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.DirectoryServices
Imports System.Collections.Generic

' For Mysql import:
'Imports System.Data
'Imports MySql.Data
'Imports MySql.Data.MySqlClient
'Imports MySql.Web

'end mysql import

Friend Class frmOrder

    Inherits System.Windows.Forms.Form
    Public iCompteur As Long = 0
    Public dgvCurrentRow As DataGridViewRow
    'Public g_oOrdline As cOrdLine

    Public FoundElement As Object
    Public SearchElement As String

    Private m_MinHeight As Integer
    Private m_MinWidth As Integer

    Private m_lLabels1Width As Integer
    Private m_lLabels2Width As Integer
    Private m_lText1Width As Integer
    Private m_lText2Width As Integer

    Private m_strSqlSelect As String
    Private m_strSqlFields As String
    Private m_strSqlFrom As String
    Private m_strSqlInnerJoin As String
    Private m_strSqlDefaultWhere As String
    Private m_strSqlSearchWhere As String
    Private m_strSqlOrderBy As String

    Private m_lElementCount As Integer
    Private m_lOrderByCount As Integer

    'Private m_oOrder As New cOrder

    Private blnLoading As Boolean = False
    Private blnDeleting As Boolean = False
    Private blnInserting As Boolean = False

    'Public dsDataSet As New DataSet()
    Public dtDataTable As New DataTable

    ' blnFill is used to prevent events from happening in some dgvItems subs
    Private blnFill As Boolean = False

    'ignoreQtyMismatch used to ignore when qty ordered exceeds stock levels (used for assorted color pricing adjustments)
    Public Shared ignoreQtyMismatch As Boolean = False

    'Prop 65 variable checker
    Public Shared shipToProp65State As Boolean = False

    Public Shared ignoreImprintMsgCheck As Boolean = False

    Private oCopyFrom As cImprint
    Private intCopyRouteID As Integer 'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
    Private strCopyRoute As String ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
    Private intCopyProductProof As Integer
    Private intCopyAutoCompleteReship As Integer

    Private strWebItem As String

    Private mblnNoMsgBox As Boolean = False

    Delegate Sub MyDelegate()
    Delegate Sub SetColumnIndex(ByVal i As Integer)
    'Delegate Sub ProductLineInsertDelegate()
    'Delegate Sub ProductLineDeleteDelegate()
    Delegate Sub SetQty_To_Ship()
    Delegate Sub Qty_On_Hand_Kits_Delegate(ByVal iRow As Integer) ' , ByVal iCol As Integer)
    Delegate Sub Del_SetProductionAvailabilityColor(ByRef oCell As DataGridViewCell) ' , ByVal iCol As Integer)
    Delegate Sub Del_Set_Num_Imprint_To_Grid(ByVal iCurrentRowIndex As Integer, ByRef oOrdline As cOrdLine)

    Delegate Sub Del_PerformProductionAvailalbilitySimulation(ByRef oCell As DataGridViewCell, ByRef poImprint As cImprint, ByVal checkMultiDays As Boolean, ByVal OsrEligibiltyList As Dictionary(Of String, Boolean), ByVal kitUnits As Integer) 'delegate notixia (thermo) availalbility day check

    Private WithEvents TestWorker As System.ComponentModel.BackgroundWorker

    Private validRouteList As ArrayList = Nothing

    Dim oDB As New cDBA

    Public Enum Columns
        FirstCol = 0
        Line_Seq_No
        Image
        Item_No
        Location
        Qty_Ordered
        Qty_Overpick
        Qty_To_Ship
        Qty_Inventory
        Qty_Allocated
        Qty_On_Hand
        Qty_Prev_To_Ship
        Qty_Prev_BkOrd ' Add to properties
        Extra_9 ' ETA Date
        Extra_10 ' ETA Quantity
        Unit_Price ' ?? not in Ordline ??
        Discount_Pct
        Calc_Price
        Request_Dt
        Promise_Dt
        Req_Ship_Dt
        Item_Desc_1
        Item_Desc_2
        Extra_1 ' Exact_Qty
        Imprint
        End_User
        Imprint_Location
        Imprint_Color
        Num_Impr_1
        Num_Impr_2
        Num_Impr_3
        Packaging
        Refill
        Laser_Setup
        Industry
        Comments
        Special_Comments
        Printer_Comment
        Packer_Comment
        Printer_Instructions
        Packer_Instructions
        Repeat_From_ID
        Route
        RouteID
        ProductProof
        AutoCompleteReship ' NEW FIELD
        UOM
        Tax_Sched
        Revision_No
        BkOrd_Fg ' BackorderItem
        Recalc_SW ' Recalc
        Item_Guid
        Kit_Item_Guid ' Kit_Item_Guid is used to get the main product guid of an item in a kit
        PictureHeight
        Item_Cd
        Mix_Group
        LastCol
    End Enum

    Private Enum History
        FirstCol
        Ord_Type
        OEI_Ord_No
        Ord_No
        Ord_Dt
        Cus_No
        Bill_To_Name
        OE_PO_No
        Pending
        FileCreated
        FileExported
        TravelerCreated
        LastCol
    End Enum

    Private Enum ReservedItems
        FirstCol
        OEI_Ord_No
        Ord_Dt
        Item_No
        Qty_Ordered
        LastCol
    End Enum

    Public Enum Tabs
        Header
        CustomerContact
        Lines
        Documents
        HeaderAll
        Taxes
        Salesperson
        CreditInfo
        Extra
        History
        ReservedItems

    End Enum


    Private Sub cmdClose_Click()

        Me.Close()

    End Sub

    Private Sub frmOrder_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

        If Not (m_oOrder.Ordhead Is Nothing) Then
            m_oOrder.Ordhead.UnsetOpenTS()
            m_oOrder.DeleteEmptyOrderIfNew()
        End If

    End Sub

    Private Sub frmOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        ' If Ctl-Number, sets the new tab.
        If e.Control Then
            Select Case e.KeyCode
                Case Keys.D1
                    oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Header)
                Case Keys.D2
                    oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CustomerContact)
                Case Keys.D3
                    oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Lines)
                Case Keys.D4
                    oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Documents)
                Case Keys.D5
                    oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.HeaderAll)
                Case Keys.D6
                    oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Taxes)
                Case Keys.D7
                    oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
                Case Keys.D8
                    oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
                Case Keys.D9
                    oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
            End Select

            Call sstOrder_Click(sstOrder, New System.EventArgs)

        End If

        If e.Alt Then
            Select Case e.KeyCode

                ' EXPORT TO EXCEL
                Case Keys.X
                    tsbExportToMacola.PerformClick()

                    ' NEW ORDER
                Case Keys.N
                    tsbNew.PerformClick()

            End Select

        End If

    End Sub

    Private Sub frmOrder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            CreateRouteCollections()

            'm_oOrder = New cOrder
            tsbXmlExport.Visible = True ' False ' g_User.XML_Access
            Call LoadProductGrid()
            Call LoadHistoryGrid()
            Call LoadReservedItemsGrid()
            'm_oOrder.OrderLines = UcOrder.OrderLines

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#Region "    OrderLines ########################################## "

    Public ReadOnly Property Items() As DataGridView
        Get
            Items = dgvItems
        End Get
        'Set(ByVal value As DataGridView)

        'End Set
    End Property

    'Public ReadOnly Property Order() As cOrder
    '    Get
    '        Order = m_oOrder
    '    End Get
    'End Property

    Public Sub LoadHistoryGrid()

        Try

            With dgvHistory.Columns
                .Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
                .Add(DGVTextBoxColumn("Ord_Type", "Ord Type", 40))
                .Add(DGVTextBoxColumn("OEI_Ord_No", "OEI No", 80))
                .Add(DGVTextBoxColumn("Ord_No", "Macola No", 70))
                .Add(DGVCalendarColumn("Ord_Dt", "Date", 70))
                .Add(DGVTextBoxColumn("Cus_No", "Customer", 80))
                .Add(DGVTextBoxColumn("Bill_To_Name", "Bill To", 250))
                .Add(DGVTextBoxColumn("OE_PO_No", "P/O", 100))
                .Add(DGVCheckBoxColumn("Pending", "Pending", 60))
                .Add(DGVCheckBoxColumn("FileCreated", "File Created", 60))
                .Add(DGVCheckBoxColumn("FileExported", "File Exported", 60))
                .Add(DGVCheckBoxColumn("TravelerCreated", "Traveler Created", 60))
                .Add(DGVTextBoxColumn("LastCol", "LastCol", 0))

            End With

            dgvHistory.Columns(History.FirstCol).Visible = False
            dgvHistory.Columns(History.LastCol).Visible = False

            'Call LoadHistoryGridValues()

            Dim oCellStyle As New DataGridViewCellStyle()

            oCellStyle = New DataGridViewCellStyle()
            oCellStyle.Format = "mm/dd/yyyy"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvHistory.Columns(History.Ord_Dt).DefaultCellStyle = oCellStyle

            dgvHistory.Columns(History.FirstCol).Frozen = True
            dgvHistory.Columns(History.OEI_Ord_No).Frozen = True
            dgvHistory.Columns(History.Ord_No).Frozen = True

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub LoadHistoryGridValues()

        'Dim strSql As String = "" & _
        '"SELECT		'' AS FirstCol, Ord_Type, Oei_Ord_No, Ord_No, " & _
        '"           Ord_Dt, Cus_No, Bill_To_Name, OE_PO_No, " & _
        '"	 		CASE WHEN ExportTS IS NULL THEN 1 ELSE 0 END AS Pending, " & _
        '"	 		CASE WHEN ExportTS IS NULL THEN 0 ELSE 1 END AS FileCreated, " & _
        '"			CASE WHEN MacolaTS IS NULL THEN 0 ELSE 1 END AS FileExported, " & _
        '"			CASE WHEN TriggerTS IS NULL THEN 0 ELSE 1 END AS TravelerCreated, '' AS LastCol " & _
        '"FROM 		OEI_Ordhdr WITH (Nolock) " & _
        '"WHERE 		UserID = '" & Trim(g_User.Usr_ID) & "' " & _
        '"ORDER BY 	OEI_Ord_No DESC "

        Try
            Dim strSql As String

            strSql = _
            "SELECT		TOP 500 '' AS FirstCol, cast(O.User_Def_Fld_2 AS CHAR(3)) AS Ord_Type, O.Oei_Ord_No, O.Ord_No, " & _
            "           O.Ord_Dt, O.Cus_No, O.Bill_To_Name, O.OE_PO_No, " & _
            "	 		CASE WHEN O.ExportTS IS NULL THEN 1 ELSE 0 END AS Pending, " & _
            "	 		CASE WHEN O.ExportTS IS NULL THEN 0 ELSE 1 END AS FileCreated, " & _
            "			CASE WHEN O.MacolaTS IS NULL THEN 0 ELSE 1 END AS FileExported, " & _
            "			CASE WHEN O.TriggerTS IS NULL THEN 0 WHEN COUNT(M.Ord_No) = 0 THEN 1 WHEN ISNULL(M.Inv_No, '') <> '' THEN 1 WHEN M.[STATUS] = 'L' THEN 1 WHEN COUNT(H.HeaderID) = 0 THEN 0 ELSE 1 END AS TravelerCreated, " & _
            "           '' AS LastCol " & _
            "FROM 		OEI_Ordhdr O WITH (Nolock) " & _
            "LEFT JOIN	OEOrdHdr_Sql M WITH (Nolock) ON O.Ord_Type = M.Ord_Type AND O.Ord_No = M.Ord_No " & _
            "LEFT JOIN	EXACT_TRAVELER_HEADER H WITH (Nolock) ON O.Ord_No = H.Ord_No AND O.Cus_No = H.Cus_No " & _
            "WHERE 		O.UserID = '" & Trim(g_User.Usr_ID) & "' " & _
            "GROUP BY   O.User_Def_Fld_2, O.Oei_Ord_No, O.Ord_No, O.Ord_Dt, O.Cus_No, O.Bill_To_Name, O.OE_PO_No, " & _
            "           O.ExportTS, O.MacolaTS, O.TriggerTS, M.Inv_No, M.[STATUS] " & _
            "ORDER BY 	CASE WHEN O.ExportTS IS NULL THEN 1 ELSE 0 END DESC, " & _
            "			CASE WHEN O.TriggerTS IS NULL THEN 0 WHEN COUNT(M.Ord_No) = 0 THEN 1 WHEN ISNULL(M.Inv_No, '') <> '' THEN 1 WHEN M.[STATUS] = 'L' THEN 1 WHEN COUNT(H.HeaderID) = 0 THEN 0 ELSE 1 END, " & _
            "           O.OEI_Ord_No DESC "

            'strSql = _
            '"SELECT		TOP 500 '' AS FirstCol, cast(O.User_Def_Fld_2 AS CHAR(3)) AS Ord_Type, O.Oei_Ord_No, O.Ord_No, " & _
            '"           O.Ord_Dt, O.Cus_No, O.Bill_To_Name, O.OE_PO_No, " & _
            '"	 		CASE WHEN O.ExportTS IS NULL THEN 1 ELSE 0 END AS Pending, " & _
            '"	 		CASE WHEN O.ExportTS IS NULL THEN 0 ELSE 1 END AS FileCreated, " & _
            '"			CASE WHEN O.MacolaTS IS NULL THEN 0 ELSE 1 END AS FileExported, " & _
            '"			CASE WHEN O.TriggerTS IS NULL THEN 0 WHEN COUNT(M.Ord_No) = 0 THEN 1 WHEN ISNULL(M.Inv_No, '') <> '' THEN 1 WHEN M.[STATUS] = 'L' THEN 1 WHEN COUNT(H.HeaderID) = 0 THEN 0 ELSE 1 END AS TravelerCreated, " & _
            '"           '' AS LastCol " & _
            '"FROM 		OEI_Ordhdr O WITH (Nolock) " & _
            '"LEFT JOIN	OEOrdHdr_Sql M WITH (Nolock) ON O.Ord_Type = M.Ord_Type AND O.Ord_No = M.Ord_No " & _
            '"LEFT JOIN	EXACT_TRAVELER_HEADER H WITH (Nolock) ON O.Ord_No = H.Ord_No AND O.Cus_No = H.Cus_No " & _
            '"WHERE 		O.UserID = 'LENI' " & _
            '"GROUP BY   O.User_Def_Fld_2, O.Oei_Ord_No, O.Ord_No, O.Ord_Dt, O.Cus_No, O.Bill_To_Name, O.OE_PO_No, " & _
            '"           O.ExportTS, O.MacolaTS, O.TriggerTS, M.Inv_No, M.[STATUS] " & _
            '"ORDER BY 	CASE WHEN O.ExportTS IS NULL THEN 1 ELSE 0 END DESC, " & _
            '"			CASE WHEN O.TriggerTS IS NULL THEN 0 WHEN COUNT(M.Ord_No) = 0 THEN 1 WHEN ISNULL(M.Inv_No, '') <> '' THEN 1 WHEN M.[STATUS] = 'L' THEN 1 WHEN COUNT(H.HeaderID) = 0 THEN 0 ELSE 1 END, " & _
            '"           O.OEI_Ord_No DESC "

            'strSql = "" & _
            '"SELECT		'' AS FirstCol, cast(O.User_Def_Fld_2 AS CHAR(3)) AS Ord_Type, O.Oei_Ord_No, O.Ord_No, " & _
            '"           O.Ord_Dt, O.Cus_No, O.Bill_To_Name, O.OE_PO_No, " & _
            '"	 		CASE WHEN O.ExportTS IS NULL THEN 1 ELSE 0 END AS Pending, " & _
            '"	 		CASE WHEN O.ExportTS IS NULL THEN 0 ELSE 1 END AS FileCreated, " & _
            '"			CASE WHEN O.MacolaTS IS NULL THEN 0 ELSE 1 END AS FileExported, " & _
            '"			CASE WHEN O.TriggerTS IS NULL THEN 0 WHEN COUNT(M.Ord_No) = 0 THEN 1 WHEN ISNULL(M.Inv_No, '') <> '' THEN 1 WHEN COUNT(H.HeaderID) = 0 THEN 0 ELSE 1 END AS TravelerCreated, " & _
            '"           '' AS LastCol " & _
            '"FROM 		OEI_Ordhdr O WITH (Nolock) " & _
            '"LEFT JOIN	OEOrdHdr_Sql M WITH (Nolock) ON O.Ord_Type = M.Ord_Type AND O.Ord_No = M.Ord_No " & _
            '"LEFT JOIN	EXACT_TRAVELER_HEADER H WITH (Nolock) ON O.Ord_No = H.Ord_No AND O.Cus_No = H.Cus_No " & _
            '"WHERE 		O.UserID = 'LORIE' " & _
            '"GROUP BY   O.User_Def_Fld_2, O.Oei_Ord_No, O.Ord_No, O.Ord_Dt, O.Cus_No, O.Bill_To_Name, O.OE_PO_No, " & _
            '"           O.ExportTS, O.MacolaTS, O.TriggerTS, M.Inv_No " & _
            '"ORDER BY 	CASE WHEN O.ExportTS IS NULL THEN 1 ELSE 0 END DESC, " & _
            '"			CASE WHEN O.TriggerTS IS NULL THEN 0 WHEN COUNT(M.Ord_No) = 0 THEN 1 WHEN ISNULL(M.Inv_No, '') <> '' THEN 1 WHEN COUNT(H.HeaderID) = 0 THEN 0 ELSE 1 END, " & _
            '"           O.OEI_Ord_No DESC "

            Dim dt As DataTable
            Dim db As New cDBA

            dt = db.DataTable(strSql)
            dgvHistory.DataSource = dt

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub LoadReservedItemsGrid()

        Try

            With dgvReservedItems.Columns
                .Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
                .Add(DGVTextBoxColumn("OEI_Ord_No", "OEI No", 100))
                .Add(DGVCalendarColumn("Ord_Dt", "Date", 100))
                .Add(DGVTextBoxColumn("Item_No", "Item No", 150))
                .Add(DGVTextBoxColumn("Qty_Ordered", "Qty Ordered", 100))
                .Add(DGVTextBoxColumn("LastCol", "LastCol", 0))

            End With

            dgvReservedItems.Columns(ReservedItems.FirstCol).Visible = False
            dgvReservedItems.Columns(ReservedItems.LastCol).Visible = False

            'Call LoadReservedItemsValues()

            Dim oCellStyle As New DataGridViewCellStyle()

            oCellStyle = New DataGridViewCellStyle()
            oCellStyle.Format = "mm/dd/yyyy"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvReservedItems.Columns(ReservedItems.Ord_Dt).DefaultCellStyle = oCellStyle

            dgvReservedItems.Columns(ReservedItems.FirstCol).Frozen = True
            dgvReservedItems.Columns(ReservedItems.OEI_Ord_No).Frozen = True

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub LoadReservedItemsValues()

        'Dim strSql As String = "" & _
        '"SELECT		'' AS FirstCol, Ord_Type, Oei_Ord_No, Ord_No, " & _
        '"           Ord_Dt, Cus_No, Bill_To_Name, OE_PO_No, " & _
        '"	 		CASE WHEN ExportTS IS NULL THEN 1 ELSE 0 END AS Pending, " & _
        '"	 		CASE WHEN ExportTS IS NULL THEN 0 ELSE 1 END AS FileCreated, " & _
        '"			CASE WHEN MacolaTS IS NULL THEN 0 ELSE 1 END AS FileExported, " & _
        '"			CASE WHEN TriggerTS IS NULL THEN 0 ELSE 1 END AS TravelerCreated, '' AS LastCol " & _
        '"FROM 		OEI_Ordhdr WITH (Nolock) " & _
        '"WHERE 		UserID = '" & Trim(g_User.Usr_ID) & "' " & _
        '"ORDER BY 	OEI_Ord_No DESC "

        Try
            Dim strSql As String = "" & _
            "SELECT		'' AS FirstCol, O.Oei_Ord_No, O.Ord_Dt, L.Item_No, " & _
            "           L.Qty_Ordered, '' AS LastCol " & _
            "FROM 		OEI_Ordhdr O WITH (Nolock) " & _
            "INNER JOIN OEI_OrdLin L WITH (Nolock) ON O.Ord_Guid = L.Ord_Guid " & _
            "WHERE 		ISNULL(O.Ord_No, '') = '' " & _
            "ORDER BY 	L.Item_No, L.Qty_Ordered DESC "

            Dim dt As DataTable
            Dim db As New cDBA

            dt = db.DataTable(strSql)
            dgvReservedItems.DataSource = dt

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub LoadProductGrid()

        'Dim rsProducts As New ADODB.Recordset
        Dim strSql As String

        Try
            blnLoading = True

            dgvItems.CausesValidation = False

            dgvItems.Columns.Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("Line_Seq_No", "Line", 0))
            dgvItems.Columns.Add(DGVImageColumn("Image", "Image", 125))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_No", "Item", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Location", "Loc", 40))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_Ordered", "Qty Ordered", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_Overpick", "Qty Overpick", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Qty_To_Ship", "Qty Shipped", 70))
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
            'dgvItems.Columns.Add(DGVCalendarColumn("Req_Ship_Dt", "Req Ship", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Desc_1", "Description", 200))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Desc_2", "Item Description Line 2", 200))

            dgvItems.Columns.Add(DGVCheckBoxColumn("Extra_1", "Exact Qty", 50))

            dgvItems.Columns.Add(DGVTextBoxColumn("Imprint", "Imprint", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("End_User", "End_User", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Imprint_Location", "Imprint Location", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Imprint_Colour", "Imprint Colour", 75))
            dgvItems.Columns.Add(DGVTextBoxColumn("Num_Impr_1", "Num Impr 1", 75))
            dgvItems.Columns.Add(DGVTextBoxColumn("Num_Impr_2", "Num Impr 2", 75))
            dgvItems.Columns.Add(DGVTextBoxColumn("Num_Impr_3", "Num Impr 3", 75))
            dgvItems.Columns.Add(DGVTextBoxColumn("Packaging", "Packaging", 75))
            dgvItems.Columns.Add(DGVTextBoxColumn("Refill", "Refill", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Laser_Setup", "Laser Setup", 100))

            Dim comboboxColumn As DataGridViewComboBoxColumn
            'dgvItems.Columns.Add(DGVTextBoxColumn("Industry", "Industry", 100))
            comboboxColumn = DGVComboBoxColumn("Industry", "Industry", 200)
            With comboboxColumn
                .DataSource = m_oOrder.GetCompleteCboIndustryList()
                .ValueMember = "Industry"
                .DisplayMember = "Industry"
            End With
            dgvItems.Columns.Add(comboboxColumn)

            dgvItems.Columns.Add(DGVTextBoxColumn("Comments", "Comments", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Special_Comments", "Spec. Comments", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Printer_Comment", "Printer Comment", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Packer_Comment", "Packer Comment", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Printer_Instructions", "Printer Instructions", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Packer_Instructions", "Packer Instructions", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Repeat_From_ID", "Repeat_From_ID", 0))

            comboboxColumn = DGVComboBoxColumn("Route", "Route", 200)
            With comboboxColumn
                .DataSource = m_oOrder.GetCompleteCboRouteList()
                .ValueMember = "Route"
                .DisplayMember = "Route"
            End With
            dgvItems.Columns.Add(comboboxColumn)

            dgvItems.Columns.Add(DGVTextBoxColumn("RouteID", "RouteID", 0))

            dgvItems.Columns.Add(DGVCheckBoxColumn("ProductProof", "Product Proof", 70))
            dgvItems.Columns.Add(DGVCheckBoxColumn("AutoCompleteReship", "Auto Complete Reship", 100))
            'AutoCompleteReship
            dgvItems.Columns.Add(DGVTextBoxColumn("UOM", "UOM", 70))

            dgvItems.Columns.Add(DGVTextBoxColumn("Tax_Sched", "Tax Schd", 70))
            dgvItems.Columns.Add(DGVTextBoxColumn("Revision_No", "ECM Revision", 70))
            dgvItems.Columns.Add(DGVCheckBoxColumn("BkOrd_Fg", "Backorder Item", 70))
            dgvItems.Columns.Add(DGVCheckBoxColumn("Recalc_SW", "Recalc", 70))

            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Guid", "Item_Guid", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("Kit_Item_Guid", "Kit_Item_Guid", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("PictureHeight", "PictureHeight", 0))

            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Cd", "Item Style", 100)) 'add item style for price adjustments
            dgvItems.Columns.Add(DGVTextBoxColumn("Mix_Group", "Mix Group", 100)) 'add item group

            dgvItems.Columns.Add(DGVTextBoxColumn("LastCol", "LastCol", 0))

            dgvItems.Columns(Columns.FirstCol).Visible = False
            dgvItems.Columns(Columns.Line_Seq_No).Visible = False
            dgvItems.Columns(Columns.Qty_Overpick).Visible = False
            'dgvItems.Columns(Columns.Laser_Setup).Visible = False
            dgvItems.Columns(Columns.Item_Guid).Visible = False
            dgvItems.Columns(Columns.Kit_Item_Guid).Visible = False
            dgvItems.Columns(Columns.Extra_1).Visible = False

            dgvItems.Columns(Columns.Qty_Prev_To_Ship).Visible = False

            dgvItems.Columns(Columns.Printer_Comment).Visible = False
            'dgvItems.Columns(Columns.Printer_Instructions).Visible = False
            dgvItems.Columns(Columns.Packer_Comment).Visible = False
            dgvItems.Columns(Columns.AutoCompleteReship).Visible = False
            'dgvItems.Columns(Columns.Packer_Instructions).Visible = False
            dgvItems.Columns(Columns.Repeat_From_ID).Visible = False
            dgvItems.Columns(Columns.RouteID).Visible = False
            dgvItems.Columns(Columns.PictureHeight).Visible = False
            dgvItems.Columns(Columns.Item_Cd).Visible = False
            dgvItems.Columns(Columns.Mix_Group).Visible = True
            dgvItems.Columns(Columns.LastCol).Visible = False

            'strSql = "" & _
            '"SELECT " & _
            '"'' AS 'Line Seq No', '' as 'Item No', '' as Location, CAST(0 AS FLOAT) as 'Qty Ordered', " & _
            '"CAST(0 AS FLOAT) as Qty_On_Hand, CAST(0 AS FLOAT) as 'Qty Shipped', CAST(0 AS FLOAT) as 'Qty BO', " & _
            '"CAST(0 AS FLOAT) AS 'Unit price', CAST(0 AS FLOAT) AS 'Disc Pct', CAST(0 AS FLOAT) AS 'Ext Price', " & _
            '"GETDATE() AS Request_Dt, GETDATE() AS Promise_Dt, " & _
            '"GETDATE() AS Req_Ship_Dt, '' AS 'Description 1', '' AS 'Description 2', " & _
            '"'' AS 'UOM', '' AS 'Tax Schd', '' AS 'ECM Revision', '' AS 'Create/Modify Kit', " & _
            '"'' AS 'Note/Text', '' AS 'Backorder Item', '' AS 'Selected', '' AS 'Recalc', '' AS 'Line Comments', '' AS Route "

            'strSql = "" & _
            '"SELECT TOP 1 " & _
            '"'' AS FirstCol, '' AS Line_Seq_No, C.Empty_Picture AS Image, '' as Item_No, '' as Location, CAST(0 AS FLOAT) as Qty_Ordered, CAST(0 AS FLOAT) as Qty_To_Ship, " & _
            '"CAST(0 AS FLOAT) as Qty_On_Hand, CAST(0 AS FLOAT) as Qty_Prev_To_Ship, CAST(0 AS FLOAT) as Qty_Prev_BkOrd, " & _
            '"CAST(0 AS FLOAT) AS Unit_Price, CAST(0 AS FLOAT) AS Discount_Pct, CAST(0 AS FLOAT) AS Calc_Price, " & _
            '"GETDATE() AS Request_Dt, GETDATE() AS Promise_Dt, GETDATE() AS Req_Ship_Dt, " & _
            '"'' AS Item_Desc_1, '' AS Item_Desc_2, '' AS UOM, '' AS Tax_Sched, '' AS Revision_No, " & _
            '"CAST(0 AS BIT) AS BkOrd_Fg, CAST(0 AS BIT) AS Recalc_SW, 0 AS RouteID, '' AS Route, " & _
            '"'' AS Imprint, '' AS End_User, '' AS Imprint_Location, '' AS Imprint_Colour, '' AS Num_Imprints, " & _
            '"'' AS Packaging, '' AS Refill, '' AS Laser_Setup, '' AS Industry, '' AS Comments, '' AS Special_Comments," & _
            '"CAST('' AS CHAR(30)) AS Item_Guid, CAST('' AS CHAR(30)) AS Kit_Item_Guid, '' AS LastCol " & _
            '"FROM OEI_Config C WITH (NoLock) WHERE ID = 1 "

            strSql = "" & _
            "SELECT TOP 1 " & _
            "'' AS FirstCol, '' AS Line_Seq_No, C.Empty_Picture AS Image, '' as Item_No, '' as Location, CAST(0 AS FLOAT) as Qty_Ordered, CAST(0 AS FLOAT) as Qty_Overpick, CAST(0 AS FLOAT) as Qty_To_Ship, " & _
            "CAST(0 AS FLOAT) as Qty_Inventory, CAST(0 AS FLOAT) as Qty_Allocated, CAST(0 AS FLOAT) as Qty_On_Hand, CAST(0 AS FLOAT) as Qty_Prev_To_Ship, CAST(0 AS FLOAT) as Qty_Prev_BkOrd, CAST('' AS CHAR(12))  AS Extra_9, CAST(0 AS FLOAT) AS Extra_10, " & _
            "CAST(0 AS FLOAT) AS Unit_Price, CAST(0 AS FLOAT) AS Discount_Pct, CAST(0 AS FLOAT) AS Calc_Price, " & _
            "GETDATE() AS Request_Dt, GETDATE() AS Promise_Dt, GETDATE() AS Req_Ship_Dt, " & _
            "'' AS Item_Desc_1, '' AS Item_Desc_2, '' AS Extra_1, " & _
            "'' AS Imprint, '' AS End_User, '' AS Imprint_Location, '' AS Imprint_Colour, '' AS Num_Impr_1, '' AS Num_Impr_2, '' AS Num_Impr_3, " & _
            "'' AS Packaging, '' AS Refill, '' AS Laser_Setup, '' AS Industry, '' AS Comments, '' AS Special_Comments, " & _
            "'' AS Printer_Comment, '' AS Packer_Comment, '' AS Printer_Instructions, '' AS Packer_Instructions, CAST(0 AS INT) AS Repeat_From_ID, " & _
            "'' AS Route, 0 AS RouteID, CAST(0 AS BIT) AS ProductProof, CAST(0 AS BIT) AS AutoCompleteReship, '' AS UOM, '' AS Tax_Sched, '' AS Revision_No, " & _
            "CAST(0 AS BIT) AS BkOrd_Fg, CAST(0 AS BIT) AS Recalc_SW, CAST('' AS CHAR(30)) AS Item_Guid, CAST('' AS CHAR(30)) AS Kit_Item_Guid, 0 AS PictureHeight, " & _
            "'' As Item_Cd, 0 as Mix_Group, '' AS LastCol " & _
            "FROM OEI_Config C WITH (NoLock) WHERE ID = 1 "

            ' , 0 AS ColIndex_Fg "

            'Imprint()
            'Imprint_Location()
            'Imprint_Colour()
            'Num_Imprints()
            'Packaging()
            'Refill()
            'Laser_Setup()
            'Comments()
            'Special_Comments()


            'strSql = "" & _
            '"SELECT " & _
            '"'' AS 'Line Seq No', '' as 'Item No', '' as Location, '' as 'Qty Ordered', '' as Qty_On_Hand, '' as 'Qty Shipped', '' as 'Qty BO', " & _
            '"'' AS 'Unit price', '' AS 'Disc Pct', '' AS 'Ext Price', '' AS 'Date Requested', " & _
            '"'' AS 'Date Promised', '' AS 'Req Ship', '' AS 'Description 1', '' AS 'Description 2', " & _
            '"'' AS 'UOM', '' AS 'Tax Schd', '' AS 'ECM Revision', '' AS 'Create/Modify Kit', " & _
            '"'' AS 'Note/Text', '' AS 'Backorder Item', '' AS 'Selected', '' AS 'Recalc', '' AS 'Line Comments', '' AS Route "

            '"SELECT " & _
            '"'' as Item, '' as Loc, '' as 'Qty Ordered/Credited', '' as 'Qty Shipped/Returned', " & _
            '"'' AS 'Unit price', '' AS 'Disc Pct', '' AS 'Ext Price', '' AS 'Requested', " & _
            '"'' AS 'Promised', '' AS 'Req Ship', '' AS 'Description', '' AS 'Item Description Line 2', " & _
            '"'' AS 'UOM', '' AS 'Tax Schd', '' AS 'ECM Revision', '' AS 'Create/Modify Kit', " & _
            '"'' AS 'Note/Text', '' AS 'Backorder Item', '' AS 'Selected', '' AS 'Recalc', '' AS 'Line Comments', '' AS Route "

            'Dim myConnection As SqlConnection = New SqlConnection("Server=thor;Initial Catalog=100;Integrated Security=SSPI")
            'Dim mySelectCommand As SqlCommand = New SqlCommand(strSql, myConnection)
            'Dim mySqlDataAdapter As SqlDataAdapter = New SqlDataAdapter(mySelectCommand)

            If dgvCurrentRow Is Nothing Then dgvCurrentRow = New DataGridViewRow

            'dsDataSet = Nothing
            'dsDataSet = New DataSet()

            Dim db As New cDBA

            dtDataTable = Nothing
            dtDataTable = db.DataTable(strSql)

            'mySqlDataAdapter.Fill(dsDataSet, "Items")
            dgvItems.DataSource = dtDataTable ' dsDataSet.Tables("Items")

            dgvItems.AllowUserToAddRows = False
            dgvItems.AllowUserToOrderColumns = False

            For lPos As Integer = 0 To dgvItems.Columns.Count - 1
                dgvItems.Columns(lPos).SortMode = DataGridViewColumnSortMode.NotSortable
            Next lPos

            dgvItems.Columns(Columns.Promise_Dt).Width = 125
            dgvItems.Columns(Columns.Item_No).Width = 125

            'dgvItems.Columns(Columns.Item_Desc_1).Width = 200
            'dgvItems.Columns(Columns.Item_Desc_2).Width = 200

            'dgvItems.Columns(Columns.Promise_Dt).Width = 100
            'dgvItems.Columns(Columns.Request_Dt).Width = 100
            'dgvItems.Columns(Columns.Req_Ship_Dt).Width = 100

            dgvItems.CausesValidation = True

            Dim oCellStyle As New DataGridViewCellStyle()
            oCellStyle.Format = "##,##0.000000"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvItems.Columns(Columns.Unit_Price).DefaultCellStyle = oCellStyle

            oCellStyle = New DataGridViewCellStyle()
            oCellStyle.Format = "##,##0.00"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvItems.Columns(Columns.Calc_Price).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Discount_Pct).DefaultCellStyle = oCellStyle

            dgvItems.Columns(Columns.Qty_Ordered).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Overpick).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_To_Ship).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Inventory).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Allocated).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_On_Hand).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Prev_To_Ship).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Qty_Prev_BkOrd).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Extra_10).DefaultCellStyle = oCellStyle

            dgvItems.Columns(Columns.Item_Guid).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Kit_Item_Guid).DefaultCellStyle = oCellStyle

            oCellStyle = New DataGridViewCellStyle()
            oCellStyle.Format = "mm/dd/yyyy"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvItems.Columns(Columns.Promise_Dt).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Request_Dt).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Req_Ship_Dt).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Extra_9).DefaultCellStyle = oCellStyle

            dgvItems.Columns(Columns.FirstCol).Frozen = True
            dgvItems.Columns(Columns.Line_Seq_No).Frozen = True
            dgvItems.Columns(Columns.Image).Frozen = True
            dgvItems.Columns(Columns.Item_No).Frozen = True

            'dgvItems.Columns(Columns.Item_Cd).Frozen = True

            dgvItems.Columns(Columns.Qty_Inventory).Visible = False
            dgvItems.Columns(Columns.Qty_Allocated).Visible = False

            Call ProductLineRestart(False)

            blnLoading = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "dgvItems Private events ##########################################"

    'Private Sub dgvItems_CancelRowEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.QuestionEventArgs) Handles dgvItems.CancelRowEdit

    '    iCompteur += 1 ''MB+
    '    Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
    '    If Not dgvItems.CurrentRow Is Nothing Then Debug.Print(dgvItems.CurrentRow.Cells(Columns.Item_No).Value.ToString)
    '    Debug.Print(g_oOrdline.Item_No)

    'End Sub

    Private Sub dgvItems_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItems.CellBeginEdit

        Debug.Print("dgvItems_CellBeginEdit")

        'iCompteur += 1 ''MB+
        'Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
        'If Not dgvItems.CurrentRow Is Nothing Then Debug.Print("")

        Try
            If blnFill Then Exit Sub

            ' Les Colonnes doivent etre programm‚es pour etre activ‚es. 
            ' Le ELSE fait un cancel = true

            Select Case e.ColumnIndex

                Case Columns.Line_Seq_No
                    e.Cancel = True

                Case Columns.Image
                    e.Cancel = True

                Case Columns.Item_No
                    If Not (dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.Equals(DBNull.Value)) Then
                        If Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref Then
                            e.Cancel = True
                        End If
                    End If
                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Qty_Ordered).Value.Equals(DBNull.Value) Then Exit Sub
                    If Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Qty_Ordered).Value) > "0" Then
                        e.Cancel = True
                    End If

                Case Columns.Location
                    If Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref Then
                        e.Cancel = True
                    End If
                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Qty_Ordered).Value.Equals(DBNull.Value) Then Exit Sub
                    If Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Qty_Ordered).Value) > "0" Then
                        e.Cancel = True
                    End If

                    'Case Columns.Qty_Ordered
                    'Case Columns.Qty_To_Ship

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

                Case Columns.Extra_9
                    e.Cancel = True

                Case Columns.Extra_10
                    e.Cancel = True

                Case Columns.Unit_Price
                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.Equals(DBNull.Value) Then
                        e.Cancel = True
                    ElseIf Not (Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref) Then
                        If dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.Equals(DBNull.Value) Then
                            e.Cancel = True
                        ElseIf Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.ToString) = "" Then
                            e.Cancel = True
                        End If
                    End If

                Case Columns.Discount_Pct
                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.Equals(DBNull.Value) Then
                        e.Cancel = True
                    ElseIf Not (Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref) Then
                        If dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.Equals(DBNull.Value) Then
                            e.Cancel = True
                        ElseIf Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.ToString) = "" Then
                            e.Cancel = True
                        End If
                    End If

                Case Columns.Calc_Price
                    e.Cancel = True

                    'Case Columns.Request_Dt
                    'Case Columns.Promise_Dt
                    'Case Columns.Req_Ship_Dt
                    'Case Columns.Item_Desc_1
                    'Case Columns.Item_Desc_2

                    'Case Columns.Comments
                    '    If Not (Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref) Then
                    '        If dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Equals(DBNull.Value) Then
                    '            e.Cancel = True
                    '        ElseIf Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value) = "" Then
                    '            e.Cancel = True
                    '        End If
                    '    End If

                Case Columns.Item_Desc_1, Columns.Item_Desc_2
                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.Equals(DBNull.Value) Then
                        e.Cancel = True
                    ElseIf Not (Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref) Then
                        If dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.Equals(DBNull.Value) Then
                            e.Cancel = True
                        ElseIf Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.ToString) = "" Then
                            e.Cancel = True
                        End If
                    End If

                Case Columns.UOM
                    e.Cancel = True

                Case Columns.Tax_Sched
                    e.Cancel = True

                Case Columns.Revision_No
                    e.Cancel = True

                    ' Traveler data columns
                Case Columns.Route

                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.Equals(DBNull.Value) Then
                        If Mid(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value, 1, 3) <> g_Comp_Pref Then
                            e.Cancel = True
                        End If
                    ElseIf Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.ToString) = "" And _
                    Not (Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref) _
                    Then
                        e.Cancel = True
                    End If

                Case Columns.ProductProof

                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.Equals(DBNull.Value) Then
                        e.Cancel = True
                    ElseIf g_oOrdline.End_Item_Cd = "K" Then ' Not (Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref) Then
                        'If dgvItems.Rows(e.RowIndex).Cells(Columns.).Equals(DBNull.Value) Then
                        e.Cancel = True
                        'ElseIf Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value) = "" Then
                        '    e.Cancel = True
                        'End If
                    End If


                Case Columns.AutoCompleteReship

                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.Equals(DBNull.Value) Then
                        e.Cancel = True
                    ElseIf g_oOrdline.End_Item_Cd = "K" Then ' Not (Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref) Then
                        e.Cancel = True
                    End If


                    ' Imprint data columns
                    ' Traveler data columns
                Case Columns.Industry

                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.Equals(DBNull.Value) Then
                        If Mid(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value, 1, 3) <> g_Comp_Pref Then
                            e.Cancel = True
                        End If
                    ElseIf Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.ToString) = "" And _
                    Not (Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref) _
                    Then
                        e.Cancel = True
                    End If

                Case _
                Columns.Imprint, Columns.End_User, Columns.Imprint_Location, Columns.Imprint_Color, Columns.Num_Impr_1, Columns.Num_Impr_2, Columns.Num_Impr_3, _
                Columns.Packaging, Columns.Refill, Columns.Comments, Columns.Special_Comments, Columns.Laser_Setup, _
                Columns.Printer_Comment, Columns.Packer_Comment, Columns.Printer_Instructions, Columns.Packer_Instructions, Columns.Repeat_From_ID ' Columns.Industry

                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.Equals(DBNull.Value) Then
                        e.Cancel = True
                    ElseIf Not (Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref) Then
                        If dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.Equals(DBNull.Value) Then
                            e.Cancel = True
                        ElseIf Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.ToString.ToString) = "" Then
                            e.Cancel = True
                        End If
                    End If

                    ' If it is num imprint, we must enter data in num imprint 1 before num imprint 2, and 2 before 3...
                    If e.ColumnIndex = Columns.Num_Impr_2 Then
                        If dgvItems.Rows(e.RowIndex).Cells(Columns.Num_Impr_1).Value.Equals(DBNull.Value) Then
                            e.Cancel = True
                        End If
                    Else
                        If e.ColumnIndex = Columns.Num_Impr_3 Then
                            If dgvItems.Rows(e.RowIndex).Cells(Columns.Num_Impr_2).Value.Equals(DBNull.Value) Then
                                e.Cancel = True
                            End If
                        End If
                    End If

                Case Else
                    If dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.Equals(DBNull.Value) Then
                        e.Cancel = True
                    ElseIf Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Location).Value.ToString.ToString) = "" Then
                        e.Cancel = True
                    End If
                    If Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref Then
                        e.Cancel = True
                    End If

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvItems_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.CellEnter

        'iCompteur += 1 ''MB+
        'Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
        'If Not dgvItems.CurrentRow Is Nothing Then Debug.Print("")
        Debug.Print("dgvItems_CellEnter")

        Try

10:         If blnDeleting Then Exit Sub

20:         Call SetColumnShortcuts()

            'If e.ColumnIndex = Columns.Route Then Call SetRouteCombo(dgvItems.Columns(Columns.Route), "All", dgvItems.Rows(e.RowIndex).Cells(Columns.Route).Value.ToString)
30:         If e.ColumnIndex = Columns.Route Then

40:             If m_oOrder.OrdLines.Contains(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_Guid).Value.ToString) Then

                    Dim oitem As cOrdLine
50:                 If Not dgvItems.Rows(e.RowIndex).Cells(Columns.Kit_Item_Guid).Value.ToString = String.Empty Then
55:                     If m_oOrder.OrdLines.Contains(dgvItems.Rows(e.RowIndex).Cells(Columns.Kit_Item_Guid).Value.ToString) Then
60:                         oitem = m_oOrder.OrdLines(dgvItems.Rows(e.RowIndex).Cells(Columns.Kit_Item_Guid).Value.ToString)
                        Else
65:                         oitem = m_oOrder.OrdLines(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_Guid).Value.ToString)
                        End If
                    Else
70:                     oitem = m_oOrder.OrdLines(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_Guid).Value.ToString)
                    End If
                    ' THIS IS A BYPASS TO FIX A BUG, NEED TO FIND IT TO FIX THE CORRECT WAY
80:                 Dim strCell As String = m_oOrder.Ordhead.Mfg_Loc
90:                 If Not dgvItems.CurrentRow.Cells(Columns.Location).Value.Equals(DBNull.Value) Then
100:                    strCell = dgvItems.CurrentRow.Cells(Columns.Location).Value
                    End If
                    ' END BYPASS

110:                If oitem.Qty_Prev_Bkord > 0 Or oitem.Qty_Ordered <> oitem.Qty_To_Ship Then
120:                    If Not mblnNoMsgBox And strCell <> "3N" Then
130:                        MsgBox("This item is no stock.")
                        End If
                    End If

                End If

140:            Call SetRouteCombo(dgvItems.Rows(e.RowIndex).Cells(Columns.Route), "All", dgvItems.Rows(e.RowIndex).Cells(Columns.Route).Value.ToString)

            End If

150:        If dgvItems.Rows.Count = 0 Then Exit Sub
            'If dgvItems.CurrentRow.Index = 0 Then Exit Sub

160:        If e.ColumnIndex < Columns.Item_No Then

170:            Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
180:            dgvItems.BeginInvoke(oInvoke, Columns.Item_No)

190:        ElseIf e.ColumnIndex = Columns.LastCol Then

200:            Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
210:            dgvItems.BeginInvoke(oInvoke, Columns.Item_No)

                'ElseIf e.ColumnIndex > Columns.Item_No And IIf(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Equals(DBNull.Value), "", dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value) = "" Then
220:        ElseIf e.ColumnIndex > Columns.Item_No And dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value Is Nothing Then
                'ElseIf e.ColumnIndex > Columns.Item_No And IIf(dgvItems.CurrentRow.Cells(Columns.Item_No).Equals(DBNull.Value), "", dgvItems.CurrentRow.Cells(Columns.Item_No).Value) = "" Then

230:            Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
240:            dgvItems.BeginInvoke(oInvoke, Columns.Item_No)

250:        ElseIf e.ColumnIndex > Columns.Item_No And Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.ToString) = "" Then

260:            Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
270:            dgvItems.BeginInvoke(oInvoke, Columns.Item_No)

            Else

                Try
                    If e.ColumnIndex = Columns.Qty_Ordered Then
                        If Mid(Trim(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value), 1, 3) = g_Comp_Pref Then
                            Exit Sub
                        End If
                        If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then
                            dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 1
                        End If
                    End If
                Catch er As Exception
                    dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 1
                End Try

            End If

350:        If dgvItems.CurrentCell.ColumnIndex > Columns.Item_No And dgvItems.CurrentCell.ColumnIndex <= Columns.Item_Desc_2 And Not (dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value Is Nothing) Then
                'If dgvItems.CurrentCell.ColumnIndex > Columns.Item_No And dgvItems.CurrentCell.ColumnIndex <= Columns.Item_Desc_2 And dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.ToString <> "" Then
360:            dgvItems.BeginEdit(True)
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & Erl() & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub dgvItems_CellValidating( _
        ByVal sender As Object, _
        ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) _
        Handles dgvItems.CellValidating


        Debug.Print("dgvItems_CellValidating")
        'iCompteur += 1 ''MB+
        'Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
        'If Not dgvItems.CurrentRow Is Nothing Then Debug.Print(dgvItems.CurrentRow.Cells(Columns.Item_No).Value.ToString)

        Dim blnValueChanged As Boolean = False

        If blnLoading Then Exit Sub
        If blnInserting Then Exit Sub

        If blnFill Then Exit Sub

        If g_oOrdline Is Nothing Then
            g_oOrdline = New cOrdLine(False) ' THAT CODE SHOULD NEVER HAPPEN
        End If

        'If e.FormattedValue.ToString = "" Then Exit Sub
        Try
            If e.ColumnIndex = Columns.Item_No Or (Mid(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, 1, 3) <> "" And e.ColumnIndex > Columns.Item_No) Then
                Select Case e.ColumnIndex
                    Case Columns.Item_No
                        Try
                            'blnValueChanged = Not (Trim(g_oOrdline.Item_No.ToString) = Trim(e.FormattedValue).ToUpper)
                            If g_oOrdline.Item_No Is Nothing Then ' .ToString <> Trim(e.FormattedValue).ToString Then
                                blnValueChanged = Not (Trim(g_oOrdline.Item_No) = Trim(e.FormattedValue).ToString)
                                g_oOrdline.Item_No = Trim(e.FormattedValue).ToString
                                If e.FormattedValue <> "" Then g_oOrdline.Pop_MessageBox(Trim(e.FormattedValue).ToString)

                                If m_oOrder.Reload Then
                                    dgvItems.BeginInvoke(New MyDelegate(AddressOf LoadPreOrder))
                                Else
                                    ' NEED TO SET IMAGE TO DATA GRID
                                    ' CHECK IN frmProductLineEntry comment je remet l'image apres 
                                    ' l'avoir redimensionnée.
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Image).Value = g_oOrdline.Image
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value = g_oOrdline.Item_No
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Location).Value = g_oOrdline.Loc
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_Desc_1).Value = g_oOrdline.Item_Desc_1
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_Desc_2).Value = g_oOrdline.Item_Desc_2
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Ordered).Value = g_oOrdline.Qty_Ordered
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Overpick).Value = g_oOrdline.Qty_Overpick
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_To_Ship).Value = g_oOrdline.Qty_To_Ship

                                    If Mid(g_oOrdline.Item_No, 1, 2) <> "44" Then

                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Inventory).Value = g_oOrdline.Qty_Inventory
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Allocated).Value = g_oOrdline.Qty_Allocated
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_To_Ship).Value = g_oOrdline.Qty_Prev_To_Ship
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_BkOrd).Value = g_oOrdline.Qty_Prev_Bkord
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_On_Hand).Value = g_oOrdline.Qty_On_Hand
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_9).Value = g_oOrdline.Extra_9
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_10).Value = g_oOrdline.Extra_10

                                    End If

                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = g_oOrdline.Unit_Price
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Discount_Pct).Value = g_oOrdline.Discount_Pct
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = g_oOrdline.Calc_Price
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Request_Dt).Value = g_oOrdline.Request_Dt
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Promise_Dt).Value = g_oOrdline.Promise_Dt
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Req_Ship_Dt).Value = g_oOrdline.Req_Ship_Dt
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.ProductProof).Value = g_oOrdline.ProductProof
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.AutoCompleteReship).Value = g_oOrdline.AutoCompleteReship
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.UOM).Value = g_oOrdline.Uom
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.BkOrd_Fg).Value = g_oOrdline.BkOrd_Fg_Bln
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_Guid).Value = g_oOrdline.Item_Guid
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Recalc_SW).Value = g_oOrdline.Recalc_Sw_Bln
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_1).Value = g_oOrdline.Extra_1_Bln
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_Cd).Value = g_oOrdline.Item_Cd
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Mix_Group).Value = g_oOrdline.Mix_Group

                                    If g_oOrdline.RouteID <> 0 Then
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value = g_oOrdline.Route
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value = g_oOrdline.RouteID
                                    End If

                                    If Not g_oOrdline.ImprintLine Is Nothing Then
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Industry).Value = g_oOrdline.ImprintLine.Industry
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Packaging).Value = g_oOrdline.ImprintLine.Packaging
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Comments).Value = g_oOrdline.ImprintLine.Comments
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Special_Comments).Value = g_oOrdline.ImprintLine.Special_Comments
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Printer_Comment).Value = g_oOrdline.ImprintLine.Printer_Comment
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Packer_Comment).Value = g_oOrdline.ImprintLine.Packer_Comment
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Printer_Instructions).Value = g_oOrdline.ImprintLine.Printer_Instructions
                                        dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Packer_Instructions).Value = g_oOrdline.ImprintLine.Packer_Instructions
                                    End If
                                    'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Kit_Item_Guid).Value = dgvItems.CurrentRow.Index

                                    If blnValueChanged Then Call SaveOrderLine()
                                    'Call SaveOrderLine()

                                    If g_oOrdline.m_OEI_W_Pixels <> 400 Then
                                        Dim dblHeight As Double
                                        dblHeight = IIf(400 / g_oOrdline.m_OEI_W_Pixels > 5, 5, 400 / g_oOrdline.m_OEI_W_Pixels)

                                        If chkLowRow.Checked Then
                                            dgvItems.Rows(dgvItems.CurrentRow.Index).Height = 22
                                        Else
                                            dgvItems.Rows(dgvItems.CurrentRow.Index).Height = 120 / dblHeight
                                        End If
                                        'dgvItems.Rows(dgvItems.CurrentRow.Index).Height = 120 / dblHeight
                                    Else
                                        Call ResizeImage(dgvItems.CurrentRow.Index, g_oOrdline.Image)
                                    End If
                                    If g_oOrdline.Kit.IsAKit And g_oOrdline.Qty_Ordered = 0 Then ' And blnValueChanged Then
                                        Dim oKit As New frmKitSelection
                                        oKit.Kit = g_oOrdline.Kit
                                        oKit.ShowDialog()
                                        oKit.Dispose()

                                        dgvItems.BeginInvoke(New MyDelegate(AddressOf LoadPreOrder))

                                    End If

                                    Call RefreshImprint(e.RowIndex)

                                    '' Change the color of cells test
                                    'If dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString.Contains("EC1040") Then
                                    '    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Style.BackColor = Color.Cyan
                                    'ElseIf dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString.Contains("EC1050") Then
                                    '    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Style.BackColor = Color.Magenta
                                    'Else
                                    '    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Style.BackColor = Color.Empty
                                    'End If

                                End If

                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Location

                        Try
                            If Mid(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, 1, 3) <> g_Comp_Pref Then

                                'If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Location_Must_Be_Numeric, True)

                                g_oOrdline.Loc = (e.FormattedValue)

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Ordered).Value = g_oOrdline.Qty_Ordered
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Overpick).Value = g_oOrdline.Qty_Overpick
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_To_Ship).Value = g_oOrdline.Qty_To_Ship

                                If Mid(g_oOrdline.Item_No, 1, 2) <> "44" Then

                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Inventory).Value = g_oOrdline.Qty_Inventory
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Allocated).Value = g_oOrdline.Qty_Allocated
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_To_Ship).Value = g_oOrdline.Qty_Prev_To_Ship
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_BkOrd).Value = g_oOrdline.Qty_Prev_Bkord
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_On_Hand).Value = g_oOrdline.Qty_On_Hand

                                End If

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = g_oOrdline.Calc_Price
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = g_oOrdline.Unit_Price

                                If g_oOrdline.End_Item_Cd = "K" Then
                                    dgvItems.BeginInvoke(New Qty_On_Hand_Kits_Delegate(AddressOf Set_Qty_On_Hand_Kits), dgvItems.CurrentCell.RowIndex) ' , dgvItems.CurrentCell.ColumnIndex)
                                End If

                                'Call SaveOrderLine()

                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Qty_Ordered

                        Try
                            If Mid(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, 1, 3) <> g_Comp_Pref Then

                                If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)

                                If Mid(g_oOrdline.Item_No, 1, 2) = "44" Then
                                    If Not m_oOrder.CheckSetupQty(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, CDbl(e.FormattedValue)) Then
                                        Throw New OEException(OEError.Too_Many_Setups_Entered, True, False) 'Too_Many_Setups_Entered
                                    End If
                                End If

                                g_oOrdline.Qty_Ordered = CDbl(e.FormattedValue)

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Ordered).Value = g_oOrdline.Qty_Ordered
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Overpick).Value = g_oOrdline.Qty_Overpick
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_To_Ship).Value = g_oOrdline.Qty_To_Ship

                                If Mid(g_oOrdline.Item_No, 1, 2) <> "44" Then
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Inventory).Value = g_oOrdline.Qty_Inventory
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Allocated).Value = g_oOrdline.Qty_Allocated
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_To_Ship).Value = g_oOrdline.Qty_Prev_To_Ship
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_BkOrd).Value = g_oOrdline.Qty_Prev_Bkord
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_On_Hand).Value = g_oOrdline.Qty_On_Hand
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_9).Value = g_oOrdline.Extra_9
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_10).Value = g_oOrdline.Extra_10
                                End If

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = g_oOrdline.Unit_Price
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = g_oOrdline.Calc_Price

                                Call UpdateOrderTotals()
                                'Call SaveOrderLine()

                                Call execThermoSimulation()

                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            'If oe_er.ShowMessage Then MsgBox(oe_er.Message)
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message, MsgBoxStyle.OkOnly, oe_er.Header)

                        End Try

                    Case Columns.Qty_To_Ship

                        Try
                            If Mid(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, 1, 3) <> g_Comp_Pref Then
                                If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)

                                Dim oResult As New MsgBoxResult
                                If CDbl(e.FormattedValue) > g_oOrdline.Qty_Ordered Then
                                    oResult = MsgBox(New OEExceptionMessage(OEError.Qty_To_Ship_GT_Qty_Ordered).Message, MsgBoxStyle.YesNo, "Error")
                                Else
                                    oResult = MsgBoxResult.Yes
                                End If

                                If oResult = MsgBoxResult.Yes Then
                                    g_oOrdline.Qty_To_Ship = CDbl(e.FormattedValue)
                                Else
                                    g_oOrdline.Qty_To_Ship = g_oOrdline.Qty_Ordered
                                    Dim oInvoke As New SetQty_To_Ship(AddressOf SetQty_To_ShipSub)
                                    dgvItems.BeginInvoke(oInvoke)
                                End If

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_To_Ship).Value = g_oOrdline.Qty_To_Ship

                                If Mid(g_oOrdline.Item_No, 1, 2) <> "44" Then

                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_To_Ship).Value = g_oOrdline.Qty_Prev_To_Ship
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Qty_Prev_BkOrd).Value = g_oOrdline.Qty_Prev_Bkord
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_9).Value = g_oOrdline.Extra_9
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Extra_10).Value = g_oOrdline.Extra_10
                                End If

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = g_oOrdline.Calc_Price

                                Call UpdateOrderTotals()

                                Call execThermoSimulation()

                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            'If oe_er.ShowMessage Then MsgBox(oe_er.Message)
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message, MsgBoxStyle.OkOnly, oe_er.Header)

                        End Try

                    Case Columns.Qty_Prev_To_Ship

                    Case Columns.Extra_9

                    Case Columns.Extra_10

                    Case Columns.Unit_Price

                        Try
                            If Mid(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, 1, 3) <> g_Comp_Pref Then

                                If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                                g_oOrdline.Unit_Price = CDbl(e.FormattedValue)

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Unit_Price).Value = g_oOrdline.Unit_Price
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = g_oOrdline.Calc_Price

                                Call UpdateOrderTotals()

                                'Call SaveOrderLine()

                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Discount_Pct

                        Try
                            If Mid(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, 1, 3) <> g_Comp_Pref Then

                                If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                                g_oOrdline.Discount_Pct = CDbl(e.FormattedValue)

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Discount_Pct).Value = g_oOrdline.Discount_Pct
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Calc_Price).Value = g_oOrdline.Calc_Price

                                Call UpdateOrderTotals()

                                'Call SaveOrderLine()
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Calc_Price
                        'Call UpdateOrderTotals()

                    Case Columns.Request_Dt

                        Try
                            If Mid(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, 1, 3) <> g_Comp_Pref Then

                                If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
                                g_oOrdline.Request_Dt = CDate(e.FormattedValue)

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Promise_Dt).Value = g_oOrdline.Promise_Dt
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Req_Ship_Dt).Value = g_oOrdline.Req_Ship_Dt

                                'Call SaveOrderLine()
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Promise_Dt

                        Try
                            If Mid(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, 1, 3) <> g_Comp_Pref Then

                                If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
                                g_oOrdline.Promise_Dt = CDate(e.FormattedValue)

                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Req_Ship_Dt).Value = g_oOrdline.Req_Ship_Dt

                                Call execThermoSimulation()
                                'Call SaveOrderLine()

                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Req_Ship_Dt
                        Try
                            If Mid(dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString, 1, 3) <> g_Comp_Pref Then

                                If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
                                g_oOrdline.Req_Ship_Dt = CDate(e.FormattedValue)

                                'Call SaveOrderLine()

                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Item_Desc_1

                        Try
                            g_oOrdline.Item_Desc_1 = (e.FormattedValue)

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Item_Desc_2

                        Try
                            g_oOrdline.Item_Desc_2 = (e.FormattedValue)

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

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
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                'If Trim(e.FormattedValue) <> "" Then
                                g_oOrdline.SetRoute(e.FormattedValue)

                                g_oOrdline.Set_Spec_Instructions(g_oOrdline.Item_No, g_oOrdline.RouteID, g_oOrdline.ImprintLine)

                                g_oOrdline.ImprintLine.SetNumImprints(g_oOrdline.RouteID, getCurrentLineItems())
                                If Not g_oOrdline.ImprintLine.IsEmpty() Then g_oOrdline.ImprintLine.Save()

                                'save orderline
                                g_oOrdline.Save()

                                'update cell values immediately for thermo simulation
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value = g_oOrdline.RouteID
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value = g_oOrdline.Route
                                If g_oOrdline.ImprintLine.Num_Impr_1 = 0 Then
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_1).Value = DBNull.Value
                                Else
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_1).Value = g_oOrdline.ImprintLine.Num_Impr_1
                                End If

                                If g_oOrdline.ImprintLine.Num_Impr_2 = 0 Then
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_2).Value = DBNull.Value
                                Else
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_2).Value = g_oOrdline.ImprintLine.Num_Impr_2
                                End If

                                If g_oOrdline.ImprintLine.Num_Impr_3 = 0 Then
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_3).Value = DBNull.Value
                                Else
                                    dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_3).Value = g_oOrdline.ImprintLine.Num_Impr_3
                                End If

                                ' '' ''dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_2).Value = g_oOrdline.ImprintLine.Num_Impr_2
                                ' '' ''dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Num_Impr_3).Value = g_oOrdline.ImprintLine.Num_Impr_3

                                Call RefreshImprint(dgvItems.CurrentRow.Index)

                                Call execThermoSimulation()
                                dgvItems.BeginInvoke(New Del_Set_Num_Imprint_To_Grid(AddressOf Set_Num_Imprint_To_Grid), dgvItems.CurrentRow.Index, g_oOrdline) ' , dgvItems.CurrentCell.ColumnIndex)

                                ' '' ''Del_SetProductionAvailabilityColor
                                ' '' ''Me.BeginInvoke(New MethodInvoker(addressof g_oOrdline.SetProductionAvailabilityColor(dgvItems.CurrentRow.Cells(Columns.Item_No)))
                                ' '' ''End If
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.ProductProof
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ProductProof_Bln = (e.FormattedValue)
                                g_oOrdline.Save()
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.AutoCompleteReship
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.AutoCompleteReship_Bln = (e.FormattedValue)
                                g_oOrdline.Save()
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Extra_1
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.Extra_1_Bln = (e.FormattedValue)
                                g_oOrdline.Save()
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Imprint
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Imprint = (e.FormattedValue)
                                g_oOrdline.Set_Spec_Imprint_Instructions(m_oOrder.Ordhead.Cus_No, g_oOrdline.ImprintLine)
                                '  If g_oOrdline.Set_Spec_Imprint_Instructions(m_oOrder.Ordhead.Cus_No, g_oOrdline.ImprintLine) Then
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Printer_Instructions).Value = g_oOrdline.ImprintLine.Printer_Instructions
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Packer_Instructions).Value = g_oOrdline.ImprintLine.Packer_Instructions
                                'End If
                                txtImprint.Text = (e.FormattedValue)
                                dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Imprint).Value = txtImprint.Text
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.End_User
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.End_User = (e.FormattedValue)
                                txtEnd_User.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Imprint_Location
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Location = (e.FormattedValue)
                                txtImprint_Location.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Imprint_Color
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Color = (e.FormattedValue)
                                txtImprint_Color.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Num_Impr_1
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                If Trim(e.FormattedValue) <> "" Then
                                    g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                    If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                                    g_oOrdline.ImprintLine.Num_Impr_1 = CInt(e.FormattedValue)
                                    txtNum_Impr_1.Text = (e.FormattedValue)
                                Else
                                    g_oOrdline.ImprintLine.Num_Impr_1 = 0
                                    txtNum_Impr_1.Text = ""
                                End If

                                Call execThermoSimulation()
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Num_Impr_2
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                If Trim(e.FormattedValue) <> "" Then
                                    g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                    If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                                    g_oOrdline.ImprintLine.Num_Impr_2 = CInt(e.FormattedValue)
                                    txtNum_Impr_2.Text = (e.FormattedValue)
                                Else
                                    g_oOrdline.ImprintLine.Num_Impr_2 = 0
                                    txtNum_Impr_2.Text = ""
                                End If

                                Call execThermoSimulation()
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Num_Impr_3
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                If Trim(e.FormattedValue) <> "" Then
                                    g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                    If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                                    g_oOrdline.ImprintLine.Num_Impr_3 = CInt(e.FormattedValue)
                                    txtNum_Impr_3.Text = (e.FormattedValue)
                                Else
                                    g_oOrdline.ImprintLine.Num_Impr_3 = 0
                                    txtNum_Impr_3.Text = ""
                                End If

                                Call execThermoSimulation()
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Packaging
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Packaging = (e.FormattedValue)
                                txtPackaging.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Refill
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Refill = (e.FormattedValue)
                                txtRefill.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Laser_Setup
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Laser_Setup = (e.FormattedValue)
                                'txtRefill.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Industry
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Industry = (e.FormattedValue)
                                cboIndustry.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Comments
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Comments = (e.FormattedValue)
                                txtComments.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Special_Comments
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Special_Comments = (e.FormattedValue)
                                txtSpecial_Comments.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Printer_Comment
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Printer_Comment = (e.FormattedValue)
                                'txtPrinter_Comment.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Packer_Comment
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Packer_Comment = (e.FormattedValue)
                                'txtPacker_Comment.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Printer_Instructions
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Printer_Instructions = (e.FormattedValue)
                                'txtPacker_Comment.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Packer_Instructions
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                g_oOrdline.ImprintLine.Packer_Instructions = (e.FormattedValue)
                                'txtPacker_Comment.Text = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Repeat_From_ID
                        Try
                            If Not (g_oOrdline.Item_No Is Nothing) Then
                                If Trim(e.FormattedValue) <> "" Then
                                    g_oOrdline.ImprintLine.Item_No = g_oOrdline.Item_No
                                    If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                                    g_oOrdline.ImprintLine.Repeat_From_ID = CInt(e.FormattedValue)
                                    txtRepeat_From_ID.Text = (e.FormattedValue)
                                End If
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)

                        End Try

                    Case Columns.Mix_Group
                        Try
                            If Not (g_oOrdline.Mix_Group Is Nothing) Then
                                g_oOrdline.Mix_Group = (e.FormattedValue)
                            End If

                        Catch oe_er As OEException
                            e.Cancel = oe_er.Cancel
                            If oe_er.ShowMessage Then MsgBox(oe_er.Message)
                        End Try


                    Case Columns.Item_Guid, Columns.Kit_Item_Guid

                End Select
            End If

        Catch er_oe As OEException
            MsgBox(er_oe.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function getCurrentLineItems() As String()
        Dim ITEM_CD As String = ""
        Dim ITEM_NO As String = ""

        Dim rows As Integer = dgvItems.Rows.Count - 1

        For i As Integer = 0 To rows
            ITEM_CD = ITEM_CD + "'" + dgvItems.Rows(i).Cells(Columns.Item_Cd).Value.ToString + "',"
            ITEM_NO = ITEM_NO + "'" + dgvItems.Rows(i).Cells(Columns.Item_No).Value.ToString + "',"
        Next

        'Remove unwanted comma character
        ITEM_CD = ITEM_CD.Substring(0, ITEM_CD.Length - 1)
        ITEM_NO = ITEM_NO.Substring(0, ITEM_NO.Length - 1)

        Return New String() {ITEM_CD, ITEM_NO}
    End Function

    Private Sub Set_Num_Imprint_To_Grid(ByVal iCurrentRowIndex As Integer, ByRef oOrdline As cOrdLine)

        Try
            If iCurrentRowIndex >= dgvItems.Rows.Count Then Exit Sub

            With dgvItems.Rows(iCurrentRowIndex)

                .Cells(Columns.RouteID).Value = g_oOrdline.RouteID

                If oOrdline.ImprintLine.Num_Impr_1 = 0 Then
                    .Cells(Columns.Num_Impr_1).Value = DBNull.Value
                Else
                    .Cells(Columns.Num_Impr_1).Value = oOrdline.ImprintLine.Num_Impr_1
                End If

                If oOrdline.ImprintLine.Num_Impr_2 = 0 Then
                    .Cells(Columns.Num_Impr_2).Value = DBNull.Value
                Else
                    .Cells(Columns.Num_Impr_2).Value = oOrdline.ImprintLine.Num_Impr_2
                End If

                If oOrdline.ImprintLine.Num_Impr_3 = 0 Then
                    .Cells(Columns.Num_Impr_3).Value = DBNull.Value
                Else
                    .Cells(Columns.Num_Impr_3).Value = oOrdline.ImprintLine.Num_Impr_3
                End If

                If oOrdline.ImprintLine.Printer_Comment = "" Then
                    .Cells(Columns.Printer_Comment).Value = DBNull.Value
                Else
                    .Cells(Columns.Printer_Comment).Value = oOrdline.ImprintLine.Printer_Comment
                End If

                If oOrdline.ImprintLine.Printer_Instructions = "" Then
                    .Cells(Columns.Printer_Instructions).Value = DBNull.Value
                Else
                    .Cells(Columns.Printer_Instructions).Value = oOrdline.ImprintLine.Printer_Instructions
                End If

                If oOrdline.ImprintLine.Packer_Comment = "" Then
                    .Cells(Columns.Packer_Comment).Value = DBNull.Value
                Else
                    .Cells(Columns.Packer_Comment).Value = oOrdline.ImprintLine.Packer_Comment
                End If

                If oOrdline.ImprintLine.Packer_Instructions = "" Then
                    .Cells(Columns.Packer_Instructions).Value = DBNull.Value
                Else
                    .Cells(Columns.Packer_Instructions).Value = oOrdline.ImprintLine.Packer_Instructions
                End If

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub dgvItems_EditingControlShowing(ByVal sender As Object, _
                                                    ByVal e As DataGridViewEditingControlShowingEventArgs) Handles dgvItems.EditingControlShowing
        'iCompteur += 1 ''MB+
        'Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
        'If Not dgvItems.CurrentRow Is Nothing Then Debug.Print("")

        'Make sure we are in the first column.
        'If Me.dgvItems.CurrentCellAddress.X = 0 Then
        '    With e.Control
        Try
            If TypeOf e.Control Is TextBox Then

                Dim tb As TextBox = CType(e.Control, TextBox)
                'Remove the existing handler if there is one.
                RemoveHandler tb.KeyDown, AddressOf XdgvItems_KeyDown ' TextBox_TextChanged

                'Add a new handler.
                AddHandler tb.KeyDown, AddressOf XdgvItems_KeyDown ' TextBox_TextChanged
                '    End With
                'End If

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvItems_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItems.Leave

        Debug.Print("dgvItems_Leave")

        'iCompteur += 1 ''MB+
        'Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
        'If Not dgvItems.CurrentRow Is Nothing Then Debug.Print(dgvItems.CurrentRow.Cells(Columns.Item_No).Value.ToString)

        Try
            If dgvItems.Rows.Count = 0 Then Exit Sub
            If dgvItems.CurrentRow.Cells(Columns.Item_No).Value.Equals(DBNull.Value) Then Exit Sub
            If dgvItems.CurrentRow.Cells(Columns.Item_No).Value.ToString = "" Then Exit Sub

            dgvItems.UpdateCellValue(dgvItems.CurrentCell.ColumnIndex, dgvItems.CurrentCell.RowIndex)

            Call SaveOrderLine(g_oOrdline)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvItems_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.RowEnter

        Debug.Print("dgvItems_RowEnter")

        'iCompteur += 1 ''MB+
        'Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
        'If Not dgvItems.CurrentRow Is Nothing Then Debug.Print("")

        Try
            If blnDeleting Then Exit Sub
            'If dgvItems.CurrentRow Is Nothing Then Exit Sub

            If Not g_oOrdline Is Nothing Then
                If g_oOrdline.Qty_Ordered <> 0 And g_oOrdline.SaveToDB Then
                    g_oOrdline.Save()
                End If
            End If

            If Not (dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value Is Nothing) Then
                'Debug.Print("031")
                cmdDeleteLine.Enabled = (Mid(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value.ToString, 3, 1) <> " ")
            End If

            If dgvItems.Rows(e.RowIndex).Cells(Columns.Item_Guid).Value Is Nothing Then
                g_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID) ' m_oOrder)
            ElseIf m_oOrder.OrdLines.Contains(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_Guid).Value.ToString) Then ' (e.RowIndex <= (m_oOrder.OrdLines.Count - 1)) Then
                g_oOrdline = m_oOrder.OrdLines(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_Guid).Value.ToString) ' e.RowIndex + 1)
            Else
                g_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID) ' m_oOrder)
            End If

            Call RefreshImprint(e.RowIndex)

            'Call ChangeWebPage(e.RowIndex)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvItems_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.RowLeave

        Debug.Print("dgvItems_RowLeave")

        'iCompteur += 1 ''MB+
        'Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
        'If Not dgvItems.CurrentRow Is Nothing Then Debug.Print("")

        Try
            If blnFill Then Exit Sub

            'Call SaveOrderLine()
            Call SaveOrderLine(g_oOrdline)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub XdgvItems_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) 'Handles dgvItems.KeyDown

        'iCompteur += 1 ''MB+
        'Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
        'If Not dgvItems.CurrentRow Is Nothing Then Debug.Print(dgvItems.CurrentRow.Cells(Columns.Item_No).Value.ToString)

        Try

            If dgvItems.Rows.Count = 0 Then Exit Sub

            If e.KeyCode = Keys.Escape Then
                If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value.Equals(DBNull.Value) Then
                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then
                    cmdDeleteLine.PerformClick()
                    Exit Sub

                ElseIf dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then
                    cmdDeleteLine.PerformClick()
                    Exit Sub

                Else
                    Exit Sub
                End If


            ElseIf e.KeyCode = Keys.Return Then
                Call ProductLineInsert()
                Exit Sub
            End If

            Select Case dgvItems.CurrentCell.ColumnIndex

                Case Columns.Item_No

                    If e.KeyCode = Keys.F5 Then

                        If Trim(dgvItems.CurrentCell.Value.ToString) <> "" Then
                            Call PerformItemInfoClick()
                        End If

                    ElseIf e.KeyCode = Keys.F6 Then

                        If Trim(dgvItems.CurrentRow.Cells(Columns.Location).Value.ToString) <> "" Then
                            Call PerformItemLocQtyClick()
                        End If

                    ElseIf e.KeyCode = Keys.F7 Then

                        If Trim(dgvItems.CurrentCell.Value.ToString) = "" Then
                            Call PerformGridSearchClick()
                        End If

                    ElseIf e.KeyCode = Keys.F8 Then

                        If Trim(dgvItems.CurrentCell.Value.ToString) <> "" Then
                            Call PerformGridNotesClick()
                        End If

                    ElseIf e.Control Then

                        Select Case e.KeyCode
                            Case Keys.D
                                dgvItems.CurrentCell.Value = "44LOGO00DBS"
                                dgvItems.BeginEdit(True)
                            Case Keys.F
                                dgvItems.CurrentCell.Value = "44PAPR00PRF"
                                dgvItems.BeginEdit(True)
                            Case Keys.S
                                dgvItems.CurrentCell.Value = "44LOGO0001C"
                                dgvItems.BeginEdit(True)
                            Case Keys.Q
                                dgvItems.CurrentCell.Value = "44EXCT00QNT"
                                dgvItems.BeginEdit(True)
                            Case Keys.L
                                dgvItems.CurrentCell.Value = "44LOGO00LAZ"
                                dgvItems.BeginEdit(True)
                            Case Keys.P
                                dgvItems.CurrentCell.Value = "44P2P00CRG"
                                dgvItems.BeginEdit(True)
                            Case Keys.R
                                dgvItems.CurrentCell.Value = "44LAZR00OPM"
                                dgvItems.BeginEdit(True)
                            Case Keys.C
                                dgvItems.CurrentCell.Value = "44COLR00MCH"
                                dgvItems.BeginEdit(True)
                            Case Keys.M
                                dgvItems.CurrentCell.Value = "44LOGOMET1C"
                                dgvItems.BeginEdit(True)


                        End Select

                    End If

                Case Columns.Location

                    Select Case e.KeyCode

                        Case Keys.F5
                            Call PerformItemInfoClick()

                        Case Keys.F6
                            Call PerformItemLocQtyClick()

                        Case Keys.F7
                            If Trim(dgvItems.CurrentCell.Value.ToString) = "" Then
                                Call PerformGridSearchClick()
                            End If

                        Case Keys.F8
                            If Trim(dgvItems.CurrentCell.Value.ToString) <> "" Then
                                Call PerformGridNotesClick()
                            End If

                    End Select

                Case Columns.Qty_Ordered, Columns.Qty_To_Ship, Columns.Qty_Inventory, _
                     Columns.Qty_Allocated, Columns.Qty_On_Hand, Columns.Qty_Prev_To_Ship, Columns.Qty_Prev_BkOrd

                    Select Case e.KeyCode

                        Case Keys.F5
                            Call PerformItemInfoClick()

                        Case Keys.F6
                            Call PerformItemLocQtyClick()

                    End Select

                Case Columns.Extra_9, Columns.Extra_10

                    Select Case e.KeyCode

                        Case Keys.F3
                            Call PerformETAClick()

                        Case Keys.F5
                            Call PerformItemInfoClick()

                    End Select

                Case Columns.Unit_Price, Columns.Discount_Pct, Columns.Calc_Price

                    Select Case e.KeyCode

                        Case Keys.F4
                            Call PerformPriceBreaksClick()

                        Case Keys.F5
                            Call PerformItemInfoClick()

                    End Select

                Case Columns.Route

                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    ''If e.Control Then

                    ''    Dim strValue As String = dgvItems.CurrentCell.Value.ToString

                    ''    Select Case e.KeyCode
                    ''        Case Keys.A
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "All", strValue)

                    ''        Case Keys.F
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Flat", strValue)

                    ''        Case Keys.S
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Silk", strValue)

                    ''        Case Keys.P
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Pad", strValue)

                    ''        Case Keys.T
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Silk/Flat", strValue)

                    ''        Case Keys.L
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Laser/Silk", strValue)

                    ''        Case Keys.D
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Silk/Pad", strValue)

                    ''        Case Keys.O
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "OS", strValue)

                    ''        Case Keys.M
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Misc", strValue)

                    ''    End Select

                    ''End If

                Case Columns.Imprint

                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    '-- NO IMPRINT	 37765	 5%
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

                    ' MUST CHECK IF IT IS A KIT ITEM, THEN Qty_Ordered WILL BE 0 BUT WE NEED TO ALLOW !!!
                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Equals(DBNull.Value) Then Exit Sub
                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    '-- CAP			214884	33% + 10% NO VALUE
                    '--	BARREL		 95876	15%
                    '--	ENLARGED AREA55470	 7%
                    '--	STANDARD	 45962	 6%
                    '-- ENLARGED 	 33253	 5%
                    '-- FRONT 	 	 20321	 3%

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

                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Equals(DBNull.Value) Then Exit Sub
                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    '-- BLACK		175882	27% + 6% NO VALUE
                    '--	SILVER		 85516	13%	
                    '--	LASER		 71272	11%
                    '--	WHITE		 67916	10%
                    '-- DEBOSS		 24212	3%
                    '-- LASER/OXID	 22382	3%

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

                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    '-- BULK		229346	36% + 3% NO VALUE
                    '--	CELLO		 88021	13%
                    '--	STD		 	 66508	10%
                    '--	P100		 48272	 7%
                    '-- SHRINK WRAP	 12395	 3%

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

                Case Else

                    Select Case e.KeyCode
                        Case Keys.Return
                            Call ProductLineInsert()

                            'Case Keys.F5
                            '    'MsgBox("Open whatever...")

                            'Case Keys.F7
                            '    Select Case dgvItems.CurrentCell.ColumnIndex
                            '        Case Columns.Item_No
                            '            Call PerformGridSearchClick()
                            '        Case Columns.Location
                            '            Call PerformGridSearchClick()
                            '        Case Else
                            '            'Do nothing
                            '    End Select

                            'Case Keys.F8
                            '    Select Case dgvItems.CurrentCell.ColumnIndex
                            '        Case Columns.Item_No
                            '            Call PerformGridNotesClick()
                            '        Case Columns.Location
                            '            Call PerformGridNotesClick()
                            '        Case Else
                            '            'Do nothing
                            '    End Select

                    End Select

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

    Private Sub SetColumnIndexSub(ByVal columnIndex As Integer)

        dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(columnIndex)
        'dgvItems.BeginEdit(True)

    End Sub

    Private Sub SetQty_To_ShipSub()

        dgvItems.CurrentRow.Cells(Columns.Qty_To_Ship).Value = g_oOrdline.Qty_To_Ship

    End Sub

#Region "    ucContacts ########################################## "

    Private Sub UcContacts1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcContacts1.Enter

        Try

            UcContacts1.Fill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub UcContacts1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcContacts1.Leave

        Try

            UcContacts1.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "    ucDocument ########################################## "

    Private Sub ucDocument1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcDocument1.Enter

        Try

            UcDocument1.Item_Guid = g_oOrdline.Item_Guid

            UcDocument1.Fill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ucDocument1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcDocument1.Leave

        Try

            UcDocument1.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "    ucOrder ############################################# "

    Private Sub UcOrder_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcOrder.Enter

        Try

            UcOrder.Fill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub UcOrder_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcOrder.Leave

        Try

            UcOrder.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "    ucCreditInfo ######################################## "

    Private Sub UcCreditInfo1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcCreditInfo1.Enter

        Try

            UcCreditInfo1.Fill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub UcCreditInfo1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcCreditInfo1.Leave

        Try

            UcCreditInfo1.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "    ucExtra1 ############################################ "

    Private Sub UcExtra1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcExtra1.Enter

        Try

            UcExtra1.Fill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub UcExtra1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcExtra1.Leave

        Try

            UcExtra1.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "    ucHeader1 ########################################### "

    Private Sub UcHeader1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcHeader1.Enter

        Try

            UcHeader1.Fill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub UcHeader1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcHeader1.Leave

        Try

            UcHeader1.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "    ucImprint1 ########################################## "

    Private Sub UcImprint1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcImprint1.Enter

        'Try

        '    UcImprint1.ImprintFill()

        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

    Private Sub UcImprint1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcImprint1.Leave

        'Try

        '    'UcImprint1.ImprintSave()

        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

#End Region

#Region "    ucOrderTotal1 ####################################### "

    Private Sub UcOrderTotal1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcOrderTotal1.Enter

        Try

            UcOrderTotal1.Fill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub UcOrderTotal1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcOrderTotal1.Leave

        Try

            UcOrderTotal1.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "    ucSalesperson1 ###################################### "

    Private Sub UcSalesperson1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcSalesperson1.Enter

        Try

            UcSalesperson1.Fill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub UcSalesperson1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcSalesperson1.Leave

        Try

            UcSalesperson1.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "    ucTaxes1 ############################################ "

    Private Sub UcTaxes1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcTaxes1.Enter

        Try

            UcTaxes1.Fill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub UcTaxes1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles UcTaxes1.Leave

        Try

            UcTaxes1.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "    tabControl ########################################## "

    Private Sub sstOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles sstOrder.Click

        Dim tcControl As TabControl = sender
        'Dim tcPage As TabPage = tcControl.TabPages(tcControl.TabIndex)

        Try
            If m_oOrder Is Nothing Then
                m_oOrder = New cOrder()
                UcOrder.Fill() ' m_oOrder)
            End If

            'tcControl.SelectedIndex = 2
            'Dim iSelectedIndex As Integer = CInt(Mid(tcPage.Name, tcPage.Name.Length - 1, 2))

            If tcControl.SelectedIndex <> Tabs.History And tcControl.SelectedIndex <> Tabs.ReservedItems Then
                If tcControl.SelectedIndex > Tabs.Header And (Trim(m_oOrder.Ordhead.Cus_No) = "" Or m_oOrder.Ordhead.Cus_No Is Nothing) Then
                    sstOrder.SelectTab(Tabs.Header)
                    Exit Sub
                End If

                'If tcControl.SelectedIndex > Tabs.CustomerContact And (m_oOrder.Ordhead.ContactView = 0) Then
                If tcControl.SelectedIndex > Tabs.CustomerContact And (m_oOrder.Ordhead.ContactView = 0) Then
                    sstOrder.SelectTab(Tabs.CustomerContact)
                    m_oOrder.Ordhead.ContactView = 1
                    Exit Sub
                End If
            End If

            Select Case tcControl.SelectedIndex ' tcControl.SelectedIndex

                Case Tabs.Header
                    UcOrder.Fill()

                Case Tabs.CustomerContact
                    m_oOrder.Ordhead.ContactView = 1
                    UcContacts1.Fill()

                Case Tabs.Lines
                    If m_oOrder.OrdLines.Count = 0 Then
                        Call ProductLineRestart(False)
                        dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Item_No)

                        'add automatic charges when order is empty
                        If m_oOrder.Ordhead.User_Def_Fld_2.ToString.ToUpper.Trim() <> "RS" Then
                            Me.cmdPrevCharges.PerformClick()
                        End If
                    Else
                        If m_oOrder.OrdLines.Count <> 0 And dgvItems.Rows.Count <= 1 Then
                            Call LoadPreOrder()
                            dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Item_No)
                        Else
                            Call Fill()
                        End If

                    End If
                    UcOrderTotal1.From_Curr_Cd_Desc = m_oOrder.Ordhead.Curr_Cd
                    chkAutoCompleteReship.Checked = (m_oOrder.Ordhead.AutoCompleteReship <> 0)

                    dgvItems.Focus()
                    'frmOrder.char()

                Case Tabs.Documents
                        UcDocument1.Fill()

                Case Tabs.HeaderAll
                        UcHeader1.Fill()

                Case Tabs.Taxes
                        UcTaxes1.Fill()

                Case Tabs.Salesperson
                        UcSalesperson1.Fill()

                Case Tabs.CreditInfo
                        UcCreditInfo1.Fill()

                Case Tabs.Extra
                        UcExtra1.Fill()

                Case Tabs.History
                        Call LoadHistoryGridValues()

                Case Tabs.ReservedItems
                        Call LoadReservedItemsValues()

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function splitStringIntoBlocks(str As String, charLimit As Integer) As List(Of String)
        Dim blockSet As New List(Of String)
        Dim charCounter As Integer
        Dim strBlock As String = ""

        'split string
        Dim wordArr() As String = Split(str)

        For Each word As String In wordArr
            'append string block to list and reset
            If charCounter >= 40 Or (charCounter + word.Length + 1) >= 40 Then
                blockSet.Add(strBlock.Trim())
                strBlock = ""
                charCounter = 0
            End If

            strBlock = strBlock & " " & word.Trim() 'append word block
            charCounter = charCounter + word.Length + 1 'update char counter
        Next

        blockSet.Add(strBlock) 'append final string block to list
        blockSet.Add("%%%") 'append delimiter
        '   blockSet.Reverse() 'reverse order

        Return blockSet
    End Function

    Private Sub insertLineSeqCmt(Ord_Guid As String, Line_Seq_No As Integer, Item_Cd As String, Item_No As String)

        Try

            Dim db As New cDBA
            Dim dt As DataTable

            Dim contact As String = UcContacts1.txtOEContact.Text.Trim()
            'ucContacts.ControlCollection.
            Dim strSql As String = "SELECT DISTINCT " & _
                                   " CASE WHEN c.taalcode = 'FR' THEN ISNULL(s.SPEC_INSTRUCTION_FR, '') " & _
                                   "                             ELSE ISNULL(s.SPEC_INSTRUCTION, '') END AS SPEC_INSTRUCTION " & _
                                   " FROM OEI_ITEM_SPEC_INSTRUCTION as s " & _
                                   " LEFT JOIN cicntp as c ON c.FullName = '" & contact & "'" & _
                                   " WHERE s.DATA_FIELD = 'ORDER_ACK_ITEM_COMMENT' " & _
                                   "    AND ( " & _
                                   "           (s.ITEM_CD = '" & Item_Cd & "' AND ISNULL(s.ITEM_NO, '') = '') " & _
                                   "        OR (s.ITEM_CD = '" & Item_Cd & "' AND s.ITEM_NO = '" & Item_No & "') " & _
                                   "        )"

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then

                'Delete existing records (say failed export)
                'strSql = "DELETE FROM OEI_LINCMT WHERE Ord_Guid = '" & Ord_Guid & "'"
                'db.Execute(strSql)

                'Insert comments into oei_lincmt
                For Each drRow As DataRow In dt.Rows
                    Dim blockSet As List(Of String) = splitStringIntoBlocks(drRow.Item("SPEC_INSTRUCTION"), 40)

                    For Each block As String In blockSet
                        strSql = "INSERT INTO OEI_LINCMT (Ord_Guid, Line_Seq_No, Cmt, Item_Cd) VALUES " & _
                                 " ('" & Ord_Guid & "', " & Line_Seq_No & ", '" & Replace(block, "'", "''") & "', '" & Item_Cd & "')"
                        db.Execute(strSql)
                    Next
                Next
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub tsbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExportToMacola.Click

        Try

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String = _
            "SELECT CASE WHEN EXPORTTS IS NULL THEN 1 ELSE 0 END AS ExportCheck " & _
            "FROM   OEI_OrdHdr WITH (Nolock) " & _
            "WHERE  Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                If dt.Rows(0).Item("ExportCheck") = 0 Then Exit Sub
            End If

            m_oOrder.CancelExport = False

            Call SaveUncommited()

            If m_oOrder.OrdLines.Count = 0 Then Throw New OEException(OEError.Order_Not_Exported_No_Item_Lines)

            Dim oLine As cOrdLine
            Dim intUndefinedRouteCount As Integer = 0
            Dim intInvalidImprintCount As Integer = 0

            For Each oLine In m_oOrder.OrdLines

                'if stocked flag, nor route and not 88 and OEC or 44 and PRC, add to undefined counter. Exclude 88RANDOMSAMP (currently 'N', 88 and OEC)
                If (oLine.Stocked_Fg = "Y" Or oLine.Item_No.ToString.Trim() = "88RANDOMSAMP" Or oLine.Item_No.ToString.Trim() = "00STOCK0000") _
                    And oLine.RouteID = 0 And Not _
                   ( _
                       (oLine.Prod_Cat = "PRC" And oLine.Item_No.ToString.Substring(0, 2) = "44") _
                       Or _
                       (oLine.Prod_Cat = "OEC" And oLine.Item_No.ToString.Substring(0, 2) = "88" And oLine.Item_No.ToString.Trim() <> "88RANDOMSAMP") _
                   ) _
                Then
                    intUndefinedRouteCount += 1
                End If

            Next

            'check for invalid imprint counts
            For Each oLine In m_oOrder.OrdLines
                If oLine.RouteID <> 0 And oLine.Stocked_Fg = "Y" Then 'if a valid route is chosen perform check on whether
                    strSql = "  Select RouteID FROM EXACT_TRAVELER_ROUTE " _
                             & " Where RouteDescription not like '%random%' and RouteDescription not like '%no imprint%' " _
                             & " AND RouteDescription <> 'completed' AND RouteDescription not like '%SAMPLE Center%' " _
                             & " AND RouteId = " & oLine.RouteID & " And Active = 1"

                    dt = db.DataTable(strSql)
                    If dt.Rows.Count <> 0 Then 'if a valid route now check if an imprint count was set
                        strSql = "  SELECT (Num_Impr_1 + Num_Impr_2 + Num_Impr_3) as num_imprints FROM OEI_EXTRA_FIELDS " _
                            & " Where Ord_guid = '" & oLine.Ord_Guid & "' AND Item_Guid = '" & oLine.Item_Guid & "'"

                        dt = db.DataTable(strSql)
                        If dt.Rows.Count = 0 OrElse dt.Rows(0).Item("num_imprints") = 0 OrElse dt.Rows(0).Item("num_imprints").ToString = "" Then
                            intInvalidImprintCount += 1 'increment invalid imprint count and exit for (can change to count all incorrect in future if needed)
                            Exit For
                        End If
                    End If
                End If
            Next

            'check that cmp codes match (when changing an order company info)
            If m_oOrder.Ordhead.Cus_No.ToString.Trim() <> m_oOrder.Ordhead.Customer.cmp_code.ToString.Trim() _
                    Or UcContacts1.txtCusNo.Text.Trim() <> m_oOrder.Ordhead.Cus_No.ToString.Trim() _
                    Or m_oOrder.Ordhead.Cus_No.ToString.Trim() <> UcOrder.txtCus_No.Text.Trim() Then
                MsgBox("Please review and ensure that the company code entered corresponds to the provided billing address and assigned contact information")
                m_oOrder.CancelExport = True
                Exit Sub
            End If

            Dim mbrResult As New MsgBoxResult
            Dim mbrResult2 As New MsgBoxResult

            Dim currentDate As DateTime = DateTime.Parse(Date.Now.ToShortDateString)
            Dim enteredDate As DateTime = DateTime.Parse(UcOrder.txtOrd_Dt.Text.ToString)

            If intUndefinedRouteCount <> 0 Then
                mbrResult = MsgBox("Some stocked items have undefined routes. Do you wish to continue ?", MsgBoxStyle.YesNo)
            End If
            If intInvalidImprintCount <> 0 Then
                mbrResult = MsgBox("Some stocked items have routes but no imprint counts assigned to them. Do you wish to continue ?", MsgBoxStyle.YesNo)
            End If

            'export continue if valid
            If intUndefinedRouteCount = 0 And intInvalidImprintCount = 0 Then
                mbrResult = MsgBox("Order is ready to export. Do you wish to continue ?", MsgBoxStyle.YesNo)
            End If

            If currentDate > enteredDate Or currentDate < enteredDate Then
                mbrResult2 = MsgBox("The order date entered is not the same as today's date. Proceed ?", MsgBoxStyle.YesNo)
            Else
                mbrResult2 = MsgBoxResult.Yes
            End If

            If mbrResult = MsgBoxResult.Yes And mbrResult2 = MsgBoxResult.Yes Then ' Or intUndefinedRouteCount = 0 Then

                Call m_oOrder.SendOrderAck()

                If Not (m_oOrder.CancelExport) Then

                    Call Save()

                    Call m_oOrder.Ordhead.UnsetOpenTS()

                    Call m_oOrder.SetExportTS()

                    Call m_oOrder.CreateExcelFile()

                    'insert line sequence comments for order ack
                    For Each oLine In m_oOrder.OrdLines
                        Call insertLineSeqCmt(oLine.Ord_Guid, oLine.Line_Seq_No, oLine.Item_Cd, oLine.Item_No)
                    Next

                    'If g_User.XML_Access Then Call m_oOrder.XML_Export(m_oOrder.Ordhead.OEI_Ord_No)

                End If

            End If

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub tsbNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbNew.Click

        Try
            Call SaveUncommited()

            m_oOrder.Ordhead.UnsetOpenTS()

            If m_oOrder.OrdLines.Count = 0 Then

                Call m_oOrder.UnsetImportedEDIOrder()

                Call m_oOrder.DeleteEmptyOrderIfNew()

                Call Reset()

                Call RefreshImprint(0)

            Else

                Call Save()

                Call Reset()

                Call RefreshImprint(0)

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private event handlers ###########################################"

    Private Sub Validate_Cus_No_Exist() Handles UcContacts1.VisibleChanged, UcCreditInfo1.VisibleChanged, UcExtra1.VisibleChanged, UcHeader1.VisibleChanged, UcDocument1.VisibleChanged, UcImprint1.VisibleChanged, UcOrderTotal1.VisibleChanged, UcSalesperson1.VisibleChanged, UcTaxes1.VisibleChanged

        Dim tcControl As TabControl = sstOrder

        Try

            If m_oOrder Is Nothing Then
                m_oOrder = New cOrder()
                UcOrder.Fill()
            End If

            If tcControl.SelectedIndex <> 0 And (Trim(m_oOrder.Ordhead.Cus_No) = "" Or m_oOrder.Ordhead.Cus_No Is Nothing) Then
                sstOrder.SelectTab(0)
                Exit Sub
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Public order management procedures ###############################"

    Private Sub Reset()

        Try

            If m_oOrder.OrdLines.Count <> 0 Then
                dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Location)
                dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Item_No)
            End If

            m_oOrder = Nothing
            m_oOrder = New cOrder()
            g_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID)

            UcOrder.Reset()
            UcContacts1.Reset()
            UcCreditInfo1.Reset()
            UcExtra1.Reset()
            UcHeader1.Reset()
            'UcImprint1.Reset()
            UcOrderTotal1.Reset()
            UcSalesperson1.Reset()
            UcTaxes1.Reset()

            intCopyRouteID = 0 'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            strCopyRoute = String.Empty ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
            intCopyProductProof = 0
            intCopyAutoCompleteReship = 0

            oCopyFrom = Nothing

            For iPos As Integer = (dgvItems.Rows.Count - 1) To 0 Step -1
                dgvItems.Rows.RemoveAt(iPos)
            Next

            Dim oComboboxColumn As New DataGridViewComboBoxColumn

            oComboboxColumn = DirectCast(dgvItems.Columns(Columns.Route), DataGridViewComboBoxColumn)
            oComboboxColumn.DataSource = m_oOrder.GetCompleteCboRouteList()

            'cboIndustry.DataSource = m_oOrder.GetCompleteCboIndustryList()

            Call ProductLineInsert()
            sstOrder.SelectTab(0)

            UcOrder.txtOrd_No.Focus()

            dgvHistory.Refresh()

            lblWhite_Glove.Visible = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdCopyLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyLine.Click

        Try
            'Call ProductLineCopy()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdInsertLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsertLine.Click

        Try
            blnInserting = True
            Call ProductLineInsert()
            g_oOrdline = New cOrdLine(False) ' m_oOrder)
            blnInserting = False

            'Call RefreshImprint(dgvItems.CurrentRow.Index)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdDeleteLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteLine.Click

        Try
            dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Item_No)

            Call ProductLineDelete()

            Call RefreshImprint(0)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdRestart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRestart.Click

        Try
            Call ProductLineRestart(True)
            g_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID)

            Call RefreshImprint(0)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Sub cmdGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrevCharges.Click

    '    Try

    '        Dim oOrderSelect As New frmProductLineEntry() ' m_oOrder)
    '        oOrderSelect.Customer_No = m_oOrder.Ordhead.Cus_No
    '        oOrderSelect.Currency = m_oOrder.Ordhead.Curr_Cd ' "USD"
    '        oOrderSelect.Item_No = "CHARGESCALCULATION"
    '        oOrderSelect.Loc = m_oOrder.Ordhead.Mfg_Loc

    '        'oOrderSelect.ProductLineEntry.LoadProductGrid()
    '        oOrderSelect.ShowDialog()
    '        'If oOrderSelect.ProductStatus = 1 Then
    '        '    Call InsertProductsInOrder(oOrderSelect)
    '        'End If
    '        'If dgvItems.Rows.Count > 2 Then
    '        '    dsDataSet.Tables(0).Rows.RemoveAt(e.RowIndex)
    '        'End If

    '        'm_Create_Imprint = False
    '        'm_Create_Kit = False
    '        'm_Create_Traveler = False
    '        'm_SaveToDB = False

    '        'Select Case cboMoreOptions.Text
    '        '    Case "Import items from order"

    '        '    Case "Search item"

    '        'End Select

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Private Sub SaveOrderLines()

        Try

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ShowOrderLines()

        Try

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ProductLineDelete(Optional ByVal pblnAsk As Boolean = True)

        Try
            If dgvItems.Rows.Count = 0 Then Exit Sub
            If dgvItems.Rows(0).Cells(Columns.Item_No).Value.ToString = "" Then Exit Sub

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            ' ask if it is ok to delete. If not OK, then exit sub right away
            If pblnAsk Then
                If Trim(dgvItems.CurrentRow.Cells(Columns.Item_No).Value.ToString) <> "" Then
                    Dim mbrOkToDelete As MsgBoxResult
                    mbrOkToDelete = MsgBox("OK to delete ?", MsgBoxStyle.YesNo, "Order entry interface")

                    If mbrOkToDelete = MsgBoxResult.No Then Exit Sub
                End If
            End If

            Dim iPos As Integer = dgvItems.CurrentRow.Index
            Dim oOrdline As cOrdLine = Nothing

            ' Delete the line(s) from the database
            'If m_oOrder.OrdLines.Count > iPos Then
            If m_oOrder.OrdLines.Contains(dgvItems.CurrentRow.Cells(Columns.Item_Guid).Value.ToString) Then
                oOrdline = m_oOrder.OrdLines(dgvItems.CurrentRow.Cells(Columns.Item_Guid).Value.ToString)
                m_oOrder.DeleteOrderLine(oOrdline)

                'End If

                ' Delete the line(s) from the grid
                blnDeleting = True
                If oOrdline.End_Item_Cd = "K" Then
                    Dim dr As New DataGridViewRow
                    ' Browse every line to remove all kit components and the main item
                    For iPos = dgvItems.Rows.Count - 1 To 0 Step -1
                        dr = dgvItems.Rows(iPos)
                        If Not (dr.Cells(Columns.Kit_Item_Guid) Is Nothing Or dr.Cells(Columns.Kit_Item_Guid).Value.Equals(DBNull.Value) Or oOrdline.Item_Guid Is Nothing Or oOrdline.Item_Guid.Equals(DBNull.Value)) Then
                            If dr.Cells(Columns.Kit_Item_Guid).Value = oOrdline.Item_Guid Or dr.Cells(Columns.Item_Guid).Value = oOrdline.Item_Guid Then
                                dgvItems.Rows.RemoveAt(iPos)
                            End If
                        End If
                    Next iPos
                Else
                    dgvItems.Rows.RemoveAt(iPos)
                End If
            Else
                dgvItems.Rows.RemoveAt(iPos)
            End If

            'g_oOrdline = Nothing

            If dgvItems.Rows.Count > 0 Then
                If iPos >= dgvItems.Rows.Count Then
                    dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Item_No)
                    dgvItems.CurrentCell = dgvItems.Rows(iPos - 1).Cells(Columns.Item_No)
                Else
                    If iPos < 0 Then iPos = 0
                    dgvItems.CurrentCell = dgvItems.Rows(iPos).Cells(Columns.Item_No)
                End If
            End If

            oOrdline = Nothing

            If m_oOrder.OrdLines.Count > 0 Then
                If m_oOrder.OrdLines.Contains(dgvItems.CurrentRow.Cells(Columns.Item_Guid).Value.ToString) Then
                    g_oOrdline = m_oOrder.OrdLines(dgvItems.CurrentRow.Cells(Columns.Item_Guid).Value.ToString)
                Else
                    g_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID)
                End If
            Else
                g_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID)
            End If

            dgvItems.Focus()

            blnDeleting = False

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            blnDeleting = False
        End Try

    End Sub

    Private Sub ProductLineInsert()

        Try
            'g_oOrdline = Nothing

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If dgvItems.Rows.Count = 0 Then

                g_oOrdline = Nothing
                Call LineInsert()
            ElseIf Not (dgvItems.Rows(dgvItems.Rows.Count - 1).Cells(Columns.Item_No).Value.Equals(DBNull.Value)) Then
                If Not Trim(dgvItems.Rows(dgvItems.Rows.Count - 1).Cells(Columns.Item_No).Value).Equals("") Then
                    Call LineInsert()
                    g_oOrdline = Nothing
                Else
                    dgvItems.CurrentCell = dgvItems.Rows(dgvItems.Rows.Count - 1).Cells(Columns.Item_No)
                End If
            Else
                dgvItems.CurrentCell = dgvItems.Rows(dgvItems.Rows.Count - 1).Cells(Columns.Item_No)
            End If

            'g_oOrdline = New cOrdLine(m_oOrder)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LineInsert()

        Try
            Dim drNewRow As DataRow
            dgvItems.AllowUserToAddRows = True

            drNewRow = dtDataTable.NewRow

            dtDataTable.Rows.Add(drNewRow)

            dgvItems.AllowUserToAddRows = False
            dgvItems.CurrentCell = dgvItems.Rows(dgvItems.Rows.Count - 1).Cells(Columns.Item_No)

            dgvItems.Focus()
            'dgvItems.BeginEdit(True)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ProductLineRestart(ByVal pblnAsk As Boolean)

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If dgvItems.Rows.Count = 0 Then
                Call ProductLineInsert()
                Exit Sub
            End If

            If dgvItems.Rows(0).Cells(Columns.Item_No).Value.ToString = "" Then Exit Sub

            If pblnAsk Then

                Dim mbrOkToRestart As MsgBoxResult
                mbrOkToRestart = MsgBox("OK to restart ? Every item in your order will be deleted.", MsgBoxStyle.YesNo, "Order entry interface")

                If mbrOkToRestart = MsgBoxResult.No Then Exit Sub

                While dgvItems.Rows.Count > 0
                    If dgvItems.Rows(0).Cells(Columns.Item_No).Value.ToString <> "" Then

                        dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Location)
                        dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Item_No)
                        Call ProductLineDelete(False)
                    Else
                        dgvItems.Rows.RemoveAt(0)
                    End If

                End While

            Else

                For iPos As Integer = (dgvItems.Rows.Count - 1) To 0 Step -1
                    dgvItems.Rows.RemoveAt(iPos)
                Next

                For Each oOrdline As cOrdLine In m_oOrder.OrdLines
                    oOrdline.Delete()
                Next

            End If

            m_oOrder.OrdLines.Clear()

            Call ProductLineInsert()

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Set_Qty_On_Hand_Kits(ByVal pintRow As Integer) ' , ByVal pintCol As Integer)

        Try
            If blnLoading Then Exit Sub

            Dim iRow As Integer = dgvItems.CurrentCell.RowIndex
            Dim iCol As Integer = dgvItems.CurrentCell.ColumnIndex

            Dim strItemGuid As String

            If dgvItems.Rows(pintRow).Cells(Columns.Item_Guid).Value.Equals(DBNull.Value) Then Exit Sub

            strItemGuid = dgvItems.Rows(pintRow).Cells(Columns.Item_Guid).Value

            For iPos As Integer = iRow + 1 To dgvItems.Rows.Count - 1

                If dgvItems.Rows(iPos).Cells(Columns.Kit_Item_Guid).Value = strItemGuid Then
                    If m_oOrder.OrdLines.Contains(dgvItems.Rows(iPos).Cells(Columns.Item_Guid).Value) Then
                        Dim oOrdline As cOrdLine = m_oOrder.OrdLines.Item(dgvItems.Rows(iPos).Cells(Columns.Item_Guid).Value)
                        oOrdline.Loc = dgvItems.Rows(pintRow).Cells(Columns.Location).Value
                        dgvItems.Rows(iPos).Cells(Columns.Qty_On_Hand).Value = oOrdline.Qty_On_Hand
                    Else
                        dgvItems.Rows(iPos).Cells(Columns.Qty_On_Hand).Value = 0
                    End If
                Else
                    iPos = dgvItems.Rows.Count + 1
                End If

            Next iPos

            dgvItems.CurrentCell = dgvItems.Rows(iRow).Cells(iCol)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub LoadPreOrder()

        Try
            Dim intLastpos As Integer = dgvItems.CurrentRow.Index

            dgvItems.AllowUserToAddRows = False

            blnLoading = True

            cmdInsertLine.Focus()

            'For lPos As Integer = dgvItems.Rows.Count To 1 Step -1
            '    dgvItems.Rows.RemoveAt(lPos - 1)
            'Next

            If dgvItems.Rows.Count > 0 Then
                dgvItems.Rows.RemoveAt(dgvItems.Rows.Count - 1)
            End If

            dgvItems.AllowUserToAddRows = False

            'm_oOrder.OrdLines.Clear()
            'If m_oOrder.OrdLines.Count = 0 Then
            m_oOrder.OrderLinesLoad(OrderSourceEnum.OEInterface)
            'End If

            For lpos As Integer = 1 To m_oOrder.OrdLines.Count
                Dim oOrdline As cOrdLine = m_oOrder.OrdLines(lpos)
                Dim blnFound As Boolean = False
                If dgvItems.Rows.Count > 0 Then
                    For lPos2 As Integer = 0 To dgvItems.Rows.Count - 1
                        If oOrdline.Item_Guid = dgvItems.Rows(lPos2).Cells(Columns.Item_Guid).Value.ToString Then
                            blnFound = True
                            lPos2 = dgvItems.Rows.Count + 1
                        End If
                    Next lPos2
                End If

                If Not blnFound Then
                    Call LoadPreOrderInsertLine(oOrdline)
                    oOrdline = Nothing
                End If

            Next

            'For LPOS As Integer = 1 To m_oOrder.OrdLines.Count
            For LPOS As Integer = 1 To dgvItems.Rows.Count ' m_oOrder.OrdLines.Count
                Dim oOrdline As cOrdLine = m_oOrder.OrdLines(LPOS)
                dgvItems.Rows(LPOS - 1).Cells(Columns.Image).Value = oOrdline.Image

                If oOrdline.m_OEI_W_Pixels < 400 Then
                    Dim dblHeight As Double
                    dblHeight = IIf(400 / oOrdline.m_OEI_W_Pixels > 5.5, 5.5, 400 / oOrdline.m_OEI_W_Pixels)

                    If chkLowRow.Checked Then
                        dgvItems.Rows(LPOS - 1).Height = 22
                    Else
                        dgvItems.Rows(LPOS - 1).Height = 120 / dblHeight
                    End If

                    'dgvItems.Rows(LPOS - 1).Height = 120 / dblHeight
                    dgvItems.Columns(Columns.Image).Width = IIf(oOrdline.Image.Width > dgvItems.Rows(LPOS - 1).Cells(Columns.Image).Size.Width, oOrdline.Image.Width, dgvItems.Rows(LPOS - 1).Cells(Columns.Image).Size.Width)

                Else
                    'Call ResizeImage(dgvItems.CurrentRow.Index, g_oOrdline.Image)
                    Call ResizeImage(LPOS - 1, oOrdline.Image)
                End If

                dgvItems.Rows(LPOS - 1).Cells(Columns.PictureHeight).Value = dgvItems.Rows(LPOS - 1).Height

               Call execThermoSimulation()
                'dgvItems.BeginInvoke(New Del_SetProductionAvailabilityColor(AddressOf oOrdline.SetProductionAvailabilityColor), dgvItems.Rows(LPOS - 1).Cells(Columns.Item_No)) ' , dgvItems.CurrentCell.ColumnIndex)

                'g_oOrdline.SetProductionAvailabilityColor(dgvItems.Rows(LPOS - 1).Cells(Columns.Item_No)) ' , dgvItems.CurrentCell.ColumnIndex)

                'ResizeImage(LPOS - 1, oOrdline.Image)
                'call m_oOrder.ChangeCboRouteList("All")
            Next

            'Call SetRouteCombo(dgvItems.Columns(Columns.Route), "All")

            m_oOrder.Reload = False
            blnLoading = False

            dgvItems.AllowUserToAddRows = False
            If dgvItems.Rows.Count - 1 > intLastpos Then
                dgvItems.CurrentCell = dgvItems.Rows(intLastpos).Cells(Columns.Item_No)
            End If
            'dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Item_No)
            dgvItems.Focus()

            ignoreImprintMsgCheck = False
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub LoadPreOrderInsertLine(ByRef pOrdLine As cOrdLine)

        ' We have to do this as we can't had a row in a binded DataGridView
        Dim drNewRow As DataRow

        Try
            blnFill = True

            'drNewRow = dsDataSet.Tables(0).NewRow()
            drNewRow = dtDataTable.NewRow()

            'If pOrdLine.Kit_Element Then
            '    drNewRow.Item(Columns.Item_No) = pOrdLine.Item_No
            '    drNewRow.Item(Columns.Qty_On_Hand) = pOrdLine.Qty_On_Hand
            '    drNewRow.Item(Columns.Item_Desc_1) = pOrdLine.Item_Desc_1
            '    drNewRow.Item(Columns.Item_Desc_2) = pOrdLine.Item_Desc_2
            '    drNewRow.Item(Columns.Route) = pOrdLine.Route
            'Else
            If pOrdLine.Kit_Component Then
                drNewRow.Item(Columns.Item_No) = g_Comp_Pref & Trim(pOrdLine.Item_No)
            Else
                drNewRow.Item(Columns.Item_No) = pOrdLine.Item_No
            End If
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Location) = pOrdLine.Loc
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Unit_Price) = pOrdLine.Unit_Price
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Qty_Ordered) = pOrdLine.Qty_Ordered
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Qty_Overpick) = pOrdLine.Qty_Overpick
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Qty_To_Ship) = pOrdLine.Qty_To_Ship
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Qty_Inventory) = pOrdLine.Qty_Inventory
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Qty_Allocated) = pOrdLine.Qty_Allocated
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Qty_Prev_To_Ship) = pOrdLine.Qty_Prev_To_Ship
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Qty_Prev_BkOrd) = pOrdLine.Qty_Prev_Bkord
            drNewRow.Item(Columns.Qty_On_Hand) = pOrdLine.Qty_On_Hand
            drNewRow.Item(Columns.Extra_9) = pOrdLine.Extra_9
            drNewRow.Item(Columns.Extra_10) = pOrdLine.Extra_10
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Discount_Pct) = pOrdLine.Discount_Pct
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Calc_Price) = pOrdLine.Calc_Price
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Request_Dt) = pOrdLine.Request_Dt
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Promise_Dt) = pOrdLine.Promise_Dt
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Req_Ship_Dt) = pOrdLine.Req_Ship_Dt
            drNewRow.Item(Columns.Item_Desc_1) = pOrdLine.Item_Desc_1
            drNewRow.Item(Columns.Item_Desc_2) = pOrdLine.Item_Desc_2
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Tax_Sched) = pOrdLine.Tax_Sched
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Revision_No) = pOrdLine.Revision_No
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.UOM) = pOrdLine.Uom
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.BkOrd_Fg) = pOrdLine.BkOrd_Fg_Bln
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.Recalc_SW) = pOrdLine.Recalc_Sw_Bln
            drNewRow.Item(Columns.Extra_1) = pOrdLine.Extra_1_Bln

            If pOrdLine.RouteID <> 0 Then
                drNewRow.Item(Columns.Route) = pOrdLine.Route
                drNewRow.Item(Columns.RouteID) = pOrdLine.RouteID
            End If
            If Not (pOrdLine.End_Item_Cd = "K") Then drNewRow.Item(Columns.ProductProof) = pOrdLine.ProductProof_Bln
            If Not (pOrdLine.Kit_Component) Then drNewRow.Item(Columns.AutoCompleteReship) = pOrdLine.AutoCompleteReship_Bln

            ' Show imprint columns
            drNewRow.Item(Columns.Imprint) = pOrdLine.ImprintLine.Imprint
            drNewRow.Item(Columns.End_User) = pOrdLine.ImprintLine.End_User
            drNewRow.Item(Columns.Imprint_Location) = pOrdLine.ImprintLine.Location
            drNewRow.Item(Columns.Imprint_Color) = pOrdLine.ImprintLine.Color
            drNewRow.Item(Columns.Num_Impr_1) = IIf(pOrdLine.ImprintLine.Num_Impr_1 <> 0, pOrdLine.ImprintLine.Num_Impr_1, DBNull.Value)
            drNewRow.Item(Columns.Num_Impr_2) = IIf(pOrdLine.ImprintLine.Num_Impr_2 <> 0, pOrdLine.ImprintLine.Num_Impr_2, DBNull.Value)
            drNewRow.Item(Columns.Num_Impr_3) = IIf(pOrdLine.ImprintLine.Num_Impr_3 <> 0, pOrdLine.ImprintLine.Num_Impr_3, DBNull.Value)
            'If pOrdLine.ImprintLine.Num_Impr_1 <> 0 Then drNewRow.Item(Columns.Num_Impr_1) = pOrdLine.ImprintLine.Num_Impr_1
            'If pOrdLine.ImprintLine.Num_Impr_2 <> 0 Then drNewRow.Item(Columns.Num_Impr_2) = pOrdLine.ImprintLine.Num_Impr_2
            'If pOrdLine.ImprintLine.Num_Impr_3 <> 0 Then drNewRow.Item(Columns.Num_Impr_3) = pOrdLine.ImprintLine.Num_Impr_3
            drNewRow.Item(Columns.Packaging) = pOrdLine.ImprintLine.Packaging
            drNewRow.Item(Columns.Refill) = pOrdLine.ImprintLine.Refill
            drNewRow.Item(Columns.Laser_Setup) = pOrdLine.ImprintLine.Laser_Setup
            drNewRow.Item(Columns.Industry) = pOrdLine.ImprintLine.Industry
            drNewRow.Item(Columns.Comments) = pOrdLine.ImprintLine.Comments
            drNewRow.Item(Columns.Special_Comments) = pOrdLine.ImprintLine.Special_Comments
            drNewRow.Item(Columns.Printer_Comment) = pOrdLine.ImprintLine.Printer_Comment
            drNewRow.Item(Columns.Packer_Comment) = pOrdLine.ImprintLine.Packer_Comment
            drNewRow.Item(Columns.Printer_Instructions) = pOrdLine.ImprintLine.Printer_Instructions
            drNewRow.Item(Columns.Packer_Instructions) = pOrdLine.ImprintLine.Packer_Instructions
            drNewRow.Item(Columns.Repeat_From_ID) = pOrdLine.ImprintLine.Repeat_From_ID
            drNewRow.Item(Columns.Item_Guid) = pOrdLine.Item_Guid
            drNewRow.Item(Columns.Kit_Item_Guid) = pOrdLine.Kit_Item_Guid
            drNewRow.Item(Columns.Item_Cd) = pOrdLine.Item_Cd 'new
            drNewRow.Item(Columns.Mix_Group) = pOrdLine.Mix_Group 'new
            'If pOrdLine.Kit.IsAKit Or pOrdLine.Kit_Component Then
            '    drNewRow.Item(Columns.Kit_Item_Guid) = pOrdLine.Comp_Item_Guid
            'Else
            '    drNewRow.Item(Columns.Kit_Item_Guid) = ""
            'End If
            'End If

            If pOrdLine.Kit.IsAKit Then
                drNewRow.Item(Columns.Kit_Item_Guid) = pOrdLine.Item_Guid
            ElseIf pOrdLine.Kit_Component Then
                drNewRow.Item(Columns.Kit_Item_Guid) = pOrdLine.Kit_Item_Guid
                'drNewRow.Item(Columns.Kit_Item_Guid) = pOrdLine.Comp_Item_Guid
            Else
                drNewRow.Item(Columns.Kit_Item_Guid) = ""
            End If

            'dsDataSet.Tables(0).Rows.Add(drNewRow)
            dtDataTable.Rows.Add(drNewRow)

            blnFill = False
            ''dgvItems.Rows(dsDataSet.Tables(0).Rows.Count - 1).Cells(Columns.Image).Value = pOrdLine.Image
            'dgvItems.Rows(3).Cells(Columns.Image).Value = pOrdLine.Image

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ResizeImage(ByVal pintRow As Integer, ByVal pImage As Image)

        Try
            ' in case there is no image
            If pImage Is Nothing Then Exit Sub

            Dim myRow As New DataGridViewRow
            myRow = dgvItems.Rows(pintRow)

            myRow.Height = pImage.Height ' IIf(oResizedImage.Height > 25, oResizedImage.Height + 5, 25)
            If chkLowRow.Checked Then
                myRow.Cells(Columns.PictureHeight).Value = 22
                dgvItems.Rows(myRow.Index).Height = 22
            Else
                myRow.Cells(Columns.PictureHeight).Value = pImage.Height
                dgvItems.Rows(myRow.Index).Height = pImage.Height
            End If
            'dgvItems.Columns(Columns.Image).Width = IIf(dgvItems.Columns(Columns.Image).Width > myRow.Cells(Columns.Image).Size.Width, dgvItems.Columns(Columns.Image).Width, myRow.Cells(Columns.Image).Size.Width)
            dgvItems.Columns(Columns.Image).Width = IIf(pImage.Width > myRow.Cells(Columns.Image).Size.Width, pImage.Width, myRow.Cells(Columns.Image).Size.Width)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ResizeImageFromMaster(ByVal pintRow As Integer, ByVal pImage As Image, Optional preservePictureQuality As Boolean = True)

        Dim newWidth, newHeight As Integer
        Dim myRow As New DataGridViewRow
        Dim size As Size

        'set size for images in the dgv
        size.Height = 125
        size.Width = 125

        Try
            'preserve picture quality and get percentage difference of original image
            If preservePictureQuality Then
                Dim originalWidth As Integer = pImage.Width
                Dim originalHeight As Integer = pImage.Height
                Dim percentWidth As Single = CSng(size.Width) / CSng(originalWidth)
                Dim percentHeight As Single = CSng(size.Height) / CSng(originalHeight)
                Dim percent As Single = If(percentHeight > percentWidth, percentWidth, percentHeight)
                newWidth = CInt(originalWidth * percent)
                newHeight = CInt(originalHeight * percent)
            Else
                newWidth = size.Width
                newHeight = size.Height
            End If

            'create bitmap using new dimensions
            Dim pNewImage As Image = New Bitmap(newWidth, newHeight)
            Using g As Graphics = Graphics.FromImage(pNewImage)
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.DrawImage(pImage, 0, 0, newWidth, newHeight)
            End Using

            Dim ms As MemoryStream = New MemoryStream()
            pNewImage.Save(ms, Imaging.ImageFormat.Jpeg)
            pNewImage.Dispose()

            myRow = dgvItems.Rows(pintRow)

            myRow.Height = pImage.Height ' IIf(oResizedImage.Height > 25, oResizedImage.Height + 5, 25)
            If chkLowRow.Checked Then
                myRow.Cells(Columns.PictureHeight).Value = 22
                dgvItems.Rows(myRow.Index).Height = 22
            Else
                myRow.Cells(Columns.PictureHeight).Value = pImage.Height
                dgvItems.Rows(myRow.Index).Height = pImage.Height
            End If

            'dgvItems.Columns(Columns.Image).Width = IIf(dgvItems.Columns(Columns.Image).Width > myRow.Cells(Columns.Image).Size.Width, dgvItems.Columns(Columns.Image).Width, myRow.Cells(Columns.Image).Size.Width)
            dgvItems.Columns(Columns.Image).Width = IIf(pImage.Width > myRow.Cells(Columns.Image).Size.Width, pImage.Width, myRow.Cells(Columns.Image).Size.Width)

        Catch ex As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Sub

    Public Sub ResizeRow(ByVal pintRow As Integer)

        Try
            If chkLowRow.Checked Then
                dgvItems.Rows(pintRow).Height = dgvItems.Rows(pintRow).Cells(Columns.PictureHeight).Value
            Else
                dgvItems.Rows(pintRow).Height = dgvItems.Rows(pintRow).Cells(Columns.PictureHeight).Value
            End If
            'dgvItems.Columns(Columns.Image).Width = IIf(dgvItems.Columns(Columns.Image).Width > myRow.Cells(Columns.Image).Size.Width, dgvItems.Columns(Columns.Image).Width, myRow.Cells(Columns.Image).Size.Width)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Fill() ' ByRef pOrder As cOrder)

        Try
            blnLoading = True

            If Trim(m_oOrder.Ordhead.Ord_No).Length <> 0 Then
            Else
                Reset() ' pOrder)
                LoadProductGrid()
            End If

            'End If
            blnLoading = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub NewOrder(ByRef pOrder As cOrder)

        'lblAmt_Cus_Balance.Text = ""

    End Sub

    Public Sub Save()

        Try

            UcOrder.Save() ' m_oOrder)
            UcContacts1.Save()
            UcCreditInfo1.Save()
            UcExtra1.Save()
            UcHeader1.Save()
            'UcImprint1.ImprintSave()
            UcOrderTotal1.Save()
            UcSalesperson1.Save()
            UcTaxes1.Save()
            UcDocument1.Save()

            m_oOrder.SetHoldFgFromAvailCredit()

            'm_oOrder.OrdLines.Clear()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SaveOrderLine(ByRef poOrdline As cOrdLine)

        Try
            'If dgvItems.CurrentRow.Cells(Columns.Item_No).Equals(DBNull.Value) Then g_oOrdline = Nothing
            'If Trim(dgvItems.CurrentRow.Cells(Columns.Item_No).Value) = "" Then g_oOrdline = Nothing

            If blnDeleting Then Exit Sub

            If m_oOrder Is Nothing Then Exit Sub
            If poOrdline Is Nothing Then Exit Sub

            'If dgvItems.CurrentRow.Index >= m_oOrder.OrdLines.Count And Not (poOrdline.Item_No Is Nothing) Then ' Or (dgvItems.CurrentRow.Index = 0 And m_oOrder.OrdLines.Count = 0) Then
            '    'm_oOrder.OrdLines.Add(poOrdline, poOrdline.Item_Guid)
            '    If poOrdline.Kit_Component Then
            '        m_oOrder.OrdLines.Add(poOrdline, poOrdline.Item_Guid)
            '    Else
            m_oOrder.AddOrderLine(poOrdline)

            '        'm_oOrder.AddOrderLine(poOrdline)
            '        'ElseIf dgvItems.CurrentRow.Index <= m_oOrder.OrdLines.Count And Not (g_oOrdline.Item_No Is Nothing) Then
            '        '    If Trim(g_oOrdline.Item_No) <> "" Then g_oOrdline.Save()
            '        'Else
            '        '    If Not (g_oOrdline.Item_No Is Nothing) Then g_oOrdline.Save()
            '    End If
            'End If

            Dim iOrder_Item_Count As Integer = 0
            Dim dblOrder_Amt As Double = 0
            Dim dblShip_Amt As Double = 0
            Dim intOrder_Qty As Integer = 0
            Dim intShip_Qty As Integer = 0
            Dim dblConv_Curr_Amt As Double = 0
            Dim oOrdline As New cOrdLine(m_oOrder.Ordhead.Ord_GUID) ' m_oOrder)

            'For lPos As Integer = 1 To m_oOrder.OrdLines.Count
            For Each oOrdline In m_oOrder.OrdLines
                If Not (oOrdline.Kit_Component) Then iOrder_Item_Count = iOrder_Item_Count + 1

                'Dim oOrdline As cOrdLine = m_oOrder.OrdLines(lPos)
                dblOrder_Amt = dblOrder_Amt + Math.Round(((oOrdline.Qty_Ordered * oOrdline.Unit_Price) * (100 - oOrdline.Discount_Pct) / 100), 2)
                dblShip_Amt = dblShip_Amt + Math.Round(oOrdline.Calc_Price, 2)
                intOrder_Qty = intOrder_Qty + oOrdline.Qty_Ordered
                intShip_Qty = intShip_Qty + oOrdline.Qty_To_Ship
            Next 'lPos

            dblConv_Curr_Amt = dblShip_Amt * m_oOrder.Ordhead.Curr_Trx_Rt

            UcOrderTotal1.Order_Item_Count = iOrder_Item_Count
            UcOrderTotal1.Curr_Trx_Rt = m_oOrder.Ordhead.Curr_Trx_Rt
            UcOrderTotal1.Order_Amt = dblOrder_Amt
            UcOrderTotal1.Order_Qty = intOrder_Qty
            UcOrderTotal1.Ship_Amt = dblShip_Amt
            UcOrderTotal1.Ship_Qty = intShip_Qty

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SaveOrderLine()

        Try
            If blnDeleting Then Exit Sub

            If m_oOrder Is Nothing Then Exit Sub
            If g_oOrdline Is Nothing Then Exit Sub

            'If Not (m_oOrder.OrdLines.Contains(g_oOrdline.Item_Guid)) Then
            m_oOrder.AddOrderLine(g_oOrdline)
            'End If

            UcOrderTotal1.Order_Item_Count = m_oOrder.Total_Order_Item_Count
            UcOrderTotal1.Curr_Trx_Rt = m_oOrder.Total_Curr_Trx_Rt
            UcOrderTotal1.Order_Amt = m_oOrder.Total_Order_Amt
            UcOrderTotal1.Order_Qty = m_oOrder.Total_Order_Qty
            UcOrderTotal1.Ship_Amt = m_oOrder.Total_Ship_Amt
            UcOrderTotal1.Ship_Qty = m_oOrder.Total_Ship_Qty

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub UpdateOrderTotals()

        Try

            UcOrderTotal1.Order_Item_Count = m_oOrder.Total_Order_Item_Count
            UcOrderTotal1.Curr_Trx_Rt = m_oOrder.Total_Curr_Trx_Rt
            UcOrderTotal1.Order_Amt = m_oOrder.Total_Order_Amt
            UcOrderTotal1.Order_Qty = m_oOrder.Total_Order_Qty
            UcOrderTotal1.Ship_Amt = m_oOrder.Total_Ship_Amt
            UcOrderTotal1.Ship_Qty = m_oOrder.Total_Ship_Qty

            'experimental price adjuster (only when ignore is false to avoid loops) - disabled for now as requested. Will be manual trigger
            'If Not ignoreQtyMismatch And Not dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value = "" Then
            '    adjustUnitPricing()
            'End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private Notes control functions ##################################"

    ' Get Field_ID for SEARCHES_SQL table for a control name
    Private Function GetNotesElementByColumnIndex() As String

        GetNotesElementByColumnIndex = ""

        Try
            Select Case dgvItems.CurrentCell.ColumnIndex

                Case Columns.Item_No ' "txtOrder"
                    GetNotesElementByColumnIndex = "OE-ITEM"

                Case Columns.Location
                    GetNotesElementByColumnIndex = "OE-ITEM-LOC"

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function


    ' Search buttons click - open search window for element
    Private Sub Notes_Element_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _
        cmdNotes.Click

        'If blnForceFocus Then Exit Sub

        Dim btnElement As Button

        Try

            If TypeOf eventSender Is Button Then

                btnElement = New Button
                btnElement = DirectCast(eventSender, Button)

                'Dim txtElement As Object
                'txtElement = GetElementByControlName(btnElement.Name)

                Dim oNotes As New cNotes(GetNotesElementByColumnIndex(), dgvItems.CurrentCell.Value) ' txtElement.Text.ToString)
                oNotes = Nothing

            End If
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

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


#End Region


    Private Sub dgvItems_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.CellValueChanged

        Debug.Print("dgvItems_CellValueChanged")

        Try
            Debug.Print("Row: " & e.RowIndex.ToString & " Column: " & e.ColumnIndex.ToString & " Value: " & dgvItems.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString.ToUpper)
            Debug.Print("Row: " & e.RowIndex.ToString & " Column: 42 Value: " & dgvItems.Rows(e.RowIndex).Cells(42).Value.ToString.ToUpper)

            Select Case e.ColumnIndex

                Case Columns.Item_No
                    Dim s As String
                    s = dgvItems.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString.ToUpper
                    dgvItems.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = s

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvItems_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItems.GotFocus

        Debug.Print("dgvItems_GotFocus")
        Try

            If dgvItems.CurrentRow Is Nothing Then
                If dgvItems.Rows.Count > 0 Then
                    dgvItems.CurrentCell = dgvItems.Rows(0).Cells(3)
                Else
                    Exit Sub
                End If
            End If

            If dgvItems.CurrentRow.Cells(Columns.Item_Guid).Value Is Nothing Then
                g_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID) ' m_oOrder)
            ElseIf m_oOrder.OrdLines.Contains(dgvItems.CurrentRow.Cells(Columns.Item_Guid).Value.ToString) Then ' (dgvItems.CurrentRow.Index <= (m_oOrder.OrdLines.Count - 1)) Then
                g_oOrdline = m_oOrder.OrdLines(dgvItems.CurrentRow.Cells(Columns.Item_Guid).Value.ToString)
            Else
                g_oOrdline = New cOrdLine(m_oOrder.Ordhead.Ord_GUID) ' m_oOrder)
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub dgvItems_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvItems.KeyDown
        ' If e.KeyCode = Keys.A And Keys.ControlKey 

        Debug.Print("dgvItems_KeyDown")

        iCompteur += 1 ''MB+
        'Debug.Print((iCompteur).ToString & New Diagnostics.StackFrame(0).GetMethod.Name)
        'If Not dgvItems.CurrentRow Is Nothing Then Debug.Print("")

        Try
            If dgvItems.Rows.Count = 0 Then Exit Sub

            If e.KeyCode = Keys.Escape Then
                If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value.Equals(DBNull.Value) Then
                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then
                    cmdDeleteLine.PerformClick()
                    Exit Sub

                ElseIf dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then
                    cmdDeleteLine.PerformClick()
                    Exit Sub

                Else
                    e.SuppressKeyPress = True
                    Exit Sub ''Debug.Print("")
                End If

            ElseIf e.KeyCode = Keys.Return Then
                Call ProductLineInsert()
                Exit Sub
            End If

            Select Case dgvItems.CurrentCell.ColumnIndex

                Case Columns.Item_No

                    If e.KeyCode = Keys.F5 Then

                        If Trim(dgvItems.CurrentCell.Value.ToString) <> "" Then
                            Call PerformItemInfoClick()
                        End If

                    ElseIf e.KeyCode = Keys.F6 Then

                        If Trim(dgvItems.CurrentRow.Cells(Columns.Location).Value.ToString) <> "" Then
                            Call PerformItemLocQtyClick()
                        End If

                    ElseIf e.KeyCode = Keys.F7 Then

                        If Trim(dgvItems.CurrentCell.Value.ToString) = "" Then
                            Call PerformGridSearchClick()
                        End If

                    ElseIf e.KeyCode = Keys.F8 Then

                        If Trim(dgvItems.CurrentCell.Value.ToString) <> "" Then
                            Call PerformGridNotesClick()
                        End If

                    ElseIf e.Control Then

                        Select Case e.KeyCode
                            Case Keys.D
                                dgvItems.CurrentCell.Value = "44LOGO00DBS"
                                dgvItems.BeginEdit(True)
                            Case Keys.F
                                dgvItems.CurrentCell.Value = "44PAPR00PRF"
                                dgvItems.BeginEdit(True)
                            Case Keys.S
                                dgvItems.CurrentCell.Value = "44LOGO0001C"
                                dgvItems.BeginEdit(True)
                            Case Keys.Q
                                dgvItems.CurrentCell.Value = "44EXCT00QNT"
                                dgvItems.BeginEdit(True)
                            Case Keys.L
                                dgvItems.CurrentCell.Value = "44LOGO00LAZ"
                                dgvItems.BeginEdit(True)
                            Case Keys.P
                                dgvItems.CurrentCell.Value = "44P2P00CRG"
                                dgvItems.BeginEdit(True)
                            Case Keys.R
                                dgvItems.CurrentCell.Value = "44LAZR00OPM"
                                dgvItems.BeginEdit(True)
                            Case Keys.C
                                dgvItems.CurrentCell.Value = "44COLR00MCH"
                                dgvItems.BeginEdit(True)
                            Case Keys.M
                                dgvItems.CurrentCell.Value = "44LOGOMET1C"
                                dgvItems.BeginEdit(True)

                        End Select

                    End If

                Case Columns.Location

                    Select Case e.KeyCode

                        Case Keys.F5
                            Call PerformItemInfoClick()

                        Case Keys.F6
                            Call PerformItemLocQtyClick()

                        Case Keys.F7
                            If Trim(dgvItems.CurrentCell.Value.ToString) = "" Then
                                Call PerformGridSearchClick()
                            End If

                        Case Keys.F8
                            If Trim(dgvItems.CurrentCell.Value.ToString) <> "" Then
                                Call PerformGridNotesClick()
                            End If

                    End Select

                Case Columns.Qty_Ordered, Columns.Qty_To_Ship, Columns.Qty_Inventory, _
                     Columns.Qty_Allocated, Columns.Qty_On_Hand, Columns.Qty_Prev_To_Ship, Columns.Qty_Prev_BkOrd

                    Select Case e.KeyCode

                        Case Keys.F5
                            Call PerformItemInfoClick()

                        Case Keys.F6
                            Call PerformItemLocQtyClick()

                    End Select

                Case Columns.Extra_9, Columns.Extra_10

                    Select Case e.KeyCode

                        Case Keys.F3
                            Call PerformETAClick()

                        Case Keys.F5
                            Call PerformItemInfoClick()

                    End Select

                Case Columns.Unit_Price, Columns.Discount_Pct, Columns.Calc_Price

                    Select Case e.KeyCode

                        Case Keys.F4
                            Call PerformPriceBreaksClick()

                        Case Keys.F5
                            Call PerformItemInfoClick()

                    End Select

                Case Columns.Route

                    If e.Control Then
                        Select Case e.KeyCode
                            Case Keys.Space
                                dgvItems.CurrentCell.Value = " "
                        End Select
                    End If

                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    ''If e.Control Then

                    ''    Dim strValue As String = dgvItems.CurrentCell.Value.ToString

                    ''    Select Case e.KeyCode
                    ''        Case Keys.A
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "All", strValue)

                    ''        Case Keys.F
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Flat", strValue)

                    ''        Case Keys.S
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Silk", strValue)

                    ''        Case Keys.P
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Pad", strValue)

                    ''        Case Keys.T
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Silk/Flat", strValue)

                    ''        Case Keys.L
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Laser/Silk", strValue)

                    ''        Case Keys.D
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Silk/Pad", strValue)

                    ''        Case Keys.O
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "OS", strValue)

                    ''        Case Keys.M
                    ''            Call SetRouteCombo(dgvItems.Columns(Columns.Route), "Misc", strValue)

                    ''    End Select

                    ''End If

                Case Columns.Imprint

                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    '-- NO IMPRINT	 37765	 5%
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

                    ' MUST CHECK IF IT IS A KIT ITEM, THEN Qty_Ordered WILL BE 0 BUT WE NEED TO ALLOW !!!
                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Equals(DBNull.Value) Then Exit Sub
                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    '-- CAP			214884	33% + 10% NO VALUE
                    '--	BARREL		 95876	15%
                    '--	ENLARGED AREA55470	 7%
                    '--	STANDARD	 45962	 6%
                    '-- ENLARGED 	 33253	 5%
                    '-- FRONT 	 	 20321	 3%

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

                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Equals(DBNull.Value) Then Exit Sub
                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    '-- BLACK		175882	27% + 6% NO VALUE
                    '--	SILVER		 85516	13%	
                    '--	LASER		 71272	11%
                    '--	WHITE		 67916	10%
                    '-- DEBOSS		 24212	3%
                    '-- LASER/OXID	 22382	3%

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

                    'If dgvItems.CurrentRow.Cells(Columns.Qty_Ordered).Value = 0 Then Exit Sub

                    '-- BULK		229346	36% + 3% NO VALUE
                    '--	CELLO		 88021	13%
                    '--	STD		 	 66508	10%
                    '--	P100		 48272	 7%
                    '-- SHRINK WRAP	 12395	 3%

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

                Case Else

                    Select Case e.KeyCode
                        Case Keys.Return
                            Call ProductLineInsert()

                    End Select

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SetStatusBarPanelText(ByVal pintPanel As Integer, ByVal pstrText As String)

        Try
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

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#Region "Imprints #############################################################"

    Private blnImprintLoading As Boolean = False
    Private m_ComboFill As Boolean = False

    Private Sub SetColumnShortcuts()

        Try

            Select Case dgvItems.CurrentCell.ColumnIndex
                Case Columns.Item_No
                    If dgvItems.CurrentCell.Value.ToString = "" Then
                        _sbOrder_Panel1.Text = g_Search_Shortcut & g_Item_No_Shortcut
                    Else
                        _sbOrder_Panel1.Text = g_ItemInfo_Shortcut & g_ItemLocQty_Shortcut & g_Notes_Shortcut
                    End If

                Case Columns.Location
                    If dgvItems.CurrentCell.Value.ToString = "" Then
                        _sbOrder_Panel1.Text = g_Search_Shortcut
                    Else
                        _sbOrder_Panel1.Text = g_ItemInfo_Shortcut & g_ItemLocQty_Shortcut & g_Notes_Shortcut
                    End If

                Case Columns.Qty_Ordered, Columns.Qty_To_Ship, Columns.Qty_Inventory, _
                     Columns.Qty_Allocated, Columns.Qty_On_Hand, Columns.Qty_Prev_To_Ship, Columns.Qty_Prev_BkOrd
                    _sbOrder_Panel1.Text = g_ItemInfo_Shortcut & g_ItemLocQty_Shortcut

                Case Columns.Extra_9, Columns.Extra_10
                    _sbOrder_Panel1.Text = g_ETA_Shortcut & g_ItemInfo_Shortcut

                Case Columns.Unit_Price, Columns.Discount_Pct, Columns.Calc_Price
                    _sbOrder_Panel1.Text = g_PriceBreaks_Shortcut

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

                Case Columns.Route
                    _sbOrder_Panel1.Text = g_Route_Shortcut

                Case Else
                    _sbOrder_Panel1.Text = " "

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdClear_Imprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear_Imprint.Click

        Try
            Call ImprintReset(dgvItems.CurrentRow.Index)

            'UcImprint1.Reset()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ChangeWebPage(ByVal piRow As Integer)

        Try
            Dim strItemNo As String

            strItemNo = dgvItems.Rows(piRow).Cells(Columns.Item_No).Value.ToString
            strItemNo = strItemNo.Replace(g_Comp_Pref, "")

            Dim dt As New DataTable
            Dim db As New cDBA

            Dim strsql As String = _
            "SELECT * " & _
            "FROM   Exact_Traveler_Item_Description_Web WITH (NoLock) " & _
            "WHERE  Item_Num = '" & SqlCompliantString(Mid(strItemNo, 3, 6)) & "'"

            dt = db.DataTable(strsql)

            If dt.Rows.Count = 1 Then

                If strWebItem <> Mid(strItemNo, 3, 6) Then

                    strWebItem = Mid(strItemNo, 3, 6)
                    wbItem.Url = New Uri("http://www.spectorandco.ca/en/products/products.php?item_num=" & strWebItem)

                End If

            Else
                strsql = _
                "SELECT * " & _
                "FROM   Exact_Traveler_Item_Description_Web WITH (NoLock) " & _
                "WHERE  Item_Num = '" & SqlCompliantString(Mid(strItemNo, 3, 5)) & "'"

                dt = db.DataTable(strsql)

                If dt.Rows.Count = 1 Then

                    If strWebItem <> Mid(strItemNo, 3, 5) Then

                        strWebItem = Mid(strItemNo, 3, 5)
                        wbItem.Url = New Uri("http://www.spectorandco.ca/en/products/products.php?item_num=" & strWebItem)

                    End If

                Else

                    If strWebItem <> Mid(strItemNo, 3, 4) Then

                        strWebItem = Mid(strItemNo, 3, 4)
                        wbItem.Url = New Uri("http://www.spectorandco.ca/en/products/products.php?item_num=" & strWebItem)

                    End If

                End If

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub RefreshImprint(ByVal pintRow As Integer)

        Dim iPos As Integer

        Try
            iPos = 1

            If dgvItems.Rows.Count > 0 Then
                If Not (dgvItems.Rows(pintRow).Cells(Columns.Item_No).Value Is Nothing) Then
                    txtImprint_Item_No.Text = dgvItems.Rows(pintRow).Cells(Columns.Item_No).Value.ToString
                Else
                    txtImprint_Item_No.Text = ""
                End If
            Else
                txtImprint_Item_No.Text = ""
            End If

            iPos = 2
            Call Enable_Imprint_Fields()
            iPos = 3

            If g_oOrdline.ImprintLine Is Nothing Then
                Call ImprintReset(pintRow)
            Else
                Call ImprintFill(g_oOrdline.ImprintLine)
            End If

            iPos = 4

            If Trim(txtImprint_Item_No.Text.ToString).Length > 3 Then
                If txtImprint_Item_No.ToString.Contains("╘═ ") Then
                    If txtImprint_Item_No.Text.ToString.Length > 11 Then
                        cmdUpdateAllItem.Text = "Update All " & Mid(txtImprint_Item_No.Text.ToString, 4, 8)
                    Else
                        cmdUpdateAllItem.Text = "Update All " & Mid(txtImprint_Item_No.Text.ToString, 4, Trim(txtImprint_Item_No.Text.ToString).Length - 3)
                    End If
                Else
                    If txtImprint_Item_No.Text.ToString.Length > 11 Then
                        cmdUpdateAllItem.Text = "Update All " & Mid(txtImprint_Item_No.Text.ToString, 1, 8)
                    Else
                        cmdUpdateAllItem.Text = "Update All " & Mid(txtImprint_Item_No.Text.ToString, 1, Trim(txtImprint_Item_No.Text.ToString).Length - 3)
                    End If
                End If
                'cmdUpdateAllItem.Enabled = Not (Mid(txtImprint_Item_No.Text.ToString, 1, 3) = "╘═ ")
                'If txtImprint_Item_No.Text.ToString.Length > 11 Then
                '    cmdUpdateAllItem.Text = "Update All " & Mid(txtImprint_Item_No.Text.ToString, 1, 8)
                'Else
                '    cmdUpdateAllItem.Text = "Update All " & Mid(txtImprint_Item_No.Text.ToString, 4, Trim(txtImprint_Item_No.Text.ToString).Length - 3)
                'End If

            Else
                cmdUpdateAllItem.Enabled = False
                cmdUpdateAllItem.Text = "Update All xxxx"
            End If

            iPos = 5

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
            txtRepeat_From_ID.Enabled = blnEnabled
            txtComments.Enabled = blnEnabled
            txtSpecial_Comments.Enabled = blnEnabled

            cboImprint_Location.Enabled = blnEnabled
            cboImprint_Color.Enabled = blnEnabled
            cboPackaging.Enabled = blnEnabled
            cboRefill.Enabled = blnEnabled
            cboIndustry.Enabled = blnEnabled

            cmdClear_Imprint.Enabled = blnEnabled
            cmdGetRepeatData.Enabled = blnEnabled
            cmdUpdateAllItem.Enabled = blnEnabled
            cmdUpdateAll.Enabled = blnEnabled
            cmdPaste.Enabled = blnEnabled
            cmdCopyDown.Enabled = blnEnabled

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ImprintFill() ' ByRef pOrder As cOrder)

        Try
            Call ImprintComboFill()

            blnImprintLoading = True

            txtImprint.Text = g_oOrdline.ImprintLine.Imprint
            txtEnd_User.Text = g_oOrdline.ImprintLine.End_User
            txtImprint_Location.Text = g_oOrdline.ImprintLine.Location
            txtNum_Impr_1.Text = IIf(g_oOrdline.ImprintLine.Num_Impr_1 <> 0, g_oOrdline.ImprintLine.Num_Impr_1, String.Empty)
            txtNum_Impr_2.Text = IIf(g_oOrdline.ImprintLine.Num_Impr_2 <> 0, g_oOrdline.ImprintLine.Num_Impr_2, String.Empty)
            txtNum_Impr_3.Text = IIf(g_oOrdline.ImprintLine.Num_Impr_3 <> 0, g_oOrdline.ImprintLine.Num_Impr_3, String.Empty)
            'txtNum_Impr_2.Text = g_oOrdline.ImprintLine.Num_Impr_2
            'txtNum_Impr_3.Text = g_oOrdline.ImprintLine.Num_Impr_3
            txtImprint_Color.Text = g_oOrdline.ImprintLine.Color
            txtPackaging.Text = g_oOrdline.ImprintLine.Packaging
            txtRefill.Text = g_oOrdline.ImprintLine.Refill
            cboIndustry.Text = g_oOrdline.ImprintLine.Industry
            txtComments.Text = g_oOrdline.ImprintLine.Comments
            txtSpecial_Comments.Text = g_oOrdline.ImprintLine.Special_Comments

            cboImprint_Color.Text = ""
            cboImprint_Location.Text = ""
            cboRefill.Text = ""
            cboPackaging.Text = ""
            'cboIndustry.Text = ""

            blnImprintLoading = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ImprintFill(ByRef poImprint As cImprint)

        Try
            g_oOrdline.ImprintLine = poImprint
            Call ImprintFill()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ImprintReset(ByVal pintRow As Integer) ' ByRef pOrder As cOrder)

        Try
            If g_oOrdline.ImprintLine Is Nothing Then
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
                txtRepeat_From_ID.Text = String.Empty
                txtComments.Text = String.Empty
                txtSpecial_Comments.Text = String.Empty
                'txtSpecial_Comments.Text = String.Empty

            Else
                g_oOrdline.ImprintLine.Reset()
                If Not g_oOrdline.ImprintLine.IsEmpty() Then g_oOrdline.ImprintLine.Save()
                Call ImprintFill()
                'With dgvItems.Rows(dgvItems.CurrentRow.Index)
                With dgvItems.Rows(pintRow)
                    .Cells(Columns.Imprint).Value = g_oOrdline.ImprintLine.Imprint
                    .Cells(Columns.End_User).Value = g_oOrdline.ImprintLine.End_User
                    .Cells(Columns.Imprint_Location).Value = g_oOrdline.ImprintLine.Location
                    .Cells(Columns.Imprint_Color).Value = g_oOrdline.ImprintLine.Color
                    .Cells(Columns.Num_Impr_1).Value = IIf(g_oOrdline.ImprintLine.Num_Impr_1 <> 0, g_oOrdline.ImprintLine.Num_Impr_1, DBNull.Value)
                    .Cells(Columns.Num_Impr_2).Value = IIf(g_oOrdline.ImprintLine.Num_Impr_2 <> 0, g_oOrdline.ImprintLine.Num_Impr_2, DBNull.Value)
                    .Cells(Columns.Num_Impr_3).Value = IIf(g_oOrdline.ImprintLine.Num_Impr_3 <> 0, g_oOrdline.ImprintLine.Num_Impr_3, DBNull.Value)
                    '.Cells(Columns.Num_Impr_1).Value = g_oOrdline.ImprintLine.Num_Impr_1
                    '.Cells(Columns.Num_Impr_2).Value = g_oOrdline.ImprintLine.Num_Impr_2
                    '.Cells(Columns.Num_Impr_3).Value = g_oOrdline.ImprintLine.Num_Impr_3
                    .Cells(Columns.Packaging).Value = g_oOrdline.ImprintLine.Packaging
                    .Cells(Columns.Refill).Value = g_oOrdline.ImprintLine.Refill
                    .Cells(Columns.Laser_Setup).Value = g_oOrdline.ImprintLine.Laser_Setup
                    .Cells(Columns.Industry).Value = g_oOrdline.ImprintLine.Industry
                    .Cells(Columns.Comments).Value = g_oOrdline.ImprintLine.Comments
                    .Cells(Columns.Special_Comments).Value = g_oOrdline.ImprintLine.Special_Comments
                    '.Cells(Columns.Printer_Comment).Value = g_oOrdline.ImprintLine.Printer_Comment
                    '.Cells(Columns.Packer_Comment).Value = g_oOrdline.ImprintLine.Packer_Comment
                    '.Cells(Columns.Printer_Instructions).Value = g_oOrdline.ImprintLine.Printer_Instructions
                    '.Cells(Columns.Packer_Instructions).Value = g_oOrdline.ImprintLine.Packer_Instructions
                    .Cells(Columns.Repeat_From_ID).Value = g_oOrdline.ImprintLine.Repeat_From_ID
                End With
            End If

            txtRepeat_Ord_No.Text = String.Empty
            cboImprint_Color.Text = ""
            cboImprint_Location.Text = ""
            cboRefill.Text = ""
            cboPackaging.Text = ""
            cboIndustry.Text = ""

            g_oOrdline.Route = ""
            g_oOrdline.RouteID = 0
            g_oOrdline.ProductProof = 0

            If Not (dgvItems.CurrentRow Is Nothing) Then
                'With dgvItems.Rows(dgvItems.CurrentRow.Index)
                With dgvItems.Rows(pintRow)
                    .Cells(Columns.RouteID).Value = 0
                    .Cells(Columns.Route).Value = ""
                    .Cells(Columns.ProductProof).Value = False
                End With
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ImprintSave()

        Try
            ' We don't save OrdGUID and ItemGUID as they always stay the same here.
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Imprint = txtImprint.Text
            g_oOrdline.ImprintLine.End_User = txtEnd_User.Text
            g_oOrdline.ImprintLine.Location = txtImprint_Location.Text
            If IsNumeric(txtNum_Impr_1.Text) Then
                g_oOrdline.ImprintLine.Num_Impr_1 = CInt(txtNum_Impr_1.Text)
            Else
                g_oOrdline.ImprintLine.Num_Impr_1 = 0
            End If
            If IsNumeric(txtNum_Impr_2.Text) Then
                g_oOrdline.ImprintLine.Num_Impr_2 = CInt(txtNum_Impr_2.Text)
            Else
                g_oOrdline.ImprintLine.Num_Impr_2 = 0
            End If
            If IsNumeric(txtNum_Impr_3.Text) Then
                g_oOrdline.ImprintLine.Num_Impr_3 = CInt(txtNum_Impr_3.Text)
            Else
                g_oOrdline.ImprintLine.Num_Impr_3 = 0
            End If
            g_oOrdline.ImprintLine.Color = txtImprint_Color.Text
            g_oOrdline.ImprintLine.Packaging = txtPackaging.Text
            g_oOrdline.ImprintLine.Refill = txtRefill.Text
            'g_oOrdline.ImprintLine.LaserSetup = txtLaserSetup.text
            g_oOrdline.ImprintLine.Industry = txtIndustry.Text
            g_oOrdline.ImprintLine.Comments = txtComments.Text
            g_oOrdline.ImprintLine.Special_Comments = txtSpecial_Comments.Text
            If IsNumeric(txtRepeat_From_ID.Text) Then
                g_oOrdline.ImprintLine.Repeat_From_ID = CInt(txtRepeat_From_ID.Text)
            Else
                g_oOrdline.ImprintLine.Repeat_From_ID = 0
            End If

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

            g_oOrdline.ImprintLine.AddLocation(cboImprint_Location.SelectedItem.ToString, txtImprint_Location.Text)
            txtImprint_Location.Text = g_oOrdline.ImprintLine.Location
            dgvItems.CurrentRow.Cells(Columns.Imprint_Location).Value = txtImprint_Location.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cboImprintColor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboImprint_Color.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            g_oOrdline.ImprintLine.AddColor(cboImprint_Color.SelectedItem.ToString, txtImprint_Color.Text)
            txtImprint_Color.Text = g_oOrdline.ImprintLine.Color
            dgvItems.CurrentRow.Cells(Columns.Imprint_Color).Value = txtImprint_Color.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cboPackaging_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPackaging.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            g_oOrdline.ImprintLine.AddPackaging(cboPackaging.SelectedItem.ToString, txtPackaging.Text)
            txtPackaging.Text = g_oOrdline.ImprintLine.Packaging
            dgvItems.CurrentRow.Cells(Columns.Packaging).Value = txtPackaging.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cboRefill_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRefill.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub

            g_oOrdline.ImprintLine.AddRefill(cboRefill.SelectedItem.ToString, txtRefill.Text)
            txtRefill.Text = g_oOrdline.ImprintLine.Refill
            dgvItems.CurrentRow.Cells(Columns.Refill).Value = txtRefill.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtImprint_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImprint.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Imprint = txtImprint.Text
            dgvItems.CurrentRow.Cells(Columns.Imprint).Value = txtImprint.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtEnd_User_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEnd_User.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.End_User = txtEnd_User.Text
            dgvItems.CurrentRow.Cells(Columns.End_User).Value = txtEnd_User.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtImprintLocation_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImprint_Location.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Location = txtImprint_Location.Text
            dgvItems.CurrentRow.Cells(Columns.Imprint_Location).Value = txtImprint_Location.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtImprintColor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImprint_Color.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Color = txtImprint_Color.Text
            dgvItems.CurrentRow.Cells(Columns.Imprint_Color).Value = txtImprint_Color.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtNum_Impr_1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNum_Impr_1.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            If IsNumeric(txtNum_Impr_1.Text) Then g_oOrdline.ImprintLine.Num_Impr_1 = txtNum_Impr_1.Text
            dgvItems.CurrentRow.Cells(Columns.Num_Impr_1).Value = IIf(IsNumeric(txtNum_Impr_1.Text), txtNum_Impr_1.Text, DBNull.Value)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtNum_Impr_2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNum_Impr_2.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            If IsNumeric(txtNum_Impr_2.Text) Then g_oOrdline.ImprintLine.Num_Impr_2 = txtNum_Impr_2.Text
            dgvItems.CurrentRow.Cells(Columns.Num_Impr_2).Value = IIf(IsNumeric(txtNum_Impr_2.Text), txtNum_Impr_2.Text, DBNull.Value)
            'g_oOrdline.ImprintLine.Num_Impr_2 = txtNum_Impr_2.Text
            'dgvItems.CurrentRow.Cells(Columns.Num_Impr_2).Value = txtNum_Impr_2.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtNum_Impr_3_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNum_Impr_3.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            If IsNumeric(txtNum_Impr_3.Text) Then g_oOrdline.ImprintLine.Num_Impr_3 = txtNum_Impr_3.Text
            dgvItems.CurrentRow.Cells(Columns.Num_Impr_3).Value = IIf(IsNumeric(txtNum_Impr_3.Text), txtNum_Impr_3.Text, DBNull.Value)

            'g_oOrdline.ImprintLine.Num_Impr_3 = txtNum_Impr_3.Text
            'dgvItems.CurrentRow.Cells(Columns.Num_Impr_3).Value = txtNum_Impr_3.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtPackaging_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPackaging.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Packaging = txtPackaging.Text
            dgvItems.CurrentRow.Cells(Columns.Packaging).Value = txtPackaging.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtRefill_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRefill.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Refill = txtRefill.Text
            dgvItems.CurrentRow.Cells(Columns.Refill).Value = txtRefill.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtIndustry_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndustry.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Industry = txtIndustry.Text
            dgvItems.CurrentRow.Cells(Columns.Industry).Value = txtIndustry.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtComments_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtComments.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Comments = txtComments.Text
            dgvItems.CurrentRow.Cells(Columns.Comments).Value = txtComments.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtSpecialComments_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSpecial_Comments.Leave

        Try
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Special_Comments = txtSpecial_Comments.Text
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

        Try
            Call Reset()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdUpdateAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateAll.Click

        Try
            If dgvItems.Rows.Count < 2 Then Exit Sub

            Dim oResult As MsgBoxResult
            oResult = MsgBox("Are you sure you want to copy imprint to all lines ?", MsgBoxStyle.YesNo)

            If oResult = MsgBoxResult.No Then Exit Sub

            'Dim intRouteID As Integer = dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            'Dim strRoute As String = dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString

            Dim intRouteID As Integer = g_oOrdline.RouteID 'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            Dim strRoute As String = g_oOrdline.Route ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
            Dim intProductProof As Integer = g_oOrdline.ProductProof
            Dim intAutoCompleteReship As Integer = g_oOrdline.AutoCompleteReship

            Dim oUpdateFrom As cImprint
            oUpdateFrom = GetImprintRow()

            For lPos As Integer = 0 To dgvItems.Rows.Count - 1

                If Not (dgvItems.Rows(lPos).Cells(Columns.Item_Guid).Value.ToString = oUpdateFrom.Item_Guid) And dgvItems.Rows(lPos).Cells(Columns.Kit_Item_Guid).Value.ToString <> dgvItems.Rows(lPos).Cells(Columns.Item_Guid).Value.ToString Then

                    If dgvItems.Rows(lPos).Cells(Columns.Item_No).Value.ToString <> "" Then

                        SetImprintRow(lPos, oUpdateFrom, intRouteID, strRoute, intProductProof, intAutoCompleteReship)

                        'trip save validation for imprint data
                        Call updateImprintTripValidation(lPos)

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

            Dim intRouteID As Integer = g_oOrdline.RouteID 'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            Dim strRoute As String = g_oOrdline.Route ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
            Dim intProductproof As Integer = g_oOrdline.ProductProof ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
            Dim intAutoCompleteReship As Integer = g_oOrdline.AutoCompleteReship
            Dim item_cd = g_oOrdline.Item_Cd.Trim()

            Dim oUpdateFrom As cImprint

            oUpdateFrom = GetImprintRow()

            For lPos As Integer = 0 To dgvItems.Rows.Count - 1

                'If Not (dgvItems.Rows(lPos).Cells(Columns.Item_Guid).Value.ToString = oUpdateFrom.Item_Guid) Then

                If dgvItems.Rows(lPos).Cells(Columns.Item_No).Value.ToString <> "" Then

                    If dgvItems.Rows(lPos).Cells(Columns.Item_Cd).Value.ToString.Trim() = item_cd And dgvItems.Rows(lPos).Cells(Columns.Kit_Item_Guid).Value.ToString <> dgvItems.Rows(lPos).Cells(Columns.Item_Guid).Value.ToString Then

                        SetImprintRow(lPos, oUpdateFrom, intRouteID, strRoute, intProductproof, intAutoCompleteReship)

                        'trip save validation for imprint data
                        Call updateImprintTripValidation(lPos)

                    End If

                    'trip promise date thermo check
                    Call updatePromiseTripValidation(lPos)

                End If

                'End If

            Next lPos

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdCopyDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyDown.Click

        Try
            ' Exit if only 1 row
            If dgvItems.Rows.Count < 2 Then Exit Sub

            ' Exit if on the last row (cannot copy down...)
            If dgvItems.CurrentRow.Index + 1 = dgvItems.Rows.Count Then Exit Sub

            ' Exit if on the next to last row and last row item_no is empty (cannot copy down...)
            If dgvItems.CurrentRow.Index + 2 = dgvItems.Rows.Count And dgvItems.Rows(dgvItems.CurrentRow.Index + 1).Cells(Columns.Item_No).Value.ToString = "" Then Exit Sub

            intCopyRouteID = g_oOrdline.RouteID 'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            strCopyRoute = g_oOrdline.Route ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString
            intCopyProductProof = g_oOrdline.ProductProof
            intCopyAutoCompleteReship = g_oOrdline.AutoCompleteReship

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

        '    Dim intRouteID As Integer = g_oOrdline.RouteID 'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
        '    Dim strRoute As String = g_oOrdline.Route ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString

        '    Dim oUpdateFrom As cImprint

        '    oUpdateFrom = GetImprintRow()

        '    SetImprintRow(dgvItems.CurrentRow.Index + 1, oUpdateFrom, intRouteID, strRoute)

        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

    Private Sub cmdPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPaste.Click

        Try
            If oCopyFrom Is Nothing Then Exit Sub

            '' Exit if on the next to last row and last row item_no is empty (cannot copy down...)
            If dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value.ToString = "" Then Exit Sub

            'Dim intRouteID As Integer = g_oOrdline.RouteID 'dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.RouteID).Value.ToString
            'Dim strRoute As String = g_oOrdline.Route ' dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Route).Value.ToString

            'Dim oUpdateFrom As cImprint

            'oUpdateFrom = GetImprintRow()

            SetImprintRow(dgvItems.CurrentRow.Index, oCopyFrom, intCopyRouteID, strCopyRoute, intCopyProductProof, intCopyAutoCompleteReship)

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

            GetImprintRow.Ord_Guid = m_oOrder.Ordhead.Ord_GUID.ToString
            GetImprintRow.Item_Guid = dgvItems.CurrentRow.Cells(Columns.Item_Guid).Value.ToString
            GetImprintRow.Item_No = dgvItems.CurrentRow.Cells(Columns.Item_No).Value.ToString
            GetImprintRow.Imprint = dgvItems.CurrentRow.Cells(Columns.Imprint).Value.ToString
            GetImprintRow.End_User = dgvItems.CurrentRow.Cells(Columns.End_User).Value.ToString
            GetImprintRow.Location = dgvItems.CurrentRow.Cells(Columns.Imprint_Location).Value.ToString
            GetImprintRow.Color = dgvItems.CurrentRow.Cells(Columns.Imprint_Color).Value.ToString
            If IsNumeric(dgvItems.CurrentRow.Cells(Columns.Num_Impr_1).Value.ToString) Then
                GetImprintRow.Num_Impr_1 = dgvItems.CurrentRow.Cells(Columns.Num_Impr_1).Value
            Else
                GetImprintRow.Num_Impr_1 = 0
            End If
            If IsNumeric(dgvItems.CurrentRow.Cells(Columns.Num_Impr_2).Value.ToString) Then
                GetImprintRow.Num_Impr_2 = dgvItems.CurrentRow.Cells(Columns.Num_Impr_2).Value
            Else
                GetImprintRow.Num_Impr_2 = 0
            End If
            If IsNumeric(dgvItems.CurrentRow.Cells(Columns.Num_Impr_3).Value.ToString) Then
                GetImprintRow.Num_Impr_3 = dgvItems.CurrentRow.Cells(Columns.Num_Impr_3).Value
            Else
                GetImprintRow.Num_Impr_3 = 0
            End If
            GetImprintRow.Num = GetImprintRow.Num_Impr_1
            GetImprintRow.Packaging = dgvItems.CurrentRow.Cells(Columns.Packaging).Value.ToString
            GetImprintRow.Refill = dgvItems.CurrentRow.Cells(Columns.Refill).Value.ToString
            GetImprintRow.Laser_Setup = dgvItems.CurrentRow.Cells(Columns.Laser_Setup).Value.ToString
            GetImprintRow.Industry = dgvItems.CurrentRow.Cells(Columns.Industry).Value.ToString
            GetImprintRow.Comments = dgvItems.CurrentRow.Cells(Columns.Comments).Value.ToString
            GetImprintRow.Special_Comments = dgvItems.CurrentRow.Cells(Columns.Special_Comments).Value.ToString
            GetImprintRow.Printer_Comment = dgvItems.CurrentRow.Cells(Columns.Printer_Comment).Value.ToString
            GetImprintRow.Packer_Comment = dgvItems.CurrentRow.Cells(Columns.Packer_Comment).Value.ToString
            GetImprintRow.Printer_Instructions = dgvItems.CurrentRow.Cells(Columns.Printer_Instructions).Value.ToString
            GetImprintRow.Packer_Instructions = dgvItems.CurrentRow.Cells(Columns.Packer_Instructions).Value.ToString
            If IsNumeric(dgvItems.CurrentRow.Cells(Columns.Repeat_From_ID).Value.ToString) Then
                GetImprintRow.Repeat_From_ID = dgvItems.CurrentRow.Cells(Columns.Repeat_From_ID).Value
            Else
                GetImprintRow.Repeat_From_ID = 0
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    'Private Sub SetImprintRow(ByVal pintRow As Integer, ByRef poSetFrom As cImprint)

    '    Try
    '        With dgvItems.Rows(pintRow)

    '            If m_oOrder.OrdLines.Contains(.Cells(Columns.Item_Guid).Value.ToString) Then
    '                Dim oOrdline As cOrdLine
    '                oOrdline = m_oOrder.OrdLines(.Cells(Columns.Item_Guid).Value.ToString)
    '                If oOrdline.ImprintLine Is Nothing Then oOrdline.ImprintLine = New cImprint(m_oOrder.Ordhead.Ord_GUID, .Cells(Columns.Item_Guid).Value.ToString)

    '                oOrdline.ImprintLine.Imprint = poSetFrom.Imprint
    '                oOrdline.ImprintLine.Location = poSetFrom.Location
    '                oOrdline.ImprintLine.Color = poSetFrom.Color
    '                oOrdline.ImprintLine.Num = poSetFrom.Num
    '                oOrdline.ImprintLine.Packaging = poSetFrom.Packaging
    '                oOrdline.ImprintLine.Refill = poSetFrom.Refill
    '                oOrdline.ImprintLine.Industry = poSetFrom.Industry
    '                oOrdline.ImprintLine.Comments = poSetFrom.Comments
    '                oOrdline.ImprintLine.Special_Comments = poSetFrom.Special_Comments

    '            End If

    '            .Cells(Columns.Imprint).Value = poSetFrom.Imprint
    '            .Cells(Columns.Imprint_Location).Value = poSetFrom.Location
    '            .Cells(Columns.Imprint_Color).Value = poSetFrom.Color
    '            .Cells(Columns.Num_Impr_1).Value = poSetFrom.Num_Impr_1
    '            .Cells(Columns.Num_Impr_2).Value = poSetFrom.Num_Impr_2
    '            .Cells(Columns.Num_Impr_3).Value = poSetFrom.Num_Impr_3
    '            .Cells(Columns.Packaging).Value = poSetFrom.Packaging
    '            .Cells(Columns.Refill).Value = poSetFrom.Refill
    '            .Cells(Columns.Industry).Value = poSetFrom.Industry
    '            .Cells(Columns.Comments).Value = poSetFrom.Comments
    '            .Cells(Columns.Special_Comments).Value = poSetFrom.Special_Comments

    '        End With

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Private Sub SetImprintRow(ByVal pintRow As Integer, ByRef poSetFrom As cImprint, ByVal pintRouteID As Integer, ByVal pstrRoute As String, ByVal pintProductProof As Integer, ByVal pintAutoCompleteReship As Integer)

        Try
            With dgvItems.Rows(pintRow)

                If m_oOrder.OrdLines.Contains(.Cells(Columns.Item_Guid).Value.ToString) Then
                    Dim oOrdline As cOrdLine
                    oOrdline = m_oOrder.OrdLines(.Cells(Columns.Item_Guid).Value.ToString)
                    If oOrdline.ImprintLine Is Nothing Then oOrdline.ImprintLine = New cImprint(m_oOrder.Ordhead.Ord_GUID, .Cells(Columns.Item_Guid).Value.ToString)

                    oOrdline.ImprintLine.Imprint = poSetFrom.Imprint
                    oOrdline.ImprintLine.End_User = poSetFrom.End_User
                    oOrdline.ImprintLine.Location = poSetFrom.Location
                    oOrdline.ImprintLine.Color = poSetFrom.Color
                    oOrdline.ImprintLine.Num_Impr_1 = poSetFrom.Num_Impr_1
                    oOrdline.ImprintLine.Num_Impr_2 = poSetFrom.Num_Impr_2
                    oOrdline.ImprintLine.Num_Impr_3 = poSetFrom.Num_Impr_3
                    oOrdline.ImprintLine.Num = poSetFrom.Num_Impr_1
                    oOrdline.ImprintLine.Packaging = poSetFrom.Packaging
                    oOrdline.ImprintLine.Refill = poSetFrom.Refill
                    oOrdline.ImprintLine.Laser_Setup = poSetFrom.Laser_Setup
                    oOrdline.ImprintLine.Industry = poSetFrom.Industry
                    oOrdline.ImprintLine.Comments = poSetFrom.Comments
                    oOrdline.ImprintLine.Special_Comments = poSetFrom.Special_Comments
                    oOrdline.ImprintLine.Printer_Comment = poSetFrom.Printer_Comment
                    oOrdline.ImprintLine.Packer_Comment = poSetFrom.Packer_Comment
                    oOrdline.ImprintLine.Printer_Instructions = poSetFrom.Printer_Instructions
                    oOrdline.ImprintLine.Packer_Instructions = poSetFrom.Packer_Instructions
                    oOrdline.ImprintLine.Repeat_From_ID = poSetFrom.Repeat_From_ID

                    oOrdline.RouteID = pintRouteID
                    oOrdline.Route = pstrRoute
                    oOrdline.ProductProof = pintProductProof
                    oOrdline.AutoCompleteReship = pintAutoCompleteReship

                    oOrdline.Save()

                End If

                ' This will set the special instructions for each item. Must be put here because
                ' each item no of the copy is different, can't take from the source.
                g_oOrdline.Set_Spec_Instructions(.Cells(Columns.Item_No).Value, poSetFrom)

                .Cells(Columns.Imprint).Value = poSetFrom.Imprint
                .Cells(Columns.End_User).Value = poSetFrom.End_User
                .Cells(Columns.Imprint_Location).Value = poSetFrom.Location
                .Cells(Columns.Imprint_Color).Value = poSetFrom.Color
                .Cells(Columns.Num_Impr_1).Value = IIf(poSetFrom.Num_Impr_1 = 0, DBNull.Value, poSetFrom.Num_Impr_1)
                .Cells(Columns.Num_Impr_2).Value = IIf(poSetFrom.Num_Impr_2 = 0, DBNull.Value, poSetFrom.Num_Impr_2)
                .Cells(Columns.Num_Impr_3).Value = IIf(poSetFrom.Num_Impr_3 = 0, DBNull.Value, poSetFrom.Num_Impr_3)
                '.Cells(Columns.Num_Impr_2).Value = poSetFrom.Num_Impr_2
                '.Cells(Columns.Num_Impr_3).Value = poSetFrom.Num_Impr_3
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

                .Cells(Columns.RouteID).Value = pintRouteID
                .Cells(Columns.Route).Value = pstrRoute
                .Cells(Columns.ProductProof).Value = (pintProductProof = 1)
                .Cells(Columns.AutoCompleteReship).Value = (pintAutoCompleteReship = 1)

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Sub SetRouteCombo(ByVal cboColumn As DataGridViewComboBoxColumn, ByVal pstrType As String)

    '    Try
    '        With cboColumn
    '            '.DataSource = RetrieveAlternativeTitles()

    '            '.ValueMember = "RouteId" ' ColumnName.TitleOfCourtesy.ToString()
    '            .ValueMember = "Route" ' ColumnName.TitleOfCourtesy.ToString()
    '            .DisplayMember = "Route"

    '            .DataSource = m_oOrder.ChangeCboRouteList(pstrType)

    '        End With

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub SetRouteCombo(ByVal cboColumn As DataGridViewComboBoxColumn, ByVal pstrType As String, ByVal pstrValue As String)

    '    Try
    '        With cboColumn
    '            '.DataSource = RetrieveAlternativeTitles()

    '            '.ValueMember = "RouteId" ' ColumnName.TitleOfCourtesy.ToString()
    '            .ValueMember = "Route" ' ColumnName.TitleOfCourtesy.ToString()
    '            .DisplayMember = "Route"

    '            .DataSource = m_oOrder.ChangeCboRouteList(pstrType, pstrValue)

    '        End With

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Private Sub SetRouteCombo(ByVal cboCell As DataGridViewComboBoxCell, ByVal pstrType As String, ByVal pstrValue As String)

        Try
            With cboCell
                '.DataSource = RetrieveAlternativeTitles()

                '.ValueMember = "RouteId" ' ColumnName.TitleOfCourtesy.ToString()
                .ValueMember = "Route" ' ColumnName.TitleOfCourtesy.ToString()
                .DisplayMember = "Route"

                .DataSource = m_oOrder.ChangeCboRouteList(pstrType, pstrValue)

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


    Private Sub PerformGridSearchClick()

        Try

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

                Case Columns.Item_No
                    GetSearchElementByColumnIndex = "IM-ITEM"

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
            If dgvItems.CurrentCell.Value.ToString = "" Then Exit Sub

            If GetNoteElementByColumnIndex(dgvItems.CurrentCell.ColumnIndex).ToString = "" Then Exit Sub

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

    Private Sub PerformETAClick()

        Try
            Dim oFormItemETA As New frmItemETA(dgvItems.CurrentRow.Cells(Columns.Item_No).Value)
            oFormItemETA.ShowDialog()

            oFormItemETA.Dispose()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub PerformPriceBreaksClick()

        Try
            Dim oFormItemPriceBreaks As New frmItemPriceBreaks(m_oOrder.Ordhead.Curr_Cd, m_oOrder.Ordhead.Cus_No, dgvItems.CurrentRow.Cells(Columns.Item_No).Value, "", m_oOrder.Ordhead.Ord_Dt.ToString)
            oFormItemPriceBreaks.ShowDialog()

            oFormItemPriceBreaks.Dispose()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub PerformItemInfoClick()

        Try
            Dim oFormItemLocQty As New frmItemLocQty(dgvItems.CurrentRow.Cells(Columns.Item_No).Value, dgvItems.CurrentRow.Cells(Columns.Location).Value)
            oFormItemLocQty.ShowDialog()

            oFormItemLocQty.Dispose()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub PerformItemLocQtyClick()

        Try
            Dim oFormItemInfo As New frmItemInfo(dgvItems.CurrentRow.Cells(Columns.Item_No).Value)
            oFormItemInfo.ShowDialog()

            oFormItemInfo.Dispose()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Property TabWindow() As TabControl
        Get
            TabWindow = sstOrder
        End Get
        Set(ByVal value As TabControl)
            ' Dont want to set for now...
            Exit Property
            sstOrder = value
        End Set
    End Property

    Private Sub sstOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles sstOrder.KeyDown

        'Dim tcControl As TabControl = sender
        'Dim iNothing As Integer = 99
        'Dim iSelectedTab As Integer = 99

        'Try
        '    If e.Control Then
        '        Select Case e.KeyCode
        '            Case Keys.D1
        '                iSelectedTab = frmOrder.Tabs.Header
        '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Header)
        '            Case Keys.D2
        '                iSelectedTab = frmOrder.Tabs.CustomerContact
        '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CustomerContact)
        '            Case Keys.D3
        '                iSelectedTab = frmOrder.Tabs.Lines
        '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Lines)
        '            Case Keys.D4
        '                iSelectedTab = frmOrder.Tabs.Documents
        '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Documents)
        '            Case Keys.D5
        '                iSelectedTab = frmOrder.Tabs.HeaderAll
        '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.HeaderAll)
        '            Case Keys.D6
        '                iSelectedTab = frmOrder.Tabs.Taxes
        '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Taxes)
        '            Case Keys.D7
        '                iSelectedTab = frmOrder.Tabs.Salesperson
        '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
        '            Case Keys.D8
        '                iSelectedTab = frmOrder.Tabs.CreditInfo
        '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
        '            Case Keys.D9
        '                iSelectedTab = frmOrder.Tabs.Extra
        '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
        '        End Select

        '        If iSelectedTab <> iNothing Then

        '            If m_oOrder Is Nothing Then
        '                m_oOrder = New cOrder()
        '                UcOrder.Fill() ' m_oOrder)
        '            End If

        '            If iSelectedTab > Tabs.Header And (Trim(m_oOrder.Ordhead.Cus_No) = "" Or m_oOrder.Ordhead.Cus_No Is Nothing) Then
        '                sstOrder.SelectTab(Tabs.Header)
        '                Exit Sub
        '            End If

        '            If iSelectedTab > Tabs.CustomerContact And (m_oOrder.Ordhead.ContactView = 0) Then
        '                sstOrder.SelectTab(Tabs.CustomerContact)
        '                m_oOrder.Ordhead.ContactView = 1
        '                Exit Sub
        '            End If

        '            Select Case iSelectedTab

        '                Case Tabs.Header
        '                    UcOrder.Fill()

        '                Case Tabs.CustomerContact
        '                    m_oOrder.Ordhead.ContactView = 1
        '                    UcContacts1.Fill()

        '                Case Tabs.Lines
        '                    If m_oOrder.OrdLines.Count = 0 Then
        '                        Call ProductLineRestart(False)
        '                    Else
        '                        If m_oOrder.OrdLines.Count <> 0 And dgvItems.Rows.Count <= 1 Then
        '                            Call LoadPreOrder()
        '                        Else
        '                            Call Fill()
        '                        End If

        '                    End If
        '                    UcOrderTotal1.From_Curr_Cd_Desc = m_oOrder.Ordhead.Curr_Cd
        '                    dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Item_No)
        '                    dgvItems.Focus()

        '                Case Tabs.Documents
        '                    UcDocument1.Fill()

        '                Case Tabs.HeaderAll
        '                    UcHeader1.Fill()

        '                Case Tabs.Taxes
        '                    UcTaxes1.Fill()

        '                Case Tabs.Salesperson
        '                    UcSalesperson1.Fill()

        '                Case Tabs.CreditInfo
        '                    UcCreditInfo1.Fill()

        '                Case Tabs.Extra
        '                    UcExtra1.Fill()

        '            End Select

        '        End If

        '    End If

        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

        Try
            ' If Ctl-Number, sets the new tab.
            If e.Control Then

                If m_oOrder Is Nothing Then
                    m_oOrder = New cOrder()
                    UcOrder.Fill() ' m_oOrder)
                End If

                Select Case e.KeyCode
                    Case Keys.D1
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Header)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                    Case Keys.D2
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CustomerContact)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                    Case Keys.D3
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Lines)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                    Case Keys.D4
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Documents)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                    Case Keys.D5
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.HeaderAll)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                    Case Keys.D6
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Taxes)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                    Case Keys.D7
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                    Case Keys.D8
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                    Case Keys.D9
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                    Case Keys.D0
                        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.History)
                        Call sstOrder_Click(sstOrder, New System.EventArgs)
                End Select

            End If

            If e.Alt Then
                Select Case e.KeyCode

                    ' EXPORT TO EXCEL
                    Case Keys.X
                        tsbExportToMacola.PerformClick()

                        ' NEW ORDER
                    Case Keys.N
                        tsbNew.PerformClick()

                End Select

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Dim ofrm As New cComments(m_oOrder.Ordhead.Ord_GUID)

    'End Sub

    Private Sub cmdGetRepeatData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetRepeatData.Click

        Try
            If Trim(txtRepeat_Ord_No.Text) <> "" Then

                Dim oFrmRepeatImprint As New frmRepeatImprint(txtRepeat_Ord_No.Text)
                oFrmRepeatImprint.ShowDialog()
                g_oOrdline.Set_Spec_Instructions(txtImprint_Item_No.Text, oFrmRepeatImprint.Imprint)

                txtImprint.Text = oFrmRepeatImprint.Imprint.Imprint
                txtNum_Impr_1.Text = IIf(oFrmRepeatImprint.Imprint.Num_Impr_1 <> 0, oFrmRepeatImprint.Imprint.Num_Impr_1, "")
                txtNum_Impr_2.Text = IIf(oFrmRepeatImprint.Imprint.Num_Impr_2 <> 0, oFrmRepeatImprint.Imprint.Num_Impr_2, "")
                txtNum_Impr_3.Text = IIf(oFrmRepeatImprint.Imprint.Num_Impr_3 <> 0, oFrmRepeatImprint.Imprint.Num_Impr_3, "")
                'txtNum_Impr_2.Text = oFrmRepeatImprint.Imprint.Num_Impr_2
                'txtNum_Impr_3.Text = oFrmRepeatImprint.Imprint.Num_Impr_3
                txtComments.Text = oFrmRepeatImprint.Imprint.Comments
                txtSpecial_Comments.Text = oFrmRepeatImprint.Imprint.Special_Comments
                txtEnd_User.Text = oFrmRepeatImprint.Imprint.End_User
                txtImprint_Location.Text = oFrmRepeatImprint.Imprint.Location
                txtImprint_Color.Text = oFrmRepeatImprint.Imprint.Color
                txtPackaging.Text = oFrmRepeatImprint.Imprint.Packaging
                txtRefill.Text = oFrmRepeatImprint.Imprint.Refill
                txtIndustry.Text = oFrmRepeatImprint.Imprint.Industry
                txtRepeat_From_ID.Text = oFrmRepeatImprint.Imprint.Repeat_From_ID

                With dgvItems.CurrentRow
                    .Cells(Columns.Imprint).Value = oFrmRepeatImprint.Imprint.Imprint
                    .Cells(Columns.Num_Impr_1).Value = IIf(oFrmRepeatImprint.Imprint.Num_Impr_1 <> 0, oFrmRepeatImprint.Imprint.Num_Impr_1, DBNull.Value)
                    .Cells(Columns.Num_Impr_2).Value = IIf(oFrmRepeatImprint.Imprint.Num_Impr_2 <> 0, oFrmRepeatImprint.Imprint.Num_Impr_2, DBNull.Value)
                    .Cells(Columns.Num_Impr_3).Value = IIf(oFrmRepeatImprint.Imprint.Num_Impr_3 <> 0, oFrmRepeatImprint.Imprint.Num_Impr_3, DBNull.Value)
                    '.Cells(Columns.Num_Impr_1).Value = oFrmRepeatImprint.Imprint.Num_Impr_1
                    '.Cells(Columns.Num_Impr_2).Value = oFrmRepeatImprint.Imprint.Num_Impr_2
                    '.Cells(Columns.Num_Impr_3).Value = oFrmRepeatImprint.Imprint.Num_Impr_3
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
                End With

                Call ImprintSave()

                oFrmRepeatImprint.Dispose()

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub SaveUncommited()

        Try
            If dgvItems.Rows.Count = 0 Then Exit Sub
            If dgvItems.CurrentRow.Cells(Columns.Item_No).Value.Equals(DBNull.Value) Then Exit Sub
            If dgvItems.CurrentRow.Cells(Columns.Item_No).Value = "" Then Exit Sub

            dgvItems.CurrentCell = dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No)
            Call SaveOrderLine(g_oOrdline)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Sub Validate_Cus_No_Exist(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UcTaxes1.VisibleChanged, UcSalesperson1.VisibleChanged, UcOrderTotal1.VisibleChanged, UcImprint1.VisibleChanged, UcHeader1.VisibleChanged, UcExtra1.VisibleChanged, UcCreditInfo1.VisibleChanged, UcContacts1.VisibleChanged

    'End Sub


    Private Sub cboIndustry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboIndustry.SelectedIndexChanged

        Try
            If blnImprintLoading Then Exit Sub
            If g_oOrdline.ImprintLine Is Nothing Then Exit Sub

            g_oOrdline.ImprintLine.Industry = cboIndustry.SelectedItem.ToString ' txtIndustry.Text)
            txtIndustry.Text = g_oOrdline.ImprintLine.Industry
            dgvItems.CurrentRow.Cells(Columns.Industry).Value = txtIndustry.Text

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdWebpage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Item_No)

            wbItem.Url = New Uri("http://www.spectorandco.ca/en/products/products.php?item_num=ST5030")
            wbItem.Visible = True

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub cmdRepeatSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRepeatSearch.Click

        Try
            Dim oSearch As New cSearch()
            oSearch.SearchElement = "OEI-REPEAT"
            oSearch.SetDefaultSearchElement("RTRIM(OEHDRHST_SQL.cus_no)", RTrim(m_oOrder.Ordhead.Cus_No), "200")
            oSearch.Form.ShowDialog()
            If oSearch.Form.FoundElementValue <> String.Empty Then
                'Dim oSearch As New cSearch("AR-ALT-ADDR-COD")
                txtRepeat_Ord_No.Text = oSearch.Form.FoundElementValue
                'cboShipTo.Text = txtCus_Alt_Adr_Cd.Text
            End If
            oSearch.Dispose()
            oSearch = Nothing
            txtRepeat_Ord_No.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEmail.Click

        'Dim msOutlook As New Microsoft.Office.Interop.Outlook.Application
        'Dim msEmail As Microsoft.Office.Interop.Outlook.MailItem

        'msEmail = msOutlook.CreateItem(Outlook.OlItemType.olMailItem)
        'msEmail.Display()

        Try
            Dim oFrmEmail As New frmEmail
            oFrmEmail.ShowDialog()

            oFrmEmail.Dispose()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub btnODBCConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnODBCConnect.Click

        Try
            Dim dt As DataTable
            Dim db As New cODBC

            dt = db.DataTable("SELECT * FROM OEI_ORDHDR WITH (Nolock)")

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub chkNoPictures_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLowRow.CheckedChanged

        Try
            For Each row As DataGridViewRow In dgvItems.Rows
                If chkLowRow.Checked Then
                    row.Height = 22
                Else
                    If row.Cells(Columns.PictureHeight).Value.ToString = "" Then
                        row.Height = 22
                    Else
                        row.Height = row.Cells(Columns.PictureHeight).Value
                    End If
                End If
            Next

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdImprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImprint.Click

        Try
            If dgvItems.Rows.Count > 0 Then
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Packer_Instructions)
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Imprint)
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdTraveler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTraveler.Click

        Try
            If dgvItems.Rows.Count > 0 Then

                mblnNoMsgBox = True
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Recalc_SW)
                'dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.ProductProof)
                'dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Packer_Instructions)
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Route)
                mblnNoMsgBox = False

                dgvItems.Focus()

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdQty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQty.Click

        Try
            If dgvItems.Rows.Count > 0 Then
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Packer_Instructions)
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Qty_Ordered)
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdRequest_Dt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRequest_Dt.Click

        Try
            If dgvItems.Rows.Count > 0 Then
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Packer_Instructions)
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Request_Dt)
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdDirServices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDirServices.Click

        Try
            Dim a1
            a1 = returngal()
            'Dim oDirectory As System.DirectoryServices.ActiveDirectory.
            'oDirectory()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function returngal() As ArrayList

        Dim arrGal As New ArrayList()
        Try
            Dim objsearch As New DirectorySearcher()
            Dim strrootdse As String = objsearch.SearchRoot.Path
            Dim objdirentry As New DirectoryEntry(strrootdse)
            objsearch.Filter = "(& (mailnickname=*)(objectClass=user))"
            objsearch.SearchScope = System.DirectoryServices.SearchScope.Subtree
            objsearch.PropertiesToLoad.Add("cn")
            objsearch.PropertyNamesOnly = True
            objsearch.Sort.Direction = System.DirectoryServices.SortDirection.Ascending
            objsearch.Sort.PropertyName = "cn"
            Dim colresults As SearchResultCollection = objsearch.FindAll()
            For Each objresult As SearchResult In colresults
                Dim de As DirectoryEntry = objresult.GetDirectoryEntry()
                'DisplayPropertyNames(de)
                'Debug.Print(objresult.GetDirectoryEntry().Properties("displayName").Value & " " & objresult.GetDirectoryEntry().Properties("mail").Value & " " & objresult.GetDirectoryEntry().Properties("cn").Value)
                arrGal.Add(objresult.GetDirectoryEntry().Properties("cn").Value)
                Debug.Print(objresult.GetDirectoryEntry().Properties("cn").Value & " -- " & objresult.GetDirectoryEntry().Properties("displayName").Value)
                'arrGal.Add(objresult.GetDirectoryEntry().Properties("mail").Value)
                'Debug.Print(objresult.GetDirectoryEntry.Properties.PropertyNames.GetEnumerator.Current.ToString)
                'Debug.Print(a1)
            Next
            objsearch.Dispose()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
        Return arrGal

    End Function

    Private Sub DisplayPropertyNames(ByVal de As DirectoryEntry)

        Try
            Dim output As String = ""
            Dim IColl As System.Collections.ICollection = de.Properties.PropertyNames
            Dim IEnum As IEnumerator = IColl.GetEnumerator
            While IEnum.MoveNext = True
                Debug.Print("PROPERTY: " & IEnum.Current.ToString & ".....  " & de.Properties(IEnum.Current.ToString).Value.ToString)
                'output &= IEnum.Current.ToString & ControlChars.NewLine
            End While
            'MsgBox(output)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub chkAutoCompleteReship_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAutoCompleteReship.CheckedChanged

        Try
            If chkAutoCompleteReship.Checked Then
                Dim oResultat As New MsgBoxResult
                oResultat = MsgBox("Are you sure you want to autocomplete for reship this order ?", MsgBoxStyle.YesNo)
                If oResultat = MsgBoxResult.Yes Then
                    m_oOrder.Ordhead.AutoCompleteReship = IIf(chkAutoCompleteReship.Checked, 1, 0)
                Else
                    chkAutoCompleteReship.Checked = False
                    m_oOrder.Ordhead.AutoCompleteReship = IIf(chkAutoCompleteReship.Checked, 1, 0)
                End If
            Else
                m_oOrder.Ordhead.AutoCompleteReship = IIf(chkAutoCompleteReship.Checked, 1, 0)
            End If

            'Debug.Print(m_oOrder.Ordhead.AutoCompleteReship.ToString)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdPrice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrice.Click

        Try
            If dgvItems.Rows.Count > 0 Then
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Packer_Instructions)
                dgvItems.CurrentCell = dgvItems.CurrentRow.Cells(Columns.Unit_Price)
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub



    Private Sub cmdPrevCharges_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrevCharges.Click
        Call add44Charges()

        'Try

        '    If m_oOrder.OrdLines.Count = 0 Then Exit Sub

        '    Dim oOrderSelect As New frmProductLineEntry() ' m_oOrder)
        '    oOrderSelect.Customer_No = m_oOrder.Ordhead.Cus_No
        '    oOrderSelect.Currency = m_oOrder.Ordhead.Curr_Cd ' "USD"
        '    oOrderSelect.Item_No = "CHARGESCALCULATION"
        '    oOrderSelect.Loc = m_oOrder.Ordhead.Mfg_Loc

        '    oOrderSelect.ShowDialog()

        '    Call LoadPreOrder()

        '    'Select Case cboMoreOptions.Text
        '    '    Case "Import items from order"

        '    '    Case "Search item"

        '    'End Select

        'Catch oe_er As OEException
        '    MsgBox(oe_er.Message)
        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub



    Private Sub cmdOEI_Config_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOEI_Config.Click

        Try
            Dim m_Path As String = ""
            Dim dt As New DataTable
            Dim db As New cDBA
            Dim strSql As String = _
            "SELECT * " & _
            "FROM   OEI_Config C WITH (Nolock) "

            dt = db.DataTable(strSql)
            'If dt.Rows.Count <> 0 Then Call RestartProgram()
            If dt.Rows.Count <> 0 Then

                m_Path = dt.Rows(0).Item("HV_Import_Files_Folder")
                If Mid(Trim(m_Path), Trim(m_Path).Length, 1) <> "\" Then m_Path = Trim(m_Path) & "\"

            End If

            dt.Dispose()

        Catch er As Exception
            ' Do nothing
        End Try

    End Sub

    'Private Sub Validate_Cus_No_Exist(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UcTaxes1.VisibleChanged, UcSalesperson1.VisibleChanged, UcOrderTotal1.VisibleChanged, UcImprint1.VisibleChanged, UcHeader1.VisibleChanged, UcExtra1.VisibleChanged, UcDocument1.VisibleChanged, UcCreditInfo1.VisibleChanged, UcContacts1.VisibleChanged

    'End Sub

    Private Sub btnEDI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEDI.Click

        Dim oEDI As cOrderEDI
        oEDI = New cOrderEDI("850_STAPLES_001.xml")
        oEDI = New cOrderEDI("850_STAPLES_002.xml")
        oEDI = New cOrderEDI("850_STAPLES_003.xml")


    End Sub


    Private Sub btnFileList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFileList.Click

        Dim strException As String = ""

        Try

            Dim diXmlFolder As New IO.DirectoryInfo("c:\")
            Dim diXmlFolderFiles As IO.FileInfo() = diXmlFolder.GetFiles("*.xml")
            Dim fiXMLFile As IO.FileInfo

            'list the names of all files in the specified directory
            For Each fiXMLFile In diXmlFolderFiles

                Try

                    Dim oEDI As cOrderEDI
                    oEDI = New cOrderEDI(fiXMLFile.Name)

                Catch subEr As Exception
                    strException &= "Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & subEr.Message & vbCrLf
                End Try

            Next

            If strException <> "" Then Throw New Exception(strException)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvItems_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.CellDoubleClick

        Debug.Print("dgvItems_CellDoubleClick")

        Try
            If e.ColumnIndex = Columns.Item_No Then
                If dgvItems.CurrentRow.Cells(e.ColumnIndex).Style.BackColor = Color.Red Then
                    Dim oThermo As New cThermometer(cThermometer.StyleEnum.Route, CInt(dgvItems.CurrentRow.Cells(Columns.RouteID).Value))
                    oThermo.Show()
                End If
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub btnThermo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnThermo.Click

        Dim oThermo As New cThermometer(cThermometer.StyleEnum.Full)
        oThermo.Show()

    End Sub


    Private Sub cmdMySqlTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMySqlTest.Click

        ' This proc allows to create a call to mysql and read table products_data_ca

        'Try

        '    ' Dim db As New cMySqlDBA
        '    ' Dim dt As DataTable
        '    Dim strSql As String

        '    strSql = "select table_name from information_schema.tables where table_schema='products2'"
        '    strSql = "SELECT * FROM products_data_ca "
        '    '  dt = db.DataTable(strSql)

        '    If dt.Rows.Count <> 0 Then
        '        For Each dtRow As DataRow In dt.Rows
        '            For Each oColumn As DataColumn In dtRow.Table.Columns
        '                Debug.Print(oColumn.ColumnName.ToString & "   " & oColumn.DataType.ToString & "   " & oColumn.MaxLength.ToString & "   " & dtRow.Item(oColumn.ColumnName))
        '            Next
        '        Next
        '    End If
        '    dt.Dispose()

        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim strMessage As String = ""
        Dim db As New cDBA()
        Dim dt As DataTable
        Dim dtOrderCheck As DataTable

        Try


            Dim strSql As String = _
            "SELECT O.Ord_Guid, O.OEI_Ord_No, O.ExportTS, C.HV_ControlTS, ISNULL(O.Trigger_Message, 0) AS Trigger_Message, " & _
            "       DateDiff(Second, O.ExportTS, GetDate()) AS TimeElapsed " & _
            "FROM   OEI_Config C WITH (Nolock), OEI_OrdHdr O WITH (Nolock) " & _
            "WHERE 	((O.ExportTS IS NOT NULL AND O.MacolaTS IS NULL AND (ABS(DateDiff(Second, O.ExportTS, GetDate()))	> (2 * 60))) OR " & _
            "		(O.MacolaTS IS NOT NULL AND O.TriggerTS IS NULL AND (ABS(DateDiff(Second, O.MacolaTS, GetDate()))	> (2 * 60)))) AND " & _
            "		(DateDiff(Second, C.HV_ControlTS, GetDate()) > 120) "

            '"WHERE  O.ExportTS IS NOT NULL AND O.MacolaTS IS NULL AND " & _
            '"       (ABS(DateDiff(Second, O.ExportTS, GetDate()))	> (3 * 60))"

            dt = db.DataTable(strSql)
            'If dt.Rows.Count <> 0 Then Call RestartProgram()
            If dt.Rows.Count <> 0 Then

                For Each drRow As DataRow In dt.Rows
                    strMessage = strMessage & drRow.Item("OEI_Ord_No").ToString & " " & drRow.Item("ExportTS").ToString & " " & drRow.Item("TimeElapsed").ToString & " seconds elapsed. Step: " & drRow.Item("Trigger_Message").ToString & "<br>"

                    If drRow.Item("Trigger_Message").trim() = "Step 0000" Then
                        strSql = "SELECT * FROM OEORDHDR_SQL WITH (NOLOCK) WHERE Filler_0001 = '" & drRow.Item("Ord_Guid").ToString.Trim & "' "
                        dtOrderCheck = db.DataTable(strSql)
                        If dtOrderCheck.Rows.Count <> 0 Then
                            ' OEI_IMPORT_TRIGGER NEVER REALLY STARTED, RESTART TRIGGER AFTER 15 MINUTES IN CASE OF.
                            If drRow.Item("TimeElapsed") > 900 Then
                                strSql = "EXEC DBO.OEI_IMPORT_TRIGGER '" & Trim(drRow.Item("Ord_Guid").ToString) & "' "
                                db.Execute(strSql)
                            End If
                        Else
                            ' ORDER NOT IMPORTED IN MACOLA. IF TIME IS OVER 15 MINUTES, UNFLAG EVERYTHING.
                            If drRow.Item("TimeElapsed") > 900 Then
                                ' TRIGGER COMPLETED BUT ORDER NOT FOUND, WE RESET THE FLAG TO FILE NOT CREATED AND NOTHING EXPORTED.
                                strSql = "UPDATE OEI_ORDHDR SET EXPORTTS = NULL, Trigger_Message = NULL, TriggerTS = NULL WHERE ORD_GUID = '" & drRow.Item("Ord_Guid").ToString.Trim & "' "
                                db.Execute(strSql)
                            End If

                        End If
                    ElseIf drRow.Item("Trigger_Message").trim() = "Complete" Then
                        strSql = "SELECT * FROM OEORDHDR_SQL WITH (NOLOCK) WHERE Filler_0001 = '" & drRow.Item("Ord_Guid").ToString.Trim & "' "
                        dtOrderCheck = db.DataTable(strSql)
                        If dtOrderCheck.Rows.Count = 0 Then
                            If drRow.Item("TimeElapsed") > 900 Then
                                ' TRIGGER COMPLETED BUT ORDER NOT FOUND, WE RESET THE FLAG TO FILE NOT CREATED AND NOTHING EXPORTED.
                                ' THIS WILL HAPPEN IF THE SP IS RUN MANUALLY WHEN THE ORDER DOES NOT EXIST
                                strSql = "UPDATE OEI_ORDHDR SET EXPORTTS = NULL, Trigger_Message = NULL, TriggerTS = NULL WHERE ORD_GUID = '" & drRow.Item("Ord_Guid").ToString.Trim & "' "
                                db.Execute(strSql)
                            End If
                        End If
                    Else
                        ' OEI_IMPORT_TRIGGER IS STARTED BUT STOPPED SOMEWHERE. 
                        ' RESTART TRIGGER.
                        strSql = "EXEC DBO.OEI_IMPORT_TRIGGER '" & drRow.Item("Ord_Guid").ToString.Trim & "' "
                        db.Execute(strSql)

                    End If

                Next

                strMessage = strMessage & " ""</p>"" "

            Else


            End If

            dt.Dispose()

        Catch er As Exception
            strMessage = "Error in HVControlService." & New Diagnostics.StackFrame(1).GetMethod.Name & "<br>" & er.Message
        End Try


    End Sub

    Private Sub Validate_Cus_No_Exist(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UcTaxes1.VisibleChanged, UcSalesperson1.VisibleChanged, UcOrderTotal1.VisibleChanged, UcImprint1.VisibleChanged, UcHeader1.VisibleChanged, UcExtra1.VisibleChanged, UcDocument1.VisibleChanged, UcCreditInfo1.VisibleChanged, UcContacts1.VisibleChanged

    End Sub


    Private Function SplitLargeSentence(ByVal pstrSentence As String, ByVal pintWidth As Integer, ByVal pstrSplitChar As String) As String

        Dim iPos As Integer

        SplitLargeSentence = ""

        While pstrSentence.Length > pintWidth
            If pstrSentence.Contains(" ") Then
                iPos = pstrSentence.LastIndexOf(" ", pintWidth)
                SplitLargeSentence = SplitLargeSentence & Mid(pstrSentence, 1, iPos) & pstrSplitChar
                If iPos + 1 >= pstrSentence.Length Then
                    pstrSentence = ""
                Else
                    pstrSentence = Mid(pstrSentence, iPos + 1).Trim
                End If
            Else
                iPos = pstrSentence.Length
            End If

        End While

    End Function


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim strTexte As String
        Dim strReponse As String

        strTexte = "Comme vous l'avez certainement déjà remarqué, nous  avons installé des panneaux « Réservé » dans l'aire de stationnement réservé pour les employés de Spector & Co.  De plus, Maggie a distribué, avec la dernière paie, un autocollant à tous les employés qui utilisent le stationnement.  Cet autocollant doit être placé dans le pare-brise arrière ou avant de votre véhicule (selon les indications incluses dans la lettre).  Si vous utilisez le stationnement et n’avez pas reçu l’autocollant, svp, envoyez un courriel à Maggie et elle vous en remettra un immédiatement."
        strReponse = SplitLargeSentence(strTexte, 80, vbCrLf)

    End Sub

    Private Sub cmdEDI_Alert_New_Inv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEDI_Alert_New_Inv.Click

        Try

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String
            Dim strSend_To As String

            strSql = _
            "SELECT     ISNULL(edi_alert_new_inv, '') AS edi_alert_new_inv, " & _
            "           ISNULL(edi_send_alert_to, '') AS edi_send_alert_to " & _
            "FROM       OEI_CONFIG WITH (Nolock) "

            dt = db.DataTable(strSql)

            If dt.Rows(0).Item("edi_alert_new_inv").ToString.Trim = "Y" Then

                strSend_To = dt.Rows(0).Item("edi_send_alert_to").ToString.Trim

                Dim strEmail As String = "New invoices were found for the following orders:<br>"
                strEmail &= "Ord_Guid".PadRight(40) & "Ord_No".PadRight(15) & "Inv_No".PadRight(15) & "<br>"

                strSql = _
                "SELECT      H.Ord_GUID, H.ord_no, I.INV_NO " & _
                "FROM		_EDISOURCE_HISTORY E WITH (NOLOCK) " & _
                "INNER JOIN	OEI_ORDHDR H WITH (NOLOCK) ON E.ord_no = H.ord_no AND H.ord_type = 'O' " & _
                "INNER JOIN	oehdrhst_sql I WITH (NOLOCK) ON H.ord_no = I.ord_no AND H.ord_type = I.ord_type " & _
                "WHERE		H.inv_no = '' "

                dt = db.DataTable(strSql)

                If dt.Rows.Count <> 0 Then

                    For Each dtRow As DataRow In dt.Rows

                        strEmail &= dtRow.Item("Ord_Guid").ToString.Trim.PadRight(40) & dtRow.Item("Ord_No").ToString.Trim.PadRight(15) & dtRow.Item("Inv_No").ToString.Trim.PadRight(15) & "<br>"

                        strSql = "UPDATE OEI_ORDHDR SET Inv_No = '" & dtRow.Item("Inv_No").ToString & "' WHERE ORD_GUID = '" & dtRow.Item("Ord_Guid").ToString & "' "
                        db.Execute(strSql)

                    Next

                End If

            End If

        Catch er As Exception
            'strMessage = "Error in HVControlService." & New Diagnostics.StackFrame(1).GetMethod.Name & "<br>" & er.Message
        End Try

    End Sub

    Private Sub cmdReset_Cus_No_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset_Cus_No.Click

        Try

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String

            strSql = _
            "SELECT		h.ord_guid, H.OEI_ORD_NO, h.ord_no, h.ord_type, h.cus_no AS Old_Cus_No, O.cus_no AS New_Cus_No " & _
            "FROM		oei_ordhdr h WITH (Nolock) " & _
            "INNER JOIN	oeordhdr_sql o with (nolock) on h.ord_no = o.ord_no and h.ord_type = o.ord_type " & _
            "WHERE      H.cus_no <> O.cus_no "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                For Each dtRow As DataRow In dt.Rows

                    strSql = "UPDATE OEI_ORDHDR SET Cus_No = '" & dtRow.Item("New_Cus_No").ToString.Trim & "' WHERE ORD_GUID = '" & dtRow.Item("Ord_Guid").ToString & "' "
                    db.Execute(strSql)

                Next

            End If

        Catch er As Exception
            'strMessage = "Error in HVControlService." & New Diagnostics.StackFrame(1).GetMethod.Name & "<br>" & er.Message
        End Try

    End Sub

    Private Sub cmdEDIEditSuppItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEDIEditSuppItem.Click

        If g_oOrdline.Extra_2_Bln Then

            Dim db As New cDBA
            Dim dt As DataTable

            Dim strsql As String = _
            "SELECT * FROM MDB_CUS_DEC WHERE SHORT_CTL_NO = '" & g_oOrdline.User_Def_Fld_4.Trim & "' "

            dt = db.DataTable(strsql)
            If dt.Rows.Count <> 0 Then

                Dim ofrmSupp_Item As New frmMDB_CUS_DEC_SUPP_ITEM()
                ofrmSupp_Item.Cus_Dec_ID = 1
                ofrmSupp_Item.ShowDialog()

                ofrmSupp_Item.Dispose()

            End If

        End If

    End Sub


    Private Sub dgvItems_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.CellContentClick

        Debug.Print("dgvItems_CellContentClick")

        Try
            Debug.Print(e.ColumnIndex)

            Select Case e.ColumnIndex
                Case Columns.Route, Columns.RouteID

                    dgvItems.EndEdit()

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdETA_Calculator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdETA_Calculator.Click

        Dim oCalc As New cETA_Calculator
        oCalc.Item_No = "00G12390GMT"
        oCalc.Location = "1"
        oCalc.Qty_Available = -8872
        oCalc.Qty_On_Hand = -8872
        oCalc.Qty_Ordered = 200000
        oCalc.Qty_Prev_Bkord = 191128
        oCalc.Curr_Cd = "USA"
        oCalc.Calculate_ETA_For_Item()
        Debug.Print(oCalc.Extra_9 & " " & oCalc.Extra_10)
        Debug.Print("")

    End Sub

    Private Sub dgvItems_CellErrorTextNeeded(sender As Object, e As System.Windows.Forms.DataGridViewCellErrorTextNeededEventArgs) Handles dgvItems.CellErrorTextNeeded

    End Sub

    Private Sub dgvItems_RowContextMenuStripChanged(sender As Object, e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItems.RowContextMenuStripChanged

    End Sub

    Private Sub UcOrder_White_Flag_Customer(sender As Object, e As EventArgs) Handles UcOrder.White_Flag_Customer

        lblWhite_Glove.Visible = True

    End Sub

    Private Sub tsbXmlExport_Click(sender As System.Object, e As System.EventArgs) Handles tsbXmlExport.Click

        Try

            Call m_oOrder.XML_Export(m_oOrder.Ordhead.OEI_Ord_No)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub XML_Export()

        Try

            Call SaveUncommited()

            If m_oOrder.OrdLines.Count = 0 Then Throw New OEException(OEError.Order_Not_Exported_No_Item_Lines)

            Dim oLine As cOrdLine
            Dim intUndefinedRouteCount As Integer = 0

            For Each oLine In m_oOrder.OrdLines

                'if stocked flag, nor route and not 88 and OEC or 44 and PRC, add to undefined counter. Exclude 88RANDOMSAMP (currently 'N', 88 and OEC)
                If (oLine.Stocked_Fg = "Y" Or oLine.Item_No.ToString.Trim() = "88RANDOMSAMP") And oLine.RouteID = 0 And Not _
                   ( _
                       (oLine.Prod_Cat = "PRC" And oLine.Item_No.ToString.Substring(0, 2) = "44") _
                       Or _
                       (oLine.Prod_Cat = "OEC" And oLine.Item_No.ToString.Substring(0, 2) = "88" And oLine.Item_No.ToString.Trim() <> "88RANDOMSAMP") _
                   ) _
                Then
                    intUndefinedRouteCount += 1
                End If

            Next

            Dim mbrResult As New MsgBoxResult

            If intUndefinedRouteCount <> 0 Then
                mbrResult = MsgBox("Some stock items have routes undefined. Do you wish to continue ?", MsgBoxStyle.YesNo)
            Else
                mbrResult = MsgBox("Order is ready to export. Do you wish to continue ?", MsgBoxStyle.YesNo)
            End If

            If mbrResult = MsgBoxResult.Yes Then ' Or intUndefinedRouteCount = 0 Then

                If Not (m_oOrder.CancelExport) Then

                    Call Save()

                    Call m_oOrder.XML_Export(m_oOrder.Ordhead.OEI_Ord_No)

                End If

            End If

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdTest_Click(sender As System.Object, e As System.EventArgs) Handles cmdTest.Click

        CreateExcelFile1()
        CreateExcelFile2()

    End Sub

    Private Sub updateImprintTripValidation(cCell As Integer)
        'enter qty ordered cell
        dgvItems.CurrentCell = dgvItems.Rows(cCell).Cells(Columns.Imprint)
        dgvItems.BeginEdit(True)
        dgvItems.EndEdit(True)

        'switch cells to force validation
        dgvItems.CurrentCell = dgvItems.Rows(cCell).Cells(Columns.Imprint_Color)
        dgvItems.BeginEdit(True)
        dgvItems.EndEdit(True)
    End Sub

    Private Sub updatePromiseTripValidation(cCell As Integer)
        'enter qty ordered cell
        dgvItems.CurrentCell = dgvItems.Rows(cCell).Cells(Columns.Promise_Dt)
        dgvItems.BeginEdit(True)
        dgvItems.EndEdit(True)

        'switch cells to force validation
        dgvItems.CurrentCell = dgvItems.Rows(cCell).Cells(Columns.Req_Ship_Dt)
        dgvItems.BeginEdit(True)
        dgvItems.EndEdit(True)
    End Sub


    Private Sub updateQtyAdjusterTripValidation(cCell As Integer)
        'enter qty ordered cell
        dgvItems.CurrentCell = dgvItems.Rows(cCell).Cells(Columns.Qty_Ordered)
        dgvItems.BeginEdit(True)
        dgvItems.EndEdit(True)

        'switch cells to force validation
        dgvItems.CurrentCell = dgvItems.Rows(cCell).Cells(Columns.Qty_To_Ship)
        dgvItems.BeginEdit(True)
        dgvItems.EndEdit(True)

        'experimental price adjuster (only when ignore is false to avoid loops) - copy this to logical area
        'If Not ignoreQtyMismatch And Not dgvItems.Rows(dgvItems.CurrentRow.Index).Cells(Columns.Item_No).Value = "" Then
        '    adjustUnitPricing()
        'End If
    End Sub

    Private Sub add44Charges()
        Try

            Dim i As Integer 'counter
            Dim currIndex As Integer

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String
            Dim itemExcludeList As String = ""
            Dim excludeFilter As String = ""

            If Not ignoreQtyMismatch Then ignoreQtyMismatch = True 'ignore quantity mismtach warning when price adjusting

            'Build item exclusion string
            For i = 0 To dgvItems.RowCount - 1
                If Not String.IsNullOrEmpty(dgvItems.Rows(i).Cells(Columns.Item_No).Value.ToString) Then
                    itemExcludeList = itemExcludeList & "'" & dgvItems.Rows(i).Cells(Columns.Item_No).Value.ToString & "',"
                End If
            Next

            'append item exclude list if appliable
            If itemExcludeList <> "" Then
                excludeFilter = " AND cd_tp_1_item_no NOT IN (" & itemExcludeList.Substring(0, itemExcludeList.Length - 1) & ")"
            End If

            'remove charges for order type GUC and GUU


            'query to find all automatic 44 charges
            strSql = "SELECT cd_tp_1_item_no, minimum_qty_1 as qty, prc_or_disc_1 as price " & _
                     "FROM oeprcfil_sql " & _
                     "WHERE cd_tp = 1 AND cd_tp_3_cust_type = '' AND cd_tp_2_prod_cat = '' " & _
                     "  AND extra_5 = 1 AND cd_tp_1_item_no like '44%' " & _
                     "  AND ISNULL(start_dt, GETDATE()) <= CAST(GETDATE() AS DATE) AND ISNULL(end_dt, GETDATE()) >= CAST(GETDATE() AS DATE) " & _
                     "  AND cd_tp_1_cust_no = '" & m_oOrder.Ordhead.Cus_No & "' AND cd_tp_1_item_no IN ('44PAPR00PRF', '44EXCT00QNT') " & _
                     "  " & excludeFilter & " GROUP BY cd_tp_1_item_no, minimum_qty_1, prc_or_disc_1"


            dt = db.DataTable(strSql)

            'add each charge found
            If dt.Rows.Count <> 0 Then
                For Each dtRow As DataRow In dt.Rows
                    cmdInsertLine.PerformClick()
                    dgvItems.CurrentRow.Cells(Columns.Item_No).Value = dtRow.Item("cd_tp_1_item_no")
                    currIndex = dgvItems.CurrentRow.Index
                    updateQtyAdjusterTripValidation(currIndex) 'trip item addition and validation
                Next
            End If

        Catch er As Exception
            MsgBox("44 code acquisition error: " & er.Message)
        End Try

        'reset mismatch var
        ignoreQtyMismatch = False
    End Sub

    Private Sub adjustUnitPricing()
        'OEError.Unit_Price_Has_Been_Changed()
        m_lElementCount = 6
        Dim i As Integer 'counter
        Dim curr_item_style As String
        Dim curr_item_qty As Integer
        Dim styleQtyCnt As New Dictionary(Of String, Integer) 'create generic associative container

        'original values for quantities
        Dim oriOrdQuantity As Integer
        Dim oriShipQuantity As Integer
        Dim newPricePoint As Double

        'calculate total quantity for each style 
        For i = 0 To dgvItems.RowCount - 1

            'ignore kit components and assorted items
            If dgvItems.Rows(i).Cells(Columns.Location).Value.ToString = "" Or dgvItems.Rows(i).Cells(Columns.Item_No).Value.ToString.Contains("AST") Or Mid(dgvItems.Rows(i).Cells(Columns.Item_No).Value.ToString, 1, 2) = "44" Then Continue For

            If Not ignoreQtyMismatch Then ignoreQtyMismatch = True 'ignore quantity mismtach warning when price adjusting

            curr_item_style = dgvItems.Rows(i).Cells(Columns.Item_Cd).Value.ToString.Trim()
            curr_item_qty = dgvItems.Rows(i).Cells(Columns.Qty_Ordered).Value

            If styleQtyCnt.ContainsKey(curr_item_style) Then
                styleQtyCnt(curr_item_style) = styleQtyCnt(curr_item_style) + curr_item_qty
            Else
                styleQtyCnt.Add(curr_item_style, curr_item_qty)
            End If
        Next

        'TODO: determine price bracket for each item
        For i = 0 To dgvItems.RowCount - 1

            'get style
            curr_item_style = dgvItems.Rows(i).Cells(Columns.Item_Cd).Value.ToString.Trim()
            If Not styleQtyCnt.ContainsKey(curr_item_style) Or dgvItems.Rows(i).Cells(Columns.Location).Value.ToString = "" OrElse styleQtyCnt(curr_item_style) <= 0 Then
                Continue For 'ignore if key not found (should not occur)
            End If

            'get original ship values
            oriOrdQuantity = dgvItems.Rows(i).Cells(Columns.Qty_Ordered).Value
            oriShipQuantity = dgvItems.Rows(i).Cells(Columns.Qty_To_Ship).Value

            'assign new quantity and validate
            dgvItems.Rows(i).Cells(Columns.Qty_Ordered).Value = styleQtyCnt(curr_item_style)
            Call updateQtyAdjusterTripValidation(i) 'validate

            newPricePoint = dgvItems.Rows(i).Cells(Columns.Unit_Price).Value 'get adjusted unit price

            'Restore original values and assign new pricing'
            dgvItems.Rows(i).Cells(Columns.Qty_Ordered).Value = oriOrdQuantity
            dgvItems.Rows(i).Cells(Columns.Qty_To_Ship).Value = oriShipQuantity

            Call updateQtyAdjusterTripValidation(i) 'validate again

            dgvItems.Rows(i).Cells(Columns.Unit_Price).Value = newPricePoint

            'Save Price
            g_oOrdline.m_Unit_Price = newPricePoint
        Next

        'reset mismatch var
        ignoreQtyMismatch = False
        MsgBox("Pricing has been adjusted where applicable")
    End Sub

    Private Function determineKitQty() As Integer
        Dim kit_amt = Nothing
        Dim i As Integer

        If g_oOrdline.Kit_Item_Guid IsNot Nothing AndAlso g_oOrdline.Kit_Item_Guid.ToString <> "" Then
            For i = 0 To dgvItems.RowCount - 1
                If dgvItems.Rows(i).Cells(Columns.Item_Guid).Value = g_oOrdline.Kit_Item_Guid Then
                    kit_amt = dgvItems.Rows(i).Cells(Columns.Qty_Ordered).Value
                End If
            Next
        End If

        Return kit_amt
    End Function

    Private Function determineReqDay() As Date
        Dim reqDay = Nothing
        Dim i As Integer

        If g_oOrdline.Kit_Item_Guid IsNot Nothing AndAlso g_oOrdline.Kit_Item_Guid.ToString.Trim() <> "" Then
            For i = 0 To dgvItems.RowCount - 1
                If dgvItems.Rows(i).Cells(Columns.Item_Guid).Value = g_oOrdline.Kit_Item_Guid Then
                    reqDay = dgvItems.Rows(i).Cells(Columns.Promise_Dt).Value
                    Exit For
                End If
            Next
        ElseIf g_oOrdline.m_Promise_Dt = Nothing Then
            reqDay = dgvItems.CurrentRow.Cells(Columns.Promise_Dt).Value
        Else
            reqDay = g_oOrdline.m_Promise_Dt
        End If

        Return reqDay
    End Function

    Private Sub execThermoSimulation()
        'Exit Sub 'thermo temporarily offline. Please remove once re-enabled

        'simulate if not empty route (0), completed (39), random sample(16) or no imprint (17). It should also not run during price adjustment procedure (ignoreQtyMismatch)
        If g_oOrdline.m_RouteID > 0 AndAlso ignoreQtyMismatch = False AndAlso validateRoute(g_oOrdline.m_RouteID) = True Then
            g_oOrdline.m_Promise_Dt = determineReqDay()
            dgvItems.BeginInvoke(New Del_PerformProductionAvailalbilitySimulation(AddressOf g_oOrdline.PerformProductionAvailalbilitySimulation), dgvItems.CurrentRow.Cells(Columns.Promise_Dt), g_oOrdline.ImprintLine, False, determineOSREligibilityList(), determineKitQty())
            'dgvItems.BeginInvoke(New Del_SetProductionAvailabilityColor(AddressOf g_oOrdline.SetProductionAvailabilityColor), dgvItems.CurrentRow.Cells(Columns.Item_No)) ' , dgvItems.CurrentCell.ColumnIndex)
        End If
    End Sub

    Private Function validateRoute(ByVal routeId As Integer) As Boolean
        Dim isValid = False

        'populate routelist if empty
        If validRouteList Is Nothing Then
            validRouteList = cImprint.getAllowableRouteList()
        End If

        'determine if chosen route is in the valid list
        For Each validRoute As Integer In validRouteList
            If validRoute = routeId Then
                isValid = True
                Exit For
            End If
        Next

        Return isValid
    End Function

    Private Function determineOSREligibilityList() As Dictionary(Of String, Boolean)
        Dim ITEM_CD = dgvItems.CurrentRow.Cells(Columns.Item_Cd).Value.ToString
        Dim i As Integer 'counter

        'define dictionaries
        Dim styleImpOsrElig As New Dictionary(Of String, Boolean) 'create generic associative container for elgibility
        Dim styleImpOsrTot As New Dictionary(Of String, Integer) 'create generic associative container for imprint totals


        For i = 0 To dgvItems.RowCount - 1
            'ignore kit components and assorted items and items where the item_cd does not match the current orderline in question
            If dgvItems.Rows(i).Cells(Columns.Location).Value.ToString = "" Or dgvItems.Rows(i).Cells(Columns.Item_No).Value.ToString.Contains("AST") _
                Or Mid(dgvItems.Rows(i).Cells(Columns.Item_No).Value.ToString, 1, 2) = "44" Or dgvItems.Rows(i).Cells(Columns.Item_Cd).Value <> ITEM_CD _
                Or IsDBNull(dgvItems.Rows(i).Cells(Columns.RouteID).Value) Then Continue For

            'get imrpint count totals
            Dim impr_counter = 1
            Dim impMethods As ArrayList = mThermometer.getAllApplicableRoutes(dgvItems.Rows(i).Cells(Columns.RouteID).Value)
            For Each impMethod As String In impMethods

                'determine which imprint count cell field to multiply the quantity by
                Dim ordLineNumImprCellIndex = 0
                Select Case impr_counter
                    Case 1
                        ordLineNumImprCellIndex = Columns.Num_Impr_1
                    Case 2
                        ordLineNumImprCellIndex = Columns.Num_Impr_2
                    Case Else
                        ordLineNumImprCellIndex = Columns.Num_Impr_3
                End Select

                'get total number of imprints for the deco method for a given orderline
                Dim imprints = Convert.ToInt32(dgvItems.Rows(i).Cells(Columns.Qty_Ordered).Value) '* dgvItems.Rows(i).Cells(ordLineNumImprCellIndex).Value
                ' MsgBox(dgvItems.Rows(i).Cells(ordLineNumImprCellIndex).Value)
                If styleImpOsrTot.ContainsKey(impMethod) Then
                    styleImpOsrTot(impMethod) = styleImpOsrTot(impMethod) + imprints
                Else
                    styleImpOsrTot.Add(impMethod, imprints)
                End If

                impr_counter = impr_counter + 1
            Next
        Next

        'build eligibility list
        For Each kvp As KeyValuePair(Of String, Integer) In styleImpOsrTot
            Dim maxUnits = getAllowableOrderUnitSizeByImprint(kvp.Key, 3, True)

            'eligible if maxUnits = 0 (Not defined) or totImprint <= maxUnits threshold
            Dim is_eligible = IIf(maxUnits < 0 OrElse (kvp.Value > maxUnits And kvp.Value > 0), False, True)
            styleImpOsrElig.Add(kvp.Key, is_eligible)
        Next


        Return IIf(styleImpOsrElig.Count = 0, Nothing, styleImpOsrElig)
    End Function

    Public Sub CreateExcelFile1()

        '10:     Dim errPos As Integer
20:     Dim xlApp As Excel.Application = New Excel.ApplicationClass
30:     Dim xlWorkBook As Excel.Workbook = Nothing

        Try

40:         Dim xlWorkSheet As Excel.Worksheet
50:         Dim misValue As Object = System.Reflection.Missing.Value

60:         xlWorkBook = xlApp.Workbooks.Add(misValue)
70:         xlWorkSheet = xlWorkBook.Sheets(1)

80:         xlWorkSheet.Cells(1, 1) = "Marc"
90:         xlWorkSheet.Cells(1, 2) = "Beauregard"

100:        Dim m_path As String = "c:\ExactTemp\"
110:        xlWorkBook.SaveAs(Filename:="" & m_path & "Fichier1.xls")

160:        MsgBox("File 1 created.")

            'CreateExcelFile = Settings.Export.SavePath & Ordhead.Ord_GUID & Settings.Export.FileExtension

        Catch er As Exception
            MsgBox("Error in frmOrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " Line: " & Erl())
        Finally
            Try
120:            xlWorkBook.Close(False)
130:            xlWorkBook = Nothing

140:            xlApp.Quit()
150:            xlApp = Nothing
            Catch e2 As Exception
                MsgBox("Error in frmOrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & e2.Message & " Line: " & Erl())
            End Try
        End Try

    End Sub

    Public Sub CreateExcelFile2()

        '10:     Dim errPos As Integer
20:     Dim xlApp As Excel.Application = New Excel.ApplicationClass
30:     Dim xlWorkBook As Excel.Workbook = Nothing

        Try

40:         Dim xlWorkSheet As Excel.Worksheet

50:         xlWorkBook = xlApp.Workbooks.Add()
60:         xlWorkSheet = xlWorkBook.Sheets(1)

70:         xlWorkSheet.Cells(1, 1) = "Marc"
80:         xlWorkSheet.Cells(1, 2) = "Beauregard"

90:         Dim m_path As String = "\\nickel\Harvest Ventures\HV_MacolaImport\"
100:        xlWorkBook.SaveAs(Filename:="" & m_path & "Fichier2.xls")

110:        MsgBox("File 2 created.")

            'CreateExcelFile = Settings.Export.SavePath & Ordhead.Ord_GUID & Settings.Export.FileExtensio

        Catch er As Exception
            MsgBox("Error in frmOrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " Line: " & Erl())
        Finally
            Try
120:            xlWorkBook.Close(False)
130:            xlWorkBook = Nothing

140:            xlApp.Quit()
150:            xlApp = Nothing
            Catch e2 As Exception
                MsgBox("Error in frmOrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & e2.Message & " Line: " & Erl())
            End Try
        End Try

    End Sub



    Private Sub btnDbgPrxAdjust_Click(sender As System.Object, e As System.EventArgs) Handles btnDbgPrxAdjust.Click
        Call adjustUnitPricing() 'debug price adjuster
    End Sub

    Private Sub testSimulateThermo_Click(sender As System.Object, e As System.EventArgs) Handles testSimulateThermo.Click
        Dim impMethod As String = "Laser"
        Dim units As Integer = 100
        Dim item_no As String = "00B10400BLK"
        Dim defaultMachine As String = "D42" ' mThermometer.getDefaultMachine(item_no, impMethod)

        Dim newDate As Date = mThermometer.dateAdjuster(Date.Now.AddDays(1).ToShortDateString, 2)
        MsgBox(newDate.ToString)

        Dim timeSpecs() As Decimal = mThermometer.getTimeSpecs(item_no, impMethod)
        Dim checkdate = Date.Now().ToShortDateString 'date to try against

        'machine data
        Dim machineData = mThermometer.getMachineDataViaRest(defaultMachine, checkdate)

        'check current day and machine only
        Dim singleFinalResult As String = checkAvailability(Date.Now().ToShortDateString, machineData, units, timeSpecs, False)
        MsgBox("Raw data for the day for default machine: " & vbCrLf & vbCrLf & singleFinalResult.ToString)

        'check current day only all relevant machines
        Dim singleFinalResultAllMachines() = mThermometer.simulateNextAvailDateAndMachine(impMethod, units, timeSpecs, defaultMachine, False, checkdate, False).Split("_")
        MsgBox("Analysis result for single day: " & vbCrLf & vbCrLf & singleFinalResultAllMachines(1).ToString)

        'full analsysis sim check
        Dim groupFinalResult() = mThermometer.simulateNextAvailDateAndMachine(impMethod, units, timeSpecs, defaultMachine, False, checkdate).Split("_")
        MsgBox("Full analysis result: " & vbCrLf & vbCrLf & groupFinalResult(1).ToString)
    End Sub

    Private Sub UcOrder_Load(sender As Object, e As EventArgs) Handles UcOrder.Load

    End Sub

    Private Sub tbOrder_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles tbOrder.ItemClicked

    End Sub
End Class

