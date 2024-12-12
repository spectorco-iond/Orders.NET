Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Windows.Media.Imaging
Imports System.Collections.Generic

Public Class cOrdLine

    ' Pour isoler les produits qui sont seulement des charges, il faut degarder dans:
    ' Le champ PROD_CAT de IMITMIDX_SQL pour les valeurs 'PRC', 'OEC', 'VOL'

    Public m_Ord_Guid As String = ""

    ' FirstLoad is used to get all data according to the entry, as in a normal entry
    Public blnFirstLoad As Boolean = False

    ' Loading is used when we already have the data, like for refreshing a grid
    Public blnLoading As Boolean

    'Popup message log
    Public Shared popMessageLog As New Dictionary(Of String, String)

    ' Flags to activate creation of different components of the item
    Public m_Create_Traveler As Boolean = False
    Public m_Create_Imprint As Boolean = False
    Public m_Create_Kit As Boolean = False

    ' Item guid is the current item, Kit_Item_Guid is the parent of a kit component
    Public m_Item_Guid As String
    Public m_Kit_Item_Guid As String
    Public m_Item_Exist As Boolean = False

    Public m_Record_Type As String = "L" 'Char (1) 'L' for Line Item' NOT TO BE RELOADED
    ' Private m_oOrder As cOrder

    Public m_Ord_Type As String 'Char (8)
    Public m_Ord_No As String 'Char (8)       Order No
    Public m_Line_Seq_No As Integer 'Int            Line Sequence No
    Public m_Item_No As String 'Char (15) 'Public Sub Item_No_Changed()

    'UPGRADE_NOTE: Loc was upgraded to Loc_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public m_Loc As String 'Char (3) 'Public Sub Loc_Changed()

    Public m_Pick_Seq As String 'Char (8)
    Public m_Cus_Item_No As String 'Char (30)
    Public m_Item_Desc_1 As String 'Char (30)      Item Description Line 1
    Public m_Item_Desc_2 As String 'Char (30)      Item Description Line 2
    Public m_Qty_Ordered As Double 'Decimal (13,4) 
    Public m_Qty_To_Ship As Double 'Decimal (13,4) 
    Public m_Qty_To_Ship_Calc As Double
    ' Qty_To_Ship agit comme Qty_Shipped pour l'automation.
    'Private m_Qty_Prev_To_Ship As Double 'Decimal (13,4)

    '==============================
    'ADDED June 6, 2017 - T. Louzon
    Public m_Unit_Price_BeforeSpecial As Double
    '==============================
    '==============================
    'ADDED June 27, 2017 - T. Louzon
    Public m_Mfg_Item_No As String
    Public m_RouteExtension As String
    '==============================
    'ADDED June 29, 2017 - T. Louzon
    Public m_IsNoStock As Integer
    Public m_IsNoArtwork As Integer
    '==============================


    Public m_Unit_Price As Double 'Decimal (16,6) Unit Price  If not populated, default 'Price' from IM Loca
    'Public Sub Unit_Price_Changed()
    Public m_Discount_Pct As Double 'Numeric NULL
    Public m_Request_Dt As Date 'DateTime NULL
    Public m_Qty_Prev_To_Ship As Double 'Decimal (13,4) 'Public Sub Qty_To_Ship_Changed() 
    Public m_Qty_Prev_Bkord As Double 'Numeric NULL
    Public m_Qty_Return_To_Stk As Double 'Numeric NULL
    Public m_Bkord_Fg As String 'Char(1) NULL ' BkOrd_Fg
    Public m_Uom As String 'Char(2) NULL
    Public m_Uom_Ratio As Double 'Numeric NULL
    Public m_Unit_Cost As Double 'Numeric NULL
    Public m_Unit_Weight As Double 'Numeric NULL
    Public m_Comm_Calc_Type As String 'Char(1) NULL ' Recalc ??
    Public m_Comm_Pct_Or_Amt As Double 'Numeric NULL
    Public m_Promise_Dt As Date 'DateTime NULL
    Public m_Tax_Fg As String 'Char(1) NULL
    Public m_Stocked_Fg As String = "" 'Char(1) NULL
    Public m_Controlled_Fg As String = "" 'Char(1) NULL
    Public m_Select_Cd As String 'Char(1) NOT NULL ' Selected
    Public m_Tot_Qty_Ordered As Double 'Numeric NULL
    Public m_Tot_Qty_Shipped As Double 'Numeric NULL
    Public m_Tax_Fg_1 As String 'Char(1) NULL
    Public m_Tax_Fg_2 As String 'Char(1) NULL
    Public m_Tax_Fg_3 As String 'Char(1) NULL
    Public m_Orig_Price As Double 'Numeric NULL
    Public m_Copy_To_Bm_Fg As String 'Char(1) NULL
    Public m_Explode_Kit As String 'Char(1) NULL
    Public m_Mfg_Ord_No As String 'Char(8) NULL
    Public m_Allocate_Dt As Date 'DateTime NULL
    Public m_Last_Post_Dt As Date 'DateTime NULL
    Public m_Post_To_Inv_Qty As Double 'Numeric NULL
    Public m_Posted_To_Inv As Double 'Numeric NULL
    Public m_Tot_Qty_Posted As Double 'Numeric NULL
    Public m_Qty_Allocated As Double 'Numeric NULL
    Public m_Components_Alloc As Double 'Numeric NULL
    Public m_Bin_Fg As String 'Char(1) NULL
    Public m_Cost_Meth As String 'Char(1) NULL
    Public m_Ser_Lot_Cd As String 'Char(1) NULL
    Public m_Mult_Ftr_Fg As String 'Char(1) NULL
    Public m_Line_Type As String 'Char(1) NULL
    Public m_Prod_Cat As String 'Char(3) NULL
    Public m_End_Item_Cd As String 'Char(1) NULL
    Public m_Reason_Cd As String 'Char(3) NULL
    Public m_Feature_Return As String 'Char(1) NULL
    Public m_Rec_Inspection As String 'Char(1) NULL
    Public m_Ship_From_Stk As String 'Char(1) NULL
    Public m_Mult_Release As String 'Char(1) NULL
    Public m_Req_Ship_Dt As Date 'DateTime NULL
    Public m_Qty_From_Stk As Double 'Numeric NULL
    Public m_User_Def_Fld_1 As String 'Char(30) NULL
    Public m_User_Def_Fld_2 As String 'Char(30) NULL
    Public m_User_Def_Fld_3 As String 'Char(30) NULL
    Public m_User_Def_Fld_4 As String 'Char(30) NULL
    Public m_User_Def_Fld_5 As String 'Char(30) NULL
    Public m_Picked_Dt As Date 'DateTime NULL
    Public m_Shipped_Dt As Date 'DateTime NULL
    Public m_Billed_Dt As Date 'DateTime NULL
    Public m_Update_Fg As String 'Char(1) NULL
    Public m_Prc_Cd_Orig_Price As Double 'Numeric NULL
    Public m_Tax_Sched As String = "" 'Char(5) NULL
    Public m_Cus_No As String = "" 'Char(20) NULL
    Public m_Tax_Amt As Double 'Numeric NULL
    Public m_Qty_Prev_Bkord_Fg As String 'Char(1) NOT NULL
    Public m_Line_No As Short = 1 'SmallInt NOT NULL
    Public m_Mfg_Method As String = "" 'Char(2) NULL
    Public m_Forced_Demand As String 'Char(1) NULL
    Public m_Conf_Pick_Dt As Date 'DateTime NULL
    Public m_Item_Release_No As String 'Char(8) NULL
    Public m_Bin_Ser_Lot_Comp As String 'Char(1) NULL
    Public m_Offset_Used_Fg As String 'Char(1) NULL
    Public m_Ecs_Space As String 'Char(20) NULL
    Public m_Sfc_Order_Status As String 'Char(1) NULL
    Public m_Total_Cost As Double 'Numeric NULL
    Public m_Po_Ord_No As String = "" 'Char(8) NULL
    Public m_Rma_Seq As Short 'SmallInt NOT NULL
    Public m_Vendor_No As String 'Char(20) NULL
    Public m_Posted_Unit_Cost As Double 'Numeric NULL
    Public m_Extra_1 As String 'Char(1) NULL
    Public m_Extra_2 As String = "" 'Char(1) NULL
    Public m_Extra_3 As String 'Char(1) NULL
    Public m_Extra_4 As String 'Char(1) NULL
    Public m_Extra_5 As String 'Char(1) NULL
    Public m_Extra_6 As String = "" 'Char(8) NULL
    Public m_Extra_7 As String 'Char(8) NULL
    Public m_Extra_8 As String 'Char(12) NULL
    Public m_Extra_9 As String = "" 'Char(12) NULL
    Public m_Extra_10 As Double = 0 'Numeric NULL
    Public m_Extra_11 As Double 'Numeric NULL
    Public m_Extra_12 As Double 'Numeric NULL
    Public m_Extra_13 As Double 'Numeric NULL
    Public m_Extra_14 As Integer 'Int NULL
    Public m_Extra_15 As Integer 'Int NULL
    Public m_Warranty_Date As Date 'DateTime NULL
    Public m_Revision_No As String = "" 'Char(8) NULL
    Public m_Cm_Post_Fg As String 'Char(1) NULL
    Public m_Recalc_Sw As String 'Char(1) NOT NULL Recalc
    Public m_Filler_0004 As String 'Char(132) NULL
    Public m_Id As Double 'Numeric NOT NULL

    Public m_Refill_Fg As Boolean
    Public m_Is_Refill_Fg As Boolean
    Public m_Has_Refills_Fg As Boolean

    Public m_Image As Image
    Public m_ImageWidth As Integer
    Public m_ImageHeight As Integer

    ' Additionnal fields 
    'Additional fields:
    Public m_Qty_On_Hand As Double 'Numeric NOT NULL
    Public m_Qty_Inventory As Double
    Public m_Qty_Lines As Double = 1

    Public m_Calc_Price As Double

    Public m_Route As String = ""
    Public m_RouteID As Integer
    Public m_ProductProof As Integer
    Public m_AutoCompleteReship As Integer

    Public m_Activity_Cd As String

    Public m_SaveToDB As Boolean = False
    Public m_Dirty As Boolean = False

    Public m_Imprint As cImprint
    Public m_Traveler As CTraveler

    ' If item is a kit Item, get kit elements
    Public m_Kit As cKit

    'If item is a kit element, set KitElement to true
    Public m_Kit_Component As Boolean = False

    ' for frmProductLineEntry, we don't want kits to trigger.
    'Public m_No_Kits As Boolean = False

    Private m_Source As OEI_SourceEnum '= OEI_SourceEnum.frmOrder

    Public m_OEI_W_Pixels As Integer = 400
    Public m_OEI_H_Pixels As Integer = 400
    Public m_OEI_Image_Rotate As String = ""

    Public m_OEOrdLin_Id As Integer = 0

    Public m_Comp_Item_Guid As String
    Public m_Item_Cd As String

    Public m_Qty_Scrap As Double
    Public m_Qty_Overpick As Double

    Private ETA_From_Item_No As String

    Public m_Loc_Status As String
    Public m_isMixMatch As Boolean
    Public m_Mix_Group As Integer

    'checks for oei thermo simulation
    Private lastThermoCheckedDate As Date = Nothing
    Private lastThermoRouteId As Integer = Nothing
    Private lastThermoQty As Double = Nothing
    Private lastThermo_Num_Impr_1 As Integer = Nothing
    Private lastThermo_Num_Impr_2 As Integer = Nothing
    Private lastThermo_Num_Impr_3 As Integer = Nothing

    'Public Sub New()
    '    m_oOrder = New cOrder
    'End Sub

#Region "Public constructors ##################################################"

    Public Sub New() ' ByRef poOrder As cOrder)

        m_Ord_Guid = m_oOrder.Ordhead.Ord_GUID
        Call GenerateGuidForOrdline()

        Call Set_New_Line_No()

        'Call Reset()

        ' m_oOrder = poOrder
        If m_Create_Imprint Then m_Imprint = New cImprint(m_Item_Guid)
        If m_Create_Traveler Then m_Traveler = New CTraveler(m_Item_Guid)
        ' We must create the Kit to init IsAKit to false
        m_Kit = New cKit()

    End Sub

    'Public Sub New(ByVal pSource As OEI_SourceEnum) ' ByRef poOrder As cOrder)

    '    Call GenerateGuidForOrdline()
    '    m_Source = pSource

    '    Call Set_New_Line_No()

    '    'Call Reset()

    '    ' m_oOrder = poOrder
    '    If m_Create_Imprint Then m_Imprint = New cImprint(m_Item_Guid)
    '    If m_Create_Traveler Then m_Traveler = New CTraveler(m_Item_Guid)
    '    ' We must create the Kit to init IsAKit to false
    '    m_Kit = New cKit()

    'End Sub

    Public Sub New(ByVal pOrd_Guid As String, ByVal pSource As OEI_SourceEnum) ' ByRef poOrder As cOrder)

        m_Ord_Guid = pOrd_Guid
        m_Source = pSource

        Call GenerateGuidForOrdline()

        Call Set_New_Line_No()

        'Call Reset()

        ' m_oOrder = poOrder
        If m_Create_Imprint Then m_Imprint = New cImprint(m_Item_Guid)
        If m_Create_Traveler Then m_Traveler = New CTraveler(m_Item_Guid)
        ' We must create the Kit to init IsAKit to false
        m_Kit = New cKit(Me)

    End Sub

    Public Sub New(ByVal pOrd_Guid As String) ' ByRef poOrder As cOrder)

        '   popMessageLog = New Dictionary(Of String, String)

        m_Ord_Guid = pOrd_Guid
        m_Source = m_oOrder.Source

        Call GenerateGuidForOrdline()

        Call Set_New_Line_No()

        'Call Reset()

        ' m_oOrder = poOrder
        If m_Create_Imprint Then m_Imprint = New cImprint(m_Item_Guid)
        If m_Create_Traveler Then m_Traveler = New CTraveler(m_Item_Guid)
        ' We must create the Kit to init IsAKit to false
        m_Kit = New cKit(Me)

    End Sub

    Public Sub New(ByVal pOrd_Guid As String, ByVal pSaveToDB As Boolean) ' ByRef poOrder As cOrder, ByVal pSaveToDB As Boolean)

        m_Ord_Guid = pOrd_Guid

        Call GenerateGuidForOrdline()
        m_Source = m_oOrder.Source

        Call Set_New_Line_No()

        'm_oOrder = poOrder
        'Call Reset()
        m_SaveToDB = pSaveToDB

        If m_Create_Imprint Then m_Imprint = New cImprint(m_Item_Guid)
        If m_Create_Traveler Then m_Traveler = New CTraveler(m_Item_Guid)
        ' We must create the Kit to init IsAKit to false
        m_Kit = New cKit(Me)

    End Sub

    Public Sub New(ByVal pSaveToDB As Boolean) ' ByRef poOrder As cOrder, ByVal pSaveToDB As Boolean)

        Call GenerateGuidForOrdline()

        Call Set_New_Line_No()

        'm_oOrder = poOrder
        'Call Reset()
        m_SaveToDB = pSaveToDB

        If m_Create_Imprint Then m_Imprint = New cImprint(m_Item_Guid)
        If m_Create_Traveler Then m_Traveler = New CTraveler(m_Item_Guid)
        ' We must create the Kit to init IsAKit to false
        m_Kit = New cKit()

    End Sub


    'Public Sub New(ByVal pSource As OEI_SourceEnum) ' ByRef poOrder As cOrder, ByVal pSaveToDB As Boolean)

    '    Call GenerateGuidForOrdline()
    '    'm_oOrder = poOrder
    '    m_Source = pSource
    '    m_SaveToDB = (m_Source = OEI_SourceEnum.frmProductLineEntry)

    '    If m_Create_Imprint Then m_Imprint = New cImprint(m_Item_Guid)
    '    If m_Create_Traveler Then m_Traveler = New CTraveler(m_Item_Guid)
    '    If m_Create_Kit Then m_Kit = New cKit

    'End Sub

#End Region

    Private Sub GenerateGuidForOrdline()

        Dim vg As New VariableGuid
        m_Item_Guid = vg.Guid(30) ' Guid.NewGuid.ToString

    End Sub

    Private Sub Set_New_Line_No()

        If m_Ord_Guid <> "" Then
            m_Line_No = m_oOrder.GetMaxLine_No(m_Ord_Guid) + 1
        Else
            m_Line_No = 0
        End If

    End Sub

    Private Sub LoadItem(ByRef poDescriptions As cOrdheadDesc)

        Dim dt As DataTable
        Dim db As New cDBA

        Try
            blnLoading = True

            Dim strSql As String

            strSql = "" & _
            "SELECT TOP 1 * " & _
            "FROM   OEOrdHdr_Sql WITH (Nolock) " & _
            "WHERE  Ord_No = '" & Ord_No & "' "

            strSql = "" & _
            "SELECT 	TOP 1 O.*, ISNULL(Loc.Loc_Desc, '') as Location_Desc, ISNULL(SV.Code_Desc, '') AS Ship_Via_Desc, " & _
            "			CASE ISNULL(O.Status, '') WHEN '1' THEN 'Booked Order' WHEN '4' THEN 'Pick Ticket printed' WHEN '8' THEN 'Selected for Billing' WHEN 'C' THEN 'Credit Hold' WHEN 'I' THEN 'Order Incomplete' WHEN 'L' THEN 'Closed Order' ELSE 'Unknown' END AS Status_Desc, " & _
            "			ISNULL(Term.Description, '') as Term_Description, ISNULL(TS.Tax_Sched_Desc, '') AS Tax_Sched_Desc, " & _
            "			ISNULL(Tax1.Tax_Cd_Description, '') as Tax_Cd_1_Desc, " & _
            "			ISNULL(Tax2.Tax_Cd_Description, '') as Tax_Cd_2_Desc, " & _
            "			ISNULL(Tax3.Tax_Cd_Description, '') as Tax_Cd_3_Desc, " & _
            "			ISNULL(Sls1.Slspsn_Name, '') AS Slspsn_1_Desc, " & _
            "			ISNULL(Sls2.Slspsn_Name, '') AS Slspsn_2_Desc, " & _
            "			ISNULL(Sls3.Slspsn_Name, '') AS Slspsn_3_Desc " & _
            "FROM   	OEOrdHdr_Sql O    WITH (Nolock)  " & _
            "LEFT JOIN	IMLOCFIL_SQL Loc  WITH (Nolock) ON O.Mfg_Loc = Loc.Loc " & _
            "LEFT JOIN 	SYCDEFIL_SQL SV   WITH (Nolock) ON O.Ship_Via_Cd = SV.sy_code " & _
            "LEFT JOIN 	SYTRMFIL_SQL Term WITH (Nolock) ON O.AR_Terms_CD = Term.Term_Code " & _
            "LEFT JOIN 	TAXSCHED_SQL TS   WITH (Nolock) ON O.Tax_Sched = TS.Tax_Sched " & _
            "LEFT JOIN 	TAXDETL_SQL  Tax1 WITH (Nolock) ON O.Tax_Cd = Tax1.Tax_Cd " & _
            "LEFT JOIN 	TAXDETL_SQL  Tax2 WITH (Nolock) ON O.Tax_Cd_2 = Tax2.Tax_Cd " & _
            "LEFT JOIN 	TAXDETL_SQL  Tax3 WITH (Nolock) ON O.Tax_Cd_3 = Tax3.Tax_Cd " & _
            "LEFT JOIN 	ARSLMFIL_SQL Sls1 WITH (Nolock) ON O.SlsPsn_No = Sls1.HumRes_ID " & _
            "LEFT JOIN 	ARSLMFIL_SQL Sls2 WITH (Nolock) ON O.SlsPsn_No_2 = Sls2.HumRes_ID " & _
            "LEFT JOIN 	ARSLMFIL_SQL Sls3 WITH (Nolock) ON O.SlsPsn_No_3 = Sls3.HumRes_ID " & _
            "WHERE  Ord_No = '" & Ord_No & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                If Not dt.Rows(0).Item("Ord_Type").Equals(DBNull.Value) Then Ord_Type = dt.Rows(0).Item("Ord_Type") 'Char (8)
                If Not dt.Rows(0).Item("Ord_No").Equals(DBNull.Value) Then Ord_No = dt.Rows(0).Item("Ord_No") 'Char (8)       Order No
                If Not dt.Rows(0).Item("Line_Seq_No").Equals(DBNull.Value) Then Line_Seq_No = dt.Rows(0).Item("Line_Seq_No") 'Int            Line Sequence No
                If Not dt.Rows(0).Item("Item_No").Equals(DBNull.Value) Then Item_No = dt.Rows(0).Item("Item_No") 'Char (15)      Item No

                If Not dt.Rows(0).Item("Loc").Equals(DBNull.Value) Then Loc = dt.Rows(0).Item("Loc") 'Char (3)
                If Not dt.Rows(0).Item("Pick_Seq").Equals(DBNull.Value) Then Pick_Seq = dt.Rows(0).Item("Pick_Seq") 'Char (8)
                If Not dt.Rows(0).Item("Cus_Item_No").Equals(DBNull.Value) Then Cus_Item_No = dt.Rows(0).Item("Cus_Item_No") 'Char (30)
                If Not dt.Rows(0).Item("Item_Desc_1").Equals(DBNull.Value) Then Item_Desc_1 = dt.Rows(0).Item("Item_Desc_1") 'Char (30)      Item Description Line 1
                If Not dt.Rows(0).Item("Item_Desc_2").Equals(DBNull.Value) Then Item_Desc_2 = dt.Rows(0).Item("Item_Desc_2") 'Char (30)      Item Description Line 2
                If Not dt.Rows(0).Item("Qty_Ordered").Equals(DBNull.Value) Then Qty_Ordered = dt.Rows(0).Item("Qty_Ordered") 'Decimal (13,4) Qty Ordered
                If Not dt.Rows(0).Item("Qty_To_Ship").Equals(DBNull.Value) Then Qty_To_Ship = dt.Rows(0).Item("Qty_To_Ship") 'Decimal (13,4)
                If Not dt.Rows(0).Item("Unit_Price").Equals(DBNull.Value) Then Unit_Price = dt.Rows(0).Item("Unit_Price") 'Decimal (16,6) Unit Price  If not populated, default 'Price' from IM Loca
                If Not dt.Rows(0).Item("Discount_Pct").Equals(DBNull.Value) Then Discount_Pct = dt.Rows(0).Item("Discount_Pct") 'Numeric NULL
                If Not dt.Rows(0).Item("Request_Dt").Equals(DBNull.Value) Then Request_Dt = dt.Rows(0).Item("Request_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Qty_Bkord").Equals(DBNull.Value) Then Qty_Prev_Bkord = dt.Rows(0).Item("Qty_Bkord") 'Numeric NULL
                If Not dt.Rows(0).Item("Qty_Return_To_Stk").Equals(DBNull.Value) Then Qty_Return_To_Stk = dt.Rows(0).Item("Qty_Return_To_Stk") 'Numeric NULL
                If Not dt.Rows(0).Item("Bkord_Fg").Equals(DBNull.Value) Then Bkord_Fg = dt.Rows(0).Item("Bkord_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Uom").Equals(DBNull.Value) Then Uom = dt.Rows(0).Item("Uom") 'Char(2) NULL
                If Not dt.Rows(0).Item("Uom_Ratio").Equals(DBNull.Value) Then Uom_Ratio = dt.Rows(0).Item("Uom_Ratio") 'Numeric NULL
                If Not dt.Rows(0).Item("Unit_Cost").Equals(DBNull.Value) Then Unit_Cost = dt.Rows(0).Item("Unit_Cost") 'Numeric NULL
                If Not dt.Rows(0).Item("Unit_Weight").Equals(DBNull.Value) Then Unit_Weight = dt.Rows(0).Item("Unit_Weight") 'Numeric NULL
                If Not dt.Rows(0).Item("Comm_Calc_Type").Equals(DBNull.Value) Then Comm_Calc_Type = dt.Rows(0).Item("Comm_Calc_Type") 'Char(1) NULL
                If Not dt.Rows(0).Item("Comm_Pct_Or_Amt").Equals(DBNull.Value) Then Comm_Pct_Or_Amt = dt.Rows(0).Item("Comm_Pct_Or_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Promise_Dt").Equals(DBNull.Value) Then Promise_Dt = dt.Rows(0).Item("Promise_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Tax_Fg").Equals(DBNull.Value) Then Tax_Fg = dt.Rows(0).Item("Tax_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Stocked_Fg").Equals(DBNull.Value) Then Stocked_Fg = dt.Rows(0).Item("Stocked_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Controlled_Fg").Equals(DBNull.Value) Then Controlled_Fg = dt.Rows(0).Item("Controlled_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Select_Cd").Equals(DBNull.Value) Then Select_Cd = dt.Rows(0).Item("Select_Cd") 'Char(1) NOT NULL
                If Not dt.Rows(0).Item("Tot_Qty_Ordered").Equals(DBNull.Value) Then Tot_Qty_Ordered = dt.Rows(0).Item("Tot_Qty_Ordered") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Qty_Shipped").Equals(DBNull.Value) Then Tot_Qty_Shipped = dt.Rows(0).Item("Tot_Qty_Shipped") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Fg_1").Equals(DBNull.Value) Then Tax_Fg_1 = dt.Rows(0).Item("Tax_Fg_1") 'Char(1) NULL
                If Not dt.Rows(0).Item("Tax_Fg_2").Equals(DBNull.Value) Then Tax_Fg_2 = dt.Rows(0).Item("Tax_Fg_2") 'Char(1) NULL
                If Not dt.Rows(0).Item("Tax_Fg_3").Equals(DBNull.Value) Then Tax_Fg_3 = dt.Rows(0).Item("Tax_Fg_3") 'Char(1) NULL
                If Not dt.Rows(0).Item("Orig_Price").Equals(DBNull.Value) Then Orig_Price = dt.Rows(0).Item("Orig_Price") 'Numeric NULL
                If Not dt.Rows(0).Item("Copy_To_Bm_Fg").Equals(DBNull.Value) Then Copy_To_Bm_Fg = dt.Rows(0).Item("Copy_To_Bm_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Explode_Kit").Equals(DBNull.Value) Then Explode_Kit = dt.Rows(0).Item("Explode_Kit") 'Char(1) NULL
                If Not dt.Rows(0).Item("Mfg_Ord_No").Equals(DBNull.Value) Then Mfg_Ord_No = dt.Rows(0).Item("Mfg_Ord_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Allocate_Dt").Equals(DBNull.Value) Then Allocate_Dt = dt.Rows(0).Item("Allocate_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Last_Post_Dt").Equals(DBNull.Value) Then Last_Post_Dt = dt.Rows(0).Item("Last_Post_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Post_To_Inv_Qty").Equals(DBNull.Value) Then Post_To_Inv_Qty = dt.Rows(0).Item("Post_To_Inv_Qty") 'Numeric NULL
                If Not dt.Rows(0).Item("Posted_To_Inv").Equals(DBNull.Value) Then Posted_To_Inv = dt.Rows(0).Item("Posted_To_Inv") 'Numeric NULL
                If Not dt.Rows(0).Item("Tot_Qty_Posted").Equals(DBNull.Value) Then Tot_Qty_Posted = dt.Rows(0).Item("Tot_Qty_Posted") 'Numeric NULL
                If Not dt.Rows(0).Item("Qty_Allocated").Equals(DBNull.Value) Then Qty_Allocated = dt.Rows(0).Item("Qty_Allocated") 'Numeric NULL
                If Not dt.Rows(0).Item("Components_Alloc").Equals(DBNull.Value) Then Components_Alloc = dt.Rows(0).Item("Components_Alloc") 'Numeric NULL
                If Not dt.Rows(0).Item("Bin_Fg").Equals(DBNull.Value) Then Bin_Fg = dt.Rows(0).Item("Bin_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Cost_Meth").Equals(DBNull.Value) Then Cost_Meth = dt.Rows(0).Item("Cost_Meth") 'Char(1) NULL
                If Not dt.Rows(0).Item("Ser_Lot_Cd").Equals(DBNull.Value) Then Ser_Lot_Cd = dt.Rows(0).Item("Ser_Lot_Cd") 'Char(1) NULL
                If Not dt.Rows(0).Item("Mult_Ftr_Fg").Equals(DBNull.Value) Then Mult_Ftr_Fg = dt.Rows(0).Item("Mult_Ftr_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Line_Type").Equals(DBNull.Value) Then Line_Type = dt.Rows(0).Item("Line_Type") 'Char(1) NULL
                If Not dt.Rows(0).Item("Prod_Cat").Equals(DBNull.Value) Then Prod_Cat = dt.Rows(0).Item("Prod_Cat") 'Char(3) NULL
                If Not dt.Rows(0).Item("End_Item_Cd").Equals(DBNull.Value) Then End_Item_Cd = dt.Rows(0).Item("End_Item_Cd") 'Char(1) NULL
                If Not dt.Rows(0).Item("Reason_Cd").Equals(DBNull.Value) Then Reason_Cd = dt.Rows(0).Item("Reason_Cd") 'Char(3) NULL
                If Not dt.Rows(0).Item("Feature_Return").Equals(DBNull.Value) Then Feature_Return = dt.Rows(0).Item("Feature_Return") 'Char(1) NULL
                If Not dt.Rows(0).Item("Rec_Inspection").Equals(DBNull.Value) Then Rec_Inspection = dt.Rows(0).Item("Rec_Inspection") 'Char(1) NULL
                If Not dt.Rows(0).Item("Ship_From_Stk").Equals(DBNull.Value) Then Ship_From_Stk = dt.Rows(0).Item("Ship_From_Stk") 'Char(1) NULL
                If Not dt.Rows(0).Item("Mult_Release").Equals(DBNull.Value) Then Mult_Release = dt.Rows(0).Item("Mult_Release") 'Char(1) NULL
                If Not dt.Rows(0).Item("Req_Ship_Dt").Equals(DBNull.Value) Then Req_Ship_Dt = dt.Rows(0).Item("Req_Ship_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Qty_From_Stk").Equals(DBNull.Value) Then Qty_From_Stk = dt.Rows(0).Item("Qty_From_Stk") 'Numeric NULL
                If Not dt.Rows(0).Item("User_Def_Fld_1").Equals(DBNull.Value) Then User_Def_Fld_1 = dt.Rows(0).Item("User_Def_Fld_1") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_2").Equals(DBNull.Value) Then User_Def_Fld_2 = dt.Rows(0).Item("User_Def_Fld_2") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_3").Equals(DBNull.Value) Then User_Def_Fld_3 = dt.Rows(0).Item("User_Def_Fld_3") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_4").Equals(DBNull.Value) Then User_Def_Fld_4 = dt.Rows(0).Item("User_Def_Fld_4") 'Char(30) NULL
                If Not dt.Rows(0).Item("User_Def_Fld_5").Equals(DBNull.Value) Then User_Def_Fld_5 = dt.Rows(0).Item("User_Def_Fld_5") 'Char(30) NULL
                If Not dt.Rows(0).Item("Picked_Dt").Equals(DBNull.Value) Then Picked_Dt = dt.Rows(0).Item("Picked_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Shipped_Dt").Equals(DBNull.Value) Then Shipped_Dt = dt.Rows(0).Item("Shipped_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Billed_Dt").Equals(DBNull.Value) Then Billed_Dt = dt.Rows(0).Item("Billed_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Update_Fg").Equals(DBNull.Value) Then Update_Fg = dt.Rows(0).Item("Update_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Prc_Cd_Orig_Price").Equals(DBNull.Value) Then Prc_Cd_Orig_Price = dt.Rows(0).Item("Prc_Cd_Orig_Price") 'Numeric NULL
                If Not dt.Rows(0).Item("Tax_Sched").Equals(DBNull.Value) Then Tax_Sched = dt.Rows(0).Item("Tax_Sched") 'Char(5) NULL
                If Not dt.Rows(0).Item("Cus_No").Equals(DBNull.Value) Then Cus_No = dt.Rows(0).Item("Cus_No") 'Char(20) NULL
                If Not dt.Rows(0).Item("Tax_Amt").Equals(DBNull.Value) Then Tax_Amt = dt.Rows(0).Item("Tax_Amt") 'Numeric NULL
                If Not dt.Rows(0).Item("Qty_Bkord_Fg").Equals(DBNull.Value) Then Qty_Bkord_Fg = dt.Rows(0).Item("Qty_Bkord_Fg") 'Char(1) NOT NULL
                If Not dt.Rows(0).Item("Line_No").Equals(DBNull.Value) Then Line_No = dt.Rows(0).Item("Line_No") 'SmallInt NOT NULL
                If Not dt.Rows(0).Item("Mfg_Method").Equals(DBNull.Value) Then Mfg_Method = dt.Rows(0).Item("Mfg_Method") 'Char(2) NULL
                If Not dt.Rows(0).Item("Forced_Demand").Equals(DBNull.Value) Then Forced_Demand = dt.Rows(0).Item("Forced_Demand") 'Char(1) NULL
                If Not dt.Rows(0).Item("Conf_Pick_Dt").Equals(DBNull.Value) Then Conf_Pick_Dt = dt.Rows(0).Item("Conf_Pick_Dt") 'DateTime NULL
                If Not dt.Rows(0).Item("Item_Release_No").Equals(DBNull.Value) Then Item_Release_No = dt.Rows(0).Item("Item_Release_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Bin_Ser_Lot_Comp").Equals(DBNull.Value) Then Bin_Ser_Lot_Comp = dt.Rows(0).Item("Bin_Ser_Lot_Comp") 'Char(1) NULL
                If Not dt.Rows(0).Item("Offset_Used_Fg").Equals(DBNull.Value) Then Offset_Used_Fg = dt.Rows(0).Item("Offset_Used_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Ecs_Space").Equals(DBNull.Value) Then Ecs_Space = dt.Rows(0).Item("Ecs_Space") 'Char(20) NULL
                If Not dt.Rows(0).Item("Sfc_Order_Status").Equals(DBNull.Value) Then Sfc_Order_Status = dt.Rows(0).Item("Sfc_Order_Status") 'Char(1) NULL
                If Not dt.Rows(0).Item("Total_Cost").Equals(DBNull.Value) Then Total_Cost = dt.Rows(0).Item("Total_Cost") 'Numeric NULL
                If Not dt.Rows(0).Item("Po_Ord_No").Equals(DBNull.Value) Then Po_Ord_No = dt.Rows(0).Item("Po_Ord_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Rma_Seq").Equals(DBNull.Value) Then Rma_Seq = dt.Rows(0).Item("Rma_Seq") 'SmallInt NOT NULL
                If Not dt.Rows(0).Item("Vendor_No").Equals(DBNull.Value) Then Vendor_No = dt.Rows(0).Item("Vendor_No") 'Char(20) NULL
                If Not dt.Rows(0).Item("Posted_Unit_Cost").Equals(DBNull.Value) Then Posted_Unit_Cost = dt.Rows(0).Item("Posted_Unit_Cost") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_1").Equals(DBNull.Value) Then Extra_1 = dt.Rows(0).Item("Extra_1") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_2").Equals(DBNull.Value) Then Extra_2 = dt.Rows(0).Item("Extra_2") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_3").Equals(DBNull.Value) Then Extra_3 = dt.Rows(0).Item("Extra_3") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_4").Equals(DBNull.Value) Then Extra_4 = dt.Rows(0).Item("Extra_4") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_5").Equals(DBNull.Value) Then Extra_5 = dt.Rows(0).Item("Extra_5") 'Char(1) NULL
                If Not dt.Rows(0).Item("Extra_6").Equals(DBNull.Value) Then Extra_6 = dt.Rows(0).Item("Extra_6") 'Char(8) NULL
                If Not dt.Rows(0).Item("Extra_7").Equals(DBNull.Value) Then Extra_7 = dt.Rows(0).Item("Extra_7") 'Char(8) NULL
                If Not dt.Rows(0).Item("Extra_8").Equals(DBNull.Value) Then Extra_8 = dt.Rows(0).Item("Extra_8") 'Char(12) NULL
                If Not dt.Rows(0).Item("Extra_9").Equals(DBNull.Value) Then Extra_9 = dt.Rows(0).Item("Extra_9") 'Char(12) NULL
                If Not dt.Rows(0).Item("Extra_10").Equals(DBNull.Value) Then Extra_10 = dt.Rows(0).Item("Extra_10") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_11").Equals(DBNull.Value) Then Extra_11 = dt.Rows(0).Item("Extra_11") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_12").Equals(DBNull.Value) Then Extra_12 = dt.Rows(0).Item("Extra_12") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_13").Equals(DBNull.Value) Then Extra_13 = dt.Rows(0).Item("Extra_13") 'Numeric NULL
                If Not dt.Rows(0).Item("Extra_14").Equals(DBNull.Value) Then Extra_14 = dt.Rows(0).Item("Extra_14") 'Int NULL
                If Not dt.Rows(0).Item("Extra_15").Equals(DBNull.Value) Then Extra_15 = dt.Rows(0).Item("Extra_15") 'Int NULL
                If Not dt.Rows(0).Item("Warranty_Date").Equals(DBNull.Value) Then Warranty_Date = dt.Rows(0).Item("Warranty_Date") 'DateTime NULL
                If Not dt.Rows(0).Item("Revision_No").Equals(DBNull.Value) Then Revision_No = dt.Rows(0).Item("Revision_No") 'Char(8) NULL
                If Not dt.Rows(0).Item("Cm_Post_Fg").Equals(DBNull.Value) Then Cm_Post_Fg = dt.Rows(0).Item("Cm_Post_Fg") 'Char(1) NULL
                If Not dt.Rows(0).Item("Recalc_Sw").Equals(DBNull.Value) Then Recalc_Sw = dt.Rows(0).Item("Recalc_Sw") 'Char(1) NOT NULL
                If Not dt.Rows(0).Item("Filler_0004").Equals(DBNull.Value) Then Filler_0004 = dt.Rows(0).Item("Filler_0004") 'Char(132) NULL
                If Not dt.Rows(0).Item("Id").Equals(DBNull.Value) Then Id = dt.Rows(0).Item("Id") 'Numeric NOT NULL
                If Not dt.Rows(0).Item("Refill_Fg").Equals(DBNull.Value) Then Refill_Fg = dt.Rows(0).Item("Refill_Fg") 'Numeric NOT NULL
                'If Not dt.Rows(0).Item("ProductProof").Equals(DBNull.Value) Then ProductProof = dt.Rows(0).Item("ProductProof")'Char(2) NULL

            End If

        Catch er As Exception
            MsgBox("Error occured : " & er.Message & " in cOrdLine.Load")

        Finally
            blnLoading = False

        End Try

    End Sub


#Region "Properties ##################################################"

    Public Property Allocate_Dt() As Date
        Get
            Allocate_Dt = m_Allocate_Dt
        End Get
        Set(ByVal Value As Date)
            m_Allocate_Dt = Value
        End Set
    End Property
    Public Property AutoCompleteReship() As Integer
        Get
            AutoCompleteReship = m_AutoCompleteReship
        End Get
        Set(ByVal Value As Integer)
            m_AutoCompleteReship = Value
        End Set
    End Property
    Public Property AutoCompleteReship_Bln() As Boolean
        Get
            AutoCompleteReship_Bln = (m_AutoCompleteReship = 1)
        End Get
        Set(ByVal Value As Boolean)
            m_AutoCompleteReship = IIf(Value, 1, 0)
        End Set
    End Property
    Public Property Billed_Dt() As Date
        Get
            Billed_Dt = m_Billed_Dt
        End Get
        Set(ByVal Value As Date)
            m_Billed_Dt = Value
        End Set
    End Property
    Public Property Bin_Fg() As String
        Get
            Bin_Fg = m_Bin_Fg
        End Get
        Set(ByVal Value As String)
            m_Bin_Fg = Value
        End Set
    End Property
    Public Property Bin_Ser_Lot_Comp() As String
        Get
            Bin_Ser_Lot_Comp = m_Bin_Ser_Lot_Comp
        End Get
        Set(ByVal Value As String)
            m_Bin_Ser_Lot_Comp = Value
        End Set
    End Property
    Public Property Bkord_Fg() As String
        Get
            Bkord_Fg = m_Bkord_Fg
        End Get
        Set(ByVal Value As String)
            Try
                If Value.Length > 1 Then Throw New OEException(OEError.Entry_Too_Long, True)
                If m_Bkord_Fg <> Value Then m_Dirty = True
                m_Bkord_Fg = Value
            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set
    End Property
    Public Property BkOrd_Fg_Bln() As Boolean
        Set(ByVal value As Boolean)
            m_Bkord_Fg = IIf(value, "Y", "N")
        End Set
        Get
            BkOrd_Fg_Bln = (m_Bkord_Fg = "Y")
        End Get
    End Property
    Public Property Calc_Price() As Double
        Get
            Calc_Price = m_Calc_Price
        End Get
        Set(ByVal Value As Double)
            m_Calc_Price = Value
        End Set
    End Property
    Public Property Cm_Post_Fg() As String
        Get
            Cm_Post_Fg = m_Cm_Post_Fg
        End Get
        Set(ByVal Value As String)
            m_Cm_Post_Fg = Value
        End Set
    End Property
    Public Property Comm_Calc_Type() As String
        Get
            Comm_Calc_Type = m_Comm_Calc_Type
        End Get
        Set(ByVal Value As String)
            m_Comm_Calc_Type = Value
        End Set
    End Property
    Public Property Comm_Pct_Or_Amt() As Double
        Get
            Comm_Pct_Or_Amt = m_Comm_Pct_Or_Amt
        End Get
        Set(ByVal Value As Double)
            m_Comm_Pct_Or_Amt = Value
        End Set
    End Property
    Public Property Comp_Item_Guid() As String
        Get
            Comp_Item_Guid = m_Comp_Item_Guid
        End Get
        Set(ByVal Value As String)
            m_Comp_Item_Guid = Value
        End Set
    End Property
    Public Property Components_Alloc() As Double
        Get
            Components_Alloc = m_Components_Alloc
        End Get
        Set(ByVal Value As Double)
            m_Components_Alloc = Value
        End Set
    End Property
    Public Property Conf_Pick_Dt() As Date
        Get
            Conf_Pick_Dt = m_Conf_Pick_Dt
        End Get
        Set(ByVal Value As Date)
            m_Conf_Pick_Dt = Value
        End Set
    End Property
    Public Property Controlled_Fg() As String
        Get
            Controlled_Fg = m_Controlled_Fg
        End Get
        Set(ByVal Value As String)
            m_Controlled_Fg = Value
        End Set
    End Property
    Public Property Copy_To_Bm_Fg() As String
        Get
            Copy_To_Bm_Fg = m_Copy_To_Bm_Fg
        End Get
        Set(ByVal Value As String)
            m_Copy_To_Bm_Fg = Value
        End Set
    End Property
    Public Property Cost_Meth() As String
        Get
            Cost_Meth = m_Cost_Meth
        End Get
        Set(ByVal Value As String)
            m_Cost_Meth = Value
        End Set
    End Property
    Public Property Create_Imprint() As Boolean
        Get
            Create_Imprint = m_Create_Imprint
        End Get
        Set(ByVal Value As Boolean)
            m_Create_Imprint = Value
            If m_Create_Imprint Then m_Imprint = New cImprint(m_Item_Guid)
        End Set
    End Property
    Public Property Create_Traveler() As Boolean
        Get
            Create_Traveler = m_Create_Traveler
        End Get
        Set(ByVal Value As Boolean)
            m_Create_Traveler = Value
            If m_Create_Traveler Then m_Traveler = New CTraveler(m_Item_Guid)
        End Set
    End Property
    Public Property Create_Kit() As Boolean
        Get
            Create_Kit = m_Create_Kit
        End Get
        Set(ByVal Value As Boolean)
            m_Create_Kit = Value
            If m_Create_Kit Then m_Kit = New cKit(Me)
        End Set
    End Property
    Public Property Cus_Item_No() As String
        Get
            Cus_Item_No = m_Cus_Item_No
        End Get
        Set(ByVal Value As String)
            m_Cus_Item_No = Value
        End Set
    End Property
    Public Property Cus_No() As String
        Get
            Cus_No = m_Cus_No
        End Get
        Set(ByVal Value As String)
            m_Cus_No = Value
        End Set
    End Property
    Public Property Discount_Pct() As Double
        Get
            Discount_Pct = m_Discount_Pct
        End Get
        Set(ByVal Value As Double)

            Try
                If (m_Discount_Pct = Value) And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)

                If Value > 100 Then Throw New OEException(OEError.Disc_Pct_GT_100_Pct, True)
                If Value < 0 Then Throw New OEException(OEError.Disc_Pct_LT_0_Pct, True)
                m_Discount_Pct = Value
                Call Price_Changed()
                m_Dirty = True
            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try

        End Set
    End Property
    Public Property SaveToDB() As Boolean
        Get
            SaveToDB = m_SaveToDB
        End Get
        Set(ByVal value As Boolean)
            m_SaveToDB = value
        End Set
    End Property
    Public Property Ecs_Space() As String
        Get
            Ecs_Space = m_Ecs_Space
        End Get
        Set(ByVal Value As String)
            m_Ecs_Space = Value
        End Set
    End Property
    Public Property End_Item_Cd() As String
        Get
            End_Item_Cd = m_End_Item_Cd
        End Get
        Set(ByVal Value As String)
            m_End_Item_Cd = Value
        End Set
    End Property
    Public Property Explode_Kit() As String
        Get
            Explode_Kit = m_Explode_Kit
        End Get
        Set(ByVal Value As String)
            m_Explode_Kit = Value
        End Set
    End Property
    Public Property Extra_1() As String
        Get
            Extra_1 = m_Extra_1
        End Get
        Set(ByVal Value As String)
            m_Extra_1 = Value
        End Set
    End Property
    Public Property Extra_1_Bln() As Boolean
        Set(ByVal value As Boolean)
            m_Extra_1 = IIf(value, "Y", "N")
        End Set
        Get
            If m_Extra_1 Is Nothing Then
                Extra_1_Bln = False
            Else
                Extra_1_Bln = (m_Extra_1 = "Y")
            End If
        End Get
    End Property
    Public Property Extra_2() As String
        Get
            Extra_2 = m_Extra_2
        End Get
        Set(ByVal Value As String)
            m_Extra_2 = Value
        End Set
    End Property
    Public Property Extra_2_Bln() As Boolean
        Set(ByVal value As Boolean)
            m_Extra_2 = IIf(value, "Y", "")
        End Set
        Get
            If m_Extra_2 Is Nothing Then
                Extra_2_Bln = False
            Else
                Extra_2_Bln = (m_Extra_2 = "Y")
            End If
        End Get
    End Property
    Public Property Extra_3() As String
        Get
            Extra_3 = m_Extra_3
        End Get
        Set(ByVal Value As String)
            m_Extra_3 = Value
        End Set
    End Property
    Public Property Extra_4() As String
        Get
            Extra_4 = m_Extra_4
        End Get
        Set(ByVal Value As String)
            m_Extra_4 = Value
        End Set
    End Property
    Public Property Extra_5() As String
        Get
            Extra_5 = m_Extra_5
        End Get
        Set(ByVal Value As String)
            m_Extra_5 = Value
        End Set
    End Property
    Public Property Extra_6() As String
        Get
            Extra_6 = m_Extra_6
        End Get
        Set(ByVal Value As String)
            m_Extra_6 = Value
        End Set
    End Property
    Public Property Extra_7() As String
        Get
            Extra_7 = m_Extra_7
        End Get
        Set(ByVal Value As String)
            m_Extra_7 = Value
        End Set
    End Property
    Public Property Extra_8() As String
        Get
            Extra_8 = m_Extra_8
        End Get
        Set(ByVal Value As String)
            m_Extra_8 = Value
        End Set
    End Property
    Public Property Extra_9() As String
        Get
            Extra_9 = m_Extra_9
        End Get
        Set(ByVal Value As String)
            m_Extra_9 = Value
        End Set
    End Property
    Public Property Extra_10() As Double
        Get
            Extra_10 = m_Extra_10
        End Get
        Set(ByVal Value As Double)
            m_Extra_10 = Value
        End Set
    End Property
    Public Property Extra_11() As Double
        Get
            Extra_11 = m_Extra_11
        End Get
        Set(ByVal Value As Double)
            m_Extra_11 = Value
        End Set
    End Property
    Public Property Extra_12() As Double
        Get
            Extra_12 = m_Extra_12
        End Get
        Set(ByVal Value As Double)
            m_Extra_12 = Value
        End Set
    End Property
    Public Property Extra_13() As Double
        Get
            Extra_13 = m_Extra_13
        End Get
        Set(ByVal Value As Double)
            m_Extra_13 = Value
        End Set
    End Property
    Public Property Extra_14() As Integer
        Get
            Extra_14 = m_Extra_14
        End Get
        Set(ByVal Value As Integer)
            m_Extra_14 = Value
        End Set
    End Property
    Public Property Extra_15() As Integer
        Get
            Extra_15 = m_Extra_15
        End Get
        Set(ByVal Value As Integer)
            m_Extra_15 = Value
        End Set
    End Property
    Public Property Feature_Return() As String
        Get
            Feature_Return = m_Feature_Return
        End Get
        Set(ByVal Value As String)
            m_Feature_Return = Value
        End Set
    End Property
    Public Property Filler_0004() As String
        Get
            Filler_0004 = m_Filler_0004
        End Get
        Set(ByVal Value As String)
            m_Filler_0004 = Value
        End Set
    End Property
    Public Property FirstLoad() As Boolean
        Get
            FirstLoad = blnFirstLoad
        End Get
        Set(ByVal value As Boolean)
            blnFirstLoad = value
        End Set
    End Property
    Public Property Forced_Demand() As String
        Get
            Forced_Demand = m_Forced_Demand
        End Get
        Set(ByVal Value As String)
            m_Forced_Demand = Value
        End Set
    End Property
    Public Property Item_Guid() As String
        Get
            Item_Guid = m_Item_Guid
        End Get
        Set(ByVal Value As String)
            m_Item_Guid = Value
        End Set
    End Property
    Public Property Id() As Double
        Get
            Id = m_Id
        End Get
        Set(ByVal Value As Double)
            m_Id = Value
        End Set
    End Property
    Public Property Image() As Image
        Get
            Image = m_Image
        End Get
        Set(ByVal value As Image)
            m_Image = value
        End Set
    End Property
    Public Property ImprintLine() As cImprint
        Get
            ImprintLine = m_Imprint
        End Get
        Set(ByVal value As cImprint)
            m_Imprint = value
            If m_Item_No IsNot Nothing AndAlso m_Item_No <> "" AndAlso frmOrder.ignoreImprintMsgCheck = False Then
                Call Set_Spec_Instructions(m_Item_No, m_Imprint)

            End If
            'If Not (blnLoading Or blnFirstLoad) And m_SaveToDB And (m_Source <> OEI_SourceEnum.frmProductLineEntry) And (m_Source <> OEI_SourceEnum.Macola_ExactRepeat) Then Call Save()
        End Set
    End Property
    Public Property Item_Cd() As String
        Get
            Item_Cd = Mid(m_Item_Cd, 1, 10)
        End Get
        Set(ByVal Value As String)
            m_Item_Cd = Mid(Value, 1, 10)
        End Set
    End Property
    Public Property Mix_Group() As String
        Get
            Mix_Group = m_Mix_Group
        End Get
        Set(ByVal Value As String)
            m_Mix_Group = Value
        End Set
    End Property
    Public Property Item_Desc_1() As String
        Get
            Item_Desc_1 = m_Item_Desc_1
        End Get
        Set(ByVal Value As String)
            Try
                If Value.Length > 30 Then Throw New OEException(OEError.Entry_Too_Long, True)
                If m_Item_Desc_1 <> Value Then m_Dirty = True
                m_Item_Desc_1 = Value
            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set
    End Property
    Public Property Item_Desc_2() As String
        Get
            Item_Desc_2 = m_Item_Desc_2
        End Get
        Set(ByVal Value As String)
            Try
                If Value.Length > 30 Then Throw New OEException(OEError.Entry_Too_Long, True)
                If m_Item_Desc_2 <> Value Then m_Dirty = True
                m_Item_Desc_2 = Value
            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set
    End Property
    Public Property Item_No() As String
        Get
            Item_No = m_Item_No
        End Get
        Set(ByVal Value As String)
            Try
                If (Trim(m_Item_No) = Trim(Value)) And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)
                Value = Trim(Value.ToUpper)
                If Value.Length > 30 Then Throw New OEException(OEError.Entry_Too_Long, True)
                If Not m_oOrder.Validation.ItemIsActive(Value) And Not blnLoading Then
                    Throw New OEException(OEError.Item_Inactive, True)
                End If
                Call Item_No_Changed(Value)
                m_Dirty = True
                m_Item_No = Value
                Call Set_Item_Image()
                Call Set_Item_Kit(m_SaveToDB)
                Call Set_Spec_Instructions(m_Item_No, m_Imprint)
                '  Call MissingProp65(m_oOrder, m_oOrder)
                Call MissingProp65(m_Item_No, m_Item_Guid, m_oOrder, m_Kit_Component)
                If Not (blnLoading Or blnFirstLoad) And m_SaveToDB And (m_Source <> OEI_SourceEnum.frmProductLineEntry) And (m_Source <> OEI_SourceEnum.Macola_ExactRepeat) Then Call Save()
            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set
    End Property
    Public Property Item_Release_No() As String
        Get
            Item_Release_No = m_Item_Release_No
        End Get
        Set(ByVal Value As String)
            m_Item_Release_No = Value
        End Set
    End Property
    Public Property Kit() As cKit
        Get
            Kit = m_Kit
        End Get
        Set(ByVal value As cKit)
            m_Kit = value
        End Set
    End Property
    Public Property Kit_Component() As Boolean
        Get
            Kit_Component = m_Kit_Component
        End Get
        Set(ByVal value As Boolean)
            m_Kit_Component = value
        End Set
    End Property
    Public Property Kit_Feat_Fg_Bln() As Boolean
        Set(ByVal value As Boolean)
            m_Bkord_Fg = IIf(value, "Y", "N")
        End Set
        Get
            BkOrd_Fg_Bln = (m_Bkord_Fg = "Y")
        End Get
    End Property
    Public Property Kit_Item_Guid() As String
        Get
            Kit_Item_Guid = m_Kit_Item_Guid
        End Get
        Set(ByVal Value As String)
            m_Kit_Item_Guid = Value
        End Set
    End Property
    Public Property Last_Post_Dt() As Date
        Get
            Last_Post_Dt = m_Last_Post_Dt
        End Get
        Set(ByVal Value As Date)
            m_Last_Post_Dt = Value
        End Set
    End Property
    Public Property Line_No() As Short
        Get
            Line_No = m_Line_No
        End Get
        Set(ByVal Value As Short)
            m_Line_No = Value
        End Set
    End Property
    Public Property Line_Seq_No() As Integer
        Get
            Line_Seq_No = m_Line_Seq_No
        End Get
        Set(ByVal Value As Integer)
            m_Line_Seq_No = Value
        End Set
    End Property
    Public Property Line_Type() As String
        Get
            Line_Type = m_Line_Type
        End Get
        Set(ByVal Value As String)
            m_Line_Type = Value
        End Set
    End Property
    'Public Property No_Kits() As Boolean
    '    Get
    '        No_Kits = m_No_Kits
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        m_No_Kits = Value
    '    End Set
    'End Property
    Public Property Loc() As String
        Get
            Loc = m_Loc
        End Get
        Set(ByVal Value As String)
            Try
                If Trim(Value) = Trim(m_Loc) And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)
                If Value.Length > 3 Then Throw New OEException(OEError.Entry_Too_Long, True)

                Call Loc_Changed(Value)
                If Value = "" Then
                    m_Loc = m_oOrder.Ordhead.Mfg_Loc
                Else
                    m_Loc = Value
                End If
                m_Dirty = True
                If Not (blnLoading Or blnFirstLoad) And m_SaveToDB Then Call Save()

                'If (m_Qty_On_Hand <= 0 Or m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST") Then
                If (m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST") Then
                    Throw New OEException(OEError.Low_Stock_For_All_Star_Customer, False, True)
                End If

            Catch oe_er As OEException
                If oe_er.Cancel Or oe_er.ShowMessage Then Throw New OEException(oe_er.Message, oe_er)
            End Try

        End Set
    End Property
    Public Property Mfg_Method() As String
        Get
            Mfg_Method = m_Mfg_Method
        End Get
        Set(ByVal Value As String)
            m_Mfg_Method = Value
        End Set
    End Property
    Public Property Mfg_Ord_No() As String
        Get
            Mfg_Ord_No = m_Mfg_Ord_No
        End Get
        Set(ByVal Value As String)
            m_Mfg_Ord_No = Value
        End Set
    End Property
    Public Property Mult_Ftr_Fg() As String
        Get
            Mult_Ftr_Fg = m_Mult_Ftr_Fg
        End Get
        Set(ByVal Value As String)
            m_Mult_Ftr_Fg = Value
        End Set
    End Property
    Public Property Mult_Release() As String
        Get
            Mult_Release = m_Mult_Release
        End Get
        Set(ByVal Value As String)
            m_Mult_Release = Value
        End Set
    End Property
    Public Property OEOrdLin_ID() As Integer
        Get
            OEOrdLin_ID = m_OEOrdLin_Id
        End Get
        Set(ByVal value As Integer)
            m_OEOrdLin_Id = value
        End Set
    End Property
    Public Property Offset_Used_Fg() As String
        Get
            Offset_Used_Fg = m_Offset_Used_Fg
        End Get
        Set(ByVal Value As String)
            m_Offset_Used_Fg = Value
        End Set
    End Property
    Public Property Ord_Guid() As String
        Get
            Ord_Guid = m_Ord_Guid
        End Get
        Set(ByVal Value As String)
            m_Ord_Guid = Value
        End Set
    End Property
    Public Property Ord_No() As String
        Get
            Ord_No = m_Ord_No
        End Get
        Set(ByVal Value As String)
            m_Ord_No = Value
        End Set
    End Property
    Public Property Ord_Type() As String
        Get
            Ord_Type = m_Ord_Type
        End Get
        Set(ByVal Value As String)
            m_Ord_Type = Value
        End Set
    End Property
    Public Property Orig_Price() As Double
        Get
            Orig_Price = m_Orig_Price
        End Get
        Set(ByVal Value As Double)
            m_Orig_Price = Value
        End Set
    End Property
    Public Property Pick_Seq() As String
        Get
            Pick_Seq = m_Pick_Seq
        End Get
        Set(ByVal Value As String)
            m_Pick_Seq = Value
        End Set
    End Property
    Public Property Picked_Dt() As Date
        Get
            Picked_Dt = m_Picked_Dt
        End Get
        Set(ByVal Value As Date)
            m_Picked_Dt = Value
        End Set
    End Property
    Public Property Po_Ord_No() As String
        Get
            Po_Ord_No = m_Po_Ord_No
        End Get
        Set(ByVal Value As String)
            m_Po_Ord_No = Value
        End Set
    End Property
    Public Property Post_To_Inv_Qty() As Double
        Get
            Post_To_Inv_Qty = m_Post_To_Inv_Qty
        End Get
        Set(ByVal Value As Double)
            m_Post_To_Inv_Qty = Value
        End Set
    End Property
    Public Property Posted_To_Inv() As Double
        Get
            Posted_To_Inv = m_Posted_To_Inv
        End Get
        Set(ByVal Value As Double)
            m_Posted_To_Inv = Value
        End Set
    End Property
    Public Property Posted_Unit_Cost() As Double
        Get
            Posted_Unit_Cost = m_Posted_Unit_Cost
        End Get
        Set(ByVal Value As Double)
            m_Posted_Unit_Cost = Value
        End Set
    End Property
    Public Property Prc_Cd_Orig_Price() As Double
        Get
            Prc_Cd_Orig_Price = m_Prc_Cd_Orig_Price
        End Get
        Set(ByVal Value As Double)
            m_Prc_Cd_Orig_Price = Value
        End Set
    End Property
    Public Property Prod_Cat() As String
        Get
            Prod_Cat = m_Prod_Cat
        End Get
        Set(ByVal Value As String)
            m_Prod_Cat = Value
        End Set
    End Property
    Public Property ProductProof() As Integer
        Get
            ProductProof = m_ProductProof
        End Get
        Set(ByVal Value As Integer)
            m_ProductProof = Value
        End Set
    End Property
    Public Property ProductProof_Bln() As Boolean
        Get
            ProductProof_Bln = (m_ProductProof = 1)
        End Get
        Set(ByVal Value As Boolean)
            m_ProductProof = IIf(Value, 1, 0)
        End Set
    End Property
    Public Property Promise_Dt() As Date
        Get
            Promise_Dt = m_Promise_Dt
        End Get
        Set(ByVal Value As Date)
            Try
                If m_Promise_Dt.Equals(Value) And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)

                'If Value.Year < (Now.Year - 1) And Not blnLoading Then
                '    Throw New OEException(OEError.Year_Cannot_Be_LT_Previous_Year, True)
                'End If

                'If Value.Year > (Now.Year + 1) And Not blnLoading Then
                '    Throw New OEException(OEError.Year_Cannot_Be_GT_Follow_Year, True)
                'End If

                m_Promise_Dt = Value
                Call Promise_Dt_Changed()

                m_Dirty = True

                If Value.Date < Date.Now.Date And Not blnLoading Then
                    Throw New OEException(OEError.Promise_Date_Before_System_Date)
                End If

            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try

        End Set
    End Property
    Public Property Qty_Allocated() As Double
        Get
            Qty_Allocated = m_Qty_Allocated
        End Get
        Set(ByVal Value As Double)
            m_Qty_Allocated = Value
        End Set
    End Property
    Public Property Qty_Prev_Bkord() As Double
        Get
            Qty_Prev_Bkord = m_Qty_Prev_Bkord
        End Get
        Set(ByVal Value As Double)
            m_Qty_Prev_Bkord = Value
        End Set
    End Property
    Public Property Qty_Bkord_Fg() As String
        Get
            Qty_Bkord_Fg = m_Qty_Prev_Bkord_Fg
        End Get
        Set(ByVal Value As String)
            m_Qty_Prev_Bkord_Fg = Value
        End Set
    End Property
    Public Property Qty_From_Stk() As Double
        Get
            Qty_From_Stk = m_Qty_From_Stk
        End Get
        Set(ByVal Value As Double)
            m_Qty_From_Stk = Value
        End Set
    End Property
    Public Property Qty_Inventory() As Double
        Get
            Qty_Inventory = m_Qty_Inventory
        End Get
        Set(ByVal Value As Double)
            m_Qty_Inventory = Value
        End Set
    End Property
    Public Property Qty_Lines() As Double
        Get
            Qty_Lines = m_Qty_Lines
        End Get
        Set(ByVal Value As Double)
            Try
                If m_Qty_Lines = Value And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)

                If Value < 0 And Not blnLoading Then Throw New OEException(OEError.Negative_Qty_Not_Allowed, True)

                m_Qty_Lines = Value
                If m_Qty_Lines > 1 Then
                    Call Multiline_Qty_To_Ship_Changed(m_Item_No, m_Loc, False)
                Else
                    Call Qty_To_Ship_Changed(m_Item_No, m_Loc, False)
                End If

                'Call Price_Changed()

                'If Not (blnLoading Or blnFirstLoad) And m_SaveToDB Then Call Save()

                m_Dirty = True

                'If (m_Qty_On_Hand <= 0 Or m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST") Then
                If Not blnLoading And ((m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST")) Then
                    Throw New OEException(OEError.Low_Stock_For_All_Star_Customer, False, True)
                End If

            Catch oe_er As OEException
                If oe_er.Cancel Or oe_er.ShowMessage Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set

    End Property
    Public Property Qty_On_Hand() As Double
        Get
            Qty_On_Hand = m_Qty_On_Hand
        End Get
        Set(ByVal Value As Double)
            m_Qty_On_Hand = Value
        End Set
    End Property
    Public Property Qty_Ordered() As Double
        Get
            Qty_Ordered = m_Qty_Ordered
        End Get
        Set(ByVal Value As Double)
            Try
                If m_Qty_Ordered = Value And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)
                If Value = 0 And Not blnLoading Then Throw New OEException(OEError.Qty_Ordered_Cannot_Be_Zero, True)

                Call Qty_Ordered_Changed(Value) ' Put before to compare with current value
                m_Qty_Ordered = Value
                m_Qty_To_Ship = Value


                'project closed 88REWARDS
                '-------------------------------
                'Dim dtT As DataTable = Nothing

                'dtT = TenPercent(m_oOrder.Ordhead.Cus_No, Item_No, m_Qty_Ordered, m_oOrder.Ordhead.Cus_Type_Cd)

                '    If dtT.Rows.Count <> 0 Then

                '    If dtT.Rows(0).Item(enumTenPercent.eFinnalPriceRevised).ToString <> -1 Then

                '        m_Discount_Pct = dtT.Rows(0).Item(enumTenPercent.eTenPercent).ToString
                '    Else
                '        m_Discount_Pct = 0.00
                '    End If

                'End If
                '-------------------------------

                'Call Qty_To_Ship_Changed(m_Item_No, m_Loc, False)
                If m_Qty_Lines > 1 Then
                    Call Multiline_Qty_To_Ship_Changed(m_Item_No, m_Loc, False)
                Else
                    m_Qty_Lines = 1
                    Call Qty_To_Ship_Changed(m_Item_No, m_Loc, False)
                End If

                ' On qty ordered, we force qty to ship to be prev to ship. they can overwrite after.
                m_Qty_To_Ship = m_Qty_Prev_To_Ship

                'Qty_To_Ship_Calc = Value
                'If m_Qty_To_Ship <> Value Then Qty_To_Ship = Value
                Call Price_Changed()
                'm_Calc_Price = m_Qty_To_Ship * m_Unit_Price
                If Not (blnLoading Or blnFirstLoad) And m_SaveToDB Then Call Save()
                m_Dirty = True

                'If (m_Qty_On_Hand <= 0 Or m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST") Then
                If Not blnLoading And ((m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST")) Then
                    '++ID changed value inserted in constructor from False to True because False close all procedure for load info in time when you change the qty and 
                    '  Throw New OEException(OEError.Low_Stock_For_All_Star_Customer, False, True)
                    Throw New OEException(OEError.Low_Stock_For_All_Star_Customer, True, True)
                End If

            Catch oe_er As OEException
                If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
            End Try
        End Set
    End Property
    Public Property Qty_Ordered_Allows_Zero() As Double
        Get
            Qty_Ordered = m_Qty_Ordered
        End Get
        Set(ByVal Value As Double)
            Try
                If m_Qty_Ordered = Value And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)

                Call Qty_Ordered_Changed(Value) ' Put before to compare with current value
                m_Qty_Ordered = Value
                m_Qty_To_Ship = m_Qty_Ordered

                'Call Qty_To_Ship_Changed(m_Item_No, m_Loc, False)
                If m_Qty_Lines > 1 Then
                    Call Multiline_Qty_To_Ship_Changed(m_Item_No, m_Loc, False)
                    ' On qty ordered, we force qty to ship to be prev to ship. they can overwrite after.
                    m_Qty_To_Ship = m_Qty_Prev_To_Ship / m_Qty_Lines
                Else
                    m_Qty_Lines = 1
                    Call Qty_To_Ship_Changed(m_Item_No, m_Loc, False)
                    ' On qty ordered, we force qty to ship to be prev to ship. they can overwrite after.
                    m_Qty_To_Ship = m_Qty_Prev_To_Ship
                End If

                '' On qty ordered, we force qty to ship to be prev to ship. they can overwrite after.
                'm_Qty_To_Ship = m_Qty_Prev_To_Ship

                'Qty_To_Ship_Calc = Value
                'If m_Qty_To_Ship > Value Then Qty_To_Ship = Value
                Call Price_Changed()
                'm_Calc_Price = m_Qty_To_Ship * m_Unit_Price
                'If Not blnLoading Then Call Save()
                If Not (blnLoading Or blnFirstLoad) And m_SaveToDB Then Call Save()
                m_Dirty = True

                'If (m_Qty_On_Hand <= 0 Or m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST") Then
                If Not blnLoading And ((m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST")) Then

                    Throw New OEException(OEError.Low_Stock_For_All_Star_Customer, False, True)
                End If

            Catch oe_er As OEException
                If oe_er.Cancel Or oe_er.ShowMessage Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set
    End Property
    'Public Property Qty_Ordered_Load() As Double
    '    Get
    '        Qty_Ordered = m_Qty_Ordered
    '    End Get
    '    Set(ByVal Value As Double)
    '        Try
    '            If m_Qty_Ordered = Value And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)
    '            If Value = 0 And Not blnLoading Then Throw New OEException(OEError.Qty_Ordered_Cannot_Be_Zero, True)

    '            Call Qty_Ordered_Changed(Value) ' Put before to compare with current value
    '            m_Qty_Ordered = Value
    '            Call Qty_To_Ship_Changed(m_Item_No, m_Loc)
    '            'Qty_To_Ship_Calc = Value
    '            If m_Qty_To_Ship > Value Then Qty_To_Ship = Value
    '            Call Price_Changed()
    '            'm_Calc_Price = m_Qty_To_Ship * m_Unit_Price
    '            'If Not blnLoading Then Call Save()
    '            If Not (blnLoading Or blnFirstLoad) And m_SaveToDB Then Call Save()
    '            m_Dirty = True
    '        Catch oe_er As OEException
    '            If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
    '        End Try
    '    End Set
    'End Property
    Public Property Qty_Overpick() As Double
        Get
            Qty_Overpick = m_Qty_Overpick
        End Get
        Set(ByVal value As Double)
            m_Qty_Overpick = value
        End Set
    End Property
    Public Property Qty_Scrap() As Double
        Get
            Qty_Scrap = m_Qty_Scrap
        End Get
        Set(ByVal value As Double)
            m_Qty_Scrap = value
        End Set
    End Property
    Public Property Qty_To_Ship() As Double
        Get
            Qty_To_Ship = m_Qty_To_Ship
        End Get
        Set(ByVal Value As Double)
            Try
                If m_Qty_To_Ship = Value And ((Not blnLoading) Or (blnFirstLoad)) Then 'Throw New OEException(OEError.Exit_Property, False, False)
                    Exit Property
                End If
                If Value <0 And Not blnLoading Then Throw New OEException(OEError.Negative_Qty_Not_Allowed, True)

                m_Qty_To_Ship = Value

                'Call Qty_To_Ship_Changed(m_Item_No, m_Loc, True)
                        If m_Qty_Lines > 1 Then
                    Call Multiline_Qty_To_Ship_Changed(m_Item_No, m_Loc, True)
                Else
                    Call Qty_To_Ship_Changed(m_Item_No, m_Loc, True)
                End If

                'Call Price_Changed()

                If Not (blnLoading Or blnFirstLoad) And m_SaveToDB Then Call Save()

                m_Dirty = True

                'If (m_Qty_On_Hand <= 0 Or m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST") Then
                If Not blnLoading And ((m_Qty_Prev_Bkord > 0) And (m_oOrder.Ordhead.Customer.ClassificationId = "6ST" Or m_oOrder.Ordhead.Customer.ClassificationId = "7ST")) Then
                    Throw New OEException(OEError.Low_Stock_For_All_Star_Customer, False, True)
                End If

            Catch oe_er As OEException
                If oe_er.Cancel Or oe_er.ShowMessage Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set

    End Property
    Public Property Qty_Return_To_Stk() As Double
        Get
            Qty_Return_To_Stk = m_Qty_Return_To_Stk
        End Get
        Set(ByVal Value As Double)
            m_Qty_Return_To_Stk = Value
        End Set
    End Property
    Public Property Qty_Ship_Prev() As Double
        Get
            Qty_Ship_Prev = m_Qty_Prev_To_Ship
        End Get
        Set(ByVal Value As Double)
            m_Qty_Prev_To_Ship = Value
        End Set
    End Property
    Public Property Qty_Prev_To_Ship() As Double
        Get
            Qty_Prev_To_Ship = m_Qty_Prev_To_Ship
        End Get
        Set(ByVal Value As Double)
            'Call Qty_To_Ship_Calc_Changed(Value) ' Put before to compare with current value
            m_Qty_Prev_To_Ship = Value
            'Call Price_Changed()
            'Call Save()
        End Set
    End Property
    Public Property Reason_Cd() As String
        Get
            Reason_Cd = m_Reason_Cd
        End Get
        Set(ByVal Value As String)
            m_Reason_Cd = Value
        End Set
    End Property
    Public Property Rec_Inspection() As String
        Get
            Rec_Inspection = m_Rec_Inspection
        End Get
        Set(ByVal Value As String)
            m_Rec_Inspection = Value
        End Set
    End Property
    Public Property Recalc_Sw() As String
        Get
            Recalc_Sw = m_Recalc_Sw
        End Get
        Set(ByVal Value As String)
            m_Recalc_Sw = Value
        End Set
    End Property
    Public Property Recalc_Sw_Bln() As Boolean
        Set(ByVal value As Boolean)
            m_Recalc_Sw = IIf(value, "Y", "N")
        End Set
        Get
            Recalc_Sw_Bln = (m_Recalc_Sw = "Y")
        End Get
    End Property
    Public Property Refill_Fg() As Boolean
        Set(ByVal value As Boolean)
            m_Refill_Fg = value
        End Set
        Get
            Refill_Fg = m_Refill_Fg
        End Get
    End Property
    Public Property Is_Refill_Fg() As Boolean
        Set(ByVal value As Boolean)
            m_Is_Refill_Fg = value
        End Set
        Get
            Is_Refill_Fg = m_Is_Refill_Fg
        End Get
    End Property
    Public Property Has_Refills_Fg() As Boolean
        Set(ByVal value As Boolean)
            m_Has_Refills_Fg = value
        End Set
        Get
            Has_Refills_Fg = m_Has_Refills_Fg
        End Get
    End Property
    Public Property Req_Ship_Dt() As Date
        Get
            Req_Ship_Dt = m_Req_Ship_Dt
        End Get
        Set(ByVal Value As Date)
            Try
                If m_Req_Ship_Dt.Equals(Value) And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)

                'If Value.Year < (Now.Year - 1) And Not blnLoading Then
                '    Throw New OEException(OEError.Year_Cannot_Be_LT_Previous_Year, True)
                'End If

                'If Value.Year > (Now.Year + 1) And Not blnLoading Then
                '    Throw New OEException(OEError.Year_Cannot_Be_GT_Follow_Year, True)
                'End If

                m_Req_Ship_Dt = Value

                If m_Req_Ship_Dt.Date > m_Promise_Dt.Date And Not blnLoading Then
                    Throw New OEException(OEError.Req_Ship_Date_GT_Promise_Date)
                End If
                m_Dirty = True
            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set
    End Property
    Public Property Request_Dt() As Date
        Get
            Request_Dt = m_Request_Dt
        End Get
        Set(ByVal Value As Date)
            Try
                If m_Request_Dt.Date.Equals(Value) And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)

                'If Value.Year < (Now.Year - 1) And Not blnLoading Then
                '    Throw New OEException(OEError.Year_Cannot_Be_LT_Previous_Year, True)
                'End If

                'If Value.Year > (Now.Year + 1) And Not blnLoading Then
                '    Throw New OEException(OEError.Year_Cannot_Be_GT_Follow_Year, True)
                'End If

                m_Request_Dt = Value
                Call Request_Dt_Changed()

                If m_Request_Dt.Date < Date.Now.Date And Not blnLoading Then
                    Throw New OEException(OEError.Req_Date_Before_System_Date)
                End If
                m_Dirty = True
            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set
    End Property
    Public Property Revision_No() As String
        Get
            Revision_No = m_Revision_No
        End Get
        Set(ByVal Value As String)
            Try
                If Value.Length > 8 Then Throw New OEException(OEError.Entry_Too_Long, True)
                If m_Revision_No <> Value Then m_Dirty = True
                m_Revision_No = Value
            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try
        End Set
    End Property
    Public Property Rma_Seq() As Short
        Get
            Rma_Seq = m_Rma_Seq
        End Get
        Set(ByVal Value As Short)
            m_Rma_Seq = Value
        End Set
    End Property
    Public Property Route() As String
        Get
            Route = m_Route
        End Get
        Set(ByVal Value As String)
            ' m_RouteID = CInt(m_oOrder.RouteCollection(Value))
            m_Route = Value
            ' m_Dirty = True
        End Set
    End Property
    Public Sub SetRoute(ByVal pstrRoute As String)
        m_Route = pstrRoute
        If RouteIDcollection.Contains(CStr(Trim(pstrRoute))) Then
            m_RouteID = RouteIDcollection(CStr(Trim(pstrRoute)))
        Else
            m_RouteID = 0
        End If
        m_Dirty = True
    End Sub
    Public Property RouteID() As Integer
        Get
            RouteID = m_RouteID
        End Get
        Set(ByVal Value As Integer)
            m_RouteID = Value
            If Value = 0 Then
                m_Route = ""
            Else
                m_Route = RouteCollection(CStr(Trim(Value)))

                Try
                    Call Set_Spec_Instructions(String.Empty, m_RouteID, m_Imprint)

                Catch oe_er As OEException
                    If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
                End Try

            End If
            m_Dirty = True

            '**********************************************************************************************
            'REMOVED SEPT 11, 2017 - Make Bags the same
            'Added July 17, 2017 - T. Louzon - So prices is re-checked for Bags when route Changes
            'RouteExtension = f_GetRouteExtensionAndData(Value)
            'Mfg_Item_No = m_Item_No & "-" & RouteExtension

            'Call SetUnit_Price(m_Qty_Ordered)
            '**********************************************************************************************


        End Set
    End Property
    Public Property Select_Cd() As String
        Get
            Select_Cd = m_Select_Cd
        End Get
        Set(ByVal Value As String)
            m_Select_Cd = Value
        End Set
    End Property
    Public Property Ser_Lot_Cd() As String
        Get
            Ser_Lot_Cd = m_Ser_Lot_Cd
        End Get
        Set(ByVal Value As String)
            m_Ser_Lot_Cd = Value
        End Set
    End Property
    Public Property Sfc_Order_Status() As String
        Get
            Sfc_Order_Status = m_Sfc_Order_Status
        End Get
        Set(ByVal Value As String)
            m_Sfc_Order_Status = Value
        End Set
    End Property
    Public Property Ship_From_Stk() As String
        Get
            Ship_From_Stk = m_Ship_From_Stk
        End Get
        Set(ByVal Value As String)
            m_Ship_From_Stk = Value
        End Set
    End Property
    Public Property Shipped_Dt() As Date
        Get
            Shipped_Dt = m_Shipped_Dt
        End Get
        Set(ByVal Value As Date)
            m_Shipped_Dt = Value
        End Set
    End Property
    Public Property Stocked_Fg() As String
        Get
            Stocked_Fg = m_Stocked_Fg
        End Get
        Set(ByVal Value As String)
            m_Stocked_Fg = Value
        End Set
    End Property
    Public Property Source() As OEI_SourceEnum
        Get
            Source = m_Source
        End Get
        Set(ByVal value As OEI_SourceEnum)
            m_Source = value
        End Set
    End Property
    Public Property Tax_Amt() As Double
        Get
            Tax_Amt = m_Tax_Amt
        End Get
        Set(ByVal Value As Double)
            m_Tax_Amt = Value
        End Set
    End Property
    Public Property Tax_Fg() As String
        Get
            Tax_Fg = m_Tax_Fg
        End Get
        Set(ByVal Value As String)
            m_Tax_Fg = Value
        End Set
    End Property
    Public Property Tax_Fg_1() As String
        Get
            Tax_Fg_1 = m_Tax_Fg_1
        End Get
        Set(ByVal Value As String)
            m_Tax_Fg_1 = Value
        End Set
    End Property
    Public Property Tax_Fg_2() As String
        Get
            Tax_Fg_2 = m_Tax_Fg_2
        End Get
        Set(ByVal Value As String)
            m_Tax_Fg_2 = Value
        End Set
    End Property
    Public Property Tax_Fg_3() As String
        Get
            Tax_Fg_3 = m_Tax_Fg_3
        End Get
        Set(ByVal Value As String)
            m_Tax_Fg_3 = Value
        End Set
    End Property
    Public Property Tax_Sched() As String
        Get
            Tax_Sched = m_Tax_Sched
        End Get
        Set(ByVal Value As String)
            m_Tax_Sched = Value
        End Set
    End Property
    Public Property Tot_Qty_Ordered() As Double
        Get
            Tot_Qty_Ordered = m_Tot_Qty_Ordered
        End Get
        Set(ByVal Value As Double)
            m_Tot_Qty_Ordered = Value
        End Set
    End Property
    Public Property Tot_Qty_Posted() As Double
        Get
            Tot_Qty_Posted = m_Tot_Qty_Posted
        End Get
        Set(ByVal Value As Double)
            m_Tot_Qty_Posted = Value
        End Set
    End Property
    Public Property Tot_Qty_Shipped() As Double
        Get
            Tot_Qty_Shipped = m_Tot_Qty_Shipped
        End Get
        Set(ByVal Value As Double)
            m_Tot_Qty_Shipped = Value
        End Set
    End Property
    Public Property Total_Cost() As Double
        Get
            Total_Cost = m_Total_Cost
        End Get
        Set(ByVal Value As Double)
            m_Total_Cost = Value
        End Set
    End Property
    Public Property Traveler() As CTraveler
        Get
            Traveler = m_Traveler
        End Get
        Set(ByVal value As CTraveler)
            m_Traveler = value
        End Set
    End Property
    Public Property Unit_Cost() As Double
        Get
            Unit_Cost = m_Unit_Cost
        End Get
        Set(ByVal Value As Double)
            m_Unit_Cost = Value
        End Set
    End Property

    '======================================================================
    'ADDED May 6, 2017 - T. Louzon
    '======================================================================
    Public Property Unit_Price_BeforeSpecial() As Double
        Get
            Unit_Price_BeforeSpecial = m_Unit_Price_BeforeSpecial
        End Get
        Set(ByVal Value As Double)
            m_Unit_Price_BeforeSpecial = Value
        End Set
    End Property

    '======================================================================
    'END -  May 6, 2017 - T. Louzon
    '======================================================================
    '======================================================================
    'ADDED June 27, 2017 - T. Louzon
    '======================================================================
    Public Property Mfg_Item_No() As String
        Get
            Mfg_Item_No = m_Mfg_Item_No
        End Get
        Set(ByVal Value As String)
            m_Mfg_Item_No = Value
        End Set
    End Property
    Public Property RouteExtension() As String
        Get
            RouteExtension = m_RouteExtension
        End Get
        Set(ByVal Value As String)
            m_RouteExtension = Value
        End Set
    End Property

    '======================================================================
    'END -  June 27, 2017 - T. Louzon
    '======================================================================

    '======================================================================
    'ADDED June 29, 2017 - T. Louzon
    '======================================================================
    Public Property IsNoArtwork() As Integer
        Get
            IsNoArtwork = m_IsNoArtwork
        End Get
        Set(ByVal Value As Integer)
            m_IsNoArtwork = Value
        End Set
    End Property
    Public Property IsNoStock() As Integer
        Get
            IsNoStock = m_IsNoStock
        End Get
        Set(ByVal Value As Integer)
            m_IsNoStock = Value
        End Set
    End Property

    '======================================================================
    'END -  June 29, 2017 - T. Louzon
    '======================================================================


    'Created June 27, 2017 - T. Louzon
    'Modified June 29, 2017 - T. Louzon
    Public Function f_GetRouteExtensionAndData(RouteIDFromTraveler As String) As String

        f_GetRouteExtensionAndData = ""
        Try
            'GET THE Item Extension for the Route ID
            Dim dt2 As DataTable
            Dim db2 As New cDBA
            Dim strSql2 As String

            Dim strGetRouteExtensionAndData As String = ""


            'Added Isnull Sept 13, 2017
            strSql2 = " SELECT ItemExtension = isnull(T.ItemExtension,'XX'), T.RouteCategory, T.RouteDescription "
            strSql2 = strSql2 & " FROM Exact_traveler_Route T "
            'strSql2 = strSql2 & " INNER JOIN HV_OEI_RouteCategory_XREF XREF "    'Removed July 17, 2017 - T. Louzon - Not using this table any more.
            'strSql2 = strSql2 & " On T.RouteCategory = XREF.RouteCategory "
            strSql2 = strSql2 & " WHERE routeID = " & RouteIDFromTraveler

            dt2 = db2.DataTable(strSql2)
            If dt2.Rows.Count <> 0 Then

                ' f_GetRouteExtensionAndData = dt2.Rows(0).Item("ItemExtension").ToString

                strGetRouteExtensionAndData = dt2.Rows(0).Item("ItemExtension").ToString

                'GET IsNoArtwork and IsNoExtension using RouteCategory
                If InStr(UCase(dt2.Rows(0).Item("RouteDescription")), "NO ARTWORK") Then
                    IsNoArtwork = 1
                Else
                    IsNoArtwork = 0
                End If

                'GET IsNoArtwork and IsNoExtension using RouteCategory
                If InStr(UCase(dt2.Rows(0).Item("RouteDescription")), "NO STOCK") Or InStr(UCase(dt2.Rows(0).Item("RouteDescription")), "NOSTK") Then
                    IsNoStock = 1
                Else
                    IsNoStock = 0
                End If
                'MsgBox("X" + dt2.Rows(0).Item("RouteDescription"))

            Else
                ' f_GetRouteExtensionAndData = "XX"
                strGetRouteExtensionAndData = "XX"

                IsNoArtwork = 0
                IsNoStock = 0
            End If
            dt2.Dispose()


            Return strGetRouteExtensionAndData

        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Function

    Public Property Unit_Price() As Double
        Get
            Unit_Price = m_Unit_Price
        End Get
        Set(ByVal Value As Double)
            Try
                If Value = m_Unit_Price And ((Not blnLoading) Or (blnFirstLoad)) Then Exit Property 'Throw New OEException(OEError.Exit_Property, False, False)

                Dim ptOverride As PermissionType
                Dim ptLessThanCost As PermissionType

                ' First check if we can override price without problem.
                    ptOverride = m_oOrder.Validation.PriceOverridePermission()

                If Value <> m_Unit_Price And m_Unit_Price <> 0 Then
                    ' Confirmation needed when system preventing
                    If ptOverride = PermissionType.Prevent And Not blnLoading Then
                        Dim mbRes As MsgBoxResult
                        mbRes = MsgBox(New OEExceptionMessage(OEError.Unit_Price_Has_Been_Changed_Is_This_OK).Message, MsgBoxStyle.YesNo)

                        ' If user says no, then throw to UI for a cancel event witout showing a message
                        If mbRes = MsgBoxResult.No Then
                            Throw New OEException(OEError.Unit_Price_Has_Been_Changed, True, False)
                        End If
                    End If
                End If
                ' Now we check if price less than cost
                ptLessThanCost = m_oOrder.Validation.PriceLessThanCostPermission(m_Item_No, m_Loc, Value)

                ' Confirmation needed when system preventing
                If ptLessThanCost = PermissionType.Prevent And Not blnLoading Then
                    Dim mbRes As MsgBoxResult
                    mbRes = MsgBox(New OEExceptionMessage(OEError.Unit_Price_Below_Unit_Cost_Is_This_OK).Message, MsgBoxStyle.YesNo)

                    ' If user says no, then throw to UI for a cancel event witout showing a message
                    If mbRes = MsgBoxResult.No And Not blnLoading Then
                        Throw New OEException(OEError.Unit_Price_Below_Unit_Cost, True, False)
                    End If
                End If

                'Everything is ok, we can change the price
                Try
                    'If m_Unit_Price <> 0 And m_Unit_Price <> Value Then
                    If m_Unit_Price <> Value Then
                        m_Unit_Price = Value
                        Call Price_Changed()
                        If ptOverride = PermissionType.Warning And Not blnLoading Then Throw New OEException(OEError.Unit_Price_Has_Been_Changed)
                    End If
                Catch oe_er As OEException
                    MsgBox(oe_er.Message)
                Finally
                    If ptLessThanCost = PermissionType.Warning And Not blnLoading Then
                        Throw New OEException(OEError.Unit_Price_Below_Unit_Cost)
                    End If
                End Try

                m_Dirty = True

            Catch oe_er As OEException
                If oe_er.Cancel Then Throw New OEException(oe_er.Message, oe_er)
            End Try

        End Set

    End Property
    Public Property Unit_Weight() As Double
        Get
            Unit_Weight = m_Unit_Weight
        End Get
        Set(ByVal Value As Double)
            m_Unit_Weight = Value
        End Set
    End Property
    Public Property Uom() As String
        Get
            Uom = m_Uom
        End Get
        Set(ByVal Value As String)
            m_Uom = Value
        End Set
    End Property
    Public Property Uom_Ratio() As Double
        Get
            Uom_Ratio = m_Uom_Ratio
        End Get
        Set(ByVal Value As Double)
            m_Uom_Ratio = Value
        End Set
    End Property
    Public Property Update_Fg() As String
        Get
            Update_Fg = m_Update_Fg
        End Get
        Set(ByVal Value As String)
            m_Update_Fg = Value
        End Set
    End Property
    Public Property User_Def_Fld_1() As String
        Get
            User_Def_Fld_1 = m_User_Def_Fld_1
        End Get
        Set(ByVal Value As String)
            m_User_Def_Fld_1 = Value
        End Set
    End Property
    Public Property User_Def_Fld_2() As String
        Get
            User_Def_Fld_2 = m_User_Def_Fld_2
        End Get
        Set(ByVal Value As String)
            m_User_Def_Fld_2 = Value
        End Set
    End Property
    Public Property User_Def_Fld_3() As String
        Get
            User_Def_Fld_3 = m_User_Def_Fld_3
        End Get
        Set(ByVal Value As String)
            m_User_Def_Fld_3 = Value
        End Set
    End Property
    Public Property User_Def_Fld_4() As String
        Get
            User_Def_Fld_4 = m_User_Def_Fld_4
        End Get
        Set(ByVal Value As String)
            m_User_Def_Fld_4 = Value
        End Set
    End Property
    Public Property User_Def_Fld_5() As String
        Get
            User_Def_Fld_5 = m_User_Def_Fld_5
        End Get
        Set(ByVal Value As String)
            m_User_Def_Fld_5 = Value
        End Set
    End Property
    Public Property Vendor_No() As String
        Get
            Vendor_No = m_Vendor_No
        End Get
        Set(ByVal Value As String)
            m_Vendor_No = Value
        End Set
    End Property
    Public Property Warranty_Date() As Date
        Get
            Warranty_Date = m_Warranty_Date
        End Get
        Set(ByVal Value As Date)
            m_Warranty_Date = Value
        End Set
    End Property

    Public Property Loading() As Boolean
        Get
            Loading = blnLoading
        End Get
        Set(ByVal value As Boolean)
            blnLoading = value
        End Set
    End Property

    'Public Sub Load_Item_Data()


    '    Dim dt As New DataTable
    '    Dim db As New cDBA

    '    ' Attention, si on modifie la procedure ici, aussi modifier la procedure dans COrder.OrderLinesLoad.
    '    ' Fait separement pour grandement accelerer le processus, afin que le load ne passe pas au travers de
    '    ' tout les declencheurs.
    '    Dim strSql As String = _
    '    "SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, " & _
    '    "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(PQV.Qty_Ordered,0)) AS FLOAT) AS Qty_On_Hand, " & _
    '    "           DBO.OEI_Item_Price('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', 1) as Unit_Price, " & _
    '    "           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, ISNULL(I.UOM, '') AS UOM, " & _
    '    "           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, " & _
    '    "           ISNULL(I.BkOrd_Fg, 'N') AS BkOrd_Fg, '' AS Route, I.Activity_Cd, ISNULL(I.Prod_Cat, '') AS Prod_Cat, " & _
    '    "FROM 		IMITMIDX_SQL I WITH (nolock) " & _
    '    "LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _
    '    "LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
    '    "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " & _
    '    "WHERE 		I.Item_No = '" & m_Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
    '    "ORDER BY   I.Item_No "

    '    dt = db.DataTable(strsql)

    '    If dt.Rows.Count > 0 Then
    '        Dim dtRow As DataRow = dt.Rows(0)

    '        With dtRow
    '            m_Item_Exist = True

    '            m_Loc = m_oOrder.Ordhead.Mfg_Loc
    '            m_Item_Desc_1 = IIf(.Item("Item_Desc_1").Equals(DBNull.Value), "", .Item("Item_Desc_1"))
    '            m_Item_Desc_2 = IIf(.Item("Item_Desc_2").Equals(DBNull.Value), "", .Item("Item_Desc_2"))
    '            m_Qty_Ordered = 0
    '            m_Qty_To_Ship = 0 ' rsItem.Fields("Qty_Ordered").Value
    '            m_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_To_Ship").Value
    '            m_Qty_On_Hand = .Item("Qty_On_Hand")
    '            m_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_Ship_Prev").Value
    '            m_Qty_Prev_Bkord = .Item("Qty_BkOrd")
    '            m_Unit_Price = .Item("Unit_Price")
    '            m_Orig_Price = m_Unit_Price
    '            m_Calc_Price = 0
    '            m_Discount_Pct = m_oOrder.Ordhead.Discount_Pct
    '            m_Promise_Dt = IIf(FormatDateTime(m_oOrder.Ordhead.Shipping_Dt, DateFormat.ShortDate) = "1/1/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
    '            m_Request_Dt = m_Promise_Dt
    '            m_Req_Ship_Dt = m_Promise_Dt
    '            m_Uom = IIf(.Item("UOM").Equals(DBNull.Value), "", .Item("UOM").Value)
    '            m_Bkord_Fg = IIf(.Item("Bkord_Fg").Equals(DBNull.Value), "N", .Item("Bkord_Fg"))
    '            m_Prod_Cat = IIf(.Item("Prod_Cat").Equals(DBNull.Value), "", .Item("Prod_Cat"))

    '            'm_Explode_Kit
    '            'm_Prc_Cd_Orig_Price()
    '            'm_Bin_Fg
    '            'm_Item_Release_No
    '            'm_Mfg_Ord_No

    '            ' Necessary info to load for import
    '            m_Unit_Weight = IIf(.Item("item_Weight").Equals(DBNull.Value), 0, .Item("item_Weight").Value)

    '            m_Recalc_Sw = "Y" ' IIf(rsItem.Fields("Recalc_Sw").Equals(DBNull.Value), "N", rsItem.Fields("Recalc_Sw").Value)

    '        End With

    '    End If
    '    dt.Dispose()

    'End Sub

#End Region

#Region "Elements change #############################################"

    Private Sub Item_No_Changed(ByVal value As String)

        If Trim(m_Item_No) = Trim(value) Then Exit Sub

        Dim m_OrderType As String = ""
        Dim m_byRegion As String = ""
        Dim m_star As String = ""
        Dim m_Mfg_Loc As String = ""

        Dim m_ItemNo As String = ""
        '  Dim m_Prod_Cat As String = ""
        '  Dim m_Calc_Price As Double = 0


        Try
            If (blnLoading And Not (blnFirstLoad)) Or value Is Nothing Or Trim(value) = "" Then Exit Sub

            If Not (Kit_Component) Then
                m_oOrder.WorkLine = m_oOrder.GetMaxLine_No(m_oOrder.Ordhead.Ord_GUID)
                m_oOrder.WorkLine += 1
                m_Line_No = m_oOrder.WorkLine
            End If

            Dim dt As DataTable
            Dim db As New cDBA

            'dim dt as datatable
            ' Attention, si on modifie la procedure ici, aussi modifier la procedure dans COrder.OrderLinesLoad.
            ' Fait separement pour grandement accelerer le processus, afin que le load ne passe pas au travers de
            ' tout les declencheurs.

            Dim strsql As String =
            "SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd, " &
            "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, " &
            "           ISNULL(L.Qty_On_Hand,0) AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0) AS Qty_Allocated, " &
            "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS FLOAT) AS Qty_On_Hand, " &
            "           DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &
            "           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, " &
            "           ISNULL(I.UOM, '') AS UOM, ISNULL(I.Item_Weight, 0) AS Item_Weight, " &
            "           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, " &
            "           '' AS Route, 0 as ProductProof, 0 as AutoCompleteReship, I.Activity_Cd, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, " &
            "           ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(L.Avg_Cost, 0) AS Avg_Cost, " &
            "           CASE WHEN ISNULL(X.Stocked_Fg, '') = '' THEN ISNULL(I.Stocked_FG, 'N') ELSE X.Stocked_Fg END AS Stocked_Fg, " &
            "           CASE WHEN ISNULL(X.BkOrd_Fg, '') = '' THEN ISNULL(I.BkOrd_Fg, 'N') ELSE X.BkOrd_Fg END AS BkOrd_Fg, " &
            "           CASE WHEN SUBSTRING (I.Item_No, 1, 2) = '10' THEN 1 ELSE 0 END AS Is_Refill_Fg, " &
            "           ISNULL(PK.Pack_CD, '') AS Packaging, L.status as Loc_Status " &
            "FROM 		IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  MDB_ITEM_MASTER MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd " &
            "LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No " &
            "LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " &
            "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
            "LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MI.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
            "LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
            "WHERE 		I.Item_No = '" & value & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " &
            "ORDER BY   I.Item_No "
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _

            strsql =
            "SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd,  ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(m_oOrder.Ordhead.Mfg_Loc) & "'), '' ) as Loc," &
            "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, " &
            "           ISNULL(L.Qty_On_Hand,0) AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0) AS Qty_Allocated, " &
            "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS FLOAT) AS Qty_On_Hand, " &
            "           DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &
            "           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, " &
            "           ISNULL(I.UOM, '') AS UOM, ISNULL(I.Item_Weight, 0) AS Item_Weight, " &
            "           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, " &
            "           '' AS Route, 0 as ProductProof, 0 as AutoCompleteReship, I.Activity_Cd, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, " &
            "           ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(L.Avg_Cost, 0) AS Avg_Cost, " &
            "           CASE WHEN ISNULL(X.Stocked_Fg, '') = '' THEN ISNULL(I.Stocked_FG, 'N') ELSE X.Stocked_Fg END AS Stocked_Fg, " &
            "           CASE WHEN ISNULL(X.BkOrd_Fg, '') = '' THEN ISNULL(I.BkOrd_Fg, 'N') ELSE X.BkOrd_Fg END AS BkOrd_Fg, " &
            "           CASE WHEN SUBSTRING (I.Item_No, 1, 2) = '10' THEN 1 ELSE 0 END AS Is_Refill_Fg, " &
            "           ISNULL(PK.Pack_CD, '') AS Packaging, L.status as Loc_Status " &
            "FROM 		IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  MDB_ITEM_MASTER MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd " &
            "LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No " &
            "LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " &
            "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
            "LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MI.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
            "LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
            "WHERE 		I.Item_No = '" & value & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " &
            "ORDER BY   I.Item_No "
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _

            '++ID 07222024
            strsql =
            "SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd,  ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(m_oOrder.Ordhead.Mfg_Loc) & "'), '' ) as Loc," _
          & "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, " _
          & "           ISNULL(L.Qty_On_Hand,0) AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0) AS Qty_Allocated, " _
          & "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS FLOAT) AS Qty_On_Hand, " _
          & " CASE WHEN '" & m_oOrder.Ordhead.Mfg_Loc & "' = '3N' THEN dbo.OE_Item_Price_Type10 (10,'" & m_oOrder.Ordhead.Curr_Cd & "',I.Item_No) ElSE " _
          & "           DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1)  " _
          & " END as Unit_Price,  " _
          & "           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, " _
          & "           ISNULL(I.UOM, '') AS UOM, ISNULL(I.Item_Weight, 0) AS Item_Weight, " _
          & "           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, " _
          & "           '' AS Route, 0 as ProductProof, 0 as AutoCompleteReship, I.Activity_Cd, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, " _
          & "           ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(L.Avg_Cost, 0) AS Avg_Cost, " _
          & "           CASE WHEN ISNULL(X.Stocked_Fg, '') = '' THEN ISNULL(I.Stocked_FG, 'N') ELSE X.Stocked_Fg END AS Stocked_Fg, " _
          & "           CASE WHEN ISNULL(X.BkOrd_Fg, '') = '' THEN ISNULL(I.BkOrd_Fg, 'N') ELSE X.BkOrd_Fg END AS BkOrd_Fg, " _
          & "           CASE WHEN SUBSTRING (I.Item_No, 1, 2) = '10' THEN 1 ELSE 0 END AS Is_Refill_Fg, " _
          & "           ISNULL(PK.Pack_CD, '') AS Packaging, L.status as Loc_Status " _
          & " FROM 		IMITMIDX_SQL I WITH (nolock) " _
          & " LEFT JOIN  MDB_ITEM_MASTER MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd " _
          & " LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No " _
          & " LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " _
          & " LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " _
          & " LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MI.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " _
          & " LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " _
          & " WHERE 		I.Item_No = '" & value & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " _
          & " ORDER BY   I.Item_No "
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _



            If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                strsql =
          "SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd,  ISNULL([dbo].[Return_Location_By_ProdCat](RTrim(LTrim(I.Item_No)),'" & Trim(m_oOrder.Ordhead.Profit_Center) & "','" & Trim(m_oOrder.Ordhead.Mfg_Loc) & "'), '' ) as Loc," _
       & "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, " _
       & "           ISNULL(L.Qty_On_Hand,0) AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0) AS Qty_Allocated, " _
       & "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS FLOAT) AS Qty_On_Hand, " _
       & "           DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " _
       & "           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, " _
       & "           ISNULL(I.UOM, '') AS UOM, ISNULL(I.Item_Weight, 0) AS Item_Weight, " _
       & "           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, " _
       & "           '' AS Route, 0 as ProductProof, 0 as AutoCompleteReship, I.Activity_Cd, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, " _
       & "           ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(L.Avg_Cost, 0) AS Avg_Cost, " _
       & "           CASE WHEN ISNULL(X.Stocked_Fg, '') = '' THEN ISNULL(I.Stocked_FG, 'N') ELSE X.Stocked_Fg END AS Stocked_Fg, " _
       & "           CASE WHEN ISNULL(X.BkOrd_Fg, '') = '' THEN ISNULL(I.BkOrd_Fg, 'N') ELSE X.BkOrd_Fg END AS BkOrd_Fg, " _
       & "           CASE WHEN SUBSTRING (I.Item_No, 1, 2) = '10' THEN 1 ELSE 0 END AS Is_Refill_Fg, " _
       & "           ISNULL(PK.Pack_CD, '') AS Packaging, L.status as Loc_Status " _
       & "FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) " _
       & "LEFT JOIN  MDB_ITEM_MASTER MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd " _
       & "LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No " _
       & "LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No " _
       & "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " _
       & "LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MI.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " _
       & "LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " _
       & "WHERE 		I.Item_No = '" & value & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " _
       & "ORDER BY   I.Item_No "
            End If




            dt = db.DataTable(strsql)

            ' si on a trouvé aucun record, alors on ouvre la fenetre de recherche
            If dt.Rows.Count = 0 OrElse dt.Rows(0).Item("Loc_Status") = "H" Then
                If value.Length < 4 And Not (value.Trim = "44" Or value.Trim = "66" Or value.Trim = "88") Then
                    ' Entry too short
                    m_Item_Exist = False
                    m_Loc = m_oOrder.Ordhead.Mfg_Loc
                    m_Item_Desc_1 = String.Empty
                    m_Item_Desc_2 = String.Empty
                    m_Qty_Ordered = 0
                    m_Qty_Overpick = 0
                    m_Qty_Scrap = 0
                    m_Qty_To_Ship = 0 ' rsItem.Fields("Qty_Ordered").Value
                    m_Qty_Lines = 1 ' rsItem.Fields("Qty_Ordered").Value
                    m_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_To_Ship").Value
                    m_Qty_Inventory = 0
                    m_Qty_Allocated = 0
                    m_Qty_On_Hand = 0
                    m_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_Ship_Prev").Value
                    m_Qty_Prev_Bkord = 0
                    '==============================
                    'Added June 6, 2017 - T. Louzon
                    m_Unit_Price_BeforeSpecial = 0
                    'Added July 17, 2017 - T. Louzon
                    m_Mfg_Item_No = ""
                    m_RouteExtension = ""
                    '==============================

                    m_Unit_Price = 0
                    m_Orig_Price = 0
                    m_Calc_Price = 0
                    m_Discount_Pct = m_oOrder.Ordhead.Discount_Pct
                    m_Promise_Dt = IIf(m_oOrder.Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") = "01/01/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
                    'm_Promise_Dt = IIf(FormatDateTime(m_oOrder.Ordhead.Shipping_Dt, DateFormat.ShortDate) = "1/1/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
                    'm_Request_Dt = m_Promise_Dt
                    If m_oOrder.Ordhead.In_Hands_Date.Year = 0 Then
                        m_Request_Dt = m_Promise_Dt
                    Else
                        m_Request_Dt = m_oOrder.Ordhead.In_Hands_Date
                    End If
                    'm_Request_Dt = m_oOrder.Ordhead.In_Hands_Date
                    m_Req_Ship_Dt = m_Promise_Dt
                    m_Uom = ""
                    m_Bkord_Fg = "N"
                    m_Stocked_Fg = "N"
                    m_Prod_Cat = String.Empty
                    m_Unit_Cost = 0
                    m_Unit_Weight = 0
                    m_End_Item_Cd = String.Empty
                    m_Recalc_Sw = "Y"
                    m_ProductProof = 0
                    m_AutoCompleteReship = 0
                    m_Create_Imprint = False
                    m_Create_Kit = False
                    m_Create_Traveler = False
                    m_Is_Refill_Fg = False
                    m_Has_Refills_Fg = False
                    m_SaveToDB = False
                    m_Loc_Status = String.Empty
                    m_isMixMatch = False
                    m_Mix_Group = 0
                Else
                    Dim oOrderSelect As New frmProductLineEntry() ' m_oOrder)
                    oOrderSelect.Customer_No = m_oOrder.Ordhead.Cus_No
                    oOrderSelect.Currency = m_oOrder.Ordhead.Curr_Cd ' "USD"
                    oOrderSelect.Item_No = value
                    oOrderSelect.Loc = m_oOrder.Ordhead.Mfg_Loc

                    'oOrderSelect.ProductLineEntry.LoadProductGrid()
                    oOrderSelect.ShowDialog()
                    'If oOrderSelect.ProductStatus = 1 Then
                    '    Call InsertProductsInOrder(oOrderSelect)
                    'End If
                    'If dgvItems.Rows.Count > 2 Then
                    '    dsDataSet.Tables(0).Rows.RemoveAt(e.RowIndex)
                    'End If

                    m_Create_Imprint = False
                    m_Create_Kit = False
                    m_Create_Traveler = False
                    m_SaveToDB = False

                End If
            Else

                m_Item_Exist = True
                m_Loc = Trim(dt.Rows(0).Item("Loc").ToString) 'm_oOrder.Ordhead.Mfg_Loc

                m_Item_Desc_1 = IIf(dt.Rows(0).Item("Item_Desc_1").Equals(DBNull.Value), "", dt.Rows(0).Item("Item_Desc_1"))
                m_Item_Desc_2 = IIf(dt.Rows(0).Item("Item_Desc_2").Equals(DBNull.Value), "", dt.Rows(0).Item("Item_Desc_2"))
                m_Qty_Ordered = 0
                m_Qty_Overpick = 0
                m_Qty_Scrap = 0
                m_Qty_To_Ship = 0 ' rsItem.Fields("Qty_Ordered").Value
                m_Qty_Lines = 1 ' rsItem.Fields("Qty_Ordered").Value
                m_Qty_Prev_To_Ship = 1 ' rsItem.Fields("Qty_To_Ship").Value
                Dim oPrice As cOEPriceFile = New cOEPriceFile(value, m_oOrder.Ordhead.Curr_Cd)
                Dim intInv_Buffer As Integer = IIf(m_Loc.Trim = "1", oPrice.Get_Inventory_Buffer, 0)
                m_Qty_Inventory = dt.Rows(0).Item("Qty_Inventory") - intInv_Buffer
                m_Qty_Allocated = dt.Rows(0).Item("Qty_Allocated")
                m_Qty_On_Hand = dt.Rows(0).Item("Qty_On_Hand") - intInv_Buffer
                m_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_Ship_Prev").Value
                m_Qty_Prev_Bkord = dt.Rows(0).Item("Qty_BkOrd")
                m_Unit_Price = dt.Rows(0).Item("Unit_Price")


                m_Orig_Price = m_Unit_Price
                m_Calc_Price = 0

                'here I need add return of discount from function

                '++ID 5.3.2018  -----Discount-----
                If String.IsNullOrEmpty(m_oOrder.Ordhead.User_Def_Fld_2) <> True Then m_OrderType = m_oOrder.Ordhead.User_Def_Fld_2

                If String.IsNullOrEmpty(m_oOrder.Ordhead.Customer.region) <> True Then m_byRegion = m_oOrder.Ordhead.Customer.region

                If String.IsNullOrEmpty(m_oOrder.Ordhead.Customer.ClassificationId) <> True Then m_star = m_oOrder.Ordhead.Customer.ClassificationId.ToUpper

                If String.IsNullOrEmpty(m_oOrder.Ordhead.Mfg_Loc) <> True Then m_Mfg_Loc = m_oOrder.Ordhead.Mfg_Loc



                ' If String.IsNullOrEmpty(dgvItems.Rows(e.RowIndex).Cells(Columns.Calc_Price).Value.ToString) <> True Then m_Calc_Price = dgvItems.Rows(e.RowIndex).Cells(Columns.Calc_Price).Value


                If String.IsNullOrEmpty(value) <> True Then
                    m_ItemNo = value

                    ' If String.IsNullOrEmpty(Return_Prod_Cat(m_ItemNo)) <> True Then m_Prod_Cat = Return_Prod_Cat(m_ItemNo) ' come from insertt line in the function
                End If

                m_Prod_Cat = IIf(dt.Rows(0).Item("Prod_Cat").Equals(DBNull.Value), "", dt.Rows(0).Item("Prod_Cat"))

                'I put in comments this variable because is main variable for discountr
                '  m_oOrder.Ordhead.Discount_Pct = Ret_Discount(m_OrderType, m_byRegion, m_star, m_ItemNo, m_Prod_Cat, m_Calc_Price, m_Mfg_Loc)

                '---------------------------- NEW -------------------------------
                '(ByVal pcus_no As String, ByVal pitem_no As String, ByVal pqty_order As Decimal, ByVal pord_type As String)

                'project closed 88REWARDS
                '++ID 1.17.2019
                'Dim dtT As DataTable = Nothing

                'dtT = TenPercent(m_oOrder.Ordhead.Cus_No, m_ItemNo, m_Tot_Qty_Ordered, m_OrderType)

                'If dtT.Rows.Count <> 0 Then

                '    If dtT.Rows(0).Item(enumTenPercent.eFinnalPriceRevised).ToString <> -1 Then
                '        m_Discount_Pct = dtT.Rows(0).Item(enumTenPercent.eTenPercent).ToString
                '    Else
                '        m_Discount_Pct = Ret_Discount(m_OrderType, m_byRegion, m_star, m_ItemNo, m_Prod_Cat, m_Calc_Price, m_Mfg_Loc) 'm_oOrder.Ordhead.Discount_Pct
                '    End If

                'Else
                '    m_Discount_Pct = Ret_Discount(m_OrderType, m_byRegion, m_star, m_ItemNo, m_Prod_Cat, m_Calc_Price, m_Mfg_Loc) 'm_oOrder.Ordhead.Discount_Pct
                'End If

                    m_Discount_Pct = Ret_Discount(m_OrderType, m_byRegion, m_star, m_ItemNo, m_Prod_Cat, m_Calc_Price, m_Mfg_Loc) 'm_oOrder.Ordhead.Discount_Pct

                m_Promise_Dt = IIf(m_oOrder.Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") = "01/01/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
                m_Promise_Dt = IIf(m_oOrder.Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0") & "/" & m_oOrder.Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") = "01/01/0001", Now.Date, m_oOrder.Ordhead.Shipping_Dt)
                'm_Promise_Dt = IIf(FormatDateTime(m_oOrder.Ordhead.Shipping_Dt, DateFormat.ShortDate) = "1/1/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
                If m_oOrder.Ordhead.In_Hands_Date.Year = 1 Then
                    m_Request_Dt = m_Promise_Dt
                Else
                    m_Request_Dt = m_oOrder.Ordhead.In_Hands_Date
                End If
                'm_Request_Dt = m_Promise_Dt
                m_Req_Ship_Dt = m_Promise_Dt
                m_Uom = IIf(dt.Rows(0).Item("UOM").Equals(DBNull.Value), "", dt.Rows(0).Item("UOM"))
                m_Bkord_Fg = IIf(dt.Rows(0).Item("Bkord_Fg").Equals(DBNull.Value), "N", dt.Rows(0).Item("Bkord_Fg"))

                m_Unit_Cost = dt.Rows(0).Item("Avg_Cost")
                m_End_Item_Cd = dt.Rows(0).Item("Kit_Feat_Fg")
                '++ID 7.19.2018 put in comments
                'm_Stocked_Fg = IIf(m_End_Item_Cd = "K", "Y", IIf(dt.Rows(0).Item("Stocked_Fg").Equals(DBNull.Value), "N", dt.Rows(0).Item("Stocked_Fg")))
                '++ID 17.19.2018 display value from imitdidx_sql
                m_Stocked_Fg = dt.Rows(0).Item("Stocked_Fg").ToString ' IIf(m_End_Item_Cd = "K", "N", IIf(dt.Rows(0).Item("Stocked_Fg").Equals(DBNull.Value), "N", dt.Rows(0).Item("Stocked_Fg")))

                m_Item_Cd = IIf(dt.Rows(0).Item("Item_Cd").Equals(DBNull.Value), "", dt.Rows(0).Item("Item_Cd"))

                ETA_From_Item_No = m_Item_No

                'm_Explode_Kit
                'm_Prc_Cd_Orig_Price()
                'm_Bin_Fg
                'm_Item_Release_No
                'm_Mfg_Ord_No

                ' Necessary info to load for import
                m_Unit_Weight = IIf(dt.Rows(0).Item("Item_Weight").Equals(DBNull.Value), 0, dt.Rows(0).Item("Item_Weight"))

                m_Recalc_Sw = "Y" ' IIf(rsItem.Fields("Recalc_Sw").Equals(DBNull.Value), "N", rsItem.Fields("Recalc_Sw").Value)
                m_ProductProof = 0
                m_AutoCompleteReship = 0

                m_Is_Refill_Fg = dt.Rows(0).Item("Is_Refill_Fg")
                m_Has_Refills_Fg = 0

                m_Create_Imprint = True
                m_Create_Kit = True
                m_Create_Traveler = True

                If (m_oOrder.Ordhead.Mfg_Loc.Trim = "3" Or m_oOrder.Ordhead.Mfg_Loc.Trim = "3N") And m_RouteID = 0 Then
                    If value.ToUpper.Trim = "88RANDOMSAMP" Then
                        m_Route = "Completed"
                        m_RouteID = 39
                    ElseIf value.ToUpper.Trim = "88SPECSAMPL" Then
                        m_Route = "Completed"
                        m_RouteID = 39
                    ElseIf value.ToUpper.Trim = "88SELFPROMO" Then
                        m_Route = "Completed"
                        m_RouteID = 39
                    Else
                        m_Route = "Random Samples"
                        m_RouteID = 16
                    End If
                    'm_Route = "Random Samples"
                    'm_RouteID = 16
                End If

                '==============================
                'Added June 6, 2017 - T. Louzon   - If item number changes, make the BeforeSpecial match the Unit_Price.
                m_Unit_Price_BeforeSpecial = dt.Rows(0).Item("Unit_Price")
                '==============================
                'Added June 27, 2017 - T. Louzon
                m_RouteExtension = f_GetRouteExtensionAndData(m_RouteID)
                If m_RouteExtension = "XX" Then
                    m_Mfg_Item_No = m_Item_No
                Else
                    m_Mfg_Item_No = m_Item_No & "-" & m_RouteExtension
                End If
                '==============================


                m_SaveToDB = True

                m_Imprint = New cImprint(m_Item_Guid)
                m_Imprint.SaveToDB = True

                If Kit_Component Then

                Else
                    m_Imprint.Packaging = IIf(dt.Rows(0).Item("Packaging").Equals(DBNull.Value), "", dt.Rows(0).Item("Packaging"))
                End If

                m_Traveler = New CTraveler(m_Item_Guid)
                'm_Imprint.SaveToDB = True

                'm_Kit = New cKit(Me)

            End If

        Catch er As Exception
            MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub
    'closed project for 10% 88REWARDs
    'Public Enum enumTenPercent
    '    eitem_no
    '    ePriceEQP
    '    ePriceNoEQP
    '    eFinnalPriceRevised
    '    eqty_order
    '    eTenPercent
    'End Enum
    '++ID 1.17.2019 add percent in discound cell 88REWARDs
    'Public Function TenPercent(ByVal pcus_no As String, ByVal pitem_no As String, ByVal pqty_order As Decimal, ByVal pord_type As String) As DataTable
    '    TenPercent = Nothing
    '    Try
    '        Dim dt As DataTable
    '        Dim db As New cDBA
    '        Dim strSQl As String = ""

    '        strSQl = "Select * from [dbo].[SelfPromoTenPercentAdd]( '" & pcus_no & "', '" & pitem_no & "',Default, " & pqty_order & ", '" & Ord_Type & "',Default)"

    '        dt = db.DataTable(strSQl)

    '        Return dt


    '        'when start use the function before need check rows count
    '        'if zero this means missing SP orders  or item selected not in SP order
    '        'if is one but FinnalPriceRevized is -1  this means min_qty in accomplished or order type   is one of this (@ord_type = 'RS' or @ord_type = 'SP' or @ord_type = 'SB' or @ord_type = 'SS')
    '    Catch ex As Exception

    '    End Try

    'End Function

    ' Use this sub when reloading Item data from an OEI Order to get item default values without overwriting or popping alarms
    Public Sub Item_No_Reload()

        Try

            Dim dt As DataTable
            Dim db As New cDBA

            'dim dt as datatable

            Dim strsql As String =
            "Select 	P.Picture As Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, " &
            "           CAST(0 As FLOAT) As Qty_Ordered, CAST(0 As FLOAT) As Qty_To_Ship, " &
            "           ISNULL(L.Qty_On_Hand,0) As Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0) AS Qty_Allocated, " &
            "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(m_Item_Guid) & "'),0)) AS FLOAT) AS Qty_On_Hand, " &
            "           DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &
            "           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, " &
            "           ISNULL(I.UOM, '') AS UOM, ISNULL(I.Item_Weight, 0) AS Item_Weight, " &
            "           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, " &
            "           '' AS Route, 0 as ProductProof, 0 AS AutoCompleteReship, I.Activity_Cd, ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg, " &
            "           ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(L.Avg_Cost, 0) AS Avg_Cost, " &
            "           CASE WHEN ISNULL(X.Stocked_Fg, '') = '' THEN ISNULL(I.Stocked_FG, 'N') ELSE X.Stocked_Fg END AS Stocked_Fg, " &
            "           CASE WHEN ISNULL(X.BkOrd_Fg, '') = '' THEN ISNULL(I.BkOrd_Fg, 'N') ELSE X.BkOrd_Fg END AS BkOrd_Fg, " &
            "           CASE WHEN SUBSTRING (I.Item_No, 1, 2) = '10' THEN 1 ELSE 0 END AS Is_Refill_Fg " '&
            '"FROM 		IMITMIDX_SQL I WITH (nolock) " &
            '"LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No " &
            '"LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " &
            '"LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
            '"WHERE 		I.Item_No = '" & m_Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_Loc & "' " &
            '"ORDER BY   I.Item_No "
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _

            '++ID 02.19.2023 
            If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then
                strsql &= " " &
                     "FROM 		IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No " &
            "LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " &
            "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
            "WHERE 		I.Item_No = '" & m_Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_Loc & "' " &
            "ORDER BY   I.Item_No "

            Else
                strsql &= " " &
                    "FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) " &
           "LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No " &
           "LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No " &
           "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
           "WHERE 		I.Item_No = '" & m_Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_Loc & "' " &
           "ORDER BY   I.Item_No "

            End If




            dt = db.DataTable(strsql)

            ' si on a trouvé aucun record, alors on ouvre la fenetre de recherche
            If dt.Rows.Count <> 0 Then
                'm_Item_Exist = True
                'm_Loc = m_oOrder.Ordhead.Mfg_Loc
                'm_Item_Desc_1 = IIf(rsItem.Fields("Item_Desc_1").Equals(DBNull.Value), "", rsItem.Fields("Item_Desc_1").Value)
                'm_Item_Desc_2 = IIf(rsItem.Fields("Item_Desc_2").Equals(DBNull.Value), "", rsItem.Fields("Item_Desc_2").Value)
                'm_Qty_Ordered = 0
                'm_Qty_To_Ship = 0 ' rsItem.Fields("Qty_Ordered").Value
                'm_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_To_Ship").Value
                m_End_Item_Cd = dt.Rows(0).Item("Kit_Feat_Fg")
                Dim oPrice As cOEPriceFile = New cOEPriceFile(m_Item_No, m_oOrder.Ordhead.Curr_Cd)
                Dim intInv_Buffer As Integer = IIf(m_Loc.Trim = "1", oPrice.Get_Inventory_Buffer, 0)
                m_Qty_Inventory = dt.Rows(0).Item("Qty_Inventory") - intInv_Buffer
                'm_Qty_Inventory = dt.Rows(0).Item("Qty_Inventory")
                m_Qty_Allocated = dt.Rows(0).Item("Qty_Allocated")
                m_Qty_On_Hand = dt.Rows(0).Item("Qty_On_Hand") - intInv_Buffer
                'm_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_Ship_Prev").Value
                'm_Qty_Prev_Bkord = rsItem.Fields("Qty_BkOrd").Value
                'm_Unit_Price = rsItem.Fields("Unit_Price").Value
                'm_Orig_Price = m_Unit_Price
                'm_Calc_Price = 0
                'm_Discount_Pct = m_oOrder.Ordhead.Discount_Pct
                'm_Promise_Dt = IIf(FormatDateTime(m_oOrder.Ordhead.Shipping_Dt, DateFormat.ShortDate) = "1/1/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
                'm_Request_Dt = m_Promise_Dt
                'm_Req_Ship_Dt = m_Promise_Dt
                'm_Uom = IIf(rsItem.Fields("UOM").Equals(DBNull.Value), "", rsItem.Fields("UOM").Value)
                'm_Bkord_Fg = IIf(rsItem.Fields("Bkord_Fg").Equals(DBNull.Value), "N", rsItem.Fields("Bkord_Fg").Value)
                'm_Stocked_Fg = IIf(rsItem.Fields("Stocked_Fg").Equals(DBNull.Value), "N", rsItem.Fields("Stocked_Fg").Value)
                'm_Prod_Cat = IIf(rsItem.Fields("Prod_Cat").Equals(DBNull.Value), "", rsItem.Fields("Prod_Cat").Value)
                'm_Unit_Cost = rsItem.Fields("Avg_Cost").Value

                'm_Explode_Kit
                'm_Prc_Cd_Orig_Price()
                'm_Bin_Fg
                'm_Item_Release_No
                'm_Mfg_Ord_No

                ' Necessary info to load for import
                'm_Unit_Weight = IIf(rsItem.Fields("Item_Weight").Equals(DBNull.Value), 0, rsItem.Fields("Item_Weight").Value)

                'm_Recalc_Sw = "Y" ' IIf(rsItem.Fields("Recalc_Sw").Equals(DBNull.Value), "N", rsItem.Fields("Recalc_Sw").Value)

                'm_Create_Imprint = True
                'm_Create_Kit = True
                'm_Create_Traveler = True
                'm_SaveToDB = True

                'm_Imprint = New cImprint(m_Item_Guid)
                'm_Imprint.SaveToDB = True

                'm_Traveler = New CTraveler(m_Item_Guid)
                'm_Imprint.SaveToDB = True

                'm_Kit = New cKit()

            End If

            Call Qty_To_Ship_Reload()

        Catch er As Exception
            MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Loc_Changed(ByVal value As String)

        If blnLoading And Not blnFirstLoad Then Exit Sub

        Dim dt As DataTable
        Dim db As New cDBA

        Try
            'Dim strsql As String = _
            '"SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, " & _
            '"           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(PQV.Qty_Ordered,0)) AS INT) AS VARCHAR) AS Qty_On_Hand, " & _
            '"           DBO.OEI_Item_Price('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', 1) as Unit_Price, " & _
            '"           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, " & _
            '"           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, '' AS Route " & _
            '"FROM 		IMITMIDX_SQL I WITH (nolock) " & _
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _
            '"LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
            '"LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " & _
            '"WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & value & "' " & _
            '"ORDER BY   I.Item_No "
            Dim strsql As String =
            "SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, " &
            "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, " &
            "           ISNULL(L.Qty_On_Hand,0) AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0) AS Qty_Allocated, " &
            "           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS INT) AS VARCHAR) AS Qty_On_Hand, " &
            "           DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &
            "           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, " &
            "           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, '' AS Route, L.status as Loc_Status " &
            "FROM 		IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " &
            "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
            "WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & value & "' " &
            "ORDER BY   I.Item_No "
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _


            '++ID 07292024 added case for unit price in case if 
            strsql =
            "SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, " &
            "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, " &
            "           ISNULL(L.Qty_On_Hand,0) AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0) AS Qty_Allocated, " &
            "           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS INT) AS VARCHAR) AS Qty_On_Hand, " &
            "  CASE WHEN '3N' = '" & value & "' THEN dbo.OE_Item_Price_Type10 (10,'" & m_oOrder.Ordhead.Curr_Cd & "',I.Item_No) ElSE  " &
            "           DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) END as Unit_Price, " &
            "           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, " &
            "           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, '' AS Route, L.status as Loc_Status " &
            " FROM 		IMITMIDX_SQL I WITH (nolock) " &
            " LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " &
            " LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
            " WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & value & "' " &
            " ORDER BY   I.Item_No "




            If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                strsql =
            "SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, " &
            "           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, " &
            "           ISNULL(L.Qty_On_Hand,0) AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0) AS Qty_Allocated, " &
            "           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS INT) AS VARCHAR) AS Qty_On_Hand, " &
            "           DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Unit_Price, " &
            "           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, " &
            "           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, '' AS Route, L.status as Loc_Status " &
            "FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No " &
            "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
            "WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & value & "' " &
            "ORDER BY   I.Item_No "
            End If




            dt = db.DataTable(strsql)

            ' si on a trouvé aucun record, alors on ouvre la fenetre de recherche

            If dt.Rows.Count <> 0 Then
                'update location status
                m_Loc_Status = dt.Rows(0).Item("Loc_Status")

                '   If value = "3N" Then
                m_Unit_Price = dt.Rows(0).Item("Unit_Price")

                m_Qty_Inventory = dt.Rows(0).Item("Qty_Inventory")
                m_Qty_Allocated = dt.Rows(0).Item("Qty_Allocated")
                m_Qty_On_Hand = dt.Rows(0).Item("Qty_on_Hand")


                ' Else
                'm_Unit_Price = dt.Rows(0).Item("Unit_Price")

                '    m_Qty_Inventory = dt.Rows(0).Item("Qty_Inventory")
                '    m_Qty_Allocated = dt.Rows(0).Item("Qty_Allocated")
                '    m_Qty_On_Hand = dt.Rows(0).Item("Qty_on_Hand")

                'End If


                If m_End_Item_Cd = "K" Then Set_Item_Kit_Qty(m_Item_No, value)

                'Call Qty_To_Ship_Changed(Item_No, value, False) ' Qty_Ordered = tmpQty_Ordered
                If m_Qty_Lines > 1 Then
                    Call Multiline_Qty_To_Ship_Changed(Item_No, value, False)
                Else
                    Call Qty_To_Ship_Changed(Item_No, value, False)
                End If

            Else
                If Not blnLoading Then Throw New OEException(OEError.Item_Location_Not_Found, True)
            End If
            ' We catch system error, we throw back custom errors to UI.
        Catch oe_er As OEException
            Throw oe_er ' New OEException(oe_er.Number)
        Catch er As Exception
            MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Qty_Ordered_Changed(ByVal value As Double)
        'm_Qty_Prev_To_Ship is actual, value is new value

        'If m_Qty_To_Ship = value Then Exit Sub

        Call SetUnit_Price(value)





    End Sub

    Public Sub Qty_To_Ship_Changed(ByVal Item_No As String, ByVal pstrLoc As String, ByVal pblnForce As Boolean)
        'm_Qty_Prev_To_Ship is actual, value is new value

        If blnLoading And Not blnFirstLoad Then Exit Sub

        Dim dt As DataTable
        Dim db As New cDBA

        Try
            If ETA_From_Item_No Is Nothing Then ETA_From_Item_No = ""

            'Dim strsql As String = _
            '"SELECT 	I.Item_No, " & _
            '"           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(PQV.Qty_Ordered,0)) AS INT) AS VARCHAR) AS Qty_On_Hand " & _
            '"FROM 		IMITMIDX_SQL I WITH (nolock) " & _
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _
            '"LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
            '"WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & pstrLoc & "' " & _
            '"ORDER BY   I.Item_No "

            Dim strsql As String =
            "SELECT 	I.Item_No, " &
            "           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS INT) AS VARCHAR) AS Qty_On_Hand, L.status as Loc_Status " &
            "FROM 		IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " &
            "WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & pstrLoc & "' " &
            "ORDER BY   I.Item_No "
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _

            '++ID 02.17.2023  
            If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                strsql =
            "SELECT 	I.Item_No, " &
            "           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS INT) AS VARCHAR) AS Qty_On_Hand, L.status as Loc_Status " &
            "FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No " &
            "WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & pstrLoc & "' " &
            "ORDER BY   I.Item_No "
            End If


            dt = db.DataTable(strsql)
            ' si on a trouvé aucun record, alors on ouvre la fenetre de recherche

            If dt.Rows.Count <> 0 Then

                If Not (m_Kit.IsAKit Or m_End_Item_Cd = "K") Then
                    'If (Not (m_Kit.IsAKit)) Then
                    Dim oPrice As cOEPriceFile = New cOEPriceFile(m_Item_No, m_oOrder.Ordhead.Curr_Cd)

                    'Dim intInv_Buffer As Integer = IIf(m_Loc.Trim = "1"  , oPrice.Get_Inventory_Buffer, 0)

                    '++ID 07302024 commented aboveline for to add to check line location
                    Dim intInv_Buffer As Integer = IIf(m_Loc.Trim = "1" And g_oOrdline.Loc = "1", oPrice.Get_Inventory_Buffer, 0)

                    'm_Qty_Inventory = dt.Rows(0).Item("Qty_Inventory") - oPrice.Get_Inventory_Buffer


                    m_Qty_On_Hand = dt.Rows(0).Item("Qty_On_Hand") - intInv_Buffer


                End If

                'If item is no stock, we do not calculate a BO.
                'If (m_Stocked_Fg = "N")  Then
                If (m_Stocked_Fg = "N") And Not (m_Kit.IsAKit Or m_End_Item_Cd = "K") Then

                    m_Qty_Prev_To_Ship = Qty_To_Ship
                    m_Qty_Prev_Bkord = 0

                ElseIf pblnForce Then

                    m_Qty_Prev_To_Ship = Qty_To_Ship

                    If Qty_Ordered > Qty_To_Ship Then
                        m_Qty_Prev_Bkord = Qty_Ordered - Qty_To_Ship
                        ETA_From_Item_No = m_Item_No
                    Else
                        m_Qty_Prev_Bkord = 0
                    End If

                Else

                    If m_End_Item_Cd = "K" Then
                        Set_Item_Kit_Qty(m_Item_No, m_Loc)
                    End If

                    If m_Qty_On_Hand <= 0 Then

                        m_Qty_Prev_To_Ship = 0 'm_Qty_On_Hand
                        m_Qty_Prev_Bkord = Qty_Ordered
                        If m_End_Item_Cd.Trim = "" Then ETA_From_Item_No = m_Item_No

                    Else

                        'MB++ MOD 2014/04/04
                        'If Qty_To_Ship > m_Qty_On_Hand Then ' old code mode to :
                        If Qty_To_Ship > m_Qty_On_Hand Then ' And Qty_Ordered > m_Qty_On_Hand Then

                            m_Qty_Prev_To_Ship = m_Qty_On_Hand
                            m_Qty_Prev_Bkord = IIf(Qty_Ordered > Qty_To_Ship, Qty_Ordered, Qty_To_Ship) - m_Qty_On_Hand
                            If m_End_Item_Cd.Trim = "" Then ETA_From_Item_No = m_Item_No

                        Else

                            m_Qty_Prev_To_Ship = IIf(Qty_Ordered > Qty_To_Ship, Qty_Ordered, Qty_To_Ship)
                            m_Qty_Prev_Bkord = Qty_Ordered - m_Qty_Prev_To_Ship
                            'm_Qty_Prev_Bkord = 0

                        End If

                    End If

                End If

                Call Price_Changed()

                If ETA_From_Item_No <> "" Then
                    'Dim oCalc As New cETA_Calculator
                    'oCalc.Item_No = ETA_From_Item_No.Trim
                    'oCalc.Location = m_Loc
                    'oCalc.Qty_Available = m_Qty_On_Hand
                    'oCalc.Qty_On_Hand = m_Qty_On_Hand
                    'oCalc.Qty_Ordered = Qty_Ordered
                    'oCalc.Qty_Prev_Bkord = m_Qty_Prev_Bkord
                    'oCalc.Calculate_ETA_For_Item()
                    'm_Extra_9 = oCalc.Extra_9
                    'm_Extra_10 = oCalc.Extra_10
                    'Debug.Print(oCalc.Extra_9 & " " & oCalc.Extra_10)
                    'Debug.Print("")
                    Call Calculate_ETA()
                Else
                    m_Extra_9 = ""
                    m_Extra_10 = 0
                End If

            Else
                If Not blnLoading Then Throw New OEException(OEError.Item_Location_Not_Found, True)
                m_Qty_On_Hand = 0

            End If

            If Not blnLoading And m_Stocked_Fg.ToString = "Y" And (m_Qty_To_Ship > m_Qty_On_Hand And m_Qty_To_Ship > 0 And m_Loc <> "3N" And frmOrder.ignoreQtyMismatch.Equals(False)) Then
                'If Not blnLoading And (m_Qty_To_Ship > m_Qty_On_Hand And m_Qty_To_Ship > 0) And Not (m_oOrder.FromHistory) Then
                '    MsgBox("Qty to ship is greater than qty available.", MsgBoxStyle.OkOnly, "Error")

                '++ID 02.20.2023
                If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then
                    MsgBox("Qty to ship is greater than qty available.", MsgBoxStyle.OkOnly, "Error")
                Else
                    Dim OEI_PRE_READ_ORDER_EXCEPTION As cOEI_PRE_READ_ORDER_EXCEPTION = New cOEI_PRE_READ_ORDER_EXCEPTION(m_oOrder.Ordhead.Oe_Po_No, m_oOrder.Ordhead.OEI_Ord_No, Item_No, 2789)
                    Dim OEI_PRE_READ_ORDER_EXCEPTION_DAL As cOEI_PRE_READ_ORDER_EXCEPTION_DAL = New cOEI_PRE_READ_ORDER_EXCEPTION_DAL()
                    OEI_PRE_READ_ORDER_EXCEPTION_DAL.Save(OEI_PRE_READ_ORDER_EXCEPTION)
                    MsgBox("For US1 location, if no stock, " & vbCrLf & " please check Available inventory in Canadian company and change to location 1.", MsgBoxStyle.OkOnly, "Error")
                End If



            End If

            If m_Loc_Status = "H" Then
                m_Qty_To_Ship = 0.0
                MsgBox("You cannot order this item from location " & m_Loc.Trim() & ". It is currently on hold", MsgBoxStyle.OkOnly, "Error")
            End If


            ' We catch system error, we throw back custom errors to UI.
        Catch oe_er As OEException
            Throw oe_er ' New OEException(oe_er.Number)
        Catch er As Exception
            MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Multiline_Qty_To_Ship_Changed(ByVal Item_No As String, ByVal pstrLoc As String, ByVal pblnForce As Boolean)
        'm_Qty_Prev_To_Ship is actual, value is new value

        If blnLoading And Not blnFirstLoad Then Exit Sub

        Dim dt As DataTable
        Dim db As New cDBA

        Try
            'Dim strsql As String = _
            '"SELECT 	I.Item_No, " & _
            '"           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(PQV.Qty_Ordered,0)) AS INT) AS VARCHAR) AS Qty_On_Hand " & _
            '"FROM 		IMITMIDX_SQL I WITH (nolock) " & _
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _
            '"LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
            '"WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & pstrLoc & "' " & _
            '"ORDER BY   I.Item_No "

            Dim strsql As String =
            "SELECT 	I.Item_No, " &
            "           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS INT) AS VARCHAR) AS Qty_On_Hand " &
            "FROM 		IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " &
            "WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & pstrLoc & "' " &
            "ORDER BY   I.Item_No "
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _


            '++ID 02.19.2023
            If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                strsql =
            "SELECT 	I.Item_No, " &
            "           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS INT) AS VARCHAR) AS Qty_On_Hand " &
            "FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No " &
            "WHERE 		I.Item_No = '" & Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & pstrLoc & "' " &
            "ORDER BY   I.Item_No "
            End If



            dt = db.DataTable(strsql)
            ' si on a trouvé aucun record, alors on ouvre la fenetre de recherche

            If dt.Rows.Count <> 0 Then

                If Not (m_Kit.IsAKit Or m_End_Item_Cd = "K") Then
                    'If (Not (m_Kit.IsAKit)) Then
                    Dim oPrice As cOEPriceFile = New cOEPriceFile(m_Item_No, m_oOrder.Ordhead.Curr_Cd)
                    Dim intInv_Buffer As Integer = IIf(m_Loc.Trim = "1", oPrice.Get_Inventory_Buffer, 0)
                    m_Qty_On_Hand = dt.Rows(0).Item("Qty_On_Hand") - intInv_Buffer
                End If

                ' If item is no stock, we do not calculate a BO.
                If (m_Stocked_Fg = "N") And Not (m_Kit.IsAKit Or m_End_Item_Cd = "K") Then

                    m_Qty_Prev_To_Ship = Qty_To_Ship * Qty_Lines
                    m_Qty_Prev_Bkord = 0

                ElseIf pblnForce Then

                    m_Qty_Prev_To_Ship = Qty_To_Ship * Qty_Lines

                    If Qty_Ordered > Qty_To_Ship Then
                        m_Qty_Prev_Bkord = (Qty_Ordered - Qty_To_Ship) * Qty_Lines
                    Else
                        m_Qty_Prev_Bkord = 0
                    End If

                Else

                    If m_Qty_On_Hand <= 0 Then

                        m_Qty_Prev_To_Ship = 0 'm_Qty_On_Hand
                        m_Qty_Prev_Bkord = Qty_Ordered * Qty_Lines

                    Else

                        If Qty_To_Ship * Qty_Lines > m_Qty_On_Hand Then
                            m_Qty_Prev_To_Ship = m_Qty_On_Hand
                            m_Qty_Prev_Bkord = (IIf(Qty_Ordered > Qty_To_Ship, Qty_Ordered, Qty_To_Ship) * Qty_Lines) - m_Qty_On_Hand
                        Else
                            m_Qty_Prev_To_Ship = IIf(Qty_Ordered > Qty_To_Ship, Qty_Ordered, Qty_To_Ship) * Qty_Lines
                            m_Qty_Prev_Bkord = 0
                        End If

                    End If
                End If

                Call Calculate_ETA()

            Else
                If Not blnLoading Then Throw New OEException(OEError.Item_Location_Not_Found, True)
                m_Qty_On_Hand = 0
            End If

            If Not blnLoading And (m_Qty_To_Ship > m_Qty_On_Hand And m_Qty_To_Ship > 0 And m_Loc <> "3N") Then
                ' MsgBox("Qty to ship is greater than qty available.", MsgBoxStyle.OkOnly, "Error")

                '++ID 02.20.2023
                If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then
                    MsgBox("Qty to ship is greater than qty available.", MsgBoxStyle.OkOnly, "Error")
                Else
                    ' MsgBox("Qty to ship is greater than qty available.", MsgBoxStyle.OkOnly, "Error")

                    MsgBox("For US1 location, if no stock, " & vbCrLf & " please check Available inventory in Canadian company and change to location 1.", MsgBoxStyle.OkOnly, "Error")
                End If

            End If

            ' We catch system error, we throw back custom errors to UI.
        Catch oe_er As OEException
            Throw oe_er ' New OEException(oe_er.Number)
        Catch er As Exception
            MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Use this function to get new reload values and new ETA date/qty if needed
    ' It will remove BO's when new quantities appear.
    Private Sub Qty_To_Ship_Reload()
        'm_Qty_Prev_To_Ship is actual, value is new value

        Dim dt As DataTable
        Dim db As New cDBA

        Try

            If ETA_From_Item_No Is Nothing Then ETA_From_Item_No = ""

            Dim strsql As String =
            "SELECT 	I.Item_No, " &
            "           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS INT) AS VARCHAR) AS Qty_On_Hand " &
            "FROM 		IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No " &
            "WHERE 		I.Item_No = '" & m_Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_Loc & "' " &
            "ORDER BY   I.Item_No "
            '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _

            '++ID 02.19.2023
            If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                strsql =
            "SELECT 	I.Item_No, " &
            "           CAST(CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, '" & Trim(Item_Guid) & "'),0)) AS INT) AS VARCHAR) AS Qty_On_Hand " &
            "FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) " &
            "LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No " &
            "WHERE 		I.Item_No = '" & m_Item_No & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_Loc & "' " &
            "ORDER BY   I.Item_No "
            End If






            dt = db.DataTable(strsql)
            ' si on a trouvé aucun record, alors on ouvre la fenetre de recherche

            If dt.Rows.Count <> 0 Then

                If Not (m_Kit.IsAKit Or m_End_Item_Cd = "K") Then
                    'If (Not (m_Kit.IsAKit)) Then
                    Dim oPrice As cOEPriceFile = New cOEPriceFile(m_Item_No, m_oOrder.Ordhead.Curr_Cd)
                    Dim intInv_Buffer As Integer = IIf(m_Loc.Trim = "1", oPrice.Get_Inventory_Buffer, 0)
                    m_Qty_On_Hand = dt.Rows(0).Item("Qty_On_Hand") - intInv_Buffer
                End If

                ' If item is no stock, we do not calculate a BO.
                If (m_Stocked_Fg = "N") And Not (m_Kit.IsAKit Or m_End_Item_Cd = "K") Then

                    m_Qty_Prev_To_Ship = Qty_To_Ship
                    m_Qty_Prev_Bkord = 0

                Else

                    If m_End_Item_Cd = "K" Then
                        Set_Item_Kit_Qty(m_Item_No, m_Loc)
                    End If

                    If m_Qty_On_Hand <= 0 Then

                        m_Qty_Prev_To_Ship = 0 'm_Qty_On_Hand
                        m_Qty_Prev_Bkord = Qty_Ordered
                        If m_End_Item_Cd.Trim = "" Then ETA_From_Item_No = m_Item_No

                    Else

                        If Qty_To_Ship >= m_Qty_On_Hand Then
                            m_Qty_Prev_To_Ship = m_Qty_On_Hand
                            m_Qty_Prev_Bkord = IIf(Qty_Ordered > Qty_To_Ship, Qty_Ordered, Qty_To_Ship) - m_Qty_On_Hand
                            If m_End_Item_Cd.Trim = "" Then ETA_From_Item_No = m_Item_No
                        Else
                            m_Qty_Prev_To_Ship = IIf(Qty_Ordered > Qty_To_Ship, Qty_Ordered, Qty_To_Ship)
                            m_Qty_Prev_Bkord = 0
                        End If

                    End If

                End If

            End If

            Call Calculate_ETA()

        Catch er As Exception
            MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Price_Changed()

        If blnLoading And Not blnFirstLoad Then Exit Sub
        m_Calc_Price = Math.Round(((100 - m_Discount_Pct) / 100) * (m_Unit_Price * m_Qty_To_Ship), 2)

    End Sub

    Private Sub Request_Dt_Changed()

        'm_Promise_Dt = m_Request_Dt
        'm_Req_Ship_Dt = m_Request_Dt

    End Sub

    Private Sub Promise_Dt_Changed()

        m_Req_Ship_Dt = m_Promise_Dt

    End Sub

#End Region

    'Public Sub Calculate_ETA_For_Item()

    '    Try

    '        If m_Source = OEI_SourceEnum.frmProductLineEntry Or m_Source = OEI_SourceEnum.Macola_ExactRepeat Then
    '            Call Calculate_ETA_ProductLineEntry()
    '            Exit Sub
    '        End If

    '        ' From now on, we show ETA for everything not location 1 just in case...
    '        If m_Qty_Prev_Bkord <= 0 And m_Loc = "1" Then
    '            m_Extra_9 = " "
    '            m_Extra_10 = 0
    '            Exit Sub
    '        End If

    '        Dim strDate As String = ""
    '        Dim dblQty As Double = 0
    '        Dim dblQtyRequired As Double = 0

    '        Dim dtItems As DataTable = Nothing
    '        Dim dt As DataTable
    '        Dim db As New cDBA

    '        Dim strsql As String = ""

    '        'If m_Source = OEI_SourceEnum.frmOrder Then

    '        strsql = _
    '        "SELECT     L.ORD_GUID, L.ITEM_GUID, ISNULL(B.COMP_ITEM_GUID, '') AS COMP_ITEM_GUID, " & _
    '        "           CASE WHEN ISNULL(B.COMP_ITEM_GUID, '') = '' THEN L.Item_No ELSE B.Item_No END AS Item_No " & _
    '        "FROM 		OEI_ORDLIN L WITH (NOLOCK) " & _
    '        "LEFT JOIN	OEI_ORDBLD B WITH (NOLOCK) ON L.ORD_GUID = B.ORD_GUID AND L.ITEM_GUID = B.KIT_ITEM_GUID " & _
    '        "WHERE	    L.ORD_GUID = '" & m_oOrder.Ordhead.Ord_GUID & "' AND L.ITEM_GUID = '" & m_Item_Guid & "' "

    '        dtItems = db.DataTable(strsql)

    '        If dtItems.Rows.Count = 0 Then Exit Sub

    '        Dim strLaterDate As String = ""
    '        Dim dtLater_Promise_Dt As DateTime = Nothing ' CDate("01/01/2010")
    '        Dim dblLater_Qty As Double = 0

    '        ' IF IT IS A KIT, THERE WILL BE MORE THAN 1 LINE. IF AN ITEM, ALWAYS ONLY 1.
    '        For Each dtItemRow As DataRow In dtItems.Rows

    '            strDate = ""
    '            dblQty = 0
    '            dblQtyRequired = 0

    '            If dtItemRow.Item("Item_Guid") <> dtItemRow.Item("Comp_Item_Guid") Then

    '                strsql = "" & _
    '                "SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
    '                "           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
    '                "           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date, " & _
    '                "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0)) AS FLOAT) AS Qty_On_Hand " & _
    '                "FROM       IMItmIdx_Sql I WITH (Nolock) " & _
    '                "LEFT JOIN  MDB_ITEM MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd " & _
    '                "INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
    '                "LEFT JOIN 	OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
    '                "LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No AND Q.Stk_Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
    '                "WHERE      I.Item_No = '" & Trim(dtItemRow.Item("Item_No")) & "' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
    '                "ORDER BY   Q.Promise_Dt "

    '                'strsql = "" & _
    '                '"SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
    '                '"           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
    '                '"           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date, " & _
    '                '"           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0)) AS FLOAT) AS Qty_On_Hand " & _
    '                '"FROM       IMItmIdx_Sql I WITH (Nolock) " & _
    '                '"LEFT JOIN  MDB_ITEM MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd " & _
    '                '"INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & m_Loc & "' " & _
    '                '"LEFT JOIN 	OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
    '                '"LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No AND Q.Stk_Loc = '" & m_Loc & "' " & _
    '                '"WHERE      I.Item_No = '" & Trim(dtItemRow.Item("Item_No")) & "' AND L.Loc = '" & g_oOrdline.Loc & "' " & _
    '                '"ORDER BY   Q.Promise_Dt "

    '                'strsql = "" & _
    '                '"SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
    '                '"           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
    '                '"           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date " & _
    '                '"FROM       IMItmIdx_Sql I WITH (Nolock) " & _
    '                '"LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No " & _
    '                '"WHERE      I.Item_No = '" & Trim(m_Item_No) & "' AND Q.Stk_Loc = '" & m_Loc & "' " & _
    '                '"ORDER BY   Q.Promise_Dt "
    '                'Debug.Print(Trim(dtItemRow.Item("Item_No")))
    '                dt = db.DataTable(strsql)

    '                'If m_Qty_On_Hand < 0 Then dblQtyRequired = Math.Abs(m_Qty_On_Hand)
    '                'dblQtyRequired += m_Qty_Prev_Bkord

    '                Dim oOrdline As cOrdLine = Nothing

    '                If dtItems.Rows.Count > 1 And Trim(dtItemRow.Item("COMP_ITEM_GUID")) <> "" Then
    '                    If m_oOrder.OrdLines.Contains(Trim(dtItemRow.Item("COMP_ITEM_GUID"))) Then
    '                        oOrdline = m_oOrder.OrdLines.Item(Trim(dtItemRow.Item("COMP_ITEM_GUID")))
    '                    End If
    '                Else
    '                    If m_oOrder.OrdLines.Contains(Trim(dtItemRow.Item("ITEM_GUID"))) Then
    '                        oOrdline = m_oOrder.OrdLines.Item(Trim(dtItemRow.Item("ITEM_GUID")))
    '                    End If
    '                End If

    '                If dt.Rows.Count <> 0 And Not (oOrdline Is Nothing) Then ' And (m_Qty_Prev_Bkord > 0) Then  

    '                    If dt.Rows(0).Item("Qty_On_Hand") < m_Qty_Ordered Then

    '                        If m_Qty_On_Hand < 0 Then dblQtyRequired = Math.Abs(m_Qty_On_Hand)
    '                        dblQtyRequired += m_Qty_Prev_Bkord

    '                        Dim iCountRows As Integer = 0

    '                        For Each dtRow As DataRow In dt.Rows

    '                            iCountRows += 1

    '                            If (dblQty < dblQtyRequired And m_Qty_Ordered > dtRow.Item("Qty_On_Hand")) Or (m_Loc <> "1" And dblQty < dblQtyRequired) Then

    '                                dblQty = dblQty + dtRow.Item("QtyRemaining")

    '                                If dtRow.Item("Activity_Cd").ToString = "O" Then
    '                                    strDate = "DISCONTINUED"
    '                                    strLaterDate = "DISCONTINUED"
    '                                ElseIf dblQty < dblQtyRequired And iCountRows = dt.Rows.Count Then
    '                                    strDate = "QTY MISSING"
    '                                    strLaterDate = "QTY MISSING"
    '                                Else
    '                                    If IsDBNull(dtRow.Item("Promise_Dt")) Then
    '                                        strDate = "" ' dtRow.Item("Promise_Date").ToString
    '                                    Else

    '                                        Dim Promise_Dt As DateTime = dtRow.Item("Promise_Dt")

    '                                        If dtLater_Promise_Dt.Year = 1 Or (Promise_Dt >= dtLater_Promise_Dt And dtLater_Promise_Dt.Year <> 1) And strLaterDate = "" Then
    '                                            'If strLaterDate <> "" Then
    '                                            dtLater_Promise_Dt = Promise_Dt
    '                                            dblLater_Qty = dblQty
    '                                            'End If
    '                                        End If

    '                                        strDate = Promise_Dt.Month.ToString.PadLeft(2, "0") & "/" & Promise_Dt.Day.ToString.PadLeft(2, "0") & "/" & Promise_Dt.Year.ToString.PadLeft(4, "0")

    '                                    End If

    '                                End If

    '                            End If

    '                        Next dtRow

    '                        If Not (oOrdline Is Nothing) Then
    '                            oOrdline.Extra_9 = strDate
    '                            oOrdline.Extra_10 = dblQty
    '                        Else
    '                            Extra_9 = strDate
    '                            Extra_10 = dblQty
    '                        End If

    '                    End If

    '                Else

    '                    If Not (oOrdline Is Nothing) Then
    '                        oOrdline.Extra_9 = "NOT FOUND"
    '                        oOrdline.Extra_10 = 0
    '                    End If

    '                    If strLaterDate = "" And m_Qty_Prev_Bkord > 0 Then
    '                        strLaterDate = "QTY MISSING" ' & Trim(dtItemRow.Item("Item_No"))
    '                    End If

    '                End If

    '            End If

    '        Next dtItemRow

    '        If dtItems.Rows.Count > 1 Then
    '            If dblLater_Qty > 0 Then
    '                If strLaterDate = "" Then
    '                    Extra_9 = dtLater_Promise_Dt.Month.ToString.PadLeft(2, "0") & "/" & dtLater_Promise_Dt.Day.ToString.PadLeft(2, "0") & "/" & dtLater_Promise_Dt.Year.ToString.PadLeft(4, "0")
    '                Else
    '                    Extra_9 = strLaterDate
    '                End If
    '                Extra_10 = dblLater_Qty
    '            Else
    '                If m_Qty_Prev_Bkord > 0 Then
    '                    Extra_9 = "QTY MISSING"
    '                Else
    '                    Extra_9 = ""
    '                End If
    '                Extra_10 = 0
    '            End If
    '        End If

    '    Catch er As Exception
    '        MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub


    Public Sub Calculate_ETA()

        Try

            If m_Kit_Component Then Exit Sub

            If m_Source = OEI_SourceEnum.frmProductLineEntry Or m_Source = OEI_SourceEnum.Macola_ExactRepeat Then
                If m_Qty_Prev_Bkord <= 0 Then
                    m_Extra_9 = " "
                    m_Extra_10 = 0
                    Exit Sub
                End If
            End If

            Dim oCalc As New cETA_Calculator

            If m_End_Item_Cd = "K" Then
                oCalc.Item_No = ETA_From_Item_No.Trim
            Else
                If m_Stocked_Fg = "Y" Then
                    oCalc.Item_No = m_Item_No
                Else
                    Exit Sub
                End If
            End If

            oCalc.Location = m_Loc
            oCalc.Qty_Available = m_Qty_On_Hand
            oCalc.Qty_On_Hand = m_Qty_On_Hand
            oCalc.Qty_Ordered = Qty_Ordered
            oCalc.Qty_Prev_Bkord = m_Qty_Prev_Bkord
            If ETA_From_Item_No Is Nothing Then
                oCalc.Item_No = m_Item_No.ToString.Trim
            Else
                oCalc.Item_No = ETA_From_Item_No.Trim
            End If
            oCalc.Curr_Cd = m_oOrder.Ordhead.Curr_Cd

            oCalc.Calculate_ETA_For_Item()
            m_Extra_9 = oCalc.Extra_9
            m_Extra_10 = oCalc.Extra_10

            'If m_Source = OEI_SourceEnum.frmProductLineEntry Or m_Source = OEI_SourceEnum.Macola_ExactRepeat Then
            '    Call Calculate_ETA_ProductLineEntry()
            '    Exit Sub
            'End If

            '' From now on, we show ETA for everything not location 1 just in case...
            'If m_Qty_Prev_Bkord <= 0 And m_Loc = "1" Then
            '    m_Extra_9 = " "
            '    m_Extra_10 = 0
            '    Exit Sub
            'End If

            'Dim strDate As String = ""
            'Dim dblQty As Double = 0
            'Dim dblQtyRequired As Double = 0

            'Dim dtItems As DataTable = Nothing
            'Dim dt As DataTable
            'Dim db As New cDBA

            'Dim strsql As String = ""

            ''If m_Source = OEI_SourceEnum.frmOrder Then

            'strsql = _
            '"SELECT     L.ORD_GUID, L.ITEM_GUID, ISNULL(B.COMP_ITEM_GUID, '') AS COMP_ITEM_GUID, " & _
            '"           CASE WHEN ISNULL(B.COMP_ITEM_GUID, '') = '' THEN L.Item_No ELSE B.Item_No END AS Item_No " & _
            '"FROM 		OEI_ORDLIN L WITH (NOLOCK) " & _
            '"LEFT JOIN	OEI_ORDBLD B WITH (NOLOCK) ON L.ORD_GUID = B.ORD_GUID AND L.ITEM_GUID = B.KIT_ITEM_GUID " & _
            '"WHERE	    L.ORD_GUID = '" & m_oOrder.Ordhead.Ord_GUID & "' AND L.ITEM_GUID = '" & m_Item_Guid & "' "

            'dtItems = db.DataTable(strsql)

            'If dtItems.Rows.Count = 0 Then Exit Sub

            'Dim strLaterDate As String = ""
            'Dim dtLater_Promise_Dt As DateTime = Nothing ' CDate("01/01/2010")
            'Dim dblLater_Qty As Double = 0

            '' IF IT IS A KIT, THERE WILL BE MORE THAN 1 LINE. IF AN ITEM, ALWAYS ONLY 1.
            'For Each dtItemRow As DataRow In dtItems.Rows

            '    strDate = ""
            '    dblQty = 0
            '    dblQtyRequired = 0

            '    If dtItemRow.Item("Item_Guid") <> dtItemRow.Item("Comp_Item_Guid") Then

            '        strsql = "" & _
            '        "SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
            '        "           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
            '        "           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date, " & _
            '        "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0)) AS FLOAT) AS Qty_On_Hand " & _
            '        "FROM       IMItmIdx_Sql I WITH (Nolock) " & _
            '        "LEFT JOIN  MDB_ITEM MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd " & _
            '        "INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
            '        "LEFT JOIN 	OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
            '        "LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No AND Q.Stk_Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
            '        "WHERE      I.Item_No = '" & Trim(dtItemRow.Item("Item_No")) & "' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
            '        "ORDER BY   Q.Promise_Dt "

            '        'strsql = "" & _
            '        '"SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
            '        '"           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
            '        '"           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date, " & _
            '        '"           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0)) AS FLOAT) AS Qty_On_Hand " & _
            '        '"FROM       IMItmIdx_Sql I WITH (Nolock) " & _
            '        '"LEFT JOIN  MDB_ITEM MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd " & _
            '        '"INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & m_Loc & "' " & _
            '        '"LEFT JOIN 	OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
            '        '"LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No AND Q.Stk_Loc = '" & m_Loc & "' " & _
            '        '"WHERE      I.Item_No = '" & Trim(dtItemRow.Item("Item_No")) & "' AND L.Loc = '" & g_oOrdline.Loc & "' " & _
            '        '"ORDER BY   Q.Promise_Dt "

            '        'strsql = "" & _
            '        '"SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
            '        '"           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
            '        '"           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date " & _
            '        '"FROM       IMItmIdx_Sql I WITH (Nolock) " & _
            '        '"LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No " & _
            '        '"WHERE      I.Item_No = '" & Trim(m_Item_No) & "' AND Q.Stk_Loc = '" & m_Loc & "' " & _
            '        '"ORDER BY   Q.Promise_Dt "
            '        'Debug.Print(Trim(dtItemRow.Item("Item_No")))
            '        dt = db.DataTable(strsql)

            '        'If m_Qty_On_Hand < 0 Then dblQtyRequired = Math.Abs(m_Qty_On_Hand)
            '        'dblQtyRequired += m_Qty_Prev_Bkord

            '        Dim oOrdline As cOrdLine = Nothing

            '        If dtItems.Rows.Count > 1 And Trim(dtItemRow.Item("COMP_ITEM_GUID")) <> "" Then
            '            If m_oOrder.OrdLines.Contains(Trim(dtItemRow.Item("COMP_ITEM_GUID"))) Then
            '                oOrdline = m_oOrder.OrdLines.Item(Trim(dtItemRow.Item("COMP_ITEM_GUID")))
            '            End If
            '        Else
            '            If m_oOrder.OrdLines.Contains(Trim(dtItemRow.Item("ITEM_GUID"))) Then
            '                oOrdline = m_oOrder.OrdLines.Item(Trim(dtItemRow.Item("ITEM_GUID")))
            '            End If
            '        End If

            '        If dt.Rows.Count <> 0 And Not (oOrdline Is Nothing) Then ' And (m_Qty_Prev_Bkord > 0) Then  

            '            If dt.Rows(0).Item("Qty_On_Hand") < m_Qty_Ordered Then

            '                If m_Qty_On_Hand < 0 Then dblQtyRequired = Math.Abs(m_Qty_On_Hand)
            '                dblQtyRequired += m_Qty_Prev_Bkord

            '                Dim iCountRows As Integer = 0

            '                For Each dtRow As DataRow In dt.Rows

            '                    iCountRows += 1

            '                    If (dblQty < dblQtyRequired And m_Qty_Ordered > dtRow.Item("Qty_On_Hand")) Or (m_Loc <> "1" And dblQty < dblQtyRequired) Then

            '                        dblQty = dblQty + dtRow.Item("QtyRemaining")

            '                        If dtRow.Item("Activity_Cd").ToString = "O" Then
            '                            strDate = "DISCONTINUED"
            '                            strLaterDate = "DISCONTINUED"
            '                        ElseIf dblQty < dblQtyRequired And iCountRows = dt.Rows.Count Then
            '                            strDate = "QTY MISSING"
            '                            strLaterDate = "QTY MISSING"
            '                        Else
            '                            If IsDBNull(dtRow.Item("Promise_Dt")) Then
            '                                strDate = "" ' dtRow.Item("Promise_Date").ToString
            '                            Else

            '                                Dim Promise_Dt As DateTime = dtRow.Item("Promise_Dt")

            '                                If dtLater_Promise_Dt.Year = 1 Or (Promise_Dt >= dtLater_Promise_Dt And dtLater_Promise_Dt.Year <> 1) And strLaterDate = "" Then
            '                                    'If strLaterDate <> "" Then
            '                                    dtLater_Promise_Dt = Promise_Dt
            '                                    dblLater_Qty = dblQty
            '                                    'End If
            '                                End If

            '                                strDate = Promise_Dt.Month.ToString.PadLeft(2, "0") & "/" & Promise_Dt.Day.ToString.PadLeft(2, "0") & "/" & Promise_Dt.Year.ToString.PadLeft(4, "0")

            '                            End If

            '                        End If

            '                    End If

            '                Next dtRow

            '                If Not (oOrdline Is Nothing) Then
            '                    oOrdline.Extra_9 = strDate
            '                    oOrdline.Extra_10 = dblQty
            '                Else
            '                    Extra_9 = strDate
            '                    Extra_10 = dblQty
            '                End If

            '            End If

            '        Else

            '            If Not (oOrdline Is Nothing) Then
            '                oOrdline.Extra_9 = "NOT FOUND"
            '                oOrdline.Extra_10 = 0
            '            End If

            '            If strLaterDate = "" And m_Qty_Prev_Bkord > 0 Then
            '                strLaterDate = "QTY MISSING" ' & Trim(dtItemRow.Item("Item_No"))
            '            End If

            '        End If

            '    End If

            'Next dtItemRow

            'If dtItems.Rows.Count > 1 Then
            '    If dblLater_Qty > 0 Then
            '        If strLaterDate = "" Then
            '            Extra_9 = dtLater_Promise_Dt.Month.ToString.PadLeft(2, "0") & "/" & dtLater_Promise_Dt.Day.ToString.PadLeft(2, "0") & "/" & dtLater_Promise_Dt.Year.ToString.PadLeft(4, "0")
            '        Else
            '            Extra_9 = strLaterDate
            '        End If
            '        Extra_10 = dblLater_Qty
            '    Else
            '        If m_Qty_Prev_Bkord > 0 Then
            '            Extra_9 = "QTY MISSING"
            '        Else
            '            Extra_9 = ""
            '        End If
            '        Extra_10 = 0
            '    End If
            'End If

        Catch er As Exception
            MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Sub Calculate_ETA_ProductLineEntry()

    '    Try
    '        If m_Qty_Prev_Bkord <= 0 And m_Loc = "1" Then
    '            m_Extra_9 = " "
    '            m_Extra_10 = 0
    '            Exit Sub
    '        End If

    '        Dim strDate As String = ""
    '        Dim dblQty As Double = 0
    '        Dim dblQtyRequired As Double = 0

    '        Dim dtItems As DataTable = Nothing
    '        Dim dt As DataTable
    '        Dim db As New cDBA

    '        Dim strsql As String = ""

    '        If m_Source = OEI_SourceEnum.frmOrder Then

    '            strsql = _
    '            "SELECT     L.ORD_GUID, L.ITEM_GUID, ISNULL(B.COMP_ITEM_GUID, '') AS COMP_ITEM_GUID, " & _
    '            "           CASE WHEN ISNULL(B.COMP_ITEM_GUID, '') = '' THEN L.Item_No ELSE B.Item_No END AS Item_No " & _
    '            "FROM 		OEI_ORDLIN L WITH (NOLOCK) " & _
    '            "LEFT JOIN	OEI_ORDBLD B WITH (NOLOCK) ON L.ORD_GUID = B.ORD_GUID AND L.ITEM_GUID = B.KIT_ITEM_GUID " & _
    '            "WHERE	    L.ORD_GUID = '" & m_oOrder.Ordhead.Ord_GUID & "' AND L.ITEM_GUID = '" & m_Item_Guid & "' AND L.ITEM_GUID <> ISNULL(B.COMP_ITEM_GUID, '') "

    '            dtItems = db.DataTable(strsql)

    '        Else

    '            strsql = _
    '            "SELECT 	'1' AS Ord_Guid, I.Item_No AS Item_Guid, K.Comp_Item_No AS Comp_Item_Guid, K.Comp_Item_No AS Item_No " & _
    '            "FROM 		IMITMIDX_SQL I WITH (Nolock) " & _
    '            "INNER JOIN IMKITFIL_SQL K WITH (Nolock) ON I.Item_No = K.Item_No " & _
    '            "WHERE		I.Item_No = '" & m_Item_No & "' "

    '            dtItems = db.DataTable(strsql)

    '            If dtItems.Rows.Count = 0 Then

    '                strsql = _
    '                "SELECT 	'1' AS Ord_Guid, I.Item_No AS Item_Guid, '' AS Comp_Item_Guid, I.Item_No " & _
    '                "FROM 		IMITMIDX_SQL I WITH (Nolock) " & _
    '                "WHERE		I.Item_No = '" & m_Item_No & "' "

    '                dtItems = db.DataTable(strsql)

    '            End If

    '        End If

    '        If dtItems.Rows.Count = 0 Then Exit Sub

    '        Dim strLaterDate As String = ""
    '        Dim dtLater_Promise_Dt As DateTime = Nothing ' CDate("01/01/2010")
    '        Dim dblLater_Qty As Double = 0

    '        ' IF IT IS A KIT, THERE WILL BE MORE THAN 1 LINE. IF AN ITEM, ALWAYS ONLY 1.
    '        For Each dtItemRow As DataRow In dtItems.Rows

    '            strDate = ""
    '            dblQty = 0
    '            dblQtyRequired = 0

    '            strsql = "" & _
    '            "SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
    '            "           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
    '            "           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date, " & _
    '            "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0)) AS FLOAT) AS Qty_On_Hand " & _
    '            "FROM       IMItmIdx_Sql I WITH (Nolock) " & _
    '            "INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
    '            "LEFT JOIN 	OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
    '            "LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No AND Q.Stk_Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
    '            "WHERE      I.Item_No = '" & Trim(dtItemRow.Item("Item_No")) & "' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
    '            "ORDER BY   Q.Promise_Dt "

    '            'strsql = "" & _
    '            '"SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
    '            '"           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
    '            '"           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date " & _
    '            '"FROM       IMItmIdx_Sql I WITH (Nolock) " & _
    '            '"LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No " & _
    '            '"WHERE      I.Item_No = '" & Trim(m_Item_No) & "' AND Q.Stk_Loc = '" & m_Loc & "' " & _
    '            '"ORDER BY   Q.Promise_Dt "
    '            'Debug.Print(Trim(dtItemRow.Item("Item_No")))
    '            dt = db.DataTable(strsql)

    '            If m_Qty_On_Hand < 0 Then dblQtyRequired = Math.Abs(m_Qty_On_Hand)
    '            dblQtyRequired += m_Qty_Prev_Bkord

    '            Dim oOrdline As cOrdLine = Nothing

    '            If m_Source = OEI_SourceEnum.frmOrder Then
    '                If dtItems.Rows.Count > 1 And Trim(dtItemRow.Item("COMP_ITEM_GUID")) <> "" Then
    '                    If m_oOrder.OrdLines.Contains(Trim(dtItemRow.Item("COMP_ITEM_GUID"))) Then
    '                        oOrdline = m_oOrder.OrdLines.Item(Trim(dtItemRow.Item("COMP_ITEM_GUID")))
    '                    End If
    '                Else
    '                    If m_oOrder.OrdLines.Contains(Trim(dtItemRow.Item("ITEM_GUID"))) Then
    '                        oOrdline = m_oOrder.OrdLines.Item(Trim(dtItemRow.Item("ITEM_GUID")))
    '                    End If
    '                End If
    '            End If

    '            If dt.Rows.Count <> 0 And (m_Qty_Prev_Bkord > 0) Then ' And Not (oOrdline Is Nothing) Then

    '                Dim iCountRows As Integer = 0

    '                For Each dtRow As DataRow In dt.Rows

    '                    iCountRows += 1

    '                    If (dblQty < dblQtyRequired And m_Qty_Ordered > dtRow.Item("Qty_On_Hand")) Or (m_Loc <> "1" And dblQty < dblQtyRequired) Then

    '                        dblQty = dblQty + dtRow.Item("QtyRemaining")

    '                        If dtRow.Item("Activity_Cd").ToString = "O" Then
    '                            strDate = "DISCONTINUED"
    '                            strLaterDate = "DISCONTINUED"
    '                        ElseIf dblQty < dblQtyRequired And iCountRows = dt.Rows.Count Then
    '                            strDate = "QTY MISSING"
    '                            strLaterDate = "QTY MISSING"
    '                        Else
    '                            If IsDBNull(dtRow.Item("Promise_Dt")) Then
    '                                strDate = "" ' dtRow.Item("Promise_Date").ToString
    '                            Else

    '                                Dim Promise_Dt As DateTime = dtRow.Item("Promise_Dt")

    '                                If dtLater_Promise_Dt.Year = 1 Or (Promise_Dt > dtLater_Promise_Dt And dtLater_Promise_Dt.Year <> 1) And strLaterDate = "" Then
    '                                    'If strLaterDate <> "" Then
    '                                    dtLater_Promise_Dt = Promise_Dt
    '                                    dblLater_Qty = dblQty
    '                                    'End If
    '                                End If

    '                                strDate = Promise_Dt.Month.ToString.PadLeft(2, "0") & "/" & Promise_Dt.Day.ToString.PadLeft(2, "0") & "/" & Promise_Dt.Year.ToString.PadLeft(4, "0")

    '                            End If

    '                        End If

    '                    End If

    '                Next dtRow

    '                If Not (oOrdline Is Nothing) Then
    '                    oOrdline.Extra_9 = strDate
    '                    oOrdline.Extra_10 = dblQty
    '                Else
    '                    Extra_9 = strDate
    '                    Extra_10 = dblQty
    '                End If

    '            Else

    '                If Not (oOrdline Is Nothing) Then
    '                    oOrdline.Extra_9 = "NOT FOUND"
    '                    oOrdline.Extra_10 = 0
    '                End If

    '                If strLaterDate = "" And m_Qty_Prev_Bkord > 0 Then
    '                    strLaterDate = "QTY MISSING" ' & Trim(dtItemRow.Item("Item_No"))
    '                End If

    '            End If

    '        Next dtItemRow

    '        If dtItems.Rows.Count > 1 Then
    '            If dblLater_Qty > 0 Then
    '                If strLaterDate = "" Then
    '                    Extra_9 = dtLater_Promise_Dt.Month.ToString.PadLeft(2, "0") & "/" & dtLater_Promise_Dt.Day.ToString.PadLeft(2, "0") & "/" & dtLater_Promise_Dt.Year.ToString.PadLeft(4, "0")
    '                Else
    '                    Extra_9 = strLaterDate
    '                End If
    '                Extra_10 = dblLater_Qty
    '            Else
    '                If m_Qty_Prev_Bkord > 0 Then
    '                    Extra_9 = "QTY MISSING"
    '                Else
    '                    Extra_9 = ""
    '                End If
    '                Extra_10 = 0
    '            End If
    '        End If

    '    Catch er As Exception
    '        MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Private Sub Set_Item_Kit(ByVal pblnSave As Boolean)

        If (m_Source = OEI_SourceEnum.frmProductLineEntry) Then

            m_Kit = New cKit(Me) ' m_oOrder.Ordhead.Ord_GUID, m_Item_Guid, Me, pblnSave) ' m_Item_No, m_Loc, m_oOrder.Ordhead.Cus_No, m_oOrder.Ordhead.Curr_Cd)
            m_Kit.SetIsAKit(m_Item_No, m_Loc)

        Else

            Call Save()
            m_Kit = New cKit(m_oOrder.Ordhead.Ord_GUID, m_Item_Guid, Me, pblnSave) ' m_Item_No, m_Loc, m_oOrder.Ordhead.Cus_No, m_oOrder.Ordhead.Curr_Cd)

            If m_Kit.IsAKit Then

                Call Set_Item_Kit_Qty(m_Item_No, m_oOrder.Ordhead.Mfg_Loc)

                'Dim strSql As String
                'strSql = "SELECT * FROM OEI_Kit_Item_Loc_Qty_View WHERE Kit_Item_No = '" & m_Item_No & "' AND Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' "

                'Dim dt As New DataTable
                'Dim db As New cDBA

                'Dim dblInventory As Double = 99999999
                'Dim dblQty_Allocated As Double = 99999999
                'Dim dblQty_On_Hand As Double = 99999999

                'dt = db.DataTable(strSql)
                'If dt.Rows.Count <> 0 Then

                '    For Each dtRow As DataRow In dt.Rows
                '        If dblQty_On_Hand > (dtRow.Item("Qty_Available") - dtRow.Item("OEI_Qty_On_Ord")) Then

                '            dblInventory = dtRow.Item("Qty_On_Hand")
                '            dblQty_Allocated = (dtRow.Item("Qty_Allocated") + dtRow.Item("OEI_Qty_On_Ord"))
                '            dblQty_On_Hand = (dtRow.Item("Qty_Available") - dtRow.Item("OEI_Qty_On_Ord"))

                '        End If
                '    Next
                'End If

                'm_Qty_Inventory = dblInventory
                'm_Qty_Allocated = dblQty_Allocated
                'm_Qty_On_Hand = dblQty_On_Hand

            End If
        End If

    End Sub

    Public Sub Set_Item_Kit_Qty(ByVal pstrItem_No As String, ByVal pstrLoc As String)

        Dim strSql As String
        strSql = "SELECT (ISNULL(Qty_Available,0) - Isnull(OEI_Qty_On_Ord,0)) / qty_per_par as 'QtyOnHand_SCR',* FROM OEI_Kit_Item_Loc_Qty_View WHERE Kit_Item_No = '" & pstrItem_No & "' AND Loc = '" & pstrLoc & "' "

        Dim dt As New DataTable
        Dim db As New cDBA

        Dim dblInventory As Double = 99999999
        Dim dblQty_Allocated As Double = 99999999
        Dim dblQty_On_Hand As Double = 99999999
        Dim strItemForETA As String = ""

        dt = db.DataTable(strSql)
        If dt.Rows.Count <> 0 Then

            Dim qty_comp_hand(dt.Rows.Count - 1) As Integer
            Dim cpt As Int16 = 0
            Dim leftQty As Integer = 0

            For Each dtRow As DataRow In dt.Rows
                Dim oPrice As cOEPriceFile = New cOEPriceFile(dtRow.Item("Item_No").ToString, m_oOrder.Ordhead.Curr_Cd)
                Dim intInv_Buffer As Integer = IIf(m_Loc.Trim = "1", oPrice.Get_Inventory_Buffer, 0)
                If dblQty_On_Hand > (dtRow.Item("Qty_Available") - dtRow.Item("OEI_Qty_On_Ord") - intInv_Buffer) Then

                    dblInventory = dtRow.Item("Qty_On_Hand") - intInv_Buffer
                    dblQty_Allocated = (dtRow.Item("Qty_Allocated") + dtRow.Item("OEI_Qty_On_Ord"))
                    dblQty_On_Hand = (dtRow.Item("Qty_Available") - dtRow.Item("OEI_Qty_On_Ord") - intInv_Buffer)
                    strItemForETA = dtRow.Item("Item_No")

                    '++ID 12.16.2020 ------------- for SCR prod Cat fill array for to be compared after -------------

                    If Prod_Cat = "SCR" Then
                        qty_comp_hand(cpt) = (dtRow.Item("QtyOnHand_SCR") - intInv_Buffer)
                        cpt += 1
                    End If

                    '------------------------------------------------------------------------------------------------
                End If
            Next

            '----------------------------------------------------------------------------
            If Prod_Cat = "SCR" Then
                Dim min As Integer
                Dim max As Integer

                '-----------static test min-max------------
                Dim qty__hand(2) As Integer
                qty__hand(0) = 16878
                qty__hand(1) = 3547
                qty__hand(2) = 3548

                min = qty__hand(0)
                max = qty__hand(0)
                '------------------------------------------

                min = qty_comp_hand(0)
                max = qty_comp_hand(0)

                'loop array comp hand qty , for leave the less qty 
                For Each _qty1 As Integer In qty__hand 'ty_comp_hand

                    If _qty1 < min Then
                        min = _qty1
                    End If

                Next

                leftQty = min
                If leftQty >= 0 Then dblQty_On_Hand = leftQty Else dblQty_On_Hand = 0

            End If
            '-------------------------------------------------------------------------- 

        End If


        m_Qty_Inventory = dblInventory
        m_Qty_Allocated = dblQty_Allocated
        m_Qty_On_Hand = dblQty_On_Hand
        ETA_From_Item_No = strItemForETA

    End Sub

    Public Sub Set_Kit_Inventory()

        Try

            Dim strSql As String
            strSql = "SELECT * FROM OEI_Kit_Item_Loc_Qty_View WHERE Kit_Item_No = '" & m_Item_No & "' AND Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' "

            '++ID 02.19.2023
            If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                strSql = "SELECT * FROM [200].dbo.OEI_Kit_Item_Loc_Qty_View_US WHERE Kit_Item_No = '" & m_Item_No & "' AND Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' "

            End If


            Dim dt As New DataTable
            Dim db As New cDBA

            Dim dblInventory As Double = 99999999
            Dim dblQty_Allocated As Double = 99999999
            Dim dblQty_On_Hand As Double = 99999999
            Dim strItemForETA As String = ""

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then

                For Each dtRow As DataRow In dt.Rows
                    Dim oPrice As cOEPriceFile = New cOEPriceFile(dtRow.Item("Item_No").ToString, m_oOrder.Ordhead.Curr_Cd)
                    '++ID 02.20.2023 Or m_Loc.Trim = "US1"
                    Dim intInv_Buffer As Integer = IIf((m_Loc.Trim = "1" Or m_Loc.Trim = "US1"), oPrice.Get_Inventory_Buffer, 0)

                    If dblQty_On_Hand > (dtRow.Item("Qty_Available") - dtRow.Item("OEI_Qty_On_Ord") - intInv_Buffer) Then

                        dblInventory = dtRow.Item("Qty_On_Hand") - intInv_Buffer
                        dblQty_Allocated = (dtRow.Item("Qty_Allocated") + dtRow.Item("OEI_Qty_On_Ord"))
                        dblQty_On_Hand = (dtRow.Item("Qty_Available") - dtRow.Item("OEI_Qty_On_Ord") - intInv_Buffer)
                        strItemForETA = dtRow.Item("Item_No")

                    End If
                Next
            End If

            m_Qty_Inventory = dblInventory
            m_Qty_Allocated = dblQty_Allocated
            m_Qty_On_Hand = dblQty_On_Hand
            ETA_From_Item_No = strItemForETA

        Catch er As Exception
            MsgBox("Error in Ordline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Set_Item_Image()

        Dim data As Byte()
        'Dim connString As String = "YourConnectionString"

        'MemoryStream that will hold the byte array
        'before converting to a Bitmap
        Dim stream As MemoryStream = Nothing
        Dim bmp As Bitmap
        Dim img As Image = Nothing

        Dim imgBytes As Byte() = Nothing

        'Table that will hold the data returned
        Dim table As New DataTable("Images")
        m_Image = Nothing

        Try

            Dim strSql As String =
            "SELECT     * " &
            "FROM       Item_Pictures WITH (Nolock) " &
            "WHERE      RTRIM(Item_No) = '" & SqlCompliantString(m_Item_No) & "' "

            Dim dt As New DataTable
            Dim conn As New cDBA
            dt = conn.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                If Not (dt.Rows(0).Item("OEI_H_Pixels").Equals(DBNull.Value)) Then m_OEI_H_Pixels = dt.Rows(0).Item("OEI_H_Pixels")
                If Not (dt.Rows(0).Item("OEI_W_Pixels").Equals(DBNull.Value)) Then m_OEI_W_Pixels = dt.Rows(0).Item("OEI_W_Pixels")
                If Not (dt.Rows(0).Item("Rotate").Equals(DBNull.Value)) Then m_OEI_Image_Rotate = dt.Rows(0).Item("Rotate")

                data = CType(dt.Rows(0).Item("Picture"), Byte())
                stream = New MemoryStream(data, 0, data.Length)

                bmp = New Bitmap(stream)
                img = bmp
                m_Image = img '' New Bitmap(stream)
                Call ResizeImage()
            Else
                bmp = New Bitmap(124, 22)
                'Dim oResizedImage As Image
                img = bmp
                'm_Image = img ' oImage
                m_Image = img '' New Bitmap(stream)
                m_ImageHeight = 22
                m_ImageWidth = 124
            End If

            dt.Dispose()

        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub ResizeImage()

        Dim oImage, oResizedImage As Image
        'Dim data As Byte()
        'Dim ms As MemoryStream
        Dim dblRatio As Double

        Try
            oImage = m_Image ' .FromStream(ms)
            If Not (m_Image.Equals(System.DBNull.Value)) Then

                Dim bmp As Bitmap

                If m_OEI_W_Pixels < 400 Then

                    If m_OEI_W_Pixels < m_OEI_H_Pixels Then
                        If Not (m_OEI_Image_Rotate = "N") Then
                            oImage.RotateFlip(RotateFlipType.Rotate270FlipX)
                        End If
                    End If

                    bmp = New Bitmap(120, 120)

                    Using g As Graphics = Graphics.FromImage(bmp)
                        g.DrawImage(oImage, 0, 0, bmp.Width, bmp.Height)
                    End Using
                    oResizedImage = bmp
                    m_Image = oResizedImage

                    'myRow.Cells(Columns.Image).Value = oResizedImage ' oImage
                    m_ImageHeight = m_OEI_H_Pixels
                    'myRow.Height = oResizedImage.Height ' IIf(oResizedImage.Height > 25, oResizedImage.Height + 5, 25)
                    m_ImageWidth = m_OEI_W_Pixels

                Else

                    dblRatio = oImage.Height / oImage.Width
                    If dblRatio > 2 Then
                        If Not (m_OEI_Image_Rotate = "N") Then
                            oImage.RotateFlip(RotateFlipType.Rotate270FlipX)
                        End If
                        dblRatio = oImage.Width / oImage.Height
                    End If

                    If dblRatio > 1 Then
                        If dblRatio > 5.5 Then dblRatio = 5.5
                        bmp = New Bitmap(120, CInt(120 / dblRatio))
                    Else
                        bmp = New Bitmap(CInt(120 * dblRatio), 120)
                    End If

                    Using g As Graphics = Graphics.FromImage(bmp)
                        g.DrawImage(oImage, 0, 0, bmp.Width, bmp.Height)
                    End Using
                    oResizedImage = bmp
                    m_Image = oResizedImage

                    'myRow.Cells(Columns.Image).Value = oResizedImage ' oImage
                    m_ImageHeight = oResizedImage.Height
                    'myRow.Height = oResizedImage.Height ' IIf(oResizedImage.Height > 25, oResizedImage.Height + 5, 25)
                    m_ImageWidth = oResizedImage.Width

                End If

                'g.DrawImage(bit, 0, 0, rect, GraphicsUnit.Pixel)
                'PreviewPictureBox.Image = cropBitmap
                'oImage.Clone()

                'oImage = myRow.Cells(Columns.Image).Value
                'oImage = Image.FromHbitmap(myRow.Cells(Columns.Image).Value)

                'dgvItems.Rows(myRow.Index).Height = oResizedImage.Height
                'dgvItems.Columns(0).Width = IIf(dgvItems.Columns(0).Width > myRow.Cells(0).Size.Width, dgvItems.Columns(0).Width, myRow.Cells(0).Size.Width)
            Else
                Dim bmp As Bitmap = New Bitmap(124, 22)
                oResizedImage = bmp
                m_Image = oResizedImage ' oImage
                'dgvItems.Columns(0).Width = IIf(dgvItems.Columns(0).Width > myRow.Cells(0).Size.Width, dgvItems.Columns(0).Width, myRow.Cells(0).Size.Width)
                m_ImageWidth = m_Image.Width
                m_ImageHeight = m_Image.Height

            End If

            'Next myRow

        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '    Public Sub SaveDataRow(ByRef pDataRow As DataRow, ByRef lpos As Integer)
    Public Sub SaveDataRow(ByRef pDataRow As DataRow)

        Try
            m_oOrder.Ordhead.Save()

            pDataRow.Item("Ord_Type") = "O" ' m_oOrder.Ordhead.Ord_Type
            pDataRow.Item("Ord_No") = "" ' m_oOrder.Ordhead.Ord_No
            pDataRow.Item("Line_Seq_No") = Line_No
            pDataRow.Item("Item_No") = Item_No
            If m_Loc = "" Then
                pDataRow.Item("Loc") = m_oOrder.Ordhead.Mfg_Loc
            Else
                pDataRow.Item("Loc") = m_Loc
            End If
            pDataRow.Item("Pick_Seq") = "" ' Pick_Seq
            'pDataRow.Item("Cus_Item_No") = Cus_Item_No
            pDataRow.Item("Item_Desc_1") = Item_Desc_1
            pDataRow.Item("Item_Desc_2") = Item_Desc_2
            pDataRow.Item("Qty_Ordered") = Qty_Ordered
            pDataRow.Item("Qty_To_Ship") = Qty_To_Ship

            '===================================
            'Added June 6, 2017 - T. Louzon
            'ADDED TEMPORARILY - MAY be able to TAKE OUT LATER AFTER LOCK UP THE FIELDS
            pDataRow.Item("Unit_Price_BeforeSpecial") = Unit_Price_BeforeSpecial
            '==============================
            'Added June 27, 2017 - T. Louzon
            m_RouteExtension = f_GetRouteExtensionAndData(m_RouteID)
            If m_RouteExtension = "" Then m_RouteExtension = "XX"          'Added Sept 13, 2017 - T. Louzon
            pDataRow.Item("RouteExtension") = m_RouteExtension
            If m_RouteExtension = "XX" Then
                Mfg_Item_No = m_Item_No
            Else
                Mfg_Item_No = m_Item_No & "-" & m_RouteExtension
            End If
            pDataRow.Item("Mfg_Item_No") = m_Mfg_Item_No
            '==============================
            'Added June 29, 2017 - T. Louzon
            pDataRow.Item("IsNoStock") = m_IsNoStock
            pDataRow.Item("IsNoArtwork") = m_IsNoArtwork
            '==============================


            pDataRow.Item("Unit_Price") = Unit_Price
            pDataRow.Item("Discount_Pct") = Discount_Pct
            If Request_Dt.Year <> 1 Then pDataRow.Item("Request_Dt") = Request_Dt
            pDataRow.Item("Qty_Bkord") = Qty_Ordered - Qty_To_Ship ' Qty_Prev_Bkord
            'pDataRow.Item("Qty_Return_To_Stk") = Qty_Return_To_Stk
            pDataRow.Item("Bkord_Fg") = Bkord_Fg
            pDataRow.Item("Uom") = Uom
            'pDataRow.Item("Uom_Ratio") = Uom_Ratio
            pDataRow.Item("Unit_Cost") = Unit_Cost
            pDataRow.Item("Unit_Weight") = Unit_Weight
            'pDataRow.Item("Comm_Calc_Type") = Comm_Calc_Type
            'pDataRow.Item("Comm_Pct_Or_Amt") = Comm_Pct_Or_Amt
            If Promise_Dt.Year <> 1 Then pDataRow.Item("Promise_Dt") = Promise_Dt
            'pDataRow.Item("Tax_Fg") = Tax_Fg
            pDataRow.Item("Stocked_Fg") = Stocked_Fg
            'pDataRow.Item("Controlled_Fg") = Controlled_Fg
            pDataRow.Item("Select_Cd") = "" ' Select_Cd
            'pDataRow.Item("Tot_Qty_Ordered") = Tot_Qty_Ordered
            'pDataRow.Item("Tot_Qty_Shipped") = Tot_Qty_Shipped
            'pDataRow.Item("Tax_Fg_1") = Tax_Fg_1
            'pDataRow.Item("Tax_Fg_2") = Tax_Fg_2
            'pDataRow.Item("Tax_Fg_3") = Tax_Fg_3
            pDataRow.Item("Orig_Price") = Orig_Price
            'pDataRow.Item("Copy_To_Bm_Fg") = Copy_To_Bm_Fg
            'pDataRow.Item("Explode_Kit") = Explode_Kit
            'pDataRow.Item("Mfg_Ord_No") = Mfg_Ord_No
            'pDataRow.Item("Allocate_Dt") = Allocate_Dt
            'pDataRow.Item("Last_Post_Dt") = Last_Post_Dt
            'pDataRow.Item("Post_To_Inv_Qty") = Post_To_Inv_Qty
            'pDataRow.Item("Posted_To_Inv") = Posted_To_Inv
            'pDataRow.Item("Tot_Qty_Posted") = Tot_Qty_Posted
            'pDataRow.Item("Qty_Allocated") = Qty_Allocated
            'pDataRow.Item("Components_Alloc") = Components_Alloc
            'pDataRow.Item("Bin_Fg") = Bin_Fg
            'pDataRow.Item("Cost_Meth") = Cost_Meth
            'pDataRow.Item("Ser_Lot_Cd") = Ser_Lot_Cd
            'pDataRow.Item("Mult_Ftr_Fg") = Mult_Ftr_Fg
            'pDataRow.Item("Line_Type") = Line_Type
            pDataRow.Item("Prod_Cat") = Prod_Cat
            pDataRow.Item("End_Item_Cd") = End_Item_Cd
            'pDataRow.Item("Reason_Cd") = Reason_Cd
            'pDataRow.Item("Feature_Return") = Feature_Return
            'pDataRow.Item("Rec_Inspection") = Rec_Inspection
            'pDataRow.Item("Ship_From_Stk") = Ship_From_Stk
            'pDataRow.Item("Mult_Release") = Mult_Release
            If Req_Ship_Dt.Year <> 1 Then pDataRow.Item("Req_Ship_Dt") = Req_Ship_Dt
            'pDataRow.Item("Qty_From_Stk") = Qty_From_Stk
            pDataRow.Item("User_Def_Fld_1") = User_Def_Fld_1
            pDataRow.Item("User_Def_Fld_2") = User_Def_Fld_2
            pDataRow.Item("User_Def_Fld_3") = User_Def_Fld_3
            pDataRow.Item("User_Def_Fld_4") = User_Def_Fld_4
            pDataRow.Item("User_Def_Fld_5") = User_Def_Fld_5
            'pDataRow.Item("Picked_Dt") = Picked_Dt
            'pDataRow.Item("Shipped_Dt") = Shipped_Dt
            'pDataRow.Item("Billed_Dt") = Billed_Dt
            'pDataRow.Item("Update_Fg") = Update_Fg
            pDataRow.Item("Prc_Cd_Orig_Price") = Prc_Cd_Orig_Price
            'pDataRow.Item("Tax_Sched") = Tax_Sched
            'pDataRow.Item("Cus_No") = Cus_No
            'pDataRow.Item("Tax_Amt") = Tax_Amt
            pDataRow.Item("Qty_Bkord_Fg") = "" ' Qty_Bkord_Fg
            pDataRow.Item("Line_No") = Line_No
            'pDataRow.Item("Mfg_Method") = Mfg_Method
            'pDataRow.Item("Forced_Demand") = Forced_Demand
            'pDataRow.Item("Conf_Pick_Dt") = Conf_Pick_Dt
            'pDataRow.Item("Item_Release_No") = Item_Release_No
            'pDataRow.Item("Bin_Ser_Lot_Comp") = Bin_Ser_Lot_Comp
            'pDataRow.Item("Offset_Used_Fg") = Offset_Used_Fg
            'pDataRow.Item("Ecs_Space") = Ecs_Space
            'pDataRow.Item("Sfc_Order_Status") = Sfc_Order_Status
            'pDataRow.Item("Total_Cost") = Total_Cost
            'pDataRow.Item("Po_Ord_No") = m_oOrder.Ordhead.Oe_Po_No '  Po_Ord_No
            pDataRow.Item("Rma_Seq") = 0 ' Rma_Seq
            'pDataRow.Item("Vendor_No") = Vendor_No
            'pDataRow.Item("Posted_Unit_Cost") = Posted_Unit_Cost
            pDataRow.Item("Extra_1") = Extra_1
            pDataRow.Item("Extra_2") = IIf(Extra_2.Trim = "", DBNull.Value, Extra_2) ' IIf(Extra_2_Bln, "Y", DBNull.Value)
            'pDataRow.Item("Extra_3") = Extra_3
            'pDataRow.Item("Extra_4") = Extra_4
            'pDataRow.Item("Extra_5") = Extra_5
            pDataRow.Item("Extra_6") = IIf(Extra_6.Trim = "", DBNull.Value, Extra_6)
            'pDataRow.Item("Extra_7") = Extra_7
            'pDataRow.Item("Extra_8") = Extra_8
            pDataRow.Item("Extra_9") = Extra_9
            pDataRow.Item("Extra_10") = Extra_10
            pDataRow.Item("Extra_11") = Extra_11
            'pDataRow.Item("Extra_12") = Extra_12
            'pDataRow.Item("Extra_13") = Extra_13
            'pDataRow.Item("Extra_14") = Extra_14
            'pDataRow.Item("Extra_15") = Extra_15
            'pDataRow.Item("Warranty_Date") = Warranty_Date
            'pDataRow.Item("Revision_No") = Revision_No
            'pDataRow.Item("Cm_Post_Fg") = Cm_Post_Fg
            pDataRow.Item("Recalc_Sw") = "Y" ' Recalc_Sw
            'pDataRow.Item("Filler_0004") = Filler_0004
            'pDataRow.Item("Id") = Id ' Never save this field, it is set to identity
            pDataRow.Item("Ord_GUID") = m_oOrder.Ordhead.Ord_GUID
            pDataRow.Item("Item_Guid") = m_Item_Guid
            pDataRow.Item("RouteID") = m_RouteID ' 0 ' Route
            'Debug.Print(m_Item_No & " " & m_RouteID.ToString)
            pDataRow.Item("ProductProof") = ProductProof
            pDataRow.Item("AutoCompleteReship") = AutoCompleteReship

            pDataRow.Item("OEOrdLin_ID") = OEOrdLin_ID
            pDataRow.Item("Item_Cd") = Item_Cd
            pDataRow.Item("Mix_Group") = Mix_Group

        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Delete_Extra_Functions()

        Dim db As New cDBA

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            Dim strSqlOrdLin As String

            strSqlOrdLin = "SELECT TOP 1 * FROM OEI_Extra_Functions WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND Item_Guid = '" & m_Item_Guid & "' "
            strSqlOrdLin = "DELETE FROM OEI_Extra_Functions WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND Item_Guid = '" & m_Item_Guid & "' "

            db.Execute(strSqlOrdLin)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Delete_Traveler_Header()

        Dim db As New cDBA

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            Dim strSqlOrdLin As String

            strSqlOrdLin = "DELETE FROM OEI_TRAVELER_HEADER WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND Item_Guid = '" & m_Item_Guid & "' "

            db.Execute(strSqlOrdLin)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Delete_Traveler_Details()

        Dim db As New cDBA

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            Dim strSqlOrdLin As String

            strSqlOrdLin = "DELETE FROM OEI_TRAVELER_DETAILS WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND Item_Guid = '" & m_Item_Guid & "' "

            db.Execute(strSqlOrdLin)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Delete_Traveler_Document()

        Dim db As New cDBA

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            Dim strSqlOrdLin As String

            strSqlOrdLin = "DELETE FROM OEI_ORDLIN WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND Item_Guid = '" & m_Item_Guid & "' "

            db.Execute(strSqlOrdLin)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Delete_Traveler_Extra_Fields()

        Dim DB As New cDBA

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            Dim strSqlOrdLin As String

            strSqlOrdLin = "SELECT TOP 1 * FROM OEI_EXTRA_FIELDS WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND Item_Guid = '" & m_Item_Guid & "' "

            DB.Execute(strSqlOrdLin)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Delete()

        Dim db As New cDBA
        Dim strSql As String

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            strSql = _
            "DELETE FROM OEI_EXTRA_FIELDS " & _
            "WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND " & _
            "Item_Guid IN (SELECT Comp_Item_Guid FROM OEI_ORDBLD WITH (NOLOCK) WHERE Kit_Item_Guid = '" & Item_Guid & "') "

            db.Execute(strSql)

            strSql = _
            "DELETE FROM OEI_EXTRA_FUNCTIONS " & _
            "WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND " & _
            "Item_Guid IN (SELECT Comp_Item_Guid FROM OEI_ORDBLD WITH (NOLOCK) WHERE Kit_Item_Guid = '" & Item_Guid & "') "

            db.Execute(strSql)

            strSql = _
            "DELETE FROM OEI_ORDLIN " & _
            "WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND " & _
            "Item_Guid IN (SELECT Comp_Item_Guid FROM OEI_ORDBLD WITH (NOLOCK) WHERE Kit_Item_Guid = '" & Item_Guid & "') "

            db.Execute(strSql)

            strSql = _
            "INSERT INTO OEI_ORDBLD_AUDIT " & _
            "SELECT * FROM OEI_ORDBLD WITH (NOLOCK) " & _
            "WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND " & _
            "Item_Guid IN (SELECT Comp_Item_Guid FROM OEI_ORDBLD WITH (NOLOCK) WHERE Kit_Item_Guid = '" & Item_Guid & "') "

            db.Execute(strSql)

            strSql = _
            "DELETE FROM OEI_ORDBLD " & _
            "WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND " & _
            "Kit_Item_Guid = '" & Item_Guid & "' "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Reset()

        ' Flags
        blnFirstLoad = False
        blnLoading = False

        ' Flags to activate creation of different components of the item
        m_Create_Traveler = False
        m_Create_Imprint = False
        m_Create_Kit = False

        ' Other flags
        m_SaveToDB = False
        m_Dirty = False
        'm_No_Kits = False

        ' Item guid is the current item, Kit_Item_Guid is the parent of a kit component
        'm_Item_Guid = string.empty
        'm_Kit_Item_Guid = string.empty

        m_Record_Type = "L"

        m_Ord_Type = String.Empty       'Char (8)
        m_Ord_No = String.Empty         'Char (8)       Order No
        m_Line_Seq_No = 0               'Int            Line Sequence No
        m_Item_No = String.Empty        'Char (15)      ' Sub Item_No_Changed()

        m_Loc = String.Empty            'Char (3)       ' Sub Loc_Changed()

        m_Pick_Seq = String.Empty       'Char (8)
        m_Cus_Item_No = String.Empty    'Char (30)
        m_Item_Desc_1 = String.Empty    'Char (30)      Item Description Line 1
        m_Item_Desc_2 = String.Empty    'Char (30)      Item Description Line 2
        m_Qty_Ordered = 0               'Decimal (13,4) 
        m_Qty_To_Ship = 0               'Decimal (13,4) 
        m_Qty_Lines = 1                 'Decimal (13,4) 
        m_Qty_To_Ship_Calc = 0
        m_Qty_Prev_To_Ship = 0          'Decimal (13,4)
        m_Unit_Price = 0                'Decimal (16,6) Unit Price  If not populated, default 'Price' from IM Loca

        '==============================
        'Added June 6, 2017 - T. Louzon
        m_Unit_Price_BeforeSpecial = 0
        '==============================
        'Added July 17, 2017 - T. Louzon
        m_Mfg_Item_No = ""
        m_RouteExtension = ""
        '==============================



        m_Discount_Pct = 0              'Numeric NULL
        m_Request_Dt = Now()            'DateTime NULL
        m_Qty_Prev_To_Ship = 0          'Decimal (13,4) ' Sub Qty_To_Ship_Changed() 
        m_Qty_Prev_Bkord = 0            'Numeric NULL
        m_Qty_Return_To_Stk = 0         'Numeric NULL
        m_Bkord_Fg = String.Empty       'Char(1) NULL ' BkOrd_Fg
        m_Uom = String.Empty            'Char(2) NULL
        m_Uom_Ratio = 0                 'Numeric NULL
        m_Unit_Cost = 0                 'Numeric NULL
        m_Unit_Weight = 0               'Numeric NULL
        m_Comm_Calc_Type = String.Empty 'Char(1) NULL ' Recalc ??
        m_Comm_Pct_Or_Amt = 0           'Numeric NULL
        m_Promise_Dt = Now()            'DateTime NULL
        m_Tax_Fg = String.Empty         'Char(1) NULL
        m_Stocked_Fg = String.Empty     'Char(1) NULL
        m_Controlled_Fg = String.Empty  'Char(1) NULL
        m_Select_Cd = String.Empty      'Char(1) NOT NULL ' Selected
        m_Tot_Qty_Ordered = 0           'Numeric NULL
        m_Tot_Qty_Shipped = 0           'Numeric NULL
        m_Tax_Fg_1 = String.Empty       'Char(1) NULL
        m_Tax_Fg_2 = String.Empty       'Char(1) NULL
        m_Tax_Fg_3 = String.Empty       'Char(1) NULL
        m_Orig_Price = 0                'Numeric NULL
        m_Copy_To_Bm_Fg = String.Empty  'Char(1) NULL
        m_Explode_Kit = String.Empty    'Char(1) NULL
        m_Mfg_Ord_No = String.Empty     'Char(8) NULL
        m_Allocate_Dt = Now()           'DateTime NULL
        m_Last_Post_Dt = Now()          'DateTime NULL
        m_Post_To_Inv_Qty = 0           'Numeric NULL
        m_Posted_To_Inv = 0             'Numeric NULL
        m_Tot_Qty_Posted = 0            'Numeric NULL
        m_Qty_Allocated = 0             'Numeric NULL
        m_Components_Alloc = 0          'Numeric NULL
        m_Bin_Fg = String.Empty         'Char(1) NULL
        m_Cost_Meth = String.Empty      'Char(1) NULL
        m_Ser_Lot_Cd = String.Empty     'Char(1) NULL
        m_Mult_Ftr_Fg = String.Empty    'Char(1) NULL
        m_Line_Type = String.Empty      'Char(1) NULL
        m_Prod_Cat = String.Empty       'Char(3) NULL
        m_End_Item_Cd = String.Empty    'Char(1) NULL
        m_Reason_Cd = String.Empty      'Char(3) NULL
        m_Feature_Return = String.Empty 'Char(1) NULL
        m_Rec_Inspection = String.Empty 'Char(1) NULL
        m_Ship_From_Stk = String.Empty  'Char(1) NULL
        m_Mult_Release = String.Empty   'Char(1) NULL
        m_Req_Ship_Dt = Now()           'DateTime NULL
        m_Qty_From_Stk = 0              'Numeric NULL
        m_User_Def_Fld_1 = String.Empty 'Char(30) NULL
        m_User_Def_Fld_2 = String.Empty 'Char(30) NULL
        m_User_Def_Fld_3 = String.Empty 'Char(30) NULL
        m_User_Def_Fld_4 = String.Empty 'Char(30) NULL
        m_User_Def_Fld_5 = String.Empty 'Char(30) NULL
        m_Picked_Dt = Now()             'DateTime NULL
        m_Shipped_Dt = Now()            'DateTime NULL
        m_Billed_Dt = Now()             'DateTime NULL
        m_Update_Fg = String.Empty      'Char(1) NULL
        m_Prc_Cd_Orig_Price = 0         'Numeric NULL
        m_Tax_Sched = String.Empty      'Char(5) NULL
        m_Cus_No = String.Empty         'Char(20) NULL
        m_Tax_Amt = 0                   'Numeric NULL
        m_Qty_Prev_Bkord_Fg = String.Empty 'Char(1) NOT NULL
        m_Line_No = 1                   'SmallInt NOT NULL
        m_Mfg_Method = String.Empty     'Char(2) NULL
        m_Forced_Demand = String.Empty  'Char(1) NULL
        m_Conf_Pick_Dt = Now()          'DateTime NULL
        m_Item_Release_No = String.Empty 'Char(8) NULL
        m_Bin_Ser_Lot_Comp = String.Empty 'Char(1) NULL
        m_Offset_Used_Fg = String.Empty 'Char(1) NULL
        m_Ecs_Space = String.Empty      'Char(20) NULL
        m_Sfc_Order_Status = String.Empty 'Char(1) NULL
        m_Total_Cost = 0                'Numeric NULL
        m_Po_Ord_No = String.Empty      'Char(8) NULL
        m_Rma_Seq = 0                   'SmallInt NOT NULL
        m_Vendor_No = String.Empty      'Char(20) NULL
        m_Posted_Unit_Cost = 0          'Numeric NULL
        m_Extra_1 = String.Empty        'Char(1) NULL
        m_Extra_2 = String.Empty        'Char(1) NULL
        m_Extra_3 = String.Empty        'Char(1) NULL
        m_Extra_4 = String.Empty        'Char(1) NULL
        m_Extra_5 = String.Empty        'Char(1) NULL
        m_Extra_6 = String.Empty        'Char(8) NULL
        m_Extra_7 = String.Empty        'Char(8) NULL
        m_Extra_8 = String.Empty        'Char(12) NULL
        m_Extra_9 = String.Empty        'Char(12) NULL
        m_Extra_10 = 0                  'Numeric NULL
        m_Extra_11 = 0                  'Numeric NULL
        m_Extra_12 = 0                  'Numeric NULL
        m_Extra_13 = 0                  'Numeric NULL
        m_Extra_14 = 0                  'Int NULL
        m_Extra_15 = 0                  'Int NULL
        m_Warranty_Date = Now()         'DateTime NULL
        m_Revision_No = String.Empty    'Char(8) NULL
        m_Cm_Post_Fg = String.Empty     'Char(1) NULL
        m_Recalc_Sw = String.Empty      'Char(1) NOT NULL Recalc
        m_Filler_0004 = String.Empty    'Char(132) NULL
        m_Id = 0                        'Numeric NOT NULL
        'm_Ord_Guid As String
        m_Image = Nothing
        m_ImageWidth = 0
        m_ImageHeight = 0

        ' Additionnal fields 
        'Additional fields:
        m_Qty_Inventory = 0
        m_Qty_Allocated = 0
        m_Qty_On_Hand = 0               'Numeric NOT NULL
        m_Calc_Price = 0

        m_Route = String.Empty
        m_ProductProof = 0
        m_AutoCompleteReship = 0

        m_Activity_Cd = String.Empty

        m_Imprint = New cImprint()
        m_Traveler = New CTraveler()
        m_Kit = New cKit(Me)

        m_Kit_Component = False

        m_Source = OEI_SourceEnum.None

        m_OEI_H_Pixels = 400
        m_OEI_W_Pixels = 400
        m_OEI_Image_Rotate = ""

        m_Comp_Item_Guid = ""

    End Sub

    Public Sub Save()

        Dim db As New cDBA
        Dim dt As DataTable
        Dim drRow As DataRow

        If m_Item_No Is Nothing Or m_Item_No.Equals(DBNull.Value) Then Exit Sub

        If Not (m_SaveToDB) Or (m_Source = OEI_SourceEnum.frmProductLineEntry) Then Exit Sub

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            If Line_No = 0 Then Exit Sub

            Dim strSql As String = "SELECT TOP 1 * FROM OEI_ORDLIN WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND Item_Guid = '" & Item_Guid & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count = 0 Then
                drRow = dt.NewRow()
            Else
                drRow = dt.Rows(0)
            End If

            Call SaveDataRow(drRow)

            If dt.Rows.Count = 0 Then
                db.DBDataTable.Rows.Add(drRow)
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            Else
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            End If

            db.DBDataAdapter.Update(db.DBDataTable)

        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function FieldsRequired() As Boolean

        FieldsRequired = False

        Try
            If Trim(m_Item_No) = "" Then Throw New Exception()
            'If Trim(m_Loc) = "" Then Throw New Exception()

            FieldsRequired = True

        Catch er As Exception
            ' Do nothing, everything that gets here is when value is nothing or ""
        End Try

    End Function

    Private Shared Sub showAppendMessageLog(key As String, msgText As String, Optional ByVal forceShowMessage As Boolean = False)
        Dim isShownAlready = False

        Try
            'check if message was already shown
            If popMessageLog.ContainsKey(key) AndAlso popMessageLog(key).Contains(msgText.Trim()) Then
                isShownAlready = True
            End If

            'update and show message
            If Not isShownAlready Or forceShowMessage = True Then
                If popMessageLog.ContainsKey(key) Then
                    popMessageLog(key) = popMessageLog(key) & "\n" & msgText.Trim()
                Else
                    popMessageLog.Add(key, msgText)
                End If


                MsgBox(Replace(msgText, "\n", vbCrLf))
            End If
        Catch er As Exception
            MsgBox("That was weird" & er.Message)
        End Try
    End Sub

    Private Sub SetUnit_Price(ByVal plQty As Double)

        Dim dt As DataTable
        Dim db As New cDBA

        Dim strSQL As String

        Try
            ''**********************************************************************************************
            ''CHANGED SEPT 11, 2017 - Make Bags the same
            ''ADDED IF STATEMENT TO GET PRICE DIFFERENTLY FOR BAGS - For Bags, use the MFG_Item_No. For everything else use the Item_no   - July 17, 2017
            'If UCase(m_Prod_Cat) = "BAG" And m_RouteExtension <> "" And m_Mfg_Item_No <> "" Then
            '    strSQL = "SELECT DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', '" & m_Mfg_Item_No & "', '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", " & plQty & ") as Unit_Price "
            'Else
            '    'Find price the normal way, using the RAW Item Number (No Item Extension)
            '    strSQL = "SELECT DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', '" & m_Item_No & "', '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", " & plQty & ") as Unit_Price "
            'End If
            '**********************************************************************************************
            strSQL = "SELECT DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', '" & m_Item_No & "', '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", " & plQty & ") as Unit_Price "
            '**********************************************************************************************
            'g_oOrdline.Loc 
            If Trim(m_Loc.ToUpper()) = "3N" Then strSQL = " select dbo.OE_Item_Price_Type10 (10,'" & m_oOrder.Ordhead.Curr_Cd & "','" & m_Item_No & "')  as Unit_Price "



            dt = db.DataTable(strSQL)
            If dt.Rows.Count <> 0 Then
                m_Unit_Price = dt.Rows(0).Item("Unit_Price")
                '==============================
                'Added June 6, 2017 - T. Louzon
                m_Unit_Price_BeforeSpecial = m_Unit_Price
                '==============================
            End If
            'm_Orig_Price = m_Unit_Price

            dt.Dispose()

        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'perform availalbility simulation for a given line item
    Public Sub PerformProductionAvailalbilitySimulation(ByRef pCell As DataGridViewCell, ByRef poImprint As cImprint, Optional checkMultiDays As Boolean = False, Optional ByVal OsrEligibiltyList As Dictionary(Of String, Boolean) = Nothing, Optional ByVal kitUnits As Integer = Nothing)
        'temp disable
        ' Exit Sub
        Try
            Dim req_units = IIf(kitUnits = Nothing, m_Qty_Ordered, kitUnits)

            'do not compare if date is invalid
            If RouteID = 0 OrElse (Stocked_Fg = "N" And Item_No <> "00STOCK0000") OrElse Date.Compare(m_Promise_Dt.ToShortDateString, Date.Now().ToShortDateString) < 0 _
                OrElse Date.Compare(m_Promise_Dt.ToShortDateString, Date.Parse("1/1/2030")) >= 0 OrElse req_units Is Nothing OrElse req_units <= 0 _
                OrElse (
                    Date.Compare(m_Promise_Dt.ToShortDateString, lastThermoCheckedDate) = 0 And lastThermoRouteId = m_RouteID And lastThermoQty = m_Qty_To_Ship And checkMultiDays = False _
                    And (poImprint IsNot Nothing AndAlso lastThermo_Num_Impr_1 = poImprint.Num_Impr_1 AndAlso lastThermo_Num_Impr_2 = poImprint.Num_Impr_2 AndAlso lastThermo_Num_Impr_3 = poImprint.Num_Impr_3)
                ) Then

                Exit Sub
            End If

            If poImprint Is Nothing Then
                poImprint = New cImprint(m_Item_Guid)
            End If

            'update lastechked values
            If checkMultiDays = False Then
                lastThermoCheckedDate = Date.Parse(m_Promise_Dt.ToShortDateString)
            End If
            lastThermoRouteId = m_RouteID
            lastThermoQty = req_units
            lastThermo_Num_Impr_1 = poImprint.Num_Impr_1
            lastThermo_Num_Impr_2 = poImprint.Num_Impr_2
            lastThermo_Num_Impr_3 = poImprint.Num_Impr_3

            Dim bestDate As Date = m_Promise_Dt.ToShortDateString
            Dim oriDate As Date = m_Promise_Dt.ToShortDateString

            Dim validitySet As ArrayList = New ArrayList
            Dim detailsSet As ArrayList = New ArrayList
            Dim isValid As Boolean = False

            Dim isDateChanged As Boolean = False

            'if rocket service orders, you can check up until red threshold while regular orders stop at the yellow threshold
            Dim ignoreIntermediateThreshold = IIf(m_oOrder.Ordhead.User_Def_Fld_2 = "RKC" Or m_oOrder.Ordhead.User_Def_Fld_2 = "RKU", True, False)

            'Get all imprint methods pertaining to the route
            Dim impMethods As ArrayList = mThermometer.getAllApplicableRoutes(RouteID.ToString)

            'Max loop count (TODO: check if necessary based on test results)
            Dim loopCount = IIf(impMethods.Count > 1 And checkMultiDays = True, 5, 1)

            Dim cptimprint As Int32 = 0
            Dim qty_units As Int32 = 0

            Do While (isValid = False And checkMultiDays = True And isDateChanged = True And impMethods.Count > 1) Or (isValid = False And loopCount > 0)
                isDateChanged = False
                loopCount = loopCount - 1

                'check each imprint method for either a single day or multi day logic
                For Each impMethod As String In impMethods

                    '++ID 12.17.2019 multiply by imprint number the quantity for the first three decorating
                    Select Case cptimprint
                        Case 0
                            qty_units = req_units * poImprint.Num_Impr_1
                        Case 1
                            qty_units = req_units * poImprint.Num_Impr_2
                        Case 2
                            qty_units = req_units * poImprint.Num_Impr_3
                        Case Else
                            qty_units = req_units
                    End Select


                    'if backordered quantity exceeds
                    'If m_Qty_Prev_Bkord > 0 AndAlso m_Bkord_Fg = "Y" AndAlso (m_Qty_Prev_Bkord + m_Qty_To_Ship) > mThermometer.getAllowableOrderUnitSizeByImprint(impMethod, 3, True) Then
                    '    Exit Sub
                    'End If

                    Dim defaultMachine As String = mThermometer.getDefaultMachine(Item_No, impMethod)
                    Dim timeSpecs() As Decimal = mThermometer.getTimeSpecs(m_Item_No, impMethod)

                    Dim osrEligible As Boolean = False

                    If OsrEligibiltyList IsNot Nothing AndAlso OsrEligibiltyList.ContainsKey(impMethod) Then
                        osrEligible = OsrEligibiltyList(impMethod)
                    End If

                    '  Dim resultPieces() = mThermometer.simulateNextAvailDateAndMachine(impMethod, req_units, timeSpecs, defaultMachine, ignoreIntermediateThreshold, bestDate, checkMultiDays, osrEligible).Split("_")

                    '++ID 12.17.2019 commented above and added same line except req_units changed to qty_units, becasue req_units multiplied with imprint count fro simulate correct qty of the items
                    Dim resultPieces() = mThermometer.simulateNextAvailDateAndMachine(impMethod, qty_units, timeSpecs, defaultMachine, ignoreIntermediateThreshold, bestDate, checkMultiDays, osrEligible).Split("_")


                    Dim imprintState = False

                    validitySet.Add(IIf(resultPieces(0).ToString = Date.MinValue, False, True))
                    detailsSet.Add(resultPieces(1).ToString)

                    If resultPieces(0).ToString = Date.MinValue Then
                        Exit For
                    ElseIf Date.Compare(Date.Parse(resultPieces(0).ToString).ToShortDateString, bestDate) > 0 Then
                        bestDate = Date.Parse(resultPieces(0).ToString).ToShortDateString
                        isDateChanged = True
                    End If
                Next
                ' MsgBox(DateTime.Parse("12:00:00 AM"))
                'check that all dates are valid
                For Each validity As Boolean In validitySet
                    isValid = validity
                    If isValid = False Then
                        Exit For
                    End If
                Next

                cptimprint += 1

            Loop

            'Populate results to the audit results the audit table
            Call appendThermoResultsToAudit(validitySet, detailsSet, bestDate.ToShortDateString, Integer.Parse(m_Qty_Ordered.ToString))

            'set validity and ask if comprehensive search desired
            If isValid = True Then
                m_Promise_Dt = bestDate.ToShortDateString
                pCell.Value = bestDate.ToShortDateString
                pCell.Style.BackColor = Color.LightGreen

                If checkMultiDays = True Then
                    lastThermoCheckedDate = Date.Parse(m_Promise_Dt.ToShortDateString)
                    MsgBox("Thermo simulation has suggested: " & bestDate.ToShortDateString & " as an acceptable production date")
                    ' Else
                    '    MsgBox("Thermo simulation has determined that: " & bestDate.ToShortDateString & " is an acceptable production date")
                End If
            Else
                pCell.Style.BackColor = Color.Crimson
                frmOrder.cmdRequest_Dt.PerformClick()

                If checkMultiDays = False Then
                    'ask if compreshensive search wanted
                    Dim multiRequest = MsgBox("We do not recommend the promise day for item " & Item_No & "." & vbCrLf & _
                                              "Do you want the system to try and find the next best date?", MsgBoxStyle.YesNoCancel)
                    If multiRequest = MsgBoxResult.Yes Then
                        'bring up promise date (up to two days) if possible (TODO: consider making the -2 days value come from SQL table config)

                        'determine if we can look for an earlier date
                        Dim daysdiff = mThermometer.daysGap(Date.Parse(mThermometer.dateAdjuster(m_Promise_Dt.ToShortDateString, -2)))
                        If daysdiff >= 4 Then
                            m_Promise_Dt = mThermometer.dateAdjuster(m_Promise_Dt.ToShortDateString, -2)
                        ElseIf daysdiff >= 3 Then
                            m_Promise_Dt = mThermometer.dateAdjuster(m_Promise_Dt.ToShortDateString, -1)
                        End If

                        Call PerformProductionAvailalbilitySimulation(pCell, poImprint, True, OsrEligibiltyList, kitUnits)
                    Else 'Ask if they want to keep date (current day check only)
                        multiRequest = MsgBox("Are you keeping the current promise date of " & m_Promise_Dt _
                                              & "for item " & Item_No & "?", MsgBoxStyle.YesNoCancel)
                        If multiRequest = MsgBoxResult.Yes Then
                            Dim fullDetailsSet = parseSimDetails(detailsSet)
                            Call sendProductionOverrideEmail(fullDetailsSet)
                        End If
                    End If

                Else 'Ask if they want to keep date (Full simulation check)
                    If lastThermoCheckedDate <> Nothing Then
                        m_Promise_Dt = lastThermoCheckedDate
                    End If

                    Dim failReminder = MsgBox("Full simulation could still not find a suitable date. Are you keeping the" _
                                             & " current promise date of " & m_Promise_Dt & " for item " & Item_No & "?", MsgBoxStyle.YesNoCancel)
                    If failReminder = MsgBoxResult.Yes And m_oOrder.Ordhead.User_Def_Fld_2.ToString <> "RKC" And m_oOrder.Ordhead.User_Def_Fld_2.ToString <> "RKU" Then
                        Dim fullDetailsSet = parseSimDetails(detailsSet)
                        Call sendProductionOverrideEmail(fullDetailsSet)
                    End If
                End If
            End If
        Catch er As Exception
            'MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub appendThermoResultsToAudit(ByVal validitySet As ArrayList, ByVal detailsSet As ArrayList, ByVal requestDt As Date, ByVal requestQty As Integer)
        For i As Integer = 0 To validitySet.Count - 1
            Dim detailPieces() = detailsSet.Item(i).ToString.Split(vbCrLf)
            Dim machineNo As String = ""
            Dim imp_method As String = ""

            For Each detail As String In detailPieces
                If detail.Contains("Machine") = True Then
                    Dim keyValPair() = detail.ToString.Split(": ")
                    machineNo = keyValPair(1).ToString.ToUpper.Trim()

                ElseIf detail.Contains("Imprint") = True Then
                    Dim keyValPair() = detail.ToString.Split(": ")
                    imp_method = keyValPair(1).ToString.Trim()
                ElseIf machineNo <> "" And imp_method <> "" Then
                    Exit For
                End If
            Next

            mThermometer.appendToAudit(machineNo, validitySet.Item(i), requestDt, m_oOrder.Ordhead.OEI_Ord_No.Trim(), Item_Guid, detailsSet.Item(i).ToString.Trim(), "OEI", Item_No, m_RouteID, imp_method, requestQty)
        Next
    End Sub

    Private Function parseSimDetails(detailsSet As ArrayList) As String
        Dim fullDetailsSet As String = ""
        For Each Details As String In detailsSet
            fullDetailsSet = fullDetailsSet & Details.ToString.Trim() & vbCrLf & vbCrLf
        Next
        Return fullDetailsSet
    End Function

    Private Function determineOSREligibility(ByVal impMethod As String) As Boolean
        Dim is_osr_eligible As Boolean = False
        Dim maxUnits = getAllowableOrderUnitSizeByImprint(impMethod, 3, True)
        'Dim impMethods As ArrayList = mThermometer.getAllApplicableRoutes(RouteID.ToString)

        If (maxUnits > 0 AndAlso m_Qty_Ordered <= maxUnits) OrElse maxUnits = 0 Then
            is_osr_eligible = True
        End If

        Return is_osr_eligible
    End Function

    Private Sub sendProductionOverrideEmail(ByVal simDetails As String)
        Dim fromAddr As String = "Spectorandco OEI Service <it@spectorandco.ca>"
        Dim subject = "OEI promise date overflow override from user " & Environment.UserName & " for item: " & Item_No & " and route: " & Route

        Dim addresses(2) As String

        addresses(0) = "irene@spectorandco.ca"
        addresses(1) = "pinky@spectorandco.ca"
        addresses(2) = "leni@spectorandco.com" '++ID added 01262024 ticket 37559 from Irene 
        addresses(3) = "iond@spectorandco.ca"

        Dim body = "User has ignored the system warning of a promise date not being recommended." & vbCrLf & vbCrLf _
                   & "User: <b>" & Environment.UserName & "</b>" & vbCrLf _
                   & "OEI Order Number: <b>" & m_oOrder.Ordhead.OEI_Ord_No & " </b>" & vbCrLf _
                   & "OE PO Number: <b>" & m_oOrder.Ordhead.Oe_Po_No & " </b>" & vbCrLf _
                   & "OEI Order GUID: <b>" & m_oOrder.Ordhead.Ord_GUID & " </b>" & vbCrLf _
                   & "OEI Order Line GUID: <b>" & Ord_Guid & " </b>" & vbCrLf _
                   & "Item Number: <b>" & Item_No & "</b> " & vbCrLf _
                   & "Item Qty: <b>" & m_Qty_To_Ship & "</b> " & vbCrLf _
                   & "Route:  <b>" & Route.ToString & "</b>" & vbCrLf _
                   & "Override Date: <b>" & m_Promise_Dt.ToShortDateString & "</b>" & vbCrLf & vbCrLf _
                   & "<b>Simulation Details: </b>" & vbCrLf _
                   & simDetails.ToString() _
                   & vbCrLf & vbCrLf _
                   & "*If you notice 0 hours, it either means that the sum of decorating group is not green or that the machine checked is offline."

        'turn vb breaklines into HTML breaklines
        body = Replace(body, vbCrLf, "<BR>")

        'send email
        mEmailer.SendGenericEmail(fromAddr, addresses, subject, body)

    End Sub

    'old thermo version, please remove when done and all delegate references
    Public Sub SetProductionAvailabilityColor(ByRef pCell As DataGridViewCell)

        Try

            Dim strSql As String
            Dim db As New cDBA
            Dim dtRouteCat As DataTable
            Dim dtAvailDate As DataTable
            Dim dtRow As DataRow
            Dim intThermoColor As Integer = 0

            Dim routeEl As String ' Object

            strSql = "SELECT ISNULL(RouteCategory, '') AS RouteCategory FROM EXACT_TRAVELER_ROUTE WHERE RouteID = " & RouteID.ToString

            dtRouteCat = db.DataTable(strSql)
            If dtRouteCat.Rows.Count <> 0 Then

                Dim arrRoutes() As String
                arrRoutes = dtRouteCat.Rows(0).Item("RouteCategory").ToString.Split("/")

                For Each routeEl In arrRoutes

                    If Not routeEl.Equals(DBNull.Value) Then

                        strSql = _
                        "SELECT TOP 1 Color " & _
                        "FROM   EXACT_TRAVELER_THERMO_AVAIL_DATE " & _
                        "WHERE  Method_Name = '" & routeEl & "' AND " & _
                        "       Next_Avail_Date = '" & m_Promise_Dt & "' "

                        dtAvailDate = db.DataTable(strSql)
                        '0-Green,1-Orange,2-Yellow,3-Blue,4-Red
                        If dtAvailDate.Rows.Count <> 0 Then
                            dtRow = dtAvailDate.Rows(0)
                            If dtRow.Item("Color") > intThermoColor Then intThermoColor = dtRow.Item("Color")

                        End If

                    End If

                Next

            End If

            Select Case intThermoColor
                Case 0
                    pCell.Style.BackColor = Color.Empty
                Case 1
                    pCell.Style.BackColor = Color.Empty ' Color.Orange
                Case 2
                    pCell.Style.BackColor = Color.Empty ' Color.Yellow
                Case 3
                    pCell.Style.BackColor = Color.Empty ' Color.Cyan
                Case 4
                    pCell.Style.BackColor = Color.Red ' Color.Magenta
                Case Else
                    pCell.Style.BackColor = Color.Empty

            End Select

        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Sub Calc_Overpick()

    '    Dim db As New cDBA
    '    Dim dtCustomer As DataTable
    '    Dim dtItem As DataTable

    '    Dim strSql As String

    '    Dim Cus_No As String

    '    strSql = "SELECT * FROM OEI_OrdHdr WITH (Nolock) WHERE ORD_Guid = '" & m_Ord_Guid & "' "
    '    dt = db.DataTable(strSql)

    '    If dt.Rows.Count <> 0 Then
    '        Cus_No = dt.Rows(0).Item("Cus_No").ToString.Trim

    '    End If

    'End Sub

    Public Shared Sub customerInstructions(ByVal datafield As String, Optional ByVal orderType As String = "")
        Dim db As New cDBA
        Dim dt As DataTable

        If datafield Is Nothing Then Exit Sub
        Dim strUnion As String = ""
        Dim subStrsql As String = ""
        If orderType <> "" Then

            subStrsql = " AND ORDER_TYPE = '" & orderType & "' "


            '++ID 8.14.2018
            If Trim(m_oOrder.Ordhead.Customer.textfield3.Replace("'", "''")) <> String.Empty Then
                strUnion = "   UNION " &
           " SELECT  oU.DATA_FIELD, ISNULL(oU.SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, oU.SHOW_MESSAGEBOX, oU.ORDER_TYPE " &
           " FROM  OEI_ITEM_SPEC_INSTRUCTION as oU WITH (NOLOCK) " &
           " WHERE oU.DATA_FIELD = '" & datafield & "' AND oU.SHOW_MESSAGEBOX = 1 " & subStrsql &
           " AND ISNULL(oU.CUS_GROUP,'') = '" & Trim(m_oOrder.Ordhead.Customer.textfield3.Replace("'", "''")) & "'  "
            End If

        End If

        Dim strsqlMain As String =
           "SELECT     o.DATA_FIELD, ISNULL(o.SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, o.SHOW_MESSAGEBOX, o.ORDER_TYPE " &
           "FROM       OEI_ITEM_SPEC_INSTRUCTION as o WITH (NOLOCK) " &
           "WHERE      o.CUS_NO = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "' AND " &
           "           DATA_FIELD = '" & datafield & "' AND SHOW_MESSAGEBOX = 1 " & subStrsql &
         strUnion '&
        '  "ORDER BY   ITEM_NO, DATA_FIELD "

        dt = db.DataTable(strsqlMain)
        If dt.Rows.Count <> 0 Then

            For Each dtRow As DataRow In dt.Rows
                'show popup
                showAppendMessageLog(datafield, dtRow.Item("SPEC_INSTRUCTION").ToString)
            Next

        End If
    End Sub

    '++ID function adeded 3.25.2020
    Public Shared Sub customerInstructions1(ByVal orderType As String, Optional ByRef _mess As String = "0")
        Dim db As New cDBA
        Dim dt As DataTable
        Dim strLevelStar As String

        strLevelStar = m_oOrder.Ordhead.Customer.ClassificationId

        If strLevelStar Is Nothing Then Exit Sub
        If orderType Is Nothing Then Exit Sub

        Dim strsqlMain As String =
           "SELECT     o.DATA_FIELD, ISNULL(o.SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, o.SHOW_MESSAGEBOX, o.ORDER_TYPE " &
           "FROM       OEI_ITEM_SPEC_INSTRUCTION as o WITH (NOLOCK) " &
           "WHERE   levelstar = '" & strLevelStar & "' AND ORDER_TYPE = '" & orderType & "' AND  SHOW_MESSAGEBOX = 1 "

        ' add this if need to look the customer number  o.CUS_NO = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "' AND " &

        dt = db.DataTable(strsqlMain)
        If dt.Rows.Count <> 0 Then

            For Each dtRow As DataRow In dt.Rows
                'show popup
                If _mess = "0" Then
                    showAppendMessageLog("Customer Instruction", dtRow.Item("SPEC_INSTRUCTION").ToString)
                ElseIf _mess = "1" Then
                    _mess = dtRow.Item("SPEC_INSTRUCTION").ToString
                End If
            Next

        End If
    End Sub


    Public Sub Pop_MessageBox(ByVal pstrItem_No)

        Dim db As New cDBA
        Dim dt As DataTable
        Dim productStyle As String

        If pstrItem_No Is Nothing Then Exit Sub

        'Dim strsql As String = _
        '"SELECT     ITEM_NO, DATA_FIELD, ISNULL(SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, ITEM_CD " & _
        '"FROM       OEI_ITEM_SPEC_INSTRUCTION WITH (NOLOCK) " & _
        '"WHERE      ITEM_NO = '" & pstrItem_No.Replace("'", "''") & "' AND " & _
        '"           DEC_MET_GROUP_ID IS NULL AND " & _
        '"           (ISNULL(CUS_NO, '') = '' OR ISNULL(CUS_NO, '') = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "') AND " & _
        '"           DATA_FIELD = 'ITEM_NO' AND SHOW_MESSAGEBOX = 1 " & _
        '"ORDER BY   ITEM_NO, DATA_FIELD "

        Dim strsql As String = _
       "SELECT     o.ITEM_NO, o.DATA_FIELD, ISNULL(o.SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, o.SHOW_MESSAGEBOX, o.ITEM_CD " & _
       "FROM       OEI_ITEM_SPEC_INSTRUCTION as o WITH (NOLOCK) " & _
       "   INNER JOIN imitmidx_sql as i ON o.item_cd = i.user_def_fld_1 " & _
       "WHERE      i.ITEM_NO = '" & pstrItem_No.Replace("'", "''") & "' AND (i.item_no = o.item_no or ISNULL(o.ITEM_NO, '') = '') AND " & _
       "           o.DEC_MET_GROUP_ID IS NULL AND " & _
       "           (ISNULL(o.CUS_NO, '') = '' OR ISNULL(o.CUS_NO, '') = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "') AND " & _
       "           DATA_FIELD = 'ITEM_NO' AND SHOW_MESSAGEBOX = 1 " & _
       "           AND (ISNULL(o.ORDER_TYPE, '') = '' OR o.ORDER_TYPE = '" & m_oOrder.Ordhead.User_Def_Fld_2 & "') " & _
       "ORDER BY   ITEM_NO, DATA_FIELD "

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then

            'If poImprint Is Nothing Then
            '    poImprint = New cImprint(m_Item_Guid)
            'End If

            For Each dtRow As DataRow In dt.Rows
                'show popup
                productStyle = dtRow.Item("ITEM_CD").Trim()
                showAppendMessageLog(productStyle, dtRow.Item("SPEC_INSTRUCTION").ToString)
            Next

        End If

    End Sub

    '------------------------ function for customer PEERNET Group : ticket N#:23893 ------------------------

    Private Function SaveProp65Alert(ByVal p_strSubject As String, ByVal p_Body As String, ByVal p_item_guid As String) As Boolean


        Try
            Dim strName As String = ""
            Dim strNames As String = ""
            Dim strAddresses As String = ""
            Dim strSubject As String = p_strSubject
            Dim _body As String = p_Body



            Dim oMail As New cEmail_Alert_Prop65

            oMail.Email_From = Trim(g_User.Mail)
            oMail.Email_To = Trim(g_User.Mail)
            oMail.Email_To_Name = Trim(g_User.FullName)
            '  oMail.Email_CC = strAddresses

            oMail.Subject = strSubject
            oMail.Body = _body
            ' oMail.Ord_Guid = ""


            oMail.UserID = g_User.Usr_ID
            oMail.Item_GUID = p_item_guid
            ' oMail.CreateTS = Now
            oMail.Save()



        Catch er As System.Exception
            MsgBox("Error in " & " cOrdLine " & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    '  Private Function AuditMissingProp65(ByVal m_oOrder As cOrder, ByVal p_ordLine As cOrdLine, ByVal _cus_group As String, ByRef poImprint As cImprint) As Boolean
    '  Private Function AuditMissingProp65(ByVal m_oOrder As cOrder, ByVal p_ordLine As cOrdLine, ByVal _cus_group As String) As Boolean
    Private Function AuditMissingProp65(ByVal m_oOrder As cOrder, ByVal p_item_no As String, ByVal p_ordLineItemGuid As String, ByVal _cus_group As String, ByVal _message As String) As Boolean

        AuditMissingProp65 = False

        Dim fromAddr As String = Trim(g_User.Mail) '"Spectorandco OEI <it@spectorandco.ca>"
        Dim subject = m_oOrder.Ordhead.OEI_Ord_No & " - PROP 65 item : " & p_item_no & " put on missing data and advise CS to contact the client " '& _cus_group  '"Missing label Prop65" '

        Dim body As String = "In entered order : " & m_oOrder.Ordhead.OEI_Ord_No & " ItemNo : " & p_item_no & " is in PROP 65 label . "
        Dim addresses(3) As String
        addresses(0) = Trim(g_User.Mail) '"iond@spectorandco.ca"
        Dim oMessage As New cMail()


        If _message.Length <> 0 Then
            ' subject = m_oOrder.Ordhead.OEI_Ord_No & " PROP 65 for " & _message
            subject = m_oOrder.Ordhead.OEI_Ord_No & " - " & _message
        End If

        Try

            'need create function for return DML email, who enter order, for now until testing will be mine iond

            'addresses(1) = "iond@spectorandco.ca"
            'addresses(2) = "dmlusern@spectorandco.ca"

            'turn vb breaklines into HTML breaklines
            'body = Replace(body, vbCrLf, "<BR> Merci / Thank you .")

            oMessage.ToAddr = addresses(0) '"iond@spectorandco.ca"
            oMessage.FromAddr = "OEI Service <it@spectorandco.ca>"
            oMessage.Subject = subject
            oMessage.Message = body & "<br> " & _cus_group & " group prop65 item, put on missing data and advise CS to contact the client on how to proceed.  <br><br> IT "

            oMessage.Send()



            'send email
            'mEmailer.SendGenericEmail65(fromAddr, addresses, subject, body)

        Catch er As Exception
            MsgBox("Error in cOrdLine." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        Finally
            'not need to populate Packer_Instruction
            '   PopulatePacker_Comm(poImprint)
            Call SaveProp65Alert(subject, body, p_ordLineItemGuid)
        End Try

    End Function
    'function disabled, not need to populate packer instruction , will be added Route : Missing Data
    Private Sub PopulatePacker_Comm(ByRef poImprint As cImprint)
        Try

            If poImprint Is Nothing Then
                poImprint = New cImprint(m_Item_Guid)
            End If
            Dim pack_instr As String = ""
            If poImprint.Packer_Instructions.Contains("APPLY PROP 65 LABELS") Then

            Else
                pack_instr = IIf(poImprint.Packer_Instructions = "", "*APPLY PROP 65 LABELS", " | *APPLY PROP 65 LABELS")
            End If

            poImprint.Packer_Instructions &= pack_instr

        Catch er As Exception
            MsgBox("Error in cOrdLine." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub
    'need to check about kit
    '  Public Sub MissingProp65(ByVal p_OrdLine As cOrdLine, ByVal m_oOrder As cOrder, ByRef poImprint As cImprint)
    ' Public Sub MissingProp65(ByVal p_OrdLine As cOrdLine, ByVal m_oOrder As cOrder)
    Public Sub MissingProp65(ByVal p_Item_No As String, ByVal p_itemLine_Guid As String, ByVal m_oOrder As cOrder, ByVal p_Kit_Component As Boolean)
        Dim strSql As String = ""
        Dim _stateprovince As String = ""

        'check if item exist (added because we need to find in OEI_Instruction item not added it mean any half item will be displayed like is not there)
        If CheckExistItem(p_Item_No) = False Then Exit Sub
        'if it try to double message 
        If CheckDoublemessage(p_itemLine_Guid) Then Exit Sub

        'm_oOrder.Ordhead.Cus_No, m_oOrder.Ordhead.Ship_To_State
        Dim Cus_Group As String = CUSGroupForToCheck(m_oOrder.Ordhead.Cus_No, m_oOrder.Ordhead.Ship_To_State)

        If Cus_Group.Length <> 0 Then
            Dim db As New cDBA
            Dim dt As DataTable
            Dim _message As String = ""
            If p_Kit_Component <> True Then

                Try
                    If checkItemException(p_Item_No) = True Then


                        If m_Kit.IsAKit <> True Then

                            'need to validate if item exist in imitimidex in case was entered half of item and item full come from production item lines something like that
                            strSql = " select * from OEI_MIssing_Prop65_Ion  " &
                " where DATA_FIELD = 'PROP 65' and Ltrim(Rtrim(ITEM_NO)) = '" & Trim(p_Item_No) & "'  and textfield3 = '" & Cus_Group & "' and isnull(CUS_NO,'') = '" & m_oOrder.Ordhead.Cus_No & "' "

                            dt = db.DataTable(strSql)


                            'if item missing in prop 65
                            'If dt.Rows.Count = 0 Then
                            'latest changes or missunderstood, need to activate message in case if item is in list Prop 65
                            If dt.Rows.Count <> 0 Then
                                'send Email 
                                'AuditMissingProp65(m_oOrder, p_OrdLine, Cus_Group, poImprint)
                                AuditMissingProp65(m_oOrder, p_Item_No, p_itemLine_Guid, Cus_Group, _message)
                                'show message
                                '    MsgBox("For Cus_no: " & Trim(m_oOrder.Ordhead.Cus_No) & " from Groupe: " & Cus_Group & "  missing Prop65 label on item : " & Trim(p_Item_No))
                                MsgBox(Cus_Group & " group prop65 item, put on missing data and advise CS to contact the client on how to proceed : " & Trim(p_Item_No) & vbCrLf & " Please follow instruction.")

                            End If

                        Else

                            Dim dbk As New cDBA
                            Dim dtk As DataTable

                            'need to validate if item exist in imitimidex in case was entered half of item and item full come from production item lines something like that
                            strSql = " select * from OEI_MIssing_Prop65_Ion  " &
                " where DATA_FIELD = 'PROP 65' and Ltrim(Rtrim(ITEM_NO)) = '" & Trim(p_Item_No) & "'  and textfield3 = '" & Cus_Group & "' and isnull(CUS_NO,'') = '" & m_oOrder.Ordhead.Cus_No & "' "

                            dtk = dbk.DataTable(strSql)

                            'If dtk.Rows.Count = 0 Then
                            If dtk.Rows.Count <> 0 Then
                                Dim dbc As New cDBA
                                Dim dtc As DataTable

                                _message &= vbCrLf & "Kit Item : " & Trim(p_Item_No) & vbCrLf

                                'if is true find components and add in alert message
                                strSql = "select Item_no,seq_no,comp_item_no from imkitfil_sql where item_no = '" & Trim(p_Item_No) & "'"

                                dtc = dbc.DataTable(strSql)

                                If dtc.Rows.Count <> 0 Then

                                    Dim dbm As New cDBA
                                    Dim dtm As DataTable
                                    Dim _cpt As Int16 = 0


                                    For Each dr As DataRow In dtc.Rows
                                        'check if components missing too
                                        strSql = " select * from OEI_MIssing_Prop65_Ion  " &
                " where DATA_FIELD = 'PROP 65' and Ltrim(Rtrim(ITEM_NO)) = '" & Trim(dr.Item("comp_item_no").ToString) & "'  and textfield3 = '" & Cus_Group & "' and isnull(CUS_NO,'') = '" & m_oOrder.Ordhead.Cus_No & "' "

                                        dtm = dbc.DataTable(strSql)
                                        If dtm.Rows.Count <> 0 Then
                                            _cpt = _cpt + 1
                                            '   _message &= _cpt & ".Component : " & Trim(dr.Item("comp_item_no").ToString) & " Missing PROP65 Label." & vbCrLf
                                            _message &= _cpt & ".Component : " & Trim(dr.Item("comp_item_no").ToString) & vbCrLf
                                        End If

                                    Next

                                End If

                                'send Email 
                                'AuditMissingProp65(m_oOrder, p_OrdLine, Cus_Group, poImprint)
                                AuditMissingProp65(m_oOrder, p_Item_No, p_itemLine_Guid, Cus_Group, _message)
                                'show message
                                ' MsgBox("For Cus_no: " & Trim(m_oOrder.Ordhead.Cus_No) & " from Groupe: " & Cus_Group & vbCrLf & "  Missing Prop65 label on  : " & _message)
                                MsgBox(Cus_Group & " group prop65 item, put on missing data and advise CS to contact the client on how to proceed" & _message & vbCrLf & " Please follow instruction.")

                            Else ' in case when Kit is in PPROP65 OEI_ITEM_SPEC_INSTRUCTION Table, need to check for components

                                Dim dbc As New cDBA
                                Dim dtc As DataTable

                                'if is true find components and add in alert message
                                strSql = "select Item_no,seq_no,comp_item_no from imkitfil_sql where item_no = '" & Trim(p_Item_No) & "'"

                                dtc = dbc.DataTable(strSql)

                                If dtc.Rows.Count <> 0 Then

                                    Dim dbm As New cDBA
                                    Dim dtm As DataTable
                                    Dim _cpt As Int16 = 0
                                    _message &= "Kit Item : " & Trim(p_Item_No) & vbCrLf

                                    For Each dr As DataRow In dtc.Rows
                                        'check if components missing too
                                        strSql = " select * from OEI_MIssing_Prop65_Ion  " &
                " where DATA_FIELD = 'PROP 65' and Ltrim(Rtrim(ITEM_NO)) = '" & Trim(dr.Item("comp_item_no").ToString) & "'  and textfield3 = '" & Cus_Group & "' and isnull(CUS_NO,'') = '" & m_oOrder.Ordhead.Cus_No & "' "

                                        dtm = dbc.DataTable(strSql)

                                        ' If dtm.Rows.Count = 0 Then
                                        If dtm.Rows.Count <> 0 Then
                                            _cpt = _cpt + 1
                                            _message &= _cpt & ".Component : " & Trim(dr.Item("comp_item_no").ToString) & vbCrLf

                                        End If

                                    Next

                                    If _cpt <> 0 Then
                                        'send Email 
                                        'AuditMissingProp65(m_oOrder, p_OrdLine, Cus_Group, poImprint)
                                        AuditMissingProp65(m_oOrder, p_Item_No, p_itemLine_Guid, Cus_Group, _message)
                                        'show message
                                        'MsgBox("For Cus_no: " & Trim(m_oOrder.Ordhead.Cus_No) & " from Groupe: " & Cus_Group & vbCrLf & "  Missing Prop65 label on  : " & _message)
                                        MsgBox(Cus_Group & " group prop65 item, put on missing data and advise CS to contact the client on how to proceed " & _message & vbCrLf & " Please follow instruction.")
                                    End If

                                End If

                                '------------------------------------

                            End If

                        End If

                    End If

                Catch er As Exception
                    MsgBox("Error in cOrdLine." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
                End Try
            Else
                'work with kit components,all components must be included in one message
            End If


        End If

    End Sub

    Private Function CheckDoublemessage(ByVal p_OrdLineGuid As String) As Boolean
        '    CheckDoublemessage = False
        Dim m_doubleMessage As Boolean = False
        Dim strSql As String = ""
        Try
            Dim db As New cDBA
            Dim dt As DataTable


            strSql = " select * from OEI_Email_Alert_Prop65  " &
        " where  item_guid = '" & p_OrdLineGuid & "'"

            dt = db.DataTable(strSql)
            'if find record this mean no need to mention about Prop65, already inserted and alert was sent and displayed
            If dt.Rows.Count <> 0 Then
                m_doubleMessage = True
            End If

        Catch er As Exception
            MsgBox("Error in cOrdLine." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        CheckDoublemessage = m_doubleMessage

    End Function

    Private Function CUSGroupForToCheck(ByVal _p_cus_no As String, ByVal _p_prov_ois_instr As String) As String

        CUSGroupForToCheck = ""
        Dim strSql As String = ""
        Dim _group As String = ""
        ' Dim _state_prov As String = ""
        Try

            Dim db As New cDBA
            Dim dt As DataTable

            strSql = " select ig.DATA_FIELD, ig.STATE_PROV ,cmp_code, cmp_name,  region, ISNULL(textfield3,'') as textfield3 , numberfield3  " &
                "  From cicmpy c inner Join OEI_ITEM_SPEC_INSTRUCTION ig on ltrim(rtrim(c.textfield3)) = ltrim(rtrim(ig.CUS_GROUP)) " &
                "  where DATA_FIELD = 'MISSING_PROP65' and Ltrim(Rtrim(cmp_code)) =  '" & Trim(_p_cus_no) & "' and ig.STATE_PROV = '" & _p_prov_ois_instr & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                _group = dt.Rows(0).Item("textfield3").ToString
                '  _state_prov = dt.Rows(0).Item("STATE_PROV").ToString
            End If

        Catch er As Exception
            MsgBox("Error in cOrdLine." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
        ' _p_prov_ois_instr = _state_prov
        CUSGroupForToCheck = _group

    End Function
    'need function for validate if item exist in imitimdx_sql
    Private Function CheckExistItem(ByVal _p_Item As String) As Boolean

        CheckExistItem = False
        Dim strSql As String = ""
        Dim _eXist As Boolean = False
        ' Dim _state_prov As String = ""
        Try

            Dim db As New cDBA
            Dim dt As DataTable

            strSql = " Select * from imitmidx_sql where Ltrim(Rtrim(item_no)) = '" & Trim(_p_Item) & "'"

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                _eXist = True
            End If

        Catch er As Exception
            MsgBox("Error in cOrdLine." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
        ' _p_prov_ois_instr = _state_prov
        CheckExistItem = _eXist

    End Function

    Private Function checkItemException(ByVal p_item As String) As Boolean


        checkItemException = False
        Dim strSql As String = ""
        Dim _eXist As Boolean = False
        ' Dim _state_prov As String = ""
        Try

            Dim db As New cDBA
            Dim dt As DataTable
            'created view with eliminate all 44 code also in view we can to add more exception
            strSql = " Select * from View_imitmidx_prop65_ID where Ltrim(Rtrim(item_no)) = '" & Trim(p_item) & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                _eXist = True
            End If

        Catch er As Exception
            MsgBox("Error in cOrdLine." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
        ' _p_prov_ois_instr = _state_prov
        checkItemException = _eXist

    End Function

    '----------------------------------------------------------------------------

    Public Sub Set_Spec_Instructions(ByVal pstrItem_No As String, ByRef poImprint As cImprint)

        Dim db As New cDBA
        Dim dt As DataTable
        Dim productStyle As String
        Dim count As Int32 = 0
        Dim pack_instr As String = ""
        Dim pack_pro65 As String = ""
        Dim unionPrdCatShipTo As String = ""

        Dim nonItemFilter As String = ""

        If m_oOrder.Ordhead.Ship_To_State IsNot Nothing AndAlso Trim(m_oOrder.Ordhead.Ship_To_State) <> "" Then
            nonItemFilter = "OR (ISNULL(o.STATE_PROV, '') = '" & Trim(m_oOrder.Ordhead.Ship_To_State) & "' AND ISNULL(o.CUS_NO, '') = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "' )"
            If frmOrder.shipToProp65State = True Then

                'union string to sql

                unionPrdCatShipTo = " union " _
                 & " SELECT  ITEM_NO,  DATA_FIELD, '*' + ISNULL( SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION,  ISNULL(SHOW_MESSAGEBOX,0) as SHOW_MESSAGEBOX, ISNULL( ITEM_CD, 'GEN') as ITEM_CD,  " _
                 & " ISNULL( PROD_CAT, '') as PROD_CAT, ISNULL(IMPRINT, '') as IMPRINT,  STATE_PROV FROM OEI_ITEM_SPEC_INSTRUCTION where  PROD_CAT = '" & spec_instruction_by_Prod_cat(Trim(pstrItem_No)) & "' " _
                 & " AND ISNULL(STATE_PROV, '') = '" & Trim(m_oOrder.Ordhead.Ship_To_State) & "' " _
                 & " UNION " _
                 & " SELECT ITEM_NO,  DATA_FIELD, '*' + ISNULL( SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION,  ISNULL(SHOW_MESSAGEBOX,0) as SHOW_MESSAGEBOX, ITEM_CD,  " _
                 & " ISNULL(PROD_CAT, '') as PROD_CAT, ISNULL(IMPRINT, '') as IMPRINT,  STATE_PROV FROM OEI_ITEM_SPEC_INSTRUCTION  WITH (NOLOCK)  " _
                 & " WHERE  ISNULL(CUS_NO, '') = '" & Trim(m_oOrder.Ordhead.Cus_No.Replace("'", "''")) & "' " _
                 & " AND SHIP_TO like '%" & Trim(m_oOrder.Ordhead.Ship_To_Zip.Replace("'", "''")) & "%' and isnull(ITEM_CD,'') = ''  "

                If Not poImprint Is Nothing Then
                    If Trim(poImprint.Imprint) <> "" Then
                        unionPrdCatShipTo &= " UNION " _
                                       & " SELECT Distinct ITEM_NO,  DATA_FIELD, '*' + ISNULL( SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION,  ISNULL(SHOW_MESSAGEBOX,0) as SHOW_MESSAGEBOX, ISNUll(ITEM_CD,'') As ITEM_CD,  " _
                                       & " ISNULL(PROD_CAT, '') as PROD_CAT, ISNULL(IMPRINT, '') as IMPRINT,  STATE_PROV FROM OEI_ITEM_SPEC_INSTRUCTION  WITH (NOLOCK)  " _
                                       & " WHERE  ISNULL(CUS_NO, '') = '" & Trim(m_oOrder.Ordhead.Cus_No.Replace("'", "''")) & "' " _
                                       & " AND (ISNULL(STATE_PROV, '') = '" & Trim(m_oOrder.Ordhead.Ship_To_State.Replace("'", "''")) & "' or ISNULL(STATE_PROV, '') = '' ) " _
                                       & " AND ISNULL(IMPRINT, '') like '%" & poImprint.Imprint.Replace("'", "''") & "%' " _
                                       & " AND (ISNULL(ITEM_CD,'') = (select Ltrim(Rtrim(user_def_fld_1)) as Item_cd from imitmidx_sql where item_no = '" & Trim(pstrItem_No) & "') " _
                                       & " Or ISNULL(ITEM_CD,'') = '')  "


                        '++ID combined 7.10.2019
                        '& " AND (ISNULL(ITEM_CD,'') = '" & Trim(pstrItem_No) & "' Or ISNULL(ITEM_CD,'') = '')  "

                        '                  & " UNION " _
                        '& " Select  ISNULL(Item_No,'') as ITEM_NO,  DATA_FIELD, '*' + ISNULL( SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, ISNULL(SHOW_MESSAGEBOX,0) as SHOW_MESSAGEBOX, ISNULL(ITEM_CD,'') As ITEM_CD, " _
                        '& " ISNULL(PROD_CAT, '') as PROD_CAT, ISNULL(IMPRINT, '') as IMPRINT, STATE_PROV FROM OEI_ITEM_SPEC_INSTRUCTION  With (NOLOCK)   WHERE  ISNULL(CUS_NO, '') = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "' " _
                        '& " And  isnull(IMPRINT,'') <> '' "

                    End If



                End If




                ' put in comment 5.31.2018
                'If pstrItem_No.Replace("'", "''").Contains("00BG") Then
                '    nonItemFilter = nonItemFilter & " Or (ISNULL(o.STATE_PROV, '') = '" & m_oOrder.Ordhead.Ship_To_State & "' AND ISNULL(o.PROD_CAT, '') = 'BAG')"

                'Else
                '    'Dim oReponse = MsgBox("Enable automatic Prop 65 message on order line?", MsgBoxStyle.YesNo, "OE Interface")
                '    ' If oReponse = MsgBoxResult.Yes Then
                '    'nonItemFilter = nonItemFilter & " Or ISNULL(o.STATE_PROV, '') = '" & m_oOrder.Ordhead.Ship_To_State & "'"
                '    'MsgBox("Prop 65 automatic instructions have been enabled. Prop 65 instructions are automatially disabled when setting the ship state to something other than 'CA'")
                '    ' Else
                '    ' ' frmOrder.shipToProp65State = False
                '    'End If
                'End If



            End If
        End If

        Dim isNotItem = False
        Dim prefix = Mid(pstrItem_No.Replace("'", "''").Trim(), 1, 2)
        If (prefix = "44" Or prefix = "88") Then
            isNotItem = True
        End If



        Dim strsql As String = ""
        '"SELECT     o.ITEM_NO, o.DATA_FIELD, '*' + ISNULL(o.SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, o.SHOW_MESSAGEBOX, " &
        '" ISNULL(o.ITEM_CD, 'GEN') as ITEM_CD, ISNULL(o.PROD_CAT, '') as PROD_CAT, ISNULL(IMPRINT, '') as IMPRINT, o.STATE_PROV " &
        '"FROM       OEI_ITEM_SPEC_INSTRUCTION as o WITH (NOLOCK) " &
        '"   LEFT JOIN imitmidx_sql as i ON o.item_cd = i.user_def_fld_1 " &
        '"WHERE      ((i.ITEM_NO = '" & pstrItem_No.Replace("'", "''") & "' AND (i.item_no = o.item_no or ISNULL(o.ITEM_NO, '') = '') AND " &
        '"           o.DEC_MET_GROUP_ID IS NULL) OR (o.item_no IS NULL AND o.IS_GLOBAL = 1) " & nonItemFilter & ") AND " &
        '"           (ISNULL(o.CUS_NO, '') = '' OR ISNULL(o.CUS_NO, '') = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "') " &
        '"           AND (ISNULL(o.ORDER_TYPE, '') = '' OR o.ORDER_TYPE = '" & m_oOrder.Ordhead.User_Def_Fld_2.Replace("'", "''") & "') " &
        ' "ORDER BY   ITEM_NO, DATA_FIELD "
        ' Debug.Print(m_oOrder.Ordhead.Ship_To_Name)

        strsql =
        " SELECT     o.ITEM_NO, o.DATA_FIELD, '*' + ISNULL(o.SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, ISNULL(o.SHOW_MESSAGEBOX,0) as SHOW_MESSAGEBOX, " &
        " ISNULL(o.ITEM_CD, 'GEN') as ITEM_CD, ISNULL(o.PROD_CAT, '') as PROD_CAT, ISNULL(IMPRINT, '') as IMPRINT, o.STATE_PROV " &
        " FROM       OEI_ITEM_SPEC_INSTRUCTION as o WITH (NOLOCK) " &
        " LEFT JOIN imitmidx_sql as i ON o.item_cd = i.user_def_fld_1 " &
        " WHERE      ((i.ITEM_NO = '" & pstrItem_No.Replace("'", "''") & "' AND (i.item_no = o.item_no or ISNULL(o.ITEM_NO, '') = '') AND " &
        "           o.DEC_MET_GROUP_ID IS NULL) OR (o.item_no IS NULL AND o.IS_GLOBAL = 1) ) AND " &
        "           (ISNULL(o.CUS_NO, '') = '' OR ISNULL(o.CUS_NO, '') = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "') " &
        "           AND (ISNULL(o.ORDER_TYPE, '') = '' OR o.ORDER_TYPE = '" & m_oOrder.Ordhead.User_Def_Fld_2.Replace("'", "''") & "') and " &
        "  (ISNULL(o.STATE_PROV, '') = '" & m_oOrder.Ordhead.Ship_To_State.Replace("'", "''") & "' or ISNULL(o.STATE_PROV, '') = '')  " &
       unionPrdCatShipTo

        '++ID 2.21.2019 added union sql for identify message for prod cat = pencil and sub_cat = Aluminium and cmp_fctry = CA 
        strsql &= " UNION " _
      & " Select  TOP 1  INS.ITEM_NO,  INS.DATA_FIELD, ISNULL(INS.SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION,  ISNULL(INS.SHOW_MESSAGEBOX,0) as SHOW_MESSAGEBOX, " _
      & " ISNULL( P.ITEM_CD, 'GEN') as ITEM_CD,    ISNULL( INS.PROD_CAT, '') as PROD_CAT,  ISNULL(INS.IMPRINT, '') as IMPRINT,  INS.STATE_PROV FROM OEI_ITEM_SPEC_INSTRUCTION INS INNER JOIN " _
      & " ( " _
      & " select M.ITEM_CD,i.prod_cat, item_note_2, Item_No  from imitmidx_sql i inner join " _
      & " (  " _
      & " Select m.Item_Cd,ENUM_VALUE  from MDB_ITEM_MASTER m inner join MDB_CFG_ENUM e On m.ITEM_CLASSIFICATION_ID = e.ID " _
      & " where ITEM_CLASSIFICATION_ID = 86 /* Writing Instrument*/  And ISNULL(m.ITEM_CD,'') <> '' " _
      & " ) as M on lTRIM(RTRIM(ISNULL(i.user_def_fld_1,''))) = LTRIM(RTRIM(isnull(M.ITEM_CD,''))) " _
      & " ) AS P ON LTRIM(RTRIM(ins.PROD_CAT)) = LTRIM(RTRIM(p.prod_cat)) And LTRIM(RTRIM(INS.SUB_PROD_CAT)) = LTRIM(RTRIM(P.item_note_2)) " _
      & " where ISNULL(INS.Cmp_Fctry, '') = ISNULL('" & Trim(m_oOrder.Ordhead.Customer.cmp_fctry) & "','') AND ISNULL(INS.ITEM_NO,'') = '' AND ISNULL(DATA_FIELD,'') = ''  " _
      & " And ISNULL(INS.Item_Cd,'') = '' AND ISNULL(INS.IMPRINT ,'') = '' AND ISNULL(INS.STATE_PROV,'') = '' And P.item_no = '" & Trim(pstrItem_No) & "'  "

        'added exception by deco met finally find below exeist function on line 5910
        '  & " UNION " _
        '& " SELECT ITEM_NO,  DATA_FIELD, '*' + ISNULL( SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION,  ISNULL(SHOW_MESSAGEBOX,0) as SHOW_MESSAGEBOX, ITEM_CD,   ISNULL(PROD_CAT, '') as PROD_CAT,  " _
        '& " ISNULL(IMPRINT, '') as IMPRINT,  STATE_PROV    " _
        '& " FROM OEI_ITEM_SPEC_INSTRUCTION i WITH (NOLOCK) inner join EXACT_TRAVELER_ROUTE r WITH (NOLOCK) on i.DEC_MET_GROUP_ID = r.Dec_Met_ID  " _
        '& "  WHERE  ISNULL(CUS_NO, '') = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "'   and isnull(ITEM_CD,'') = '" & Trim(Item_Cd) & "' and ISNUll(DEC_MET_GROUP_ID,0) <> '' " _
        '& " And r.RouteID in (" & RouteID & ")  " _


        '  & " ORDER BY ITEM_NO, DATA_FIELD  " _

        dt = db.DataTable(strsql)

        If dt.Rows.Count <> 0 Then

            If poImprint Is Nothing Then
                poImprint = New cImprint(m_Item_Guid)
            End If
            'I need initialise packer properties because some time return with old value
            '  poImprint.Packer_Instructions = ""

            For Each dtRow As DataRow In dt.Rows

                Select Case Trim(dtRow.Item("DATA_FIELD").ToString.ToUpper)

                    Case "COMMENT"
                        poImprint.Comments = IIf(poImprint.Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Comments.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Comments, (poImprint.Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "COMMENT2"
                        poImprint.Special_Comments = IIf(poImprint.Special_Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Special_Comments.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Special_Comments, (poImprint.Special_Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "PRINTER_COMMENT"
                        poImprint.Printer_Comment = IIf(poImprint.Printer_Comment = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Printer_Comment.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Printer_Comment, (poImprint.Printer_Comment & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "PACKER_COMMENT"
                        poImprint.Packer_Comment = IIf(poImprint.Packer_Comment = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Packer_Comment.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Packer_Comment, (poImprint.Packer_Comment & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "PRINTER_INSTRUCTIONS"
                        poImprint.Printer_Instructions = IIf(poImprint.Printer_Instructions = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Printer_Instructions.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Printer_Instructions, (poImprint.Printer_Instructions & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "PACKER_INSTRUCTIONS"
                        If (dtRow.Item("IMPRINT").ToString.Trim() = "" OrElse (dtRow.Item("IMPRINT").ToString.ToLower().Trim().Contains(poImprint.Imprint.ToLower()) And poImprint.Imprint <> "")) And isNotItem = False Then
                            pack_instr = IIf(poImprint.Packer_Instructions = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Packer_Instructions.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Packer_Instructions, (poImprint.Packer_Instructions & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                            poImprint.Packer_Instructions = pack_instr
                        End If
                    Case "PROP 65"
                        If (dtRow.Item("IMPRINT").ToString.Trim() = "" OrElse (dtRow.Item("IMPRINT").ToString.ToLower().Trim().Contains(poImprint.Imprint.ToLower().Trim()) And poImprint.Imprint <> "")) And isNotItem = False Then
                            pack_instr = IIf(poImprint.Packer_Instructions = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Packer_Instructions.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Packer_Instructions, (poImprint.Packer_Instructions & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                            poImprint.Packer_Instructions = pack_instr
                        End If
                    Case "ITEM_NO"

                        If dtRow.Item("SHOW_MESSAGEBOX") = 0 Then
                            poImprint.Comments = IIf(poImprint.Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Comments.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Comments, (poImprint.Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                        End If

                        'Case "SUPER SPEED"
                        '    '    poImprint.Special_Comments = IIf(poImprint.Special_Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, (poImprint.Special_Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString))
                        '    'must to fill urgent field
                        '    m_oOrder.Ordhead.User_Def_Fld_1 = Trim(m_oOrder.Ordhead.Cus_Name.Replace("'", "''"))    
                End Select

                '    If Trim(pack_instr) <> "" Then poImprint.Packer_Instructions = pack_instr




                'show popup if specified
                If Not IsDBNull(dtRow.Item("SHOW_MESSAGEBOX")) AndAlso dtRow.Item("SHOW_MESSAGEBOX") Then
                    productStyle = dtRow.Item("ITEM_CD").ToString.Trim()
                    showAppendMessageLog(productStyle, dtRow.Item("SPEC_INSTRUCTION").ToString)
                End If

            Next

        End If

    End Sub

    Private Function spec_instruction_by_Prod_cat(ByVal pItemCd As String) As String
        spec_instruction_by_Prod_cat = ""
        Try
            Dim pProd As String = ""

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSQL As String = ""

            strSQL = "Select prod_cat from imitmidx_sql where item_no = '" & Trim(pItemCd) & "'"

            dt = db.DataTable(strSQL)

            If dt.Rows.Count Then
                pProd = dt.Rows(0).Item(0).ToString
            End If

            Return Trim(pProd)

        Catch ex As Exception
            MsgBox("Error in cOrdLine." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Public Sub Set_Spec_Instructions(ByVal pstrItem_No As String, ByVal piRouteID As Integer, ByRef poImprint As cImprint)

        Dim db As New cDBA
        Dim dt As DataTable
        'use item_no
        Dim strsql As String =
        "SELECT     I.ITEM_NO, I.DATA_FIELD, '*' + ISNULL(I.SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, SHOW_MESSAGEBOX " &
        "FROM       OEI_ITEM_SPEC_INSTRUCTION I WITH (NOLOCK) " &
        "   INNER JOIN imitmidx_sql as m ON I.item_cd = m.user_def_fld_1 " &
        "INNER JOIN MDB_CFG_DEC_MET DM WITH (NOLOCK) ON I.DEC_MET_GROUP_ID = DM.DEC_MET_GROUP_ID " &
        "INNER JOIN EXACT_TRAVELER_ROUTE R WITH (NOLOCK) ON DM.DEC_MET_ID = R.DEC_MET_ID " &
        "WHERE      m.ITEM_NO = '" & pstrItem_No.Replace("'", "''") & "' AND (I.item_no = m.item_no or ISNULL(I.ITEM_NO, '') = '') AND " &
        "           ISNULL(R.ROUTEID, 0) = " & piRouteID.ToString & " AND " &
        "           (ISNULL(CUS_NO, '') = '' OR ISNULL(CUS_NO, '') = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "')" &
        "           AND (ISNULL(ORDER_TYPE, '') = '' OR ORDER_TYPE = '" & m_oOrder.Ordhead.User_Def_Fld_2 & "') " &
        "ORDER BY   ITEM_NO, DATA_FIELD "

        If Trim(pstrItem_No) = String.Empty Then
            strsql = " select DATA_FIELD, SPEC_INSTRUCTION, SHOW_MESSAGEBOX, USER_LOGIN,  RouteID from OEI_ITEM_SPEC_INSTRUCTION  where  RouteID = " & piRouteID & " "
        End If

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then

            If poImprint Is Nothing Then
                poImprint = New cImprint(m_Item_Guid)
            End If

            For Each dtRow As DataRow In dt.Rows

                Select Case dtRow.Item("DATA_FIELD").ToString.ToUpper

                    Case "COMMENT"
                        poImprint.Comments = IIf(poImprint.Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Comments.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Comments, (poImprint.Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "COMMENT2"
                        poImprint.Special_Comments = IIf(poImprint.Special_Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Special_Comments.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Special_Comments, (poImprint.Special_Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "PRINTER_COMMENT"
                        poImprint.Printer_Comment = IIf(poImprint.Printer_Comment = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Printer_Comment.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Printer_Comment, (poImprint.Printer_Comment & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "PACKER_COMMENT"
                        poImprint.Packer_Comment = IIf(poImprint.Packer_Comment = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Packer_Comment.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Packer_Comment, (poImprint.Packer_Comment & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "PRINTER_INSTRUCTIONS"
                        poImprint.Printer_Instructions = IIf(poImprint.Printer_Instructions = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Printer_Instructions.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Printer_Instructions, (poImprint.Printer_Instructions & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "PACKER_INSTRUCTIONS"
                        poImprint.Packer_Instructions = IIf(poImprint.Packer_Instructions = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Packer_Instructions.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Packer_Instructions, (poImprint.Packer_Instructions & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                    Case "ITEM_NO"
                        poImprint.Comments = IIf(poImprint.Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Comments.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Comments, (poImprint.Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                        'Case "ORDER_ACK_ITEM_COMMENT"
                        '    poImprint.Special_Comments = IIf(poImprint.Special_Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, (poImprint.Special_Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString))
                    Case "SUPER SPEED"

                        '++ID 08.08.2022 message box yes/no
                        Dim result As DialogResult = MessageBox.Show(dtRow.Item("SPEC_INSTRUCTION").ToString, "SUPPER SPEED NOTE", MessageBoxButtons.YesNo)

                        If result = DialogResult.No Then

                        ElseIf result = DialogResult.Yes Then

                            'must to fill urgent field
                            m_oOrder.Ordhead.User_Def_Fld_1 = Trim(dtRow.Item("DATA_FIELD").ToString.ToUpper)

                        End If
                    Case "NFC"
                        poImprint.Packer_Instructions = IIf(poImprint.Packer_Instructions = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Packer_Instructions.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Packer_Instructions, (poImprint.Packer_Instructions & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                End Select

                'show popup if specified
                If Not IsDBNull(dtRow.Item("SHOW_MESSAGEBOX")) AndAlso dtRow.Item("SHOW_MESSAGEBOX") Then MsgBox(dtRow.Item("SPEC_INSTRUCTION").ToString)
            Next
        End If
    End Sub

    Public Function Set_Spec_Imprint_Instructions(ByVal pstrCus_No As String, ByRef poImprint As cImprint) As Boolean

        Set_Spec_Imprint_Instructions = False

        Dim db As New cDBA
        Dim dt As DataTable

        Dim strsql As String = _
        "SELECT     ITEM_NO, DATA_FIELD, '*' + ISNULL(SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, SHOW_MESSAGEBOX " & _
        "FROM       OEI_ITEM_SPEC_INSTRUCTION WITH (NOLOCK) " & _
        "WHERE      CUS_IMPRINT = 1 AND " & _
        "           DEC_MET_GROUP_ID IS NULL AND " & _
        "           ITEM_CD = '" & pstrCus_No.Replace("'", "''") & "' AND " & _
        "           ITEM_NO = '" & poImprint.Imprint.Replace("'", "''") & "' " & _
        "ORDER BY   ITEM_NO, DATA_FIELD "

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then

            If poImprint Is Nothing Then
                poImprint = New cImprint(m_Item_Guid)
            End If

            For Each dtRow As DataRow In dt.Rows

                Select Case dtRow.Item("DATA_FIELD").ToString.ToUpper

                    Case "COMMENT"
                        poImprint.Comments = IIf(poImprint.Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Comments.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Comments, (poImprint.Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                        Set_Spec_Imprint_Instructions = True
                    Case "COMMENT2"
                        poImprint.Special_Comments = IIf(poImprint.Special_Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Special_Comments.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Special_Comments, (poImprint.Special_Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                        Set_Spec_Imprint_Instructions = True
                    Case "PRINTER_COMMENT"
                        poImprint.Printer_Comment = IIf(poImprint.Printer_Comment = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Printer_Comment.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Printer_Comment, (poImprint.Printer_Comment & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                        Set_Spec_Imprint_Instructions = True
                    Case "PACKER_COMMENT"
                        poImprint.Packer_Comment = IIf(poImprint.Packer_Comment = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Packer_Comment.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Packer_Comment, (poImprint.Packer_Comment & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                        Set_Spec_Imprint_Instructions = True
                    Case "PRINTER_INSTRUCTIONS"
                        poImprint.Printer_Instructions = IIf(poImprint.Printer_Instructions = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Printer_Instructions.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Printer_Instructions, (poImprint.Printer_Instructions & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                        Set_Spec_Imprint_Instructions = True
                    Case "PACKER_INSTRUCTIONS"
                        poImprint.Packer_Instructions = IIf(poImprint.Packer_Instructions = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Packer_Instructions.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Packer_Instructions, (poImprint.Packer_Instructions & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                        Set_Spec_Imprint_Instructions = True
                    Case "ITEM_NO"
                        poImprint.Comments = IIf(poImprint.Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, IIf(poImprint.Comments.Contains(dtRow.Item("SPEC_INSTRUCTION").ToString.Trim()), poImprint.Comments, (poImprint.Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString)))
                        Set_Spec_Imprint_Instructions = True
                        'Case "ORDER_ACK_ITEM_COMMENT"
                        'poImprint.Special_Comments = IIf(poImprint.Special_Comments = "", dtRow.Item("SPEC_INSTRUCTION").ToString, (poImprint.Special_Comments & " | " & dtRow.Item("SPEC_INSTRUCTION").ToString))

                End Select

                'show popup if specified
                If Not IsDBNull(dtRow.Item("SHOW_MESSAGEBOX")) AndAlso dtRow.Item("SHOW_MESSAGEBOX") Then MsgBox(dtRow.Item("SPEC_INSTRUCTION").ToString)

            Next

        End If

    End Function

End Class

