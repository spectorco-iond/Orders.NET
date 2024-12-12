Option Strict Off
Option Explicit On

Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient

Public Class cOrder

    'Private _Ordhead As cOrdHead
    'Private _OrdLines As Collection 'Collection of cOrdLine
    'Private _CreditInfo As cCreditInfo

    Public Source As OEI_SourceEnum = OEI_SourceEnum.frmOrder

    Public Ordhead As cOrdHead
    Public OrdLines As Collection 'Collection of cOrdLine
    Public CreditInfo As cCreditInfo
    Public Descriptions As cOrdheadDesc
    Public Validation As cOrderValidation
    Public Settings As cSettings
    Public Line_No As Integer

    Public Reload As Boolean = False
    Public NewRowsToAdd As Long = 0
    Public FromHistory As Boolean = False

    Public WithEvents Orders As DataSet
    Public WithEvents OrderLines As DataTable
    Public WithEvents OrderLine As DataRow

    Public WorkLine As Integer = 0
    Public UserID As String
    Public OrderSource As OrderSourceEnum

    Public CancelExport As Boolean = False

    Private colDecMetList As Collection
    Private colItemCharges As Collection
    'Private colKitItems As Collection
    Private colLogos As Collection
    Private colItemRC As Collection
    Private cptPricingType As ChargesPricingType
    Private oDecMet As cMDBCfgDecMet

    Private colColorCount As System.Collections.Generic.Dictionary(Of String, Integer) ' Both color collections used for color change charge
    Private colColorList As System.Collections.Generic.Dictionary(Of String, String) ' Both color collections used for color change charge
    Private colColorEnum As Collection
    Private colItemExtraByDecMetID As Collection


    'THESE ARE THE FIELDS EXPORTED TO EXCEL
    Public Enum xlsExport

        Ord_Type = 1
        Ord_Dt
        OE_PO_No
        Cus_No
        Bill_To_Name
        Bill_To_Addr_1
        Bill_To_Addr_2
        Bill_To_Addr_3
        Bill_To_Addr_4
        Bill_To_Country
        Cus_Alt_Adr_Cd
        Ship_To_Name
        Ship_To_Addr_1
        Ship_To_Addr_2
        Ship_To_Addr_3
        Ship_To_Addr_4
        Ship_To_Country
        Shipping_Dt
        Ship_Via_Cd
        Ar_Terms_Cd = 20
        Ship_Instruction_1
        Ship_Instruction_2
        SlsPsn_No
        Discount_Pct ' 24
        Job_No
        Mfg_Loc
        Profit_Center
        Filler_0001
        Curr_Cd
        Orig_Trx_Rt ' 30
        Curr_Trx_Rt
        Tax_Sched
        Contact_1
        Phone_Number ' 34
        Fax_Number
        Email_Address
        Ship_To_City
        Ship_To_State
        Ship_To_Zip
        Bill_To_City ' 40
        Bill_To_State
        Bill_To_Zip
        Tax_Cd ' NEED TO ADD THIS LINE TO CONFIG 43
        Tax_Pct ' NEED TO ADD THIS LINE TO CONFIG 44
        Tax_Cd_2 ' NEED TO ADD THIS LINE TO CONFIG 45
        Tax_Pct_2 ' NEED TO ADD THIS LINE TO CONFIG 46
        Tax_Cd_3 ' NEED TO ADD THIS LINE TO CONFIG 47
        Tax_Pct_3 ' NEED TO ADD THIS LINE TO CONFIG 48
        Hdr_User_Def_Fld_1 ' NEED TO ADD THIS LINE TO CONFIG 49
        Hdr_User_Def_Fld_2 ' NEED TO ADD THIS LINE TO CONFIG 50
        Hdr_User_Def_Fld_3 ' NEED TO ADD THIS LINE TO CONFIG 51
        Hdr_User_Def_Fld_4 ' NEED TO ADD THIS LINE TO CONFIG 52
        Hdr_User_Def_Fld_5 ' NEED TO ADD THIS LINE TO CONFIG 53
        Form_No
        Deter_Rate_By
        Tax_Fg
        Frt_Pay_Cd
        Selection_Cd
        Hold_Fg            '59

        '++ID 07312024 adding inv_batch_id
        inv_batch_id ' 60

        'Misc_Mn_No ' Must set these 4 fields to NULL, completed when invoicing.
        'Misc_Sb_No
        'Frt_Mn_No
        'Frt_Sb_No
        'Item_51 = 51
        'Item_52 ' AZ

        Item_No = 79
        Item_Loc               '80
        Item_Qty_Ordered
        Item_Qty_To_Ship
        Item_Qty_BkOrd
        Item_Unit_Price
        Item_Discount_Pct
        Item_Request_Dt
        Item_Promise_Dt
        Item_Req_Ship_Date
        Item_Unit_Weight ' NEED TO ADD THIS LINE TO CONFIG 63

        Item_BkOrd_Fg       '90
        Item_Bin_Fg
        Item_Prc_Cd_Orig_Price
        Item_Qty_BkOrd_Fg
        Item_Unit_Cost
        Item_Release_No
        Item_Desc_1
        Item_Desc_2
        'Extra_2
        'Extra_6
        'Item_2

        User_Def_Fld_4 = 104
        Item_Guid            ' 105
        EOF_Marker            ' 106

        'Added June 16, 2017 - T. Louzon
        TravelerRouteID         '107
        TravelerRoute
        RouteCategory
        ItemExtension    'Refers to Route and is added to end of item
        Mfg_Item_No
        IsNoStock
        IsNoArtwork
        Prod_Cat       '114
        End_Item_Cd   'K for Kit  115

        'Added July 18, 2017 - T. Louzon
        Num_Impr_1   '116
        Num_Impr_2
        Num_Impr_3
        Imprint
        Imprint_Location     '120
        Imprint_Color
        Comments            '122
        Special_Comments    '123
        Printer_Instructions    '124
        Packer_Instructions      '125




    End Enum

#Region "Constructors #####################################################"

    Public Sub New()

        Validation = New cOrderValidation
        OrderSource = OrderSourceEnum.OEInterface
        Ordhead = New cOrdHead
        OrdLines = New Collection
        CreditInfo = New cCreditInfo
        Descriptions = New cOrdheadDesc
        UserID = GetNTUserID()

    End Sub

    Public Sub New(ByVal pstrOrd_No As String, ByVal pSource As OrderSourceEnum)

        Try
            Validation = New cOrderValidation
            OrderSource = pSource
            Descriptions = New cOrdheadDesc
            Ordhead = New cOrdHead(pstrOrd_No, Descriptions, pSource)
            OrdLines = New Collection
            m_oOrder = Me
            Call OrderLinesLoad(pSource)
            'CreditInfo = New cCreditInfo(Ordhead.Cus_No, pstrOrd_No)
            CreditInfo = New cCreditInfo(Ordhead.Cus_No)
            UserID = g_User.Usr_ID ' GetNTUserID()
        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

    Public Function OrderInUse(ByVal pstrOEI_Ord_No As String) As String

        OrderInUse = ""
        Try
            Dim dt As DataTable
            Dim db As New cDBA

            Dim strSql As String = _
            "SELECT *   " & _
            "FROM   OEI_ORDHDR " & _
            "WHERE  OEI_Ord_No = '" & SqlCompliantString(pstrOEI_Ord_No) & "' AND " & _
            "       OpenTS IS NOT NULL "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                OrderInUse = dt.Rows(0).Item("OpenTS").ToString
            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Private Sub CreateExactRepeatOrder(ByVal pstrOrd_No As String)

        Try
            Dim dtHeader As New DataTable
            Dim dtLines As New DataTable
            Dim db As New cDBA
            Dim dtRow As DataRow = Nothing

            ' First part - create order header

            Ordhead.OEI_Ord_No = g_User.NextOrderNumber()
            Ordhead.Ord_Dt = Now.Date
            Ordhead.Shipping_Dt = Now.Date
            Ordhead.Extra_6 = Trim(pstrOrd_No).PadLeft(8)

            Dim strSql As String = ""
            strSql = _
            "SELECT * " & _
            "FROM   OEHDRHST_SQL WITH (Nolock) " & _
            "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

            dtHeader = db.DataTable(strSql)

            If dtHeader.Rows.Count = 0 Then

                strSql = _
                "SELECT * " & _
                "FROM   OEORDHDR_SQL WITH (Nolock) " & _
                "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

                dtHeader = db.DataTable(strSql)

            Else
                FromHistory = True
            End If

            If dtHeader.Rows.Count <> 0 Then

                dtRow = dtHeader.Rows(0)

                With Ordhead

                    .Ord_Type = Trim(dtRow.Item("Ord_Type").ToString)
                    .Cus_No = Trim(dtRow.Item("Cus_No").ToString)
                    Call .GetCustomerDefaultValues()

                End With

                Select Case m_oOrder.Ordhead.Customer.cmp_status
                    Case "A", "S" ' Passive and active, ok to proceed
                    Case "B" ' Blocked, show message. On message if not OK, restart order.
                        Dim mbResult As New MsgBoxResult
                        mbResult = MsgBox(New OEExceptionMessage(OEError.Cust_On_Credit_Hold).Message, MsgBoxStyle.YesNo, "Warning.")
                        If mbResult = MsgBoxResult.No Then
                            Throw New OEException(OEError.Cust_On_Credit_Hold, True, False)
                            ' Cancel the order
                            'txtCus_No.Text = String.Empty
                            'txtCus_No.Focus()
                            'Exit Sub
                        Else
                            Ordhead.Hold_Fg = "H"
                        End If
                    Case Else '"E" - treat as inactive
                        Throw New OEException(OEError.Cust_Is_Inactive, True, True)
                        'MsgBox("Customer is inactive.", MsgBoxStyle.OkOnly, "Error.")
                        'txtCus_No.Text = String.Empty
                        'txtCus_No.Focus()
                        'Exit Sub
                End Select
            End If

            'Specific instruction popup (if applicable)
            If m_oOrder.Ordhead.Customer.Specific_Instructions <> "" Then
                MsgBox(New OEExceptionMessage(OEError.Specific_Instructions).Message & ":" & vbCrLf & vbCrLf & m_oOrder.Ordhead.Customer.Specific_Instructions, MsgBoxStyle.OkOnly)
            End If


            If dtHeader.Rows.Count <> 0 Then

                'Dim dtRow As DataRow = dtHeader.Rows(0)
                With Ordhead
                    '.Ord_Type = Trim(dtRow.Item("Ord_Type").ToString)
                    '.Cus_No = Trim(dtRow.Item("Cus_No").ToString)
                    'Call .GetCustomerDefaultValues()
                    '.Cus_Alt_Adr_Cd = Trim(dtRow.Item("Cus_Alt_Adr_Cd").ToString)
                    '.Slspsn_No = Trim(dtRow.Item("Slspsn_No").ToString)
                    '.Slspsn_Pct_Comm = dtRow.Item("Slspsn_Pct_Comm")


                    .User_Def_Fld_4 = Trim(dtRow.Item("User_Def_Fld_4").ToString)
                    .Cus_Email_Address = Trim(dtRow.Item("Email_Address").ToString)
                    '.Mfg_Loc = Trim(dtRow.Item("Mfg_Loc").ToString)
                    '.Ship_Via_Cd = Trim(dtRow.Item("Ship_Via_Cd").ToString)
                    '.Ar_Terms_Cd = Trim(dtRow.Item("Ar_Terms_Cd").ToString)
                    '.Profit_Center = Trim(dtRow.Item("Profit_Center").ToString)
                    '.Dept = Trim(dtRow.Item("Dept").ToString)
                    '.User_Def_Fld_2 = Trim(dtRow.Item("User_Def_Fld_2").ToString)
                    '.Discount_Pct = Trim(dtRow.Item("Discount_Pct").ToString)
                    .User_Def_Fld_1 = Trim(dtRow.Item("User_Def_Fld_1").ToString)
                    '.User_Def_Fld_3 = Trim(dtRow.Item("User_Def_Fld_3").ToString)
                    .User_Def_Fld_5 = Trim(dtRow.Item("User_Def_Fld_5").ToString)
                    .Ship_Instruction_1 = Trim(dtRow.Item("Ship_Instruction_1").ToString)
                    .Ship_Instruction_2 = Trim(dtRow.Item("Ship_Instruction_2").ToString)
                    '.Bill_To_Name = Trim(dtRow.Item("Bill_To_Name").ToString)
                    '.Bill_To_Addr_1 = Trim(dtRow.Item("Bill_To_Addr_1").ToString)
                    '.Bill_To_Addr_2 = Trim(dtRow.Item("Bill_To_Addr_2").ToString)
                    '.Bill_To_Addr_3 = Trim(dtRow.Item("Bill_To_Addr_3").ToString)
                    '.Bill_To_Addr_4 = Trim(dtRow.Item("Bill_To_Addr_4").ToString)
                    '.Bill_To_City = Trim(dtRow.Item("Bill_To_City").ToString)
                    '.Bill_To_State = Trim(dtRow.Item("Bill_To_State").ToString)
                    '.Bill_To_Zip = Trim(dtRow.Item("Bill_To_Zip").ToString)
                    '.Bill_To_Country = Trim(dtRow.Item("Bill_To_Country").ToString)
                    .Ship_To_Name = Trim(dtRow.Item("Ship_To_Name").ToString)
                    .Ship_To_Addr_1 = Trim(dtRow.Item("Ship_To_Addr_1").ToString)
                    .Ship_To_Addr_2 = Trim(dtRow.Item("Ship_To_Addr_2").ToString)
                    .Ship_To_Addr_3 = Trim(dtRow.Item("Ship_To_Addr_3").ToString)
                    .Ship_To_Addr_4 = Trim(dtRow.Item("Ship_To_Addr_4").ToString)
                    .Ship_To_City = Trim(dtRow.Item("Ship_To_City").ToString)
                    .Ship_To_State = Trim(dtRow.Item("Ship_To_State").ToString)
                    .Ship_To_Zip = Trim(dtRow.Item("Ship_To_Zip").ToString)
                    .Ship_To_Country = Trim(dtRow.Item("Ship_To_Country").ToString)

                    '.Curr_Cd = Trim(dtRow.Item("Curr_Cd").ToString)
                    .Orig_Trx_Rt = dtRow.Item("Curr_Trx_Rt")
                    '.Curr_Trx_Rt = dtRow.Item("Curr_Trx_Rt")
                    '.Form_No = dtRow.Item("Form_No")
                    '.Deter_Rate_By = dtRow.Item("Deter_Rate_By")
                    '.Tax_Fg = Trim(dtRow.Item("Tax_Fg").ToString)
                    .Contact_1 = dtRow.Item("Contact_1").ToString.Trim
                    .Phone_Number = Trim(dtRow.Item("Phone_Number").ToString)
                    .Fax_Number = Trim(dtRow.Item("Fax_Number").ToString)
                    .Email_Address = Trim(dtRow.Item("Email_Address").ToString)

                    .RepeatOrd_No = pstrOrd_No.ToString.Trim

                    .Save()

                End With

            End If

            ' second part - create order lines

            If FromHistory Then

                strSql = _
                "SELECT * " & _
                "FROM   OELINHST_SQL WITH (Nolock) " & _
                "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

            Else

                strSql = _
                "SELECT * " & _
                "FROM   OEORDLIN_SQL WITH (Nolock) " & _
                "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

            End If

            dtLines = db.DataTable(strSql)

            If dtLines.Rows.Count <> 0 And dtHeader.Rows.Count <> 0 Then
                For Each dtRow In dtLines.Rows

                    Dim oOrdline As New cOrdLine(Ordhead.Ord_GUID)
                    oOrdline.Source = OEI_SourceEnum.Macola_ExactRepeat
                    oOrdline.Line_Seq_No = dtRow.Item("Line_Seq_No")
                    oOrdline.OEOrdLin_ID = dtRow.Item("Line_Seq_No")
                    oOrdline.Item_No = dtRow.Item("Item_No").ToString
                    oOrdline.Loc = dtRow.Item("Loc").ToString

                    ' Qty to ship is set to qty ordered, because of overpicks.
                    Try
                        oOrdline.Qty_Ordered = dtRow.Item("Qty_Ordered")
                    Catch oe_er As OEException
                        'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                    End Try

                    Try
                        oOrdline.Qty_To_Ship = dtRow.Item("Qty_Ordered") - oOrdline.Qty_Prev_Bkord
                        oOrdline.Qty_To_Ship_Changed(oOrdline.Item_No, oOrdline.Loc, False)
                        'If oOrdline.Qty_Prev_Bkord > 0 Then
                        '    oOrdline.Calculate_ETA()
                        'End If
                    Catch oe_er As OEException
                        'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                    End Try

                    ' Do not enter unit price, let calculate new price on date
                    ' oOrdline.Unit_Price = dtRow.Item("Unit_Price").ToString
                    oOrdline.Discount_Pct = dtRow.Item("Discount_Pct")
                    oOrdline.Request_Dt = Now.Date
                    oOrdline.Promise_Dt = Now.Date
                    oOrdline.Req_Ship_Dt = Now.Date
                    oOrdline.Item_Desc_1 = dtRow.Item("Item_Desc_1").ToString
                    oOrdline.Item_Desc_2 = dtRow.Item("Item_Desc_2").ToString
                    oOrdline.Extra_1 = dtRow.Item("Extra_1").ToString

                    oOrdline.Source = OEI_SourceEnum.frmOrder
                    oOrdline.Save()

                Next dtRow

            End If

            ' Third part - create traveler. We cannot create in lines
            ' because of kit items.
            strSql = "" & _
            "SELECT 	L.*, ISNULL(B.Seq_No, 0) AS Kit_No " & _
            "FROM 		OEI_OrdLin L " & _
            "LEFT JOIN	OEI_OrdBld B WITH (NoLock) ON L.Ord_Guid = B.Ord_Guid AND L.Item_Guid = B.Comp_Item_Guid " & _
            "WHERE      L.Ord_Guid = '" & Trim(Ordhead.Ord_GUID) & "' " & _
            "ORDER BY   L.Line_No, B.Seq_No "

            dtLines = db.DataTable(strSql)

            If dtLines.Rows.Count <> 0 Then

                For Each dtRow In dtLines.Rows

                    'If m_oOrder.OrdLines.Contains(dtRow.Item("Item_Guid").ToString) Then

                    Dim dtItem As DataTable

                    strSql = "" & _
                    "SELECT 	ISNULL(RouteID, 0) AS RouteID " & _
                    "FROM 		Exact_Traveler_Header WITH (Nolock) " & _
                    "WHERE		Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND " & _
                    "   		Line_No = " & dtRow.Item("OEOrdLin_ID") & " AND " & _
                    "           Component_Seq_No = " & dtRow.Item("Kit_No") & " " & _
                    "ORDER BY 	Ord_No, Line_No, Component_Seq_No "

                    dtItem = db.DataTable(strSql)
                    If dtItem.Rows.Count <> 0 Then

                        '************************************* SET ROUTE ID *********************************************** 
                        strSql = _
                        "UPDATE OEI_OrdLin " & _
                        "SET    RouteID = " & dtItem.Rows(0).Item("RouteID") & " " & _
                        "WHERE  Ord_Guid = '" & dtRow.Item("Ord_Guid") & "' AND " & _
                        "       Item_Guid = '" & dtRow.Item("Item_Guid") & "'"

                        db.Execute(strSql)

                        ''==============================
                        ''ADDED June 16, 2017 - T. Louzon ---------------- Will add in stuff later to update the mfg_item_no and RouteExtension
                        ''==============================
                        'strSql =
                        '"UPDATE OEI_OrdLin " &
                        '"SET    RouteID = " & dtItem.Rows(0).Item("RouteID") & " " &
                        '"WHERE  Ord_Guid = '" & dtRow.Item("Ord_Guid") & "' AND " &
                        '"       Item_Guid = '" & dtRow.Item("Item_Guid") & "'"

                        'db.Execute(strSql)



                        '==============================


                        'db.DBCommand.CommandText = strSql
                        'db.DBCommand.ExecuteNonQuery()
                        'db.DBCommand.CommandText = ""

                        'Dim oImprint As New cImprint(Trim(m_oOrder.Ordhead.Ord_GUID), Trim(dtItem.Rows(0).Item("Item_Guid").ToString))
                        ''oImprint.Ord_Type = dtRow.Item("Ord_Type").ToString
                        ''oImprint.OE_Line_No = dtRow.Item("OE_Line_No").ToString
                        ''oImprint.Comp_Seq_No = dtRow.Item("Comp_Seq_No").ToString
                        'oImprint.Item_No = dtRow.Item("Item_No").ToString
                        ''oImprint.Par_Item_No = dtRow.Item("Par_Item_No").ToString
                        'oImprint.Imprint = dtRow.Item("Extra_1").ToString
                        'oImprint.Location = dtRow.Item("Extra_2").ToString
                        'oImprint.Color = dtRow.Item("Extra_3").ToString
                        'If IsNumeric(dtRow.Item("Extra_4").ToString) Then
                        '    oImprint.Num = dtRow.Item("Extra_4").ToString
                        'Else
                        '    oImprint.Num = 0
                        'End If
                        'oImprint.Packaging = dtRow.Item("Extra_5").ToString
                        ''oImprint.Extra_6 = dtRow.Item("Extra_6").ToString
                        ''oImprint.Extra_7 = dtRow.Item("Extra_7").ToString
                        'oImprint.Refill = dtRow.Item("Extra_8").ToString
                        'oImprint.Laser_Setup = dtRow.Item("Extra_9").ToString
                        'oImprint.Industry = dtRow.Item("Industry").ToString
                        'oImprint.Comments = dtRow.Item("Comment").ToString
                        'oImprint.Special_Comments = "REPEAT ORDER # " & Trim(pstrOrd_No) ' dtRow.Item("Comment2").ToString
                        'oImprint.Save()

                    End If
                    'End If

                Next dtRow

            End If


            ' fourt part - create order imprint data. We cannot create in lines
            ' because of kit items.

            'strSql = "" & _
            '"SELECT 	EF.*, ISNULL(L.Line_No, 0) AS OEOrdLin_ID, ISNULL(EF.Comp_Seq_No, 0) AS Kit_No " & _
            '"FROM 		OELinHst_Sql L WITH (Nolock) " & _
            '"LEFT JOIN	Exact_Traveler_Extra_Fields EF WITH (Nolock) ON L.Ord_No = EF.Ord_No AND L.Line_Seq_No = EF.OE_Line_No " & _
            '"LEFT JOIN	IMOrdHst_Sql B WITH (Nolock) ON EF.Ord_No = B.Ord_No AND EF.OE_Line_No = B.Line_No AND EF.Comp_Seq_No = B.Seq_No " & _
            '"WHERE		L.Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' " & _
            '"ORDER BY 	L.Ord_No, L.Line_Seq_No, EF.Comp_Seq_No "

            ' No need to redo, will use the same one as previous query

            'dtLines = db.DataTable(strSql)

            If dtLines.Rows.Count <> 0 Then

                For Each dtRow In dtLines.Rows

                    Dim dtItem As DataTable

                    'strSql = "" & _
                    '"SELECT 	L.* " & _
                    '"FROM 		OEI_OrdLin L WITH (NoLock) " & _
                    '"LEFT JOIN	OEI_OrdBld B WITH (NoLock) ON L.Ord_Guid = B.Ord_Guid AND L.Item_Guid = B.Comp_Item_Guid " & _
                    '"WHERE      L.Ord_Guid = '" & Trim(m_oOrder.Ordhead.Ord_GUID) & "' AND " & _
                    '"           L.Line_Seq_No = " & dtRow.Item("OEOrdLin_ID") & " " & _
                    'IIf(dtRow.Item("Kit_No") = 0, "", " AND B.Seq_No = " & dtRow.Item("Kit_No") & " ") & _
                    '"ORDER BY   L.Line_No, B.Seq_No "

                    strSql = "" & _
                    "SELECT 	*, CASE WHEN ISNULL(Num_Impr_1, 0) = 0 THEN CONVERT(INT, ISNULL(Extra_4, '0')) ELSE ISNULL(Num_Impr_1, 0) END AS Comp_Num_Impr " & _
                    "FROM 		Exact_Traveler_Extra_Fields WITH (Nolock) " & _
                    "WHERE		Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND " & _
                    "   		OE_Line_No = " & dtRow.Item("OEOrdLin_ID") & " AND " & _
                    "           Comp_Seq_No = " & dtRow.Item("Kit_No") & " " & _
                    "ORDER BY 	Ord_No, OE_Line_No, Comp_Seq_No "

                    dtItem = db.DataTable(strSql)
                    If dtItem.Rows.Count <> 0 Then

                        Dim oImprint As New cImprint(Trim(Ordhead.Ord_GUID), Trim(dtRow.Item("Item_Guid").ToString))
                        'oImprint.Ord_Type = dtRow.Item("Ord_Type").ToString
                        'oImprint.OE_Line_No = dtRow.Item("OE_Line_No").ToString
                        'oImprint.Comp_Seq_No = dtRow.Item("Comp_Seq_No").ToString
                        oImprint.SaveToDB = False
                        oImprint.Item_No = dtItem.Rows(0).Item("Item_No").ToString
                        'oImprint.Par_Item_No = dtRow.Item("Par_Item_No").ToString
                        oImprint.Imprint = dtItem.Rows(0).Item("Extra_1").ToString
                        oImprint.Location = dtItem.Rows(0).Item("Extra_2").ToString
                        oImprint.Color = dtItem.Rows(0).Item("Extra_3").ToString
                        If IsNumeric(dtItem.Rows(0).Item("Comp_Num_Impr").ToString) Then
                            oImprint.Num_Impr_1 = dtItem.Rows(0).Item("Comp_Num_Impr")
                        Else
                            oImprint.Num_Impr_1 = 0
                        End If
                        If IsNumeric(dtItem.Rows(0).Item("Num_Impr_2").ToString) Then
                            oImprint.Num_Impr_2 = dtItem.Rows(0).Item("Num_Impr_2")
                        Else
                            oImprint.Num_Impr_2 = 0
                        End If
                        If IsNumeric(dtItem.Rows(0).Item("Num_Impr_3").ToString) Then
                            oImprint.Num_Impr_3 = dtItem.Rows(0).Item("Num_Impr_3")
                        Else
                            oImprint.Num_Impr_3 = 0
                        End If
                        oImprint.Packaging = dtItem.Rows(0).Item("Extra_5").ToString
                        'oImprint.Extra_6 = dtRow.Item("Extra_6").ToString
                        'oImprint.Extra_7 = dtRow.Item("Extra_7").ToString
                        oImprint.Refill = dtItem.Rows(0).Item("Extra_8").ToString
                        oImprint.Laser_Setup = dtItem.Rows(0).Item("Extra_9").ToString
                        oImprint.Industry = dtItem.Rows(0).Item("Industry").ToString
                        oImprint.Comments = dtItem.Rows(0).Item("Comment").ToString
                        oImprint.Repeat_From_Ord_No = dtItem.Rows(0).Item("Ord_No").ToString
                        oImprint.Repeat_From_ID = dtItem.Rows(0).Item("ID")
                        oImprint.Special_Comments = "REPEAT ORDER # " & Trim(pstrOrd_No) ' dtRow.Item("Comment2").ToString
                        oImprint.Printer_Comment = dtItem.Rows(0).Item("Printer_Comment").ToString
                        oImprint.Packer_Comment = dtItem.Rows(0).Item("Packer_Comment").ToString
                        oImprint.Printer_Instructions = dtItem.Rows(0).Item("Printer_Instructions").ToString
                        oImprint.Packer_Instructions = dtItem.Rows(0).Item("Packer_Instructions").ToString
                        oImprint.End_User = dtItem.Rows(0).Item("End_User").ToString
                        g_oOrdline.Set_Spec_Instructions(oImprint.Item_No, oImprint)
                        oImprint.SaveToDB = True

                        If Not oImprint.IsEmpty() Then oImprint.Save()

                    End If

                Next dtRow

            End If

            strSql = "DELETE FROM OEI_Order_Contacts WHERE ORD_GUID = '" & m_oOrder.Ordhead.Ord_GUID & "' "
            db.Execute(strSql)

            strSql = _
            "INSERT INTO OEI_Order_Contacts (Ord_Guid, DefContact, DefMethod, ContactType) " & _
            "SELECT 	'" & m_oOrder.Ordhead.Ord_GUID & "', DefContact, DefMethod, ContactType  " & _
            "FROM 		Exact_Traveler_Order_Customer_Contacts_New " & _
            "WHERE 		OrderNo = '" & Trim(pstrOrd_No) & "'"
            db.Execute(strSql)

            m_oOrder.FromHistory = False

        Catch oe_er As OEException
            Throw oe_er ' New OEException(oe_er.Message, oe_er)
        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Shared Sub sendEDIAbortedEmail(abortType As String)

        Try

            Dim strSql As String
            Dim db As New cDBA
            Dim dtEdi As New DataTable
            Dim dtRow As DataRow = Nothing
            'error fields
            Dim eFields() As String

            Dim oXmlFile As New XmlDocument()
            Dim mailer = mEmailer.CreateEmailConnection()
            Dim oe_po_no As String
            Dim contact As String

            Dim oNode As XmlNode

            'layout = EdiFailureCd - EdiFailureDesc - strCus_No - piOrdEDI_ID.ToString()
            eFields = abortType.Split("-")

            strSql = _
                 "SELECT Xml_Text " & _
                 "FROM   OEI_OrdEDI WITH (Nolock) " & _
                 "WHERE  OrdEDI_ID = " & eFields(3).ToString() & " AND " & _
                 "       OEI_Status = 'R' "

            dtEdi = db.DataTable(strSql)

            'Get OE_PO_NO and contact from XML file
            If dtEdi.Rows.Count <> 0 Then
                oXmlFile.LoadXml(dtEdi.Rows(0).Item("Xml_Text").ToString)
                oNode = oXmlFile.SelectSingleNode("/ExactRequestSet/MacolaSalesOrderAddRequest/OrderHeader")

                If Not oNode.SelectSingleNode("oe_po_no") Is Nothing And oNode.SelectSingleNode("oe_po_no").InnerText <> "" Then
                    oe_po_no = oNode.SelectSingleNode("oe_po_no").InnerText.ToString.Trim()
                Else
                    oe_po_no = "N/A"
                End If

                If Not oNode.SelectSingleNode("contact_1") Is Nothing And oNode.SelectSingleNode("contact_1").InnerText <> "" Then
                    contact = oNode.SelectSingleNode("contact_1").InnerText.ToString.Trim()
                Else
                    oe_po_no = "N/A"
                    contact = "--"
                End If

            Else
                oe_po_no = "N/A"
                contact = "--"
            End If


            'Create and send email

            mailer.Message.FromAddr = "Spectorandco EDI Service <it@spectorandco.ca>"
            mailer.Message.ToAddr = "sharon@spectorandco.com" '"andrew@spectorandco.ca"
            ' mailer.Message.CCAddr = "andrew@spectorandco.ca" '"pinky@spectorandco.ca, leni@spectorandco.ca"
            mailer.Message.BCCAddr = "iond@spectorandco.com" '"itdept@spectorandco.ca"

            mailer.Message.Subject = "Aborted/Inocrrect EDI Order for P.O. " & oe_po_no
            mailer.Message.BodyText = "EDI order #" & eFields(3) & " has not been imported due to errors found in the entry." & vbCrLf & vbCrLf & _
                                      "EDI Failure Code: <b>" & eFields(0) & vbCrLf & "</b>" & _
                                      "OE_PO_NO: <b>" & oe_po_no & vbCrLf & "</b>" & _
                                      "Customer Number: <b>" & eFields(2) & vbCrLf & "</b>" & _
                                      "Customer Contact: <b>" & contact & vbCrLf & "</b>" & _
                                      "Reason: <b>" & eFields(1) & vbCrLf & vbCrLf & "</b>" & _
                                      "Please contact the distributor that sent this EDI order so that they can resend it if applicable." & _
                                      "" & vbCrLf & vbCrLf & _
                                      "Thanks, " & vbCrLf & vbCrLf & "OEI System"

            mailer.Message.BodyText = Replace(mailer.Message.BodyText, vbCrLf, "<BR>")

            mailer.Send()

            'close connection
            mailer = Nothing

            'Update edi status and show error message
            Call updateEDIOrderStatus("B", eFields(3).ToString())
            MsgBox("EDI upload failure" & vbCrLf & "Base reason: " & eFields(1) & vbCrLf & vbCrLf & "EDI order has been removed and an email has been sent to the Account manager.")

            'Reset form
            For Each f As Form In My.Application.OpenForms()
                If TypeOf f Is frmOrder Then
                    Dim topForm As frmOrder = f
                    topForm.tsbNew.PerformClick()
                End If
            Next

        Catch er As Exception
            MsgBox("Error in cOrder->sendEDIAbortedEmail." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Shared Sub updateEDIOrderStatus(edi_status As Char, ordEDI_ID As Integer)

        Dim strSql As String
        Dim db As New cDBA

        Try
            strSql = _
                 "UPDATE OEI_OrdEDI " & _
                 "SET  OEI_Status = '" & edi_status & "'" & _
                 "WHERE  OrdEDI_ID = " & ordEDI_ID

            db.Execute(strSql)

        Catch er As Exception
            MsgBox("Error in cOrder->updateEDIOrderStatus." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub CreateEDIOrder(ByVal pstrOrd_No As String)

        Try
            Dim dtHeader As New DataTable
            Dim dtLines As New DataTable
            Dim db As New cDBA
            Dim dtRow As DataRow = Nothing

            ' First part - create order header

            Ordhead.OEI_Ord_No = g_User.NextOrderNumber()
            Ordhead.Ord_Dt = Now.Date
            Ordhead.Shipping_Dt = Now.Date
            'Ordhead.Extra_6 = Trim(pstrOrd_No).PadLeft(8)

            Dim strSql As String = ""
            strSql = _
            "SELECT * " & _
            "FROM   OEHDRHST_SQL WITH (Nolock) " & _
            "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

            dtHeader = db.DataTable(strSql)

            If dtHeader.Rows.Count = 0 Then

                strSql = _
                "SELECT * " & _
                "FROM   OEORDHDR_SQL WITH (Nolock) " & _
                "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

                dtHeader = db.DataTable(strSql)

            Else
                FromHistory = True
            End If

            If dtHeader.Rows.Count <> 0 Then

                dtRow = dtHeader.Rows(0)

                With Ordhead

                    .Ord_Type = Trim(dtRow.Item("Ord_Type").ToString)
                    .Cus_No = Trim(dtRow.Item("Cus_No").ToString)
                    Call .GetCustomerDefaultValues()

                End With

                Select Case m_oOrder.Ordhead.Customer.cmp_status
                    Case "A", "S" ' Passive and active, ok to proceed
                    Case "B" ' Blocked, show message. On message if not OK, restart order.
                        Dim mbResult As New MsgBoxResult
                        mbResult = MsgBox(New OEExceptionMessage(OEError.Cust_On_Credit_Hold).Message, MsgBoxStyle.YesNo, "Warning.")
                        If mbResult = MsgBoxResult.No Then
                            Throw New OEException(OEError.Cust_On_Credit_Hold, True, False)
                            ' Cancel the order
                            'txtCus_No.Text = String.Empty
                            'txtCus_No.Focus()
                            'Exit Sub
                        Else
                            Ordhead.Hold_Fg = "H"
                        End If
                    Case Else '"E" - treat as inactive
                        Throw New OEException(OEError.Cust_Is_Inactive, True, True)
                        'MsgBox("Customer is inactive.", MsgBoxStyle.OkOnly, "Error.")
                        'txtCus_No.Text = String.Empty
                        'txtCus_No.Focus()
                        'Exit Sub
                End Select
            End If

            'Specific instruction popup (if applicable)
            If m_oOrder.Ordhead.Customer.Specific_Instructions <> "" Then
                MsgBox(New OEExceptionMessage(OEError.Specific_Instructions).Message & ":" & vbCrLf & vbCrLf & m_oOrder.Ordhead.Customer.Specific_Instructions, MsgBoxStyle.OkOnly)
            End If

            If dtHeader.Rows.Count <> 0 Then

                'Dim dtRow As DataRow = dtHeader.Rows(0)
                With Ordhead

                    '.Ord_Type = Trim(dtRow.Item("Ord_Type").ToString)
                    '.Cus_No = Trim(dtRow.Item("Cus_No").ToString)
                    'Call .GetCustomerDefaultValues()
                    '.Cus_Alt_Adr_Cd = Trim(dtRow.Item("Cus_Alt_Adr_Cd").ToString)
                    '.Slspsn_No = Trim(dtRow.Item("Slspsn_No").ToString)
                    '.Slspsn_Pct_Comm = dtRow.Item("Slspsn_Pct_Comm")


                    .User_Def_Fld_4 = Trim(dtRow.Item("User_Def_Fld_4").ToString)
                    .Cus_Email_Address = Trim(dtRow.Item("Email_Address").ToString)
                    '.Mfg_Loc = Trim(dtRow.Item("Mfg_Loc").ToString)
                    .Ship_Via_Cd = Trim(dtRow.Item("Ship_Via_Cd").ToString)
                    .Ar_Terms_Cd = Trim(dtRow.Item("Ar_Terms_Cd").ToString)
                    .Profit_Center = Trim(dtRow.Item("Profit_Center").ToString)
                    .Dept = Trim(dtRow.Item("Dept").ToString)
                    '.User_Def_Fld_2 = Trim(dtRow.Item("User_Def_Fld_2").ToString)
                    '.Discount_Pct = Trim(dtRow.Item("Discount_Pct").ToString)
                    .User_Def_Fld_1 = Trim(dtRow.Item("User_Def_Fld_1").ToString)
                    '.User_Def_Fld_3 = Trim(dtRow.Item("User_Def_Fld_3").ToString)
                    .User_Def_Fld_5 = Trim(dtRow.Item("User_Def_Fld_5").ToString)
                    .Ship_Instruction_1 = Trim(dtRow.Item("Ship_Instruction_1").ToString)
                    .Ship_Instruction_2 = Trim(dtRow.Item("Ship_Instruction_2").ToString)
                    '.Bill_To_Name = Trim(dtRow.Item("Bill_To_Name").ToString)
                    '.Bill_To_Addr_1 = Trim(dtRow.Item("Bill_To_Addr_1").ToString)
                    '.Bill_To_Addr_2 = Trim(dtRow.Item("Bill_To_Addr_2").ToString)
                    '.Bill_To_Addr_3 = Trim(dtRow.Item("Bill_To_Addr_3").ToString)
                    '.Bill_To_Addr_4 = Trim(dtRow.Item("Bill_To_Addr_4").ToString)
                    '.Bill_To_City = Trim(dtRow.Item("Bill_To_City").ToString)
                    '.Bill_To_State = Trim(dtRow.Item("Bill_To_State").ToString)
                    '.Bill_To_Zip = Trim(dtRow.Item("Bill_To_Zip").ToString)
                    '.Bill_To_Country = Trim(dtRow.Item("Bill_To_Country").ToString)
                    .Ship_To_Name = Trim(dtRow.Item("Ship_To_Name").ToString)
                    .Ship_To_Addr_1 = Trim(dtRow.Item("Ship_To_Addr_1").ToString)
                    .Ship_To_Addr_2 = Trim(dtRow.Item("Ship_To_Addr_2").ToString)
                    .Ship_To_Addr_3 = Trim(dtRow.Item("Ship_To_Addr_3").ToString)
                    .Ship_To_Addr_4 = Trim(dtRow.Item("Ship_To_Addr_4").ToString)
                    .Ship_To_City = Trim(dtRow.Item("Ship_To_City").ToString)
                    .Ship_To_State = Trim(dtRow.Item("Ship_To_State").ToString)
                    .Ship_To_Zip = Trim(dtRow.Item("Ship_To_Zip").ToString)
                    .Ship_To_Country = Trim(dtRow.Item("Ship_To_Country").ToString)

                    '.Curr_Cd = Trim(dtRow.Item("Curr_Cd").ToString)
                    .Orig_Trx_Rt = dtRow.Item("Curr_Trx_Rt")
                    '.Curr_Trx_Rt = dtRow.Item("Curr_Trx_Rt")
                    '.Form_No = dtRow.Item("Form_No")
                    '.Deter_Rate_By = dtRow.Item("Deter_Rate_By")
                    '.Tax_Fg = Trim(dtRow.Item("Tax_Fg").ToString)
                    .Contact_1 = Trim(dtRow.Item("Contact_1").ToString)
                    .Phone_Number = Trim(dtRow.Item("Phone_Number").ToString)
                    .Fax_Number = Trim(dtRow.Item("Fax_Number").ToString)
                    .Email_Address = Trim(dtRow.Item("Email_Address").ToString)

                    .Save()

                End With

            End If

            ' second part - create order lines

            If FromHistory Then

                strSql = _
                "SELECT * " & _
                "FROM   OELINHST_SQL WITH (Nolock) " & _
                "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

            Else

                strSql = _
                "SELECT * " & _
                "FROM   OEORDLIN_SQL WITH (Nolock) " & _
                "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

            End If

            dtLines = db.DataTable(strSql)

            If dtLines.Rows.Count <> 0 And dtHeader.Rows.Count <> 0 Then
                For Each dtRow In dtLines.Rows

                    Dim oOrdline As New cOrdLine(Ordhead.Ord_GUID)
                    oOrdline.Source = OEI_SourceEnum.Macola_ExactRepeat
                    oOrdline.Line_Seq_No = dtRow.Item("Line_Seq_No")
                    oOrdline.OEOrdLin_ID = dtRow.Item("Line_Seq_No")
                    oOrdline.Item_No = dtRow.Item("Item_No").ToString
                    oOrdline.Loc = dtRow.Item("Loc").ToString

                    ' Qty to ship is set to qty ordered, because of overpicks.
                    Try
                        oOrdline.Qty_Ordered = dtRow.Item("Qty_Ordered")
                    Catch oe_er As OEException
                        'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                    End Try

                    Try
                        oOrdline.Qty_To_Ship = dtRow.Item("Qty_Ordered")
                        oOrdline.Qty_To_Ship_Changed(oOrdline.Item_No, oOrdline.Loc, False)
                        'If oOrdline.Qty_Prev_Bkord > 0 Then
                        '    oOrdline.Calculate_ETA()
                        'End If
                    Catch oe_er As OEException
                        'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                    End Try

                    ' Do not enter unit price, let calculate new price on date
                    ' oOrdline.Unit_Price = dtRow.Item("Unit_Price").ToString
                    oOrdline.Discount_Pct = dtRow.Item("Discount_Pct")
                    oOrdline.Request_Dt = Now.Date
                    oOrdline.Promise_Dt = Now.Date
                    oOrdline.Req_Ship_Dt = Now.Date
                    oOrdline.Item_Desc_1 = dtRow.Item("Item_Desc_1").ToString
                    oOrdline.Item_Desc_2 = dtRow.Item("Item_Desc_2").ToString
                    oOrdline.Extra_1 = dtRow.Item("Extra_1").ToString

                    oOrdline.Source = OEI_SourceEnum.frmOrder
                    oOrdline.Save()

                Next dtRow

            End If

            ' Third part - create traveler. We cannot create in lines
            ' because of kit items.
            strSql = "" & _
            "SELECT 	L.*, ISNULL(B.Seq_No, 0) AS Kit_No " & _
            "FROM 		OEI_OrdLin L " & _
            "LEFT JOIN	OEI_OrdBld B WITH (NoLock) ON L.Ord_Guid = B.Ord_Guid AND L.Item_Guid = B.Comp_Item_Guid " & _
            "WHERE      L.Ord_Guid = '" & Trim(Ordhead.Ord_GUID) & "' " & _
            "ORDER BY   L.Line_No, B.Seq_No "

            dtLines = db.DataTable(strSql)

            If dtLines.Rows.Count <> 0 Then

                For Each dtRow In dtLines.Rows

                    'If m_oOrder.OrdLines.Contains(dtRow.Item("Item_Guid").ToString) Then

                    Dim dtItem As DataTable

                    strSql = "" & _
                    "SELECT 	ISNULL(RouteID, 0) AS RouteID " & _
                    "FROM 		Exact_Traveler_Header WITH (Nolock) " & _
                    "WHERE		Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND " & _
                    "   		Line_No = " & dtRow.Item("OEOrdLin_ID") & " AND " & _
                    "           Component_Seq_No = " & dtRow.Item("Kit_No") & " " & _
                    "ORDER BY 	Ord_No, Line_No, Component_Seq_No "

                    dtItem = db.DataTable(strSql)
                    If dtItem.Rows.Count <> 0 Then
                        '******************************************************************** SET ROUTE ID *******************
                        strSql = _
                        "UPDATE OEI_OrdLin " & _
                        "SET    RouteID = " & dtItem.Rows(0).Item("RouteID") & " " & _
                        "WHERE  Ord_Guid = '" & dtRow.Item("Ord_Guid") & "' AND " & _
                        "       Item_Guid = '" & dtRow.Item("Item_Guid") & "'"

                        db.Execute(strSql)
                        'db.DBCommand.CommandText = strSql
                        'db.DBCommand.ExecuteNonQuery()
                        'db.DBCommand.CommandText = ""

                        'Dim oImprint As New cImprint(Trim(m_oOrder.Ordhead.Ord_GUID), Trim(dtItem.Rows(0).Item("Item_Guid").ToString))
                        ''oImprint.Ord_Type = dtRow.Item("Ord_Type").ToString
                        ''oImprint.OE_Line_No = dtRow.Item("OE_Line_No").ToString
                        ''oImprint.Comp_Seq_No = dtRow.Item("Comp_Seq_No").ToString
                        'oImprint.Item_No = dtRow.Item("Item_No").ToString
                        ''oImprint.Par_Item_No = dtRow.Item("Par_Item_No").ToString
                        'oImprint.Imprint = dtRow.Item("Extra_1").ToString
                        'oImprint.Location = dtRow.Item("Extra_2").ToString
                        'oImprint.Color = dtRow.Item("Extra_3").ToString
                        'If IsNumeric(dtRow.Item("Extra_4").ToString) Then
                        '    oImprint.Num = dtRow.Item("Extra_4").ToString
                        'Else
                        '    oImprint.Num = 0
                        'End If
                        'oImprint.Packaging = dtRow.Item("Extra_5").ToString
                        ''oImprint.Extra_6 = dtRow.Item("Extra_6").ToString
                        ''oImprint.Extra_7 = dtRow.Item("Extra_7").ToString
                        'oImprint.Refill = dtRow.Item("Extra_8").ToString
                        'oImprint.Laser_Setup = dtRow.Item("Extra_9").ToString
                        'oImprint.Industry = dtRow.Item("Industry").ToString
                        'oImprint.Comments = dtRow.Item("Comment").ToString
                        'oImprint.Special_Comments = "REPEAT ORDER # " & Trim(pstrOrd_No) ' dtRow.Item("Comment2").ToString
                        'oImprint.Save()

                    End If
                    'End If

                Next dtRow

            End If

            ' fourth part - create order imprint data. We cannot create in lines
            ' because of kit items.

            'strSql = "" & _
            '"SELECT 	EF.*, ISNULL(L.Line_No, 0) AS OEOrdLin_ID, ISNULL(EF.Comp_Seq_No, 0) AS Kit_No " & _
            '"FROM 		OELinHst_Sql L WITH (Nolock) " & _
            '"LEFT JOIN	Exact_Traveler_Extra_Fields EF WITH (Nolock) ON L.Ord_No = EF.Ord_No AND L.Line_Seq_No = EF.OE_Line_No " & _
            '"LEFT JOIN	IMOrdHst_Sql B WITH (Nolock) ON EF.Ord_No = B.Ord_No AND EF.OE_Line_No = B.Line_No AND EF.Comp_Seq_No = B.Seq_No " & _
            '"WHERE		L.Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' " & _
            '"ORDER BY 	L.Ord_No, L.Line_Seq_No, EF.Comp_Seq_No "

            ' No need to redo, will use the same one as previous query

            'dtLines = db.DataTable(strSql)

            If dtLines.Rows.Count <> 0 Then

                For Each dtRow In dtLines.Rows

                    Dim dtItem As DataTable

                    'strSql = "" & _
                    '"SELECT 	L.* " & _
                    '"FROM 		OEI_OrdLin L WITH (NoLock) " & _
                    '"LEFT JOIN	OEI_OrdBld B WITH (NoLock) ON L.Ord_Guid = B.Ord_Guid AND L.Item_Guid = B.Comp_Item_Guid " & _
                    '"WHERE      L.Ord_Guid = '" & Trim(m_oOrder.Ordhead.Ord_GUID) & "' AND " & _
                    '"           L.Line_Seq_No = " & dtRow.Item("OEOrdLin_ID") & " " & _
                    'IIf(dtRow.Item("Kit_No") = 0, "", " AND B.Seq_No = " & dtRow.Item("Kit_No") & " ") & _
                    '"ORDER BY   L.Line_No, B.Seq_No "

                    strSql = "" & _
                    "SELECT 	*, CASE WHEN ISNULL(Num_Impr_1, 0) = 0 THEN CONVERT(INT, ISNULL(Extra_4, '0')) ELSE ISNULL(Num_Impr_1, 0) END AS Comp_Num_Impr " & _
                    "FROM 		Exact_Traveler_Extra_Fields WITH (Nolock) " & _
                    "WHERE		Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND " & _
                    "   		OE_Line_No = " & dtRow.Item("OEOrdLin_ID") & " AND " & _
                    "           Comp_Seq_No = " & dtRow.Item("Kit_No") & " " & _
                    "ORDER BY 	Ord_No, OE_Line_No, Comp_Seq_No "

                    dtItem = db.DataTable(strSql)
                    If dtItem.Rows.Count <> 0 Then

                        Dim oImprint As New cImprint(Trim(Ordhead.Ord_GUID), Trim(dtRow.Item("Item_Guid").ToString))
                        'oImprint.Ord_Type = dtRow.Item("Ord_Type").ToString
                        'oImprint.OE_Line_No = dtRow.Item("OE_Line_No").ToString
                        'oImprint.Comp_Seq_No = dtRow.Item("Comp_Seq_No").ToString
                        oImprint.SaveToDB = False
                        oImprint.Item_No = dtItem.Rows(0).Item("Item_No").ToString
                        'oImprint.Par_Item_No = dtRow.Item("Par_Item_No").ToString
                        oImprint.Imprint = dtItem.Rows(0).Item("Extra_1").ToString
                        oImprint.Location = dtItem.Rows(0).Item("Extra_2").ToString
                        oImprint.Color = dtItem.Rows(0).Item("Extra_3").ToString
                        If IsNumeric(dtItem.Rows(0).Item("Comp_Num_Impr").ToString) Then
                            oImprint.Num_Impr_1 = dtItem.Rows(0).Item("Comp_Num_Impr")
                        Else
                            oImprint.Num_Impr_1 = 0
                        End If
                        If IsNumeric(dtItem.Rows(0).Item("Num_Impr_2").ToString) Then
                            oImprint.Num_Impr_2 = dtItem.Rows(0).Item("Num_Impr_2")
                        Else
                            oImprint.Num_Impr_2 = 0
                        End If
                        If IsNumeric(dtItem.Rows(0).Item("Num_Impr_3").ToString) Then
                            oImprint.Num_Impr_3 = dtItem.Rows(0).Item("Num_Impr_3")
                        Else
                            oImprint.Num_Impr_3 = 0
                        End If
                        oImprint.Packaging = dtItem.Rows(0).Item("Extra_5").ToString
                        'oImprint.Extra_6 = dtRow.Item("Extra_6").ToString
                        'oImprint.Extra_7 = dtRow.Item("Extra_7").ToString
                        oImprint.Refill = dtItem.Rows(0).Item("Extra_8").ToString
                        oImprint.Laser_Setup = dtItem.Rows(0).Item("Extra_9").ToString
                        oImprint.Industry = dtItem.Rows(0).Item("Industry").ToString
                        oImprint.Comments = dtItem.Rows(0).Item("Comment").ToString
                        oImprint.Repeat_From_Ord_No = dtItem.Rows(0).Item("Ord_No").ToString
                        oImprint.Repeat_From_ID = dtItem.Rows(0).Item("ID")
                        oImprint.Special_Comments = "REPEAT ORDER # " & Trim(pstrOrd_No) ' dtRow.Item("Comment2").ToString

                        oImprint.SaveToDB = True

                        If Not oImprint.IsEmpty() Then oImprint.Save()

                    End If

                Next dtRow

            End If

            strSql = "DELETE FROM OEI_Order_Contacts WHERE ORD_GUID = '" & m_oOrder.Ordhead.Ord_GUID & "' "
            db.Execute(strSql)

            strSql = _
            "INSERT INTO OEI_Order_Contacts (Ord_Guid, DefContact, DefMethod, ContactType) " & _
            "SELECT 	'" & m_oOrder.Ordhead.Ord_GUID & "', DefContact, DefMethod, ContactType  " & _
            "FROM 		Exact_Traveler_Order_Customer_Contacts_New " & _
            "WHERE 		OrderNo = '" & Trim(pstrOrd_No) & "'"
            db.Execute(strSql)

            m_oOrder.FromHistory = False

        Catch oe_er As OEException
            Throw oe_er ' New OEException(oe_er.Message, oe_er)
        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub OrderLinesLoad(ByVal pSource As OrderSourceEnum)

        Dim strSql As String = ""

        Try
            If pSource = OrderSourceEnum.Macola Then

                strSql =
                "SELECT     '' AS Item_Guid,  ISNULL(O.Line_No, 0) AS Line_No, " &
                "           ISNULL(O.Item_No, '') AS Item_No, " &
                "           ISNULL(O.Loc, '') AS Loc, ISNULL(O.Qty_Ordered, 0) AS Qty_Ordered, " &
                "           ISNULL(O.Qty_To_Ship, 0) AS Qty_To_Ship, ISNULL(O.Ord_Type, '') AS Ord_Type, " &
                "           ISNULL(O.Line_Seq_No, 0) AS Line_Seq_No, ISNULL(O.Pick_Seq, '') AS Pick_Seq, " &
                "           ISNULL(O.Ord_No, '') AS Ord_No, ISNULL(O.Item_Desc_1, '') AS Item_Desc_1, " &
                "           ISNULL(O.Item_Desc_2, '') AS Item_Desc_2, ISNULL(O.Unit_Price, 0) AS Unit_Price, " &
                "           ISNULL(O.Discount_Pct, 0) AS Discount_Pct, ISNULL(O.Request_Dt, 0) AS Request_Dt, " &
                "           ISNULL(O.Promise_Dt, 0) AS Promise_Dt, ISNULL(O.Select_Cd, 0) AS Select_Cd , " &
                "           ISNULL(O.Req_Ship_Dt, 0) AS Req_Ship_Dt, ISNULL(O.User_Def_Fld_1, '') AS User_Def_Fld_1, " &
                "           ISNULL(O.User_Def_Fld_2, '') AS User_Def_Fld_2, ISNULL(O.User_Def_Fld_3, '') AS User_Def_Fld_3, " &
                "           ISNULL(O.User_Def_Fld_4, '') AS User_Def_Fld_4, ISNULL(O.User_Def_Fld_5, '') AS User_Def_Fld_5, " &
                "           ISNULL(O.Qty_Bkord_Fg, '') AS Qty_Bkord_Fg , ISNULL(O.Po_Ord_No, '') AS Po_Ord_No, " &
                "           ISNULL(O.Rma_Seq, 0) AS Rma_Seq, ISNULL(O.Recalc_Sw, '') AS Recalc_Sw, ISNULL(O.Orig_Price, 0) AS Orig_Price, " &
                "           -1 AS Seq_No, '' AS Kit_Item_Guid, '' AS Comp_Item_Guid, " &
                "           0 AS Comp_Unit_Price, ISNULL(O.Unit_Cost, 0) AS Unit_Cost, 0 AS Comp_Unit_Cost, " &
                "           P.Picture AS Image, ISNULL(L.Qty_On_Hand,0) AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(I.Item_No, L.Loc, 'X'),0) AS Qty_Allocated, CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(O.Item_No, O.Loc, 'X'),0)) AS FLOAT) AS Qty_On_Hand, " &
                "           ISNULL(I.UOM, '') AS UOM, ISNULL(I.BkOrd_Fg, 'N') AS BkOrd_Fg, ISNULL(O.End_Item_Cd, ' ') AS End_Item_Cd, " &
                "           '' AS Route, 0 AS RouteID, ISNULL(O.Unit_Weight, 0) AS Unit_Weight, ISNULL(O.Prc_Cd_Orig_Price, 0) AS Prc_Cd_Orig_Price, " &
                "           ISNULL(I.Activity_Cd, '') AS Activity_Cd, ISNULL(O.Prod_Cat, '') AS Prod_Cat, " &
                "           ISNULL(P.OEI_W_Pixels, 400) AS OEI_W_Pixels, ISNULL(P.OEI_H_Pixels, 400) AS OEI_H_Pixels, " &
                "           ISNULL(O.Extra_9, '') AS Extra_9, ISNULL(O.Extra_10, 0) AS Extra_10, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd, 0 as Mix_Group " ' &
                '"FROM       " & Ordhead.Lines_Table & " O WITH (NOLOCK) " &
                '"INNER JOIN IMITMIDX_SQL I WITH (nolock) ON O.Item_No = I.Item_No " &
                '"LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No AND O.Loc = L.Loc " &
                '"LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                '"WHERE      O.Ord_No = '" & Ordhead.Ord_No & "' AND I.ACTIVITY_CD = 'A' " &
                '"ORDER BY   O.Line_No " ' , K.Seq_No "



                If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then

                    strSql &= " " &
                         "FROM       " & Ordhead.Lines_Table & " O WITH (NOLOCK) " &
                "INNER JOIN IMITMIDX_SQL I WITH (nolock) ON O.Item_No = I.Item_No " &
                "LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No AND O.Loc = L.Loc " &
                "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                "WHERE      O.Ord_No = '" & Ordhead.Ord_No & "' AND I.ACTIVITY_CD = 'A' " &
                "ORDER BY   O.Line_No " ' , K.Seq_No "

                Else
                    strSql &= " " &
                         "FROM   " & Ordhead.Lines_Table & " O WITH (NOLOCK) " &
                "INNER JOIN [200].dbo.IMITMIDX_SQL I WITH (nolock) ON O.Item_No = I.Item_No " &
                "LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No AND O.Loc = L.Loc " &
                "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                "WHERE      O.Ord_No = '" & Ordhead.Ord_No & "' AND I.ACTIVITY_CD = 'A' " &
                "ORDER BY   O.Line_No " ' , K.Seq_No "

                End If







                '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON O.Item_No = L.Item_No AND O.Loc = L.Loc " & _

                ' "LEFT JOIN  OEI_OrdBld K WITH (Nolock) ON O.Item_Guid = K.Comp_Item_Guid " & _
                ' "LEFT JOIN  Exact_Traveler_Route R WITH (Nolock) ON O.RouteID = R.RouteID " & _

            ElseIf pSource = OrderSourceEnum.OEInterface Then

                strSql =
                "SELECT     ISNULL(O.Item_Guid, '') AS Item_Guid,  ISNULL(O.Line_No, 0) AS Line_No, " &
                "           ISNULL(O.Item_No, '') AS Item_No, " &
                "           ISNULL(O.Loc, '') AS Loc, ISNULL(O.Qty_Ordered, 0) AS Qty_Ordered, " &
                "           ISNULL(O.Qty_To_Ship, 0) AS Qty_To_Ship, ISNULL(O.Ord_Type, '') AS Ord_Type, " &
                "           ISNULL(O.Line_Seq_No, 0) AS Line_Seq_No, ISNULL(O.Pick_Seq, '') AS Pick_Seq, " &
                "           ISNULL(O.Ord_No, '') AS Ord_No, ISNULL(O.Item_Desc_1, '') AS Item_Desc_1, "

                '===================================
                'Modified June 6, 2017 - T. Louzon
                'Added Unit_Price_BeforeSpecial 
                '===================================
                ''''''"           ISNULL(O.Item_Desc_2, '') AS Item_Desc_2, ISNULL(O.Unit_Price, 0) AS Unit_Price, " &  
                strSql = strSql &
                "           ISNULL(O.Item_Desc_2, '') AS Item_Desc_2, " &
                "           ISNULL(O.Unit_Price, 0) AS Unit_Price, " &
                "            ISNULL(O.Unit_Price_BeforeSpecial, 0) AS Unit_Price_BeforeSpecial, " &
                "            ISNULL(O.mfg_Item_no, 0) AS Mfg_Item_no, " &
                "            ISNULL(O.RouteExtension, 0) AS RouteExtension, " &
                "            ISNULL(O.IsNoArtwork, 0) AS IsNoArtwork, " &
                "            ISNULL(O.IsNoStock, 0) AS IsNoStock, "
                '===================================

                strSql &= "  " &
                "           ISNULL(O.Discount_Pct, 0) AS Discount_Pct, ISNULL(O.Request_Dt, 0) AS Request_Dt, " &
                "           ISNULL(O.Promise_Dt, 0) AS Promise_Dt, ISNULL(O.Select_Cd, 0) AS Select_Cd , " &
                "           ISNULL(O.Req_Ship_Dt, 0) AS Req_Ship_Dt, ISNULL(O.User_Def_Fld_1, '') AS User_Def_Fld_1, " &
                "           ISNULL(O.User_Def_Fld_2, '') AS User_Def_Fld_2, ISNULL(O.User_Def_Fld_3, '') AS User_Def_Fld_3, " &
                "           ISNULL(O.User_Def_Fld_4, '') AS User_Def_Fld_4, ISNULL(O.User_Def_Fld_5, '') AS User_Def_Fld_5, " &
                "           ISNULL(O.Qty_Bkord_Fg, '') AS Qty_Bkord_Fg , ISNULL(O.Po_Ord_No, '') AS Po_Ord_No, " &
                "           ISNULL(O.Rma_Seq, 0) AS Rma_Seq, ISNULL(O.Recalc_Sw, '') AS Recalc_Sw, ISNULL(O.Orig_Price, 0) AS Orig_Price, " &
                "           ISNULL(K.Seq_No, -1) AS Seq_No, ISNULL(K.Kit_Item_Guid,'') AS Kit_Item_Guid, ISNULL(K.Comp_Item_Guid,'') AS Comp_Item_Guid, " &
                "           ISNULL(K.Unit_Price, 0) AS Comp_Unit_Price, ISNULL(O.Unit_Cost, 0) AS Unit_Cost, ISNULL(K.Unit_Cost, 0) AS Comp_Unit_Cost, " &
                "           P.Picture AS Image, ISNULL(L.Qty_On_Hand,0) AS Qty_Inventory, ISNULL(L.Qty_Allocated,0) + ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(O.Item_No, O.Loc, O.Item_Guid),0) AS Qty_Allocated, " &
                "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(DBO.OEI_Pending_Qty_Ordered_Item(O.Item_No, O.Loc, O.Item_Guid),0)) AS FLOAT) AS Qty_On_Hand, " &
                "           ISNULL(I.UOM, '') AS UOM, ISNULL(O.BkOrd_Fg, 'N') AS BkOrd_Fg, ISNULL(O.Stocked_Fg, 'N') AS Stocked_Fg, ISNULL(O.End_Item_Cd, ' ') AS End_Item_Cd, " &
                "           ISNULL(R.RouteDescription, '') AS Route, ISNULL(O.RouteID, 0) AS RouteID, ISNULL(O.Unit_Weight, 0) AS Unit_Weight, ISNULL(O.Prc_Cd_Orig_Price, 0) AS Prc_Cd_Orig_Price, " &
                "           ISNULL(I.Activity_Cd, '') AS Activity_Cd, ISNULL(O.Prod_Cat, '') AS Prod_Cat, ISNULL(O.ProductProof, 0) AS ProductProof, ISNULL(O.AutoCompleteReship, 0) AS AutoCompleteReship, " &
                "           ISNULL(P.OEI_W_Pixels, 400) AS OEI_W_Pixels, ISNULL(P.OEI_H_Pixels, 400) AS OEI_H_Pixels,  " &
                "           ISNULL(O.Extra_1, '') AS Extra_1, ISNULL(O.Extra_2, '') AS Extra_2, ISNULL(O.Extra_6, '') AS Extra_6, ISNULL(O.Extra_9, '') AS Extra_9, ISNULL(O.Extra_10, 0) AS Extra_10, ISNULL(O.Item_Cd, '') AS Item_Cd, ISNULL(O.Mix_Group, 0) as Mix_Group "  '&
                '"FROM       OEI_ORDLIN O WITH (NOLOCK) " &
                '"INNER JOIN IMITMIDX_SQL I WITH (nolock) ON O.Item_No = I.Item_No " &
                '"LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No AND O.Loc = L.Loc " &
                '"LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                '"LEFT JOIN  OEI_OrdBld K WITH (Nolock) ON O.Item_Guid = K.Comp_Item_Guid " &
                '"LEFT JOIN  Exact_Traveler_Route R WITH (Nolock) ON O.RouteID = R.RouteID " &
                '"WHERE      O.Ord_GUID = '" & Ordhead.Ord_GUID & "' AND I.ACTIVITY_CD = 'A' " &
                '"ORDER BY   O.Line_No, K.Seq_No "

                If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then

                    strSql &= " FROM  OEI_ORDLIN O WITH (NOLOCK) " &
                "INNER JOIN IMITMIDX_SQL I WITH (nolock) ON O.Item_No = I.Item_No " &
                "LEFT JOIN  OEI_Item_Loc_Qty_View L ON I.Item_No = L.Item_No AND O.Loc = L.Loc " &
                "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                "LEFT JOIN  OEI_OrdBld K WITH (Nolock) ON O.Item_Guid = K.Comp_Item_Guid " &
                "LEFT JOIN  Exact_Traveler_Route R WITH (Nolock) ON O.RouteID = R.RouteID " &
                "WHERE      O.Ord_GUID = '" & Ordhead.Ord_GUID & "' AND I.ACTIVITY_CD = 'A' " &
                "ORDER BY   O.Line_No, K.Seq_No "
                Else
                    strSql &= " FROM  OEI_ORDLIN O WITH (NOLOCK) " &
                "INNER JOIN [200].dbo.IMITMIDX_SQL I WITH (nolock) ON O.Item_No = I.Item_No " &
                "LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No AND O.Loc = L.Loc " &
                "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " &
                "LEFT JOIN  OEI_OrdBld K WITH (Nolock) ON O.Item_Guid = K.Comp_Item_Guid " &
                "LEFT JOIN  Exact_Traveler_Route R WITH (Nolock) ON O.RouteID = R.RouteID " &
                "WHERE      O.Ord_GUID = '" & Ordhead.Ord_GUID & "' AND I.ACTIVITY_CD = 'A' " &
                "ORDER BY   O.Line_No, K.Seq_No "

                End If




                '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON O.Item_No = L.Item_No AND O.Loc = L.Loc " & _

                '"WHERE      O.Ord_GUID = '" & Ordhead.Ord_GUID & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
                '"ORDER BY   O.Line_No, K.Seq_No "

                'strSql = _
                '"SELECT 	P.Picture AS Image, I.Item_No, I.Item_Desc_1, I.Item_Desc_2, " & _
                '"           CAST(0 AS FLOAT) AS Qty_Ordered, CAST(0 AS FLOAT) AS Qty_To_Ship, CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(PQV.Qty_Ordered,0)) AS FLOAT) AS Qty_On_Hand, " & _
                '"           DBO.OEI_Item_Price('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', 1) as Unit_Price, " & _
                '"           CAST(0 AS FLOAT) AS Qty_Ship_Prev, CAST(0 AS FLOAT) AS Qty_BkOrd, ISNULL(I.UOM, '') AS UOM, " & _
                '"           CONVERT(VARCHAR, GETDATE(), 103) AS Request_Dt, CONVERT(VARCHAR, GETDATE(), 103) AS Promise_DT, " & _
                '"           ISNULL(I.BkOrd_Fg, 'N') AS BkOrd_Fg, '' AS Route, I.Activity_Cd " & _
                '"FROM 		IMITMIDX_SQL I WITH (nolock) " & _
                '"LEFT JOIN  IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No " & _
                '"LEFT JOIN  OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
                '"LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " & _
                '"WHERE 		I.Item_No = '" & value & "' AND I.ACTIVITY_CD = 'A' AND L.Loc = '" & m_oOrder.Ordhead.Mfg_Loc & "' " & _
                '"ORDER BY   I.Item_No "

                '"SELECT 	P.Picture AS Image, I.Item_No AS 'Item No', I.Item_Desc_1 AS 'Description', " & _
                '"           I.Item_Desc_2 AS 'Description 2', CAST(0 AS VARCHAR) AS 'Qty Ordered', " & _
                '"           DBO.OEI_Item_Price('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', 1) as 'Price', " & _
                '"           CAST(CAST(l.Qty_On_Hand AS INT) AS VARCHAR) AS 'Qty available', CAST(0 AS VARCHAR) AS 'Qty Shipped', CAST(0 AS VARCHAR) AS 'Qty B.O.', " & _
                '"           CONVERT(VARCHAR, GETDATE(), 103) AS 'Date Requested', CONVERT(VARCHAR, GETDATE(), 103) AS 'Date Promised', '' AS 'Route' " & _
                '"FROM 		IMITMIDX_SQL I WITH (nolock) " & _
                '"INNER JOIN IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '1 '" & _
                '"LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " & _
                '"WHERE 		I.Item_No LIKE '" & Item_No & "%' AND I.ACTIVITY_CD = 'A' " & _
                '"ORDER BY   I.Item_No "

                '"SELECT 	 I.Item_No AS 'Item No', I.Item_Desc_1, " & _
                '"           I.Item_Desc_2, CAST(0 AS VARCHAR) AS 'Qty Ordered', " & _
                '"           CAST(CAST(l.Qty_On_Hand AS INT) AS VARCHAR) AS 'Qty available', CAST(0 AS VARCHAR) AS 'Qty Shipped', CAST(0 AS VARCHAR) AS 'Qty B.O.', " & _
                '"           CONVERT(VARCHAR, GETDATE(), 103) AS 'Date Requested', CONVERT(VARCHAR, GETDATE(), 103) AS 'Date Promised', '' AS 'Route' " & _
                '"FROM 		 IMITMIDX_SQL I WITH (nolock) " & _
                '"INNER JOIN IMINVLOC_SQL L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '1 '" & _
                '"WHERE 	    I.Item_No LIKE '" & SqlCompliantString(m_Item_No) & "%' AND I.ACTIVITY_CD = 'A' " & _
                '"ORDER BY   I.Item_No "

                '' si on a trouvé aucun record, alors on ouvre la fenetre de recherche
                'If rsItem.RecordCount = 0 Then
                '    If value.Length < 4 Then
                '        ' Entry too short
                '    Else
                '        Dim oOrderSelect As New frmProductLineEntry() ' m_oOrder)
                '        oOrderSelect.Customer_No = m_oOrder.Ordhead.Cus_No
                '        oOrderSelect.Currency = m_oOrder.Ordhead.Curr_Cd ' "USD"
                '        oOrderSelect.Item_No = value
                '        'oOrderSelect.ProductLineEntry.LoadProductGrid()
                '        oOrderSelect.ShowDialog()
                '        'If oOrderSelect.ProductStatus = 1 Then
                '        '    Call InsertProductsInOrder(oOrderSelect)
                '        'End If
                '        'If dgvItems.Rows.Count > 2 Then
                '        '    dsDataSet.Tables(0).Rows.RemoveAt(e.RowIndex)
                '        'End If

                '    End If
                'Else
                '    m_Loc = m_oOrder.Ordhead.Mfg_Loc
                '    m_Item_Desc_1 = IIf(rsItem.Fields("Item_Desc_1").Equals(DBNull.Value), "", rsItem.Fields("Item_Desc_1").Value)
                '    m_Item_Desc_2 = IIf(rsItem.Fields("Item_Desc_2").Equals(DBNull.Value), "", rsItem.Fields("Item_Desc_2").Value)
                '    m_Qty_Ordered = 0
                '    m_Qty_To_Ship = 0 ' rsItem.Fields("Qty_Ordered").Value
                '    m_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_To_Ship").Value
                '    m_Qty_On_Hand = rsItem.Fields("Qty_On_Hand").Value
                '    m_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_Ship_Prev").Value
                '    m_Qty_Prev_Bkord = rsItem.Fields("Qty_BkOrd").Value
                '    m_Unit_Price = rsItem.Fields("Unit_Price").Value
                '    m_Calc_Price = 0
                '    m_Discount_Pct = m_oOrder.Ordhead.Discount_Pct
                '    m_Promise_Dt = IIf(FormatDateTime(m_oOrder.Ordhead.Shipping_Dt, DateFormat.ShortDate) = "1/1/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
                '    m_Request_Dt = m_Promise_Dt
                '    m_Req_Ship_Dt = m_Promise_Dt
                '    m_Uom = IIf(rsItem.Fields("UOM").Equals(DBNull.Value), "", rsItem.Fields("UOM").Value)
                '    m_Bkord_Fg = IIf(rsItem.Fields("Bkord_Fg").Equals(DBNull.Value), "N", rsItem.Fields("Bkord_Fg").Value)
                '    m_Recalc_Sw = "Y" ' IIf(rsItem.Fields("Recalc_Sw").Equals(DBNull.Value), "N", rsItem.Fields("Recalc_Sw").Value)

            Else
                    ' Do nothing    
                    ' Call IDEOrderLines()
                End If

            Dim db As New cDBA
            Dim dt As New DataTable

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                For Each row As DataRow In dt.Rows
                    If Not (m_oOrder.OrdLines.Contains(row.Item("Item_Guid"))) Then
                        Dim olLine As New cOrdLine(Ordhead.Ord_GUID)

                        '' Utiliser la methode ici pour faire le LOAD.
                        olLine.blnLoading = True
                        olLine.blnFirstLoad = True

                        olLine.m_Ord_Guid = Ordhead.Ord_GUID

                        If pSource = OrderSourceEnum.OEInterface Then olLine.m_Item_Guid = row.Item("Item_Guid")

                        olLine.m_Line_No = row.Item("Line_No")
                        'olLine.m_Kit_Component = (row.Item("Seq_No") > 0)
                        olLine.m_Kit_Component = (row.Item("Kit_Item_Guid") <> row.Item("Comp_Item_Guid"))
                        olLine.m_Item_No = Trim(row.Item("Item_No"))
                        olLine.m_Loc = row.Item("Loc")
                        olLine.m_Stocked_Fg = IIf(row.Item("Stocked_Fg").Equals(DBNull.Value), "N", row.Item("Stocked_Fg"))
                        olLine.m_Qty_Ordered = row.Item("Qty_Ordered")
                        olLine.Qty_To_Ship = row.Item("Qty_To_Ship")
                        'olLine.Qty_Prev_Bkord = olLine.m_Qty_Ordered - olLine.Qty_To_Ship
                        olLine.Extra_9 = row.Item("Extra_9")
                        olLine.Extra_10 = row.Item("Extra_10")
                        'olLine.m_Qty_On_Hand = 0
                        'olLine.m_Qty_Prev_To_Ship = 0
                        'olLine.m_Qty_Prev_Bkord = 0
                        olLine.m_Ord_Type = row.Item("Ord_Type")
                        olLine.m_Line_Seq_No = row.Item("Line_Seq_No")
                        olLine.m_Pick_Seq = row.Item("Pick_Seq")
                        olLine.m_Ord_No = row.Item("Ord_No")
                        olLine.m_Item_Desc_1 = row.Item("Item_Desc_1")
                        olLine.m_Item_Desc_2 = row.Item("Item_Desc_2")
                        olLine.m_Orig_Price = row.Item("Orig_Price")


                        olLine.m_Unit_Price = row.Item("Unit_Price")

                        'June 22, 2017 - T. Louzon  **************************************************************
                        olLine.m_Unit_Price_BeforeSpecial = row.Item("Unit_Price_BeforeSpecial")
                        'June 27, 2017 - T. Louzon  
                        olLine.m_Mfg_Item_No = row.Item("Mfg_Item_no")
                        olLine.m_RouteExtension = row.Item("RouteExtension")

                        olLine.m_Unit_Cost = row.Item("Unit_Cost")
                        olLine.m_Discount_Pct = row.Item("Discount_Pct")
                        olLine.m_Request_Dt = row.Item("Request_Dt")
                        olLine.m_Promise_Dt = row.Item("Promise_Dt")
                        olLine.m_Select_Cd = row.Item("Select_Cd")
                        olLine.m_Req_Ship_Dt = row.Item("Req_Ship_Dt")
                        olLine.m_User_Def_Fld_1 = row.Item("User_Def_Fld_1")
                        olLine.m_User_Def_Fld_2 = row.Item("User_Def_Fld_2")
                        olLine.m_User_Def_Fld_3 = row.Item("User_Def_Fld_3")
                        olLine.m_User_Def_Fld_4 = row.Item("User_Def_Fld_4")
                        olLine.m_User_Def_Fld_5 = row.Item("User_Def_Fld_5")
                        olLine.m_Qty_Prev_Bkord_Fg = row.Item("Qty_Bkord_Fg")
                        olLine.m_Po_Ord_No = row.Item("Po_Ord_No")
                        olLine.m_Rma_Seq = row.Item("Rma_Seq")
                        olLine.m_Recalc_Sw = row.Item("Recalc_Sw")

                        olLine.m_Qty_On_Hand = IIf(row.Item("Qty_On_Hand").Equals(DBNull.Value), 0, row.Item("Qty_On_Hand"))
                        olLine.m_Uom = IIf(row.Item("UOM").Equals(DBNull.Value), "", row.Item("UOM"))
                        olLine.m_Bkord_Fg = IIf(row.Item("BkOrd_Fg").Equals(DBNull.Value), "N", row.Item("BkOrd_Fg"))
                        olLine.m_Route = IIf(row.Item("Route").Equals(DBNull.Value), "", row.Item("Route"))
                        olLine.m_RouteID = IIf(row.Item("Routeid").Equals(DBNull.Value), 0, row.Item("RouteID"))
                        olLine.m_Activity_Cd = IIf(row.Item("Activity_Cd").Equals(DBNull.Value), 0, row.Item("Activity_Cd"))
                        olLine.m_Prod_Cat = IIf(row.Item("Prod_Cat").Equals(DBNull.Value), "", row.Item("Prod_Cat"))
                        olLine.m_Unit_Weight = IIf(row.Item("Unit_Weight").Equals(DBNull.Value), 0, row.Item("Unit_Weight"))
                        olLine.m_Prc_Cd_Orig_Price = IIf(row.Item("Prc_Cd_Orig_Price").Equals(DBNull.Value), 0, row.Item("Prc_Cd_Orig_Price"))
                        olLine.m_End_Item_Cd = IIf(row.Item("End_Item_Cd").Equals(DBNull.Value), " ", row.Item("End_Item_Cd"))

                        'olLine.m_Recalc_Sw = "Y" ' Prc_Cd_Orig_Price
                        olLine.m_OEI_W_Pixels = row.Item("OEI_W_Pixels")
                        olLine.m_OEI_H_Pixels = row.Item("OEI_H_Pixels")
                        'olLine.m_Recalc_Sw = "Y"

                        olLine.m_ProductProof = row.Item("ProductProof")
                        olLine.m_AutoCompleteReship = row.Item("AutoCompleteReship")

                        olLine.m_Calc_Price = ((100 - olLine.m_Discount_Pct) / 100) * (olLine.m_Unit_Price * olLine.m_Qty_To_Ship)

                        'MB++ HAVE TO LOAD IF IT IS A KIT !!!!

                        olLine.Kit.IsAKit = (Trim(row.Item("Comp_Item_Guid")) <> "" And Not (olLine.m_Kit_Component))
                        olLine.m_Kit_Item_Guid = row.Item("Kit_Item_Guid").ToString
                        'olLine.m_Kit_Item_Guid = row.Item("Comp_Item_Guid").ToString

                        olLine.Item_No_Reload()

                        olLine.Comp_Item_Guid = (Trim(row.Item("Comp_Item_Guid")))
                        olLine.Item_Cd = Trim(row.Item("Item_Cd"))
                        olLine.Mix_Group = row.Item("Mix_Group")

                        olLine.Extra_1 = IIf(row.Item("Extra_1").Equals(DBNull.Value), "", row.Item("Extra_1"))
                        olLine.Extra_2 = IIf(row.Item("Extra_2").Equals(DBNull.Value), "", row.Item("Extra_2"))
                        olLine.Extra_6 = IIf(row.Item("Extra_6").Equals(DBNull.Value), "", row.Item("Extra_6"))

                        '' Calcul de la quantite disponible. Maintenant fait par Item_No_Reload.
                        'If olLine.m_Qty_On_Hand < 0 Then
                        '    olLine.m_Qty_Prev_To_Ship = 0 'm_Qty_On_Hand
                        '    olLine.m_Qty_Prev_Bkord = olLine.m_Qty_Ordered
                        'Else
                        '    If olLine.m_Qty_Ordered > olLine.m_Qty_On_Hand Then
                        '        olLine.m_Qty_Prev_To_Ship = olLine.m_Qty_On_Hand
                        '        olLine.m_Qty_Prev_Bkord = olLine.m_Qty_Ordered - olLine.m_Qty_On_Hand
                        '    Else
                        '        olLine.m_Qty_Prev_To_Ship = olLine.m_Qty_Ordered
                        '        olLine.m_Qty_Prev_Bkord = 0
                        '    End If
                        'End If

                        ' Traitement de l'image
                        Dim data As Byte()

                        Dim stream As MemoryStream = Nothing
                        Dim bmp As Bitmap
                        Dim img As Image = Nothing

                        olLine.m_Image = Nothing

                        If Not (row.Item("Image").Equals(DBNull.Value)) Then
                            data = CType(row.Item("Image"), Byte())
                            stream = New MemoryStream(data, 0, data.Length)

                            bmp = New Bitmap(stream)
                            img = bmp
                            olLine.m_Image = img '' New Bitmap(stream)
                            olLine.ResizeImage()
                        Else
                            bmp = New Bitmap(124, 24)
                            'Dim oResizedImage As Image
                            img = bmp
                            'm_Image = img ' oImage
                            olLine.m_Image = img '' New Bitmap(stream)
                            olLine.m_ImageHeight = 24
                            olLine.m_ImageWidth = 124
                        End If

                        '"           P.Picture AS Image, CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(PQV.Qty_Ordered,0)) AS FLOAT) AS Qty_On_Hand, " & _
                        '"           ISNULL(I.UOM, '') AS UOM, ISNULL(I.BkOrd_Fg, 'N') AS BkOrd_Fg, '' AS Route, I.Activity_Cd " & _

                        '    m_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_To_Ship").Value
                        '    m_Qty_On_Hand = rsItem.Fields("Qty_On_Hand").Value
                        '    m_Qty_Prev_To_Ship = 0 ' rsItem.Fields("Qty_Ship_Prev").Value
                        '    m_Qty_Prev_Bkord = rsItem.Fields("Qty_BkOrd").Value
                        '    m_Unit_Price = rsItem.Fields("Unit_Price").Value
                        '    m_Calc_Price = 0
                        '    m_Discount_Pct = m_oOrder.Ordhead.Discount_Pct
                        '    m_Promise_Dt = IIf(FormatDateTime(m_oOrder.Ordhead.Shipping_Dt, DateFormat.ShortDate) = "1/1/0001", DateAdd("d", 1, Now), m_oOrder.Ordhead.Shipping_Dt)
                        '    m_Request_Dt = m_Promise_Dt
                        '    m_Req_Ship_Dt = m_Promise_Dt
                        '    m_Uom = IIf(rsItem.Fields("UOM").Equals(DBNull.Value), "", rsItem.Fields("UOM").Value)
                        '    m_Bkord_Fg = IIf(rsItem.Fields("Bkord_Fg").Equals(DBNull.Value), "N", rsItem.Fields("Bkord_Fg").Value)
                        '    m_Recalc_Sw = "Y" ' IIf(rsItem.Fields("Recalc_Sw").Equals(DBNull.Value), "N", rsItem.Fields("Recalc_Sw").Value)

                        ''olLine.Item_Guid = row.Item("Item_Guid")
                        ''olLine.Line_No = IIf(row.Item("Line_No").Equals(DBNull.Value), 0, row.Item("Line_No"))
                        ''olLine.Kit_Component = (row.Item("Seq_No") > 0)
                        ''olLine.Item_No = Trim(row.Item("Item_No"))
                        ''olLine.Loc = row.Item("Loc")
                        ''olLine.Qty_Ordered = IIf(row.Item("Qty_Ordered").Equals(DBNull.Value), 0, row.Item("Qty_Ordered"))
                        ''olLine.Qty_To_Ship = IIf(row.Item("Qty_To_Ship").Equals(DBNull.Value), 0, row.Item("Qty_To_Ship"))
                        ''olLine.Ord_Type = IIf(row.Item("Ord_Type").Equals(DBNull.Value), "", row.Item("Ord_Type"))
                        ''olLine.Line_Seq_No = IIf(row.Item("Line_Seq_No").Equals(DBNull.Value), 0, row.Item("Line_Seq_No"))
                        ''olLine.Pick_Seq = IIf(row.Item("Pick_Seq").Equals(DBNull.Value), "", row.Item("Pick_Seq"))
                        ''olLine.Ord_No = IIf(row.Item("Ord_No").Equals(DBNull.Value), "", row.Item("Ord_No"))
                        ''olLine.Item_Desc_1 = IIf(row.Item("Item_Desc_1").Equals(DBNull.Value), "", row.Item("Item_Desc_1"))
                        ''olLine.Item_Desc_2 = IIf(row.Item("Item_Desc_2").Equals(DBNull.Value), "", row.Item("Item_Desc_2"))
                        ''olLine.Unit_Price = IIf(row.Item("Unit_Price").Equals(DBNull.Value), 0, row.Item("Unit_Price"))
                        ''olLine.Discount_Pct = IIf(row.Item("Discount_Pct").Equals(DBNull.Value), 0, row.Item("Discount_Pct"))
                        ''olLine.Request_Dt = IIf(row.Item("Request_Dt").Equals(DBNull.Value), 0, row.Item("Request_Dt"))
                        ''olLine.Promise_Dt = IIf(row.Item("Promise_Dt").Equals(DBNull.Value), 0, row.Item("Promise_Dt"))
                        ''olLine.Select_Cd = IIf(row.Item("Select_Cd").Equals(DBNull.Value), "", row.Item("Select_Cd"))
                        ''olLine.Req_Ship_Dt = IIf(row.Item("Req_Ship_Dt").Equals(DBNull.Value), 0, row.Item("Req_Ship_Dt"))
                        ''olLine.User_Def_Fld_1 = IIf(row.Item("User_Def_Fld_1").Equals(DBNull.Value), "", row.Item("User_Def_Fld_1"))
                        ''olLine.User_Def_Fld_2 = IIf(row.Item("User_Def_Fld_2").Equals(DBNull.Value), "", row.Item("User_Def_Fld_2"))
                        ''olLine.User_Def_Fld_3 = IIf(row.Item("User_Def_Fld_3").Equals(DBNull.Value), "", row.Item("User_Def_Fld_3"))
                        ''olLine.User_Def_Fld_4 = IIf(row.Item("User_Def_Fld_4").Equals(DBNull.Value), "", row.Item("User_Def_Fld_4"))
                        ''olLine.User_Def_Fld_5 = IIf(row.Item("User_Def_Fld_5").Equals(DBNull.Value), "", row.Item("User_Def_Fld_5"))
                        ''olLine.Qty_Bkord_Fg = IIf(row.Item("Qty_Bkord_Fg").Equals(DBNull.Value), "", row.Item("Qty_Bkord_Fg"))
                        ''olLine.Po_Ord_No = IIf(row.Item("Po_Ord_No").Equals(DBNull.Value), "", row.Item("Po_Ord_No"))
                        ''olLine.Rma_Seq = IIf(row.Item("Rma_Seq").Equals(DBNull.Value), 0, row.Item("Rma_Seq"))
                        ''olLine.Recalc_Sw = IIf(row.Item("Recalc_Sw").Equals(DBNull.Value), "", row.Item("Recalc_Sw"))

                        olLine.blnLoading = False
                        olLine.blnFirstLoad = False

                        ' Add the line first so the workline follows for the imprint and the traveler

                        'Dim oLastKitLine As cOrdLine = Nothing
                        Dim iLastPos As Integer = 0

                        If olLine.Kit_Item_Guid <> olLine.Item_Guid And Trim(olLine.Kit_Item_Guid) <> "" Then
                            Dim oKitLine As cOrdLine
                            For iPos As Integer = 1 To OrdLines.Count
                                oKitLine = OrdLines.Item(iPos)
                                If oKitLine.Item_Guid = olLine.Kit_Item_Guid Then
                                    iLastPos = iPos
                                ElseIf Not (oKitLine.Kit_Item_Guid Is Nothing) Then
                                    If oKitLine.Kit_Item_Guid.ToString = olLine.Kit_Item_Guid Then ' oKitLine.Kit_Item_Guid = olLine.Kit_Item_Guid Then
                                        'oLastKitLine = OrdLines.Item(iPos)
                                        iLastPos = iPos
                                    End If
                                End If
                            Next iPos
                        End If

                        olLine.Source = OEI_SourceEnum.frmOrder

                        If iLastPos = 0 Then ' oLastKitLine Is Nothing Then
                            OrdLines.Add(olLine, olLine.Item_Guid)
                        Else
                            OrdLines.Add(olLine, olLine.Item_Guid, , iLastPos)
                        End If

                        olLine.Create_Imprint = True

                        olLine.ImprintLine = New cImprint(m_oOrder.Ordhead.Ord_GUID, olLine.Item_Guid)
                        olLine.ImprintLine.Load()

                        olLine.Create_Traveler = True

                        olLine.Traveler = New CTraveler(m_oOrder.Ordhead.Ord_GUID, olLine.Item_Guid)
                        olLine.Traveler.Load()

                        'olLine.m_Route = row.Item("Route")
                        'olLine.m_RouteID = row.Item("RouteID")

                        olLine.Create_Kit = True

                        If olLine.End_Item_Cd = "K" Then
                            Call olLine.Set_Kit_Inventory()
                        End If

                        olLine.SaveToDB = True

                    End If

                Next

            End If

            WorkLine = GetMaxLine_No(Ordhead.Ord_GUID) + 1

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub Save()

        'Does nothing, is not accessed.

    End Sub

    Public Function CreateExcelFile() As String

        CreateExcelFile = ""
        Dim errPos As Integer

        errPos = 1
        Dim xlApp As Excel.Application
        xlApp = New Excel.ApplicationClass

        errPos = 2
        Dim xlWorkBook As Excel.Workbook = Nothing
        errPos = 3
        Dim xlWorkSheet As Excel.Worksheet
        errPos = 4
        Dim misValue As Object = System.Reflection.Missing.Value

        errPos = 5
        Try
            errPos = 6
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            errPos = 7
            xlWorkSheet = xlWorkBook.Sheets(1)
            errPos = 8

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String

            'errPos = 1
            'Dim xlApp As Object ' Excel.Application
            'errPos = 2
            'Dim xlWorkBook As Object ' Excel.Workbook
            'errPos = 3
            'Dim xlWorkSheet As Object ' Excel.Worksheet
            'errPos = 4
            'Dim misValue As Object = System.Reflection.Missing.Value
            'errPos = 5
            'xlApp = CreateObject("Excel.Application")
            'errPos = 6
            'xlWorkBook = xlApp.Workbooks.Add(misValue)
            'errPos = 7
            'xlWorkSheet = xlWorkBook.Sheets(1)
            'errPos = 8

            errPos = 9
            Dim strLastGuid As String = ""
            errPos = 10

            'Try

            errPos = 11
            Dim iPos As Integer = 0
            errPos = 12
            Dim iCount As Integer = 0
            errPos = 13
            For Each oOrdline As cOrdLine In OrdLines
                ' As we do not have now line validation, validate empty items here 
                ' and do not include them in the excel export file.
                ' iCount and iPos are different because of kits. We do not insert line for
                ' kit components but they are part of the order because of traveler and imprint.
                ' iCount is the number of lines in the order
                errPos = 21
                iCount += 1
                errPos = 22
                If Trim(oOrdline.Item_No) <> "" And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
                    strSql = "SELECT ITEM_NO FROM IMITMIDX_SQL WITH (NOLOCK) WHERE ITEM_NO='" & SqlCompliantString(Trim(oOrdline.Item_No)) & "' "
                    dt = db.DataTable(strSql)
                    If dt.Rows.Count <> 0 Then
                        ' iPos is the Y row position to write to in the Excel file.
                        errPos = 30
                        strLastGuid = oOrdline.Item_Guid
                        errPos = 31
                        iPos += 1
                        errPos = 32

                        'forrced to enterd Type O  , it was from the start of OEI
                        xlWorkSheet.Cells(iPos, xlsExport.Ord_Type) = "O"

                        '++ID 08.06.2024 below used to separate and take just O otr I, requested to commented changes fro 08.05.2024
                        'xlWorkSheet.Cells(iPos, xlsExport.Ord_Type) = Mid(Ordhead.Ord_Type, 1, 1) '  function just take '0' or 'I'
                        errPos = 33
                        If Ordhead.Ord_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Ord_Dt) = _
                            Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") & _
                            Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") & _
                            Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
                        errPos = 34
                        '        Selection.NumberFormat = "@"
                        xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No).Style.NumberFormat = "@"
                        errPos = 35
                        xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No) = Ordhead.Oe_Po_No
                        errPos = 36
                        xlWorkSheet.Cells(iPos, xlsExport.Cus_No) = Ordhead.Cus_No
                        errPos = 37
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Name) = Ordhead.Bill_To_Name
                        errPos = 38
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_1) = Ordhead.Bill_To_Addr_1
                        errPos = 39
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_2) = Ordhead.Bill_To_Addr_2
                        errPos = 40
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_3) = Ordhead.Bill_To_Addr_3
                        errPos = 41
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_4) = Ordhead.Bill_To_Addr_4
                        errPos = 42
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Country) = Ordhead.Bill_To_Country
                        errPos = 43

                        If Ordhead.Cus_Alt_Adr_Cd <> "SAME" And Ordhead.Cus_Alt_Adr_Cd <> "" Then
                            errPos = 44
                            xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
                        Else
                            errPos = 45
                            xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = ""
                        End If

                        errPos = 46
                        'xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Name) = Ordhead.Ship_To_Name
                        errPos = 47
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_1) = Ordhead.Ship_To_Addr_1
                        errPos = 48
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_2) = Ordhead.Ship_To_Addr_2
                        errPos = 49
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_3) = Ordhead.Ship_To_Addr_3
                        errPos = 50
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_4) = Ordhead.Ship_To_Addr_4
                        errPos = 51
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Country) = Ordhead.Ship_To_Country
                        errPos = 52
                        If Ordhead.Shipping_Dt.Date.Equals(NoDate()) Then
                            errPos = 53
                            xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) = "" ' "21111111"
                        Else
                            errPos = 54
                            xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) = _
                            Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") & _
                            Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") & _
                            Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0")
                            'If Entered_Dt.Year <> 1 Then pDataRow.Item("Shipping_Dt") = Shipping_Dt.Date
                        End If

                        'If Ordhead.Shipping_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) = _
                        '    Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") & _
                        '    Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") & _
                        '    Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0")
                        errPos = 55
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_Via_Cd) = Ordhead.Ship_Via_Cd
                        errPos = 56
                        xlWorkSheet.Cells(iPos, xlsExport.Ar_Terms_Cd) = Ordhead.Ar_Terms_Cd
                        errPos = 57
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_1) = Ordhead.Ship_Instruction_1
                        errPos = 58
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_2) = Ordhead.Ship_Instruction_2
                        errPos = 59
                        xlWorkSheet.Cells(iPos, xlsExport.SlsPsn_No) = Ordhead.Slspsn_No
                        errPos = 60
                        xlWorkSheet.Cells(iPos, xlsExport.Discount_Pct) = Ordhead.Discount_Pct
                        errPos = 61
                        xlWorkSheet.Cells(iPos, xlsExport.Job_No) = Ordhead.Job_No
                        errPos = 62
                        xlWorkSheet.Cells(iPos, xlsExport.Mfg_Loc) = Ordhead.Mfg_Loc
                        errPos = 63
                        xlWorkSheet.Cells(iPos, xlsExport.Profit_Center) = Ordhead.Profit_Center
                        errPos = 64
                        xlWorkSheet.Cells(iPos, xlsExport.Filler_0001) = Ordhead.Ord_GUID
                        errPos = 65
                        xlWorkSheet.Cells(iPos, xlsExport.Curr_Cd) = Ordhead.Curr_Cd
                        errPos = 66
                        xlWorkSheet.Cells(iPos, xlsExport.Orig_Trx_Rt) = Ordhead.Orig_Trx_Rt
                        errPos = 67
                        xlWorkSheet.Cells(iPos, xlsExport.Curr_Trx_Rt) = Ordhead.Curr_Trx_Rt
                        errPos = 68
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Sched) = Ordhead.Tax_Sched
                        errPos = 69
                        xlWorkSheet.Cells(iPos, xlsExport.Contact_1) = Ordhead.Contact_1
                        errPos = 70
                        xlWorkSheet.Cells(iPos, xlsExport.Phone_Number) = Ordhead.Phone_Number
                        errPos = 71
                        xlWorkSheet.Cells(iPos, xlsExport.Fax_Number) = Ordhead.Fax_Number
                        errPos = 72
                        xlWorkSheet.Cells(iPos, xlsExport.Email_Address) = Ordhead.Email_Address
                        errPos = 73
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_City) = Ordhead.Ship_To_City
                        errPos = 74
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_State) = Ordhead.Ship_To_State
                        errPos = 75
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Zip) = Ordhead.Ship_To_Zip
                        errPos = 76
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_City) = Ordhead.Bill_To_City
                        errPos = 77
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_State) = Ordhead.Bill_To_State
                        errPos = 78
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Zip) = Ordhead.Bill_To_Zip
                        errPos = 79

                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd) = Ordhead.Tax_Cd
                        errPos = 80
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct) = Ordhead.Tax_Pct
                        errPos = 81
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_2) = Ordhead.Tax_Cd_2
                        errPos = 82
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_2) = Ordhead.Tax_Pct_2
                        errPos = 83
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_3) = Ordhead.Tax_Cd_3
                        errPos = 84
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_3) = Ordhead.Tax_Pct_3
                        errPos = 85
                        xlWorkSheet.Cells(iPos, xlsExport.Form_No) = Ordhead.Form_No
                        errPos = 86
                        xlWorkSheet.Cells(iPos, xlsExport.Deter_Rate_By) = Ordhead.Deter_Rate_By
                        errPos = 87
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_1) = Ordhead.User_Def_Fld_1
                        errPos = 88
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_2) = Ordhead.User_Def_Fld_2
                        errPos = 89
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_3) = Ordhead.User_Def_Fld_3
                        errPos = 90
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_4) = Ordhead.User_Def_Fld_4
                        errPos = 91
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_5) = Ordhead.User_Def_Fld_5
                        errPos = 92
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Fg) = Ordhead.Tax_Fg
                        errPos = 93
                        xlWorkSheet.Cells(iPos, xlsExport.Frt_Pay_Cd) = Ordhead.Frt_Pay_Cd
                        errPos = 94
                        xlWorkSheet.Cells(iPos, xlsExport.Selection_Cd) = Ordhead.Selection_Cd
                        errPos = 94
                        xlWorkSheet.Cells(iPos, xlsExport.Hold_Fg) = Ordhead.Hold_Fg
                        errPos = 95

                        'ID 07312024 ticket# 40519
                        xlWorkSheet.Cells(iPos, xlsExport.inv_batch_id) = Ordhead.Inv_Batch_Id
                        errPos = 955

                        'xlWorkSheet.Cells(iPos, xlsExport.Misc_Mn_No) = "" ' Must set these 4 fields to NULL, completed when invoicing.
                        'xlWorkSheet.Cells(iPos, xlsExport.Misc_Sb_No) = ""
                        'xlWorkSheet.Cells(iPos, xlsExport.Frt_Mn_No) = ""
                        'xlWorkSheet.Cells(iPos, xlsExport.Frt_Sb_No) = ""

                        xlWorkSheet.Cells(iPos, xlsExport.Item_No) = oOrdline.Item_No
                        errPos = 96
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Loc) = oOrdline.Loc
                        errPos = 97
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_Ordered) = oOrdline.Qty_Ordered
                        errPos = 98
                        'xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = oOrdline.Qty_Prev_To_Ship
                        'errPos = 99
                        'xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = oOrdline.Qty_Prev_Bkord
                        'errPos = 100
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = oOrdline.Qty_To_Ship
                        errPos = 99
                        If oOrdline.Qty_To_Ship >= oOrdline.Qty_Ordered Then
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = 0
                        Else
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = oOrdline.Qty_Ordered - oOrdline.Qty_To_Ship
                        End If
                        errPos = 100
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Price) = oOrdline.Unit_Price
                        errPos = 101
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Discount_Pct) = oOrdline.Discount_Pct
                        errPos = 102
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Weight) = oOrdline.Unit_Weight

                        errPos = 103
                        If oOrdline.Request_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Request_Dt) = _
                            oOrdline.Request_Dt.Year.ToString.PadLeft(4, "0") & _
                            oOrdline.Request_Dt.Month.ToString.PadLeft(2, "0") & _
                            oOrdline.Request_Dt.Day.ToString.PadLeft(2, "0")
                        'oOrdline.Request_Dt
                        errPos = 104
                        If oOrdline.Promise_Dt.Equals(Date.MinValue) Then
                            oOrdline.Promise_Dt = Ordhead.Ord_Dt_Shipped
                        End If

                        If oOrdline.Promise_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Promise_Dt) = _
                            oOrdline.Promise_Dt.Year.ToString.PadLeft(4, "0") & _
                            oOrdline.Promise_Dt.Month.ToString.PadLeft(2, "0") & _
                            oOrdline.Promise_Dt.Day.ToString.PadLeft(2, "0")
                        'oOrdline.Promise_Dt
                        errPos = 105
                        If oOrdline.Req_Ship_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Req_Ship_Date) = _
                            oOrdline.Req_Ship_Dt.Year.ToString.PadLeft(4, "0") & _
                            oOrdline.Req_Ship_Dt.Month.ToString.PadLeft(2, "0") & _
                            oOrdline.Req_Ship_Dt.Day.ToString.PadLeft(2, "0")

                        errPos = 106
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Guid) = oOrdline.Item_Guid

                        ' The EOF_Marker to guid is set the line after the Next of the For Each Command
                        errPos = 107
                        xlWorkSheet.Cells(iPos, xlsExport.EOF_Marker) = "N"

                        errPos = 108
                        xlWorkSheet.Cells(iPos, xlsExport.User_Def_Fld_4) = oOrdline.User_Def_Fld_4
                        errPos = 109
                        xlWorkSheet.Cells(iPos, xlsExport.Item_BkOrd_Fg) = oOrdline.Bkord_Fg
                        errPos = 110
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Bin_Fg) = oOrdline.Bin_Fg
                        errPos = 111
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Prc_Cd_Orig_Price) = oOrdline.Orig_Price
                        errPos = 112
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd_Fg) = IIf(oOrdline.Qty_Prev_Bkord <> 0, "Y", "N") ' oOrdline.Qty_Bkord_Fg
                        errPos = 113
                        'oOrdline()
                        ' Prevent division by zero
                        If Ordhead.Curr_Trx_Rt = 0 Then
                            errPos = 113
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = oOrdline.Unit_Cost  ' Forced at 1, inventory incremental 
                        Else
                            errPos = 114
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = Format(oOrdline.Unit_Cost / Ordhead.Curr_Trx_Rt, "##,##0.000000") ' Forced at 1, inventory incremental 
                        End If

                        errPos = 115
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Release_No) = "     1".ToString ' Forced at 1, inventory incremental 
                        errPos = 116
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_1) = oOrdline.Item_Desc_1 ' Forced at 1, inventory incremental 
                        errPos = 117
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_2) = oOrdline.Item_Desc_2 ' Forced at 1, inventory incremental 
                        '98
                        'xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_2) = oOrdline.Extra_1 ' Forced at 1, inventory incremental 
                        '99
                        xlWorkSheet.Cells(iPos, 98) = oOrdline.Extra_2 ' Forced at 1, inventory incremental 
                        '100
                        xlWorkSheet.Cells(iPos, 99) = oOrdline.Extra_6 ' Forced at 1, inventory incremental 
                        errPos = 118


                        '==============================================
                        'Added June 17, 2017 - T. Louzon

                        'GET THE Item Extension for the Route ID
                        Dim dt2 As DataTable
                        Dim db2 As New cDBA
                        Dim strSql2 As String
                        Dim sItemExtension As String
                        Dim sRouteCategory As String
                        strSql2 = " SELECT XREF.ItemExtension, T.RouteCategory "
                        strSql2 = strSql2 & " FROM Exact_traveler_Route T "
                        strSql2 = strSql2 & " INNER JOIN EXACT_Traveler_Route XREF  "    'HV_OEI_RouteCategory_XREF XREF "  'Changed July 18, 2017
                        strSql2 = strSql2 & " On T.RouteCategory = XREF.RouteCategory "
                        strSql2 = strSql2 & " WHERE XREF.routeID = " & oOrdline.m_RouteID

                        errPos = 119
                        dt2 = db2.DataTable(strSql2)
                        If dt2.Rows.Count <> 0 Then

                            If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                sItemExtension = "XX"
                            Else
                                sItemExtension = dt2.Rows(0).Item("ItemExtension")
                            End If

                            If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                sRouteCategory = "XX"
                            Else
                                sRouteCategory = dt2.Rows(0).Item("RouteCategory")
                            End If

                        Else
                            sItemExtension = "XX"
                            sRouteCategory = "XX"
                        End If

                        dt2.Dispose()

                        errPos = 120
                        xlWorkSheet.Cells(iPos, xlsExport.TravelerRouteID) = oOrdline.m_RouteID
                        xlWorkSheet.Cells(iPos, xlsExport.TravelerRoute) = oOrdline.m_Route
                        xlWorkSheet.Cells(iPos, xlsExport.RouteCategory) = sRouteCategory ' 
                        xlWorkSheet.Cells(iPos, xlsExport.Mfg_Item_No) = oOrdline.Item_No & "-" & sItemExtension ' Changed June 26, 2017
                        xlWorkSheet.Cells(iPos, xlsExport.ItemExtension) = sItemExtension '  

                        errPos = 121
                        'Added June 29, 2017 - T. Louzon
                        xlWorkSheet.Cells(iPos, xlsExport.IsNoStock) = oOrdline.IsNoStock
                        xlWorkSheet.Cells(iPos, xlsExport.IsNoArtwork) = oOrdline.IsNoArtwork

                        errPos = 122
                        'Added July 12, 2017 - T. Louzon
                        xlWorkSheet.Cells(iPos, xlsExport.Prod_Cat) = oOrdline.Prod_Cat
                        xlWorkSheet.Cells(iPos, xlsExport.End_Item_Cd) = oOrdline.End_Item_Cd


                        errPos = 123
                        'GET THE EXTRA FIELDS TO EXPORT
                        '==============================================
                        'Added June 17, 2017 - T.Louzon
                        Dim dt3 As DataTable
                        Dim db3 As New cDBA
                        Dim strSql3 As String
                        'Dim sItemExtension As String
                        'Dim sRouteCategory As String
                        strSql3 = " SELECT   "
                        strSql3 = strSql3 & " item_no, "
                        strSql3 = strSql3 & " num_impr_1, "
                        strSql3 = strSql3 & " num_impr_2, "
                        strSql3 = strSql3 & " num_impr_3, "
                        strSql3 = strSql3 & " imprint = extra_1, "
                        strSql3 = strSql3 & " Imprint_Location = Extra_2, "
                        strSql3 = strSql3 & " imprint_Color = Extra_3, "
                        strSql3 = strSql3 & " NUM = Extra_4,  "       '-- Number of Imprints
                        strSql3 = strSql3 & " Packaging = Extra_5, "
                        strSql3 = strSql3 & " Refill = Extra_8, "
                        strSql3 = strSql3 & " Laser_Setup = Extra_9, "
                        strSql3 = strSql3 & " comments = Comment, "
                        strSql3 = strSql3 & " special_Comments = Comment2, "
                        strSql3 = strSql3 & " Industry = Industry, "
                        strSql3 = strSql3 & " Printer_Instructions, "
                        strSql3 = strSql3 & " Packer_Instructions "

                        strSql3 = strSql3 & " FROM OEI_EXTRA_FIELDS "
                        strSql3 = strSql3 & " WHERE ITEM_GUID = '" & oOrdline.m_Item_Guid & "'"

                        errPos = 124
                        dt3 = db3.DataTable(strSql3)

                        errPos = 125

                        If dt3.Rows.Count <> 0 Then
                            errPos = 126

                            'Added July 18, 2017 - T. Louzon
                            xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_1) = dt3.Rows(0).Item("Num_Impr_1")
                            xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_2) = dt3.Rows(0).Item("Num_Impr_2")
                            xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_3) = dt3.Rows(0).Item("Num_Impr_3")
                            xlWorkSheet.Cells(iPos, xlsExport.Imprint) = dt3.Rows(0).Item("Imprint")    ''
                            xlWorkSheet.Cells(iPos, xlsExport.Imprint_Location) = dt3.Rows(0).Item("Imprint_Location")
                            xlWorkSheet.Cells(iPos, xlsExport.Imprint_Color) = dt3.Rows(0).Item("Imprint_Color")
                            xlWorkSheet.Cells(iPos, xlsExport.Comments) = dt3.Rows(0).Item("Comments")
                            xlWorkSheet.Cells(iPos, xlsExport.Special_Comments) = dt3.Rows(0).Item("Special_Comments")
                            xlWorkSheet.Cells(iPos, xlsExport.Printer_Instructions) = dt3.Rows(0).Item("Printer_Instructions")
                            xlWorkSheet.Cells(iPos, xlsExport.Packer_Instructions) = dt3.Rows(0).Item("Packer_Instructions")
                        Else
                            'Nothing
                        End If
                        errPos = 127
                        dt3.Dispose()

                        '==============================================



                    End If
                Else
                    iCount -= 1
                End If

                errPos = 150
            Next

            ' This must be used so when the last item is a kit component, it sets correctly the last line to guid
            errPos = 151
            '** IF THE PROGRAM CRASHES HERE, THEN EVERY LINE HAS A QTY OF 0 **
            xlWorkSheet.Cells(iPos, xlsExport.EOF_Marker) = Ordhead.Ord_GUID

            '\\bankersnt03\Harvest Ventures\HV_MacolaImport\Files

            'xlWorkBook.SaveAs(Filename:=Settings.Export.SavePath & Ordhead.Ord_GUID & Settings.Export.FileExtension, FileFormat:=Excel.XlFileFormat.xlExcel8)
            'xlWorkBook.SaveAs(Filename:="C:\HVMacolaImport\" & Ordhead.Ord_GUID & ".xls")
            errPos = 152
            'xlWorkBook.SaveAs(Filename:="C:\" & Ordhead.Ord_GUID & ".xls")

            If Trim(m_oSettings.Login.DSNName.ToString.ToUpper) = "THOR" Then
                'xlWorkBook.SaveAs("C:\HV_MacolaImport\" & Trim(Ordhead.OEI_Ord_No) & ".xls")
                'Dim m_path As String = m_oOrder.HV_Import_Files_Folder
                'If Mid(Trim(m_path), m_path.Length, 1) <> "\" Then m_path = Trim(m_path) & "\"
                'xlWorkBook.SaveAs(Filename:="" & m_path & "Files\Files\" & Trim(Ordhead.OEI_Ord_No) & Ordhead.Ord_GUID & ".xls")
                'xlWorkBook.SaveAs(Filename:="" & m_path & "Files\" & Ordhead.Ord_GUID & ".xls")

                xlWorkBook.SaveAs(Filename:="C:\HV_MacolaImport\" & Trim(Ordhead.OEI_Ord_No) & ".xls")

            Else
                'xlWorkBook.SaveAs("C:\HV_MacolaImport\" & Trim(Ordhead.OEI_Ord_No) & ".xls")
                'xlWorkBook.SaveAs(Filename:="\\bankersnt03\Harvest Ventures\HV_MacolaImport\Files\" & Ordhead.Ord_GUID & ".xls")


                '++ID 02.20.2023
                If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                    ''justin add 20230224
                    Call m_oOrder.CreateUSExcelFile()  'HV_MacolaImport1CAD_OE
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''20230808 version 2 change file 2/3's cus#
                    Dim UsOrder_Ordhead As cOrdHead = New cOrdHead
                    Dim CaPo_Ordhead As cOrdHead = New cOrdHead

                    If Trim(Ordhead.Cus_No) = "SPECT00" Then
                        'Initialize header with cus#->SPECTCA...............if needed other field of header will be filled 20230809
                        UsOrder_Ordhead.Cus_No = "SPECTCA"
                        UsOrder_Ordhead.Mfg_Loc = "US1"

                        CaPo_Ordhead.Cus_No = "SPECTUS"
                        CaPo_Ordhead.Mfg_Loc = "GIT"
                    Else
                        'Initialize header with cus#->SPECCA................
                        UsOrder_Ordhead.Cus_No = "SPECCA"
                        UsOrder_Ordhead.Mfg_Loc = "US1"

                        CaPo_Ordhead.Cus_No = "SPECUS"
                        CaPo_Ordhead.Mfg_Loc = "US1"
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Call m_oOrder.CreateUsOrderExcelFile(UsOrder_Ordhead) 'HV_MacolaImport2USD_OE
                    Call m_oOrder.CreateCAPoExcelFile(CaPo_Ordhead) 'HV_MacolaImport3CA_PO

                Else
                    Dim m_path As String = HV_Import_Files_Folder()
                    'm_path = "C:\ExactTemp\4\"
                    'If Mid(Trim(m_path), m_path.Length, 1) <> "\" Then m_path = Trim(m_path) & "\"
                    xlWorkBook.SaveAs(Filename:="" & m_path & Ordhead.Ord_GUID & ".xls")
                    errPos = 153
                    xlWorkBook.Close()
                    xlWorkBook = Nothing
                    errPos = 154

                    xlApp.Quit()
                    xlApp = Nothing

                End If
            End If
            errPos = 155
            MsgBox("File created.")
            'CreateExcelFile = Settings.Export.SavePath & Ordhead.Ord_GUID & Settings.Export.FileExtension
            'Call UnsetExportTS() 'for test repeat export excel!!!  justin 20230809
        Catch er As Exception
            Call UnsetExportTS()
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " " & errPos.ToString)
            xlWorkBook.Close(False)
            xlApp.Quit()
        End Try

    End Function


    'Justin add below 20230221
    Public Function CreateUSExcelFile() As String


        Dim errPos As Integer

        errPos = 1
        Dim xlApp As Excel.Application
        xlApp = New Excel.ApplicationClass

        errPos = 2
        Dim xlWorkBook As Excel.Workbook = Nothing
        errPos = 3
        Dim xlWorkSheet As Excel.Worksheet
        errPos = 4
        Dim misValue As Object = System.Reflection.Missing.Value

        errPos = 5
        Try
            errPos = 6
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            errPos = 7
            xlWorkSheet = xlWorkBook.Sheets(1)
            errPos = 8

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String


            errPos = 9
            Dim strLastGuid As String = ""
            errPos = 10

            'Try

            errPos = 11
            Dim iPos As Integer = 0
            errPos = 12
            Dim iCount As Integer = 0
            errPos = 13
            For Each oOrdline As cOrdLine In OrdLines
                ' As we do not have now line validation, validate empty items here 
                ' and do not include them in the excel export file.
                ' iCount and iPos are different because of kits. We do not insert line for
                ' kit components but they are part of the order because of traveler and imprint.
                ' iCount is the number of lines in the order
                errPos = 21
                iCount += 1
                errPos = 22
                If Trim(oOrdline.Item_No) <> "" And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
                    strSql = "SELECT ITEM_NO FROM IMITMIDX_SQL WITH (NOLOCK) WHERE ITEM_NO='" & SqlCompliantString(Trim(oOrdline.Item_No)) & "' "
                    dt = db.DataTable(strSql)
                    If dt.Rows.Count <> 0 Then
                        ' iPos is the Y row position to write to in the Excel file.
                        errPos = 30
                        strLastGuid = oOrdline.Item_Guid
                        errPos = 31
                        iPos += 1
                        errPos = 32
                        xlWorkSheet.Cells(iPos, xlsExport.Ord_Type) = "O" ' Ordhead.Ord_Type
                        errPos = 33
                        If Ordhead.Ord_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Ord_Dt) =
                            Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") &
                            Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") &
                            Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
                        errPos = 34
                        '        Selection.NumberFormat = "@"
                        xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No).Style.NumberFormat = "@"
                        errPos = 35
                        xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No) = Ordhead.Oe_Po_No
                        errPos = 36
                        xlWorkSheet.Cells(iPos, xlsExport.Cus_No) = Ordhead.Cus_No
                        errPos = 37
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Name) = Ordhead.Bill_To_Name
                        errPos = 38
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_1) = Ordhead.Bill_To_Addr_1
                        errPos = 39
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_2) = Ordhead.Bill_To_Addr_2
                        errPos = 40
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_3) = Ordhead.Bill_To_Addr_3
                        errPos = 41
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_4) = Ordhead.Bill_To_Addr_4
                        errPos = 42
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Country) = Ordhead.Bill_To_Country
                        errPos = 43

                        If Ordhead.Cus_Alt_Adr_Cd <> "SAME" And Ordhead.Cus_Alt_Adr_Cd <> "" Then
                            errPos = 44
                            xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
                        Else
                            errPos = 45
                            xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = ""
                        End If

                        errPos = 46
                        'xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Name) = Ordhead.Ship_To_Name
                        errPos = 47
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_1) = Ordhead.Ship_To_Addr_1
                        errPos = 48
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_2) = Ordhead.Ship_To_Addr_2
                        errPos = 49
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_3) = Ordhead.Ship_To_Addr_3
                        errPos = 50
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_4) = Ordhead.Ship_To_Addr_4
                        errPos = 51
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Country) = Ordhead.Ship_To_Country
                        errPos = 52
                        If Ordhead.Shipping_Dt.Date.Equals(NoDate()) Then
                            errPos = 53
                            xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) = "" ' "21111111"
                        Else
                            errPos = 54
                            xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) =
                            Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") &
                            Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") &
                            Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0")
                            'If Entered_Dt.Year <> 1 Then pDataRow.Item("Shipping_Dt") = Shipping_Dt.Date
                        End If


                        errPos = 55
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_Via_Cd) = Ordhead.Ship_Via_Cd
                        errPos = 56
                        xlWorkSheet.Cells(iPos, xlsExport.Ar_Terms_Cd) = Ordhead.Ar_Terms_Cd
                        errPos = 57
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_1) = Ordhead.Ship_Instruction_1
                        errPos = 58
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_2) = Ordhead.Ship_Instruction_2
                        errPos = 59
                        xlWorkSheet.Cells(iPos, xlsExport.SlsPsn_No) = Ordhead.Slspsn_No
                        errPos = 60
                        xlWorkSheet.Cells(iPos, xlsExport.Discount_Pct) = Ordhead.Discount_Pct
                        errPos = 61
                        xlWorkSheet.Cells(iPos, xlsExport.Job_No) = Ordhead.Job_No
                        errPos = 62
                        xlWorkSheet.Cells(iPos, xlsExport.Mfg_Loc) = Ordhead.Mfg_Loc
                        errPos = 63
                        xlWorkSheet.Cells(iPos, xlsExport.Profit_Center) = Ordhead.Profit_Center
                        errPos = 64
                        xlWorkSheet.Cells(iPos, xlsExport.Filler_0001) = Ordhead.Ord_GUID
                        errPos = 65
                        xlWorkSheet.Cells(iPos, xlsExport.Curr_Cd) = Ordhead.Curr_Cd
                        errPos = 66
                        xlWorkSheet.Cells(iPos, xlsExport.Orig_Trx_Rt) = Ordhead.Orig_Trx_Rt
                        errPos = 67
                        xlWorkSheet.Cells(iPos, xlsExport.Curr_Trx_Rt) = Ordhead.Curr_Trx_Rt
                        errPos = 68
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Sched) = Ordhead.Tax_Sched
                        errPos = 69
                        xlWorkSheet.Cells(iPos, xlsExport.Contact_1) = Ordhead.Contact_1
                        errPos = 70
                        xlWorkSheet.Cells(iPos, xlsExport.Phone_Number) = Ordhead.Phone_Number
                        errPos = 71
                        xlWorkSheet.Cells(iPos, xlsExport.Fax_Number) = Ordhead.Fax_Number
                        errPos = 72
                        xlWorkSheet.Cells(iPos, xlsExport.Email_Address) = Ordhead.Email_Address
                        errPos = 73
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_City) = Ordhead.Ship_To_City
                        errPos = 74
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_State) = Ordhead.Ship_To_State
                        errPos = 75
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Zip) = Ordhead.Ship_To_Zip
                        errPos = 76
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_City) = Ordhead.Bill_To_City
                        errPos = 77
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_State) = Ordhead.Bill_To_State
                        errPos = 78
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Zip) = Ordhead.Bill_To_Zip
                        errPos = 79

                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd) = Ordhead.Tax_Cd
                        errPos = 80
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct) = Ordhead.Tax_Pct
                        errPos = 81
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_2) = Ordhead.Tax_Cd_2
                        errPos = 82
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_2) = Ordhead.Tax_Pct_2
                        errPos = 83
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_3) = Ordhead.Tax_Cd_3
                        errPos = 84
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_3) = Ordhead.Tax_Pct_3
                        errPos = 85
                        xlWorkSheet.Cells(iPos, xlsExport.Form_No) = Ordhead.Form_No
                        errPos = 86
                        xlWorkSheet.Cells(iPos, xlsExport.Deter_Rate_By) = Ordhead.Deter_Rate_By
                        errPos = 87
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_1) = Ordhead.User_Def_Fld_1
                        errPos = 88
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_2) = Ordhead.User_Def_Fld_2
                        errPos = 89
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_3) = Ordhead.User_Def_Fld_3
                        errPos = 90
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_4) = Ordhead.User_Def_Fld_4
                        errPos = 91
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_5) = Ordhead.User_Def_Fld_5
                        errPos = 92
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Fg) = Ordhead.Tax_Fg
                        errPos = 93
                        xlWorkSheet.Cells(iPos, xlsExport.Frt_Pay_Cd) = Ordhead.Frt_Pay_Cd
                        errPos = 94
                        xlWorkSheet.Cells(iPos, xlsExport.Selection_Cd) = Ordhead.Selection_Cd
                        errPos = 94
                        xlWorkSheet.Cells(iPos, xlsExport.Hold_Fg) = Ordhead.Hold_Fg
                        errPos = 95

                        xlWorkSheet.Cells(iPos, xlsExport.Item_No) = oOrdline.Item_No
                        errPos = 96
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Loc) = oOrdline.Loc
                        errPos = 97
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_Ordered) = oOrdline.Qty_Ordered
                        errPos = 98
                        'xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = oOrdline.Qty_Prev_To_Ship
                        'errPos = 99
                        'xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = oOrdline.Qty_Prev_Bkord
                        'errPos = 100
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = oOrdline.Qty_To_Ship
                        errPos = 99
                        If oOrdline.Qty_To_Ship >= oOrdline.Qty_Ordered Then
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = 0
                        Else
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = oOrdline.Qty_Ordered - oOrdline.Qty_To_Ship
                        End If
                        errPos = 100
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Price) = oOrdline.Unit_Price
                        errPos = 101
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Discount_Pct) = oOrdline.Discount_Pct
                        errPos = 102
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Weight) = oOrdline.Unit_Weight

                        errPos = 103
                        If oOrdline.Request_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Request_Dt) =
                            oOrdline.Request_Dt.Year.ToString.PadLeft(4, "0") &
                            oOrdline.Request_Dt.Month.ToString.PadLeft(2, "0") &
                            oOrdline.Request_Dt.Day.ToString.PadLeft(2, "0")
                        'oOrdline.Request_Dt
                        errPos = 104
                        If oOrdline.Promise_Dt.Equals(Date.MinValue) Then
                            oOrdline.Promise_Dt = Ordhead.Ord_Dt_Shipped
                        End If

                        If oOrdline.Promise_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Promise_Dt) =
                            oOrdline.Promise_Dt.Year.ToString.PadLeft(4, "0") &
                            oOrdline.Promise_Dt.Month.ToString.PadLeft(2, "0") &
                            oOrdline.Promise_Dt.Day.ToString.PadLeft(2, "0")
                        'oOrdline.Promise_Dt
                        errPos = 105
                        If oOrdline.Req_Ship_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Req_Ship_Date) =
                            oOrdline.Req_Ship_Dt.Year.ToString.PadLeft(4, "0") &
                            oOrdline.Req_Ship_Dt.Month.ToString.PadLeft(2, "0") &
                            oOrdline.Req_Ship_Dt.Day.ToString.PadLeft(2, "0")

                        errPos = 106
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Guid) = oOrdline.Item_Guid

                        ' The EOF_Marker to guid is set the line after the Next of the For Each Command
                        errPos = 107
                        xlWorkSheet.Cells(iPos, xlsExport.EOF_Marker) = "N"

                        errPos = 108
                        xlWorkSheet.Cells(iPos, xlsExport.User_Def_Fld_4) = oOrdline.User_Def_Fld_4
                        errPos = 109
                        xlWorkSheet.Cells(iPos, xlsExport.Item_BkOrd_Fg) = oOrdline.Bkord_Fg
                        errPos = 110
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Bin_Fg) = oOrdline.Bin_Fg
                        errPos = 111
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Prc_Cd_Orig_Price) = oOrdline.Orig_Price
                        errPos = 112
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd_Fg) = IIf(oOrdline.Qty_Prev_Bkord <> 0, "Y", "N") ' oOrdline.Qty_Bkord_Fg
                        errPos = 113
                        'oOrdline()
                        ' Prevent division by zero
                        If Ordhead.Curr_Trx_Rt = 0 Then
                            errPos = 113
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = oOrdline.Unit_Cost  ' Forced at 1, inventory incremental 
                        Else
                            errPos = 114
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = Format(oOrdline.Unit_Cost / Ordhead.Curr_Trx_Rt, "##,##0.000000") ' Forced at 1, inventory incremental 
                        End If

                        errPos = 115
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Release_No) = "     1".ToString ' Forced at 1, inventory incremental 
                        errPos = 116
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_1) = oOrdline.Item_Desc_1 ' Forced at 1, inventory incremental 
                        errPos = 117
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_2) = oOrdline.Item_Desc_2 ' Forced at 1, inventory incremental 
                        '98
                        'xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_2) = oOrdline.Extra_1 ' Forced at 1, inventory incremental 
                        '99
                        xlWorkSheet.Cells(iPos, 98) = oOrdline.Extra_2 ' Forced at 1, inventory incremental 
                        '100
                        xlWorkSheet.Cells(iPos, 99) = oOrdline.Extra_6 ' Forced at 1, inventory incremental 
                        errPos = 118


                        '==============================================
                        'Added June 17, 2017 - T. Louzon

                        'GET THE Item Extension for the Route ID
                        Dim dt2 As DataTable
                        Dim db2 As New cDBA
                        Dim strSql2 As String
                        Dim sItemExtension As String
                        Dim sRouteCategory As String
                        strSql2 = " SELECT XREF.ItemExtension, T.RouteCategory "
                        strSql2 = strSql2 & " FROM Exact_traveler_Route T "
                        strSql2 = strSql2 & " INNER JOIN EXACT_Traveler_Route XREF  "    'HV_OEI_RouteCategory_XREF XREF "  'Changed July 18, 2017
                        strSql2 = strSql2 & " On T.RouteCategory = XREF.RouteCategory "
                        strSql2 = strSql2 & " WHERE XREF.routeID = " & oOrdline.m_RouteID

                        errPos = 119
                        dt2 = db2.DataTable(strSql2)
                        If dt2.Rows.Count <> 0 Then

                            If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                sItemExtension = "XX"
                            Else
                                sItemExtension = dt2.Rows(0).Item("ItemExtension")
                            End If

                            If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                sRouteCategory = "XX"
                            Else
                                sRouteCategory = dt2.Rows(0).Item("RouteCategory")
                            End If

                        Else
                            sItemExtension = "XX"
                            sRouteCategory = "XX"
                        End If

                        dt2.Dispose()

                        errPos = 120
                        xlWorkSheet.Cells(iPos, xlsExport.TravelerRouteID) = oOrdline.m_RouteID
                        xlWorkSheet.Cells(iPos, xlsExport.TravelerRoute) = oOrdline.m_Route
                        xlWorkSheet.Cells(iPos, xlsExport.RouteCategory) = sRouteCategory ' 
                        xlWorkSheet.Cells(iPos, xlsExport.Mfg_Item_No) = oOrdline.Item_No & "-" & sItemExtension ' Changed June 26, 2017
                        xlWorkSheet.Cells(iPos, xlsExport.ItemExtension) = sItemExtension '  

                        errPos = 121
                        'Added June 29, 2017 - T. Louzon
                        xlWorkSheet.Cells(iPos, xlsExport.IsNoStock) = oOrdline.IsNoStock
                        xlWorkSheet.Cells(iPos, xlsExport.IsNoArtwork) = oOrdline.IsNoArtwork

                        errPos = 122
                        'Added July 12, 2017 - T. Louzon
                        xlWorkSheet.Cells(iPos, xlsExport.Prod_Cat) = oOrdline.Prod_Cat
                        xlWorkSheet.Cells(iPos, xlsExport.End_Item_Cd) = oOrdline.End_Item_Cd


                        errPos = 123
                        'GET THE EXTRA FIELDS TO EXPORT
                        '==============================================
                        'Added June 17, 2017 - T.Louzon
                        Dim dt3 As DataTable
                        Dim db3 As New cDBA
                        Dim strSql3 As String
                        'Dim sItemExtension As String
                        'Dim sRouteCategory As String
                        strSql3 = " SELECT   "
                        strSql3 = strSql3 & " item_no, "
                        strSql3 = strSql3 & " num_impr_1, "
                        strSql3 = strSql3 & " num_impr_2, "
                        strSql3 = strSql3 & " num_impr_3, "
                        strSql3 = strSql3 & " imprint = extra_1, "
                        strSql3 = strSql3 & " Imprint_Location = Extra_2, "
                        strSql3 = strSql3 & " imprint_Color = Extra_3, "
                        strSql3 = strSql3 & " NUM = Extra_4,  "       '-- Number of Imprints
                        strSql3 = strSql3 & " Packaging = Extra_5, "
                        strSql3 = strSql3 & " Refill = Extra_8, "
                        strSql3 = strSql3 & " Laser_Setup = Extra_9, "
                        strSql3 = strSql3 & " comments = Comment, "
                        strSql3 = strSql3 & " special_Comments = Comment2, "
                        strSql3 = strSql3 & " Industry = Industry, "
                        strSql3 = strSql3 & " Printer_Instructions, "
                        strSql3 = strSql3 & " Packer_Instructions "

                        strSql3 = strSql3 & " FROM OEI_EXTRA_FIELDS "
                        strSql3 = strSql3 & " WHERE ITEM_GUID = '" & oOrdline.m_Item_Guid & "'"

                        errPos = 124
                        dt3 = db3.DataTable(strSql3)

                        errPos = 125

                        If dt3.Rows.Count <> 0 Then
                            errPos = 126

                            'Added July 18, 2017 - T. Louzon
                            xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_1) = dt3.Rows(0).Item("Num_Impr_1")
                            xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_2) = dt3.Rows(0).Item("Num_Impr_2")
                            xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_3) = dt3.Rows(0).Item("Num_Impr_3")
                            xlWorkSheet.Cells(iPos, xlsExport.Imprint) = dt3.Rows(0).Item("Imprint")    ''
                            xlWorkSheet.Cells(iPos, xlsExport.Imprint_Location) = dt3.Rows(0).Item("Imprint_Location")
                            xlWorkSheet.Cells(iPos, xlsExport.Imprint_Color) = dt3.Rows(0).Item("Imprint_Color")
                            xlWorkSheet.Cells(iPos, xlsExport.Comments) = dt3.Rows(0).Item("Comments")
                            xlWorkSheet.Cells(iPos, xlsExport.Special_Comments) = dt3.Rows(0).Item("Special_Comments")
                            xlWorkSheet.Cells(iPos, xlsExport.Printer_Instructions) = dt3.Rows(0).Item("Printer_Instructions")
                            xlWorkSheet.Cells(iPos, xlsExport.Packer_Instructions) = dt3.Rows(0).Item("Packer_Instructions")
                        Else
                            'Nothing
                        End If
                        errPos = 127
                        dt3.Dispose()

                        '==============================================



                    End If
                Else
                    iCount -= 1
                End If

                errPos = 150
            Next

            ' This must be used so when the last item is a kit component, it sets correctly the last line to guid
            errPos = 151
            '** IF THE PROGRAM CRASHES HERE, THEN EVERY LINE HAS A QTY OF 0 **
            xlWorkSheet.Cells(iPos, xlsExport.EOF_Marker) = Ordhead.Ord_GUID

            errPos = 152
            'xlWorkBook.SaveAs(Filename:="C:\" & Ordhead.Ord_GUID & ".xls")


            'xlWorkBook.SaveAs("C:\HV_MacolaImport\" & Trim(Ordhead.OEI_Ord_No) & ".xls")
            'xlWorkBook.SaveAs(Filename:="\\bankersnt03\Harvest Ventures\HV_MacolaImport\Files\" & Ordhead.Ord_GUID & ".xls")


            Dim m_path As String = HV_US_MACOLAIMPORT1CAD_Folder()
            '''
            'm_path = "C:\ExactTemp\1\" 'for test
            '''

            xlWorkBook.SaveAs(Filename:="" & m_path & Trim(Ordhead.OEI_Ord_No) & Ordhead.Ord_GUID & ".xls")

            errPos = 153
            xlWorkBook.Close()
            xlWorkBook = Nothing
            errPos = 154

            xlApp.Quit()
            xlApp = Nothing

            'xlApp.DisplayAlerts = False
            'xlApp.ActiveWorkbook.Save()
            'xlApp.ActiveWorkbook.Close()
            'xlApp.UserControl = False
            'xlApp.DisplayAlerts = True
            'xlApp.Quit()
            'xlApp = Nothing

            errPos = 155

            'MsgBox("File_1 created.")

            'CreateExcelFile = Settings.Export.SavePath & Ordhead.Ord_GUID & Settings.Export.FileExtension

        Catch er As Exception
            Call UnsetExportTS()
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " " & errPos.ToString)
            xlWorkBook.Close(False)
            xlApp.Quit()
        End Try

    End Function
    Public Function Get_nqg_price_by_kits(ByVal in_Item_no As String, ByVal original_Price As Double) As Double

        Try
            Get_nqg_price_by_kits = original_Price

            Dim orderline_kits_row_line_no As Int16 = 0
            Dim orderline_kits_row_ord_guid As String = ""
            Dim orderline_kits_row_ord_qty_ordered As Int16 = 0

            For Each oOrdline As cOrdLine In OrdLines
                If Trim(oOrdline.Item_No) = in_Item_no And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
                    orderline_kits_row_ord_guid = oOrdline.Ord_Guid
                    orderline_kits_row_line_no = oOrdline.Line_No
                    Get_nqg_price_by_kits = 0
                    For Each kits_components_oOrdline As cOrdLine In OrdLines
                        'If (LTrim(RTrim(kits_components_oOrdline.Ord_Guid)) = orderline_kits_row_ord_guid And
                        If (kits_components_oOrdline.Line_No = orderline_kits_row_line_no And
                        If_inventory_item_no(kits_components_oOrdline.Item_No)) And
                        Not If_kits_item_no(kits_components_oOrdline.Item_No) Then
                            Get_nqg_price_by_kits = Get_nqg_price_by_kits + Get_nqg_price_by_item_no(kits_components_oOrdline.Item_No, kits_components_oOrdline.Unit_Price)
                        End If
                    Next
                    orderline_kits_row_ord_guid = ""
                    orderline_kits_row_line_no = 0
                End If
            Next
        Catch er As Exception
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try
    End Function

    Public Function Get_commodity_cd_by_item_no(ByVal in_Item_no As String) As String
        Try
            Get_commodity_cd_by_item_no = ""
            Dim Item_no = LTrim(RTrim(in_Item_no))

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String


            strSql = "Select top 1 ltrim(rtrim(isnull(upper(commodity_cd),''))) as commodity_cd
                        From dbo.imitmidx_sql
                        Where  LTrim(RTrim(Item_no)) = '" & in_Item_no & "' "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                Get_commodity_cd_by_item_no = dt.Rows(0).Item("commodity_cd")
            End If

        Catch er As Exception
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try

    End Function

    Public Function Get_nqg_price_by_item_no(ByVal in_Item_no As String, ByVal original_Unit_Price As Double) As Double
        Try
            Get_nqg_price_by_item_no = original_Unit_Price
            Dim Item_no = LTrim(RTrim(in_Item_no))

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String


            strSql = "Select top 1 Item_no, vend_no, curr_cd, approved_cd, neg_price_1
                        From dbo.poitmvnd_sql
                        Where (vend_no = 'SPECUS')
                        And LTrim(RTrim(Item_no)) = '" & in_Item_no & "' "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                Get_nqg_price_by_item_no = dt.Rows(0).Item("neg_price_1")
            End If

        Catch er As Exception
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try

    End Function

    Public Function Get_avg_cost_by_kits_In200(ByVal in_Item_no As String, ByVal original_Price As Double) As Double

        Try
            Get_avg_cost_by_kits_In200 = original_Price

            Dim orderline_kits_row_line_no As Int16 = 0
            Dim orderline_kits_row_ord_guid As String = ""
            Dim orderline_kits_row_ord_qty_ordered As Int16 = 0

            For Each oOrdline As cOrdLine In OrdLines
                If Trim(oOrdline.Item_No) = in_Item_no And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
                    orderline_kits_row_ord_guid = oOrdline.Ord_Guid
                    orderline_kits_row_line_no = oOrdline.Line_No
                    Get_avg_cost_by_kits_In200 = 0
                    For Each kits_components_oOrdline As cOrdLine In OrdLines
                        'If (LTrim(RTrim(kits_components_oOrdline.Ord_Guid)) = orderline_kits_row_ord_guid And
                        If (kits_components_oOrdline.Line_No = orderline_kits_row_line_no And
                        If_inventory_item_no(kits_components_oOrdline.Item_No)) And
                        Not If_kits_item_no(kits_components_oOrdline.Item_No) Then
                            Get_avg_cost_by_kits_In200 = Get_avg_cost_by_kits_In200 + Get_avg_cost_by_item_no_In200(kits_components_oOrdline.Item_No, kits_components_oOrdline.Unit_Price)
                        End If
                    Next
                    orderline_kits_row_ord_guid = ""
                    orderline_kits_row_line_no = 0
                End If
            Next
        Catch er As Exception
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try
    End Function

    Public Function Get_avg_cost_by_item_no_In200(ByVal in_Item_no As String, ByVal original_Unit_Price As Double) As Double
        Try
            Get_avg_cost_by_item_no_In200 = original_Unit_Price
            Dim Item_no = LTrim(RTrim(in_Item_no))

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String

            strSql = "Select top 1 item_no, loc, avg_cost
                        From [200].dbo.iminvloc_sql
                        Where (loc = 'US1') and LTrim(RTrim(Item_no)) = '" & in_Item_no & "' "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                Get_avg_cost_by_item_no_In200 = dt.Rows(0).Item("avg_cost")
            End If

        Catch er As Exception
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try

    End Function

    'Public Function Get_avg_cost_by_kits_In200_git(ByVal in_Item_no As String, ByVal original_Price As Double) As Double

    '    Try
    '        Get_avg_cost_by_kits_In200_git = original_Price

    '        Dim orderline_kits_row_line_no As Int16 = 0
    '        Dim orderline_kits_row_ord_guid As String = ""
    '        Dim orderline_kits_row_ord_qty_ordered As Int16 = 0

    '        For Each oOrdline As cOrdLine In OrdLines
    '            If Trim(oOrdline.Item_No) = in_Item_no And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
    '                orderline_kits_row_ord_guid = oOrdline.Ord_Guid
    '                orderline_kits_row_line_no = oOrdline.Line_No
    '                Get_avg_cost_by_kits_In200_git = 0
    '                For Each kits_components_oOrdline As cOrdLine In OrdLines
    '                    'If (LTrim(RTrim(kits_components_oOrdline.Ord_Guid)) = orderline_kits_row_ord_guid And
    '                    If (kits_components_oOrdline.Line_No = orderline_kits_row_line_no And
    '                    If_inventory_item_no(kits_components_oOrdline.Item_No)) And
    '                    Not If_kits_item_no(kits_components_oOrdline.Item_No) Then
    '                        Get_avg_cost_by_kits_In200_git = Get_avg_cost_by_kits_In200_git + Get_avg_cost_by_item_no_In200_git(kits_components_oOrdline.Item_No, kits_components_oOrdline.Unit_Price)
    '                    End If
    '                Next
    '                orderline_kits_row_ord_guid = ""
    '                orderline_kits_row_line_no = 0
    '            End If
    '        Next
    '    Catch er As Exception
    '        MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
    '    End Try
    'End Function

    'Public Function Get_avg_cost_by_item_no_In200_git(ByVal in_Item_no As String, ByVal original_Unit_Price As Double) As Double
    '    Try
    '        Get_avg_cost_by_item_no_In200_git = original_Unit_Price
    '        Dim Item_no = LTrim(RTrim(in_Item_no))

    '        Dim db As New cDBA
    '        Dim dt As DataTable
    '        Dim strSql As String

    '        strSql = "Select top 1 item_no, loc, avg_cost
    '                    From [200].dbo.iminvloc_sql
    '                    Where (loc = 'GIT') and LTrim(RTrim(Item_no)) = '" & in_Item_no & "' "

    '        dt = db.DataTable(strSql)
    '        If dt.Rows.Count <> 0 Then
    '            Get_avg_cost_by_item_no_In200_git = dt.Rows(0).Item("avg_cost")
    '        End If

    '    Catch er As Exception
    '        MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
    '    End Try

    'End Function

    Public Function Get_all_inventory_item_total_fee() As Double
        Try
            Dim item_total_fee As Double = 0
            Dim orderline_kits_row_line_seq_no As Int16 = 0
            Dim orderline_kits_row_ord_guid As String = ""
            Dim orderline_kits_row_ord_qty_ordered As Int16 = 0
            For Each oOrdline As cOrdLine In OrdLines
                If Trim(oOrdline.Item_No) <> "" And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
                    If If_inventory_item_no(oOrdline.Item_No) And Not If_kits_item_no(oOrdline.Item_No) Then  'is inventory and is not kits
                        item_total_fee = item_total_fee + oOrdline.Qty_Ordered * Get_nqg_price_by_item_no(oOrdline.Item_No, oOrdline.Unit_Price)
                    End If
                    If If_kits_item_no(oOrdline.Item_No) Then 'is inventory and is kits
                        orderline_kits_row_line_seq_no = oOrdline.Line_Seq_No
                        orderline_kits_row_ord_guid = LTrim(RTrim(oOrdline.Ord_Guid))
                        orderline_kits_row_ord_qty_ordered = oOrdline.Qty_Ordered
                        For Each kits_components_oOrdline As cOrdLine In OrdLines
                            If (LTrim(RTrim(kits_components_oOrdline.Ord_Guid)) = orderline_kits_row_ord_guid And
                                kits_components_oOrdline.Line_Seq_No = orderline_kits_row_line_seq_no And
                                If_inventory_item_no(kits_components_oOrdline.Item_No)) And
                                Not If_kits_item_no(kits_components_oOrdline.Item_No) Then

                                item_total_fee = item_total_fee + orderline_kits_row_ord_qty_ordered * Get_nqg_price_by_item_no(kits_components_oOrdline.Item_No, kits_components_oOrdline.Unit_Price)

                            End If
                        Next
                        orderline_kits_row_line_seq_no = 0
                        orderline_kits_row_ord_guid = ""
                        orderline_kits_row_ord_qty_ordered = 0
                    End If
                End If
            Next
            Get_all_inventory_item_total_fee = item_total_fee
        Catch er As Exception
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try
    End Function

    'Public Function Get_all_inventory_item_total_unit_fee() As Double
    '    Try
    '        Get_all_inventory_item_total_unit_fee = 0
    '        Dim item_total_unit_fee As Double = 0
    '        Dim orderline_kits_row_line_no As Int16 = 0
    '        Dim orderline_kits_row_ord_guid As String = ""
    '        Dim orderline_kits_row_ord_qty_ordered As Int16 = 0
    '        For Each oOrdline As cOrdLine In OrdLines
    '            If Trim(oOrdline.Item_No) <> "" And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
    '                If If_inventory_item_no(oOrdline.Item_No) And Not If_kits_item_no(oOrdline.Item_No) Then  'is inventory and is not kits
    '                    item_total_unit_fee = item_total_unit_fee + Get_nqg_price_by_item_no(oOrdline.Item_No, oOrdline.Unit_Price)
    '                End If
    '                If If_kits_item_no(oOrdline.Item_No) Then 'is inventory and is kits
    '                    orderline_kits_row_line_no = oOrdline.Line_No
    '                    orderline_kits_row_ord_guid = LTrim(RTrim(oOrdline.Ord_Guid))
    '                    orderline_kits_row_ord_qty_ordered = oOrdline.Qty_Ordered
    '                    For Each kits_components_oOrdline As cOrdLine In OrdLines
    '                        If (kits_components_oOrdline.Line_No = orderline_kits_row_line_no And
    '                            If_inventory_item_no(kits_components_oOrdline.Item_No)) And
    '                            Not If_kits_item_no(kits_components_oOrdline.Item_No) Then

    '                            item_total_unit_fee = item_total_unit_fee + Get_nqg_price_by_item_no(kits_components_oOrdline.Item_No, kits_components_oOrdline.Unit_Price)

    '                        End If
    '                    Next
    '                    orderline_kits_row_line_no = 0
    '                    orderline_kits_row_ord_guid = ""
    '                    orderline_kits_row_ord_qty_ordered = 0
    '                End If
    '            End If
    '        Next
    '        Get_all_inventory_item_total_unit_fee = item_total_unit_fee
    '    Catch er As Exception
    '        MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
    '    End Try
    'End Function

    'Public Function Get_all_inventory_item_total_price() As Double 'justin add 20230510
    '    Try
    '        Get_all_inventory_item_total_price = 0
    '        Dim item_total_price As Double = 0
    '        Dim orderline_kits_row_line_no As Int16 = 0
    '        Dim orderline_kits_row_ord_guid As String = ""
    '        Dim orderline_kits_row_ord_qty_ordered As Int16 = 0
    '        For Each oOrdline As cOrdLine In OrdLines
    '            If Trim(oOrdline.Item_No) <> "" And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
    '                If If_inventory_item_no(oOrdline.Item_No) And Not If_kits_item_no(oOrdline.Item_No) Then  'is inventory and is not kits
    '                    item_total_price = item_total_price + Get_nqg_price_by_item_no(oOrdline.Item_No, oOrdline.Unit_Price) * oOrdline.Qty_Ordered
    '                End If
    '                If If_kits_item_no(oOrdline.Item_No) Then 'is inventory and is kits
    '                    orderline_kits_row_line_no = oOrdline.Line_No
    '                    orderline_kits_row_ord_guid = LTrim(RTrim(oOrdline.Ord_Guid))
    '                    orderline_kits_row_ord_qty_ordered = oOrdline.Qty_Ordered
    '                    For Each kits_components_oOrdline As cOrdLine In OrdLines
    '                        If (kits_components_oOrdline.Line_No = orderline_kits_row_line_no And
    '                            If_inventory_item_no(kits_components_oOrdline.Item_No)) And
    '                            Not If_kits_item_no(kits_components_oOrdline.Item_No) Then

    '                            item_total_price = item_total_price + Get_nqg_price_by_item_no(kits_components_oOrdline.Item_No, kits_components_oOrdline.Unit_Price) * kits_components_oOrdline.Qty_Ordered

    '                        End If
    '                    Next
    '                    orderline_kits_row_line_no = 0
    '                    orderline_kits_row_ord_guid = ""
    '                    orderline_kits_row_ord_qty_ordered = 0
    '                End If
    '            End If
    '        Next
    '        Get_all_inventory_item_total_price = item_total_price
    '    Catch er As Exception
    '        MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
    '    End Try
    'End Function

    Public Function Add_44markup_per_ordline(ByRef _excel_Worksheet As Excel.Worksheet, ByRef _Oordline As cOrdLine, ByVal row_number As Int32, ByVal us_ng_unit_price As Double, ByVal _Qty As Double, ByVal in_Ordhead As cOrdHead) As Int16  'justin add 20230511 here return ipos
        Try
            Add_44markup_per_ordline = row_number

            Dim item_Markup_QTY_per_line As Double = 0
            Dim item_Markup_unit_price_per_line As Double = 0
            'if its inventory_line
            'If Trim(_Oordline.Item_No) <> "" And _Oordline.Qty_Ordered <> 0 And Not (_Oordline.Kit_Component) And If_inventory_item_no(_Oordline.Item_No) Then
            If Trim(_Oordline.Item_No) <> "" And _Qty <> 0 And (If_inventory_item_no(_Oordline.Item_No) Or If_kits_item_no(_Oordline.Item_No)) Then
                item_Markup_unit_price_per_line = us_ng_unit_price * 0.02
                item_Markup_QTY_per_line = _Qty
                Add_44markup_per_ordline = insert_44markup_row_by_line(_excel_Worksheet, _Oordline, row_number, item_Markup_QTY_per_line, item_Markup_unit_price_per_line, in_Ordhead)
            End If
        Catch er As Exception
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try
    End Function



    Public Function insert_44markup_row_by_line(ByRef excel_Worksheet As Excel.Worksheet, ByRef _Oordline As cOrdLine, ByVal row_number As Int32, ByVal _QTY As Double, ByVal _unit_price As Double, ByVal in_Ordhead As cOrdHead) As Int16  'here return Ipos
        Try
            insert_44markup_row_by_line = row_number
            Dim vg As New VariableGuid
            Dim Item_Guid As String = ""
            Item_Guid = vg.Guid(30)


            Dim errPos As Int16 = 0
            Dim iPos As Int16 = 0
            iPos = row_number
            Dim strSql As String
            Dim dt As DataTable
            Dim db As New cDBA

            strSql = " " &
                "SELECT 	I.Item_No, I.Item_Desc_1, I.Item_Desc_2, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd,  
                    ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg,  ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(L.Avg_Cost, 0) AS Avg_Cost,            
                    CASE WHEN ISNULL(X.Stocked_Fg, '') = '' THEN ISNULL(I.Stocked_FG, 'N') ELSE X.Stocked_Fg END AS Stocked_Fg,  
                    CASE WHEN ISNULL(X.BkOrd_Fg, '') = '' THEN ISNULL(I.BkOrd_Fg, 'N') ELSE X.BkOrd_Fg END AS BkOrd_Fg,
                    CASE WHEN SUBSTRING (I.Item_No, 1, 2) = '10' THEN 1 ELSE 0 END AS Is_Refill_Fg,  L.Loc,              
                    ISNULL(PK.Pack_CD, '') AS Packaging, L.status as Loc_Status 
			                    FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) 
			                    LEFT JOIN  MDB_ITEM_MASTER MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd 
			                    LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No 
			                    LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No 
			                    LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) 
			                    LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MI.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = 'US ' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 
			                    LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID 
                    WHERE 		I.Item_No = '44MARKUPUS1' AND I.ACTIVITY_CD = 'A' AND L.Loc = 'US1' ORDER BY   I.Item_No "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                ' iPos is the Y row position to write to in the Excel file.
                iPos += 1
                excel_Worksheet.Cells(iPos, xlsExport.Ord_Type) = "O" ' Ordhead.Ord_Type
                If Ordhead.Ord_Dt.Year <> 1 Then excel_Worksheet.Cells(iPos, xlsExport.Ord_Dt) =
                    Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") &
                    Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") &
                    Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
                excel_Worksheet.Cells(iPos, xlsExport.OE_PO_No).Style.NumberFormat = "@"
                excel_Worksheet.Cells(iPos, xlsExport.OE_PO_No) = Ordhead.Oe_Po_No
                errPos = 36

                'excel_Worksheet.Cells(iPos, xlsExport.Cus_No) = Ordhead.Cus_No
                excel_Worksheet.Cells(iPos, xlsExport.Cus_No) = in_Ordhead.Cus_No

                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Name) = Ordhead.Bill_To_Name
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_1) = Ordhead.Bill_To_Addr_1
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_2) = Ordhead.Bill_To_Addr_2
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_3) = Ordhead.Bill_To_Addr_3
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_4) = Ordhead.Bill_To_Addr_4
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Country) = Ordhead.Bill_To_Country

                If Ordhead.Cus_Alt_Adr_Cd <> "SAME" And Ordhead.Cus_Alt_Adr_Cd <> "" Then
                    excel_Worksheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
                Else
                    excel_Worksheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = ""
                End If

                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Name) = Ordhead.Ship_To_Name
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_1) = Ordhead.Ship_To_Addr_1
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_2) = Ordhead.Ship_To_Addr_2
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_3) = Ordhead.Ship_To_Addr_3
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_4) = Ordhead.Ship_To_Addr_4
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Country) = Ordhead.Ship_To_Country
                If Ordhead.Shipping_Dt.Date.Equals(NoDate()) Then
                    excel_Worksheet.Cells(iPos, xlsExport.Shipping_Dt) = "" ' "21111111"
                Else
                    excel_Worksheet.Cells(iPos, xlsExport.Shipping_Dt) =
                    Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") &
                    Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") &
                    Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0")
                End If

                excel_Worksheet.Cells(iPos, xlsExport.Ship_Via_Cd) = Ordhead.Ship_Via_Cd
                errPos = 56
                excel_Worksheet.Cells(iPos, xlsExport.Ar_Terms_Cd) = Ordhead.Ar_Terms_Cd
                excel_Worksheet.Cells(iPos, xlsExport.Ship_Instruction_1) = Ordhead.Ship_Instruction_1
                errPos = 58
                excel_Worksheet.Cells(iPos, xlsExport.Ship_Instruction_2) = Ordhead.Ship_Instruction_2
                excel_Worksheet.Cells(iPos, xlsExport.SlsPsn_No) = Ordhead.Slspsn_No
                excel_Worksheet.Cells(iPos, xlsExport.Discount_Pct) = Ordhead.Discount_Pct
                excel_Worksheet.Cells(iPos, xlsExport.Job_No) = Ordhead.Job_No

                excel_Worksheet.Cells(iPos, xlsExport.Mfg_Loc) = in_Ordhead.Mfg_Loc '20230809 by justin

                excel_Worksheet.Cells(iPos, xlsExport.Profit_Center) = Ordhead.Profit_Center
                excel_Worksheet.Cells(iPos, xlsExport.Filler_0001) = Ordhead.Ord_GUID
                excel_Worksheet.Cells(iPos, xlsExport.Curr_Cd) = Ordhead.Curr_Cd
                excel_Worksheet.Cells(iPos, xlsExport.Orig_Trx_Rt) = Ordhead.Orig_Trx_Rt
                excel_Worksheet.Cells(iPos, xlsExport.Curr_Trx_Rt) = Ordhead.Curr_Trx_Rt
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Sched) = Ordhead.Tax_Sched
                excel_Worksheet.Cells(iPos, xlsExport.Contact_1) = Ordhead.Contact_1
                excel_Worksheet.Cells(iPos, xlsExport.Phone_Number) = Ordhead.Phone_Number
                excel_Worksheet.Cells(iPos, xlsExport.Fax_Number) = Ordhead.Fax_Number
                excel_Worksheet.Cells(iPos, xlsExport.Email_Address) = Ordhead.Email_Address
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_City) = Ordhead.Ship_To_City
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_State) = Ordhead.Ship_To_State
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Zip) = Ordhead.Ship_To_Zip
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_City) = Ordhead.Bill_To_City
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_State) = Ordhead.Bill_To_State
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Zip) = Ordhead.Bill_To_Zip
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Cd) = Ordhead.Tax_Cd
                errPos = 80
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Pct) = Ordhead.Tax_Pct
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Cd_2) = Ordhead.Tax_Cd_2
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Pct_2) = Ordhead.Tax_Pct_2
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Cd_3) = Ordhead.Tax_Cd_3
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Pct_3) = Ordhead.Tax_Pct_3
                excel_Worksheet.Cells(iPos, xlsExport.Form_No) = Ordhead.Form_No
                errPos = 86
                excel_Worksheet.Cells(iPos, xlsExport.Deter_Rate_By) = Ordhead.Deter_Rate_By
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_1) = Ordhead.User_Def_Fld_1
                errPos = 88
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_2) = Ordhead.User_Def_Fld_2
                errPos = 89
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_3) = Ordhead.User_Def_Fld_3
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_4) = Ordhead.User_Def_Fld_4
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_5) = Ordhead.User_Def_Fld_5
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Fg) = Ordhead.Tax_Fg
                excel_Worksheet.Cells(iPos, xlsExport.Frt_Pay_Cd) = Ordhead.Frt_Pay_Cd
                excel_Worksheet.Cells(iPos, xlsExport.Selection_Cd) = Ordhead.Selection_Cd
                excel_Worksheet.Cells(iPos, xlsExport.Hold_Fg) = Ordhead.Hold_Fg

                excel_Worksheet.Cells(iPos, xlsExport.Item_No) = dt.Rows(0).Item("Item_No")
                errPos = 96
                excel_Worksheet.Cells(iPos, xlsExport.Item_Loc) = in_Ordhead.Mfg_Loc '20230809 by justin
                excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_Ordered) = _QTY
                errPos = 98
                excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = _QTY
                errPos = 99

                excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = 0

                '    errPos = 100

                excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Price) = _unit_price

                excel_Worksheet.Cells(iPos, xlsExport.Item_Discount_Pct) = 0
                excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Weight) = 0

                Dim _44mark_dt As String = Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") & Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") & Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
                excel_Worksheet.Cells(iPos, xlsExport.Item_Request_Dt) = _44mark_dt

                excel_Worksheet.Cells(iPos, xlsExport.Item_Promise_Dt) = _44mark_dt
                excel_Worksheet.Cells(iPos, xlsExport.Item_Req_Ship_Date) = _44mark_dt
                excel_Worksheet.Cells(iPos, xlsExport.Item_Guid) = Item_Guid

                ' The EOF_Marker to guid is set the line after the Next of the For Each Command
                excel_Worksheet.Cells(iPos, xlsExport.EOF_Marker) = "N"

                excel_Worksheet.Cells(iPos, xlsExport.Item_BkOrd_Fg) = dt.Rows(0).Item("BkOrd_Fg")

                excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_BkOrd_Fg) = "N" ' oOrdline.Qty_Bkord_Fg

                excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Cost) = 0

                excel_Worksheet.Cells(iPos, xlsExport.Item_Release_No) = "     1".ToString ' Forced at 1, inventory incremental 
                errPos = 116
                excel_Worksheet.Cells(iPos, xlsExport.Item_Desc_1) = dt.Rows(0).Item("Item_Desc_1") ' Forced at 1, inventory incremental 
                excel_Worksheet.Cells(iPos, xlsExport.Item_Desc_2) = dt.Rows(0).Item("Item_Desc_2") ' Forced at 1, inventory incremental 

                errPos = 118

                errPos = 120
                excel_Worksheet.Cells(iPos, xlsExport.TravelerRouteID) = "XX"
                excel_Worksheet.Cells(iPos, xlsExport.TravelerRoute) = "XX"
                excel_Worksheet.Cells(iPos, xlsExport.RouteCategory) = "XX"
                excel_Worksheet.Cells(iPos, xlsExport.Mfg_Item_No) = dt.Rows(0).Item("Item_No") & "-" & "XX" ' Changed June 26, 2017
                excel_Worksheet.Cells(iPos, xlsExport.ItemExtension) = "XX"

                errPos = 121
                'Added June 29, 2017 - T. Louzon
                excel_Worksheet.Cells(iPos, xlsExport.IsNoStock) = 0
                excel_Worksheet.Cells(iPos, xlsExport.IsNoArtwork) = 0

                errPos = 122
                'Added July 12, 2017 - T. Louzon
                excel_Worksheet.Cells(iPos, xlsExport.Prod_Cat) = dt.Rows(0).Item("Prod_Cat")


                errPos = 125

                'Added July 18, 2017 - T. Louzon
                excel_Worksheet.Cells(iPos, xlsExport.Num_Impr_1) = 0
                excel_Worksheet.Cells(iPos, xlsExport.Num_Impr_2) = 0
                excel_Worksheet.Cells(iPos, xlsExport.Num_Impr_3) = 0
                errPos = 127
            End If
            insert_44markup_row_by_line = iPos
        Catch er As Exception
            MsgBox("Error in COrder->insert_44markup_row_by_line." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Function

    Public Function Add_44MARKUPINV_per_ordline(ByRef _excel_Worksheet As Excel.Worksheet, ByRef _Oordline As cOrdLine, ByVal row_number As Int32, ByVal in_price As Double, ByVal _Qty As Double, ByVal in_Ordhead As cOrdHead) As Int16  'justin add 20230511 here return ipos
        Try
            Add_44MARKUPINV_per_ordline = row_number

            Dim item_Markup_QTY_per_line As Double = 0
            Dim item_Markup_unit_price_per_line As Double = 0
            If Trim(_Oordline.Item_No) <> "" And _Qty <> 0 And (If_inventory_item_no(_Oordline.Item_No) Or If_kits_item_no(_Oordline.Item_No)) Then
                item_Markup_unit_price_per_line = in_price * 0.08
                item_Markup_QTY_per_line = _Qty
                Add_44MARKUPINV_per_ordline = insert_44MARKUPINV_row_by_line(_excel_Worksheet, _Oordline, row_number, item_Markup_QTY_per_line, item_Markup_unit_price_per_line, in_Ordhead)
            End If
        Catch er As Exception
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try
    End Function
    Public Function insert_44MARKUPINV_row_by_line(ByRef excel_Worksheet As Excel.Worksheet, ByRef _Oordline As cOrdLine, ByVal row_number As Int32, ByVal _QTY As Double, ByVal _unit_price As Double, ByVal in_Ordhead As cOrdHead) As Int16  'here return Ipos
        Try
            insert_44MARKUPINV_row_by_line = row_number
            Dim vg As New VariableGuid
            Dim Item_Guid As String = ""
            Item_Guid = vg.Guid(30)


            Dim errPos As Int16 = 0
            Dim iPos As Int16 = 0
            iPos = row_number
            Dim strSql As String
            Dim dt As DataTable
            Dim db As New cDBA

            strSql = " " &
                "SELECT 	I.Item_No, I.Item_Desc_1, I.Item_Desc_2, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd,  
                    ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg,  ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(L.Avg_Cost, 0) AS Avg_Cost,            
                    CASE WHEN ISNULL(X.Stocked_Fg, '') = '' THEN ISNULL(I.Stocked_FG, 'N') ELSE X.Stocked_Fg END AS Stocked_Fg,  
                    CASE WHEN ISNULL(X.BkOrd_Fg, '') = '' THEN ISNULL(I.BkOrd_Fg, 'N') ELSE X.BkOrd_Fg END AS BkOrd_Fg,
                    CASE WHEN SUBSTRING (I.Item_No, 1, 2) = '10' THEN 1 ELSE 0 END AS Is_Refill_Fg,  L.Loc,              
                    ISNULL(PK.Pack_CD, '') AS Packaging, L.status as Loc_Status 
			                    FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) 
			                    LEFT JOIN  MDB_ITEM_MASTER MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd 
			                    LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No 
			                    LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No 
			                    LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) 
			                    LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MI.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = 'US ' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 
			                    LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID 
                    WHERE 		I.Item_No = '44MARKUPINV' AND I.ACTIVITY_CD = 'A' AND L.Loc = 'US1' ORDER BY   I.Item_No "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                ' iPos is the Y row position to write to in the Excel file.
                iPos += 1
                excel_Worksheet.Cells(iPos, xlsExport.Ord_Type) = "O" ' Ordhead.Ord_Type
                If Ordhead.Ord_Dt.Year <> 1 Then excel_Worksheet.Cells(iPos, xlsExport.Ord_Dt) =
                    Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") &
                    Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") &
                    Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
                excel_Worksheet.Cells(iPos, xlsExport.OE_PO_No).Style.NumberFormat = "@"
                excel_Worksheet.Cells(iPos, xlsExport.OE_PO_No) = Ordhead.Oe_Po_No
                errPos = 36
                excel_Worksheet.Cells(iPos, xlsExport.Cus_No) = Ordhead.Cus_No
                excel_Worksheet.Cells(iPos, xlsExport.Cus_No) = in_Ordhead.Cus_No
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Name) = Ordhead.Bill_To_Name
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_1) = Ordhead.Bill_To_Addr_1
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_2) = Ordhead.Bill_To_Addr_2
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_3) = Ordhead.Bill_To_Addr_3
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_4) = Ordhead.Bill_To_Addr_4
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Country) = Ordhead.Bill_To_Country

                If Ordhead.Cus_Alt_Adr_Cd <> "SAME" And Ordhead.Cus_Alt_Adr_Cd <> "" Then
                    excel_Worksheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
                Else
                    excel_Worksheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = ""
                End If

                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Name) = Ordhead.Ship_To_Name
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_1) = Ordhead.Ship_To_Addr_1
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_2) = Ordhead.Ship_To_Addr_2
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_3) = Ordhead.Ship_To_Addr_3
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_4) = Ordhead.Ship_To_Addr_4
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Country) = Ordhead.Ship_To_Country
                If Ordhead.Shipping_Dt.Date.Equals(NoDate()) Then
                    excel_Worksheet.Cells(iPos, xlsExport.Shipping_Dt) = "" ' "21111111"
                Else
                    excel_Worksheet.Cells(iPos, xlsExport.Shipping_Dt) =
                    Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") &
                    Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") &
                    Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0")
                End If

                excel_Worksheet.Cells(iPos, xlsExport.Ship_Via_Cd) = Ordhead.Ship_Via_Cd
                errPos = 56
                excel_Worksheet.Cells(iPos, xlsExport.Ar_Terms_Cd) = Ordhead.Ar_Terms_Cd
                excel_Worksheet.Cells(iPos, xlsExport.Ship_Instruction_1) = Ordhead.Ship_Instruction_1
                errPos = 58
                excel_Worksheet.Cells(iPos, xlsExport.Ship_Instruction_2) = Ordhead.Ship_Instruction_2
                excel_Worksheet.Cells(iPos, xlsExport.SlsPsn_No) = Ordhead.Slspsn_No
                excel_Worksheet.Cells(iPos, xlsExport.Discount_Pct) = Ordhead.Discount_Pct
                excel_Worksheet.Cells(iPos, xlsExport.Job_No) = Ordhead.Job_No
                excel_Worksheet.Cells(iPos, xlsExport.Mfg_Loc) = in_Ordhead.Mfg_Loc '20230809 justin
                excel_Worksheet.Cells(iPos, xlsExport.Profit_Center) = Ordhead.Profit_Center
                excel_Worksheet.Cells(iPos, xlsExport.Filler_0001) = Ordhead.Ord_GUID
                excel_Worksheet.Cells(iPos, xlsExport.Curr_Cd) = Ordhead.Curr_Cd
                excel_Worksheet.Cells(iPos, xlsExport.Orig_Trx_Rt) = Ordhead.Orig_Trx_Rt
                excel_Worksheet.Cells(iPos, xlsExport.Curr_Trx_Rt) = Ordhead.Curr_Trx_Rt
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Sched) = Ordhead.Tax_Sched
                excel_Worksheet.Cells(iPos, xlsExport.Contact_1) = Ordhead.Contact_1
                excel_Worksheet.Cells(iPos, xlsExport.Phone_Number) = Ordhead.Phone_Number
                excel_Worksheet.Cells(iPos, xlsExport.Fax_Number) = Ordhead.Fax_Number
                excel_Worksheet.Cells(iPos, xlsExport.Email_Address) = Ordhead.Email_Address
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_City) = Ordhead.Ship_To_City
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_State) = Ordhead.Ship_To_State
                excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Zip) = Ordhead.Ship_To_Zip
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_City) = Ordhead.Bill_To_City
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_State) = Ordhead.Bill_To_State
                excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Zip) = Ordhead.Bill_To_Zip
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Cd) = Ordhead.Tax_Cd
                errPos = 80
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Pct) = Ordhead.Tax_Pct
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Cd_2) = Ordhead.Tax_Cd_2
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Pct_2) = Ordhead.Tax_Pct_2
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Cd_3) = Ordhead.Tax_Cd_3
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Pct_3) = Ordhead.Tax_Pct_3
                excel_Worksheet.Cells(iPos, xlsExport.Form_No) = Ordhead.Form_No
                errPos = 86
                excel_Worksheet.Cells(iPos, xlsExport.Deter_Rate_By) = Ordhead.Deter_Rate_By
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_1) = Ordhead.User_Def_Fld_1
                errPos = 88
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_2) = Ordhead.User_Def_Fld_2
                errPos = 89
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_3) = Ordhead.User_Def_Fld_3
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_4) = Ordhead.User_Def_Fld_4
                excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_5) = Ordhead.User_Def_Fld_5
                excel_Worksheet.Cells(iPos, xlsExport.Tax_Fg) = Ordhead.Tax_Fg
                excel_Worksheet.Cells(iPos, xlsExport.Frt_Pay_Cd) = Ordhead.Frt_Pay_Cd
                excel_Worksheet.Cells(iPos, xlsExport.Selection_Cd) = Ordhead.Selection_Cd
                excel_Worksheet.Cells(iPos, xlsExport.Hold_Fg) = Ordhead.Hold_Fg

                excel_Worksheet.Cells(iPos, xlsExport.Item_No) = dt.Rows(0).Item("Item_No")
                errPos = 96
                excel_Worksheet.Cells(iPos, xlsExport.Item_Loc) = in_Ordhead.Mfg_Loc '20230809 justin
                excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_Ordered) = _QTY
                errPos = 98
                excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = _QTY
                errPos = 99

                excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = 0

                '    errPos = 100

                excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Price) = _unit_price

                excel_Worksheet.Cells(iPos, xlsExport.Item_Discount_Pct) = 0
                excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Weight) = 0

                Dim _44mark_dt As String = Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") & Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") & Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
                excel_Worksheet.Cells(iPos, xlsExport.Item_Request_Dt) = _44mark_dt

                excel_Worksheet.Cells(iPos, xlsExport.Item_Promise_Dt) = _44mark_dt
                excel_Worksheet.Cells(iPos, xlsExport.Item_Req_Ship_Date) = _44mark_dt
                excel_Worksheet.Cells(iPos, xlsExport.Item_Guid) = Item_Guid

                ' The EOF_Marker to guid is set the line after the Next of the For Each Command
                excel_Worksheet.Cells(iPos, xlsExport.EOF_Marker) = "N"

                excel_Worksheet.Cells(iPos, xlsExport.Item_BkOrd_Fg) = dt.Rows(0).Item("BkOrd_Fg")

                excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_BkOrd_Fg) = "N" ' oOrdline.Qty_Bkord_Fg

                excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Cost) = 0

                excel_Worksheet.Cells(iPos, xlsExport.Item_Release_No) = "     1".ToString ' Forced at 1, inventory incremental 
                errPos = 116
                excel_Worksheet.Cells(iPos, xlsExport.Item_Desc_1) = dt.Rows(0).Item("Item_Desc_1") ' Forced at 1, inventory incremental 
                excel_Worksheet.Cells(iPos, xlsExport.Item_Desc_2) = dt.Rows(0).Item("Item_Desc_2") ' Forced at 1, inventory incremental 

                errPos = 118

                errPos = 120
                excel_Worksheet.Cells(iPos, xlsExport.TravelerRouteID) = "XX"
                excel_Worksheet.Cells(iPos, xlsExport.TravelerRoute) = "XX"
                excel_Worksheet.Cells(iPos, xlsExport.RouteCategory) = "XX"
                excel_Worksheet.Cells(iPos, xlsExport.Mfg_Item_No) = dt.Rows(0).Item("Item_No") & "-" & "XX" ' Changed June 26, 2017
                excel_Worksheet.Cells(iPos, xlsExport.ItemExtension) = "XX"

                errPos = 121
                'Added June 29, 2017 - T. Louzon
                excel_Worksheet.Cells(iPos, xlsExport.IsNoStock) = 0
                excel_Worksheet.Cells(iPos, xlsExport.IsNoArtwork) = 0

                errPos = 122
                'Added July 12, 2017 - T. Louzon
                excel_Worksheet.Cells(iPos, xlsExport.Prod_Cat) = dt.Rows(0).Item("Prod_Cat")


                errPos = 125

                'Added July 18, 2017 - T. Louzon
                excel_Worksheet.Cells(iPos, xlsExport.Num_Impr_1) = 0
                excel_Worksheet.Cells(iPos, xlsExport.Num_Impr_2) = 0
                excel_Worksheet.Cells(iPos, xlsExport.Num_Impr_3) = 0
                errPos = 127
            End If
            insert_44MARKUPINV_row_by_line = iPos
        Catch er As Exception
            MsgBox("Error in COrder->insert_44MARKUPINV_line." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Function

    Public Function Get_all_inventory_item_total_qty() As Double
        Try
            Dim item_total_qty As Double = 0
            Dim orderline_kits_row_ord_guid As String = ""
            Dim orderline_kits_row_ord_qty_ordered As Int16 = 0
            For Each oOrdline As cOrdLine In OrdLines
                If Trim(oOrdline.Item_No) <> "" And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
                    If If_inventory_item_no(oOrdline.Item_No) And Not If_kits_item_no(oOrdline.Item_No) Then  'is inventory and is not kits
                        item_total_qty = item_total_qty + oOrdline.Qty_Ordered
                    End If
                    If If_kits_item_no(oOrdline.Item_No) Then 'is inventory and is kits
                        item_total_qty = item_total_qty + oOrdline.Qty_Ordered
                    End If
                End If
            Next
            Get_all_inventory_item_total_qty = item_total_qty
        Catch er As Exception
            MsgBox("Error In COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try
    End Function


    Public Function If_inventory_item_no(ByVal in_Item_no As String) As Boolean
        Try
            If_inventory_item_no = False
            Dim Item_no = LTrim(RTrim(in_Item_no))

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String

            strSql = "Select top 1 ltrim(rtrim(isnull(stocked_fg,''))) as stocked_fg FROM [200].dbo.IMITMIDX_SQL WITH (NOLOCK) WHERE ltrim(rtrim(isnull(stocked_fg,''))) = 'Y' and ITEM_NO='" & in_Item_no & "' "
            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                If dt.Rows(0).Item("stocked_fg") = "Y" Then
                    If_inventory_item_no = True
                End If
            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try


    End Function

    Public Function If_kits_item_no(ByVal in_Item_no As String) As Boolean
        Try
            If_kits_item_no = False
            Dim Item_no = LTrim(RTrim(in_Item_no))

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String

            strSql = "SELECT top 1 ltrim(rtrim(isnull(kit_feat_fg,''))) as kit_feat_fg FROM [200].dbo.IMITMIDX_SQL WITH (NOLOCK) WHERE ltrim(rtrim(isnull(stocked_fg,''))) != 'Y' and ITEM_NO='" & in_Item_no & "' "
            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                If dt.Rows(0).Item("kit_feat_fg") = "K" Then
                    If_kits_item_no = True
                End If
            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " ")
        End Try


    End Function



    'justin add for us_order excel file 20230221
    Public Function CreateUsOrderExcelFile(ByRef UsOrder_Ordhead As cOrdHead) As String

        Dim errPos As Integer

        Dim xlApp As Excel.Application
        xlApp = New Excel.ApplicationClass

        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Try
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets(1)

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String

            Dim strLastGuid As String = ""
            Dim iPos As Integer = 0
            Dim iCount As Integer = 0


            For Each oOrdline As cOrdLine In OrdLines
                ' As we do not have now line validation, validate empty items here 
                ' and do not include them in the excel export file.
                ' iCount and iPos are different because of kits. We do not insert line for
                ' kit components but they are part of the order because of traveler and imprint.
                ' iCount is the number of lines in the order
                iCount += 1
                If Trim(oOrdline.Item_No) <> "" And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
                    strSql = "SELECT ITEM_NO FROM  [200].dbo.IMITMIDX_SQL WITH (NOLOCK) WHERE ITEM_NO='" & SqlCompliantString(Trim(oOrdline.Item_No)) & "' "
                    dt = db.DataTable(strSql)
                    If dt.Rows.Count <> 0 Then
                        ' iPos is the Y row position to write to in the Excel file.
                        strLastGuid = oOrdline.Item_Guid
                        iPos += 1
                        xlWorkSheet.Cells(iPos, xlsExport.Ord_Type) = "O" ' Ordhead.Ord_Type
                        If Ordhead.Ord_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Ord_Dt) =
                            Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") &
                            Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") &
                            Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
                        xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No).Style.NumberFormat = "@"
                        xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No) = Ordhead.Oe_Po_No
                        errPos = 36
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'xlWorkSheet.Cells(iPos, xlsExport.Cus_No) = Ordhead.Cus_No 'justin 20230809 update to 
                        xlWorkSheet.Cells(iPos, xlsExport.Cus_No) = UsOrder_Ordhead.Cus_No
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Name) = Ordhead.Bill_To_Name
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_1) = Ordhead.Bill_To_Addr_1
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_2) = Ordhead.Bill_To_Addr_2
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_3) = Ordhead.Bill_To_Addr_3
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_4) = Ordhead.Bill_To_Addr_4
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Country) = Ordhead.Bill_To_Country

                        If Ordhead.Cus_Alt_Adr_Cd <> "SAME" And Ordhead.Cus_Alt_Adr_Cd <> "" Then
                            xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
                        Else
                            xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = ""
                        End If

                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Name) = Ordhead.Ship_To_Name
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_1) = Ordhead.Ship_To_Addr_1
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_2) = Ordhead.Ship_To_Addr_2
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_3) = Ordhead.Ship_To_Addr_3
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_4) = Ordhead.Ship_To_Addr_4
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Country) = Ordhead.Ship_To_Country
                        If Ordhead.Shipping_Dt.Date.Equals(NoDate()) Then
                            xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) = "" ' "21111111"
                        Else
                            xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) =
                            Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") &
                            Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") &
                            Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0")
                        End If

                        xlWorkSheet.Cells(iPos, xlsExport.Ship_Via_Cd) = Ordhead.Ship_Via_Cd
                        errPos = 56
                        xlWorkSheet.Cells(iPos, xlsExport.Ar_Terms_Cd) = Ordhead.Ar_Terms_Cd
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_1) = Ordhead.Ship_Instruction_1
                        errPos = 58
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_2) = Ordhead.Ship_Instruction_2
                        xlWorkSheet.Cells(iPos, xlsExport.SlsPsn_No) = Ordhead.Slspsn_No

                        '20230727 justin modify par ticket 34701  commented on 20230906 because all header's discount pct is 0 in oei_ordheader
                        'xlWorkSheet.Cells(iPos, xlsExport.Discount_Pct) = 0
                        xlWorkSheet.Cells(iPos, xlsExport.Discount_Pct) = Ordhead.Discount_Pct

                        xlWorkSheet.Cells(iPos, xlsExport.Job_No) = Ordhead.Job_No
                        'justin 20230809
                        'xlWorkSheet.Cells(iPos, xlsExport.Mfg_Loc) = Ordhead.Mfg_Loc
                        xlWorkSheet.Cells(iPos, xlsExport.Mfg_Loc) = UsOrder_Ordhead.Mfg_Loc

                        xlWorkSheet.Cells(iPos, xlsExport.Profit_Center) = Ordhead.Profit_Center
                        xlWorkSheet.Cells(iPos, xlsExport.Filler_0001) = Ordhead.Ord_GUID
                        xlWorkSheet.Cells(iPos, xlsExport.Curr_Cd) = Ordhead.Curr_Cd
                        xlWorkSheet.Cells(iPos, xlsExport.Orig_Trx_Rt) = Ordhead.Orig_Trx_Rt
                        xlWorkSheet.Cells(iPos, xlsExport.Curr_Trx_Rt) = Ordhead.Curr_Trx_Rt
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Sched) = Ordhead.Tax_Sched
                        xlWorkSheet.Cells(iPos, xlsExport.Contact_1) = Ordhead.Contact_1
                        xlWorkSheet.Cells(iPos, xlsExport.Phone_Number) = Ordhead.Phone_Number
                        xlWorkSheet.Cells(iPos, xlsExport.Fax_Number) = Ordhead.Fax_Number
                        xlWorkSheet.Cells(iPos, xlsExport.Email_Address) = Ordhead.Email_Address
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_City) = Ordhead.Ship_To_City
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_State) = Ordhead.Ship_To_State
                        xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Zip) = Ordhead.Ship_To_Zip
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_City) = Ordhead.Bill_To_City
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_State) = Ordhead.Bill_To_State
                        xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Zip) = Ordhead.Bill_To_Zip
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd) = Ordhead.Tax_Cd
                        errPos = 80
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct) = Ordhead.Tax_Pct
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_2) = Ordhead.Tax_Cd_2
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_2) = Ordhead.Tax_Pct_2
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_3) = Ordhead.Tax_Cd_3
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_3) = Ordhead.Tax_Pct_3
                        xlWorkSheet.Cells(iPos, xlsExport.Form_No) = Ordhead.Form_No
                        errPos = 86
                        xlWorkSheet.Cells(iPos, xlsExport.Deter_Rate_By) = Ordhead.Deter_Rate_By
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_1) = Ordhead.User_Def_Fld_1
                        errPos = 88
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_2) = Ordhead.User_Def_Fld_2
                        errPos = 89
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_3) = Ordhead.User_Def_Fld_3
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_4) = Ordhead.User_Def_Fld_4
                        xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_5) = Ordhead.User_Def_Fld_5
                        xlWorkSheet.Cells(iPos, xlsExport.Tax_Fg) = Ordhead.Tax_Fg
                        xlWorkSheet.Cells(iPos, xlsExport.Frt_Pay_Cd) = Ordhead.Frt_Pay_Cd
                        xlWorkSheet.Cells(iPos, xlsExport.Selection_Cd) = Ordhead.Selection_Cd
                        xlWorkSheet.Cells(iPos, xlsExport.Hold_Fg) = Ordhead.Hold_Fg
                        xlWorkSheet.Cells(iPos, xlsExport.Item_No) = oOrdline.Item_No
                        errPos = 96
                        'justin 20230809
                        'xlWorkSheet.Cells(iPos, xlsExport.Item_Loc) = oOrdline.Loc
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Loc) = UsOrder_Ordhead.Mfg_Loc

                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_Ordered) = oOrdline.Qty_Ordered
                        errPos = 98
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = oOrdline.Qty_To_Ship
                        errPos = 99
                        If oOrdline.Qty_To_Ship >= oOrdline.Qty_Ordered Then
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = 0
                        Else
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = oOrdline.Qty_Ordered - oOrdline.Qty_To_Ship
                        End If
                        errPos = 100

                        Dim _oOrdline_US_Unit_Price As Double = 0

                        Dim _oOrdline_US_avg_cost As Double = 0
                        Dim _oOrdline_US_nqg_price As Double = 0
                        '------------------------------------------------------------------------------------------------------
                        'justin add 20230808 for inventory transfers ticket:34919
                        If UsOrder_Ordhead.Cus_No = "SPECTCA" Then
                            If If_kits_item_no(oOrdline.Item_No) Then
                                _oOrdline_US_Unit_Price = Get_avg_cost_by_kits_In200(oOrdline.Item_No, oOrdline.Unit_Price)
                            Else
                                If (Get_commodity_cd_by_item_no(oOrdline.Item_No) <> "Y" And If_inventory_item_no(oOrdline.Item_No) = False) Then
                                    _oOrdline_US_Unit_Price = 0
                                Else
                                    If If_inventory_item_no(oOrdline.Item_No) = False Then '20230512
                                        _oOrdline_US_Unit_Price = oOrdline.Unit_Price
                                    Else
                                        _oOrdline_US_Unit_Price = Get_avg_cost_by_item_no_In200(oOrdline.Item_No, oOrdline.Unit_Price)
                                    End If
                                End If
                            End If
                            _oOrdline_US_avg_cost = _oOrdline_US_Unit_Price
                        Else '''''''''''''''''''''''''''''''''''''''''''''''''''
                            If If_kits_item_no(oOrdline.Item_No) Then
                                _oOrdline_US_Unit_Price = Get_nqg_price_by_kits(oOrdline.Item_No, oOrdline.Unit_Price) 'justin add 20230221
                            Else
                                If (Get_commodity_cd_by_item_no(oOrdline.Item_No) <> "Y" And If_inventory_item_no(oOrdline.Item_No) = False) Then
                                    _oOrdline_US_Unit_Price = 0
                                Else
                                    If If_inventory_item_no(oOrdline.Item_No) = False Then '20230512
                                        _oOrdline_US_Unit_Price = oOrdline.Unit_Price
                                    Else
                                        _oOrdline_US_Unit_Price = Get_nqg_price_by_item_no(oOrdline.Item_No, oOrdline.Unit_Price)
                                    End If
                                End If
                            End If
                            _oOrdline_US_nqg_price = _oOrdline_US_Unit_Price
                        End If '''''''''''''''''''''''''''''''''''''''''''''''''''
                        '------------------------------------------------------------------------------------------------------



                        xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Price) = _oOrdline_US_Unit_Price

                        '20230906 justin modify par ticket 34701  
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Discount_Pct) = 0
                        'xlWorkSheet.Cells(iPos, xlsExport.Item_Discount_Pct) = oOrdline.Discount_Pct
                        '

                        xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Weight) = oOrdline.Unit_Weight

                        If oOrdline.Request_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Request_Dt) =
                                oOrdline.Request_Dt.Year.ToString.PadLeft(4, "0") &
                                oOrdline.Request_Dt.Month.ToString.PadLeft(2, "0") &
                                oOrdline.Request_Dt.Day.ToString.PadLeft(2, "0")
                        If oOrdline.Promise_Dt.Equals(Date.MinValue) Then
                            oOrdline.Promise_Dt = Ordhead.Ord_Dt_Shipped
                        End If

                        If oOrdline.Promise_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Promise_Dt) =
                                oOrdline.Promise_Dt.Year.ToString.PadLeft(4, "0") &
                                oOrdline.Promise_Dt.Month.ToString.PadLeft(2, "0") &
                                oOrdline.Promise_Dt.Day.ToString.PadLeft(2, "0")
                        If oOrdline.Req_Ship_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Req_Ship_Date) =
                                oOrdline.Req_Ship_Dt.Year.ToString.PadLeft(4, "0") &
                                oOrdline.Req_Ship_Dt.Month.ToString.PadLeft(2, "0") &
                                oOrdline.Req_Ship_Dt.Day.ToString.PadLeft(2, "0")
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Guid) = oOrdline.Item_Guid

                        ' The EOF_Marker to guid is set the line after the Next of the For Each Command
                        xlWorkSheet.Cells(iPos, xlsExport.EOF_Marker) = "N"

                        xlWorkSheet.Cells(iPos, xlsExport.User_Def_Fld_4) = oOrdline.User_Def_Fld_4
                        xlWorkSheet.Cells(iPos, xlsExport.Item_BkOrd_Fg) = oOrdline.Bkord_Fg
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Bin_Fg) = oOrdline.Bin_Fg
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Prc_Cd_Orig_Price) = oOrdline.Orig_Price
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd_Fg) = IIf(oOrdline.Qty_Prev_Bkord <> 0, "Y", "N") ' oOrdline.Qty_Bkord_Fg
                        'oOrdline()
                        ' Prevent division by zero
                        If Ordhead.Curr_Trx_Rt = 0 Then
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = oOrdline.Unit_Cost  ' Forced at 1, inventory incremental 
                        Else
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = Format(oOrdline.Unit_Cost / Ordhead.Curr_Trx_Rt, "##,##0.000000") ' Forced at 1, inventory incremental 
                        End If
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Release_No) = "     1".ToString ' Forced at 1, inventory incremental 
                        errPos = 116
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_1) = oOrdline.Item_Desc_1 ' Forced at 1, inventory incremental 
                        xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_2) = oOrdline.Item_Desc_2 ' Forced at 1, inventory incremental 
                        xlWorkSheet.Cells(iPos, 98) = oOrdline.Extra_2 ' Forced at 1, inventory incremental 
                        xlWorkSheet.Cells(iPos, 99) = oOrdline.Extra_6 ' Forced at 1, inventory incremental 
                        errPos = 118


                        '==============================================
                        'Added June 17, 2017 - T. Louzon

                        'GET THE Item Extension for the Route ID
                        Dim dt2 As DataTable
                        Dim db2 As New cDBA
                        Dim strSql2 As String
                        Dim sItemExtension As String
                        Dim sRouteCategory As String
                        strSql2 = " SELECT XREF.ItemExtension, T.RouteCategory "
                        strSql2 = strSql2 & " FROM Exact_traveler_Route T "
                        strSql2 = strSql2 & " INNER JOIN EXACT_Traveler_Route XREF  "    'HV_OEI_RouteCategory_XREF XREF "  'Changed July 18, 2017
                        strSql2 = strSql2 & " On T.RouteCategory = XREF.RouteCategory "
                        strSql2 = strSql2 & " WHERE XREF.routeID = " & oOrdline.m_RouteID

                        errPos = 119
                        dt2 = db2.DataTable(strSql2)
                        If dt2.Rows.Count <> 0 Then

                            If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                sItemExtension = "XX"
                            Else
                                sItemExtension = dt2.Rows(0).Item("ItemExtension")
                            End If

                            If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                sRouteCategory = "XX"
                            Else
                                sRouteCategory = dt2.Rows(0).Item("RouteCategory")
                            End If

                        Else
                            sItemExtension = "XX"
                            sRouteCategory = "XX"
                        End If

                        dt2.Dispose()

                        errPos = 120
                        xlWorkSheet.Cells(iPos, xlsExport.TravelerRouteID) = oOrdline.m_RouteID
                        xlWorkSheet.Cells(iPos, xlsExport.TravelerRoute) = oOrdline.m_Route
                        xlWorkSheet.Cells(iPos, xlsExport.RouteCategory) = sRouteCategory ' 
                        xlWorkSheet.Cells(iPos, xlsExport.Mfg_Item_No) = oOrdline.Item_No & "-" & sItemExtension ' Changed June 26, 2017
                        xlWorkSheet.Cells(iPos, xlsExport.ItemExtension) = sItemExtension '  

                        errPos = 121
                        'Added June 29, 2017 - T. Louzon
                        xlWorkSheet.Cells(iPos, xlsExport.IsNoStock) = oOrdline.IsNoStock
                        xlWorkSheet.Cells(iPos, xlsExport.IsNoArtwork) = oOrdline.IsNoArtwork

                        errPos = 122
                        'Added July 12, 2017 - T. Louzon
                        xlWorkSheet.Cells(iPos, xlsExport.Prod_Cat) = oOrdline.Prod_Cat
                        xlWorkSheet.Cells(iPos, xlsExport.End_Item_Cd) = oOrdline.End_Item_Cd

                        'GET THE EXTRA FIELDS TO EXPORT
                        '==============================================
                        'Added June 17, 2017 - T.Louzon
                        Dim dt3 As DataTable
                        Dim db3 As New cDBA
                        Dim strSql3 As String
                        'Dim sItemExtension As String
                        'Dim sRouteCategory As String
                        strSql3 = " SELECT   "
                        strSql3 = strSql3 & " item_no, "
                        strSql3 = strSql3 & " num_impr_1, "
                        strSql3 = strSql3 & " num_impr_2, "
                        strSql3 = strSql3 & " num_impr_3, "
                        strSql3 = strSql3 & " imprint = extra_1, "
                        strSql3 = strSql3 & " Imprint_Location = Extra_2, "
                        strSql3 = strSql3 & " imprint_Color = Extra_3, "
                        strSql3 = strSql3 & " NUM = Extra_4,  "       '-- Number of Imprints
                        strSql3 = strSql3 & " Packaging = Extra_5, "
                        strSql3 = strSql3 & " Refill = Extra_8, "
                        strSql3 = strSql3 & " Laser_Setup = Extra_9, "
                        strSql3 = strSql3 & " comments = Comment, "
                        strSql3 = strSql3 & " special_Comments = Comment2, "
                        strSql3 = strSql3 & " Industry = Industry, "
                        strSql3 = strSql3 & " Printer_Instructions, "
                        strSql3 = strSql3 & " Packer_Instructions "

                        strSql3 = strSql3 & " FROM OEI_EXTRA_FIELDS "
                        strSql3 = strSql3 & " WHERE ITEM_GUID = '" & oOrdline.m_Item_Guid & "'"

                        dt3 = db3.DataTable(strSql3)

                        errPos = 125

                        If dt3.Rows.Count <> 0 Then
                            errPos = 126

                            'Added July 18, 2017 - T. Louzon
                            xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_1) = dt3.Rows(0).Item("Num_Impr_1")
                            xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_2) = dt3.Rows(0).Item("Num_Impr_2")
                            xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_3) = dt3.Rows(0).Item("Num_Impr_3")
                            xlWorkSheet.Cells(iPos, xlsExport.Imprint) = dt3.Rows(0).Item("Imprint")    ''
                            xlWorkSheet.Cells(iPos, xlsExport.Imprint_Location) = dt3.Rows(0).Item("Imprint_Location")
                            xlWorkSheet.Cells(iPos, xlsExport.Imprint_Color) = dt3.Rows(0).Item("Imprint_Color")
                            xlWorkSheet.Cells(iPos, xlsExport.Comments) = dt3.Rows(0).Item("Comments")
                            xlWorkSheet.Cells(iPos, xlsExport.Special_Comments) = dt3.Rows(0).Item("Special_Comments")
                            xlWorkSheet.Cells(iPos, xlsExport.Printer_Instructions) = dt3.Rows(0).Item("Printer_Instructions")
                            xlWorkSheet.Cells(iPos, xlsExport.Packer_Instructions) = dt3.Rows(0).Item("Packer_Instructions")
                        Else
                            'Nothing
                        End If
                        errPos = 127
                        dt3.Dispose()

                        '''''''''''''
                        'iPos = Add_44markup_per_ordline(xlWorkSheet, oOrdline, iPos, _oOrdline_US_Unit_Price, oOrdline.Qty_Ordered, UsOrder_Ordhead) 'here beed confirm from neg_price or avg_cost?
                        'if customer is SPECTCA _oOrdline_US_Unit_Price should be _oOrdline_US_avg_cost  above line 3503
                        'if other customer      _oOrdline_US_Unit_Priceshould be _oOrdline_US_nqg_price  above line 3503
                        '''''''''''''

                        If UsOrder_Ordhead.Cus_No = "SPECTCA" Then
                            iPos = Add_44MARKUPINV_per_ordline(xlWorkSheet, oOrdline, iPos, _oOrdline_US_avg_cost, oOrdline.Qty_Ordered, UsOrder_Ordhead)
                        Else
                            iPos = Add_44markup_per_ordline(xlWorkSheet, oOrdline, iPos, _oOrdline_US_Unit_Price, oOrdline.Qty_Ordered, UsOrder_Ordhead) 'here beed confirm from neg_price or avg_cost?
                        End If
                        '==============================================
                    End If
                Else
                    iCount -= 1
                End If

                errPos = 150

            Next

            ' This must be used so when the last item is a kit component, it sets correctly the last line to guid
            errPos = 151
            '** IF THE PROGRAM CRASHES HERE, THEN EVERY LINE HAS A QTY OF 0 **
            xlWorkSheet.Cells(iPos, xlsExport.EOF_Marker) = Ordhead.Ord_GUID


            errPos = 152

            'justin 20230221 add 44MARKUPUS1 item row in the end of row of excel
            'Call Add_44MARKUPUS1(xlWorkSheet, iPos)

            Dim m_USpath As String = HV_US_Import_Files_Folder()

            ''''for test
            'm_USpath = "C:\ExactTemp\2\"
            'xlWorkBook.SaveAs(Filename:="" & m_USpath & Ordhead.Ord_GUID & "_2_SO_US.xls")
            '''
            xlWorkBook.SaveAs(Filename:="" & m_USpath & Ordhead.Ord_GUID & ".xls")



            errPos = 153
            xlWorkBook.Close()
            xlWorkBook = Nothing
            errPos = 154

            xlApp.Quit()
            xlApp = Nothing

            errPos = 155

            'MsgBox("File created.")
        Catch er As Exception
            Call UnsetExportTS()
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " " & errPos.ToString)
            xlWorkBook.Close(False)
            xlApp.Quit()
        End Try

    End Function
    '

    Public Function GetUsOrder_Unit_price(ByVal Item_No As String) As Double 'justin add 20230221
        Try

            GetUsOrder_Unit_price = 0

            'get from Betty's method
            GetUsOrder_Unit_price = 9

        Catch er As Exception
            MsgBox("Error in COrder->GetUsOrder_Unit_price." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Function

    'Public Sub Add_44MARKUPUS1(ByRef excel_Worksheet As Excel.Worksheet, ByVal row_number As Int32)  'justin add 20230221
    '    Try

    '        Dim vg As New VariableGuid
    '        Dim Item_Guid As String = ""
    '        Item_Guid = vg.Guid(30)


    '        Dim errPos As Int16 = 0
    '        Dim iPos As Int16 = 0
    '        iPos = row_number
    '        Dim strSql As String
    '        Dim dt As DataTable
    '        Dim db As New cDBA

    '        strSql = " " &
    '            "SELECT 	I.Item_No, I.Item_Desc_1, I.Item_Desc_2, ISNULL(I.User_Def_Fld_1, '') AS Item_Cd,  
    '                ISNULL(I.Kit_Feat_Fg, '') AS Kit_Feat_Fg,  ISNULL(I.Prod_Cat, '') AS Prod_Cat, ISNULL(L.Avg_Cost, 0) AS Avg_Cost,            
    '                CASE WHEN ISNULL(X.Stocked_Fg, '') = '' THEN ISNULL(I.Stocked_FG, 'N') ELSE X.Stocked_Fg END AS Stocked_Fg,  
    '                CASE WHEN ISNULL(X.BkOrd_Fg, '') = '' THEN ISNULL(I.BkOrd_Fg, 'N') ELSE X.BkOrd_Fg END AS BkOrd_Fg,
    '                CASE WHEN SUBSTRING (I.Item_No, 1, 2) = '10' THEN 1 ELSE 0 END AS Is_Refill_Fg,  L.Loc,              
    '                ISNULL(PK.Pack_CD, '') AS Packaging, L.status as Loc_Status 
    '                   FROM 		[200].dbo.IMITMIDX_SQL I WITH (nolock) 
    '                   LEFT JOIN  MDB_ITEM_MASTER MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd 
    '                   LEFT JOIN  OEI_ItmIdx X WITH (Nolock) ON I.Item_No = X.Item_No 
    '                   LEFT JOIN  [200].dbo.OEI_Item_Loc_Qty_View_US L ON I.Item_No = L.Item_No 
    '                   LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) 
    '                   LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MI.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = 'US ' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 
    '                   LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID 
    '                WHERE 		I.Item_No = '44MARKUPUS1' AND I.ACTIVITY_CD = 'A' AND L.Loc = 'US1' ORDER BY   I.Item_No "

    '        dt = db.DataTable(strSql)
    '        If dt.Rows.Count <> 0 Then
    '            ' iPos is the Y row position to write to in the Excel file.
    '            iPos += 1
    '            excel_Worksheet.Cells(iPos, xlsExport.Ord_Type) = "O" ' Ordhead.Ord_Type
    '            If Ordhead.Ord_Dt.Year <> 1 Then excel_Worksheet.Cells(iPos, xlsExport.Ord_Dt) =
    '                Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") &
    '                Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") &
    '                Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
    '            excel_Worksheet.Cells(iPos, xlsExport.OE_PO_No).Style.NumberFormat = "@"
    '            excel_Worksheet.Cells(iPos, xlsExport.OE_PO_No) = Ordhead.Oe_Po_No
    '            errPos = 36
    '            excel_Worksheet.Cells(iPos, xlsExport.Cus_No) = Ordhead.Cus_No
    '            excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Name) = Ordhead.Bill_To_Name
    '            excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_1) = Ordhead.Bill_To_Addr_1
    '            excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_2) = Ordhead.Bill_To_Addr_2
    '            excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_3) = Ordhead.Bill_To_Addr_3
    '            excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Addr_4) = Ordhead.Bill_To_Addr_4
    '            excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Country) = Ordhead.Bill_To_Country

    '            If Ordhead.Cus_Alt_Adr_Cd <> "SAME" And Ordhead.Cus_Alt_Adr_Cd <> "" Then
    '                excel_Worksheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
    '            Else
    '                excel_Worksheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = ""
    '            End If

    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Name) = Ordhead.Ship_To_Name
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_1) = Ordhead.Ship_To_Addr_1
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_2) = Ordhead.Ship_To_Addr_2
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_3) = Ordhead.Ship_To_Addr_3
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Addr_4) = Ordhead.Ship_To_Addr_4
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Country) = Ordhead.Ship_To_Country
    '            If Ordhead.Shipping_Dt.Date.Equals(NoDate()) Then
    '                excel_Worksheet.Cells(iPos, xlsExport.Shipping_Dt) = "" ' "21111111"
    '            Else
    '                excel_Worksheet.Cells(iPos, xlsExport.Shipping_Dt) =
    '                Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") &
    '                Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") &
    '                Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0")
    '            End If

    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_Via_Cd) = Ordhead.Ship_Via_Cd
    '            errPos = 56
    '            excel_Worksheet.Cells(iPos, xlsExport.Ar_Terms_Cd) = Ordhead.Ar_Terms_Cd
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_Instruction_1) = Ordhead.Ship_Instruction_1
    '            errPos = 58
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_Instruction_2) = Ordhead.Ship_Instruction_2
    '            excel_Worksheet.Cells(iPos, xlsExport.SlsPsn_No) = Ordhead.Slspsn_No
    '            excel_Worksheet.Cells(iPos, xlsExport.Discount_Pct) = Ordhead.Discount_Pct
    '            excel_Worksheet.Cells(iPos, xlsExport.Job_No) = Ordhead.Job_No
    '            excel_Worksheet.Cells(iPos, xlsExport.Mfg_Loc) = Ordhead.Mfg_Loc
    '            excel_Worksheet.Cells(iPos, xlsExport.Profit_Center) = Ordhead.Profit_Center
    '            excel_Worksheet.Cells(iPos, xlsExport.Filler_0001) = Ordhead.Ord_GUID
    '            excel_Worksheet.Cells(iPos, xlsExport.Curr_Cd) = Ordhead.Curr_Cd
    '            excel_Worksheet.Cells(iPos, xlsExport.Orig_Trx_Rt) = Ordhead.Orig_Trx_Rt
    '            excel_Worksheet.Cells(iPos, xlsExport.Curr_Trx_Rt) = Ordhead.Curr_Trx_Rt
    '            excel_Worksheet.Cells(iPos, xlsExport.Tax_Sched) = Ordhead.Tax_Sched
    '            excel_Worksheet.Cells(iPos, xlsExport.Contact_1) = Ordhead.Contact_1
    '            excel_Worksheet.Cells(iPos, xlsExport.Phone_Number) = Ordhead.Phone_Number
    '            excel_Worksheet.Cells(iPos, xlsExport.Fax_Number) = Ordhead.Fax_Number
    '            excel_Worksheet.Cells(iPos, xlsExport.Email_Address) = Ordhead.Email_Address
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_To_City) = Ordhead.Ship_To_City
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_To_State) = Ordhead.Ship_To_State
    '            excel_Worksheet.Cells(iPos, xlsExport.Ship_To_Zip) = Ordhead.Ship_To_Zip
    '            excel_Worksheet.Cells(iPos, xlsExport.Bill_To_City) = Ordhead.Bill_To_City
    '            excel_Worksheet.Cells(iPos, xlsExport.Bill_To_State) = Ordhead.Bill_To_State
    '            excel_Worksheet.Cells(iPos, xlsExport.Bill_To_Zip) = Ordhead.Bill_To_Zip
    '            excel_Worksheet.Cells(iPos, xlsExport.Tax_Cd) = Ordhead.Tax_Cd
    '            errPos = 80
    '            excel_Worksheet.Cells(iPos, xlsExport.Tax_Pct) = Ordhead.Tax_Pct
    '            excel_Worksheet.Cells(iPos, xlsExport.Tax_Cd_2) = Ordhead.Tax_Cd_2
    '            excel_Worksheet.Cells(iPos, xlsExport.Tax_Pct_2) = Ordhead.Tax_Pct_2
    '            excel_Worksheet.Cells(iPos, xlsExport.Tax_Cd_3) = Ordhead.Tax_Cd_3
    '            excel_Worksheet.Cells(iPos, xlsExport.Tax_Pct_3) = Ordhead.Tax_Pct_3
    '            excel_Worksheet.Cells(iPos, xlsExport.Form_No) = Ordhead.Form_No
    '            errPos = 86
    '            excel_Worksheet.Cells(iPos, xlsExport.Deter_Rate_By) = Ordhead.Deter_Rate_By
    '            excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_1) = Ordhead.User_Def_Fld_1
    '            errPos = 88
    '            excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_2) = Ordhead.User_Def_Fld_2
    '            errPos = 89
    '            excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_3) = Ordhead.User_Def_Fld_3
    '            excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_4) = Ordhead.User_Def_Fld_4
    '            excel_Worksheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_5) = Ordhead.User_Def_Fld_5
    '            excel_Worksheet.Cells(iPos, xlsExport.Tax_Fg) = Ordhead.Tax_Fg
    '            excel_Worksheet.Cells(iPos, xlsExport.Frt_Pay_Cd) = Ordhead.Frt_Pay_Cd
    '            excel_Worksheet.Cells(iPos, xlsExport.Selection_Cd) = Ordhead.Selection_Cd
    '            excel_Worksheet.Cells(iPos, xlsExport.Hold_Fg) = Ordhead.Hold_Fg

    '            excel_Worksheet.Cells(iPos, xlsExport.Item_No) = dt.Rows(0).Item("Item_No")
    '            errPos = 96
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Loc) = dt.Rows(0).Item("Loc")
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_Ordered) = Get_all_inventory_item_total_qty()
    '            errPos = 98
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = 1
    '            errPos = 99

    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = 0

    '            '    errPos = 100

    '            'Dim _oOrdline_US_Unit_Price As Double = Get_all_inventory_item_total_unit_fee() * 0.02 'justin add 20230221 , ->****20230509 need think per row???
    '            'excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Price) = _oOrdline_US_Unit_Price
    '            Dim _oOrdline_US_Total_Price As Double = Get_all_inventory_item_total_price() * 0.02 'justin add 20230221 , ->****20230509 need think per row???
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Price) = _oOrdline_US_Total_Price

    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Discount_Pct) = 0
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Weight) = 0

    '            Dim _44mark_dt As String = Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") & Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") & Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Request_Dt) = _44mark_dt



    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Promise_Dt) = _44mark_dt
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Req_Ship_Date) = _44mark_dt
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Guid) = Item_Guid

    '            ' The EOF_Marker to guid is set the line after the Next of the For Each Command
    '            excel_Worksheet.Cells(iPos, xlsExport.EOF_Marker) = "N"


    '            excel_Worksheet.Cells(iPos, xlsExport.Item_BkOrd_Fg) = dt.Rows(0).Item("BkOrd_Fg")

    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Qty_BkOrd_Fg) = "N" ' oOrdline.Qty_Bkord_Fg

    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Unit_Cost) = 0

    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Release_No) = "     1".ToString ' Forced at 1, inventory incremental 
    '            errPos = 116
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Desc_1) = dt.Rows(0).Item("Item_Desc_1") ' Forced at 1, inventory incremental 
    '            excel_Worksheet.Cells(iPos, xlsExport.Item_Desc_2) = dt.Rows(0).Item("Item_Desc_2") ' Forced at 1, inventory incremental 

    '            errPos = 118



    '            errPos = 120
    '            excel_Worksheet.Cells(iPos, xlsExport.TravelerRouteID) = "XX"
    '            excel_Worksheet.Cells(iPos, xlsExport.TravelerRoute) = "XX"
    '            excel_Worksheet.Cells(iPos, xlsExport.RouteCategory) = "XX"
    '            excel_Worksheet.Cells(iPos, xlsExport.Mfg_Item_No) = dt.Rows(0).Item("Item_No") & "-" & "XX" ' Changed June 26, 2017
    '            excel_Worksheet.Cells(iPos, xlsExport.ItemExtension) = "XX"

    '            errPos = 121
    '            'Added June 29, 2017 - T. Louzon
    '            excel_Worksheet.Cells(iPos, xlsExport.IsNoStock) = 0
    '            excel_Worksheet.Cells(iPos, xlsExport.IsNoArtwork) = 0

    '            errPos = 122
    '            'Added July 12, 2017 - T. Louzon
    '            excel_Worksheet.Cells(iPos, xlsExport.Prod_Cat) = dt.Rows(0).Item("Prod_Cat")


    '            errPos = 125

    '            'Added July 18, 2017 - T. Louzon
    '            excel_Worksheet.Cells(iPos, xlsExport.Num_Impr_1) = 0
    '            excel_Worksheet.Cells(iPos, xlsExport.Num_Impr_2) = 0
    '            excel_Worksheet.Cells(iPos, xlsExport.Num_Impr_3) = 0
    '            errPos = 127
    '        End If
    '    Catch er As Exception
    '        MsgBox("Error in COrder->Add_44MARKUPUS1." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try
    'End Sub


    Public Function CreateCAPoExcelFile(ByRef CaPo_Ordhead As cOrdHead) As String

        Dim errPos As Integer

        Dim xlApp As Excel.Application
        xlApp = New Excel.ApplicationClass

        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Try
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets(1)

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String

            Dim strLastGuid As String = ""
            Dim iPos As Integer = 0
            Dim iCount As Integer = 0


            For Each oOrdline As cOrdLine In OrdLines
                ' As we do not have now line validation, validate empty items here 
                ' and do not include them in the excel export file.
                ' iCount and iPos are different because of kits. We do not insert line for
                ' kit components but they are part of the order because of traveler and imprint.
                ' iCount is the number of lines in the order
                iCount += 1
                If Trim(oOrdline.Item_No) <> "" And oOrdline.Qty_Ordered <> 0 And Not (oOrdline.Kit_Component) Then
                    strSql = "SELECT ITEM_NO FROM  [200].dbo.IMITMIDX_SQL WITH (NOLOCK) WHERE ITEM_NO='" & SqlCompliantString(Trim(oOrdline.Item_No)) & "' "
                    dt = db.DataTable(strSql)
                    If dt.Rows.Count <> 0 Then
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'if is not kits
                        If Not If_kits_item_no(Trim(oOrdline.Item_No)) Then
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ' iPos is the Y row position to write to in the Excel file.
                            strLastGuid = oOrdline.Item_Guid
                            iPos += 1
                            xlWorkSheet.Cells(iPos, xlsExport.Ord_Type) = "O" ' Ordhead.Ord_Type
                            If Ordhead.Ord_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Ord_Dt) =
                                Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") &
                                Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") &
                                Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
                            xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No).Style.NumberFormat = "@"
                            xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No) = Ordhead.Oe_Po_No
                            errPos = 36
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            'xlWorkSheet.Cells(iPos, xlsExport.Cus_No) = Ordhead.Cus_No  'justin comment this line and add below on 20230809
                            xlWorkSheet.Cells(iPos, xlsExport.Cus_No) = CaPo_Ordhead.Cus_No
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Name) = Ordhead.Bill_To_Name
                            xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_1) = Ordhead.Bill_To_Addr_1
                            xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_2) = Ordhead.Bill_To_Addr_2
                            xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_3) = Ordhead.Bill_To_Addr_3
                            xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_4) = Ordhead.Bill_To_Addr_4
                            xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Country) = Ordhead.Bill_To_Country

                            If Ordhead.Cus_Alt_Adr_Cd <> "SAME" And Ordhead.Cus_Alt_Adr_Cd <> "" Then
                                xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
                            Else
                                xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = ""
                            End If

                            xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Name) = Ordhead.Ship_To_Name
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_1) = Ordhead.Ship_To_Addr_1
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_2) = Ordhead.Ship_To_Addr_2
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_3) = Ordhead.Ship_To_Addr_3
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_4) = Ordhead.Ship_To_Addr_4
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Country) = Ordhead.Ship_To_Country
                            If Ordhead.Shipping_Dt.Date.Equals(NoDate()) Then
                                xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) = "" ' "21111111"
                            Else
                                xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) =
                                Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") &
                                Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") &
                                Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0")
                            End If

                            xlWorkSheet.Cells(iPos, xlsExport.Ship_Via_Cd) = Ordhead.Ship_Via_Cd
                            errPos = 56
                            xlWorkSheet.Cells(iPos, xlsExport.Ar_Terms_Cd) = Ordhead.Ar_Terms_Cd
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_1) = Ordhead.Ship_Instruction_1
                            errPos = 58
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_2) = Ordhead.Ship_Instruction_2
                            xlWorkSheet.Cells(iPos, xlsExport.SlsPsn_No) = Ordhead.Slspsn_No
                            xlWorkSheet.Cells(iPos, xlsExport.Discount_Pct) = Ordhead.Discount_Pct
                            xlWorkSheet.Cells(iPos, xlsExport.Job_No) = Ordhead.Job_No

                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''justin comment this line and add below on 20230809 here should be US1 or GIT
                            'xlWorkSheet.Cells(iPos, xlsExport.Mfg_Loc) = Ordhead.Mfg_Loc
                            xlWorkSheet.Cells(iPos, xlsExport.Mfg_Loc) = CaPo_Ordhead.Mfg_Loc

                            xlWorkSheet.Cells(iPos, xlsExport.Profit_Center) = Ordhead.Profit_Center
                            xlWorkSheet.Cells(iPos, xlsExport.Filler_0001) = Ordhead.Ord_GUID
                            xlWorkSheet.Cells(iPos, xlsExport.Curr_Cd) = Ordhead.Curr_Cd
                            xlWorkSheet.Cells(iPos, xlsExport.Orig_Trx_Rt) = Ordhead.Orig_Trx_Rt
                            xlWorkSheet.Cells(iPos, xlsExport.Curr_Trx_Rt) = Ordhead.Curr_Trx_Rt
                            xlWorkSheet.Cells(iPos, xlsExport.Tax_Sched) = Ordhead.Tax_Sched
                            xlWorkSheet.Cells(iPos, xlsExport.Contact_1) = Ordhead.Contact_1
                            xlWorkSheet.Cells(iPos, xlsExport.Phone_Number) = Ordhead.Phone_Number
                            xlWorkSheet.Cells(iPos, xlsExport.Fax_Number) = Ordhead.Fax_Number
                            xlWorkSheet.Cells(iPos, xlsExport.Email_Address) = Ordhead.Email_Address
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_To_City) = Ordhead.Ship_To_City
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_To_State) = Ordhead.Ship_To_State
                            xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Zip) = Ordhead.Ship_To_Zip
                            xlWorkSheet.Cells(iPos, xlsExport.Bill_To_City) = Ordhead.Bill_To_City
                            xlWorkSheet.Cells(iPos, xlsExport.Bill_To_State) = Ordhead.Bill_To_State
                            xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Zip) = Ordhead.Bill_To_Zip
                            xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd) = Ordhead.Tax_Cd
                            errPos = 80
                            xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct) = Ordhead.Tax_Pct
                            xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_2) = Ordhead.Tax_Cd_2
                            xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_2) = Ordhead.Tax_Pct_2
                            xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_3) = Ordhead.Tax_Cd_3
                            xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_3) = Ordhead.Tax_Pct_3
                            xlWorkSheet.Cells(iPos, xlsExport.Form_No) = Ordhead.Form_No
                            errPos = 86
                            xlWorkSheet.Cells(iPos, xlsExport.Deter_Rate_By) = Ordhead.Deter_Rate_By
                            xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_1) = Ordhead.User_Def_Fld_1
                            errPos = 88
                            xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_2) = Ordhead.User_Def_Fld_2
                            errPos = 89
                            xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_3) = Ordhead.User_Def_Fld_3
                            xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_4) = Ordhead.User_Def_Fld_4
                            xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_5) = Ordhead.User_Def_Fld_5
                            xlWorkSheet.Cells(iPos, xlsExport.Tax_Fg) = Ordhead.Tax_Fg
                            xlWorkSheet.Cells(iPos, xlsExport.Frt_Pay_Cd) = Ordhead.Frt_Pay_Cd
                            xlWorkSheet.Cells(iPos, xlsExport.Selection_Cd) = Ordhead.Selection_Cd
                            xlWorkSheet.Cells(iPos, xlsExport.Hold_Fg) = Ordhead.Hold_Fg
                            xlWorkSheet.Cells(iPos, xlsExport.Item_No) = oOrdline.Item_No
                            errPos = 96

                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''justin comment this line and add below on 20230809 here should be US1 or GIT
                            'xlWorkSheet.Cells(iPos, xlsExport.Item_Loc) = oOrdline.Loc
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Loc) = CaPo_Ordhead.Mfg_Loc

                            xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_Ordered) = oOrdline.Qty_Ordered
                            errPos = 98
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = oOrdline.Qty_To_Ship
                            errPos = 99
                            If oOrdline.Qty_To_Ship >= oOrdline.Qty_Ordered Then
                                xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = 0
                            Else
                                xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = oOrdline.Qty_Ordered - oOrdline.Qty_To_Ship
                            End If
                            errPos = 100

                            Dim _oOrdline_US_Unit_Price As Double = 0
                            ''
                            Dim _oOrdline_US_avg_cost As Double = 0
                            Dim _oOrdline_US_nqg_price As Double = 0
                            '------------------------------------------------------------------------------------------------------
                            'justin add 20230808 for inventory transfers ticket:34919
                            If CaPo_Ordhead.Cus_No = "SPECTUS" Then
                                If (Get_commodity_cd_by_item_no(oOrdline.Item_No) <> "Y" And If_inventory_item_no(oOrdline.Item_No) = False) Then
                                    _oOrdline_US_Unit_Price = 0
                                Else
                                    If If_inventory_item_no(oOrdline.Item_No) = False Then '20230512
                                        _oOrdline_US_Unit_Price = oOrdline.Unit_Price
                                    Else
                                        _oOrdline_US_Unit_Price = Get_avg_cost_by_item_no_In200(oOrdline.Item_No, oOrdline.Unit_Price)
                                    End If
                                End If
                                _oOrdline_US_avg_cost = _oOrdline_US_Unit_Price
                            Else '''''''''''''''''''''''''''''''''''''''''''''''''''
                                If (Get_commodity_cd_by_item_no(oOrdline.Item_No) <> "Y" And If_inventory_item_no(oOrdline.Item_No) = False) Then
                                    _oOrdline_US_Unit_Price = 0
                                Else
                                    If If_inventory_item_no(oOrdline.Item_No) = False Then '20230512
                                        _oOrdline_US_Unit_Price = oOrdline.Unit_Price
                                    Else
                                        _oOrdline_US_Unit_Price = Get_nqg_price_by_item_no(oOrdline.Item_No, oOrdline.Unit_Price)
                                    End If

                                End If
                                _oOrdline_US_nqg_price = _oOrdline_US_Unit_Price
                            End If '''''''''''''''''''''''''''''''''''''''''''''''''''
                            '------------------------------------------------------------------------------------------------------
                            ''

                            xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Price) = _oOrdline_US_Unit_Price

                            xlWorkSheet.Cells(iPos, xlsExport.Item_Discount_Pct) = oOrdline.Discount_Pct
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Weight) = oOrdline.Unit_Weight

                            If oOrdline.Request_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Request_Dt) =
                                oOrdline.Request_Dt.Year.ToString.PadLeft(4, "0") &
                                oOrdline.Request_Dt.Month.ToString.PadLeft(2, "0") &
                                oOrdline.Request_Dt.Day.ToString.PadLeft(2, "0")
                            If oOrdline.Promise_Dt.Equals(Date.MinValue) Then
                                oOrdline.Promise_Dt = Ordhead.Ord_Dt_Shipped
                            End If

                            If oOrdline.Promise_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Promise_Dt) =
                                oOrdline.Promise_Dt.Year.ToString.PadLeft(4, "0") &
                                oOrdline.Promise_Dt.Month.ToString.PadLeft(2, "0") &
                                oOrdline.Promise_Dt.Day.ToString.PadLeft(2, "0")
                            If oOrdline.Req_Ship_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Req_Ship_Date) =
                                oOrdline.Req_Ship_Dt.Year.ToString.PadLeft(4, "0") &
                                oOrdline.Req_Ship_Dt.Month.ToString.PadLeft(2, "0") &
                                oOrdline.Req_Ship_Dt.Day.ToString.PadLeft(2, "0")
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Guid) = oOrdline.Item_Guid

                            ' The EOF_Marker to guid is set the line after the Next of the For Each Command
                            xlWorkSheet.Cells(iPos, xlsExport.EOF_Marker) = "N"

                            xlWorkSheet.Cells(iPos, xlsExport.User_Def_Fld_4) = oOrdline.User_Def_Fld_4
                            xlWorkSheet.Cells(iPos, xlsExport.Item_BkOrd_Fg) = oOrdline.Bkord_Fg
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Bin_Fg) = oOrdline.Bin_Fg
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Prc_Cd_Orig_Price) = oOrdline.Orig_Price
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd_Fg) = IIf(oOrdline.Qty_Prev_Bkord <> 0, "Y", "N") ' oOrdline.Qty_Bkord_Fg
                            'oOrdline()
                            ' Prevent division by zero
                            If Ordhead.Curr_Trx_Rt = 0 Then
                                xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = oOrdline.Unit_Cost  ' Forced at 1, inventory incremental 
                            Else
                                xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = Format(oOrdline.Unit_Cost / Ordhead.Curr_Trx_Rt, "##,##0.000000") ' Forced at 1, inventory incremental 
                            End If
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Release_No) = "     1".ToString ' Forced at 1, inventory incremental 
                            errPos = 116
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_1) = oOrdline.Item_Desc_1 ' Forced at 1, inventory incremental 
                            xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_2) = oOrdline.Item_Desc_2 ' Forced at 1, inventory incremental 
                            xlWorkSheet.Cells(iPos, 98) = oOrdline.Extra_2 ' Forced at 1, inventory incremental 
                            xlWorkSheet.Cells(iPos, 99) = oOrdline.Extra_6 ' Forced at 1, inventory incremental 
                            errPos = 118


                            '==============================================
                            'Added June 17, 2017 - T. Louzon

                            'GET THE Item Extension for the Route ID
                            Dim dt2 As DataTable
                            Dim db2 As New cDBA
                            Dim strSql2 As String
                            Dim sItemExtension As String
                            Dim sRouteCategory As String
                            strSql2 = " SELECT XREF.ItemExtension, T.RouteCategory "
                            strSql2 = strSql2 & " FROM Exact_traveler_Route T "
                            strSql2 = strSql2 & " INNER JOIN EXACT_Traveler_Route XREF  "    'HV_OEI_RouteCategory_XREF XREF "  'Changed July 18, 2017
                            strSql2 = strSql2 & " On T.RouteCategory = XREF.RouteCategory "
                            strSql2 = strSql2 & " WHERE XREF.routeID = " & oOrdline.m_RouteID

                            errPos = 119
                            dt2 = db2.DataTable(strSql2)
                            If dt2.Rows.Count <> 0 Then

                                If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                    sItemExtension = "XX"
                                Else
                                    sItemExtension = dt2.Rows(0).Item("ItemExtension")
                                End If

                                If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                    sRouteCategory = "XX"
                                Else
                                    sRouteCategory = dt2.Rows(0).Item("RouteCategory")
                                End If

                            Else
                                sItemExtension = "XX"
                                sRouteCategory = "XX"
                            End If

                            dt2.Dispose()

                            errPos = 120
                            xlWorkSheet.Cells(iPos, xlsExport.TravelerRouteID) = oOrdline.m_RouteID
                            xlWorkSheet.Cells(iPos, xlsExport.TravelerRoute) = oOrdline.m_Route
                            xlWorkSheet.Cells(iPos, xlsExport.RouteCategory) = sRouteCategory ' 
                            xlWorkSheet.Cells(iPos, xlsExport.Mfg_Item_No) = oOrdline.Item_No & "-" & sItemExtension ' Changed June 26, 2017
                            xlWorkSheet.Cells(iPos, xlsExport.ItemExtension) = sItemExtension '  

                            errPos = 121
                            'Added June 29, 2017 - T. Louzon
                            xlWorkSheet.Cells(iPos, xlsExport.IsNoStock) = oOrdline.IsNoStock
                            xlWorkSheet.Cells(iPos, xlsExport.IsNoArtwork) = oOrdline.IsNoArtwork

                            errPos = 122
                            'Added July 12, 2017 - T. Louzon
                            xlWorkSheet.Cells(iPos, xlsExport.Prod_Cat) = oOrdline.Prod_Cat
                            xlWorkSheet.Cells(iPos, xlsExport.End_Item_Cd) = oOrdline.End_Item_Cd

                            'GET THE EXTRA FIELDS TO EXPORT
                            '==============================================
                            'Added June 17, 2017 - T.Louzon
                            Dim dt3 As DataTable
                            Dim db3 As New cDBA
                            Dim strSql3 As String
                            'Dim sItemExtension As String
                            'Dim sRouteCategory As String
                            strSql3 = " SELECT   "
                            strSql3 = strSql3 & " item_no, "
                            strSql3 = strSql3 & " num_impr_1, "
                            strSql3 = strSql3 & " num_impr_2, "
                            strSql3 = strSql3 & " num_impr_3, "
                            strSql3 = strSql3 & " imprint = extra_1, "
                            strSql3 = strSql3 & " Imprint_Location = Extra_2, "
                            strSql3 = strSql3 & " imprint_Color = Extra_3, "
                            strSql3 = strSql3 & " NUM = Extra_4,  "       '-- Number of Imprints
                            strSql3 = strSql3 & " Packaging = Extra_5, "
                            strSql3 = strSql3 & " Refill = Extra_8, "
                            strSql3 = strSql3 & " Laser_Setup = Extra_9, "
                            strSql3 = strSql3 & " comments = Comment, "
                            strSql3 = strSql3 & " special_Comments = Comment2, "
                            strSql3 = strSql3 & " Industry = Industry, "
                            strSql3 = strSql3 & " Printer_Instructions, "
                            strSql3 = strSql3 & " Packer_Instructions "

                            strSql3 = strSql3 & " FROM OEI_EXTRA_FIELDS "
                            strSql3 = strSql3 & " WHERE ITEM_GUID = '" & oOrdline.m_Item_Guid & "'"

                            dt3 = db3.DataTable(strSql3)

                            errPos = 125

                            If dt3.Rows.Count <> 0 Then
                                errPos = 126

                                'Added July 18, 2017 - T. Louzon
                                xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_1) = dt3.Rows(0).Item("Num_Impr_1")
                                xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_2) = dt3.Rows(0).Item("Num_Impr_2")
                                xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_3) = dt3.Rows(0).Item("Num_Impr_3")
                                xlWorkSheet.Cells(iPos, xlsExport.Imprint) = dt3.Rows(0).Item("Imprint")    ''
                                xlWorkSheet.Cells(iPos, xlsExport.Imprint_Location) = dt3.Rows(0).Item("Imprint_Location")
                                xlWorkSheet.Cells(iPos, xlsExport.Imprint_Color) = dt3.Rows(0).Item("Imprint_Color")
                                xlWorkSheet.Cells(iPos, xlsExport.Comments) = dt3.Rows(0).Item("Comments")
                                xlWorkSheet.Cells(iPos, xlsExport.Special_Comments) = dt3.Rows(0).Item("Special_Comments")
                                xlWorkSheet.Cells(iPos, xlsExport.Printer_Instructions) = dt3.Rows(0).Item("Printer_Instructions")
                                xlWorkSheet.Cells(iPos, xlsExport.Packer_Instructions) = dt3.Rows(0).Item("Packer_Instructions")
                            Else
                                'Nothing
                            End If
                            errPos = 127
                            dt3.Dispose()
                            'iPos = Add_44markup_per_ordline(xlWorkSheet, oOrdline, iPos, _oOrdline_US_Unit_Price, oOrdline.Qty_Ordered, CaPo_Ordhead)
                            'if customer is SPECTCA _oOrdline_US_Unit_Price should be _oOrdline_US_avg_cost  above line 3503
                            'if other customer      _oOrdline_US_Unit_Priceshould be _oOrdline_US_nqg_price  above line 3503
                            '''''''''''''

                            If CaPo_Ordhead.Cus_No = "SPECTUS" Then
                                iPos = Add_44MARKUPINV_per_ordline(xlWorkSheet, oOrdline, iPos, _oOrdline_US_avg_cost, oOrdline.Qty_Ordered, CaPo_Ordhead)
                            Else
                                iPos = Add_44markup_per_ordline(xlWorkSheet, oOrdline, iPos, _oOrdline_US_Unit_Price, oOrdline.Qty_Ordered, CaPo_Ordhead)
                            End If
                            '==============================================
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            'if is not kits end
                        Else
                            'if it is kit
                            'get line_no
                            Dim kits_line_no As Int16 = oOrdline.Line_No
                            Dim kits_line_qty_order As Int16 = oOrdline.Qty_Ordered

                            Dim kits_line_qty_to_Ship = oOrdline.Qty_To_Ship
                            'loop each orderline
                            For Each kits_oOrdline As cOrdLine In OrdLines
                                'if line_seq_is above and not kits then
                                If kits_oOrdline.Line_No = kits_line_no And Not If_kits_item_no(Trim(kits_oOrdline.Item_No)) Then
                                    'Write excel
                                    ' iPos is the Y row position to write to in the Excel file.
                                    strLastGuid = oOrdline.Item_Guid
                                    iPos += 1
                                    xlWorkSheet.Cells(iPos, xlsExport.Ord_Type) = "O" ' Ordhead.Ord_Type
                                    If Ordhead.Ord_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Ord_Dt) =
                                        Ordhead.Ord_Dt.Year.ToString.PadLeft(4, "0") &
                                        Ordhead.Ord_Dt.Month.ToString.PadLeft(2, "0") &
                                        Ordhead.Ord_Dt.Day.ToString.PadLeft(2, "0")
                                    xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No).Style.NumberFormat = "@"
                                    xlWorkSheet.Cells(iPos, xlsExport.OE_PO_No) = Ordhead.Oe_Po_No
                                    errPos = 36
                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    'xlWorkSheet.Cells(iPos, xlsExport.Cus_No) = Ordhead.Cus_No  'justin comment this line and add below on 20230809
                                    xlWorkSheet.Cells(iPos, xlsExport.Cus_No) = CaPo_Ordhead.Cus_No
                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Name) = Ordhead.Bill_To_Name
                                    xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_1) = Ordhead.Bill_To_Addr_1
                                    xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_2) = Ordhead.Bill_To_Addr_2
                                    xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_3) = Ordhead.Bill_To_Addr_3
                                    xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Addr_4) = Ordhead.Bill_To_Addr_4
                                    xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Country) = Ordhead.Bill_To_Country

                                    If Ordhead.Cus_Alt_Adr_Cd <> "SAME" And Ordhead.Cus_Alt_Adr_Cd <> "" Then
                                        xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = Ordhead.Cus_Alt_Adr_Cd
                                    Else
                                        xlWorkSheet.Cells(iPos, xlsExport.Cus_Alt_Adr_Cd) = ""
                                    End If

                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Name) = Ordhead.Ship_To_Name
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_1) = Ordhead.Ship_To_Addr_1
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_2) = Ordhead.Ship_To_Addr_2
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_3) = Ordhead.Ship_To_Addr_3
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Addr_4) = Ordhead.Ship_To_Addr_4
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Country) = Ordhead.Ship_To_Country
                                    If Ordhead.Shipping_Dt.Date.Equals(NoDate()) Then
                                        xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) = "" ' "21111111"
                                    Else
                                        xlWorkSheet.Cells(iPos, xlsExport.Shipping_Dt) =
                                        Ordhead.Shipping_Dt.Year.ToString.PadLeft(4, "0") &
                                        Ordhead.Shipping_Dt.Month.ToString.PadLeft(2, "0") &
                                        Ordhead.Shipping_Dt.Day.ToString.PadLeft(2, "0")
                                    End If

                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_Via_Cd) = Ordhead.Ship_Via_Cd
                                    errPos = 56
                                    xlWorkSheet.Cells(iPos, xlsExport.Ar_Terms_Cd) = Ordhead.Ar_Terms_Cd
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_1) = Ordhead.Ship_Instruction_1
                                    errPos = 58
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_Instruction_2) = Ordhead.Ship_Instruction_2
                                    xlWorkSheet.Cells(iPos, xlsExport.SlsPsn_No) = Ordhead.Slspsn_No
                                    xlWorkSheet.Cells(iPos, xlsExport.Discount_Pct) = Ordhead.Discount_Pct
                                    xlWorkSheet.Cells(iPos, xlsExport.Job_No) = Ordhead.Job_No

                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    ''justin comment this line and add below on 20230809 here should be US1 or GIT
                                    'xlWorkSheet.Cells(iPos, xlsExport.Mfg_Loc) = Ordhead.Mfg_Loc
                                    xlWorkSheet.Cells(iPos, xlsExport.Mfg_Loc) = CaPo_Ordhead.Mfg_Loc



                                    xlWorkSheet.Cells(iPos, xlsExport.Profit_Center) = Ordhead.Profit_Center
                                    xlWorkSheet.Cells(iPos, xlsExport.Filler_0001) = Ordhead.Ord_GUID
                                    xlWorkSheet.Cells(iPos, xlsExport.Curr_Cd) = Ordhead.Curr_Cd
                                    xlWorkSheet.Cells(iPos, xlsExport.Orig_Trx_Rt) = Ordhead.Orig_Trx_Rt
                                    xlWorkSheet.Cells(iPos, xlsExport.Curr_Trx_Rt) = Ordhead.Curr_Trx_Rt
                                    xlWorkSheet.Cells(iPos, xlsExport.Tax_Sched) = Ordhead.Tax_Sched
                                    xlWorkSheet.Cells(iPos, xlsExport.Contact_1) = Ordhead.Contact_1
                                    xlWorkSheet.Cells(iPos, xlsExport.Phone_Number) = Ordhead.Phone_Number
                                    xlWorkSheet.Cells(iPos, xlsExport.Fax_Number) = Ordhead.Fax_Number
                                    xlWorkSheet.Cells(iPos, xlsExport.Email_Address) = Ordhead.Email_Address
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_To_City) = Ordhead.Ship_To_City
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_To_State) = Ordhead.Ship_To_State
                                    xlWorkSheet.Cells(iPos, xlsExport.Ship_To_Zip) = Ordhead.Ship_To_Zip
                                    xlWorkSheet.Cells(iPos, xlsExport.Bill_To_City) = Ordhead.Bill_To_City
                                    xlWorkSheet.Cells(iPos, xlsExport.Bill_To_State) = Ordhead.Bill_To_State
                                    xlWorkSheet.Cells(iPos, xlsExport.Bill_To_Zip) = Ordhead.Bill_To_Zip
                                    xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd) = Ordhead.Tax_Cd
                                    errPos = 80
                                    xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct) = Ordhead.Tax_Pct
                                    xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_2) = Ordhead.Tax_Cd_2
                                    xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_2) = Ordhead.Tax_Pct_2
                                    xlWorkSheet.Cells(iPos, xlsExport.Tax_Cd_3) = Ordhead.Tax_Cd_3
                                    xlWorkSheet.Cells(iPos, xlsExport.Tax_Pct_3) = Ordhead.Tax_Pct_3
                                    xlWorkSheet.Cells(iPos, xlsExport.Form_No) = Ordhead.Form_No
                                    errPos = 86
                                    xlWorkSheet.Cells(iPos, xlsExport.Deter_Rate_By) = Ordhead.Deter_Rate_By
                                    xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_1) = Ordhead.User_Def_Fld_1
                                    errPos = 88
                                    xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_2) = Ordhead.User_Def_Fld_2
                                    errPos = 89
                                    xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_3) = Ordhead.User_Def_Fld_3
                                    xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_4) = Ordhead.User_Def_Fld_4
                                    xlWorkSheet.Cells(iPos, xlsExport.Hdr_User_Def_Fld_5) = Ordhead.User_Def_Fld_5
                                    xlWorkSheet.Cells(iPos, xlsExport.Tax_Fg) = Ordhead.Tax_Fg
                                    xlWorkSheet.Cells(iPos, xlsExport.Frt_Pay_Cd) = Ordhead.Frt_Pay_Cd
                                    xlWorkSheet.Cells(iPos, xlsExport.Selection_Cd) = Ordhead.Selection_Cd
                                    xlWorkSheet.Cells(iPos, xlsExport.Hold_Fg) = Ordhead.Hold_Fg
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_No) = kits_oOrdline.Item_No
                                    errPos = 96

                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    ''justin comment this line and add below on 20230809 here should be US1 or GIT
                                    'xlWorkSheet.Cells(iPos, xlsExport.Item_Loc) = kits_oOrdline.Loc
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Loc) = CaPo_Ordhead.Mfg_Loc


                                    'xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_Ordered) = kits_oOrdline.Qty_Ordered
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_Ordered) = kits_line_qty_order
                                    errPos = 98
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_To_Ship) = kits_line_qty_to_Ship
                                    errPos = 99
                                    If kits_oOrdline.Qty_To_Ship >= kits_oOrdline.Qty_Ordered Then
                                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = 0
                                    Else
                                        xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd) = kits_oOrdline.Qty_Ordered - kits_oOrdline.Qty_To_Ship
                                    End If
                                    errPos = 100

                                    Dim _oOrdline_US_Unit_Price As Double = 0

                                    Dim _oOrdline_US_avg_cost As Double = 0
                                    Dim _oOrdline_US_nqg_price As Double = 0
                                    '------------------------------------------------------------------------------------------------------
                                    'justin add 20230808 for inventory transfers ticket:34919
                                    If CaPo_Ordhead.Cus_No = "SPECTUS" Then
                                        _oOrdline_US_Unit_Price = Get_avg_cost_by_item_no_In200(kits_oOrdline.Item_No, kits_oOrdline.Unit_Price)
                                        _oOrdline_US_avg_cost = _oOrdline_US_Unit_Price
                                    Else
                                        _oOrdline_US_Unit_Price = Get_nqg_price_by_item_no(kits_oOrdline.Item_No, kits_oOrdline.Unit_Price)
                                        _oOrdline_US_nqg_price = _oOrdline_US_Unit_Price
                                    End If

                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Price) = _oOrdline_US_Unit_Price

                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Discount_Pct) = kits_oOrdline.Discount_Pct
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Weight) = kits_oOrdline.Unit_Weight

                                    If kits_oOrdline.Request_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Request_Dt) =
                                        kits_oOrdline.Request_Dt.Year.ToString.PadLeft(4, "0") &
                                        kits_oOrdline.Request_Dt.Month.ToString.PadLeft(2, "0") &
                                        kits_oOrdline.Request_Dt.Day.ToString.PadLeft(2, "0")
                                    If kits_oOrdline.Promise_Dt.Equals(Date.MinValue) Then
                                        kits_oOrdline.Promise_Dt = Ordhead.Ord_Dt_Shipped
                                    End If

                                    If kits_oOrdline.Promise_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Promise_Dt) =
                                        kits_oOrdline.Promise_Dt.Year.ToString.PadLeft(4, "0") &
                                        kits_oOrdline.Promise_Dt.Month.ToString.PadLeft(2, "0") &
                                        kits_oOrdline.Promise_Dt.Day.ToString.PadLeft(2, "0")
                                    If kits_oOrdline.Req_Ship_Dt.Year <> 1 Then xlWorkSheet.Cells(iPos, xlsExport.Item_Req_Ship_Date) =
                                        kits_oOrdline.Req_Ship_Dt.Year.ToString.PadLeft(4, "0") &
                                        kits_oOrdline.Req_Ship_Dt.Month.ToString.PadLeft(2, "0") &
                                        kits_oOrdline.Req_Ship_Dt.Day.ToString.PadLeft(2, "0")
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Guid) = kits_oOrdline.Item_Guid

                                    ' The EOF_Marker to guid is set the line after the Next of the For Each Command
                                    xlWorkSheet.Cells(iPos, xlsExport.EOF_Marker) = "N"

                                    xlWorkSheet.Cells(iPos, xlsExport.User_Def_Fld_4) = kits_oOrdline.User_Def_Fld_4
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_BkOrd_Fg) = kits_oOrdline.Bkord_Fg
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Bin_Fg) = kits_oOrdline.Bin_Fg
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Prc_Cd_Orig_Price) = kits_oOrdline.Orig_Price
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Qty_BkOrd_Fg) = IIf(kits_oOrdline.Qty_Prev_Bkord <> 0, "Y", "N") ' oOrdline.Qty_Bkord_Fg
                                    'oOrdline()
                                    ' Prevent division by zero
                                    If Ordhead.Curr_Trx_Rt = 0 Then
                                        xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = kits_oOrdline.Unit_Cost  ' Forced at 1, inventory incremental 
                                    Else
                                        xlWorkSheet.Cells(iPos, xlsExport.Item_Unit_Cost) = Format(kits_oOrdline.Unit_Cost / Ordhead.Curr_Trx_Rt, "##,##0.000000") ' Forced at 1, inventory incremental 
                                    End If
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Release_No) = "     1".ToString ' Forced at 1, inventory incremental 
                                    errPos = 116
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_1) = kits_oOrdline.Item_Desc_1 ' Forced at 1, inventory incremental 
                                    xlWorkSheet.Cells(iPos, xlsExport.Item_Desc_2) = kits_oOrdline.Item_Desc_2 ' Forced at 1, inventory incremental 
                                    xlWorkSheet.Cells(iPos, 98) = kits_oOrdline.Extra_2 ' Forced at 1, inventory incremental 
                                    xlWorkSheet.Cells(iPos, 99) = kits_oOrdline.Extra_6 ' Forced at 1, inventory incremental 
                                    errPos = 118


                                    '==============================================
                                    'Added June 17, 2017 - T. Louzon

                                    'GET THE Item Extension for the Route ID
                                    Dim dt2 As DataTable
                                    Dim db2 As New cDBA
                                    Dim strSql2 As String
                                    Dim sItemExtension As String
                                    Dim sRouteCategory As String
                                    strSql2 = " SELECT XREF.ItemExtension, T.RouteCategory "
                                    strSql2 = strSql2 & " FROM Exact_traveler_Route T "
                                    strSql2 = strSql2 & " INNER JOIN EXACT_Traveler_Route XREF  "    'HV_OEI_RouteCategory_XREF XREF "  'Changed July 18, 2017
                                    strSql2 = strSql2 & " On T.RouteCategory = XREF.RouteCategory "
                                    strSql2 = strSql2 & " WHERE XREF.routeID = " & oOrdline.m_RouteID

                                    errPos = 119
                                    dt2 = db2.DataTable(strSql2)
                                    If dt2.Rows.Count <> 0 Then

                                        If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                            sItemExtension = "XX"
                                        Else
                                            sItemExtension = dt2.Rows(0).Item("ItemExtension")
                                        End If

                                        If IsDBNull(dt2.Rows(0).Item("ItemExtension")) Then
                                            sRouteCategory = "XX"
                                        Else
                                            sRouteCategory = dt2.Rows(0).Item("RouteCategory")
                                        End If

                                    Else
                                        sItemExtension = "XX"
                                        sRouteCategory = "XX"
                                    End If

                                    dt2.Dispose()

                                    errPos = 120
                                    xlWorkSheet.Cells(iPos, xlsExport.TravelerRouteID) = kits_oOrdline.m_RouteID
                                    xlWorkSheet.Cells(iPos, xlsExport.TravelerRoute) = kits_oOrdline.m_Route
                                    xlWorkSheet.Cells(iPos, xlsExport.RouteCategory) = sRouteCategory ' 
                                    xlWorkSheet.Cells(iPos, xlsExport.Mfg_Item_No) = kits_oOrdline.Item_No & "-" & sItemExtension ' Changed June 26, 2017
                                    xlWorkSheet.Cells(iPos, xlsExport.ItemExtension) = sItemExtension '  

                                    errPos = 121
                                    'Added June 29, 2017 - T. Louzon
                                    xlWorkSheet.Cells(iPos, xlsExport.IsNoStock) = kits_oOrdline.IsNoStock
                                    xlWorkSheet.Cells(iPos, xlsExport.IsNoArtwork) = kits_oOrdline.IsNoArtwork

                                    errPos = 122
                                    'Added July 12, 2017 - T. Louzon
                                    xlWorkSheet.Cells(iPos, xlsExport.Prod_Cat) = kits_oOrdline.Prod_Cat
                                    xlWorkSheet.Cells(iPos, xlsExport.End_Item_Cd) = kits_oOrdline.End_Item_Cd

                                    'GET THE EXTRA FIELDS TO EXPORT
                                    '==============================================
                                    'Added June 17, 2017 - T.Louzon
                                    Dim dt3 As DataTable
                                    Dim db3 As New cDBA
                                    Dim strSql3 As String
                                    'Dim sItemExtension As String
                                    'Dim sRouteCategory As String
                                    strSql3 = " SELECT   "
                                    strSql3 = strSql3 & " item_no, "
                                    strSql3 = strSql3 & " num_impr_1, "
                                    strSql3 = strSql3 & " num_impr_2, "
                                    strSql3 = strSql3 & " num_impr_3, "
                                    strSql3 = strSql3 & " imprint = extra_1, "
                                    strSql3 = strSql3 & " Imprint_Location = Extra_2, "
                                    strSql3 = strSql3 & " imprint_Color = Extra_3, "
                                    strSql3 = strSql3 & " NUM = Extra_4,  "       '-- Number of Imprints
                                    strSql3 = strSql3 & " Packaging = Extra_5, "
                                    strSql3 = strSql3 & " Refill = Extra_8, "
                                    strSql3 = strSql3 & " Laser_Setup = Extra_9, "
                                    strSql3 = strSql3 & " comments = Comment, "
                                    strSql3 = strSql3 & " special_Comments = Comment2, "
                                    strSql3 = strSql3 & " Industry = Industry, "
                                    strSql3 = strSql3 & " Printer_Instructions, "
                                    strSql3 = strSql3 & " Packer_Instructions "

                                    strSql3 = strSql3 & " FROM OEI_EXTRA_FIELDS "
                                    strSql3 = strSql3 & " WHERE ITEM_GUID = '" & kits_oOrdline.m_Item_Guid & "'"

                                    dt3 = db3.DataTable(strSql3)

                                    errPos = 125

                                    If dt3.Rows.Count <> 0 Then
                                        errPos = 126

                                        'Added July 18, 2017 - T. Louzon
                                        xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_1) = dt3.Rows(0).Item("Num_Impr_1")
                                        xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_2) = dt3.Rows(0).Item("Num_Impr_2")
                                        xlWorkSheet.Cells(iPos, xlsExport.Num_Impr_3) = dt3.Rows(0).Item("Num_Impr_3")
                                        xlWorkSheet.Cells(iPos, xlsExport.Imprint) = dt3.Rows(0).Item("Imprint")    ''
                                        xlWorkSheet.Cells(iPos, xlsExport.Imprint_Location) = dt3.Rows(0).Item("Imprint_Location")
                                        xlWorkSheet.Cells(iPos, xlsExport.Imprint_Color) = dt3.Rows(0).Item("Imprint_Color")
                                        xlWorkSheet.Cells(iPos, xlsExport.Comments) = dt3.Rows(0).Item("Comments")
                                        xlWorkSheet.Cells(iPos, xlsExport.Special_Comments) = dt3.Rows(0).Item("Special_Comments")
                                        xlWorkSheet.Cells(iPos, xlsExport.Printer_Instructions) = dt3.Rows(0).Item("Printer_Instructions")
                                        xlWorkSheet.Cells(iPos, xlsExport.Packer_Instructions) = dt3.Rows(0).Item("Packer_Instructions")
                                    Else
                                        'Nothing
                                    End If
                                    errPos = 127
                                    dt3.Dispose()

                                    'iPos = Add_44markup_per_ordline(xlWorkSheet, kits_oOrdline, iPos, _oOrdline_US_Unit_Price, kits_line_qty_order, CaPo_Ordhead)
                                    'if customer is SPECTCA _oOrdline_US_Unit_Price should be _oOrdline_US_avg_cost  above line 3503
                                    'if other customer      _oOrdline_US_Unit_Priceshould be _oOrdline_US_nqg_price  above line 3503
                                    '''''''''''''

                                    If CaPo_Ordhead.Cus_No = "SPECTUS" Then
                                        iPos = Add_44MARKUPINV_per_ordline(xlWorkSheet, oOrdline, iPos, _oOrdline_US_avg_cost, oOrdline.Qty_Ordered, CaPo_Ordhead)
                                    Else
                                        iPos = Add_44markup_per_ordline(xlWorkSheet, kits_oOrdline, iPos, _oOrdline_US_Unit_Price, kits_line_qty_order, CaPo_Ordhead)
                                    End If
                                    '
                                End If
                            Next
                            '
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    End If
                Else
                    iCount -= 1
                End If

                errPos = 150
            Next

            ' This must be used so when the last item is a kit component, it sets correctly the last line to guid
            errPos = 151
            '** IF THE PROGRAM CRASHES HERE, THEN EVERY LINE HAS A QTY OF 0 **
            xlWorkSheet.Cells(iPos, xlsExport.EOF_Marker) = Ordhead.Ord_GUID


            errPos = 152

            'justin 20230221 add 44MARKUPUS1 item row in the end of row of excel
            'Call Add_44MARKUPUS1(xlWorkSheet, iPos)


            Dim m_USpath As String = HV_US_PO_Import_Files_Folder()
            '''''
            'm_USpath = "C:\ExactTemp\3\"
            'xlWorkBook.SaveAs(Filename:="" & m_USpath & Ordhead.Ord_GUID & "_3_PO_US.xls")
            'Call UnsetExportTS()
            '''''
            xlWorkBook.SaveAs(Filename:="" & m_USpath & Ordhead.Ord_GUID & ".xls")



            errPos = 153
            xlWorkBook.Close()
            xlWorkBook = Nothing
            errPos = 154

            xlApp.Quit()
            xlApp = Nothing

            errPos = 155

            'MsgBox("File created.")
        Catch er As Exception
            Call UnsetExportTS()
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " " & errPos.ToString)
            xlWorkBook.Close(False)
            xlApp.Quit()
        End Try

    End Function


    Public Sub SendOrderAck()

        Dim mbrResult As New MsgBoxResult
        Dim strSql As String

        Dim dt As DataTable
        Dim db As New cDBA

        Ordhead.SendOrderAck = 0

        Try
            Dim p_ord_Ack As Boolean = False

            p_ord_Ack = Return_Order_Ack_Exception()

            '++ID 09.06.2018
            'True is exception, when client doesn't want receive the order ack confirmation
            ' Ordhead.SendOrderAck = 0 variable is already has defauslt value above
            If p_ord_Ack = False Then

                'SELECT * FROM OEI_Order_Contacts WITH (Nolock) WHERE Ord_Guid = 'E3A09FC8-03C3-4576-BAF9-D17541' AND contactType = 4 

                strSql =
            "SELECT *  " &
            "FROM   OEI_Order_Contacts WITH (Nolock) " &
            "WHERE  Ord_Guid = '" & Ordhead.Ord_GUID & "' AND " &
            "       ContactType = 4 "

                dt = db.DataTable(strSql)

                If dt.Rows.Count <> 0 Then

                    'MsgBoxStyle.YesNoCancel()

                    mbrResult = MsgBox("Do you wish to send the Order Ack right away?", MsgBoxStyle.YesNoCancel)

                    If mbrResult = MsgBoxResult.Yes Then

                        Ordhead.SendOrderAck = 1

                    ElseIf mbrResult = MsgBoxResult.No Then

                        Ordhead.SendOrderAck = 0

                        MsgBox("The Order Ack will not be sent.", MsgBoxStyle.OkOnly, "Order Ack not sent")

                    ElseIf mbrResult = MsgBoxResult.Cancel Then

                        CancelExport = True

                    End If

                Else

                    MsgBox("The Order Ack will not be sent." & vbCrLf & "You must first choose a contact for the order.", MsgBoxStyle.OkOnly, "Order Ack not sent")

                End If
            End If

            Call Ordhead.Save()

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SetExportTS()

        Dim db As New cDBA

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola, False, False)

            Dim strSql As String

            strSql = _
            "UPDATE OEI_ORDHDR SET ExportTS = GetDate() " & _
            "WHERE  Ord_Guid = '" & Ordhead.Ord_GUID & "' "

            db.Execute(strSql)

            Ordhead.ExportTS = Date.Now.ToString

        Catch oe_er As OEException
            If oe_er.ShowMessage Then MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Public Sub UnsetExportTS()

        Dim db As New cDBA

        Try

            'If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola, False, False)

            Dim strSql As String

            strSql = _
            "UPDATE OEI_ORDHDR SET ExportTS = NULL " & _
            "WHERE  Ord_Guid = '" & Ordhead.Ord_GUID & "' "

            db.Execute(strSql)

        Catch oe_er As OEException
            If oe_er.ShowMessage Then MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrdhead." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        Ordhead.ExportTS = ""

    End Sub

    'Public Sub DeleteOrderLine(ByRef poOrdline As cOrdLine)

    '    If poOrdline.Item_Guid.Equals(DBNull.Value) Then Exit Sub

    '    Call DeleteOrderLineKit(poOrdline)

    '    If poOrdline.Kit.IsAKit Then
    '        Call DeleteOrderLineKit(poOrdline)
    '    Else
    '        poOrdline.Delete()
    '        OrdLines.Remove(poOrdline.Item_Guid)
    '        poOrdline = Nothing
    '    End If

    'End Sub

    Public Sub UnsetImportedEDIOrder()

        Dim oOrderEDI As New cOrderEDI(Ordhead.Ord_GUID)

        If oOrderEDI.OrdEDI_ID <> 0 Then
            oOrderEDI.OEI_Status = "R"
            oOrderEDI.Ord_Guid = ""
            oOrderEDI.Save()
        End If

    End Sub

    Public Sub DeleteEmptyOrderIfNew()

        Try
            ' Delete the order if no line item and order is the last order made by that user.
            ' We can do reverse deleting here, as we can create 10 orders, and 1 by one from
            ' the last one, delete every item of the order and delete the order then.

            If OrdLines.Count <> 0 Then Exit Sub

            If Ordhead.OEI_Ord_No Is Nothing Then

                g_User.PreviousOrderNumber()

            ElseIf Trim(Ordhead.OEI_Ord_No) = Trim(g_User.OrderNumberGenerated) Then

                DeletePendingLines(Ordhead.Ord_GUID)
                Ordhead.Delete()
                g_User.PreviousOrderNumber()

            End If

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub DeletePendingLines(ByVal pOrd_Guid As String)

        Try
            If pOrd_Guid.Trim.Equals(String.Empty) Then Exit Sub

            Dim db As New cDBA()
            Dim strSqlWhere As String = "WHERE Ord_Guid = '" & pOrd_Guid & "' "

            db.Execute("DELETE FROM OEI_Extra_Functions " & strSqlWhere)
            db.Execute("DELETE FROM OEI_Extra_Fields " & strSqlWhere)
            db.Execute("DELETE FROM OEI_Traveler_Details " & strSqlWhere)
            db.Execute("DELETE FROM OEI_Traveler_Header " & strSqlWhere)
            db.Execute("DELETE FROM OEI_Order_Contacts " & strSqlWhere)
            db.Execute("DELETE FROM OEI_OrdBld " & strSqlWhere)
            db.Execute("DELETE FROM OEI_EMAIL " & strSqlWhere)
            'db.Execute("DELETE FROM EXACT_TRAVELER_DOCUMENT " & strSqlWhere)

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub DeleteOrderLine(ByRef poOrdline As cOrdLine)

        Dim strSql As String
        Dim dt As DataTable
        Dim db As New cDBA
        Dim strSqlWhere As String

        Try
            strSqlWhere = "WHERE Item_Guid = '" & poOrdline.Item_Guid & "' "

            If poOrdline.Item_Guid.Equals(DBNull.Value) Then Exit Sub

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If poOrdline.Item_Guid.Equals(DBNull.Value) Then Exit Sub

            strSql = _
            "       SELECT  Comp_Item_Guid " & _
            "       FROM    OEI_OrdBld " & _
            "       WHERE   Kit_Item_Guid = '" & poOrdline.Item_Guid & "' "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                For Each dr As DataRow In dt.Rows
                    If OrdLines.Contains(dr.Item("Comp_Item_Guid").ToString) Then
                        OrdLines.Remove(dr.Item("Comp_Item_Guid").ToString)
                    End If
                    'OrdLines.Remove(dr.Item("Comp_Item_Guid"))
                Next dr

                strSqlWhere = _
                "WHERE Item_Guid in " & _
                "   (   " & _
                "       SELECT  Comp_Item_Guid " & _
                "       FROM    OEI_OrdBld " & _
                "       WHERE   Kit_Item_Guid = '" & poOrdline.Item_Guid & "' " & _
                "   ) OR Item_Guid = '" & poOrdline.Item_Guid & "' "
            Else
                If OrdLines.Contains(poOrdline.Item_Guid) Then OrdLines.Remove(poOrdline.Item_Guid)
            End If

            db.Execute("DELETE FROM OEI_Extra_Functions " & strSqlWhere)
            db.Execute("DELETE FROM OEI_Extra_Fields " & strSqlWhere)
            db.Execute("DELETE FROM OEI_Traveler_Details " & strSqlWhere)
            db.Execute("DELETE FROM OEI_Traveler_Header " & strSqlWhere)
            db.Execute("DELETE FROM OEI_OrdLin " & strSqlWhere)

            ' DELETE FROM OEI_ORDBLD - ALL CASES, DO LAST !
            db.Execute("INSERT INTO OEI_ORDBLD_DELETED SELECT * FROM OEI_OrdBld WHERE Kit_Item_Guid = '" & poOrdline.Item_Guid & "' ")

            ' DELETE FROM OEI_ORDBLD - ALL CASES, DO LAST !
            db.Execute("DELETE FROM OEI_OrdBld WHERE Kit_Item_Guid = '" & poOrdline.Item_Guid & "' ")

            'poOrdline = Nothing

            'strSql = _
            '"DELETE FROM OEI_Extra_Functions " & strSqlWhere

            'db.DBCommand.CommandText = strSql
            'If db.DBCommand.Connection.State <> ConnectionState.Open Then db.DBCommand.Connection.Open()
            'db.DBCommand.ExecuteNonQuery()
            'If db.DBCommand.Connection.State <> ConnectionState.Closed Then db.DBCommand.Connection.Close()

            'strSql = _
            '"DELETE FROM OEI_Extra_Fields " & strSqlWhere

            'db.DBCommand.CommandText = strSql
            'If db.DBCommand.Connection.State <> ConnectionState.Open Then db.DBCommand.Connection.Open()
            'db.DBCommand.ExecuteNonQuery()
            'If db.DBCommand.Connection.State <> ConnectionState.Closed Then db.DBCommand.Connection.Close()

            'strSql = _
            '"DELETE FROM OEI_Traveler_Details " & strSqlWhere

            'db.DBCommand.CommandText = strSql
            'If db.DBCommand.Connection.State <> ConnectionState.Open Then db.DBCommand.Connection.Open()
            'db.DBCommand.ExecuteNonQuery()
            'If db.DBCommand.Connection.State <> ConnectionState.Closed Then db.DBCommand.Connection.Close()

            'strSql = _
            '"DELETE FROM OEI_Traveler_Header " & strSqlWhere

            'db.DBCommand.CommandText = strSql
            'If db.DBCommand.Connection.State <> ConnectionState.Open Then db.DBCommand.Connection.Open()
            'db.DBCommand.ExecuteNonQuery()
            'If db.DBCommand.Connection.State <> ConnectionState.Closed Then db.DBCommand.Connection.Close()

            'strSql = _
            '"DELETE FROM OEI_OrdLin " & strSqlWhere

            'db.DBCommand.CommandText = strSql
            'If db.DBCommand.Connection.State <> ConnectionState.Open Then db.DBCommand.Connection.Open()
            'db.DBCommand.ExecuteNonQuery()
            'If db.DBCommand.Connection.State <> ConnectionState.Closed Then db.DBCommand.Connection.Close()

            '' DELETE FROM OEI_ORDBLD - ALL CASES, DO LAST !
            'strSql = "DELETE FROM OEI_OrdBld WHERE Item_Guid = '" & poOrdline.Item_Guid & "' "

            'db.DBCommand.CommandText = strSql
            'If db.DBCommand.Connection.State <> ConnectionState.Open Then db.DBCommand.Connection.Open()
            'db.DBCommand.ExecuteNonQuery()
            'If db.DBCommand.Connection.State <> ConnectionState.Closed Then db.DBCommand.Connection.Close()

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub AddExternalLine(ByRef poOrdline As cOrdLine)

        Dim oNewOrdline As New cOrdLine(poOrdline.Ord_Guid)

        Try

            With oNewOrdline

                .Reset()
                .Loading = True

                .Ord_Guid = poOrdline.Ord_Guid
                .Item_Guid = poOrdline.Item_Guid
                .Source = poOrdline.Source

                ' here we must fix the bug for the kit insert for line no.
                'If it is as kit, it must not be workline + 1.
                m_oOrder.WorkLine = m_oOrder.GetMaxLine_No(m_oOrder.Ordhead.Ord_GUID)
                If Not poOrdline.Kit.IsAKit Then m_oOrder.WorkLine += 1
                oNewOrdline.Line_No = m_oOrder.WorkLine

                .SaveToDB = True

                .Create_Imprint = True
                .ImprintLine = New cImprint(.Item_Guid)

                .Create_Traveler = True
                .Traveler = New CTraveler(.Item_Guid)

                .Create_Kit = True

                .m_Loc = poOrdline.Loc ' For kits, must enter Loc first else will not find item.

                .Item_No = poOrdline.Item_No
                .Loc = poOrdline.Loc
                .Qty_Ordered = poOrdline.Qty_Ordered
                .Qty_To_Ship = poOrdline.Qty_To_Ship
                .Qty_Inventory = poOrdline.Qty_Inventory
                .Qty_Allocated = poOrdline.Qty_Allocated
                .Qty_On_Hand = poOrdline.Qty_On_Hand
                .Qty_Prev_To_Ship = poOrdline.Qty_Prev_To_Ship
                .Qty_Prev_Bkord = poOrdline.Qty_Prev_Bkord
                .Qty_On_Hand = poOrdline.Qty_On_Hand
                .Unit_Price = poOrdline.Unit_Price
                .Unit_Price_BeforeSpecial = poOrdline.Unit_Price       'Added June 22, 2017 T. Louzon *******
                .Extra_9 = poOrdline.Extra_9
                .Extra_10 = poOrdline.Extra_10
                .Extra_11 = poOrdline.Extra_11
                .Discount_Pct = poOrdline.Discount_Pct
                .Calc_Price = poOrdline.Calc_Price
                .Request_Dt = poOrdline.Request_Dt
                .Promise_Dt = poOrdline.Promise_Dt
                .Req_Ship_Dt = poOrdline.Req_Ship_Dt
                .Item_Desc_1 = poOrdline.Item_Desc_1
                .Item_Desc_2 = poOrdline.Item_Desc_2
                .Uom = poOrdline.Uom
                .Tax_Sched = poOrdline.Tax_Sched
                .Revision_No = poOrdline.Revision_No
                .Bkord_Fg = poOrdline.Bkord_Fg
                .Recalc_Sw = poOrdline.Recalc_Sw
                .Route = poOrdline.Route
                .RouteID = poOrdline.RouteID
                .Prod_Cat = poOrdline.Prod_Cat
                .Stocked_Fg = poOrdline.Stocked_Fg
                .End_Item_Cd = poOrdline.End_Item_Cd
                .ProductProof = poOrdline.ProductProof
                .AutoCompleteReship = poOrdline.AutoCompleteReship
                .Item_Cd = poOrdline.Item_Cd
                .Mix_Group = poOrdline.Mix_Group

                ' add imprint columns

                .ImprintLine.Imprint = poOrdline.ImprintLine.Imprint
                .ImprintLine.Location = poOrdline.ImprintLine.Location
                .ImprintLine.Color = poOrdline.ImprintLine.Color
                .ImprintLine.Num_Impr_1 = poOrdline.ImprintLine.Num_Impr_1
                .ImprintLine.Num_Impr_2 = poOrdline.ImprintLine.Num_Impr_2
                .ImprintLine.Num_Impr_3 = poOrdline.ImprintLine.Num_Impr_3
                If poOrdline.Kit.IsAKit Then
                    Dim oItems As DataTable
                    Dim db As New cDBA()
                    Dim sSql As String = "SELECT * FROM OEI_ORDBLD WHERE ORD_GUID = '" & .Ord_Guid & "' AND KIT_ITEM_GUID = '" & .Item_Guid & "' "
                    oItems = db.DataTable(sSql)
                    If oItems.Rows.Count <> 0 Then
                        For Each oKRow As DataRow In oItems.Rows
                            sSql = "UPDATE OEI_EXTRA_FIELDS SET PACKAGING = '" & poOrdline.ImprintLine.Packaging & "' WHERE ORD_GUID = '" & .Ord_Guid & "' AND ITEM_GUID = '" & oKRow.Item("COMP_ITEM_GUID").ToString.Trim & "' "
                            'db.Connexion.Open()
                            'db.Execute(sSql)
                            'db.Connexion.Close()
                        Next
                    End If
                Else
                    .ImprintLine.Packaging = poOrdline.ImprintLine.Packaging
                End If
                .ImprintLine.Refill = poOrdline.ImprintLine.Refill
                .ImprintLine.Laser_Setup = poOrdline.ImprintLine.Laser_Setup
                .ImprintLine.Comments = poOrdline.ImprintLine.Comments
                .ImprintLine.Special_Comments = poOrdline.ImprintLine.Special_Comments
                .ImprintLine.Industry = poOrdline.ImprintLine.Industry
                .ImprintLine.Printer_Comment = poOrdline.ImprintLine.Printer_Comment
                .ImprintLine.Printer_Instructions = poOrdline.ImprintLine.Printer_Instructions
                .ImprintLine.Packer_Comment = poOrdline.ImprintLine.Packer_Comment


                .ImprintLine.Packer_Instructions = poOrdline.ImprintLine.Packer_Instructions

                .ImprintLine.End_User = poOrdline.ImprintLine.End_User

                '++ID 06.09.2020 scribl project
                .ImprintLine.Lamination_Scribl = poOrdline.ImprintLine.Lamination_Scribl
                '++ID 12.09.2021 Foil Color
                .ImprintLine.Foil_Color = poOrdline.ImprintLine.Foil_Color

                .ImprintLine.Thread_Color_Scribl = poOrdline.ImprintLine.Thread_Color_Scribl
                .ImprintLine.Rounded_corners_Scribl = poOrdline.ImprintLine.Rounded_corners_Scribl
                '------------------------------------------------------------------------------------
                '+ID 08.06.2020----------------------------------------------------------------------
                .ImprintLine.Tip_In_PageTxt = poOrdline.ImprintLine.Tip_In_PageTxt
                '-------------------------------------------------------------------------------------
                .Loading = False

            End With

            If poOrdline.FieldsRequired() Then

                If Not (OrdLines.Contains(oNewOrdline.Item_Guid)) Then
                    If oNewOrdline.SaveToDB Then oNewOrdline.Save()
                    OrdLines.Add(oNewOrdline, oNewOrdline.Item_Guid)
                Else
                    If oNewOrdline.SaveToDB Then oNewOrdline.Save()
                End If

            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub AddOrderLine(ByRef poOrdline As cOrdLine)

        Try

            ' Be careful with that function as kit items and kit components don't need the same fields
            If poOrdline.FieldsRequired() Then

                If Not (OrdLines.Contains(poOrdline.Item_Guid)) Then

                    If poOrdline.SaveToDB Then poOrdline.Save()

                    OrdLines.Add(poOrdline, poOrdline.Item_Guid)

                    If poOrdline.End_Item_Cd = "K" Then ' poOrdline.Kit.IsAKit Then

                        For Each iOrdline As cOrdLine In g_oKitCollection
                            If Not (OrdLines.Contains(iOrdline.Item_Guid)) Then
                                OrdLines.Add(iOrdline, iOrdline.Item_Guid)
                            End If
                        Next

                    End If

                Else
                    If poOrdline.SaveToDB Then poOrdline.Save()
                End If

            End If

            'If poOrdline.End_Item_Cd = "K" Then ' poOrdline.Kit.IsAKit Then

            '    For Each iOrdline As cOrdLine In g_oKitCollection
            '        If Not (OrdLines.Contains(iOrdline.Item_Guid)) Then
            '            OrdLines.Add(iOrdline, iOrdline.Item_Guid)
            '        End If
            '    Next

            'End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function GetMaxLine_No(ByVal pstrOrder_Guid As String) As Integer

        GetMaxLine_No = 0

        Try
            Dim strSql As String = _
            "SELECT ISNULL(MAX(Line_No),0) AS MaxLine_No " & _
            "FROM   OEI_OrdLin WITH (Nolock) " & _
            "WHERE  Ord_Guid = '" & pstrOrder_Guid & "' "

            Dim dt As DataTable
            Dim db As New cDBA
            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                'If dt.Rows(0).Item("MaxLine_No") = 0 Then
                'GetMaxLine_No = 1
                'Else
                GetMaxLine_No = dt.Rows(0).Item("MaxLine_No")
                'End If
            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function Total_Order_Item_Count() As Integer

        Total_Order_Item_Count = 0

        Try
            Dim oOrdline As cOrdLine ' m_oOrder)

            'For lPos As Integer = 1 To m_oOrder.OrdLines.Count
            For Each oOrdline In m_oOrder.OrdLines
                If Not (oOrdline.Kit_Component) Then Total_Order_Item_Count = Total_Order_Item_Count + 1
            Next oOrdline

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function Total_Curr_Trx_Rt() As Double

        Total_Curr_Trx_Rt = 0

        Try
            For Each oOrdline As cOrdLine In m_oOrder.OrdLines
                Total_Curr_Trx_Rt = m_oOrder.Ordhead.Curr_Trx_Rt
            Next oOrdline

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function Total_Order_Amt() As Double

        Total_Order_Amt = 0

        Try
            For Each oOrdline As cOrdLine In m_oOrder.OrdLines
                If Not (oOrdline.Kit_Component) Then Total_Order_Amt = Total_Order_Amt + Math.Round(((oOrdline.Qty_Ordered * oOrdline.Unit_Price) * (100 - oOrdline.Discount_Pct) / 100), 2)
            Next oOrdline

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function Total_Order_Qty() As Double

        Total_Order_Qty = 0

        Try
            For Each oOrdline As cOrdLine In m_oOrder.OrdLines
                If Not (oOrdline.Kit_Component) Then Total_Order_Qty = Total_Order_Qty + oOrdline.Qty_Ordered
            Next oOrdline

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function Total_Ship_Amt() As Double

        Total_Ship_Amt = 0

        Try
            For Each oOrdline As cOrdLine In m_oOrder.OrdLines
                If Not (oOrdline.Kit_Component) Then Total_Ship_Amt = Total_Ship_Amt + oOrdline.Calc_Price
            Next oOrdline

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function Total_Ship_Qty() As Double

        Total_Ship_Qty = 0

        Try
            For Each oOrdline As cOrdLine In m_oOrder.OrdLines
                If Not (oOrdline.Kit_Component) Then Total_Ship_Qty = Total_Ship_Qty + oOrdline.Qty_To_Ship
            Next oOrdline

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function Get_Product_category_Macola(ByVal in_item_no As String) As String 'justin add 20230301
        Try
            Get_Product_category_Macola = ""
            Dim item_no As String = LTrim(RTrim(in_item_no))
            If in_item_no = "" Then
                Return Get_Product_category_Macola
            End If

            Dim sql As String = ""
            sql =
                " SELECT  top 1 " &
                "         RTrim([200].dbo.imitmidx_sql.prod_cat) As cat, " &
                "         isnull(UPPER(rtrim([200].dbo.imitmidx_sql.item_note_2)),'') AS sub, " &
                "         [200].dbo.imcatfil_sql.prod_cat_desc as name_c " &
                "          FROM [200].dbo.imitmidx_sql " &
                "          Left Join [200].dbo.imcatfil_sql on rtrim([200].dbo.imitmidx_sql.prod_cat) = rtrim([200].dbo.imcatfil_sql.prod_cat) " &
                "          WHERE  [200].dbo.imitmidx_sql.activity_cd = 'A' and substring([200].dbo.imitmidx_sql.item_no,1,2) <>'AC' " &
                "          And substring([200].dbo.imitmidx_sql.item_no,1,2) <>'MK' and UPPER(rtrim([200].dbo.imitmidx_sql.item_note_2)) not like '%DEFECTIV%'  " &
                "          And ltrim(rtrim([200].dbo.imitmidx_sql.item_no)) = '" & item_no & "' "
            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(sql)
            If dt.Rows.Count > 0 Then
                Get_Product_category_Macola = dt.Rows(0).Item("cat")
            End If
        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Function

    Public Function Get_Product_sub_category_Macola(ByVal in_item_no As String) As String  'justin add 20230301
        Try
            Get_Product_sub_category_Macola = ""
            Dim item_no As String = LTrim(RTrim(in_item_no))
            If in_item_no = "" Then
                Return Get_Product_sub_category_Macola
            End If

            Dim sql As String = ""
            sql =
              " SELECT  top 1 " &
                "         RTrim([200].dbo.imitmidx_sql.prod_cat) As cat, " &
                "         isnull(UPPER(rtrim([200].dbo.imitmidx_sql.item_note_2)),'') AS sub, " &
                "         [200].dbo.imcatfil_sql.prod_cat_desc as name_c " &
                "          FROM [200].dbo.imitmidx_sql " &
                "          Left Join [200].dbo.imcatfil_sql on rtrim([200].dbo.imitmidx_sql.prod_cat) = rtrim([200].dbo.imcatfil_sql.prod_cat) " &
                "          WHERE  [200].dbo.imitmidx_sql.activity_cd = 'A' and substring([200].dbo.imitmidx_sql.item_no,1,2) <>'AC' " &
                "          And substring([200].dbo.imitmidx_sql.item_no,1,2) <>'MK' and UPPER(rtrim([200].dbo.imitmidx_sql.item_note_2)) not like '%DEFECTIV%'  " &
                "          And ltrim(rtrim([200].dbo.imitmidx_sql.item_no)) = '" & item_no & "' "
            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(sql)
            If dt.Rows.Count > 0 Then
                Get_Product_sub_category_Macola = dt.Rows(0).Item("sub")
            End If
        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Function

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Function ChangeCboRouteList(ByVal strGroup As String, ByVal pstrValue As String, ByVal pstr_item_no As String) As DataTable ' , ByVal pstrProd_Cat As String)

        ChangeCboRouteList = New DataTable

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Route must match category from line item.
            Dim SQL As String
            Dim groupCondition As String

            Dim Us_route_Condition As String  'justin add 20230301

            '==================================================
            'Added If statement July 14, 2017 - T. Louzon
            If gbchkLimitRoutes = False Then

                ' Extra conditions to sort correct routes
                Select Case strGroup

                    Case "All"
                        groupCondition = ""

                    Case "Misc"
                        groupCondition =
                        "   AND ((Rte.RouteDescription   like '%marketing%') " &
                        "   OR (Rte.RouteDescription not like '%flat%' " &
                        "   AND Rte.RouteDescription not like '%laser/silk%' " &
                        "   AND Rte.RouteDescription not like '%pad%' " &
                        "   AND Rte.RouteDescription not like '%silk%' " &
                        "   AND Rte.RouteDescription not like '%factory direct%' " &
                        "   AND Rte.RouteDescription not like '%overseas direct%' " &
                        "   AND Rte.RouteDescription not like '%silk/flat%')) "

                    Case "OS"
                        groupCondition =
                        "   AND (Rte.RouteDescription like '%factory direct%' " &
                        "   OR  Rte.RouteDescription  like '%overseas direct%') "

                    Case "Flat"
                        groupCondition =
                        "   AND Rte.RouteDescription     like '%flat%' " &
                        "   AND Rte.RouteDescription not like '%silk/flat%' "

                    Case "Silk"
                        groupCondition =
                        "   AND Rte.RouteDescription     like '%silk%' " &
                        "   AND Rte.RouteDescription not like '%laser/silk%' " &
                        "   AND Rte.RouteDescription not like '%silk/pad%' " &
                        "   AND Rte.RouteDescription not like '%silk/flat%' "

                    Case "Pad"
                        groupCondition =
                        "   AND Rte.RouteDescription     like '%pad%' " &
                        "   AND Rte.RouteDescription not like '%silk/pad%' "

                    Case Else
                        groupCondition =
                        "   AND Rte.RouteDescription like '%" & strGroup & "%' "

                End Select

                ' SQL Query to sort routes by group
                SQL =
                " SELECT     DISTINCT Rte.RouteDescription AS Route " &
                " FROM       exact_Traveler_Route Rte " &
                " INNER JOIN EXACT_TRAVELER_Xref_ProductCategoryRoute Xref " &
                "           ON Rte.RouteID = Xref.RouteID " &
                " WHERE      Rte.active <> 0 " & groupCondition &
                "           AND xref.Prod_cat = '" & g_oOrdline.Prod_Cat & "'" &
                " UNION " &
                " SELECT     '" & SqlCompliantString(pstrValue) & "' AS Route " &
                " UNION " &
                " SELECT     ' ' AS Route " &
                " ORDER BY   RouteDescription "

                'justin add 20230227 for us route
                If LTrim(RTrim(Ordhead.Mfg_Loc)) = "US1" Then
                    Us_route_Condition = ""
                    If LTrim(RTrim(pstr_item_no)).Substring(0, 3) = "╘═ " Then pstr_item_no = LTrim(RTrim(pstr_item_no)).Substring(3, LTrim(RTrim(pstr_item_no)).Length - 3)
                    'justin add above for filter component's ╘═  before "╘═ 00ST4237BLK"   20230601
                    Select Case Get_Product_category_Macola(LTrim(RTrim(pstr_item_no)))

                        Case "BAG"
                            Us_route_Condition = " and cat in ('Ashbury')  "
                        Case "JRN"
                            Us_route_Condition = " and cat in ('Journal')  "
                        Case "ORA"
                            Us_route_Condition = " and cat in ('ORA')  "
                        Case "MET"
                            If Get_Product_sub_category_Macola(LTrim(RTrim(pstr_item_no))) = "STYLUS" Then
                                Us_route_Condition = "  and cat in ('PEN Metal')  "
                            End If
                        Case "PAI"
                            If Get_Product_sub_category_Macola(LTrim(RTrim(pstr_item_no))) = "STYLUS" Then
                                Us_route_Condition = "  and cat in ('PEN Plastic')  "
                            End If
                        Case "PKG"
                            Us_route_Condition = " and cat in ('PKG')  "
                        Case Else
                            Us_route_Condition = " "
                    End Select
                    '
                    SQL = " SELECT     DISTINCT Rte.RouteDescription AS Route " &
                    " From US_Exact_Traveler_Route_Accepted Rte " &
                    " where (1=1) " & Us_route_Condition &
                    " UNION " &
                    " SELECT     '" & SqlCompliantString(pstrValue) & "' AS Route " &
                    " UNION " &
                    " SELECT     ' ' AS Route " &
                    " ORDER BY   RouteDescription "

                    'justin 20231018 add 
                    SQL = " SELECT     DISTINCT Rte.RouteDescription AS Route " &
                    " From US_Exact_Traveler_Route_Accepted_table Rte " &
                    " where (1=1) " & Us_route_Condition &
                    " UNION " &
                    " SELECT     '" & SqlCompliantString(pstrValue) & "' AS Route " &
                    " UNION " &
                    " SELECT     ' ' AS Route " &
                    " ORDER BY   RouteDescription "

                    '++ID 05262023 added exception in case if was not added exception yet. 
                    ' will show as in Canada
                    If Us_route_Condition = " " Then
                        ' SQL Query to sort routes by group
                        SQL =
                " SELECT     DISTINCT Rte.RouteDescription AS Route " &
                " FROM       exact_Traveler_Route Rte " &
                " INNER JOIN EXACT_TRAVELER_Xref_ProductCategoryRoute Xref " &
                "           ON Rte.RouteID = Xref.RouteID " &
                " WHERE      Rte.active <> 0 " & groupCondition &
                "           AND xref.Prod_cat = '" & g_oOrdline.Prod_Cat & "'" &
                " UNION " &
                " SELECT     '" & SqlCompliantString(pstrValue) & "' AS Route " &
                " UNION " &
                " SELECT     ' ' AS Route " &
                " ORDER BY   RouteDescription "
                    End If






                    'SQL =
                    '"SELECT     DISTINCT Rte.RouteDescription AS Route " &
                    '"FROM       US_Exact_Traveler_Route_Accepted Rte " &
                    '"INNER JOIN EXACT_TRAVELER_Xref_ProductCategoryRoute Xref " &
                    '"           ON Rte.RouteID = Xref.RouteID " &
                    '"WHERE      Rte.active <> 0 " & groupCondition &
                    '"           AND xref.Prod_cat = '" & g_oOrdline.Prod_Cat & "'" &
                    '" " & Us_route_Condition &
                    '"UNION " &
                    '"SELECT     '" & SqlCompliantString(pstrValue) & "' AS Route " &
                    '"UNION " &
                    '"SELECT     ' ' AS Route " &
                    '"ORDER BY   RouteDescription "
                End If
            Else
                '==================================================
                'ADDED July 14, 2017 - T. Louzon
                ' SQL QUERY TO ONLY DISPLAY ROUTES WITH MATCHING ITEM NUMBERS
                'Only used when chkLimitRoutes is checked.
                SQL =
                " SELECT     DISTINCT Rte.RouteDescription AS Route " &
                " FROM       exact_Traveler_Route Rte " &
                " INNER JOIN IMITMIDX_SQL IM  " &
                "           ON '" & Trim(g_oOrdline.Item_No) & "-' + RTE.ItemExtension = IM.item_no  " &
                " WHERE      Rte.active <> 0 " &
                " UNION " &
                " SELECT     ' ' AS Route " &
                " ORDER BY   RouteDescription "

            End If




            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(SQL)

            ChangeCboRouteList = dt


        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Function ChangeCboRouteList(ByVal strGroup As String, ByVal pstrProd_Cat As String, ByVal pstrValue As String, ByVal pstr_item_no As String) As DataTable ' , ByVal pstrProd_Cat As String)

        ChangeCboRouteList = New DataTable

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Route must match category from line item.
            Dim SQL As String
            Dim groupCondition As String


            Dim Us_route_Condition As String = "  "


            ' Dont attempt to load route if no step is loaded
            '    If gridLineItems.Columns(0).Text = "" Or gridLineItems.Columns(0).Text = Null Then
            '        Exit Sub
            '    End If

            ' Extra conditions to sort correct routes
            Select Case strGroup

                Case "All"
                    groupCondition = ""

                Case "Misc"
                    groupCondition =
                    "   AND ((Rte.RouteDescription   like '%marketing%') " &
                    "   OR (Rte.RouteDescription not like '%flat%' " &
                    "   AND Rte.RouteDescription not like '%laser/silk%' " &
                    "   AND Rte.RouteDescription not like '%pad%' " &
                    "   AND Rte.RouteDescription not like '%silk%' " &
                    "   AND Rte.RouteDescription not like '%factory direct%' " &
                    "   AND Rte.RouteDescription not like '%overseas direct%' " &
                    "   AND Rte.RouteDescription not like '%silk/flat%')) "

                Case "OS"
                    groupCondition =
                    "   AND (Rte.RouteDescription like '%factory direct%' " &
                    "   OR  Rte.RouteDescription  like '%overseas direct%') "

                Case "Flat"
                    groupCondition =
                    "   AND Rte.RouteDescription     like '%flat%' " &
                    "   AND Rte.RouteDescription not like '%silk/flat%' "

                Case "Silk"
                    groupCondition =
                    "   AND Rte.RouteDescription     like '%silk%' " &
                    "   AND Rte.RouteDescription not like '%laser/silk%' " &
                    "   AND Rte.RouteDescription not like '%silk/pad%' " &
                    "   AND Rte.RouteDescription not like '%silk/flat%' "

                Case "Pad"
                    groupCondition =
                    "   AND Rte.RouteDescription     like '%pad%' " &
                    "   AND Rte.RouteDescription not like '%silk/pad%' "

                Case Else
                    groupCondition =
                    "   AND Rte.RouteDescription like '%" & strGroup & "%' "

            End Select

            ' SQL Query to sort routes by group
            SQL =
            "SELECT     DISTINCT Rte.RouteDescription AS Route " &
            "FROM       exact_Traveler_Route Rte " &
            "INNER JOIN EXACT_TRAVELER_Xref_ProductCategoryRoute Xref " &
            "           ON Rte.RouteID = Xref.RouteID " &
            "WHERE      Rte.active <> 0 " & groupCondition &
            "           AND xref.Prod_cat = '" & pstrProd_Cat & "'" &
            "UNION " &
            "SELECT     '" & SqlCompliantString(pstrValue) & "' AS Route " &
            "ORDER BY   RouteDescription "

            If LTrim(RTrim(Ordhead.Mfg_Loc)) = "US1" Then 'justin add 20230227 for us route
                'justin add 20230227 for us route
                Us_route_Condition = ""
                Select Case Get_Product_category_Macola(LTrim(RTrim(pstr_item_no)))

                    Case "BAG"
                        Us_route_Condition = " and cat in ('Ashbury')  "
                    Case "JRN"
                        Us_route_Condition = " and cat in ('Journal')  "
                    Case "ORA"
                        Us_route_Condition = " and cat in ('ORA')  "
                    Case "MET"
                        If Get_Product_sub_category_Macola(LTrim(RTrim(pstr_item_no))) = "STYLUS" Then
                            Us_route_Condition = "  and cat in ('PEN Metal')  "
                        End If
                    Case "PAI"
                        If Get_Product_sub_category_Macola(LTrim(RTrim(pstr_item_no))) = "STYLUS" Then
                            Us_route_Condition = "  and cat in ('PEN Plastic')  "
                        End If
                    Case Else
                        Us_route_Condition = " "
                End Select
                '
                SQL = " SELECT     DISTINCT Rte.RouteDescription AS Route " &
                    " From US_Exact_Traveler_Route_Accepted Rte " &
                    " where (1=1) " & Us_route_Condition &
                    "UNION " &
                    "SELECT     '" & SqlCompliantString(pstrValue) & "' AS Route " &
                    "ORDER BY   RouteDescription "

                'justin 20231018 add
                SQL = " SELECT     DISTINCT Rte.RouteDescription AS Route " &
                    " From US_Exact_Traveler_Route_Accepted_table Rte " &
                    " where (1=1) " & Us_route_Condition &
                    "UNION " &
                    "SELECT     '" & SqlCompliantString(pstrValue) & "' AS Route " &
                    "ORDER BY   RouteDescription "

                'SQL =
                '"SELECT     DISTINCT Rte.RouteDescription AS Route " &
                '"FROM       US_Exact_Traveler_Route_Accepted Rte " &
                '"INNER JOIN EXACT_TRAVELER_Xref_ProductCategoryRoute Xref " &
                '"           ON Rte.RouteID = Xref.RouteID " &
                '"WHERE      Rte.active <> 0 " & groupCondition &
                '"           AND xref.Prod_cat = '" & pstrProd_Cat & "'" &
                '" " & Us_route_Condition &
                '" UNION " &
                '"SELECT     '" & SqlCompliantString(pstrValue) & "' AS Route " &
                '"ORDER BY   RouteDescription "
            End If
            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(SQL)

            ChangeCboRouteList = dt

            'If dt.Rows.Count <> 0 Then
            '    For Each oRow As DataRow In dt.Rows
            '        RouteCollection.Add(oRow.Item("RouteDescription").ToString, CStr(Trim(oRow.Item("RouteID"))))
            '        RouteIDCollection.Add(CInt(oRow.Item("RouteID")), oRow.Item("RouteDescription").ToString)
            '    Next oRow
            'End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Function GetCompleteCboRouteList() As DataTable ' , ByVal pstrProd_Cat As String)

        GetCompleteCboRouteList = New DataTable

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Route must match category from line item.
            Dim SQL As String

            ' SQL Query to sort routes by group
            SQL = _
            "SELECT     DISTINCT RouteDescription AS Route " & _
            "FROM       Exact_Traveler_Route Rte " & _
            "WHERE      Active = 1 " & _
            "ORDER BY   RouteDescription "

            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(SQL)

            GetCompleteCboRouteList = dt

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function GetDflt_Form() As Integer

        GetDflt_Form = 0

        Dim dt As DataTable
        Dim db As New cDBA
        Dim strSql As String = "SELECT ISNULL(Dflt_Form, 0) AS Dflt_Form FROM oectlfil_sql WITH (Nolock) "

        Try

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                GetDflt_Form = dt.Rows(0).Item("Dflt_Form")
            End If
            dt.Dispose()

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function GetOE_Ctl_OE_Exchange_Rate_Flag() As String

        GetOE_Ctl_OE_Exchange_Rate_Flag = ""

        Dim dt As DataTable
        Dim db As New cDBA
        Dim strSql As String = "SELECT ISNULL(OE_Ctl_OE_Exchange_Rate_Flag, 0) AS OE_Ctl_OE_Exchange_Rate_Flag FROM oectlfil_sql WITH (Nolock) "

        Try

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                GetOE_Ctl_OE_Exchange_Rate_Flag = dt.Rows(0).Item("OE_Ctl_OE_Exchange_Rate_Flag")
            End If
            dt.Dispose()

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Function GetCompleteCboIndustryList() As DataTable ' , ByVal pstrProd_Cat As String)

        GetCompleteCboIndustryList = New DataTable

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Route must match category from line item.
            Dim strSQL As String

            strSQL =
            "SELECT     '' AS Industry " &
            "UNION " &
            "SELECT     Industry_Type AS Industry " &
            "FROM       Industry_Classification WITH (Nolock) " &
            "ORDER BY   Industry ASC "

            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(strSQL)

            GetCompleteCboIndustryList = dt

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function Get_Scribl_Fields(ByVal _enum_cat As String) As DataTable ' , ByVal pstrProd_Cat As String)

        Get_Scribl_Fields = New DataTable

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Route must match category from line item.
            Dim strSQL As String

            strSQL =
            "SELECT     '-1' AS ID,'' as Enum_Value " &
            "UNION " &
            "SELECT     ID, Enum_Value " &
            "FROM    mdb_cfg_enum    WITH (Nolock) where enum_cat = '" & _enum_cat & "' " &
            "ORDER BY   Enum_value ASC "

            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(strSQL)

            Get_Scribl_Fields = dt

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function GetEQPTypes() As DataTable ' , ByVal pstrProd_Cat As String)

        GetEQPTypes = New DataTable

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Route must match category from line item.
            Dim strSQL As String

            strSQL = _
            "SELECT     'Standard' AS EQP_Type, 1 as SortOrder " & _
            "UNION " & _
            "SELECT     'EQP' AS EQP_Type, 2 as SortOrder " & _
            "UNION " & _
            "SELECT     'EQP Plus' AS EQP_Type, 3 as SortOrder " & _
            "UNION " & _
            "SELECT     'EQP Minus Pct' AS EQP_Type, 4 as SortOrder  " & _
            "ORDER BY   SortOrder "

            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(strSQL)

            GetEQPTypes = dt

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Sub Create_Preview_Item_Charges()

        Dim db As New cDBA
        Dim dt As DataTable
        Dim strSql As String

        'Dim blnContainsKits As Boolean = False
        Dim dblUnitPrice As Double = 0
        Dim intKitOrdered As Integer
        Dim oItem As cChargePreviewElement

        colDecMetList = New Collection
        colItemCharges = New Collection
        'colKitItems = New Collection
        colLogos = New Collection
        colItemRC = New Collection
        colColorCount = New System.Collections.Generic.Dictionary(Of String, Integer)
        colColorList = New System.Collections.Generic.Dictionary(Of String, String)
        'colColorCount = New System.Collections.Generic.Dictionary(Of String, Integer)
        'colColorList = New System.Collections.Generic.Dictionary(Of String, String)
        colColorEnum = New Collection
        colItemExtraByDecMetID = New Collection

        Try

            Call Get_Dec_Met_Configuration(colDecMetList)

            cptPricingType = GetPricingType()

            Call Delete_Preview_Item_Charges()

            strSql = "exec dbo.OEI_ITEM_CHARGES_PRECALC '" & m_oOrder.Ordhead.Ord_GUID & "' "
            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                For Each drRow As DataRow In dt.Rows

                    If drRow.Item("SqlPart") = 1 Then

                        oDecMet = New cMDBCfgDecMet(drRow.Item("Dec_Met_ID"))
                        Call Preview_Charges_Calc_Add_Kit_Logos(drRow)

                    End If

                Next

                For Each drRow As DataRow In dt.Rows

                    intKitOrdered = 1

                    If drRow.Item("SqlPart") = 2 Then

                        Dim DefDecMet As Integer = drRow.Item("Dec_Met_ID")

                        If drRow.Item("Logo_Qty") > intKitOrdered Or drRow.Item("Default_Opt") = 0 Then ' Dec_Met_ID
                            'WE MUST ADD ADDITIONAL CHARGE
                            Call Preview_Charges_Calc_Add_Logos(drRow)
                        End If

                    End If

                Next

                For Each drRow As DataRow In dt.Rows

                    intKitOrdered = 1

                    If drRow.Item("SqlPart") = 2 Then

                        Call Preview_Charges_Calc_Add_Logos(drRow)
                        Call Preview_Charges_Calc_Add_Loc(drRow)
                        Call Preview_Charges_Calc_Add_Colors(drRow)

                    End If

                Next

                For Each drRow As DataRow In dt.Rows

                    If drRow.Item("SqlPart") = 3 Then

                        'Dim DefDecMet As Integer = drRow.Item("Dec_Met_ID")
                        'If colKitItems.Contains(Trim(drRow.Item("Kit_Item_Cd"))) Then
                        '    DefDecMet = colKitItems(Trim(drRow.Item("Kit_Item_Cd")))
                        'End If

                        'If DefDecMet <> dtRow.Item("Dec_Met_ID") Then
                        '    If colItemCharges.Contains(1) Then
                        '    End If
                        'End If

                        Call Preview_Charges_Calc_Add_Logos(drRow)
                        Call Preview_Charges_Calc_Add_Loc(drRow)
                        Call Preview_Charges_Calc_Add_Colors(drRow)

                    End If

                Next

            End If

            Call Preview_Charges_Calc_Add_Color_Change_Charge()

            Call Preview_Charges_Always_Charges("PAPER_PROOF")
            Call Preview_Charges_Always_Charges("EXACT_QTY")
            'Call Preview_Charges_Always_Charges("PRODUCT_PROOF")
            Call Preview_Charges_ProductProof()

            Dim i As Integer

            For i = 1 To colItemCharges.Count
                oItem = colItemCharges(i)
                oItem.Save()
            Next

            'If Not colLogos.Contains(oItem.Item_No & "<" & oItem.Unit_Price.ToString & ">") Then
            '    colLogos.Add(oItem.Item_No & "<" & oItem.Unit_Price.ToString & ">", oItem.Item_No & "<" & oItem.Unit_Price.ToString & ">")
            'End If

            'If Not colLogos.Contains(oItem.Item_No & "<" & oDecMet.Dec_Met_ID.ToString & ">") Then
            '    colLogos.Add(oItem.Item_No & "<" & oDecMet.Dec_Met_ID.ToString & ">", oItem.Item_No & "<" & oDecMet.Dec_Met_ID.ToString & ">")
            'End If

            'If Not colKitItems.Contains(oItem.Item_No) Then
            '    colLogos.Add(dtRow.Item("Dec_Met_ID"), oItem.Item_No)
            'End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Preview_Charges_Calc_Add_Color_Change_Charge()

        Dim db As New cDBA
        Dim dt As DataTable
        Dim strSql As String
        'Dim strItemCharge As String
        Dim iCount As Integer = 0
        Dim oItem As cChargePreviewElement
        Dim dblUnitPrice As Double
        Dim strItem_No As String

        Try

            strSql = _
            "SELECT Item_No, DBO.OEI_Item_Price_20140101('" & Ordhead.Curr_Cd & "', '" & Ordhead.Cus_No & "', Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Item_Net " & _
            "FROM MDB_CFG_CHARGE WITH (Nolock) " & _
            "WHERE CHARGE_CD = 'COLOR_CHANGE' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                strItem_No = Trim(dt.Rows(0).Item("Item_No").ToString.ToUpper)

                'If cptPricingType = ChargesPricingType.Net Then
                dblUnitPrice = dt.Rows(0).Item("Item_Net")
                'Else
                '    dblUnitPrice = dt.Rows(0).Item("Item_Prc")
                'End If

                For Each strKey As String In colColorCount.Keys

                    If colColorCount(strKey) > 1 Then

                        If (colItemCharges.Contains(strItem_No & "<" & Trim(dblUnitPrice.ToString) & ">")) Then
                            oItem = colItemCharges(strItem_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                            If oItem.Description.Contains(strKey) Then
                                'oItem.Description &= ", " & strMainColor
                            Else
                                '                                oItem.Description &= ", " & Trim(strKey) & " (" & colColorList(strKey) & ")"
                                oItem.Description &= ", " & Trim(strKey) & " (" & Mid(colColorList(strKey), InStr(colColorList(strKey), ",") + 1) & ")"
                            End If
                            oItem.Qty_Ordered += (colColorCount(strKey) - 1) ' intKitOrdered
                        Else
                            oItem = New cChargePreviewElement
                            oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                            oItem.Item_No = Trim(strItem_No) ' Trim(dtRow.Item("Item_No"))
                            oItem.Unit_Price = dblUnitPrice
                            oItem.Qty_Ordered = (colColorCount(strKey) - 1) ' intKitOrdered
                            'oItem.Description = "COLOR CHANGE FOR " & Trim(strKey) & " (" & colColorList(strKey) & ")"
                            oItem.Description = "COLOR CHANGE FOR " & Trim(strKey) & " (" & Mid(colColorList(strKey), InStr(colColorList(strKey), ",") + 1) & ")"
                            colItemCharges.Add(oItem, strItem_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                        End If

                    End If

                Next strKey

                '    ' Once it's all done, now we run add color charges.
                '    If colColors.Count > 0 And Not (colColors.Contains(Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">")) Then
                '        'Else
                '    End If
            End If

            'If pdtRow.Item("Logo_Qty") <= 1 Then Exit Sub
            'If pdtRow.Item("Logo_Count") <= 1 Then Exit Sub

            'strItemCharge = Trim(oDec_Met.Setup_Item_No)

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Preview_Charges_Calc_Add_Colors(ByRef pdtRow As DataRow)

        Dim db As New cDBA
        Dim dt As DataTable
        Dim strSql As String
        Dim oDec_Met As cMDBCfgDecMet
        'Dim strItemCharge As String
        Dim iCount As Integer = 0
        Dim oItem As cChargePreviewElement
        Dim dblUnitPrice As Double
        Dim dblRCPrice As Double

        Try
            oDec_Met = colDecMetList(pdtRow.Item("Dec_Met_ID"))

            'If cptPricingType = ChargesPricingType.Net Then
            dblUnitPrice = pdtRow.Item("Add_Col_Setup_Net")
            dblRCPrice = pdtRow.Item("Add_Col_RC_Net")
            'Else
            '    dblUnitPrice = pdtRow.Item("Add_Col_Setup_Prc")
            '    dblRCPrice = pdtRow.Item("Add_Col_RC_Prc")
            'End If

            If Trim(oDec_Met.Setup_Item_No = "") Then Exit Sub

            If colItemExtraByDecMetID.Contains(Trim(pdtRow.Item("Item_Cd")) & "EXTRA_3" & pdtRow.Item("Dec_Met_ID").ToString) Then Exit Sub

            strSql = "EXEC DBO.OEI_Item_Charge_Elements '" & m_oOrder.Ordhead.Ord_GUID & "', '" & Trim(pdtRow.Item("Item_Cd")) & "', 'EXTRA_3', " & pdtRow.Item("Dec_Met_ID").ToString
            dt = db.DataTable(strSql)

            'Debug.Print(Trim(pdtRow.Item("Item_Cd")))

            If dt.Rows.Count <> 0 Then

                For Each dtRow As DataRow In dt.Rows

                    ' Here checks for the color change charge. Color change has no running charge.
                    Dim strMainColor As String
                    strMainColor = Trim(dtRow.Item("Element_Value")).ToUpper
                    If strMainColor.Contains("CMYK") Then strMainColor = "CMYK"
                    If strMainColor.Contains("/") Then strMainColor = Mid(strMainColor, 1, strMainColor.IndexOf("/"))

                    'If pdtRow.Item("Loc_Qty") >= pdtRow.Item("Logo_Qty") Then
                    If dtRow.Item("Field_Qty") > 1 Then

                        ' We do not add the charge on a repeat, only the running charges
                        If dtRow.Item("Repeat_From_ID") = 0 Then

                            If colItemCharges.Contains(Trim(oDec_Met.Add_Col_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                                oItem = colItemCharges(Trim(oDec_Met.Add_Col_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                                If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                                    oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd"))
                                End If
                                oItem.Qty_Ordered += (dtRow.Item("Field_Qty") - 1) ' intKitOrdered
                            Else
                                oItem = New cChargePreviewElement
                                oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                                oItem.Item_No = Trim(oDec_Met.Add_Col_Item_No) ' Trim(dtRow.Item("Item_No"))
                                oItem.Unit_Price = dblUnitPrice
                                oItem.Qty_Ordered = (dtRow.Item("Field_Qty") - 1) ' intKitOrdered
                                oItem.Description = "ADD'L COLOR FOR " & Trim(pdtRow.Item("Item_Cd"))
                                colItemCharges.Add(oItem, Trim(oDec_Met.Add_Col_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                            End If

                        End If

                        ' Supplementary or optional element, we must add running charges
                        If Trim(oDec_Met.RC_Col_Item_No) <> "" Then

                            If colItemCharges.Contains(Trim(oDec_Met.RC_Col_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                                oItem = colItemCharges(Trim(oDec_Met.RC_Col_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                                If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                                    oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd"))
                                End If
                                oItem.Qty_Ordered += ((dtRow.Item("Field_Qty") - 1) * dtRow.Item("Qty_Ordered")) ' intKitOrdered
                            Else
                                oItem = New cChargePreviewElement
                                oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                                oItem.Item_No = Trim(oDec_Met.RC_Col_Item_No) ' Trim(dtRow.Item("Item_No"))
                                oItem.Unit_Price = dblRCPrice
                                oItem.Qty_Ordered = ((dtRow.Item("Field_Qty") - 1) * dtRow.Item("Qty_Ordered")) ' intKitOrdered
                                oItem.Description = "ADD'L COLOR RC FOR " & Trim(pdtRow.Item("Item_Cd"))
                                colItemCharges.Add(oItem, Trim(oDec_Met.RC_Col_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                            End If

                        End If

                    End If

                    If Not (colColorCount.ContainsKey(Trim(pdtRow.Item("Item_Cd")))) Then
                        colColorCount.Add(Trim(pdtRow.Item("Item_Cd")), 1)
                        colColorList.Add(Trim(pdtRow.Item("Item_Cd")), strMainColor)
                        If Not (colColorEnum.Contains(Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">")) Then
                            colColorEnum.Add(Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">", Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">")
                        End If
                    Else
                        If Not (colColorEnum.Contains(Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">")) Then
                            colColorEnum.Add(Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">", Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">")
                            If (colColorCount.ContainsKey(Trim(pdtRow.Item("Item_Cd")))) Then
                                colColorCount.Item(Trim(pdtRow.Item("Item_Cd"))) += 1
                                colColorList.Item(Trim(pdtRow.Item("Item_Cd"))) &= "," & strMainColor
                            End If
                        End If
                    End If

                    'If Not (colColors.Contains(Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">")) Then
                    '    colColors.Add(Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">", Trim(pdtRow.Item("Item_Cd")) & "<" & strMainColor & ">")
                    'End If

                Next

            End If

            colItemExtraByDecMetID.Add(Trim(pdtRow.Item("Item_Cd")) & "EXTRA_3" & pdtRow.Item("Dec_Met_ID").ToString, Trim(pdtRow.Item("Item_Cd")) & "EXTRA_3" & pdtRow.Item("Dec_Met_ID").ToString)

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Preview_Charges_Calc_Add_Loc(ByRef pdtRow As DataRow)

        Dim db As New cDBA
        Dim dt As DataTable
        Dim strSql As String
        Dim oDec_Met As cMDBCfgDecMet
        'Dim strItemCharge As String
        Dim iCount As Integer = 0
        Dim iQty As Integer = 0
        Dim oItem As cChargePreviewElement
        Dim dblUnitPrice As Double
        Dim dblRCPrice As Double
        Dim strLocRCValue As String
        Dim strLocationsList As String = String.Empty
        Dim strLocationsRCList As String = String.Empty

        Try

            oDec_Met = colDecMetList(pdtRow.Item("Dec_Met_ID"))

            'If cptPricingType = ChargesPricingType.Net Then
            dblUnitPrice = pdtRow.Item("Add_Loc_Setup_Net")
            dblRCPrice = pdtRow.Item("Add_Loc_RC_Net")
            'Else
            '    dblUnitPrice = pdtRow.Item("Add_Loc_Setup_Prc")
            '    dblRCPrice = pdtRow.Item("Add_Loc_RC_Prc")
            'End If

            If Trim(oDec_Met.Setup_Item_No = "") Then Exit Sub

            If colItemCharges.Contains(Trim(oDecMet.Add_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                oItem = colItemCharges(Trim(oDecMet.Add_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                If oItem.Description.Contains(pdtRow.Item("Item_CD")) Then
                    Exit Sub
                End If
            End If

            If colItemExtraByDecMetID.Contains(Trim(pdtRow.Item("Item_Cd")) & "EXTRA_2" & pdtRow.Item("Dec_Met_ID").ToString) Then Exit Sub

            strSql = "EXEC DBO.OEI_Item_Charge_Elements '" & m_oOrder.Ordhead.Ord_GUID & "', '" & Trim(pdtRow.Item("Item_Cd")) & "', 'EXTRA_LOC', " & pdtRow.Item("Dec_Met_ID").ToString
            'strSql = "EXEC DBO.OEI_Item_Charge_Elements '" & m_oOrder.Ordhead.Ord_GUID & "', '" & Trim(pdtRow.Item("Item_Cd")) & "', 'EXTRA_2', " & pdtRow.Item("Dec_Met_ID").ToString
            dt = db.DataTable(strSql)
            For Each dtRow As DataRow In dt.Rows

                strLocRCValue = dtRow.Item("Item_Guid").ToString & "<LOC><" & Trim(dtRow.Item("Element_Value")).ToString & ">"

                If Not (colItemRC.Contains(strLocRCValue)) Then

                    ' We do not count the rows when there are repeats, only running charges.
                    If dtRow.Item("Repeat_From_ID") = 0 Then
                        If dtRow.Item("OrderID") > 1 Then ' if dtRow.Item("Field_Qty") > dtRow.Item("Qty_Line") Then
                            iCount += 1 ' (dtRow.Item("Field_Qty") - 1) ' (dtRow.Item("Field_Qty") - dtRow.Item("Qty_Line"))
                            'iQty += (dtRow.Item("Field_Qty") - dtRow.Item("Qty_Line")) * dtRow.Item("Qty_Ordered")
                            strLocationsList &= Trim(dtRow.Item("Element_Value")).ToString & ", "
                        End If
                    End If
                    'If dtRow.Item("Field_Qty") > 1 Then
                    If dtRow.Item("OrderID") > 1 Then
                        'iCount += (dtRow.Item("Field_Qty") - dtRow.Item("Qty_Line"))
                        iQty += dtRow.Item("Qty_Ordered") ' (dtRow.Item("Field_Qty") - 1) * dtRow.Item("Qty_Ordered")
                        strLocationsRCList &= Trim(dtRow.Item("Element_Value")).ToString & ", "
                    End If

                    colItemRC.Add(strLocRCValue, strLocRCValue)
                    'Debug.Print(strLocRCValue)

                End If

            Next

            If strLocationsList.Length > 2 Then strLocationsList = Mid(strLocationsList, 1, strLocationsList.Length - 2)
            If strLocationsRCList.Length > 2 Then strLocationsRCList = Mid(strLocationsRCList, 1, strLocationsRCList.Length - 2)

            If iCount > 0 Then

                If colItemCharges.Contains(Trim(oDecMet.Add_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                    oItem = colItemCharges(Trim(oDecMet.Add_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                    If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                        oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd"))
                    End If
                    oItem.Qty_Ordered += iCount ' intKitOrdered
                Else
                    oItem = New cChargePreviewElement
                    oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                    oItem.Item_No = Trim(oDecMet.Add_Loc_Item_No) ' Trim(dtRow.Item("Item_No"))
                    oItem.Unit_Price = dblUnitPrice
                    oItem.Qty_Ordered = iCount ' 1 ' intKitOrdered
                    oItem.Description = "ADD'L LOCATION FOR " & Trim(pdtRow.Item("Item_Cd")) & " (" & Trim(strLocationsList) & ")"
                    colItemCharges.Add(oItem, Trim(oDecMet.Add_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                End If

            End If

            If iQty > 0 Then

                ' Supplementary or optional element, we must add running charges
                If Trim(oDecMet.RC_Loc_Item_No) <> "" Then

                    If colItemCharges.Contains(Trim(oDecMet.RC_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                        oItem = colItemCharges(Trim(oDecMet.RC_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                        If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                            oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd"))
                        End If
                        oItem.Qty_Ordered += iQty ' intKitOrdered
                    Else
                        oItem = New cChargePreviewElement
                        oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                        oItem.Item_No = Trim(oDecMet.RC_Loc_Item_No) ' Trim(dtRow.Item("Item_No"))
                        oItem.Unit_Price = dblRCPrice
                        oItem.Qty_Ordered = iQty ' intKitOrdered
                        oItem.Description = "ADD'L LOCATION RC FOR " & Trim(pdtRow.Item("Item_Cd")) & " (" & Trim(strLocationsRCList) & ")"
                        colItemCharges.Add(oItem, Trim(oDecMet.RC_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                    End If

                End If

            End If

            colItemExtraByDecMetID.Add(Trim(pdtRow.Item("Item_Cd")) & "EXTRA_2" & pdtRow.Item("Dec_Met_ID").ToString, Trim(pdtRow.Item("Item_Cd")) & "EXTRA_2" & pdtRow.Item("Dec_Met_ID").ToString)

            'End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Sub Preview_Charges_Calc_Add_Loc(ByRef pdtRow As DataRow)

    '    Dim db As New cDBA
    '    Dim dt As DataTable
    '    Dim strSql As String
    '    Dim oDec_Met As cMDBCfgDecMet
    '    'Dim strItemCharge As String
    '    Dim iCount As Integer = 0
    '    Dim oItem As cChargePreviewElement
    '    Dim dblUnitPrice As Double
    '    Dim dblRCPrice As Double

    '    Try
    '        oDec_Met = colDecMetList(pdtRow.Item("Dec_Met_ID"))

    '        If cptPricingType = ChargesPricingType.Net Then
    '            dblUnitPrice = pdtRow.Item("Add_Loc_Setup_Net")
    '            dblRCPrice = pdtRow.Item("Add_Loc_RC_Net")
    '        Else
    '            dblUnitPrice = pdtRow.Item("Add_Loc_Setup_Prc")
    '            dblRCPrice = pdtRow.Item("Add_Loc_RC_Prc")
    '        End If

    '        If Trim(oDec_Met.Setup_Item_No = "") Then Exit Sub

    '        strSql = "EXEC DBO.OEI_Item_Charge_Elements '" & m_oOrder.Ordhead.Ord_GUID & "', '" & Trim(pdtRow.Item("Item_Cd")) & "', 'EXTRA_2', " & pdtRow.Item("Dec_Met_ID").ToString
    '        dt = db.DataTable(strSql)

    '        If dt.Rows.Count <> 0 Then

    '            For Each dtRow As DataRow In dt.Rows

    '                'If pdtRow.Item("Loc_Qty") >= pdtRow.Item("Logo_Qty") Then
    '                If dtRow.Item("Field_Qty") > pdtRow.Item("Logo_Qty") Then

    '                    If colItemCharges.Contains(Trim(oDecMet.Setup_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
    '                        oItem = colItemCharges(Trim(oDecMet.Setup_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
    '                        oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd"))
    '                        oItem.Qty_Ordered += 1 ' intKitOrdered
    '                    Else
    '                        oItem = New cChargePreviewElement
    '                        oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
    '                        oItem.Item_No = Trim(oDecMet.Add_Loc_Item_No) ' Trim(dtRow.Item("Item_No"))
    '                        oItem.Unit_Price = dblUnitPrice
    '                        oItem.Qty_Ordered = 1 ' intKitOrdered
    '                        oItem.Description = "ADD LOC FOR " & Trim(pdtRow.Item("Item_Cd"))
    '                        colItemCharges.Add(oItem, Trim(oDecMet.Add_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
    '                    End If

    '                    ' Supplementary or optional element, we must add running charges
    '                    If Trim(oDecMet.RC_Item_No) <> "" Then

    '                        If colItemCharges.Contains(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
    '                            oItem = colItemCharges(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
    '                            oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd"))
    '                            oItem.Qty_Ordered += pdtRow.Item("Qty_Ordered") ' intKitOrdered
    '                        Else
    '                            oItem = New cChargePreviewElement
    '                            oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
    '                            oItem.Item_No = Trim(oDecMet.RC_Loc_Item_No) ' Trim(dtRow.Item("Item_No"))
    '                            oItem.Unit_Price = dblUnitPrice
    '                            oItem.Qty_Ordered = pdtRow.Item("Qty_Ordered") ' intKitOrdered
    '                            oItem.Description = "ADD LOC RC FOR " & Trim(pdtRow.Item("Item_Cd"))
    '                            colItemCharges.Add(oItem, Trim(oDecMet.RC_Loc_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
    '                        End If

    '                    End If

    '                End If

    '            Next

    '        End If

    '        'If pdtRow.Item("Logo_Qty") <= 1 Then Exit Sub
    '        'If pdtRow.Item("Logo_Count") <= 1 Then Exit Sub

    '        'strItemCharge = Trim(oDec_Met.Setup_Item_No)

    '    Catch er As Exception
    '        MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Private Sub Preview_Charges_Calc_Add_Logos(ByRef pdtRow As DataRow, Optional ByVal pintKitOrdered As Integer = 0)

        Dim db As New cDBA
        Dim dt As DataTable
        Dim strSql As String
        'Dim oDec_Met As cMDBCfgDecMet
        'Dim strItemCharge As String
        Dim iCount As Integer = 0
        'Dim itemCount As Integer = 0
        Dim oItem As cChargePreviewElement
        Dim dblUnitPrice As Double
        Dim strLogoRCValue As String

        Try
            oDecMet = colDecMetList(pdtRow.Item("Dec_Met_ID"))

            If Trim(oDecMet.Setup_Item_No = "") Then Exit Sub

            strSql = "EXEC DBO.OEI_Item_Charge_Elements '" & m_oOrder.Ordhead.Ord_GUID & "', '" & Trim(pdtRow.Item("Item_Cd")) & "', 'EXTRA_1', " & pdtRow.Item("Dec_Met_ID").ToString
            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                'iCount = 0

                For Each dtRow As DataRow In dt.Rows

                    strLogoRCValue = dtRow.Item("Item_Guid").ToString & "<LOGO><" & Trim(dtRow.Item("Element_Value")).ToString & ">"

                    If Not (colLogos.Contains(Trim(dtRow.Item("Element_Value")) & "<" & pdtRow.Item("Dec_Met_ID").ToString & ">")) Then

                        ' IF ITEM IS NOT IN LIST, ADD IT AND CHECK IF CHARGES APPLY
                        colLogos.Add(Trim(dtRow.Item("Element_Value")) & "<" & pdtRow.Item("Dec_Met_ID").ToString & ">", Trim(dtRow.Item("Element_Value")) & "<" & pdtRow.Item("Dec_Met_ID").ToString & ">")
                        'iCount += 1

                        'If dtRow.Item("From_Kit_Item") = 1 Then
                        ' A kit
                        'If iCount > pintKitOrdered Then
                        ' Supplementary element, we must add a setup charge

                        'If cptPricingType = ChargesPricingType.Net Then
                        dblUnitPrice = pdtRow.Item("Setup_Net")
                        'Else
                        '    dblUnitPrice = pdtRow.Item("Setup_Prc")
                        'End If

                        'MB++ IF NOT A REPEAT ORDER THEN
                        If dtRow.Item("Repeat_From_ID") = 0 Then

                            If colItemCharges.Contains(Trim(oDecMet.Setup_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                                oItem = colItemCharges(Trim(oDecMet.Setup_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                                If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                                    oItem.Description &= ", " & Trim(dtRow.Item("Element_Value")) ' Trim(pdtRow.Item("Item_Cd"))
                                Else
                                    Dim iPosTextInsert As Integer
                                    iPosTextInsert = oItem.Description.IndexOf(")", oItem.Description.IndexOf(Trim(pdtRow.Item("Item_Cd")) & " ("))
                                    oItem.Description = oItem.Description.Substring(0, iPosTextInsert) & ", " & Trim(dtRow.Item("Element_Value")) & oItem.Description.Substring(iPosTextInsert)
                                End If
                                oItem.Qty_Ordered += 1 ' intKitOrdered
                            Else
                                ' If it is not a default option, or logo order > 1 means it is an additional logo
                                If dtRow.Item("Default_Opt") = 0 Or dtRow.Item("OrderID") > 1 Or dtRow.Item("From_Kit_Item") = 0 Then
                                    oItem = New cChargePreviewElement
                                    oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                                    oItem.Item_No = Trim(oDecMet.Setup_Item_No) ' Trim(dtRow.Item("Item_No"))
                                    oItem.Unit_Price = dblUnitPrice
                                    ' If it is a kit, the first item is always included.
                                    'If dtRow.Item("From_Kit_Item") = 1 Then
                                    '    oItem.Qty_Ordered = 0 ' intKitOrdered
                                    '    oItem.Description = "SETUP FOR " & Trim(dtRow.Item("Element_Value")) & " (INCL.)" ' Trim(pdtRow.Item("Item_Cd"))
                                    'Else
                                    oItem.Qty_Ordered = 1 ' intKitOrdered
                                    oItem.Description = "SETUP FOR " & Trim(pdtRow.Item("Item_Cd")) & " (" & Trim(dtRow.Item("Element_Value")) & ")" ' Trim(pdtRow.Item("Item_Cd"))
                                    'End If
                                    colItemCharges.Add(oItem, oItem.Item_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                                End If
                            End If
                            'MB++ END IF NOT A REPEAT ORDER
                        End If

                        If dtRow.Item("Repeat_From_ID") <> 0 Or dtRow.Item("Default_Opt") = 0 Or dtRow.Item("OrderID") > 1 Then
                            ' Supplementary element, we also must add running charges
                            If Trim(oDecMet.RC_Item_No) <> "" Then
                                If Not (colItemRC.Contains(strLogoRCValue)) Then
                                    'If cptPricingType = ChargesPricingType.Net Then
                                    dblUnitPrice = pdtRow.Item("RC_Net")
                                    'Else
                                    '    dblUnitPrice = pdtRow.Item("RC_Prc")
                                    'End If

                                    If pdtRow.Item("RC_Net") <> 0 Then
                                        If colItemCharges.Contains(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                                            oItem = colItemCharges(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                                            If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                                                oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd")) & " (" & Trim(dtRow.Item("Element_Value")) & ")" ' ", " & Trim(dtRow.Item("Element_Value")) ' Trim(pdtRow.Item("Item_Cd"))
                                            Else
                                                Dim iPosTextInsert As Integer
                                                Dim iPosTextStart As Integer
                                                iPosTextStart = oItem.Description.IndexOf(Trim(pdtRow.Item("Item_Cd")) & " (")
                                                iPosTextInsert = oItem.Description.IndexOf(")", iPosTextStart)
                                                If Not oItem.Description.Substring(iPosTextStart, iPosTextInsert - iPosTextStart).Contains(Trim(dtRow.Item("Element_Value"))) Then
                                                    oItem.Description = oItem.Description.Substring(0, iPosTextInsert) & ", " & Trim(dtRow.Item("Element_Value")) & oItem.Description.Substring(iPosTextInsert)
                                                End If
                                            End If
                                            oItem.Qty_Ordered += dtRow.Item("Qty_Ordered") ' intKitOrdered
                                        Else
                                            oItem = New cChargePreviewElement
                                            oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                                            oItem.Item_No = Trim(oDecMet.RC_Item_No) ' Trim(dtRow.Item("Item_No"))
                                            oItem.Unit_Price = dblUnitPrice
                                            oItem.Qty_Ordered = dtRow.Item("Qty_Ordered") ' intKitOrdered
                                            oItem.Description = "RC FOR " & Trim(pdtRow.Item("Item_Cd")) & " (" & Trim(dtRow.Item("Element_Value")) & ")" ' Trim(pdtRow.Item("Item_Cd"))

                                            colItemCharges.Add(oItem, oItem.Item_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                                        End If
                                    End If

                                End If

                            End If

                        End If

                        'End If

                        'Else

                        '' ''If iCount > 1 Or pdtRow.Item("Default_Opt") = 0 Then
                        '' '' Supplementary or optional element, we add a setup charge

                        '' ''If cptPricingType = ChargesPricingType.Net Then
                        ' ''dblUnitPrice = pdtRow.Item("Setup_Net")
                        '' ''Else
                        '' ''    dblUnitPrice = pdtRow.Item("Setup_Prc")
                        '' ''End If

                        '' ''MB++ IF NOT A REPEAT ORDER THEN
                        ' ''If dtRow.Item("Repeat_From_ID") = 0 Then
                        ' ''    If colItemCharges.Contains(Trim(oDecMet.Setup_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                        ' ''        oItem = colItemCharges(Trim(oDecMet.Setup_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                        ' ''        If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                        ' ''            oItem.Description &= ", " & Trim(dtRow.Item("Element_Value")) ' Trim(pdtRow.Item("Item_Cd"))
                        ' ''        End If
                        ' ''        If iCount > pdtRow.Item("Default_Loc_Qty") Then
                        ' ''            oItem.Qty_Ordered += 1 ' intKitOrdered
                        ' ''        End If
                        ' ''    Else
                        ' ''        oItem = New cChargePreviewElement
                        ' ''        oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                        ' ''        oItem.Item_No = Trim(oDecMet.Setup_Item_No) ' Trim(dtRow.Item("Item_No"))
                        ' ''        oItem.Unit_Price = dblUnitPrice
                        ' ''        oItem.Qty_Ordered = 1 ' intKitOrdered
                        ' ''        'End If
                        ' ''        oItem.Description = "SETUP FOR " & Trim(dtRow.Item("Element_Value")) ' Trim(pdtRow.Item("Item_Cd"))
                        ' ''        colItemCharges.Add(oItem, oItem.Item_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                        ' ''    End If
                        ' ''End If
                        '' ''MB++ END IF NOT A REPEAT ORDER THEN


                        '' '' Supplementary or optional element, we must add running charges
                        ' ''If Trim(oDecMet.RC_Item_No) <> "" Then
                        ' ''    If Not (colItemRC.Contains(dtRow.Item("Item_Guid").ToString)) Then

                        ' ''        If iCount > 1 Or pdtRow.Item("Default_Opt") = 0 Then
                        ' ''            'If cptPricingType = ChargesPricingType.Net Then
                        ' ''            dblUnitPrice = pdtRow.Item("RC_Net")
                        ' ''            'Else
                        ' ''            '    dblUnitPrice = pdtRow.Item("RC_Prc")
                        ' ''            'End If

                        ' ''            If colItemCharges.Contains(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                        ' ''                oItem = colItemCharges(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                        ' ''                If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                        ' ''                    oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd"))
                        ' ''                End If
                        ' ''                oItem.Qty_Ordered += dtRow.Item("Qty_Ordered") ' intKitOrdered
                        ' ''            Else
                        ' ''                oItem = New cChargePreviewElement
                        ' ''                oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                        ' ''                oItem.Item_No = Trim(oDecMet.RC_Item_No) ' Trim(dtRow.Item("Item_No"))
                        ' ''                oItem.Unit_Price = dblUnitPrice
                        ' ''                oItem.Qty_Ordered = dtRow.Item("Qty_Ordered") ' intKitOrdered
                        ' ''                oItem.Description = "RC FOR " & Trim(pdtRow.Item("Item_Cd"))
                        ' ''                colItemCharges.Add(oItem, oItem.Item_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                        ' ''            End If

                        ' ''        End If
                        ' ''        colItemRC.Add(dtRow.Item("Item_Guid").ToString, dtRow.Item("Item_Guid").ToString)

                        ' ''    End If

                        ' ''End If

                        'End If

                    Else

                        If Not (colItemRC.Contains(strLogoRCValue)) Then

                            ' IF ITEM ALREADY IN LIST, CHECK IF RUNNING CHARGES APPLY
                            If dtRow.Item("Repeat_From_ID") <> 0 Or pdtRow.Item("Default_Opt") = 0 Or dtRow.Item("OrderID") > 1 Or pintKitOrdered <> 0 Then
                                ' It is not the default option, we must add running charges

                                If Trim(oDecMet.RC_Item_No) <> "" Then
                                    'If cptPricingType = ChargesPricingType.Net Then
                                    dblUnitPrice = pdtRow.Item("RC_Net")
                                    'Else
                                    '    dblUnitPrice = pdtRow.Item("RC_Prc")
                                    'End If

                                    If pdtRow.Item("RC_Net") <> 0 Then
                                        If colItemCharges.Contains(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                                            oItem = colItemCharges(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                                            If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                                                oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd")) & " (" & Trim(dtRow.Item("Element_Value")) & ")" ' ", " & Trim(pdtRow.Item("Item_Cd"))
                                            Else
                                                Dim iPosTextInsert As Integer
                                                Dim iPosTextStart As Integer
                                                iPosTextStart = oItem.Description.IndexOf(Trim(pdtRow.Item("Item_Cd")) & " (")
                                                iPosTextInsert = oItem.Description.IndexOf(")", iPosTextStart)
                                                If Not oItem.Description.Substring(iPosTextStart, iPosTextInsert - iPosTextStart).Contains(Trim(dtRow.Item("Element_Value"))) Then
                                                    oItem.Description = oItem.Description.Substring(0, iPosTextInsert) & ", " & Trim(dtRow.Item("Element_Value")) & oItem.Description.Substring(iPosTextInsert)
                                                End If
                                            End If
                                            oItem.Qty_Ordered += dtRow.Item("Qty_Ordered") ' intKitOrdered
                                        Else
                                            oItem = New cChargePreviewElement
                                            oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                                            oItem.Item_No = Trim(oDecMet.RC_Item_No) ' Trim(dtRow.Item("Item_No"))
                                            oItem.Unit_Price = dblUnitPrice
                                            oItem.Qty_Ordered = dtRow.Item("Qty_Ordered") ' intKitOrdered
                                            oItem.Description = "RC FOR " & Trim(pdtRow.Item("Item_Cd")) & " (" & Trim(dtRow.Item("Element_Value")) & ")" ' Trim(pdtRow.Item("Item_Cd"))
                                            colItemCharges.Add(oItem, oItem.Item_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                                        End If
                                    End If

                                End If
                                'End If

                                'ElseIf pintKitOrdered <> 0 Or dtRow.Item("OrderID") > 1 Then

                                '    If dtRow.Item("OrderID") > 1 Then
                                '        ' It is a second logo, we must add running charges

                                '        If Trim(oDecMet.RC_Item_No) <> "" Then
                                '            'If cptPricingType = ChargesPricingType.Net Then
                                '            dblUnitPrice = pdtRow.Item("RC_Net")
                                '            'Else
                                '            '    dblUnitPrice = pdtRow.Item("RC_Prc")
                                '            'End If

                                '            If colItemCharges.Contains(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                                '                oItem = colItemCharges(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                                '                If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                                '                    oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd"))
                                '                End If
                                '                oItem.Qty_Ordered += dtRow.Item("Qty_Ordered") ' intKitOrdered
                                '            Else
                                '                oItem = New cChargePreviewElement
                                '                oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                                '                oItem.Item_No = Trim(oDecMet.RC_Item_No) ' Trim(dtRow.Item("Item_No"))
                                '                oItem.Unit_Price = dblUnitPrice
                                '                oItem.Qty_Ordered = dtRow.Item("Qty_Ordered") ' intKitOrdered
                                '                oItem.Description = "RC FOR " & Trim(pdtRow.Item("Item_Cd"))
                                '                colItemCharges.Add(oItem, oItem.Item_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                                '            End If

                                '        End If

                                '    End If

                                'Else

                                '    If dtRow.Item("OrderID") > 1 Then
                                '        ' It is a second logo, we must add running charges

                                '        If Trim(oDecMet.RC_Item_No) <> "" Then
                                '            'If cptPricingType = ChargesPricingType.Net Then
                                '            dblUnitPrice = pdtRow.Item("RC_Net")
                                '            'Else
                                '            '    dblUnitPrice = pdtRow.Item("RC_Prc")
                                '            'End If

                                '            If colItemCharges.Contains(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                                '                oItem = colItemCharges(Trim(oDecMet.RC_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                                '                If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                                '                    oItem.Description &= ", " & Trim(pdtRow.Item("Item_Cd"))
                                '                End If
                                '                oItem.Qty_Ordered += dtRow.Item("Qty_Ordered") ' intKitOrdered
                                '            Else
                                '                oItem = New cChargePreviewElement
                                '                oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                                '                oItem.Item_No = Trim(oDecMet.RC_Item_No) ' Trim(dtRow.Item("Item_No"))
                                '                oItem.Unit_Price = dblUnitPrice
                                '                oItem.Qty_Ordered = dtRow.Item("Qty_Ordered") ' intKitOrdered
                                '                oItem.Description = "RC FOR " & Trim(pdtRow.Item("Item_Cd"))
                                '                colItemCharges.Add(oItem, oItem.Item_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                                '            End If

                                '        End If

                                '    End If

                            End If

                        End If

                    End If

                    Call Add_Corr_Location_RC_To_Collection(dtRow, pdtRow.Item("Item_Cd"), pdtRow.Item("Dec_Met_ID"))

                    If Not colItemRC.Contains(strLogoRCValue) Then
                        colItemRC.Add(strLogoRCValue, strLogoRCValue)
                        'Debug.Print(strLogoRCValue)
                    End If

                Next

            End If

            'If pdtRow.Item("Logo_Qty") <= 1 Then Exit Sub
            'If pdtRow.Item("Logo_Count") <= 1 Then Exit Sub

            'strItemCharge = Trim(oDec_Met.Setup_Item_No)

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Add_Corr_Location_RC_To_Collection(ByRef pdtRow As DataRow, ByVal pstrItem_Cd As String, ByVal pDec_Met_ID As Integer)
        ' This procedure is to include lines from additional logos to simulate
        ' RC for additional locations related to that logo. 
        ' When 2 logos, 1 logo add RC, no add location RC because charged on logo.
        Try
            Dim db As New cDBA
            Dim dt As DataTable
            Dim strLocRCValue As String

            Dim strSql As String = ""
            strSql = "EXEC DBO.OEI_Item_Charge_Elements '" & m_oOrder.Ordhead.Ord_GUID & "', '" & Trim(pstrItem_Cd) & "', 'EXTRA_LOC', " & pDec_Met_ID.ToString

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                For Each dtRow As DataRow In dt.Rows
                    If dtRow.Item("Item_Guid") = pdtRow.Item("Item_Guid") And dtRow.Item("OrderID") = pdtRow.Item("OrderID") Then
                        strLocRCValue = dtRow.Item("Item_Guid").ToString & "<LOC><" & Trim(dtRow.Item("Element_Value")).ToString & ">"
                        If Not colItemRC.Contains(strLocRCValue) Then
                            colItemRC.Add(strLocRCValue, strLocRCValue)
                            'Debug.Print(strLocRCValue)
                        End If
                    End If
                Next
            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Preview_Charges_Calc_Add_Kit_Logos(ByRef pdtRow As DataRow)

        Dim db As New cDBA
        Dim dt As DataTable
        Dim dtLines As DataTable
        Dim strSql As String
        'Dim oDec_Met As cMDBCfgDecMet
        'Dim strItemCharge As String
        Dim iCount As Integer = 0
        'Dim itemCount As Integer = 0
        Dim oItem As cChargePreviewElement
        Dim dblUnitPrice As Double = 0
        Dim strLogo As String = String.Empty

        Try

            ' First, add the item charges
            strSql = _
            "SELECT		DISTINCT K.KIT_DEC_MET_ID, K2.Kit_Setup_Net, DM.SETUP_ITEM_NO, " & _
            "			DM.SETUP_NET, MAX(LOGO_QTY) AS LOGO_QTY " & _
            "FROM		OEI_UNIQUE_IMPRINT_KITS K " & _
            "INNER JOIN " & _
            "	( " & _
            "		SELECT		KIT_DEC_MET_ID, DEC_MET_ID, Kit_Setup_Net, COUNT(EXTRA_1) AS LOGO_QTY " & _
            "		FROM		OEI_UNIQUE_IMPRINT_KITS " & _
            "		WHERE		Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' " & _
            "		GROUP BY	KIT_DEC_MET_ID, DEC_MET_ID, Kit_Setup_Net " & _
            "	) K2 ON	K.KIT_DEC_MET_ID = K2.KIT_DEC_MET_ID " & _
            "INNER JOIN	MDB_CFG_DEC_MET DM ON K.KIT_DEC_MET_ID = DM.DEC_MET_ID " & _
            "WHERE		K.Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' " & _
            "GROUP BY	K.KIT_DEC_MET_ID, K2.Kit_Setup_Net, DM.SETUP_ITEM_NO, DM.SETUP_NET "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then

                For Each dtRow As DataRow In dt.Rows

                    oDecMet = colDecMetList(dtRow.Item("Kit_Dec_Met_ID"))

                    If colItemCharges.Contains(Trim(oDecMet.Setup_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">") Then
                        ' Should not happen because the query always return 1 of each kit_dec_met_id per price
                        oItem = colItemCharges(Trim(oDecMet.Setup_Item_No) & "<" & Trim(dblUnitPrice.ToString) & ">")
                        If Not oItem.Description.Contains(Trim(pdtRow.Item("Item_Cd"))) Then
                            oItem.Description &= ", " & Trim(dtRow.Item("Element_Value")) ' Trim(pdtRow.Item("Item_Cd"))
                        End If
                        oItem.Qty_Ordered += dtRow.Item("Kit_Dec_Met_ID") ' intKitOrdered
                    Else
                        oItem = New cChargePreviewElement
                        oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                        oItem.Item_No = Trim(oDecMet.Setup_Item_No) ' Trim(dtRow.Item("Item_No"))

                        strSql = "select * from mdb_item_charges where item_cd = '" & Trim(pdtRow.Item("Item_Cd")) & "'  "
                        dtLines = db.DataTable(strSql)

                        If dtLines.Rows.Count <> 0 Then
                            dblUnitPrice = dtLines.Rows(0).Item("Setup_Net")
                        End If

                        'strSql = "SELECT dbo.OEI_Item_Price('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', '" & oItem.Item_No & "', '', '', 1) as Unit_Price "
                        'dtLines = db.DataTable(strSql)

                        'If dtLines.Rows.Count <> 0 Then
                        '    dblUnitPrice = dtLines.Rows(0).Item("Unit_Price")
                        'End If

                        'dblUnitPrice = oDecMet.Setup_Net

                        oItem.Unit_Price = dblUnitPrice
                        oItem.Qty_Ordered = dtRow.Item("LOGO_QTY")
                        oItem.Description = "KIT SETUP FOR " ' & Trim(pdtRow.Item("Kit_Item_Cd")) & " (INCL.)" ' Trim(pdtRow.Item("Item_Cd"))

                        strSql = _
                        "SELECT * " & _
                        "FROM   OEI_UNIQUE_IMPRINT_KITS " & _
                        "WHERE  Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND " & _
                        "       Kit_Dec_Met_ID = " & dtRow.Item("Kit_Dec_Met_ID") & " AND " & _
                        "       Kit_Setup_Net = " & dtRow.Item("Kit_Setup_Net") & "  "

                        dtLines = db.DataTable(strSql)
                        If dtLines.Rows.Count <> 0 Then

                            For Each dtLineRow As DataRow In dtLines.Rows

                                strLogo = Trim(dtLineRow.Item("Extra_1")).ToUpper

                                If InStr(strLogo, "/") > 0 Then
                                    strLogo = Mid(strLogo, 1, InStr(strLogo, "/") - 1)
                                End If

                                If Not (colLogos.Contains(strLogo & "<" & dtLineRow.Item("Dec_Met_ID").ToString & ">")) Then

                                    ' IF ITEM IS NOT IN LIST, ADD IT AND CHECK IF CHARGES APPLY
                                    colLogos.Add(strLogo & "<" & dtLineRow.Item("Dec_Met_ID").ToString & ">", strLogo & "<" & dtLineRow.Item("Dec_Met_ID").ToString & ">")

                                End If

                                If Not oItem.Description.Contains(Trim(dtLineRow.Item("Kit_Item_Cd").ToString.ToUpper)) Then
                                    oItem.Description &= Trim(dtLineRow.Item("Kit_Item_Cd").ToString.ToUpper) & ", "
                                End If

                            Next dtLineRow

                        End If

                    End If

                    oItem.Description = Mid(oItem.Description, 1, oItem.Description.Length - 2)

                    colItemCharges.Add(oItem, oItem.Item_No & "<" & Trim(dblUnitPrice.ToString) & ">")

                Next dtRow

            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function GetKitQtyOrdered(ByRef pdt As DataTable, ByVal pstrItem_No As String) As Integer

        'strSql = "exec dbo.OEI_ITEM_CHARGES_PRECALC '" & m_oOrder.Ordhead.Ord_GUID & "' "

        GetKitQtyOrdered = 999999

        Try
            For Each drRow As DataRow In pdt.Rows

                If drRow.Item("Repeat_From_ID") = 0 Then

                    If drRow.Item("SqlPart") = 2 Then

                        If Trim(drRow.Item("Kit_Item_Cd")) = Trim(pstrItem_No) And Trim(drRow.Item("Kit_Item_Cd")) <> Trim(drRow.Item("Item_Cd")) Then

                            If drRow.Item("Logo_Count") < GetKitQtyOrdered Then

                                GetKitQtyOrdered = drRow.Item("Logo_Count")

                            End If

                        End If

                    End If

                End If

            Next

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        If GetKitQtyOrdered = 999999 Or GetKitQtyOrdered = 0 Then GetKitQtyOrdered = 1

    End Function

    Private Sub Get_Dec_Met_Configuration(ByRef oDec_Met_Col As Collection)

        Try
            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String = _
            "SELECT Dec_Met_ID " & _
            "FROM MDB_CFG_DEC_MET WITH (Nolock) " & _
            "ORDER BY Dec_Met_ID "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                For Each dtRow As DataRow In dt.Rows

                    Dim oDecMet As New cMDBCfgDecMet(dtRow.Item("Dec_Met_ID"))
                    oDec_Met_Col.Add(oDecMet, oDecMet.Dec_Met_ID.ToString)

                Next

            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function GetPricingType() As ChargesPricingType

        'Dim db As New cDBA
        'Dim dt As DataTable
        'Dim dtPrice As DataTable
        'Dim strSql As String
        'Dim strItem_No As String = ""
        'Dim dblUnit_Price As Double

        GetPricingType = ChargesPricingType.Standard

        'Try
        '    strSql = "SELECT TOP 1 * FROM MDB_CFG_DEC_MET WITH (Nolock) ORDER BY Dec_Met_ID "

        '    dt = db.DataTable(strSql)

        '    If dt.Rows.Count <> 0 Then
        '        strItem_No = dt.Rows(0).Item("Setup_Item_No")
        '    End If

        '    If strItem_No <> "" Then

        '        strSql = "SELECT dbo.OEI_Item_Price('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', '" & strItem_No & "', '', '', 1) as Unit_Price "
        '        dtPrice = db.DataTable(strSql)

        '        If dtPrice.Rows.Count <> 0 Then
        '            dblUnit_Price = dtPrice.Rows(0).Item("Unit_Price")
        '        End If

        '        'If dblUnit_Price = dt.Rows(0).Item("Setup_Net") Then
        '        GetPricingType = ChargesPricingType.Net
        '        'Else
        '        '    GetPricingType = ChargesPricingType.Standard
        '        'End If

        '    End If

        'Catch er As Exception
        '    MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Function

    Private Sub Delete_Preview_Item_Charges()

        Dim db As New cDBA
        Dim strSql As String

        Try

            strSql = "DELETE FROM OEI_CHARGES_PREVIEW WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' "

            db.Execute(strSql)

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Preview_Charges_ProductProof()

        Dim blnProductproof As Boolean = False
        Dim dt As DataTable
        Dim strsql As String
        Dim oItem As cChargePreviewElement

        Try

            For Each oOrdline As cOrdLine In OrdLines

                If oOrdline.ProductProof_Bln Then blnProductproof = True

            Next

            If Not blnProductproof Then Exit Sub

            Dim db As New cDBA

            strsql = _
            "SELECT     C.Item_No, C.Item_Prc, C.Item_Net, C.Description, ISNULL(CU.No_Charge, 0) AS No_Charge " & _
            "FROM       MDB_CFG_CHARGE C WITH (Nolock) " & _
            "LEFT JOIN  MDB_CFG_CHARGE_USAGE CU WITH (Nolock) ON CU.Charge_CD = C.Charge_CD " & _
            "WHERE      (CU.Cus_No = '" & Ordhead.Cus_No & "' OR CU.Cus_No IS NULL) AND " & _
            "           C.Charge_CD = 'PRODUCT_PROOF' "

            dt = db.DataTable(strsql)
            If dt.Rows.Count = 0 Then

                strsql = _
                "SELECT     C.Item_No, C.Item_Prc, C.Item_Net, C.Description, 0 AS No_Charge " & _
                "FROM       MDB_CFG_CHARGE C WITH (Nolock) " & _
                "WHERE      C.Charge_CD = 'PRODUCT_PROOF' "

                dt = db.DataTable(strsql)
            End If

            If dt.Rows.Count <> 0 Then

                Dim drRow As DataRow = dt.Rows(0)
                Dim strItem_No As String = Trim(drRow.Item("Item_No").ToString.ToUpper)
                Dim dblUnitPrice As Double

                'If cptPricingType = ChargesPricingType.Net Then
                'If dt.Rows(0).Item("No_Charge") = 1 Then
                '    dblUnitPrice = 0
                'Else
                dblUnitPrice = dt.Rows(0).Item("Item_Net")
                'End If

                If Not (colItemCharges.Contains(strItem_No & "<" & Trim(dblUnitPrice.ToString) & ">")) Then
                    '    oItem = colItemCharges(strItem_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                    '    oItem.Qty_Ordered = 1
                    'Else
                    oItem = New cChargePreviewElement
                    oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                    oItem.Item_No = Trim(strItem_No) ' Trim(dtRow.Item("Item_No"))
                    If Not drRow.Item("NO_CHARGE") Then
                        oItem.Unit_Price = dblUnitPrice
                    End If
                    oItem.Qty_Ordered = 1
                    oItem.Description = Trim(drRow.Item("DESCRIPTION")) & " " & IIf(oItem.Unit_Price = 0, "AT N/C", "")
                    colItemCharges.Add(oItem, strItem_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                End If

            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Preview_Charges_Always_Charges(ByVal pstrCharge_CD As String)

        Dim dt As DataTable
        Dim db As New cDBA
        Dim strsql As String
        Dim oItem As cChargePreviewElement

        Try
            '"           DBO.OEI_Item_Price('" & Currency & "', '" & Customer_No & "', I.Item_No, '', '', 1) as Unit_Price, " & _

            'strsql = _
            '"SELECT     CU.*, C.Item_No, C.Item_Prc, C.Item_Net, C.Description " & _
            '"FROM       MDB_CFG_CHARGE_USAGE CU WITH (NOLOCK) " & _
            '"INNER JOIN MDB_CFG_CHARGE C WITH (Nolock) ON CU.Charge_CD = C.Charge_CD " & _
            '"WHERE      CU.Cus_No = '" & Ordhead.Cus_No & "' AND " & _
            '"           C.Charge_CD = '" & pstrCharge_CD & "' AND " & _
            '"           CU.Always_Use = 1 "

            strsql = _
            "SELECT     CU.*, C.Item_No, C.Description, " & _
            "           DBO.OEI_Item_Price_20140101('" & Ordhead.Curr_Cd & "', '" & Ordhead.Cus_No & "', C.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Item_Net " & _
            "FROM       MDB_CFG_CHARGE_USAGE CU WITH (NOLOCK) " & _
            "INNER JOIN MDB_CFG_CHARGE C WITH (Nolock) ON CU.Charge_CD = C.Charge_CD " & _
            "WHERE      CU.Cus_No = '" & Ordhead.Cus_No & "' AND " & _
            "           C.Charge_CD = '" & pstrCharge_CD & "' AND " & _
            "           CU.Always_Use = 1 "

            dt = db.DataTable(strsql)
            If dt.Rows.Count <> 0 Then

                Dim drRow As DataRow = dt.Rows(0)
                Dim strItem_No As String = Trim(drRow.Item("Item_No").ToString.ToUpper)
                Dim dblUnitPrice As Double

                'If cptPricingType = ChargesPricingType.Net Then
                dblUnitPrice = dt.Rows(0).Item("Item_Net")
                'Else
                '    dblUnitPrice = dt.Rows(0).Item("Item_Prc")
                'End If

                If (colItemCharges.Contains(strItem_No & "<" & Trim(dblUnitPrice.ToString) & ">")) Then
                    oItem = colItemCharges(strItem_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                    oItem.Qty_Ordered += 1
                Else
                    oItem = New cChargePreviewElement
                    oItem.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                    oItem.Item_No = Trim(strItem_No) ' Trim(dtRow.Item("Item_No"))
                    If Not drRow.Item("NO_CHARGE") Then
                        oItem.Unit_Price = dblUnitPrice
                    End If
                    oItem.Qty_Ordered += 1
                    oItem.Description = "ALWAYS " & Trim(drRow.Item("DESCRIPTION")) & " " & IIf(oItem.Unit_Price = 0, "AT N/C", "")
                    colItemCharges.Add(oItem, strItem_No & "<" & Trim(dblUnitPrice.ToString) & ">")
                End If

            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Set_Charge_Active(ByVal pstrCharge_CD As String, ByRef pFlag As Boolean)

        Try

            Dim dt As DataTable
            Dim db As New cDBA
            Dim strsql As String

            strsql = _
            "SELECT     CU.*, C.Item_No, C.Description, " & _
            "           DBO.OEI_Item_Price_20140101('" & Ordhead.Curr_Cd & "', '" & Ordhead.Cus_No & "', C.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) as Item_Net " & _
            "FROM       MDB_CFG_CHARGE_USAGE CU WITH (NOLOCK) " & _
            "INNER JOIN MDB_CFG_CHARGE C WITH (Nolock) ON CU.Charge_ID = C.Charge_ID " & _
            "WHERE      CU.Cus_No = '" & Ordhead.Cus_No & "' AND " & _
            "           C.Charge_CD = '" & pstrCharge_CD & "' AND " & _
            "           CU.Always_Use = 1 "

            dt = db.DataTable(strsql)
            If dt.Rows.Count <> 0 Then pFlag = True

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function HV_US_MACOLAIMPORT1CAD_Folder() As String

        HV_US_MACOLAIMPORT1CAD_Folder = "\\zeus\Harvest Import Service\"

        Try

            Dim strPath As String

            Dim dt As New DataTable
            Dim db As New cDBA

            Dim strSql As String =
            "SELECT * " &
            "FROM   OEI_Config C WITH (Nolock) " &
            "WHERE  ISNULL(HV_US_MacolaImport1CAD_OE, '') <> '' "

            dt = db.DataTable(strSql)
            'If dt.Rows.Count <> 0 Then Call RestartProgram()
            If dt.Rows.Count <> 0 Then

                strPath = Trim(dt.Rows(0).Item("HV_US_MacolaImport1CAD_OE"))
                If Mid(strPath, strPath.Length, 1) <> "\" Then strPath = strPath & "\"

                '==============================
                'Added May 16, 2017
                'TEMPORARILY MODIFIED SO WILL NOT PRINT OUT TO LIVE DIRECTORY
                'Modified July 28, 2017
                '==============================
                'THIS IS THE CORRECT DIRECTORY - PUT IT BACK IN  HV_Import_Files_Folder = strPath & "Files\"
                If gsServerName = "BARIUM" Then
                    HV_US_MACOLAIMPORT1CAD_Folder = strPath  '****************************** PUT THIS LINE BACK In
                Else
                    '++ID 07.15.2020
                    HV_US_MACOLAIMPORT1CAD_Folder = strPath
                    '++ID 07.15.2020 we will follow info from the table OEi_CONFIG
                    '' HV_Import_Files_Folder = "\\EXACT\Harvest Ventures\Utilities and BATs\HV_MacolaImport_Platinum\TEST_PLATINUM_OUTPUT\"
                    '  MsgBox("x  - JUST TESTING SO USE A DIFFERENT DIRECTORY" & vbCrLf & HV_Import_Files_Folder)

                End If
                '==============================

            End If

            dt.Dispose()

        Catch er As Exception
            ' Do nothing
        End Try

    End Function
    Public Function HV_Import_Files_Folder() As String

        HV_Import_Files_Folder = "\\zeus\Harvest Import Service\"

        Try

            Dim strPath As String

            Dim dt As New DataTable
            Dim db As New cDBA

            Dim strSql As String = _
            "SELECT * " & _
            "FROM   OEI_Config C WITH (Nolock) " & _
            "WHERE  ISNULL(HV_Import_Files_Folder, '') <> '' "

            dt = db.DataTable(strSql)
            'If dt.Rows.Count <> 0 Then Call RestartProgram()
            If dt.Rows.Count <> 0 Then

                strPath = Trim(dt.Rows(0).Item("HV_Import_Files_Folder"))
                If Mid(strPath, strPath.Length, 1) <> "\" Then strPath = strPath & "\"

                '==============================
                'Added May 16, 2017
                'TEMPORARILY MODIFIED SO WILL NOT PRINT OUT TO LIVE DIRECTORY
                'Modified July 28, 2017
                '==============================
                'THIS IS THE CORRECT DIRECTORY - PUT IT BACK IN  HV_Import_Files_Folder = strPath & "Files\"
                If gsServerName = "BARIUM" Then
                    HV_Import_Files_Folder = strPath & "Files\" '****************************** PUT THIS LINE BACK In
                Else
                    '++ID 07.15.2020
                    HV_Import_Files_Folder = strPath & "Files\"
                    '++ID 07.15.2020 we will follow info from the table OEi_CONFIG
                    '' HV_Import_Files_Folder = "\\EXACT\Harvest Ventures\Utilities and BATs\HV_MacolaImport_Platinum\TEST_PLATINUM_OUTPUT\"
                    '  MsgBox("x  - JUST TESTING SO USE A DIFFERENT DIRECTORY" & vbCrLf & HV_Import_Files_Folder)

                End If
                '==============================

            End If

            dt.Dispose()

        Catch er As Exception
            ' Do nothing
        End Try

    End Function

    '++ID 02.20.2023

    Public Function HV_US_Import_Files_Folder() As String

        HV_US_Import_Files_Folder = "\\cobalt1\Harvest Ventures\HV_BuySell\Files\"

        Try

            Dim strPath As String

            Dim dt As New DataTable
            Dim db As New cDBA

            Dim strSql As String =
            "SELECT * " &
            "FROM   OEI_Config C WITH (Nolock) " &
            "WHERE  ISNULL(HV_US_Import_Files_Folder, '') <> '' "

            dt = db.DataTable(strSql)
            'If dt.Rows.Count <> 0 Then Call RestartProgram()
            If dt.Rows.Count <> 0 Then

                strPath = Trim(dt.Rows(0).Item("HV_US_Import_Files_Folder"))
                If Mid(strPath, strPath.Length, 1) <> "\" Then strPath = strPath & "\"

                '==============================
                'Added May 16, 2017
                'TEMPORARILY MODIFIED SO WILL NOT PRINT OUT TO LIVE DIRECTORY
                'Modified July 28, 2017
                '==============================
                'THIS IS THE CORRECT DIRECTORY - PUT IT BACK IN  HV_Import_Files_Folder = strPath & "Files\"
                If gsServerName = "BARIUM" Then
                    HV_US_Import_Files_Folder = strPath '& "Files\" '****************************** PUT THIS LINE BACK In
                Else
                    '++ID 07.15.2020
                    HV_US_Import_Files_Folder = strPath '& "Files\"
                    '++ID 07.15.2020 we will follow info from the table OEi_CONFIG
                    '' HV_Import_Files_Folder = "\\EXACT\Harvest Ventures\Utilities and BATs\HV_MacolaImport_Platinum\TEST_PLATINUM_OUTPUT\"
                    '  MsgBox("x  - JUST TESTING SO USE A DIFFERENT DIRECTORY" & vbCrLf & HV_Import_Files_Folder)

                End If
                '==============================

            End If

            dt.Dispose()

        Catch er As Exception
            ' Do nothing
        End Try

    End Function

    Public Function HV_US_PO_Import_Files_Folder() As String

        HV_US_PO_Import_Files_Folder = "\\cobalt1\Harvest Ventures\HV_BuySell\FilesPO\"

        Try

            Dim strPath As String

            Dim dt As New DataTable
            Dim db As New cDBA

            Dim strSql As String =
            "SELECT * " &
            "FROM   OEI_Config C WITH (Nolock) " &
            "WHERE  ISNULL(HV_US_Import_Files_Folder, '') <> '' "

            dt = db.DataTable(strSql)
            'If dt.Rows.Count <> 0 Then Call RestartProgram()
            If dt.Rows.Count <> 0 Then

                strPath = Trim(dt.Rows(0).Item("HV_US_PO_Import_Files_Folder"))
                If Mid(strPath, strPath.Length, 1) <> "\" Then strPath = strPath & "\"

                '==============================
                'Added May 16, 2017
                'TEMPORARILY MODIFIED SO WILL NOT PRINT OUT TO LIVE DIRECTORY
                'Modified July 28, 2017
                '==============================
                'THIS IS THE CORRECT DIRECTORY - PUT IT BACK IN  HV_Import_Files_Folder = strPath & "Files\"
                If gsServerName = "BARIUM" Then
                    HV_US_PO_Import_Files_Folder = strPath '& "Files\" '****************************** PUT THIS LINE BACK In
                Else
                    '++ID 07.15.2020
                    HV_US_PO_Import_Files_Folder = strPath '& "Files\"
                    '++ID 07.15.2020 we will follow info from the table OEi_CONFIG
                    '' HV_Import_Files_Folder = "\\EXACT\Harvest Ventures\Utilities and BATs\HV_MacolaImport_Platinum\TEST_PLATINUM_OUTPUT\"
                    '  MsgBox("x  - JUST TESTING SO USE A DIFFERENT DIRECTORY" & vbCrLf & HV_Import_Files_Folder)

                End If
                '==============================

            End If

            dt.Dispose()

        Catch er As Exception
            ' Do nothing
        End Try

    End Function

    Public Sub LoadOrderFromXmlFile(ByVal piOrdEDI_ID As Integer)
        Dim EdiFailureCd As String = "EDI_IME"
        Dim shipArr() As String = {
                "ship_to_name",
                "ship_to_addr_1",
                "ship_to_addr_2",
                "ship_to_city",
                "ship_to_zip",
                "ship_to_state",
                "ship_to_country",
                "ship_via_cd",
                "ship_instruction_1",
                "ship_instruction_2",
                "contact_1",
                "phone_number",
                "fax_number",
                "email_address",
                "shipping_dt"
            }

        Try
            Dim oOrderEdi As New cOrderEDI(piOrdEDI_ID)

            Dim oXmlFile As New XmlDocument()

            Dim dt As DataTable
            Dim dtHeader As New DataTable
            Dim dtLines As New DataTable
            Dim dtLineRow As DataRow

            Dim db As New cDBA
            Dim dtRow As DataRow = Nothing
            Dim strCus_No As String

            Dim strSql As String
            Dim iQtyOrdered As Integer = 0
            Dim iQtyRemainAssorted As Integer = 0

            Dim blnExact_Qty As Boolean = False
            Dim blnExact_Qty_Flag As Boolean = False

            Dim EdiFailureDesc As String

            strSql = _
            "SELECT TOP 1 * " & _
            "FROM   OEI_OrdEDI WITH (Nolock) " & _
            "WHERE  OrdEDI_ID = " & piOrdEDI_ID.ToString() & " AND " & _
            "       OEI_Status = 'R' "

            ' First part - create order header

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                oXmlFile.LoadXml(dt.Rows(0).Item("Xml_Text").ToString)
            Else
                MsgBox("There is no EDI order pending.")
                Exit Sub
            End If

            Dim oNode As XmlNode
            oNode = oXmlFile.SelectSingleNode("/ExactRequestSet/MacolaSalesOrderAddRequest/OrderHeader")
            Debug.Print(oNode.SelectSingleNode("cus_no").InnerText)
            'throw error before beginning
            If oNode.SelectSingleNode("cus_no") Is Nothing OrElse oNode.SelectSingleNode("cus_no").InnerText = "" Then
                EdiFailureDesc = "Customer number was not found."
                Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & "Empty" & "-" & piOrdEDI_ID.ToString())
            End If

            strSql = "SELECT * FROM MDB_CUSTOMER WHERE EDI_Cus_Name = '" & oNode.SelectSingleNode("cus_no").InnerText & "' "
            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                strCus_No = dt.Rows(0).Item("Cus_No").ToString()
            Else
                MsgBox("No customer available for name " & oNode.SelectSingleNode("cus_no").InnerText & " in MDB_CUSTOMER table.")
                Exit Sub
            End If

            'For Each forecastNode As XmlNode In doc.SelectNodes("/response/hourly_forecast/forecast")
            'For Each oNode As XmlNode In oXmlFile.SelectNodes("/ExactRequestSet/MacolaSalesOrderAddRequest/OrderHeader")

            With Ordhead

                .OEI_Ord_No = g_User.NextOrderNumber()
                .Ord_Dt = Now.Date
                .Shipping_Dt = NoDate()

                .Ord_Type = "O" ' Trim(dtRow.Item("Ord_Type").ToString)

                ' HERE - CALL THE PROC TO GET THE CUSTOMER NUMBER

                .Cus_No = strCus_No
                Call .GetCustomerDefaultValues()

                '.Cus_Alt_Adr_Cd = string.Empty
                If Not oNode.SelectSingleNode("oe_po_no") Is Nothing AndAlso oNode.SelectSingleNode("oe_po_no").InnerText <> "" Then
                    .Oe_Po_No = oNode.SelectSingleNode("oe_po_no").InnerText

                    'check to see if duplicate
                    If m_oOrder.Validation.POExists(strCus_No, .Oe_Po_No, True) = True Then
                        EdiFailureDesc = "Order po no " & .Oe_Po_No & " already exists in the system."
                        Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                    End If
                Else
                    EdiFailureDesc = "Order po no was not found."
                    Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                End If

                'Try this mofo -------------
                'EdiFailureDesc = "Whatever is"
                'Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                '---------------------------

                '.Slspsn_No = String.Empty
                '.Slspsn_Pct_Comm = String.Empty

                '.User_Def_Fld_4 = String.Empty
                '.Cus_Email_Address = String.Empty
                '.User_Def_Fld_1 = String.Empty
                '.User_Def_Fld_5 = String.Empty
                '.Ship_Instruction_1 = String.Empty
                '.Ship_Instruction_2 = String.Empty

                'check shipping destination fields
                For Each field As String In shipArr
                    Debug.Print(oNode.SelectSingleNode(field).InnerText)
                    If oNode.SelectSingleNode(field) Is Nothing OrElse oNode.SelectSingleNode(field).InnerText = "" Then
                        If field = "ship_to_addr_2" Or field = "ship_via_cd" Or field = "ship_instruction_1" Or field = "fax_number" Or field = "shipping_dt" Then
                            Continue For
                        Else
                            EdiFailureDesc = "Shipping field " & field & "was not found."
                            Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                        End If
                    End If
                Next

                ' Ship instructions are sent in reverse so we flip the fields for OEI
                .Ship_Instruction_1 = oNode.SelectSingleNode("ship_instruction_2").InnerText
                If Not oNode.SelectSingleNode("ship_instruction_1") Is Nothing Then .Ship_Instruction_2 = oNode.SelectSingleNode("ship_instruction_1").InnerText
                If Not oNode.SelectSingleNode("ship_instruction_2") Is Nothing Then .Ship_Instruction_1 = oNode.SelectSingleNode("ship_instruction_2").InnerText

                .Ship_To_Name = oNode.SelectSingleNode("ship_to_name").InnerText
                .Ship_To_Addr_1 = oNode.SelectSingleNode("ship_to_addr_1").InnerText
                If Not oNode.SelectSingleNode("ship_to_addr_2") Is Nothing Then .Ship_To_Addr_2 = oNode.SelectSingleNode("ship_to_addr_2").InnerText
                '.Ship_To_Addr_3 = String.Empty
                '.Ship_To_Addr_4 = String.Empty
                .Ship_To_City = oNode.SelectSingleNode("ship_to_city").InnerText
                .Ship_To_State = oNode.SelectSingleNode("ship_to_state").InnerText
                .Ship_To_Zip = oNode.SelectSingleNode("ship_to_zip").InnerText
                .Ship_To_Country = Replace(oNode.SelectSingleNode("ship_to_country").InnerText.Trim(), "USA", "US")

                '.Orig_Trx_Rt = dtRow.Item("Curr_Trx_Rt")
                .Contact_1 = oNode.SelectSingleNode("contact_1").InnerText
                .User_Def_Fld_4 = oNode.SelectSingleNode("contact_1").InnerText

                .Phone_Number = oNode.SelectSingleNode("phone_number").InnerText
                If Not oNode.SelectSingleNode("fax_number") Is Nothing Then .Fax_Number = oNode.SelectSingleNode("fax_number").InnerText

                .Email_Address = oNode.SelectSingleNode("email_address").InnerText
                .Cus_Email_Address = oNode.SelectSingleNode("email_address").InnerText

                'requested date (may omit for now)
                'If Not oNode.SelectSingleNode("shipping_dt") Is Nothing AndAlso oNode.SelectSingleNode("shipping_dt").InnerText <> "" Then
                '    If IsDate(oNode.SelectSingleNode("shipping_dt").InnerText) Then
                '        .In_Hands_Date = Date.Parse(oNode.SelectSingleNode("shipping_dt").InnerText)
                '    End If
                'Else
                '    .In_Hands_Date = Nothing
                '    ' MsgBox("Specify that order is in under 48 hour rocket service criteria for request date")
                'End If

                .Edi_Fg = "S" ' S for STAPLES FILE. TRIGGER WILL DO SOMETHING ABOUT IT.

                .Save() 'TODO: needs to get moved as the order is not even validated yet

                Set_Charge_Active("EXACT_QTY", blnExact_Qty)

                WorkLine = 1

                For Each oLineNode As XmlNode In oNode.SelectNodes("OrderLines/OrderLine")

                    If oLineNode.SelectSingleNode("qty_ordered") Is Nothing OrElse Not IsNumeric(oLineNode.SelectSingleNode("qty_ordered").InnerText) OrElse oLineNode.SelectSingleNode("qty_ordered").InnerText <= 0 Then
                        EdiFailureDesc = "Invalid quantity was entered."
                        Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                    End If

                    If Not oLineNode.SelectSingleNode("item_no") Is Nothing AndAlso oLineNode.SelectSingleNode("item_no").InnerText <> "" Then
                        strSql = _
                        "SELECT     DISTINCT C.CUS_DEC_ID, C.SHORT_CTL_NO, D.KIT_ITEM_NO " & _
                        "FROM       MDB_CUS_DEC C WITH (NOLOCK) " & _
                        "INNER JOIN MDB_CUS_DEC_DETAIL D WITH (NOLOCK) ON C.CUS_DEC_ID = D.CUS_DEC_ID " & _
                        "WHERE      SHORT_CTL_NO = '" & oLineNode.SelectSingleNode("item_no").InnerText.Replace("'", "''") & "' AND " & _
                        "           C.ASSORTED = 1 AND D.SELECTED = 1 "

                        dtLines = db.DataTable(strSql)
                        Select Case dtLines.Rows.Count
                            Case 0
                                ' Not an assorted
                                strSql = _
                                "SELECT     DISTINCT C.CUS_DEC_ID, C.SHORT_CTL_NO, D.KIT_ITEM_NO " & _
                                "FROM       MDB_CUS_DEC C WITH (NOLOCK) " & _
                                "INNER JOIN MDB_CUS_DEC_DETAIL D WITH (NOLOCK) ON C.CUS_DEC_ID = D.CUS_DEC_ID " & _
                                "WHERE SHORT_CTL_NO = '" & oLineNode.SelectSingleNode("item_no").InnerText.Replace("'", "''") & "' "
                                dtLines = db.DataTable(strSql)

                                If dtLines.Rows.Count = 0 Then
                                    EdiFailureDesc = "No repeat order data was found"
                                    Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                                End If

                                If Not oLineNode.SelectSingleNode("qty_ordered") Is Nothing Then
                                    If IsNumeric(oLineNode.SelectSingleNode("qty_ordered").InnerText) Then
                                        iQtyOrdered = Integer.Parse(oLineNode.SelectSingleNode("qty_ordered").InnerText)
                                    End If
                                    iQtyRemainAssorted = 0
                                End If

                            Case 1
                                ' An assorted with only 1 selected.
                                If Not oLineNode.SelectSingleNode("qty_ordered") Is Nothing Then
                                    If IsNumeric(oLineNode.SelectSingleNode("qty_ordered").InnerText) Then
                                        iQtyOrdered = Integer.Parse(oLineNode.SelectSingleNode("qty_ordered").InnerText)
                                    End If
                                    iQtyRemainAssorted = 0
                                End If

                            Case Else
                                ' An assorted with 2 or more elements
                                If Not oLineNode.SelectSingleNode("qty_ordered") Is Nothing Then
                                    If IsNumeric(oLineNode.SelectSingleNode("qty_ordered").InnerText) Then
                                        iQtyOrdered = Integer.Parse(oLineNode.SelectSingleNode("qty_ordered").InnerText)
                                        iQtyRemainAssorted = iQtyOrdered Mod dtLines.Rows.Count
                                        iQtyOrdered = iQtyOrdered \ dtLines.Rows.Count
                                    End If
                                End If

                        End Select

                    Else
                        EdiFailureDesc = "Invalid item number was entered."
                        Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                    End If

                    Dim iLinePos As Integer = 1 ' This will count the lines in cases of assort. First line will be Y in EXTRA_2, others will be A
                    Dim iLink_Line_Seq_No As Integer

                    For Each dtLineRow In dtLines.Rows

                        Dim oOrdline As New cOrdLine(Ordhead.Ord_GUID)
                        'oOrdline.Source = OEI_SourceEnum.ExternalFile
                        'oOrdline.Ord_Guid = Ordhead.Ord_GUID

                        oOrdline.Line_Seq_No = WorkLine
                        If iLinePos = 1 Then iLink_Line_Seq_No = oOrdline.Line_Seq_No

                        'oOrdline.OEOrdLin_ID = 1 ' Integer.Parse(oNode.SelectSingleNode("line_no").InnerText)
                        'oOrdline.SaveToDB = True
                        ' HERE WE MUST GET ITEM NO FROM ITEM LINK TO MDB_CUS_DEC
                        oOrdline.Extra_11 = dtLineRow.Item("Cus_Dec_ID")
                        oOrdline.Item_No = dtLineRow.Item("Kit_Item_No").ToString()

                        ' We put in field Extra_11 the Cus_Dec_ID we will use when filling the imprint and the traveler

                        ' Location will be default for item
                        ' oOrdline.Loc = oNode.SelectSingleNode("loc").InnerText

                        If Not oLineNode.SelectSingleNode("cus_item_no") Is Nothing AndAlso oLineNode.SelectSingleNode("cus_item_no").InnerText <> "" Then
                            oOrdline.Cus_Item_No = oLineNode.SelectSingleNode("cus_item_no").InnerText.ToString()
                        Else
                            EdiFailureDesc = "Invalid customer item number was entered."
                            Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                        End If

                        'If Not oLineNode.SelectSingleNode("user_def_fld_4") Is Nothing Then oOrdline.User_Def_Fld_4 = oLineNode.SelectSingleNode("user_def_fld_4").InnerText.ToString()
                        oOrdline.User_Def_Fld_4 = oLineNode.SelectSingleNode("item_no").InnerText.ToString()

                        ' Qty to ship is set to qty ordered, because of overpicks.
                        Try
                            oOrdline.Qty_Ordered = iQtyOrdered
                            If iQtyRemainAssorted > 0 Then
                                oOrdline.Qty_Ordered += 1
                                iQtyRemainAssorted -= 1
                            End If
                        Catch oe_er As OEException
                            'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                        End Try

                        Try
                            oOrdline.Qty_To_Ship = oOrdline.Qty_Ordered
                            oOrdline.Qty_To_Ship_Changed(oOrdline.Item_No, oOrdline.Loc, False)
                            'If oOrdline.Qty_Prev_Bkord > 0 Then
                            '    oOrdline.Calculate_ETA()
                            'End If
                        Catch oe_er As OEException
                            'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                        End Try

                        ' Do not enter unit price, let calculate new price on date
                        'oOrdline.Unit_Price = Integer.Parse(oNode.SelectSingleNode("unit_price").InnerText)

                        'oOrdline.Discount_Pct = dtRow.Item("Discount_Pct")
                        'oOrdline.Request_Dt = Now.Date
                        'oOrdline.Promise_Dt = Now.Date

                        If Not oLineNode.SelectSingleNode("req_ship_dt") Is Nothing Then
                            If IsDate(oLineNode.SelectSingleNode("req_ship_dt").InnerText) Then
                                oOrdline.Req_Ship_Dt = Date.Parse(oLineNode.SelectSingleNode("req_ship_dt").InnerText)
                                'Else
                                '    EdiFailureDesc = "Invalid shipping date entered."
                                '    Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                            End If
                            'Else
                            '    EdiFailureDesc = "Requested shipping date has not been entered."
                            '    Throw New Exception(EdiFailureCd & "-" & EdiFailureDesc & "-" & strCus_No & "-" & piOrdEDI_ID.ToString())
                        End If

                        'THIS FLAG WILL ALLOW ONLY THIS LINE TO BE PRINTED INTO THE EDI INVOICE ONLY
                        If iLinePos = 1 Then
                            oOrdline.Extra_2 = "Y"
                        Else
                            oOrdline.Extra_2 = "A"
                            oOrdline.Extra_6 = iLink_Line_Seq_No.ToString.Trim
                        End If

                        ''oOrdline.Source = OEI_SourceEnum.frmOrder
                        oOrdline.Save()

                        'AddExternalLine(oOrdline)
                        AddOrderLine(oOrdline)

                        ' Only add additional items one time per item, or one time per assort.
                        ' Must check quantity entered is the total qty, not qty per item.
                        If iLinePos = dtLines.Rows.Count Then

                            'Now check if there are additional line items we must insert 
                            'like charges or else.
                            Dim dtSupp_Item As DataTable
                            strSql = _
                            "SELECT * FROM MDB_CUS_DEC_SUPP_ITEM WITH (NOLOCK) WHERE CUS_DEC_ID = " & dtLineRow.Item("Cus_Dec_ID").ToString & " ORDER BY CUS_DEC_SUPP_ITEM_ID "

                            dtSupp_Item = db.DataTable(strSql)

                            If dtSupp_Item.Rows.Count <> 0 Then

                                For Each drItem As DataRow In dtSupp_Item.Rows

                                    Dim oOrdlineSupp = New cOrdLine(Ordhead.Ord_GUID)

                                    oOrdlineSupp.Line_Seq_No = WorkLine
                                    ' HERE WE MUST GET ITEM NO FROM ITEM LINK TO MDB_CUS_DEC
                                    oOrdlineSupp.Extra_11 = dtLineRow.Item("Cus_Dec_ID")
                                    oOrdlineSupp.Item_No = drItem.Item("Item_No").ToString()

                                    ' Qty to ship is set to qty ordered, because of overpicks.
                                    Try
                                        If drItem.Item("Qty_Ordered") = 0 Then
                                            oOrdlineSupp.Qty_Ordered = Integer.Parse(oLineNode.SelectSingleNode("qty_ordered").InnerText)
                                        Else
                                            oOrdlineSupp.Qty_Ordered = drItem.Item("Qty_Ordered") ' iQtyOrdered
                                        End If

                                    Catch oe_er As OEException
                                        'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                                    End Try

                                    Try
                                        oOrdlineSupp.Qty_To_Ship = oOrdlineSupp.Qty_Ordered
                                        oOrdlineSupp.Qty_To_Ship_Changed(oOrdlineSupp.Item_No, oOrdlineSupp.Loc, False)
                                    Catch oe_er As OEException
                                        'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                                    End Try


                                    Try
                                        If drItem.Item("Unit_Price") <> 0 Then
                                            oOrdlineSupp.Unit_Price = drItem.Item("Unit_Price") ' iQtyOrdered
                                        End If
                                        oOrdlineSupp.Request_Dt = oOrdline.Request_Dt
                                    Catch oe_er As OEException
                                        'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                                    End Try

                                    'THIS FLAG WILL ALLOW ONLY THIS LINE TO BE PRINTED INTO THE EDI INVOICE ONLY
                                    oOrdlineSupp.Extra_2_Bln = False
                                    oOrdlineSupp.Extra_6 = iLink_Line_Seq_No.ToString.Trim

                                    ''oOrdline.Source = OEI_SourceEnum.frmOrder
                                    oOrdlineSupp.Save()

                                    'AddExternalLine(oOrdline)
                                    AddOrderLine(oOrdlineSupp)
                                    WorkLine += 1

                                Next

                            End If

                        End If

                        iLinePos += 1

                    Next dtLineRow

                Next oLineNode

                ' Now we create imprint data for lines
                For Each oOrdline As cOrdLine In OrdLines

                    Dim dtSpecs As DataTable
                    strSql = _
                    "SELECT * " & _
                    "FROM   DBO.OEI_ORD_CUS_DEC_PREVIEW " & _
                    "WHERE  ORD_GUID = '" & oOrdline.Ord_Guid & "' AND " & _
                    "       ITEM_GUID = '" & oOrdline.Item_Guid & "' "

                    dtSpecs = db.DataTable(strSql)

                    If dtSpecs.Rows.Count <> 0 Then

                        Dim drSpec As DataRow
                        drSpec = dtSpecs.Rows(0)

                        oOrdline.Traveler.Load()
                        If Not dtSpecs.Rows(0).Item("RouteID").Equals(DBNull.Value) Then

                            oOrdline.Traveler.RouteID = dtSpecs.Rows(0).Item("RouteID")
                            oOrdline.RouteID = dtSpecs.Rows(0).Item("RouteID")

                        End If
                        oOrdline.Traveler.Save()

                        oOrdline.ImprintLine.Load()

                        With oOrdline.ImprintLine

                            If Not drSpec.Item("Item_No").Equals(DBNull.Value) Then .Item_No = drSpec.Item("Item_No").ToString()
                            If Not drSpec.Item("Color_List").Equals(DBNull.Value) Then .Color = drSpec.Item("Color_List").ToString()
                            If Not drSpec.Item("Logo_List").Equals(DBNull.Value) Then .Imprint = drSpec.Item("Logo_List").ToString()
                            If Not drSpec.Item("Imp_Size_List").Equals(DBNull.Value) Then .Laser_Setup = drSpec.Item("Imp_Size_List").ToString()
                            If Not drSpec.Item("Loc_List").Equals(DBNull.Value) Then .Location = drSpec.Item("Loc_List").ToString()
                            If Not drSpec.Item("Num_Impr_1").Equals(DBNull.Value) Then .Num_Impr_1 = drSpec.Item("Num_Impr_1")
                            If Not drSpec.Item("Num_Impr_2").Equals(DBNull.Value) Then .Num_Impr_2 = drSpec.Item("Num_Impr_2")
                            If Not drSpec.Item("Num_Impr_3").Equals(DBNull.Value) Then .Num_Impr_3 = drSpec.Item("Num_Impr_3")
                            If Not drSpec.Item("Pack_List").Equals(DBNull.Value) Then .Packaging = drSpec.Item("Pack_List").ToString()
                            If Not drSpec.Item("Packer_Comment").Equals(DBNull.Value) Then .Packer_Comment = drSpec.Item("Packer_Comment").ToString()
                            If Not drSpec.Item("Packer_Instructions").Equals(DBNull.Value) Then .Packer_Instructions = drSpec.Item("Packer_Instructions").ToString()
                            If Not drSpec.Item("Printer_Comment").Equals(DBNull.Value) Then .Printer_Comment = drSpec.Item("Printer_Comment").ToString()
                            If Not drSpec.Item("Printer_Instructions").Equals(DBNull.Value) Then .Printer_Instructions = drSpec.Item("Printer_Instructions").ToString()
                            If Not drSpec.Item("Refill_Color").Equals(DBNull.Value) Then .Refill = drSpec.Item("Refill_Color").ToString()
                            If Not drSpec.Item("Comment").Equals(DBNull.Value) Then .Comments = drSpec.Item("Comment").ToString()
                            If Not drSpec.Item("Comment2").Equals(DBNull.Value) Then .Special_Comments = drSpec.Item("Comment2").ToString()

                            ' IF EXACT QTY, WE SET FLAG IN SPECIAL COMMENTS.
                            If blnExact_Qty Then

                                ' IF THERE IS SOMETHING IN SPECIAL COMMENTS, ADVISE THAT WE OVERWRITE.
                                If .Special_Comments <> String.Empty Then
                                    If Not blnExact_Qty_Flag Then
                                        blnExact_Qty_Flag = True
                                        MsgBox("Import warning: EXACT QTY flag overwrites data '" & .Special_Comments.Trim() & "' in Special comments field.")
                                    End If

                                End If

                                .Special_Comments = "EXACT QTY"

                            End If

                        End With

                        If Not oOrdline.ImprintLine.IsEmpty() Then oOrdline.ImprintLine.Save()
                        oOrdline.Save()

                        ' insert into exact_traveler_document (ord_no, headerid, doctype, docdescription, docname, document, createdate, artworkpurged, doctypeid, ord_guid, item_guid, docfile, ondrive)
                        ' values (
                        ' '', 0, 'P.O.', 'test', 'test.txt', 'THIS IS A TEST', GETDATE(), 0, 0, NULL, NULL, NULL, 0)

                        ' SELECT	CAST(CAST(Document AS VARBINARY(MAX)) AS VARCHAR(MAX)) AS OE_PO_Data
                        ' FROM	Exact_Traveler_Document
                        ' WHERE	DocID = 2132311

                    End If

                Next oOrdline

            End With

            Ordhead.Contacts.SetOrderContacts(Ordhead)

            CopyDocumentsFromProgramProfile(oNode)

            oOrderEdi.Ord_Guid = Ordhead.Ord_GUID
            oOrderEdi.OEI_Status = "I" ' IMPORTED
            oOrderEdi.Save()

            ' insert into exact_traveler_document (ord_no, headerid, doctype, docdescription, docname, document, createdate, artworkpurged, doctypeid, ord_guid, item_guid, docfile, ondrive)
            ' values (
            ' '', 0, 'P.O.', 'test', 'test.txt', 'THIS IS A TEST', GETDATE(), 0, 0, NULL, NULL, NULL, 0)

        Catch oe_er As OEException
            Throw oe_er ' New OEException(oe_er.Message, oe_er)
        Catch er As Exception
            If Mid(er.Message, 1, 7) = EdiFailureCd Then
                Call sendEDIAbortedEmail(er.Message)
            Else
                MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End If
        End Try

    End Sub

    Private Sub CopyDocumentsFromProgramProfile(ByRef poNode As XmlNode)

        Try

            Dim db As New cDBA
            Dim dtLines As DataTable
            Dim strSql As String

            Dim strCus_Dec_ID_List As String = ""
            Dim oLogo As cMDB_Logo
            Dim oDocument As CDocument

            For Each oLineNode As XmlNode In poNode.SelectNodes("OrderLines/OrderLine")

                strSql = "" & _
                "SELECT     DISTINCT C.CUS_DEC_ID " & _
                "FROM       MDB_CUS_DEC C WITH (NOLOCK) " & _
                "WHERE      SHORT_CTL_NO ='" & oLineNode.SelectSingleNode("item_no").InnerText.Replace("'", "''") & "' "

                dtLines = db.DataTable(strSql)
                If dtLines.Rows.Count <> 0 Then
                    For Each dtRow As DataRow In dtLines.Rows
                        If strCus_Dec_ID_List <> "" Then strCus_Dec_ID_List += ", "
                        strCus_Dec_ID_List += dtRow.Item("Cus_Dec_ID").ToString().Trim()
                    Next dtRow
                End If

            Next oLineNode

            If strCus_Dec_ID_List = String.Empty Then Exit Sub

            strSql = "" & _
            "SELECT     DISTINCT D.Logo_ID, L.Logo_Name, L.File_Url " & _
            "FROM       MDB_CUS_DEC_DETAIL D WITH (NOLOCK) " & _
            "INNER JOIN OEI_ORDLIN O WITH (NOLOCK) ON D.COMP_ITEM_NO = O.ITEM_NO AND O.ORD_GUID = '" & Ordhead.Ord_GUID & "' " & _
            "INNER JOIN MDB_LOGO L WITH (NoLock) ON D.Logo_ID = L.Logo_ID " & _
            "WHERE      D.Cus_Dec_ID IN (" & strCus_Dec_ID_List & ") "

            dtLines = db.DataTable(strSql)
            If dtLines.Rows.Count <> 0 Then
                For Each dtRow As DataRow In dtLines.Rows
                    'Dim oLogo As cMDB_Logo
                    oLogo = New cMDB_Logo(dtRow.Item("Logo_ID"))
                    If oLogo.File_Url <> "" Then
                        oDocument = New CDocument

                        oDocument.Ord_Guid = Ordhead.Ord_GUID
                        oDocument.DocTypeID = 1
                        oDocument.DocType = "Final Artwork"

                        oDocument.DocDescription = oLogo.Logo_Name
                        oDocument.DocFile = oLogo.File_Url
                        oDocument.Document = oLogo.Logo_Doc
                        oDocument.Document_Array = oLogo.Logo_Doc_Array
                        oDocument.AutoSave = True
                        oDocument.Save()
                    End If
                Next dtRow

            End If

            strSql = "" & _
"SELECT     ISNULL(PO_Instructions, '') AS PO_Instructions " & _
"FROM       MDB_CUS_DEC D WITH (NOLOCK) " & _
"WHERE      D.Cus_Dec_ID IN (" & strCus_Dec_ID_List & ") "

            Dim strPO_Instructions As String = String.Empty

            dtLines = db.DataTable(strSql)
            If dtLines.Rows.Count <> 0 Then

                strPO_Instructions = _
                SplitLargeSentence("THIS PURCHASE ORDER IS AN AUTOMATED PURCHASE ORDER CREATED BY OUR EDI PROCESS. IF THERE IS ANY PROBLEM, CUSTOMER SERVICE MUST BE NOTIFIED. BECAUSE THIS IS AN AUTOMATED PROCESS, DATA MAY BE INCORRECT OR MAY BECOME INCORRECT WITH TIME. PLEASE ADVISE OF ANY PROBLEM.", 80, vbCrLf) & vbCrLf

                For Each dtRow As DataRow In dtLines.Rows

                    strPO_Instructions = strPO_Instructions & SplitLargeSentence(dtRow.Item("PO_Instructions").ToString, 80, vbCrLf)

                Next dtRow

            End If

            If strPO_Instructions.Length > 0 Then

                strSql = _
                "INSERT INTO EXACT_TRAVELER_DOCUMENT (ord_no, headerid, doctype, docdescription, docname, document, createdate, artworkpurged, doctypeid, ord_guid, item_guid, docfile, ondrive) " & _
                "VALUES " & _
                "('', 0, 'Customer P.O.', 'Automated EDI Customer PO', 'Automated EDI Customer PO.txt', '" & strPO_Instructions.Replace("'", "''") & "', GETDATE(), 0, 10, '" & Ordhead.Ord_GUID & "', NULL, NULL, 2) "

                db.Execute(strSql)

            End If

            'strSql = _
            '            "SELECT     DISTINCT C.CUS_DEC_ID, C.SHORT_CTL_NO, D.KIT_ITEM_NO " & _
            '            "FROM       MDB_CUS_DEC C WITH (NOLOCK) " & _
            '            "INNER JOIN MDB_CUS_DEC_DETAIL D WITH (NOLOCK) ON C.CUS_DEC_ID = D.CUS_DEC_ID " & _
            '            "WHERE      SHORT_CTL_NO = '" & oLineNode.SelectSingleNode("item_no").InnerText.Replace("'", "''") & "' AND " & _
            '            "           C.ASSORTED = 1 AND D.SELECTED = 1 "

            ''    dtLines = db.DataTable(strSql)
            ''    Select Case dtLines.Rows.Count
            ''        Case 0
            ''            ' Not an assorted
            ''            strSql = _
            ''            "SELECT     DISTINCT C.CUS_DEC_ID, C.SHORT_CTL_NO, D.KIT_ITEM_NO " & _
            ''            "FROM       MDB_CUS_DEC C WITH (NOLOCK) " & _
            ''            "INNER JOIN MDB_CUS_DEC_DETAIL D WITH (NOLOCK) ON C.CUS_DEC_ID = D.CUS_DEC_ID " & _
            ''            "WHERE SHORT_CTL_NO = '" & oLineNode.SelectSingleNode("item_no").InnerText.Replace("'", "''") & "' "
            ''            dtLines = db.DataTable(strSql)

            ''            If IsNumeric(oLineNode.SelectSingleNode("qty_ordered").InnerText) Then
            ''                iQtyOrdered = Integer.Parse(oLineNode.SelectSingleNode("qty_ordered").InnerText)
            ''            End If
            ''            iQtyRemainAssorted = 0

            ''        Case 1
            ''            ' An assorted with only 1 selected.
            ''            If IsNumeric(oLineNode.SelectSingleNode("qty_ordered").InnerText) Then
            ''                iQtyOrdered = Integer.Parse(oLineNode.SelectSingleNode("qty_ordered").InnerText)
            ''            End If
            ''            iQtyRemainAssorted = 0

            ''        Case Else
            ''            ' An assorted with 2 or more elements
            ''            If IsNumeric(oLineNode.SelectSingleNode("qty_ordered").InnerText) Then
            ''                iQtyOrdered = Integer.Parse(oLineNode.SelectSingleNode("qty_ordered").InnerText)
            ''                iQtyRemainAssorted = iQtyOrdered Mod dtLines.Rows.Count
            ''                iQtyOrdered = iQtyOrdered \ dtLines.Rows.Count
            ''            End If

            ''    End Select

            ''    For Each dtLineRow In dtLines.Rows

            ''        Dim oOrdline As New cOrdLine(Ordhead.Ord_GUID)
            ''        'oOrdline.Source = OEI_SourceEnum.ExternalFile
            ''        'oOrdline.Ord_Guid = Ordhead.Ord_GUID

            ''        oOrdline.Line_Seq_No = WorkLine
            ''        'oOrdline.OEOrdLin_ID = 1 ' Integer.Parse(oNode.SelectSingleNode("line_no").InnerText)
            ''        'oOrdline.SaveToDB = True
            ''        ' HERE WE MUST GET ITEM NO FROM ITEM LINK TO MDB_CUS_DEC
            ''        oOrdline.Extra_11 = dtLineRow.Item("Cus_Dec_ID")
            ''        oOrdline.Item_No = dtLineRow.Item("Kit_Item_No").ToString()

            ''        ' We put in field Extra_11 the Cus_Dec_ID we will use when filling the imprint and the traveler

            ''        ' Location will be default for item
            ''        ' oOrdline.Loc = oNode.SelectSingleNode("loc").InnerText

            ''        oOrdline.Cus_Item_No = oLineNode.SelectSingleNode("cus_item_no").InnerText.ToString()

            ''        ' Qty to ship is set to qty ordered, because of overpicks.
            ''        Try
            ''            oOrdline.Qty_Ordered = iQtyOrdered
            ''            If iQtyRemainAssorted > 0 Then
            ''                oOrdline.Qty_Ordered += 1
            ''                iQtyRemainAssorted -= 1
            ''            End If
            ''        Catch oe_er As OEException
            ''            'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
            ''        End Try

            ''        Try
            ''            oOrdline.Qty_To_Ship = oOrdline.Qty_Ordered
            ''            oOrdline.Qty_To_Ship_Changed(oOrdline.Item_No, oOrdline.Loc, False)
            ''            'If oOrdline.Qty_Prev_Bkord > 0 Then
            ''            '    oOrdline.Calculate_ETA()
            ''            'End If
            ''        Catch oe_er As OEException
            ''            'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
            ''        End Try

            ''        ' Do not enter unit price, let calculate new price on date
            ''        'oOrdline.Unit_Price = Integer.Parse(oNode.SelectSingleNode("unit_price").InnerText)

            ''        'oOrdline.Discount_Pct = dtRow.Item("Discount_Pct")
            ''        'oOrdline.Request_Dt = Now.Date
            ''        'oOrdline.Promise_Dt = Now.Date

            ''        If Not oLineNode.SelectSingleNode("req_ship_date") Is Nothing Then
            ''            If IsDate(oLineNode.SelectSingleNode("req_ship_date").InnerText) Then
            ''                oOrdline.Req_Ship_Dt = Date.Parse(oLineNode.SelectSingleNode("req_ship_date").InnerText)
            ''            End If
            ''        End If

            ''        ''oOrdline.Source = OEI_SourceEnum.frmOrder
            ''        oOrdline.Save()

            ''        'AddExternalLine(oOrdline)
            ''        AddOrderLine(oOrdline)

            ''    Next dtLineRow

            'Next oLineNode

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

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

        If pstrSentence.Trim.Length > 0 Then
            SplitLargeSentence = SplitLargeSentence & pstrSentence & pstrSplitChar
        End If

    End Function

    '    Public Sub SetPickingQty(ByRef po_Row As DataGridViewRow, ByVal pdblQty_Orderred As Double)

    '        Dim strItem_Guid As String = po_Row.Cells("Item_Guid").ToString

    '        Dim strSql As String
    '        Dim db As New cDBA
    '        Dim dt As DataTable

    '        If po_Row.Cells("End_Item_Cd").ToString = "K" Then

    '            ' Is a kit, then we must fill each lines separately
    '            strSql = "SELECT BLA BLA BLA"

    '            dt = db.DataTable(strSql)
    '            If dt.Rows.Count <> 0 Then

    '            End If

    '        Else

    '            ' Is not a kit, we can do it straight in that line !
    '            strSql = "SELECT BLA BLA BLA"

    '            dt = db.DataTable(strSql)
    '            If dt.Rows.Count <> 0 Then

    '            End If

    '        End If

    '    End Sub

    Public Function OneShotOrderUsed(ByVal pProg_Spector_Cd As String) As String

        OneShotOrderUsed = String.Empty

        Dim db As New cDBA()
        Dim dt As DataTable
        Dim strSql As String = _
        "SELECT     CASE WHEN ISNULL(O.ORD_NO, '') = '' THEN O.OEI_ORD_NO ELSE O.ORD_NO END AS ORD_NO " & _
        "FROM       OEI_ORDHDR O WITH (Nolock) " & _
        "INNER JOIN MDB_CUS_PROG P WITH (Nolock) ON O.Prog_Spector_Cd = P.Spector_Cd " & _
        "WHERE      O.PROG_SPECTOR_CD = '" & pProg_Spector_Cd.Replace("'", "''") & "' " & _
        "       AND ISNULL(P.ONE_SHOT, 0) = 1 "

        dt = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            OneShotOrderUsed = dt.Rows(0).Item("Ord_No").ToString
        End If

    End Function

    Public Function CheckSetupQty(ByVal pstrItem_No As String, pdblQty As Double) As Boolean

        CheckSetupQty = False

        Dim db As New cDBA()
        Dim dt As DataTable

        Dim strSql As String = _
        "SELECT ITEM_NO, EXTRA_14 " & _
        "FROM   imitmidx_sql " & _
        "WHERE  item_no = '" & pstrItem_No.Trim & "' And Extra_14 = 27 "

        dt = db.DataTable(strSql)

        If dt.Rows.Count <> 0 And pdblQty >= 5 Then

            Dim ombResult As New MsgBoxResult
            ombResult = MsgBox("Are you certain the order requires this many setups ? ", MsgBoxStyle.YesNo)

            If ombResult = MsgBoxResult.Yes Then
                CheckSetupQty = True
            End If

        Else
            CheckSetupQty = True
        End If

    End Function

    Public Sub SetHoldFgFromAvailCredit()

        Try
            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String = "SELECT * FROM OEI_CONFIG WITH (Nolock) WHERE Check_Avail_Credit = 1 "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                If m_oOrder.Ordhead.Tot_Sls_Amt > m_oOrder.CreditInfo.Amt_Avail_Credit Then
                    m_oOrder.Ordhead.Hold_Fg = "H"
                Else
                    m_oOrder.Ordhead.Hold_Fg = ""
                End If
            End If

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function Is_White_Flag() As Boolean

        Is_White_Flag = False

        If Ordhead.Cus_No Is Nothing Then Exit Function

        Try
            Dim db As New cDBA
            Dim dt As DataTable

            Dim strSql As String = _
            "SELECT * " & _
            "FROM	MDB_CUSTOMER WITH (NOLOCK) " & _
            "WHERE	CUS_NO = '" & Ordhead.Cus_No.Trim & "' AND " & _
            "   WHITE_GLOVE_END_DATE >= CAST(GETDATE() as DATE) "
            '"   WHITE_GLOVE_END_DATE >= CAST(GETDATE() as DATE) AND ISNULL(WHITE_GLOVE_ORDER_COUNT, 0) > 0 "

            dt = db.DataTable(strSql)

            Is_White_Flag = (dt.Rows.Count <> 0)
            m_oOrder.Ordhead.White_Glove = Is_White_Flag

        Catch er As Exception
            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function






    Public Sub XML_Export_Old(no As String)
        Try
            Dim sql As String
            Dim dt As New DataTable
            Dim db As New cDBA
            Dim strPath As String
            Dim strSql As String = "select * from OEI_CONFIG"
            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                strPath = Trim(dt.Rows(0).Item("XML_Export_File_Path"))
                'strPath = "C:\"
                sql = "select	HEADER.OEI_ORD_NO, HEADER.OE_PO_NO, HEADER.CUS_NO, HEADER.Bill_To_Name, " & _
                               "HEADER.Bill_To_Addr_1, HEADER.Bill_To_Addr_2, HEADER.Bill_To_Addr_3, " & _
                               "HEADER.Bill_To_Addr_4, HEADER.Bill_To_Country, HEADER.Cus_Alt_Adr_Cd, " & _
                               "HEADER.Ship_To_Name, HEADER.Ship_To_Addr_1, HEADER.Ship_To_Addr_2," & _
                               "HEADER.Ship_To_Addr_3, HEADER.Ship_To_Addr_4, HEADER.Ship_To_Country," & _
                               "HEADER.Shipping_Dt, HEADER.Ship_Via_Cd, HEADER.Ar_Terms_Cd, HEADER.Ship_Instruction_1, " & _
                               "HEADER.Ship_Instruction_2, HEADER.Slspsn_No, HEADER.Discount_Pct, HEADER.Job_No, " & _
                               "HEADER.Mfg_Loc, HEADER.Profit_Center, HEADER.Curr_Cd, HEADER.Contact_1, " & _
                               "HEADER.Phone_Number, HEADER.Fax_Number, HEADER.Email_Address, HEADER.Ship_To_City, " & _
                               "HEADER.Ship_To_State, HEADER.Ship_To_Zip, HEADER.Bill_To_City, HEADER.Bill_To_State, " & _
                               "HEADER.Bill_To_Zip, HEADER.Tax_Cd, HEADER.User_Def_Fld_1 AS Urgent, " & _
                               "HEADER.User_Def_Fld_2 AS Order_Type, HEADER.User_Def_Fld_3 AS CSR, " & _
                               "HEADER.User_Def_Fld_4 AS Cus_Contact, HEADER.User_Def_Fld_5 AS Rep, " & _
                               "LINE.Item_No, LINE.Loc, LINE.Qty_Ordered, LINE.Qty_To_Ship, LINE.Unit_Price, " & _
                               "LINE.Discount_Pct, LINE.Request_Dt, LINE.Promise_Dt, LINE.Req_Ship_Dt, " & _
                               "TRAVELER.RouteID, TRAVELER.RouteDescription, IMPRINT.Extra_1 AS Imprint, " & _
                               "IMPRINT.Extra_2 AS Location, IMPRINT.Extra_3 AS Color, IMPRINT.Extra_4 AS Num, " & _
                               "IMPRINT.Extra_5 AS Packaging, IMPRINT.Extra_8 AS Refill, IMPRINT.Extra_9 AS Laser_Setup, " & _
                               "IMPRINT.Comment AS Comments, IMPRINT.Comment2 AS Special_Comments, " & _
                               "IMPRINT.Industry, IMPRINT.Printer_Comment, IMPRINT.Packer_Comment, " & _
                               "IMPRINT.Printer_Instructions, IMPRINT.Packer_Instructions, " & _
                               "IMPRINT.Num_Impr_1, IMPRINT.Num_Impr_2, IMPRINT.Num_Impr_3 " & _
                               "FROM OEI_ORDHDR HEADER WITH (NOLOCK) " & _
                               "INNER JOIN	OEI_ORDLIN LINE WITH (NOLOCK) ON HEADER.ORD_GUID = LINE.ORD_GUID " & _
                               "LEFT JOIN	EXACT_TRAVELER_ROUTE TRAVELER WITH (NOLOCK) ON LINE.ROUTEID = TRAVELER.ROUTEID " & _
                               "LEFT JOIN	OEI_EXTRA_FIELDS IMPRINT WITH (NOLOCK) ON LINE.ORD_GUID = IMPRINT.ORD_GUID AND LINE.ITEM_GUID = IMPRINT.ITEM_GUID " & _
                               "WHERE HEADER.OEI_ORD_NO = '" & no.Trim & "'" & _
                               "FOR XML AUTO"

                strPath = Trim(strPath) + "OEI_ORD_" + Trim(no) + ".xml"

                Dim dtXml As DataTable
                dtXml = db.DataTable(sql)
                If dtXml.Rows.Count <> 0 Then
                    Using write As New XmlTextWriter(strPath, System.Text.Encoding.UTF8)
                        write.WriteStartDocument()
                        write.WriteStartElement(dtXml.Rows(0).Item(0).ToString)
                        ' End document.
                        write.WriteEndElement()
                        write.WriteEndDocument()
                    End Using
                    'MsgBox("Done")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub XML_Export(no As String)
        Try
            Dim sql As String
            Dim dt As New DataTable
            Dim db As New cDBA
            Dim strPath As String
            Dim strSql As String = "select * from OEI_CONFIG"
            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                strPath = Trim(dt.Rows(0).Item("XML_Export_File_Path"))
                strPath = "C:\"
                sql = "select	HEADER.OEI_ORD_NO, HEADER.OE_PO_NO, HEADER.CUS_NO, HEADER.Bill_To_Name, " & _
                               "HEADER.Bill_To_Addr_1, HEADER.Bill_To_Addr_2, HEADER.Bill_To_Addr_3, " & _
                               "HEADER.Bill_To_Addr_4, HEADER.Bill_To_Country, HEADER.Cus_Alt_Adr_Cd, " & _
                               "HEADER.Ship_To_Name, HEADER.Ship_To_Addr_1, HEADER.Ship_To_Addr_2," & _
                               "HEADER.Ship_To_Addr_3, HEADER.Ship_To_Addr_4, HEADER.Ship_To_Country," & _
                               "HEADER.Shipping_Dt, HEADER.Ship_Via_Cd, HEADER.Ar_Terms_Cd, HEADER.Ship_Instruction_1, " & _
                               "HEADER.Ship_Instruction_2, HEADER.Slspsn_No, HEADER.Discount_Pct, HEADER.Job_No, " & _
                               "HEADER.Mfg_Loc, HEADER.Profit_Center, HEADER.Curr_Cd, HEADER.Contact_1, " & _
                               "HEADER.Phone_Number, HEADER.Fax_Number, HEADER.Email_Address, HEADER.Ship_To_City, " & _
                               "HEADER.Ship_To_State, HEADER.Ship_To_Zip, HEADER.Bill_To_City, HEADER.Bill_To_State, " & _
                               "HEADER.Bill_To_Zip, HEADER.Tax_Cd, HEADER.User_Def_Fld_1 AS Urgent, " & _
                               "HEADER.User_Def_Fld_2 AS Order_Type, HEADER.User_Def_Fld_3 AS CSR, " & _
                               "HEADER.User_Def_Fld_4 AS Cus_Contact, HEADER.User_Def_Fld_5 AS Rep, " & _
                               "LINE.Item_No, LINE.Loc, LINE.Qty_Ordered, LINE.Qty_To_Ship, LINE.Unit_Price, " & _
                               "LINE.Discount_Pct, LINE.Request_Dt, LINE.Promise_Dt, LINE.Req_Ship_Dt, " & _
                               "TRAVELER.RouteID, TRAVELER.RouteDescription, IMPRINT.Extra_1 AS Imprint, " & _
                               "IMPRINT.Extra_2 AS Location, IMPRINT.Extra_3 AS Color, IMPRINT.Extra_4 AS Num, " & _
                               "IMPRINT.Extra_5 AS Packaging, IMPRINT.Extra_8 AS Refill, IMPRINT.Extra_9 AS Laser_Setup, " & _
                               "IMPRINT.Comment AS Comments, IMPRINT.Comment2 AS Special_Comments, " & _
                               "IMPRINT.Industry, IMPRINT.Printer_Comment, IMPRINT.Packer_Comment, " & _
                               "IMPRINT.Printer_Instructions, IMPRINT.Packer_Instructions, " & _
                               "IMPRINT.Num_Impr_1, IMPRINT.Num_Impr_2, IMPRINT.Num_Impr_3 " & _
                               "FROM OEI_ORDHDR HEADER WITH (NOLOCK) " & _
                               "INNER JOIN	OEI_ORDLIN LINE WITH (NOLOCK) ON HEADER.ORD_GUID = LINE.ORD_GUID " & _
                               "LEFT JOIN	EXACT_TRAVELER_ROUTE TRAVELER WITH (NOLOCK) ON LINE.ROUTEID = TRAVELER.ROUTEID " & _
                               "LEFT JOIN	OEI_EXTRA_FIELDS IMPRINT WITH (NOLOCK) ON LINE.ORD_GUID = IMPRINT.ORD_GUID AND LINE.ITEM_GUID = IMPRINT.ITEM_GUID " & _
                               "WHERE HEADER.OEI_ORD_NO = '" & no.Trim & "'" & _
                               "FOR XML AUTO"

                'strPath = "C:\ExactTemp\"

                strPath = Trim(strPath) + "OEI_ORD_" + Trim(no) + ".xml"

                Dim dtXml As DataTable
                dtXml = db.DataTable(sql)
                If dtXml.Rows.Count <> 0 Then

                    '++ID I put in comment 28.1.15
                    'Using write As New XmlTextWriter(strPath, System.Text.Encoding.UTF8) ' UTF8

                    '    write.WriteStartDocument()

                    '    'write.WriteStartElement(dtXml.Rows(0).Item(0).ToString)
                    '    ' End document.

                    '    Dim strXml As String = ""
                    '    For Each dtRow As DataRow In dtXml.Rows
                    '        strXml &= dtRow.Item(0).ToString
                    '    Next

                    '    write.WriteString(strXml)
                    '    'write.WriteChars(strXml.ToCharArray, 0, strXml.Length - 1)
                    '    'write.WriteEndElement()
                    '    write.WriteEndDocument()

                    'End Using

                    'MsgBox("Done")

                    ''++ID 28.1.15

                    Dim strXml As String = ""

                    For i As Int32 = 0 To dtXml.Rows.Count - 1

                        strXml &= dtXml.Rows(i).Item(0)

                    Next

                    Using wr As StreamWriter = New StreamWriter(strPath)
                        wr.WriteLine("<?xml version=""1.0"" encoding=""utf-8""?>")
                        wr.WriteLine(strXml)

                    End Using

                    MsgBox("Done")

                End If

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class

Public Enum OEI_SourceEnum
    None        ' When initialized with reset.
    frmOrder
    frmProductLineEntry
    frmPLEntrySelectedItem
    Macola_Redo
    Macola_ExactRepeat
    OEI_Reload
    ExternalFile
End Enum

Public Enum OrderSourceEnum
    None        ' When initialized with reset.
    Macola
    MacolaHistory
    OEInterface
    WebImport
End Enum

Public Enum ChargesPricingType
    Standard
    Net
End Enum