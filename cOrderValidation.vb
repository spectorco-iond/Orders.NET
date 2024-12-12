Imports System.Text.RegularExpressions

Public Class cOrderValidation

    Private dt As DataTable
    Private db As New cDBA

    Public Function AccountingPeriodExist(ByVal year As String, ByVal per As String) As Boolean

        AccountingPeriodExist = True

        Try

            Dim strSql As String =
                "SELECT 	TOP 1 * " &
                "FROM 	    perdat WITH (Nolock) " &
                "WHERE      bkjrcode = " & year & " And " &
                "           per_fin  = " & per & ""

            '++ID 02.17.2023 not need 
            'If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
            '    strSql =
            '    "SELECT 	TOP 1 * " &
            '    "FROM 	    [200].dbo.perdat WITH (Nolock) " &
            '    "WHERE      bkjrcode = " & year & " And " &
            '    "           per_fin  = " & per & ""
            'End If




            dt = db.DataTable(strSql)

            AccountingPeriodExist = (dt.Rows.Count <> 0)

        Catch er As Exception
            MsgBox("Error in COrderValidation." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)

        End Try

    End Function

    Public Function CustomerAllowsSubst(ByVal pstrCus_No) As Boolean

        CustomerAllowsSubst = False

        Try

            Dim strSql As String = _
            "SELECT * " & _
            "FROM   ArCusFil_Sql WITH (Nolock) " & _
            "WHERE  Cus_No = '" & SqlCompliantString(pstrCus_No) & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                CustomerAllowsSubst = (dt.Rows(0).Item("Allow_SB_Item") = "Y")
            End If

        Catch er As Exception
            MsgBox("Error in COrderValidation." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function CustomerAllowsBO(ByVal pstrCus_No) As Boolean

        CustomerAllowsBO = False

        Try

            Dim strSql As String = _
            "SELECT * " & _
            "FROM   ArCusFil_Sql WITH (Nolock) " & _
            "WHERE  Cus_No = '" & SqlCompliantString(pstrCus_No) & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                CustomerAllowsBO = (dt.Rows(0).Item("Allow_BO") <> "Y")
            End If

        Catch er As Exception
            MsgBox("Error in COrderValidation." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function DateIsOutsideOfValidTrxDates(ByVal year As String, ByVal per As String, ByVal bknr As String) As Boolean

        DateIsOutsideOfValidTrxDates = True

        Try
            Dim strSql As String = _
            "SELECT 	TOP 1 * " & _
            "FROM 	    AfgPer WITH (Nolock) " & _
            "WHERE      bkjrcode = " & year & " And " & _
            "           Periode  = " & per & " And " & _
            "           dagbknr  = " & bknr & ""

            dt = db.DataTable(strSql)

            DateIsOutsideOfValidTrxDates = (dt.Rows.Count <> 0)

        Catch er As Exception
            MsgBox("Error in COrderValidation." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)

        End Try

    End Function

    Public Function GetPeriodDates(ByVal year As String, ByVal per As String) As String

        GetPeriodDates = ""

        Try
            Dim strSql As String = _
            "SELECT 	TOP 1 * " & _
            "FROM 	    perdat WITH (Nolock) " & _
            "WHERE      bkjrcode = " & year & " And " & _
            "           per_fin  = " & per & ""

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                GetPeriodDates = CDate(dt.Rows(0).Item("bgdatum")).ToShortDateString & "-" & CDate(dt.Rows(0).Item("eddatum")).ToShortDateString
            End If

        Catch er As Exception
            MsgBox("Error in COrderValidation." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)

        End Try

    End Function

    Public Function IsValidEmail(ByRef strEmail As String) As Boolean

        IsValidEmail = False

        Try

            Dim pattern As String = "[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
            Dim m As Match = Regex.Match(strEmail, pattern, RegexOptions.IgnoreCase)

            If m.Success Then
                IsValidEmail = True
            End If

        Catch er As Exception
            MsgBox("Error in COrderValidation." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function IsValidPhoneNumber(ByRef TelNo As String) As Boolean

        IsValidPhoneNumber = False

        Try
            Dim strTemp As String
            Dim strPhone As String
            Dim strExtension As String = ""
            Dim intResult As Short

            ' Remove all the grouping characters for now. We'll add them back in later.
            strTemp = Replace(TelNo, "(", "")
            strTemp = Replace(strTemp, ")", "")
            strTemp = Replace(strTemp, "-", "")
            strTemp = Replace(strTemp, " ", "")
            strTemp = Replace(strTemp, "X", "x")

            ' Break up the digits into the number and the extension, if any.
            intResult = InStr(1, strTemp, "x", CompareMethod.Text)
            If intResult > 0 Then
                strExtension = Mid(strTemp, intResult + 1)
                strPhone = Left(strTemp, intResult - 1)
            Else
                strPhone = strTemp
            End If

            If Left(strPhone, 1) = "1" Then
                strPhone = Mid(strPhone, 2)
            End If

            If Len(strPhone) <> 10 Then Exit Function

            ' Build the new phone number
            TelNo = "(" & Left(strPhone, 3) & ") " & Mid(strPhone, 4, 3) & "-" & Right(strPhone, 4)

            ' Add the extension, if any
            If strExtension <> "" Then
                TelNo = TelNo & " x" & strExtension
            End If

            IsValidPhoneNumber = True

        Catch er As Exception
            MsgBox("Error in COrderValidation." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function ItemIsActive(ByVal pstrItem_No As String) As Boolean

        ItemIsActive = True

        Try
            Dim strSql As String = _
            "SELECT * " & _
            "FROM   ImItmIdx_Sql WITH (Nolock) " & _
            "WHERE  Item_No = '" & SqlCompliantString(pstrItem_No) & "' "

            '++ID 02.17.2023
            If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                strSql =
            "SELECT * " &
            "FROM   [200].dbo.ImItmIdx_Sql WITH (Nolock) " &
            "WHERE  Item_No = '" & SqlCompliantString(pstrItem_No) & "' "
            End If


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                ItemIsActive = (dt.Rows(0).Item("Activity_Cd") = "A")
            End If

        Catch er As Exception
            MsgBox("Error in COrderValidation." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    'Public Sub OrderExists2(ByVal pstrOrd_No As String)

    '    'OrderExists = False

    '    Dim strSql As String = _
    '    "SELECT Ord_No " & _
    '    "FROM   OEOrdHdr_Sql WITH (Nolock) " & _
    '    "WHERE  LTRIM(RTRIM(ISNULL(Ord_No, ''))) = '" & SqlCompliantString(pstrOrd_No) & "' " & _
    '    "UNION  " & _
    '    "SELECT Ord_No " & _
    '    "FROM   OEI_ORDHDR WITH (Nolock) " & _
    '    "WHERE  LTRIM(RTRIM(ISNULL(Ord_No, ''))) = '" & SqlCompliantString(pstrOrd_No) & "' "

    '    Dim db As New cDBA
    '    Dim dt As New DataTable
    '    dt = db.DataTable(strSql)

    '    'OrderExists = (dt.Rows.Count <> 0)

    'End Sub

    Public Sub OrderExists(ByVal pstrOrd_No As String)

        If Trim(pstrOrd_No) = "" Then Exit Sub

        Dim strSql As String = _
        "SELECT * " & _
        "FROM   OEOrdHdr_Sql WITH (Nolock) " & _
        "WHERE  LTRIM(RTRIM(ISNULL(Ord_No, ''))) = '" & SqlCompliantString(pstrOrd_No) & "' "

        dt = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Select Case dt.Rows(0).Item("Ord_Type")
                Case "O"
                    Throw New OEException(OEError.Order_No_From_Macola_Entry_Not_Allowed, True, True)
                Case "I"
                    Throw New OEException(OEError.OrderTypeInvoiceIsActive, True, True)
                Case Else
                    Throw New OEException(OEError.Order_No_Found_In_History_Files_Entry_Not_Allowed, True, True)
            End Select
        End If

        ' Afficher ce message pour tous les types. Le seul qui est different est pour les
        ' type 'I', le message est: Order already invoiced. Access denied.
        'OrderExistsInHistory = (dt.Rows.Count <> 0)

        strSql = _
        "SELECT * " & _
        "FROM   OEHdrHst_Sql WITH (Nolock) " & _
        "WHERE  LTRIM(RTRIM(ISNULL(Ord_No, ''))) = '" & SqlCompliantString(pstrOrd_No) & "' "

        dt = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Throw New OEException(OEError.Order_No_Found_In_History_Files_Entry_Not_Allowed, True, True)
        End If
        ' Afficher ce message pour tous les types. Le seul qui est different est pour les
        ' type 'I', le message est: Order already invoiced. Access denied.
        'OrderExistsInHistory = (dt.Rows.Count <> 0)

        strSql = _
        "SELECT ISNULL(Ord_No, '') AS Ord_No " & _
        "FROM   OEI_ORDHDR WITH (Nolock) " & _
        "WHERE  LTRIM(RTRIM(ISNULL(OEI_Ord_No, ''))) = '" & SqlCompliantString(pstrOrd_No) & "' "

        dt = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then

            ' If Ord_No in OEI_ORDHDR, then order from OEI is already in Macola.
            If Trim(dt.Rows(0).Item("Ord_No").ToString) <> "" Then
                Throw New OEException(OEError.Order_No_Found_In_OEI_Exported_And_Processed, _
                                      vbCrLf & "Macola order number: " & Trim(dt.Rows(0).Item("Ord_No").ToString), True, True)
            Else

                strSql = _
                "SELECT *   " & _
                "FROM   OEI_ORDHDR " & _
                "WHERE  OEI_Ord_No = '" & SqlCompliantString(pstrOrd_No) & "' AND " & _
                "       ExportTS IS NOT NULL "

                dt = db.DataTable(strSql)

                If dt.Rows.Count <> 0 Then
                    Throw New OEException(OEError.Order_No_Found_In_OEI_Exported_Not_Processed, True, True)
                End If

                strSql = _
                "SELECT *   " & _
                "FROM   OEI_ORDHDR " & _
                "WHERE  OEI_Ord_No = '" & SqlCompliantString(pstrOrd_No) & "' AND " & _
                "       OpenTS IS NOT NULL "

                dt = db.DataTable(strSql)

                If dt.Rows.Count <> 0 Then
                    Throw New OEException(OEError.Order_No_In_Use, _
                              vbCrLf & "TS: " & Trim(dt.Rows(0).Item("OpenTS").ToString), True, True)
                Else
                    Throw New OEException(OEError.Order_No_Found_In_OEI, False, False)
                End If

            End If

        End If

        'If g_User.OrderNumberGenerated() = pstrOrd_No Then
        If g_User.OrderNumberGenerated() = pstrOrd_No Then
            Throw New OEException(OEError.Order_No_Is_New, False, False)
        Else
            Throw New OEException(OEError.Order_No_Not_Found, True, True)
        End If

    End Sub

    ' Checks if an order with the same PO Number exists in Macola and OEI
    Public Function POExists(ByVal pstrCus_No As String, ByVal pstrPO_No As String, Optional ByVal suppressMsg As Boolean = False) As Boolean

        POExists = False

        pstrCus_No = pstrCus_No.Trim
        pstrPO_No = pstrPO_No.Trim

        'Determine all accounts
        Dim accounts As String = "'" & pstrCus_No & "'"

        Dim strSql As String = "select c.cmp_code from cicmpy as c " & _
                                " INNER JOIN ( " & _
                                    " select c2.cmp_code, c2.cmp_wwn " & _
                                    " from cicmpy as c1 " & _
                                    " INNER JOIN cicmpy as c2 " & _
                                    "   ON c1.cmp_parent = c2.cmp_wwn " & _
                                    " where c1.cmp_code = '" & SqlCompliantString(pstrCus_No) & "' " & _
                                 " ) as tbl " & _
                                    " ON tbl.cmp_code = c.cmp_code or tbl.cmp_wwn = c.cmp_parent "

        dt = db.DataTable(strSql)
        If dt.Rows.Count <> 0 Then
            For Each row As DataRow In dt.Rows
                accounts = accounts + "," + "'" & row.Item("cmp_code").ToString().Trim() & "'"
            Next
        End If

        strSql = _
        "SELECT Ord_No, OE_PO_No " & _
        "FROM   OEOrdHdr_Sql WITH (Nolock) " & _
        "WHERE  LTRIM(RTRIM(ISNULL(Cus_No, ''))) IN (" & accounts & ") AND LTRIM(RTRIM(ISNULL(OE_PO_No, ''))) = '" & SqlCompliantString(pstrPO_No) & "' " & _
        "UNION ALL " & _
        "SELECT CASE WHEN ISNULL(Ord_No, '') = '' THEN OEI_Ord_No ELSE Ord_No END AS Ord_No, OE_PO_No " & _
        "FROM   OEI_ORDHDR WITH (Nolock) " & _
        "WHERE  LTRIM(RTRIM(ISNULL(Cus_No, ''))) IN (" & accounts & ") AND LTRIM(RTRIM(ISNULL(OE_PO_No, ''))) = '" & SqlCompliantString(pstrPO_No) & "' " & _
        "ORDER BY ord_no"

        dt = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            'Throw New OEException(OEError.Cust_PO_Exists)
            POExists = True

            If suppressMsg = False Then
                Throw New OEException(OEError.Cust_PO_Exists, vbCrLf & " Order: " & dt.Rows(0).Item("Ord_No").ToString, False, True)
            End If
        End If

        If POExists = False Then
            strSql = _
            "SELECT Ord_No, OE_PO_No " & _
            "FROM   OEHdrHst_Sql WITH (Nolock) " & _
            "WHERE  LTRIM(RTRIM(ISNULL(Cus_No, ''))) IN (" & accounts & ") AND LTRIM(RTRIM(ISNULL(OE_PO_No, ''))) = '" & SqlCompliantString(pstrPO_No) & "' "

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                POExists = True
                'Throw New OEException(OEError.Cust_PO_Exists_In_History)
                If suppressMsg = False Then
                    Throw New OEException(OEError.Cust_PO_Exists_In_History, vbCrLf & " Order: " & dt.Rows(0).Item("Ord_No").ToString, False, True)
                End If
            End If
        End If

        Return POExists
    End Function

    'Public Function POExistsInHistory(ByVal pstrCus_No As String, ByVal pstrPO_No As String) As Boolean

    '    POExistsInHistory = False

    '    pstrCus_No = pstrCus_No.Trim
    '    pstrPO_No = pstrPO_No.Trim

    '    Dim strSql As String = _
    '    "SELECT OE_PO_No " & _
    '    "FROM   OEHdrHst_Sql WITH (Nolock) " & _
    '    "WHERE  LTRIM(RTRIM(ISNULL(Cus_No, ''))) = '" & SqlCompliantString(pstrCus_No) & "' AND LTRIM(RTRIM(ISNULL(OE_PO_No, ''))) = '" & SqlCompliantString(pstrPO_No) & "' "

    '    Dim db As New cDBA
    '    Dim dt As New DataTable
    '    dt = db.DataTable(strSql)
    '    POExistsInHistory = (dt.Rows.Count <> 0)

    'End Function

    Public Function OrderOnHoldPermission() As PermissionType

        OrderOnHoldPermission = PermissionType.Neither

        Dim strSql As String = _
        "SELECT ISNULL(WARN_PREV, '') AS WARN_PREV " & _
        "FROM   SYSFIELD_SQL WITH (Nolock) " & _
        "WHERE  Package = 'XX' AND Description = 'O/E Order On Hold' "

        dt = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Select Case dt.Rows(0).Item("WARN_PREV")
                Case "W"
                    OrderOnHoldPermission = PermissionType.Prevent
                Case "P"
                    OrderOnHoldPermission = PermissionType.Warning
            End Select
        End If

    End Function

    ' Resulting in Neither, Prevent or Warning.
    ' In neither, no msgbox, no cancel. In warning, msgbox with no cancel. 
    ' in prevent, msgbox with cancel.
    Public Function PriceLessThanCostPermission(ByVal pstrItem_No As String, ByVal pstrLocation As String, ByVal pdblPrice As Double) As PermissionType

        PriceLessThanCostPermission = PermissionType.Neither

        'ID 02.17.2023 modified because prise consernig for customer from CA, we don't look in US prices
        If SqlCompliantString(pstrLocation) = "US1" Then
            pstrLocation = "1"
        End If


        ' First get item cost from location
        Dim strSql As String = _
        "SELECT ISNULL(Last_Cost, 0) AS Last_Cost " & _
        "FROM   IMINVLOC_SQL WITH (Nolock) " & _
        "WHERE  Item_No = '" & SqlCompliantString(pstrItem_No) & "' AND Loc = '" & SqlCompliantString(pstrLocation) & "' "


        '++ID 02.17.2023 need to check if is used that function , for to know if need to add exception for US
        'price for customer is not form US 
        '
        dt = db.DataTable(strSql)
        If dt.Rows.Count <> 0 Then

            ' If item cost > price, then check what kind of message to give
            If dt.Rows(0).Item("Last_Cost") > pdblPrice Then

                strSql = _
                "SELECT ISNULL(WARN_PREV, '') AS WARN_PREV " & _
                "FROM   SYSFIELD_SQL WITH (Nolock) " & _
                "WHERE  Package = 'XX' AND Description = 'Price Less Than Cost' "

                Dim dtSysField As New DataTable
                dtSysField = db.DataTable(strSql)

                If dtSysField.Rows.Count <> 0 Then

                    Select Case dtSysField.Rows(0).Item("WARN_PREV")
                        Case "P"
                            PriceLessThanCostPermission = PermissionType.Prevent
                        Case "W"
                            PriceLessThanCostPermission = PermissionType.Warning
                    End Select

                End If

            End If

        End If

    End Function

    Public Function PriceOverridePermission() As PermissionType

        PriceOverridePermission = PermissionType.Neither

        Dim strSql As String = _
        "SELECT ISNULL(WARN_PREV, '') AS WARN_PREV " & _
        "FROM   SYSFIELD_SQL WITH (Nolock) " & _
        "WHERE  Package = 'XX' AND Description = 'Price Override' "

        dt = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then

            Select Case dt.Rows(0).Item("WARN_PREV")
                Case "P"
                    PriceOverridePermission = PermissionType.Prevent
                Case "W"
                    PriceOverridePermission = PermissionType.Warning
            End Select

        End If

    End Function

End Class

Public Enum PermissionType
    Neither
    Prevent
    Warning
End Enum
