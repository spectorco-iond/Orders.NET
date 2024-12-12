Public Class cETA_Calculator

    Public Item_No As String
    Public Location As String
    Public Qty_Ordered As Double
    Public Qty_On_Hand As Double
    Public Qty_Prev_Bkord As Double
    Public Qty_Available As Double
    Public Extra_9 As String = ""
    Public Extra_10 As Double = 0
    Public Curr_Cd As String = ""

    Public Sub Calculate_ETA_For_Item()

        Try

            'If Source = OEI_SourceEnum.frmProductLineEntry Or Source = OEI_SourceEnum.Macola_ExactRepeat Then
            '    Call Calculate_ETA_ProductLineEntry()
            '    Exit Sub
            'End If

            ' From now on, we show ETA for everything not location 1 just in case...
            If Qty_Prev_Bkord <= 0 And Location = "1" Then
                Extra_9 = " "
                Extra_10 = 0
                Exit Sub
            End If

            Dim strDate As String = ""
            Dim dblQty As Double = 0
            Dim dblQtyRequired As Double = 0

            Dim dtItems As DataTable = Nothing
            Dim dt As DataTable
            Dim db As New cDBA

            Dim strsql As String = ""

            ''If Source = OEI_SourceEnum.frmOrder Then

            'strsql = _
            '"SELECT     L.ORD_GUID, L.ITEM_GUID, ISNULL(B.COMP_ITEM_GUID, '') AS COMP_ITEM_GUID, " & _
            '"           CASE WHEN ISNULL(B.COMP_ITEM_GUID, '') = '' THEN L.Item_No ELSE B.Item_No END AS Item_No " & _
            '"FROM 		OEI_ORDLIN L WITH (NOLOCK) " & _
            '"LEFT JOIN	OEI_ORDBLD B WITH (NOLOCK) ON L.ORD_GUID = B.ORD_GUID AND L.ITEM_GUID = B.KIT_ITEM_GUID " & _
            '"WHERE	    L.ORD_GUID = '" & oOrder.Ordhead.Ord_GUID & "' AND L.ITEM_GUID = '" & Item_Guid & "' "

            'dtItems = db.DataTable(strsql)

            'If dtItems.Rows.Count = 0 Then Exit Sub

            Dim strLaterDate As String = ""
            Dim dtLater_Promise_Dt As DateTime = Nothing ' CDate("01/01/2010")
            Dim dblLater_Qty As Double = 0

            ' IF IT IS A KIT, THERE WILL BE MORE THAN 1 LINE. IF AN ITEM, ALWAYS ONLY 1.
            'For Each dtItemRow As DataRow In dtItems.Rows

            strDate = ""
            dblQty = 0
            dblQtyRequired = 0

            'If dtItemRow.Item("Item_Guid") <> dtItemRow.Item("Comp_Item_Guid") Then

            strsql = "" & _
            "SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
            "           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
            "           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date, " & _
            "           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0)) " & IIf(Location.Trim = "1", "- DBO.OEI_INV_BUFFER(I.Item_No, 'CAD')", "") & " AS FLOAT) AS Qty_On_Hand, " & _
            "           CASE WHEN UPPER(SUBSTRING(ISNULL(I.Item_Note_2, ''), 1, 5)) = 'DISCO' OR UPPER(SUBSTRING(ISNULL(I.Item_Note_4, ''), 1, 5)) = 'DISCO' THEN 1 ELSE 0 END AS DISCONTINUED " & _
            "FROM       IMItmIdx_Sql I WITH (Nolock) " & _
            "INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Location & "' " & _
            "LEFT JOIN 	OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
            "LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No AND Q.Stk_Loc = '" & Location & "' " & _
            "WHERE      I.Item_No = '" & Item_No & "' AND L.Loc = '" & Location & "' " & _
            "ORDER BY   Q.Promise_Dt "

            'strsql = "" & _
            '"SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
            '"           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
            '"           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date, " & _
            '"           CAST((ISNULL(L.Qty_On_Hand,0) - ISNULL(L.Qty_Allocated,0) - ISNULL(PQV.Qty_Ordered,0)) AS FLOAT) AS Qty_On_Hand " & _
            '"FROM       IMItmIdx_Sql I WITH (Nolock) " & _
            '"LEFT JOIN  MDB_ITEM MI WITH (Nolock) ON I.User_Def_Fld_1 = MI.Item_Cd " & _
            '"INNER JOIN OEI_Item_Loc_Qty_View L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & Loc & "' " & _
            '"LEFT JOIN 	OEI_ORDLIN_PENDING_QTY_VIEW PQV WITH (Nolock) ON I.Item_No = PQV.Item_No AND I.Loc = PQV.Loc " & _
            '"LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No AND Q.Stk_Loc = '" & Loc & "' " & _
            '"WHERE      I.Item_No = '" & Trim(dtItemRow.Item("Item_No")) & "' AND L.Loc = '" & g_oOrdline.Loc & "' " & _
            '"ORDER BY   Q.Promise_Dt "

            'strsql = "" & _
            '"SELECT     I.Item_No, I.Activity_Cd, ISNULL(Q.qtyremaining, 0) AS QtyRemaining, " & _
            '"           CASE WHEN Promise_Dt IS NULL THEN NULL ELSE Promise_Dt END AS Promise_Dt, " & _
            '"           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as Promise_Date " & _
            '"FROM       IMItmIdx_Sql I WITH (Nolock) " & _
            '"LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No " & _
            '"WHERE      I.Item_No = '" & Trim(m_Item_No) & "' AND Q.Stk_Loc = '" & Loc & "' " & _
            '"ORDER BY   Q.Promise_Dt "
            'Debug.Print(Trim(dtItemRow.Item("Item_No")))
            dt = db.DataTable(strsql)

            'If Qty_On_Hand < 0 Then dblQtyRequired = Math.Abs(m_Qty_On_Hand)
            'dblQtyRequired += Qty_Prev_Bkord

            'Dim oOrdline As cOrdLine = Nothing

            'If dtItems.Rows.Count > 1 And Trim(dtItemRow.Item("COMP_ITEM_GUID")) <> "" Then
            '    If oOrder.OrdLines.Contains(Trim(dtItemRow.Item("COMP_ITEM_GUID"))) Then
            '        oOrdline = oOrder.OrdLines.Item(Trim(dtItemRow.Item("COMP_ITEM_GUID")))
            '    End If
            'Else
            '    If oOrder.OrdLines.Contains(Trim(dtItemRow.Item("ITEM_GUID"))) Then
            '        oOrdline = oOrder.OrdLines.Item(Trim(dtItemRow.Item("ITEM_GUID")))
            '    End If
            'End If

            'If dt.Rows.Count <> 0 And Not (oOrdline Is Nothing) Then ' And (m_Qty_Prev_Bkord > 0) Then  

            If dt.Rows.Count <> 0 Then

                If dt.Rows(0).Item("Qty_On_Hand") < Qty_Ordered Then

                    If Qty_On_Hand < 0 Then dblQtyRequired = Math.Abs(Qty_On_Hand)
                    dblQtyRequired += Qty_Prev_Bkord

                    Dim iCountRows As Integer = 0

                    For Each dtRow As DataRow In dt.Rows

                        iCountRows += 1

                        If (dblQty < dblQtyRequired And Qty_Ordered > dtRow.Item("Qty_On_Hand")) Or (Location <> "1" And dblQty < dblQtyRequired) Then

                            dblQty = dblQty + dtRow.Item("QtyRemaining")

                            If dtRow.Item("Activity_Cd").ToString = "O" Then
                                strDate = "DISCO"
                                strLaterDate = "DISCO"
                            ElseIf dblQty < dblQtyRequired And iCountRows = dt.Rows.Count Then
                                strDate = "QTY MISSING"
                                strLaterDate = "QTY MISSING"
                            Else
                                If IsDBNull(dtRow.Item("Promise_Dt")) Then
                                    strDate = "" ' dtRow.Item("Promise_Date").ToString
                                Else

                                    Dim Promise_Dt As DateTime = dtRow.Item("Promise_Dt")

                                    If dtLater_Promise_Dt.Year = 1 Or (Promise_Dt >= dtLater_Promise_Dt And dtLater_Promise_Dt.Year <> 1) And strLaterDate = "" Then
                                        'If strLaterDate <> "" Then
                                        dtLater_Promise_Dt = Promise_Dt
                                        dblLater_Qty = dblQty
                                        'End If
                                    End If

                                    strDate = Promise_Dt.Month.ToString.PadLeft(2, "0") & "/" & Promise_Dt.Day.ToString.PadLeft(2, "0") & "/" & Promise_Dt.Year.ToString.PadLeft(4, "0")

                                End If

                            End If

                        End If

                    Next dtRow

                    Extra_9 = strDate
                    Extra_10 = dblQty

                End If

                If dt.Rows(0).Item("Discontinued") <> 0 Then
                    If Extra_10 = 0 Then
                        Extra_9 = "DISCO"
                    Else
                        If Extra_9 = "QTY MISSING" Then
                            Extra_9 = "MISS, DISCO"
                        End If
                    End If
                End If

            Else

                Extra_9 = " "
                Extra_10 = 0

            End If

            'Else

            'If Not (oOrdline Is Nothing) Then
            '    Extra_9 = "NOT FOUND"
            '    Extra_10 = 0
            'End If

            If strLaterDate = "" And Qty_Prev_Bkord > 0 Then

                If Extra_9 <> "DISCO" Then
                    strLaterDate = "QTY MISSING" ' & Trim(dtItemRow.Item("Item_No"))
                Else
                    strLaterDate = "DISCO"
                End If

            End If

                'End If

                'End If

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
                '        If Qty_Prev_Bkord > 0 Then
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

End Class
