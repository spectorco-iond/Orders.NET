Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics

' Insert Include here


Public Class cOEPriceFile


#Region "private variables #################################################"


    Private m_bCd_Tp As Byte
    Private m_strCurr_Cd As String
    Private m_strCd_Tp_1_Cust_No As String
    Private m_strCd_Tp_3_Cust_Type As String
    Private m_strCd_Tp_1_Item_No As String
    Private m_strCd_Tp_2_Prod_Cat As String
    Private m_dtStart_Dt As DateTime
    Private m_dtEnd_Dt As DateTime
    Private m_strCd_Prc_Basis As String
    Private m_decMinimum_Qty_1 As Decimal
    Private m_decPrc_Or_Disc_1 As Decimal
    Private m_decMinimum_Qty_2 As Decimal
    Private m_decPrc_Or_Disc_2 As Decimal
    Private m_decMinimum_Qty_3 As Decimal
    Private m_decPrc_Or_Disc_3 As Decimal
    Private m_decMinimum_Qty_4 As Decimal
    Private m_decPrc_Or_Disc_4 As Decimal
    Private m_decMinimum_Qty_5 As Decimal
    Private m_decPrc_Or_Disc_5 As Decimal
    Private m_decMinimum_Qty_6 As Decimal
    Private m_decPrc_Or_Disc_6 As Decimal
    Private m_decMinimum_Qty_7 As Decimal
    Private m_decPrc_Or_Disc_7 As Decimal
    Private m_decMinimum_Qty_8 As Decimal
    Private m_decPrc_Or_Disc_8 As Decimal
    Private m_decMinimum_Qty_9 As Decimal
    Private m_decPrc_Or_Disc_9 As Decimal
    Private m_decMinimum_Qty_10 As Decimal
    Private m_decPrc_Or_Disc_10 As Decimal
    Private m_strExtra_1 As String
    Private m_strExtra_2 As String
    Private m_strExtra_3 As String
    Private m_strExtra_4 As String
    Private m_strExtra_5 As String
    Private m_strExtra_6 As String
    Private m_strExtra_7 As String
    Private m_strExtra_8 As String
    Private m_strExtra_9 As String
    Private m_decExtra_10 As Decimal
    Private m_decExtra_11 As Decimal
    Private m_decExtra_12 As Decimal
    Private m_decExtra_13 As Decimal
    Private m_intExtra_14 As Int32
    Private m_intExtra_15 As Int32
    Private m_strFiller_0004 As String
    Private m_decId As Decimal


#End Region


#Region "Public constructors ##############################################"


    Public Sub New()

        Try

            Call Init()

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Public Sub New(ByVal pId As Integer)

        Try

            Call Init()
            Call Load(pId)

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Public Sub New(ByVal pstrItem_No As String, ByVal pstrCurr_Cd As String)

        Try

            Call Init()
            Call Load(pstrItem_No, pstrCurr_Cd)

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Public Sub New(ByVal pstrItem_No As String, ByVal pstrCus_No As String, ByVal pstrCurr_Cd As String)

        Try

            Call Init()
            Call Load(pstrItem_No, pstrCus_No, pstrCurr_Cd)

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    'Public Sub New(ByVal pstrItem_No As String, ByVal pCus_Prog_ID As Integer, ByVal pCus_Prog_Item_List_ID As Integer)
    'Public Sub New(ByVal pstrItem_No As String, ByRef oProgram As cMdb_Cus_Prog)

    '    Try

    '        Call Init()
    '        Call Load(pstrItem_No, oProgram)

    '    Catch ex As Exception
    '        MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
    '    End Try

    'End Sub


#End Region


#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_bCd_Tp = 0
            m_strCurr_Cd = String.Empty
            m_strCd_Tp_1_Cust_No = String.Empty
            m_strCd_Tp_3_Cust_Type = String.Empty
            m_strCd_Tp_1_Item_No = String.Empty
            m_strCd_Tp_2_Prod_Cat = String.Empty
            m_dtStart_Dt = NoDate()
            m_dtEnd_Dt = NoDate()
            m_strCd_Prc_Basis = String.Empty
            m_decMinimum_Qty_1 = 0
            m_decPrc_Or_Disc_1 = 0
            m_decMinimum_Qty_2 = 0
            m_decPrc_Or_Disc_2 = 0
            m_decMinimum_Qty_3 = 0
            m_decPrc_Or_Disc_3 = 0
            m_decMinimum_Qty_4 = 0
            m_decPrc_Or_Disc_4 = 0
            m_decMinimum_Qty_5 = 0
            m_decPrc_Or_Disc_5 = 0
            m_decMinimum_Qty_6 = 0
            m_decPrc_Or_Disc_6 = 0
            m_decMinimum_Qty_7 = 0
            m_decPrc_Or_Disc_7 = 0
            m_decMinimum_Qty_8 = 0
            m_decPrc_Or_Disc_8 = 0
            m_decMinimum_Qty_9 = 0
            m_decPrc_Or_Disc_9 = 0
            m_decMinimum_Qty_10 = 0
            m_decPrc_Or_Disc_10 = 0
            m_strExtra_1 = String.Empty
            m_strExtra_2 = String.Empty
            m_strExtra_3 = String.Empty
            m_strExtra_4 = String.Empty
            m_strExtra_5 = String.Empty
            m_strExtra_6 = String.Empty
            m_strExtra_7 = String.Empty
            m_strExtra_8 = String.Empty
            m_strExtra_9 = String.Empty
            m_decExtra_10 = 0
            m_decExtra_11 = 0
            m_decExtra_12 = 0
            m_decExtra_13 = 0
            m_intExtra_14 = 0
            m_intExtra_15 = 0
            m_strFiller_0004 = String.Empty
            m_decId = 0

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try

            If Not (pdrRow.Item("Cd_Tp").Equals(DBNull.Value)) Then m_bCd_Tp = pdrRow.Item("Cd_Tp")
            If Not (pdrRow.Item("Curr_Cd").Equals(DBNull.Value)) Then m_strCurr_Cd = pdrRow.Item("Curr_Cd").ToString
            If Not (pdrRow.Item("Cd_Tp_1_Cust_No").Equals(DBNull.Value)) Then m_strCd_Tp_1_Cust_No = pdrRow.Item("Cd_Tp_1_Cust_No").ToString
            If Not (pdrRow.Item("Cd_Tp_3_Cust_Type").Equals(DBNull.Value)) Then m_strCd_Tp_3_Cust_Type = pdrRow.Item("Cd_Tp_3_Cust_Type").ToString
            If Not (pdrRow.Item("Cd_Tp_1_Item_No").Equals(DBNull.Value)) Then m_strCd_Tp_1_Item_No = pdrRow.Item("Cd_Tp_1_Item_No").ToString
            If Not (pdrRow.Item("Cd_Tp_2_Prod_Cat").Equals(DBNull.Value)) Then m_strCd_Tp_2_Prod_Cat = pdrRow.Item("Cd_Tp_2_Prod_Cat").ToString
            If Not (pdrRow.Item("Start_Dt").Equals(DBNull.Value)) Then m_dtStart_Dt = pdrRow.Item("Start_Dt")
            If Not (pdrRow.Item("End_Dt").Equals(DBNull.Value)) Then m_dtEnd_Dt = pdrRow.Item("End_Dt")
            If Not (pdrRow.Item("Cd_Prc_Basis").Equals(DBNull.Value)) Then m_strCd_Prc_Basis = pdrRow.Item("Cd_Prc_Basis").ToString
            If Not (pdrRow.Item("Minimum_Qty_1").Equals(DBNull.Value)) Then m_decMinimum_Qty_1 = pdrRow.Item("Minimum_Qty_1")
            If Not (pdrRow.Item("Prc_Or_Disc_1").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_1 = pdrRow.Item("Prc_Or_Disc_1")
            If Not (pdrRow.Item("Minimum_Qty_2").Equals(DBNull.Value)) Then m_decMinimum_Qty_2 = pdrRow.Item("Minimum_Qty_2")
            If Not (pdrRow.Item("Prc_Or_Disc_2").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_2 = pdrRow.Item("Prc_Or_Disc_2")
            If Not (pdrRow.Item("Minimum_Qty_3").Equals(DBNull.Value)) Then m_decMinimum_Qty_3 = pdrRow.Item("Minimum_Qty_3")
            If Not (pdrRow.Item("Prc_Or_Disc_3").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_3 = pdrRow.Item("Prc_Or_Disc_3")
            If Not (pdrRow.Item("Minimum_Qty_4").Equals(DBNull.Value)) Then m_decMinimum_Qty_4 = pdrRow.Item("Minimum_Qty_4")
            If Not (pdrRow.Item("Prc_Or_Disc_4").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_4 = pdrRow.Item("Prc_Or_Disc_4")
            If Not (pdrRow.Item("Minimum_Qty_5").Equals(DBNull.Value)) Then m_decMinimum_Qty_5 = pdrRow.Item("Minimum_Qty_5")
            If Not (pdrRow.Item("Prc_Or_Disc_5").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_5 = pdrRow.Item("Prc_Or_Disc_5")
            If Not (pdrRow.Item("Minimum_Qty_6").Equals(DBNull.Value)) Then m_decMinimum_Qty_6 = pdrRow.Item("Minimum_Qty_6")
            If Not (pdrRow.Item("Prc_Or_Disc_6").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_6 = pdrRow.Item("Prc_Or_Disc_6")
            If Not (pdrRow.Item("Minimum_Qty_7").Equals(DBNull.Value)) Then m_decMinimum_Qty_7 = pdrRow.Item("Minimum_Qty_7")
            If Not (pdrRow.Item("Prc_Or_Disc_7").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_7 = pdrRow.Item("Prc_Or_Disc_7")
            If Not (pdrRow.Item("Minimum_Qty_8").Equals(DBNull.Value)) Then m_decMinimum_Qty_8 = pdrRow.Item("Minimum_Qty_8")
            If Not (pdrRow.Item("Prc_Or_Disc_8").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_8 = pdrRow.Item("Prc_Or_Disc_8")
            If Not (pdrRow.Item("Minimum_Qty_9").Equals(DBNull.Value)) Then m_decMinimum_Qty_9 = pdrRow.Item("Minimum_Qty_9")
            If Not (pdrRow.Item("Prc_Or_Disc_9").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_9 = pdrRow.Item("Prc_Or_Disc_9")
            If Not (pdrRow.Item("Minimum_Qty_10").Equals(DBNull.Value)) Then m_decMinimum_Qty_10 = pdrRow.Item("Minimum_Qty_10")
            If Not (pdrRow.Item("Prc_Or_Disc_10").Equals(DBNull.Value)) Then m_decPrc_Or_Disc_10 = pdrRow.Item("Prc_Or_Disc_10")
            If Not (pdrRow.Item("Extra_1").Equals(DBNull.Value)) Then m_strExtra_1 = pdrRow.Item("Extra_1").ToString
            If Not (pdrRow.Item("Extra_2").Equals(DBNull.Value)) Then m_strExtra_2 = pdrRow.Item("Extra_2").ToString
            If Not (pdrRow.Item("Extra_3").Equals(DBNull.Value)) Then m_strExtra_3 = pdrRow.Item("Extra_3").ToString
            If Not (pdrRow.Item("Extra_4").Equals(DBNull.Value)) Then m_strExtra_4 = pdrRow.Item("Extra_4").ToString
            If Not (pdrRow.Item("Extra_5").Equals(DBNull.Value)) Then m_strExtra_5 = pdrRow.Item("Extra_5").ToString
            If Not (pdrRow.Item("Extra_6").Equals(DBNull.Value)) Then m_strExtra_6 = pdrRow.Item("Extra_6").ToString
            If Not (pdrRow.Item("Extra_7").Equals(DBNull.Value)) Then m_strExtra_7 = pdrRow.Item("Extra_7").ToString
            If Not (pdrRow.Item("Extra_8").Equals(DBNull.Value)) Then m_strExtra_8 = pdrRow.Item("Extra_8").ToString
            If Not (pdrRow.Item("Extra_9").Equals(DBNull.Value)) Then m_strExtra_9 = pdrRow.Item("Extra_9").ToString
            If Not (pdrRow.Item("Extra_10").Equals(DBNull.Value)) Then m_decExtra_10 = pdrRow.Item("Extra_10")
            If Not (pdrRow.Item("Extra_11").Equals(DBNull.Value)) Then m_decExtra_11 = pdrRow.Item("Extra_11")
            If Not (pdrRow.Item("Extra_12").Equals(DBNull.Value)) Then m_decExtra_12 = pdrRow.Item("Extra_12")
            If Not (pdrRow.Item("Extra_13").Equals(DBNull.Value)) Then m_decExtra_13 = pdrRow.Item("Extra_13")
            If Not (pdrRow.Item("Extra_14").Equals(DBNull.Value)) Then m_intExtra_14 = pdrRow.Item("Extra_14")
            If Not (pdrRow.Item("Extra_15").Equals(DBNull.Value)) Then m_intExtra_15 = pdrRow.Item("Extra_15")
            If Not (pdrRow.Item("Filler_0004").Equals(DBNull.Value)) Then m_strFiller_0004 = pdrRow.Item("Filler_0004").ToString
            If Not (pdrRow.Item("Id").Equals(DBNull.Value)) Then m_decId = pdrRow.Item("Id")

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            'pdrRow.Item("Id") = m_decId
            pdrRow.Item("Cd_Tp") = m_bCd_Tp
            pdrRow.Item("Curr_Cd") = m_strCurr_Cd
            pdrRow.Item("Cd_Tp_1_Cust_No") = m_strCd_Tp_1_Cust_No
            pdrRow.Item("Cd_Tp_3_Cust_Type") = m_strCd_Tp_3_Cust_Type
            pdrRow.Item("Cd_Tp_1_Item_No") = m_strCd_Tp_1_Item_No
            pdrRow.Item("Cd_Tp_2_Prod_Cat") = m_strCd_Tp_2_Prod_Cat
            If m_dtStart_Dt.Year <> 1 Then pdrRow.Item("Start_Dt") = m_dtStart_Dt
            If m_dtEnd_Dt.Year <> 1 Then pdrRow.Item("End_Dt") = m_dtEnd_Dt
            pdrRow.Item("Cd_Prc_Basis") = m_strCd_Prc_Basis
            pdrRow.Item("Minimum_Qty_1") = m_decMinimum_Qty_1
            pdrRow.Item("Prc_Or_Disc_1") = m_decPrc_Or_Disc_1
            pdrRow.Item("Minimum_Qty_2") = m_decMinimum_Qty_2
            pdrRow.Item("Prc_Or_Disc_2") = m_decPrc_Or_Disc_2
            pdrRow.Item("Minimum_Qty_3") = m_decMinimum_Qty_3
            pdrRow.Item("Prc_Or_Disc_3") = m_decPrc_Or_Disc_3
            pdrRow.Item("Minimum_Qty_4") = m_decMinimum_Qty_4
            pdrRow.Item("Prc_Or_Disc_4") = m_decPrc_Or_Disc_4
            pdrRow.Item("Minimum_Qty_5") = m_decMinimum_Qty_5
            pdrRow.Item("Prc_Or_Disc_5") = m_decPrc_Or_Disc_5
            pdrRow.Item("Minimum_Qty_6") = m_decMinimum_Qty_6
            pdrRow.Item("Prc_Or_Disc_6") = m_decPrc_Or_Disc_6
            pdrRow.Item("Minimum_Qty_7") = m_decMinimum_Qty_7
            pdrRow.Item("Prc_Or_Disc_7") = m_decPrc_Or_Disc_7
            pdrRow.Item("Minimum_Qty_8") = m_decMinimum_Qty_8
            pdrRow.Item("Prc_Or_Disc_8") = m_decPrc_Or_Disc_8
            pdrRow.Item("Minimum_Qty_9") = m_decMinimum_Qty_9
            pdrRow.Item("Prc_Or_Disc_9") = m_decPrc_Or_Disc_9
            pdrRow.Item("Minimum_Qty_10") = m_decMinimum_Qty_10
            pdrRow.Item("Prc_Or_Disc_10") = m_decPrc_Or_Disc_10
            'pdrRow.Item("Extra_1") = m_strExtra_1
            'pdrRow.Item("Extra_2") = m_strExtra_2
            'pdrRow.Item("Extra_3") = m_strExtra_3
            'pdrRow.Item("Extra_4") = m_strExtra_4
            'pdrRow.Item("Extra_5") = m_strExtra_5
            'pdrRow.Item("Extra_6") = m_strExtra_6
            'pdrRow.Item("Extra_7") = m_strExtra_7
            'pdrRow.Item("Extra_8") = m_strExtra_8
            'pdrRow.Item("Extra_9") = m_strExtra_9
            'pdrRow.Item("Extra_10") = m_decExtra_10
            'pdrRow.Item("Extra_11") = m_decExtra_11
            'pdrRow.Item("Extra_12") = m_decExtra_12
            'pdrRow.Item("Extra_13") = m_decExtra_13
            pdrRow.Item("Extra_14") = m_intExtra_14 ' We only save this field, it links the special pricing Cus_Prog_ID
            'pdrRow.Item("Extra_15") = m_intExtra_15
            'pdrRow.Item("Filler_0004") = m_strFiller_0004

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


#End Region


#Region "Public maintenance procedures ####################################"

    ' Deletes the current Comment from the database, if it exists.
    Public Sub Delete()

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            ' We make sure only to remove data linked to a program, special pricing or quote.
            strSql = _
            "DELETE FROM Oeprcfil_Sql " & _
            "WHERE  Id = " & m_decId & " AND ISNULL(Extra_14, 0) <> 0 AND ISNULL(Extra_14, 0) = " & m_intExtra_14 & " AND CD_TP_1_CUST_NO = '" & m_strCd_Tp_1_Cust_No.Trim & "' AND CD_TP_1_ITEM_NO = '" & m_strCd_Tp_1_Item_No.Trim & "' "

            db.Execute(strSql)

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub Load(ByVal pintID As Integer)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   Oeprcfil_Sql WITH (Nolock) " & _
            "WHERE  Id = " & pintID & " "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub Load(ByVal pstrItem_No As String, ByVal pstrCurr_Cd As String)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   Oeprcfil_Sql WITH (Nolock) " & _
            "WHERE  Cd_Tp = 6 AND Cd_Tp_1_Item_No = '" & pstrItem_No & "' AND " & _
            "       Curr_Cd = '" & pstrCurr_Cd & "' AND " & _
            "       (Start_Dt IS NULL OR Start_Dt <= GETDATE()) AND (End_Dt IS NULL OR End_Dt >= GETDATE()) "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub Load(ByVal pstrItem_No As String, ByVal pstrCus_No As String, ByVal pstrCurr_Cd As String)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   Oeprcfil_Sql WITH (Nolock) " & _
            "WHERE  Cd_Tp = 6 AND " & _
            "       Cd_Tp_1_Item_No = '" & pstrItem_No & "' AND " & _
            "       Cd_Tp_1_Cust_No = '" & pstrCus_No & "' AND " & _
            "       Curr_Cd = '" & pstrItem_No & "' " & _
            "       (Start_Dt IS NULL OR Start_Dt <= GETDATE()) AND (End_Dt IS NULL OR End_Dt >= GETDATE()) "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    '++ID 07222024 function get the price by item_no, type, curr_cd  - > used for type 10 , after selected No Imprint
    Public Sub LoadPriceByType(ByVal pstrItem_No As String, ByVal pstrCurr_Cd As String, ByVal pstrType As String, Optional ByVal _cus_no As String = "", Optional ByVal _qty As Int32 = 1, Optional ByRef _priceretrun As Decimal = 0.00)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "SELECT * " &
            "FROM   Oeprcfil_Sql WITH (Nolock) " &
            "WHERE  Cd_Tp = '" & pstrType & "' AND " &
            "       Cd_Tp_1_Item_No = '" & pstrItem_No & "' AND " &
            "       Curr_Cd = '" & pstrCurr_Cd & "' AND " &
            "     End_Dt >= GETDATE()  "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            Else

                strSql =
                       "SELECT * " &
                       "FROM   Oeprcfil_Sql WITH (Nolock) " &
                       "WHERE  Cd_Tp = 6 AND " &
                       "       Cd_Tp_1_Item_No = '" & pstrItem_No & "' AND " &
                       "       Curr_Cd = '" & pstrCurr_Cd & "' AND " &
                       "     End_Dt >= GETDATE()  "

                strSql = "select  CAST(DBO.OEI_Item_Price_20140101('" & pstrCurr_Cd & "', '', '" & pstrItem_No & "', '', '', '', " & _qty & ") As Decimal(10,3)) As price  "


                dt = db.DataTable(strSql)

                If dt.Rows.Count <> 0 Then
                    ' Call LoadLine(dt.Rows(0))
                    _priceretrun = dt.Rows(0).Item("price")
                End If
            End If

        Catch ex As Exception
            MsgBox("Error in ccOeprcefil." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub
    '----------------------------------------------------------------------------------------------------------------




    'Private Sub Load(ByVal pstrItem_No As String, ByVal pCus_Prog_ID As Integer, ByVal pCus_Prog_Item_List_ID As Integer)
    'Private Sub Load(ByVal pstrItem_No As String, ByVal pCus_Prog_ID As Integer)
    'Private Sub Load(ByVal pstrItem_No As String, ByVal opProgram As cMdb_Cus_Prog)

    '    Try

    '        Dim db As New cDBA
    '        Dim dt As New DataTable

    '        Dim strSql As String
    '        strSql = _
    '        "SELECT * " & _
    '        "FROM   Oeprcfil_Sql WITH (Nolock) " & _
    '        "WHERE  Cd_Tp = " & IIf(opProgram.Prog_Type = 2, "1", "9") & " AND " & _
    '        "       Cd_Tp_1_Item_No = '" & pstrItem_No & "' AND " & _
    '        "       Extra_14 = " & opProgram.Cus_Prog_Id & " "

    '        'strSql = _
    '        '"SELECT * " & _
    '        '"FROM   Oeprcfil_Sql WITH (Nolock) " & _
    '        '"WHERE  Cd_Tp = 1 AND " & _
    '        '"       Cd_Tp_1_Item_No = '" & pstrItem_No & "' AND " & _
    '        '"       Extra_14 = " & pCus_Prog_ID & " "
    '        '"WHERE  Cd_Tp_1_Item_No = '" & pstrItem_No & "' AND Extra_14 = " & pCus_Prog_ID & " AND Extra_15 = " & pCus_Prog_Item_List_ID

    '        dt = db.DataTable(strSql)

    '        If dt.Rows.Count <> 0 Then
    '            Call LoadLine(dt.Rows(0))
    '        End If

    '    Catch ex As Exception
    '        MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
    '    End Try

    'End Sub


    ' Update the current Comment into the database, or creates it if not existing
    Public Sub Save()

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String

            strSql = _
            "SELECT * " & _
            "FROM   Oeprcfil_Sql " & _
            "WHERE  Id = " & m_decId & " "

            dt = db.DataTable(strSql)

            If dt.Rows.Count = 0 Then
                drRow = dt.NewRow()
            Else
                drRow = dt.Rows(0)
            End If

            Call SaveLine(drRow)

            If dt.Rows.Count = 0 Then

                db.DBDataTable.Rows.Add(drRow)
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand

            Else

                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand

            End If

            db.DBDataAdapter.Update(db.DBDataTable)

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


#End Region



#Region "Public procedures ################################################"

    Public Sub Set_Prc_Or_Disc(ByVal piIndex As Integer, ByVal pdecPrice As Decimal)

        Try
            Select Case piIndex
                Case 1
                    m_decPrc_Or_Disc_1 = pdecPrice
                Case 2
                    m_decPrc_Or_Disc_2 = pdecPrice
                Case 3
                    m_decPrc_Or_Disc_3 = pdecPrice
                Case 4
                    m_decPrc_Or_Disc_4 = pdecPrice
                Case 5
                    m_decPrc_Or_Disc_5 = pdecPrice
                Case 6
                    m_decPrc_Or_Disc_6 = pdecPrice
                Case 7
                    m_decPrc_Or_Disc_7 = pdecPrice
                Case 8
                    m_decPrc_Or_Disc_8 = pdecPrice
                Case 9
                    m_decPrc_Or_Disc_9 = pdecPrice
                Case 10
                    m_decPrc_Or_Disc_10 = pdecPrice

            End Select

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub

    Public Sub Set_Minimum_Qty(ByVal piIndex As Integer, ByVal pdecQty As Decimal)

        Try
            Select Case piIndex

                Case 1
                    m_decMinimum_Qty_1 = pdecQty
                Case 2
                    m_decMinimum_Qty_2 = pdecQty
                Case 3
                    m_decMinimum_Qty_3 = pdecQty
                Case 4
                    m_decMinimum_Qty_4 = pdecQty
                Case 5
                    m_decMinimum_Qty_5 = pdecQty
                Case 6
                    m_decMinimum_Qty_6 = pdecQty
                Case 7
                    m_decMinimum_Qty_7 = pdecQty
                Case 8
                    m_decMinimum_Qty_8 = pdecQty
                Case 9
                    m_decMinimum_Qty_9 = pdecQty
                Case 10
                    m_decMinimum_Qty_10 = pdecQty

            End Select

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub

    Public Function Get_Minimum_Qty(ByVal piIndex As Integer) As Integer

        Get_Minimum_Qty = 0

        Try
            Select Case piIndex

                Case 1
                    Get_Minimum_Qty = m_decMinimum_Qty_1
                Case 2
                    Get_Minimum_Qty = m_decMinimum_Qty_2
                Case 3
                    Get_Minimum_Qty = m_decMinimum_Qty_3
                Case 4
                    Get_Minimum_Qty = m_decMinimum_Qty_4
                Case 5
                    Get_Minimum_Qty = m_decMinimum_Qty_5
                Case 6
                    Get_Minimum_Qty = m_decMinimum_Qty_6
                Case 7
                    Get_Minimum_Qty = m_decMinimum_Qty_7
                Case 8
                    Get_Minimum_Qty = m_decMinimum_Qty_8
                Case 9
                    Get_Minimum_Qty = m_decMinimum_Qty_9
                Case 10
                    Get_Minimum_Qty = m_decMinimum_Qty_10

            End Select

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

    Public Function Get_Eqp_Price() As Decimal

        Get_Eqp_Price = 0

        Try
            Dim oPrice As New cOEPriceFile(Item_No, Curr_Cd)
            Get_Eqp_Price = oPrice.m_decPrc_Or_Disc_5

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

    Public Function Get_Inventory_Buffer() As Integer

        Get_Inventory_Buffer = 0

        Try
            Dim db As New cDBA
            Dim dt As DataTable

            Dim strsql As String = _
            "SELECT ISNULL(Inventory_Buffer_Pct, 0) AS Inventory_Buffer_Pct, " & _
            "       ISNULL(Inventory_Buffer_Col, 0) AS Inventory_Buffer_Col " & _
            "FROM   OEI_CONFIG WITH (NoLock)"

            dt = db.DataTable(strsql)
            If dt.Rows.Count <> 0 Then
                ' We set +1 because Column 1 is always empty
                Get_Inventory_Buffer = Get_Minimum_Qty(dt.Rows(0).Item("Inventory_Buffer_Col") + 1) / 100 * dt.Rows(0).Item("Inventory_Buffer_Pct")
            End If

        Catch ex As Exception
            MsgBox("Error in cMacolaOeprcfil_Sql." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

#End Region


#Region "Public properties ################################################"

    Public Property Cd_Tp() As Byte
        Get
            Cd_Tp = m_bCd_Tp
        End Get
        Set(ByVal value As Byte)
            m_bCd_Tp = value
        End Set
    End Property

    Public Property Curr_Cd() As String
        Get
            Curr_Cd = m_strCurr_Cd
        End Get
        Set(ByVal value As String)
            m_strCurr_Cd = value
        End Set
    End Property

    Public Property Cus_No() As String
        Get
            Cus_No = m_strCd_Tp_1_Cust_No
        End Get
        Set(ByVal value As String)
            m_strCd_Tp_1_Cust_No = value
        End Set
    End Property

    Public Property Cd_Tp_3_Cust_Type() As String
        Get
            Cd_Tp_3_Cust_Type = m_strCd_Tp_3_Cust_Type
        End Get
        Set(ByVal value As String)
            m_strCd_Tp_3_Cust_Type = value
        End Set
    End Property

    Public Property Item_No() As String
        Get
            Item_No = m_strCd_Tp_1_Item_No
        End Get
        Set(ByVal value As String)
            m_strCd_Tp_1_Item_No = value
        End Set
    End Property

    Public Property Prod_Cat() As String
        Get
            Prod_Cat = m_strCd_Tp_2_Prod_Cat
        End Get
        Set(ByVal value As String)
            m_strCd_Tp_2_Prod_Cat = value
        End Set
    End Property

    Public Property Start_Dt() As DateTime
        Get
            Start_Dt = m_dtStart_Dt
        End Get
        Set(ByVal value As DateTime)
            m_dtStart_Dt = value
        End Set
    End Property

    Public Property End_Dt() As DateTime
        Get
            End_Dt = m_dtEnd_Dt
        End Get
        Set(ByVal value As DateTime)
            m_dtEnd_Dt = value
        End Set
    End Property

    Public Property Cd_Prc_Basis() As String
        Get
            Cd_Prc_Basis = m_strCd_Prc_Basis
        End Get
        Set(ByVal value As String)
            m_strCd_Prc_Basis = value
        End Set
    End Property

    Public Property Minimum_Qty_1() As Decimal
        Get
            Minimum_Qty_1 = m_decMinimum_Qty_1
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_1 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_1() As Decimal
        Get
            Prc_Or_Disc_1 = m_decPrc_Or_Disc_1
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_1 = value
        End Set
    End Property

    Public Property Minimum_Qty_2() As Decimal
        Get
            Minimum_Qty_2 = m_decMinimum_Qty_2
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_2 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_2() As Decimal
        Get
            Prc_Or_Disc_2 = m_decPrc_Or_Disc_2
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_2 = value
        End Set
    End Property

    Public Property Minimum_Qty_3() As Decimal
        Get
            Minimum_Qty_3 = m_decMinimum_Qty_3
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_3 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_3() As Decimal
        Get
            Prc_Or_Disc_3 = m_decPrc_Or_Disc_3
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_3 = value
        End Set
    End Property

    Public Property Minimum_Qty_4() As Decimal
        Get
            Minimum_Qty_4 = m_decMinimum_Qty_4
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_4 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_4() As Decimal
        Get
            Prc_Or_Disc_4 = m_decPrc_Or_Disc_4
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_4 = value
        End Set
    End Property

    Public Property Minimum_Qty_5() As Decimal
        Get
            Minimum_Qty_5 = m_decMinimum_Qty_5
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_5 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_5() As Decimal
        Get
            Prc_Or_Disc_5 = m_decPrc_Or_Disc_5
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_5 = value
        End Set
    End Property

    Public Property Minimum_Qty_6() As Decimal
        Get
            Minimum_Qty_6 = m_decMinimum_Qty_6
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_6 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_6() As Decimal
        Get
            Prc_Or_Disc_6 = m_decPrc_Or_Disc_6
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_6 = value
        End Set
    End Property

    Public Property Minimum_Qty_7() As Decimal
        Get
            Minimum_Qty_7 = m_decMinimum_Qty_7
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_7 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_7() As Decimal
        Get
            Prc_Or_Disc_7 = m_decPrc_Or_Disc_7
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_7 = value
        End Set
    End Property

    Public Property Minimum_Qty_8() As Decimal
        Get
            Minimum_Qty_8 = m_decMinimum_Qty_8
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_8 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_8() As Decimal
        Get
            Prc_Or_Disc_8 = m_decPrc_Or_Disc_8
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_8 = value
        End Set
    End Property

    Public Property Minimum_Qty_9() As Decimal
        Get
            Minimum_Qty_9 = m_decMinimum_Qty_9
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_9 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_9() As Decimal
        Get
            Prc_Or_Disc_9 = m_decPrc_Or_Disc_9
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_9 = value
        End Set
    End Property

    Public Property Minimum_Qty_10() As Decimal
        Get
            Minimum_Qty_10 = m_decMinimum_Qty_10
        End Get
        Set(ByVal value As Decimal)
            m_decMinimum_Qty_10 = value
        End Set
    End Property

    Public Property Prc_Or_Disc_10() As Decimal
        Get
            Prc_Or_Disc_10 = m_decPrc_Or_Disc_10
        End Get
        Set(ByVal value As Decimal)
            m_decPrc_Or_Disc_10 = value
        End Set
    End Property

    Public Property Extra_1() As String
        Get
            Extra_1 = m_strExtra_1
        End Get
        Set(ByVal value As String)
            m_strExtra_1 = value
        End Set
    End Property

    Public Property Extra_2() As String
        Get
            Extra_2 = m_strExtra_2
        End Get
        Set(ByVal value As String)
            m_strExtra_2 = value
        End Set
    End Property

    Public Property Extra_3() As String
        Get
            Extra_3 = m_strExtra_3
        End Get
        Set(ByVal value As String)
            m_strExtra_3 = value
        End Set
    End Property

    Public Property Extra_4() As String
        Get
            Extra_4 = m_strExtra_4
        End Get
        Set(ByVal value As String)
            m_strExtra_4 = value
        End Set
    End Property

    Public Property Extra_5() As String
        Get
            Extra_5 = m_strExtra_5
        End Get
        Set(ByVal value As String)
            m_strExtra_5 = value
        End Set
    End Property

    Public Property Extra_6() As String
        Get
            Extra_6 = m_strExtra_6
        End Get
        Set(ByVal value As String)
            m_strExtra_6 = value
        End Set
    End Property

    Public Property Extra_7() As String
        Get
            Extra_7 = m_strExtra_7
        End Get
        Set(ByVal value As String)
            m_strExtra_7 = value
        End Set
    End Property

    Public Property Extra_8() As String
        Get
            Extra_8 = m_strExtra_8
        End Get
        Set(ByVal value As String)
            m_strExtra_8 = value
        End Set
    End Property

    Public Property Spector_Cd() As String
        Get
            Spector_Cd = m_strExtra_9
        End Get
        Set(ByVal value As String)
            m_strExtra_9 = value
        End Set
    End Property

    Public Property Extra_10() As Decimal
        Get
            Extra_10 = m_decExtra_10
        End Get
        Set(ByVal value As Decimal)
            m_decExtra_10 = value
        End Set
    End Property

    Public Property Extra_11() As Decimal
        Get
            Extra_11 = m_decExtra_11
        End Get
        Set(ByVal value As Decimal)
            m_decExtra_11 = value
        End Set
    End Property

    Public Property Extra_12() As Decimal
        Get
            Extra_12 = m_decExtra_12
        End Get
        Set(ByVal value As Decimal)
            m_decExtra_12 = value
        End Set
    End Property

    Public Property Extra_13() As Decimal
        Get
            Extra_13 = m_decExtra_13
        End Get
        Set(ByVal value As Decimal)
            m_decExtra_13 = value
        End Set
    End Property

    Public Property Cus_Prog_Id() As Integer
        Get
            Cus_Prog_Id = m_intExtra_14
        End Get
        Set(ByVal value As Int32)
            m_intExtra_14 = value
        End Set
    End Property

    Public Property Cus_Prog_Item_List_ID() As Integer
        Get
            Cus_Prog_Item_List_ID = m_intExtra_15
        End Get
        Set(ByVal value As Int32)
            m_intExtra_15 = value
        End Set
    End Property

    Public Property Filler_0004() As String
        Get
            Filler_0004 = m_strFiller_0004
        End Get
        Set(ByVal value As String)
            m_strFiller_0004 = value
        End Set
    End Property

    Public Property Id() As Decimal
        Get
            Id = m_decId
        End Get
        Set(ByVal value As Decimal)
            m_decId = value
        End Set
    End Property

#End Region


End Class ' cMacolaOeprcfil_Sql



'Public Class cOEPriceFile

'End Class


