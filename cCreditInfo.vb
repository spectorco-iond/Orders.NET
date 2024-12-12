Option Strict Off
Option Explicit On

Public Class cCreditInfo

    Public Cus_No As String = ""
    Public DebtorNumber As String = ""
    Public Ord_No As String = "X"

    Public Amt_Cus_Balance As Double = 0
    Public Amt_Cur_Orders As Double = 0
    Public Amt_Open_Orders As Double = 0
    Public Amt_Orders_On_Hold As Double = 0
    Public Amt_Open_Invoices As Double = 0
    Public Amt_Credit_Limit As Double = 0
    Public Amt_Avail_Credit As Double = 0

    Public Days_Over_Terms As Long = 0
    Public Amt_Over_Terms As Double = 0

    Public Credit_Hold As String = ""
    Public Credit_Rate As String = ""

    Public Acc_Start_Dt As DateTime

    Public Comment1 As String = ""
    Public Comment2 As String = ""

    ' When loading, use this new definition
    'Public Sub New(ByVal pstrCus_No As String, ByVal pstrOrd_No As String)

    '    Cus_No = pstrCus_No
    '    Ord_No = pstrOrd_No
    '    Call Load()

    'End Sub

    ' On a new order, use this new definition
    Public Sub New(ByVal pstrCus_No As String)

        Cus_No = pstrCus_No
        Call Load()

    End Sub

    Public Sub New()

    End Sub

    Public Sub Init()

        Cus_No = String.Empty
        DebtorNumber = String.Empty
        Ord_No = "X"

        Amt_Cus_Balance = 0
        Amt_Cur_Orders = 0
        Amt_Open_Orders = 0
        Amt_Orders_On_Hold = 0
        Amt_Open_Invoices = 0
        Amt_Credit_Limit = 0
        Amt_Avail_Credit = 0

        Days_Over_Terms = 0
        Amt_Over_Terms = 0

        Credit_Hold = String.Empty
        Credit_Rate = String.Empty

        Acc_Start_Dt = Now

        Comment1 = String.Empty
        Comment2 = String.Empty

    End Sub

    Public Sub Load()

        On Error GoTo Trap

        'Dim rsData As New ADODB.Recordset
        Dim strSql As String

        strSql = "" & _
        "SELECT TOP 1 * " & _
        "FROM   ArCusFil_Sql WITH (Nolock) " & _
        "WHERE  Cus_No = '" & Cus_No & "' "

        'rsData.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If rsData.RecordCount <> 0 Then
        '    Credit_Hold = dt.rows(0).item("Hold_Fg").Value
        '    Acc_Start_Dt = dt.rows(0).item("Start_Dt").Value
        '    DebtorNumber = dt.rows(0).item("DebNr").Value
        'End If

        'rsData.Close()

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strSql)
        If dt.Rows.Count <> 0 Then
            Credit_Hold = dt.Rows(0).Item("Hold_Fg")
            Acc_Start_Dt = dt.Rows(0).Item("Start_Dt")
            DebtorNumber = dt.Rows(0).Item("DebNr")
        End If

        Amt_Credit_Limit = GetAmt_Credit_Limit(Cus_No)
        Amt_Cus_Balance = GetAmt_Cus_Balance(DebtorNumber)
        Amt_Cur_Orders = GetAmt_Cur_Orders(Cus_No)
        Amt_Open_Orders = GetAmt_Open_Orders(Cus_No)
        Amt_Orders_On_Hold = GetAmt_Orders_On_Hold(Cus_No)

        Amt_Avail_Credit = Amt_Credit_Limit - Amt_Cus_Balance - Amt_Cur_Orders - Amt_Open_Orders

        Days_Over_Terms = GetDays_Over_Terms(DebtorNumber)
        Amt_Over_Terms = GetAmt_Over_Terms(DebtorNumber)

        Exit Sub

Trap:
        MsgBox("Error occured : " & Err.Number & " " & Err.Description & " in cCreditInfo.Load")

    End Sub

    Public Function GetAmt_Credit_Limit(ByVal pstrCus_No As String) As Double

        On Error GoTo Trap

        GetAmt_Credit_Limit = 0

        Dim strsql As String = "" & _
        "Select C.Cr_Lmt + (C.Cr_Lmt * CTL.frm_cr_limit / 100)  AS Amt_Credit_Limit " & vbCrLf & _
        "FROM 	ARCUSFIL_SQL C WITH (Nolock), OECTLFIL_SQL CTL WITH (Nolock) " & vbCrLf & _
        "WHERE 	C.Cus_No = '" & Cus_No & "' "

        'Dim rsData As New ADODB.Recordset
        'rsData.Open(strsql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If rsData.RecordCount <> 0 Then
        '    GetAmt_Credit_Limit = dt.rows(0).item("Amt_Credit_Limit").Value
        'End If

        'rsData.Close()
        'rsData = Nothing

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then
            GetAmt_Credit_Limit = dt.Rows(0).Item("Amt_Credit_Limit")
        End If

        Exit Function

Trap:
        MsgBox("Error occured : " & Err.Number & " " & Err.Description & " in cCreditInfo.GetAmt_Credit_Limit")

    End Function

    Public Function GetAmt_Cus_Balance(ByVal pstrDebtorNumber As String) As Double

        On Error GoTo Trap

        GetAmt_Cus_Balance = 0

        Dim strsql As String = "" & _
        "SELECT 	CAST(ISNULL(SUM(Gbkmut.bdr_hfl),0) AS NUMERIC(16, 2)) AS Amt_Cus_Balance " & vbCrLf & _
        "FROM       Gbkmut WITH (Nolock) " & vbCrLf & _
        "INNER JOIN Grtbk WITH (Nolock) ON Gbkmut.reknr = Grtbk.reknr " & vbCrLf & _
        "WHERE 		Gbkmut.debnr = '" & pstrDebtorNumber & "' and " & vbCrLf & _
        "			Grtbk.omzrek = 'D' and Gbkmut.transtype in ('N', 'C', 'P')"

        'Dim rsData As New ADODB.Recordset
        'rsData.Open(strsql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If rsData.RecordCount <> 0 Then
        '    GetAmt_Cus_Balance = dt.rows(0).item("Amt_Cus_Balance").Value
        'End If

        'rsData.Close()
        'rsData = Nothing

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then
            GetAmt_Cus_Balance = dt.Rows(0).Item("Amt_Cus_Balance")
        End If

        Exit Function

Trap:
        MsgBox("Error occured : " & Err.Number & " " & Err.Description & " in cCreditInfo.GetAmt_Cus_Balance")

    End Function

    Public Function GetAmt_Cur_Orders(ByVal pstrCus_No As String) As Double

        On Error GoTo Trap

        GetAmt_Cur_Orders = 0

        Dim strsql As String = "" & _
        "SELECT CAST(ISNULL(SUM(ISNULL(Tot_Dollars,0) * ISNULL(Orig_trx_Rt,0)),0) AS NUMERIC(16, 2)) AS Amt_Cur_Orders " & vbCrLf & _
        "FROM 	OEORDHDR_SQL WITH (Nolock) " & vbCrLf & _
        "WHERE 	CUS_NO = '" & pstrCus_No & "' AND " & vbCrLf & _
        "		ORD_TYPE = 'O' and status <> 'C' AND Ord_No = '" & Ord_No & "' "

        'Dim rsData As New ADODB.Recordset
        'rsData.Open(strsql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If rsData.RecordCount <> 0 Then
        '    GetAmt_Cur_Orders = dt.rows(0).item("Amt_Cur_Orders").Value
        'End If

        'rsData.Close()
        'rsData = Nothing

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then
            GetAmt_Cur_Orders = dt.Rows(0).Item("Amt_Cur_Orders")
        End If

        Exit Function

Trap:
        MsgBox("Error occured : " & Err.Number & " " & Err.Description & " in cCreditInfo.GetAmt_Cur_Orders")

    End Function

    Public Function GetAmt_Open_Orders(ByVal pstrCus_No As String) As Double

        On Error GoTo Trap

        GetAmt_Open_Orders = 0

        Dim strsql As String = "" & _
        "SELECT	CAST(ISNULL(SUM(ISNULL(Tot_Dollars,0) * ISNULL(Orig_trx_Rt,0)),0) AS NUMERIC(16, 2)) AS Amt_Open_Orders " & vbCrLf & _
        "FROM 	OEORDHDR_SQL WITH (Nolock) " & vbCrLf & _
        "WHERE	CUS_NO = '" & pstrCus_No & "'  AND " & vbCrLf & _
        "		ORD_TYPE = 'O' AND Status BETWEEN '1' AND '8' "

        'Dim rsData As New ADODB.Recordset
        'rsData.Open(strsql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If rsData.RecordCount <> 0 Then
        '    GetAmt_Open_Orders = dt.rows(0).item("Amt_Open_Orders").Value - Amt_Cur_Orders
        'End If

        'rsData.Close()
        'rsData = Nothing

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then
            GetAmt_Open_Orders = dt.Rows(0).Item("Amt_Open_Orders")
        End If

        Exit Function

Trap:
        MsgBox("Error occured : " & Err.Number & " " & Err.Description & " in cCreditInfo.GetAmt_Open_Orders")

    End Function

    Public Function GetAmt_Orders_On_Hold(ByVal pstrCus_No As String) As Double

        On Error GoTo Trap

        GetAmt_Orders_On_Hold = 0

        Dim strsql As String = "" & _
        "SELECT	CAST(ISNULL(SUM(ISNULL(Tot_Dollars,0) * ISNULL(Orig_trx_Rt,0)),0) AS NUMERIC(16, 2)) AS Amt_Orders_On_Hold " & vbCrLf & _
        "FROM 	OEORDHDR_SQL WITH (Nolock) " & vbCrLf & _
        "WHERE	CUS_NO = '" & pstrCus_No & "' AND " & vbCrLf & _
        "		ord_type = 'O' and status = 'C' "

        'Dim rsData As New ADODB.Recordset
        'rsData.Open(strsql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If rsData.RecordCount <> 0 Then
        '    GetAmt_Orders_On_Hold = dt.rows(0).item("Amt_Orders_On_Hold").Value
        'End If

        'rsData.Close()
        'rsData = Nothing

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then
            GetAmt_Orders_On_Hold = dt.Rows(0).Item("Amt_Orders_On_Hold")
        End If

        Exit Function

Trap:
        MsgBox("Error occured : " & Err.Number & " " & Err.Description & " in cCreditInfo.GetAmt_Orders_On_Hold")

    End Function

    Public Function GetAmt_Open_Invoices(ByVal pstrDebtorNumber As String) As Double

        GetAmt_Open_Invoices = 0

    End Function

    Public Function GetDays_Over_Terms(ByVal pstrDebtorNumber As String) As Long

        On Error GoTo Trap

        GetDays_Over_Terms = 0

        Dim strDate As String = Date.Now.Day & "/" & Date.Now.Month & "/" & Date.Now.Year

        Dim strSql As String = "" & _
        "SELECT CASE    WHEN ISNULL(DATEDIFF(""DD"", DATEADD(DAY, MAX(a.PaymentDays), MAX(a.InvoiceDate)), GETDATE()),0) <= 0 " & vbCrLf & _
        "			    THEN 0 " & vbCrLf & _
        "			    ELSE DATEDIFF(""DD"", DATEADD(DAY, MAX(a.PaymentDays), MAX(a.InvoiceDate)), GETDATE()) END AS Days_Over_Terms " & vbCrLf & _
        "FROM ( " & vbCrLf & _
        "		SELECT 	AmountDC, PaymentDays, InvoiceDate, EntryNumber, InvoiceNumber " & vbCrLf & _
        "		FROM 	BankTransactions BT WITH (NoLock) " & vbCrLf & _
        "		WHERE 	BT.matchid IS NULL AND debtornumber = '" & pstrDebtorNumber & "' AND amountdc <> 0 AND status <> 'V' " & vbCrLf & _
        ") A " & vbCrLf & _
        "INNER JOIN Gbkmut GB WITH (NoLock) " & vbCrLf & _
        "		ON 	GB.transtype IN ('N', 'C', 'P') AND " & vbCrLf & _
        "			GB.vervdatfak < CONVERT(SMALLDATETIME, '" & strDate & "', 103) AND " & vbCrLf & _
        "			GB.bkstnr = A.entrynumber AND " & vbCrLf & _
        "			A.InvoiceDate = GB.datum AND A.InvoiceNumber = GB.faktuurnr " & vbCrLf & _
        "INNER JOIN Oectlfil_sql OE WITH (NoLock) " & vbCrLf & _
        "		ON  OE.oe_ctl_key_1 = 1 AND " & vbCrLf & _
        "			OE.meth_fg IN ('T', 'B') AND " & vbCrLf & _
        "			OE.days_from_terms < (CONVERT(SMALLDATETIME, '" & strDate & "', 103) - GB.vervdatfak) " & vbCrLf & _
        "INNER JOIN ( " & vbCrLf & _
        "		SELECT reknr " & vbCrLf & _
        "		FROM grtbk WITH (nolock) " & vbCrLf & _
        "		WHERE omzrek = 'D' " & vbCrLf & _
        ") GR ON GR.reknr = GB.reknr "

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strSql)
        If dt.Rows.Count <> 0 Then
            GetDays_Over_Terms = dt.Rows(0).Item("Days_Over_Terms")
        End If

        'Dim rsData As New ADODB.Recordset
        'rsData.Open(strsql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If rsData.RecordCount <> 0 Then
        '    GetDays_Over_Terms = dt.rows(0).item("Days_Over_Terms").Value
        'End If

        'rsData.Close()
        'rsData = Nothing

        Exit Function

Trap:
        MsgBox("Error occured : " & Err.Number & " " & Err.Description & " in cCreditInfo.GetDays_Over_Terms")

    End Function

    Public Function GetAmt_Over_Terms(ByVal pstrDebtorNumber As String) As Double

        On Error GoTo Trap

        GetAmt_Over_Terms = 0

        Dim strDate As String = Date.Now.Day & "/" & Date.Now.Month & "/" & Date.Now.Year

        Dim strsql As String = "" & _
        "SELECT CAST(ISNULL(SUM(ISNULL(A.AmountDC,0)),0) AS NUMERIC(16, 2)) AS Amt_Over_Terms " & vbCrLf & _
        "FROM ( " & vbCrLf & _
        "		SELECT 	AmountDC, PaymentDays, InvoiceDate, EntryNumber, InvoiceNumber " & vbCrLf & _
        "		FROM 	BankTransactions BT WITH (NoLock) " & vbCrLf & _
        "		WHERE 	BT.matchid IS NULL AND debtornumber = '" & pstrDebtorNumber & "' AND amountdc <> 0 AND status <> 'V' " & vbCrLf & _
        ") A " & vbCrLf & _
        "INNER JOIN Gbkmut GB WITH (NoLock) " & vbCrLf & _
        "		ON 	GB.transtype IN ('N', 'C', 'P') AND " & vbCrLf & _
        "			GB.vervdatfak < CONVERT(SMALLDATETIME, '" & strDate & "', 103) AND " & vbCrLf & _
        "			GB.bkstnr = A.entrynumber AND " & vbCrLf & _
        "			A.InvoiceDate = GB.datum AND A.InvoiceNumber = GB.faktuurnr " & vbCrLf & _
        "INNER JOIN Oectlfil_sql OE WITH (NoLock) " & vbCrLf & _
        "		ON  OE.oe_ctl_key_1 = 1 AND " & vbCrLf & _
        "			OE.meth_fg IN ('T', 'B') AND " & vbCrLf & _
        "			OE.days_from_terms < (CONVERT(SMALLDATETIME, '" & strDate & "', 103) - GB.vervdatfak) " & vbCrLf & _
        "INNER JOIN ( " & vbCrLf & _
        "		SELECT reknr " & vbCrLf & _
        "		FROM grtbk WITH (Nolock) " & vbCrLf & _
        "		WHERE omzrek = 'D' " & vbCrLf & _
        ") GR ON GR.reknr = GB.reknr "

        'Dim rsData As New ADODB.Recordset
        'rsData.Open(strsql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If rsData.RecordCount <> 0 Then
        '    GetAmt_Over_Terms = dt.rows(0).item("Amt_Over_Terms").Value
        'End If

        'rsData.Close()
        'rsData = Nothing

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then
            GetAmt_Over_Terms = dt.Rows(0).Item("Amt_Over_Terms")
        End If

        Exit Function

Trap:
        MsgBox("Error occured : " & Err.Number & " " & Err.Description & " in cCreditInfo.GetAmt_Over_Terms")

    End Function

    Public Sub Reset()

        Amt_Cus_Balance = 0
        Amt_Cur_Orders = 0
        Amt_Open_Orders = 0
        Amt_Orders_On_Hold = 0
        Amt_Open_Invoices = 0
        Amt_Credit_Limit = 0
        Amt_Avail_Credit = 0

        Days_Over_Terms = 0
        Amt_Over_Terms = 0

        Credit_Hold = String.Empty
        Credit_Rate = String.Empty

        Acc_Start_Dt = Now

        Comment1 = String.Empty
        Comment2 = String.Empty

    End Sub

End Class


