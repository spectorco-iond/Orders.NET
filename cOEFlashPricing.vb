Public Class cOEFlashPricing

    ' Unexposed properties
    Private m_ClassName As String = "cOEFlashPricing"
    Private db As New cDBA

    ' Exposed properties
    Private m_intOE_Price_ID As Integer
    Private m_strProgram_Name As String
    Private m_strItem_CD As String
    Private m_strCurr_CD As String
    Private m_intCD_TP As Integer
    Private m_intCFG_Spc_Prc_ID As Integer
    Private m_strCD_TP_1_Cust_No As String
    Private m_strCd_Tp_2_Prod_Cat As String
    Private m_strCd_Tp_3_Cust_Type As String
    Private m_datStart_Date As Date
    Private m_datEnd_Date As Date
    Private m_strCd_Prc_Basis As String
    Private m_intMinimum_Qty_1 As Integer
    Private m_dblPrc_Or_Disc_1 As Double
    Private m_intMinimum_Qty_2 As Integer
    Private m_dblPrc_Or_Disc_2 As Double
    Private m_intMinimum_Qty_3 As Integer
    Private m_dblPrc_Or_Disc_3 As Double
    Private m_intMinimum_Qty_4 As Integer
    Private m_dblPrc_Or_Disc_4 As Double
    Private m_intMinimum_Qty_5 As Integer
    Private m_dblPrc_Or_Disc_5 As Double
    Private m_intMinimum_Qty_6 As Integer
    Private m_dblPrc_Or_Disc_6 As Double
    Private m_intMinimum_Qty_7 As Integer
    Private m_dblPrc_Or_Disc_7 As Double
    Private m_intMinimum_Qty_8 As Integer
    Private m_dblPrc_Or_Disc_8 As Double
    Private m_intMinimum_Qty_9 As Integer
    Private m_dblPrc_Or_Disc_9 As Double
    Private m_blnOEI_Warning As Boolean
    Private m_blnMacola_Import As Boolean
    Private m_strEQP_Type As String
    Private m_intEQP_Plus_Pct As Integer
    Private m_blnWhen_Req As Boolean
    Private m_strAutorized_By As String
    Private m_strComments As String

    Private m_blnDirty As Boolean = False

#Region "Public constructors ##############################################"

    Public Sub New()

        Call Init()

    End Sub

    Public Sub New(ByVal pstrCus_No As String)

        Call Init()
        m_strCD_TP_1_Cust_No = pstrCus_No

    End Sub

    Public Sub New(ByVal pintCharge_Usage_ID As Integer)

        m_intOE_Price_ID = pintCharge_Usage_ID
        Call Load()

    End Sub


    'Public Sub New(ByVal pstrItem_CD As String, ByVal pstrCurr_Cd As String)

    '    Call Init()
    '    m_strItem_CD = pstrItem_CD
    '    m_strCurr_CD = pstrCurr_Cd

    'End Sub

    'Private Sub New(ByVal pintOE_Price_ID As Integer)

    '    m_intOE_Price_ID = pintOE_Price_ID
    '    Call Load()

    'End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_intOE_Price_ID = 0

            m_strItem_CD = String.Empty
            m_strCurr_CD = String.Empty

            m_intCD_TP = 0
            m_intCFG_Spc_Prc_ID = 0
            m_strCD_TP_1_Cust_No = String.Empty
            m_strCd_Tp_2_Prod_Cat = String.Empty
            m_strCd_Tp_3_Cust_Type = String.Empty
            m_datStart_Date = NoDate()
            m_datEnd_Date = NoDate()
            m_strCd_Prc_Basis = String.Empty

            m_intMinimum_Qty_1 = 0
            m_dblPrc_Or_Disc_1 = 0
            m_intMinimum_Qty_2 = 0
            m_dblPrc_Or_Disc_2 = 0
            m_intMinimum_Qty_3 = 0
            m_dblPrc_Or_Disc_3 = 0
            m_intMinimum_Qty_4 = 0
            m_dblPrc_Or_Disc_4 = 0
            m_intMinimum_Qty_5 = 0
            m_dblPrc_Or_Disc_5 = 0
            m_intMinimum_Qty_6 = 0
            m_dblPrc_Or_Disc_6 = 0
            m_intMinimum_Qty_7 = 0
            m_dblPrc_Or_Disc_7 = 0
            m_intMinimum_Qty_8 = 0
            m_dblPrc_Or_Disc_8 = 0
            m_intMinimum_Qty_9 = 0
            m_dblPrc_Or_Disc_9 = 0

            m_blnOEI_Warning = False
            m_blnMacola_Import = False

            m_strProgram_Name = String.Empty

            m_strEQP_Type = String.Empty
            m_intEQP_Plus_Pct = 0

            m_strAutorized_By = String.Empty

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Load()

        Try

            Dim dt As DataTable

            Dim strSql As String = _
            "SELECT * " & _
            "FROM   MDB_OE_PRICE " & _
            "WHERE  OE_PRICE_ID = " & m_intOE_Price_ID.ToString

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            Else
                Call Init()
            End If

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            With pdrRow

                If Not (.Item("OE_Price_ID").Equals(DBNull.Value)) Then m_intOE_Price_ID = .Item("OE_Price_ID")
                If Not (.Item("Item_CD").Equals(DBNull.Value)) Then m_strItem_CD = .Item("Item_CD").ToString
                If Not (.Item("Curr_CD").Equals(DBNull.Value)) Then m_strCurr_CD = .Item("Curr_CD").ToString
                If Not (.Item("CD_TP").Equals(DBNull.Value)) Then m_intCD_TP = .Item("CD_TP")
                If Not (.Item("CFG_Spc_Prc_ID").Equals(DBNull.Value)) Then m_intCFG_Spc_Prc_ID = .Item("CFG_Spc_Prc_ID")
                If Not (.Item("CD_TP_1_Cust_No").Equals(DBNull.Value)) Then m_strCD_TP_1_Cust_No = .Item("CD_TP_1_Cust_No").ToString
                If Not (.Item("Cd_Tp_2_Prod_Cat").Equals(DBNull.Value)) Then m_strCd_Tp_2_Prod_Cat = .Item("Cd_Tp_2_Prod_Cat").ToString
                If Not (.Item("Cd_Tp_3_Cust_Type").Equals(DBNull.Value)) Then m_strCd_Tp_3_Cust_Type = .Item("Cd_Tp_3_Cust_Type").ToString
                If Not (.Item("Start_Date").Equals(DBNull.Value)) Then m_datStart_Date = .Item("Start_Date")
                If Not (.Item("End_Date").Equals(DBNull.Value)) Then m_datEnd_Date = .Item("End_Date")
                If Not (.Item("Cd_Prc_Basis").Equals(DBNull.Value)) Then m_strCd_Prc_Basis = .Item("Cd_Prc_Basis").ToString
                If Not (.Item("Minimum_Qty_1").Equals(DBNull.Value)) Then m_intMinimum_Qty_1 = .Item("Minimum_Qty_1")
                If Not (.Item("Prc_Or_Disc_1").Equals(DBNull.Value)) Then m_dblPrc_Or_Disc_1 = .Item("Prc_Or_Disc_1")
                If Not (.Item("Minimum_Qty_2").Equals(DBNull.Value)) Then m_intMinimum_Qty_2 = .Item("Minimum_Qty_2")
                If Not (.Item("Prc_Or_Disc_2").Equals(DBNull.Value)) Then m_dblPrc_Or_Disc_2 = .Item("Prc_Or_Disc_2")
                If Not (.Item("Minimum_Qty_3").Equals(DBNull.Value)) Then m_intMinimum_Qty_3 = .Item("Minimum_Qty_3")
                If Not (.Item("Prc_Or_Disc_3").Equals(DBNull.Value)) Then m_dblPrc_Or_Disc_3 = .Item("Prc_Or_Disc_3")
                If Not (.Item("Minimum_Qty_4").Equals(DBNull.Value)) Then m_intMinimum_Qty_4 = .Item("Minimum_Qty_4")
                If Not (.Item("Prc_Or_Disc_4").Equals(DBNull.Value)) Then m_dblPrc_Or_Disc_4 = .Item("Prc_Or_Disc_4")
                If Not (.Item("Minimum_Qty_5").Equals(DBNull.Value)) Then m_intMinimum_Qty_5 = .Item("Minimum_Qty_5")
                If Not (.Item("Prc_Or_Disc_5").Equals(DBNull.Value)) Then m_dblPrc_Or_Disc_5 = .Item("Prc_Or_Disc_5")
                If Not (.Item("Minimum_Qty_6").Equals(DBNull.Value)) Then m_intMinimum_Qty_6 = .Item("Minimum_Qty_6")
                If Not (.Item("Prc_Or_Disc_6").Equals(DBNull.Value)) Then m_dblPrc_Or_Disc_6 = .Item("Prc_Or_Disc_6")
                If Not (.Item("Minimum_Qty_7").Equals(DBNull.Value)) Then m_intMinimum_Qty_7 = .Item("Minimum_Qty_7")
                If Not (.Item("Prc_Or_Disc_7").Equals(DBNull.Value)) Then m_dblPrc_Or_Disc_7 = .Item("Prc_Or_Disc_7")
                If Not (.Item("Minimum_Qty_8").Equals(DBNull.Value)) Then m_intMinimum_Qty_8 = .Item("Minimum_Qty_8")
                If Not (.Item("Prc_Or_Disc_8").Equals(DBNull.Value)) Then m_dblPrc_Or_Disc_8 = .Item("Prc_Or_Disc_8")
                If Not (.Item("Minimum_Qty_9").Equals(DBNull.Value)) Then m_intMinimum_Qty_9 = .Item("Minimum_Qty_9")
                If Not (.Item("Prc_Or_Disc_9").Equals(DBNull.Value)) Then m_dblPrc_Or_Disc_9 = .Item("Prc_Or_Disc_9")
                If Not (.Item("OEI_Warning").Equals(DBNull.Value)) Then m_blnOEI_Warning = .Item("OEI_Warning")
                If Not (.Item("Macola_Import").Equals(DBNull.Value)) Then m_blnMacola_Import = .Item("Macola_Import")
                If Not (.Item("Program_Name").Equals(DBNull.Value)) Then m_strProgram_Name = .Item("Program_Name").ToString
                If Not (.Item("EQP_Type").Equals(DBNull.Value)) Then m_strEQP_Type = .Item("EQP_Type")
                If Not (.Item("EQP_Plus_Pct").Equals(DBNull.Value)) Then m_intEQP_Plus_Pct = .Item("EQP_Plus_Pct")
                If Not (.Item("Autorized_By").Equals(DBNull.Value)) Then m_strAutorized_By = .Item("Autorized_By").ToString
                If Not (.Item("When_Req").Equals(DBNull.Value)) Then m_blnWhen_Req = .Item("When_Req")
                If Not (.Item("Comments").Equals(DBNull.Value)) Then m_strComments = .Item("Comments").ToString

            End With

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Sub Reload()

    '    Call Reset()
    '    Call Load()

    'End Sub

    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("Item_CD") = m_strItem_CD
            pdrRow.Item("Curr_CD") = m_strCurr_CD

            ' Never save m_intCharge_Usage_ID as it is the incremental ID
            ' pdrRow.Item("OE_Price_ID") = m_intOE_Price_ID

            pdrRow.Item("CD_TP") = m_intCD_TP
            pdrRow.Item("CFG_Spc_Prc_ID") = m_intCFG_Spc_Prc_ID
            pdrRow.Item("CD_TP_1_Cust_No") = m_strCD_TP_1_Cust_No
            pdrRow.Item("Cd_Tp_2_Prod_Cat") = m_strCd_Tp_2_Prod_Cat
            pdrRow.Item("Cd_Tp_3_Cust_Type") = m_strCd_Tp_3_Cust_Type

            If m_datStart_Date.Year <> 1 Then pdrRow.Item("Start_Date") = m_datStart_Date
            If m_datEnd_Date.Year <> 1 Then pdrRow.Item("End_Date") = m_datEnd_Date

            pdrRow.Item("Cd_Prc_Basis") = m_strCd_Prc_Basis

            pdrRow.Item("Minimum_Qty_1") = m_intMinimum_Qty_1
            pdrRow.Item("Prc_Or_Disc_1") = m_dblPrc_Or_Disc_1
            pdrRow.Item("Minimum_Qty_2") = m_intMinimum_Qty_2
            pdrRow.Item("Prc_Or_Disc_2") = m_dblPrc_Or_Disc_2
            pdrRow.Item("Minimum_Qty_3") = m_intMinimum_Qty_3
            pdrRow.Item("Prc_Or_Disc_3") = m_dblPrc_Or_Disc_3
            pdrRow.Item("Minimum_Qty_4") = m_intMinimum_Qty_4
            pdrRow.Item("Prc_Or_Disc_4") = m_dblPrc_Or_Disc_4
            pdrRow.Item("Minimum_Qty_5") = m_intMinimum_Qty_5
            pdrRow.Item("Prc_Or_Disc_5") = m_dblPrc_Or_Disc_5
            pdrRow.Item("Minimum_Qty_6") = m_intMinimum_Qty_6
            pdrRow.Item("Prc_Or_Disc_6") = m_dblPrc_Or_Disc_6
            pdrRow.Item("Minimum_Qty_7") = m_intMinimum_Qty_7
            pdrRow.Item("Prc_Or_Disc_7") = m_dblPrc_Or_Disc_7
            pdrRow.Item("Minimum_Qty_8") = m_intMinimum_Qty_8
            pdrRow.Item("Prc_Or_Disc_8") = m_dblPrc_Or_Disc_8
            pdrRow.Item("Minimum_Qty_9") = m_intMinimum_Qty_9
            pdrRow.Item("Prc_Or_Disc_9") = m_dblPrc_Or_Disc_9

            pdrRow.Item("OEI_Warning") = m_blnOEI_Warning
            pdrRow.Item("Macola_Import") = m_blnMacola_Import

            pdrRow.Item("Program_Name") = m_strProgram_Name

            pdrRow.Item("EQP_Type") = m_strEQP_Type
            pdrRow.Item("EQP_Plus_Pct") = m_intEQP_Plus_Pct

            pdrRow.Item("When_Req") = m_blnWhen_Req
            pdrRow.Item("Autorized_By") = m_strAutorized_By

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    Public Sub Delete()

        Try
            If m_intOE_Price_ID = 0 Then Exit Sub

            Dim dt As New DataTable

            Dim strSql As String = _
            "DELETE FROM MDB_OE_PRICE " & _
            "WHERE  OE_PRICE_ID = " & m_intOE_Price_ID & " "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Sub Reset()

    '    Try

    '        m_intOE_Price_ID = 0

    '        m_strItem_CD = String.Empty
    '        m_strCurr_CD = String.Empty
    '        m_intCD_TP = 0
    '        m_intCFG_Spc_Prc_ID = 0
    '        m_strCD_TP_1_Cust_No = String.Empty
    '        m_strCd_Tp_2_Prod_Cat = String.Empty
    '        m_strCd_Tp_3_Cust_Type = String.Empty
    '        m_datStart_Date = NoDate()
    '        m_datEnd_Date = NoDate()
    '        m_strCd_Prc_Basis = String.Empty
    '        m_intMinimum_Qty_1 = 0
    '        m_dblPrc_Or_Disc_1 = 0
    '        m_intMinimum_Qty_2 = 0
    '        m_dblPrc_Or_Disc_2 = 0
    '        m_intMinimum_Qty_3 = 0
    '        m_dblPrc_Or_Disc_3 = 0
    '        m_intMinimum_Qty_4 = 0
    '        m_dblPrc_Or_Disc_4 = 0
    '        m_intMinimum_Qty_5 = 0
    '        m_dblPrc_Or_Disc_5 = 0
    '        m_intMinimum_Qty_6 = 0
    '        m_dblPrc_Or_Disc_6 = 0
    '        m_intMinimum_Qty_7 = 0
    '        m_dblPrc_Or_Disc_7 = 0
    '        m_intMinimum_Qty_8 = 0
    '        m_dblPrc_Or_Disc_8 = 0
    '        m_intMinimum_Qty_9 = 0
    '        m_dblPrc_Or_Disc_9 = 0
    '        m_blnOEI_Warning = False
    '        m_blnMacola_Import = False
    '        m_strProgram_Name = String.Empty
    '        m_strEQP_Type = String.Empty
    '        m_intEQP_Plus_Pct = 0
    '        m_strAutorized_By = String.Empty
    '        m_blnWhen_Req = False
    '        m_strComments = String.Empty

    '    Catch er As Exception
    '        MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Public Sub Save()

        Try

            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String = _
            "SELECT * " & _
            "FROM   MDB_OE_PRICE " & _
            "WHERE  OE_PRICE_ID = " & m_intOE_Price_ID & " "

            dt = db.DataTable(strSql)

            If dt.Rows.Count = 0 Then
                drRow = dt.NewRow()
            Else
                drRow = dt.Rows(0)
            End If
            'drRow = IIf(dt.Rows.Count = 0, dt.NewRow(), dt.Rows(0))

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

            If m_intOE_Price_ID = 0 Then m_intOE_Price_ID = db.DataTable("SELECT MAX(OE_Price_ID) AS Max_OE_Price_ID FROM MDB_OE_PRICE WITH (Nolock)").Rows(0).Item("Max_OE_Price_ID")

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public ReadOnly Properties #######################################"

    Public ReadOnly Property Charge_Exist(ByVal pstrCharge_Cd As String) As Boolean
        Get
            Charge_Exist = False
            Try
                Dim strSql As String = _
                "SELECT Charge_ID " & _
                "FROM   MDB_CFG_CHARGE WITH (Nolock) " & _
                "WHERE  Charge_CD = '" & SqlCompliantString(pstrCharge_Cd) & "' "

                Dim dt As DataTable = db.DataTable(strSql)

                Charge_Exist = (dt.Rows.Count > 0)

            Catch er As Exception
                MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Get
    End Property
    Public ReadOnly Property Dirty() As Boolean
        Get
            Dirty = m_blnDirty
        End Get
        'Set(ByVal value As Boolean)
        '    m_blnDirty = value
        'End Set
    End Property

#End Region

#Region "Public properties ################################################"

    Public Property OE_Price_ID() As Integer
        Get
            OE_Price_ID = m_intOE_Price_ID
        End Get
        Set(ByVal value As Integer)
            m_intOE_Price_ID = value
        End Set
    End Property
    Public Property Item_CD() As String
        Get
            Item_CD = m_strItem_CD
        End Get
        Set(ByVal value As String)
            If Not (m_strItem_CD = value) Then
                m_strItem_CD = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Curr_CD() As String
        Get
            Curr_CD = m_strCurr_CD
        End Get
        Set(ByVal value As String)
            If Not (m_strCurr_CD = value) Then
                m_strCurr_CD = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property CD_TP() As Integer
        Get
            CD_TP = m_intCD_TP
        End Get
        Set(ByVal value As Integer)
            If Not (m_intCD_TP = value) Then
                m_intCD_TP = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property CFG_Spc_Prc_ID() As Integer
        Get
            CFG_Spc_Prc_ID = m_intCFG_Spc_Prc_ID
        End Get
        Set(ByVal value As Integer)
            If Not (m_intCFG_Spc_Prc_ID = value) Then
                m_intCFG_Spc_Prc_ID = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Cus_No() As String
        Get
            Cus_No = m_strCD_TP_1_Cust_No
        End Get
        Set(ByVal value As String)
            If Not (m_strCD_TP_1_Cust_No = value) Then
                m_strCD_TP_1_Cust_No = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prod_Cat() As String
        Get
            Prod_Cat = m_strCd_Tp_2_Prod_Cat
        End Get
        Set(ByVal value As String)
            If Not (m_strCd_Tp_2_Prod_Cat = value) Then
                m_strCd_Tp_2_Prod_Cat = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Cust_Type() As String
        Get
            Cust_Type = m_strCd_Tp_3_Cust_Type
        End Get
        Set(ByVal value As String)
            If Not (m_strCd_Tp_3_Cust_Type = value) Then
                m_strCd_Tp_3_Cust_Type = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Start_Date() As Date
        Get
            Start_Date = m_datStart_Date
        End Get
        Set(ByVal value As Date)
            If Not (m_datStart_Date = value) Then
                m_datStart_Date = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property End_Date() As Date
        Get
            End_Date = m_datEnd_Date
        End Get
        Set(ByVal value As Date)
            If Not (m_datEnd_Date = value) Then
                m_datEnd_Date = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Cd_Prc_Basis() As String
        Get
            Cd_Prc_Basis = m_strCd_Prc_Basis
        End Get
        Set(ByVal value As String)
            If Not (m_strCd_Prc_Basis = value) Then
                m_strCd_Prc_Basis = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Minimum_Qty_1() As Integer
        Get
            Minimum_Qty_1 = m_intMinimum_Qty_1
        End Get
        Set(ByVal value As Integer)
            If Not (m_intMinimum_Qty_1 = value) Then
                m_intMinimum_Qty_1 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prc_Or_Disc_1() As Double
        Get
            Prc_Or_Disc_1 = m_dblPrc_Or_Disc_1
        End Get
        Set(ByVal value As Double)
            If Not (m_dblPrc_Or_Disc_1 = value) Then
                m_dblPrc_Or_Disc_1 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Minimum_Qty_2() As Integer
        Get
            Minimum_Qty_2 = m_intMinimum_Qty_2
        End Get
        Set(ByVal value As Integer)
            If Not (m_intMinimum_Qty_2 = value) Then
                m_intMinimum_Qty_2 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prc_Or_Disc_2() As Double
        Get
            Prc_Or_Disc_2 = m_dblPrc_Or_Disc_2
        End Get
        Set(ByVal value As Double)
            If Not (m_dblPrc_Or_Disc_2 = value) Then
                m_dblPrc_Or_Disc_2 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Minimum_Qty_3() As Integer
        Get
            Minimum_Qty_3 = m_intMinimum_Qty_3
        End Get
        Set(ByVal value As Integer)
            If Not (m_intMinimum_Qty_3 = value) Then
                m_intMinimum_Qty_3 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prc_Or_Disc_3() As Double
        Get
            Prc_Or_Disc_3 = m_dblPrc_Or_Disc_3
        End Get
        Set(ByVal value As Double)
            If Not (m_dblPrc_Or_Disc_3 = value) Then
                m_dblPrc_Or_Disc_3 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Minimum_Qty_4() As Integer
        Get
            Minimum_Qty_4 = m_intMinimum_Qty_4
        End Get
        Set(ByVal value As Integer)
            If Not (m_intMinimum_Qty_4 = value) Then
                m_intMinimum_Qty_4 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prc_Or_Disc_4() As Double
        Get
            Prc_Or_Disc_4 = m_dblPrc_Or_Disc_4
        End Get
        Set(ByVal value As Double)
            If Not (m_dblPrc_Or_Disc_4 = value) Then
                m_dblPrc_Or_Disc_4 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Minimum_Qty_5() As Integer
        Get
            Minimum_Qty_5 = m_intMinimum_Qty_5
        End Get
        Set(ByVal value As Integer)
            If Not (m_intMinimum_Qty_5 = value) Then
                m_intMinimum_Qty_5 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prc_Or_Disc_5() As Double
        Get
            Prc_Or_Disc_5 = m_dblPrc_Or_Disc_5
        End Get
        Set(ByVal value As Double)
            If Not (m_dblPrc_Or_Disc_5 = value) Then
                m_dblPrc_Or_Disc_5 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Minimum_Qty_6() As Integer
        Get
            Minimum_Qty_6 = m_intMinimum_Qty_6
        End Get
        Set(ByVal value As Integer)
            If Not (m_intMinimum_Qty_6 = value) Then
                m_intMinimum_Qty_6 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prc_Or_Disc_6() As Double
        Get
            Prc_Or_Disc_6 = m_dblPrc_Or_Disc_6
        End Get
        Set(ByVal value As Double)
            If Not (m_dblPrc_Or_Disc_6 = value) Then
                m_dblPrc_Or_Disc_6 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Minimum_Qty_7() As Integer
        Get
            Minimum_Qty_7 = m_intMinimum_Qty_7
        End Get
        Set(ByVal value As Integer)
            If Not (m_intMinimum_Qty_7 = value) Then
                m_intMinimum_Qty_7 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prc_Or_Disc_7() As Double
        Get
            Prc_Or_Disc_7 = m_dblPrc_Or_Disc_7
        End Get
        Set(ByVal value As Double)
            If Not (m_dblPrc_Or_Disc_7 = value) Then
                m_dblPrc_Or_Disc_7 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Minimum_Qty_8() As Integer
        Get
            Minimum_Qty_8 = m_intMinimum_Qty_8
        End Get
        Set(ByVal value As Integer)
            If Not (m_intMinimum_Qty_8 = value) Then
                m_intMinimum_Qty_8 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prc_Or_Disc_8() As Double
        Get
            Prc_Or_Disc_8 = m_dblPrc_Or_Disc_8
        End Get
        Set(ByVal value As Double)
            If Not (m_dblPrc_Or_Disc_8 = value) Then
                m_dblPrc_Or_Disc_8 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Minimum_Qty_9() As Integer
        Get
            Minimum_Qty_9 = m_intMinimum_Qty_9
        End Get
        Set(ByVal value As Integer)
            If Not (m_intMinimum_Qty_9 = value) Then
                m_intMinimum_Qty_9 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Prc_Or_Disc_9() As Double
        Get
            Prc_Or_Disc_9 = m_dblPrc_Or_Disc_9
        End Get
        Set(ByVal value As Double)
            If Not (m_dblPrc_Or_Disc_9 = value) Then
                m_dblPrc_Or_Disc_9 = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property OEI_Warning() As Boolean
        Get
            OEI_Warning = m_blnOEI_Warning
        End Get
        Set(ByVal value As Boolean)
            If Not (m_blnOEI_Warning = value) Then
                m_blnOEI_Warning = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Macola_Import() As Boolean
        Get
            Macola_Import = m_blnMacola_Import
        End Get
        Set(ByVal value As Boolean)
            If Not (m_blnMacola_Import = value) Then
                m_blnMacola_Import = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Program_Name() As String
        Get
            Program_Name = m_strProgram_Name
        End Get
        Set(ByVal value As String)
            If Not (m_strProgram_Name = value) Then
                m_strProgram_Name = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property EQP_Type() As String
        Get
            EQP_Type = m_strEQP_Type
        End Get
        Set(ByVal value As String)
            If Not (m_strEQP_Type = value) Then
                m_strEQP_Type = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property EQP_Plus_Pct() As Integer
        Get
            EQP_Plus_Pct = m_intEQP_Plus_Pct
        End Get
        Set(ByVal value As Integer)
            If Not (m_intEQP_Plus_Pct = value) Then
                m_intEQP_Plus_Pct = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property When_Req() As Boolean
        Get
            When_Req = m_blnWhen_Req
        End Get
        Set(ByVal value As Boolean)
            If Not (m_blnWhen_Req = value) Then
                m_blnWhen_Req = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Autorized_By() As String
        Get
            Autorized_By = m_strAutorized_By
        End Get
        Set(ByVal value As String)
            If Not (m_strAutorized_By = value) Then
                m_strAutorized_By = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Comments() As String
        Get
            Comments = m_strComments
        End Get
        Set(ByVal value As String)
            If Not (m_strComments = value) Then
                m_strComments = value
                m_blnDirty = True
            End If
        End Set
    End Property

#End Region

End Class
