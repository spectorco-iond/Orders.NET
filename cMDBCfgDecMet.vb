Public Class cMDBCfgDecMet

    Private m_Dec_Met_ID As Long
    Private m_Dec_Met_Cd As String
    Private m_Setup_Item_No As String
    Private m_RC_Item_No As String
    Private m_Add_Col_Item_No As String
    Private m_RC_Col_Item_No As String
    Private m_Add_Loc_Item_No As String
    Private m_RC_Loc_Item_No As String
    Private m_RC_Oxi_Item_No As String
    Private m_RC_2ndOpt_Item_No As String
    Private m_RC_2ndLoc_Item_No As String
    Private m_Perso_Item_No As String
    Private m_Setup_Prc As Double
    Private m_Setup_Net As Double
    Private m_RC_Prc As Double
    Private m_RC_Net As Double
    Private m_Loc_Qty As Long
    Private m_Col_Qty As Long
    Private m_Allow_Add_Loc_ As Boolean
    Private m_Allow_Add_Col_ As Boolean
    Private m_Add_Loc_Setup_Prc As Double
    Private m_Add_Loc_Setup_Net As Double
    Private m_Add_Loc_RC_Prc As Double
    Private m_Add_Loc_RC_Net As Double
    Private m_Add_Col_Setup_Prc As Double
    Private m_Add_Col_Setup_Net As Double
    Private m_Add_Col_RC_Prc As Double
    Private m_Add_Col_RC_Net As Double
    Private m_User_ID As String
    Private m_Create_TS As Date

#Region "Public constructors ##############################################"

    'Public Sub New()
    Public Sub New()

        Call Init()

    End Sub

    'Public Sub New(ByVal pID As Integer)
    Public Sub New(ByVal pDec_Met_ID As Integer)

        Try

            Call Init()

            m_Dec_Met_ID = pDec_Met_ID

            Call Load(pDec_Met_ID)

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_Dec_Met_ID = 0
            'm_Ord_Guid = String.Empty

            Call Reset()

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Load is made private, it is loaded automatically on New.
    Private Sub Load(ByVal pDec_Met_ID As Integer)

        Try
            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   MDB_CFG_DEC_MET C WITH (Nolock) " & _
            "WHERE  Dec_Met_ID = " & pDec_Met_ID & " "

            dt = db.DataTable(strSql)

            Call LoadLine(dt.Rows(0))

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Function Load(ByVal pstrOrd_Guid As String) As DataTable

    '    Load = New DataTable

    '    Try
    '        Dim db As New cDBA
    '        Dim dt As DataTable

    '        Dim strSql As String

    '        strSql = _
    '        "SELECT     * " & _
    '        "FROM       OEI_OrdHdr H WITH (Nolock) " & _
    '        "WHERE      Ord_Guid = '" & pstrOrd_Guid & "' "

    '        dt = db.DataTable(strSql)
    '        m_OEI_Ord_No = dt.Rows(0).Item("Oei_Ord_No")

    '        strSql = _
    '        "SELECT     S.ID, S.Ord_Guid, S.Cmt, S.Line_Seq_No, H.OEI_Ord_No, S.Ord_No " & _
    '        "FROM       OEI_LinCmt S WITH (Nolock) " & _
    '        "INNER JOIN OEI_OrdHdr H WITH (Nolock) ON H.Ord_Guid = S.Ord_Guid " & _
    '        "WHERE      H.Ord_Guid = '" & pstrOrd_Guid & "' " & _
    '        "ORDER BY   Line_Seq_No "

    '        Load = db.DataTable(strSql)

    '    Catch er As Exception
    '        MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Function

    ' Load is made private, it is loaded automatically on New.
    ' Loads the Comment fields from the DataRow record
    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            If Not (pdrRow.Item("Dec_Met_ID").Equals(DBNull.Value)) Then m_Dec_Met_ID = pdrRow.Item("Dec_Met_ID")
            If Not (pdrRow.Item("Dec_Met_Cd").Equals(DBNull.Value)) Then m_Dec_Met_Cd = pdrRow.Item("Dec_Met_Cd")
            If Not (pdrRow.Item("Setup_Item_No").Equals(DBNull.Value)) Then m_Setup_Item_No = pdrRow.Item("Setup_Item_No")
            If Not (pdrRow.Item("RC_Item_No").Equals(DBNull.Value)) Then m_RC_Item_No = pdrRow.Item("RC_Item_No").ToString
            If Not (pdrRow.Item("Add_Col_Item_No").Equals(DBNull.Value)) Then m_Add_Col_Item_No = pdrRow.Item("Add_Col_Item_No")
            If Not (pdrRow.Item("RC_Col_Item_No").Equals(DBNull.Value)) Then m_RC_Col_Item_No = pdrRow.Item("RC_Col_Item_No")
            If Not (pdrRow.Item("Add_Loc_Item_No").Equals(DBNull.Value)) Then m_Add_Loc_Item_No = pdrRow.Item("Add_Loc_Item_No")
            If Not (pdrRow.Item("RC_Loc_Item_No").Equals(DBNull.Value)) Then m_RC_Loc_Item_No = pdrRow.Item("RC_Loc_Item_No")
            If Not (pdrRow.Item("RC_Oxi_Item_No").Equals(DBNull.Value)) Then m_RC_Oxi_Item_No = pdrRow.Item("RC_Oxi_Item_No")
            If Not (pdrRow.Item("RC_2ndOpt_Item_No").Equals(DBNull.Value)) Then m_RC_2ndOpt_Item_No = pdrRow.Item("RC_2ndOpt_Item_No")
            If Not (pdrRow.Item("RC_2ndLoc_Item_No").Equals(DBNull.Value)) Then m_RC_2ndLoc_Item_No = pdrRow.Item("RC_2ndLoc_Item_No")
            If Not (pdrRow.Item("Perso_Item_No").Equals(DBNull.Value)) Then m_Perso_Item_No = pdrRow.Item("Perso_Item_No")
            If Not (pdrRow.Item("Setup_Prc").Equals(DBNull.Value)) Then m_Setup_Prc = pdrRow.Item("Setup_Prc")
            If Not (pdrRow.Item("Setup_Net").Equals(DBNull.Value)) Then m_Setup_Net = pdrRow.Item("Setup_Net")
            If Not (pdrRow.Item("RC_Prc").Equals(DBNull.Value)) Then m_RC_Prc = pdrRow.Item("RC_Prc")
            If Not (pdrRow.Item("RC_Net").Equals(DBNull.Value)) Then m_RC_Net = pdrRow.Item("RC_Net")
            If Not (pdrRow.Item("Loc_Qty").Equals(DBNull.Value)) Then m_Loc_Qty = pdrRow.Item("Loc_Qty")
            If Not (pdrRow.Item("Col_Qty").Equals(DBNull.Value)) Then m_Col_Qty = pdrRow.Item("Col_Qty")
            If Not (pdrRow.Item("Allow_Add_Loc").Equals(DBNull.Value)) Then m_Allow_Add_Loc_ = IIf(pdrRow.Item("Allow_Add_Loc") = 1, True, False)
            If Not (pdrRow.Item("Allow_Add_Col").Equals(DBNull.Value)) Then m_Allow_Add_Col_ = IIf(pdrRow.Item("Allow_Add_Col") = 1, True, False)
            If Not (pdrRow.Item("Add_Loc_Setup_Prc").Equals(DBNull.Value)) Then m_Add_Loc_Setup_Prc = pdrRow.Item("Add_Loc_Setup_Prc")
            If Not (pdrRow.Item("Add_Loc_Setup_Net").Equals(DBNull.Value)) Then m_Add_Loc_Setup_Net = pdrRow.Item("Add_Loc_Setup_Net")
            If Not (pdrRow.Item("Add_Loc_RC_Prc").Equals(DBNull.Value)) Then m_Add_Loc_RC_Prc = pdrRow.Item("Add_Loc_RC_Prc")
            If Not (pdrRow.Item("Add_Loc_RC_Net").Equals(DBNull.Value)) Then m_Add_Loc_RC_Net = pdrRow.Item("Add_Loc_RC_Net")
            If Not (pdrRow.Item("Add_Col_Setup_Prc").Equals(DBNull.Value)) Then m_Add_Col_Setup_Prc = pdrRow.Item("Add_Col_Setup_Prc")
            If Not (pdrRow.Item("Add_Col_Setup_Net").Equals(DBNull.Value)) Then m_Add_Col_Setup_Net = pdrRow.Item("Add_Col_Setup_Net")
            If Not (pdrRow.Item("Add_Col_RC_Prc").Equals(DBNull.Value)) Then m_Add_Col_RC_Prc = pdrRow.Item("Add_Col_RC_Prc")
            If Not (pdrRow.Item("Add_Col_RC_Net").Equals(DBNull.Value)) Then m_Add_Col_RC_Net = pdrRow.Item("Add_Col_RC_Net")
            If Not (pdrRow.Item("User_ID").Equals(DBNull.Value)) Then m_User_ID = pdrRow.Item("User_ID")
            If Not (pdrRow.Item("Create_TS").Equals(DBNull.Value)) Then m_Create_TS = pdrRow.Item("Create_TS")

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Saves the Comment fields into the DataRow record
    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            'pdrRow.Item("Dec_Met_ID") = m_Dec_Met_ID
            pdrRow.Item("Dec_Met_Cd") = m_Dec_Met_Cd
            pdrRow.Item("Setup_Item_No") = m_Setup_Item_No
            pdrRow.Item("RC_Item_No") = m_RC_Item_No
            pdrRow.Item("Add_Col_Item_No") = m_Add_Col_Item_No
            pdrRow.Item("RC_Col_Item_No") = m_RC_Col_Item_No
            pdrRow.Item("Add_Loc_Item_No") = m_Add_Loc_Item_No
            pdrRow.Item("RC_Loc_Item_No") = m_RC_Loc_Item_No
            pdrRow.Item("RC_Oxi_Item_No") = m_RC_Oxi_Item_No
            pdrRow.Item("RC_2ndOpt_Item_No") = m_RC_2ndOpt_Item_No
            pdrRow.Item("RC_2ndLoc_Item_No") = m_RC_2ndLoc_Item_No
            pdrRow.Item("Perso_Item_No") = m_Perso_Item_No
            pdrRow.Item("Setup_Prc") = m_Setup_Prc
            pdrRow.Item("Setup_Net") = m_Setup_Net
            pdrRow.Item("RC_Prc") = m_RC_Prc
            pdrRow.Item("RC_Net") = m_RC_Net
            pdrRow.Item("Loc_Qty") = m_Loc_Qty
            pdrRow.Item("Col_Qty") = m_Col_Qty
            pdrRow.Item("Allow_Add_Loc_") = IIf(m_ALLOW_Add_Loc_, 1, 0)
            pdrRow.Item("Allow_Add_Col_") = IIf(m_ALLOW_Add_Col_, 1, 0)
            pdrRow.Item("Add_Loc_Setup_Prc") = m_Add_Loc_Setup_Prc
            pdrRow.Item("Add_Loc_Setup_Net") = m_Add_Loc_Setup_Net
            pdrRow.Item("Add_Loc_RC_Prc") = m_Add_Loc_RC_Prc
            pdrRow.Item("Add_Loc_RC_Net") = m_Add_Loc_RC_Net
            pdrRow.Item("Add_Col_Setup_Prc") = m_Add_Col_Setup_Prc
            pdrRow.Item("Add_Col_Setup_Net") = m_Add_Col_Setup_Net
            pdrRow.Item("Add_Col_RC_Prc") = m_Add_Col_RC_Prc
            pdrRow.Item("Add_Col_RC_Net") = m_Add_Col_RC_Net
            pdrRow.Item("User_ID") = m_User_ID
            pdrRow.Item("Create_TS") = m_Create_TS

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    ' Deletes the current Comment from the database, if it exists.
    Private Sub Delete()

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If m_Dec_Met_ID = 0 Then Exit Sub

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            strSql = _
            "DELETE FROM MDB_CFG_DEC_MET " & _
            "WHERE  Dec_Met_ID = " & m_Dec_Met_ID & " "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every non mandatory field of a Comment
    Public Sub Reset()

        ' Resets the locals to empty values, all but ID, Ord_No and Ord_Guid
        m_Dec_Met_Cd = String.Empty
        m_Setup_Item_No = String.Empty
        m_RC_Item_No = String.Empty
        m_Add_Col_Item_No = String.Empty
        m_RC_Col_Item_No = String.Empty
        m_Add_Loc_Item_No = String.Empty
        m_RC_Loc_Item_No = String.Empty
        m_RC_Oxi_Item_No = String.Empty
        m_RC_2ndOpt_Item_No = String.Empty
        m_RC_2ndLoc_Item_No = String.Empty
        m_Perso_Item_No = String.Empty
        m_Setup_Prc = 0
        m_Setup_Net = 0
        m_RC_Prc = 0
        m_RC_Net = 0
        m_Loc_Qty = 0
        m_Col_Qty = 0
        m_ALLOW_Add_Loc_ = False
        m_ALLOW_Add_Col_ = False
        m_Add_Loc_Setup_Prc = 0
        m_Add_Loc_Setup_Net = 0
        m_Add_Loc_RC_Prc = 0
        m_Add_Loc_RC_Net = 0
        m_Add_Col_Setup_Prc = 0
        m_Add_Col_Setup_Net = 0
        m_Add_Col_RC_Prc = 0
        m_Add_Col_RC_Net = 0
        m_User_ID = String.Empty
        m_Create_TS = Nothing

    End Sub

    ' Update the current Comment into the database, or creates it if not existing
    Public Sub Save()

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String

            strSql = _
            "SELECT * " & _
            "FROM   MDB_CFG_DEC_MET " & _
            "WHERE  Dec_Met_ID = " & m_Dec_Met_ID & " "

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

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property Add_Col_Item_No() As String
        Get
            Add_Col_Item_No = m_Add_Col_Item_No
        End Get
        Set(ByVal value As String)
            m_Add_Col_Item_No = value
        End Set
    End Property
    Public Property Add_Col_RC_Net() As Double
        Get
            Add_Col_RC_Net = m_Add_Col_RC_Net
        End Get
        Set(ByVal value As Double)
            m_Add_Col_RC_Net = value
        End Set
    End Property
    Public Property Add_Col_RC_Prc() As Double
        Get
            Add_Col_RC_Prc = m_Add_Col_RC_Prc
        End Get
        Set(ByVal value As Double)
            m_Add_Col_RC_Prc = value
        End Set
    End Property
    Public Property Add_Col_Setup_Net() As Double
        Get
            Add_Col_Setup_Net = m_Add_Col_Setup_Net
        End Get
        Set(ByVal value As Double)
            m_Add_Col_Setup_Net = value
        End Set
    End Property
    Public Property Add_Col_Setup_Prc() As Double
        Get
            Add_Col_Setup_Prc = m_Add_Col_Setup_Prc
        End Get
        Set(ByVal value As Double)
            m_Add_Col_Setup_Prc = value
        End Set
    End Property
    Public Property Add_Loc_Item_No() As String
        Get
            Add_Loc_Item_No = m_Add_Loc_Item_No
        End Get
        Set(ByVal value As String)
            m_Add_Loc_Item_No = value
        End Set
    End Property
    Public Property Add_Loc_RC_Net() As Double
        Get
            Add_Loc_RC_Net = m_Add_Loc_RC_Net
        End Get
        Set(ByVal value As Double)
            m_Add_Loc_RC_Net = value
        End Set
    End Property
    Public Property Add_Loc_RC_Prc() As Double
        Get
            Add_Loc_RC_Prc = m_Add_Loc_RC_Prc
        End Get
        Set(ByVal value As Double)
            m_Add_Loc_RC_Prc = value
        End Set
    End Property
    Public Property Add_Loc_Setup_Net() As Double
        Get
            Add_Loc_Setup_Net = m_Add_Loc_Setup_Net
        End Get
        Set(ByVal value As Double)
            m_Add_Loc_Setup_Net = value
        End Set
    End Property
    Public Property Add_Loc_Setup_Prc() As Double
        Get
            Add_Loc_Setup_Prc = m_Add_Loc_Setup_Prc
        End Get
        Set(ByVal value As Double)
            m_Add_Loc_Setup_Prc = value
        End Set
    End Property
    Public Property Allow_Add_Col_() As Boolean
        Get
            ALLOW_Add_Col_ = m_ALLOW_Add_Col_
        End Get
        Set(ByVal value As Boolean)
            m_ALLOW_Add_Col_ = value
        End Set
    End Property
    Public Property Allow_Add_Loc_() As Boolean
        Get
            ALLOW_Add_Loc_ = m_ALLOW_Add_Loc_
        End Get
        Set(ByVal value As Boolean)
            m_ALLOW_Add_Loc_ = value
        End Set
    End Property
    Public Property Col_Qty() As Long
        Get
            Col_Qty = m_Col_Qty
        End Get
        Set(ByVal value As Long)
            m_Col_Qty = value
        End Set
    End Property
    Public Property Create_TS() As Date
        Get
            Create_TS = m_Create_TS
        End Get
        Set(ByVal value As Date)
            m_Create_TS = value
        End Set
    End Property
    Public Property Dec_Met_Cd() As String
        Get
            Dec_Met_Cd = m_Dec_Met_Cd
        End Get
        Set(ByVal value As String)
            m_Dec_Met_Cd = value
        End Set
    End Property
    Public Property Dec_Met_ID() As Long
        Get
            Dec_Met_ID = m_Dec_Met_ID
        End Get
        Set(ByVal value As Long)
            m_Dec_Met_ID = value
        End Set
    End Property
    Public Property Loc_Qty() As Long
        Get
            Loc_Qty = m_Loc_Qty
        End Get
        Set(ByVal value As Long)
            m_Loc_Qty = value
        End Set
    End Property
    Public Property Perso_Item_No() As String
        Get
            Perso_Item_No = m_Perso_Item_No
        End Get
        Set(ByVal value As String)
            m_Perso_Item_No = value
        End Set
    End Property
    Public Property RC_2ndLoc_Item_No() As String
        Get
            RC_2ndLoc_Item_No = m_RC_2ndLoc_Item_No
        End Get
        Set(ByVal value As String)
            m_RC_2ndLoc_Item_No = value
        End Set
    End Property
    Public Property RC_2ndOpt_Item_No() As String
        Get
            RC_2ndOpt_Item_No = m_RC_2ndOpt_Item_No
        End Get
        Set(ByVal value As String)
            m_RC_2ndOpt_Item_No = value
        End Set
    End Property
    Public Property RC_Col_Item_No() As String
        Get
            RC_Col_Item_No = m_RC_Col_Item_No
        End Get
        Set(ByVal value As String)
            m_RC_Col_Item_No = value
        End Set
    End Property
    Public Property RC_Item_No() As String
        Get
            RC_Item_No = m_RC_Item_No
        End Get
        Set(ByVal value As String)
            m_RC_Item_No = value
        End Set
    End Property
    Public Property RC_Loc_Item_No() As String
        Get
            RC_Loc_Item_No = m_RC_Loc_Item_No
        End Get
        Set(ByVal value As String)
            m_RC_Loc_Item_No = value
        End Set
    End Property
    Public Property RC_Net() As Double
        Get
            RC_Net = m_RC_Net
        End Get
        Set(ByVal value As Double)
            m_RC_Net = value
        End Set
    End Property
    Public Property RC_Oxi_Item_No() As String
        Get
            RC_Oxi_Item_No = m_RC_Oxi_Item_No
        End Get
        Set(ByVal value As String)
            m_RC_Oxi_Item_No = value
        End Set
    End Property
    Public Property RC_Prc() As Double
        Get
            RC_Prc = m_RC_Prc
        End Get
        Set(ByVal value As Double)
            m_RC_Prc = value
        End Set
    End Property
    Public Property Setup_Item_No() As String
        Get
            Setup_Item_No = m_Setup_Item_No
        End Get
        Set(ByVal value As String)
            m_Setup_Item_No = value
        End Set
    End Property
    Public Property Setup_Net() As Double
        Get
            Setup_Net = m_Setup_Net
        End Get
        Set(ByVal value As Double)
            m_Setup_Net = value
        End Set
    End Property
    Public Property Setup_Prc() As Double
        Get
            Setup_Prc = m_Setup_Prc
        End Get
        Set(ByVal value As Double)
            m_Setup_Prc = value
        End Set
    End Property
    Public Property User_ID() As String
        Get
            User_ID = m_User_ID
        End Get
        Set(ByVal value As String)
            m_User_ID = value
        End Set
    End Property

#End Region

End Class
