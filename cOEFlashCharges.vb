Public Class cOEFlashCharges

    ' Unexposed properties
    Private m_ClassName As String = "cOEFlashCharges"
    Private db As New cDBA

    ' Exposed properties
    Private m_strCus_No As String

    Private m_intCharge_Usage_ID As Integer
    Private m_intCharge_ID As Integer
    Private m_strCharge_CD As String
    Private m_strApply_To_Ship_To As String ' 40
    Private m_strApply_To_Program As String ' 60

    ' Replace for int Usage 0 - Nothing, 1 - Always 2 - Never
    Private m_blnAlways_Use As Boolean
    Private m_blnNever_Use As Boolean

    ' Replace for use when qty or amt > and < integer
    Private m_intWhen_Qty_LT As Integer ' MUST CHANGE IN DATABASE !
    Private m_intWhen_Qty_GT As Integer ' MUST CHANGE IN DATABASE !

    Private m_intWhen_Amt_LT As Integer ' MUST CHANGE IN DATABASE !
    Private m_intWhen_Amt_GT As Integer ' MUST CHANGE IN DATABASE !

    ' Dont use these flags, set them to correspondance.
    '[SEND_EMAIL] [bit] NULL,
    '[EMAIL_TO] [varchar](128) NULL,

    Private m_blnNo_Charge As Boolean
    Private m_datCharge_From As Date
    Private m_datCharge_To As Date
    Private m_blnWhen_Req As Boolean
    Private m_blnBlind As Boolean
    Private m_strComments As String
    Private m_strAutorized_By As String ' MUST ADD TO DATABASE !

    Private m_blnDirty As Boolean = False

#Region "Public constructors ##############################################"

    Public Sub New()

        Call Init()

    End Sub

    Public Sub New(ByVal pstrCus_No As String)

        Call Init()
        m_strCus_No = pstrCus_No

    End Sub

    Public Sub New(ByVal pintCharge_Usage_ID As Integer)

        m_intCharge_Usage_ID = pintCharge_Usage_ID
        Call Load()

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_strCus_No = String.Empty

            m_intCharge_Usage_ID = 0
            m_intCharge_ID = 0
            m_strCharge_CD = String.Empty

            m_strApply_To_Ship_To = String.Empty
            m_strApply_To_Program = String.Empty

            ' Replace for int Usage 0 - Nothing, 1 - Always 2 - Never
            m_blnAlways_Use = False
            m_blnNever_Use = False

            ' Replace for use when qty or amt > and < integer
            m_intWhen_Qty_GT = 0
            m_intWhen_Amt_GT = 0
            m_intWhen_Qty_LT = 0
            m_intWhen_Amt_LT = 0

            ' Dont use these flags, set them to correspondance.
            '[SEND_EMAIL] [bit] NULL,
            '[EMAIL_TO] [varchar](128) NULL,

            m_blnNo_Charge = False
            m_datCharge_From = NoDate()
            m_datCharge_To = NoDate()
            m_blnWhen_Req = False
            m_blnBlind = False
            m_strComments = String.Empty
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
            "FROM   MDB_CFG_CHARGE_USAGE " & _
            "WHERE  CHARGE_USAGE_ID = " & m_intCharge_Usage_ID.ToString

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

                If Not (.Item("Cus_No").Equals(DBNull.Value)) Then m_strCus_No = .Item("Cus_No").ToString
                If Not (.Item("Charge_Usage_ID").Equals(DBNull.Value)) Then m_intCharge_Usage_ID = .Item("Charge_Usage_ID")
                If Not (.Item("Charge_ID").Equals(DBNull.Value)) Then m_intCharge_ID = .Item("Charge_ID")
                If Not (.Item("Charge_CD").Equals(DBNull.Value)) Then m_strCharge_CD = .Item("Charge_CD").ToString
                If Not (.Item("Apply_To_Ship_To").Equals(DBNull.Value)) Then m_strApply_To_Ship_To = .Item("Apply_To_Ship_To").ToString
                If Not (.Item("Apply_To_Program").Equals(DBNull.Value)) Then m_strApply_To_Program = .Item("Apply_To_Program").ToString

                If Not (.Item("Always_Use").Equals(DBNull.Value)) Then m_blnAlways_Use = .Item("Always_Use")
                If Not (.Item("Never_Use").Equals(DBNull.Value)) Then m_blnNever_Use = .Item("Never_Use")
                If Not (.Item("When_Qty_GT").Equals(DBNull.Value)) Then m_intWhen_Qty_GT = .Item("When_Qty_GT")
                If Not (.Item("When_Amt_GT").Equals(DBNull.Value)) Then m_intWhen_Amt_GT = .Item("When_Amt_GT")
                If Not (.Item("When_Qty_LT").Equals(DBNull.Value)) Then m_intWhen_Qty_LT = .Item("When_Qty_LT")
                If Not (.Item("When_Amt_LT").Equals(DBNull.Value)) Then m_intWhen_Amt_LT = .Item("When_Amt_LT")
                If Not (.Item("No_Charge").Equals(DBNull.Value)) Then m_blnNo_Charge = .Item("No_Charge")
                If Not (.Item("Charge_From").Equals(DBNull.Value)) Then m_datCharge_From = .Item("Charge_From")
                If Not (.Item("Charge_To").Equals(DBNull.Value)) Then m_datCharge_To = .Item("Charge_To")
                If Not (.Item("When_Req").Equals(DBNull.Value)) Then m_blnWhen_Req = .Item("When_Req")
                If Not (.Item("Blind").Equals(DBNull.Value)) Then m_blnBlind = .Item("Blind")
                If Not (.Item("Comments").Equals(DBNull.Value)) Then m_strComments = .Item("Comments")
                If Not (.Item("Autorized_By").Equals(DBNull.Value)) Then m_strAutorized_By = .Item("Autorized_By")

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
            pdrRow.Item("Cus_No") = m_strCus_No

            ' Never save m_intCharge_Usage_ID as it is the incremental ID
            ' pdrRow.Item("Charge_Usage_ID") = m_intCharge_Usage_ID

            pdrRow.Item("Charge_ID") = m_intCharge_ID
            pdrRow.Item("Charge_CD") = m_strCharge_CD
            pdrRow.Item("Apply_To_Ship_To") = m_strApply_To_Ship_To
            pdrRow.Item("Apply_To_Program") = m_strApply_To_Program
            pdrRow.Item("Always_Use") = m_blnAlways_Use
            pdrRow.Item("Never_Use") = m_blnNever_Use

            pdrRow.Item("When_Qty_LT") = m_intWhen_Qty_LT
            pdrRow.Item("When_Qty_GT") = m_intWhen_Qty_GT
            pdrRow.Item("When_Amt_LT") = m_intWhen_Amt_LT
            pdrRow.Item("When_Amt_GT") = m_intWhen_Amt_GT

            pdrRow.Item("No_Charge") = m_blnNo_Charge
            If m_datCharge_From.Year <> 1 Then
                pdrRow.Item("Charge_From") = m_datCharge_From.Date
            Else
                pdrRow.Item("Charge_From") = DBNull.Value
            End If

            If m_datCharge_To.Year <> 1 Then
                pdrRow.Item("Charge_To") = m_datCharge_To.Date
            Else
                pdrRow.Item("Charge_To") = DBNull.Value
            End If

            pdrRow.Item("When_Req") = m_blnWhen_Req
            pdrRow.Item("Blind") = m_blnBlind
            pdrRow.Item("Comments") = m_strComments
            pdrRow.Item("Autorized_By") = m_strAutorized_By

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    Public Sub Delete()

        Try
            If m_intCharge_Usage_ID = 0 Then Exit Sub

            Dim dt As New DataTable

            Dim strSql As String = _
            "DELETE FROM MDB_CFG_CHARGE_USAGE " & _
            "WHERE  CHARGE_USAGE_ID = " & m_intCharge_Usage_ID & " "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Sub Reset()

    '    Try
    '        m_intCharge_Usage_ID = 0
    '        m_intCharge_ID = 0
    '        m_strCharge_CD = String.Empty

    '        m_strApply_To_Ship_To = String.Empty
    '        m_strApply_To_Program = String.Empty

    '        ' Replace for int Usage 0 - Nothing, 1 - Always 2 - Never
    '        m_blnAlways_Use = False
    '        m_blnNever_Use = False

    '        ' Replace for use when qty or amt > and < integer
    '        m_intWhen_Qty_GT = 0
    '        m_intWhen_Amt_GT = 0
    '        m_intWhen_Qty_LT = 0
    '        m_intWhen_Amt_LT = 0

    '        ' Dont use these flags, set them to correspondance.
    '        '[SEND_EMAIL] [bit] NULL,
    '        '[EMAIL_TO] [varchar](128) NULL,

    '        m_blnNo_Charge = False
    '        m_datCharge_From = NoDate()
    '        m_datCharge_To = NoDate()
    '        m_blnWhen_Req = False
    '        m_blnBlind = False
    '        m_strComments = String.Empty
    '        m_strAutorized_By = String.Empty

    '    Catch er As Exception
    '        MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Public Sub Save()

        Try
            If Trim(m_strCharge_CD) = "" Then Exit Sub
            If Not m_blnDirty Then Exit Sub

            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String = _
            "SELECT * " & _
            "FROM   MDB_CFG_CHARGE_USAGE " & _
            "WHERE  CHARGE_USAGE_ID = " & m_intCharge_Usage_ID & " "

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

            If m_intCharge_Usage_ID = 0 Then m_intCharge_Usage_ID = db.DataTable("SELECT MAX(Charge_Usage_ID) AS Max_Charge_Usage_ID FROM MDB_CFG_CHARGE_USAGE WITH (Nolock)").Rows(0).Item("Max_Charge_Usage_ID")

            m_blnDirty = False

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

    Public Property Cus_No() As String
        Get
            Cus_No = m_strCus_No
        End Get
        Set(ByVal value As String)
            m_strCus_No = value
        End Set
    End Property
    Public Property Charge_Usage_ID() As Integer
        Get
            Charge_Usage_ID = m_intCharge_Usage_ID
        End Get
        Set(ByVal value As Integer)
            m_intCharge_Usage_ID = value
            'm_blnDirty = True
        End Set
    End Property
    Public Property Charge_ID() As Integer
        Get
            Charge_ID = m_intCharge_ID
        End Get
        Set(ByVal value As Integer)
            m_intCharge_ID = value
            'm_blnDirty = True
        End Set
    End Property
    Public Property Charge_CD() As String
        Get
            Charge_CD = m_strCharge_CD
        End Get
        Set(ByVal value As String)
            If Not (m_strCharge_CD = value) Then
                m_strCharge_CD = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Apply_To_Ship_To() As String
        Get
            Apply_To_Ship_To = m_strApply_To_Ship_To
        End Get
        Set(ByVal value As String)
            If Not (m_strApply_To_Ship_To = value) Then
                m_strApply_To_Ship_To = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Apply_To_Program() As String
        Get
            Apply_To_Program = m_strApply_To_Program
        End Get
        Set(ByVal value As String)
            If Not (m_strApply_To_Program = value) Then
                m_strApply_To_Program = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Always_Use() As Boolean
        Get
            Always_Use = m_blnAlways_Use
        End Get
        Set(ByVal value As Boolean)
            If Not (m_blnAlways_Use = value) Then
                m_blnAlways_Use = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Never_Use() As Boolean
        Get
            Never_Use = m_blnNever_Use
        End Get
        Set(ByVal value As Boolean)
            If Not (m_blnNever_Use = value) Then
                m_blnNever_Use = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property When_Qty_LT() As Integer
        Get
            When_Qty_LT = m_intWhen_Qty_LT
        End Get
        Set(ByVal value As Integer)
            If Not (m_intWhen_Qty_LT = value) Then
                m_intWhen_Qty_LT = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property When_Qty_GT() As Integer
        Get
            When_Qty_GT = m_intWhen_Qty_GT
        End Get
        Set(ByVal value As Integer)
            If Not (m_intWhen_Qty_GT = value) Then
                m_intWhen_Qty_GT = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property When_Amt_LT() As Integer
        Get
            When_Amt_LT = m_intWhen_Amt_LT
        End Get
        Set(ByVal value As Integer)
            If Not (m_intWhen_Amt_LT = value) Then
                m_intWhen_Amt_LT = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property When_Amt_GT() As Integer
        Get
            When_Amt_GT = m_intWhen_Amt_GT
        End Get
        Set(ByVal value As Integer)
            If Not (m_intWhen_Amt_GT = value) Then
                m_intWhen_Amt_GT = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property No_Charge() As Boolean
        Get
            No_Charge = m_blnNo_Charge
        End Get
        Set(ByVal value As Boolean)
            If Not (m_blnNo_Charge = value) Then
                m_blnNo_Charge = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Charge_From() As Date
        Get
            Charge_From = m_datCharge_From
        End Get
        Set(ByVal value As Date)
            If Not (m_datCharge_From = value) Then
                m_datCharge_From = value
                m_blnDirty = True
            End If
        End Set
    End Property
    Public Property Charge_To() As Date
        Get
            Charge_To = m_datCharge_To
        End Get
        Set(ByVal value As Date)
            If Not (m_datCharge_To = value) Then
                m_datCharge_To = value
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
    Public Property Blind() As Boolean
        Get
            Blind = m_blnBlind
        End Get
        Set(ByVal value As Boolean)
            If Not (m_blnBlind = value) Then
                m_blnBlind = value
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

#End Region

End Class



