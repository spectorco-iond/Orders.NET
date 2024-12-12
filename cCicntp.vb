Public Class cCicntp

    Private m_ID As Integer = 0
    Private m_cnt_id As String
    Private m_cmp_wwn As String
    Private m_cnt_f_name As String
    Private m_cnt_l_name As String
    Private m_cnt_m_name As String
    Private m_FullName As String
    Private m_Initials As String
    Private m_Gender As String
    Private m_predcode As String
    Private m_cnt_job_desc As String
    Private m_cnt_dept As String
    Private m_taalcode As String
    Private m_cnt_f_ext As String
    Private m_cnt_f_fax As String
    Private m_cnt_f_tel As String
    Private m_cnt_f_mobile As String
    Private m_cnt_email As String
    Private m_cnt_acc_man As Integer
    Private m_cnt_note As String
    Private m_active_y As Integer
    Private m_Administration As String
    Private m_res_id As Integer
    Private m_WebAccess As Integer
    Private m_Picture As String ' image
    Private m_PictureFileName As String
    Private m_syscreated As Date
    Private m_syscreator As Integer
    Private m_sysmodified As Date
    Private m_sysmodifier As Integer
    Private m_sysguid As String
    Private m_timestamp As String ' timestamp
    Private m_datefield1 As Date
    Private m_datefield2 As Date
    Private m_datefield3 As Date
    Private m_datefield4 As Date
    Private m_datefield5 As Date
    Private m_numberfield1 As Double
    Private m_numberfield2 As Double
    Private m_numberfield3 As Double
    Private m_numberfield4 As Double
    Private m_numberfield5 As Double
    Private m_YesNoField1 As Integer
    Private m_YesNoField2 As Integer
    Private m_YesNoField3 As Integer
    Private m_YesNoField4 As Integer
    Private m_YesNoField5 As Integer
    Private m_textfield1 As String
    Private m_textfield2 As String
    Private m_textfield3 As String
    Private m_textfield4 As String
    Private m_textfield5 As String
    Private m_cntp_Directory As String
    Private m_Division As Integer
    Private m_ContactChange As Boolean
    Private m_OldFileName As String

    Private m_DataTable As DataTable

#Region "Public constructors ##############################################"

    'Public Sub New()
    Public Sub New()

        Call Init()

    End Sub

    'Public Sub New(ByVal pID As Integer)
    Public Sub New(ByVal pID As Integer)

        Try

            Call Init()

            m_ID = pID

            Call Load(pID)

        Catch er As Exception
            MsgBox("Error in cCicntp." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_ID = 0
            'm_Ord_No = String.Empty
            'm_Ord_Guid = String.Empty

            Call Reset()

        Catch er As Exception
            MsgBox("Error in cCicntp." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Load is made private, it is loaded automatically on New.
    Private Sub Load(ByVal pintID As Integer)

        Try
            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "SELECT * " &
            "FROM   cicntp C WITH (Nolock) " &
            "WHERE ISNULL(C.active_y,0) = 1 And  ID = " & pintID & " "
            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            End If

        Catch er As Exception
            MsgBox("Error in cCicntp." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Load is made private, it is loaded automatically on New.
    Public Sub LoadRes_ID(ByVal pintID As Integer)

        Try
            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "SELECT * " &
            "FROM   cicntp C WITH (Nolock) " &
            " WHERE ISNULL(C.active_y,0) = 1 And  Res_ID = " & pintID & " "

            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 
            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            End If

        Catch er As Exception
            MsgBox("Error in cCicntp." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
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
    '        MsgBox("Error in cCicntp." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Function

    ' Load is made private, it is loaded automatically on New.
    ' Loads the Comment fields from the DataRow record
    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            If Not (pdrRow.Item("ID").Equals(DBNull.Value)) Then m_ID = pdrRow.Item("ID")
            If Not (pdrRow.Item("cnt_id").Equals(DBNull.Value)) Then m_cnt_id = pdrRow.Item("cnt_id").ToString
            If Not (pdrRow.Item("cmp_wwn").Equals(DBNull.Value)) Then m_cmp_wwn = pdrRow.Item("cmp_wwn").ToString
            If Not (pdrRow.Item("cnt_f_name").Equals(DBNull.Value)) Then m_cnt_f_name = pdrRow.Item("cnt_f_name")
            If Not (pdrRow.Item("cnt_l_name").Equals(DBNull.Value)) Then m_cnt_l_name = pdrRow.Item("cnt_l_name")
            If Not (pdrRow.Item("cnt_m_name").Equals(DBNull.Value)) Then m_cnt_m_name = pdrRow.Item("cnt_m_name")
            If Not (pdrRow.Item("FullName").Equals(DBNull.Value)) Then m_FullName = pdrRow.Item("FullName")
            If Not (pdrRow.Item("Initials").Equals(DBNull.Value)) Then m_Initials = pdrRow.Item("Initials")
            If Not (pdrRow.Item("Gender").Equals(DBNull.Value)) Then m_Gender = pdrRow.Item("Gender")
            If Not (pdrRow.Item("predcode").Equals(DBNull.Value)) Then m_predcode = pdrRow.Item("predcode")
            If Not (pdrRow.Item("cnt_job_desc").Equals(DBNull.Value)) Then m_cnt_job_desc = pdrRow.Item("cnt_job_desc")
            If Not (pdrRow.Item("cnt_dept").Equals(DBNull.Value)) Then m_cnt_dept = pdrRow.Item("cnt_dept")
            If Not (pdrRow.Item("taalcode").Equals(DBNull.Value)) Then m_taalcode = pdrRow.Item("taalcode")
            If Not (pdrRow.Item("cnt_f_ext").Equals(DBNull.Value)) Then m_cnt_f_ext = pdrRow.Item("cnt_f_ext")
            If Not (pdrRow.Item("cnt_f_fax").Equals(DBNull.Value)) Then m_cnt_f_fax = pdrRow.Item("cnt_f_fax")
            If Not (pdrRow.Item("cnt_f_tel").Equals(DBNull.Value)) Then m_cnt_f_tel = pdrRow.Item("cnt_f_tel")
            If Not (pdrRow.Item("cnt_f_mobile").Equals(DBNull.Value)) Then m_cnt_f_mobile = pdrRow.Item("cnt_f_mobile")
            If Not (pdrRow.Item("cnt_email").Equals(DBNull.Value)) Then m_cnt_email = pdrRow.Item("cnt_email")
            If Not (pdrRow.Item("cnt_acc_man").Equals(DBNull.Value)) Then m_cnt_acc_man = pdrRow.Item("cnt_acc_man")
            If Not (pdrRow.Item("cnt_note").Equals(DBNull.Value)) Then m_cnt_note = pdrRow.Item("cnt_note")
            If Not (pdrRow.Item("active_y").Equals(DBNull.Value)) Then m_active_y = pdrRow.Item("active_y")
            If Not (pdrRow.Item("Administration").Equals(DBNull.Value)) Then m_Administration = pdrRow.Item("Administration")
            If Not (pdrRow.Item("res_id").Equals(DBNull.Value)) Then m_res_id = pdrRow.Item("res_id")
            If Not (pdrRow.Item("WebAccess").Equals(DBNull.Value)) Then m_WebAccess = pdrRow.Item("WebAccess")
            If Not (pdrRow.Item("Picture").Equals(DBNull.Value)) Then m_Picture = pdrRow.Item("Picture")
            If Not (pdrRow.Item("PictureFileName").Equals(DBNull.Value)) Then m_PictureFileName = pdrRow.Item("PictureFileName")
            If Not (pdrRow.Item("syscreated").Equals(DBNull.Value)) Then m_syscreated = pdrRow.Item("syscreated")
            If Not (pdrRow.Item("syscreator").Equals(DBNull.Value)) Then m_syscreator = pdrRow.Item("syscreator")
            If Not (pdrRow.Item("sysmodified").Equals(DBNull.Value)) Then m_sysmodified = pdrRow.Item("sysmodified")
            If Not (pdrRow.Item("sysmodifier").Equals(DBNull.Value)) Then m_sysmodifier = pdrRow.Item("sysmodifier")
            If Not (pdrRow.Item("sysguid").Equals(DBNull.Value)) Then m_sysguid = pdrRow.Item("sysguid").ToString
            If Not (pdrRow.Item("timestamp").Equals(DBNull.Value)) Then m_timestamp = pdrRow.Item("timestamp").ToString
            If Not (pdrRow.Item("datefield1").Equals(DBNull.Value)) Then m_datefield1 = pdrRow.Item("datefield1")
            If Not (pdrRow.Item("datefield2").Equals(DBNull.Value)) Then m_datefield2 = pdrRow.Item("datefield2")
            If Not (pdrRow.Item("datefield3").Equals(DBNull.Value)) Then m_datefield3 = pdrRow.Item("datefield3")
            If Not (pdrRow.Item("datefield4").Equals(DBNull.Value)) Then m_datefield4 = pdrRow.Item("datefield4")
            If Not (pdrRow.Item("datefield5").Equals(DBNull.Value)) Then m_datefield5 = pdrRow.Item("datefield5")
            If Not (pdrRow.Item("numberfield1").Equals(DBNull.Value)) Then m_numberfield1 = pdrRow.Item("numberfield1")
            If Not (pdrRow.Item("numberfield2").Equals(DBNull.Value)) Then m_numberfield2 = pdrRow.Item("numberfield2")
            If Not (pdrRow.Item("numberfield3").Equals(DBNull.Value)) Then m_numberfield3 = pdrRow.Item("numberfield3")
            If Not (pdrRow.Item("numberfield4").Equals(DBNull.Value)) Then m_numberfield4 = pdrRow.Item("numberfield4")
            If Not (pdrRow.Item("numberfield5").Equals(DBNull.Value)) Then m_numberfield5 = pdrRow.Item("numberfield5")
            If Not (pdrRow.Item("YesNoField1").Equals(DBNull.Value)) Then m_YesNoField1 = pdrRow.Item("YesNoField1")
            If Not (pdrRow.Item("YesNoField2").Equals(DBNull.Value)) Then m_YesNoField2 = pdrRow.Item("YesNoField2")
            If Not (pdrRow.Item("YesNoField3").Equals(DBNull.Value)) Then m_YesNoField3 = pdrRow.Item("YesNoField3")
            If Not (pdrRow.Item("YesNoField4").Equals(DBNull.Value)) Then m_YesNoField4 = pdrRow.Item("YesNoField4")
            If Not (pdrRow.Item("YesNoField5").Equals(DBNull.Value)) Then m_YesNoField5 = pdrRow.Item("YesNoField5")
            If Not (pdrRow.Item("textfield1").Equals(DBNull.Value)) Then m_textfield1 = pdrRow.Item("textfield1")
            If Not (pdrRow.Item("textfield2").Equals(DBNull.Value)) Then m_textfield2 = pdrRow.Item("textfield2")
            If Not (pdrRow.Item("textfield3").Equals(DBNull.Value)) Then m_textfield3 = pdrRow.Item("textfield3")
            If Not (pdrRow.Item("textfield4").Equals(DBNull.Value)) Then m_textfield4 = pdrRow.Item("textfield4")
            If Not (pdrRow.Item("textfield5").Equals(DBNull.Value)) Then m_textfield5 = pdrRow.Item("textfield5")
            If Not (pdrRow.Item("cntp_Directory").Equals(DBNull.Value)) Then m_cntp_Directory = pdrRow.Item("cntp_Directory")
            If Not (pdrRow.Item("Division").Equals(DBNull.Value)) Then m_Division = pdrRow.Item("Division")
            If Not (pdrRow.Item("ContactChange").Equals(DBNull.Value)) Then m_ContactChange = pdrRow.Item("ContactChange")
            If Not (pdrRow.Item("OldFileName").Equals(DBNull.Value)) Then m_OldFileName = pdrRow.Item("OldFileName")

        Catch er As Exception
            MsgBox("Error in cCicntp." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Saves the Comment fields into the DataRow record
    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("ID") = m_ID
            pdrRow.Item("cnt_id") = m_cnt_id
            pdrRow.Item("cmp_wwn") = m_cmp_wwn
            pdrRow.Item("cnt_f_name") = m_cnt_f_name
            pdrRow.Item("cnt_l_name") = m_cnt_l_name
            pdrRow.Item("cnt_m_name") = m_cnt_m_name
            pdrRow.Item("FullName") = m_FullName
            pdrRow.Item("Initials") = m_Initials
            pdrRow.Item("Gender") = m_Gender
            pdrRow.Item("predcode") = m_predcode
            pdrRow.Item("cnt_job_desc") = m_cnt_job_desc
            pdrRow.Item("cnt_dept") = m_cnt_dept
            pdrRow.Item("taalcode") = m_taalcode
            pdrRow.Item("cnt_f_ext") = m_cnt_f_ext
            pdrRow.Item("cnt_f_fax") = m_cnt_f_fax
            pdrRow.Item("cnt_f_tel") = m_cnt_f_tel
            pdrRow.Item("cnt_f_mobile") = m_cnt_f_mobile
            pdrRow.Item("cnt_email") = m_cnt_email
            pdrRow.Item("cnt_acc_man") = m_cnt_acc_man
            pdrRow.Item("cnt_note") = m_cnt_note
            pdrRow.Item("active_y") = m_active_y
            pdrRow.Item("Administration") = m_Administration
            pdrRow.Item("res_id") = m_res_id
            pdrRow.Item("WebAccess") = m_WebAccess
            pdrRow.Item("Picture") = m_Picture
            pdrRow.Item("PictureFileName") = m_PictureFileName
            pdrRow.Item("syscreated") = m_syscreated
            pdrRow.Item("syscreator") = m_syscreator
            pdrRow.Item("sysmodified") = m_sysmodified
            pdrRow.Item("sysmodifier") = m_sysmodifier
            pdrRow.Item("sysguid") = m_sysguid
            pdrRow.Item("timestamp") = m_timestamp
            pdrRow.Item("datefield1") = m_datefield1
            pdrRow.Item("datefield2") = m_datefield2
            pdrRow.Item("datefield3") = m_datefield3
            pdrRow.Item("datefield4") = m_datefield4
            pdrRow.Item("datefield5") = m_datefield5
            pdrRow.Item("numberfield1") = m_numberfield1
            pdrRow.Item("numberfield2") = m_numberfield2
            pdrRow.Item("numberfield3") = m_numberfield3
            pdrRow.Item("numberfield4") = m_numberfield4
            pdrRow.Item("numberfield5") = m_numberfield5
            pdrRow.Item("YesNoField1") = m_YesNoField1
            pdrRow.Item("YesNoField2") = m_YesNoField2
            pdrRow.Item("YesNoField3") = m_YesNoField3
            pdrRow.Item("YesNoField4") = m_YesNoField4
            pdrRow.Item("YesNoField5") = m_YesNoField5
            pdrRow.Item("textfield1") = m_textfield1
            pdrRow.Item("textfield2") = m_textfield2
            pdrRow.Item("textfield3") = m_textfield3
            pdrRow.Item("textfield4") = m_textfield4
            pdrRow.Item("textfield5") = m_textfield5
            pdrRow.Item("cntp_Directory") = m_cntp_Directory
            pdrRow.Item("Division") = m_Division
            pdrRow.Item("ContactChange") = m_ContactChange
            pdrRow.Item("OldFileName") = m_OldFileName

        Catch er As Exception
            MsgBox("Error in cCicntp." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    ' Deletes the current Comment from the database, if it exists.
    Public Sub Delete()

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If m_ID = 0 Then Exit Sub

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            strSql = _
            "DELETE FROM Cicntp " & _
            "WHERE  ID = " & m_ID & " " 

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in cCicntp." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every non mandatory field of a Comment
    Public Sub Reset()

        ' Resets the locals to empty values, all but ID, Ord_No and Ord_Guid
        'm_Ord_No = String.Empty
        'm_Ord_Guid = String.Empty
        'm_Cmt = String.Empty
        'm_Line_Seq_No = 0

        m_numberfield1 = 0
        m_textfield1 = String.Empty

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
            "FROM   Cicntp " & _
            "WHERE  ID = " & m_ID & " "

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
            MsgBox("Error in cCicntp." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property ID() As Integer
        Get
            ID = m_ID
        End Get
        Set(ByVal value As Integer)
            m_ID = value
        End Set
    End Property
    Public Property cnt_id() As String
        Get
            cnt_id = m_cnt_id
        End Get
        Set(ByVal value As String)
            m_cnt_id = value
        End Set
    End Property
    Public Property cmp_wwn() As String
        Get
            cmp_wwn = m_cmp_wwn
        End Get
        Set(ByVal value As String)
            m_cmp_wwn = value
        End Set
    End Property
    Public Property cnt_f_name() As String
        Get
            cnt_f_name = m_cnt_f_name
        End Get
        Set(ByVal value As String)
            m_cnt_f_name = value
        End Set
    End Property
    Public Property cnt_l_name() As String
        Get
            cnt_l_name = m_cnt_l_name
        End Get
        Set(ByVal value As String)
            m_cnt_l_name = value
        End Set
    End Property
    Public Property cnt_m_name() As String
        Get
            cnt_m_name = m_cnt_m_name
        End Get
        Set(ByVal value As String)
            m_cnt_m_name = value
        End Set
    End Property
    Public Property FullName() As String
        Get
            FullName = m_FullName
        End Get
        Set(ByVal value As String)
            m_FullName = value
        End Set
    End Property
    Public Property Initials() As String
        Get
            Initials = m_Initials
        End Get
        Set(ByVal value As String)
            m_Initials = value
        End Set
    End Property
    Public Property Gender() As String
        Get
            Gender = m_Gender
        End Get
        Set(ByVal value As String)
            m_Gender = value
        End Set
    End Property
    Public Property predcode() As String
        Get
            predcode = m_predcode
        End Get
        Set(ByVal value As String)
            m_predcode = value
        End Set
    End Property
    Public Property cnt_job_desc() As String
        Get
            cnt_job_desc = m_cnt_job_desc
        End Get
        Set(ByVal value As String)
            m_cnt_job_desc = value
        End Set
    End Property
    Public Property cnt_dept() As String
        Get
            cnt_dept = m_cnt_dept
        End Get
        Set(ByVal value As String)
            m_cnt_dept = value
        End Set
    End Property
    Public Property taalcode() As String
        Get
            taalcode = m_taalcode
        End Get
        Set(ByVal value As String)
            m_taalcode = value
        End Set
    End Property
    Public Property cnt_f_ext() As String
        Get
            cnt_f_ext = m_cnt_f_ext
        End Get
        Set(ByVal value As String)
            m_cnt_f_ext = value
        End Set
    End Property
    Public Property cnt_f_fax() As String
        Get
            cnt_f_fax = m_cnt_f_fax
        End Get
        Set(ByVal value As String)
            m_cnt_f_fax = value
        End Set
    End Property
    Public Property cnt_f_tel() As String
        Get
            cnt_f_tel = m_cnt_f_tel
        End Get
        Set(ByVal value As String)
            m_cnt_f_tel = value
        End Set
    End Property
    Public Property cnt_f_mobile() As String
        Get
            cnt_f_mobile = m_cnt_f_mobile
        End Get
        Set(ByVal value As String)
            m_cnt_f_mobile = value
        End Set
    End Property
    Public Property cnt_email() As String
        Get
            cnt_email = m_cnt_email
        End Get
        Set(ByVal value As String)
            m_cnt_email = value
        End Set
    End Property
    Public Property cnt_acc_man() As Integer
        Get
            cnt_acc_man = m_cnt_acc_man
        End Get
        Set(ByVal value As Integer)
            m_cnt_acc_man = value
        End Set
    End Property
    Public Property cnt_note() As String
        Get
            cnt_note = m_cnt_note
        End Get
        Set(ByVal value As String)
            m_cnt_note = value
        End Set
    End Property
    Public Property active_y() As Integer
        Get
            active_y = m_active_y
        End Get
        Set(ByVal value As Integer)
            m_active_y = value
        End Set
    End Property
    Public Property Administration() As String
        Get
            Administration = m_Administration
        End Get
        Set(ByVal value As String)
            m_Administration = value
        End Set
    End Property
    Public Property res_id() As Integer
        Get
            res_id = m_res_id
        End Get
        Set(ByVal value As Integer)
            m_res_id = value
        End Set
    End Property
    Public Property WebAccess() As Integer
        Get
            WebAccess = m_WebAccess
        End Get
        Set(ByVal value As Integer)
            m_WebAccess = value
        End Set
    End Property
    Public Property Picture() As String
        Get
            Picture = m_Picture
        End Get
        Set(ByVal value As String)
            m_Picture = value
        End Set
    End Property
    Public Property PictureFileName() As String
        Get
            PictureFileName = m_PictureFileName
        End Get
        Set(ByVal value As String)
            m_PictureFileName = value
        End Set
    End Property
    Public Property syscreated() As Date
        Get
            syscreated = m_syscreated
        End Get
        Set(ByVal value As Date)
            m_syscreated = value
        End Set
    End Property
    Public Property syscreator() As Integer
        Get
            syscreator = m_syscreator
        End Get
        Set(ByVal value As Integer)
            m_syscreator = value
        End Set
    End Property
    Public Property sysmodified() As Date
        Get
            sysmodified = m_sysmodified
        End Get
        Set(ByVal value As Date)
            m_sysmodified = value
        End Set
    End Property
    Public Property sysmodifier() As Integer
        Get
            sysmodifier = m_sysmodifier
        End Get
        Set(ByVal value As Integer)
            m_sysmodifier = value
        End Set
    End Property
    Public Property Alternate_Rep() As Integer
        Get
            Alternate_Rep = m_numberfield1
        End Get
        Set(ByVal value As Integer)
            m_numberfield1 = value
        End Set
    End Property
    Public Property Use_Account_No() As String
        Get
            Use_Account_No = m_textfield1
        End Get
        Set(ByVal value As String)
            m_textfield1 = value
        End Set
    End Property

#End Region

End Class
