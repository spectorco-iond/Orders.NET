Public Class cImprint

#Region "Private variables ################################################"

    Private m_Ord_Guid As String
    Private m_Item_Guid As String
    Private m_Item_No As String
    Private m_ID As Integer

    Private m_Imprint As String
    Private m_Location As String
    Private m_Num As Integer
    Private m_Color As String
    Private m_Packaging As String
    Private m_Refill As String
    Private m_Laser_Setup As String
    Private m_Comments As String
    Private m_Special_Comments As String
    Private m_Industry As String
    Private m_Printer_Comment As String
    Private m_Packer_Comment As String
    Private m_Printer_Instructions As String
    Private m_Packer_Instructions As String

    Private m_Repeat_From_Ord_No As String
    Private m_Repeat_From_ID As Integer
    Private m_Repeat_From_RouteID As Integer

    Private m_Num_Impr_1 As Integer
    Private m_Num_Impr_2 As Integer
    Private m_Num_Impr_3 As Integer

    Private m_End_User As String

    Private m_Dirty As Boolean = False
    Private m_SaveToDB As Boolean = True

    '+++ID 06.08.2020 for Scribl project
    Private m_Lamination_Scribl As String
    '++ID 12.09.2021 Foil_color
    Private m_Foil_Color As String

    Private m_Thread_Color_Scribl As String

    '++ID 07.28.2020
    Private m_Rounded_Corners_Scribl As String
    '++ID08.05.2020
    Private m_Tip_In_PageTxt As String
    '--------------------------------------
#End Region

#Region "Public constructors ##############################################"

    Public Sub New()

        Call Init()

    End Sub

    Public Sub New(ByVal pstrItem_Guid As String)

        m_Ord_Guid = m_oOrder.Ordhead.Ord_GUID ' .Empty
        m_Item_Guid = pstrItem_Guid

        Call Reset()

        If m_oOrder.Ordhead.User_Def_Fld_2 = "SP" Then m_Industry = "Self Promo"

    End Sub

    Public Sub New(ByVal pstrOrd_Guid As String, ByVal pstrItem_Guid As String)

        m_Ord_Guid = pstrOrd_Guid
        m_Item_Guid = pstrItem_Guid

        Call Reset()

        If m_oOrder.Ordhead.User_Def_Fld_2 = "SP" Then m_Industry = "Self Promo"

    End Sub

#End Region

#Region "Public I/O procedures ############################################"

    Public Function IsEmpty() As Boolean

        Dim blnEmpty As Boolean = True

        blnEmpty = (m_ID = 0)

        If blnEmpty Then
            blnEmpty = (m_Imprint.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Location.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Color.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Packaging.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Refill.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Laser_Setup.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Comments.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Special_Comments.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Industry.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Printer_Comment.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Packer_Comment.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Printer_Instructions.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Packer_Instructions.Trim = String.Empty) And blnEmpty

            '++ID 06.09.2020 csribl project
            blnEmpty = (Lamination_Scribl.Trim = String.Empty) And blnEmpty
            '++ID 12.09.2021 Foil_Color
            blnEmpty = (Foil_Color.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Thread_Color_Scribl.Trim = String.Empty) And blnEmpty
            blnEmpty = (m_Rounded_Corners_Scribl = String.Empty) And blnEmpty
            '-----------------------------------------------------------------
            '++ID 08.06.2020 
            blnEmpty = (m_Tip_In_PageTxt = String.Empty) And blnEmpty

        End If

        IsEmpty = blnEmpty

    End Function

    Public Sub Delete()

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            Dim conn As New cDBA
            Dim dt As New DataTable

            ' MB++ Faire aussi un DELETE de masse pour le master Item_Guid, bien que pas 
            ' necessairement besoin vu que ca sera probablement appelé d'un niveau plus haut.

            Dim strSql As String = _
            "SELECT     * " & _
            "FROM       OEI_EXTRA_FIELDS " & _
            "WHERE      TRIM(Ord_Guid) = '" & Trim(Ord_Guid) & "' AND " & _
            "           TRIM(Item_Guid) = '" & Trim(Item_Guid) & "' "

            dt = conn.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                conn.DBDataTable.Rows.RemoveAt(0)
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(conn.DBDataAdapter)
                conn.DBDataAdapter.DeleteCommand = cmd.GetDeleteCommand
            End If

            conn.DBDataAdapter.Update(conn.DBDataTable)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Init()

        m_ID = 0
        m_Ord_Guid = String.Empty
        m_Item_Guid = String.Empty
        m_Item_No = String.Empty

        Call Reset()

        'm_Imprint = String.Empty
        'm_Location = String.Empty
        'm_Num = 0 ' String.Empty
        'm_Num_Impr_1 = 0
        'm_Num_Impr_2 = 0
        'm_Num_Impr_3 = 0
        'm_Color = String.Empty
        'm_Packaging = String.Empty
        'm_Refill = String.Empty
        'm_Laser_Setup = String.Empty
        'm_Comments = String.Empty
        'm_Special_Comments = String.Empty
        'm_Industry = String.Empty

        'm_Printer_Comment = String.Empty
        'm_Packer_Comment = String.Empty
        'm_Printer_Instructions = String.Empty
        'm_Packer_Instructions = String.Empty

        'm_Repeat_From_Ord_No = String.Empty
        'm_Repeat_From_ID = 0
        'm_Repeat_From_RouteID = 0

    End Sub

    Public Sub Load()

        Try
            Dim conn As New cDBA
            Dim dt As New DataTable
            Dim drDataRow As DataRow

            Dim strSql As String =
            "SELECT     1 AS ID, ISNULL(Extra_1, '') AS Extra_1, ISNULL(Extra_2, '') AS Extra_2, " &
            "           ISNULL(Extra_3, '') AS Extra_3, ISNULL(Extra_4, '0') AS Extra_4, " &
            "           ISNULL(Extra_5, '') AS Extra_5, ISNULL(Extra_6, '') AS Extra_6, " &
            "           ISNULL(Extra_7, '') AS Extra_7, ISNULL(Extra_8, '') AS Extra_8, " &
            "           ISNULL(Extra_9, '') AS Extra_9, ISNULL(Item_No, '') AS Item_No, " &
            "           ISNULL(Comment, '') AS Comment, ISNULL(Comment2, '') AS Comment2, " &
            "           ISNULL(Industry, '') AS Industry, " &
            "           ISNULL(Printer_Comment, '') AS Printer_Comment, " &
            "           ISNULL(Packer_Comment, '') AS Packer_Comment, " &
            "           ISNULL(Printer_Instructions, '') AS Printer_Instructions, " &
            "           ISNULL(Packer_Instructions, '') AS Packer_Instructions,  " &
            "           ISNULL(Repeat_From_Ord_No, '') AS Repeat_From_Ord_No, " &
            "           ISNULL(Repeat_From_ID, 0) AS Repeat_From_ID, " &
            "           ISNULL(Repeat_From_RouteID, 0) AS Repeat_From_RouteID, " &
            "           CASE WHEN ISNULL(Num_Impr_1, 0) = 0 THEN CONVERT(INT, ISNULL(Extra_4, '0')) ELSE ISNULL(Num_Impr_1, 0) END AS Num_Impr_1, " &
            "           ISNULL(Num_Impr_2, 0) AS Num_Impr_2, ISNULL(Num_Impr_3, 0) AS Num_Impr_3,  " &
            "           ISNULL(End_User, '') AS End_User, " &
            " ISNULL(Lamination_Scribl,'') AS Lamination_Scribl, ISNULL(Foil_Color,'') AS Foil_Color, ISNULL(Thread_Color_Scribl,'') AS Thread_Color_Scribl, ISNULL(Rounded_Corners_Scribl ,'') AS Rounded_corners_Scribl, " &
            " ISNULL(Tip_In_PageTxt, '') AS Tip_In_PageTxt  " & '++ID 08.06.2020
            " FROM       OEI_EXTRA_FIELDS WITH (Nolock) " &
            " WHERE      Ord_Guid = '" & Trim(Ord_Guid) & "' AND " &
            "           Item_Guid = '" & Trim(Item_Guid) & "' "
            'CAST(0 AS BIT) AS Rounded_corners_Scribl
            '++ID 12.09.2021 Foil_Color added in sql

            dt = conn.DataTable(strSql)
            If dt.Rows.Count <> 0 Then

                drDataRow = dt.Rows(0)
                With drDataRow

                    m_ID = .Item("ID")
                    m_Item_No = .Item("Item_No").ToString
                    m_Imprint = .Item("Extra_1").ToString
                    m_Location = .Item("Extra_2").ToString
                    m_Color = .Item("Extra_3").ToString
                    m_Num = .Item("Extra_4").ToString
                    m_Packaging = .Item("Extra_5").ToString
                    'drDataRow.Item("Extra_6") = Item_Guid ' Extra_6 n'est pas utilisé
                    'drDataRow.Item("Extra_7") = Item_Guid ' Extra_7 n'est pas utilisé
                    m_Refill = .Item("Extra_8").ToString
                    m_Laser_Setup = .Item("Extra_9").ToString
                    m_Comments = .Item("Comment").ToString
                    m_Special_Comments = .Item("Comment2").ToString
                    m_Industry = .Item("Industry").ToString

                    m_Printer_Comment = .Item("Printer_Comment").ToString
                    m_Packer_Comment = .Item("Packer_Comment").ToString
                    m_Printer_Instructions = .Item("Printer_Instructions").ToString
                    m_Packer_Instructions = .Item("Packer_Instructions").ToString

                    m_Repeat_From_Ord_No = .Item("Repeat_From_Ord_No")
                    m_Repeat_From_ID = .Item("Repeat_From_ID")
                    m_Repeat_From_RouteID = .Item("Repeat_From_RouteID")

                    m_Num_Impr_1 = .Item("Num_Impr_1")
                    m_Num_Impr_2 = .Item("Num_Impr_2")
                    m_Num_Impr_3 = .Item("Num_Impr_3")

                    m_End_User = .Item("End_User").ToString

                    m_Lamination_Scribl = .Item("Lamination_Scribl").ToString
                    m_Foil_Color = .Item("Foil_Color").ToString '++ID 12.09.2021 foil_color
                    m_Thread_Color_Scribl = .Item("Thread_Color_Scribl").ToString
                    m_Rounded_Corners_Scribl = .Item("Rounded_Corners_Scribl").ToString
                    '++ID 08.06.2020
                    m_Tip_In_PageTxt = .Item("Tip_In_PageTxt")

                End With

            End If

        Catch er As Exception
            MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Reset()

        m_Imprint = String.Empty
        m_Location = String.Empty
        m_Num = 0 ' String.Empty
        m_Color = String.Empty
        m_Packaging = String.Empty
        m_Refill = String.Empty
        m_Laser_Setup = String.Empty
        m_Comments = String.Empty
        m_Special_Comments = String.Empty
        m_Industry = String.Empty

        m_Printer_Comment = String.Empty
        m_Packer_Comment = String.Empty
        m_Printer_Instructions = String.Empty
        m_Packer_Instructions = String.Empty

        m_Repeat_From_Ord_No = String.Empty
        m_Repeat_From_ID = 0
        m_Repeat_From_RouteID = 0

        m_Num_Impr_1 = 0
        m_Num_Impr_2 = 0
        m_Num_Impr_3 = 0

        m_End_User = String.Empty
        '++ID 06.08.2020 ------ new field for scribl project
        m_Lamination_Scribl = String.Empty
        m_Foil_Color = String.Empty '++ID 12.09.2021 Foil_Color
        m_Thread_Color_Scribl = String.Empty
        m_Rounded_Corners_Scribl = String.Empty
        '++ID 08.06.2020
        m_Tip_In_PageTxt = String.Empty
        '---------------------------------------------------
        'If m_oOrder.Ordhead.User_Def_Fld_2 = "SP" Then m_Industry = "Self Promo"

    End Sub

    'Public Sub SaveIfNotEmpty()

    '    If Not IsEmpty() Then
    '        Save()
    '    End If

    'End Sub

    'Public Sub SaveIfImprint()

    '    If m_Imprint <> "" Then
    '        Save()
    '    End If

    'End Sub

    Public Sub checkSaveEligibility()
        'If Not m_Item_No Is Nothing And m_Item_No <> "" And Mid(m_Item_No, 1, 2) <> "44" And Mid(m_Item_No, 1, 2) <> "88" Then
        If Not m_Item_No Is Nothing And m_Item_No <> "" Then
            Call Save()
        End If
    End Sub

    Public Sub Save()

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub
        'If Empty() Then Exit Sub
        'If m_Imprint = "" Then Exit Sub

        Try
            Dim conn As New cDBA
            Dim dt As New DataTable
            Dim drDataRow As DataRow

            If Not m_Dirty Then Exit Sub
            If Not m_SaveToDB Then Exit Sub

            Dim strSql As String = _
            "SELECT     * " & _
            "FROM       OEI_EXTRA_FIELDS " & _
            "WHERE      Ord_Guid = '" & Trim(Ord_Guid) & "' AND " & _
            "           Item_Guid = '" & Trim(Item_Guid) & "' "

            dt = conn.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                drDataRow = dt.Rows(0)
            Else
                drDataRow = dt.NewRow
                drDataRow.Item("Ord_Guid") = Ord_Guid
                drDataRow.Item("Item_Guid") = Item_Guid
            End If

            drDataRow.Item("Item_No") = m_Item_No
            drDataRow.Item("Extra_1") = m_Imprint ''''
            drDataRow.Item("Extra_2") = m_Location ''''
            drDataRow.Item("Extra_3") = m_Color ''''
            drDataRow.Item("Extra_4") = m_Num ' m_Num ''''
            drDataRow.Item("Extra_5") = m_Packaging ''''
            'drDataRow.Item("Extra_6") = Item_Guid
            'drDataRow.Item("Extra_7") = Item_Guid
            drDataRow.Item("Extra_8") = m_Refill ''''
            drDataRow.Item("Extra_9") = m_Laser_Setup
            drDataRow.Item("Comment") = m_Comments ''''
            drDataRow.Item("Comment2") = m_Special_Comments ''''
            drDataRow.Item("Industry") = m_Industry
            drDataRow.Item("Printer_Comment") = m_Printer_Comment
            drDataRow.Item("Packer_Comment") = m_Packer_Comment
            drDataRow.Item("Printer_Instructions") = m_Printer_Instructions
            drDataRow.Item("Packer_Instructions") = m_Packer_Instructions

            drDataRow.Item("Repeat_From_Ord_No") = m_Repeat_From_Ord_No
            drDataRow.Item("Repeat_From_ID") = m_Repeat_From_ID
            drDataRow.Item("Repeat_From_RouteID") = m_Repeat_From_RouteID

            drDataRow.Item("Num_Impr_1") = m_Num_Impr_1
            drDataRow.Item("Num_Impr_2") = m_Num_Impr_2
            drDataRow.Item("Num_Impr_3") = m_Num_Impr_3

            drDataRow.Item("End_User") = m_End_User

            '++ID 06.08.2020 ------ new field for scribl project
            drDataRow.Item("Lamination_Scribl") = m_Lamination_Scribl
            drDataRow.Item("Foil_Color") = m_Foil_Color '++ID 12.09.2021 Foil_Color
            drDataRow.Item("Thread_Color_Scribl") = m_Thread_Color_Scribl
            drDataRow.Item("Rounded_Corners_Scribl") = m_Rounded_Corners_Scribl
            '++ID 08.06.2020 Tip_In_PageTxt
            drDataRow.Item("Tip_In_PageTxt") = m_Tip_In_PageTxt
            '------------------------------------------------------------------
            'drDataRow.Item("PrimaryContact") = m_PrimaryContact
            'drDataRow.Item("DateAdded") = DateAdded

            If dt.Rows.Count = 0 Then
                conn.DBDataTable.Rows.Add(drDataRow)
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(conn.DBDataAdapter)
                conn.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            Else
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(conn.DBDataAdapter)
                conn.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            End If

            conn.DBDataAdapter.Update(conn.DBDataTable)

        Catch er As Exception
            MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

    Public Sub SetNumImprints(ByRef iRouteID As Integer, ByVal ordLineItems As String())

        Dim strSql As String
        Dim db As New cDBA
        Dim dt As DataTable

        Try

            strSql = _
            "SELECT		R.RouteID, R.RouteDescription, " & _
            "			CASE WHEN RC.Deco_Meth_1 IS NULL THEN 0 ELSE CASE WHEN RCI.Num_Impr_1 IS NOT NULL THEN RCI.Num_Impr_1 ELSE 1 END END AS Num_Impr_1, " & _
            "			CASE WHEN RC.Deco_Meth_2 IS NULL THEN 0 ELSE CASE WHEN RCI.Num_Impr_2 IS NOT NULL THEN RCI.Num_Impr_2 ELSE 1 END END AS Num_Impr_2, " & _
            "			CASE WHEN RC.Deco_Meth_3 IS NULL THEN 0 ELSE CASE WHEN RCI.Num_Impr_3 IS NOT NULL THEN RCI.Num_Impr_3 ELSE 1 END END AS Num_Impr_3 " & _
            "FROM		Exact_Traveler_Route R WITH (Nolock) " & _
            "INNER JOIN	Exact_Traveler_XRef_RouteCategory RC WITH (Nolock) ON R.RouteCategory = RC.RouteCategory " & _
            "LEFT JOIN  EXACT_TRAVELER_Xref_RouteCategory_ITEM_IMP_COUNT_DEFAULTS RCI " & _
                " ON RC.RouteCategory = RCI.RouteCategory AND " & _
                "    (RCI.ITEM_CD IN (" & ordLineItems(0).ToString.Trim() & ") OR RCI.ITEM_NO IN (" & ordLineItems(1).ToString.Trim() & ")) " & _
            "WHERE      R.RouteID = " & iRouteID.ToString

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then

                If m_Num_Impr_1 <= 1 Then

                    m_Num_Impr_1 = dt.Rows(0).Item("Num_Impr_1")
                    m_Num_Impr_2 = dt.Rows(0).Item("Num_Impr_2")
                    m_Num_Impr_3 = dt.Rows(0).Item("Num_Impr_3")

                End If

                m_Num = m_Num_Impr_1 + m_Num_Impr_2 + m_Num_Impr_3

            Else

                ' Do nothing, don't overwrite

            End If

            m_Dirty = True

        Catch er As Exception
            MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub
#Region "Public Properties ################################################"

    Public Property Color() As String
        Get
            Color = m_Color
        End Get
        Set(ByVal value As String)
            If m_Color <> value Then m_Dirty = True
            m_Color = value
            Call checkSaveEligibility()
            'If m_Imprint <> "" Then Call Save() 'old logic (repeated everywhere)
        End Set
    End Property
    Public Property Comments() As String
        Get
            Comments = m_Comments
        End Get
        Set(ByVal value As String)
            If m_Comments <> value Then m_Dirty = True
            m_Comments = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Imprint() As String
        Get
            Imprint = m_Imprint
        End Get
        Set(ByVal value As String)
            If m_Imprint <> value Then m_Dirty = True
            m_Imprint = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Industry() As String
        Get
            Industry = m_Industry
        End Get
        Set(ByVal value As String)
            If m_Industry <> value Then m_Dirty = True
            m_Industry = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Item_Guid() As String
        Get
            Item_Guid = m_Item_Guid
        End Get
        Set(ByVal value As String)
            If m_Item_Guid <> value Then m_Dirty = True
            m_Item_Guid = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Item_No() As String
        Get
            Item_No = m_Item_No
        End Get
        Set(ByVal value As String)
            If m_Item_No <> value Then m_Dirty = True
            m_Item_No = value
            'Call Save()
        End Set
    End Property
    Public Property Laser_Setup() As String
        Get
            Laser_Setup = m_Laser_Setup
        End Get
        Set(ByVal value As String)
            If m_Laser_Setup <> value Then m_Dirty = True
            m_Laser_Setup = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Location() As String
        Get
            Location = m_Location
        End Get
        Set(ByVal value As String)
            If m_Location <> value Then m_Dirty = True
            m_Location = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Num() As Integer
        Get
            Num = m_Num
        End Get
        Set(ByVal value As Integer)
            If m_Num <> value Then m_Dirty = True
            m_Num = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Num_Impr_1() As Integer
        Get
            Num_Impr_1 = m_Num_Impr_1
        End Get
        Set(ByVal value As Integer)
            If m_Num_Impr_1 <> value Then m_Dirty = True
            m_Num_Impr_1 = value
            m_Num = m_Num_Impr_1 + m_Num_Impr_2 + m_Num_Impr_3
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Num_Impr_2() As Integer
        Get
            Num_Impr_2 = m_Num_Impr_2
        End Get
        Set(ByVal value As Integer)
            If m_Num_Impr_2 <> value Then m_Dirty = True
            m_Num_Impr_2 = value
            m_Num = m_Num_Impr_1 + m_Num_Impr_2 + m_Num_Impr_3
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Num_Impr_3() As Integer
        Get
            Num_Impr_3 = m_Num_Impr_3
        End Get
        Set(ByVal value As Integer)
            If m_Num_Impr_3 <> value Then m_Dirty = True
            m_Num_Impr_3 = value
            m_Num = m_Num_Impr_1 + m_Num_Impr_2 + m_Num_Impr_3
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property End_User() As String
        Get
            End_User = m_End_User
        End Get
        Set(ByVal value As String)
            If m_End_User <> value Then m_Dirty = True
            m_End_User = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Ord_Guid() As String
        Get
            Ord_Guid = m_Ord_Guid
        End Get
        Set(ByVal value As String)
            If m_Ord_Guid <> value Then m_Dirty = True
            m_Ord_Guid = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Packaging() As String
        Get
            Packaging = m_Packaging
        End Get
        Set(ByVal value As String)
            If m_Packaging <> value Then m_Dirty = True
            m_Packaging = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Packer_Comment() As String
        Get
            Packer_Comment = m_Packer_Comment
        End Get
        Set(ByVal value As String)
            If m_Packer_Comment <> value Then m_Dirty = True
            m_Packer_Comment = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Packer_Instructions() As String
        Get
            Packer_Instructions = m_Packer_Instructions
        End Get
        Set(ByVal value As String)
            If m_Packer_Instructions <> value Then m_Dirty = True
            m_Packer_Instructions = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Printer_Comment() As String
        Get
            Printer_Comment = m_Printer_Comment
        End Get
        Set(ByVal value As String)
            If m_Printer_Comment <> value Then m_Dirty = True
            m_Printer_Comment = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Printer_Instructions() As String
        Get
            Printer_Instructions = m_Printer_Instructions
        End Get
        Set(ByVal value As String)
            If m_Printer_Instructions <> value Then m_Dirty = True
            m_Printer_Instructions = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Refill() As String
        Get
            Refill = m_Refill
        End Get
        Set(ByVal value As String)
            If m_Refill <> value Then m_Dirty = True
            m_Refill = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Repeat_From_Ord_No() As String
        Get
            Repeat_From_Ord_No = m_Repeat_From_Ord_No
        End Get
        Set(ByVal value As String)
            If m_Repeat_From_Ord_No <> value Then m_Dirty = True
            m_Repeat_From_Ord_No = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Repeat_From_ID() As Integer
        Get
            Repeat_From_ID = m_Repeat_From_ID
        End Get
        Set(ByVal value As Integer)
            If m_Repeat_From_ID <> value Then m_Dirty = True
            m_Repeat_From_ID = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Repeat_From_RouteID() As Integer
        Get
            Repeat_From_RouteID = m_Repeat_From_RouteID
        End Get
        Set(ByVal value As Integer)
            If m_Repeat_From_RouteID <> value Then m_Dirty = True
            m_Repeat_From_RouteID = value
            Call checkSaveEligibility()
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
    Public Property Special_Comments() As String
        Get
            Special_Comments = m_Special_Comments
        End Get
        Set(ByVal value As String)
            If m_Special_Comments <> value Then m_Dirty = True
            m_Special_Comments = value
            Call checkSaveEligibility()
        End Set
    End Property

    '++ID 06.08.2020 fpr scrabl project added 3 fields -----------------------------
    Public Property Lamination_Scribl As String
        Get
            Return m_Lamination_Scribl
        End Get
        Set(value As String)
            m_Lamination_Scribl = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Foil_Color As String
        Get
            Return m_Foil_Color
        End Get
        Set(value As String)
            m_Foil_Color = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Thread_Color_Scribl As String
        Get
            Return m_Thread_Color_Scribl
        End Get
        Set(value As String)
            m_Thread_Color_Scribl = value
            Call checkSaveEligibility()
        End Set
    End Property
    Public Property Rounded_corners_Scribl As String
        Get
            Return m_Rounded_Corners_Scribl
        End Get
        Set(value As String)
            m_Rounded_Corners_Scribl = value
            Call checkSaveEligibility()
        End Set
    End Property
    '++ID 08.06.2020
    Public Property Tip_In_PageTxt As String
        Get
            Return m_Tip_In_PageTxt
        End Get
        Set(value As String)
            m_Tip_In_PageTxt = value
            Call checkSaveEligibility()
        End Set
    End Property
    '--------------------------------------------------

    Public Sub AddColor(ByRef pValue As String, ByRef pstrColor As String)

        Dim s As String
        s = Trim(pstrColor)
        If Not (InStr(pstrColor, pValue)) > 0 Then
            If s <> "" Then s = s & "/"
            s = s & Trim(pValue)
            ' By using the property, it saves.
        End If
        Color = s

    End Sub

    Public Sub AddIndustry(ByRef pValue As String, ByRef pstrIndustry As String)

        Dim s As String
        s = Trim(pstrIndustry)
        If Not (InStr(pstrIndustry, pValue)) > 0 Then
            If s <> "" Then s = s & "/"
            s = s & Trim(pValue)
            ' By using the property, it saves.
        End If
        Industry = s

    End Sub

    Public Sub AddLocation(ByRef pValue As String, ByRef pstrLocation As String)

        Dim s As String
        s = Trim(pstrLocation)
        If Not (InStr(pstrLocation, pValue)) > 0 Then
            If s <> "" Then s = s & "/"
            s = s & Trim(pValue)
            ' By using the property, it saves.
        End If
        Location = s

    End Sub

    '++ID 06.29.2020 ---------------------------------------------
    Public Sub AddLamination(ByRef pValue As String, ByRef pstrLamination As String)

        Dim s As String
        s = Trim(pstrLamination)
        If Not (InStr(pstrLamination, pValue)) > 0 Then
            '   If s <> "" Then s = s & "/"
            '  s = s & Trim(pValue)
            s = Trim(pValue)
            ' By using the property, it saves.
        End If
        Lamination_Scribl = s

    End Sub

    '++ID 12.09.2021 Foil_Color
    Public Sub AddFoil_Color(ByRef pValue As String, ByRef pstrFoil_Color As String)

        Dim s As String
        s = Trim(pstrFoil_Color)
        If Not (InStr(pstrFoil_Color, pValue)) > 0 Then
            '   If s <> "" Then s = s & "/"
            '  s = s & Trim(pValue)
            s = Trim(pValue)
            ' By using the property, it saves.
        End If
        Foil_Color = s

    End Sub

    Public Sub AddThreadColor_Scribl(ByRef pValue As String, ByRef pstrThreadColorScribl As String)

        Dim s As String
        s = Trim(pstrThreadColorScribl)
        If Not (InStr(pstrThreadColorScribl, pValue)) > 0 Then
            '    If s <> "" Then s = s & "/"
            ' s = s & Trim(pValue)
            s = Trim(pValue)
            ' By using the property, it saves.
        End If
        Thread_Color_Scribl = s

    End Sub

    Public Sub AddRoundedCorners(ByRef pValue As String, ByRef pstrRoundedCorners As String)

        Dim s As String
        s = Trim(pstrRoundedCorners)
        If Not (InStr(pstrRoundedCorners, pValue)) > 0 Then
            '   If s <> "" Then s = s & "/"
            ' s = s & Trim(pValue)
            s = Trim(pValue)
            ' By using the property, it saves.
        End If
        Rounded_corners_Scribl = s

    End Sub

    '++ID 08.06.2020
    Public Sub AddTip_In_PageTxt(ByRef pValue As String, ByRef pstrTip_In_PageTxt As String)

        Dim s As String
        s = Trim(pstrTip_In_PageTxt)
        If Not (InStr(pstrTip_In_PageTxt, pValue)) > 0 Then
            '   If s <> "" Then s = s & "/"
            's = s & Trim(pValue)
            s = Trim(pValue)
            ' By using the property, it saves.
        End If
        Tip_In_PageTxt = s

    End Sub
    '-------------------------------------------------------------

    Public Sub AddPackaging(ByRef pValue As String, ByRef pstrPackaging As String)

        Dim s As String
        s = Trim(pstrPackaging)
        If Not (InStr(pstrPackaging, pValue)) > 0 Then
            If s <> "" Then s = s & "/"
            s = s & Trim(pValue)
            ' By using the property, it saves.
        End If
        Packaging = s

    End Sub

    Public Sub AddRefill(ByRef pValue As String, ByRef pstrRefill As String)

        Dim s As String
        s = Trim(pstrRefill)
        If Not (InStr(pstrRefill, pValue)) > 0 Then
            If s <> "" Then s = s & "/"
            s = s & Trim(pValue)
            ' By using the property, it saves.
        End If
        Refill = s

    End Sub




    Public Shared Function getAllowableRouteList() As ArrayList
        Dim allowableRoutes As ArrayList = New ArrayList

        Try

            Dim conn As New cDBA
            Dim dtRoutes As New DataTable

            ' MB++ Faire aussi un DELETE de masse pour le master Item_Guid, bien que pas 
            ' necessairement besoin vu que ca sera probablement appelé d'un niveau plus haut.

            Dim strSql As String = _
            "SELECT     RouteId " & _
            "FROM       EXACT_TRAVELER_ROUTE " & _
            "WHERE      Active = 1 AND ThermoSimEligible = 1"

            dtRoutes = conn.DataTable(strSql)
            If dtRoutes.Rows.Count <> 0 Then
                For Each row As DataRow In dtRoutes.Rows
                    allowableRoutes.Add(Integer.Parse(row.Item("RouteId").ToString))
                Next

            End If

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        Return IIf(allowableRoutes.Count > 0, allowableRoutes, Nothing)
    End Function

#End Region

End Class
