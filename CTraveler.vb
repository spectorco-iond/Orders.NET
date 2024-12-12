Public Class CTraveler

    ' m_HeaderID is pulled from system after first save.
    Private m_HeaderID As Integer '  IDENTITY (1, 1) NOT NULL ,
    Private m_RouteID As Integer '  NULL ,
    Private m_Cus_No As String '  (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,

    ' Ord_No is pushed via trigger on import, for sync purposes only.
    Private m_Ord_No As String '  (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,

    Private m_Ord_Guid As String
    Private m_Component_Seq_No As Integer '  NULL CONSTRAINT [DF_EXACT_TRAVELER_HEADER_Component_seq_no DEFAULT (0),

    ' Line_No is pushed via trigger on import, for sync purposes.
    Private m_Line_No As Integer ' smallint ' NOT NULL ,

    Private m_Item_No As String '  (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Private m_Head_Extra1 As String '  (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Private m_NumTravelers As Integer '  NULL ,
    Private m_SlsPsnName As String '  (64) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Private m_CreditManager As String '  (64) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Private m_Parent_Item_No As String '  (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Private m_Parent_Orig_Qty_Ord As Double '(13, 4) NULL ,
    Private m_Orig_Qty_Ord As Double '(13, 4) NULL CONSTRAINT [DF_EXACT_TRAVELER_HEADER_Orig_Qty_Ord DEFAULT (0),
    Private m_OrderStatus As String '  (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Private m_RushOrder As Integer ' smallint ' NULL CONSTRAINT [DF_EXACT_TRAVELER_HEADER_RushOrder DEFAULT (0),
    Private m_ManagerOverridePriority As Integer ' smallint ' NULL CONSTRAINT [DF_EXACT_TRAVELER_HEADER_ManagerOverridePriority DEFAULT (0),

    Private m_Item_Guid As String ' Before import, use this

    ' OEOrdLin_ID is pushed via trigger on import, for sync purposes only.
    Private m_Oeordlin_ID As Double ' After import,  use this

    Private m_TS As Date

    Private m_TravelerDetails As CTravelerDetails
    Private m_TravelerDocument As CTravelerDocument
    Private m_TravelerDocuments As Collection
    Private m_TravelerExtraFunctions As CTravelerExtraFunctions
    Private m_TravelerNewOEEntries As CTravelerNewOEEntries
    Private m_TravelerOrderAckAudit As CTravelerOrderAckAudit

    Private m_Dirty As Boolean = False
    Private m_SaveToDB As Boolean = False

    Public Class CTravelerDetails

        Private m_ID As Integer '  IDENTITY (1, 1) NOT NULL ,
        'Private m_HeaderID As Integer '  NULL ,
        'Private m_Ord_No As String '  (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_StepID As Integer '  NOT NULL ,
        Private m_RestartCount As Integer '  NOT NULL CONSTRAINT [DF_EXACT_TRAVELER_DETAILS_RestartCount DEFAULT (0),
        Private m_Sequence As Integer '  NULL ,
        Private m_Description As String '  (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_StepCategory As String '  (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_UserID As String '  (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_StartDate As Date '  NULL ,
        Private m_CompleteDate As Date '  NULL ,
        Private m_QtyCompleted As Integer '  NULL ,
        Private m_QtyScrap As Integer '  NULL ,
        Private m_Comments As String '  (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_Step_Extra1 As String '  (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_Step_Extra2 As String '  (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_Step_Extra3 As String '  (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_RouteID As Integer '  NULL ,
        Private m_RouteCategory As String '  (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_RouteDescription As String '  (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_UserID_SCRAP As String '  (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_PO_Transferred As String '  (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_OE_Inv_No As String '  (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_PickStatus As String '  (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_OrderStatus As String '  (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_ScrapCalcDate As Date '  NULL ,

#Region "PublicProperties ################################################"

        Public Property Comments() As String
            Get
                Comments = m_Comments
            End Get
            Set(ByVal Value As String)
                m_Comments = Value
            End Set
        End Property
        Public Property CompleteDate() As Date '
            Get
                CompleteDate = m_CompleteDate
            End Get
            Set(ByVal Value As Date)
                m_CompleteDate = Value
            End Set
        End Property
        Public Property Description() As String
            Get
                Description = m_Description
            End Get
            Set(ByVal Value As String)
                m_Description = Value
            End Set
        End Property
        Public Property ID() As Integer
            Get
                ID = m_ID
            End Get
            Set(ByVal Value As Integer)
                m_ID = Value
            End Set
        End Property
        Public Property OE_Inv_No() As String
            Get
                OE_Inv_No = m_OE_Inv_No
            End Get
            Set(ByVal Value As String)
                m_OE_Inv_No = Value
            End Set
        End Property
        Public Property OrderStatus() As String
            Get
                OrderStatus = m_OrderStatus
            End Get
            Set(ByVal Value As String)
                m_OrderStatus = Value
            End Set
        End Property
        Public Property PickStatus() As String
            Get
                PickStatus = m_PickStatus
            End Get
            Set(ByVal Value As String)
                m_PickStatus = Value
            End Set
        End Property
        Public Property PO_Transferred() As String
            Get
                PO_Transferred = m_PO_Transferred
            End Get
            Set(ByVal Value As String)
                m_PO_Transferred = Value
            End Set
        End Property
        Public Property QtyCompleted() As Integer
            Get
                QtyCompleted = m_QtyCompleted
            End Get
            Set(ByVal Value As Integer)
                m_QtyCompleted = Value
            End Set
        End Property
        Public Property QtyScrap() As Integer
            Get
                QtyScrap = m_QtyScrap
            End Get
            Set(ByVal Value As Integer)
                m_QtyScrap = Value
            End Set
        End Property
        Public Property RestartCount() As Integer
            Get
                RestartCount = m_RestartCount
            End Get
            Set(ByVal Value As Integer)
                m_RestartCount = Value
            End Set
        End Property
        Public Property RouteCategory() As String
            Get
                RouteCategory = m_RouteCategory
            End Get
            Set(ByVal Value As String)
                m_RouteCategory = Value
            End Set
        End Property
        Public Property RouteDescription() As String
            Get
                RouteDescription = m_RouteDescription
            End Get
            Set(ByVal Value As String)
                m_RouteDescription = Value
            End Set
        End Property
        Public Property RouteID() As Integer
            Get
                RouteID = m_RouteID
            End Get
            Set(ByVal Value As Integer)
                m_RouteID = Value
            End Set
        End Property
        Public Property ScrapCalcDate() As DateTime
            Get
                ScrapCalcDate = m_ScrapCalcDate
            End Get
            Set(ByVal Value As DateTime)
                m_ScrapCalcDate = Value
            End Set
        End Property
        Public Property Sequence() As Integer
            Get
                Sequence = m_Sequence
            End Get
            Set(ByVal Value As Integer)
                m_Sequence = Value
            End Set
        End Property
        Public Property StartDate() As Date '
            Get
                StartDate = m_StartDate
            End Get
            Set(ByVal Value As DateTime)
                m_StartDate = Value
            End Set
        End Property
        Public Property StepCategory() As String
            Get
                StepCategory = m_StepCategory
            End Get
            Set(ByVal Value As String)
                m_StepCategory = Value
            End Set
        End Property
        Public Property Step_Extra1() As String
            Get
                Step_Extra1 = m_Step_Extra1
            End Get
            Set(ByVal Value As String)
                m_Step_Extra1 = Value
            End Set
        End Property
        Public Property Step_Extra2() As String
            Get
                Step_Extra2 = m_Step_Extra2
            End Get
            Set(ByVal Value As String)
                m_Step_Extra2 = Value
            End Set
        End Property
        Public Property Step_Extra3() As String
            Get
                Step_Extra3 = m_Step_Extra3
            End Get
            Set(ByVal Value As String)
                m_Step_Extra3 = Value
            End Set
        End Property
        Public Property StepID() As Integer
            Get
                StepID = m_StepID
            End Get
            Set(ByVal Value As Integer)
                m_StepID = Value
            End Set
        End Property
        Public Property UserID() As String
            Get
                UserID = m_UserID
            End Get
            Set(ByVal Value As String)
                m_UserID = Value
            End Set
        End Property
        Public Property UserID_SCRAP() As String
            Get
                UserID_SCRAP = m_UserID_SCRAP
            End Get
            Set(ByVal Value As String)
                m_UserID_SCRAP = Value
            End Set
        End Property

#End Region

    End Class

    Public Class CTravelerDocument

        Private m_DocID As Integer '  IDENTITY (1, 1) NOT NULL ,
        'Private m_Ord_no As String '  (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        'Private m_HeaderID As Integer '  NOT NULL ,
        Private m_DocType As String '  (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
        Private m_DocDescription As String '  (400) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
        Private m_DocName As String '  (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
        Private m_Document As Object ' [image  NOT NULL ,
        Private m_CreateDate As Date '  NULL ,
        Private m_ArtworkPurged As Integer ' smallint ' NOT NULL CONSTRAINT [DF_EXACT_TRAVELER_DOCUMENT_ArtworkPurged DEFAULT (0),

#Region "Public Properties ################################################"

        Public Property ArtworkPurged() As Integer
            Get
                ArtworkPurged = m_ArtworkPurged
            End Get
            Set(ByVal Value As Integer)
                m_ArtworkPurged = Value
            End Set
        End Property
        Public Property CreateDate() As Date  '
            Get
                CreateDate = m_CreateDate
            End Get
            Set(ByVal Value As Date)
                m_CreateDate = Value
            End Set
        End Property
        Public Property DocDescription() As String
            Get
                DocDescription = m_DocDescription
            End Get
            Set(ByVal Value As String)
                m_DocDescription = Value
            End Set
        End Property
        Public Property DocID() As Integer
            Get
                DocID = m_DocID
            End Get
            Set(ByVal Value As Integer)
                m_DocID = Value
            End Set
        End Property
        Public Property DocName() As String
            Get
                DocName = m_DocName
            End Get
            Set(ByVal Value As String)
                m_DocName = Value
            End Set
        End Property
        Public Property DocType() As String
            Get
                DocType = m_DocType
            End Get
            Set(ByVal Value As String)
                m_DocType = Value
            End Set
        End Property
        Public Property Document() As Object
            Get
                Document = m_Document
            End Get
            Set(ByVal Value As Object)
                m_Document = Value
            End Set
        End Property

#End Region

    End Class

    Public Class CTravelerExtraFunctions

        'Private m_ord_no As String '  (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
        'Private m_HeaderID As Integer '  NOT NULL ,
        Private m_SendPreShippingAdvice As Integer ' smallint ' NULL CONSTRAINT [DF_EXACT_TRAVELER_EXTRA_FUNCTIONS_SendPreShippingAdvice DEFAULT (1),
        Private m_Send90DayReminder As Integer ' smallint ' NULL CONSTRAINT [DF_EXACT_TRAVELER_EXTRA_FUNCTIONS_Send90DayReminder DEFAULT (1)

#Region "Public Properties ################################################"

        Public Property SendPreShippingAdvice() As Integer
            Get
                SendPreShippingAdvice = m_SendPreShippingAdvice
            End Get
            Set(ByVal Value As Integer)
                m_SendPreShippingAdvice = Value
            End Set
        End Property
        Public Property Send90DayReminder() As Integer
            Get
                Send90DayReminder = m_Send90DayReminder
            End Get
            Set(ByVal Value As Integer)
                m_Send90DayReminder = Value
            End Set
        End Property

#End Region

    End Class

    Public Class CTravelerNewOEEntries

        'Private OrderNo As String '  (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
        Private m_Oeordlin_ID As Double '(10, 0) NOT NULL ,
        'Private m_CusNo As String '  (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
        Private m_AddingUser As String '  (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
        Private m_DateAdded As Date '  NOT NULL CONSTRAINT [DF_EXACT_TRAVELER_NEW_OE_ENTRIES_DateAdded DEFAULT (getdate())

#Region "Public Properties ################################################"

        Public Property AddingUser() As String
            Get
                AddingUser = m_AddingUser
            End Get
            Set(ByVal Value As String)
                m_AddingUser = Value
            End Set
        End Property
        Public Property DateAdded() As Date
            Get
                DateAdded = m_DateAdded
            End Get
            Set(ByVal Value As Date)
                m_DateAdded = Value
            End Set
        End Property

#End Region

    End Class

    Public Class CTravelerOrderAckAudit

        Private m_id As Integer '  IDENTITY (1, 1) NOT NULL ,
        'Private m_ord_no As String '  (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
        Private m_dateSent As Date '  NOT NULL CONSTRAINT [DF_EXACT_TRAVELER_ORDER_ACK_AUDIT_dateSent DEFAULT (getdate()),
        Private m_sent As Integer ' smallint ' NOT NULL CONSTRAINT [DF_EXACT_TRAVELER_ORDER_ACK_AUDIT_sent DEFAULT (0),

#Region "Public Properties ################################################"

        Public Property DateSent() As Date
            Get
                DateSent = m_dateSent
            End Get
            Set(ByVal Value As Date)
                m_dateSent = Value
            End Set
        End Property
        Public Property Id() As Integer
            Get
                Id = m_id
            End Get
            Set(ByVal Value As Integer)
                m_id = Value
            End Set
        End Property
        Public Property Sent() As Integer
            Get
                Sent = m_sent
            End Get
            Set(ByVal Value As Integer)
                m_sent = Value
            End Set
        End Property

#End Region

    End Class

#Region "Constructors #####################################################"

    Public Sub New()

        m_Ord_Guid = String.Empty
        m_Item_Guid = String.Empty

        Call InitializeSubClasses()

    End Sub

    Public Sub New(ByVal pstrItem_Guid As String)

        m_Ord_Guid = m_oOrder.Ordhead.Ord_GUID ' String.Empty
        m_Item_Guid = pstrItem_Guid

        Call InitializeSubClasses()

        Call LoadDefaultValues()

    End Sub

    Public Sub New(ByVal pstrOrd_Guid As String, ByVal pstrItem_Guid As String)

        m_Ord_Guid = pstrOrd_Guid
        m_Item_Guid = pstrItem_Guid

        Call InitializeSubClasses()

        Call LoadDefaultValues()

    End Sub

#End Region

    Private Sub LoadDefaultValues()

        Try

            If m_Ord_No Is Nothing Then Exit Sub

            Dim strsql As String = _
            "SELECT     H.Cus_No, L.* " & _
            "FROM       OEI_ORDHDR H WITH (Nolock) " & _
            "INNER JOIN OEI_ORDLIN L WITH (Nolock) ON H.Ord_Guid = L.Ord_Guid " & _
            "WHERE      H.Ord_Guid = '" & m_Ord_No & "' AND L.Item_Guid = '" & m_Item_Guid & "' "

            Dim conn As New cDBA
            Dim dt As New DataTable

            dt = conn.DataTable(strsql)

            If dt.Rows.Count <> 0 Then
                Dim dr As DataRow
                dr = dt.Rows(0)

                m_Ord_No = IIf(dr.Item("Ord_No").Equals(DBNull.Value), "", dr.Item("Ord_No"))
                m_Cus_No = IIf(dr.Item("Cus_No").Equals(DBNull.Value), "", dr.Item("Cus_No"))
                m_HeaderID = IIf(dr.Item("HeaderID").Equals(DBNull.Value), 0, dr.Item("HeaderID"))
                m_RouteID = IIf(dr.Item("RouteID").Equals(DBNull.Value), 0, dr.Item("RouteID"))
                m_Ord_Guid = IIf(dr.Item("Ord_Guid").Equals(DBNull.Value), "", dr.Item("Ord_Guid"))
                m_Component_Seq_No = IIf(dr.Item("Component_Seq_No").Equals(DBNull.Value), 0, dr.Item("Component_Seq_No"))
                m_Line_No = IIf(dr.Item("Line_No").Equals(DBNull.Value), 0, dr.Item("Line_No"))
                m_Item_No = IIf(dr.Item("Item_No").Equals(DBNull.Value), "", dr.Item("Item_No"))
                m_Head_Extra1 = IIf(dr.Item("Head_Extra1").Equals(DBNull.Value), "", dr.Item("Head_Extra1"))
                m_NumTravelers = IIf(dr.Item("NumTravelers").Equals(DBNull.Value), 0, dr.Item("NumTravelers"))
                m_SlsPsnName = IIf(dr.Item("SlsPsnName").Equals(DBNull.Value), "", dr.Item("SlsPsnName"))
                m_CreditManager = IIf(dr.Item("CreditManager").Equals(DBNull.Value), "", dr.Item("CreditManager"))
                m_Parent_Item_No = IIf(dr.Item("Parent_Item_No").Equals(DBNull.Value), "", dr.Item("RouteID"))
                m_Parent_Orig_Qty_Ord = IIf(dr.Item("Parent_Orig_Qty_Ord").Equals(DBNull.Value), 0, dr.Item("Parent_Orig_Qty_Ord"))
                m_Orig_Qty_Ord = IIf(dr.Item("Orig_Qty_Ord").Equals(DBNull.Value), 0, dr.Item("Orig_Qty_Ord"))
                m_OrderStatus = IIf(dr.Item("OrderStatus").Equals(DBNull.Value), "", dr.Item("OrderStatus"))
                m_RushOrder = IIf(dr.Item("RouteID").Equals(DBNull.Value), 0, dr.Item("RushOrder"))
                m_ManagerOverridePriority = IIf(dr.Item("ManagerOverridePriority").Equals(DBNull.Value), 0, dr.Item("ManagerOverridePriority"))
                m_Item_Guid = IIf(dr.Item("RouteID").Equals(DBNull.Value), "", dr.Item("RouteID"))
                m_RouteID = IIf(dr.Item("Item_Guid").Equals(DBNull.Value), 0, dr.Item("Item_Guid"))
                m_Oeordlin_ID = IIf(dr.Item("OEOrdLin_ID").Equals(DBNull.Value), 0, dr.Item("OEOrdLin_ID"))

            Else
            End If

        Catch er As Exception
            MsgBox(er.Message)
        End Try

    End Sub

    Private Sub InitializeSubClasses()

        m_TravelerDetails = New CTravelerDetails
        m_TravelerDocument = New CTravelerDocument
        m_TravelerDocuments = New Collection
        m_TravelerExtraFunctions = New CTravelerExtraFunctions
        m_TravelerNewOEEntries = New CTravelerNewOEEntries
        m_TravelerOrderAckAudit = New CTravelerOrderAckAudit

    End Sub


#Region "Public I/O procedures ############################################"

    Public Sub Delete()

        'Try
        '    Dim conn As New cDBA
        '    Dim dt As New DataTable

        '    ' MB++ Faire aussi un DELETE de masse pour le master Item_Guid, bien que pas 
        '    ' necessairement besoin vu que ca sera probablement appelé d'un niveau plus haut.

        '    Dim strSql As String = _
        '    "SELECT     * " & _
        '    "FROM       OEI_EXTRA_FIELDS " & _
        '    "WHERE      TRIM(Ord_Guid) = '" & Trim(Ord_Guid) & "' AND " & _
        '    "           TRIM(Item_Guid) = '" & Trim(Item_Guid) & "' "

        '    dt = conn.DataTable(strSql)
        '    If dt.Rows.Count <> 0 Then
        '        conn.DBDataTable.Rows.RemoveAt(0)
        '        Dim cmd As OleDb.OleDbCommandBuilder = New OleDb.OleDbCommandBuilder(conn.DBDataAdapter)
        '        conn.DBDataAdapter.DeleteCommand = cmd.GetDeleteCommand
        '    End If

        '    conn.DBDataAdapter.Update(conn.DBDataTable)

        'Catch er As Exception
        '    MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

    Public Sub Load()

        'Try
        '    Dim conn As New cDBA
        '    Dim dt As New DataTable
        '    Dim drDataRow As DataRow

        '    Dim strSql As String = _
        '    "SELECT     * " & _
        '    "FROM       OEI_EXTRA_FIELDS WITH (Nolock) " & _
        '    "WHERE      TRIM(Ord_Guid) = '" & Trim(Ord_Guid) & "' AND " & _
        '    "           TRIM(Item_Guid) = '" & Trim(Item_Guid) & "' "

        '    dt = conn.DataTable(strSql)
        '    If dt.Rows.Count <> 0 Then

        '        drDataRow = dt.Rows(0)
        '        With drDataRow

        '            m_Imprint = IIf(.Item("Extra_1").Equals(DBNull.Value), "", .Item("Extra_1"))
        '            m_Location = IIf(.Item("Extra_2").Equals(DBNull.Value), "", .Item("Extra_2"))
        '            m_Color = IIf(.Item("Extra_3").Equals(DBNull.Value), "", .Item("Extra_3"))
        '            m_Num = IIf(.Item("Extra_4").Equals(DBNull.Value), 0, .Item("Extra_4"))
        '            m_Packaging = IIf(.Item("Extra_5").Equals(DBNull.Value), "", .Item("Extra_5"))
        '            'drDataRow.Item("Extra_6") = Item_Guid ' Extra_6 n'est pas utilisé
        '            'drDataRow.Item("Extra_7") = Item_Guid ' Extra_7 n'est pas utilisé
        '            m_Refill = IIf(.Item("Extra_8").Equals(DBNull.Value), "", .Item("Extra_8"))
        '            m_LaserSetup = IIf(.Item("Extra_9").Equals(DBNull.Value), "", .Item("Extra_9"))
        '            m_Comments = IIf(.Item("Comment").Equals(DBNull.Value), "", .Item("Comment"))
        '            m_Special_Comments = IIf(.Item("Comment2").Equals(DBNull.Value), "", .Item("Comment2"))
        '            'drDataRow.Item("PrimaryContact") = m_PrimaryContact
        '            'drDataRow.Item("DateAdded") = DateAdded

        '        End With

        '    End If
        '    dt.Dispose()

        'Catch er As Exception
        '    MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

    Public Sub Reset()

        'm_Imprint = String.Empty
        'm_Location = String.Empty
        'm_Num = 0 ' String.Empty
        'm_Color = String.Empty
        'm_Packaging = String.Empty
        'm_Refill = String.Empty
        'm_LaserSetup = String.Empty
        'm_Comments = String.Empty
        'm_Special_Comments = String.Empty

    End Sub

    Public Sub Save()

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        If Not m_Dirty Then Exit Sub
        If Not m_SaveToDB Then Exit Sub

        'Try
        '    Dim conn As New cDBA
        '    Dim dt As New DataTable
        '    Dim drDataRow As DataRow

        '    Dim strSql As String = _
        '    "SELECT     * " & _
        '    "FROM       OEI_EXTRA_FIELDS " & _
        '    "WHERE      Ord_Guid = '" & Trim(Ord_Guid) & "' AND " & _
        '    "           Item_Guid = '" & Trim(Item_Guid) & "' "

        '    dt = conn.DataTable(strSql)
        '    If dt.Rows.Count <> 0 Then
        '        drDataRow = dt.Rows(0)
        '    Else
        '        drDataRow = dt.NewRow
        '        drDataRow.Item("Ord_Guid") = Ord_Guid
        '        drDataRow.Item("Item_Guid") = Item_Guid
        '    End If

        '    drDataRow.Item("Extra_1") = m_Imprint
        '    drDataRow.Item("Extra_2") = m_Location
        '    drDataRow.Item("Extra_3") = m_Color
        '    drDataRow.Item("Extra_4") = m_Num
        '    drDataRow.Item("Extra_5") = m_Packaging
        '    'drDataRow.Item("Extra_6") = Item_Guid
        '    'drDataRow.Item("Extra_7") = Item_Guid
        '    drDataRow.Item("Extra_8") = m_Refill
        '    drDataRow.Item("Extra_9") = m_LaserSetup
        '    drDataRow.Item("Comment") = m_Comments
        '    drDataRow.Item("Comment2") = m_Special_Comments

        '    'drDataRow.Item("PrimaryContact") = m_PrimaryContact
        '    'drDataRow.Item("DateAdded") = DateAdded

        '    If dt.Rows.Count = 0 Then
        '        conn.DBDataTable.Rows.Add(drDataRow)
        '        Dim cmd As OleDb.OleDbCommandBuilder = New OleDb.OleDbCommandBuilder(conn.DBDataAdapter)
        '        conn.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
        '    Else
        '        Dim cmd As OleDb.OleDbCommandBuilder = New OleDb.OleDbCommandBuilder(conn.DBDataAdapter)
        '        conn.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
        '    End If

        '    conn.DBDataAdapter.Update(conn.DBDataTable)

        'Catch er As Exception
        '    MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

    'Private Sub ComboFill()

    '    Dim strSql As String
    '    Dim conn As New cDBA
    '    Dim dtCombo As New DataTable

    '    Try
    '        strSql = "Select * from EXACT_TRAVELER_EXTRA_FIELDS_xRef order by FieldValue Asc "
    '        dtCombo = conn.DataTable(strSql)

    '        If dtCombo.Rows.Count <> 0 Then

    '            Dim drComboElement As DataRow

    '            For Each drComboElement In dtCombo.Rows

    '                Select Case drComboElement.Item("Field")
    '                    Case "ImprintLocation"
    '                cboImprintLocation.AddItem !FieldValue & ""
    '                    Case "ImprintColour"
    '                cboImprintColour.AddItem !FieldValue & ""
    '                    Case "Packaging"
    '                cboPackaging.AddItem !FieldValue & ""
    '                    Case "Refill"
    '                cboRefill.AddItem !FieldValue & ""
    '                End Select

    '            Next

    '        End If

    '    Catch er As Exception
    '        MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

#End Region

#Region "Public Properties ################################################"

    Public Property Component_Seq_No() As Integer
        Get
            Component_Seq_No = m_Component_Seq_No
        End Get
        Set(ByVal Value As Integer)
            m_Component_Seq_No = Value
        End Set
    End Property
    Public Property CreditManager() As String
        Get
            CreditManager = m_CreditManager
        End Get
        Set(ByVal Value As String)
            m_CreditManager = Value
        End Set
    End Property
    Public Property Cus_No() As String
        Get
            Cus_No = m_Cus_No
        End Get
        Set(ByVal Value As String)
            m_Cus_No = Value
        End Set
    End Property
    Public Property Head_Extra1() As String
        Get
            Head_Extra1 = m_Head_Extra1
        End Get
        Set(ByVal Value As String)
            m_Head_Extra1 = Value
        End Set
    End Property
    Public Property HeaderID() As Integer
        Get
            HeaderID = m_HeaderID
        End Get
        Set(ByVal Value As Integer)
            m_HeaderID = Value
        End Set
    End Property
    Public Property Item_No() As String
        Get
            Item_No = m_Item_No
        End Get
        Set(ByVal Value As String)
            m_Item_No = Value
        End Set
    End Property
    Public Property Line_No() As Integer
        Get
            Line_No = m_Line_No
        End Get
        Set(ByVal Value As Integer)
            m_Line_No = Value
        End Set
    End Property
    Public Property ManagerOverridePriority() As Integer
        Get
            ManagerOverridePriority = m_ManagerOverridePriority
        End Get
        Set(ByVal Value As Integer)
            m_ManagerOverridePriority = Value
        End Set
    End Property
    Public Property NumTravelers() As Integer
        Get
            NumTravelers = m_NumTravelers
        End Get
        Set(ByVal Value As Integer)
            m_NumTravelers = Value
        End Set
    End Property
    Public Property Oeordlin_ID() As Double
        Get
            Oeordlin_ID = m_Oeordlin_ID
        End Get
        Set(ByVal Value As Double)
            m_Oeordlin_ID = Value
        End Set
    End Property
    Public Property Ord_No() As String
        Get
            Ord_No = m_Ord_No
        End Get
        Set(ByVal Value As String)
            m_Ord_No = Value
        End Set
    End Property
    Public Property OrderStatus() As String
        Get
            OrderStatus = m_OrderStatus
        End Get
        Set(ByVal Value As String)
            m_OrderStatus = Value
        End Set
    End Property
    Public Property Item_Guid() As String
        Get
            Item_Guid = m_Item_Guid
        End Get
        Set(ByVal Value As String)
            m_Item_Guid = Value
        End Set
    End Property
    Public Property Orig_Qty_Ord() As Double
        Get
            Orig_Qty_Ord = m_Orig_Qty_Ord
        End Get
        Set(ByVal Value As Double)
            m_Orig_Qty_Ord = Value
        End Set
    End Property
    Public Property Parent_Item_No() As String
        Get
            Parent_Item_No = m_Parent_Item_No
        End Get
        Set(ByVal Value As String)
            m_Parent_Item_No = Value
        End Set
    End Property
    Public Property Parent_Orig_Qty_Ord() As Double
        Get
            Parent_Orig_Qty_Ord = m_Parent_Orig_Qty_Ord
        End Get
        Set(ByVal Value As Double)
            m_Parent_Orig_Qty_Ord = Value
        End Set
    End Property
    Public Property RouteID() As Integer
        Get
            RouteID = m_RouteID
        End Get
        Set(ByVal Value As Integer)
            m_RouteID = Value
        End Set
    End Property
    Public Property RushOrder() As Integer
        Get
            RushOrder = m_RushOrder
        End Get
        Set(ByVal Value As Integer)
            m_RushOrder = Value
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
    Public Property SlsPsnName() As String
        Get
            SlsPsnName = m_SlsPsnName
        End Get
        Set(ByVal Value As String)
            m_SlsPsnName = Value
        End Set
    End Property
    Public Property TravelerDetails() As CTravelerDetails
        Get
            TravelerDetails = m_TravelerDetails
        End Get
        Set(ByVal value As CTravelerDetails)
            m_TravelerDetails = value
        End Set
    End Property
    Public Property TravelerDocument() As CTravelerDocument
        Get
            TravelerDocument = m_TravelerDocument
        End Get
        Set(ByVal value As CTravelerDocument)
            m_TravelerDocument = value
        End Set
    End Property
    Public Property TravelerDocuments() As Collection
        Get
            TravelerDocuments = m_TravelerDocuments
        End Get
        Set(ByVal value As Collection)
            m_TravelerDocuments = value
        End Set
    End Property
    Public Property TravelerExtraFunctions() As CTravelerExtraFunctions
        Get
            TravelerExtraFunctions = m_TravelerExtraFunctions
        End Get
        Set(ByVal value As CTravelerExtraFunctions)
            m_TravelerExtraFunctions = value
        End Set
    End Property
    Public Property TravelerNewOEEntries() As CTravelerNewOEEntries
        Get
            TravelerNewOEEntries = m_TravelerNewOEEntries
        End Get
        Set(ByVal value As CTravelerNewOEEntries)
            m_TravelerNewOEEntries = value
        End Set
    End Property
    Public Property TravelerOrderAckAudit() As CTravelerOrderAckAudit
        Get
            TravelerOrderAckAudit = m_TravelerOrderAckAudit
        End Get
        Set(ByVal value As CTravelerOrderAckAudit)
            m_TravelerOrderAckAudit = value
        End Set
    End Property

#End Region

End Class
