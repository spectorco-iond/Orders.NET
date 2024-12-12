Public Class cKit

    Private m_Ord_Guid As String
    Private m_Item_Guid As String
    Private m_Source As OEI_SourceEnum

    Private m_Curr_Cd As String
    Private m_Cus_No As String
    Private m_Item_Desc_1 As String
    Private m_Item_Desc_2 As String
    Private m_Item_No As String
    Private m_Item_Price As Double
    Private m_Loc As String = "1"
    Private m_Sql As String
    Private m_IsAKit As Boolean = False
    Private m_Line_No As Integer
    Private m_OEOrdLin_ID As Integer
    Private m_Unit_Cost As Double
    Private m_Stocked_Fg As String
    Private m_Controlled_Fg As String
    Private m_Mfg_Method As String

    ' These fields will calculate the accurate inventory according to the components.
    Private m_Inventory As Double
    Private m_Qty_Allocated As Double
    Private m_Qty_On_Hand As Double

    Private m_Item_Cd As String
    Private m_Extra_11 As Integer

    'Private m_PO_Parent As cOrdLine

    'Private m_OrdLin As cOrdLine

    Public Items As Collection ' of cKitItem

    Public Class cKitItem

        'Ord_Type, Ord_Guid, both not mandatory
        'Kit_Item_Guid, Comp_Item_Guid, Seq_No, Last_Cost, 
        'Stocked_Fg, Controlled_Fg, Qty_Per_Par

        Private m_Comp_Item_Guid As String
        Private m_Comp_Item_No As String
        Private m_Controlled_Fg As String
        Private m_Cus_No As String
        Private m_Item_Desc_1 As String
        Private m_Item_Desc_2 As String
        'Private m_Item_Guid As String
        Private m_Item_No As String
        Private m_Item_Price As Double
        Private m_Kit_Component As Boolean
        Private m_Kit_Item_Guid As String
        Private m_Last_Cost As Double
        Private m_Loc As String
        Private m_Qty_Per_Par As Double
        Private m_Seq_No As Integer
        Private m_Stocked_Fg As String
        Private m_Kit_Item_Cd As String
        Private m_Comp_Item_Cd As String

        'Private m_Item As cOrdLine

        Public Sub New()

            'm_Item = New cOrdLine(false)

        End Sub

#Region "Public properties ################################################"

        Public Property Comp_Item_Guid() As String
            Get
                Comp_Item_Guid = m_Comp_Item_Guid
            End Get
            Set(ByVal value As String)
                m_Comp_Item_Guid = value
            End Set
        End Property
        Public Property Comp_Item_Cd() As String
            Get
                Comp_Item_Cd = Mid(m_Comp_Item_Cd, 1, 10)
            End Get
            Set(ByVal value As String)
                m_Comp_Item_Cd = Mid(value, 1, 10)
            End Set
        End Property
        Public Property Comp_Item_No() As String
            Get
                Comp_Item_No = m_Comp_Item_No
            End Get
            Set(ByVal value As String)
                m_Comp_Item_No = value
            End Set
        End Property
        Public Property Controlled_Fg() As String
            Get
                Controlled_Fg = m_Controlled_Fg
            End Get
            Set(ByVal value As String)
                m_Controlled_Fg = value
            End Set
        End Property
        Public Property Cus_No() As String
            Get
                Cus_No = m_Cus_No
            End Get
            Set(ByVal value As String)
                m_Cus_No = value
            End Set
        End Property
        'Public Property Item() As cOrdLine
        '    Get
        '        Item = m_Item
        '    End Get
        '    Set(ByVal value As cOrdLine)
        '        m_Item = value
        '    End Set
        'End Property
        Public Property Item_Desc_1() As String
            Get
                Item_Desc_1 = m_Item_Desc_1
            End Get
            Set(ByVal value As String)
                m_Item_Desc_1 = value
            End Set
        End Property
        Public Property Item_Desc_2() As String
            Get
                Item_Desc_2 = m_Item_Desc_2
            End Get
            Set(ByVal value As String)
                m_Item_Desc_2 = value
            End Set
        End Property
        'Public Property Item_Guid() As String
        '    Get
        '        Item_Guid = m_Item_Guid
        '    End Get
        '    Set(ByVal value As String)
        '        m_Item_Guid = value
        '    End Set
        'End Property
        Public Property Item_No() As String
            Get
                Item_No = m_Item_No
            End Get
            Set(ByVal value As String)
                m_Item_No = value
            End Set
        End Property
        Public Property Item_Price() As Double
            Get
                Item_Price = m_Item_Price
            End Get
            Set(ByVal value As Double)
                m_Item_Price = value
            End Set
        End Property
        Public Property Kit_Component() As Boolean
            Get
                Kit_Component = m_Kit_Component
            End Get
            Set(ByVal value As Boolean)
                m_Kit_Component = value
            End Set
        End Property
        Public Property Kit_Item_Cd() As String
            Get
                Kit_Item_Cd = Mid(m_Kit_Item_Cd, 1, 10)
            End Get
            Set(ByVal value As String)
                m_Kit_Item_Cd = Mid(value, 1, 10)
            End Set
        End Property
        Public Property Kit_Item_Guid() As String
            Get
                Kit_Item_Guid = m_Kit_Item_Guid
            End Get
            Set(ByVal value As String)
                m_Kit_Item_Guid = value
            End Set
        End Property
        Public Property Last_Cost() As Double
            Get
                Last_Cost = m_Last_Cost
            End Get
            Set(ByVal value As Double)
                m_Last_Cost = value
            End Set
        End Property
        Public Property Loc() As String
            Get
                Loc = m_Loc
            End Get
            Set(ByVal value As String)
                m_Loc = value
            End Set
        End Property
        Public Property Qty_Per_Par() As Double
            Get
                Qty_Per_Par = m_Qty_Per_Par
            End Get
            Set(ByVal value As Double)
                m_Qty_Per_Par = value
            End Set
        End Property
        Public Property Seq_No() As Integer
            Get
                Seq_No = m_Seq_No
            End Get
            Set(ByVal value As Integer)
                m_Seq_No = value
            End Set
        End Property
        Public Property Stocked_Fg() As String
            Get
                Stocked_Fg = m_Stocked_Fg
            End Get
            Set(ByVal value As String)
                m_Stocked_Fg = value
            End Set
        End Property

#End Region

    End Class

    Public Sub New()

        Call Init()

    End Sub

    Public Sub New(ByVal pSource As OEI_SourceEnum)

        Call Init()
        m_Source = pSource

    End Sub

    Public Sub New(ByVal pstrOrd_Guid As String)

        Call Init()
        m_Ord_Guid = pstrOrd_Guid

    End Sub

    Public Sub New(ByVal pstrOrd_Guid As String, ByVal pSource As OEI_SourceEnum)

        Call Init()

        m_Ord_Guid = pstrOrd_Guid
        m_Source = pSource

    End Sub

    'Public Sub New(ByVal pstrItem_No As String, ByVal pstrLoc As String, ByVal pstrCus_No As String, ByVal pstrCurr_Cd As String)

    '    Call Init()

    '    m_Item_No = pstrItem_No
    '    m_Loc = pstrLoc
    '    m_Cus_No = pstrCus_No
    '    m_Curr_Cd = pstrCurr_Cd

    '    Call LoadSave()

    'End Sub

    Public Sub New(ByRef poParent_Ordlin As cOrdLine)

        Call Init()

        m_Ord_Guid = poParent_Ordlin.Ord_Guid
        m_Item_Guid = poParent_Ordlin.Item_Guid
        m_Source = poParent_Ordlin.Source

        If Not poParent_Ordlin.Cus_No Is Nothing Then m_Cus_No = poParent_Ordlin.Cus_No
        m_Item_Desc_1 = poParent_Ordlin.Item_Desc_1
        m_Item_Desc_2 = poParent_Ordlin.Item_Desc_2
        m_Item_No = poParent_Ordlin.Item_No
        m_Item_Price = poParent_Ordlin.Unit_Price


        m_Loc = poParent_Ordlin.Loc
        m_Line_No = poParent_Ordlin.Line_No
        m_OEOrdLin_ID = poParent_Ordlin.OEOrdLin_ID
        m_Unit_Cost = poParent_Ordlin.Unit_Cost


        m_Item_Cd = poParent_Ordlin.Item_Cd
        m_Extra_11 = poParent_Ordlin.Extra_11

        If Not poParent_Ordlin.Stocked_Fg Is Nothing Then m_Stocked_Fg = poParent_Ordlin.Stocked_Fg
        If Not poParent_Ordlin.Controlled_Fg Is Nothing Then m_Controlled_Fg = poParent_Ordlin.Controlled_Fg
        If Not poParent_Ordlin.Mfg_Method Is Nothing Then m_Mfg_Method = poParent_Ordlin.Mfg_Method
        'm_Line_No = m_oOrder.WorkLine ' poParent_Ordlin.Line_No

        'm_OrdLin = poParent_Ordlin

        If Not (poParent_Ordlin.Kit_Component) Then
            Call LoadSave(poParent_Ordlin)
        End If

        'If pblnSave Then
        '    If Not (poParent_Ordlin.Kit_Component) Then
        '        Call LoadSave()
        '    End If
        'Else
        '    Call Load()
        'End If

    End Sub

    Public Sub New(ByVal pstrOrd_Guid As String, ByVal pstrItem_Guid As String, ByVal poParent_Ordlin As cOrdLine, ByVal pblnSave As Boolean)

        Call Init()

        m_Ord_Guid = pstrOrd_Guid
        m_Item_Guid = pstrItem_Guid
        m_Source = poParent_Ordlin.Source

        If Not poParent_Ordlin.Cus_No Is Nothing Then m_Cus_No = poParent_Ordlin.Cus_No
        m_Item_Desc_1 = poParent_Ordlin.Item_Desc_1
        m_Item_Desc_2 = poParent_Ordlin.Item_Desc_2
        m_Item_No = poParent_Ordlin.Item_No
        m_Item_Price = poParent_Ordlin.Unit_Price
        m_Loc = poParent_Ordlin.Loc
        m_Line_No = poParent_Ordlin.Line_No
        m_OEOrdLin_ID = poParent_Ordlin.OEOrdLin_ID
        m_Unit_Cost = poParent_Ordlin.Unit_Cost
        m_Item_Cd = poParent_Ordlin.Item_Cd
        m_Extra_11 = poParent_Ordlin.Extra_11

        If Not poParent_Ordlin.Stocked_Fg Is Nothing Then m_Stocked_Fg = poParent_Ordlin.Stocked_Fg
        If Not poParent_Ordlin.Controlled_Fg Is Nothing Then m_Controlled_Fg = poParent_Ordlin.Controlled_Fg
        If Not poParent_Ordlin.Mfg_Method Is Nothing Then m_Mfg_Method = poParent_Ordlin.Mfg_Method
        'm_Line_No = m_oOrder.WorkLine ' poParent_Ordlin.Line_No

        'm_OrdLin = poParent_Ordlin

        If Not (poParent_Ordlin.Kit_Component) Then
            Call LoadSave(poParent_Ordlin)
        End If

        'If pblnSave Then
        '    If Not (poParent_Ordlin.Kit_Component) Then
        '        Call LoadSave()
        '    End If
        'Else
        '    Call Load()
        'End If

    End Sub

    ' Initializes every field to blank
    Public Sub Init()

        m_Source = m_oOrder.Source

        m_Item_No = String.Empty
        m_Cus_No = String.Empty
        m_Curr_Cd = String.Empty
        m_Loc = String.Empty
        m_Item_Desc_1 = String.Empty
        m_Item_Desc_2 = String.Empty
        m_Item_Price = 0

        Call Reset()

    End Sub

    Public Sub SetIsAKit(ByVal m_item_no As String, ByVal m_loc As String)

        Dim strSql As String

        Dim dtKit As New DataTable
        Dim db As New cDBA

        Try

            strSql = "" &
            "SELECT 		I.Item_No, I.Item_Desc_1, I.Item_Desc_2, " &
            "			DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) AS Item_Price, K.Comp_Item_No, " &
            "			DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', KI.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) AS Comp_Item_Price, " &
            "			KI.Item_Desc_1 AS Comp_Item_Desc_1, KI.Item_Desc_2 AS Comp_Item_Desc_2, K.Qty_Per_Par AS Comp_Qty_Per_Par, " &
            "			ISNULL(KINV.Qty_On_Hand, 0) - ISNULL(KINV.Qty_Allocated, 0) AS Comp_Qty_Available, " &
            "           K.Seq_No, K.Qty_Per_Par, ISNULL(KINV.Last_Cost, 0) AS Last_Cost, I.Stocked_Fg, I.Controlled_Fg, " &
            "           I.Mfg_Method " '&

            '"FROM 		IMITMIDX_SQL I WITH (Nolock) " &
            '"INNER JOIN IMINVLOC_SQL IINV WITH (Nolock) ON I.Item_No = IINV.Item_No " &
            '"INNER JOIN IMKITFIL_SQL K WITH (Nolock) ON I.Item_No = K.Item_No " &
            '"INNER JOIN IMITMIDX_SQL KI WITH (Nolock) ON K.Comp_item_No = KI.Item_No " &
            '"INNER JOIN IMINVLOC_SQL KINV WITH (Nolock) ON KI.Item_No = KINV.Item_No " &
            '"WHERE		I.Item_No = '" & m_item_no & "' AND IINV.Loc = '" & m_loc & "' AND KINV.Loc = '" & m_loc & "' "

            '++ID 02.20.2023
            If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then
                strSql &= "FROM IMITMIDX_SQL I WITH (Nolock) " &
            "INNER JOIN IMINVLOC_SQL IINV WITH (Nolock) ON I.Item_No = IINV.Item_No " &
            "INNER JOIN IMKITFIL_SQL K WITH (Nolock) ON I.Item_No = K.Item_No " &
            "INNER JOIN IMITMIDX_SQL KI WITH (Nolock) ON K.Comp_item_No = KI.Item_No " &
            "INNER JOIN IMINVLOC_SQL KINV WITH (Nolock) ON KI.Item_No = KINV.Item_No " &
            "WHERE		I.Item_No = '" & m_item_no & "' AND IINV.Loc = '" & m_loc & "' AND KINV.Loc = '" & m_loc & "' "

            Else
                strSql &= " FROM  [200].dbo.IMITMIDX_SQL I WITH (Nolock) " &
            "INNER JOIN [200].dbo.IMINVLOC_SQL IINV WITH (Nolock) ON I.Item_No = IINV.Item_No " &
            "INNER JOIN [200].dbo.IMKITFIL_SQL K WITH (Nolock) ON I.Item_No = K.Item_No " &
            "INNER JOIN [200].dbo.IMITMIDX_SQL KI WITH (Nolock) ON K.Comp_item_No = KI.Item_No " &
            "INNER JOIN [200].dbo.IMINVLOC_SQL KINV WITH (Nolock) ON KI.Item_No = KINV.Item_No " &
            "WHERE		I.Item_No = '" & m_item_no & "' AND IINV.Loc = '" & m_loc & "' AND KINV.Loc = '" & m_loc & "' "
            End If


            dtKit = New DataTable
            dtKit = db.DataTable(strSql)

            ' If it is not a kit, we leave now.
            If dtKit.Rows.Count = 0 Then Exit Sub

            m_IsAKit = True

        Catch er As Exception
            MsgBox("Error in CKit." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub load()

        m_Sql = "" & _
        "SELECT 	I.Item_No, I.Item_Desc_1, I.Item_Desc_2, " & _
        "			DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) AS Item_Price, K.Comp_Item_No, " & _
        "			DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', KI.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) AS Comp_Item_Price, " & _
        "			KI.Item_Desc_1 AS Comp_Item_Desc_1, KI.Item_Desc_2 AS Comp_Item_Desc_2, K.Qty_Per_Par AS Comp_Qty_Per_Par, " & _
        "			ISNULL(KINV.Qty_On_Hand, 0) - ISNULL(KINV.Qty_Allocated, 0) AS Comp_Qty_Available " & _
        "FROM 		IMITMIDX_SQL I WITH (Nolock) " & _
        "INNER JOIN IMINVLOC_SQL IINV WITH (Nolock) ON I.Item_No = IINV.Item_No " & _
        "INNER JOIN IMKITFIL_SQL K WITH (Nolock) ON I.Item_No = K.Item_No " & _
        "INNER JOIN IMITMIDX_SQL KI WITH (Nolock) ON K.Comp_item_No = KI.Item_No " & _
        "INNER JOIN IMINVLOC_SQL KINV WITH (Nolock) ON KI.Item_No = KINV.Item_No " & _
        "WHERE		I.Item_No = '" & m_Item_No & "' AND IINV.Loc = '" & m_Loc & "' AND KINV.Loc = '" & m_Loc & "' "

    End Sub

    Private Sub LoadSave(ByRef poOrdline As cOrdLine)

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Dim strSql As String

        Dim dt As New DataTable
        Dim dtKit As New DataTable
        Dim dtItem As New DataTable

        Dim db As New cDBA
        Dim oOrdline As cOrdLine

        Dim dblInventory As Double = 9999999
        Dim dblQty_Allocated As Double = 9999999
        Dim dblQty_On_Hand As Double = 9999999
        Dim strDefPack As String

        Try

            m_Sql = "" &
            "SELECT 	I.Item_No, I.Item_Desc_1, I.Item_Desc_2, I.User_Def_Fld_1 AS Item_Cd, " &
            "			DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', I.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) AS Item_Price, K.Comp_Item_No, " &
            "			DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', KI.Item_No, '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1) AS Comp_Item_Price, " &
            "			KI.Item_Desc_1 AS Comp_Item_Desc_1, KI.Item_Desc_2 AS Comp_Item_Desc_2, K.Qty_Per_Par AS Comp_Qty_Per_Par, " &
            "			ISNULL(KINV.Qty_On_Hand, 0) - ISNULL(KINV.Qty_Allocated, 0) AS Comp_Qty_Available " &
            "FROM 		IMITMIDX_SQL I WITH (Nolock) " &
            "INNER JOIN IMINVLOC_SQL IINV WITH (Nolock) ON I.Item_No = IINV.Item_No " &
            "INNER JOIN IMKITFIL_SQL K WITH (Nolock) ON I.Item_No = K.Item_No " &
            "INNER JOIN IMITMIDX_SQL KI WITH (Nolock) ON K.Comp_item_No = KI.Item_No " &
            "INNER JOIN IMINVLOC_SQL KINV WITH (Nolock) ON KI.Item_No = KINV.Item_No " &
            "WHERE		I.Item_No = '" & m_Item_No & "' AND IINV.Loc = '" & m_Loc & "' AND KINV.Loc = '" & m_Loc & "' "

            ' First, we get a recordset with kit items. We check if it is a kit. 
            ' we get every item from ItmKit link to item to get kit data, and we create a line for each.
            ' "SELECT 	K.Comp_Item_no, K.Seq_No, K.Qty_Per_Par, I.* "
            strSql = "" &
            "SELECT 	K.Comp_Item_no, K.Seq_No, K.Qty_Per_Par, ISNULL(L.Last_Cost, 0) AS Last_Cost, " &
            "           I.Stocked_Fg, I.Controlled_Fg, I.Mfg_Method, KI.User_Def_Fld_1 AS Kit_Item_Cd, I.User_Def_Fld_1 AS Comp_Item_Cd " &
            "FROM 		ImKitFil_Sql K WITH (Nolock) " &
            "INNER JOIN ImItmIdx_Sql I WITH (Nolock) ON K.Comp_Item_No = I.Item_No " &
            "INNER JOIN ImItmIdx_Sql KI WITH (Nolock) ON K.Item_No = KI.Item_No " &
            "LEFT JOIN  ImInvLoc_SQL L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & m_Loc & "' " &
            "WHERE 		K.Item_No = '" & m_Item_No & "' "

            ' First, we get a recordset with kit items. We check if it is a kit. 
            ' we get every item from ItmKit link to item to get kit data, and we create a line for each.
            ' We join twice with imitmidx_sql twice. Once for kit component and the other for the kit istelf
            ' to get the packaging for the skit style, not by component
            ' "SELECT 	K.Comp_Item_no, K.Seq_No, K.Qty_Per_Par, I.* "
            strSql = "" &
            "SELECT 	K.Comp_Item_no, K.Seq_No, K.Qty_Per_Par, ISNULL(L.Last_Cost, 0) AS Last_Cost, " &
            "           I.Stocked_Fg, I.Controlled_Fg, I.Mfg_Method, KI.User_Def_Fld_1 AS Kit_Item_Cd, " &
            "           I.User_Def_Fld_1 AS Comp_Item_Cd, ISNULL(PK.Pack_CD, '') as Packaging " &
            "FROM 		ImKitFil_Sql K WITH (Nolock) " &
            "INNER JOIN ImItmIdx_Sql I WITH (Nolock) ON K.Comp_Item_No = I.Item_No " &
            "INNER JOIN ImItmIdx_Sql KI WITH (Nolock) ON K.Item_No = KI.Item_No " &
            "LEFT JOIN  ImInvLoc_SQL L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & m_Loc & "' " &
            "LEFT JOIN ImItmIdx_Sql I2 WITH (Nolock) ON K.item_no = I2.Item_No " &
            "LEFT JOIN  MDB_ITEM_MASTER MD WITH (Nolock) ON I2.user_def_fld_1 = MD.ITEM_CD " &
            "LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MD.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
            "LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
            "WHERE 		K.Item_No = '" & m_Item_No & "' "


            If m_oOrder.Ordhead.Mfg_Loc = "US1" Then

                strSql = "" &
            "SELECT 	K.Comp_Item_no, K.Seq_No, K.Qty_Per_Par, ISNULL(L.Last_Cost, 0) AS Last_Cost, " &
            "           I.Stocked_Fg, I.Controlled_Fg, I.Mfg_Method, KI.User_Def_Fld_1 AS Kit_Item_Cd, " &
            "           I.User_Def_Fld_1 AS Comp_Item_Cd, ISNULL(PK.Pack_CD, '') as Packaging " &
            "FROM 		[200].dbo.ImKitFil_Sql K WITH (Nolock) " &
            "INNER JOIN [200].dbo.ImItmIdx_Sql I WITH (Nolock) ON K.Comp_Item_No = I.Item_No " &
            "INNER JOIN [200].dbo.ImItmIdx_Sql KI WITH (Nolock) ON K.Item_No = KI.Item_No " &
            "LEFT JOIN  [200].dbo.ImInvLoc_SQL L WITH (Nolock) ON I.Item_No = L.Item_No AND L.Loc = '" & m_Loc & "' " &
            "LEFT JOIN [200].dbo.ImItmIdx_Sql I2 WITH (Nolock) ON K.item_no = I2.Item_No " &
            "LEFT JOIN  MDB_ITEM_MASTER MD WITH (Nolock) ON I2.user_def_fld_1 = MD.ITEM_CD " &
            "LEFT JOIN  MDB_ITEM_PKG IP WITH (Nolock) ON MD.ITEM_MASTER_ID = IP.ITEM_MASTER_ID AND IP.Country_CD = '" & m_oOrder.Ordhead.Ship_To_Country & "' AND IP.Bound_Type = 'out' AND IP.ITEM_MASTER_ID IS NOT NULL AND IP.STANDARD = 1 " &
            "LEFT JOIN  MDB_CFG_PACK PK WITH (Nolock) ON IP.PACK_ID = PK.PACK_ID " &
            "WHERE 		K.Item_No = '" & m_Item_No & "' "


            End If








            'TODO: stop the insanity of duplicate querying over and over again on the same item row

            dtKit = New DataTable
            dtKit = db.DataTable(strSql)

            ' If it is not a kit, we leave now.
            If dtKit.Rows.Count = 0 Then Exit Sub

            m_IsAKit = True

            Item_Cd = dtKit.Rows(0).Item("Kit_Item_Cd")
            strDefPack = dtKit.Rows(0).Item("Packaging")

            ' Second, we check if an entry already exists.
            strSql = "" &
            "SELECT 	* " &
            "FROM 		OEI_ORDBLD WITH (Nolock) " &
            "WHERE 		Kit_Item_Guid = '" & m_Item_Guid & "' "

            dt = db.DataTable(strSql)

            ' If it exists, we leave now.
            If dt.Rows.Count <> 0 Then Exit Sub

            ' Third, we insert the current item into OEI_ORDBLD so we can retrieve it with the INNER JOIN to this table.
            strSql = "" &
            "INSERT INTO OEI_ORDBLD " &
            "   ( " &
            "       ORD_TYPE, ORD_GUID, Kit_Item_Guid, Comp_Item_Guid, CTL_NO, LINE_NO, SEQ_NO, PAR_ITEM_NO, " &
            "       SEQ_FTR_NO,ITEM_NO, TYPE, LOC, UNIT_PRICE, UNIT_COST, LVL_NO, STOCKED_FG, CONTROLLED_FG, " &
            "       PUR_OR_MFG, QTY, QTY_PER, ACTIVITY_FG, EFFECTIVITY_DT, OBSOLETE_DT, QTY_BO, " &
            "       PICK_SEQ, LINE_SEQ, QTY_TO_SHIP, PAR_FG, DRAWING_REL_NO, DRAWING_REV_NO, " &
            "       LAST_ITEM_REV, ATTACH_OPER_NO, SCRP_FCTR, MFG_UOM, BACKFLUSH_FG, BULK_ISSUE_FG, " &
            "       SCRP_QTY, NEG_PER_FG, MPR_PHANTOM_FG, PO_NO, PO_LINE_NO, VENDOR_NO, DIST_PRICE, " &
            "       EXTRA_1, EXTRA_2, EXTRA_3, EXTRA_4, EXTRA_5, EXTRA_6, EXTRA_7, EXTRA_8, EXTRA_9, " &
            "       EXTRA_10, EXTRA_11, EXTRA_12, EXTRA_13, EXTRA_14, EXTRA_15, ISSUED_QTY, FILLER_0001, Kit_Item_Cd, Comp_Item_Cd " &
            "   ) VALUES " &
            "   ( " &
            "        '" & Mid(m_oOrder.Ordhead.Ord_Type, 1, 1) & "', '" & m_oOrder.Ordhead.Ord_GUID & "', '" & m_Item_Guid & "', '" & m_Item_Guid & "', " &
            "        '        ', " & m_Line_No & ", 0, '" & m_Item_No & "', 0, '" & m_Item_No & "', NULL, '" & m_Loc & "', " &
            "        DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', '" & m_Item_No & "', '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1), " &
            "        " & m_Unit_Cost & ", 0, '" & m_Stocked_Fg & "', '" & m_Controlled_Fg & "', " &
            "        '" & m_Mfg_Method & "', 0, 1, NULL, NULL, NULL, 0, '        ', 0, 0, " &
            "        NULL, NULL, NULL,NULL, 0, 0, NULL, NULL, NULL, 0, NULL, NULL, NULL, 0, NULL, 0, " &
            "        NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, NULL, '" & Item_Cd & "', '" & Item_Cd & "' " &
            "   )    "

            db.Execute(strSql)

            Dim lPos As Integer = 1 ' Unit_Cost, Stocked_Fg, Controlled_Fg, Mfg_Method
            'g_oOrdline.Kit.Items.Count

            g_oKitCollection = New Collection

            For Each drRow As DataRow In dtKit.Rows

                oOrdline = New cOrdLine(m_Ord_Guid, poOrdline.Source)
                'oOrdline.Source = m_Source

                oOrdline.Kit_Component = True

                ' Will prevent infinite loops
                oOrdline.Kit.IsAKit = False

                oOrdline.Item_No = drRow.Item("Comp_Item_no").ToString
                oOrdline.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                oOrdline.Kit_Item_Guid = m_Item_Guid
                oOrdline.Loc = m_Loc
                oOrdline.Line_No = m_Line_No
                oOrdline.OEOrdLin_ID = m_OEOrdLin_ID
                oOrdline.Extra_11 = m_Extra_11

                oOrdline.ImprintLine = New cImprint
                oOrdline.ImprintLine.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                oOrdline.ImprintLine.Item_No = oOrdline.Item_No
                oOrdline.ImprintLine.Item_Guid = oOrdline.Item_Guid
                oOrdline.ImprintLine.Packaging = strDefPack

                oOrdline.Set_Spec_Instructions(oOrdline.Item_No, oOrdline.ImprintLine)
                If oOrdline.Source = OEI_SourceEnum.frmPLEntrySelectedItem Then oOrdline.ImprintLine.Save()

                oOrdline.Save()

                strSql = "" &
                "INSERT INTO OEI_ORDBLD " &
                "   ( " &
                "       ORD_TYPE, ORD_GUID, Kit_Item_Guid, Comp_Item_Guid, CTL_NO, LINE_NO, SEQ_NO, PAR_ITEM_NO, SEQ_FTR_NO, " &
                "       ITEM_NO, TYPE, LOC, UNIT_PRICE, UNIT_COST, LVL_NO, STOCKED_FG, CONTROLLED_FG, " &
                "       PUR_OR_MFG, QTY, QTY_PER, ACTIVITY_FG, EFFECTIVITY_DT, OBSOLETE_DT, QTY_BO, " &
                "       PICK_SEQ, LINE_SEQ, QTY_TO_SHIP, PAR_FG, DRAWING_REL_NO, DRAWING_REV_NO, " &
                "       LAST_ITEM_REV, ATTACH_OPER_NO, SCRP_FCTR, MFG_UOM, BACKFLUSH_FG, BULK_ISSUE_FG, " &
                "       SCRP_QTY, NEG_PER_FG, MPR_PHANTOM_FG, PO_NO, PO_LINE_NO, VENDOR_NO, DIST_PRICE, " &
                "       EXTRA_1, EXTRA_2, EXTRA_3, EXTRA_4, EXTRA_5, EXTRA_6, EXTRA_7, EXTRA_8, EXTRA_9, " &
                "       EXTRA_10, EXTRA_11, EXTRA_12, EXTRA_13, EXTRA_14, EXTRA_15, ISSUED_QTY, FILLER_0001, Kit_Item_Cd, Comp_Item_Cd  " &
                "   ) VALUES " &
                "   ( " &
                "        '" & Mid(m_oOrder.Ordhead.Ord_Type, 1, 1) & "', '" & m_oOrder.Ordhead.Ord_GUID & "', " &
                "        '" & m_Item_Guid & "', '" & oOrdline.Item_Guid & "', '        ', " & m_Line_No & ", " &
                "        " & lPos & ", '" & m_Item_No & "', 0, '" & oOrdline.Item_No & "', NULL, '" & m_Loc & "', " &
                "        DBO.OEI_Item_Price_20140101('" & m_oOrder.Ordhead.Curr_Cd & "', '" & m_oOrder.Ordhead.Cus_No & "', '" & oOrdline.Item_No & "', '', '', " & m_oOrder.Ordhead.Cus_Prog_ID & ", 1), " &
                "        " & drRow.Item("Last_Cost") & ", 0, '" & drRow.Item("Stocked_Fg").ToString & "', '" & drRow.Item("Controlled_Fg").ToString & "', " &
                "        'P', 0, " & drRow.Item("Qty_Per_Par").ToString & ", NULL, NULL, NULL, 0, '        ', 0, 0, " &
                "        NULL, NULL, NULL, NULL, 0, 0, NULL, NULL, NULL, 0, NULL, NULL, NULL, 0, NULL, 0, " &
                "        NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, NULL, '" & Item_Cd & "', '" & Mid(drRow.Item("Comp_Item_Cd").ToString, 1, 10) & "' " &
                "   ) "
                '        '" & drRow.Item("Mfg_Method").ToString & "', 0, " & drRow.Item("Qty_Per_Par").ToString & ", NULL, NULL, NULL, 0, '        ', 0, 0, " & _

                db.Execute(strSql)

                lPos += 1

                g_oKitCollection.Add(oOrdline, oOrdline.Item_Guid)

            Next drRow

            dt.Dispose()

        Catch er As Exception
            MsgBox("Error in CKit." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every field to blank. For a complete reset (as in New), use Init.
    Public Sub Reset()

        m_Sql = String.Empty
        m_IsAKit = False

        Items = New Collection

    End Sub








#Region "Public properties ################################################"

    Public Property Curr_Cd() As String
        Get
            Curr_Cd = m_Curr_Cd
        End Get
        Set(ByVal value As String)
            m_Curr_Cd = value
        End Set
    End Property
    Public Property Cus_No() As String
        Get
            Cus_No = m_Cus_No
        End Get
        Set(ByVal value As String)
            m_Cus_No = value
        End Set
    End Property
    Public Property IsAKit() As Boolean
        Get
            IsAKit = m_IsAKit
        End Get
        Set(ByVal value As Boolean)
            m_IsAKit = value
        End Set
    End Property
    Public Property Item_Desc_1() As String
        Get
            Item_Desc_1 = m_Item_Desc_1
        End Get
        Set(ByVal value As String)
            m_Item_Desc_1 = value
        End Set
    End Property
    Public Property Item_Desc_2() As String
        Get
            Item_Desc_2 = m_Item_Desc_2
        End Get
        Set(ByVal value As String)
            m_Item_Desc_2 = value
        End Set
    End Property
    Public Property Item_Guid() As String
        Get
            Item_Guid = m_Item_Guid
        End Get
        Set(ByVal value As String)
            m_Item_Guid = value
        End Set
    End Property
    Public Property Item_Cd() As String
        Get
            Item_Cd = Mid(m_Item_Cd, 1, 10)
        End Get
        Set(ByVal value As String)
            m_Item_Cd = Mid(value, 1, 10)
        End Set
    End Property
    Public Property Item_No() As String
        Get
            Item_No = m_Item_No
        End Get
        Set(ByVal value As String)
            m_Item_No = value
        End Set
    End Property
    Public Property Item_Price() As Double
        Get
            Item_Price = m_Item_Price
        End Get
        Set(ByVal value As Double)
            m_Item_Price = value
        End Set
    End Property
    Public Property Loc() As String
        Get
            Loc = m_Loc
        End Get
        Set(ByVal value As String)
            m_Loc = value
        End Set
    End Property
    Public Property Sql() As String
        Get
            Sql = m_Sql
        End Get
        Set(ByVal value As String)
            m_Sql = value
        End Set
    End Property

#End Region

End Class
