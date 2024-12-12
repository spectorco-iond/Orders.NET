Public Class frmShipLinks

    Private m_OE_PO_No As String
    Private m_Ord_No As String
    Private m_Cus_No As String

#Region "Constructors #####################################################"

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Call LoadShipLinkView()

    End Sub

    Public Sub New(ByVal pstrOE_PO_No As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        OE_PO_No = pstrOE_PO_No

        Call LoadShipLinkView()

    End Sub

    Public Sub New(ByVal pstrCus_No As String, ByVal pstrOrd_No As String, ByVal pstrOE_PO_No As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        OE_PO_No = pstrOE_PO_No
        Ord_No = pstrOrd_No
        Cus_No = pstrCus_No

        Call LoadShipLinkView()

    End Sub

#End Region

    Private Sub LoadShipLinkView()

        Dim strSql As String

        'SELECT      H.Cus_No, H.Ord_No, H.OE_PO_No, L.OrdNo, L.LinkGroupID, LS.OrdNo, LS.ID
        'FROM        OEORDHDR_SQL H WITH (Nolock) 
        'INNER JOIN  EXACT_TRAVELER_LINKED_SHIPMENT L WITH (NOLOCK) ON H.Ord_No = L.OrdNo
        'INNER JOIN  EXACT_TRAVELER_LINKED_SHIPMENT LS WITH (NOLOCK) ON L.LinkGroupID = LS.LinkGroupID
        'WHERE       H.OE_PO_No = 'L0000175205' -- 402834 403529
        'ORDER BY    L.LinkGroupID, H.Ord_No, LS.OrdNo

        ' from OEI_OrdHdr, we only take orders without Ord_No as when they are transferred, 
        ' they also are put in EXACT_TRAVELER_LINKED_SHIPMENT
        strSql = _
        "SELECT     LS.OrdNo AS 'Linked Orders' " & _
        "FROM       OEORDHDR_SQL H WITH (Nolock) " & _
        "INNER JOIN EXACT_TRAVELER_LINKED_SHIPMENT L WITH (NOLOCK) ON H.Ord_No = L.OrdNo " & _
        "INNER JOIN EXACT_TRAVELER_LINKED_SHIPMENT LS WITH (NOLOCK) ON L.LinkGroupID = LS.LinkGroupID " & _
        "WHERE      H.Cus_No = '" & m_Cus_No & "' AND H.OE_PO_No = '" & m_OE_PO_No & "'  " & _
        "UNION " & _
        "SELECT		H.OEI_Ord_No AS 'Linked Orders' " & _
        "FROM		OEI_OrdHdr H WITH (Nolock) " & _
        "WHERE		H.Cus_No = '" & m_Cus_No & "' AND H.Ship_Link = '" & m_OE_PO_No & "' AND H.Ord_No IS NULL " & _
        "ORDER BY 	'Linked Orders' "

        'SELECT 		LS.OrdNo AS 'Linked Orders' 
        'FROM        OEORDHDR_SQL H WITH (Nolock) 
        'INNER JOIN  EXACT_TRAVELER_LINKED_SHIPMENT L WITH (NOLOCK) ON H.Ord_No = L.OrdNo 
        'INNER JOIN  EXACT_TRAVELER_LINKED_SHIPMENT LS WITH (NOLOCK) ON L.LinkGroupID = LS.LinkGroupID 
        'WHERE       H.OE_PO_No = '3736'  -- ORDER BY   L.LinkGroupID, H.Ord_No, LS.OrdNo 
        'UNION
        'SELECT		H.OEI_Ord_No AS 'Linked Orders' 
        'FROM		OEI_OrdHdr H WITH (Nolock)
        'WHERE		LTRIM(RTRIM(H.Ship_Link)) = '3736'
        'ORDER BY 	'Linked Orders' 

        Try

            Dim conn As New cDBA()
            Dim dt As New DataTable

            dt = conn.DataTable(strSql)

            dgvShipLinks.DataSource = dt

            dgvShipLinks.RowHeadersWidth = 16
            dgvShipLinks.Columns(0).Width = dgvShipLinks.Width - dgvShipLinks.RowHeadersWidth - 2

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

#Region "Public properties ################################################"

    Public Property Cus_No() As String
        Get
            Cus_No = m_Cus_No
        End Get
        Set(ByVal value As String)
            m_Cus_No = value
        End Set
    End Property
    Public Property OE_PO_No() As String
        Get
            OE_PO_No = m_OE_PO_No
        End Get
        Set(ByVal value As String)
            m_OE_PO_No = value
            txtShip_Link_No.Text = m_OE_PO_No
        End Set
    End Property
    Public Property Ord_No() As String
        Get
            Ord_No = m_Ord_No
        End Get
        Set(ByVal value As String)
            m_Ord_No = value
        End Set
    End Property

#End Region

    Private Sub dgvShipLinks_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvShipLinks.CellBeginEdit

        e.Cancel = True

    End Sub

    Private Sub txtShip_Link_No_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShip_Link_No.GotFocus

        dgvShipLinks.Focus()

    End Sub

    Private Sub cmdAddShip_Link_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddShip_Link.Click

    End Sub
End Class