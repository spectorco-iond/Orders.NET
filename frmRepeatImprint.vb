Public Class frmRepeatImprint

    Private m_Order_No As String = ""
    Private m_Imprint As New cImprint

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

    End Sub

    Public Sub New(ByVal pstrOrder_No As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        m_Order_No = Trim(pstrOrder_No).PadLeft(8)

    End Sub

    Private Sub frmRepeatImprint_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Try

            Call SetExtraFieldsGridColumns()
            Call CreateExtraFieldsRecordset()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdCopyImprintData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopyImprintData.Click

        Try

            m_Imprint = New cImprint

            If dgvExtraFields.Rows.Count <> 0 Then

                With dgvExtraFields.CurrentRow

                    m_Imprint.SaveToDB = False

                    m_Imprint.Item_No = .Cells("Item_No").Value.ToString
                    m_Imprint.Imprint = .Cells("Extra_1").Value.ToString
                    m_Imprint.End_User = .Cells("End_User").Value.ToString
                    m_Imprint.Location = .Cells("Extra_2").Value.ToString
                    m_Imprint.Color = .Cells("Extra_3").Value.ToString
                    If IsNumeric(.Cells("Num_Impr_1").Value.ToString) Then
                        m_Imprint.Num_Impr_1 = .Cells("Num_Impr_1").Value
                    Else
                        m_Imprint.Num_Impr_1 = 0
                    End If
                    If IsNumeric(.Cells("Num_Impr_2").Value.ToString) Then
                        m_Imprint.Num_Impr_2 = .Cells("Num_Impr_2").Value.ToString
                    Else
                        m_Imprint.Num_Impr_2 = 0
                    End If
                    If IsNumeric(.Cells("Num_Impr_3").Value.ToString) Then
                        m_Imprint.Num_Impr_3 = .Cells("Num_Impr_3").Value.ToString
                    Else
                        m_Imprint.Num_Impr_3 = 0
                    End If
                    m_Imprint.Packaging = .Cells("Extra_5").Value.ToString
                    m_Imprint.Refill = .Cells("Extra_8").Value.ToString
                    m_Imprint.Laser_Setup = .Cells("Extra_9").Value.ToString
                    m_Imprint.Comments = .Cells("Comment").Value.ToString
                    m_Imprint.Special_Comments = .Cells("Comment2").Value.ToString
                    m_Imprint.Industry = .Cells("Industry").Value.ToString
                    m_Imprint.Printer_Comment = .Cells("Printer_Comment").Value.ToString
                    m_Imprint.Packer_Comment = .Cells("Packer_Comment").Value.ToString
                    m_Imprint.Printer_Instructions = .Cells("Printer_Instructions").Value.ToString
                    m_Imprint.Packer_Instructions = .Cells("Packer_Instructions").Value.ToString
                    m_Imprint.Repeat_From_ID = .Cells("Repeat_From_ID").Value

                    g_oOrdline.Set_Spec_Instructions(m_Imprint.Item_No, m_Imprint)

                    m_Imprint.SaveToDB = True

                    If Not m_Imprint.IsEmpty() Then m_Imprint.Save()

                    Me.Close()

                End With

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

    Private Sub CreateExtraFieldsRecordset()

        Dim dt As DataTable
        Dim db As New cDBA

        Try

            Dim strSql As String

            strSql = _
            "SELECT EX.Item_no, EX.Par_Item_no, EX.Extra_1, EX.End_User, EX.Extra_2, " & _
            "       EX.Extra_3, CASE WHEN ISNULL(Num_Impr_1, 0) = 0 THEN CONVERT(INT, ISNULL(Extra_4, '0')) ELSE ISNULL(Num_Impr_1, 0) END AS Num_Impr_1, " & _
            "       ISNULL(Num_Impr_2, 0) AS Num_Impr_2, ISNULL(Num_Impr_3, 0) AS Num_Impr_3, EX.Extra_5, " & _
            "       CASE WHEN OEL.qty_ordered IS NULL THEN 'Completed' ELSE CAST(CAST(OEL.qty_ordered AS INT) AS CHAR(20)) END AS QtyOnOrder, " & _
            "       EX.Extra_8, EX.Extra_9, EX.Industry, EX.Comment, EX.Comment2, " & _
            "       EX.Printer_Comment, EX.Packer_Comment, EX.Printer_Instructions, EX.Packer_Instructions, EX.ID AS Repeat_From_ID " & _
            "FROM   Exact_Traveler_Extra_Fields EX " & _
            "LEFT JOIN " & _
            "   (" & _
            "       SELECT  qty_ordered, ord_no, line_no " & _
            "       FROM    oeordlin_sql " & _
            "       WHERE   ord_type in ('O', 'I')" & _
            "   ) OEL ON EX.ord_no = OEL.ord_no AND EX.OE_Line_no = OEL.Line_no " & _
            "WHERE  EX.ord_no = '" & m_Order_No & "' " & _
            "ORDER BY EX.ord_no, EX.OE_Line_no, EX.Comp_seq_no "

            'strSql = _
            '"SELECT '' AS FirstCol, EX.ID, LTRIM(RTRIM(EX.Ord_no)), EX.OE_Line_no, EX.Comp_seq_no, " & _
            '"       EX.Par_Item_no, " & " EX.Item_no, EX.Extra_1, EX.Extra_2, " & _
            '"       EX.Extra_3, EX.Extra_4, EX.Extra_5, EX.Extra_6, EX.Extra_7, " & _
            '"       CASE WHEN OEL.qty_ordered IS NULL THEN 'Completed' ELSE CAST(CAST(OEL.qty_ordered AS INT) AS CHAR(20)) END AS QtyOnOrder, " & _
            '"       EX.Extra_8, EX.Extra_9, EX.Industry, EX.Comment, EX.Comment2, '' AS LastCol " & _
            '"FROM   Exact_Traveler_Extra_Fields EX " & _
            '"LEFT JOIN (" & _
            '"       SELECT  qty_ordered, ord_no, line_no " & _
            '"       FROM    oeordlin_sql " & _
            '"       WHERE   ord_type in ('O', 'I')" & _
            '") OEL ON EX.ord_no = OEL.ord_no AND EX.OE_Line_no = OEL.Line_no " & _
            '"WHERE  ltrim(rtrim(EX.ord_no)) = '" & m_Order_No & "' " & _
            '"ORDER BY EX.ord_no, EX.OE_Line_no, EX.Comp_seq_no "

            dt = db.DataTable(strSql)

            dgvExtraFields.DataSource = dt

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SetExtraFieldsGridColumns()

        Try

            With dgvExtraFields.Columns


                '.Add(DGVTextBoxColumn("FirstCol", "FirstCol", 50))
                '.Add(DGVTextBoxColumn("ID", "ID", 50))
                '.Add(DGVTextBoxColumn("Ord_no", "Ord_no", 50))
                '.Add(DGVTextBoxColumn("OE_Line_no", "OE_Line_no", 50))

                '.Add(DGVTextBoxColumn("Comp_seq_no", "Comp seq no", 50))
                .Add(DGVTextBoxColumn("Item_no", "Item no", 100))
                .Add(DGVTextBoxColumn("Par_Item_no", "Par Item no", 100))
                .Add(DGVTextBoxColumn("Extra_1", "Imprint", 80))
                .Add(DGVTextBoxColumn("End_User", "End User", 80))
                .Add(DGVTextBoxColumn("Extra_2", "Location", 80))
                .Add(DGVTextBoxColumn("Extra_3", "Color", 80))
                .Add(DGVTextBoxColumn("Num_Impr_1", "Num Impr 1", 80))
                .Add(DGVTextBoxColumn("Num_Impr_2", "Num Impr 2", 80))
                .Add(DGVTextBoxColumn("Num_Impr_3", "Num Impr 3", 80))
                .Add(DGVTextBoxColumn("Extra_5", "Packaging", 80))
                '.Add(DGVTextBoxColumn("Extra_6", "Extra6", 80))
                '.Add(DGVTextBoxColumn("Extra_7", "Extra7", 80))
                .Add(DGVTextBoxColumn("QtyOnOrder", "QtyOnOrder", 80))
                .Add(DGVTextBoxColumn("Extra_8", "Refill", 80))
                .Add(DGVTextBoxColumn("Extra_9", "Laser Setup", 80))
                .Add(DGVTextBoxColumn("Industry", "Industry", 80))
                .Add(DGVTextBoxColumn("Comment", "Comment 1", 100))
                .Add(DGVTextBoxColumn("Comment2", "Comment 2", 100))
                .Add(DGVTextBoxColumn("Printer_Comment", "Printer Comment", 100))
                .Add(DGVTextBoxColumn("Packer_Comment", "Packer Comment", 100))
                .Add(DGVTextBoxColumn("Printer_Instructions", "Printer Instructions", 100))
                .Add(DGVTextBoxColumn("Packer_Instructions", "Packer Instructions", 100))
                .Add(DGVTextBoxColumn("Repeat_From_ID", "Repeat_From_ID", 100))

                '.Add(DGVTextBoxColumn("LastCol", "LastCol", 50))
                'Repeat_From_ID
            End With

            dgvExtraFields.Columns(Columns.Repeat_From_ID).Visible = False
            'dgvExtraFields.Columns(Columns.Extra_7).Visible = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Property Order_No() As String
        Get
            Order_No = m_Order_No
        End Get
        Set(ByVal value As String)
            m_Order_No = value
        End Set
    End Property

    Public Property Imprint() As cImprint
        Get
            Imprint = m_Imprint
        End Get
        Set(ByVal value As cImprint)
            m_Imprint = value
        End Set
    End Property

    Private Enum Columns
        Item_no
        Par_Item_no
        Extra_1
        End_User
        Extra_2
        Extra_3
        Num_Impr_1
        Nun_Impr_2
        Extra_5
        QtyOnOrder
        Extra_8
        Extra_9
        Industry
        Comment
        Comment2
        Printer_Comment
        Packer_Comment
        Printer_Instructions
        Packer_Instructions
        Repeat_From_ID

        'FirstCol = 0
        'ID
        'Ord_no
        'OE_Line_no
        'Comp_seq_no
        'Par_Item_no
        'Item_no
        'Extra_1
        'Extra_2
        'Extra_3
        'Extra_4
        'Extra_5
        'Extra_6
        'Extra_7
        'QtyOnOrder
        'Extra_8
        'Extra_9
        'Industry
        'Comment
        'Comment2
        'LastCol
    End Enum

End Class