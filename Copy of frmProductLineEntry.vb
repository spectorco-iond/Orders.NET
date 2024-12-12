Imports System.Data.SqlClient
Imports System.IO

Public Class frmProductLineEntry

    Private _Item_No As String
    Private blnInvalidate As Boolean = False
    Private iRowPos As Integer = 0
    Private iColPos As Integer = 0

    Private _ProductStatus As Integer = 0

    Public Enum _Columns
        _Picture = 0
        _Item_No
        _Description
        _Description2
        _Qty_Ordered
        _Qty_Available
        _Qty_BO
        _Route
    End Enum

    Public Enum Status
        Cancel = 0
        Insert
    End Enum

    Public Sub LoadProductGrid()

        Dim rsProducts As New ADODB.Recordset
        Dim strSql As String

        If RTrim(Trim(_Item_No)).Length < 4 Then Exit Sub

        Try
            strSql = "" & _
            "SELECT 	P.Picture AS Image, I.Item_No AS 'Item No', I.Item_Desc_1 AS 'Description', I.Item_Desc_2 AS 'Description 2', CAST(0 AS VARCHAR) AS 'Qty Ordered', CAST(1000 AS VARCHAR) AS 'Qty available', CAST(0 AS VARCHAR) AS 'Qty B.O.', '' AS 'Route' " & _
            "FROM 		IMITMIDX_SQL I WITH (nolock) " & _
            "LEFT JOIN 	Item_Pictures P WITH (Nolock) ON RTRIM(I.Item_No) = RTRIM(P.Item_No) " & _
            "WHERE 		I.Item_No LIKE '" & _Item_No & "%' " & _
            "ORDER BY   I.Item_No "

            Dim myConnection As SqlConnection = New SqlConnection("Server=thor;Initial Catalog=100;Integrated Security=SSPI")
            Dim mySelectCommand As SqlCommand = New SqlCommand(strSql, myConnection)
            Dim mySqlDataAdapter As SqlDataAdapter = New SqlDataAdapter(mySelectCommand)
            Dim myDataSet As New DataSet()
            mySqlDataAdapter.Fill(myDataSet)
            dgvOrderLines.DataSource = Nothing
            dgvOrderLines.DataSource = myDataSet.Tables(0)

            dgvOrderLines.AllowUserToAddRows = False
            dgvOrderLines.AllowUserToOrderColumns = False

            'For lPos As Integer = 1 To 5
            '    dgvOrderLines.Columns(lPos).SortMode = DataGridViewColumnSortMode.NotSortable
            'Next lPos

            Dim oImage, oResizedImage As Image
            Dim data As Byte()
            Dim ms As MemoryStream
            Dim dblRatio As Double

            For Each myRow As DataGridViewRow In dgvOrderLines.Rows

                If Not (myRow.Cells(_Columns._Picture).Value.Equals(System.DBNull.Value)) Then
                    data = myRow.Cells(_Columns._Picture).Value ' (byte[]) dt.Rows[0]["IMAGE"];
                    ms = New MemoryStream(data)
                    oImage = Image.FromStream(ms)

                    dblRatio = oImage.Height / oImage.Width
                    If dblRatio > 2 Then
                        oImage.RotateFlip(RotateFlipType.Rotate270FlipX)
                        dblRatio = oImage.Width / oImage.Height
                    End If

                    Dim bmp As Bitmap
                    If dblRatio > 1 Then
                        If dblRatio > 5.5 Then dblRatio = 5.5
                        bmp = New Bitmap(120, CInt(120 / dblRatio))
                    Else
                        bmp = New Bitmap(CInt(120 * dblRatio), 120)
                    End If

                    Using g As Graphics = Graphics.FromImage(bmp)
                        g.DrawImage(oImage, 0, 0, bmp.Width, bmp.Height)
                    End Using
                    oResizedImage = bmp

                    myRow.Cells(_Columns._Picture).Value = oResizedImage ' oImage
                    myRow.Height = oResizedImage.Height ' IIf(oResizedImage.Height > 25, oResizedImage.Height + 5, 25)
                    dgvOrderLines.Rows(myRow.Index).Height = oResizedImage.Height
                    Debug.Print(myRow.Cells(0).Size.Width)
                    dgvOrderLines.Columns(0).Width = IIf(dgvOrderLines.Columns(0).Width > myRow.Cells(0).Size.Width, dgvOrderLines.Columns(0).Width, myRow.Cells(0).Size.Width)
                Else
                    Dim bmp As Bitmap = New Bitmap(124, 24)
                    oResizedImage = bmp
                    myRow.Cells(_Columns._Picture).Value = oResizedImage ' oImage
                    dgvOrderLines.Columns(0).Width = IIf(dgvOrderLines.Columns(0).Width > myRow.Cells(0).Size.Width, dgvOrderLines.Columns(0).Width, myRow.Cells(0).Size.Width)
                End If

            Next myRow

            dgvOrderLines.Columns.Remove("Route")

            Call AddComboBoxColumn("Route", _Columns._Route)

            For lPos As Integer = 0 To dgvOrderLines.Columns.Count - 1
                dgvOrderLines.Columns(lPos).SortMode = DataGridViewColumnSortMode.NotSortable
            Next lPos

            If dgvOrderLines.Columns(0).Width < 120 Then dgvOrderLines.Columns(0).Width = 120
            Debug.Print(dgvOrderLines.Columns(0).Width)

        Catch er As Exception
            MsgBox(er.Message)
        End Try

    End Sub

    Public ReadOnly Property Col_Picture() As Integer
        Get
            Return _Columns._Picture
        End Get
    End Property

    Public ReadOnly Property Col_Item_No() As Integer
        Get
            Return _Columns._Item_No
        End Get
    End Property

    Public ReadOnly Property Col_Description() As Integer
        Get
            Return _Columns._Description
        End Get
    End Property

    Public ReadOnly Property Col_Description2() As Integer
        Get
            Return _Columns._Description2
        End Get
    End Property

    Public ReadOnly Property Col_Qty_Ordered() As Integer
        Get
            Return _Columns._Qty_Ordered
        End Get
    End Property

    Public ReadOnly Property Col_Qty_Available() As Integer
        Get
            Return _Columns._Qty_Available
        End Get
    End Property

    Public ReadOnly Property Col_Qty_BO() As Integer
        Get
            Return _Columns._Qty_BO
        End Get
    End Property

    Public ReadOnly Property Col_Route() As Integer
        Get
            Return _Columns._Route
        End Get
    End Property

    Public ReadOnly Property ProductStatus() As Integer
        Get
            Return _ProductStatus
        End Get
    End Property


    Public Property Item_No() As String
        Get
            Return _Item_No
        End Get
        Set(ByVal value As String)
            _Item_No = value
        End Set
    End Property

    Private Sub dgvOrderLines_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvOrderLines.CellBeginEdit

        If dgvOrderLines.CurrentCell.ColumnIndex <> Col_Qty_Ordered And dgvOrderLines.CurrentCell.ColumnIndex <> _Columns._Route Then
            e.Cancel = True
            Exit Sub
        End If

        If dgvOrderLines.Rows(iRowPos).Cells(iColPos).Style.BackColor = Color.FromArgb(255, 192, 192) Then
            blnInvalidate = False
            dgvOrderLines.Rows(iRowPos).Cells(iColPos).Style.BackColor = Color.White
        End If

    End Sub


    Public Sub ChangeLineQty( _
        ByVal sender As Object, _
        ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) _
        Handles dgvOrderLines.CellValidating

        Dim strFormattedFalue As String

        e.Cancel = False

        If e.ColumnIndex <> _Columns._Qty_Ordered And e.ColumnIndex <> _Columns._Route Then Exit Sub

        'MB++ Actuellement, il y a un trou dans cette procedure.
        ' -------------------------------------------------------
        ' Si la QTY BO vient d'une autre commande et qu'on traite une commande récupérée (ce qui
        ' ne devrait pas arriver parce qu'on ne traitera que de nouvelles commandes), alors lors
        ' d'un retrait de qté de commandes BO, il faut rajouter du code pour valider la QTÉ BO
        ' de la commande actuelle parce qu'on traite la QTÉ BO comme étant globale pour le moment.
        '--------------------------------------------------------

        strFormattedFalue = e.FormattedValue

        Select Case e.ColumnIndex
            Case _Columns._Qty_Ordered
                If Trim(strFormattedFalue) = "" Then strFormattedFalue = "0"

                ' Pour la ligne courante
                With dgvOrderLines.Rows(e.RowIndex)

                    'If IsNumeric(.Cells(e.ColumnIndex).Value.ToString) Then
                    If IsNumeric(strFormattedFalue) Then

                        Dim dblDifference As Double

                        If .Cells(e.ColumnIndex).Tag Is Nothing Then
                            ' dblDifference = CDbl(.Cells(e.ColumnIndex).Value)
                            dblDifference = CDbl(strFormattedFalue)
                            If dblDifference > 0 Then
                                If dblDifference > CDbl(.Cells(_Columns._Qty_Available).Value) Then
                                    If CDbl(.Cells(_Columns._Qty_BO).Value) > 0 Then
                                        .Cells(_Columns._Qty_BO).Value = CDbl(.Cells(_Columns._Qty_BO).Value) + dblDifference
                                    Else
                                        .Cells(_Columns._Qty_Available).Value = CDbl(.Cells(_Columns._Qty_Available).Value) - dblDifference
                                        If CDbl(.Cells(_Columns._Qty_Available).Value) < 0 Then
                                            .Cells(_Columns._Qty_BO).Value = CDbl(.Cells(_Columns._Qty_BO).Value) - CDbl(.Cells(_Columns._Qty_Available).Value)
                                            .Cells(_Columns._Qty_Available).Value = 0
                                        End If
                                    End If
                                    .Cells(_Columns._Qty_Available).Value = 0
                                Else
                                    .Cells(_Columns._Qty_Available).Value = CDbl(.Cells(_Columns._Qty_Available).Value) - dblDifference
                                End If
                            Else

                                'dblDifference = Math.Abs(CDbl(.Cells(e.ColumnIndex).Value) - CDbl(.Cells(e.ColumnIndex).Tag))
                                dblDifference = Math.Abs(CDbl(strFormattedFalue) - CDbl(.Cells(e.ColumnIndex).Tag))

                                If dblDifference > CDbl(.Cells(_Columns._Qty_BO).Value) Then
                                    If CDbl(.Cells(_Columns._Qty_BO).Value) = 0 Then
                                        .Cells(_Columns._Qty_Available).Value = CDbl(.Cells(_Columns._Qty_Available).Value) + dblDifference
                                    Else
                                        .Cells(_Columns._Qty_BO).Value = CDbl(.Cells(_Columns._Qty_BO).Value) - dblDifference
                                        If CDbl(.Cells(_Columns._Qty_BO).Value) < 0 Then
                                            .Cells(_Columns._Qty_Available).Value = CDbl(.Cells(_Columns._Qty_Available).Value) - CDbl(.Cells(_Columns._Qty_BO).Value)
                                            .Cells(_Columns._Qty_BO).Value = 0
                                        End If
                                    End If
                                Else
                                    .Cells(_Columns._Qty_BO).Value = CDbl(.Cells(_Columns._Qty_BO).Value) - dblDifference
                                End If

                            End If

                        ElseIf dgvOrderLines.Rows(e.RowIndex).Cells(e.ColumnIndex).Tag.ToString <> strFormattedFalue Then

                            'dblDifference = CDbl(.Cells(e.ColumnIndex).Value) - CDbl(.Cells(e.ColumnIndex).Tag)
                            dblDifference = CDbl(strFormattedFalue) - CDbl(.Cells(e.ColumnIndex).Tag)

                            ' Difference is positive we add items
                            If dblDifference > 0 Then
                                If dblDifference > CDbl(.Cells(_Columns._Qty_Available).Value) Then
                                    If CDbl(.Cells(_Columns._Qty_BO).Value) > 0 Then
                                        .Cells(_Columns._Qty_BO).Value = CDbl(.Cells(_Columns._Qty_BO).Value) + dblDifference
                                    Else
                                        .Cells(_Columns._Qty_Available).Value = CDbl(.Cells(_Columns._Qty_Available).Value) - dblDifference
                                        If CDbl(.Cells(_Columns._Qty_Available).Value) < 0 Then
                                            .Cells(_Columns._Qty_BO).Value = CDbl(.Cells(_Columns._Qty_BO).Value) - CDbl(.Cells(_Columns._Qty_Available).Value)
                                            .Cells(_Columns._Qty_Available).Value = 0
                                        End If
                                    End If
                                    .Cells(_Columns._Qty_Available).Value = 0
                                Else
                                    .Cells(_Columns._Qty_Available).Value = CDbl(.Cells(_Columns._Qty_Available).Value) - dblDifference
                                End If

                                ' difference is negative, we remove items   
                            Else
                                'dblDifference = Math.Abs(CDbl(.Cells(e.ColumnIndex).Value) - CDbl(.Cells(e.ColumnIndex).Tag))
                                dblDifference = Math.Abs(CDbl(strFormattedFalue) - CDbl(.Cells(e.ColumnIndex).Tag))

                                If dblDifference > CDbl(.Cells(_Columns._Qty_BO).Value) Then
                                    If CDbl(.Cells(_Columns._Qty_BO).Value) = 0 Then
                                        .Cells(_Columns._Qty_Available).Value = CDbl(.Cells(_Columns._Qty_Available).Value) + dblDifference
                                    Else
                                        .Cells(_Columns._Qty_BO).Value = CDbl(.Cells(_Columns._Qty_BO).Value) - dblDifference
                                        If CDbl(.Cells(_Columns._Qty_BO).Value) < 0 Then
                                            .Cells(_Columns._Qty_Available).Value = CDbl(.Cells(_Columns._Qty_Available).Value) - CDbl(.Cells(_Columns._Qty_BO).Value)
                                            .Cells(_Columns._Qty_BO).Value = 0
                                        End If
                                    End If
                                Else
                                    .Cells(_Columns._Qty_BO).Value = CDbl(.Cells(_Columns._Qty_BO).Value) - dblDifference
                                End If
                            End If

                        End If

                        ' On se sert du TAG pour enregistrer une OldValue
                        'dgvOrderLines.Rows(e.RowIndex).Cells(e.ColumnIndex).Tag = dgvOrderLines.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                        dgvOrderLines.Rows(e.RowIndex).Cells(e.ColumnIndex).Tag = strFormattedFalue

                    Else

                        blnInvalidate = True
                        iRowPos = e.RowIndex
                        iColPos = e.ColumnIndex
                        dgvOrderLines.Invalidate()

                    End If

                End With

        End Select

    End Sub

    Public Sub dgvOrderLines_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles dgvOrderLines.Paint

        If Not (dgvOrderLines.Rows.Count = 0) Then

            If blnInvalidate Then
                dgvOrderLines.Rows(iRowPos).Cells(iColPos).Style.BackColor = Color.FromArgb(255, 192, 192)
                dgvOrderLines.Rows(iRowPos).Cells(iColPos).ToolTipText = "Value must be numeric"
            Else
                dgvOrderLines.Rows(iRowPos).Cells(iColPos).Style.BackColor = Color.White
                dgvOrderLines.Rows(iRowPos).Cells(iColPos).ToolTipText = ""
            End If

            iColPos = 0
            iRowPos = 0

        End If

    End Sub

    Private Sub dgvOrderLines_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvOrderLines.CellClick

        If e.ColumnIndex = _Columns._Route And e.RowIndex <> -1 Then

            'If dgvOrderLines.Rows(e.RowIndex).Cells(e.ColumnIndex).Tag = "" Then

            Call SetRouteCombo(dgvOrderLines.Columns(_Columns._Route))

            'dgvOrderLines.Rows(e.RowIndex).Cells(e.ColumnIndex).Tag = "Listed"
            'End If

        End If

    End Sub

    Private Sub SetRouteCombo(ByVal cboColumn As DataGridViewComboBoxColumn)

        With cboColumn
            ' .DataSource = RetrieveAlternativeTitles()

            .DataSource = PopulateComboBox( _
            "SELECT     Rte.RouteId, Rte.RouteDescription " & _
            "FROM       Exact_Route Rte " & _
            "INNER JOIN EXACT_TRAVELER_Xref_ProductCategoryRoute Xref " & _
            "           ON Rte.RouteID = Xref.RouteID " & _
            "INNER JOIN imitmidx_sql I WITH (Nolock) ON XRef.Prod_Cat = I.Prod_Cat " & _
            "WHERE      Rte.active <> 0 AND " & _
            "           I.Item_No = '" & dgvOrderLines.Rows(dgvOrderLines.CurrentCell.RowIndex).Cells(_Columns._Item_No).Value & "' " & _
            "ORDER BY   RouteDescription ")

            '.ValueMember = "RouteId" ' ColumnName.TitleOfCourtesy.ToString()
            .DisplayMember = "RouteDescription"

        End With

    End Sub

    Private Sub AddComboBoxColumn(ByVal pstrColumnName As String, ByVal pintPos As Integer)

        Dim cboColumn As New DataGridViewComboBoxColumn

        With cboColumn
            .HeaderText = pstrColumnName
            .DataPropertyName = pstrColumnName
            .DropDownWidth = 200
            .Width = 200
            .MaxDropDownItems = 12
            .FlatStyle = FlatStyle.Standard
        End With

        dgvOrderLines.Columns.Insert(pintPos, cboColumn)

    End Sub

    Private Function PopulateComboBox(ByVal pstrSqlCommand As String) As DataTable

        Dim myConnection As SqlConnection = New SqlConnection("Server=thor;Initial Catalog=100;Integrated Security=SSPI")
        Dim mySelectCommand As SqlCommand = New SqlCommand(pstrSqlCommand, myConnection)
        Dim mySqlDataAdapter As SqlDataAdapter = New SqlDataAdapter(mySelectCommand)

        Dim myDataSet As New DataSet()
        mySqlDataAdapter.Fill(myDataSet)
        Return myDataSet.Tables(0) ' dgvOrderLines.DataSource = myDataSet.Tables(0)

    End Function

    Private Sub cmdResetAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResetAll.Click

        If dgvOrderLines.Rows.Count > 0 Then

            For Each oRow As DataGridViewRow In dgvOrderLines.Rows
                oRow.Cells(Col_Qty_Available).Value = 0
                oRow.Cells(Col_Qty_Ordered).Value = 0
                oRow.Cells(Col_Qty_BO).Value = 0
            Next

        End If

    End Sub

    Private Sub cmdApplyToAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApplyToAll.Click

        Dim lRow As Long = dgvOrderLines.CurrentCell.RowIndex
        Dim dblQtyOrdered As Double = dgvOrderLines.Rows(lRow).Cells(Col_Qty_Ordered).Value
        Dim strItem As String = dgvOrderLines.Rows(lRow).Cells(Col_Item_No).Value

        If dgvOrderLines.Rows.Count > 0 Then

            For Each oRow As DataGridViewRow In dgvOrderLines.Rows

                If oRow.Cells(Col_Item_No).Value <> strItem Then
                    oRow.Cells(Col_Qty_Ordered).Value = dblQtyOrdered ' .dgvOrderLines.RaiseEvent ChangeLineQty()
                End If
            Next

        End If

    End Sub

    Private Sub frmProductLineEntry_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        LoadProductGrid()

    End Sub

    Private Sub frmProductLineEntry_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        If Me.Width < 500 Then Me.Width = 500
        If Me.Height < 200 Then Me.Height = 200

        dgvOrderLines.Height = Me.Height - 90
        dgvOrderLines.Width = Me.Width - 30

        cmdOK.Top = Me.Height - 66
        cmdCancel.Top = cmdOK.Top
        cmdApplyToAll.Top = cmdOK.Top
        cmdResetAll.Top = cmdOK.Top

        cmdOK.Left = Me.Width - 225
        cmdCancel.Left = Me.Width - 120

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        _ProductStatus = Status.Insert
        Close()

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        _ProductStatus = Status.Cancel
        Close()

    End Sub

End Class