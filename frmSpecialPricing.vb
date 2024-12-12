Public Class frmSpecialPricing



#Region "################## FORM EVENTS  ################################# "
    '========================================================================================================================
    '========================================================================================================================
    Private Sub frmSpecialPricing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '==============================
        'SET CONTROL VALUES IN FORM
        '==============================
        Me.Height = cmdViewSpecialPricesHistory.Top + cmdViewSpecialPricesHistory.Height + 40

        txtCusNo.Text = spCusNo
        txtItemNo.Text = spItemNo

        'Load the Date Picker with date 1 year previous to today
        dtpEndDate.Text = DateAdd(DateInterval.Year, -1, Today())
        '==============================
        ' LOAD GRIDS
        '==============================
        'Call GridPriceCodes_Format() No need to display this for now - Clayton
        Call GridSpecialPrices_Format()

        'gridPriceCodes.DefaultCellStyle.Font = New Drawing.Font("Microsoft Sans Serif", 8)
        'gridPriceCodes.RowTemplate.Height = 14
    End Sub
    '====================================================================================================
    Private Sub frmSpecialPricing_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        '==============================
        'RESIZE GRIDS WHEN FORM WIDTH CHANGES
        '==============================
        'Changes width of grids when form is resized.
        gridPriceCodes.Width = Me.Width - (4 * gridPriceCodes.Location.X)
        gridSpecialPrices.Width = Me.Width - (4 * gridSpecialPrices.Location.X)
        gridSpecialPricesHISTORY.Width = Me.Width - (4 * gridSpecialPricesHISTORY.Location.X)
    End Sub
    '====================================================================================================
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        '==============================
        'CLOSE THE FORM AND CLEAR MEMORY
        '==============================
        Me.Close()
        Me.Dispose()
    End Sub

    '========================================================================================================================
    '========================================================================================================================
#End Region '"################## END FORM EVENTS  ################################# "

#Region "################## PRICE CODES GRID ################################# "
    'Private Sub gridPriceCodes_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles gridPriceCodes.CellFormatting
    '    '================================
    '    'REPLACE 0's with empty cells
    '    '================================
    '    If IsNumeric(e.Value) = True Then
    '        If e.Value = 0 Then
    '            e.Value = ""
    '            e.FormattingApplied = True
    '        End If
    '    End If
    'End Sub

    'Public Sub GridPriceCodes_Format()
    '    Try
    '        '================================
    '        'SET GRID LINE COLOURS AND PROPERTIES
    '        '================================
    '        With Me.gridPriceCodes
    '            .RowsDefaultCellStyle.BackColor = Color.White
    '            .AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue    'LightCyan
    '            .ReadOnly = True
    '            .AllowUserToAddRows = False
    '            .AllowUserToDeleteRows = False
    '            .AllowUserToOrderColumns = True

    '        End With
    '        '================================
    '        'CREATE COLUMN WIDTHS AND HEADINGS
    '        '================================

    '        ' Public Function AddGridColumn(FieldType As GridColTypeEnum, FieldName As String, Title As String, Visible As Boolean, Locked As Boolean, Optional Width As Long = 100, Optional Alignment As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleRight, Optional Format As GridFormatEnum = 0) As Object
    '        With gridPriceCodes.Columns
    '            .Clear()
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "FirstCol", "FirstCol", cVisible_FALSE, cLocked_TRUE, 30, DataGridViewContentAlignment.MiddleLeft))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "CodeType", "CodeType", cVisible_TRUE, cLocked_TRUE, 60, DataGridViewContentAlignment.MiddleLeft))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Currency", "Currency", cVisible_TRUE, cLocked_TRUE, 60, DataGridViewContentAlignment.MiddleLeft))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "CusNo", "CusNo", cVisible_FALSE, cLocked_TRUE, 100, DataGridViewContentAlignment.MiddleLeft))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "CusType", "CusType", cVisible_FALSE, cLocked_TRUE, 50, DataGridViewContentAlignment.MiddleLeft))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "ItemNo", "ItemNo", cVisible_FALSE, cLocked_TRUE, 140, DataGridViewContentAlignment.MiddleLeft))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "ProdCat", "Program/SP", cVisible_FALSE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleLeft))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Extra_14", "Prog Code", cVisible_FALSE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleLeft))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "StartDt", "StartDt", cVisible_TRUE, cLocked_TRUE, 70, DataGridViewContentAlignment.MiddleLeft, GridFormatEnum.ShortDate))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "EndDt", "EndDt", cVisible_TRUE, cLocked_TRUE, 70, DataGridViewContentAlignment.MiddleLeft, GridFormatEnum.ShortDate))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "PriceBasis", "PriceBasis", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty1", "Qty1", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price1", "Price1", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty2", "Qty2", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price2", "Price2", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty3", "Qty3", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price3", "Price3", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty4", "Qty4", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price4", "Price4", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty5", "Qty5", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price5", "Price5", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty6", "Qty6", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price6", "Price6", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty7", "Qty7", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price7", "Price7", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty8", "Qty8", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price8", "Price8", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty9", "Qty9", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price9", "Price9", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty10", "Qty10", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price10", "Price10", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
    '            .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "LastCol", "LastCol", cVisible_FALSE, cLocked_TRUE, 30, DataGridViewContentAlignment.MiddleLeft))
    '        End With

    '        '================================
    '        'Fill the GRID
    '        '================================
    '        Call GridPriceCodes_LoadValues()

    '        '================================
    '        'SET FROZEN COLUMNS
    '        '================================
    '        gridPriceCodes.Columns(PriceCodesEnum.FirstCol).Frozen = True


    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try
    'End Sub

    'Public Sub GridPriceCodes_LoadValues()
    '    Try

    '        Dim sqlLOC As String

    '        sqlLOC = "SELECT  "
    '        sqlLOC = sqlLOC & " FirstCol = '0', " & vbCrLf
    '        sqlLOC = sqlLOC & " CodeType = PRC.cd_tp, " & vbCrLf
    '        sqlLOC = sqlLOC & " Currency = PRC.curr_cd, " & vbCrLf
    '        sqlLOC = sqlLOC & " CusNo = PRC.cd_tp_1_cust_no, " & vbCrLf
    '        sqlLOC = sqlLOC & " CusType = PRC.cd_tp_3_cust_type, " & vbCrLf
    '        sqlLOC = sqlLOC & " ItemNo = PRC.cd_tp_1_item_no, " & vbCrLf
    '        sqlLOC = sqlLOC & " ProdCat = PRC.cd_tp_2_prod_cat, " & vbCrLf
    '        sqlLOC = sqlLOC & " StartDt = PRC.start_dt, " & vbCrLf
    '        sqlLOC = sqlLOC & " EndDt = PRC.end_dt, " & vbCrLf
    '        sqlLOC = sqlLOC & " PriceBasis = PRC.cd_prc_basis, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty1 = PRC.minimum_qty_1, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price1 = PRC.prc_or_disc_1, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty2 = PRC.minimum_qty_2, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price2 = PRC.prc_or_disc_2, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty3 = PRC.minimum_qty_3, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price3 = PRC.prc_or_disc_3, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty4 = PRC.minimum_qty_4, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price4 = PRC.prc_or_disc_4, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty5 = PRC.minimum_qty_5, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price5 = PRC.prc_or_disc_5, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty6 = PRC.minimum_qty_6, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price6 = PRC.prc_or_disc_6, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty7 = PRC.minimum_qty_7, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price7 = PRC.prc_or_disc_7, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty8 = PRC.minimum_qty_8, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price8 = PRC.prc_or_disc_8, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty9 = PRC.minimum_qty_9, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price9 = PRC.prc_or_disc_9, " & vbCrLf
    '        sqlLOC = sqlLOC & " Qty10 = PRC.minimum_qty_10, " & vbCrLf
    '        sqlLOC = sqlLOC & " Price10 = PRC.prc_or_disc_10, " & vbCrLf
    '        sqlLOC = sqlLOC & " Extra_14 = PRC.extra_14, " & vbCrLf
    '        sqlLOC = sqlLOC & " LastCol = '0' " & vbCrLf
    '        sqlLOC = sqlLOC & " FROM oeprcfil_Sql PRC"

    '        sqlLOC = sqlLOC & " WHERE 0=0 " & vbCrLf
    '        sqlLOC = sqlLOC & " AND PRC.curr_cd = (Select Currency from cicmpy where debcode = '" & spCusNo & "')" & vbCrLf
    '        sqlLOC = sqlLOC & " AND getdate() between isnull(PRC.start_dt,'1900-01-01') and PRC.end_dt " & vbCrLf
    '        sqlLOC = sqlLOC & " AND (    ( PRC.Cd_tp = 1 AND PRC.cd_tp_1_cust_no = '" & spCusNo & "' AND PRC.cd_tp_1_item_no = '" & sqlSAFE(spItemNo) & "')" & vbCrLf
    '        sqlLOC = sqlLOC & "       OR ( PRC.Cd_tp = 2 AND PRC.cd_tp_1_cust_no = '" & spCusNo & "' AND PRC.cd_tp_2_prod_cat = (Select prod_cat from imitmidx_sql where item_no ='" & sqlSAFE(spItemNo) & "'))" & vbCrLf
    '        sqlLOC = sqlLOC & "       OR ( PRC.Cd_tp = 3 AND PRC.cd_tp_3_cust_type = (Select cmp_type from cicmpy where debcode = '" & spCusNo & "') AND cd_tp_1_item_no = '" & sqlSAFE(spItemNo) & "')" & vbCrLf
    '        sqlLOC = sqlLOC & "       OR ( PRC.Cd_tp = 4 AND PRC.cd_tp_3_cust_type = (Select cmp_type from cicmpy where debcode = '" & spCusNo & "') AND cd_tp_2_prod_cat = (Select prod_cat from imitmidx_sql where item_no ='" & sqlSAFE(spItemNo) & "'))" & vbCrLf
    '        sqlLOC = sqlLOC & "       OR ( PRC.Cd_tp = 5 AND PRC.cd_tp_1_cust_no = '" & spCusNo & "' )" & vbCrLf
    '        sqlLOC = sqlLOC & "       OR ( PRC.Cd_tp = 6 AND PRC.cd_tp_1_item_no = '" & sqlSAFE(spItemNo) & "')" & vbCrLf
    '        sqlLOC = sqlLOC & "       OR ( PRC.Cd_tp = 7 AND PRC.cd_tp_3_cust_type = (Select cmp_type from cicmpy where debcode = '" & spCusNo & "')) " & vbCrLf
    '        sqlLOC = sqlLOC & "       OR ( PRC.Cd_tp = 8 AND PRC.cd_tp_2_prod_cat = (Select prod_cat from imitmidx_sql where item_no ='" & sqlSAFE(spItemNo) & "')) " & vbCrLf
    '        sqlLOC = sqlLOC & "     ) " & vbCrLf

    '        Dim dt As DataTable
    '        Dim db As New cDBA

    '        dt = db.DataTable(sqlLOC)
    '        gridPriceCodes.DataSource = dt

    '    Catch er As Exception
    '        MsgBox("Error In " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub
#End Region '"################## END PRICE CODES GRID ################################# "

#Region "################## SPECIAL PRICES GRID ################################# "
    Private Sub gridSpecialPrices_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles gridSpecialPrices.CellFormatting
        '================================
        'REPLACE 0's with empty cells
        '================================
        If IsNumeric(e.Value) = True Then
            If e.Value = 0 Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        End If
        'If e.Value.ToString = "0.0000" Or e.Value.ToString = "0.000000" Then
        '    e.Value = ""
        '    e.FormattingApplied = True
        'End If
    End Sub

    Public Sub GridSpecialPrices_Format()

        Try
            '================================
            'SET GRID LINE COLOURS AND PROPERTIES
            '================================
            With Me.gridSpecialPrices
                .RowsDefaultCellStyle.BackColor = Color.White
                '.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue    'LightCyan
                .ReadOnly = True
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToOrderColumns = True
            End With
            '================================
            'CREATE COLUMN WIDTHS AND HEADINGS
            '================================

            ' Public Function AddGridColumn(FieldType As GridColTypeEnum, FieldName As String, Title As String, Visible As Boolean, Locked As Boolean, Optional Width As Long = 100, Optional Alignment As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleRight, Optional Format As GridFormatEnum = 0) As Object

            With gridSpecialPrices.Columns
                .Clear()
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "FirstCol", "FirstCol", cVisible_FALSE, cLocked_TRUE, 30, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "CodeType", "CodeType", cVisible_TRUE, cLocked_TRUE, 60, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Currency", "Currency", cVisible_TRUE, cLocked_TRUE, 60, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "CusNo", "CusNo", cVisible_FALSE, cLocked_TRUE, 100, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "CusType", "CusType", cVisible_FALSE, cLocked_TRUE, 50, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "ItemNo", "ItemNo", cVisible_TRUE, cLocked_TRUE, 140, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "ProdCat", "Program/SP", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Extra_14", "Prog Code", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "StartDt", "StartDt", cVisible_TRUE, cLocked_TRUE, 70, DataGridViewContentAlignment.MiddleLeft, GridFormatEnum.ShortDate))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "EndDt", "EndDt", cVisible_TRUE, cLocked_TRUE, 70, DataGridViewContentAlignment.MiddleLeft, GridFormatEnum.ShortDate))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "PriceBasis", "PriceBasis", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty1", "Qty1", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price1", "Price1", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty2", "Qty2", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price2", "Price2", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty3", "Qty3", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price3", "Price3", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "DEC_MET_ID", "Decorating", cVisible_TRUE, cLocked_TRUE, 100, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.Text))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Setup_waived", "Setup Waived", cVisible_TRUE, cLocked_TRUE, 80, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.Text))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Setup_price", "Special Setup Price", cVisible_TRUE, cLocked_TRUE, 80, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Charge_Id", "Charge", cVisible_TRUE, cLocked_TRUE, 100, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.Text))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Run_Charge", "RunCharge", cVisible_TRUE, cLocked_TRUE, 80, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price6", "Price6", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty7", "Qty7", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price7", "Price7", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty4", "Qty4", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price4", "Price4", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty5", "Qty5", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price5", "Price5", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty6", "Qty6", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price6", "Price6", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty7", "Qty7", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price7", "Price7", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty8", "Qty8", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price8", "Price8", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty9", "Qty9", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price9", "Price9", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty10", "Qty10", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price10", "Price10", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "LastCol", "LastCol", cVisible_FALSE, cLocked_TRUE, 30, DataGridViewContentAlignment.MiddleLeft))
            End With

            '================================
            'SET FROZEN COLUMNS
            '================================
            gridSpecialPrices.Columns(PriceCodesEnum.FirstCol).Frozen = True

            '================================
            'Fill the GRID
            '================================
            Call GridSpecialPrices_LoadValues()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub GridSpecialPrices_LoadValues()
        Try

            Dim sqlLOC As String

            sqlLOC = "SELECT DISTINCT "
            sqlLOC = sqlLOC & " FirstCol = '0', " & vbCrLf
            sqlLOC = sqlLOC & " CodeType = PRC.cd_tp, " & vbCrLf
            sqlLOC = sqlLOC & " Currency = PRC.curr_cd, " & vbCrLf
            sqlLOC = sqlLOC & " CusNo = PRC.cd_tp_1_cust_no, " & vbCrLf
            sqlLOC = sqlLOC & " CusType = PRC.cd_tp_3_cust_type, " & vbCrLf
            sqlLOC = sqlLOC & " ItemNo = PRC.cd_tp_1_item_no, " & vbCrLf
            sqlLOC = sqlLOC & " ProdCat = PRC.cd_tp_2_prod_cat, " & vbCrLf
            sqlLOC = sqlLOC & " Extra_14 = PRC.extra_14, " & vbCrLf
            sqlLOC = sqlLOC & " StartDt = PRC.start_dt, " & vbCrLf
            sqlLOC = sqlLOC & " EndDt = PRC.end_dt, " & vbCrLf
            sqlLOC = sqlLOC & " PriceBasis = PRC.cd_prc_basis, " & vbCrLf
            sqlLOC = sqlLOC & " Qty1 = PRC.minimum_qty_1, " & vbCrLf
            sqlLOC = sqlLOC & " Price1 = PRC.prc_or_disc_1, " & vbCrLf
            sqlLOC = sqlLOC & " Qty2 = PRC.minimum_qty_2, " & vbCrLf
            sqlLOC = sqlLOC & " Price2 = PRC.prc_or_disc_2, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty3 = PRC.minimum_qty_3, " & vbCrLf
            'sqlLOC = sqlLOC & " Price3 = PRC.prc_or_disc_3, " & vbCrLf
            'sqlLOC = sqlLOC & " Decorating =  M.DEC_MET_CD, " & vbCrLf
            sqlLOC = sqlLOC & " Setup_waived = I.Setup_waived, " & vbCrLf
            sqlLOC = sqlLOC & " Setup_price = I.Setup_price, " & vbCrLf
            ' sqlLOC = sqlLOC & " Charge = C.DESCRIPTION, " & vbCrLf
            sqlLOC = sqlLOC & " RunCharge =  I.Run_Charge, " & vbCrLf
            sqlLOC = sqlLOC & " Comments =  P.Prog_Comments, " & vbCrLf 'to get comments
            'sqlLOC = sqlLOC & " Qty4 = PRC.minimum_qty_4, " & vbCrLf
            'sqlLOC = sqlLOC & " Price4 = PRC.prc_or_disc_4, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty5 = PRC.minimum_qty_5, " & vbCrLf
            'sqlLOC = sqlLOC & " Price5 = PRC.prc_or_disc_5, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty6 = PRC.minimum_qty_6, " & vbCrLf
            'sqlLOC = sqlLOC & " Price6 = PRC.prc_or_disc_6, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty7 = PRC.minimum_qty_7, " & vbCrLf
            'sqlLOC = sqlLOC & " Price7 = PRC.prc_or_disc_7, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty8 = PRC.minimum_qty_8, " & vbCrLf
            'sqlLOC = sqlLOC & " Price8 = PRC.prc_or_disc_8, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty9 = PRC.minimum_qty_9, " & vbCrLf
            'sqlLOC = sqlLOC & " Price9 = PRC.prc_or_disc_9, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty10 = PRC.minimum_qty_10, " & vbCrLf
            'sqlLOC = sqlLOC & " Price10 = PRC.prc_or_disc_10, " & vbCrLf
            sqlLOC = sqlLOC & " LastCol = '0' " & vbCrLf
            sqlLOC = sqlLOC & " FROM  MDB_CUS_PROG P  LEFT JOIN " & vbCrLf
            sqlLOC = sqlLOC & " MDB_CUS_PROG_ITEM_LIST I  ON I.CUS_PROG_ID = P.CUS_PROG_ID LEFT JOIN " & vbCrLf
            sqlLOC = sqlLOC & " oeprcfil_Sql PRC On PRC.extra_14 = P.CUS_PROG_ID " & vbCrLf
            sqlLOC = sqlLOC & " WHERE PRC.curr_cd = (Select Currency from cicmpy where debcode = '" & spCusNo & "')" & vbCrLf
            sqlLOC = sqlLOC & " AND getdate() between isnull(PRC.start_dt,'1900-01-01') and PRC.end_dt " & vbCrLf
            sqlLOC = sqlLOC & " AND PRC.cd_tp_1_cust_no = '" & spCusNo & "' AND PRC.cd_tp_1_item_no = '" & sqlSAFE(spItemNo) & "'" & vbCrLf
            'AND PRC.Cd_tp = 9  removed from criteria by Clayton to allow wider search of program and special prices in Macola.

            Dim dt As DataTable
            Dim db As New cDBA

            dt = db.DataTable(sqlLOC)
            gridSpecialPrices.DataSource = dt

            If dt.Rows.Count > 0 Then
                txtProg_Comments.Text = dt.Rows(0).Item("Comments").ToString
            Else
                MessageBox.Show("Unable to locate a program pricing for item " & spItemNo)
                Me.Close() 'close the form if no data is found
            End If

        Catch er As Exception
            MsgBox("Error In " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region '"################## END SPECIAL PRICES GRID ################################# "

#Region "################## SPECIAL PRICES HISTORY GRID ################################# "
    Private Sub gridSpecialPricesHISTORY_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles gridSpecialPricesHISTORY.CellFormatting
        '================================
        'REPLACE 0's with empty cells
        '================================
        If IsNumeric(e.Value) = True Then
            If e.Value = 0 Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        End If
        'If e.Value.ToString = "0.0000" Or e.Value.ToString = "0.000000" Then
        '    e.Value = ""
        '    e.FormattingApplied = True
        'End If
    End Sub

    Public Sub gridSpecialPricesHISTORY_Format()

        Try
            '================================
            'SET GRID LINE COLOURS AND PROPERTIES
            '================================
            With Me.gridSpecialPricesHISTORY
                .RowsDefaultCellStyle.BackColor = Color.White
                .AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue    'LightCyan
                .ReadOnly = True
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToOrderColumns = True
            End With
            '================================
            'CREATE COLUMN WIDTHS AND HEADINGS
            '================================

            ' Public Function AddGridColumn(FieldType As GridColTypeEnum, FieldName As String, Title As String, Visible As Boolean, Locked As Boolean, Optional Width As Long = 100, Optional Alignment As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleRight, Optional Format As GridFormatEnum = 0) As Object

            With gridSpecialPricesHISTORY.Columns
                .Clear()
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "FirstCol", "FirstCol", cVisible_FALSE, cLocked_TRUE, 30, DataGridViewContentAlignment.MiddleLeft))
                gridSpecialPricesHISTORY.Columns(PriceCodesEnum.FirstCol).Visible = False            'Added back in Sept 12, 2017  T. Louzon
                'MsgBox(gridSpecialPricesHISTORY.Columns(PriceCodesEnum.FirstCol).ToString & "   " & gridSpecialPricesHISTORY.Columns(PriceCodesEnum.ItemNo).ToString)

                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "CodeType", "CodeType", cVisible_TRUE, cLocked_TRUE, 60, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Currency", "Currency", cVisible_TRUE, cLocked_TRUE, 60, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "CusNo", "CusNo", cVisible_FALSE, cLocked_TRUE, 100, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "CusType", "CusType", cVisible_FALSE, cLocked_TRUE, 50, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "ItemNo", "ItemNo", cVisible_TRUE, cLocked_TRUE, 140, DataGridViewContentAlignment.MiddleLeft))     '   Make this visible in History
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "ProdCat", "Program/SP", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Extra_14", "Prog Code", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleLeft))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "StartDt", "StartDt", cVisible_TRUE, cLocked_TRUE, 70, DataGridViewContentAlignment.MiddleLeft, GridFormatEnum.ShortDate))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "EndDt", "EndDt", cVisible_TRUE, cLocked_TRUE, 70, DataGridViewContentAlignment.MiddleLeft, GridFormatEnum.ShortDate))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "PriceBasis", "PriceBasis", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty1", "Qty1", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price1", "Price1", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty2", "Qty2", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price2", "Price2", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty3", "Qty3", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price3", "Price3", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty4", "Qty4", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price4", "Price4", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty5", "Qty5", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price5", "Price5", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty6", "Qty6", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price6", "Price6", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty7", "Qty7", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price7", "Price7", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty8", "Qty8", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price8", "Price8", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty9", "Qty9", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price9", "Price9", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Qty10", "Qty10", cVisible_TRUE, cLocked_TRUE, 45, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N0))
                '.Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "Price10", "Price10", cVisible_TRUE, cLocked_TRUE, 65, DataGridViewContentAlignment.MiddleRight, GridFormatEnum.N4))
                .Add(AddGridColumn(GridColTypeEnum.TextBoxColumn, "LastCol", "LastCol", cVisible_FALSE, cLocked_TRUE, 30, DataGridViewContentAlignment.MiddleLeft))
            End With

            '================================
            'Fill the GRID
            '================================
            Call gridSpecialPricesHISTORY_LoadValues()

            '================================
            'SET FROZEN COLUMNS
            '================================
            gridSpecialPricesHISTORY.Columns(PriceCodesEnum.FirstCol).Frozen = True

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Public Sub gridSpecialPricesHISTORY_LoadValues()
        Try

            Dim sqlLOC As String

            sqlLOC = "SELECT  "
            sqlLOC = sqlLOC & " FirstCol = '0', " & vbCrLf
            sqlLOC = sqlLOC & " CodeType = PRC.cd_tp, " & vbCrLf
            sqlLOC = sqlLOC & " Currency = PRC.curr_cd, " & vbCrLf
            sqlLOC = sqlLOC & " CusNo = PRC.cd_tp_1_cust_no, " & vbCrLf
            sqlLOC = sqlLOC & " CusType = PRC.cd_tp_3_cust_type, " & vbCrLf
            sqlLOC = sqlLOC & " ItemNo = PRC.cd_tp_1_item_no, " & vbCrLf
            sqlLOC = sqlLOC & " ProdCat = PRC.cd_tp_2_prod_cat, " & vbCrLf
            sqlLOC = sqlLOC & " StartDt = PRC.start_dt, " & vbCrLf
            sqlLOC = sqlLOC & " EndDt = PRC.end_dt, " & vbCrLf
            sqlLOC = sqlLOC & " PriceBasis = PRC.cd_prc_basis, " & vbCrLf
            sqlLOC = sqlLOC & " Qty1 = PRC.minimum_qty_1, " & vbCrLf
            sqlLOC = sqlLOC & " Price1 = PRC.prc_or_disc_1, " & vbCrLf
            sqlLOC = sqlLOC & " Qty2 = PRC.minimum_qty_2, " & vbCrLf
            sqlLOC = sqlLOC & " Price2 = PRC.prc_or_disc_2, " & vbCrLf
            sqlLOC = sqlLOC & " Qty3 = PRC.minimum_qty_3, " & vbCrLf
            sqlLOC = sqlLOC & " Price3 = PRC.prc_or_disc_3, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty4 = PRC.minimum_qty_4, " & vbCrLf
            'sqlLOC = sqlLOC & " Price4 = PRC.prc_or_disc_4, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty5 = PRC.minimum_qty_5, " & vbCrLf
            'sqlLOC = sqlLOC & " Price5 = PRC.prc_or_disc_5, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty6 = PRC.minimum_qty_6, " & vbCrLf
            'sqlLOC = sqlLOC & " Price6 = PRC.prc_or_disc_6, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty7 = PRC.minimum_qty_7, " & vbCrLf
            'sqlLOC = sqlLOC & " Price7 = PRC.prc_or_disc_7, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty8 = PRC.minimum_qty_8, " & vbCrLf
            'sqlLOC = sqlLOC & " Price8 = PRC.prc_or_disc_8, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty9 = PRC.minimum_qty_9, " & vbCrLf
            'sqlLOC = sqlLOC & " Price9 = PRC.prc_or_disc_9, " & vbCrLf
            'sqlLOC = sqlLOC & " Qty10 = PRC.minimum_qty_10, " & vbCrLf
            'sqlLOC = sqlLOC & " Price10 = PRC.prc_or_disc_10, " & vbCrLf
            'sqlLOC = sqlLOC & " Extra_14 = PRC.extra_14, " & vbCrLf
            sqlLOC = sqlLOC & " LastCol = '0' " & vbCrLf
            sqlLOC = sqlLOC & " FROM oeprcfil_Sql PRC "
            sqlLOC = sqlLOC & " WHERE 0=0 " & vbCrLf
            sqlLOC = sqlLOC & " AND PRC.curr_cd = (Select Currency from cicmpy where debcode = '" & spCusNo & "')" & vbCrLf
            sqlLOC = sqlLOC & " AND isnull(PRC.start_dt,'1900-01-01') <= getdate() " & vbCrLf
            sqlLOC = sqlLOC & " AND PRC.end_dt >= '" & dtpEndDate.Value.Date & "'" & vbCrLf
            'sqlLOC = sqlLOC & " AND PRC.Cd_tp = 9 " & vbCrLf removed from criteria by Clayton to allow wider search on pricing
            sqlLOC = sqlLOC & " AND PRC.cd_tp_1_cust_no = '" & spCusNo & "'" & vbCrLf

            If optThisItem.Checked = True Then
                'Use only this item
                sqlLOC = sqlLOC & " AND  PRC.cd_tp_1_item_no = '" & sqlSAFE(spItemNo) & "'" & vbCrLf
            ElseIf optThisStyle.Checked = True Then
                'Get ALL items with the same Style as shown
                sqlLOC = sqlLOC & " AND  PRC.cd_tp_1_item_no IN (SELECT item_no from imitmidx_Sql where User_def_fld_1 = (Select user_def_fld_1 from imitmidx_Sql WHERE item_no = '" & sqlSAFE(spItemNo) & "'))" & vbCrLf
            ElseIf optAllItems.Checked = True Then
                'Not item criteria
            End If
            sqlLOC = sqlLOC & "Order by PRC.end_dt desc, PRC.start_dt desc" & vbCrLf


            Dim dt As DataTable
            Dim db As New cDBA

            dt = db.DataTable(sqlLOC)
            gridSpecialPricesHISTORY.DataSource = dt

        Catch er As Exception
            MsgBox("Error In " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdViewSpecialPricesHistory_Click(sender As Object, e As EventArgs) Handles cmdViewSpecialPricesHistory.Click
        Me.Height = gridSpecialPricesHISTORY.Top + gridSpecialPricesHISTORY.Height + 40

        gridSpecialPricesHISTORY.ClearSelection()

        Call gridSpecialPricesHISTORY_Format()
        cmdViewSpecialPricesHistory.Text = "Refresh History"
    End Sub

    Private Sub cmdRefreshHistory_Click(sender As Object, e As EventArgs)
        Call gridSpecialPricesHISTORY_Format()
    End Sub


    Private Sub gridSpecialPrices_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridSpecialPrices.CellContentClick
        '  Dim x As Integer

        '    MsgBox(sender.name)
        '    MsgBox(sender.columns)

        '    x = 1
    End Sub

    Private Sub gridPriceCodes_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)

    End Sub

#End Region '"################## END SPECIAL PRICES HISTORY GRID ################################# "

#Region "################## CODE NOT USED ################################# "
    'Public Sub GridPriceCodes_Format()

    '    Try
    '        '================================
    '        'SET GRID LINE COLOURS AND PROPERTIES
    '        '================================
    '        With Me.gridPriceCodes
    '            .RowsDefaultCellStyle.BackColor = Color.White
    '            .AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue    'LightCyan
    '            .ReadOnly = True
    '            .AllowUserToAddRows = False
    '            .AllowUserToDeleteRows = False
    '            .AllowUserToOrderColumns = True

    '        End With
    '        '================================
    '        'CREATE COLUMN WIDTHS AND HEADINGS
    '        '================================
    '        With gridPriceCodes.Columns
    '            .Add(DGVTextBoxColumn("FirstCol", "FirstCol", 30))
    '            .Add(DGVTextBoxColumn("CodeType", "CodeType", 60))
    '            .Add(DGVTextBoxColumn("Currency", "Currency", 60))
    '            .Add(DGVTextBoxColumn("CusNo", "CusNo", 100))
    '            .Add(DGVTextBoxColumn("CusType", "CusType", 50))
    '            .Add(DGVTextBoxColumn("ItemNo", "ItemNo", 140))
    '            .Add(DGVTextBoxColumn("ProdCat", "ProdCat", 65))
    '            .Add(DGVTextBoxColumn("StartDt", "StartDt", 70))
    '            .Add(DGVTextBoxColumn("EndDt", "EndDt", 70))
    '            .Add(DGVTextBoxColumn("PriceBasis", "PriceBasis", 65))
    '            .Add(DGVTextBoxColumn("Qty1", "Qty1", 45))
    '            .Add(DGVTextBoxColumn("Price1", "Price1", 65))
    '            .Add(DGVTextBoxColumn("Qty2", "Qty2", 45))
    '            .Add(DGVTextBoxColumn("Price2", "Price2", 65))
    '            .Add(DGVTextBoxColumn("Qty3", "Qty3", 45))
    '            .Add(DGVTextBoxColumn("Price3", "Price3", 65))
    '            .Add(DGVTextBoxColumn("Qty4", "Qty4", 45))
    '            .Add(DGVTextBoxColumn("Price4", "Price4", 65))
    '            .Add(DGVTextBoxColumn("Qty5", "Qty5", 45))
    '            .Add(DGVTextBoxColumn("Price5", "Price5", 65))
    '            .Add(DGVTextBoxColumn("Qty6", "Qty6", 45))
    '            .Add(DGVTextBoxColumn("Price6", "Price6", 65))
    '            .Add(DGVTextBoxColumn("Qty7", "Qty7", 45))
    '            .Add(DGVTextBoxColumn("Price7", "Price7", 65))
    '            .Add(DGVTextBoxColumn("Qty8", "Qty8", 45))
    '            .Add(DGVTextBoxColumn("Price8", "Price8", 65))
    '            .Add(DGVTextBoxColumn("Qty9", "Qty9", 45))
    '            .Add(DGVTextBoxColumn("Price9", "Price9", 65))
    '            .Add(DGVTextBoxColumn("Qty10", "Qty10", 45))
    '            .Add(DGVTextBoxColumn("Price10", "Price10", 65))
    '            .Add(DGVTextBoxColumn("Extra_14", "Extra_14", 65))
    '            .Add(DGVTextBoxColumn("LastCol", "LastCol", 30))
    '        End With

    '        gridPriceCodes.Columns(PriceCodesEnum.FirstCol).Visible = False
    '        gridPriceCodes.Columns(PriceCodesEnum.CodeType).Visible = False
    '        'gridPriceCodes.Columns(PriceCodesEnum.Currency).Visible = False
    '        gridPriceCodes.Columns(PriceCodesEnum.CusNo).Visible = False
    '        gridPriceCodes.Columns(PriceCodesEnum.ItemNo).Visible = False
    '        gridPriceCodes.Columns(PriceCodesEnum.CusType).Visible = False
    '        gridPriceCodes.Columns(PriceCodesEnum.ProdCat).Visible = False
    '        gridPriceCodes.Columns(PriceCodesEnum.Extra_14).Visible = False
    '        gridPriceCodes.Columns(PriceCodesEnum.LastCol).Visible = False

    '        Call GridPriceCodes_LoadValues()


    '        '================================
    '        'CREATE style for Date
    '        '================================
    '        Dim oCellStyle_DATE As New DataGridViewCellStyle()
    '        oCellStyle_DATE = New DataGridViewCellStyle()
    '        oCellStyle_DATE.Format = "mm/dd/yyyy"
    '        oCellStyle_DATE.Alignment = DataGridViewContentAlignment.MiddleLeft

    '        gridPriceCodes.Columns(PriceCodesEnum.StartDt).DefaultCellStyle = oCellStyle_DATE
    '        gridPriceCodes.Columns(PriceCodesEnum.EndDt).DefaultCellStyle = oCellStyle_DATE

    '        '================================
    '        'SET FROZEN COLUMNS
    '        '================================
    '        'Freezing locks all the columns to the left of the column chosen.
    '        gridPriceCodes.Columns(PriceCodesEnum.FirstCol).Frozen = True
    '        'gridPriceCodes.Columns(PriceCodesEnum.OEI_Ord_No).Frozen = True

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Public Sub GridSpecialPrices_LoadValues()

    '    Try
    '        '================================
    '        'SET GRID LINE COLOURS AND PROPERTIES
    '        '================================
    '        With Me.gridSpecialPrices
    '            .RowsDefaultCellStyle.BackColor = Color.White
    '            .AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue    'LightCyan
    '            .ReadOnly = True
    '            .AllowUserToAddRows = False
    '            .AllowUserToDeleteRows = False
    '            .AllowUserToOrderColumns = True
    '        End With
    '        '================================
    '        'CREATE COLUMN WIDTHS AND HEADINGS
    '        '================================
    '        With gridSpecialPrices.Columns
    '            .Add(DGVTextBoxColumn("FirstCol", "FirstCol", 30))
    '            .Add(DGVTextBoxColumn("CodeType", "CodeType", 60))
    '            .Add(DGVTextBoxColumn("Currency", "Currency", 60))
    '            .Add(DGVTextBoxColumn("CusNo", "CusNo", 100))
    '            .Add(DGVTextBoxColumn("CusType", "CusType", 50))
    '            .Add(DGVTextBoxColumn("ItemNo", "ItemNo", 140))
    '            .Add(DGVTextBoxColumn("ProdCat", "ProdCat", 65))
    '            .Add(DGVTextBoxColumn("StartDt", "StartDt", 70))
    '            .Add(DGVTextBoxColumn("EndDt", "EndDt", 70))
    '            .Add(DGVTextBoxColumn("PriceBasis", "PriceBasis", 65))
    '            .Add(DGVTextBoxColumn("Qty1", "Qty1", 45))
    '            .Add(DGVTextBoxColumn("Price1", "Price1", 65))
    '            .Add(DGVTextBoxColumn("Qty2", "Qty2", 45))
    '            .Add(DGVTextBoxColumn("Price2", "Price2", 65))
    '            .Add(DGVTextBoxColumn("Qty3", "Qty3", 45))
    '            .Add(DGVTextBoxColumn("Price3", "Price3", 65))
    '            .Add(DGVTextBoxColumn("Qty4", "Qty4", 45))
    '            .Add(DGVTextBoxColumn("Price4", "Price4", 65))
    '            .Add(DGVTextBoxColumn("Qty5", "Qty5", 45))
    '            .Add(DGVTextBoxColumn("Price5", "Price5", 65))
    '            .Add(DGVTextBoxColumn("Qty6", "Qty6", 45))
    '            .Add(DGVTextBoxColumn("Price6", "Price6", 65))
    '            .Add(DGVTextBoxColumn("Qty7", "Qty7", 45))
    '            .Add(DGVTextBoxColumn("Price7", "Price7", 65))
    '            .Add(DGVTextBoxColumn("Qty8", "Qty8", 45))
    '            .Add(DGVTextBoxColumn("Price8", "Price8", 65))
    '            .Add(DGVTextBoxColumn("Qty9", "Qty9", 45))
    '            .Add(DGVTextBoxColumn("Price9", "Price9", 65))
    '            .Add(DGVTextBoxColumn("Qty10", "Qty10", 45))
    '            .Add(DGVTextBoxColumn("Price10", "Price10", 65))
    '            .Add(DGVTextBoxColumn("Extra_14", "Extra_14", 65))
    '            .Add(DGVTextBoxColumn("LastCol", "LastCol", 30))
    '        End With

    '        With gridSpecialPrices
    '            .Columns(PriceCodesEnum.FirstCol).Visible = False
    '            .Columns(PriceCodesEnum.CodeType).Visible = False
    '            .Columns(PriceCodesEnum.CusNo).Visible = False
    '            .Columns(PriceCodesEnum.ItemNo).Visible = False
    '            '.Columns(PriceCodesEnum.Currency).Visible = False
    '            .Columns(PriceCodesEnum.CusType).Visible = False
    '            .Columns(PriceCodesEnum.ProdCat).Visible = False
    '            .Columns(PriceCodesEnum.Extra_14).Visible = False
    '            .Columns(PriceCodesEnum.LastCol).Visible = False
    '        End With
    '        Call GridSpecialPrices_LoadValues()


    '        '================================
    '        'CREATE style for Date
    '        '================================
    '        Dim oCellStyle_DATE As New DataGridViewCellStyle()
    '        oCellStyle_DATE = New DataGridViewCellStyle()
    '        oCellStyle_DATE.Format = "mm/dd/yyyy"
    '        oCellStyle_DATE.Alignment = DataGridViewContentAlignment.MiddleLeft

    '        gridSpecialPrices.Columns(PriceCodesEnum.StartDt).DefaultCellStyle = oCellStyle_DATE
    '        gridSpecialPrices.Columns(PriceCodesEnum.EndDt).DefaultCellStyle = oCellStyle_DATE

    '        '================================
    '        'CREATE style for 0 decimals (qty)
    '        '================================
    '        Dim oCellStyle_N0 As New DataGridViewCellStyle()
    '        oCellStyle_N0 = New DataGridViewCellStyle()
    '        oCellStyle_N0.Format = "N0"
    '        oCellStyle_DATE.Alignment = DataGridViewContentAlignment.MiddleRight

    '        With gridSpecialPrices
    '            .Columns(PriceCodesEnum.Qty1).DefaultCellStyle = oCellStyle_N0
    '            .Columns(PriceCodesEnum.Qty2).DefaultCellStyle = oCellStyle_N0
    '            .Columns(PriceCodesEnum.Qty3).DefaultCellStyle = oCellStyle_N0
    '            .Columns(PriceCodesEnum.Qty4).DefaultCellStyle = oCellStyle_N0
    '            .Columns(PriceCodesEnum.Qty5).DefaultCellStyle = oCellStyle_N0
    '            .Columns(PriceCodesEnum.Qty6).DefaultCellStyle = oCellStyle_N0
    '            .Columns(PriceCodesEnum.Qty7).DefaultCellStyle = oCellStyle_N0
    '            .Columns(PriceCodesEnum.Qty8).DefaultCellStyle = oCellStyle_N0
    '            .Columns(PriceCodesEnum.Qty9).DefaultCellStyle = oCellStyle_N0
    '            .Columns(PriceCodesEnum.Qty10).DefaultCellStyle = oCellStyle_N0
    '        End With

    '        '================================
    '        'CREATE style for 4 decimals (Price)
    '        '================================
    '        Dim oCellStyle_N4 As New DataGridViewCellStyle()
    '        oCellStyle_N4 = New DataGridViewCellStyle()
    '        oCellStyle_N4.Format = "N4"
    '        oCellStyle_DATE.Alignment = DataGridViewContentAlignment.MiddleRight

    '        With gridSpecialPrices
    '            .Columns(PriceCodesEnum.Price1).DefaultCellStyle = oCellStyle_N4
    '            .Columns(PriceCodesEnum.Price2).DefaultCellStyle = oCellStyle_N4
    '            .Columns(PriceCodesEnum.Price3).DefaultCellStyle = oCellStyle_N4
    '            .Columns(PriceCodesEnum.Price4).DefaultCellStyle = oCellStyle_N4
    '            .Columns(PriceCodesEnum.Price5).DefaultCellStyle = oCellStyle_N4
    '            .Columns(PriceCodesEnum.Price6).DefaultCellStyle = oCellStyle_N4
    '            .Columns(PriceCodesEnum.Price7).DefaultCellStyle = oCellStyle_N4
    '            .Columns(PriceCodesEnum.Price8).DefaultCellStyle = oCellStyle_N4
    '            .Columns(PriceCodesEnum.Price9).DefaultCellStyle = oCellStyle_N4
    '            .Columns(PriceCodesEnum.Price10).DefaultCellStyle = oCellStyle_N4
    '        End With
    '        '================================
    '        'SET FROZEN COLUMNS
    '        '================================
    '        'Freezing locks all the columns to the left of the column chosen.
    '        gridSpecialPrices.Columns(PriceCodesEnum.FirstCol).Frozen = True
    '        'gridPriceCodes.Columns(PriceCodesEnum.OEI_Ord_No).Frozen = True

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub
    'Public Sub LoadReservedItemsGrid()

    '    Try

    '        With dgvReservedItems.Columns
    '            .Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
    '            .Add(DGVTextBoxColumn("OEI_Ord_No", "OEI No", 100))
    '            .Add(DGVCalendarColumn("Ord_Dt", "Date", 100))
    '            .Add(DGVTextBoxColumn("Item_No", "Item No", 150))
    '            .Add(DGVTextBoxColumn("Qty_Ordered", "Qty Ordered", 100))
    '            .Add(DGVTextBoxColumn("LastCol", "LastCol", 0))

    '        End With

    '        dgvReservedItems.Columns(ReservedItems.FirstCol).Visible = False
    '        dgvReservedItems.Columns(ReservedItems.LastCol).Visible = False

    '        'Call LoadReservedItemsValues()

    '        Dim oCellStyle_DATE As New DataGridViewCellStyle()

    '        oCellStyle_DATE = New DataGridViewCellStyle()
    '        oCellStyle_DATE.Format = "mm/dd/yyyy"
    '        oCellStyle_DATE.Alignment = DataGridViewContentAlignment.MiddleLeft

    '        dgvReservedItems.Columns(ReservedItems.Ord_Dt).DefaultCellStyle = oCellStyle_DATE

    '        dgvReservedItems.Columns(ReservedItems.FirstCol).Frozen = True
    '        dgvReservedItems.Columns(ReservedItems.OEI_Ord_No).Frozen = True

    '    Catch er As Exception
    '        MsgBox("Error In " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Public Sub LoadReservedItemsValues()


    '    Try
    '        Dim strSql As String = "" &
    '        "Select		'' AS FirstCol, O.Oei_Ord_No, O.Ord_Dt, L.Item_No, " &
    '        "           L.Qty_Ordered, '' AS LastCol " &
    '        "FROM 		OEI_Ordhdr O WITH (Nolock) " &
    '        "INNER JOIN OEI_OrdLin L WITH (Nolock) ON O.Ord_Guid = L.Ord_Guid " &
    '        "WHERE 		ISNULL(O.Ord_No, '') = '' " &
    '        "ORDER BY 	L.Item_No, L.Qty_Ordered DESC "

    '        Dim dt As DataTable
    '        Dim db As New cDBA

    '        dt = db.DataTable(strSql)
    '        dgvReservedItems.DataSource = dt

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

#End Region '"################## END CODE NOT USED ################################# "

End Class