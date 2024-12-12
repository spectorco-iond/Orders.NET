Module mod_HV_GRID
    Private bInitializationComplete As Boolean   'True when initialization has been ran - So styles have been created.

    Public Const cVisible_TRUE = True
    Public Const cVisible_FALSE = False
    Public Const cLocked_TRUE = True
    Public Const cLocked_FALSE = False



    Public Enum GridColTypeEnum
        ButtonColumn = 1
        CalendarColumn
        CheckBoxColumn
        ComboBoxColumn
        ImageColumn
        TextBoxColumn
    End Enum

    Public Enum GridFormatEnum
        Text
        N0 = 1
        N1
        N2
        N3
        N4
        N5
        N6
        ShortDate
    End Enum

    Private Style_ShortDate As New DataGridViewCellStyle()
    Private Style_N0 As New DataGridViewCellStyle()


    Public Sub Initialize_mod_HV_GRID()
        ''We want this to run through one time so that it creates the styles used in this mod

        ''If it has already been initialized then do nothing.
        'If bInitializationComplete = True Then Exit Sub

        ''================================
        ''CREATE style for Date
        ''================================
        ''Dim Style_ShortDate As New DataGridViewCellStyle()
        'Style_ShortDate = New DataGridViewCellStyle()
        'Style_ShortDate.Format = "mm/dd/yyyy"
        ''Style_ShortDate.Alignment = DataGridViewContentAlignment.MiddleLeft

        ''================================
        ''CREATE style for 0 decimals (qty)
        ''================================
        ''Dim Style_N0 As New DataGridViewCellStyle()
        'Style_N0 = New DataGridViewCellStyle()
        'Style_N0.Format = "N0"
        ''Style_N0.Alignment = DataGridViewContentAlignment.MiddleRight
        ''================================
        ''CREATE style for 4 decimals (qty)
        ''================================
        'Dim Style_N4 As New DataGridViewCellStyle()
        'Style_N4 = New DataGridViewCellStyle()
        'Style_N4.Format = "N4"
        ''Style_N0.Alignment = DataGridViewContentAlignment.MiddleRight
        'bInitializationComplete = True
    End Sub

    Public Function AddGridColumn(FieldType As GridColTypeEnum, FieldName As String, Title As String, Visible As Boolean, Locked As Boolean, Optional Width As Long = 100, Optional Alignment As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleLeft, Optional Format As GridFormatEnum = 0) As Object
        Dim ObjCOL As Object

        Try
            'This will only run one time - sets up the formats
            Call Initialize_mod_HV_GRID()

            '================================
            'Create the column ObjCOLect
            '================================
            Select Case FieldType
                Case GridColTypeEnum.ButtonColumn
                    ObjCOL = New DataGridViewButtonColumn
                Case GridColTypeEnum.CalendarColumn
                    ObjCOL = New CalendarColumn
                    ObjCOL.DefaultCellStyle.Format = "mm/dd/yyyy"
                Case GridColTypeEnum.CheckBoxColumn
                    ObjCOL = New DataGridViewCheckBoxColumn
                    'Set true/false values
                    ObjCOL.FalseValue = 0 '"False"
                    ObjCOL.TrueValue = 1 '"True"
                    ObjCOL.IndeterminateValue = 0 '"False"
                Case GridColTypeEnum.ComboBoxColumn
                    ObjCOL = New DataGridViewComboBoxColumn
                Case GridColTypeEnum.ImageColumn
                    ObjCOL = New DataGridViewImageColumn
                Case GridColTypeEnum.TextBoxColumn
                    ObjCOL = New DataGridViewTextBoxColumn
                Case Else
                    MsgBox("Invalid fieldType passed to AddGridColumn", vbOKOnly, "INVALID CODE")
                    Return vbNull
                    Exit Function
            End Select


            '================================
            'SET TEXT FORMAT
            '================================
            Select Case Format
                Case GridFormatEnum.Text
    '-- use normal formatting
                Case GridFormatEnum.N0
                    ObjCOL.DefaultCellStyle.format = "N0"
                Case GridFormatEnum.N1
                    ObjCOL.DefaultCellStyle.format = "N1"
                Case GridFormatEnum.N2
                    ObjCOL.DefaultCellStyle.format = "N2"
                Case GridFormatEnum.N3
                    ObjCOL.DefaultCellStyle.format = "N3"
                Case GridFormatEnum.N4
                    ObjCOL.DefaultCellStyle.format = "N4"
                Case GridFormatEnum.N5
                    ObjCOL.DefaultCellStyle.format = "N5"
                Case GridFormatEnum.N6
                    ObjCOL.DefaultCellStyle.format = "N6"
                Case GridFormatEnum.ShortDate
                    ObjCOL.DefaultCellStyle.Format = "mm/dd/yyyy"
            End Select

            '================================
            'SET TEXT ALIGNMENT
            '================================
            ObjCOL.DefaultCellStyle.Alignment = Alignment

            'With gridSpecialPrices
            '    .Columns(PriceCodesEnum.Qty1).DefaultCellStyle = oCellStyle_N0
            '    .Columns(PriceCodesEnum.Qty2).DefaultCellStyle = oCellStyle_N0
            '    .Columns(PriceCodesEnum.Qty3).DefaultCellStyle = oCellStyle_N0
            '    .Columns(PriceCodesEnum.Qty4).DefaultCellStyle = oCellStyle_N0
            '================================
            'SET OTHER VALUES
            '================================
            ObjCOL.HeaderText = Title
            ObjCOL.DataPropertyName = FieldName
            ObjCOL.Name = FieldName

            ObjCOL.Visible = Visible
            ObjCOL.readonly = Locked
            ObjCOL.Width = Width

            'ObjCOL.Alignment = DataGridViewContentAlignment.MiddleRight
            ' If Format <> 0 Then ObjCOL.DefaultCellStyle = Format



            '================================
            '================================
            Return ObjCOL
            '================================
            '================================
            '================================
            '================================
            '================================
            '================================
            '================================
            '================================
            '================================

        Catch er As Exception
            MsgBox("Error in mod_HV_Grid." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            Return vbNull
        End Try


    End Function

    'Public Function hvButtonColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewButtonColumn

    '    DGVButtonColumn = New DataGridViewButtonColumn

    '    DGVButtonColumn.HeaderText = pstrHeaderText
    '    DGVButtonColumn.DataPropertyName = pstrName
    '    DGVButtonColumn.Name = pstrName

    '    If plWidth <> 0 Then DGVButtonColumn.Width = plWidth

    'End Function

    'Public Function hvCalendarColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As CalendarColumn

    '    DGVCalendarColumn = New CalendarColumn

    '    DGVCalendarColumn.HeaderText = pstrHeaderText
    '    DGVCalendarColumn.DataPropertyName = pstrName
    '    DGVCalendarColumn.Name = pstrName
    '    'DGVCalendarColumn.DefaultCellStyle = "mm/dd/yyyy"
    '    DGVCalendarColumn.DefaultCellStyle.Format = "mm/dd/yyyy"

    '    If plWidth <> 0 Then DGVCalendarColumn.Width = plWidth

    'End Function

    'Public Function DGVCheckBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewCheckBoxColumn

    '    DGVCheckBoxColumn = New DataGridViewCheckBoxColumn

    '    DGVCheckBoxColumn.HeaderText = pstrHeaderText
    '    DGVCheckBoxColumn.DataPropertyName = pstrName
    '    DGVCheckBoxColumn.Name = pstrName

    '    DGVCheckBoxColumn.FalseValue = 0 '"False"
    '    DGVCheckBoxColumn.TrueValue = 1 '"True"
    '    DGVCheckBoxColumn.IndeterminateValue = 0 '"False"

    '    If plWidth <> 0 Then DGVCheckBoxColumn.Width = plWidth

    'End Function

    'Public Function DGVComboBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewComboBoxColumn

    '    DGVComboBoxColumn = New DataGridViewComboBoxColumn

    '    DGVComboBoxColumn.HeaderText = pstrHeaderText
    '    DGVComboBoxColumn.DataPropertyName = pstrName
    '    DGVComboBoxColumn.Name = pstrName
    '    'DGVComboBoxColumn.DropDownWidth = 160

    '    If plWidth <> 0 Then DGVComboBoxColumn.Width = plWidth

    'End Function

    'Public Function DGVImageColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, ByVal plWidth As Long) As DataGridViewImageColumn

    '    DGVImageColumn = New DataGridViewImageColumn
    '    DGVImageColumn.Name = pstrName
    '    DGVImageColumn.DataPropertyName = pstrName
    '    DGVImageColumn.HeaderText = pstrHeaderText

    'End Function

    'Public Function DGVTextBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewTextBoxColumn

    '    DGVTextBoxColumn = New DataGridViewTextBoxColumn

    '    DGVTextBoxColumn.HeaderText = pstrHeaderText
    '    DGVTextBoxColumn.DataPropertyName = pstrName
    '    DGVTextBoxColumn.Name = pstrName

    '    If plWidth <> 0 Then DGVTextBoxColumn.Width = plWidth

    'End Function
End Module
