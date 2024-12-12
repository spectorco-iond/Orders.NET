Imports System.IO

Public Class frmMixMatchSelector
    Public mix_item_cd As String
    Public item_loc As String
    Private chosenSKuSet As ArrayList
    Private chosenInventory As Integer = 0
    Const resizeVal = 100

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Enum Columns
        image
        component
        item_no
        loc
    End Enum

    Public Sub GenerateComponentGroups()
        Dim strSql As String
        Dim conn As New cDBA
        Dim dtMixComponents As New DataTable

        strSql = "Select itemNo, loc, desc_en AS cmp_type " _
                & "FROM MWAV_INVENTORY_LEVELS_COMBINED_WITH_DESC " _
                & "WHERE itemNo like '%" & mix_item_cd.Trim() & "%' AND loc = '" & item_loc & "'"


        strSql = "Select itemNo, loc, desc_en AS cmp_type " _
                & "FROM MWAV_INVENTORY_LEVELS_COMBINED_WITH_DESC_ion " _
                & "WHERE itemNo like '%" & mix_item_cd.Trim() & "%' AND loc = '" & item_loc & "'"

        'was modified the original view for
        strSql = "Select ISNULL(itemNo,'') As ItemNo, Ltrim(Rtrim(loc)) as loc,Isnull(desc_en,'') AS cmp_type, isnull(COMPONENT_ORDER_NO,0) As COMPONENT_ORDER_NO " _
                & "FROM dbo.[MWAV_INVENTORY_LEVELS_COMBINED_WITH_DESC_Ion2.07.2019] " _
              & "WHERE isnull(KitItemNo,'') like '%" & mix_item_cd.Trim() & "%' AND loc = '" & Trim(item_loc) & "'    order by COMPONENT_ORDER_NO"


        Try
            dtMixComponents = conn.DataTable(strSql)
            Call populateGrid(dtMixComponents)
        Catch er As Exception
            MsgBox("Error in frmMixMatchSelector->GenerateComponentGroups: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    'Function BlankImage() As Image
    '    Try
    '        Dim oBM As New Bitmap(1, 1)
    '        oBM.SetPixel(0, 0, Color.Transparent)
    '        Return oBM
    '    Catch ex As Exception
    '        Return Nothing
    '    End Try

    Private Sub populateGrid(componentTable As DataTable)
        Dim componentType As String = ""
        Dim componentOrder As Int32 = -1

        Dim currentItemSkuComboCell As DataGridViewComboBoxCell = Nothing

        dgvComponents.DataSource = Nothing 'reset grid

        Try

            For Each component As DataRow In componentTable.Rows
                '++ID commented 1.07.2019
                If componentType <> component.Item("cmp_type").ToString And componentOrder <> component.Item("COMPONENT_ORDER_NO") Then
                    componentType = component.Item("cmp_type").ToString
                    componentOrder = component.Item("COMPONENT_ORDER_NO")

                    Dim row As String() = {Nothing, componentType, Nothing, item_loc}

                    dgvComponents.Rows.Add(row)
                    dgvComponents.Rows(dgvComponents.Rows.Count - 1).Selected = True
                    dgvComponents.Rows(dgvComponents.Rows.Count - 1).Height = resizeVal

                    If Not currentItemSkuComboCell Is Nothing AndAlso currentItemSkuComboCell.Items.Count = 1 Then
                        currentItemSkuComboCell.Value = currentItemSkuComboCell.Items(0).ToString
                    End If

                    currentItemSkuComboCell = Nothing
                    currentItemSkuComboCell = dgvComponents.Rows(dgvComponents.Rows.Count - 1).Cells(Columns.item_no)

                    'add blank image
                    Dim oBM As New Bitmap(1, 1)
                    oBM.SetPixel(0, 0, Color.Transparent)
                    dgvComponents.Rows(dgvComponents.Rows.Count - 1).Cells(Columns.image).Value = oBM


                ElseIf componentType = component.Item("cmp_type").ToString And componentOrder <> component.Item("COMPONENT_ORDER_NO") Then

                    componentType = component.Item("cmp_type").ToString
                    componentOrder = component.Item("COMPONENT_ORDER_NO")

                    Dim row As String() = {Nothing, componentType, Nothing, item_loc}

                    dgvComponents.Rows.Add(row)
                    dgvComponents.Rows(dgvComponents.Rows.Count - 1).Selected = True
                    dgvComponents.Rows(dgvComponents.Rows.Count - 1).Height = resizeVal

                    If Not currentItemSkuComboCell Is Nothing AndAlso currentItemSkuComboCell.Items.Count = 1 Then
                        currentItemSkuComboCell.Value = currentItemSkuComboCell.Items(0).ToString
                    End If

                    currentItemSkuComboCell = Nothing
                    currentItemSkuComboCell = dgvComponents.Rows(dgvComponents.Rows.Count - 1).Cells(Columns.item_no)

                    'add blank image
                    Dim oBM As New Bitmap(1, 1)
                    oBM.SetPixel(0, 0, Color.Transparent)
                    dgvComponents.Rows(dgvComponents.Rows.Count - 1).Cells(Columns.image).Value = oBM

                End If
                currentItemSkuComboCell.Items.Add(component.Item("itemNo").ToString)




            Next

            'handle final set if applicable
            If Not currentItemSkuComboCell Is Nothing AndAlso currentItemSkuComboCell.Items.Count = 1 Then
                currentItemSkuComboCell.Value = currentItemSkuComboCell.Items(0).ToString
            End If

        Catch er As Exception
            MsgBox("Error in frmMixMatchSelector->populateGrid: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    ' ---------------------------------------

    Private Sub populateGrid1(componentTable As DataTable)
        Dim componentType As String = ""
        ' Dim componentOrder As Int32 = -1

        Dim currentItemSkuComboCell As DataGridViewComboBoxCell = Nothing

        dgvComponents.DataSource = Nothing 'reset grid

        Try

            For Each component As DataRow In componentTable.Rows
                '++ID commented 1.07.2019
                If componentType <> component.Item("cmp_type").ToString Then
                    componentType = component.Item("cmp_type").ToString


                    Dim row As String() = {Nothing, componentType, Nothing, item_loc}

                    dgvComponents.Rows.Add(row)
                    dgvComponents.Rows(dgvComponents.Rows.Count - 1).Selected = True
                    dgvComponents.Rows(dgvComponents.Rows.Count - 1).Height = resizeVal

                    If Not currentItemSkuComboCell Is Nothing AndAlso currentItemSkuComboCell.Items.Count = 1 Then
                        currentItemSkuComboCell.Value = currentItemSkuComboCell.Items(0).ToString
                    End If

                    currentItemSkuComboCell = Nothing
                    currentItemSkuComboCell = dgvComponents.Rows(dgvComponents.Rows.Count - 1).Cells(Columns.item_no)

                    'add blank image
                    Dim oBM As New Bitmap(1, 1)
                    oBM.SetPixel(0, 0, Color.Transparent)
                    dgvComponents.Rows(dgvComponents.Rows.Count - 1).Cells(Columns.image).Value = oBM



                End If
                currentItemSkuComboCell.Items.Add(component.Item("itemNo").ToString)




            Next

            'handle final set if applicable
            If Not currentItemSkuComboCell Is Nothing AndAlso currentItemSkuComboCell.Items.Count = 1 Then
                currentItemSkuComboCell.Value = currentItemSkuComboCell.Items(0).ToString
            End If

        Catch er As Exception
            MsgBox("Error in frmMixMatchSelector->populateGrid: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    '----------------------------------------
    Private Sub txtQty_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtQty.TextChanged
        If Not IsNumeric(txtQty.Text) OrElse txtQty.Text <= 0 Then
            MsgBox("Invalid quantity entered")
            txtQty.Text = 1
        End If
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        'reset values and close
        chosenSKuSet = Nothing
        chosenInventory = 0
        Me.Close()
    End Sub

    Private Sub cmdSubmit_Click(sender As System.Object, e As System.EventArgs) Handles cmdSubmit.Click
        chosenSKuSet = New ArrayList
        Dim isValid As Boolean = True

        For Each row As DataGridViewRow In dgvComponents.Rows
            If Not row.Cells.Item(Columns.item_no).Value Is Nothing AndAlso row.Cells.Item(Columns.item_no).Value.ToString <> "" AndAlso IsNumeric(txtQty.Text) AndAlso txtQty.Text > 0 Then
                chosenSKuSet.Add(row.Cells.Item(Columns.item_no).Value.ToString)
            Else
                MsgBox("Not all components have been selected or invalid quantity")
                isValid = False
                Exit For
            End If
        Next
        chosenInventory = txtQty.Text.ToString
        If isValid = True Then
            Me.Close()
        End If
    End Sub

    Public Function getChosenSkuSet() As ArrayList
        Return chosenSKuSet
    End Function

    Public Function getChosenInventory() As Integer
        Return chosenInventory
    End Function

    Public Function getLocation() As String
        Return item_loc
    End Function


    Private Sub dgvComponents_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvComponents.CellEndEdit
        Try
            If dgvComponents.CurrentCell.ColumnIndex = Columns.item_no Then
                Call changeImage(dgvComponents.CurrentCell.Value.ToString)
            End If
        Catch er As Exception
            MsgBox("Error in frmMixMatchSelector->CellValidating: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub changeImage(sku As String)
        Dim strSql As String
        Dim conn As New cDBA
        Dim dtPictures As New DataTable


        strSql = "Select picture,ISNULL(OEI_W_Pixels, 400) as OEI_W_Pixels, ISNULL(OEI_H_Pixels, 400) as OEI_H_Pixels, ISNULL(Rotate, '') as Rotate " _
                & "FROM item_pictures " _
                & "WHERE item_no = '" & sku & "'"

        Try
            dtPictures = conn.DataTable(strSql)
            If dtPictures.Rows.Count > 0 Then
                '  Dim ColImage As New DataGridViewImageColumn
                '  Dim Img As New DataGridViewImageCell
                Dim oImage As Image
                Dim data As Byte() = dtPictures.Rows(0).Item("picture")
                Dim ms As MemoryStream

                'Set image value
                ms = New MemoryStream(data)
                oImage = Image.FromStream(ms)
                '  oImage.HorizontalResolution = 400


                'Add the image cell to a row
                dgvComponents.Rows(dgvComponents.CurrentRow.Index).Cells(Columns.image).Value = New Bitmap(oImage, New Size(resizeVal, resizeVal))
                dgvComponents.Rows(dgvComponents.CurrentRow.Index).Height = resizeVal
                ' dgvComponents.Rows(dgvComponents.CurrentRow.Index).Cells(Columns.image).Size.Width.MaxValue = 200
                'dgvComponents.Rows(dgvComponents.CurrentRow.Index).Cells(Columns.image).Size.Height.MaxValue = 200

            End If
        Catch er As Exception
            MsgBox("Error in frmMixMatchSelector->GenerateComponentGroups: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

End Class