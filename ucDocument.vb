Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Text
'Imports Microsoft.Office.Interop.Outlook

Friend Class ucDocument
    Inherits System.Windows.Forms.UserControl

#Region "Private variables ################################################"

    Private m_Document As New CDocument

    Private m_Ord_Guid As String
    Private m_Item_Guid As String
    Private m_Cus_No As String

    Private blnLoading As Boolean = False
    Private dgvRow As DataGridViewRow
    Private intListDropIndex As Integer
    Private blnRenameFile As Boolean = False
    Private m_DateButton As Button

    Private m_Path As String = "c:\ExactTemp\"

    'Private objOL As New Microsoft.Office.Interop.Outlook.Application

    Private Enum Columns

        FirstCol = 0
        DocID
        HeaderID
        Ord_Guid
        Item_Guid
        DocTypeID
        DocType
        DocFile
        Ord_No
        OE_Po_No
        DocName
        DocDescription

    End Enum

    Private Enum EmailColumns

        FirstCol
        ID
        Email_From
        Email_To
        Email_To_Name
        Email_CC
        Subject
        Body
        Ord_Guid
        Ord_No
        CreateTS
        SendTS
        UserID

    End Enum

#End Region

#Region "Private DateTimePicker control functions #########################"

    ' Date buttons click - open DateControl for element
    Private Sub Element_DateButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDateTP.Click, cmdShipDateTP.Click

        Try
            m_DateButton = New Button
            m_DateButton = DirectCast(sender, Button)

            mcCalendar.Top = m_DateButton.Top + pnlHistory.Top + m_DateButton.Height
            mcCalendar.Left = m_DateButton.Left + pnlHistory.Left - mcCalendar.Width + m_DateButton.Width
            mcCalendar.Visible = True
            mcCalendar.SetDate(Date.Now)

            Select Case m_DateButton.Name
                Case "cmdDateTP"
                    If IsDate(txtOrd_Dt.Text) Then mcCalendar.SetDate(txtOrd_Dt.Text)

                Case "cmdShipDateTP"
                    If IsDate(txtOrd_Dt_Shipped.Text) Then mcCalendar.SetDate(txtOrd_Dt_Shipped.Text)

            End Select
            mcCalendar.Select()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Returns the search button control associated with a textbox or a combobox
    Private Function GetDateControlByControlName(ByVal txtElement As String) As Button

        GetDateControlByControlName = New Button

        Try
            Select Case txtElement

                Case "txtOrd_Dt"
                    GetDateControlByControlName = cmdDateTP

                Case "txtOrd_Dt_Shipped"
                    GetDateControlByControlName = cmdShipDateTP

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Puts the selected date from the calendar info the linked textbox.
    Private Sub mcCalendar_DateSelected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles mcCalendar.DateSelected

        Try
            'If mcCalendar.SelectionRange.Start.Year < (Now.Year - 1) Then
            '    Throw New OEException(OEError.Year_Cannot_Be_LT_Previous_Year)
            'End If

            'If mcCalendar.SelectionRange.Start.Year > (Now.Year + 1) Then
            '    Throw New OEException(OEError.Year_Cannot_Be_GT_Follow_Year)
            'End If

            Select Case m_DateButton.Name
                Case "cmdDateTP"
                    txtOrd_Dt.Text = mcCalendar.SelectionRange.Start
                    txtOrd_Dt.Focus()

                Case "cmdShipDateTP"
                    txtOrd_Dt_Shipped.Text = mcCalendar.SelectionRange.Start
                    txtOrd_Dt_Shipped.Focus()
            End Select

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        Finally
            mcCalendar.Visible = False
            m_DateButton = Nothing
        End Try

    End Sub

    ' When date time picker gets Escape button, it closes.
    Private Sub mcCalendar_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mcCalendar.KeyDown

        Try
            If e.KeyCode = Keys.Escape Then ' e.Control And e.KeyValue = Keys.F Then
                mcCalendar.Visible = False
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private Date elements private events #############################"

    ' Logical sequence get focus to force focus on needed fields
    Private Sub Element_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                    txtOrd_Dt.GotFocus, cmdDateTP.GotFocus, _
                    txtOrd_Dt_Shipped.GotFocus, cmdShipDateTP.GotFocus

        Dim lTabIndex As Integer

        Try
            If TypeOf sender Is xTextBox Then

                Dim txtElement As New xTextBox
                txtElement = DirectCast(sender, xTextBox)
                lTabIndex = txtElement.TabIndex

                Select Case txtElement.Name
                    Case "txtOrd_Dt", "txtOrd_Dt_Shipped"
                        Call oFrmOrder.SetStatusBarPanelText(1, g_Date_Shortcut)
                        'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))
                        'txtElement.BeginInvoke(New myDelegate(

                    Case Else
                        Call oFrmOrder.SetStatusBarPanelText(1, " ")
                        'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))

                End Select

            ElseIf TypeOf sender Is Button Then

                Dim txtElement As New Button
                txtElement = DirectCast(sender, Button)
                lTabIndex = txtElement.TabIndex

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Gets Key Down, to trigger F9
    Private Sub Element_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOrd_Dt.KeyDown, txtOrd_Dt_Shipped.KeyDown

        Dim txtElement As New Object

        Try

            txtElement = New xTextBox
            txtElement = DirectCast(sender, xTextBox)

            Select Case e.KeyValue

                ' Calendar
                Case Keys.F9
                    Dim oButton As Button
                    oButton = GetDateControlByControlName(sender.name)
                    If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Checks if any date is a valid date.
    Private Sub Element_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtOrd_Dt.Validating, txtOrd_Dt_Shipped.Validating

        Dim txtElement As New Object

        Try

            txtElement = New xTextBox
            txtElement = DirectCast(sender, xTextBox)

            If txtElement.text = "" Then Exit Sub

            Try
                If Not IsDate(txtElement.Text) Then
                    e.Cancel = True
                    Throw New OEException(OEError.Invalid_Date_Format)
                End If

            Catch er As OEException
                MsgBox(er.Message)
            End Try


        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private lvTypes events ###########################################"

    Private Sub lvTypes_DragEnter(ByVal sender As Object, ByVal e As  _
                System.Windows.Forms.DragEventArgs) Handles lvTypes.DragEnter

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then 'For email Drag & Drop 
            e.Effect = DragDropEffects.Copy
        End If

    End Sub

    Private Sub lvTypes_DragDrop(ByVal sender As Object, ByVal e As  _
                System.Windows.Forms.DragEventArgs) Handles lvTypes.DragDrop

        Try

            If intListDropIndex = 0 Then

                MsgBox("You cannot drop elements on 'All'. Please select a type.")
                Exit Sub

            End If

            Dim MyFiles() As String

            If e.Data.GetDataPresent(DataFormats.FileDrop) Then

                ' Assign the files to an array.
                MyFiles = e.Data.GetData(DataFormats.FileDrop)

                AddDocumentFiles(intListDropIndex, MyFiles)

            ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then

                Dim objOL As New Microsoft.Office.Interop.Outlook.Application
                Dim objMI As Microsoft.Office.Interop.Outlook.MailItem
                ReDim MyFiles(1)

                MyFiles(0) = ""
                Dim iPos As Integer = 0
                For Each objMI In objOL.ActiveExplorer.Selection()
                    ReDim Preserve MyFiles(iPos)
                    'hardcode a destination path for testing
                    Dim strFile As String = _
                                IO.Path.Combine("c:\ExactTemp", _
                                                (m_oOrder.Ordhead.OEI_Ord_No & "-" & objMI.Subject + ".msg").Replace(":", "").Replace("/", "_").Replace("*", ".").Replace("?", "."))
                    objMI.SaveAs(strFile)
                    MyFiles(iPos) = strFile
                    iPos += 1

                Next

                AddDocumentFiles(intListDropIndex, MyFiles)

            ElseIf (e.Data.GetDataPresent(GetType(CDocument))) Then

                Dim oDocument As New CDocument
                oDocument = DirectCast(e.Data.GetData(GetType(CDocument)), CDocument)

                Call AddHistoryDocument(oDocument, CInt(lvTypes.Items(intListDropIndex).SubItems(2).Text), lvTypes.Items(intListDropIndex).SubItems(1).Text)

                'oDocument.DocType = lvTypes.Items(intListDropIndex).SubItems(1).Text
                'oDocument.DocTypeID = CInt(lvTypes.Items(intListDropIndex).SubItems(2).Text)
                'If oDocument.DocID <> 0 Then
                '    oDocument.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                '    oDocument.Item_Guid = g_oOrdline.Item_Guid
                '    oDocument.Save()
                'Else
                '    oDocument.Save()
                'End If

                'Dim n As Integer = dgvFileList.Rows.Add()
                'With dgvFileList.Rows.Item(n)

                '    .Cells(Columns.DocID).Value = 0
                '    .Cells(Columns.HeaderID).Value = 0
                '    .Cells(Columns.DocDescription).Value = oDocument.DocDescription
                '    .Cells(Columns.DocFile).Value = oDocument.DocFile
                '    .Cells(Columns.DocName).Value = oDocument.DocName
                '    .Cells(Columns.DocType).Value = oDocument.DocType
                '    .Cells(Columns.DocTypeID).Value = oDocument.DocTypeID
                '    .Cells(Columns.Item_Guid).Value = oDocument.Item_Guid
                '    .Cells(Columns.Ord_Guid).Value = oDocument.Ord_Guid

                'End With

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub lvTypes_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvTypes.DragOver

        Dim this_list As ListView = DirectCast(sender, ListView)
        Dim pt As Point = this_list.PointToClient(New Point(e.X, e.Y))

        Dim selection As ListViewItem = this_list.GetItemAt(pt.X, pt.Y)

        If selection IsNot Nothing Then
            intListDropIndex = this_list.GetItemAt(pt.X, pt.Y).Index
            lvTypes.View = View.Details
            'this_list.SelectedItems.Clear()
            this_list.Items(intListDropIndex).Selected = True
            If pt.Y < 23 And intListDropIndex > 0 Then this_list.EnsureVisible(intListDropIndex - 1)
            If pt.Y > this_list.Height - 21 Then
                this_list.EnsureVisible(intListDropIndex + 1)
            End If

            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub lvTypes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvTypes.SelectedIndexChanged

        If lvTypes.SelectedItems.Count <> 0 Then
            For Each item As ListViewItem In lvTypes.SelectedItems
                intListDropIndex = item.Index
            Next item

            Call ChangeGridView(intListDropIndex)
            If chkShowHistory.Checked Then Call ShowHistoryOrders()

        End If

    End Sub

#End Region

#Region "Private events ucDocument ########################################"

    Private Sub ucDocument_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Enter

        lvTypes.AutoScrollOffset = New System.Drawing.Point(20, 20)

    End Sub

#End Region

    ' activate drag and drop copy and move 
#Region "Common Private events dgvHistory #################################"

    Private Sub DataGridView_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) _
    Handles dgvFileList.MouseDown, dgvHistory.MouseDown

        Dim oDataGrid As DataGridView
        oDataGrid = DirectCast(sender, DataGridView)

        If e.Button = MouseButtons.Left Then
            'Dim info As DataGridView.HitTestInfo = dgvFileList.HitTest(e.X, e.Y)
            Dim info As DataGridView.HitTestInfo = oDataGrid.HitTest(e.X, e.Y)
            If info.ColumnIndex <> Columns.DocDescription Then
                If info.RowIndex >= 0 Then
                    'dim View as DataRowView = (DataRowView) dgvFileList.Rows(info.RowIndex).DataBoundItem
                    'dgvRow = dgvFileList.Rows(info.RowIndex) ' .DataBoundItem, DataRowView)
                    'm_Document = New CDocument
                    'm_Document.Load(dgvRow)
                    If Not (dgvRow Is Nothing) Then
                        'Dim oDocument As New CDocument
                        'oDocument.Load(dgvRow)
                        'dgvFileList.DoDragDrop(m_Document, DragDropEffects.Copy)
                        oDataGrid.DoDragDrop(m_Document, DragDropEffects.Copy)
                    End If
                End If
            End If
        End If

    End Sub

#End Region

    ' dgvFileList events activate preview and delete
#Region "Private events dgvFileList #######################################"

    Private Sub dgvFileList_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) _
    Handles dgvFileList.CellBeginEdit

        Select Case e.ColumnIndex
            'Case Columns.DocName
            '    e.Cancel = Not (blnRenameFile)

            Case Columns.DocDescription
                ' Allow editing

            Case Else
                e.Cancel = True

        End Select

    End Sub

    Private Sub dgvFileList_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) _
    Handles dgvFileList.CellValidating

        Select Case e.ColumnIndex
            'Case Columns.DocName

            '    If Trim(e.FormattedValue) <> Trim(m_Document.DocName) Then
            '        m_Document.DocName = e.FormattedValue
            '        m_Document.Save()
            '    End If

            Case Columns.DocDescription

                If Trim(e.FormattedValue) <> Trim(m_Document.DocDescription) Then
                    m_Document.DocDescription = e.FormattedValue
                    m_Document.Save(m_Document.DocID)
                End If

        End Select

    End Sub

    Private Sub dgvFileList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles dgvFileList.KeyDown

        If dgvFileList.Rows.Count = 0 Then Exit Sub

        'Select Case dgvFileList.CurrentCell.ColumnIndex
        '    Case Columns.
        'Select Case e.KeyCode
        'Case Keys.F2
        '    blnRenameFile = True
        '    dgvFileList.CurrentCell = dgvFileList.CurrentRow.Cells(Columns.DocName)
        '    dgvFileList.BeginEdit(True)
        '    blnRenameFile = False
        '    'End Select
        'End Select

    End Sub

    Private Sub dgvFileList_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) _
    Handles dgvFileList.RowEnter, dgvHistory.RowEnter

        Dim oDataGrid As DataGridView
        oDataGrid = DirectCast(sender, DataGridView)

        'If dgvFileList.CurrentRow Is Nothing Then Exit Sub
        If oDataGrid.CurrentRow Is Nothing Then Exit Sub

        'dgvRow = dgvFileList.Rows(dgvFileList.CurrentRow.Index) ' .DataBoundItem, DataRowView)
        'dgvRow = dgvFileList.Rows(e.RowIndex) ' .DataBoundItem, DataRowView)
        dgvRow = oDataGrid.Rows(e.RowIndex) ' .DataBoundItem, DataRowView)
        m_Document = New CDocument
        m_Document.Load(dgvRow)

    End Sub

#End Region

    ' dgvHistory events activate preview 
#Region "Private events dgvHistory ########################################"

    Private Sub dgvHistory_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) _
    Handles dgvHistory.CellBeginEdit

        e.Cancel = True

    End Sub

#End Region

#Region "Private Button events ############################################"

    Private Sub chkShowHistory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowHistory.CheckedChanged

        Call ShowHistoryOrders()

    End Sub

    Private Sub cmdPreviewDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreviewDoc.Click

        Call PreviewDoc()

    End Sub

    Private Sub cmdDeleteDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteDoc.Click

        Call DeleteDoc()

    End Sub

    Private Sub cmdSearchHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearchHistory.Click

        Call ShowHistoryOrders()

    End Sub

#End Region

#Region "Button procedures ################################################"

    Private Sub DeleteDoc()

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If dgvFileList.Rows.Count > 0 Then

                m_Document.Load(dgvFileList.Rows(dgvFileList.CurrentRow.Index))
                m_Document.Delete()

                If m_Document.DeleteStatus = mGlobal.DeleteStatus.Deleted Then
                    dgvFileList.Rows.RemoveAt(dgvFileList.CurrentRow.Index)
                    If dgvFileList.Rows.Count > 0 Then
                        m_Document = New CDocument
                        dgvFileList.CurrentCell = dgvFileList.Rows(0).Cells(Columns.DocName)
                    End If

                    Throw New OEException(OEError.Delete_Successful)
                End If

            End If

        Catch oe_er As OEException
            'If oe_er.ShowMessage Then MsgBox(oe_er.Message)

        End Try

    End Sub

    Private Sub PreviewDoc()

        'Dim mStream As ADODB.Stream
        'Dim strPathAndFile As String

        Try

            If m_Document Is Nothing Then Exit Sub

            'If dgvFileList.Rows.Count <> 0 Then
            '    m_Document.Load(dgvFileList.Rows(dgvFileList.CurrentRow.Index))
            'End If

            If dgvFileList.CurrentRow Is Nothing Then
                If dgvFileList.Rows.Count >= 1 Then
                    dgvFileList.CurrentCell = dgvFileList.Rows(0).Cells(1)
                Else
                    Exit Sub
                End If
            End If

            Dim strPathAndFile As String
            'Dim strPathAndFile As String = gPath & m_Document.DocFile
            If m_Document.DocID <> 0 Then
                If Trim(m_Document.DocName) = "" Then
                    strPathAndFile = m_Path & m_Document.DocID
                Else
                    strPathAndFile = m_Path & m_Document.DocName
                End If

            Else
                strPathAndFile = m_Path & m_Document.DocFile ' Name
            End If

            'Dim strPathAndFile As String = gPath & m_Document.DocFile ' Name

            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If Not (Directory.Exists(m_Path)) Then
                'If Dir(m_Path, FileAttribute.Directory) = "" Then
                MkDir(m_Path)
            End If

            '/////////////////////////////
            'Get the document from the database and write it to a temp file
            Dim mStream As New ADODB.Stream
            mStream.Type = ADODB.StreamTypeEnum.adTypeBinary
            mStream.Open()
            mStream.Write(m_Document.GetDocumentRow.Item("Document"))
            If m_Document.DocID <> 0 Then
                'mStream.Write(m_Document.GetDocumentRow.Item("Document"))
                mStream.SaveToFile(strPathAndFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
            Else
                'mStream.Write(m_Document.GetDocumentImage())
                mStream.SaveToFile(strPathAndFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
            End If

            'mStream.SaveToFile(strPathAndFile, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)

            '////////////////////////////////
            'Show the file
            Select Case Mid(strPathAndFile.ToUpper, strPathAndFile.Length - 3)

                Case ".MSG", ".XLS", "XLSX", ".DOC", "DOCX", ".BMP"
                    pnlPreviewDoc.Visible = False

                Case Else
                    pnlPreviewDoc.Visible = True

            End Select

            wbShowFile.Navigate(strPathAndFile) ', True) '1 = NewWindow
            'wbShowFile.Navigate(New System.Uri(strPathAndFile), True) '1 = NewWindow

            'End If

        Catch er As Exception
            If er.Message = "Object reference not set to an instance of an object." Then
                MsgBox("Please select an item to make a preview.")
            Else
                MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End If
        End Try

    End Sub

#End Region

#Region "Private functions to fill both DataGridView ######################"

    Private Sub AddDocumentFiles(ByVal piType As Integer, ByRef Files As String())

        Try

            dgvFileList.AllowUserToAddRows = True

            panProgress.Visible = True

            pbProgress.Value = 0
            pbProgress.Maximum = Files.Length

            txtPBEvent.Refresh()

            ' Loop through the array and add the files to the list.
            For Each strFile As String In Files

                txtPBDocument.Text = strFile
                txtPBDocument.Refresh()

                Dim oDocument As New CDocument(m_Ord_Guid, m_Item_Guid, strFile, CInt(lvTypes.Items(intListDropIndex).SubItems(2).Text), lvTypes.Items(intListDropIndex).SubItems(1).Text)

                Dim n As Integer = dgvFileList.Rows.Add()
                With dgvFileList.Rows.Item(n)

                    .Cells(Columns.DocID).Value = oDocument.DocID
                    .Cells(Columns.HeaderID).Value = 0
                    .Cells(Columns.DocDescription).Value = oDocument.DocDescription
                    .Cells(Columns.DocFile).Value = oDocument.DocFile
                    .Cells(Columns.DocName).Value = oDocument.DocName
                    .Cells(Columns.DocType).Value = oDocument.DocType
                    .Cells(Columns.DocTypeID).Value = oDocument.DocTypeID
                    .Cells(Columns.Item_Guid).Value = oDocument.Item_Guid
                    .Cells(Columns.Ord_Guid).Value = oDocument.Ord_Guid

                End With
                pbProgress.Increment(1)
                panProgress.Invalidate()
            Next

            dgvFileList.AllowUserToAddRows = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        Finally
            panProgress.Visible = False
        End Try

    End Sub

    Private Sub AddDocumentFiles(ByRef pdtFiles As DataTable)

        Try

            dgvFileList.AllowUserToDeleteRows = True
            dgvFileList.Rows.Clear()
            dgvFileList.AllowUserToDeleteRows = False

            If pdtFiles.Rows.Count = 0 Then Exit Sub

            dgvFileList.AllowUserToAddRows = True

            ' Loop through the array and add the files to the list.
            For Each oRow As DataRow In pdtFiles.Rows

                Dim n As Integer = dgvFileList.Rows.Add()
                With dgvFileList.Rows.Item(n)

                    .Cells(Columns.DocID).Value = oRow.Item("DocID")
                    .Cells(Columns.HeaderID).Value = 0
                    .Cells(Columns.DocDescription).Value = Trim(oRow.Item("DocDescription").ToString)
                    .Cells(Columns.DocFile).Value = Trim(oRow.Item("DocFile").ToString)
                    .Cells(Columns.DocName).Value = Trim(oRow.Item("DocName").ToString)
                    .Cells(Columns.DocType).Value = Trim(oRow.Item("DocType").ToString)
                    .Cells(Columns.DocTypeID).Value = oRow.Item("DocTypeID")
                    .Cells(Columns.Item_Guid).Value = Trim(oRow.Item("Item_Guid").ToString)
                    .Cells(Columns.Ord_Guid).Value = Trim(oRow.Item("Ord_Guid").ToString)

                End With

            Next

            dgvFileList.AllowUserToAddRows = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub AddHistoryDocument(ByRef oDocument As CDocument, ByVal pintDocTypeID As Integer, ByVal pstrDocType As String)

        Try
            'Dim oDocument As New CDocument
            'oDocument = DirectCast(e.Data.GetData(GetType(CDocument)), CDocument)
            oDocument.DocType = pstrDocType ' lvTypes.Items(intListDropIndex).SubItems(1).Text
            oDocument.DocTypeID = pintDocTypeID ' CInt(lvTypes.Items(intListDropIndex).SubItems(2).Text)
            If oDocument.DocID <> 0 Then
                oDocument.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                oDocument.Item_Guid = g_oOrdline.Item_Guid
                oDocument.Save()
            Else
                oDocument.Save()
            End If

            Dim n As Integer = dgvFileList.Rows.Add()
            With dgvFileList.Rows.Item(n)

                .Cells(Columns.DocID).Value = 0
                .Cells(Columns.HeaderID).Value = 0
                .Cells(Columns.DocDescription).Value = oDocument.DocDescription
                .Cells(Columns.DocFile).Value = "c:\ExactTemp\" & oDocument.DocName ' oDocument.DocFile
                .Cells(Columns.DocName).Value = oDocument.DocName
                .Cells(Columns.DocType).Value = oDocument.DocType
                .Cells(Columns.DocTypeID).Value = oDocument.DocTypeID
                .Cells(Columns.Item_Guid).Value = oDocument.Item_Guid
                .Cells(Columns.Ord_Guid).Value = oDocument.Ord_Guid

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub CopyDeletedDocument(ByRef oDocument As CDocument, ByVal pintDocTypeID As Integer, ByVal pstrDocType As String)

        Try
            'Dim oDocument As New CDocument
            'oDocument = DirectCast(e.Data.GetData(GetType(CDocument)), CDocument)
            oDocument.DocType = pstrDocType ' lvTypes.Items(intListDropIndex).SubItems(1).Text
            oDocument.DocTypeID = pintDocTypeID ' CInt(lvTypes.Items(intListDropIndex).SubItems(2).Text)
            If oDocument.DocID <> 0 Then
                oDocument.Ord_Guid = m_oOrder.Ordhead.Ord_GUID
                oDocument.Item_Guid = g_oOrdline.Item_Guid
                oDocument.Save()
            Else
                oDocument.Save()
            End If

            Dim n As Integer = dgvFileList.Rows.Add()
            With dgvFileList.Rows.Item(n)

                .Cells(Columns.DocID).Value = 0
                .Cells(Columns.HeaderID).Value = 0
                .Cells(Columns.DocDescription).Value = oDocument.DocDescription
                .Cells(Columns.DocFile).Value = "c:\ExactTemp\" & oDocument.DocName ' oDocument.DocFile
                .Cells(Columns.DocName).Value = oDocument.DocName
                .Cells(Columns.DocType).Value = oDocument.DocType
                .Cells(Columns.DocTypeID).Value = oDocument.DocTypeID
                .Cells(Columns.Item_Guid).Value = oDocument.Item_Guid
                .Cells(Columns.Ord_Guid).Value = oDocument.Ord_Guid

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub AddHistoryFiles(ByRef pdtFiles As DataTable)

        Try

            dgvHistory.AllowUserToDeleteRows = True
            dgvHistory.Rows.Clear()
            dgvHistory.AllowUserToDeleteRows = False

            If pdtFiles.Rows.Count = 0 Then Exit Sub

            dgvHistory.AllowUserToAddRows = True

            ' Loop through the array and add the files to the list.
            For Each oRow As DataRow In pdtFiles.Rows

                Dim n As Integer = dgvHistory.Rows.Add()
                With dgvHistory.Rows.Item(n)

                    .Cells(Columns.Ord_No).Value = Trim(oRow.Item("Ord_No"))
                    .Cells(Columns.OE_Po_No).Value = Trim(oRow.Item("OE_Po_No"))
                    .Cells(Columns.DocID).Value = Trim(oRow.Item("DocID"))
                    .Cells(Columns.HeaderID).Value = Trim(oRow.Item("HeaderID"))
                    .Cells(Columns.DocDescription).Value = Trim(oRow.Item("DocDescription"))
                    '.Cells(Columns.DocFile).Value = Trim(oRow.Item("DocFile"))
                    .Cells(Columns.DocName).Value = Trim(oRow.Item("DocName"))
                    .Cells(Columns.DocType).Value = Trim(oRow.Item("DocType"))
                    .Cells(Columns.DocTypeID).Value = oRow.Item("DocTypeID")
                    .Cells(Columns.Item_Guid).Value = String.Empty
                    .Cells(Columns.Ord_Guid).Value = String.Empty

                End With

            Next

            dgvHistory.AllowUserToAddRows = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ChangeGridView(ByVal pintCodeTypeID As Integer)

        ' Check enum order for field order

        Dim strSql As String
        If pintCodeTypeID = 0 Then ' when All types is selected

            '"FROM       OEI_Documents WITH (Nolock) " & _
            strSql = _
            "SELECT     DocID, Ord_Guid, Item_Guid, DocTypeID, DocType, DocFile, Ord_No, DocName, DocDescription " & _
            "FROM       Exact_Traveler_Document WITH (Nolock) " & _
            "WHERE      Ord_Guid = '" & m_Ord_Guid & "' " & _
            "ORDER BY   DocType, DocName "

            dgvFileList.Columns(Columns.DocType).Visible = True

        Else

            '"FROM       OEI_Documents WITH (Nolock) " & _
            strSql = _
            "SELECT     DocID, Ord_Guid, Item_Guid, DocTypeID, DocType, DocFile, Ord_No, DocName, DocDescription " & _
            "FROM       Exact_Traveler_Document WITH (Nolock) " & _
            "WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND DocTypeID = " & CInt(lvTypes.Items(intListDropIndex).SubItems(2).Text) & " " & _
            "ORDER BY   DocType, DocName "

            dgvFileList.Columns(Columns.DocType).Visible = False

        End If

        Dim dt As New DataTable
        Dim conn As New cDBA

        dt = conn.DataTable(strSql)

        Call AddDocumentFiles(dt)

        lblSelectedType.Text = lvTypes.Items(intListDropIndex).SubItems(1).Text

    End Sub

    Private Sub FillFileColumns(ByRef dgvList As DataGridView)

        If dgvList.ColumnCount < 5 Then

            ' Lors du chargement quelque part, le numero de colonne n'est pas le bon
            ' Le champ Item_Guid sort la valeur du DocTypeID

            dgvList.Columns.Clear()
            dgvList.Columns.Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
            dgvList.Columns.Add(DGVTextBoxColumn("DocID", "DocID", 0))
            dgvList.Columns.Add(DGVTextBoxColumn("HeaderID", "HeaderID", 0))
            dgvList.Columns.Add(DGVTextBoxColumn("Ord_Guid", "Ord_Guid", 0))
            dgvList.Columns.Add(DGVTextBoxColumn("Item_Guid", "Item_Guid", 0))
            dgvList.Columns.Add(DGVTextBoxColumn("DocTypeID", "DocTypeID", 0))
            dgvList.Columns.Add(DGVTextBoxColumn("DocType", "DocType", 160))
            dgvList.Columns.Add(DGVTextBoxColumn("DocFile", "DocFile", 160))
            dgvList.Columns.Add(DGVTextBoxColumn("Ord_No", "Ord_No", 100))
            dgvList.Columns.Add(DGVTextBoxColumn("OE_Po_No", "OE_Po_No", 100))
            dgvList.Columns.Add(DGVTextBoxColumn("DocName", "DocName", 250))
            dgvList.Columns.Add(DGVTextBoxColumn("DocDescription", "DocDescription", 300))
            dgvList.RowHeadersWidth = 22
            dgvList.Columns(Columns.FirstCol).Visible = False
            dgvList.Columns(Columns.DocID).Visible = False
            dgvList.Columns(Columns.HeaderID).Visible = False
            dgvList.Columns(Columns.Ord_Guid).Visible = False
            dgvList.Columns(Columns.Item_Guid).Visible = False
            dgvList.Columns(Columns.DocTypeID).Visible = False
            dgvList.Columns(Columns.DocFile).Visible = False
            If dgvList.Name = "dgvFileList" Then
                dgvList.Columns(Columns.Ord_No).Visible = False
                dgvList.Columns(Columns.OE_Po_No).Visible = False
            End If
        End If

    End Sub

    Private Sub FillEmailColumns()

        ' Lors du chargement quelque part, le numero de colonne n'est pas le bon
        ' Le champ Item_Guid sort la valeur du DocTypeID

        dgvEmails.Columns.Clear()
        dgvEmails.Columns.Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
        dgvEmails.Columns.Add(DGVTextBoxColumn("ID", "ID", 0))
        dgvEmails.Columns.Add(DGVTextBoxColumn("Email_From", "From", 0))
        dgvEmails.Columns.Add(DGVTextBoxColumn("Email_To", "To", 0))
        dgvEmails.Columns.Add(DGVTextBoxColumn("Email_To_Name", "To", 165))
        dgvEmails.Columns.Add(DGVTextBoxColumn("Email_CC", "CC", 0))
        dgvEmails.Columns.Add(DGVTextBoxColumn("Subject", "Subject", 240))
        dgvEmails.Columns.Add(DGVTextBoxColumn("Body", "Body", 300))
        dgvEmails.Columns.Add(DGVTextBoxColumn("Ord_Guid", "Ord_Guid", 0))
        dgvEmails.Columns.Add(DGVTextBoxColumn("Ord_No", "Ord_No", 0))
        dgvEmails.Columns.Add(DGVTextBoxColumn("CreateTS", "Created", 135))
        dgvEmails.Columns.Add(DGVTextBoxColumn("SendTS", "Sent", 0))
        dgvEmails.Columns.Add(DGVTextBoxColumn("UserID", "User ID", 90))
        dgvEmails.Columns(EmailColumns.FirstCol).Visible = False
        dgvEmails.Columns(EmailColumns.ID).Visible = False
        dgvEmails.Columns(EmailColumns.Email_From).Visible = False
        dgvEmails.Columns(EmailColumns.Email_To).Visible = False
        dgvEmails.Columns(EmailColumns.Email_CC).Visible = False
        dgvEmails.Columns(EmailColumns.Ord_Guid).Visible = False
        dgvEmails.Columns(EmailColumns.Ord_No).Visible = False
        dgvEmails.Columns(EmailColumns.SendTS).Visible = False

    End Sub

    Private Sub FillTypes()

        Try
            If lvTypes.Items.Count <> 0 Then Exit Sub

            Dim dt As New DataTable
            dt = m_Document.DocTypeDataTable()

            lvTypes.Columns.Clear()
            lvTypes.Items.Clear()

            lvTypes.View = View.Details

            If lvTypes.Columns.Count = 0 Then
                lvTypes.Columns.Add("Z", 0, HorizontalAlignment.Left)
                lvTypes.Columns.Add("DocType", 165, HorizontalAlignment.Left)
                lvTypes.Columns.Add("DocTypeID", 0, HorizontalAlignment.Left)
                'lvTypes.Columns("DocType").Width = 100 ' 165
                'lvTypes.Columns("DocTypeID").Width = 65 ' 0
            End If

            If dt.Rows.Count <> 0 Then

                For Each drRow As DataRow In dt.Rows

                    Dim lviItem As New ListViewItem
                    lviItem.SubItems.Clear()
                    lviItem.SubItems.Add(drRow.Item("DocType").ToString)
                    lviItem.SubItems.Add(drRow.Item("DocTypeID").ToString)
                    lvTypes.Items.Add(lviItem)

                    'lviItem = lvTypes.Items.Add(drRow.Item("DocType").ToString)
                    'lviItem.SubItems(1).Text = drRow.Item("DocTypeID").ToString
                    'lviItem.SubItems("DocTypeID").Text = drRow.Item("DocTypeID").ToString
                    'lbTypes.Items.Add(drRow.Item("DocType").ToString)

                Next

            End If

            dt.Dispose()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ShowEmails()

        ' Check enum order for field order

        Dim strSql As String
        strSql = _
        "SELECT     0 AS FirstCol, ID, Email_From, Email_To,        " & _
        "           Email_To_Name, Email_CC, Subject, Body,         " & _
        "           Ord_Guid, Ord_No, CreateTS, SendTS, UserID      " & _
        "FROM       OEI_Email WITH (Nolock)                         " & _
        "WHERE      Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "'  " & _
        "ORDER BY   ID                                              " ' CreateTS

        Dim dt As New DataTable
        Dim db As New cDBA

        dt = db.DataTable(strSql)
        dgvEmails.DataSource = dt

        dgvEmails.AllowUserToAddRows = False
        dgvEmails.AllowUserToDeleteRows = False

    End Sub

    Private Sub ShowHistoryOrders()

        pnlHistory.Visible = chkShowHistory.Checked

        If chkShowHistory.Checked Then
            dgvFileList.Width = 399
        Else
            dgvFileList.Width = 807 ' 686
        End If

        If chkShowHistory.Checked Then
            If lvTypes.SelectedItems.Count <> 0 Then
                For Each item As ListViewItem In lvTypes.SelectedItems
                    intListDropIndex = item.Index
                Next item

                Call ShowHistoryOrders(intListDropIndex)
            End If
        End If

    End Sub

    Private Sub ShowHistoryOrders(ByVal pintCodeTypeID As Integer)

        Dim strSql As String
        Dim strSearchSql As String = ""

        If Trim(txtSearchElements.Text) <> "" Then
            Dim strSearchElements As String() = txtSearchElements.Text.Split(New Char() {" "c})
            For Each strSearchPart As String In strSearchElements

                strSearchPart = SqlCompliantString(strSearchPart)

                If btnOr.Checked Then
                    If strSearchSql <> "" Then strSearchSql = strSearchSql & " OR "
                Else
                    If strSearchSql <> "" Then strSearchSql = strSearchSql & " AND "
                End If

                If strSearchSql = "" Then strSearchSql = "("

                strSearchSql = strSearchSql & " ( " & _
                "D.Ord_No like '%" & strSearchPart & "%' OR " & _
                "D.DocName like '%" & strSearchPart & "%' OR " & _
                "D.DocDescription LIKE '%" & strSearchPart & "%' OR " & _
                "H.OE_Po_No LIKE '%" & strSearchPart & "%' )"

            Next
            If strSearchSql <> "" Then strSearchSql = strSearchSql & ")"

        Else
            strSearchSql = ""
        End If

        If txtOrd_Dt.Text = "" Then txtOrd_Dt.Text = Now.AddYears(-1).AddDays(1).ToShortDateString
        If txtOrd_Dt_Shipped.Text = "" Then txtOrd_Dt_Shipped.Text = Now.ToShortDateString

        'DocID()
        'HeaderID()
        'Ord_Guid()
        'Item_Guid()
        'DocTypeID()
        'DocType()
        'DocFile()
        'Ord_No()
        'DocName()
        'DocDescription()
        'OE_Po_No()

        strSql = _
        "SELECT 	D.DocID, D.HeaderID, D.DocTypeID, D.DocType, D.DocName, D.Ord_No, " & _
        "			D.DocName, D.DocDescription, H.OE_Po_No " & _
        "FROM 		EXACT_TRAVELER_DOCUMENT_DELETED D WITH (NOLOCK) " & _
        "INNER JOIN OEHdrHst_Sql H WITH (Nolock) ON D.Ord_No = H.Ord_No " & _
        "WHERE		H.Ord_No = '  445694' "

        strSql = _
        "SELECT 	D.DocID, D.HeaderID, D.DocTypeID, D.DocType, D.DocName, D.Ord_No, " & _
        "			D.DocName, D.DocDescription, H.OE_Po_No " & _
        "FROM 		EXACT_TRAVELER_DOCUMENT D WITH (NOLOCK) " & _
        "INNER JOIN OEHdrHst_Sql H WITH (Nolock) ON D.Ord_No = H.Ord_No " & _
        "WHERE		H.Cus_No = '" & m_Cus_No & "' " & _
        "			AND D.CreateDate BETWEEN '" & txtOrd_Dt.Text & "' AND '" & txtOrd_Dt_Shipped.Text & " 23:59:59' "

        If pintCodeTypeID <> 0 Then
            strSql = strSql & " AND (D.DocTypeID = " & CInt(lvTypes.Items(pintCodeTypeID).SubItems(2).Text) & ") "
        Else
        End If
        If strSearchSql <> "" Then strSql = strSql & "			AND " & strSearchSql & " "

        strSql = strSql & " ORDER BY D.DocType, D.DocDescription, D.DocName "

        '"			AND D.CreateDate BETWEEN '2011/01/01' AND '2011/12/31 23:59:59' " & _
        '"			AND " & strSearchSql
        '"			( " & _
        '"			(D.Ord_No like '%410391%' OR D.DocName like '%410391%' OR D.DocDescription LIKE '%410391%' OR H.OE_Po_No LIKE '%410391%' ) OR " & _
        '"			(D.Ord_No like '%12345%' OR D.DocName like '%12345%' OR D.DocDescription LIKE '%12345%' OR H.OE_Po_No LIKE '%12345%' ) OR " & _
        '"			(D.Ord_No like '%99887%' OR D.DocName like '%99887%' OR D.DocDescription LIKE '%99887%' OR H.OE_Po_No LIKE '%99887%' ) OR " & _
        '"			(D.Ord_No like '%5507981%' OR D.DocName like '%5507981%' OR D.DocDescription LIKE '%5507981%' OR H.OE_Po_No LIKE '%5507981%' ) " & _
        '"			) "

        If pintCodeTypeID = 0 Then ' when All types is selected

            dgvHistory.Columns(Columns.DocType).Visible = True

        Else

            dgvHistory.Columns(Columns.DocType).Visible = False

        End If

        Dim dt As New DataTable
        Dim conn As New cDBA

        dt = conn.DataTable(strSql)

        Call AddHistoryFiles(dt)
        'Call AddDocumentFiles(dt)

        'lblSelectedType.Text = lvTypes.Items(intListDropIndex).SubItems(1).Text

    End Sub

#End Region

#Region "Public procedures ################################################"

    Public Sub Fill()

        m_Ord_Guid = m_oOrder.Ordhead.Ord_GUID
        m_Cus_No = m_oOrder.Ordhead.Cus_No
        chkOrderAckSaveOnly.Checked = (m_oOrder.Ordhead.OrderAckSaveOnly = 1)

        'm_Item_Guid = g_oOrdline.Item_Guid

        Call FillTypes()

        ' Set and fill document grid
        Call FillFileColumns(dgvFileList)

        Call FillFileColumns(dgvHistory)

        Call ChangeGridView(0)

        ' Set and fill email grid
        Call FillEmailColumns()

        Call ShowEmails()

    End Sub

    Public Sub Reset()

        chkOrderAckSaveOnly.Checked = True

    End Sub

    Public Sub Save()

        m_oOrder.Ordhead.OrderAckSaveOnly = IIf(chkOrderAckSaveOnly.Checked, 1, 0)
        pnlPreviewDoc.Visible = False

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property Cus_No() As String
        Get
            Cus_No = m_Cus_No
        End Get
        Set(ByVal value As String)
            m_Cus_No = value
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
    Public Property Ord_Guid() As String
        Get
            Ord_Guid = m_Ord_Guid
        End Get
        Set(ByVal value As String)
            m_Ord_Guid = value
        End Set
    End Property

#End Region

    'Private Sub ucDocument_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

    '    ' If Ctl-Number, sets the new tab.
    '    If e.Control Then
    '        Select Case e.KeyCode
    '            Case Keys.D1
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Header)
    '            Case Keys.D2
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CustomerContact)
    '            Case Keys.D3
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Lines)
    '            Case Keys.D4
    '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Documents)
    '            Case Keys.D5
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.HeaderAll)
    '            Case Keys.D6
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Taxes)
    '            Case Keys.D7
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
    '            Case Keys.D8
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
    '            Case Keys.D9
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
    '        End Select

    '    End If

    'End Sub

    Private Sub chkOrderAckSaveOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOrderAckSaveOnly.CheckedChanged

        m_oOrder.Ordhead.OrderAckSaveOnly = IIf(chkOrderAckSaveOnly.Checked, 1, 0)

    End Sub

    Private Sub cmdPreviewClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPreviewClose.Click

        pnlPreviewDoc.Visible = False

    End Sub

    Private Sub cmdNewEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewEmail.Click

        Dim oEmail As New frmEmail
        oEmail.SetEmailToAddress(Trim(m_oOrder.Ordhead.User_Def_Fld_3))

        oEmail.ShowDialog()

        ShowEmails()

    End Sub

    Private Sub cmdDeleteEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteEmail.Click

        If dgvEmails.Rows.Count > 0 Then

            If dgvEmails.CurrentRow Is Nothing Then Exit Sub

            Dim mbrResulat As New MsgBoxResult

            Dim oEmail As New cEmail(dgvEmails.CurrentRow.Cells(EmailColumns.ID).Value)

            mbrResulat = MsgBox("Do you want to remove this email to " & oEmail.Email_To, MsgBoxStyle.YesNoCancel, "Delete an email")

            If mbrResulat = MsgBoxResult.Yes Then
                oEmail.Delete()
                dgvEmails.AllowUserToDeleteRows = True
                dgvEmails.Rows.RemoveAt(dgvEmails.CurrentRow.Index)
                dgvEmails.AllowUserToDeleteRows = False

            End If

        End If

    End Sub

    Private Sub dgvEmails_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvEmails.DoubleClick

        Dim oEmail As cEmail
        Dim ofrmEmail As New frmEmail

        Try
            If dgvEmails.Rows.Count > 0 Then

                If dgvEmails.CurrentRow Is Nothing Then Exit Sub

                Dim mbrResulat As New MsgBoxResult

                oEmail = New cEmail(dgvEmails.CurrentRow.Cells(EmailColumns.ID).Value)
                ofrmEmail.Email = oEmail
                ofrmEmail.ShowDialog()
                'ofrmEmail.SetEmailToAddress(Trim(m_oOrder.Ordhead.User_Def_Fld_3))
                ShowEmails()

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

End Class

