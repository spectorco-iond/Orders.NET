Public Class frmMDB_CUS_DEC_SUPP_ITEM

    Private m_ShowEvents As Boolean = False
    Private m_DirtyLine As Boolean = False
    Private m_Cus_Dec_ID As Integer = 0
    Dim dtGrid As DataTable

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Load")

        Try

            If dgvSupp_Item.Columns.Count > 0 Then
                For iCol As Integer = dgvSupp_Item.Columns.Count To 1 Step -1
                    dgvSupp_Item.Columns.RemoveAt(iCol - 1)
                Next
            End If

            With dgvSupp_Item.Columns

                '.Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
                .Add(DGVTextBoxColumn("Cus_Dec_Supp_Item_ID", "Cus_Dec_Supp_Item_ID", 20))
                .Add(DGVTextBoxColumn("Cus_Dec_ID", "Cus_Dec_ID", 100))
                .Add(DGVTextBoxColumn("Item_No", "Item_No", 100))
                .Add(DGVTextBoxColumn("Qty_Ordered", "Qty_Ordered", 100))
                .Add(DGVTextBoxColumn("Unit_Price", "Unit_Price", 100))
                .Add(DGVTextBoxColumn("User_Login", "User_Login", 80))
                .Add(DGVTextBoxColumn("Update_TS", "Update_TS", 80))

            End With

            Dim db As New cDBA

            Dim strSql As String = _
            "SELECT     Cus_Dec_Supp_Item_ID, Cus_Dec_ID, Item_No, Qty_Ordered, " & _
            "           Unit_Price, User_Login, Update_TS " & _
            "FROM       MDB_CUS_DEC_SUPP_ITEM WITH (Nolock) " & _
            "WHERE      CUS_DEC_ID = " & m_Cus_Dec_ID

            dtGrid = db.DataTable(strSql)

            dgvSupp_Item.DataSource = dtGrid

            dgvSupp_Item.AllowUserToAddRows = False
            dgvSupp_Item.AllowUserToOrderColumns = False

            For lPos As Integer = 0 To dgvSupp_Item.Columns.Count - 1
                dgvSupp_Item.Columns(lPos).SortMode = DataGridViewColumnSortMode.NotSortable
            Next lPos

            dgvSupp_Item.CausesValidation = True

            Dim oCellStyle As New DataGridViewCellStyle()
            oCellStyle = New DataGridViewCellStyle()
            oCellStyle.Format = "##,##0.00"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            With dgvSupp_Item
                .Columns(Columns.Qty_Ordered).DefaultCellStyle = oCellStyle
                .Columns(Columns.Unit_Price).DefaultCellStyle = oCellStyle

                '.Columns(Supp_ItemColumns.ColHeader).Frozen = True
                .Columns(Columns.Cus_Dec_Supp_Item_ID).Frozen = True
                .Columns(Columns.Cus_Dec_ID).Frozen = True
                .Columns(Columns.Item_No).Frozen = True

                '.Columns(Supp_ItemColumns.ColHeader).Visible = False
                .Columns(Columns.Cus_Dec_Supp_Item_ID).Visible = False
                .Columns(Columns.Cus_Dec_ID).Visible = False

                .Columns(Columns.User_Login).Visible = False
                .Columns(Columns.Update_TS).Visible = False

                .CurrentCell = .Rows(0).Cells(Columns.Item_No)

            End With

            cmdClose.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSupp_Item.CellBeginEdit

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellBeginEdit")

        ' HERE TO SET CANCEL = TRUE IF NEEDED
        Try

            m_DirtyLine = True
            Debug.Print("Set DirtyLine True")

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellEnter

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellEnter")

        ' HERE TO SET CURSOR TO SPECIFIC COLUMN WHEN MISSING DATA
        Try

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvSupp_Item.CellValidating

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellValidating")

        ' HERE TO VALIDATE
        Try

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_Type_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellValueChanged

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellValueChanged")

        ' HERE TO PUT DATA IN UPPERCASE AFTER ENTERED IF WANTED
        Try

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvSupp_Item.EditingControlShowing

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.EditingControlShowing")

        ' HERE TO OVERWRITE THE KEYDOWN EVENT OF THE TEXTBOX TO A CUSTOM CONTROL TEXTBOX
        Try

            If TypeOf e.Control Is TextBox Then

                Dim tb As TextBox = CType(e.Control, TextBox)
                'Remove the existing handler if there is one.
                RemoveHandler tb.KeyDown, AddressOf dgvSupp_Item_KeyDown ' TextBox_TextChanged

                'Add a new handler.
                AddHandler tb.KeyDown, AddressOf dgvSupp_Item_KeyDown ' TextBox_TextChanged
                '    End With
                'End If

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.GotFocus

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.GotFocus")

        ' HERE TO GET THE ACTUAL LINE ON FOCUS OF GRID CONTROL
        Try

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvSupp_Item.KeyDown

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.KeyDown")

        ' HERE TO CHECK FOR KEYCODES TO LAUNCH SHORTCUTS
        Try

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.Leave

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Leave")

        ' HERE TO FORCE A SAVE ON LOSS OF FOCUS OF CONTROL
        Try

            Debug.Print("dgvSupp_Item.Leave")
            'Call SaveOrderLine(dgvSupp_Item.Rows(dgvSupp_Item.CurrentRow.Index))

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.RowEnter

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowEnter")

        ' HERE TO LOAD NEW RECORD INTO OBJECT
        Try

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.RowLeave

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowLeave")

        ' HERE TO SAVE CURRENT RECORD TO OBJECT
        Try

            dgvSupp_Item.EndEdit()
            'Call SaveOrderLine(dgvSupp_Item.Rows(dgvSupp_Item.CurrentRow.Index))

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub SaveOrderLine(ByVal oRow As DataGridViewRow)

        Try

            If m_DirtyLine Then

                Debug.Print("DirtyLine Save")

                Dim oSupp_Item As New cMdb_Cus_Dec_Supp_Item(oRow.Cells(Columns.Cus_Dec_Supp_Item_ID).Value)

                oSupp_Item.Cus_Dec_ID = oRow.Cells(Columns.Cus_Dec_ID).Value
                oSupp_Item.Item_No = oRow.Cells(Columns.Item_No).Value
                oSupp_Item.Unit_Price = oRow.Cells(Columns.Unit_Price).Value
                oSupp_Item.Qty_Ordered = oRow.Cells(Columns.Qty_Ordered).Value
                oSupp_Item.Save()

                m_DirtyLine = False

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellContentClick

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellContentClick")

        ' HERE TO SET DIRTYLINE TO TRUE FOR A CHECKBOX
        Try

            'Select Case e.ColumnIndex

            'End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvSupp_Item_RowValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSupp_Item.RowValidating

        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowValidating")

        ' HERE TO FORCE A SAVE ON LOSS OF FOCUS OF CONTROL
        Try

            Debug.Print("dgvSupp_Item.RowValidating")
            Call SaveOrderLine(dgvSupp_Item.Rows(dgvSupp_Item.CurrentRow.Index))

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.FormClosed")

        Try

            dgvSupp_Item.DataSource = Nothing

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.FormClosing")

        Try

            dgvSupp_Item.EndEdit()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


#Region "dgvSupp_Item ORDER OF EVENTS #########################################################"

    Private Sub dgvSupp_Item_CancelRowEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.QuestionEventArgs) Handles dgvSupp_Item.CancelRowEdit
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CancelRowEdit")
    End Sub

    Private Sub dgvSupp_Item_AllowUserToAddRowsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.AllowUserToAddRowsChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AllowUserToAddRowsChanged")
    End Sub

    Private Sub dgvSupp_Item_CausesValidationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.CausesValidationChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CausesValidationChanged")
    End Sub

    Private Sub dgvSupp_Item_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellClick")
    End Sub

    Private Sub dgvSupp_Item_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellContentDoubleClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellContentDoubleClick")
    End Sub

    Private Sub dgvSupp_Item_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellDoubleClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellDoubleClick")
    End Sub

    Private Sub dgvSupp_Item_CellErrorTextChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellErrorTextChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellErrorTextChanged")
    End Sub

    Private Sub dgvSupp_Item_CellErrorTextNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellErrorTextNeededEventArgs) Handles dgvSupp_Item.CellErrorTextNeeded
        'If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellErrorTextNeeded")
    End Sub

    Private Sub dgvSupp_Item_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgvSupp_Item.CellFormatting
        'If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellFormatting")
    End Sub

    Private Sub dgvSupp_Item_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellLeave
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellLeave")
    End Sub

    Private Sub dgvSupp_Item_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvSupp_Item.CellPainting
        'If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellPainting")
    End Sub

    Private Sub dgvSupp_Item_CellParsing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellParsingEventArgs) Handles dgvSupp_Item.CellParsing
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellParsing")
    End Sub

    Private Sub dgvSupp_Item_CellStateChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellStateChangedEventArgs) Handles dgvSupp_Item.CellStateChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellStateChanged")
    End Sub

    Private Sub dgvSupp_Item_CellStyleChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_CellStyleContentChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellStyleContentChangedEventArgs) Handles dgvSupp_Item.CellStyleContentChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellStyleContentChanged")
    End Sub

    Private Sub dgvSupp_Item_CellToolTipTextChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellToolTipTextChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellToolTipTextChanged")
    End Sub

    Private Sub dgvSupp_Item_CellToolTipTextNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellToolTipTextNeededEventArgs) Handles dgvSupp_Item.CellToolTipTextNeeded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellToolTipTextNeeded")
    End Sub

    Private Sub dgvSupp_Item_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellValidated
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellValidated")
    End Sub

    Private Sub dgvSupp_Item_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles dgvSupp_Item.CellValueNeeded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellValueNeeded")
    End Sub

    Private Sub dgvSupp_Item_CellValuePushed(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles dgvSupp_Item.CellValuePushed
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellValuePushed")
    End Sub

    Private Sub dgvSupp_Item_ChangeUICues(ByVal sender As Object, ByVal e As System.Windows.Forms.UICuesEventArgs) Handles dgvSupp_Item.ChangeUICues
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ChangeUICues")
    End Sub

    Private Sub dgvSupp_Item_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.Click
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Click")
    End Sub

    Private Sub dgvSupp_Item_ClientSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ClientSizeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ClientSizeChanged")
    End Sub

    Private Sub dgvSupp_Item_ContextMenuChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ContextMenuChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ContextMenuChanged")
    End Sub

    Private Sub dgvSupp_Item_ContextMenuStripChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ContextMenuStripChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ContextMenuStripChanged")
    End Sub

    Private Sub dgvSupp_Item_ControlAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs) Handles dgvSupp_Item.ControlAdded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ControlAdded")
    End Sub

    Private Sub dgvSupp_Item_ControlRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs) Handles dgvSupp_Item.ControlRemoved
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ControlRemoved")
    End Sub

    Private Sub dgvSupp_Item_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.CurrentCellChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CurrentCellChanged")
    End Sub

    Private Sub dgvSupp_Item_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.CurrentCellDirtyStateChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CurrentCellDirtyStateChanged")
    End Sub

    Private Sub dgvSupp_Item_CursorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.CursorChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CursorChanged")
    End Sub

    Private Sub dgvSupp_Item_DataBindingComplete(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewBindingCompleteEventArgs) Handles dgvSupp_Item.DataBindingComplete
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DataBindingComplete")
    End Sub

    Private Sub dgvSupp_Item_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSupp_Item.DataError
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DataError")
    End Sub

    Private Sub dgvSupp_Item_DataMemberChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.DataMemberChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DataMemberChanged")
    End Sub

    Private Sub dgvSupp_Item_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.DataSourceChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DataSourceChanged")
    End Sub

    Private Sub dgvSupp_Item_DefaultCellStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.DefaultCellStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DefaultCellStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_DefaultValuesNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.DefaultValuesNeeded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DefaultValuesNeeded")
    End Sub

    Private Sub dgvSupp_Item_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.Disposed
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Disposed")
    End Sub

    Private Sub dgvSupp_Item_DockChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.DockChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DockChanged")
    End Sub

    Private Sub dgvSupp_Item_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.DoubleClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DoubleClick")
    End Sub

    Private Sub dgvSupp_Item_EditModeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.EditModeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.EditModeChanged")
    End Sub

    Private Sub dgvSupp_Item_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.EnabledChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.EnabledChanged")
    End Sub

    Private Sub dgvSupp_Item_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.Enter
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Enter")
    End Sub

    Private Sub dgvSupp_Item_GiveFeedback(ByVal sender As Object, ByVal e As System.Windows.Forms.GiveFeedbackEventArgs) Handles dgvSupp_Item.GiveFeedback
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.GiveFeedback")
    End Sub

    Private Sub dgvSupp_Item_HandleCreated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.HandleCreated
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.HandleCreated")
    End Sub

    Private Sub dgvSupp_Item_HandleDestroyed(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.HandleDestroyed
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.HandleDestroyed")
    End Sub

    Private Sub dgvSupp_Item_HelpRequested(ByVal sender As Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles dgvSupp_Item.HelpRequested
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.HelpRequested")
    End Sub

    Private Sub dgvSupp_Item_ImeModeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ImeModeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ImeModeChanged")
    End Sub

    Private Sub dgvSupp_Item_Invalidated(ByVal sender As Object, ByVal e As System.Windows.Forms.InvalidateEventArgs) Handles dgvSupp_Item.Invalidated
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Invalidated")
    End Sub

    Private Sub dgvSupp_Item_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvSupp_Item.KeyPress
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.KeyPress")
    End Sub

    Private Sub dgvSupp_Item_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvSupp_Item.KeyUp
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.KeyUp")
    End Sub

    Private Sub dgvSupp_Item_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles dgvSupp_Item.Layout
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Layout")
    End Sub

    Private Sub dgvSupp_Item_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.LocationChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.LocationChanged")
    End Sub

    Private Sub dgvSupp_Item_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.LostFocus
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.LostFocus")
    End Sub

    Private Sub dgvSupp_Item_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.Move
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Move")
    End Sub

    Private Sub dgvSupp_Item_MultiSelectChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.MultiSelectChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MultiSelectChanged")
    End Sub

    Private Sub dgvSupp_Item_NewRowNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.NewRowNeeded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.NewRowNeeded")
    End Sub

    Private Sub dgvSupp_Item_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles dgvSupp_Item.Paint
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Paint")
    End Sub

    Private Sub dgvSupp_Item_ParentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ParentChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ParentChanged")
    End Sub

    Private Sub dgvSupp_Item_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles dgvSupp_Item.PreviewKeyDown
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.PreviewKeyDown")
    End Sub

    Private Sub dgvSupp_Item_QueryAccessibilityHelp(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryAccessibilityHelpEventArgs) Handles dgvSupp_Item.QueryAccessibilityHelp
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.QueryAccessibilityHelp")
    End Sub

    Private Sub dgvSupp_Item_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs) Handles dgvSupp_Item.QueryContinueDrag
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.QueryContinueDrag")
    End Sub

    Private Sub dgvSupp_Item_ReadOnlyChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ReadOnlyChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ReadOnlyChanged")
    End Sub

    Private Sub dgvSupp_Item_RegionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.RegionChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RegionChanged")
    End Sub

    Private Sub dgvSupp_Item_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.Resize
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Resize")
    End Sub

    Private Sub dgvSupp_Item_RightToLeftChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.RightToLeftChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RightToLeftChanged")
    End Sub

    Private Sub dgvSupp_Item_RowDirtyStateNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.QuestionEventArgs) Handles dgvSupp_Item.RowDirtyStateNeeded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowDirtyStateNeeded")
    End Sub

    Private Sub dgvSupp_Item_RowDividerDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowDividerDoubleClickEventArgs) Handles dgvSupp_Item.RowDividerDoubleClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowDividerDoubleClick")
    End Sub

    Private Sub dgvSupp_Item_RowDividerHeightChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.RowDividerHeightChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowDividerHeightChanged")
    End Sub

    Private Sub dgvSupp_Item_RowErrorTextChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.RowErrorTextChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowErrorTextChanged")
    End Sub

    Private Sub dgvSupp_Item_RowErrorTextNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowErrorTextNeededEventArgs) Handles dgvSupp_Item.RowErrorTextNeeded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowErrorTextNeeded")
    End Sub

    Private Sub dgvSupp_Item_RowHeaderCellChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.RowHeaderCellChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeaderCellChanged")
    End Sub

    Private Sub dgvSupp_Item_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSupp_Item.RowHeaderMouseClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeaderMouseClick")
    End Sub

    Private Sub dgvSupp_Item_RowHeaderMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSupp_Item.RowHeaderMouseDoubleClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeaderMouseDoubleClick")
    End Sub

    Private Sub dgvSupp_Item_RowHeadersBorderStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.RowHeadersBorderStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeadersBorderStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_RowHeadersDefaultCellStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.RowHeadersDefaultCellStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeadersDefaultCellStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_RowHeadersWidthChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.RowHeadersWidthChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeadersWidthChanged")
    End Sub

    Private Sub dgvSupp_Item_RowHeadersWidthSizeModeChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewAutoSizeModeEventArgs) Handles dgvSupp_Item.RowHeadersWidthSizeModeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeadersWidthSizeModeChanged")
    End Sub

    Private Sub dgvSupp_Item_RowHeightChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.RowHeightChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeightChanged")
    End Sub

    Private Sub dgvSupp_Item_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles dgvSupp_Item.RowHeightInfoNeeded
        'If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeightInfoNeeded")
    End Sub

    Private Sub dgvSupp_Item_RowHeightInfoPushed(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoPushedEventArgs) Handles dgvSupp_Item.RowHeightInfoPushed
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowHeightInfoPushed")
    End Sub

    Private Sub dgvSupp_Item_RowMinimumHeightChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.RowMinimumHeightChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowMinimumHeightChanged")
    End Sub

    Private Sub dgvSupp_Item_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgvSupp_Item.RowPostPaint
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowPostPaint")
    End Sub

    Private Sub dgvSupp_Item_RowPrePaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPrePaintEventArgs) Handles dgvSupp_Item.RowPrePaint
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowPrePaint")
    End Sub

    Private Sub dgvSupp_Item_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvSupp_Item.RowsAdded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowsAdded")
    End Sub

    Private Sub dgvSupp_Item_RowsDefaultCellStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.RowsDefaultCellStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowsDefaultCellStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvSupp_Item.RowsRemoved
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowsRemoved")
    End Sub

    Private Sub dgvSupp_Item_RowStateChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowStateChangedEventArgs) Handles dgvSupp_Item.RowStateChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowStateChanged")
    End Sub

    Private Sub dgvSupp_Item_RowUnshared(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.RowUnshared
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowUnshared")
    End Sub

    Private Sub dgvSupp_Item_RowValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.RowValidated
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowValidated")
    End Sub

    Private Sub dgvSupp_Item_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles dgvSupp_Item.Scroll
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Scroll")
    End Sub

    Private Sub dgvSupp_Item_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.SelectionChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.SelectionChanged")
    End Sub

    Private Sub dgvSupp_Item_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.SizeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.SizeChanged")
    End Sub

    Private Sub dgvSupp_Item_SortCompare(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewSortCompareEventArgs) Handles dgvSupp_Item.SortCompare
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.SortCompare")
    End Sub

    Private Sub dgvSupp_Item_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.Sorted
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Sorted")
    End Sub

    Private Sub dgvSupp_Item_SystemColorsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.SystemColorsChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.SystemColorsChanged")
    End Sub

    Private Sub dgvSupp_Item_TabIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.TabIndexChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.TabIndexChanged")
    End Sub

    Private Sub dgvSupp_Item_TabStopChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.TabStopChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.TabStopChanged")
    End Sub

    Private Sub dgvSupp_Item_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.UserAddedRow
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.UserAddedRow")
    End Sub

    Private Sub dgvSupp_Item_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.UserDeletedRow
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.UserDeletedRow")
    End Sub

    Private Sub dgvSupp_Item_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvSupp_Item.UserDeletingRow
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.UserDeletingRow")
    End Sub

    Private Sub dgvSupp_Item_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.Validated
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Validated")
    End Sub

    Private Sub dgvSupp_Item_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dgvSupp_Item.Validating
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.Validating")
    End Sub

    Private Sub dgvSupp_Item_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.VisibleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.VisibleChanged")
    End Sub

    Private Sub dgvSupp_Item_AutoSizeColumnModeChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewAutoSizeColumnModeEventArgs) Handles dgvSupp_Item.AutoSizeColumnModeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AutoSizeColumnModeChanged")
    End Sub

    Private Sub dgvSupp_Item_AutoSizeColumnsModeChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewAutoSizeColumnsModeEventArgs) Handles dgvSupp_Item.AutoSizeColumnsModeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AutoSizeColumnsModeChanged")
    End Sub

    Private Sub dgvSupp_Item_AutoSizeRowsModeChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewAutoSizeModeEventArgs) Handles dgvSupp_Item.AutoSizeRowsModeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AutoSizeRowsModeChanged")
    End Sub

    Private Sub dgvSupp_Item_BackgroundColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.BackgroundColorChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.BackgroundColorChanged")
    End Sub

    Private Sub dgvSupp_Item_BindingContextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.BindingContextChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.BindingContextChanged")
    End Sub

    Private Sub dgvSupp_Item_BorderStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.BorderStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.BorderStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_CellBorderStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.CellBorderStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellBorderStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_CellContextMenuStripChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellContextMenuStripChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellContextMenuStripChanged")
    End Sub

    Private Sub dgvSupp_Item_CellContextMenuStripNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellContextMenuStripNeededEventArgs) Handles dgvSupp_Item.CellContextMenuStripNeeded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellContextMenuStripNeeded")
    End Sub

    Private Sub dgvSupp_Item_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSupp_Item.CellMouseClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellMouseClick")
    End Sub

    Private Sub dgvSupp_Item_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSupp_Item.CellMouseDoubleClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellMouseDoubleClick")
    End Sub

    Private Sub dgvSupp_Item_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSupp_Item.CellMouseDown
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellMouseDown")
    End Sub

    Private Sub dgvSupp_Item_CellMouseEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellMouseEnter
        'If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellMouseEnter")
    End Sub

    Private Sub dgvSupp_Item_CellMouseLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSupp_Item.CellMouseLeave
        'If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellMouseLeave")
    End Sub

    Private Sub dgvSupp_Item_CellMouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSupp_Item.CellMouseUp
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.CellMouseUp")
    End Sub

    Private Sub dgvSupp_Item_ColumnAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnAdded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnAdded")
    End Sub

    Private Sub dgvSupp_Item_ColumnContextMenuStripChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnContextMenuStripChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnContextMenuStripChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnDataPropertyNameChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnDataPropertyNameChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnDataPropertyNameChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnDefaultCellStyleChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnDefaultCellStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnDefaultCellStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnDisplayIndexChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnDisplayIndexChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnDisplayIndexChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnDividerDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnDividerDoubleClickEventArgs) Handles dgvSupp_Item.ColumnDividerDoubleClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnDividerDoubleClick")
    End Sub

    Private Sub dgvSupp_Item_ColumnDividerWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnDividerWidthChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnDividerWidthChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnHeaderCellChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnHeaderCellChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnHeaderCellChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSupp_Item.ColumnHeaderMouseClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnHeaderMouseClick")
    End Sub

    Private Sub dgvSupp_Item_ColumnHeaderMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSupp_Item.ColumnHeaderMouseDoubleClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnHeaderMouseDoubleClick")
    End Sub

    Private Sub dgvSupp_Item_ColumnHeadersBorderStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ColumnHeadersBorderStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnHeadersBorderStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnHeadersDefaultCellStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ColumnHeadersDefaultCellStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnHeadersDefaultCellStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnHeadersHeightChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ColumnHeadersHeightChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnHeadersHeightChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnHeadersHeightSizeModeChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewAutoSizeModeEventArgs) Handles dgvSupp_Item.ColumnHeadersHeightSizeModeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnHeadersHeightSizeModeChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnMinimumWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnMinimumWidthChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnMinimumWidthChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnNameChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnNameChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnNameChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnRemoved
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnRemoved")
    End Sub

    Private Sub dgvSupp_Item_ColumnSortModeChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnSortModeChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnSortModeChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnStateChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnStateChangedEventArgs) Handles dgvSupp_Item.ColumnStateChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnStateChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnToolTipTextChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnToolTipTextChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnToolTipTextChanged")
    End Sub

    Private Sub dgvSupp_Item_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvSupp_Item.ColumnWidthChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ColumnWidthChanged")
    End Sub

    Private Sub dgvSupp_Item_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles dgvSupp_Item.DragDrop
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DragDrop")
    End Sub

    Private Sub dgvSupp_Item_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles dgvSupp_Item.DragEnter
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DragEnter")
    End Sub

    Private Sub dgvSupp_Item_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.DragLeave
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DragLeave")
    End Sub

    Private Sub dgvSupp_Item_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles dgvSupp_Item.DragOver
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.DragOver")
    End Sub

    Private Sub dgvSupp_Item_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.FontChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.FontChanged")
    End Sub

    Private Sub dgvSupp_Item_ForeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.ForeColorChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.ForeColorChanged")
    End Sub

    Private Sub dgvSupp_Item_GridColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.GridColorChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.GridColorChanged")
    End Sub

    Private Sub dgvSupp_Item_MarginChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.MarginChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MarginChanged")
    End Sub

    Private Sub dgvSupp_Item_MouseCaptureChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.MouseCaptureChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MouseCaptureChanged")
    End Sub

    Private Sub dgvSupp_Item_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvSupp_Item.MouseClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MouseClick")
    End Sub

    Private Sub dgvSupp_Item_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvSupp_Item.MouseDoubleClick
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MouseDoubleClick")
    End Sub

    Private Sub dgvSupp_Item_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvSupp_Item.MouseDown
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MouseDown")
    End Sub

    Private Sub dgvSupp_Item_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.MouseEnter
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MouseEnter")
    End Sub

    Private Sub dgvSupp_Item_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.MouseLeave
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MouseLeave")
    End Sub

    Private Sub dgvSupp_Item_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvSupp_Item.MouseUp
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MouseUp")
    End Sub

    Private Sub dgvSupp_Item_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvSupp_Item.MouseWheel
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.MouseWheel")
    End Sub

    Private Sub dgvSupp_Item_RowContextMenuStripChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.RowContextMenuStripChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowContextMenuStripChanged")
    End Sub

    Private Sub dgvSupp_Item_RowContextMenuStripNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowContextMenuStripNeededEventArgs) Handles dgvSupp_Item.RowContextMenuStripNeeded
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowContextMenuStripNeeded")
    End Sub

    Private Sub dgvSupp_Item_RowDefaultCellStyleChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSupp_Item.RowDefaultCellStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.RowDefaultCellStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_AllowUserToOrderColumnsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.AllowUserToOrderColumnsChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AllowUserToOrderColumnsChanged")
    End Sub

    Private Sub dgvSupp_Item_AllowUserToResizeColumnsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.AllowUserToResizeColumnsChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AllowUserToResizeColumnsChanged")
    End Sub

    Private Sub dgvSupp_Item_AllowUserToResizeRowsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.AllowUserToResizeRowsChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AllowUserToResizeRowsChanged")
    End Sub

    Private Sub dgvSupp_Item_AlternatingRowsDefaultCellStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.AlternatingRowsDefaultCellStyleChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AlternatingRowsDefaultCellStyleChanged")
    End Sub

    Private Sub dgvSupp_Item_AllowUserToDeleteRowsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.AllowUserToDeleteRowsChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AllowUserToDeleteRowsChanged")
    End Sub

    Private Sub dgvSupp_Item_AutoGenerateColumnsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSupp_Item.AutoGenerateColumnsChanged
        If m_ShowEvents Then Debug.Print("dgvSupp_Item.AutoGenerateColumnsChanged")
    End Sub

#End Region

#Region "dgvSupp_Item ORDER OF EVENTS #########################################################"

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_AutoSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.AutoSizeChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.AutoSizeChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_AutoValidateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.AutoValidateChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.AutoValidateChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_BackColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.BackColorChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.BackColorChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_BackgroundImageChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.BackgroundImageChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.BackgroundImageChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_BackgroundImageLayoutChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.BackgroundImageLayoutChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.BackgroundImageLayoutChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_BindingContextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.BindingContextChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.BindingContextChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_CausesValidationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.CausesValidationChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.CausesValidationChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ChangeUICues(ByVal sender As Object, ByVal e As System.Windows.Forms.UICuesEventArgs) Handles Me.ChangeUICues
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ChangeUICues")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Click")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ClientSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ClientSizeChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ClientSizeChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ContextMenuChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ContextMenuChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ContextMenuChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ContextMenuStripChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ContextMenuStripChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ContextMenuStripChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ControlAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs) Handles Me.ControlAdded
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ControlAdded")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ControlRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs) Handles Me.ControlRemoved
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ControlRemoved")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_CursorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.CursorChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.CursorChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_DockChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DockChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.DockChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DoubleClick
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.DoubleClick")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.DragDrop")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.DragEnter")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DragLeave
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.DragLeave")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragOver
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.DragOver")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.EnabledChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.EnabledChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.FontChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.FontChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ForeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ForeColorChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ForeColorChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_GiveFeedback(ByVal sender As Object, ByVal e As System.Windows.Forms.GiveFeedbackEventArgs) Handles Me.GiveFeedback
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.GiveFeedback")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.GotFocus")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_HandleCreated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.HandleCreated
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.HandleCreated")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_HandleDestroyed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.HandleDestroyed
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.HandleDestroyed")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_HelpButtonClicked(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.HelpButtonClicked
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.HelpButtonClicked")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_HelpRequested(ByVal sender As Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles Me.HelpRequested
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.HelpRequested")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ImeModeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ImeModeChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ImeModeChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_InputLanguageChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.InputLanguageChangedEventArgs) Handles Me.InputLanguageChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.InputLanguageChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_InputLanguageChanging(ByVal sender As Object, ByVal e As System.Windows.Forms.InputLanguageChangingEventArgs) Handles Me.InputLanguageChanging
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.InputLanguageChanging")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Invalidated(ByVal sender As Object, ByVal e As System.Windows.Forms.InvalidateEventArgs) Handles Me.Invalidated
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Invalidated")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.KeyDown")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.KeyPress")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.KeyUp")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles Me.Layout
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Layout")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Leave")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.LocationChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MaximizedBoundsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MaximizedBoundsChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MaximizedBoundsChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MaximumSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MaximumSizeChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MaximumSizeChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MdiChildActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MdiChildActivate
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MdiChildActivate")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MenuComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MenuComplete
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MenuComplete")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MenuStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MenuStart
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MenuStart")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MinimumSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MinimumSizeChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MinimumSizeChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MouseCaptureChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseCaptureChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MouseCaptureChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MouseClick")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDoubleClick
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MouseDoubleClick")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MouseDown")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseEnter
        'If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MouseEnter")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseLeave
        'If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MouseLeave")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseHover
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MouseHover")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MouseUp")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.MouseWheel")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Move")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_PaddingChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PaddingChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.PaddingChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Paint")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ParentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ParentChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ParentChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles Me.PreviewKeyDown
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.PreviewKeyDown")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_QueryAccessibilityHelp(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryAccessibilityHelpEventArgs) Handles Me.QueryAccessibilityHelp
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.QueryAccessibilityHelp")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs) Handles Me.QueryContinueDrag
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.QueryContinueDrag")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_RegionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.RegionChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.RegionChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Resize")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ResizeBegin(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeBegin
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ResizeBegin")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.ResizeEnd")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_RightToLeftChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.RightToLeftChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.RightToLeftChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_RightToLeftLayoutChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.RightToLeftLayoutChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.RightToLeftLayoutChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles Me.Scroll
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Scroll")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Shown")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.SizeChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_StyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.StyleChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.StyleChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_SystemColorsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SystemColorsChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.SystemColorsChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.TextChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.TextChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Validated
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Validated")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.VisibleChanged")
    End Sub

    Private Sub frmMDB_CUS_DEC_SUPP_ITEM_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If m_ShowEvents Then Debug.Print("frmMDB_CUS_DEC_SUPP_ITEM.Activated")
    End Sub




#End Region

    Public Property Cus_Dec_ID() As Integer
        Get
            Cus_Dec_ID = m_Cus_Dec_ID
        End Get
        Set(ByVal value As Integer)
            m_Cus_Dec_ID = value
        End Set
    End Property

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        'dgvSupp_Item.

    End Sub

    Private Sub ProductLineInsert()

        Try

            If dgvSupp_Item.Rows.Count = 0 Then

                Call LineInsert()

            ElseIf Not (dgvSupp_Item.Rows(dgvSupp_Item.Rows.Count - 1).Cells(Columns.Item_No).Value.Equals(DBNull.Value)) Then

                If Not Trim(dgvSupp_Item.Rows(dgvSupp_Item.Rows.Count - 1).Cells(Columns.Item_No).Value).Equals("") Then
                    Call LineInsert()
                Else
                    dgvSupp_Item.CurrentCell = dgvSupp_Item.Rows(dgvSupp_Item.Rows.Count - 1).Cells(Columns.Item_No)
                End If

            Else

                dgvSupp_Item.CurrentCell = dgvSupp_Item.Rows(dgvSupp_Item.Rows.Count - 1).Cells(Columns.Item_No)

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LineInsert()

        Try
            Dim drNewRow As DataRow
            dgvSupp_Item.AllowUserToAddRows = True

            drNewRow = dtGrid.NewRow

            dtGrid.Rows.Add(drNewRow)

            dgvSupp_Item.AllowUserToAddRows = False
            dgvSupp_Item.CurrentCell = dgvSupp_Item.Rows(dgvSupp_Item.Rows.Count - 1).Cells(Columns.Item_No)

            dgvSupp_Item.Focus()
            'dgvItems.BeginEdit(True)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Enum Columns
        'ColHeader
        Cus_Dec_Supp_Item_ID
        Cus_Dec_ID
        Item_No
        Qty_Ordered
        Unit_Price
        User_Login
        Update_TS
    End Enum

End Class