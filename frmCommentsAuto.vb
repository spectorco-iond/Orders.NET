Public Class frmCommentsAuto

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       VARIABLE DECLARATIONS             ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private m_OEI_Ord_No As String
    Private m_Ord_Guid As String

    Private m_NbVars As Integer
    Private m_Var1 As String
    Private m_Var2 As String
    Private m_Var3 As String
    Private m_Var4 As String
    Private m_Var5 As String
    Private m_Description As String



    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(ByVal pstrOrd_Guid As String, ByVal pstrOEI_Ord_No As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        m_Ord_Guid = pstrOrd_Guid
        m_OEI_Ord_No = pstrOEI_Ord_No

        Call SetLanguage(m_Ord_Guid)

    End Sub

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       FORM LOAD / UNLOAD PROCEDURES      //////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '   FORM LOAD
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmOEComments_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Try
            'lblVersion.Text = gVersion

            txtOEI_Ord_No.Text = m_OEI_Ord_No ' gOrderNo
            txtCmp_Name.Text = m_oOrder.Ordhead.Customer.cmp_name

            Call LoadCategories()

            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
            'mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub cmdAddComment_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddComment.Click

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            ' Make sure all the needed vars are present
            If cboCategory.Text = "" Then
                MsgBox("Please select a category")
                Exit Sub
            End If

            If cboLang.Text = "" Then
                MsgBox("Please select a Language")
                Exit Sub
            End If

            If cboComments.Text = "" Then
                MsgBox("Please select the Comment to add")
                Exit Sub
            End If

            If m_NbVars > 0 Then
                If Not txtVar1.Visible Or txtVar1.Text = "" Then
                    MsgBox("Please enter a value for " & m_Var1)
                    Exit Sub
                End If
            End If

            If m_NbVars > 1 Then
                If Not txtVar2.Visible Or txtVar2.Text = "" Then
                    MsgBox("Please enter a value for " & m_Var2)
                    Exit Sub
                End If
            End If

            If m_NbVars > 2 Then
                If Not txtVar3.Visible Or txtVar3.Text = "" Then
                    MsgBox("Please enter a value for " & m_Var3)
                    Exit Sub
                End If
            End If

            If m_NbVars > 3 Then
                If Not txtVar4.Visible Or txtVar4.Text = "" Then
                    MsgBox("Please enter a value for " & m_Var4)
                    Exit Sub
                End If
            End If

            If m_NbVars > 4 Then
                If Not txtVar5.Visible Or txtVar5.Text = "" Then
                    MsgBox("Please enter a value for " & m_Var5)
                    Exit Sub
                End If
            End If

            ' Insert OE Comments into oelincmt_sql table
            Call InsertComment(SqlCompliantString((Replace(txtFinalMsg.Text, vbCrLf, ""))))

            MsgBox(" You have successfully added a comment to the order ")

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click, cmdClose.Click

        Try

            Me.Close()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    'UPGRADE_WARNING: Event cboCategory.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboCategory_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCategory.SelectedIndexChanged ' , cboCategory.SelectedIndexChanged

        Try
            Call LoadCatComments()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub



    'UPGRADE_WARNING: Event cboComments.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboComments_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboComments.SelectedIndexChanged, cboComments.SelectedIndexChanged

        Try
            Call LoadCommentInfo()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cboLang_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLang.SelectedIndexChanged, cboLang.SelectedIndexChanged

        Try
            Call LoadCatComments()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtVar_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVar1.Leave, txtVar2.Leave, txtVar3.Leave, txtVar4.Leave, txtVar5.Leave

        Try
            Dim txtElement As New TextBox
            Dim strNewText As String
            Dim strVarValue As String = ""

            txtElement = DirectCast(eventSender, TextBox)
            strNewText = m_Description

            If txtElement.Visible Then

                For i As Integer = 1 To m_NbVars

                    Select Case i
                        Case 1
                            strVarValue = m_Var1
                            txtElement = txtVar1
                        Case 2
                            strVarValue = m_Var2
                            txtElement = txtVar2
                        Case 3
                            strVarValue = m_Var3
                            txtElement = txtVar3
                        Case 4
                            strVarValue = m_Var4
                            txtElement = txtVar4
                        Case 5
                            strVarValue = m_Var5
                            txtElement = txtVar5
                    End Select

                    If txtElement.Text <> "" And Not (strVarValue.Equals(DBNull.Value)) Then
                        strNewText = Replace(strNewText, strVarValue, txtElement.Text)
                    End If

                Next i

            End If

            Call CommentLineBreakDown(strNewText)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadCategories()

        Dim strSql As String
        Dim dtCategories As DataTable
        Dim db As New cDBA

        Try
            strSql = _
            "SELECT     *   " & _
            "FROM       OE_Comments_Categories WITH (Nolock) " & _
            "ORDER BY   CatName "

            dtCategories = db.DataTable(strSql)

            If dtCategories.Rows.Count <> 0 Then
                For Each drRow As DataRow In dtCategories.Rows

                    cboCategory.Items.Add(New VB6.ListBoxItem(Trim(drRow.Item("catName").ToString), drRow.Item("id").ToString))

                Next drRow

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadCatComments()

        Dim strSql As String
        Dim dt As DataTable
        Dim db As New cDBA

        Try
            If cboCategory.Text <> "" And cboLang.Text <> "" Then

                strSql = _
                "SELECT     Id, Name " & _
                "FROM       OE_Comments_Messages WITH (Nolock) " & _
                "WHERE      Cat_Id = '" & VB6.GetItemData(cboCategory, cboCategory.SelectedIndex) & "' " & " AND " & _
                "           Lang = '" & Trim(cboLang.Text) & "' " & _
                "ORDER BY   Name "

                dt = db.DataTable(strSql)

                cboComments.Items.Clear()

                If dt.Rows.Count <> 0 Then

                    For Each drRow As DataRow In dt.Rows

                        cboComments.Items.Add(New VB6.ListBoxItem(Trim(drRow.Item("Name").ToString), drRow.Item("id").ToString))

                    Next drRow

                End If

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadCommentInfo()

        Dim strSql As String
        Dim dt As DataTable
        Dim db As New cDBA

        Try

            ' Clear any previous comment info
            Call ClearCommentInfo()

            strSql = _
            "SELECT     Description, NbVars, Var1, Var2, Var3, Var4, Var5 " & _
            "FROM       OE_Comments_Messages WITH (Nolock) " & _
            "WHERE      Id = '" & VB6.GetItemData(cboComments, cboComments.SelectedIndex) & "' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                m_Description = dt.Rows(0).Item("Description").ToString
                m_Var1 = dt.Rows(0).Item("Var1").ToString
                m_Var2 = dt.Rows(0).Item("Var2").ToString
                m_Var3 = dt.Rows(0).Item("Var3").ToString
                m_Var4 = dt.Rows(0).Item("Var4").ToString
                m_Var5 = dt.Rows(0).Item("Var5").ToString
                m_NbVars = dt.Rows(0).Item("NbVars")

                Call CommentLineBreakDown(m_Description)

                If m_NbVars > 0 Then
                    lblNoVars.Visible = False
                    lblVar1.Visible = True
                    lblVar1.Text = m_Var1
                    txtVar1.Text = ""
                    txtVar1.Visible = True
                End If

                If m_NbVars > 1 Then
                    lblVar2.Visible = True
                    lblVar2.Text = m_Var2
                    txtVar2.Text = ""
                    txtVar2.Visible = True
                End If

                If m_NbVars > 2 Then
                    lblVar3.Visible = True
                    lblVar3.Text = m_Var3
                    txtVar3.Text = ""
                    txtVar3.Visible = True
                End If

                If m_NbVars > 3 Then
                    lblVar4.Visible = True
                    lblVar4.Text = m_Var4
                    txtVar4.Text = ""
                    txtVar4.Visible = True
                End If

                If m_NbVars > 4 Then
                    lblVar5.Visible = True
                    lblVar5.Text = m_Var5
                    txtVar5.Text = ""
                    txtVar5.Visible = True
                End If

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ClearCommentInfo()

        Try
            txtFinalMsg.Text = ""

            lblNoVars.Visible = True

            lblVar1.Visible = False
            lblVar1.Text = "TXT_VAR_1"
            txtVar1.Text = ""
            txtVar1.Visible = False

            lblVar2.Visible = False
            lblVar2.Text = "TXT_VAR_2"
            txtVar2.Text = ""
            txtVar2.Visible = False

            lblVar3.Visible = False
            lblVar3.Text = "TXT_VAR_3"
            txtVar3.Text = ""
            txtVar3.Visible = False

            lblVar4.Visible = False
            lblVar4.Text = "TXT_VAR_4"
            txtVar4.Text = ""
            txtVar4.Visible = False

            lblVar5.Visible = False
            lblVar5.Text = "TXT_VAR_5"
            txtVar5.Text = ""
            txtVar5.Visible = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub



    Private Sub CommentLineBreakDown(ByRef strComment As String)

        Try
            Dim lastSpace As Short
            Dim nextSpace As Short
            Dim strLen As Short
            Dim lastLineFinish As Short

            lastSpace = 0
            lastLineFinish = 0
            nextSpace = -1
            strLen = Len(strComment)

            txtFinalMsg.Text = ""

            ' PLEASE VERIFY REPEAT NUMBER AS THE NUMBER ON THE ORDER IS NOT IN OUR SYSTEM. THANK YOU.

            Do While nextSpace <> 0

                nextSpace = InStr(lastSpace + 1, strComment, " ")
                'MsgBox nextSpace
                If nextSpace - lastLineFinish > 40 Then
                    'MsgBox Mid(strComment, (lastLineFinish + 1), (lastSpace - lastLineFinish))
                    txtFinalMsg.Text = txtFinalMsg.Text & Mid(strComment, lastLineFinish + 1, lastSpace - lastLineFinish) & vbCrLf

                    lastLineFinish = lastSpace

                End If

                If nextSpace <> 0 Then
                    lastSpace = nextSpace
                End If

            Loop
            'MsgBox lastSpace

            ' Finish off last portion
            If Len(Mid(strComment, lastLineFinish + 1, strLen - lastLineFinish)) > 40 Then
                txtFinalMsg.Text = txtFinalMsg.Text & Mid(strComment, lastLineFinish + 1, lastSpace - lastLineFinish) & vbCrLf
                txtFinalMsg.Text = txtFinalMsg.Text & Mid(strComment, lastSpace + 1, strLen - lastSpace)
            Else
                txtFinalMsg.Text = txtFinalMsg.Text & Mid(strComment, lastLineFinish + 1, strLen - lastLineFinish)
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub InsertComment(ByRef strComment As String)

        Try
            Dim i As Short
            Dim lastSpace As Short
            Dim nextSpace As Short
            Dim strLen As Short
            Dim lastLineFinish As Short
            Dim textToInsert As String

            lastSpace = 0
            lastLineFinish = 0
            nextSpace = -1
            strLen = Len(strComment)

            ' *** Adding Comment
            textToInsert = "**********************************"
            Call InsertCommentLine(textToInsert)
            i = i + 1

            Do While nextSpace <> 0

                nextSpace = InStr(lastSpace + 1, strComment, " ")

                If nextSpace - lastLineFinish > 40 Then

                    textToInsert = Mid(strComment, lastLineFinish + 1, lastSpace - lastLineFinish)

                    ' Insert Comment (textToInsert)
                    Call InsertCommentLine(textToInsert)
                    i = i + 1


                    lastLineFinish = lastSpace

                End If

                If nextSpace <> 0 Then
                    lastSpace = nextSpace
                End If

            Loop

            ' Finish off last portion
            If Len(Mid(strComment, lastLineFinish + 1, strLen - lastLineFinish)) > 40 Then

                textToInsert = Mid(strComment, lastLineFinish + 1, lastSpace - lastLineFinish)
                Call InsertCommentLine(textToInsert)
                i = i + 1

                textToInsert = Mid(strComment, lastSpace + 1, strLen - lastSpace)
                Call InsertCommentLine(textToInsert)
                i = i + 1

            Else
                textToInsert = Mid(strComment, lastLineFinish + 1, strLen - lastLineFinish)
                Call InsertCommentLine(textToInsert)
                i = i + 1
            End If

            Dim db As New cDBA

            Dim strSql As String = _
            "UPDATE OEI_LinCmt " & _
            "SET Line_Seq_No = ID " & _
            "WHERE Ord_Guid = '" & m_Ord_Guid & "' AND Line_Seq_No IS NULL "

            db.Execute(strSql)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub InsertCommentLine(ByRef comment As String)

        If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

        Try

            Dim db As New cDBA

            Dim strSql As String = _
            " INSERT INTO OEI_LinCmt (Ord_Guid, Cmt) " & _
            " VALUES ( '" & m_Ord_Guid & "', '" & SqlCompliantString(comment) & "') "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SetLanguage(ByVal pOrd_Guid As String)

        Try

            Dim db As New cDBA()
            Dim dt As DataTable
            Dim strsql As String =
            "SELECT		OC.Ord_Guid, OC.DefContact, C.TaalCode " &
            "FROM		OEI_Order_Contacts OC WITH (Nolock) " &
            "INNER JOIN	Cicntp C WITH (Nolock) ON OC.DefContact = C.ID " &
            "WHERE	ISNULL(C.active_y,0) = 1 And OC.ord_guid = '" & pOrd_Guid & "' AND OC.ContactType = 0 "

            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 
            dt = db.DataTable(strsql)
            If dt.Rows.Count <> 0 Then
                cboLang.Text = dt.Rows(0).Item("TaalCode").ToString.Trim
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

End Class