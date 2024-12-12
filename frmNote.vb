Public Class frmNote

    Private m_Note As New cNote

    Public Property Note() As cNote
        Get
            Note = m_Note
        End Get
        Set(ByVal value As cNote)
            m_Note = value
            Call Fill()
        End Set
    End Property

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            m_Note.Note_Topic = txtNote_Topic.Text
            m_Note.Notes = txtNotes.Text

            m_Note.Save()

            Me.Close()

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Fill()

        txtNote_Dt.Text = m_Note.Note_Dt.ToString
        txtUser_Name.Text = m_Note.User_Name
        txtNote_Topic.Text = m_Note.Note_Topic
        txtNotes.Text = m_Note.Notes

    End Sub

End Class