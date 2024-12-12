Public Class cComments

    Private oform As frmComments

    Public Sub New(ByVal pstrOrd_Guid As String)

        Try
            ' open form
            oform = New frmComments(pstrOrd_Guid)

            ' show me the money
            oform.ShowDialog()

            oform.Dispose()

        Catch er As Exception
            MsgBox("Error in cComments." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public ReadOnly Property Form() As frmComments
        Get
            Form = oform
        End Get
    End Property

End Class
