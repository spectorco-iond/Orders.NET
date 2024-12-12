Public Class cNotes

    Private oform As frmNotes

    Public Sub New()

        Try
            oform = New frmNotes

        Catch er As Exception
            MsgBox("Error in CNotes." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub New(ByVal pstrNotesElement As String, ByVal pstrNotesValue As String)

        Try
            oform = New frmNotes
            oform.NotesElement = pstrNotesElement
            oform.NotesValue = pstrNotesValue

            oform.ShowDialog()

        Catch er As Exception
            MsgBox("Error in CNotes." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Show()

        Try
            oform.ShowDialog()

        Catch er As Exception
            MsgBox("Error in CNotes." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Property NotesElement() As String
        Get
            NotesElement = oform.NotesElement
        End Get
        Set(ByVal value As String)
            oform.NotesElement = value
        End Set
    End Property

    Public ReadOnly Property Form() As frmNotes
        Get
            Form = oform
        End Get
    End Property

End Class
