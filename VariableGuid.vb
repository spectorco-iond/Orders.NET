Public Class VariableGuid

    Public Function Guid(ByVal piLength As Integer)

        Guid = Mid(System.Guid.NewGuid.ToString, 1, piLength).ToUpper

    End Function

End Class
