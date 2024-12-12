Module modHV_BASICS

    Public gsDataBaseName As String
    Public gsServerName As String
    Public Function sqlSAFE(sValue) As String
        sValue = Replace(sValue, "'", "''")
        Return sValue
    End Function
End Module
