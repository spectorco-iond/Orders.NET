
Friend Class cSearch

    Private oform As frmSearch

    Public Sub New()

        Try
            oform = New frmSearch

        Catch er As Exception
            MsgBox("Error in CSearch." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub New(ByVal pstrSearchElement As String, Optional ByVal pblnShow As Boolean = True, Optional ByVal pblnAll As Boolean = False)

        Try
            oform = New frmSearch
            oform.SearchElement = pstrSearchElement
            If pblnShow Then
                oform.ShowDialog()
            Else
                oform.GetData(pblnAll)
            End If
        Catch er As Exception
            MsgBox("Error in CSearch." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        'frmSearch.SearchElement = pstrElement
        'frmSearch.ShowDialog()

    End Sub

    Public Sub New(ByVal pstrSearchElement As String, ByVal pstrSearchFromWindow As String)

        Try
            oform = New frmSearch
            oform.SearchElement = pstrSearchElement
            oform.SearchFromWindow = pstrSearchFromWindow
            oform.ShowDialog()
        Catch er As Exception
            MsgBox("Error in CSearch." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        'frmSearch.SearchElement = pstrElement
        'frmSearch.ShowDialog()

    End Sub

    Public Sub GetResults(Optional ByVal pblnAll As Boolean = False)

        Dim ipos As Integer = 0

        Try
            ipos = 1
            Call oform.GetData(pblnAll)

            ipos = 2
            If oform.DataGrid.Rows.Count >= 1 Then
                ipos = 3
                oform.FoundElementIndex = oform.DataGrid.Rows(0).Cells(1).Value ' oform.TextMatrix(1, 1)
                ipos = 4
                If oform.DataGrid.ColumnCount >= 4 Then
                    oform.FoundElementValue = oform.DataGrid.Rows(0).Cells(3).Value ' oform.TextMatrix(1, 1)
                Else
                    oform.FoundElementIndex = oform.DataGrid.Rows(0).Cells(1).Value ' oform.TextMatrix(1, 1)
                End If
                oform.FoundRow = 1
            Else
                oform.FoundElementIndex = ""
                oform.FoundElementValue = ""
                oform.FoundRow = -1
            End If

                ipos = 5

        Catch er As Exception
            MsgBox("Error in CSearch." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " Pos: " & ipos.ToString)
        End Try

    End Sub

    Public Sub Search()

        Try
            oform.GetData()

        Catch er As Exception
            MsgBox("Error in CSearch." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SetDefaultSearchElement(ByVal pstrField As String, ByVal pstrValue As String, ByVal pstrType As String)

        Try
            oform.SetDefaultSearchElement(pstrField, pstrValue, pstrType)

        Catch er As Exception
            MsgBox("Error in CSearch." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Show()

        Try
            oform.ShowDialog()

        Catch er As Exception
            MsgBox("Error in CSearch." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Dispose()

        Try
            oform.Dispose()
        Catch er As Exception
        End Try

    End Sub

    Public Function FoundElement() As Object

        FoundElement = oform.FoundElementValue

    End Function

    Public Function FoundElementIndex() As Object

        FoundElementIndex = oform.FoundElementIndex

    End Function

    Public Property SearchElement() As String
        Get
            SearchElement = oform.SearchElement
        End Get
        Set(ByVal value As String)
            oform.SearchElement = value
        End Set
    End Property

    Public ReadOnly Property Form() As frmSearch
        Get
            Form = oform
        End Get
    End Property

End Class

