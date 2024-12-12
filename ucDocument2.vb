Option Strict Off
Option Explicit On

Friend Class ucDocument2
    Inherits System.Windows.Forms.UserControl

    Private Sub lbTypes_DragEnter(ByVal sender As Object, ByVal e As  _
                System.Windows.Forms.DragEventArgs) Handles lbTypes.DragEnter

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If

    End Sub

    'Private Sub lbTypes_DragDrop(ByVal sender As Object, ByVal e As  _
    '            System.Windows.Forms.DragEventArgs) Handles lbTypes.DragDrop

    '    If e.Data.GetDataPresent(DataFormats.FileDrop) Then

    '        Dim MyFiles() As String

    '        ' Assign the files to an array.
    '        MyFiles = e.Data.GetData(DataFormats.FileDrop)

    '        AddDocumentFiles(MyFiles)

    '    End If

    'End Sub

    'Private Sub AddDocumentFiles(ByRef Files As String())

    '    ' Loop through the array and add the files to the list.
    '    For i As Integer = 0 To Files.Length - 1

    '        dgvFileList.Rows.Add()
    '        dgvFileList.Rows(dgvFileList.Rows.Count - 1).Cells(1).Value = Files(i).ToString

    '        'AddDocumentFiles(MyFiles)
    '        'lbTypes.Items.Add(MyFiles(i))
    '    Next

    'End Sub

End Class
