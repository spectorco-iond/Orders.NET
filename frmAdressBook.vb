Public Class frmAddressBook
    Dim row, cell As Integer
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        refreshDGV()
    End Sub
 
    Private Sub Form1_FormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        save()
    End Sub

    Private Sub AddToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles AddToolStripMenuItem1.Click
        Dim Address_Book As cOei_Address_Book = New cOei_Address_Book
        Dim Address_Book_BUS As cOei_Address_Book_BUS = New cOei_Address_Book_BUS
        Dim c As cDBA = New cDBA
        Dim value As Integer
        save()
        value = Convert.ToInt32(dgvUsers.Rows(dgvUsers.Rows.Count - 1).Cells(0).Value)
        Address_Book.addrBook_Id = value + 1
        Address_Book.addrBook_Fullname = ""
        Address_Book.addrBook_Mail = ""
        Address_Book_BUS.Save(Address_Book)
        refreshDGV()
        'Focus on last row added.
        dgvUsers.Focus()
        dgvUsers.CurrentCell = dgvUsers(1, dgvUsers.Rows.Count - 1)
        dgvUsers.BeginEdit(True)
    End Sub

    Private Sub DeleteToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles DeleteToolStripMenuItem1.Click
        ' Set the SelectionMode to select multiple items.
        dgvUsers.EndEdit(True)
        dgvUsers.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        dgvUsers.Rows(row).Cells(1).ReadOnly = False
        dgvUsers.Rows(row).Selected = True

        If MessageBox.Show("Do you want to Delete the user '" + Trim(Me.dgvUsers.SelectedCells(1).Value) + "' ?", "Title", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Dim A As cOei_Address_Book_BUS = New cOei_Address_Book_BUS
            A.Delete(Convert.ToInt32(Me.dgvUsers.SelectedCells(0).Value))
        End If
        refreshDGV()
    End Sub

    Public Sub refreshDGV()
        Dim A As cOei_Address_Book_BUS = New cOei_Address_Book_BUS
        dgvUsers.DataSource = A.LoadAll()
        configureproperties()
    End Sub

    Public Sub save()
        Dim i As Integer
        Dim Address_Book As cOei_Address_Book = New cOei_Address_Book
        Dim Address_Book_BUS As cOei_Address_Book_BUS = New cOei_Address_Book_BUS

        For i = 0 To dgvUsers.Rows.Count - 1
            Address_Book.addrBook_Id = dgvUsers.Rows(i).Cells(0).Value
            Address_Book.addrBook_Fullname = dgvUsers.Rows(i).Cells(1).Value
            Address_Book.addrBook_Mail = dgvUsers.Rows(i).Cells(2).Value
            Address_Book_BUS.Save(Address_Book)
        Next i
    End Sub

    Public Sub configureproperties()
        dgvUsers.Columns(0).Visible = False
        dgvUsers.RowHeadersVisible = False
    End Sub

    Private Sub dgvUsers_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUsers.CellClick
        cell = e.ColumnIndex
        row = e.RowIndex
        dgvUsers.BeginEdit(True)
    End Sub
End Class