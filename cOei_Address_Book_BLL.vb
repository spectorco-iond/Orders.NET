Public Class cOei_Address_Book_BUS



        ' Handle to the Oei_Address_Book DBAccess class
        Private oOei_Address_Book_DA As cOei_Address_Book_DAL = Nothing


        Public Sub New()
                oOei_Address_Book_DA = New cOei_Address_Book_DAL()
        End Sub


        ' Get element with ID specified from the database 
        Public Function Load(ByVal piId As Integer) As cOei_Address_Book
                Return oOei_Address_Book_DA.Load(piId)
        End Function


    ' Get element with ID specified from the database 
    Public Function LoadAll() As DataTable
        Return oOei_Address_Book_DA.LoadAll()
    End Function

        ' Insert/Update current element in the database 
        Public Function Save(ByRef oOei_Address_Book As cOei_Address_Book) As Boolean
                Return oOei_Address_Book_DA.Save(oOei_Address_Book)
        End Function


        ' Delete element with ID specified from the database 
        Public Function Delete(ByVal piId As Integer) As Boolean
                Return oOei_Address_Book_DA.Delete(piId)
        End Function


        ' Validate current element 
        Public Function Validate(ByRef oOei_Address_Book As cOei_Address_Book) As Boolean
            Return True 
        End Function

End Class ' cOei_Address_Book


