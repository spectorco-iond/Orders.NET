Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Public Class cImage

    Public Sub New()

    End Sub

    Public Sub SetImage(ByRef pTable As DataTable, ByVal piRow As Integer, ByVal pstrCellName As String, ByRef oImage As Image)

        SetImage(pTable.Rows(piRow), pstrCellName, oImage)

    End Sub

    Public Sub SetImage(ByRef pRow As DataRow, ByVal pstrCellName As String, ByRef oImage As Image)

        SetImage(pRow.Item(pstrCellName), oImage)

    End Sub

    Public Sub SetImage(ByRef pCell As Object, ByRef oImage As Image)
        'Public Sub SetImage(ByRef pCell As cell, ByRef oImage As Image)

        Try
            'Dim oBmp As Bitmap
            Dim baImage As Byte() = Nothing

            ' Dim imageData As Byte() = DirectCast(cmd.ExecuteScalar(), Byte())
            Dim imageData As Byte() = DirectCast(pCell, Byte())
            If Not imageData Is Nothing Then
                Using ms As New MemoryStream(imageData, 0, imageData.Length)
                    ms.Write(imageData, 0, imageData.Length)
                    'oBmp = New Bitmap(ms)
                    'oImage = oBmp
                    oImage = Image.FromStream(ms, True)
                    'PictureBox1.BackgroundImage = Image.FromStream(ms, True)
                End Using
            End If

            '' '' ''SetImageAsByteArray(pCell, baImage)

            '' '' ''Dim oStream As MemoryStream = Nothing

            ' '' '' ''baImage = CType(pCell, Byte())

            '' '' ''Dim formatter As New BinaryFormatter()
            '' '' ''Dim mem As New IO.MemoryStream
            '' '' ''formatter.Serialize(mem, pCell)
            '' '' ''baImage = mem.ToArray
            '' '' ''oStream = New MemoryStream(baImage, 0, baImage.Length)

            '' '' ''Dim o2 As Stream = Nothing
            '' '' ''formatter.Serialize(mem, pCell)

            '' '' ''oStream = New MemoryStream(baImage)
            '' '' ''oImage = Image.FromStream(oStream)
            'oBmp = New Bitmap(ms)
            'oImage = oBmp

            'm_Image = img '' New Bitmap(stream)

            'MemoryStream that will hold the byte array
            'before converting to a Bitmap
            'Dim msImage As MemoryStream = Nothing

            'Dim ImgStream As New IO.MemoryStream(pCell, Byte()))
            ''Dim ImgStream As New IO.MemoryStream(CType(cmd.ExecuteScalar, Byte()))
            'msImage = Image.FromStream(ImgStream)

            'ImgStream.Dispose()

            'Dim bmpImage As Bitmap
            'Dim imgImage As Image = Nothing
            'Dim oImage As Object
            'oImage = CType(pCell, Object)
            'baImage = CType(CType(pCell, Object), Byte())
            'msImage = New MemoryStream(baImage, 0, baImage.Length)
            'pCell.
            'bmpImage = New Bitmap(msImage)
            'oImage = bmpImage

        Catch er As Exception
            MsgBox("Error in CImage." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SetImage(ByRef pCell As Object, ByRef oImage As Image, ByRef oByteArray As Byte())
        'Public Sub SetImage(ByRef pCell As cell, ByRef oImage As Image)

        Try

            oByteArray = DirectCast(pCell, Byte())

            If Not oByteArray Is Nothing Then
                Using ms As New MemoryStream(oByteArray, 0, oByteArray.Length)
                    ms.Write(oByteArray, 0, oByteArray.Length)
                    oImage = Image.FromStream(ms, True)
                End Using
            End If

        Catch er As Exception
            MsgBox("Error in CImage." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    'Public Sub SetImageAsByteArray(ByRef pCell As Object, ByRef oByte As Byte())

    '    Dim formatter As New BinaryFormatter()
    '    Dim mem As New IO.MemoryStream
    '    formatter.Serialize(mem, pCell)
    '    oByte = mem.ToArray

    '    'baImage = CType(pCell, Byte())

    'End Sub

    Public Sub SetImageAsByteArray(ByRef pTable As DataTable, ByVal piRow As Integer, ByVal pstrCellName As String, ByRef oByte As Byte())

        Dim oImage As Image = Nothing
        SetImage(pTable.Rows(piRow), pstrCellName, oImage)
        SetImageAsByteArray(oImage, oByte)

    End Sub

    Public Sub SetImageAsByteArray(ByRef pRow As DataRow, ByVal pstrCellName As String, ByRef oByte As Byte())

        Dim oImage As Image = Nothing
        SetImage(pRow.Item(pstrCellName), oImage)
        SetImageAsByteArray(oImage, oByte)

    End Sub

    Public Sub SetImageAsByteArray(ByRef pCell As DataColumn, ByRef oByte As Byte())

        Dim oImage As Image = Nothing
        SetImage(pCell, oImage)
        SetImageAsByteArray(oImage, oByte)

    End Sub

    Public Sub SetImageAsByteArray(ByRef pImage As Image, ByRef oByte As Byte())

        Dim ImageStream As MemoryStream

        ReDim oByte(0)

        If pImage IsNot Nothing Then
            Dim converter As New ImageConverter
            oByte = converter.ConvertTo(pImage, GetType(Byte()))

            ImageStream = New MemoryStream
            pImage.Save(ImageStream, ImageFormat.Jpeg)
            ReDim oByte(CInt(ImageStream.Length - 1))
            ImageStream.Position = 0
            ImageStream.Read(oByte, 0, CInt(ImageStream.Length))
        End If

    End Sub

    Public Sub ResizeImage(ByRef oImage As Image, ByRef oNewImage As Image, ByVal pX As Integer, ByVal pY As Integer)

        Dim oBmp As Bitmap

        oBmp = New Bitmap(pX, pY)

        Using g As Graphics = Graphics.FromImage(oBmp)
            g.DrawImage(oImage, 0, 0, oBmp.Width, oBmp.Height)
        End Using

        oNewImage = oBmp

    End Sub

End Class
