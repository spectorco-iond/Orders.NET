Public Class COEFlash

    Private m_strCus_No As String
    Private m_strCus_Name As String
    Private m_intStarStatus As Integer
    Private m_blnEQP As Boolean
    Private m_intEQPColumn As Integer
    Private m_intEQPPlusPct As Integer

    Private m_colOEFlashCharges As Collection ' of cOEFlashCharges
    Private m_colOEFlashCorrespondance As Collection ' of cOEFlashCorrespondance
    Private m_colOEFlashCustomer As Collection ' of cOEFlashCustomer
    Private m_colOEFlashPricing As Collection ' of cOEFlashPricing
    Private m_colOEFlashShipping As Collection ' of cOEFlashShipping

    Private db As New cDBA

    Public Sub New()

        Call Init()

    End Sub

    Public Sub New(ByVal pstrCus_No As String)

        m_strCus_No = pstrCus_No
        Call Reload()

    End Sub

    Private Sub Delete()

    End Sub

    Private Sub Init()

        Try

            m_strCus_No = String.Empty
            m_strCus_Name = String.Empty
            m_intStarStatus = 0
            m_blnEQP = False
            m_intEQPColumn = 0
            m_intEQPPlusPct = 0

            m_colOEFlashCharges = New Collection ' of cOEFlashCharges
            m_colOEFlashCorrespondance = New Collection ' of cOEFlashCorrespondance
            m_colOEFlashCustomer = New Collection ' of cOEFlashCustomer
            m_colOEFlashPricing = New Collection ' of cOEFlashPricing
            m_colOEFlashShipping = New Collection ' of cOEFlashShipping

        Catch er As Exception
            MsgBox("Error in COEFlash." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Load()

        Call LoadCharges()
        Call LoadCorrespondance()
        Call LoadCustomer()
        Call LoadPricing()
        Call LoadShipping()

    End Sub

    Private Sub LoadCharges()

        Try

            Dim dt As DataTable
            Dim strSql As String = ""

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                For Each drRow As DataRow In dt.Rows

                Next
            End If

        Catch er As Exception
            MsgBox("Error in COEFlash." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadCorrespondance()

        Try

            Dim dt As DataTable
            Dim strSql As String = ""

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                For Each drRow As DataRow In dt.Rows

                Next
            End If

        Catch er As Exception
            MsgBox("Error in COEFlash." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadCustomer()

        Try

            Dim dt As DataTable
            Dim strSql As String = ""

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                For Each drRow As DataRow In dt.Rows

                Next
            End If

        Catch er As Exception
            MsgBox("Error in COEFlash." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadPricing()

        Try

            Dim dt As DataTable
            Dim strSql As String = ""

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                For Each drRow As DataRow In dt.Rows

                Next
            End If

        Catch er As Exception
            MsgBox("Error in COEFlash." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadShipping()

        Try

            Dim dt As DataTable
            Dim strSql As String = ""

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                For Each drRow As DataRow In dt.Rows

                Next
            End If

        Catch er As Exception
            MsgBox("Error in COEFlash." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Reload()

        Call Reset()
        Call Load()

    End Sub

    Private Sub Reset()

        m_intStarStatus = 0
        m_blnEQP = False
        m_intEQPColumn = 0
        m_intEQPPlusPct = 0

        m_colOEFlashCharges = New Collection ' of cOEFlashCharges
        m_colOEFlashCorrespondance = New Collection ' of cOEFlashCorrespondance
        m_colOEFlashCustomer = New Collection ' of cOEFlashCustomer
        m_colOEFlashPricing = New Collection ' of cOEFlashPricing
        m_colOEFlashShipping = New Collection ' of cOEFlashShipping

    End Sub

    Private Sub Save()

    End Sub

End Class
