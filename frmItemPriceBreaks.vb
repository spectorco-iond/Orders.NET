Public Class frmItemPriceBreaks

    Dim m_Item_No As String
    Dim m_Cus_No As String
    Dim m_Curr_Cd As String
    Dim m_Start_Dt As String
    Dim m_End_Dt As String

    Public Sub New(ByVal pstrCurr_Cd As String, ByVal pstrCus_No As String, ByVal pstrItem_No As String, ByVal pstrStart_Dt As String, ByVal pstrEndDt As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        m_Curr_Cd = pstrCurr_Cd
        m_Cus_No = pstrCus_No
        m_Item_No = pstrItem_No
        m_Start_Dt = pstrStart_Dt
        m_End_Dt = pstrEndDt

        Call DataLoad()

    End Sub

    Private Sub DataLoad()

        Try

            Dim dt As DataTable
            Dim db As New cDBA

            Dim strSql As String
            strSql = "SELECT * FROM DBO.OEI_Item_Price_Breaks('" & m_Curr_Cd & "', '" & m_Cus_No & "', '" & m_Item_No & "', '" & m_Start_Dt & "', '" & m_End_Dt & "')"

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                txtStartDate.Text = dt.Rows(0).Item("Start_Dt")
                txtEndDate.Text = dt.Rows(0).Item("End_Dt")

                txtMinQty1.Text = dt.Rows(0).Item("MinQty1")
                txtMinQty2.Text = dt.Rows(0).Item("MinQty2")
                txtMinQty3.Text = dt.Rows(0).Item("MinQty3")
                txtMinQty4.Text = dt.Rows(0).Item("MinQty4")
                txtMinQty5.Text = dt.Rows(0).Item("MinQty5")
                txtMinQty6.Text = dt.Rows(0).Item("MinQty6")
                txtMinQty7.Text = dt.Rows(0).Item("MinQty7")
                txtMinQty8.Text = dt.Rows(0).Item("MinQty8")
                txtMinQty9.Text = dt.Rows(0).Item("MinQty9")
                txtMinQty10.Text = dt.Rows(0).Item("MinQty10")

                txtPrice1.Text = dt.Rows(0).Item("Price1")
                txtPrice2.Text = dt.Rows(0).Item("Price2")
                txtPrice3.Text = dt.Rows(0).Item("Price3")
                txtPrice4.Text = dt.Rows(0).Item("Price4")
                txtPrice5.Text = dt.Rows(0).Item("Price5")
                txtPrice6.Text = dt.Rows(0).Item("Price6")
                txtPrice7.Text = dt.Rows(0).Item("Price7")
                txtPrice8.Text = dt.Rows(0).Item("Price8")
                txtPrice9.Text = dt.Rows(0).Item("Price9")
                txtPrice10.Text = dt.Rows(0).Item("Price10")

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

End Class