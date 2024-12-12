Imports System.Collections.Generic
Public Class frm_Oei_Report
    Private Sub cb_ORDER_EXCEPTION_Enum_list_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_ORDER_EXCEPTION_Enum_list.SelectedIndexChanged
        If cb_ORDER_EXCEPTION_Enum_list.SelectedIndex = 0 Then
            'MsgBox("all")
            'Dim oOEI_PRE_READ_ORDER_EXCEPTION_DAL = New cOEI_PRE_READ_ORDER_EXCEPTION_DAL
            'DGV_US_EX_REP.DataSource = oOEI_PRE_READ_ORDER_EXCEPTION_DAL.Load_EXCEPTION_report_ALL()
        Else

            'MsgBox(cb_ORDER_EXCEPTION_Enum_list.Text & "-" & cb_ORDER_EXCEPTION_Enum_list.SelectedValue)

        End If

    End Sub

    Private Sub fem_Oei_Report_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim oList = New List(Of cMdb_Cfg_Enum)
        Dim oOEI_PRE_READ_ORDER_EXCEPTION_DAL = New cOEI_PRE_READ_ORDER_EXCEPTION_DAL
        Dim oEnum As cMdb_Cfg_Enum = New cMdb_Cfg_Enum

        oList = oOEI_PRE_READ_ORDER_EXCEPTION_DAL.Load_EXCEPTION_Enum_list()



        With cb_ORDER_EXCEPTION_Enum_list
            .DataSource = oList
            .DisplayMember = "ENUM_VALUE"
            .ValueMember = "ID"
        End With
    End Sub

    Private Sub BT_US_EX_ORDER_Search_Click(sender As Object, e As EventArgs) Handles BT_US_EX_ORDER_Search.Click
        Dim oOEI_PRE_READ_ORDER_EXCEPTION_DAL = New cOEI_PRE_READ_ORDER_EXCEPTION_DAL
        If cb_ORDER_EXCEPTION_Enum_list.SelectedIndex = 0 Then


            DGV_US_EX_REP.DataSource = oOEI_PRE_READ_ORDER_EXCEPTION_DAL.Load_EXCEPTION_report_ALL()
        Else
            DGV_US_EX_REP.DataSource = oOEI_PRE_READ_ORDER_EXCEPTION_DAL.Load_EXCEPTION_report_by_Id(cb_ORDER_EXCEPTION_Enum_list.SelectedValue)


        End If
    End Sub
End Class