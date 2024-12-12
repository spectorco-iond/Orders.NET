Public Class cCustomerInfo

    Private oform As frmCustomerInfo

    Public Sub New(ByVal pstrCus_No As String)

        Try
            ' open form
            oform = New frmCustomerInfo(pstrCus_No)

            '' load datagrid
            'oform.dgvComments.DataSource = Load(pstrOrd_Guid)

            ' show me the money
            oform.ShowDialog()

            oform.Dispose()

        Catch er As Exception
            MsgBox("Error in frmCustomerInfo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Sub Show()

    '    Try
    '        oform.ShowDialog()

    '    Catch er As Exception
    '        MsgBox("Error in cComments." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Public ReadOnly Property Form() As frmCustomerInfo
        Get
            Form = oform
        End Get
    End Property

End Class
