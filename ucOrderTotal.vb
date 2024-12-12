Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports System.IO

Friend Class ucOrderTotal
    Inherits System.Windows.Forms.UserControl

    'Private m_oOrder As New cOrder
    Private blnLoading As Boolean = False

    Public dsDataSet As New DataSet()

    Public Sub LoadOrder(ByRef rsDataset As DataSet)

        blnLoading = True

        If Trim(m_oOrder.Ordhead.Ord_No).Length <> 0 Then
            dsDataSet = rsDataset
        Else
            'NewOrder(pOrder)
        End If

        blnLoading = False

    End Sub

    Public Sub Fill() ' ByRef pOrder As cOrder)

        ' m_oOrder = pOrder

    End Sub

    Public Sub Reset() ' ByRef pOrder As cOrder)

    End Sub

    Public Sub Save()
        '-- 13.80 14.31 tot_sls_amt = 13.80 tot_tax_amt = 13.80 tot_dollars = 13.80 

        m_oOrder.Ordhead.Tot_Sls_Amt = lblOrig_Curr_Amt.Text

    End Sub

    Public ReadOnly Property Order() As cOrder
        Get
            Order = m_oOrder
        End Get
    End Property

    Public Property Order_Item_Count() As Integer
        Get
            If IsNumeric(lblOrder_Item_Count.Text) Then
                Order_Item_Count = CInt(lblOrder_Item_Count.Text)
            Else
                Order_Item_Count = 0
            End If
        End Get
        Set(ByVal value As Integer)
            lblOrder_Item_Count.Text = value
        End Set
    End Property

    Public Property Curr_Trx_Rt() As Double
        Get
            If IsNumeric(lblCurr_Trx_Rt.Text) Then
                Curr_Trx_Rt = CDbl(lblCurr_Trx_Rt.Text)
            Else
                Curr_Trx_Rt = 0
            End If
        End Get
        Set(ByVal value As Double)
            lblCurr_Trx_Rt.Text = String.Format("{0:f6}", value)
        End Set
    End Property

    Public Property Order_Amt() As Double
        Get
            If IsNumeric(lblOrder_Amt.Text) Then
                Order_Amt = CDbl(lblOrder_Amt.Text)
            Else
                Order_Amt = 0
            End If
        End Get
        Set(ByVal value As Double)
            lblOrder_Amt.Text = String.Format("{0:f2}", value)
        End Set
    End Property

    Public Property Order_Qty() As Integer
        Get
            If IsNumeric(lblOrder_Qty.Text) Then
                Order_Qty = CInt(lblOrder_Qty.Text)
            Else
                Order_Qty = 0
            End If
        End Get
        Set(ByVal value As Integer)
            lblOrder_Qty.Text = String.Format("{0:f2}", value)
        End Set
    End Property

    Public Property Ship_Amt() As Double
        Get
            If IsNumeric(lblShip_Amt.Text) Then
                Ship_Amt = CDbl(lblShip_Amt.Text)
            Else
                Ship_Amt = 0
            End If
        End Get
        Set(ByVal value As Double)
            lblShip_Amt.Text = String.Format("{0:f2}", value)
            lblOrig_Curr_Amt.Text = String.Format("{0:f2}", value)
            lblConv_Curr_Amt.Text = String.Format("{0:f2}", (value * lblCurr_Trx_Rt.Text))
        End Set
    End Property

    Public Property Ship_Qty() As Integer
        Get
            If IsNumeric(lblShip_Qty.Text) Then
                Ship_Qty = CInt(lblShip_Qty.Text)
            Else
                Ship_Qty = 0
            End If
        End Get
        Set(ByVal value As Integer)
            lblShip_Qty.Text = String.Format("{0:f2}", value)
        End Set
    End Property

    Public Property From_Curr_Cd_Desc() As String
        Get
            From_Curr_Cd_Desc = lblFrom_Curr_Cd_Desc.Text
        End Get
        Set(ByVal value As String)
            lblFrom_Curr_Cd_Desc.Text = value

            'Try
            '    'Value is currency code or string code, will search for it else will put it as is.
            '    Dim rsCurr As New ADODB.Recordset
            '    Dim strSql As String = "" & _
            '    "SELECT OMS30_0 FROM Valuta WITH (Nolock) WHERE Valcode = '" & value & "'"
            '    rsCurr.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            '    If rsCurr.RecordCount <> 0 Then
            '        lblFrom_Curr_Cd_Desc.Text = rsCurr.Fields("OMS30_0").Value
            '    Else
            '        lblFrom_Curr_Cd_Desc.Text = value
            '    End If
            '    rsCurr.Close()

            'Catch er As Exception
            '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            'End Try
        End Set
    End Property

    Public Property To_Curr_Cd_Desc() As String
        Get
            To_Curr_Cd_Desc = lblTo_Curr_Cd_Desc.Text
        End Get
        Set(ByVal value As String)
            lblTo_Curr_Cd_Desc.Text = value
        End Set
    End Property

End Class
