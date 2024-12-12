Public Class frmItemLocQty

    Private m_Item_No As String
    Private m_Loc As String

    Public Sub New(ByVal pstrItem_No As String, ByVal pstrLoc As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        m_Item_No = pstrItem_No
        m_Loc = pstrLoc

        Call DataLoad()

    End Sub

    Private Sub DataLoad()

        Try

            Dim dt As DataTable
            Dim db As New cDBA

            Dim strSql As String
            'strSql = "SELECT * FROM DBO.OEI_Item_Price_Breaks('" & m_Curr_Cd & "', '" & m_Cus_No & "', '" & m_Item_No & "', '" & m_Start_Dt & "', '" & m_End_Dt & "')"

            If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then
                strSql = "EXEC DBO.OEI_Item_Loc_Qty_Form '" & m_Item_No & "', '" & m_Loc & "'"
            Else
                'ID 02.17.2023
                strSql = "EXEC [200].DBO.OEI_Item_Loc_Qty_Form_US '" & m_Item_No & "', '" & m_Loc & "'"
            End If


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                'SELECT 		I.Item_No, I.Item_Desc_1, I.Item_Desc_2, L.Loc, LF.Loc_Desc, 
                'END AS Pur_Or_Mfg_Desc, 
                'I.Pur_Or_Mfg, I.Prod_Cat, I.Mat_Cost_Type, 
                'ISNULL(I.Stocked_Fg, 'N') AS Stocked_Fg, ISNULL(I.Controlled_Fg, 'N') AS Controlled_Fg, 
                'ISNULL(I.BkOrd_Fg, 'N') AS BkOrd_Fg, ISNULL(I.Tax_Fg, 'N') AS Tax_Fg, 
                'L.Qty_On_Hand, L.Qty_Allocated, L.Qty_BkOrd, L.Qty_On_Ord, L.Mult_Bin_Fg, 


                txtItem_No_2.Text = dt.Rows(0).Item("Item_No").ToString
                txtItem_No.Text = txtItem_No_2.Text
                txtLoc_2.Text = dt.Rows(0).Item("Loc").ToString
                txtLoc.Text = txtLoc_2.Text
                txtItem_Desc_1.Text = dt.Rows(0).Item("Item_Desc_1").ToString
                txtItem_Desc_2.Text = dt.Rows(0).Item("Item_Desc_2").ToString
                txtLoc_Desc.Text = dt.Rows(0).Item("Loc_Desc").ToString
                txtPur_Or_Mfg_Desc.Text = dt.Rows(0).Item("Pur_Or_Mfg_Desc").ToString
                txtProd_Cat.Text = dt.Rows(0).Item("Prod_Cat").ToString
                txtMat_Cost_Type.Text = dt.Rows(0).Item("Mat_Cost_Type").ToString
                chkStocked_Fg.Checked = (dt.Rows(0).Item("Stocked_Fg") = "Y")
                chkControlled_Fg.Checked = (dt.Rows(0).Item("Controlled_Fg") = "Y")
                chkBkOrd_Fg.Checked = (dt.Rows(0).Item("BkOrd_Fg") = "Y")
                chkTax_Fg.Checked = (dt.Rows(0).Item("Tax_Fg") = "Y")
                chkMult_Bin_Fg.Checked = (dt.Rows(0).Item("Mult_Bin_Fg") = "Y")

                'Dim oPrice As cOEPriceFile = New cOEPriceFile(txtItem_No.Text.Trim, m_oOrder.Ordhead.Curr_Cd)
                'Dim intInv_Buffer As Integer = IIf(txtLoc.Text.Trim = "1", oPrice.Get_Inventory_Buffer, 0)
                txtQty_On_Hand.Text = dt.Rows(0).Item("Qty_On_Hand").ToString '- intInv_Buffer
                txtQty_Allocated.Text = dt.Rows(0).Item("Qty_Allocated").ToString
                txtQty_BkOrd.Text = dt.Rows(0).Item("Qty_BkOrd").ToString
                txtQty_On_Ord.Text = dt.Rows(0).Item("Qty_On_Ord").ToString
                txtUser_Def_Fld_1.Text = dt.Rows(0).Item("User_Def_Fld_1").ToString
                txtUser_Def_Fld_2.Text = dt.Rows(0).Item("User_Def_Fld_2").ToString
                txtUser_Def_Fld_3.Text = dt.Rows(0).Item("User_Def_Fld_3").ToString
                txtUser_Def_Fld_4.Text = dt.Rows(0).Item("User_Def_Fld_4").ToString
                txtUser_Def_Fld_5.Text = dt.Rows(0).Item("User_Def_Fld_5").ToString
                txtItem_Note_1.Text = dt.Rows(0).Item("Item_Note_1").ToString
                txtItem_Note_2.Text = dt.Rows(0).Item("Item_Note_2").ToString
                txtItem_Note_3.Text = dt.Rows(0).Item("Item_Note_3").ToString
                txtItem_Note_4.Text = dt.Rows(0).Item("Item_Note_4").ToString
                txtItem_Note_5.Text = dt.Rows(0).Item("Item_Note_5").ToString
                txtLoc_Def_Fld_1.Text = dt.Rows(0).Item("Loc_Def_Fld_1").ToString
                txtLoc_Def_Fld_2.Text = dt.Rows(0).Item("Loc_Def_Fld_2").ToString
                txtLoc_Def_Fld_3.Text = dt.Rows(0).Item("Loc_Def_Fld_3").ToString
                txtLoc_Def_Fld_4.Text = dt.Rows(0).Item("Loc_Def_Fld_4").ToString
                txtLoc_Def_Fld_5.Text = dt.Rows(0).Item("Loc_Def_Fld_5").ToString

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

End Class