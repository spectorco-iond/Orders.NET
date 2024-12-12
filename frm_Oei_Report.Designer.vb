<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Oei_Report
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TP_US_EX_Order = New System.Windows.Forms.TabPage()
        Me.DGV_US_EX_REP = New System.Windows.Forms.DataGridView()
        Me.cb_ORDER_EXCEPTION_Enum_list = New System.Windows.Forms.ComboBox()
        Me.TabControl_Rep = New System.Windows.Forms.TabControl()
        Me.BT_US_EX_ORDER_Search = New System.Windows.Forms.Button()
        Me.TP_US_EX_Order.SuspendLayout()
        CType(Me.DGV_US_EX_REP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl_Rep.SuspendLayout()
        Me.SuspendLayout()
        '
        'TP_US_EX_Order
        '
        Me.TP_US_EX_Order.Controls.Add(Me.BT_US_EX_ORDER_Search)
        Me.TP_US_EX_Order.Controls.Add(Me.DGV_US_EX_REP)
        Me.TP_US_EX_Order.Controls.Add(Me.cb_ORDER_EXCEPTION_Enum_list)
        Me.TP_US_EX_Order.Location = New System.Drawing.Point(4, 22)
        Me.TP_US_EX_Order.Name = "TP_US_EX_Order"
        Me.TP_US_EX_Order.Padding = New System.Windows.Forms.Padding(3)
        Me.TP_US_EX_Order.Size = New System.Drawing.Size(654, 396)
        Me.TP_US_EX_Order.TabIndex = 0
        Me.TP_US_EX_Order.Text = "US_Exception_Order"
        Me.TP_US_EX_Order.UseVisualStyleBackColor = True
        '
        'DGV_US_EX_REP
        '
        Me.DGV_US_EX_REP.AllowUserToAddRows = False
        Me.DGV_US_EX_REP.AllowUserToDeleteRows = False
        Me.DGV_US_EX_REP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_US_EX_REP.Location = New System.Drawing.Point(15, 47)
        Me.DGV_US_EX_REP.Name = "DGV_US_EX_REP"
        Me.DGV_US_EX_REP.Size = New System.Drawing.Size(624, 327)
        Me.DGV_US_EX_REP.TabIndex = 2
        '
        'cb_ORDER_EXCEPTION_Enum_list
        '
        Me.cb_ORDER_EXCEPTION_Enum_list.FormattingEnabled = True
        Me.cb_ORDER_EXCEPTION_Enum_list.Location = New System.Drawing.Point(15, 20)
        Me.cb_ORDER_EXCEPTION_Enum_list.Name = "cb_ORDER_EXCEPTION_Enum_list"
        Me.cb_ORDER_EXCEPTION_Enum_list.Size = New System.Drawing.Size(121, 21)
        Me.cb_ORDER_EXCEPTION_Enum_list.TabIndex = 1
        '
        'TabControl_Rep
        '
        Me.TabControl_Rep.Controls.Add(Me.TP_US_EX_Order)
        Me.TabControl_Rep.Location = New System.Drawing.Point(43, 47)
        Me.TabControl_Rep.Name = "TabControl_Rep"
        Me.TabControl_Rep.SelectedIndex = 0
        Me.TabControl_Rep.Size = New System.Drawing.Size(662, 422)
        Me.TabControl_Rep.TabIndex = 3
        '
        'BT_US_EX_ORDER_Search
        '
        Me.BT_US_EX_ORDER_Search.Location = New System.Drawing.Point(162, 18)
        Me.BT_US_EX_ORDER_Search.Name = "BT_US_EX_ORDER_Search"
        Me.BT_US_EX_ORDER_Search.Size = New System.Drawing.Size(75, 23)
        Me.BT_US_EX_ORDER_Search.TabIndex = 0
        Me.BT_US_EX_ORDER_Search.Text = "Search"
        Me.BT_US_EX_ORDER_Search.UseVisualStyleBackColor = True
        '
        'fem_Oei_Report
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(732, 481)
        Me.Controls.Add(Me.TabControl_Rep)
        Me.Name = "fem_Oei_Report"
        Me.Text = "fem_Oei_Report"
        Me.TP_US_EX_Order.ResumeLayout(False)
        CType(Me.DGV_US_EX_REP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl_Rep.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TP_US_EX_Order As TabPage
    Friend WithEvents DGV_US_EX_REP As DataGridView
    Friend WithEvents cb_ORDER_EXCEPTION_Enum_list As ComboBox
    Friend WithEvents TabControl_Rep As TabControl
    Friend WithEvents BT_US_EX_ORDER_Search As Button
End Class
