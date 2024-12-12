<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_route_maintain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button_ADD = New System.Windows.Forms.Button()
        Me.ButtonDEL = New System.Windows.Forms.Button()
        Me.ButtonRefresh = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(1, 37)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(754, 367)
        Me.DataGridView1.TabIndex = 1
        '
        'Button_ADD
        '
        Me.Button_ADD.Location = New System.Drawing.Point(12, 8)
        Me.Button_ADD.Name = "Button_ADD"
        Me.Button_ADD.Size = New System.Drawing.Size(75, 23)
        Me.Button_ADD.TabIndex = 2
        Me.Button_ADD.Text = "ADD"
        Me.Button_ADD.UseVisualStyleBackColor = True
        '
        'ButtonDEL
        '
        Me.ButtonDEL.Location = New System.Drawing.Point(122, 8)
        Me.ButtonDEL.Name = "ButtonDEL"
        Me.ButtonDEL.Size = New System.Drawing.Size(75, 23)
        Me.ButtonDEL.TabIndex = 3
        Me.ButtonDEL.Text = "DELETE"
        Me.ButtonDEL.UseVisualStyleBackColor = True
        '
        'ButtonRefresh
        '
        Me.ButtonRefresh.Location = New System.Drawing.Point(237, 8)
        Me.ButtonRefresh.Name = "ButtonRefresh"
        Me.ButtonRefresh.Size = New System.Drawing.Size(75, 23)
        Me.ButtonRefresh.TabIndex = 4
        Me.ButtonRefresh.Text = "Refresh"
        Me.ButtonRefresh.UseVisualStyleBackColor = True
        '
        'frm_route_maintain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(767, 416)
        Me.Controls.Add(Me.ButtonRefresh)
        Me.Controls.Add(Me.ButtonDEL)
        Me.Controls.Add(Me.Button_ADD)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "frm_route_maintain"
        Me.Text = "frm_route_maintain"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button_ADD As Button
    Friend WithEvents ButtonDEL As Button
    Friend WithEvents ButtonRefresh As Button
End Class
