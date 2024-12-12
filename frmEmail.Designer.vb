<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEmail
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEmail))
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdTo = New System.Windows.Forms.Button
        Me.cmdCC = New System.Windows.Forms.Button
        Me.lblSubject = New System.Windows.Forms.Label
        Me.bgWorker = New System.ComponentModel.BackgroundWorker
        Me.lblAttachments = New System.Windows.Forms.Label
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.txtAttachments = New Orders.xTextBox
        Me.txtBody = New Orders.xTextBox
        Me.txtSubject = New Orders.xTextBox
        Me.txtCC = New Orders.xTextBox
        Me.txtTo = New Orders.xTextBox
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.Location = New System.Drawing.Point(10, 11)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(68, 74)
        Me.cmdSave.TabIndex = 0
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdTo
        '
        Me.cmdTo.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTo.Location = New System.Drawing.Point(82, 11)
        Me.cmdTo.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdTo.Name = "cmdTo"
        Me.cmdTo.Size = New System.Drawing.Size(68, 22)
        Me.cmdTo.TabIndex = 1
        Me.cmdTo.Text = "To..."
        Me.cmdTo.UseVisualStyleBackColor = True
        '
        'cmdCC
        '
        Me.cmdCC.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCC.Location = New System.Drawing.Point(82, 37)
        Me.cmdCC.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdCC.Name = "cmdCC"
        Me.cmdCC.Size = New System.Drawing.Size(68, 22)
        Me.cmdCC.TabIndex = 2
        Me.cmdCC.Text = "CC..."
        Me.cmdCC.UseVisualStyleBackColor = True
        '
        'lblSubject
        '
        Me.lblSubject.AutoSize = True
        Me.lblSubject.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSubject.Location = New System.Drawing.Point(84, 66)
        Me.lblSubject.Name = "lblSubject"
        Me.lblSubject.Size = New System.Drawing.Size(51, 15)
        Me.lblSubject.TabIndex = 6
        Me.lblSubject.Text = "Subject:"
        '
        'bgWorker
        '
        '
        'lblAttachments
        '
        Me.lblAttachments.AutoSize = True
        Me.lblAttachments.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAttachments.Location = New System.Drawing.Point(12, 93)
        Me.lblAttachments.Name = "lblAttachments"
        Me.lblAttachments.Size = New System.Drawing.Size(75, 15)
        Me.lblAttachments.TabIndex = 10
        Me.lblAttachments.Text = "Attachments"
        Me.lblAttachments.Visible = False
        '
        'cmdSearch
        '
        Me.cmdSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSearch.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSearch.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSearch.Image = CType(resources.GetObject("cmdSearch.Image"), System.Drawing.Image)
        Me.cmdSearch.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdSearch.Location = New System.Drawing.Point(595, 62)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSearch.Size = New System.Drawing.Size(24, 24)
        Me.cmdSearch.TabIndex = 11
        Me.cmdSearch.TabStop = False
        Me.cmdSearch.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdSearch.UseVisualStyleBackColor = False
        '
        'txtAttachments
        '
        Me.txtAttachments.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtAttachments.DataLength = CType(0, Long)
        Me.txtAttachments.DataType = Orders.CDataType.dtString
        Me.txtAttachments.DateValue = New Date(CType(0, Long))
        Me.txtAttachments.DecimalDigits = CType(0, Long)
        Me.txtAttachments.Location = New System.Drawing.Point(92, 89)
        Me.txtAttachments.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAttachments.Mask = Nothing
        Me.txtAttachments.Name = "txtAttachments"
        Me.txtAttachments.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtAttachments.OldValue = ""
        Me.txtAttachments.ReadOnly = True
        Me.txtAttachments.Size = New System.Drawing.Size(527, 22)
        Me.txtAttachments.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtAttachments.StringValue = Nothing
        Me.txtAttachments.TabIndex = 9
        Me.txtAttachments.TextDataID = Nothing
        Me.txtAttachments.Visible = False
        '
        'txtBody
        '
        Me.txtBody.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtBody.DataLength = CType(0, Long)
        Me.txtBody.DataType = Orders.CDataType.dtString
        Me.txtBody.DateValue = New Date(CType(0, Long))
        Me.txtBody.DecimalDigits = CType(0, Long)
        Me.txtBody.Location = New System.Drawing.Point(11, 89)
        Me.txtBody.Margin = New System.Windows.Forms.Padding(2)
        Me.txtBody.Mask = Nothing
        Me.txtBody.Multiline = True
        Me.txtBody.Name = "txtBody"
        Me.txtBody.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtBody.OldValue = ""
        Me.txtBody.Size = New System.Drawing.Size(608, 381)
        Me.txtBody.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtBody.StringValue = Nothing
        Me.txtBody.TabIndex = 7
        Me.txtBody.TextDataID = Nothing
        '
        'txtSubject
        '
        Me.txtSubject.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtSubject.DataLength = CType(0, Long)
        Me.txtSubject.DataType = Orders.CDataType.dtString
        Me.txtSubject.DateValue = New Date(CType(0, Long))
        Me.txtSubject.DecimalDigits = CType(0, Long)
        Me.txtSubject.Location = New System.Drawing.Point(154, 63)
        Me.txtSubject.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSubject.Mask = Nothing
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtSubject.OldValue = ""
        Me.txtSubject.Size = New System.Drawing.Size(440, 22)
        Me.txtSubject.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtSubject.StringValue = Nothing
        Me.txtSubject.TabIndex = 5
        Me.txtSubject.TextDataID = Nothing
        '
        'txtCC
        '
        Me.txtCC.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.txtCC.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.txtCC.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtCC.DataLength = CType(0, Long)
        Me.txtCC.DataType = Orders.CDataType.dtString
        Me.txtCC.DateValue = New Date(CType(0, Long))
        Me.txtCC.DecimalDigits = CType(0, Long)
        Me.txtCC.Location = New System.Drawing.Point(154, 37)
        Me.txtCC.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCC.Mask = Nothing
        Me.txtCC.Name = "txtCC"
        Me.txtCC.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtCC.OldValue = ""
        Me.txtCC.Size = New System.Drawing.Size(465, 22)
        Me.txtCC.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtCC.StringValue = Nothing
        Me.txtCC.TabIndex = 4
        Me.txtCC.TextDataID = Nothing
        '
        'txtTo
        '
        Me.txtTo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.txtTo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.txtTo.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtTo.DataLength = CType(0, Long)
        Me.txtTo.DataType = Orders.CDataType.dtString
        Me.txtTo.DateValue = New Date(CType(0, Long))
        Me.txtTo.DecimalDigits = CType(0, Long)
        Me.txtTo.Location = New System.Drawing.Point(154, 11)
        Me.txtTo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTo.Mask = Nothing
        Me.txtTo.Name = "txtTo"
        Me.txtTo.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtTo.OldValue = ""
        Me.txtTo.Size = New System.Drawing.Size(465, 22)
        Me.txtTo.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtTo.StringValue = Nothing
        Me.txtTo.TabIndex = 3
        Me.txtTo.TextDataID = Nothing
        '
        'frmEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(630, 481)
        Me.Controls.Add(Me.cmdSearch)
        Me.Controls.Add(Me.lblAttachments)
        Me.Controls.Add(Me.txtAttachments)
        Me.Controls.Add(Me.txtBody)
        Me.Controls.Add(Me.lblSubject)
        Me.Controls.Add(Me.txtSubject)
        Me.Controls.Add(Me.txtCC)
        Me.Controls.Add(Me.txtTo)
        Me.Controls.Add(Me.cmdCC)
        Me.Controls.Add(Me.cmdTo)
        Me.Controls.Add(Me.cmdSave)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEmail"
        Me.Text = "Send Email"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdTo As System.Windows.Forms.Button
    Friend WithEvents cmdCC As System.Windows.Forms.Button
    Friend WithEvents txtTo As Orders.xTextBox
    Friend WithEvents txtCC As Orders.xTextBox
    Friend WithEvents txtSubject As Orders.xTextBox
    Friend WithEvents lblSubject As System.Windows.Forms.Label
    Friend WithEvents txtBody As Orders.xTextBox
    Friend WithEvents bgWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents txtAttachments As Orders.xTextBox
    Friend WithEvents lblAttachments As System.Windows.Forms.Label
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
End Class
