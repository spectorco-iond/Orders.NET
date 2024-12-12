<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNote
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
        Me.lblDate = New System.Windows.Forms.Label
        Me.lblUser = New System.Windows.Forms.Label
        Me.lblTopic = New System.Windows.Forms.Label
        Me.lblAttachment = New System.Windows.Forms.Label
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.txtNote_Dt = New System.Windows.Forms.TextBox
        Me.txtDocument_Field_1 = New System.Windows.Forms.TextBox
        Me.txtNotes = New Orders.xTextBox
        Me.txtNote_Topic = New Orders.xTextBox
        Me.txtUser_Name = New Orders.xTextBox
        Me.SuspendLayout()
        '
        'lblDate
        '
        Me.lblDate.AutoSize = True
        Me.lblDate.Location = New System.Drawing.Point(12, 18)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(37, 16)
        Me.lblDate.TabIndex = 0
        Me.lblDate.Text = "Date"
        '
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Location = New System.Drawing.Point(262, 18)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(37, 16)
        Me.lblUser.TabIndex = 2
        Me.lblUser.Text = "User"
        '
        'lblTopic
        '
        Me.lblTopic.AutoSize = True
        Me.lblTopic.Location = New System.Drawing.Point(12, 46)
        Me.lblTopic.Name = "lblTopic"
        Me.lblTopic.Size = New System.Drawing.Size(43, 16)
        Me.lblTopic.TabIndex = 4
        Me.lblTopic.Text = "Topic"
        '
        'lblAttachment
        '
        Me.lblAttachment.AutoSize = True
        Me.lblAttachment.Location = New System.Drawing.Point(12, 74)
        Me.lblAttachment.Name = "lblAttachment"
        Me.lblAttachment.Size = New System.Drawing.Size(74, 16)
        Me.lblAttachment.TabIndex = 6
        Me.lblAttachment.Text = "Attachment"
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Location = New System.Drawing.Point(310, 247)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(81, 26)
        Me.cmdSave.TabIndex = 9
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(399, 247)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(81, 26)
        Me.cmdCancel.TabIndex = 10
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'txtNote_Dt
        '
        Me.txtNote_Dt.Location = New System.Drawing.Point(92, 15)
        Me.txtNote_Dt.Name = "txtNote_Dt"
        Me.txtNote_Dt.ReadOnly = True
        Me.txtNote_Dt.Size = New System.Drawing.Size(164, 22)
        Me.txtNote_Dt.TabIndex = 1
        '
        'txtDocument_Field_1
        '
        Me.txtDocument_Field_1.AllowDrop = True
        Me.txtDocument_Field_1.Enabled = False
        Me.txtDocument_Field_1.Location = New System.Drawing.Point(92, 71)
        Me.txtDocument_Field_1.Name = "txtDocument_Field_1"
        Me.txtDocument_Field_1.Size = New System.Drawing.Size(388, 22)
        Me.txtDocument_Field_1.TabIndex = 7
        '
        'txtNotes
        '
        Me.txtNotes.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtNotes.DataLength = CType(250, Long)
        Me.txtNotes.DataType = Orders.CDataType.dtString
        Me.txtNotes.DateValue = New Date(CType(0, Long))
        Me.txtNotes.DecimalDigits = CType(0, Long)
        Me.txtNotes.Location = New System.Drawing.Point(15, 99)
        Me.txtNotes.Mask = Nothing
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtNotes.OldValue = ""
        Me.txtNotes.Size = New System.Drawing.Size(465, 141)
        Me.txtNotes.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtNotes.StringValue = Nothing
        Me.txtNotes.TabIndex = 8
        Me.txtNotes.TextDataID = Nothing
        '
        'txtNote_Topic
        '
        Me.txtNote_Topic.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtNote_Topic.DataLength = CType(40, Long)
        Me.txtNote_Topic.DataType = Orders.CDataType.dtString
        Me.txtNote_Topic.DateValue = New Date(CType(0, Long))
        Me.txtNote_Topic.DecimalDigits = CType(0, Long)
        Me.txtNote_Topic.Location = New System.Drawing.Point(92, 43)
        Me.txtNote_Topic.Mask = Nothing
        Me.txtNote_Topic.Name = "txtNote_Topic"
        Me.txtNote_Topic.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtNote_Topic.OldValue = ""
        Me.txtNote_Topic.Size = New System.Drawing.Size(388, 22)
        Me.txtNote_Topic.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtNote_Topic.StringValue = Nothing
        Me.txtNote_Topic.TabIndex = 5
        Me.txtNote_Topic.TextDataID = Nothing
        '
        'txtUser_Name
        '
        Me.txtUser_Name.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtUser_Name.DataLength = CType(20, Long)
        Me.txtUser_Name.DataType = Orders.CDataType.dtString
        Me.txtUser_Name.DateValue = New Date(CType(0, Long))
        Me.txtUser_Name.DecimalDigits = CType(0, Long)
        Me.txtUser_Name.Location = New System.Drawing.Point(325, 15)
        Me.txtUser_Name.Mask = Nothing
        Me.txtUser_Name.Name = "txtUser_Name"
        Me.txtUser_Name.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtUser_Name.OldValue = ""
        Me.txtUser_Name.ReadOnly = True
        Me.txtUser_Name.Size = New System.Drawing.Size(155, 22)
        Me.txtUser_Name.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtUser_Name.StringValue = Nothing
        Me.txtUser_Name.TabIndex = 3
        Me.txtUser_Name.TextDataID = Nothing
        '
        'frmNote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(492, 286)
        Me.Controls.Add(Me.txtNotes)
        Me.Controls.Add(Me.txtNote_Topic)
        Me.Controls.Add(Me.txtUser_Name)
        Me.Controls.Add(Me.txtDocument_Field_1)
        Me.Controls.Add(Me.txtNote_Dt)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lblAttachment)
        Me.Controls.Add(Me.lblTopic)
        Me.Controls.Add(Me.lblUser)
        Me.Controls.Add(Me.lblDate)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(500, 320)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(500, 320)
        Me.Name = "frmNote"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Note"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents lblTopic As System.Windows.Forms.Label
    Friend WithEvents lblAttachment As System.Windows.Forms.Label
    Public WithEvents cmdSave As System.Windows.Forms.Button
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents txtNote_Dt As System.Windows.Forms.TextBox
    Friend WithEvents txtDocument_Field_1 As System.Windows.Forms.TextBox
    Friend WithEvents txtUser_Name As Orders.xTextBox
    Friend WithEvents txtNote_Topic As Orders.xTextBox
    Friend WithEvents txtNotes As Orders.xTextBox
End Class
