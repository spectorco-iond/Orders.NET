<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucImprint_2
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.cmdUpdateSelected = New System.Windows.Forms.Button
        Me.cmdLockImprintData = New System.Windows.Forms.Button
        Me.cmdUpdateAll = New System.Windows.Forms.Button
        Me.cmdGetRepeatData = New System.Windows.Forms.Button
        Me.cmdClear = New System.Windows.Forms.Button
        Me.gbRepeatData = New System.Windows.Forms.GroupBox
        Me.txtRepeatOrderNo = New System.Windows.Forms.TextBox
        Me.lblRepeatOrderNo = New System.Windows.Forms.Label
        Me.lblRefill = New System.Windows.Forms.Label
        Me.txtRefill = New System.Windows.Forms.TextBox
        Me.cboRefill = New System.Windows.Forms.ComboBox
        Me.lblPackaging = New System.Windows.Forms.Label
        Me.txtPackaging = New System.Windows.Forms.TextBox
        Me.cboPackaging = New System.Windows.Forms.ComboBox
        Me.txtImprintColor = New System.Windows.Forms.TextBox
        Me.lblImprintColour = New System.Windows.Forms.Label
        Me.cboImprintColor = New System.Windows.Forms.ComboBox
        Me.lblImprintLocation = New System.Windows.Forms.Label
        Me.txtImprintLocation = New System.Windows.Forms.TextBox
        Me.lblSpecialComments = New System.Windows.Forms.Label
        Me.txtSpecialComments = New System.Windows.Forms.TextBox
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.lblComments = New System.Windows.Forms.Label
        Me.txtNumImprints = New System.Windows.Forms.TextBox
        Me.lblNumImprints = New System.Windows.Forms.Label
        Me.cboImprintLocation = New System.Windows.Forms.ComboBox
        Me.txtImprint = New System.Windows.Forms.TextBox
        Me.lblImprint = New System.Windows.Forms.Label
        Me.gbRepeatData.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdUpdateSelected
        '
        Me.cmdUpdateSelected.Location = New System.Drawing.Point(64, 109)
        Me.cmdUpdateSelected.Name = "cmdUpdateSelected"
        Me.cmdUpdateSelected.Size = New System.Drawing.Size(118, 25)
        Me.cmdUpdateSelected.TabIndex = 49
        Me.cmdUpdateSelected.Text = "Update Selected"
        Me.cmdUpdateSelected.UseVisualStyleBackColor = True
        '
        'cmdLockImprintData
        '
        Me.cmdLockImprintData.Location = New System.Drawing.Point(279, 109)
        Me.cmdLockImprintData.Name = "cmdLockImprintData"
        Me.cmdLockImprintData.Size = New System.Drawing.Size(120, 25)
        Me.cmdLockImprintData.TabIndex = 48
        Me.cmdLockImprintData.Text = "Lock Imprint Data"
        Me.cmdLockImprintData.UseVisualStyleBackColor = True
        '
        'cmdUpdateAll
        '
        Me.cmdUpdateAll.Location = New System.Drawing.Point(188, 109)
        Me.cmdUpdateAll.Name = "cmdUpdateAll"
        Me.cmdUpdateAll.Size = New System.Drawing.Size(85, 25)
        Me.cmdUpdateAll.TabIndex = 47
        Me.cmdUpdateAll.Text = "Update All"
        Me.cmdUpdateAll.UseVisualStyleBackColor = True
        '
        'cmdGetRepeatData
        '
        Me.cmdGetRepeatData.Location = New System.Drawing.Point(9, 67)
        Me.cmdGetRepeatData.Name = "cmdGetRepeatData"
        Me.cmdGetRepeatData.Size = New System.Drawing.Size(114, 25)
        Me.cmdGetRepeatData.TabIndex = 13
        Me.cmdGetRepeatData.Text = "Get Repeat Data"
        Me.cmdGetRepeatData.UseVisualStyleBackColor = True
        '
        'cmdClear
        '
        Me.cmdClear.Location = New System.Drawing.Point(3, 109)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(55, 25)
        Me.cmdClear.TabIndex = 46
        Me.cmdClear.Text = "Clear"
        Me.cmdClear.UseVisualStyleBackColor = True
        '
        'gbRepeatData
        '
        Me.gbRepeatData.Controls.Add(Me.cmdGetRepeatData)
        Me.gbRepeatData.Controls.Add(Me.txtRepeatOrderNo)
        Me.gbRepeatData.Controls.Add(Me.lblRepeatOrderNo)
        Me.gbRepeatData.Location = New System.Drawing.Point(836, 4)
        Me.gbRepeatData.Name = "gbRepeatData"
        Me.gbRepeatData.Size = New System.Drawing.Size(132, 99)
        Me.gbRepeatData.TabIndex = 45
        Me.gbRepeatData.TabStop = False
        '
        'txtRepeatOrderNo
        '
        Me.txtRepeatOrderNo.Location = New System.Drawing.Point(9, 39)
        Me.txtRepeatOrderNo.Name = "txtRepeatOrderNo"
        Me.txtRepeatOrderNo.Size = New System.Drawing.Size(114, 22)
        Me.txtRepeatOrderNo.TabIndex = 12
        '
        'lblRepeatOrderNo
        '
        Me.lblRepeatOrderNo.AutoSize = True
        Me.lblRepeatOrderNo.Location = New System.Drawing.Point(6, 18)
        Me.lblRepeatOrderNo.Name = "lblRepeatOrderNo"
        Me.lblRepeatOrderNo.Size = New System.Drawing.Size(60, 16)
        Me.lblRepeatOrderNo.TabIndex = 11
        Me.lblRepeatOrderNo.Text = "Order No"
        '
        'lblRefill
        '
        Me.lblRefill.AutoSize = True
        Me.lblRefill.Location = New System.Drawing.Point(687, 54)
        Me.lblRefill.Name = "lblRefill"
        Me.lblRefill.Size = New System.Drawing.Size(36, 16)
        Me.lblRefill.TabIndex = 44
        Me.lblRefill.Text = "Refill"
        '
        'txtRefill
        '
        Me.txtRefill.Location = New System.Drawing.Point(687, 81)
        Me.txtRefill.Name = "txtRefill"
        Me.txtRefill.Size = New System.Drawing.Size(143, 22)
        Me.txtRefill.TabIndex = 43
        '
        'cboRefill
        '
        Me.cboRefill.FormattingEnabled = True
        Me.cboRefill.Location = New System.Drawing.Point(726, 51)
        Me.cboRefill.Name = "cboRefill"
        Me.cboRefill.Size = New System.Drawing.Size(104, 24)
        Me.cboRefill.TabIndex = 42
        '
        'lblPackaging
        '
        Me.lblPackaging.AutoSize = True
        Me.lblPackaging.Location = New System.Drawing.Point(506, 54)
        Me.lblPackaging.Name = "lblPackaging"
        Me.lblPackaging.Size = New System.Drawing.Size(69, 16)
        Me.lblPackaging.TabIndex = 41
        Me.lblPackaging.Text = "Packaging"
        '
        'txtPackaging
        '
        Me.txtPackaging.Location = New System.Drawing.Point(506, 81)
        Me.txtPackaging.Name = "txtPackaging"
        Me.txtPackaging.Size = New System.Drawing.Size(175, 22)
        Me.txtPackaging.TabIndex = 40
        '
        'cboPackaging
        '
        Me.cboPackaging.FormattingEnabled = True
        Me.cboPackaging.Location = New System.Drawing.Point(575, 51)
        Me.cboPackaging.Name = "cboPackaging"
        Me.cboPackaging.Size = New System.Drawing.Size(106, 24)
        Me.cboPackaging.TabIndex = 39
        '
        'txtImprintColor
        '
        Me.txtImprintColor.Location = New System.Drawing.Point(258, 81)
        Me.txtImprintColor.Name = "txtImprintColor"
        Me.txtImprintColor.Size = New System.Drawing.Size(242, 22)
        Me.txtImprintColor.TabIndex = 37
        '
        'lblImprintColour
        '
        Me.lblImprintColour.AutoSize = True
        Me.lblImprintColour.Location = New System.Drawing.Point(258, 54)
        Me.lblImprintColour.Name = "lblImprintColour"
        Me.lblImprintColour.Size = New System.Drawing.Size(38, 16)
        Me.lblImprintColour.TabIndex = 38
        Me.lblImprintColour.Text = "Color"
        '
        'cboImprintColor
        '
        Me.cboImprintColor.FormattingEnabled = True
        Me.cboImprintColor.Location = New System.Drawing.Point(299, 51)
        Me.cboImprintColor.Name = "cboImprintColor"
        Me.cboImprintColor.Size = New System.Drawing.Size(201, 24)
        Me.cboImprintColor.TabIndex = 36
        '
        'lblImprintLocation
        '
        Me.lblImprintLocation.AutoSize = True
        Me.lblImprintLocation.Location = New System.Drawing.Point(3, 54)
        Me.lblImprintLocation.Name = "lblImprintLocation"
        Me.lblImprintLocation.Size = New System.Drawing.Size(57, 16)
        Me.lblImprintLocation.TabIndex = 35
        Me.lblImprintLocation.Text = "Location"
        '
        'txtImprintLocation
        '
        Me.txtImprintLocation.Location = New System.Drawing.Point(3, 81)
        Me.txtImprintLocation.Name = "txtImprintLocation"
        Me.txtImprintLocation.Size = New System.Drawing.Size(249, 22)
        Me.txtImprintLocation.TabIndex = 34
        '
        'lblSpecialComments
        '
        Me.lblSpecialComments.AutoSize = True
        Me.lblSpecialComments.Location = New System.Drawing.Point(572, 4)
        Me.lblSpecialComments.Name = "lblSpecialComments"
        Me.lblSpecialComments.Size = New System.Drawing.Size(116, 16)
        Me.lblSpecialComments.TabIndex = 33
        Me.lblSpecialComments.Text = "Special comments"
        '
        'txtSpecialComments
        '
        Me.txtSpecialComments.Location = New System.Drawing.Point(575, 23)
        Me.txtSpecialComments.Name = "txtSpecialComments"
        Me.txtSpecialComments.Size = New System.Drawing.Size(255, 22)
        Me.txtSpecialComments.TabIndex = 32
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(299, 23)
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(270, 22)
        Me.txtComments.TabIndex = 31
        '
        'lblComments
        '
        Me.lblComments.AutoSize = True
        Me.lblComments.Location = New System.Drawing.Point(306, 4)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(71, 16)
        Me.lblComments.TabIndex = 30
        Me.lblComments.Text = "Comments"
        '
        'txtNumImprints
        '
        Me.txtNumImprints.Location = New System.Drawing.Point(218, 23)
        Me.txtNumImprints.Name = "txtNumImprints"
        Me.txtNumImprints.Size = New System.Drawing.Size(75, 22)
        Me.txtNumImprints.TabIndex = 29
        '
        'lblNumImprints
        '
        Me.lblNumImprints.AutoSize = True
        Me.lblNumImprints.Location = New System.Drawing.Point(215, 4)
        Me.lblNumImprints.Name = "lblNumImprints"
        Me.lblNumImprints.Size = New System.Drawing.Size(85, 16)
        Me.lblNumImprints.TabIndex = 28
        Me.lblNumImprints.Text = "Num Imprints"
        '
        'cboImprintLocation
        '
        Me.cboImprintLocation.FormattingEnabled = True
        Me.cboImprintLocation.Location = New System.Drawing.Point(64, 51)
        Me.cboImprintLocation.Name = "cboImprintLocation"
        Me.cboImprintLocation.Size = New System.Drawing.Size(188, 24)
        Me.cboImprintLocation.TabIndex = 27
        '
        'txtImprint
        '
        Me.txtImprint.Location = New System.Drawing.Point(3, 23)
        Me.txtImprint.Name = "txtImprint"
        Me.txtImprint.Size = New System.Drawing.Size(209, 22)
        Me.txtImprint.TabIndex = 26
        '
        'lblImprint
        '
        Me.lblImprint.AutoSize = True
        Me.lblImprint.Location = New System.Drawing.Point(3, 4)
        Me.lblImprint.Name = "lblImprint"
        Me.lblImprint.Size = New System.Drawing.Size(47, 16)
        Me.lblImprint.TabIndex = 25
        Me.lblImprint.Text = "Imprint"
        '
        'ucImprint_2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.cmdUpdateSelected)
        Me.Controls.Add(Me.cmdLockImprintData)
        Me.Controls.Add(Me.cmdUpdateAll)
        Me.Controls.Add(Me.cmdClear)
        Me.Controls.Add(Me.gbRepeatData)
        Me.Controls.Add(Me.lblRefill)
        Me.Controls.Add(Me.txtRefill)
        Me.Controls.Add(Me.cboRefill)
        Me.Controls.Add(Me.lblPackaging)
        Me.Controls.Add(Me.txtPackaging)
        Me.Controls.Add(Me.cboPackaging)
        Me.Controls.Add(Me.txtImprintColor)
        Me.Controls.Add(Me.lblImprintColour)
        Me.Controls.Add(Me.cboImprintColor)
        Me.Controls.Add(Me.lblImprintLocation)
        Me.Controls.Add(Me.txtImprintLocation)
        Me.Controls.Add(Me.lblSpecialComments)
        Me.Controls.Add(Me.txtSpecialComments)
        Me.Controls.Add(Me.txtComments)
        Me.Controls.Add(Me.lblComments)
        Me.Controls.Add(Me.txtNumImprints)
        Me.Controls.Add(Me.lblNumImprints)
        Me.Controls.Add(Me.cboImprintLocation)
        Me.Controls.Add(Me.txtImprint)
        Me.Controls.Add(Me.lblImprint)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "ucImprint_2"
        Me.Size = New System.Drawing.Size(977, 137)
        Me.gbRepeatData.ResumeLayout(False)
        Me.gbRepeatData.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdUpdateSelected As System.Windows.Forms.Button
    Friend WithEvents cmdLockImprintData As System.Windows.Forms.Button
    Friend WithEvents cmdUpdateAll As System.Windows.Forms.Button
    Friend WithEvents cmdGetRepeatData As System.Windows.Forms.Button
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents gbRepeatData As System.Windows.Forms.GroupBox
    Friend WithEvents txtRepeatOrderNo As System.Windows.Forms.TextBox
    Friend WithEvents lblRepeatOrderNo As System.Windows.Forms.Label
    Friend WithEvents lblRefill As System.Windows.Forms.Label
    Friend WithEvents txtRefill As System.Windows.Forms.TextBox
    Friend WithEvents cboRefill As System.Windows.Forms.ComboBox
    Friend WithEvents lblPackaging As System.Windows.Forms.Label
    Friend WithEvents txtPackaging As System.Windows.Forms.TextBox
    Friend WithEvents cboPackaging As System.Windows.Forms.ComboBox
    Friend WithEvents txtImprintColor As System.Windows.Forms.TextBox
    Friend WithEvents lblImprintColour As System.Windows.Forms.Label
    Friend WithEvents cboImprintColor As System.Windows.Forms.ComboBox
    Friend WithEvents lblImprintLocation As System.Windows.Forms.Label
    Friend WithEvents txtImprintLocation As System.Windows.Forms.TextBox
    Friend WithEvents lblSpecialComments As System.Windows.Forms.Label
    Friend WithEvents txtSpecialComments As System.Windows.Forms.TextBox
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents txtNumImprints As System.Windows.Forms.TextBox
    Friend WithEvents lblNumImprints As System.Windows.Forms.Label
    Friend WithEvents cboImprintLocation As System.Windows.Forms.ComboBox
    Friend WithEvents txtImprint As System.Windows.Forms.TextBox
    Friend WithEvents lblImprint As System.Windows.Forms.Label

End Class
