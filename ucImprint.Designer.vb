<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucImprint
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
        Me.lblRefill = New System.Windows.Forms.Label
        Me.cboRefill = New System.Windows.Forms.ComboBox
        Me.lblPackaging = New System.Windows.Forms.Label
        Me.cboPackaging = New System.Windows.Forms.ComboBox
        Me.lblImprintColour = New System.Windows.Forms.Label
        Me.cboImprintColor = New System.Windows.Forms.ComboBox
        Me.lblImprintLocation = New System.Windows.Forms.Label
        Me.lblSpecialComments = New System.Windows.Forms.Label
        Me.lblComments = New System.Windows.Forms.Label
        Me.lblNumImprints = New System.Windows.Forms.Label
        Me.cboImprintLocation = New System.Windows.Forms.ComboBox
        Me.lblImprint = New System.Windows.Forms.Label
        Me.lblIndustry = New System.Windows.Forms.Label
        Me.cboIndustry = New System.Windows.Forms.ComboBox
        Me.txtImprint = New Orders.xTextBox
        Me.txtNumImprints = New Orders.xTextBox
        Me.txtComments = New Orders.xTextBox
        Me.txtSpecialComments = New Orders.xTextBox
        Me.txtImprintLocation = New Orders.xTextBox
        Me.txtImprintColor = New Orders.xTextBox
        Me.txtPackaging = New Orders.xTextBox
        Me.txtRefill = New Orders.xTextBox
        Me.txtIndustry = New Orders.xTextBox
        Me.SuspendLayout()
        '
        'lblRefill
        '
        Me.lblRefill.AutoSize = True
        Me.lblRefill.Location = New System.Drawing.Point(501, 44)
        Me.lblRefill.Name = "lblRefill"
        Me.lblRefill.Size = New System.Drawing.Size(36, 16)
        Me.lblRefill.TabIndex = 17
        Me.lblRefill.Text = "Refill"
        '
        'cboRefill
        '
        Me.cboRefill.FormattingEnabled = True
        Me.cboRefill.Location = New System.Drawing.Point(501, 60)
        Me.cboRefill.Name = "cboRefill"
        Me.cboRefill.Size = New System.Drawing.Size(160, 24)
        Me.cboRefill.TabIndex = 18
        '
        'lblPackaging
        '
        Me.lblPackaging.AutoSize = True
        Me.lblPackaging.Location = New System.Drawing.Point(335, 44)
        Me.lblPackaging.Name = "lblPackaging"
        Me.lblPackaging.Size = New System.Drawing.Size(69, 16)
        Me.lblPackaging.TabIndex = 14
        Me.lblPackaging.Text = "Packaging"
        '
        'cboPackaging
        '
        Me.cboPackaging.FormattingEnabled = True
        Me.cboPackaging.Location = New System.Drawing.Point(335, 60)
        Me.cboPackaging.Name = "cboPackaging"
        Me.cboPackaging.Size = New System.Drawing.Size(160, 24)
        Me.cboPackaging.TabIndex = 15
        '
        'lblImprintColour
        '
        Me.lblImprintColour.AutoSize = True
        Me.lblImprintColour.Location = New System.Drawing.Point(169, 44)
        Me.lblImprintColour.Name = "lblImprintColour"
        Me.lblImprintColour.Size = New System.Drawing.Size(38, 16)
        Me.lblImprintColour.TabIndex = 11
        Me.lblImprintColour.Text = "Color"
        '
        'cboImprintColor
        '
        Me.cboImprintColor.FormattingEnabled = True
        Me.cboImprintColor.Location = New System.Drawing.Point(169, 60)
        Me.cboImprintColor.Name = "cboImprintColor"
        Me.cboImprintColor.Size = New System.Drawing.Size(160, 24)
        Me.cboImprintColor.TabIndex = 12
        '
        'lblImprintLocation
        '
        Me.lblImprintLocation.AutoSize = True
        Me.lblImprintLocation.Location = New System.Drawing.Point(3, 44)
        Me.lblImprintLocation.Name = "lblImprintLocation"
        Me.lblImprintLocation.Size = New System.Drawing.Size(57, 16)
        Me.lblImprintLocation.TabIndex = 8
        Me.lblImprintLocation.Text = "Location"
        '
        'lblSpecialComments
        '
        Me.lblSpecialComments.AutoSize = True
        Me.lblSpecialComments.Location = New System.Drawing.Point(572, 3)
        Me.lblSpecialComments.Name = "lblSpecialComments"
        Me.lblSpecialComments.Size = New System.Drawing.Size(116, 16)
        Me.lblSpecialComments.TabIndex = 6
        Me.lblSpecialComments.Text = "Special comments"
        '
        'lblComments
        '
        Me.lblComments.AutoSize = True
        Me.lblComments.Location = New System.Drawing.Point(306, 3)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(71, 16)
        Me.lblComments.TabIndex = 4
        Me.lblComments.Text = "Comments"
        '
        'lblNumImprints
        '
        Me.lblNumImprints.AutoSize = True
        Me.lblNumImprints.Location = New System.Drawing.Point(215, 3)
        Me.lblNumImprints.Name = "lblNumImprints"
        Me.lblNumImprints.Size = New System.Drawing.Size(85, 16)
        Me.lblNumImprints.TabIndex = 2
        Me.lblNumImprints.Text = "Num Imprints"
        '
        'cboImprintLocation
        '
        Me.cboImprintLocation.FormattingEnabled = True
        Me.cboImprintLocation.Location = New System.Drawing.Point(3, 60)
        Me.cboImprintLocation.Name = "cboImprintLocation"
        Me.cboImprintLocation.Size = New System.Drawing.Size(160, 24)
        Me.cboImprintLocation.TabIndex = 9
        '
        'lblImprint
        '
        Me.lblImprint.AutoSize = True
        Me.lblImprint.Location = New System.Drawing.Point(3, 3)
        Me.lblImprint.Name = "lblImprint"
        Me.lblImprint.Size = New System.Drawing.Size(47, 16)
        Me.lblImprint.TabIndex = 0
        Me.lblImprint.Text = "Imprint"
        '
        'lblIndustry
        '
        Me.lblIndustry.AutoSize = True
        Me.lblIndustry.Location = New System.Drawing.Point(667, 44)
        Me.lblIndustry.Name = "lblIndustry"
        Me.lblIndustry.Size = New System.Drawing.Size(54, 16)
        Me.lblIndustry.TabIndex = 20
        Me.lblIndustry.Text = "Industry"
        '
        'cboIndustry
        '
        Me.cboIndustry.FormattingEnabled = True
        Me.cboIndustry.Location = New System.Drawing.Point(667, 60)
        Me.cboIndustry.Name = "cboIndustry"
        Me.cboIndustry.Size = New System.Drawing.Size(160, 24)
        Me.cboIndustry.TabIndex = 21
        '
        'txtImprint
        '
        Me.txtImprint.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtImprint.DataLength = CType(0, Long)
        Me.txtImprint.DataType = Orders.CDataType.dtString
        Me.txtImprint.DateValue = New Date(CType(0, Long))
        Me.txtImprint.DecimalDigits = CType(0, Long)
        Me.txtImprint.Location = New System.Drawing.Point(3, 19)
        Me.txtImprint.Mask = Nothing
        Me.txtImprint.Name = "txtImprint"
        Me.txtImprint.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtImprint.OldValue = ""
        Me.txtImprint.Size = New System.Drawing.Size(209, 22)
        Me.txtImprint.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtImprint.StringValue = Nothing
        Me.txtImprint.TabIndex = 1
        Me.txtImprint.TextDataID = Nothing
        '
        'txtNumImprints
        '
        Me.txtNumImprints.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtNumImprints.DataLength = CType(0, Long)
        Me.txtNumImprints.DataType = Orders.CDataType.dtString
        Me.txtNumImprints.DateValue = New Date(CType(0, Long))
        Me.txtNumImprints.DecimalDigits = CType(0, Long)
        Me.txtNumImprints.Location = New System.Drawing.Point(218, 19)
        Me.txtNumImprints.Mask = Nothing
        Me.txtNumImprints.Name = "txtNumImprints"
        Me.txtNumImprints.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtNumImprints.OldValue = ""
        Me.txtNumImprints.Size = New System.Drawing.Size(75, 22)
        Me.txtNumImprints.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtNumImprints.StringValue = Nothing
        Me.txtNumImprints.TabIndex = 3
        Me.txtNumImprints.TextDataID = Nothing
        '
        'txtComments
        '
        Me.txtComments.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtComments.DataLength = CType(0, Long)
        Me.txtComments.DataType = Orders.CDataType.dtString
        Me.txtComments.DateValue = New Date(CType(0, Long))
        Me.txtComments.DecimalDigits = CType(0, Long)
        Me.txtComments.Location = New System.Drawing.Point(299, 19)
        Me.txtComments.Mask = Nothing
        Me.txtComments.Name = "txtComments"
        Me.txtComments.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtComments.OldValue = ""
        Me.txtComments.Size = New System.Drawing.Size(270, 22)
        Me.txtComments.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtComments.StringValue = Nothing
        Me.txtComments.TabIndex = 5
        Me.txtComments.TextDataID = Nothing
        '
        'txtSpecialComments
        '
        Me.txtSpecialComments.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtSpecialComments.DataLength = CType(0, Long)
        Me.txtSpecialComments.DataType = Orders.CDataType.dtString
        Me.txtSpecialComments.DateValue = New Date(CType(0, Long))
        Me.txtSpecialComments.DecimalDigits = CType(0, Long)
        Me.txtSpecialComments.Location = New System.Drawing.Point(575, 19)
        Me.txtSpecialComments.Mask = Nothing
        Me.txtSpecialComments.Name = "txtSpecialComments"
        Me.txtSpecialComments.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtSpecialComments.OldValue = ""
        Me.txtSpecialComments.Size = New System.Drawing.Size(252, 22)
        Me.txtSpecialComments.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtSpecialComments.StringValue = Nothing
        Me.txtSpecialComments.TabIndex = 7
        Me.txtSpecialComments.TextDataID = Nothing
        '
        'txtImprintLocation
        '
        Me.txtImprintLocation.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtImprintLocation.DataLength = CType(0, Long)
        Me.txtImprintLocation.DataType = Orders.CDataType.dtString
        Me.txtImprintLocation.DateValue = New Date(CType(0, Long))
        Me.txtImprintLocation.DecimalDigits = CType(0, Long)
        Me.txtImprintLocation.Location = New System.Drawing.Point(3, 86)
        Me.txtImprintLocation.Mask = Nothing
        Me.txtImprintLocation.Name = "txtImprintLocation"
        Me.txtImprintLocation.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtImprintLocation.OldValue = ""
        Me.txtImprintLocation.Size = New System.Drawing.Size(160, 22)
        Me.txtImprintLocation.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtImprintLocation.StringValue = Nothing
        Me.txtImprintLocation.TabIndex = 10
        Me.txtImprintLocation.TextDataID = Nothing
        '
        'txtImprintColor
        '
        Me.txtImprintColor.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtImprintColor.DataLength = CType(0, Long)
        Me.txtImprintColor.DataType = Orders.CDataType.dtString
        Me.txtImprintColor.DateValue = New Date(CType(0, Long))
        Me.txtImprintColor.DecimalDigits = CType(0, Long)
        Me.txtImprintColor.Location = New System.Drawing.Point(169, 86)
        Me.txtImprintColor.Mask = Nothing
        Me.txtImprintColor.Name = "txtImprintColor"
        Me.txtImprintColor.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtImprintColor.OldValue = ""
        Me.txtImprintColor.Size = New System.Drawing.Size(160, 22)
        Me.txtImprintColor.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtImprintColor.StringValue = Nothing
        Me.txtImprintColor.TabIndex = 13
        Me.txtImprintColor.TextDataID = Nothing
        '
        'txtPackaging
        '
        Me.txtPackaging.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtPackaging.DataLength = CType(0, Long)
        Me.txtPackaging.DataType = Orders.CDataType.dtString
        Me.txtPackaging.DateValue = New Date(CType(0, Long))
        Me.txtPackaging.DecimalDigits = CType(0, Long)
        Me.txtPackaging.Location = New System.Drawing.Point(335, 86)
        Me.txtPackaging.Mask = Nothing
        Me.txtPackaging.Name = "txtPackaging"
        Me.txtPackaging.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtPackaging.OldValue = ""
        Me.txtPackaging.Size = New System.Drawing.Size(160, 22)
        Me.txtPackaging.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtPackaging.StringValue = Nothing
        Me.txtPackaging.TabIndex = 16
        Me.txtPackaging.TextDataID = Nothing
        '
        'txtRefill
        '
        Me.txtRefill.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtRefill.DataLength = CType(0, Long)
        Me.txtRefill.DataType = Orders.CDataType.dtString
        Me.txtRefill.DateValue = New Date(CType(0, Long))
        Me.txtRefill.DecimalDigits = CType(0, Long)
        Me.txtRefill.Location = New System.Drawing.Point(501, 86)
        Me.txtRefill.Mask = Nothing
        Me.txtRefill.Name = "txtRefill"
        Me.txtRefill.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtRefill.OldValue = ""
        Me.txtRefill.Size = New System.Drawing.Size(160, 22)
        Me.txtRefill.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtRefill.StringValue = Nothing
        Me.txtRefill.TabIndex = 19
        Me.txtRefill.TextDataID = Nothing
        '
        'txtIndustry
        '
        Me.txtIndustry.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtIndustry.DataLength = CType(0, Long)
        Me.txtIndustry.DataType = Orders.CDataType.dtString
        Me.txtIndustry.DateValue = New Date(CType(0, Long))
        Me.txtIndustry.DecimalDigits = CType(0, Long)
        Me.txtIndustry.Location = New System.Drawing.Point(667, 86)
        Me.txtIndustry.Mask = Nothing
        Me.txtIndustry.Name = "txtIndustry"
        Me.txtIndustry.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtIndustry.OldValue = ""
        Me.txtIndustry.Size = New System.Drawing.Size(160, 22)
        Me.txtIndustry.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtIndustry.StringValue = Nothing
        Me.txtIndustry.TabIndex = 22
        Me.txtIndustry.TextDataID = Nothing
        '
        'ucImprint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.txtIndustry)
        Me.Controls.Add(Me.txtRefill)
        Me.Controls.Add(Me.txtPackaging)
        Me.Controls.Add(Me.txtImprintColor)
        Me.Controls.Add(Me.txtImprintLocation)
        Me.Controls.Add(Me.txtSpecialComments)
        Me.Controls.Add(Me.txtComments)
        Me.Controls.Add(Me.txtNumImprints)
        Me.Controls.Add(Me.txtImprint)
        Me.Controls.Add(Me.lblIndustry)
        Me.Controls.Add(Me.cboIndustry)
        Me.Controls.Add(Me.lblRefill)
        Me.Controls.Add(Me.cboRefill)
        Me.Controls.Add(Me.lblPackaging)
        Me.Controls.Add(Me.cboPackaging)
        Me.Controls.Add(Me.lblImprintColour)
        Me.Controls.Add(Me.cboImprintColor)
        Me.Controls.Add(Me.lblImprintLocation)
        Me.Controls.Add(Me.lblSpecialComments)
        Me.Controls.Add(Me.lblComments)
        Me.Controls.Add(Me.lblNumImprints)
        Me.Controls.Add(Me.cboImprintLocation)
        Me.Controls.Add(Me.lblImprint)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "ucImprint"
        Me.Size = New System.Drawing.Size(980, 555)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblRefill As System.Windows.Forms.Label
    Friend WithEvents cboRefill As System.Windows.Forms.ComboBox
    Friend WithEvents lblPackaging As System.Windows.Forms.Label
    Friend WithEvents cboPackaging As System.Windows.Forms.ComboBox
    Friend WithEvents lblImprintColour As System.Windows.Forms.Label
    Friend WithEvents cboImprintColor As System.Windows.Forms.ComboBox
    Friend WithEvents lblImprintLocation As System.Windows.Forms.Label
    Friend WithEvents lblSpecialComments As System.Windows.Forms.Label
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents lblNumImprints As System.Windows.Forms.Label
    Friend WithEvents cboImprintLocation As System.Windows.Forms.ComboBox
    Friend WithEvents lblImprint As System.Windows.Forms.Label
    Friend WithEvents lblIndustry As System.Windows.Forms.Label
    Friend WithEvents cboIndustry As System.Windows.Forms.ComboBox
    Friend WithEvents txtImprint As Orders.xTextBox
    Friend WithEvents txtNumImprints As Orders.xTextBox
    Friend WithEvents txtComments As Orders.xTextBox
    Friend WithEvents txtSpecialComments As Orders.xTextBox
    Friend WithEvents txtImprintLocation As Orders.xTextBox
    Friend WithEvents txtImprintColor As Orders.xTextBox
    Friend WithEvents txtPackaging As Orders.xTextBox
    Friend WithEvents txtRefill As Orders.xTextBox
    Friend WithEvents txtIndustry As Orders.xTextBox

End Class
