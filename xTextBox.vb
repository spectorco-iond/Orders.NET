Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms



Public Class xTextBox
    Inherits TextBox

    ''' Declare a local variable used to contain the current dollar value
    Private mNumeric As Decimal
    Private mString As String
    Private mDate As Date
    Private mTime As Date

    Private mySpacePadding As CSpacePadding
    Private MyInput As Cinput
    Private myDataType As CDataType
    Private myOldValue As String = ""
    Private myDataLength As Long
    Private myDecimalDigits As Long
    Private myMask As String
    ' Place to put the ID of some text, can be any format.
    Private myTextDataID As Object

    Private blnGotFocus As Boolean
    Private blnKeyPress As Boolean

    Public Property CharacterInput() As Cinput
        Get
            Return Me.MyInput
        End Get
        Set(ByVal value As Cinput)
            Me.MyInput = value
        End Set
    End Property

    Public Property DataLength() As Long
        Get
            DataLength = myDataLength
        End Get
        Set(ByVal value As Long)
            myDataLength = value
        End Set
    End Property

    Public Property DataType() As CDataType
        Get
            DataType = myDataType
        End Get
        Set(ByVal value As CDataType)
            myDataType = value
        End Set
    End Property

    Public Property DecimalDigits() As Long
        Get
            DecimalDigits = myDecimalDigits
        End Get
        Set(ByVal value As Long)
            myDecimalDigits = value
        End Set
    End Property

    Public Property TextDataID() As Object
        Get
            TextDataID = myTextDataID
        End Get
        Set(ByVal value As Object)
            myTextDataID = value
        End Set
    End Property

    Public Property Mask() As String
        Get
            Mask = myMask
        End Get
        Set(ByVal value As String)
            myMask = value
        End Set
    End Property

    Public Property OldValue() As String
        Get
            OldValue = myOldValue
        End Get
        Set(ByVal value As String)
            myOldValue = value
        End Set
    End Property

    ''' Declare public property used to get/set current dollar value
    Public Property NumericValue() As Decimal
        Get
            Return mNumeric
        End Get
        Set(ByVal value As Decimal)
            mNumeric = value
        End Set
    End Property

    Public Property SpacePadding() As CSpacePadding
        Get
            SpacePadding = mySpacePadding
        End Get
        Set(ByVal value As CSpacePadding)
            mySpacePadding = value
        End Set
    End Property

    ''' Declare public property used to get/set current dollar value
    Public Property StringValue() As String
        Get
            Return mString
        End Get
        Set(ByVal value As String)
            mString = value
        End Set
    End Property

    ''' Declare public property used to get/set current dollar value
    Public Property DateValue() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property


    ''' Default constructor; sets dollar value to zero when control is created
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Me.CausesValidation = True
        NumericValue = 0

    End Sub


    ''' Default OnPaint event
    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)
    End Sub

    Public Sub TextInit()
        Text = ""
        myTextDataID = 0
    End Sub

    Private Sub xTextBox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        If e.KeyCode = Keys.Escape Then
            Me.Text = myOldValue
        End If

    End Sub

    ''' Key press event handler or text; limits user inputs to numbers,
    ''' a single decimal point, and control characters to permit edits 
    ''' to the contents of the box (e.g., using the backspace key)
    Private Sub xTextBox_KeyPress(ByVal sender As System.Object, _
    ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress

        blnKeyPress = True

        '' allows only numbers, decimals and control characters
        'If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) _
        'And e.KeyChar <> "." Then
        '    e.Handled = True
        'End If

        '' allows only a single decimal point
        'If e.KeyChar = "." And Me.Text.Contains(".") Then
        '    e.Handled = True
        'End If

        '' denies entry of point without leading value (e.g., zero)
        '' disable this if you don't want to insist on a leading zero
        '' or other value
        'If e.KeyChar = "." And Me.Text.Length < 1 Then
        '    e.Handled = True
        'End If

        Select Case myDataType
            Case CDataType.dtCurrency, CDataType.dtNumWithDecimals, CDataType.dtNumWithoutDecimals
                ' allows only numbers, decimals and control characters
                If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) _
                And e.KeyChar <> "." Then
                    e.Handled = True
                End If

                ' allows only a single decimal point
                If e.KeyChar = "." And Me.Text.Contains(".") Then
                    e.Handled = True
                End If

                ' denies entry of point without leading value (e.g., zero)
                ' disable this if you don't want to insist on a leading zero
                ' or other value
                If e.KeyChar = "." And Me.Text.Length < 1 Then
                    e.Handled = True
                End If

            Case CDataType.dtDate
                ' allows only numbers, decimals and control characters
                If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) _
                And e.KeyChar <> "/" Then
                    e.Handled = True
                End If

                '' allows only a single decimal point
                'If e.KeyChar = "/" And Me.Text.contains("/") Then
                '    e.Handled = True
                'End If

                ' denies entry of point without leading value (e.g., zero)
                ' disable this if you don't want to insist on a leading zero
                ' or other value
                If e.KeyChar = "/" And Me.Text.Length < 1 Then
                    e.Handled = True
                End If

            Case CDataType.dtDateTime

            Case CDataType.dtString

            Case CDataType.dtTime

        End Select

    End Sub


    ' Update display to show decimal as currency whenver it is validated

    '    Private Sub xTextBox_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Validated
    Private Sub xTextBox_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Validated

        ValidateFormat()

    End Sub

    Private Sub ValidateFormat()

        Try

            Select Case myDataType
                Case CDataType.dtCurrency
                    Dim dTmp As Decimal
                    dTmp = Convert.ToDecimal(Me.Text.ToString())
                    Me.Text = dTmp.ToString("C") '  & myDecimalDigits)
                Case CDataType.dtDate
                    Dim dTmp As Date
                    If IsDate(Me.Text.ToString) Then
                        dTmp = Convert.ToDateTime(Me.Text.ToString())
                        Me.Text = String.Format("{0:MM/dd/yyyy}", CDate(dTmp.ToString))
                    Else
                        If Trim(Me.Text).Length <> 0 Then
                            MsgBox("Invalid date.")
                            Me.Text = myOldValue
                        End If
                    End If
                Case CDataType.dtDateTime
                Case CDataType.dtNumWithDecimals
                    Dim dTmp As Decimal
                    dTmp = Convert.ToDecimal(Me.Text.ToString())
                    Me.Text = dTmp.ToString("N" & myDecimalDigits)
                Case CDataType.dtNumWithoutDecimals
                    Dim dTmp As Decimal
                    dTmp = Convert.ToDecimal(Me.Text.ToString())
                    Me.Text = dTmp.ToString("G")
                Case CDataType.dtString
                    ' We force padding only if length <> Datalength.
                    If Me.Text.Length <> Me.DataLength And Me.Text.Length <> 0 Then
                        Select Case mySpacePadding
                            Case CSpacePadding.NoPadding
                                Me.Text = Trim(RTrim(Me.Text))

                            Case CSpacePadding.PaddingAfter
                                Me.Text = Trim(RTrim(Me.Text)).PadRight(Me.DataLength)

                            Case CSpacePadding.PaddingBefore
                                Me.Text = Trim(RTrim(Me.Text)).PadLeft(Me.DataLength)

                        End Select
                    End If
                Case CDataType.dtTime
                Case Else
            End Select

        Catch

            ' skip it

        End Try

    End Sub

    '''' Revert back to the display of numbers only whenever the user 
    '''' clicks in the box for editing purposes
    'Private Sub xTextBox_Click(ByVal sender As System.Object, _
    'ByVal e As System.EventArgs) Handles MyBase.Click

    '    Me.Text = NumericValue.ToString()

    '    If Me.Text = "0.00" Or Trim(Me.Text) = "" Then
    '        Me.Clear()
    '    End If

    '    Me.SelectionStart = Me.Text.Length

    'End Sub


    ''' Revert back to the display of numbers only whenever the user 
    ''' clicks in the box for editing purposes
        Private Sub xTextBox_GotFocus(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles MyBase.GotFocus

        blnGotFocus = True

        myOldValue = Me.Text

        Select Case Me.DataType

            Case CDataType.dtCurrency, CDataType.dtNumWithDecimals, CDataType.dtNumWithoutDecimals
                Me.Text = NumericValue '.ToString()
                '                If NumericValue = 0 Then
                ' Me.Text = 0
                ' End If
            Case CDataType.dtDate
                '                Me.Text = ""

            Case CDataType.dtDateTime
                '                Me.Text = ""

            Case CDataType.dtString
                Me.Text = LTrim(RTrim(Me.Text))

            Case CDataType.dtTime
                '                Me.Text = ""

        End Select

        'If Me.Text = "0.00" Then ' Or Trim(Me.Text) = "" Then
        '    Me.Clear()
        'End If

        'Me.SelectionStart = Me.Text.Length

    End Sub


    ''' Update the dollar value each time the value is changed
    Private Sub xTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged

        If Not (blnGotFocus Or blnKeyPress) Then ValidateFormat()

        Try

            Select Case myDataType
                Case CDataType.dtCurrency, CDataType.dtNumWithDecimals, CDataType.dtNumWithoutDecimals
                    Dim strValue = Me.Text

                    If Not (Me.Text.Length = 0) Then
                        '    NumericValue = 0
                        'Else
                        ' remove space and set allowed formatting
                        NumericValue = Decimal.Parse(Me.Text.Replace(" ", ""), _
                        Globalization.NumberStyles.AllowThousands Or _
                        Globalization.NumberStyles.AllowDecimalPoint Or _
                        Globalization.NumberStyles.AllowCurrencySymbol)
                    End If
                Case CDataType.dtDate
                    If Me.Text.Length = 2 Then
                        If Mid(Me.Text, 2, 1) = "/" Then
                            Me.Text = "0" & Me.Text
                        Else
                            Me.Text = Me.Text & "/"
                        End If
                        Me.SelectionStart = 4
                    End If
                    Me.Text.Replace("//", "/")
                    If Me.Text.Length = 5 Then
                        If Mid(Me.Text, 5, 1) = "/" Then
                            Me.Text = Mid(Me.Text, 1, 3) & "0" & Mid(Me.Text, 4, 2)
                        Else
                            Me.Text = Me.Text & "/"
                        End If
                        Me.SelectionStart = 7
                    End If

                Case CDataType.dtDateTime
                Case CDataType.dtString
                    If myDataLength <> 0 Then
                        If Text.Length > myDataLength Then
                            Dim lPos As Integer = SelectionStart
                            Text = Mid(Text, 1, myDataLength)
                            SelectionStart = lPos
                        End If
                    End If
                Case CDataType.dtTime
                Case Else
            End Select

        Catch
            ' skip it for formatting 
        End Try

        blnGotFocus = False
        blnKeyPress = False

    End Sub

End Class


Public Enum Cinput
    NumericOnly
    CharactersOnly
End Enum

Public Enum CDataType
    dtString
    dtNumWithDecimals
    dtNumWithoutDecimals
    dtCurrency
    dtDate
    dtTime
    dtDateTime
End Enum

Public Enum CSpacePadding
    NoPadding
    PaddingBefore
    PaddingAfter
End Enum

