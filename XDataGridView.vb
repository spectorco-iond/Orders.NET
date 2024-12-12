Imports System
Imports System.Windows.Forms

Public Class XDataGridView
    Inherits System.Windows.Forms.DataGridView

    '    Protected Overrides ProcessCmdKey(byref msg as System.Windows.Forms.Message , byref keyData as System.Windows.Forms.Keys ) as boolean
    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean

        Select Case keyData ' Keys.ControlKey
            Case Keys.F7
                'MsgBox("F7")

            Case Keys.F8
                'MsgBox("F8")

            Case Keys.F9
                'MsgBox("F9")

            Case Keys.Escape
                'MsgBox("F9")

            Case Else

        End Select

        ProcessCmdKey = MyBase.ProcessCmdKey(msg, keyData)

    End Function

    '    Public Overrides Property CellTemplate() As DataGridViewCell
    '        Get
    '            Return base.CellTemplate
    '        End Get
    '        Set(ByVal value As DataGridViewCell)

    '            If Not (value Is Nothing) And Not value.GetType().IsAssignableFrom(GetType(xTextBox)) Then

    '                Throw New InvalidCastException("Expected a TextCell")

    '            End If

    '            TestBase.CellTemplate = value

    '        End Set

    '    End Property

    'Read more: How to Add a Text Box to DataGridView | eHow.com http://www.ehow.com/how_8109268_add-text-box-datagridview.html#ixzz1ZH0gipEt

End Class




'public class DataGridViewXTextBoxColumn : DataGridViewTextBoxColumn 
'    public DataGridViewXTextBoxColumn() : base() { 
'        CellTemplate = new DataGridViewXTextBoxCell(); 
'    } 
'End Class

'public class DataGridViewXTextBoxCell : DataGridViewTextBoxCell 
'    public DataGridViewXTextBoxCell() : base() 
'    Public override Type EditType
'        get { 
'            EditType= typeof(DataGridViewXTextBoxEditingControl); 
'       end get
'    } 
'End Class

'public class DataGridViewXTextBoxCellEditingControl : DataGridViewTextBoxEditingControl { 
'    public DataGridViewUpperCaseTextBoxEditingControl() : base() { 
'        this.CharacterCasing = CharacterCasing.Upper; 
'    } 
'}
