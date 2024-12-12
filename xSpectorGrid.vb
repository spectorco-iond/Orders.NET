Imports System
Imports System.Windows.Forms

Public Class xSpectorGrid
    Inherits System.Windows.Forms.DataGridView

    Public Class cSPT_SPECTOR_GRID_FIELDS

        Private m_Spector_Grid_Field_ID As Integer = 0
        Private m_Spector_Grid_Setup_ID As Integer
        Private m_Field_Name As String
        Private m_Field_Label As String
        Private m_Col_Pos As Integer
        Private m_Show_Col As Boolean
        Private m_Sort_Level As Integer
        Private m_Sort_Order As Boolean
        Private m_Cell_FontName As String
        Private m_Cell_FontSize As Integer
        Private m_Cell_FontSpecial As Integer
        Private m_Cell_TextColor As Integer
        Private m_Cell_BackColor As Integer
        Private m_Cell_Align As Integer
        Private m_Cell_Format As String
        Private m_Cell_Width As Integer
        Private m_User_Login As String
        Private m_Update_TS As DateTime
        Private m_DataTable As DataTable

#Region "Public constructors ##############################################"

        'Public Sub New()
        Public Sub New()

            Call Init()

        End Sub

        'Public Sub New(ByVal pID As Integer)
        Public Sub New(ByVal pSpector_Grid_Field_ID As Integer)

            Try

                Call Init()

                m_Spector_Grid_Field_ID = pSpector_Grid_Field_ID

                Call Load(pSpector_Grid_Field_ID)

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_FIELDS." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

#End Region

#Region "Private maintenance procedures ###################################"

        Private Sub Init()

            Try

                m_Spector_Grid_Field_ID = 0

                Call Reset()

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_FIELDS." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Load is made private, it is loaded automatically on New.
        Private Sub Load(ByVal pintSpector_Grid_Field_ID As Integer)

            Try
                Dim db As New cDBA
                Dim dt As New DataTable

                Dim strSql As String
                strSql = _
                "SELECT * " & _
                "FROM   SPT_SPECTOR_GRID_FIELDS PT WITH (Nolock) " & _
                "WHERE  PT.Spector_Grid_Field_ID = " & pintSpector_Grid_Field_ID & " "

                dt = db.DataTable(strSql)

                Call LoadLine(dt.Rows(0))

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_FIELDS." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Load is made private, it is loaded automatically on New.
        ' Loads the Comment fields from the DataRow record
        Private Sub LoadLine(ByRef pdrRow As DataRow)

            Try
                ' Loads the item from the DataRow into the locals
                With pdrRow
                    If Not (.Item("Spector_Grid_Field_ID").Equals(DBNull.Value)) Then m_Spector_Grid_Field_ID = .Item("Spector_Grid_Field_ID")
                    If Not (.Item("Spector_Grid_Setup_ID").Equals(DBNull.Value)) Then m_Spector_Grid_Setup_ID = .Item("Spector_Grid_Setup_ID")
                    If Not (.Item("Col_Pos").Equals(DBNull.Value)) Then m_Col_Pos = .Item("Col_Pos")
                    If Not (.Item("Field_Name").Equals(DBNull.Value)) Then m_Field_Name = .Item("Field_Name").ToString
                    If Not (.Item("Field_Label").Equals(DBNull.Value)) Then m_Field_Label = .Item("Field_Label").ToString

                    If Not (.Item("Show_Col").Equals(DBNull.Value)) Then m_Show_Col = .Item("Show_Col")
                    If Not (.Item("Sort_Level").Equals(DBNull.Value)) Then m_Sort_Level = .Item("Sort_Level")
                    If Not (.Item("Sort_Order").Equals(DBNull.Value)) Then m_Sort_Order = .Item("Sort_Order")
                    If Not (.Item("Cell_FontName").Equals(DBNull.Value)) Then m_Cell_FontName = .Item("Cell_FontName").ToString
                    If Not (.Item("Cell_FontSize").Equals(DBNull.Value)) Then m_Cell_FontSize = .Item("Cell_FontSize")
                    If Not (.Item("Cell_FontSpecial").Equals(DBNull.Value)) Then m_Cell_FontSpecial = .Item("Cell_FontSpecial")
                    If Not (.Item("Cell_TextColor").Equals(DBNull.Value)) Then m_Cell_TextColor = .Item("Cell_TextColor")
                    If Not (.Item("Cell_BackColor").Equals(DBNull.Value)) Then m_Cell_BackColor = .Item("Cell_BackColor")
                    If Not (.Item("Cell_Align").Equals(DBNull.Value)) Then m_Cell_Align = .Item("Cell_Align")
                    If Not (.Item("Cell_Format").Equals(DBNull.Value)) Then m_Cell_Format = .Item("Cell_Format").ToString
                    If Not (.Item("Cell_Width").Equals(DBNull.Value)) Then m_Cell_Width = .Item("Cell_Width")

                    If Not (.Item("User_Login").Equals(DBNull.Value)) Then m_User_Login = .Item("User_Login").ToString
                    If Not (.Item("Update_TS").Equals(DBNull.Value)) Then m_Update_TS = .Item("Update_TS")

                End With

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_FIELDS." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Saves the Comment fields into the DataRow record
        Private Sub SaveLine(ByRef pdrRow As DataRow)

            Try

                ' Save the locals into the datarow.
                pdrRow.Item("Spector_Grid_Setup_ID") = m_Spector_Grid_Setup_ID
                pdrRow.Item("Field_Name") = m_Field_Name
                pdrRow.Item("Field_Label") = m_Field_Label
                pdrRow.Item("Col_Pos") = m_Col_Pos
                pdrRow.Item("Show_Col") = m_Show_Col
                pdrRow.Item("Sort_Level") = m_Sort_Level
                pdrRow.Item("Sort_Order") = m_Sort_Order
                pdrRow.Item("Cell_FontName") = m_Cell_FontName
                pdrRow.Item("Cell_FontSize") = m_Cell_FontSize
                pdrRow.Item("Cell_FontSpecial") = m_Cell_FontSpecial
                pdrRow.Item("Cell_TextColor") = m_Cell_TextColor
                pdrRow.Item("Cell_BackColor") = m_Cell_BackColor
                pdrRow.Item("Cell_Align") = m_Cell_Align
                pdrRow.Item("Cell_Format") = m_Cell_Format
                pdrRow.Item("Cell_Width") = m_Cell_Width
                pdrRow.Item("User_Login") = m_User_Login
                pdrRow.Item("Update_TS") = Date.Now()

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_FIELDS." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

#End Region

#Region "Public maintenance procedures ####################################"

        ' Deletes the current Comment from the database, if it exists.
        Public Sub Delete()

            Try

                If m_Spector_Grid_Field_ID = 0 Then Exit Sub

                Dim db As New cDBA
                Dim dt As New DataTable

                Dim strSql As String

                strSql = _
                "DELETE FROM SPT_SPECTOR_GRID_FIELDS " & _
                "WHERE  Spector_Grid_Field_ID = " & m_Spector_Grid_Field_ID & " "

                db.Execute(strSql)

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_FIELDS." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Reset every non mandatory field of a Comment
        Public Sub Reset()

            Try
                ' Resets the locals to empty values, all but ID, Is_Mandatory and Default_Pct
                m_Spector_Grid_Setup_ID = 0
                m_Field_Name = String.Empty
                m_Col_Pos = 0
                m_Field_Label = String.Empty
                m_Show_Col = False

                m_Sort_Level = 0
                m_Sort_Order = False
                m_Cell_FontName = String.Empty
                m_Cell_FontSize = 0
                m_Cell_FontSpecial = 0
                m_Cell_TextColor = 0
                m_Cell_BackColor = 0
                m_Cell_Align = 0
                m_Cell_Format = String.Empty
                m_Cell_Width = 0

                m_User_Login = String.Empty
                m_Update_TS = Nothing

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_FIELDS." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Update the current Comment into the database, or creates it if not existing
        Public Sub Save()

            Try

                Dim db As New cDBA
                Dim dt As New DataTable
                Dim drRow As DataRow

                Dim strSql As String

                strSql = _
                "SELECT * " & _
                "FROM   SPT_SPECTOR_GRID_FIELDS " & _
                "WHERE  Spector_Grid_Field_ID = " & m_Spector_Grid_Field_ID & " "

                dt = db.DataTable(strSql)

                If dt.Rows.Count = 0 Then
                    drRow = dt.NewRow()
                Else
                    drRow = dt.Rows(0)
                End If
                'drRow = IIf(dt.Rows.Count = 0, dt.NewRow(), dt.Rows(0))

                Call SaveLine(drRow)

                If dt.Rows.Count = 0 Then
                    db.DBDataTable.Rows.Add(drRow)
                    Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                    db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
                Else
                    Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                    db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
                End If

                db.DBDataAdapter.Update(db.DBDataTable)

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_FIELDS." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

#End Region

#Region "Public properties ################################################"

        Public Property Spector_Grid_Field_ID() As Integer
            Get
                Spector_Grid_Field_ID = m_Spector_Grid_Field_ID
            End Get
            Set(ByVal value As Integer)
                m_Spector_Grid_Field_ID = value
            End Set
        End Property
        Public Property Spector_Grid_Setup_ID() As Integer
            Get
                Spector_Grid_Setup_ID = m_Spector_Grid_Setup_ID
            End Get
            Set(ByVal value As Integer)
                m_Spector_Grid_Setup_ID = value
            End Set
        End Property
        Public Property Col_Pos() As Integer
            Get
                Col_Pos = m_Col_Pos
            End Get
            Set(ByVal value As Integer)
                m_Col_Pos = value
            End Set
        End Property
        Public Property Show_Col() As Boolean
            Get
                Show_Col = m_Show_Col
            End Get
            Set(ByVal value As Boolean)
                m_Show_Col = value
            End Set
        End Property
        Public Property Sort_Level() As Integer
            Get
                Sort_Level = m_Sort_Level
            End Get
            Set(ByVal value As Integer)
                m_Sort_Level = value
            End Set
        End Property
        Public Property Sort_Order() As Boolean
            Get
                Sort_Order = m_Sort_Order
            End Get
            Set(ByVal value As Boolean)
                m_Sort_Order = value
            End Set
        End Property
        Public Property Cell_FontName() As String
            Get
                Cell_FontName = m_Cell_FontName
            End Get
            Set(ByVal value As String)
                m_Cell_FontName = value
            End Set
        End Property
        Public Property Cell_FontSize() As Integer
            Get
                Cell_FontSize = m_Cell_FontSize
            End Get
            Set(ByVal value As Integer)
                m_Cell_FontSize = value
            End Set
        End Property
        Public Property Cell_FontSpecial() As Integer
            Get
                Cell_FontSpecial = m_Cell_FontSpecial
            End Get
            Set(ByVal value As Integer)
                m_Cell_FontSpecial = value
            End Set
        End Property
        Public Property Cell_TextColor() As Integer
            Get
                Cell_TextColor = m_Cell_TextColor
            End Get
            Set(ByVal value As Integer)
                m_Cell_TextColor = value
            End Set
        End Property
        Public Property Cell_BackColor() As Integer
            Get
                Cell_BackColor = m_Cell_BackColor
            End Get
            Set(ByVal value As Integer)
                m_Cell_BackColor = value
            End Set
        End Property
        Public Property Cell_Align() As Integer
            Get
                Cell_Align = m_Cell_Align
            End Get
            Set(ByVal value As Integer)
                m_Cell_Align = value
            End Set
        End Property
        Public Property Cell_Format() As String
            Get
                Cell_Format = m_Cell_Format
            End Get
            Set(ByVal value As String)
                m_Cell_Format = value
            End Set
        End Property
        Public Property Cell_Width() As Integer
            Get
                Cell_Width = m_Cell_Width
            End Get
            Set(ByVal value As Integer)
                m_Cell_Width = value
            End Set
        End Property
        Public Property Field_Name() As String
            Get
                Field_Name = m_Field_Name
            End Get
            Set(ByVal value As String)
                m_Field_Name = value
            End Set
        End Property
        Public Property Field_Label() As Boolean
            Get
                Field_Label = m_Field_Label
            End Get
            Set(ByVal value As Boolean)
                m_Field_Label = value
            End Set
        End Property
        Public Property User_Login() As String
            Get
                User_Login = m_User_Login
            End Get
            Set(ByVal value As String)
                m_User_Login = value
            End Set
        End Property
        Public Property Update_TS() As DateTime
            Get
                Update_TS = m_Update_TS
            End Get
            Set(ByVal value As DateTime)
                m_Update_TS = value
            End Set
        End Property

#End Region

    End Class

    Public Class cSPT_SPECTOR_GRID_SETUP

        Private m_Spector_Grid_Setup_ID As Integer = 0
        Private m_Dept_ID As Integer
        Private m_Table_Name As String
        Private m_Can_Search As Boolean
        Private m_Can_Read As Boolean
        Private m_Can_Write As Boolean
        Private m_Can_Sort As Boolean
        Private m_Can_Delete As Boolean
        Private m_Insert_Field As String
        Private m_Row_Head_Height As Integer
        Private m_Col_Head_Width As Integer
        Private m_User_Login As String
        Private m_Update_TS As DateTime
        Private m_DataTable As DataTable

#Region "Public constructors ##############################################"

        'Public Sub New()
        Public Sub New()

            Call Init()

        End Sub

        'Public Sub New(ByVal pID As Integer)
        Public Sub New(ByVal pSpector_Grid_Setup_ID As Integer)

            Try

                Call Init()

                m_Spector_Grid_Setup_ID = pSpector_Grid_Setup_ID

                Call Load(pSpector_Grid_Setup_ID)

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_SETUP." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

#End Region

#Region "Private maintenance procedures ###################################"

        Private Sub Init()

            Try

                m_Spector_Grid_Setup_ID = 0

                Call Reset()

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_SETUP." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Load is made private, it is loaded automatically on New.
        Private Sub Load(ByVal pintSpector_Grid_Setup_ID As Integer)

            Try
                Dim db As New cDBA
                Dim dt As New DataTable

                Dim strSql As String
                strSql = _
                "SELECT * " & _
                "FROM   SPT_SPECTOR_GRID_SETUP PT WITH (Nolock) " & _
                "WHERE  PT.Spector_Grid_Setup_ID = " & pintSpector_Grid_Setup_ID & " "

                dt = db.DataTable(strSql)

                Call LoadLine(dt.Rows(0))

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_SETUP." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Load is made private, it is loaded automatically on New.
        ' Loads the Comment fields from the DataRow record
        Private Sub LoadLine(ByRef pdrRow As DataRow)

            Try
                ' Loads the item from the DataRow into the locals
                With pdrRow
                    If Not (.Item("Spector_Grid_Setup_ID").Equals(DBNull.Value)) Then m_Spector_Grid_Setup_ID = .Item("Spector_Grid_Setup_ID")
                    If Not (.Item("Dept_ID").Equals(DBNull.Value)) Then m_Dept_ID = .Item("Dept_ID")
                    If Not (.Item("Table_Name").Equals(DBNull.Value)) Then m_Table_Name = .Item("Table_Name").ToString

                    If Not (.Item("Can_Search").Equals(DBNull.Value)) Then m_Can_Search = .Item("Can_Search")
                    If Not (.Item("Can_Read").Equals(DBNull.Value)) Then m_Can_Read = .Item("Can_Read")
                    If Not (.Item("Can_Write").Equals(DBNull.Value)) Then m_Can_Write = .Item("Can_Write")
                    If Not (.Item("Can_Sort").Equals(DBNull.Value)) Then m_Can_Sort = .Item("Can_Sort")
                    If Not (.Item("Can_Delete").Equals(DBNull.Value)) Then m_Can_Delete = .Item("Can_Delete")
                    If Not (.Item("Row_Head_Height").Equals(DBNull.Value)) Then m_Row_Head_Height = .Item("Row_Head_Height")
                    If Not (.Item("Col_Head_Width").Equals(DBNull.Value)) Then m_Col_Head_Width = .Item("Col_Head_Width")

                    If Not (.Item("Insert_Field").Equals(DBNull.Value)) Then m_Insert_Field = .Item("Insert_Field").ToString

                    If Not (.Item("User_Login").Equals(DBNull.Value)) Then m_User_Login = .Item("User_Login").ToString
                    If Not (.Item("Update_TS").Equals(DBNull.Value)) Then m_Update_TS = .Item("Update_TS")

                End With

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_SETUP." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Saves the Comment fields into the DataRow record
        Private Sub SaveLine(ByRef pdrRow As DataRow)

            Try

                ' Save the locals into the datarow.
                pdrRow.Item("Dept_ID") = m_Dept_ID
                pdrRow.Item("Table_Name") = m_Table_Name
                pdrRow.Item("Insert_Field") = m_Insert_Field
                pdrRow.Item("Can_Search") = m_Can_Search
                pdrRow.Item("Can_Read") = m_Can_Read
                pdrRow.Item("Can_Write") = m_Can_Write
                pdrRow.Item("Can_Sort") = m_Can_Sort
                pdrRow.Item("Can_Delete") = m_Can_Delete
                pdrRow.Item("Row_Head_Height") = m_Row_Head_Height
                pdrRow.Item("Col_Head_Width") = m_Col_Head_Width
                pdrRow.Item("User_Login") = m_User_Login
                pdrRow.Item("Update_TS") = Date.Now()

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_SETUP." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

#End Region

#Region "Public maintenance procedures ####################################"

        ' Deletes the current Comment from the database, if it exists.
        Public Sub Delete()

            Try

                If m_Spector_Grid_Setup_ID = 0 Then Exit Sub

                Dim db As New cDBA
                Dim dt As New DataTable

                Dim strSql As String

                strSql = _
                "DELETE FROM SPT_SPECTOR_GRID_SETUP " & _
                "WHERE  Spector_Grid_Setup_ID = " & m_Spector_Grid_Setup_ID & " "

                db.Execute(strSql)

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_SETUP." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Reset every non mandatory field of a Comment
        Public Sub Reset()

            Try
                ' Resets the locals to empty values, all but ID, Is_Mandatory and Default_Pct
                m_Dept_ID = 0
                m_Table_Name = String.Empty

                m_Can_Search = False
                m_Can_Read = False
                m_Can_Write = False
                m_Can_Sort = False
                m_Can_Delete = False
                m_Row_Head_Height = 0
                m_Col_Head_Width = 0

                m_Insert_Field = String.Empty

                m_User_Login = String.Empty
                m_Update_TS = Nothing

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_SETUP." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        ' Update the current Comment into the database, or creates it if not existing
        Public Sub Save()

            Try

                Dim db As New cDBA
                Dim dt As New DataTable
                Dim drRow As DataRow

                Dim strSql As String

                strSql = _
                "SELECT * " & _
                "FROM   SPT_SPECTOR_GRID_SETUP " & _
                "WHERE  Spector_Grid_Setup_ID = " & m_Spector_Grid_Setup_ID & " "

                dt = db.DataTable(strSql)

                If dt.Rows.Count = 0 Then
                    drRow = dt.NewRow()
                Else
                    drRow = dt.Rows(0)
                End If
                'drRow = IIf(dt.Rows.Count = 0, dt.NewRow(), dt.Rows(0))

                Call SaveLine(drRow)

                If dt.Rows.Count = 0 Then
                    db.DBDataTable.Rows.Add(drRow)
                    Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                    db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
                Else
                    Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                    db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
                End If

                db.DBDataAdapter.Update(db.DBDataTable)

            Catch er As Exception
                MsgBox("Error in cSPT_SPECTOR_GRID_SETUP." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

#End Region

#Region "Public properties ################################################"

        Public Property Spector_Grid_Setup_ID() As Integer
            Get
                Spector_Grid_Setup_ID = m_Spector_Grid_Setup_ID
            End Get
            Set(ByVal value As Integer)
                m_Spector_Grid_Setup_ID = value
            End Set
        End Property
        Public Property Dept_ID() As Integer
            Get
                Dept_ID = m_Dept_ID
            End Get
            Set(ByVal value As Integer)
                m_Dept_ID = value
            End Set
        End Property
        Public Property Table_Name() As String
            Get
                Table_Name = m_Table_Name
            End Get
            Set(ByVal value As String)
                m_Table_Name = value
            End Set
        End Property
        Public Property Can_Search() As Boolean
            Get
                Can_Search = m_Can_Search
            End Get
            Set(ByVal value As Boolean)
                m_Can_Search = value
            End Set
        End Property
        Public Property Can_Read() As Boolean
            Get
                Can_Read = m_Can_Read
            End Get
            Set(ByVal value As Boolean)
                m_Can_Read = value
            End Set
        End Property
        Public Property Can_Write() As Boolean
            Get
                Can_Write = m_Can_Write
            End Get
            Set(ByVal value As Boolean)
                m_Can_Write = value
            End Set
        End Property
        Public Property Can_Sort() As Boolean
            Get
                Can_Sort = m_Can_Sort
            End Get
            Set(ByVal value As Boolean)
                m_Can_Sort = value
            End Set
        End Property
        Public Property Can_Delete() As Boolean
            Get
                Can_Delete = m_Can_Delete
            End Get
            Set(ByVal value As Boolean)
                m_Can_Delete = value
            End Set
        End Property
        Public Property Row_Head_Height() As Integer
            Get
                Row_Head_Height = m_Row_Head_Height
            End Get
            Set(ByVal value As Integer)
                m_Row_Head_Height = value
            End Set
        End Property
        Public Property Col_Head_Width() As Integer
            Get
                Col_Head_Width = m_Col_Head_Width
            End Get
            Set(ByVal value As Integer)
                m_Col_Head_Width = value
            End Set
        End Property
        Public Property Insert_Field() As String
            Get
                Insert_Field = m_Insert_Field
            End Get
            Set(ByVal value As String)
                m_Insert_Field = value
            End Set
        End Property
        Public Property User_Login() As String
            Get
                User_Login = m_User_Login
            End Get
            Set(ByVal value As String)
                m_User_Login = value
            End Set
        End Property
        Public Property Update_TS() As DateTime
            Get
                Update_TS = m_Update_TS
            End Get
            Set(ByVal value As DateTime)
                m_Update_TS = value
            End Set
        End Property

#End Region

    End Class

#Region "    CalendarColumn Class ##############################################"

    Public Class CalendarColumn
        Inherits DataGridViewColumn

        Public Sub New()
            MyBase.New(New CalendarCell())
        End Sub

        Public Overrides Property CellTemplate() As DataGridViewCell
            Get
                Return MyBase.CellTemplate
            End Get
            Set(ByVal value As DataGridViewCell)

                ' Ensure that the cell used for the template is a CalendarCell.
                If (value IsNot Nothing) AndAlso _
                    Not value.GetType().IsAssignableFrom(GetType(CalendarCell)) _
                    Then
                    Throw New InvalidCastException("Must be a CalendarCell")
                End If
                MyBase.CellTemplate = value

            End Set
        End Property

    End Class

    Public Class CalendarCell
        Inherits DataGridViewTextBoxCell

        Public Sub New()
            ' Use the short date format.
            Me.Style.Format = "d"

        End Sub

        Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, _
            ByVal initialFormattedValue As Object, _
            ByVal dataGridViewCellStyle As DataGridViewCellStyle)

            ' Set the value of the editing control to the current cell value.
            MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, _
                dataGridViewCellStyle)

            Dim ctl As CalendarEditingControl = _
                CType(DataGridView.EditingControl, CalendarEditingControl)

            ' Use the default row value when Value property is null.
            If (Me.Value Is Nothing) Or (Me.Value.Equals(DBNull.Value)) Then
                ctl.Value = CType(Me.DefaultNewRowValue, DateTime)
            Else
                ctl.Value = CType(Me.Value, DateTime)
            End If
        End Sub

        Public Overrides ReadOnly Property EditType() As Type
            Get
                ' Return the type of the editing control that CalendarCell uses.
                Return GetType(CalendarEditingControl)
            End Get
        End Property

        Public Overrides ReadOnly Property ValueType() As Type
            Get
                ' Return the type of the value that CalendarCell contains.
                Return GetType(DateTime)
            End Get
        End Property

        Public Overrides ReadOnly Property DefaultNewRowValue() As Object
            Get
                ' Use the current date and time as the default value.
                Return DateTime.Now
            End Get
        End Property

    End Class

    Class CalendarEditingControl
        Inherits DateTimePicker
        Implements IDataGridViewEditingControl

        Private dataGridViewControl As DataGridView
        Private valueIsChanged As Boolean = False
        Private rowIndexNum As Integer

        Public Sub New()
            Me.Format = DateTimePickerFormat.Short
        End Sub

        Public Property EditingControlFormattedValue() As Object _
            Implements IDataGridViewEditingControl.EditingControlFormattedValue

            Get
                Return Me.Value.ToShortDateString()
            End Get

            Set(ByVal value As Object)
                Try
                    ' This will throw an exception of the string is 
                    ' null, empty, or not in the format of a date.
                    Me.Value = DateTime.Parse(CStr(value))
                Catch
                    ' In the case of an exception, just use the default
                    ' value so we're not left with a null value.
                    Me.Value = DateTime.Now
                End Try
            End Set

        End Property

        Public Function GetEditingControlFormattedValue(ByVal context _
            As DataGridViewDataErrorContexts) As Object _
            Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

            Return Me.Value.ToShortDateString()

        End Function

        Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As  _
            DataGridViewCellStyle) _
            Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

            Me.Font = dataGridViewCellStyle.Font
            Me.CalendarForeColor = dataGridViewCellStyle.ForeColor
            Me.CalendarMonthBackground = dataGridViewCellStyle.BackColor

        End Sub

        Public Property EditingControlRowIndex() As Integer _
            Implements IDataGridViewEditingControl.EditingControlRowIndex

            Get
                Return rowIndexNum
            End Get
            Set(ByVal value As Integer)
                rowIndexNum = value
            End Set

        End Property

        Public Function EditingControlWantsInputKey(ByVal key As Keys, _
            ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
            Implements IDataGridViewEditingControl.EditingControlWantsInputKey

            ' Let the DateTimePicker handle the keys listed.
            Select Case key And Keys.KeyCode
                Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, _
                    Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp

                    Return True

                Case Else
                    Return Not dataGridViewWantsInputKey
            End Select

        End Function

        Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
            Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

            ' No preparation needs to be done.

        End Sub

        Public ReadOnly Property RepositionEditingControlOnValueChange() _
            As Boolean Implements _
            IDataGridViewEditingControl.RepositionEditingControlOnValueChange

            Get
                Return False
            End Get

        End Property

        Public Property EditingControlDataGridView() As DataGridView _
            Implements IDataGridViewEditingControl.EditingControlDataGridView

            Get
                Return dataGridViewControl
            End Get
            Set(ByVal value As DataGridView)
                dataGridViewControl = value
            End Set

        End Property

        Public Property EditingControlValueChanged() As Boolean _
            Implements IDataGridViewEditingControl.EditingControlValueChanged

            Get
                Return valueIsChanged
            End Get
            Set(ByVal value As Boolean)
                valueIsChanged = value
            End Set

        End Property

        Public ReadOnly Property EditingControlCursor() As Cursor _
            Implements IDataGridViewEditingControl.EditingPanelCursor

            Get
                Return MyBase.Cursor
            End Get

        End Property

        Protected Overrides Sub OnValueChanged(ByVal eventargs As EventArgs)

            ' Notify the DataGridView that the contents of the cell have changed.
            valueIsChanged = True
            Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
            MyBase.OnValueChanged(eventargs)

        End Sub

    End Class

#End Region

#Region "    DGV Columns functions ############################### "

    Public Function DGVButtonColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewButtonColumn

        DGVButtonColumn = New DataGridViewButtonColumn

        DGVButtonColumn.HeaderText = pstrHeaderText
        DGVButtonColumn.DataPropertyName = pstrName
        DGVButtonColumn.Name = pstrName

        If plWidth <> 0 Then DGVButtonColumn.Width = plWidth

    End Function

    Public Function DGVCalendarColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As CalendarColumn

        DGVCalendarColumn = New CalendarColumn

        DGVCalendarColumn.HeaderText = pstrHeaderText
        DGVCalendarColumn.DataPropertyName = pstrName
        DGVCalendarColumn.Name = pstrName
        'DGVCalendarColumn.DefaultCellStyle = "mm/dd/yyyy"
        DGVCalendarColumn.DefaultCellStyle.Format = "mm/dd/yyyy"

        If plWidth <> 0 Then DGVCalendarColumn.Width = plWidth

    End Function

    Public Function DGVCheckBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewCheckBoxColumn

        DGVCheckBoxColumn = New DataGridViewCheckBoxColumn

        DGVCheckBoxColumn.HeaderText = pstrHeaderText
        DGVCheckBoxColumn.DataPropertyName = pstrName
        DGVCheckBoxColumn.Name = pstrName

        DGVCheckBoxColumn.FalseValue = 0 '"False"
        DGVCheckBoxColumn.TrueValue = 1 '"True"
        DGVCheckBoxColumn.IndeterminateValue = 0 '"False"

        If plWidth <> 0 Then DGVCheckBoxColumn.Width = plWidth

    End Function

    Public Function DGVComboBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewComboBoxColumn

        DGVComboBoxColumn = New DataGridViewComboBoxColumn

        DGVComboBoxColumn.HeaderText = pstrHeaderText
        DGVComboBoxColumn.DataPropertyName = pstrName
        DGVComboBoxColumn.Name = pstrName
        'DGVComboBoxColumn.DropDownWidth = 160

        If plWidth <> 0 Then DGVComboBoxColumn.Width = plWidth

    End Function

    Public Function DGVImageColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, ByVal plWidth As Long) As DataGridViewImageColumn

        DGVImageColumn = New DataGridViewImageColumn
        DGVImageColumn.Name = pstrName
        DGVImageColumn.DataPropertyName = pstrName
        DGVImageColumn.HeaderText = pstrHeaderText

    End Function

    Public Function DGVTextBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewTextBoxColumn

        DGVTextBoxColumn = New DataGridViewTextBoxColumn

        DGVTextBoxColumn.HeaderText = pstrHeaderText
        DGVTextBoxColumn.DataPropertyName = pstrName
        DGVTextBoxColumn.Name = pstrName

        If plWidth <> 0 Then DGVTextBoxColumn.Width = plWidth

    End Function

    Public Function GetDDTestsDataTable(ByVal pstrTable As String, ByVal pstrField As String) As DataTable

        Dim dt As DataTable
        Dim db As New cDBA
        Dim strsql As String = _
        "SELECT     DatabaseChar, DisplayChar, ISNULL(Description, '') AS Description " & _
        "FROM       DDTests WITH (nolock) " & _
        "WHERE      TableName = 'cicmpy' AND FieldName = 'cmp_type' " & _
        "ORDER BY   TableName, FieldName, DatabaseChar "

        dt = db.DataTable(strsql)

        Return dt

    End Function

    'Public Function GetElementDescription(ByVal pstrTable As String, ByVal pstrField As String, ByVal pstrValue As String, Optional ByVal pstrValue2 As String = "") As String

    '    GetElementDescription = ""

    '    Dim dt As DataTable
    '    Dim db As New cDBA
    '    Dim strsql As String = ""

    '    Select Case pstrTable.ToUpper

    '        Case "ADDRESSSTATES" ' States and provinces
    '            strsql = _
    '            "SELECT StateCode, ISNULL(Name, '') AS Description, CountryCode " & _
    '            "FROM   AddressStates WITH (Nolock) " & _
    '            "WHERE  CountryCode='" & pstrValue & "' AND StateCode = '" & pstrValue2 & "' "

    '        Case "ARSLMFIL_SQL"
    '            strsql = "SELECT * FROM ARSLMFIL_SQL WITH (Nolock) WHERE HumRes_ID = '" & pstrValue & "'"

    '        Case "CICMPY"

    '            Select Case pstrField.ToUpper

    '                Case "CMP_CODE"
    '                    strsql = _
    '                    "SELECT Cmp_Code, ISNULL(cmp_name, '') AS Description " & _
    '                    "FROM   Cicmpy WITH (Nolock) " & _
    '                    "WHERE 	Cmp_Code = '" & pstrValue & "' "

    '                Case "CMP_PARENT"
    '                    strsql = _
    '                    "SELECT Cmp_Code AS Description " & _
    '                    "FROM   Cicmpy WITH (Nolock) " & _
    '                    "WHERE  CMP_WWN = '" & pstrValue & "' "

    '                Case ""
    '                    'Case ""
    '                    'Case ""

    '            End Select

    '        Case "CLASSIFICATIONS"
    '            strsql = _
    '            "SELECT     C.ClassificationID, ISNULL(C.Description, '') AS Description " & _
    '            "FROM       Classifications C WITH (Nolock) " & _
    '            "LEFT JOIN  DDTests D WITH (Nolock) ON D.TableName = 'cicmpy' AND D.FieldName = 'cmp_type' AND C.Type = D.DatabaseChar " & _
    '            "WHERE      C.Type IS NULL AND C.ClassificationID = '" & pstrValue & "' "

    '        Case "LAND" ' Country
    '            strsql = _
    '            "SELECT LandCode, ISNULL(OMS60_0, '') AS Description " & _
    '            "FROM   Land WITH (Nolock) " & _
    '            "WHERE  Active = 1 AND LandCode = '" & pstrValue & "' "

    '        Case "SYCDEFIL_SQL"
    '            strsql = _
    '            "SELECT Sy_Code, ISNULL(Code_Desc, '') AS Description " & _
    '            "FROM   SYCDEFIL_SQL WITH (Nolock) " & _
    '            "WHERE  Sy_Code = '" & pstrValue & "'"

    '        Case "SYTRMFIL_SQL"
    '            strsql = _
    '            "SELECT Term_Code, ISNULL(Description, '') AS Description " & _
    '            "FROM   SYTRMFIL_SQL WITH (Nolock) " & _
    '            "WHERE  Term_Code = '" & pstrValue & "'"

    '        Case "TAAL"
    '            strsql = _
    '            "SELECT TaalCode, ISNULL(OMS30_0, '') AS Description " & _
    '            "FROM   Taal WITH (Nolock) " & _
    '            "WHERE  TaalCode = '" & pstrValue & "' "

    '        Case "TAXDETL_SQL"
    '            strsql = _
    '            "SELECT Tax_Cd, ISNULL(Tax_Cd_Description, '') AS Description " & _
    '            "FROM   TAXDETL_SQL WITH (Nolock) " & _
    '            "WHERE  Tax_Cd = '" & pstrValue & "' "

    '        Case "TAXSCHED_SQL"
    '            strsql = _
    '            "SELECT Tax_Sched, ISNULL(Tax_Sched_Desc, '') AS Description " & _
    '            "FROM   TAXSCHED_SQL WITH (Nolock) " & _
    '            "WHERE  Tax_Sched = '" & pstrValue & "'"

    '    End Select

    '    If strsql = "" Then

    '        strsql = _
    '        "SELECT     DatabaseChar, DisplayChar, ISNULL(Description, '') AS Description " & _
    '        "FROM       DDTests WITH (nolock) " & _
    '        "WHERE      TableName = '" & pstrTable & "' AND FieldName = '" & pstrField & "' AND DatabaseChar = '" & pstrValue & "' " & _
    '        "ORDER BY   TableName, FieldName, DatabaseChar "

    '    End If

    '    dt = db.DataTable(strsql)

    '    If dt.Rows.Count <> 0 Then
    '        GetElementDescription = dt.Rows(0).Item("Description").ToString
    '    End If

    '    'Dim strSql As String = "SELECT * FROM SYCDEFIL_SQL WITH (Nolock) WHERE sy_code = '" & pstrShip_Via_Cd & "'"
    '    'Dim strSql As String = "SELECT * FROM SYTRMFIL_SQL WITH (Nolock) WHERE Term_Code = '" & pstrAr_Terms_Cd & "'"
    '    'Dim strSql As String = "SELECT * FROM TAXSCHED_SQL WITH (Nolock) WHERE Tax_Sched = '" & pstrTax_Sched & "'"
    '    'Dim strSql As String = _
    '    '"SELECT ISNULL(Tax_Cd_Description, '') AS Tax_Cd_Description " & _
    '    '"FROM TAXDETL_SQL WITH (Nolock) WHERE Tax_Cd"
    '    'Dim strSql As String = "SELECT * FROM ARSLMFIL_SQL WITH (Nolock) WHERE HumRes_ID = '" & piSlspsn_No & "'"

    'End Function

#End Region

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

    'Public Overrides sub 
    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function

    Public Sub Load(ByVal pUser_Login As String, ByVal pUser_Dept As String, ByVal pTable_Source As String)

        Dim db As New cDBA("DSN=EXACT")
        Dim dt As DataTable
        Dim strSql As String
        Dim oSetup As cSPT_SPECTOR_GRID_SETUP

        strSql = _
        "SELECT * " & _
        "FROM   SPT_SPECTOR_GRID_SETUP WITH (NOLOCK) " & _
        "WHERE  USER_LOGIN = '" & pUser_Login & "' AND " & _
        "       TABLE_NAME = '" & pTable_Source & "' "

        dt = db.DataTable(strSql)
        If dt.Rows.Count <> 0 Then

            oSetup = New cSPT_SPECTOR_GRID_SETUP(dt.Rows(0).Item("SPECTOR_GRID_SETUP_ID"))

            strSql = _
            "SELECT     * " & _
            "FROM       SPT_SPECTOR_GRID_FIELDS WITH (NOLOCK) " & _
            "WHERE      SPECTOR_GRID_SETUP_ID = " & oSetup.Spector_Grid_Setup_ID.ToString & " " & _
            "ORDER BY   COL_POS, SPECTOR_GRID_FIELD_ID "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

            End If

        End If

    End Sub


    'Public Sub LoadHistoryGrid()

    '    Try

    '        With dgvHistory.Columns
    '            .Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
    '            .Add(DGVTextBoxColumn("Ord_Type", "Ord Type", 40))
    '            .Add(DGVTextBoxColumn("OEI_Ord_No", "OEI No", 80))
    '            .Add(DGVTextBoxColumn("Ord_No", "Macola No", 70))
    '            .Add(DGVCalendarColumn("Ord_Dt", "Date", 70))
    '            .Add(DGVTextBoxColumn("Cus_No", "Customer", 80))
    '            .Add(DGVTextBoxColumn("Bill_To_Name", "Bill To", 250))
    '            .Add(DGVTextBoxColumn("OE_PO_No", "P/O", 100))
    '            .Add(DGVCheckBoxColumn("Pending", "Pending", 60))
    '            .Add(DGVCheckBoxColumn("FileCreated", "File Created", 60))
    '            .Add(DGVCheckBoxColumn("FileExported", "File Exported", 60))
    '            .Add(DGVCheckBoxColumn("TravelerCreated", "Traveler Created", 60))
    '            .Add(DGVTextBoxColumn("LastCol", "LastCol", 0))

    '        End With

    '        dgvHistory.Columns(History.FirstCol).Visible = False
    '        dgvHistory.Columns(History.LastCol).Visible = False

    '        Call LoadHistoryGridValues()

    '        Dim oCellStyle As New DataGridViewCellStyle()

    '        oCellStyle = New DataGridViewCellStyle()
    '        oCellStyle.Format = "mm/dd/yyyy"
    '        oCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

    '        dgvHistory.Columns(History.Ord_Dt).DefaultCellStyle = oCellStyle

    '        dgvHistory.Columns(History.FirstCol).Frozen = True
    '        dgvHistory.Columns(History.OEI_Ord_No).Frozen = True
    '        dgvHistory.Columns(History.Ord_No).Frozen = True

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

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
