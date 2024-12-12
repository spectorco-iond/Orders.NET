Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports System.IO

'Imports System.Drawing

Public Class frmNotes
    Inherits System.Windows.Forms.Form

    Private m_Note_Name As String = ""
    Private m_Note_Name_Value As String = ""

    Private m_MinHeight As Integer
    Private m_MinWidth As Integer

    Private m_strSqlSelect As String

#Region "Private buttons events ###############################################"

    Private Sub cmdNew_Click(ByVal oSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNew.Click

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            Dim oNote As New cNote
            oNote.Note_Name = m_Note_Name
            oNote.Note_Name_Value = m_Note_Name_Value

            Dim ofrmNote As New frmNote
            ofrmNote.Note = oNote

            ofrmNote.ShowDialog()

            Call GetData()

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdDelete_Click(ByVal oSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click

        Try
            If dgvNotes.Rows.Count = 0 Then Exit Sub

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            Dim oNote As New cNote(dgvNotes.CurrentRow.Cells(Columns.ID).Value)

            Dim oResult As MsgBoxResult
            oResult = MsgBox("Do you want to delete this note ? ", MsgBoxStyle.YesNoCancel, "Delete Note")

            If oResult = MsgBoxResult.Yes Then

                oNote.Delete()
                dgvNotes.Rows.RemoveAt(dgvNotes.CurrentRow.Index)

            End If
        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "frmNotes"

    Private Sub frmNotes_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'm_MinHeight = 5320
        'm_MinWidth = 6800

        Call GetData()
        Call HeaderTitle()

    End Sub

    Private Sub frmNotes_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize


        '' ''< 716 370 680 270 16 293
        ''dgvNotes.Width = Me.Width - 36
        ''dgvNotes.Height = Me.Height - 100

        ''panButtons.Location = New Point(Me.Width - panButtons.Width - 36, Me.Height - 77)

        'panButtons.Location.X = Me.Width - panButtons.Width - 36
        'panButtons.Location.Y = Me.Height - 77

    End Sub

#End Region

#Region "Private subs and functions #######################################"

    'Private Function Connection() As ADODB.Connection
    '    ''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '    'On Error GoTo ErrHandle
    '    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '    Dim cnConnection As String
    '    cnConnection = "DRIVER={SQL Server};SERVER=" & gServerName & ";DATABASE=" & gDatabaseName & ";Trusted_Connection=YES"
    '    'cnConnection = "DRIVER={SQL Server};SERVER=thor;DATABASE=100;Trusted_Connection=YES"

    '    Dim Conn As New ADODB.Connection
    '    Conn.ConnectionString = cnConnection

    '    '////////////////////////
    '    'Test connection
    '    Conn.Open()
    '    Conn.Close()
    '    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '    Connection = Conn

    'End Function

    Private Sub GetNotes(Optional ByVal pstrButton As String = "", Optional ByVal pstrValue As String = "")

        'Dim rsNotes As New ADODB.Recordset

        Try
            m_strSqlSelect = SqlLine()

            Dim dt As New DataTable
            Dim db As New cDBA

            dt = db.DataTable(m_strSqlSelect)

            ''Dim oConnection As New Odbc.OdbcConnection("Provider=SQLodbc;" & gConn.ConnectionString)
            'Dim oConnection As New Odbc.OdbcConnection(gConn.ConnectionString)
            'Dim oCommand As New Odbc.OdbcCommand(m_strSqlSelect, oConnection)
            'Dim oDataAdapter As New Odbc.OdbcDataAdapter(oCommand)
            'Dim oDataset As New DataSet
            'oDataAdapter.Fill(oDataset, "Notes")

            'rsNotes.Open(m_strSqlSelect, Connection.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'dgvNotes.DataSource = oDataset.Tables("Notes")
            dgvNotes.DataSource = dt
            dgvNotes.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            'dgvNotes.Refresh() ' .CtlRefresh()
            'Call SetNotesGridHeaders()
            'rsNotes.Close()

            'dgvNotes.Columns(0).Width = 25
            dgvNotes.Columns(0).Width = 190
            dgvNotes.Columns(1).Width = 250
            dgvNotes.Columns(2).Width = 110
            dgvNotes.Columns(3).Width = 80

            dgvNotes.Columns(4).Width = 0
            dgvNotes.Columns(4).Visible = False

            'For Each dgvCol As DataGridViewColumn In dgvNotes.Columns
            '    dgvCol.Width = 500
            'Next

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub HeaderTitle()

        Dim strTitle As String = ""

        Select Case m_Note_Name

            Case "OE-ORDER-NO"
                strTitle = "Order"
            Case "AR-CUSTOMER"
                strTitle = "Customer"
            Case "AR-SALESPERSON"
                strTitle = "Salesperson"
            Case "SY-JOB-CODE"
                strTitle = "Project"
            Case "IM-LOCATION"
                strTitle = "Location"
            Case "AR-TERMS-CODE"
                strTitle = "Terms"
            Case "SY-ACCOUNT2"
                strTitle = "Profit center"
            Case "SY-ACCOUNT3"
                strTitle = "Cost center"
            Case "OE-COMMENT-CD-2"
                strTitle = ""
            Case "AR-ALT-ADDR-COD"
                strTitle = "Ship-To"
            Case "AR-SHIP-VIA"
                strTitle = "Ship via"

        End Select

        Me.Text = strTitle & " notes " & m_Note_Name_Value

    End Sub

    Private Sub SetNotesGridHeaders()

        ' Fixed and defined columns
        dgvNotes.Columns(0).Width = 240
        'dgvNotes.set_ColWidth(1, 0)

        'dgvNotes.Columns(1).Width = "Topic")
        'fgNotes.set_TextMatrix(0, 2, "Attachment")
        'fgNotes.set_TextMatrix(0, 3, "Entered by")
        'fgNotes.set_TextMatrix(0, 4, "Date")

        dgvNotes.Columns(1).Width = 240
        dgvNotes.Columns(1).Width = 240
        dgvNotes.Columns(1).Width = 240
        dgvNotes.Columns(1).Width = 240

    End Sub

    Private Function SqlLine() As String

        SqlLine = _
        "SELECT S.Note_Topic AS 'Topic', " & _
        "       '' AS 'Attachment', " & _
        "       S.User_Name AS 'Entered by', " & _
        "       CONVERT(CHAR(10), S.Note_Dt, 103) AS 'Date', S.ID " & _
        "FROM   SYSNOTES_SQL S WITH (Nolock) " & _
        "WHERE  Note_Name = '" & SqlCompliantString(m_Note_Name) & "' AND " & _
        "       Note_Name_Value = '" & SqlCompliantString(m_Note_Name_Value) & "'"

    End Function

#End Region

#Region "Public subs and functions ########################################"

    Public Sub GetData(Optional ByVal pblnAll As Boolean = False)

        Call GetNotes()

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property NotesElement() As String
        Get
            NotesElement = m_Note_Name
        End Get
        Set(ByVal value As String)
            m_Note_Name = value
        End Set
    End Property

    Public Property NotesValue() As String
        Get
            NotesValue = m_Note_Name_Value
        End Get
        Set(ByVal value As String)
            m_Note_Name_Value = value
        End Set
    End Property

#End Region

    Private Sub dgvNotes_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvNotes.DoubleClick

        If dgvNotes.Rows.Count = 0 Then Exit Sub

        Dim oNote As New cNote(dgvNotes.CurrentRow.Cells(Columns.ID).Value)

        Dim ofrmNote As New frmNote
        ofrmNote.Note = oNote

        ofrmNote.ShowDialog()

        Call GetData()

    End Sub

    'S.Note_Topic AS 'Topic', " & _
    '    "       ISNULL(S.Filler_001, '') AS 'Attachment', " & _
    '    "       S.User_Name AS 'Entered by', " & _
    '    "       CONVERT(CHAR(10), S.Note_Dt, 103) AS 'Date', S.ID " 
    Private Enum Columns

        Topic = 0
        Attachment
        User_Name
        Note_Dt
        ID
        LastCol

    End Enum

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        dgvNotes.DataSource = Nothing

        Me.Close()

    End Sub

End Class