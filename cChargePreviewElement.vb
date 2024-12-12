Option Strict Off
Option Explicit On

Imports System.IO

Public Class cChargePreviewElement

    Private m_ID As Integer = 0
    Private m_Ord_Guid As String
    Private m_Item_No As String
    Private m_Unit_Price As Double
    Private m_Qty_Ordered As Double
    Private m_Description As String
    'Private m_Line_Seq_No As Integer

#Region "Public constructors ##############################################"

    'Public Sub New()
    Public Sub New()

        Call Init()

    End Sub

    'Public Sub New(ByVal pID As Integer)
    Public Sub New(ByVal pID As Integer)

        Try

            Call Init()

            m_ID = pID

            Call Load(pID)

        Catch er As Exception
            MsgBox("Error in cChargePreviewElement." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_ID = 0
            m_Ord_Guid = String.Empty

            Call Reset()

        Catch er As Exception
            MsgBox("Error in cChargePreviewElement." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Load is made private, it is loaded automatically on New.
    Private Sub Load(ByVal pintID As Integer)

        Try
            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   OEI_Charges_Preview C WITH (Nolock) " & _
            "WHERE  ID = " & pintID & " "

            dt = db.DataTable(strSql)

            Call LoadLine(dt.Rows(0))

        Catch er As Exception
            MsgBox("Error in cChargePreviewElement." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Function Load(ByVal pstrOrd_Guid As String) As DataTable

    '    Load = New DataTable

    '    Try
    '        Dim db As New cDBA
    '        Dim dt As DataTable

    '        Dim strSql As String

    '        strSql = _
    '        "SELECT     * " & _
    '        "FROM       OEI_OrdHdr H WITH (Nolock) " & _
    '        "WHERE      Ord_Guid = '" & pstrOrd_Guid & "' "

    '        dt = db.DataTable(strSql)
    '        m_OEI_Ord_No = dt.Rows(0).Item("Oei_Ord_No")

    '        strSql = _
    '        "SELECT     S.ID, S.Ord_Guid, S.Cmt, S.Line_Seq_No, H.OEI_Ord_No, S.Ord_No " & _
    '        "FROM       OEI_LinCmt S WITH (Nolock) " & _
    '        "INNER JOIN OEI_OrdHdr H WITH (Nolock) ON H.Ord_Guid = S.Ord_Guid " & _
    '        "WHERE      H.Ord_Guid = '" & pstrOrd_Guid & "' " & _
    '        "ORDER BY   Line_Seq_No "

    '        Load = db.DataTable(strSql)

    '    Catch er As Exception
    '        MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Function

    ' Load is made private, it is loaded automatically on New.
    ' Loads the Comment fields from the DataRow record
    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            If Not (pdrRow.Item("ID").Equals(DBNull.Value)) Then m_ID = pdrRow.Item("ID")
            If Not (pdrRow.Item("Ord_Guid").Equals(DBNull.Value)) Then m_Ord_Guid = pdrRow.Item("Ord_Guid").ToString
            If Not (pdrRow.Item("Item_No").Equals(DBNull.Value)) Then m_Item_No = pdrRow.Item("Item_No").ToString
            If Not (pdrRow.Item("Unit_Price").Equals(DBNull.Value)) Then m_Unit_Price = pdrRow.Item("Unit_Price")
            If Not (pdrRow.Item("Qty_Ordered").Equals(DBNull.Value)) Then m_Qty_Ordered = pdrRow.Item("Qty_Ordered")
            If Not (pdrRow.Item("Description").Equals(DBNull.Value)) Then m_Description = pdrRow.Item("Description").ToString

        Catch er As Exception
            MsgBox("Error in cChargePreviewElement." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Saves the Comment fields into the DataRow record
    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("Ord_Guid") = m_Ord_Guid
            pdrRow.Item("Item_No") = m_Item_No
            pdrRow.Item("Unit_Price") = m_Unit_Price
            pdrRow.Item("Qty_Ordered") = m_Qty_Ordered
            pdrRow.Item("Description") = m_Description

        Catch er As Exception
            MsgBox("Error in cChargePreviewElement." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    ' Deletes the current Comment from the database, if it exists.
    Public Sub Delete()

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If m_ID = 0 Then Exit Sub

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            strSql = _
            "DELETE FROM OEI_CHARGES_PREVIEW " & _
            "WHERE  ID = " & m_ID & " AND " & _
            "       Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in cChargePreviewElement." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every non mandatory field of a Comment
    Public Sub Reset()

        ' Resets the locals to empty values, all but ID, Ord_No and Ord_Guid
        m_Item_No = String.Empty
        m_Unit_Price = 0
        m_Qty_Ordered = 0
        m_Description = String.Empty

    End Sub

    ' Update the current Comment into the database, or creates it if not existing
    Public Sub Save()

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String

            strSql = _
            "SELECT * " & _
            "FROM   OEI_CHARGES_PREVIEW " & _
            "WHERE  ID = " & m_ID & " AND " & _
            "       Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' "

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
            MsgBox("Error in cChargePreviewElement." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property Description() As String
        Get
            Description = m_Description
        End Get
        Set(ByVal value As String)
            m_Description = value
        End Set
    End Property
    Public Property ID() As Integer
        Get
            ID = m_ID
        End Get
        Set(ByVal value As Integer)
            m_ID = value
        End Set
    End Property
    Public Property Item_No() As String
        Get
            Item_No = m_Item_No
        End Get
        Set(ByVal value As String)
            m_Item_No = value
        End Set
    End Property
    Public Property Ord_Guid() As String
        Get
            Ord_Guid = m_Ord_Guid
        End Get
        Set(ByVal value As String)
            m_Ord_Guid = value
        End Set
    End Property
    Public Property Qty_Ordered() As Double
        Get
            Qty_Ordered = m_Qty_Ordered
        End Get
        Set(ByVal value As Double)
            m_Qty_Ordered = value
        End Set
    End Property
    Public Property Unit_Price() As Double
        Get
            Unit_Price = m_Unit_Price
        End Get
        Set(ByVal value As Double)
            m_Unit_Price = value
        End Set
    End Property

#End Region

End Class
