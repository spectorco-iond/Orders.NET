Imports System.Data.Odbc
Imports System.Collections.Generic

Public Class cTraveler_Reason_LateShipment
    Implements ICloneable

#Region " -------------------------------- Private properties ---------------------------------"

#Region "###############Private Variable###############"

    Private m_intID As Int32
    Private m_strOrdNo As String
    Private m_intHeaderId As Int32
    Private m_strCus_No As String
    Private m_intReason_TypeID As Int32
    Private m_strReasonComments As String

    Private m_strExtra1 As String
    Private m_strExtra2 As String
    Private m_strExtra3 As String

    Private m_intExtra4 As Int32
    Private m_intExtra5 As Int32

    Private m_blExtra6 As Int16
    Private m_blExtra7 As Int16

    Private m_strUSER_LOGIN As String
    Private m_dtCREATE_TS As DateTime
    Private m_dtUPDATE_TS As DateTime

    Private m_stroeiguid As String

#End Region

#End Region


    Sub New()
        Call Init()
    End Sub


    Private Sub Init()

        m_intID = 0
        m_strOrdNo = String.Empty
            m_intHeaderId = 0
            m_strCus_No = String.Empty
            m_intReason_TypeID = 0
            m_strReasonComments = String.Empty

            m_strExtra1 = String.Empty
            m_strExtra2 = String.Empty
            m_strExtra3 = String.Empty

            m_intExtra4 = 0
            m_intExtra5 = 0

            m_blExtra6 = 0
            m_blExtra7 = 0

            m_strUSER_LOGIN = String.Empty
            m_dtCREATE_TS = NoDate()
        m_dtUPDATE_TS = NoDate()

        m_stroeiguid = String.Empty

    End Sub

    Private Sub LoadLine(ByVal pClass As cTraveler_Reason_LateShipment, ByRef pdrRow As DataRow)

        Try

            If Not (pdrRow.Item("ID").Equals(DBNull.Value)) Then pClass.ID = pdrRow.Item("ID")
            If Not (pdrRow.Item("ORD_NO").Equals(DBNull.Value)) Then pClass.ORD_NO = pdrRow.Item("ORD_NO").ToString
            If Not (pdrRow.Item("HEADERID").Equals(DBNull.Value)) Then pClass.HEADERID = pdrRow.Item("HEADERID")
            If Not (pdrRow.Item("CUS_NO").Equals(DBNull.Value)) Then pClass.CUS_NO = pdrRow.Item("CUS_NO").ToString

            If Not (pdrRow.Item("REASON_TYPE_ID").Equals(DBNull.Value)) Then pClass.REASON_TYPE_ID = pdrRow.Item("REASON_TYPE_ID")
            If Not (pdrRow.Item("REASON_COMMENTS").Equals(DBNull.Value)) Then pClass.REASON_COMMENTS = pdrRow.Item("REASON_COMMENTS").ToString

            '------------------------------------------------------
            If Not (pdrRow.Item("EXTRA1").Equals(DBNull.Value)) Then pClass.EXTRA1 = pdrRow.Item("EXTRA1").ToString
            If Not (pdrRow.Item("EXTRA2").Equals(DBNull.Value)) Then pClass.EXTRA2 = pdrRow.Item("EXTRA2").ToString
            If Not (pdrRow.Item("EXTRA3").Equals(DBNull.Value)) Then pClass.EXTRA3 = pdrRow.Item("EXTRA3").ToString

            If Not (pdrRow.Item("EXTRA4").Equals(DBNull.Value)) Then pClass.EXTRA4 = pdrRow.Item("EXTRA4")
            If Not (pdrRow.Item("EXTRA5").Equals(DBNull.Value)) Then pClass.EXTRA5 = pdrRow.Item("EXTRA5")

            If Not (pdrRow.Item("EXTRA6").Equals(DBNull.Value)) Then pClass.EXTRA6 = pdrRow.Item("EXTRA6")
            If Not (pdrRow.Item("EXTRA7").Equals(DBNull.Value)) Then pClass.EXTRA7 = pdrRow.Item("EXTRA7")
            '------------------------------------------------------

            If Not (pdrRow.Item("USER_LOGIN").Equals(DBNull.Value)) Then m_strUSER_LOGIN = pdrRow.Item("USER_LOGIN").ToString
            If Not (pdrRow.Item("CREATE_TS").Equals(DBNull.Value)) Then m_dtCREATE_TS = pdrRow.Item("CREATE_TS").ToString
            If Not (pdrRow.Item("UPDATE_TS").Equals(DBNull.Value)) Then m_dtUPDATE_TS = pdrRow.Item("UPDATE_TS").ToString

            If Not (pdrRow.Item("OEI_GUID").Equals(DBNull.Value)) Then pClass.OEIGUID = pdrRow.Item("OEI_GUID").ToString

        Catch ex As Exception
            MsgBox("Error in cTraveler_Reason_LateShipment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub SaveLine(ByRef pdrRow As DataRow)
        Try


            ' pdrRow.Item("ID") = ID
            pdrRow.Item("ORD_NO") = ORD_NO
            pdrRow.Item("HEADERID") = HEADERID
            pdrRow.Item("CUS_NO") = CUS_NO
            pdrRow.Item("REASON_TYPE_ID") = REASON_TYPE_ID
            pdrRow.Item("REASON_COMMENTS") = REASON_COMMENTS

            '------------------------------------------------------
            pdrRow.Item("EXTRA1") = EXTRA1
            pdrRow.Item("EXTRA2") = EXTRA2
            pdrRow.Item("EXTRA3") = EXTRA3

            pdrRow.Item("EXTRA4") = EXTRA4
            pdrRow.Item("EXTRA5") = EXTRA5

            pdrRow.Item("EXTRA6") = EXTRA6
            pdrRow.Item("EXTRA7") = EXTRA7
            '------------------------------------------------------

            pdrRow.Item("USER_LOGIN") = Environment.UserName
            pdrRow.Item("UPDATE_TS") = Date.Now.ToString("MM/dd/yyyy HHH:mm:ss")
            pdrRow.Item("OEI_GUID") = OEIGUID

        Catch ex As Exception
            MsgBox("Error in cTraveler_Reason_LateShipment. " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Sub

    Public Function Load(ByVal _id As Int32) As cTraveler_Reason_LateShipment

        Load = New cTraveler_Reason_LateShipment()

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "SELECT * " &
            "FROM   Traveler_Reason_LateShipment WITH (Nolock) where ID = " & _id


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(Load, dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cTraveler_Reason_LateShipment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Public Function Load(ByVal _oeiguid As String) As cTraveler_Reason_LateShipment

        Load = New cTraveler_Reason_LateShipment()

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "SELECT * " &
            "FROM   Traveler_Reason_LateShipment WITH (Nolock) where oei_guid = '" & _oeiguid & "'"


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(Load, dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cTraveler_Reason_LateShipment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Public Sub Save()
        Try
            Dim drRow As DataRow
            Dim dt As New DataTable
            Dim db As cDBA
            db = New cDBA


            Dim strSql As String
            strSql =
            "SELECT * " &
            "FROM   Traveler_Reason_LateShipment WITH (Nolock) where OEI_GUID = '" & OEIGUID & "'"


            dt = db.DataTable(strSql)

            If dt.Rows.Count = 0 Then
                drRow = dt.NewRow()
            Else
                drRow = dt.Rows(0)
            End If

            Call SaveLine(drRow)

            If dt.Rows.Count = 0 Then
                db.DBDataTable.Rows.Add(drRow)
                Dim cmd = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            Else
                Dim cmd = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            End If

            db.DBDataAdapter.Update(db.DBDataTable)

        Catch ex As Exception
            MsgBox("Error in cTraveler_Reason_LateShipment. " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Sub

    Public Sub Delete()

        Dim db As New cDBA

        Try

            Dim strSql As String
            strSql =
            "DELETE from   Traveler_Reason_LateShipment where  OEI_GUID = '" & OEIGUID & "'"

            db.Execute(strSql)


        Catch er As Exception
            MsgBox("Error in cTraveler_Reason_LateShipment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


#Region "--------------------------- Public properties ---------------------- "

    Public Property ID As Int32
        Get
            Return m_intID
        End Get
        Set(ByVal value As Int32)
            m_intID = value
        End Set
    End Property
    Public Property ORD_NO As String
        Get
            Return m_strOrdNo
        End Get
        Set(ByVal value As String)
            m_strOrdNo = value
        End Set
    End Property
    Public Property HEADERID As Int32
        Get
            Return m_intHeaderId
        End Get
        Set(ByVal value As Int32)
            m_intHeaderId = value
        End Set
    End Property
    Public Property CUS_NO As String
        Get
            Return m_strCus_No
        End Get
        Set(ByVal value As String)
            m_strCus_No = value
        End Set
    End Property
    Public Property REASON_TYPE_ID As Int32
        Get
            Return m_intReason_TypeID
        End Get
        Set(ByVal value As Int32)
            m_intReason_TypeID = value
        End Set
    End Property
    Public Property REASON_COMMENTS As String
        Get
            Return m_strReasonComments
        End Get
        Set(ByVal value As String)
            m_strReasonComments = value
        End Set
    End Property
    Public Property EXTRA1 As String
        Get
            Return m_strExtra1
        End Get
        Set(ByVal value As String)
            m_strExtra1 = value
        End Set
    End Property
    Public Property EXTRA2 As String
        Get
            Return m_strExtra2
        End Get
        Set(ByVal value As String)
            m_strExtra2 = value
        End Set
    End Property
    Public Property EXTRA3 As String
        Get
            Return m_strExtra3
        End Get
        Set(ByVal value As String)
            m_strExtra3 = value
        End Set
    End Property
    Public Property EXTRA4 As Int32
        Get
            Return m_intExtra4
        End Get
        Set(ByVal value As Int32)
            m_intExtra4 = value
        End Set
    End Property
    Public Property EXTRA5 As Int32
        Get
            Return m_intExtra5
        End Get
        Set(ByVal value As Int32)
            m_intExtra5 = value
        End Set
    End Property
    Public Property EXTRA6 As Int16
        Get
            Return m_blExtra6
        End Get
        Set(ByVal value As Int16)
            m_blExtra6 = value
        End Set
    End Property
    Public Property EXTRA7 As Int16
        Get
            Return m_blExtra7
        End Get
        Set(ByVal value As Int16)
            m_blExtra7 = value
        End Set
    End Property

    Public Property USER_LOGIN As String
        Get
            Return m_strUSER_LOGIN
        End Get
        Set(ByVal value As String)
            m_strUSER_LOGIN = value
        End Set
    End Property
    Public Property CREATE_TS As Date
        Get
            Return m_dtCREATE_TS
        End Get
        Set(ByVal value As DateTime)
            m_dtCREATE_TS = value
        End Set
    End Property

    Public Property UPDATE_TS As Date
        Get
            Return m_dtUPDATE_TS
        End Get
        Set(ByVal value As DateTime)
            m_dtUPDATE_TS = value
        End Set
    End Property
    Public Property OEIGUID As String
        Get
            Return m_stroeiguid
        End Get
        Set(ByVal value As String)
            m_stroeiguid = value
        End Set
    End Property

#End Region

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone
    End Function
End Class
