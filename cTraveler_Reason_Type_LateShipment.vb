Imports System.Collections.Generic
Imports System.Data.Odbc

Public Class cTraveler_Reason_Type_LateShipment
    Implements ICloneable

#Region " -------------------------------- Private properties ---------------------------------"

    Private m_intID As Int32
    Private m_strReasonType As String
    Private m_strExtra1 As String
    Private m_strExtra2 As String
    Private m_strExtra3 As String
    Private m_strUsername As String
    Private m_dtCREATE As DateTime
    Private m_dtupdate As DateTime
    Private m_strapptype As String

#End Region


    Sub New()
        Call Init()
    End Sub


    Private Sub Init()
        m_intID = 0
        m_strReasonType = String.Empty
        m_strExtra1 = String.Empty
        m_strExtra2 = String.Empty
        m_strExtra3 = String.Empty
        m_strUsername = String.Empty
        m_dtCREATE = NoDate()
        m_dtupdate = NoDate()
        m_strapptype = String.Empty

    End Sub

    Private Sub LoadLine(ByVal pClass As cTraveler_Reason_Type_LateShipment, ByRef pdrRow As DataRow)

        Try

            If Not (pdrRow.Item("ID").Equals(DBNull.Value)) Then pClass.ID = pdrRow.Item("ID")
            If Not (pdrRow.Item("REASON_TYPE").Equals(DBNull.Value)) Then pClass.REASONTYPE = pdrRow.Item("REASON_TYPE").ToString
            If Not (pdrRow.Item("EXTRA_1").Equals(DBNull.Value)) Then pClass.EXTRA1 = pdrRow.Item("EXTRA_1").ToString
            If Not (pdrRow.Item("EXTRA_2").Equals(DBNull.Value)) Then pClass.EXTRA2 = pdrRow.Item("EXTRA_2").ToString
            If Not (pdrRow.Item("EXTRA_3").Equals(DBNull.Value)) Then pClass.EXTRA3 = pdrRow.Item("EXTRA_3").ToString
            If Not (pdrRow.Item("USER_LOGIN").Equals(DBNull.Value)) Then pClass.USERNAME = pdrRow.Item("USER_LOGIN").ToString
            If Not (pdrRow.Item("CREATE_TS").Equals(DBNull.Value)) Then pClass.CREATEDATE = pdrRow.Item("CREATE_TS").ToString
            If Not (pdrRow.Item("Update_Ts").Equals(DBNull.Value)) Then pClass.UPDATEDATE = pdrRow.Item("Update_Ts").ToString
            If Not (pdrRow.Item("AppType").Equals(DBNull.Value)) Then pClass.APPTYPE = pdrRow.Item("AppType").ToString

        Catch ex As Exception
            MsgBox("Error in cTraveler_Reason_Type_LateShipment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub



    Public Function Load(ByVal _id As Int32) As cTraveler_Reason_Type_LateShipment

        Load = New cTraveler_Reason_Type_LateShipment()

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "SELECT * " &
            "FROM   Traveler_Reason_Type_LateShipment WITH (Nolock) where ID = " & _id

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(Load, dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cTraveler_Reason_Type_LateShipment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Public Function LoadLST() As List(Of cTraveler_Reason_Type_LateShipment)
        'this function rturn same like below + color name
        '   LoadLST = New List(Of cTraveler_Reason_Type_LateShipment)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim myList = New List(Of cTraveler_Reason_Type_LateShipment)
            Dim i As Integer
            Dim sql As String


            sql = "SELECT * FROM Traveler_Reason_Type_LateShipment WITH (Nolock) where AppType = 'OEI' "

            dt = db.DataTable(sql)

            If dt.Rows.Count <> 0 Then
                For i = 0 To dt.Rows.Count - 1
                    Dim oEnum = New cTraveler_Reason_Type_LateShipment
                    Call LoadLine(oEnum, dt.Rows(i))
                    myList.Add(oEnum)
                Next
            End If

            Return myList


        Catch ex As Exception
            MsgBox("Error in cTraveler_Reason_Type_LateShipment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

#Region "--------------------------- Public properties ---------------------- "


    Public Property ID As Int32
        Get
            Return m_intID
        End Get
        Set(value As Int32)
            m_intID = value
        End Set
    End Property
    Public Property REASONTYPE As String
        Get
            Return m_strReasonType
        End Get
        Set(value As String)
            m_strReasonType = value
        End Set
    End Property
    Public Property EXTRA1 As String
        Get
            Return m_strExtra1
        End Get
        Set(value As String)
            m_strExtra1 = value
        End Set
    End Property
    Public Property EXTRA2 As String
        Get
            Return m_strExtra2
        End Get
        Set(value As String)
            m_strExtra2 = value
        End Set
    End Property
    Public Property EXTRA3 As String
        Get
            Return m_strExtra3
        End Get
        Set(value As String)
            m_strExtra3 = value
        End Set
    End Property
    Public Property CREATEDATE As DateTime
        Get
            Return m_dtCREATE
        End Get
        Set(value As DateTime)
            m_dtCREATE = value
        End Set
    End Property
    Public Property UPDATEDATE As DateTime
        Get
            Return m_dtupdate
        End Get
        Set(value As DateTime)
            m_dtupdate = value
        End Set
    End Property
    Public Property USERNAME As String
        Get
            Return m_strUsername
        End Get
        Set(value As String)
            m_strUsername = value
        End Set
    End Property
    Public Property APPTYPE As String
        Get
            Return m_strapptype
        End Get
        Set(value As String)
            m_strapptype = value
        End Set
    End Property
#End Region

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone
    End Function
End Class
