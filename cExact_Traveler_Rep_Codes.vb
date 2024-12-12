Imports System.Collections.Generic
Imports System.Data.Odbc


Public Class cExact_Traveler_Rep_Codes
    Implements ICloneable





#Region " -------------------------------- Private properties ---------------------------------"

    Private m_intID As Int32
    Private m_strrepname As String


#End Region


    Sub New()
        Call Init()
    End Sub


    Private Sub Init()
        m_intID = 0
        m_strrepname = String.Empty


    End Sub

    Private Sub LoadLine(ByVal pClass As cExact_Traveler_Rep_Codes, ByRef pdrRow As DataRow)

        Try

            If Not (pdrRow.Item("ID").Equals(DBNull.Value)) Then pClass.ID = pdrRow.Item("ID")
            If Not (pdrRow.Item("repname").Equals(DBNull.Value)) Then pClass.REPNAME = pdrRow.Item("repname").ToString



        Catch ex As Exception
            MsgBox("Error in cExact_Traveler_Rep_Codes." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub



    Public Function Load(ByVal _id As Int32) As cExact_Traveler_Rep_Codes

        Load = New cExact_Traveler_Rep_Codes()

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "SELECT ID,Enum_Value As repname " &
            "FROM   MDB_CFG_ENUM WITH (Nolock) where ID = " & _id

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(Load, dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cExact_Traveler_Rep_Codes." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Public Function LoadLST() As List(Of cExact_Traveler_Rep_Codes)
        'this function rturn same like below + color name
        '   LoadLST = New List(Of cTraveler_Reason_Type_LateShipment)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim myList = New List(Of cExact_Traveler_Rep_Codes)
            Dim i As Integer
            Dim sql As String


            sql = "SELECT ID,Enum_Value As repname  FROM  MDB_CFG_ENUM WITH (Nolock) where ENUM_CAT = 'OEI-SSP'"

            dt = db.DataTable(sql)

            '  If dt.Rows.Count <> 0 Then
            For i = 0 To dt.Rows.Count - 1
                Dim oEnum = New cExact_Traveler_Rep_Codes
                Call LoadLine(oEnum, dt.Rows(i))
                myList.Add(oEnum)
            Next
            '    End If

            Return myList


        Catch ex As Exception
            MsgBox("Error in cExact_Traveler_Rep_Codes." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
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
    Public Property REPNAME As String
        Get
            Return m_strrepname
        End Get
        Set(value As String)
            m_strrepname = value
        End Set
    End Property


#End Region

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone
    End Function



End Class
