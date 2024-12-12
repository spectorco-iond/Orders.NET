Public Class cAllStarProgram


#Region "Private variables ################################################"

    Public m_ClassificationId As String
    Public m_Cmp_Code As String
    Public m_Item_No As String
    Public m_Cmp_Fctry As String
    Public m_Cmp_Name As String
    Public m_Region As String
    Public m_KeyAccount As String
    Public m_CustNo As String
    Public m_Vol_Reb As String
    Public m_CreditIssued As String
    Public m_OrdVolAmt As String
    Public m_BalVolReb As String
    Public m_TotIssued As String


#End Region

#Region "Public constructors ##############################################"

    Public Sub New()

            Call Init()
            Call GetAllStarInformation(m_oOrder.Ordhead.Cus_No)

        End Sub

        Public Sub New(ByVal pstrItem_No As String)



        End Sub

#End Region

        Private Sub GetAllStarInformation(ByVal customer_code As String)

            Dim strSql As String
            Dim db As New cDBA
            Dim dt As DataTable

            Try

            strSql = "SELECT ClassificationId, cmp_code, cmp_name, textfield4, cmp_fctry, region " &
                "FROM cicmpy " &
                "WHERE cmp_code = '" & Trim(customer_code) & "' AND ClassificationId In('1ST', '2ST', '3ST', '4ST', '5ST', '6ST', '7ST') AND " &
                "(cmp_fctry IN('CA', 'US') OR (region LIKE '%USA%' AND region NOT LIKE '%CA%'))"

            'result is same 
            strSql = " select * from ( " _
                & " SELECT ClassificationId, cmp_code, cmp_name, textfield4, cmp_fctry, region " _
                & " FROM cicmpy WHERE  ClassificationId In('1ST', '2ST', '3ST', '4ST', '5ST', '6ST', '7ST') " _
                & " and region like '%USA%' " _
                & " Union " _
                & " SELECT ClassificationId, cmp_code, cmp_name, textfield4, cmp_fctry, region " _
                & " FROM cicmpy WHERE  ClassificationId In('1ST', '2ST', '3ST', '4ST', '5ST', '6ST', '7ST')  " _
                & " and region not like '%USA%' " _
                & "    ) as viewAllStarInformation " _
                & "  where cmp_code = '" & Trim(customer_code) & "'  order by ClassificationId, cmp_code  asc "


            'in the  exception if containe USA and ele if contain CA

            dt = db.DataTable(strSql)

                If dt.Rows.Count <> 0 Then
                    LoadlineAllStar(dt.Rows(0))
                End If

                'get marketing allowance information
                MarketingAllownace(customer_code)

            Catch er As Exception
                MsgBox("Error in cAllStarProgram." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

#Region "Public Properties ################################################"

        Public Property ClassificationId() As String
            Get
                ClassificationId = m_ClassificationId
            End Get
            Set(ByVal value As String)
                m_ClassificationId = value
            End Set
        End Property

        Public Property Cmp_Code() As String
            Get
                Cmp_Code = m_Cmp_Code
            End Get
            Set(ByVal value As String)
                m_Cmp_Code = value
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

        Public Property Cmp_Fctry() As String
            Get
                Cmp_Fctry = m_Cmp_Fctry
            End Get
            Set(ByVal value As String)
                m_Cmp_Fctry = value
            End Set
        End Property

        Public Property Cmp_Name() As String
            Get
                Cmp_Name = m_Cmp_Name
            End Get
            Set(ByVal value As String)
                m_Cmp_Name = value
            End Set
        End Property

        Public Property Region() As String
            Get
                Region = m_Region
            End Get
            Set(ByVal value As String)
                m_Region = value
            End Set
        End Property

    Public Property KeyAccount() As String
        Get
            KeyAccount = m_KeyAccount
        End Get
        Set(ByVal value As String)
            m_KeyAccount = value
        End Set
    End Property

    Public Property CustNo() As String
        Get
            CustNo = m_CustNo
        End Get
        Set(ByVal value As String)
            m_CustNo = value
        End Set
    End Property

    Public Property Vol_Reb() As String
        Get
            Vol_Reb = m_Vol_Reb
        End Get
        Set(ByVal value As String)
            m_Vol_Reb = value
        End Set
    End Property
    Public Property CreditIssued() As String
        Get
            CreditIssued = m_CreditIssued
        End Get
        Set(ByVal value As String)
            m_CreditIssued = value
        End Set
    End Property
    Public Property OrdVolAmt() As String
        Get
            OrdVolAmt = m_OrdVolAmt
        End Get
        Set(ByVal value As String)
            m_OrdVolAmt = value
        End Set
    End Property
    Public Property BalVolReb() As String
        Get
            BalVolReb = m_BalVolReb
        End Get
        Set(ByVal value As String)
            m_BalVolReb = value
        End Set
    End Property
    Public Property TotIssued() As String
        Get
            TotIssued = m_TotIssued
        End Get
        Set(ByVal value As String)
            m_TotIssued = value
        End Set
    End Property



#End Region

    Public Sub Init()

        m_ClassificationId = String.Empty
        m_Cmp_Code = String.Empty
        m_Item_No = String.Empty
        m_Cmp_Fctry = String.Empty
        m_Cmp_Name = String.Empty
        m_Region = String.Empty
        m_KeyAccount = String.Empty

        m_CustNo = String.Empty
        m_Vol_Reb = String.Empty
        m_CreditIssued = String.Empty
        m_OrdVolAmt = String.Empty
        m_BalVolReb = String.Empty
        m_TotIssued = String.Empty

    End Sub

        ''' <summary>
        '''  Returns true or false if customer is a key account client.  ''' 
        ''' </summary>
        ''' <returns>Boolean</returns>
        Public Function IsKeyAccountCustomer() As Boolean

            Dim blnIsKeyAccount As Boolean = False

            blnIsKeyAccount = (m_KeyAccount.Contains("Key Accounts")) 'There must be a consistency In the database To name key accounts this way

            IsKeyAccountCustomer = blnIsKeyAccount

        End Function

    Private Sub LoadlineAllStar(ByRef pdrRow As DataRow)

        Try

            If Not pdrRow.Item("ClassificationId").Equals(DBNull.Value) Then
                ClassificationId = pdrRow.Item("ClassificationId").ToString.Trim
            End If

            If Not pdrRow.Item("cmp_code").Equals(DBNull.Value) Then
                Cmp_Code = pdrRow.Item("cmp_code").ToString.Trim
            End If

            If Not pdrRow.Item("cmp_name").Equals(DBNull.Value) Then
                Cmp_Name = pdrRow.Item("cmp_name").ToString.Trim
            End If

            If Not pdrRow.Item("textfield4").Equals(DBNull.Value) Then
                KeyAccount = pdrRow.Item("textfield4").ToString.Trim
            End If

            If Not pdrRow.Item("cmp_fctry").Equals(DBNull.Value) Then
                Cmp_Fctry = pdrRow.Item("cmp_fctry").ToString.Trim
            End If

            If Not pdrRow.Item("region").Equals(DBNull.Value) Then
                Region = pdrRow.Item("region").ToString.Trim
            End If

        Catch er As Exception

            MsgBox("Error in cAllStarProgram. " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)

        End Try

    End Sub

    Private Sub MarketingAllownace(ByVal cusno As String)

            Dim strSql As String
            Dim db As New cDBA
            Dim dt As DataTable

            Try

                strSql = "SELECT custno, vol_reb, creditissued, ordvolamt, balvolreb, totissued " &
                " FROM Bankers_custom_VolumeRebate " &
                " WHERE custno = '" & Trim(cusno) & "' "

                dt = db.DataTable(strSql)

                If dt.Rows.Count <> 0 Then
                    LoadlineMarketing(dt.Rows(0))
                End If


            Catch er As Exception
                MsgBox("Error in cAllStarProgram." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        Private Sub LoadlineMarketing(ByRef pdrRow As DataRow)

            Try

                If Not pdrRow.Item("vol_reb").Equals(DBNull.Value) Then
                Vol_Reb = pdrRow.Item("vol_reb").ToString
            End If

                If Not pdrRow.Item("creditissued").Equals(DBNull.Value) Then
                CreditIssued = pdrRow.Item("creditissued").ToString
            End If

                If Not pdrRow.Item("ordvolamt").Equals(DBNull.Value) Then
                OrdVolAmt = pdrRow.Item("ordvolamt").ToString
            End If

                If Not pdrRow.Item("balvolreb").Equals(DBNull.Value) Then
                BalVolReb = pdrRow.Item("balvolreb").ToString
            End If

                If Not pdrRow.Item("totissued").Equals(DBNull.Value) Then
                TotIssued = pdrRow.Item("totissued").ToString
            End If

            Catch er As Exception

                MsgBox("Error in cAllStarProgram. " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)

            End Try

        End Sub
        '============================================================
        'All Star Program - December 6th 2017 - Clayton Solomon
        '============================================================

        Private Sub AllStarProgram(ByVal companyCode As String)
            Try
                Dim db As New cDBA
                Dim dt As DataTable
                Dim _25percent As Double = 0.25 'applied to 3ST and lower
                Dim _50percent As Double = 0.5  'applied to 4ST and higher
                Dim starLevel As String = ""
                Dim orderType As String = m_oOrder.Ordhead.Ord_Type

                Dim sqlQuery As String = "SELECT ClassificationId, cmp_code, cmp_name, textfield4, region " &
                "FROM cicmpy " &
                "WHERE cmp_code = '" & companyCode.Trim & "' " &
                "AND region like '%USA%' OR region like '%CA%' " &
                "ORDER BY id"

                dt = db.DataTable(sqlQuery)

                If dt.Rows.Count > 0 Then

                    starLevel = dt.Rows(0).Item("ClassificationId").ToString.ToUpper

                    Select Case starLevel

                        Case "1ST", "2ST", "3ST"

                            If orderType = "RS" Then
                                'Customer must pay freight or be billed for it
                                'Item under $20 (EQP NET) are no charge
                                'Item OVER $20 (EQP NET) are charged at EQP-25%
                                'Key Accounts: EQP - 50%
                            ElseIf orderType = "SB" Or orderType = "SS" Then
                                'Customer must pay freight or be billed for it
                                'Pricing: EQP-25% (set-up, item cost and all other charges)
                                'Key Accounts: EQP - 50%
                            ElseIf orderType = "SP" Then
                                'Pricing: EQP-25% (set-up, item cost and all other charges)
                                'Key Accounts: EQP - 50%
                            End If

                        Case "4ST", "5ST", "6ST", "7ST"

                            If orderType = "RS" Then
                                'Customer must pay freight or be billed for it
                                'Item under $20 (EQP NET) are no charge
                                'Item OVER $20 (EQP NET) are charged at EQP-25%
                                'Key Accounts: EQP - 50%
                            ElseIf orderType = "SB" Or orderType = "SS" Then
                                'Customer must pay freight or be billed for it
                                'Pricing: EQP-25% (set-up, item cost and all other charges)
                                'Key Accounts: EQP - 50%
                            ElseIf orderType = "SP" Then
                                'Customer must pay freight or be billed for it
                                'Pricing: EQP-25% (set-up, item cost and all other charges)
                                'Key Accounts: EQP - 50%
                            End If

                        Case "CA", "N/A"

                    End Select
                End If

            Catch ex As Exception
                MsgBox("Error in AllStarProgram. " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
            End Try
        End Sub


    End Class
