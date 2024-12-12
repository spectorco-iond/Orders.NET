Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module mThermometer
    'justin change from  172.16.10.201 to 172.16.11.71 as per ticket 31785 on 20230227
    'Private Const RESTPath As String = "http://172.16.10.201:8182/default/spector/"
    Private Const RESTPath As String = "http://172.16.11.71:8182/default/spector/"

    Public Function checkAvailability(ByVal requestDt As Date, ByVal machineData As String, ByVal units As Integer, ByVal timeSpecs() As Decimal, ByVal ignoreIntermediateThreshold As Boolean, Optional ByVal recurseOnFail As Boolean = False, Optional ByVal machineList As ArrayList = Nothing) As String
        Dim isAvailable As Boolean = False
        Dim isValidDate = False

        'Parse JSON data and check validity
        Try

            Dim jsonMachineData As cMachineData = JsonConvert.DeserializeObject(Of cMachineData)(machineData)

            'basic info
            Dim machineNo As String = jsonMachineData.machineID.ToString().ToUpper()
            Dim unitType As String = jsonMachineData.unitType.ToString.ToUpper()
            Dim availabilitySet As ArrayList = jsonMachineData.availabilities
            Dim thresholds As ArrayList


            'requested amounts and current amounts
            Dim actualRequestedAmount As Decimal = 0.0
            Dim currentAmount As Decimal = 0.0
            Dim topThreshold As Decimal = 0.0

            'Get max date
            Dim maxAvailSet As JArray = availabilitySet.Item(availabilitySet.Count() - 1)
            Dim maxDateFound As Date = Date.Parse(maxAvailSet.Item(0).ToString)
            maxAvailSet = Nothing

            'get machine group aggregate totals for the requested day
            Dim decoGroupDayOverload = False
            Dim groupThresholds As Decimal() = {0.0, 0.0}
            If Not machineList Is Nothing Then
                Dim aggregateTotals() As Decimal = getAggregateDayTotals(machineList, requestDt.ToShortDateString, ignoreIntermediateThreshold)
                If aggregateTotals(1) - aggregateTotals(0) <= 0 Then
                    decoGroupDayOverload = True
                End If
            End If

            'determine availablility for current date if applicable
            If decoGroupDayOverload = False Then

                For Each availability As JArray In availabilitySet
                    ' MsgBox(Date.Parse(availability.Item(0).ToString))
                    Dim availabilityDate = Date.Parse(availability.Item(0).ToString)

                    If Date.Compare(availabilityDate, Date.Parse(requestDt).ToShortDateString) = 0 Then
                        isValidDate = True

                        'get thresholds either in hours or by units (usually by hours) and current and requested production amounts
                        If Not jsonMachineData.hourThresholds Is Nothing And unitType = "TIME" Then
                            thresholds = getAdjustedThresholds(jsonMachineData.hourThresholds, getCurrentCapacityModifier(requestDt, machineNo))
                            actualRequestedAmount = timeSpecs(0) + (units / timeSpecs(1))
                            currentAmount = System.Convert.ToDecimal(availability.Item(1).ToString)
                        Else
                            thresholds = getAdjustedThresholds(jsonMachineData.unitThresholds, getCurrentCapacityModifier(requestDt, machineNo))
                            actualRequestedAmount = units
                            currentAmount = System.Convert.ToDecimal(availability.Item(2).ToString)
                        End If

                        'Thresholds.Item(2) for red threshold (rocket service orders) and thresholds.Item(1) for intermediate threshold (regular orders)
                        'Use minimum threshold thresholds.Item(0) if regular order is already in the intermediate zone
                        If ignoreIntermediateThreshold = True Then
                            topThreshold = thresholds.Item(2)
                        ElseIf ignoreIntermediateThreshold = False And currentAmount < thresholds.Item(0) Then
                            topThreshold = thresholds.Item(1)
                        Else
                            topThreshold = thresholds.Item(0)
                        End If

                        'check if new amount under maximum allowed
                        If (currentAmount + actualRequestedAmount) <= topThreshold Then
                            isAvailable = True
                        End If

                        Exit For
                    End If
                Next
            End If

            If Date.Parse(requestDt) >= maxDateFound Then
                recurseOnFail = False
            End If

            'escape recursive check is recurse set to false or if a valid date has been found
            If recurseOnFail = False Or isAvailable = True Then
                'true or false on chosen day for a total amount where the current amount is this and the max amount is that
                Return isAvailable.ToString & "_" & requestDt.ToShortDateString & "_" & (currentAmount + actualRequestedAmount) & "_" & currentAmount & "_" & topThreshold
            End If

        Catch er As Exception
            MsgBox("Error in mThermometer->checkAvailability: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            Return isAvailable.ToString & "_" & requestDt.ToShortDateString & "_0_0_0"
        End Try

        Return checkAvailability(requestDt.AddDays(1), machineData, units, timeSpecs, ignoreIntermediateThreshold, True, machineList)
    End Function

    Public Function getAvailableMachines(ByVal impMethod As String, ByVal defaultMachine As String, Optional ByVal includeOSRMachines As Boolean = False) As ArrayList
        Dim machineList As New ArrayList

        Dim strSql As String
        Dim conn As New cDBA
        Dim dtMachines As New DataTable

        Dim osrFilter = IIf(includeOSRMachines = True, "OR (ss.DEC_MET_GROUP_DESC = gc.DEC_MET_GROUP_DESC AND ISNULL(ss.GROUP_CODE, '') = '')", "")

        strSql = "SELECT station_number " _
                 & " FROM " _
                 & " ( " _
                 & "    SELECT ss.station_number, CASE WHEN ss.STATION_NUMBER LIKE '%OSR%' THEN 1 ELSE 0 END AS is_osr " _
                 & "    FROM station_substitutes AS ss " _
                 & "    INNER JOIN ( " _
                 & "	    SELECT GROUP_CODE, DEC_MET_GROUP_DESC FROM station_substitutes WHERE STATION_NUMBER = '" & defaultMachine & "' " _
                 & "	 ) AS gc " _
                 & "	    ON (gc.GROUP_CODE = ss.GROUP_CODE AND gc.GROUP_CODE <> '') " & osrFilter & " " _
                 & "    WHERE ss.station_number <> '" & defaultMachine & "' AND ss.IS_ACTIVE = 1 AND ISNULL(ss.PARENT_MACHINE, '') = '' " _
                 & "        AND ss.AGGREGATE_EXCLUDE = 0 " _
                 & " ) as machines " _
                 & " ORDER BY is_osr DESC"

        'strSql = "SELECT ss.station_number " _
        '         & " FROM station_substitutes as ss " _
        '         & "    INNER JOIN ( " _
        '         & "        SELECT GROUP_CODE, DEC_MET_GROUP_DESC FROM station_substitutes WHERE STATION_NUMBER = '" & defaultMachine & "' " _
        '         & "    ) as gc ON gc.GROUP_CODE = ss.GROUP_CODE " _
        '         & " WHERE ss.station_number <> '" & defaultMachine & "' AND gc.GROUP_CODE <> ''"

        Try

            machineList.Add(defaultMachine.ToString) 'add default first

            dtMachines = conn.DataTable(strSql)
            If dtMachines.Rows.Count <> 0 Then
                For Each dtMachine As DataRow In dtMachines.Rows
                    'Add OSR machine(s) before default machine
                    If dtMachine.Item("station_number").ToString.Contains("OSR") And includeOSRMachines Then
                        machineList.Insert(0, dtMachine.Item("station_number").ToString)
                    Else
                        machineList.Add(dtMachine.Item("station_number").ToString)
                    End If
                Next
            End If

        Catch er As Exception
            MsgBox("Error in mThermometer->getAvailableMachines: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        Return machineList
    End Function

    Public Function getAggregateDayTotals(ByVal machineList As ArrayList, ByVal requestDt As Date, ByVal ignoreIntermediateThreshold As Boolean) As Array
        Dim aggSpecs() As Decimal = {0.0, 0.0} 'current usage, max usage
        'Dim currentUsage = 0
        'Dim maxUsage = 0
        Dim thresholds As ArrayList
        Dim unitType As String = ""

        For Each machine As String In machineList

            'et machine data and deserialize
            Dim machineData As String = getMachineDataViaRest(machine, requestDt.ToShortDateString)
            Dim jsonMachineData As cMachineData = JsonConvert.DeserializeObject(Of cMachineData)(machineData)

            If jsonMachineData.availabilities Is Nothing Then
                Continue For
            End If

            'get unit type (time or units)
            unitType = jsonMachineData.unitType.ToString.ToUpper()

            'get all availabilities and append machine total aggregates for the matching day
            Dim availabilitySet As ArrayList = jsonMachineData.availabilities
            For Each availability As JArray In availabilitySet

                Dim availabilityDate = Date.Parse(availability.Item(0).ToString)

                If Date.Compare(availabilityDate, Date.Parse(requestDt).ToShortDateString) <> 0 Then
                    Continue For
                End If

                'get thresholds either in hours or by units (usually by hours) and current and requested production amounts
                If Not jsonMachineData.hourThresholds Is Nothing And unitType = "TIME" Then
                    thresholds = getAdjustedThresholds(jsonMachineData.hourThresholds, getCurrentCapacityModifier(requestDt, machine))
                    aggSpecs(0) = aggSpecs(0) + System.Convert.ToDecimal(availability.Item(1).ToString)
                Else
                    thresholds = getAdjustedThresholds(jsonMachineData.unitThresholds, getCurrentCapacityModifier(requestDt, machine))
                    aggSpecs(0) = aggSpecs(0) + System.Convert.ToDecimal(availability.Item(2).ToString)
                End If

                If ignoreIntermediateThreshold = True Then
                    aggSpecs(1) = aggSpecs(1) + thresholds.Item(2)
                Else
                    aggSpecs(1) = aggSpecs(1) + thresholds.Item(1)
                End If

            Next

        Next

        Return aggSpecs
    End Function
    Public Function simulateNextAvailDateAndMachine(ByVal impMethod As String, ByVal units As Integer, ByVal timeSpecs() As Decimal, ByVal defaultMachine As String, ByVal ignoreIntermediateThreshold As Boolean, ByVal requestDt As Date, Optional ByVal recurseOnFail As Boolean = True, Optional ByVal includeOSRMachines As Boolean = False) As String

        Dim isValid As Boolean = False
        Dim validDate As Date
        Dim validMachine As String = ""
        Dim estHours As Decimal
        Dim currentHours As Decimal
        Dim maxHours As Decimal
        Dim isOSRMachine As Boolean = False

        Dim machineList As ArrayList = getAvailableMachines(impMethod, defaultMachine, includeOSRMachines)
        For Each machine As String In machineList
            Dim machineData As String = getMachineDataViaRest(machine, requestDt.ToShortDateString)
            If Not machineData.Contains("availabilities") Then
                Continue For
            ElseIf validMachine.Contains("OSR") And machine.Contains("OSR") = False And isValid = True Then
                Exit For
            End If

            Dim retDataPieces() = checkAvailability(requestDt, machineData, units, timeSpecs, ignoreIntermediateThreshold, recurseOnFail, machineList).Split("_")

            If (validDate = Nothing OrElse Date.Compare(Convert.ToDateTime(retDataPieces(1)).ToShortDateString, validDate) <= 0) _
                And (estHours = 0.0 OrElse Convert.ToDecimal(retDataPieces(2).ToString) < estHours) Then
                'only show as valid if truly valid
                If Convert.ToBoolean(retDataPieces(0).ToString) = True Then
                    isValid = True
                    validDate = Convert.ToDateTime(retDataPieces(1)).ToShortDateString
                End If
                validMachine = machine
                estHours = Convert.ToDecimal(retDataPieces(2).ToString)
                currentHours = Convert.ToDecimal(retDataPieces(3).ToString)
                maxHours = Convert.ToDecimal(retDataPieces(4).ToString)
            End If
        Next

        Dim validationDetails As String = " Machine chosen: " & validMachine & vbCrLf _
                                        & " Imprint Method: " & impMethod & vbCrLf _
                                        & " Validated date: " & isValid.ToString & vbCrLf _
                                        & " Valid date found: " & IIf(isValid = True, validDate, "--") & vbCrLf _
                                        & " Original date requested: " & requestDt.ToShortDateString & vbCrLf _
                                        & " Estimated new total hours on machine: " & estHours & vbCrLf _
                                        & " Current hours on machine: " & currentHours & vbCrLf _
                                        & " Max hours allowed on machine: " & maxHours

        Return validDate & "_" & validationDetails
    End Function

    Public Function getMachineDataViaRest(ByVal machineNo As String, Optional ByVal requestDt As Date = Nothing, Optional ByVal formatData As Boolean = False) As String
        Dim machineData As String = ""

        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader
        Dim address As Uri
        Dim postStream As Stream = Nothing

        If requestDt = Nothing Then
            requestDt = Date.Now.ToShortDateString
        End If

        address = New Uri(RESTPath & machineNo.ToString.Trim() & "?date=" & requestDt.ToString("yyyy/M/dd"))

        Try
            ' Create the web request  
            request = DirectCast(WebRequest.Create(address), HttpWebRequest)

            ' Get response  
            response = DirectCast(request.GetResponse(), HttpWebResponse)

            ' Get the response stream into a reader  
            reader = New StreamReader(response.GetResponseStream())

            ' text application output  
            machineData = reader.ReadToEnd()

            If formatData = True Then
                'TODO: format data
            End If

        Catch er As Exception
            MsgBox("Error in thermoRequestTest: " & New Diagnostics.StackFrame(0).GetMethod.Name & Erl() & vbCrLf & er.Message & vbCrLf & er.InnerException.ToString)

        Finally
            If Not response Is Nothing Then
                response.Close()
            End If
        End Try

        Return machineData
    End Function

    Public Function getDefaultMachine(ByVal item_no As String, ByVal impMethod As String) As String
        Dim defaultMachine As String = ""

        Dim strSql As String
        Dim conn As New cDBA
        Dim dtMachine As New DataTable

        strSql = "SELECT mach_no as defaultMachine " _
                 & "FROM EXACT_TRAVELER_PRODUCTION_TIMES as pt " _
                 & " INNER JOIN imitmidx_sql as s " _
                 & "    ON s.item_no = pt.item_no " _
                 & " WHERE pt.item_no = '" & item_no & "' AND pt.decorating_method = '" & impMethod & "'" _
                 & "    AND pt.Step IN ('Printing', 'Lasering', 'Plate Stamping', 'Doming') "

        '----------------- ID 03.14.2023 --------------------------
        Dim _Facil_Loc As Int32 = 2409

        If Trim(m_oOrder.Ordhead.Mfg_Loc) = "US1" Then _Facil_Loc = 2410

        strSql = "SELECT mach_no as defaultMachine " _
                 & "FROM EXACT_TRAVELER_PRODUCTION_TIMES as pt " _
                 & " INNER JOIN imitmidx_sql as s " _
                 & " ON s.item_no = pt.item_no " _
                 & " WHERE pt.item_no = '" & item_no & "' AND pt.decorating_method = '" & impMethod & "'" _
                 & "  AND pt.Step IN ('Sewn','Trimming','Lasering','NFC Program and Assembly','Flex Transfer','Sewing','Cutting','Hotstamping','Heat Press','Sealing','Print Transfer','Doming', " _
                 & " 'Cut','Plate Stamping', 'NFC Program and Sewing','Lamination','Die Cutting','3 Way Trim','Round Cut','Printing','Apply Transfer','Digibinder','HP Printing','Glueing') " _
                 & " AND pt.Facility_Location = " & _Facil_Loc

        '-------------------------------------------------------------

        Try
            dtMachine = conn.DataTable(strSql)

            If dtMachine.Rows.Count <> 0 Then
                defaultMachine = dtMachine.Rows(0).Item("defaultMachine").ToString
            End If

        Catch er As Exception
            MsgBox("Error in mThermometer->getDefaultMachine: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        Return defaultMachine

    End Function

    Public Function getTimeSpecs(ByVal item_no As String, ByVal impMethod As String) As Array
        Dim timeSpecs() As Decimal = {0.0, 0.0}

        Dim strSql As String
        Dim conn As New cDBA
        Dim dtTimeSpecs As New DataTable

        strSql = "SELECT pt.prod_rate, pt.setup_time " _
                 & "FROM EXACT_TRAVELER_PRODUCTION_TIMES as pt " _
                 & " INNER JOIN imitmidx_sql as s " _
                 & "    ON s.item_no = pt.item_no " _
                 & " WHERE pt.item_no = '" & item_no & "' AND pt.decorating_method = '" & impMethod & "'" _
                 & "    AND pt.Step IN ('Printing', 'Lasering', 'Plate Stamping', 'Doming') "


        '--------------- ID 03.14.2023 ---------------------
        Dim _Facil_Loc As Int32 = 2409

        If Trim(m_oOrder.Ordhead.Mfg_Loc) = "US1" Then _Facil_Loc = 2410


        strSql = "SELECT pt.prod_rate, pt.setup_time " _
                 & "FROM EXACT_TRAVELER_PRODUCTION_TIMES as pt " _
                 & " INNER JOIN imitmidx_sql as s " _
                 & "    ON s.item_no = pt.item_no " _
                 & " WHERE pt.item_no = '" & item_no & "' AND pt.decorating_method = '" & impMethod & "'" _
                 & "    AND pt.Step IN ('Sewn','Trimming','Lasering','NFC Program and Assembly','Flex Transfer','Sewing','Cutting','Hotstamping','Heat Press','Sealing','Print Transfer','Doming', " _
                 & " 'Cut','Plate Stamping', 'NFC Program and Sewing','Lamination','Die Cutting','3 Way Trim','Round Cut','Printing','Apply Transfer','Digibinder','HP Printing','Glueing') " _
                 & " AND pt.Facility_Location = " & _Facil_Loc

        'm_oOrder.Ordhead.Mfg_Loc 
        '----------------------------------------------------
        Try
            dtTimeSpecs = conn.DataTable(strSql)

            If dtTimeSpecs.Rows.Count <> 0 Then
                timeSpecs(0) = dtTimeSpecs.Rows(0).Item("setup_time").ToString
                timeSpecs(1) = dtTimeSpecs.Rows(0).Item("prod_rate").ToString
            End If

        Catch er As Exception
            MsgBox("Error in mThermometer->getTimeSpecs: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        Return timeSpecs

    End Function

    Public Sub appendToAudit(ByVal machineNo As String, ByVal isValidDate As Boolean, ByVal requestDt As Date, ord_no As String,
                                ord_line As String, ByVal requestResultDetails As String, ByVal auditSource As String, ByVal item_no As String,
                                ByVal route_id As Integer, ByVal imp_method As String, ByVal qty As Integer)
        Dim strSql As String
        Dim conn As New cDBA

        Try
            Dim order_no_field As String = IIf(IsNumeric(ord_no.TrimStart) = True, "MACOLA_ORD_NO", "OTHER_ORD_NO")
            Dim order_line_field As String = IIf(IsNumeric(ord_line.TrimStart) = True, "TRAVELER_HEADERID", "OTHER_ORD_LINE")

            strSql = "INSERT INTO THERMO_SIMULATION_AUDIT (MACHINE_NO, IS_VALID_DATE, VALID_DATE, " & order_no_field & ", " _
                    & order_line_field & ", SIMULATION_DETAILS, SIMULATION_SOURCE, ITEM_NO, ROUTE_ID, IMP_METHOD, QTY, AUD_USER, AUD_DT) VALUES " _
                    & "('" & machineNo & "', " & Convert.ToInt32(isValidDate) & ", " _
                    & IIf(requestDt = Date.MinValue Or isValidDate = False, "NULL", "'" & requestDt & "'") & ", " _
                    & "'" & ord_no & "', '" & ord_line & "', '" & requestResultDetails & "', '" & auditSource & "', " _
                    & "'" & item_no & "', '" & route_id & "', '" & imp_method & "', '" & qty & "', " _
                    & "'" & "BANKERS\" & Environment.UserName & "', GETDATE())"

            conn.Execute(strSql)

        Catch er As Exception
            MsgBox("Error in mThermometer->appendToAudit: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Public Function getAllApplicableRoutes(RouteID As String) As ArrayList
        Dim impMethods As New ArrayList

        Dim strSql As String
        Dim conn As New cDBA

        Dim dtMethods As New DataTable

        strSql = "select Department as impMethod FROM EXACT_TRAVELER_STEP " _
            & " WHERE Step_Extra3 = 'Y' AND RouteID = " & RouteID _
            & " GROUP BY Department"

        Try
            dtMethods = conn.DataTable(strSql)

            If dtMethods.Rows.Count <> 0 Then
                For Each impRow As DataRow In dtMethods.Rows
                    impMethods.Add(impRow.Item("impMethod").ToString())
                Next
            End If

        Catch er As Exception
            MsgBox("Error in mThermometer->getAllApplicableRoutes: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        Return impMethods
    End Function

    Public Function daysGap(ByVal requestedDt As Date, Optional ByVal ignoreWeekends As Boolean = True) As Integer
        Dim dayGap As Integer = 0
        Dim currentDate As Date = Date.Parse(Date.Now.ToShortDateString)

        While (currentDate < requestedDt)
            If ignoreWeekends = True AndAlso (currentDate.DayOfWeek = DayOfWeek.Saturday Or currentDate.DayOfWeek = DayOfWeek.Sunday) Then
                'Continue While
            Else
                dayGap = dayGap + 1
            End If
            currentDate = currentDate.AddDays(1)
        End While
        Return dayGap
    End Function

    Public Function dateAdjuster(ByVal requestedDt As Date, ByVal daysToAdjust As Integer, Optional ByVal ignoreWeekends As Boolean = True) As Date
        Dim adjustValMultiplier As Integer = IIf(daysToAdjust < 0, -1, 1)
        Dim dayAdjust As Integer = 1

        'adjust for weekends
        If requestedDt.AddDays((1 * adjustValMultiplier)).DayOfWeek = DayOfWeek.Saturday Then
            dayAdjust = IIf(adjustValMultiplier < 0, 2, 3)
        ElseIf requestedDt.AddDays((1 * adjustValMultiplier)).DayOfWeek = DayOfWeek.Sunday Then
            dayAdjust = IIf(adjustValMultiplier < 0, 3, 2)
        End If

        'adjust date and return if necessary
        If daysToAdjust <> 0 AndAlso Date.Compare(requestedDt.AddDays(dayAdjust * adjustValMultiplier).ToShortDateString, Date.Now.ToShortDateString) >= 0 Then
            requestedDt = requestedDt.AddDays(dayAdjust * adjustValMultiplier).ToShortDateString
        Else
            Return requestedDt
        End If

        Return dateAdjuster(requestedDt, (daysToAdjust - (1 * adjustValMultiplier)), ignoreWeekends)
    End Function

    Public Function getCurrentCapacityModifier(ByVal requestedDt As Date, ByVal machineNo As String) As Double
        Dim capacityModifier As Double = 1.0

        Dim strSql As String
        Dim conn As New cDBA
        Dim dtCapacity As New DataTable


        strSql = "SELECT m.STATION_NUMBER, " _
            & " CASE WHEN m.IS_ACTIVE = 0 THEN 0 WHEN s.ADJUSTED_CAPACITY IS NOT NULL THEN s.ADJUSTED_CAPACITY ELSE 100 END AS ADJUSTED_CAPACITY " _
            & " FROM station_substitutes AS m " _
            & " LEFT JOIN station_scheduled_maintenance as s " _
            & " ON s.STATION_NUMBER = m.STATION_NUMBER AND OFFLINE_START_DATE <= '" & requestedDt.ToShortDateString & "' AND OFFLINE_END_DATE >= '" & requestedDt.ToShortDateString & "' " _
            & " WHERE m.STATION_NUMBER = '" & machineNo & "'"

        Try
            dtCapacity = conn.DataTable(strSql)

            If dtCapacity.Rows.Count <> 0 Then
                capacityModifier = Convert.ToDouble(dtCapacity.Rows(0).Item("ADJUSTED_CAPACITY").ToString) / 100.0
            End If
        Catch er As Exception
            MsgBox("Error in mThermometer->getCurrentCapacityModifier: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        Return capacityModifier
    End Function

    Public Function getAdjustedThresholds(ByVal thresholds As ArrayList, ByVal capacityModifier As Double) As ArrayList
        For i As Integer = 0 To thresholds.Count - 1
            thresholds(i) = Convert.ToInt32(Math.Floor(thresholds(i) * capacityModifier))
        Next

        Return thresholds
    End Function

    Public Function getAllowableOrderUnitSizeByImprint(ByVal impMethod As String, ByVal totalType As Integer, Optional ByVal excludeNonOSRMachines As Boolean = False, Optional ByVal machineNo As String = "") As Integer
        Dim strSql As String
        Dim conn As New cDBA
        Dim dtOrderSize As New DataTable

        'define query filters
        Dim machineFilter = IIf(machineNo <> "", " AND STATION_NUMBER = '" & machineNo & "'", "")
        machineFilter = machineFilter + IIf(excludeNonOSRMachines = True, " AND STATION_NUMBER LIKE '%OSR%'", "")

        Dim totOrderSize As Integer = -1
        Dim queryType = ""

        Select Case totalType
            Case 1
                queryType = "SUM"
            Case 2
                queryType = "MAX"
            Case Else
                queryType = "MIN"
        End Select

        strSql = "SELECT ISNULL(" & queryType & "(MAX_UNITS_PER_ORDER), 0) AS MAX_UNITS_PER_ORDER " _
            & " FROM station_substitutes " _
            & " WHERE DEC_MET_GROUP_DESC = '" & impMethod & "' " & machineFilter

        Try
            dtOrderSize = conn.DataTable(strSql)

            If dtOrderSize.Rows.Count <> 0 Then
                totOrderSize = Convert.ToInt16(dtOrderSize.Rows(0).Item("MAX_UNITS_PER_ORDER").ToString)
            End If

        Catch er As Exception
            MsgBox("Error in mThermometer->getAllowableOrderUnitSizeByImprint: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        Return totOrderSize
    End Function
End Module