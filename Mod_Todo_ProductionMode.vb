Option Explicit On

Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.Byte
Imports System.Int32
Imports System.Buffer
Imports System.BitConverter

Imports System.Linq.Expressions
Imports System.Linq

Module Mod_Todo_ProductionMode
    'Dim distinctDates() As Long
    Dim highCount() As Integer
    'Sub test_setAll_ApptTodoRecords4()
    '    'Dim x As Integer
    '    'x = DateDiff(DateInterval.Day, #1/1/0001#, Date.Today.Date)
    '    'Debug.Print(#1/1/0001# & " " & x)
    '    'Debug.Print(DaysToToday)
    '    'Exit Sub


    '    'startTime()
    '    ApptGV.UserID = 1
    '    Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
    '    set_CreateDateArray() '0.007 seconds
    '    'endTime()
    '    ' Debug.Print("CreateDateArray")
    '    'Exit Sub
    '    'startTime()
    '    'setTodoDistinctDates() '0.001 to 0.005 seconds B=NOW .003 to .006 seconds
    '    'endTime()

    '    Dim aa() As ApptTodoRecordType
    '    Dim bb() As ApptTodoRecordType
    '    '===============================================testtest
    '    Dim i As Integer
    '    startTime()

    '    'For i = 1 To ApptGV.TodoDistinctDates_Cnt
    '    '    aa = getApptRecords_AllFields_ApptTodo(ApptGV.todoDistinctDatesStr(i)) '0.0997327   '.1  seconds
    '    '    bb = getApptRecords_AllFields_ApptTodo_NEW(ApptGV.TodoDistinctDates(i)) '0.0388957  '.05 to .06 seconds abount twice as fast as above

    '    '    'Debug.Print(UBound(aa) & " " & UBound(bb))
    '    'Next
    '    'endTime()
    '    'Exit Sub
    '    'aa = getApptRecords_AllFields_ApptTodo("2019/11")
    '    Dim L As Long = #11/1/2019#.Ticks
    '    aa = getApptRecords_AllFields_ApptTodo_NEW(L)
    '    'If UBound(aa) <> UBound(bb) Then
    '    '    Debug.Print(UBound(aa) & " " & UBound(bb))
    '    '    i = i
    '    'End If
    '    For i = 1 To UBound(aa)
    '        'If aa(i).dTics <> bb(i).dTics Or aa(i).Msg <> bb(i).Msg Then
    '        '    Debug.Print(aa(i).ID & aa(i).Msg & " " & bb(i).ID & bb(i).Msg)
    '        '    Debug.Print(New Date(aa(i).dTicsOriginal) & " " & New Date(bb(i).dTicsOriginal))
    '        '    i = i
    '        'End If
    '        With aa(i)
    '            Debug.Print(i.ToString("00") & " " & .Msg & " del=" & .DeleteFlag)
    '        End With
    '    Next
    '    endTime()
    '    Debug.Print("test")
    '    '===============================================endtest
    '    Exit Sub
    '    'startTime()
    '    'aa = getApptRecords_AllFields_ApptTodo("CC") '0.04 to 0.05 seconds
    '    'endTime()
    '    'Debug.Print(UBound(aa))

    '    '==
    '    startTime()
    '    aa = getTodoRecs_IncludingDeletedToday_ApptTodo(1) '0.001
    '    endTime()
    '    Debug.Print(UBound(aa))
    '    Exit Sub
    '    '==
    '    '1262304000000000 - 1/1/0005
    '    '636884640000000000 - 3/18/2019
    '    '636918336000000000 - 4/26/2019
    '    '636919200000000000 - 4/27/2019
    '    '636920928000000000 - 4/29/2019
    '    '636921792000000000 - 4/30/2019
    '    '636922656000000000 - 5/1/2019
    '    '636923520000000000 - 5/2/2019
    '    '636925248000000000 - 5/4/2019
    '    '636926976000000000 - 5/6/2019
    '    '636927840000000000 - 5/7/2019
    '    '636929568000000000 - 5/9/2019
    '    '636936480000000000 - 5/17/2019
    '    '637045344000000000 - 9/20/2019
    '    '637047936000000000 - 9/23/2019
    '    '637049664000000000 - 9/25/2019
    '    '==
    '    '1262304000000000
    '    '636884640000000000
    '    '636918336000000000
    '    '636919200000000000
    '    '636920928000000000
    '    '636921792000000000
    '    '636922656000000000
    '    '636923520000000000
    '    '636925248000000000
    '    '636926976000000000
    '    '636927840000000000
    '    '636929568000000000
    '    '636936480000000000
    '    '637045344000000000
    '    '637047936000000000
    '    '637049664000000000


    '    Dim highCnt As Integer
    '    'startTime()
    '    'highCnt = getHighCountFromApptTodoRecords_TodoCnt() ' - 0 seconds 17 reads
    '    'endTime()

    '    startTime() '38 times faster than below routine
    '    'Dim i As Integer
    '    'Dim a() As Long = {1262304000000000, 636884640000000000, 636918336000000000, 636919200000000000, 636920928000000000, 636921792000000000, 636922656000000000, 636923520000000000, 636925248000000000, 636926976000000000, 636927840000000000, 636929568000000000, 636936480000000000, 637045344000000000, 637047936000000000, 637049664000000000}
    '    ReDim highCount(UBound(ApptGV.distinctDates))
    '    For i = 1 To UBound(ApptGV.distinctDates) '0 To a.Count - 1
    '        highCnt = getHighCountFromApptTodoRecords_TodoCnt(ApptGV.distinctDates(i) + TimeSpan.TicksPerMinute) '(a(i) + TimeSpan.TicksPerMinute) ' - 0 seconds 17 reads
    '        highCount(i) = highCnt
    '        'Debug.Print(New Date(a(i)) & "highCnt=" & highCnt)
    '        'Debug.Print(New Date(ApptGV.distinctDates(i)) & "highCnt=" & highCnt)
    '    Next
    '    endTime()
    '    ''=============

    '    'For i = 1 To 10
    '    '    Debug.Print(ApptGV.All_ApptTodoRecords(i).ApptDateStr)
    '    'Next
    '    ''==============


    '    Debug.Print("Above # seconds to get high count using bisearch - highCnt = " & highCnt)

    '    startTime()
    '    For i = 1 To UBound(ApptGV.distinctDates)
    '        highCnt = getHighCountFromApptTodoRecords(ApptGV.distinctDates(i)) ' 0.012 seconds
    '        If highCnt <> highCount(i) Then
    '            i = i
    '        End If
    '    Next

    '    endTime()
    '    Debug.Print(highCnt)
    'End Sub
    Sub test_setAll_ApptTodoRecords3()
        Dim i As Integer

        ApptGV.UserID = 1
        startTime()
        'setAll_ApptTodoRecords() '0.2 seconds
        ''above does the following
        'ApptGV.All_ApptTodoRecords = getAll_ApptTodoRecordType_UsingReader() '0.2 secs

        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds

        endTime()

        'testing
        Dim recs() As ApptTodoRecordType = ApptGV.All_ApptTodoRecords.Clone



        startTime()
        'QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))

        QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt) '0.0408907 seconds (This sort is Twice as fast as the Array.sort lambda function below)

        'Array.Sort(ApptGV.All_ApptTodoRecords, (Function(x As ApptTodoRecordType, y As ApptTodoRecordType) x.dTics.CompareTo(y.dTics))) '0.03 seconds for the sort

        endTime()

        'Dim recs() As ApptTodoRecordType = ApptGV.All_ApptTodoRecords.Clone
        startTime()
        Array.Sort(recs, (Function(x As ApptTodoRecordType, y As ApptTodoRecordType) x.dTics.CompareTo(y.dTics))) '0.03 seconds for the sort
        endTime()

        For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
            If ApptGV.All_ApptTodoRecords(i).dTics <> recs(i).dTics Then
                i = i
            End If
        Next

        Exit Sub

        Debug.Print("set all records - total=" & ApptGV.All_ApptTodoRecords_Cnt)

        Dim startDate As String = "2019/09/20"
        Dim endDate As String = "2019/09/25"

        Dim a() As ApptTodoRecordType
        'a = getPreliminaryRecs_SDateEdate(startDate, endDate, IncludeTODOs)

        startTime()
        'a = getPreliminaryRecs_SDateEdate(startDate, endDate) '0.02 to 0.05
        a = getPreliminaryRecs_SDateEdate(startDate, endDate, True) '0.02 to 0.05 - less than 0.1 second
        endTime()

        Dim nItems As Integer = UBound(a)
        Debug.Print(nItems)

        'For i = 1 To nItems
        '    With a(i)
        '        Debug.Print(.ApptDateStr & " " & .Msg)
        '    End With
        'Next

        Dim cnt As Integer
        startTime()

        'cnt = getHighCountFromApptTodoRecords() 'todo records

        endTime()
        Debug.Print("cnt=" & cnt)


    End Sub
    'Sub test_setAll_ApptTodoRecords2()
    '    startTime()
    '    setAll_ApptTodoRecords() '0.2 seconds
    '    endTime()

    '    startTime()
    '    QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger)) '0.02 seconds for 100,000 records
    '    endTime()

    '    startTime()

    '    'test_getPat()
    '    test_getPat2()
    '    endTime()
    '    Exit Sub

    '    'Debug.Print(ApptGV.All_ApptTodoRecords(1).ApptDateStr)'0005/01/01 00:00:00.0000003

    '    Dim i As Integer
    '    startTime()

    '    For i = 1 To ApptGV.All_ApptTodoRecords_Cnt '0.02 seconds
    '        If Mid(ApptGV.All_ApptTodoRecords(i).ApptDateStr, 1, 10) = "0005/01/01" Then '0.16 seconds
    '            'If InStr(ApptGV.All_ApptTodoRecords(i).Msg, "Pat") Then
    '            'Debug.Print(i & ApptGV.All_ApptTodoRecords(i).Msg)
    '        End If
    '    Next
    '    endTime()
    'End Sub

    'Private Sub SetStart_End_DatesETC()
    Function GetApptTodoRecs_SDateEDate(ByVal startDate As String, ByVal endDate As String, Optional ByVal IncludeTODOs As Boolean = False) As ApptTodoRecordType()
        Dim a() As ApptTodoRecordType
        a = getPreliminaryRecs_SDateEdate(startDate, endDate, IncludeTODOs)
        'Dim nItems As Integer = UBound(a)

        'Dim Q As String
        ''appt_ApptRecs = getApptRecords_AllFields(Q)

        'If IncludeTODOs = True Then
        '    Q = "select * from Appt where scUserID= " & ApptGV.UserID & " and deleteflag=0 and ((apptDate>= '" & appt_startDate & "' and apptDate<='" & appt_EndDate & "') or (apptdate='0005/01/01')) order by substring(apptdatestr,1,1) desc, dtics"
        'Else
        '    Q = "select * from Appt where scUserID= " & ApptGV.UserID & " and apptDate>= '" & appt_startDate & "' and apptDate<='" & appt_EndDate & "' order by substring(apptdatestr,1,1) desc, dtics"
        'End If
        Return a
    End Function
    'Function getDistinctDates(ByVal startDate As String, ByVal endDate As String) As ApptTodoRecordType()
    '    Dim a() As ApptTodoRecordType
    '    'appt_distinctDates = QgetDates(Q)
    '    'Q = "Select distinct apptdate from Appt where scUserID= " & ApptGV.UserID & " And apptDate>='" & appt_startDate & "' and apptDate<='" & appt_EndDate & "' order by 1" 'apptdatestr"
    '    Return a
    'End Function
    Sub setTodoDistinctDates() '0.003 seconds
        'startTime()
        ReDim ApptGV.TodoDistinctDates(1000000)

        Dim cnt As Integer = 0



        Dim i As Integer

        Dim startPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
        Dim endPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerMinute)


        Dim x As Long
        Dim y As Integer


        '==
        cnt += 1
        ApptGV.TodoDistinctDates(cnt) = Date.Today.Date.Ticks 'add today's date in ticks form
        x = (ApptGV.TodoDistinctDates(cnt) \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay
        y = New Date(x).Day - 1
        x = x - y * TimeSpan.TicksPerDay
        ApptGV.TodoDistinctDates(cnt) = x
        '==


        For i = startPt To endPt - 1 '1 To ApptGV.All_ApptTodoRecords_Cnt
            With ApptGV.All_ApptTodoRecords(i)
                If .apptDate = ApptGV.todoTicks AndAlso .DeleteFlag = 0 Then
                    'x = .dTicsOriginal Mod TimeSpan.TicksPerDay '+ TimeSpan.
                    x = (.dTicsOriginal \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay
                    y = New Date(x).Day - 1
                    x = x - y * TimeSpan.TicksPerDay
                    ' Debug.Print(New Date(x))
                    cnt += 1
                    ApptGV.TodoDistinctDates(cnt) = x
                End If
            End With
        Next

        ReDim Preserve ApptGV.TodoDistinctDates(cnt)
        If cnt = 1 Then
            Exit Sub
        End If

        QuickSort_Long(ApptGV.TodoDistinctDates, 1, cnt)

        Dim newCnt As Integer = 1
        Dim ss(cnt) As Long
        ss(newCnt) = ApptGV.TodoDistinctDates(cnt)
        For i = cnt - 1 To 1 Step -1
            'Debug.Print(s(i) & " " & ss(newCnt))
            If ApptGV.TodoDistinctDates(i) <> ss(newCnt) Then

                newCnt += 1
                ss(newCnt) = ApptGV.TodoDistinctDates(i)
                'Debug.Print(s(i) & " " & ss(newCnt))
            End If
        Next
        ReDim ApptGV.TodoDistinctDates(newCnt)
        ReDim ApptGV.todoDistinctDatesStr(newCnt)
        For i = 1 To newCnt
            ApptGV.TodoDistinctDates(i) = ss(i)
            ApptGV.todoDistinctDatesStr(i) = Format(New Date(ss(i)), "yyyy/MM")
        Next
        ReDim ss(0)

        ApptGV.TodoDistinctDates_Cnt = newCnt

        'For i = 1 To newCnt
        '    Debug.Print(New Date(ApptGV.TodoDistinctDates(i)))
        '    '11/1/2019
        '    '10/1/2019
        '    '9/1/2019
        '    '5/1/2019
        '    '4/1/2019
        'Next
        'endTime()


    End Sub

    Sub setTodoDistinctDatesOLD()
        startTime()
        ReDim ApptGV.TodoDistinctDates(1000000)

        Dim cnt As Integer = 0



        Dim i As Integer

        Dim startPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
        Dim endPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerMinute)


        Dim x As Long
        Dim y As Integer


        '==
        cnt += 1
        ApptGV.TodoDistinctDates(cnt) = Date.Today.Date.Ticks 'add today's date in ticks form
        x = (ApptGV.TodoDistinctDates(cnt) \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay
        y = New Date(x).Day - 1
        x = x - y * TimeSpan.TicksPerDay
        ApptGV.TodoDistinctDates(cnt) = x
        '==


        For i = startPt To endPt - 1 '1 To ApptGV.All_ApptTodoRecords_Cnt
            With ApptGV.All_ApptTodoRecords(i)
                If .apptDate = ApptGV.todoTicks AndAlso .DeleteFlag = 0 Then
                    'x = .dTicsOriginal Mod TimeSpan.TicksPerDay '+ TimeSpan.
                    x = (.dTicsOriginal \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay
                    y = New Date(x).Day - 1
                    x = x - y * TimeSpan.TicksPerDay
                    ' Debug.Print(New Date(x))
                    cnt += 1
                    ApptGV.TodoDistinctDates(cnt) = x
                    'ApptGV.TodoDistinctDates(cnt) = (.dTicsOriginal \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay '- (.dTicsOriginal Mod TimeSpan.TicksPerDay + TimeSpan.TicksPerDay)
                    'Debug.Print(ApptGV.TodoDistinctDates(cnt) & " " & New Date(ApptGV.TodoDistinctDates(cnt)) & " " & x & " : " & New Date(x))
                    'Debug.Print(New Date(x) & " " & y)
                End If
                'If Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate AndAlso .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 Then
                '    'Debug.Print(.ApptDateStr)

                '    cnt += 1
                '    a(cnt) = ApptGV.All_ApptTodoRecords(i)


                '    ApptGV.TodoDistinctDates(cnt) = (a(cnt).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '
                'End If
            End With
        Next


        ReDim Preserve ApptGV.TodoDistinctDates(cnt)
        'Debug.Print(UBound(ApptGV.TodoDistinctDates))
        'startTime()
        ApptGV.TodoDistinctDates = ApptGV.TodoDistinctDates.Distinct.ToArray '0.002 seconds for 5400 records
        'Debug.Print(UBound(ApptGV.TodoDistinctDates))
        'endTime()

        Array.Sort(ApptGV.TodoDistinctDates)
        'Array.Reverse(ApptGV.TodoDistinctDates)

        ApptGV.TodoDistinctDates_Cnt = UBound(ApptGV.TodoDistinctDates)

        Dim c() As Long = ApptGV.TodoDistinctDates.Clone

        ReDim ApptGV.todoDistinctDatesStr(ApptGV.TodoDistinctDates_Cnt)
        cnt = 0
        For i = ApptGV.TodoDistinctDates_Cnt To 1 Step -1 'ApptGV.TodoDistinctDates.Count - 1 'ApptGV.TodoDistinctDates_Cnt '- 1
            cnt += 1
            'ApptGV.todoDistinctDatesStr(cnt) = Format(New Date(ApptGV.TodoDistinctDates(i)), "yyyy/MM")
            ApptGV.todoDistinctDatesStr(cnt) = Format(New Date(c(i)), "yyyy/MM")
            ApptGV.TodoDistinctDates(cnt) = c(i)

            'Debug.Print(ApptGV.TodoDistinctDates(i) & " - " & New Date(ApptGV.TodoDistinctDates(i)))
            'Debug.Print(distinctDates(i))
        Next
        ReDim c(0)

        '        For i = 1 To ApptGV.TodoDistinctDates_Cnt
        '            Debug.Print(ApptGV.todoDistinctDatesStr(i))
        '        Next
        '2019/10
        '2019/09
        '2019/05
        '2019/04
        'For i = 1 To ApptGV.TodoDistinctDates_Cnt
        'For i = 0 To ApptGV.TodoDistinctDates.Count - 1 'ApptGV.TodoDistinctDates_Cnt '- 1
        '    Debug.Print(ApptGV.TodoDistinctDates(i) & " - " & New Date(ApptGV.TodoDistinctDates(i)))
        '    'Debug.Print(distinctDates(i))
        'Next
        'Debug.Print(ApptGV.TodoDistinctDates_Cnt)
        '636896736000000000 - 4/1/2019
        '636922656000000000 - 5/1/2019
        '637028928000000000 - 9/1/2019
        endTime()
    End Sub
    Function getPreliminaryRecs_SDateEdate(ByVal sDate As String, ByVal eDate As String, Optional ByVal IncludeTODOs As Boolean = False) As ApptTodoRecordType()
        Dim a(1000000) As ApptTodoRecordType
        ReDim ApptGV.distinctDates(1000000)
        Dim cnt As Integer = 0
        Dim i As Integer
        If IncludeTODOs = False Then
            For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
                With ApptGV.All_ApptTodoRecords(i)
                    '
                    If Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate AndAlso .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 Then
                        'Debug.Print(.ApptDateStr)

                        cnt += 1
                        a(cnt) = ApptGV.All_ApptTodoRecords(i)

                        'ApptGV.distinctDates(cnt) = a(cnt).apptDate
                        ApptGV.distinctDates(cnt) = (a(cnt).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '
                    End If
                End With
            Next
        Else
            For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
                With ApptGV.All_ApptTodoRecords(i)
                    'If Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate AndAlso .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 Then
                    If .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 And (Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate) Or (Mid(.ApptDateStr, 1, 10) = "0005/01/01") Then
                        cnt += 1
                        a(cnt) = ApptGV.All_ApptTodoRecords(i)

                        ApptGV.distinctDates(cnt) = (a(cnt).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '  a(cnt).apptDate
                        'distinctDates(cnt) = a(cnt).dTics
                    End If
                End With
            Next
        End If

        'appt_distinctDates
        ReDim Preserve a(cnt)
        ReDim Preserve ApptGV.distinctDates(cnt)

        'startTime()
        ApptGV.distinctDates = ApptGV.distinctDates.Distinct.ToArray '0.002 seconds for 5400 records
        'endTime()
        Array.Sort(ApptGV.distinctDates)
        For i = 1 To UBound(ApptGV.distinctDates)
            Debug.Print(ApptGV.distinctDates(i) & " - " & New Date(ApptGV.distinctDates(i)))
            'Debug.Print(distinctDates(i))
        Next
        Debug.Print(UBound(ApptGV.distinctDates))

        'order by substring(apptdatestr, 1, 1) desc, dtics

        'Array.Sort(a, (Function(x As LongIntegerType, y As LongIntegerType) x.L.CompareTo(y.L)))
        startTime()
        Array.Sort(a, (Function(x As ApptTodoRecordType, y As ApptTodoRecordType) x.dTics.CompareTo(y.dTics))) '0.03 seconds for the sort
        endTime()
        Debug.Print("sort")

        'Array.Sort(a, (Function(x As ApptTodoRecordType, y As ApptTodoRecordType) y.dTics.CompareTo(x.dTics))) 'decr decrement
        Return a
    End Function
#Region "AddApptTodoRecord"
    'Sub AddApptTodoRecord(ByVal rec As ApptRecAllType) ', Optional ByVal useTODO_DB As Boolean = False) 'use for recurring records  'new 9/5/2015
    Function formatApptTodoDateStr(ByVal dTics As Long) As String
        Dim x As String = Format(New DateTime(dTics), "yyyy/MM/dd HH:mm:ss.fffffff")
        Return x
    End Function
    Function createPureDate_Ticks_FromTicks(ByVal x As Long) As Long 'Date 'from createPureDateFromTicks
        'createPureDateFromTicks
        'Dim retTicks as Long = New Date((x \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay)
        Dim retTicks As Long = (x \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay
        Return retTicks
    End Function



    'Sub addApptTodoRecord_test()
    '    ',[dTicsOriginal] ,[dTicsCompleted]
    '    Dim xDate As String = formatApptTodoDateStr(rec.dTics) 'formatApptDateStr(rec.dTics)
    '    'Dim aDate As Date = createPureDateFromTicks(rec.dTics)
    '    Dim aDate As Long = createPureDate_Ticks_FromTicks(rec.dTics)
    '    Dim L As Long = Date.Now.Ticks

    '    '===========NEW
    '    Dim ff As Integer = FreeFile()
    '    Dim filename As String = xBuildFullFileName(f_Appt.FileName)
    '    FileOpen(ff, filename, OpenMode.Random, , , f_Appt.FileLen)
    '    Dim xCount As Integer = LOF(ff) / f_Appt.FileLen + 1

    '    '[ID],[dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag],[DeleteAbsolute],[scUserID],[apptDateStr],[apptDate],[dTicsOriginal],[dTicsCompleted]

    '    Dim xMsg As String = rec.Msg
    '    SQLreplace(xMsg)

    '    With rec
    '        .ID = xCount
    '        .dTics = rec.dTics
    '        .Msg = xMsg
    '        .AcctNum = 0
    '        .StrikeOut = 0
    '        .DeleteFlag = 0
    '        .DeleteAbsolute = 0
    '        .scUserID = 1
    '        .ApptDateStr = formatApptTodoDateStr(rec.dTics)
    '        .apptDate = createPureDate_Ticks_FromTicks(rec.dTics)
    '        .dTicsOriginal = L
    '        .dTicsCompleted = 0

    '        '.dTics = xTicks 'd.Ticks
    '    End With
    '    FilePut(ff, rec, xCount)
    '    FileClose(ff)

    '    '===========
    '    ApptGV.All_ApptTodoRecords_Cnt += 1
    '    ReDim Preserve ApptGV.All_ApptTodoRecords(xCount)
    '    ApptGV.All_ApptTodoRecords(xCount) = rec



    'End Sub
    Function getHighCnt(ByVal xTicks As Long) As Integer
        Dim yyyymmddhhmm As Long = (xTicks \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute
        Dim NextMinute As Long = yyyymmddhhmm + TimeSpan.TicksPerMinute
        Dim LocInFile As Integer
        LocInFile = biSearch_GTE_ApptTodoRecords(NextMinute) - 1
        Dim dif As Long = ApptGV.All_ApptTodoRecords(LocInFile).dTics - yyyymmddhhmm
        If dif > 0 Then
            Return CInt(dif + 1)
        Else
            Return 1
        End If
    End Function
    'Function getHighCountFromApptTodoRecords(Optional ByVal xDate As String = Nothing) As Integer
    Function getHighCountFromApptTodoRecords(ByVal xDate As Long) As Integer ' = 0) As Integer
        Dim i As Integer
        Dim cnt As Integer = 0
        Dim temp As Integer = 0
        'Dim sDate As String
        'Dim c As Integer = 0
        'If xDate = 0 Then 'Nothing Then
        '    sDate = "0005/01/01"
        '    For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
        '        With ApptGV.All_ApptTodoRecords(i)
        '            'apptdate='0005/01/01' and scUserID 
        '            '
        '            If Mid(.ApptDateStr, 1, 10) = sDate AndAlso .scUserID = ApptGV.UserID Then
        '                'c += 1
        '                'If c > 10 Then Exit For
        '                temp = .dTics Mod 10000000
        '                'Debug.Print(.ID & " " & temp & "=cnt " & .ApptDateStr & " " & .dTics & " " & New Date(.dTics) & " " & .Msg)
        '                If temp > cnt Then
        '                    cnt = temp
        '                    'Debug.Print(cnt)
        '                    'If cnt > 8561302 Then
        '                    '    Debug.Print(temp & " " & .ID & " " & temp & "=cnt " & .ApptDateStr & " " & .dTics & " " & New Date(.dTics) & " " & .Msg)
        '                    '    i = i
        '                    'End If

        '                    'If i = ApptGV.All_ApptTodoRecords_Cnt Then Debug.Print(.ID & " " & temp & "=cnt " & .ApptDateStr & " " & .dTics & " " & New Date(.dTics) & " " & .Msg)
        '                End If

        '            End If
        '        End With

        '    Next
        'Else
        ' sDate = xDate
        For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
            With ApptGV.All_ApptTodoRecords(i)
                '
                'If Mid(.ApptDateStr, 1, 19) = sDate AndAlso .scUserID = ApptGV.UserID Then
                If (.dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute = xDate AndAlso .scUserID = ApptGV.UserID Then
                    'Debug.Print(.ApptDateStr)
                    temp = .dTics Mod 10000000
                    If temp > cnt Then cnt = temp
                End If
            End With
        Next
        'end If
        'cnt = QInt("select cast(isnull(max(dtics % 10000000),0) as int) from [dbo].[Appt] where apptdate='" & xSelectedDate & "' and scUserID =" & ApptGV.UserID) 'new 5/2/2019
        Return cnt

    End Function
#End Region'AddApptTodoRecord
    Sub test_getAllTodoRecords()


        ApptGV.UserID = 1
        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds


        startTime()
        Dim a() As ApptTodoRecordType = getAllTodoRecords() '0.001 seconds
        endTime()

        Dim nItems As Integer = UBound(a)
        Dim i As Integer
        For i = 1 To 19 'nItems
            With a(i)
                Debug.Print(i & " " & .ID & "=id " & .ApptDateStr & " " & .DeleteFlag & "=del " & .Msg)
            End With

        Next
    End Sub
    Function getAllTodoRecords() As ApptTodoRecordType()
        '1262304000000000
        'Dim todoTicks As Long = 1262304000000000 '1/1/0005'
        Dim NextMinute As Long = ApptGV.todoTicks + TimeSpan.TicksPerMinute
        Dim startLocInFile As Integer
        Dim endLocInFile As Integer
        startLocInFile = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
        endLocInFile = biSearch_GTE_ApptTodoRecords(NextMinute) - 1
        Dim i As Integer
        Dim nItems As Integer = endLocInFile - startLocInFile + 1
        Dim recs(0) As ApptTodoRecordType
        Dim cnt As Integer = 0
        If nItems > 0 Then
            ReDim recs(nItems)
            For i = startLocInFile To endLocInFile
                If ApptGV.All_ApptTodoRecords(i).DeleteAbsolute = 0 Then '.DeleteFlag = 0 Then

                    cnt += 1
                    recs(cnt) = ApptGV.All_ApptTodoRecords(i)
                End If
            Next
        End If
        ReDim Preserve recs(cnt)
        Return recs
    End Function
    Function sortFunction1(ByVal a() As ApptTodoRecordType) 'deletes at bottom
        'Dim orderByResult = From s In studentList
        '                    Order By s.StudentName
        '                    Select s

        'Dim result = capitals.OrderBy(c >= c.Length).ThenBy(c >= c);
        'Dim result = a.OrderBy(c >= c.deleteflag).ThenBy(c >= c.dtics)
        'Dim result = a.OrderBy(Funtion(x As ApptTodoRecordType , y As ApptTodoRecordType ) x.dtics.compare2 y.dtics ))
        'Array.Sort.a(ApptTodoRecordType) '.OrderBy(Funtion(x As ApptTodoRecordType , y As ApptTodoRecordType ) x.dtics.compare2 y.dtics ))
        'Dim result = a.OrderBy(s.DeleteFlag, s.dTics * (s.DeleteFlag Mod 1), s.dTicsCompleted)


        'Dim AList As ArrayList = New ArrayList(StrArry) 'StrArry will be our array of strings
        Dim b As ArrayList = New ArrayList(a)


        'Order By s.DeleteFlag, s.dTics * (s.DeleteFlag Xor 1) Descending, s.dTicsCompleted

        'Dim orderByResults = From s In b 'a

        Dim orderByResults = From s In b
                             Order By s.DeleteFlag, s.dTics * (s.DeleteFlag Xor 1), s.dTicsCompleted
                             Select s

        Dim i As Integer
        'Dim nItems As Integer = UBound(orderByResults)
        Dim nItems As Integer = b.Count

        For i = 0 To b.Count - 1 'nItems - 1


            With orderByResults(i)
                'Debug.Print(.Msg)
                Debug.Print(.ID & " " & .Msg & " " & .ApptDateStr & " " & .DeleteFlag) '& " " & (.DeleteFlag Xor 1))
            End With

        Next



        Return orderbyresults

    End Function
    Function getTodoRecs_IncludingDeletedToday_ApptTodo(ByVal UserID As Integer, Optional ByVal DeletesAtBottom As Boolean = False) As ApptTodoRecordType() ' ApptRecAllType()
        'from: getTodoRecs_IncludingDeletedToday in Mappointment.vb

        '================
        Dim i As Integer
        Dim days2today As Integer = DaysToToday()

        Dim a(1000000) As ApptTodoRecordType
        'ReDim ApptGV.distinctDates(1000000)
        Dim cnt As Integer = 0

        Dim startPt As Integer
        Dim endPt As Integer
        'startPt = biSearch_GTE_AllAppts_Long(ApptGV.todoTicks) ' biSearch_E_AllAppt(search)
        'endPt = biSearch_GTE_AllAppts_Long(ApptGV.todoTicks + TimeSpan.TicksPerMinute) 'startPt ' BiSearch_Long_LTE(searchE)

        startPt = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks) ' biSearch_E_AllAppt(search)
        endPt = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerMinute) 'startPt ' BiSearch_Long_LTE(searchE)




        'If DeletesAtBottom = True Then
        '    For i = startPt To endPt - 1
        '        With ApptGV.All_ApptTodoRecords(i)


        '            If .apptDate = ApptGV.todoTicks AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID
        '                cnt += 1
        '                a(cnt) = ApptGV.All_ApptTodoRecords(i)


        '                'If cnt < 3 Then a(cnt).DeleteFlag = 1
        '            End If
        '        End With
        '    Next
        'Else
        For i = startPt To endPt - 1
            With ApptGV.All_ApptTodoRecords(i)
                If .apptDate = ApptGV.todoTicks AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID
                    cnt += 1
                    a(cnt) = ApptGV.All_ApptTodoRecords(i)


                    'If cnt < 3 Then a(cnt).DeleteFlag = 1
                End If
            End With
        Next
        'End If
        'Debug.Print(endPt - startPt)'24300/2700=9 16200/2700=6
        ReDim Preserve a(Math.Min(ApptGV.MaxApptTodos, cnt)) 'a(cnt)
        Return a
        '========

        Dim r() As ApptTodoRecordType 'ApptRecAllType NEW

        'Note: (deleteflag ^ 1) has the affect to turning (0 to 1) and (1 to 0)
        '*(dtics % 10000000) 
        '1. @days2today gets the number of days from 1/1/0001 to today
        '2. cast(dticsCompleted/864000000000 as int) computes number of days from 1/1/0001 to dticsCompleted date
        '3. if the diff is 0 then it's as TODAY completed record
        'Order by
        'order by deleteflag, dtics - puts deleted or completed records at the bottom
        'order by dtics - keeps existing order (deletes are intermingled with actives)



        Dim Q As String
        'difference is in the 'order by' statement
        If DeletesAtBottom = True Then
            Q = "declare @days2today int =datediff(day,'00010101',getdate())
 select * from appt where apptDate='0005/1/1' and scUserID=" & UserID & "  
 and (deleteFlag=0 or (deleteflag=1 and @days2today-cast(dticsCompleted/864000000000 as int)=0)) 
 order by deleteflag,dtics*(deleteflag ^1),dTicsCompleted"

            'order by deleteflag,dtics 'sorts inactives at bottom but retains the ORDER number so that the BOTTOM items will not be sorted by date deleted!

            'comments
            'deleteflag,dtics*(deleteflag ^1),dTicsCompleted
            '1. deleteflag separates active (top) from inactive (bottom)
            '2. dtics*(deleteflag ^1)
            '   has effect on TOP to order by order dtics % 10000000
            '   since (deleteflag ^1) is ZERO if deleteFlag is 1 - for Bottom items (inactives) this is a constant 0
            '3. finally, dticsCompleted has no effect on TOP items as dticsCompleted field is ZERO for actives but causes Bottom items to be sorted by dateTime deleted
            'NOTE: Top  Bottom
            '        0       1
            'dtics * 1  ...* 0
            '        0  dticsCompleted

            'NOTE: if i want to order by existing order ORDER but at the BOTTOM then use: 'order by deleteflag,dtics'
            'NOTE: if i want to order by existing order ORDER then use: order by deleteflag,dtics (Not at bottome)


            ' order by deleteflag, dtics"
        Else
            Q = "declare @days2today int =datediff(day,'00010101',getdate())
 select * from appt where apptDate='0005/1/1' and scUserID=" & UserID & "  
 and (deleteFlag=0 or (deleteflag=1 and @days2today-cast(dticsCompleted/864000000000 as int)=0))
 order by dtics"
        End If

        '==================
        'Dim fontReg As Font = DGV_Todo.Font 'DGV_Printing.Font
        'Dim fontForStrikeOut As Font = New Font("Microsoft Sans Serif", 8, FontStyle.Strikeout)



        'Debug.Print(Q)
        'Oct 1 2018
        r = getApptRecords_AllFields_ApptTodo("Something") 'getApptRecords_AllFields(Q) ', useTODO_DB)
        Return r
    End Function

    Function getApptRecords_AllFields_ApptTodo(ByVal SelectedYearMonth As String) As ApptTodoRecordType()
        '737339
        '737339
        Dim days2today As Integer = DaysToToday()

        Dim a(1000000) As ApptTodoRecordType
        'ReDim ApptGV.distinctDates(1000000)
        Dim cnt As Integer = 0
        Dim i As Integer
        'If IncludeTODOs = False Then
        '    For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
        '        With ApptGV.All_ApptTodoRecords(i)
        '            '
        '            If Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate AndAlso .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 Then
        '                'Debug.Print(.ApptDateStr)

        '                cnt += 1
        '                a(cnt) = ApptGV.All_ApptTodoRecords(i)

        '                'ApptGV.distinctDates(cnt) = a(cnt).apptDate
        '                ApptGV.distinctDates(cnt) = (a(cnt).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '
        '            End If
        '        End With
        '    Next
        'Else
        Dim startPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks) ' biSearch_E_AllAppt(search)
        Dim endPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerMinute) 'startPt ' BiSearch_Long_LTE(searchE)


        'For i = 1 To ApptGV.All_ApptTodoRecords_Cnt

        Dim createDateStr As String

        For i = startPt To endPt - 1
            With ApptGV.All_ApptTodoRecords(i)
                'If Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate AndAlso .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 Then
                'TimeSpan.TicksPerDay 864000000000
                'If .apptDate = ApptGV.todoTicks AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / 864000000000) = 0)) Then '.scUserID = ApptGV.UserID
                'If .apptDate = ApptGV.todoTicks AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID


                createDateStr = Format(New Date(.dTicsOriginal), "yyyy/MM")

                'Debug.Print(Mid(.ApptDateStr, 1, 7) & " " & SelectedYearMonth)
                'Debug.Print(createDateStr & " " & SelectedYearMonth)



                'If Mid(.ApptDateStr, 1, 7) = SelectedYearMonth AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID
                If createDateStr = SelectedYearMonth AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID
                    cnt += 1
                    a(cnt) = ApptGV.All_ApptTodoRecords(i)


                    'If cnt < 3 Then a(cnt).DeleteFlag = 1
                End If




                'If .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 And (Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate) Or (Mid(.ApptDateStr, 1, 10) = "0005/01/01") Then
                '            cnt += 1
                '            a(cnt) = ApptGV.All_ApptTodoRecords(i)

                '    'ApptGV.distinctDates(cnt) = (a(cnt).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '  a(cnt).apptDate
                '    'distinctDates(cnt) = a(cnt).dTics
                'End If
            End With
        Next
        'End If

        'appt_distinctDates
        ReDim Preserve a(Math.Min(ApptGV.MaxApptTodos, cnt)) 'a(6) 'a(cnt)
        'ReDim Preserve ApptGV.distinctDates(cnt)
        'a = sortFunction1(a)
        'For i = 1 To UBound(a) '10
        '    With a(i)
        '        Debug.Print(.ID & " " & .Msg & " " & .ApptDateStr & " " & .DeleteFlag)
        '    End With

        'Next
        Return a

    End Function
    Function getDeletedTodos_DeletedToday(Optional ByVal ReverseDateOrder As Boolean = False) As ApptTodoRecordType()

        Dim cnt As Integer = 0
        Dim i As Integer
        Dim d As Date = Date.Today
        Dim xday As Integer = d.Day
        Dim s As Long = getFirstOfMonthTicks(d.Ticks)

        Dim startPt As Integer = biSearch_GTE_TodoCompletedDates(s)
        Dim endPt As Integer = biSearch_GTE_TodoCompletedDates(s + TimeSpan.TicksPerDay)
        Dim nRecs As Integer = endPt - startPt

        Dim a(nRecs) As ApptTodoRecordType
        For i = startPt To endPt - 1

            With ApptGV.All_ApptTodoRecords(ApptGV.TodoCompletedDate(i).i)
                If New Date(.dTicsCompleted).Day = xday Then
                    'If .DeleteFlag = 1 Then
                    'note: these are all completed items so no need to test for being deleted!

                    'Debug.Print(cnt & " " & New Date(ApptGV.TodoCreateDate(i).L) & " " & " " & New Date(.dTicsOriginal) & " " & .Msg & " " & .DeleteFlag & " " & New Date(.dTicsCompleted))
                    cnt += 1
                    a(cnt) = ApptGV.All_ApptTodoRecords(ApptGV.TodoCompletedDate(i).i)
                    a(cnt).Msg = Trim(a(cnt).Msg)
                End If
            End With
        Next
        'ReDim Preserve a(Math.Min(ApptGV.MaxApptTodos, cnt))
        ReDim Preserve a(cnt)
        'QuickSort_ApptTodoRecords(a, 1, UBound(a))
        If ReverseDateOrder = False Then
            QuickSort_ApptTodoRecords_bydTicsCompleted(a, 1, UBound(a))
        Else
            QuickSort_ApptTodoRecords_descending_bydTicsCompleted(a, 1, UBound(a))
        End If

        Return a

    End Function
    Function getDeletedTodos_withinDateRange_NEW(ByVal sDate As Date, Optional ByVal ReverseDateOrder As Boolean = False) As ApptTodoRecordType()
        Debug.Print(sDate)
        Dim cnt As Integer = 0
        Dim i As Integer
        Dim s As Long = sDate.Ticks

        Dim startPt As Integer = biSearch_GTE_TodoCompletedDates(s)
        Dim endPt As Integer = biSearch_GTE_TodoCompletedDates(s + TimeSpan.TicksPerDay)
        Dim nRecs As Integer = endPt - startPt

        Dim a(nRecs) As ApptTodoRecordType
        For i = startPt To endPt - 1

            With ApptGV.All_ApptTodoRecords(ApptGV.TodoCompletedDate(i).i)
                'If .DeleteFlag = 1 Then
                'note: these are all completed items so no need to test for being deleted!

                'Debug.Print(cnt & " " & New Date(ApptGV.TodoCreateDate(i).L) & " " & " " & New Date(.dTicsOriginal) & " " & .Msg & " " & .DeleteFlag & " " & New Date(.dTicsCompleted))
                cnt += 1
                a(cnt) = ApptGV.All_ApptTodoRecords(ApptGV.TodoCompletedDate(i).i)
                a(cnt).Msg = TrimF(a(cnt).Msg)
                'End If
            End With
        Next
        'ReDim Preserve a(Math.Min(ApptGV.MaxApptTodos, cnt))
        ReDim Preserve a(cnt)
        'QuickSort_ApptTodoRecords(a, 1, UBound(a))
        If ReverseDateOrder = False Then
            QuickSort_ApptTodoRecords_bydTicsCompleted(a, 1, UBound(a))
        Else
            QuickSort_ApptTodoRecords_descending_bydTicsCompleted(a, 1, UBound(a))
        End If

        Return a
    End Function
    Function getApptRecords_AllFields_ApptTodo_NEW(ByVal SelectedYearMonth_LONG As Long, Optional ByVal ActiveOnly As Boolean = False) As ApptTodoRecordType()
        Dim cnt As Integer = 0
        Dim i As Integer
        Dim days2today As Integer = DaysToToday()

        Dim startPt As Integer = biSearch_GTE_TodoCreateDates(SelectedYearMonth_LONG)
        Dim endPt As Integer = biSearch_GTE_TodoCreateDates(SelectedYearMonth_LONG + TimeSpan.TicksPerDay)
        Dim nRecs As Integer = endPt - startPt

        Dim a(nRecs) As ApptTodoRecordType

        ReDim ApptGV.CreateDate_LOCATIONarray(nRecs) 'NEW

        If ActiveOnly = False Then
            For i = startPt To endPt - 1
                With ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)
                    If (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - (.dTicsCompleted \ TimeSpan.TicksPerDay) = 0)) Then
                        cnt += 1

                        ApptGV.CreateDate_LOCATIONarray(cnt) = ApptGV.TodoCreateDate(i).i 'NEW

                        a(cnt) = ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)
                        a(cnt).Msg = Trim(a(cnt).Msg)
                    End If
                End With
            Next
        Else
            'ActiveOnly
            For i = startPt To endPt - 1
                With ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)
                    If .DeleteFlag = 0 Then
                        cnt += 1

                        ApptGV.CreateDate_LOCATIONarray(cnt) = ApptGV.TodoCreateDate(i).i 'NEW

                        a(cnt) = ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)
                        a(cnt).Msg = Trim(a(cnt).Msg)
                    End If
                End With
            Next
        End If
        Dim ArraySizeToSave As Integer = Math.Min(ApptGV.MaxApptTodos, cnt)

        ReDim Preserve a(ArraySizeToSave) 'a(6) 'a(cnt)
        ReDim Preserve ApptGV.CreateDate_LOCATIONarray(ArraySizeToSave) 'NEW
        Return a

    End Function
    Function getApptRecords_AllFields_ApptTodo_NEWOLD(ByVal SelectedYearMonth_LONG As Long, Optional ByVal nonDeletesOnly As Boolean = False) As ApptTodoRecordType()

        Dim d As Date = New Date(SelectedYearMonth_LONG) 'New DateTime(d.Year, d.Month, 1).AddDays(-1)
        Dim dd As Date = d.AddMonths(1) 'New DateTime(d.Year, d.Month, 1).AddDays(-1)
        Dim NextMth As Long = dd.Ticks

        'Debug.Print(d & " " & dd) '5/1/2019 6/1/2019


        Dim days2today As Integer = DaysToToday()

        'Debug.Print(SelectedYearMonth_LONG & " " & New Date(SelectedYearMonth_LONG))

        Dim a(1000000) As ApptTodoRecordType
        Dim cnt As Integer = 0
        Dim i As Integer

        Dim startPt As Integer = biSearch_GTE_TodoCreateDates(SelectedYearMonth_LONG)
        'Debug.Print(New Date(ApptGV.TodoCreateDate(startPt - 1).L) & " ___")
        'Debug.Print(New Date(ApptGV.TodoCreateDate(startPt).L) & " !!!")

        'Dim endPt As Integer = biSearch_GTE_TodoCreateDates(SelectedYearMonth_LONG + TimeSpan.TicksPerDay)
        Dim endPt As Integer = biSearch_GTE_TodoCreateDates(NextMth)
        'Debug.Print(startPt & " " & endPt & " " & (endPt - startPt) / 2700)

        'Debug.Print(New Date(ApptGV.TodoCreateDate(endPt - 1).L) & " ______")
        'Debug.Print(New Date(ApptGV.TodoCreateDate(endPt).L) & " !!!!!")
        'biSearch_GTE_TodoCreateDates

        'For i = 1 To ApptGV.All_ApptTodoRecords_Cnt

        'Dim createDateStr As String
        If nonDeletesOnly = False Then
            For i = startPt To endPt - 1
                'With ApptGV.All_ApptTodoRecords(i)
                With ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)


                    ''''createDateStr = Format(New Date(.dTicsOriginal), "yyyy/MM")

                    'Debug.Print(Mid(.ApptDateStr, 1, 7) & " " & SelectedYearMonth)
                    'Debug.Print(createDateStr & " " & SelectedYearMonth)
                    ''''If i < 20 Then Debug.Print(createDateStr & " " & .Msg)



                    'If Mid(.ApptDateStr, 1, 7) = SelectedYearMonth AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID
                    'If createDateStr = SelectedYearMonth AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID

                    'If .DeleteFlag = 1 Then
                    '    Debug.Print(days2today & " " & CInt(.dTicsCompleted / TimeSpan.TicksPerDay))
                    '    Debug.Print(days2today & " " & CInt(.dTicsCompleted \ TimeSpan.TicksPerDay))
                    'End If

                    'If (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then
                    If (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - (.dTicsCompleted \ TimeSpan.TicksPerDay) = 0)) Then
                        cnt += 1

                        'a(cnt) = ApptGV.All_ApptTodoRecords(i)
                        a(cnt) = ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)

                        a(cnt).Msg = Trim(a(cnt).Msg)
                        'If cnt < 3 Then a(cnt).DeleteFlag = 1
                    End If




                    'If .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 And (Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate) Or (Mid(.ApptDateStr, 1, 10) = "0005/01/01") Then
                    '            cnt += 1
                    '            a(cnt) = ApptGV.All_ApptTodoRecords(i)

                    '    'ApptGV.distinctDates(cnt) = (a(cnt).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '  a(cnt).apptDate
                    '    'distinctDates(cnt) = a(cnt).dTics
                    'End If
                End With
            Next
        Else
            'nonDeletesOnly
            For i = startPt To endPt - 1
                With ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)
                    If .DeleteFlag = 0 Then
                        cnt += 1
                        a(cnt) = ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)
                        a(cnt).Msg = Trim(a(cnt).Msg)
                    End If
                End With
            Next
        End If

        'End If

        'appt_distinctDates
        ReDim Preserve a(Math.Min(ApptGV.MaxApptTodos, cnt)) 'a(6) 'a(cnt)
        'ReDim Preserve ApptGV.distinctDates(cnt)
        'a = sortFunction1(a)
        'For i = 1 To UBound(a) '10
        '    With a(i)
        '        Debug.Print(.ID & " " & .Msg & " " & .ApptDateStr & " " & .DeleteFlag)
        '    End With

        'Next
        Return a

    End Function
    Sub testAppt_apptTodo_ckForCounter()
        Dim a(1000000) As ApptTodoRecordType
        Dim cnt As Integer = 0
        Dim i As Integer

        For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
            With ApptGV.All_ApptTodoRecords(i)
                Debug.Print(.ID & " " & .ApptDateStr & " " & .dTics & " " & .Msg)
            End With

        Next i
        Exit Sub


        'Dim startPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerDay)
        Dim startPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)

        'Dim endPt As Integer = biSearch_GTE_TodoCreateDates(NextMth)
        'Dim endPt As Integer = ApptGV.All_ApptTodoRecords_Cnt ' biSearch_GTE_ApptTodoRecords(givenDate.Ticks + TimeSpan.TicksPerDay)
        Dim endPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerDay)

        'Dim createDateStr As String

        For i = startPt To endPt - 1
            With ApptGV.All_ApptTodoRecords(i)
                'If .DeleteFlag = 0 Then
                cnt += 1
                    a(cnt) = ApptGV.All_ApptTodoRecords(i)
                    a(cnt).Msg = Trim(a(cnt).Msg)
                'End If
                Debug.Print(.ID & " " & .ApptDateStr & " " & .dTics & " " & .Msg)
                If .dTics Mod TimeSpan.TicksPerSecond = 0 Then
                    Debug.Print(.ApptDateStr & " " & .dTics & " " & .Msg)
                    .dTics = .dTics + 1
                    .ApptDateStr = formatApptTodoDateStr(.dTics)
                    UpdateApptTodoRecord(ApptGV.All_ApptTodoRecords(i))
                End If
            End With
        Next
        ReDim Preserve a(cnt)
    End Sub
    Function getApptRecords_GivenDate_ApptTodoType(ByVal givenDate As Date)
        Dim a(1000000) As ApptTodoRecordType
        Dim cnt As Integer = 0
        Dim i As Integer

        ''''Dim startPt As Integer = biSearch_GTE_TodoCreateDates(SelectedYearMonth_LONG)
        Dim startPt As Integer = biSearch_GTE_ApptTodoRecords(givenDate.Ticks)

        'Dim endPt As Integer = biSearch_GTE_TodoCreateDates(NextMth)
        Dim endPt As Integer = biSearch_GTE_ApptTodoRecords(givenDate.Ticks + TimeSpan.TicksPerDay)

        'Dim createDateStr As String

        For i = startPt To endPt - 1
            With ApptGV.All_ApptTodoRecords(i)
                If .DeleteFlag = 0 Then
                    cnt += 1
                    a(cnt) = ApptGV.All_ApptTodoRecords(i)
                    a(cnt).Msg = Trim(a(cnt).Msg)
                End If
            End With
        Next
        ReDim Preserve a(Math.Min(ApptGV.MaxApptTodos, cnt)) 'a(3) 'a(cnt) 'a(6) 'a(cnt)

        'QuickSort_ApptTodoRecords(a, 1, cnt) ''==========add holidays NEW


        'ReDim Preserve ApptGV.distinctDates(cnt)
        'a = sortFunction1(a)
        'For i = 1 To UBound(a) '10
        '    With a(i)
        '        Debug.Print(.ID & " " & .Msg & " " & .ApptDateStr & " " & .DeleteFlag)
        '    End With

        'Next
        Return a


    End Function
    Function getApptRecords_AllFields_StartingDateEndingDate_NEW(ByVal sd As Date, ByVal ed As Date)
        Dim a(1000000) As ApptTodoRecordType
        Dim cnt As Integer = 0
        Dim i As Integer

        ''''Dim startPt As Integer = biSearch_GTE_TodoCreateDates(SelectedYearMonth_LONG)
        Dim startPt As Integer = biSearch_GTE_ApptTodoRecords(sd.Ticks)

        'Dim endPt As Integer = biSearch_GTE_TodoCreateDates(NextMth)
        Dim endPt As Integer = biSearch_GTE_ApptTodoRecords(ed.Ticks + TimeSpan.TicksPerDay)

        ''Dim createDateStr As String
        'Debug.Print(UBound(ApptGV.All_ApptTodoRecords) & " " & ApptGV.All_ApptTodoRecords_Cnt)
        'For i = 1 To UBound(ApptGV.All_ApptTodoRecords)
        '    With ApptGV.All_ApptTodoRecords(i)
        '        Debug.Print(i & " " & .ApptDateStr & " " & .Msg)
        '    End With

        'Next
        For i = startPt To endPt - 1
            With ApptGV.All_ApptTodoRecords(i)
                ' With ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)


                ''''createDateStr = Format(New Date(.dTicsOriginal), "yyyy/MM")

                'Debug.Print(Mid(.ApptDateStr, 1, 7) & " " & SelectedYearMonth)
                'Debug.Print(createDateStr & " " & SelectedYearMonth)
                ''''If i < 20 Then Debug.Print(createDateStr & " " & .Msg)



                'If Mid(.ApptDateStr, 1, 7) = SelectedYearMonth AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID
                'If createDateStr = SelectedYearMonth AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID
                If .DeleteFlag = 0 Then
                    cnt += 1


                    a(cnt) = ApptGV.All_ApptTodoRecords(i)
                    'a(cnt) = ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)

                    'Debug.Print(Len(.Msg))
                    '.Msg = Trim(.Msg)
                    a(cnt).Msg = Trim(a(cnt).Msg)
                    'Debug.Print(Len(.Msg))

                    'If cnt < 3 Then a(cnt).DeleteFlag = 1
                End If




                'If .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 And (Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate) Or (Mid(.ApptDateStr, 1, 10) = "0005/01/01") Then
                '            cnt += 1
                '            a(cnt) = ApptGV.All_ApptTodoRecords(i)

                '    'ApptGV.distinctDates(cnt) = (a(cnt).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '  a(cnt).apptDate
                '    'distinctDates(cnt) = a(cnt).dTics
                'End If
            End With
        Next
        'End If

        '==========add holidays NEW
        Dim hol() As HolidaysType

        hol = Holidays_GivenMonthYear(sd)
        For i = 1 To UBound(hol)
            cnt += 1
            With a(cnt)
                .dTics = hol(i).HolidayDate.Ticks
                .apptDate = .dTics
                .Msg = hol(i).HolidayName
                .ApptDateStr = formatApptTodoDateStr(.dTics)
            End With
            'Debug.Print(Format(d, "MMMM yyyy ") & .HolidayName & " " & .HolidayDate)
        Next

        'add birthdays
        Dim bdays() As BirthdayRecordType
        bdays = getBirthdaysforGivenMonth(sd.Month)
        For i = 1 To UBound(bdays)

            cnt += 1
            With a(cnt)
                .dTics = DateSerial(sd.Year, bdays(i).Month, bdays(i).Day).Ticks
                .apptDate = .dTics
                .Msg = "Birthday - " & Trim(bdays(i).Msg) 'hol(i).HolidayName
                .ApptDateStr = formatApptTodoDateStr(.dTics)
                .DeleteFlag = 20 '? 11/25/2019
            End With
            'Debug.Print(Format(d, "MMMM yyyy ") & .HolidayName & " " & .HolidayDate)
        Next

        '==========

        'appt_distinctDates
        ReDim Preserve a(Math.Min(ApptGV.MaxApptTodos, cnt)) 'a(3) 'a(cnt) 'a(6) 'a(cnt)

        QuickSort_ApptTodoRecords(a, 1, cnt) ''==========add holidays NEW


        'ReDim Preserve ApptGV.distinctDates(cnt)
        'a = sortFunction1(a)
        'For i = 1 To UBound(a) '10
        '    With a(i)
        '        Debug.Print(.ID & " " & .Msg & " " & .ApptDateStr & " " & .DeleteFlag)
        '    End With

        'Next
        Return a



    End Function
    Function getApptRecords_AllFields_NEW(ByVal d As Date)
        'Dim d As Date = SelectedDate() ' New Date(SelectedYearMonth_LONG) 'New DateTime(d.Year, d.Month, 1).AddDays(-1)
        'Dim dd As Date = d.AddMonths(1) 'New DateTime(d.Year, d.Month, 1).AddDays(-1)
        'Dim NextMth As Long = dd.Ticks

        ''Debug.Print(d & " " & dd) '5/1/2019 6/1/2019


        'Dim days2today As Integer = DaysToToday()

        'Debug.Print(SelectedYearMonth_LONG & " " & New Date(SelectedYearMonth_LONG))

        Dim a(1000000) As ApptTodoRecordType
        Dim cnt As Integer = 0
        Dim i As Integer

        ''''Dim startPt As Integer = biSearch_GTE_TodoCreateDates(SelectedYearMonth_LONG)
        Dim startPt As Integer = biSearch_GTE_ApptTodoRecords(d.Ticks)

        'Dim endPt As Integer = biSearch_GTE_TodoCreateDates(NextMth)
        Dim endPt As Integer = biSearch_GTE_ApptTodoRecords(d.Ticks + TimeSpan.TicksPerDay)

        'Dim createDateStr As String

        For i = startPt To endPt - 1
            With ApptGV.All_ApptTodoRecords(i)
                ' With ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)


                ''''createDateStr = Format(New Date(.dTicsOriginal), "yyyy/MM")

                'Debug.Print(Mid(.ApptDateStr, 1, 7) & " " & SelectedYearMonth)
                'Debug.Print(createDateStr & " " & SelectedYearMonth)
                ''''If i < 20 Then Debug.Print(createDateStr & " " & .Msg)



                'If Mid(.ApptDateStr, 1, 7) = SelectedYearMonth AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID
                'If createDateStr = SelectedYearMonth AndAlso (.DeleteFlag = 0 Or (.DeleteFlag = 1 And days2today - CInt(.dTicsCompleted / TimeSpan.TicksPerDay) = 0)) Then '.scUserID = ApptGV.UserID
                If .DeleteFlag = 0 Then
                    cnt += 1

                    a(cnt) = ApptGV.All_ApptTodoRecords(i)
                    'a(cnt) = ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)

                    'Debug.Print(Len(.Msg))
                    '.Msg = Trim(.Msg)
                    a(cnt).Msg = Trim(a(cnt).Msg)
                    'Debug.Print(Len(.Msg))

                    'If cnt < 3 Then a(cnt).DeleteFlag = 1
                End If




                'If .scUserID = ApptGV.UserID AndAlso .DeleteFlag = 0 AndAlso .DeleteAbsolute = 0 And (Mid(.ApptDateStr, 1, 10) >= sDate AndAlso Mid(.ApptDateStr, 1, 10) <= eDate) Or (Mid(.ApptDateStr, 1, 10) = "0005/01/01") Then
                '            cnt += 1
                '            a(cnt) = ApptGV.All_ApptTodoRecords(i)

                '    'ApptGV.distinctDates(cnt) = (a(cnt).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '  a(cnt).apptDate
                '    'distinctDates(cnt) = a(cnt).dTics
                'End If
            End With
        Next
        'End If

        'new as of 11/21/2019
        Dim bdays() As StringPointerType = getBirthdaysForGivenDate(d)
        Dim numBdays As Integer = UBound(bdays)
        For i = 1 To numBdays
            cnt += 1
            With a(cnt)
                .Msg = "Birthday - " & Trim(bdays(i).str)
                .dTics = d.Ticks
                .apptDate = .dTics
                .ApptDateStr = formatApptTodoDateStr(.dTics)
                .DeleteFlag = 20
                .ID = bdays(i).ptr
            End With
        Next


        'appt_distinctDates
        ReDim Preserve a(Math.Min(ApptGV.MaxApptTodos, cnt)) 'a(cnt) 'a(6) 'a(cnt)


        'ReDim Preserve ApptGV.distinctDates(cnt)
        'a = sortFunction1(a)
        'For i = 1 To UBound(a) '10
        '    With a(i)
        '        Debug.Print(.ID & " " & .Msg & " " & .ApptDateStr & " " & .DeleteFlag)
        '    End With

        'Next
        Return a


    End Function
    Sub test_findApptRecordsNEW()
        ApptGV.UserID = 1
        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds

        Dim find As String
        find = "screw" '"t" '"Haf"
        Dim r() As ApptTodoRecordType
        startTime()

        'r = findApptRecordsNEW(find) '< .1 seconds
        r = findApptRecordsNEW(find, True) '< .1 seconds
        endTime()

        Dim i As Integer
        Dim nRecs As Integer = UBound(r)
        For i = 1 To 3 'nRecs
            Debug.Print(r(i).Msg & " " & r(i).ApptDateStr & " " & r(i).dTics)
        Next
    End Sub
    Function findApptRecordsNEW(ByVal find As String, Optional ByVal includeTodos As Boolean = False) As ApptTodoRecordType()

        'includeTodos = False 'testing


        Dim r(0) As ApptTodoRecordType
        If find = "" Then
            Return r
        End If

        ReDim r(1000000)
        Dim i As Integer
        Dim cnt As Integer = 0
        Dim sPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerDay)
        Dim xfind As String = find.ToLower



        ''''Dim Q As String
        'Q = "select * from appt where apptdate<>'1/1/0005' and charindex('t',msg)>0 order by dtics"
        'select * from appt where apptdate<>'1/1/0005' and charindex('t',msg)>0 order by dtics
        'Select * From appt Where apptdate ='1/1/0005' and deleteflag=0 and charindex('t',msg)>0 order by dtics
        ''''Q = "declare @find varchar(100)='" & find & "'"


        If includeTodos = False Then
            'Q &= " Select * from appt where apptdate<>'1/1/0005' and charindex(@find,msg)>0 order by dtics"
            sPt = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerDay)

            'For i = sPt To ApptGV.All_ApptTodoRecords_Cnt
            For i = ApptGV.All_ApptTodoRecords_Cnt To sPt Step -1 'go from newest (closest date from NOW) to oldest record
                With ApptGV.All_ApptTodoRecords(i)
                    If InStr(.Msg.ToLower, xfind) Then
                        cnt += 1
                        r(cnt) = ApptGV.All_ApptTodoRecords(i)
                    End If
                End With

            Next

        Else
            'Q &= " Select * from appt where deleteflag=0 and charindex(@find,msg)>0 order by substring(apptdatestr,1,1) desc, dtics"


            sPt = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
            For i = ApptGV.All_ApptTodoRecords_Cnt To sPt Step -1 'go from newest (closest date from NOW) to oldest record
                With ApptGV.All_ApptTodoRecords(i)
                    If InStr(.Msg.ToLower, xfind) Then
                        cnt += 1
                        r(cnt) = ApptGV.All_ApptTodoRecords(i)
                    End If
                End With

            Next
        End If

        'Debug.Print(cnt)
        ReDim Preserve r(Math.Min(cnt, 100)) 'show only lastest 100 records
        'r = getApptRecords_AllFields(Q) ', useTODO_DB)
        Return r

    End Function
    Function getApptDatesTobeBolded()
        'from getApptDatesWithinCalendarMonth -same as but put in this module so it is accessible from addAppt form

        'startTime()

        If ApptGV.All_ApptTodoRecords_Cnt = 0 Then
            Dim r(0) As Date
            Return r
        End If

        Dim a(1000000) As Date
        Dim cnt As Integer = -1 ' 0
        Dim i As Integer
        Dim OldAppt As Long = 0
        Dim MaxDates As Integer = 366

        Dim LowPoint As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerDay) 'location of first non todo record (appt rec)


        For i = ApptGV.All_ApptTodoRecords_Cnt To LowPoint Step -1
            With ApptGV.All_ApptTodoRecords(i)
                If cnt = -1 Then
                    cnt += 1
                    a(cnt) = New Date(.apptDate)
                    OldAppt = .apptDate
                Else
                    If .apptDate <> OldAppt Then
                        cnt += 1
                        a(cnt) = New Date(.apptDate)
                        OldAppt = .apptDate
                    End If
                End If
                If cnt >= MaxDates Then Exit For
            End With
        Next
        If cnt < MaxDates Then
            For i = 1 To UBound(ApptGV.HolidayDates)
                cnt += 1
                a(cnt) = ApptGV.HolidayDates(i)
            Next
        End If


        If cnt = -1 Then
            ReDim a(0)
            Return a
        End If
        ReDim Preserve a(Math.Min(cnt, MaxDates)) ' highlight 1 year of dates - from highest date to lowest (trimming off those that exceed 366 dates
        ' Debug.Print(a.Count)
        Return a


    End Function
    Sub setBoldDatesNEWGeneral(ByRef mthCal As MonthCalendar)
        'from setBoldDatesNEW()
        'appt_distinctDates_for_MonthCalendar1 = testBold() 'getApptDatesWithinCalendarMonth() 'NEW .03 seconds
        Dim appt_distinctDates_for_mthCal() As Date = getApptDatesTobeBolded() 'getApptDatesWithinCalendarMonth() 'NEW .03 seconds
        'endTime()

        'updateBoldedDates_MonthCalendarForPrinting_Appt(MonthCalendar1, appt_distinctDates_for_MonthCalendar1)
        'next 3 lines replaces above one line
        mthCal.RemoveAllBoldedDates()
        mthCal.UpdateBoldedDates()
        mthCal.BoldedDates = appt_distinctDates_for_mthCal 'appt_distinctDates_for_MonthCalendar1


        'updateBoldedDates_MonthCalendarForPrinting_Appt(MonthCalendarForPrinting, appt_distinctDates_for_MonthCalendar1)
    End Sub
    Function getDeletedTodos(Optional ByVal ReverseDateOrder As Boolean = False) As ApptTodoRecordType()
        Dim cnt As Integer = 0
        Dim i As Integer

        Dim startPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
        Dim endPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerDay)
        Dim nRecs As Integer = endPt - startPt

        Dim a(nRecs) As ApptTodoRecordType
        For i = startPt To endPt - 1
            'For i = endPt - 1 To startPt Step -1

            'With ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)
            With ApptGV.All_ApptTodoRecords(i)
                If .DeleteFlag = 1 Then
                    'Debug.Print(cnt & " " & New Date(ApptGV.TodoCreateDate(i).L) & " " & " " & New Date(.dTicsOriginal) & " " & .Msg & " " & .DeleteFlag & " " & New Date(.dTicsCompleted))
                    cnt += 1
                    'a(cnt) = ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i)
                    a(cnt) = ApptGV.All_ApptTodoRecords(i)
                    'Debug.Print(Trim(a(cnt).Msg))
                    a(cnt).Msg = Trim(a(cnt).Msg)
                End If
            End With
        Next
        'ReDim Preserve a(Math.Min(ApptGV.MaxApptTodos, cnt))
        ReDim Preserve a(cnt)
        'QuickSort_ApptTodoRecords(a, 1, UBound(a))
        If ReverseDateOrder = False Then
            QuickSort_ApptTodoRecords_bydTicsCompleted(a, 1, UBound(a))
        Else
            QuickSort_ApptTodoRecords_descending_bydTicsCompleted(a, 1, UBound(a))
        End If

        Return a
    End Function
End Module

