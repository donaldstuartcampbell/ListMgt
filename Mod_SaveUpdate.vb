Option Explicit On

Module Mod_SaveUpdate
    'Sub AddApptTodoRecord(ByVal rec As ApptTodoRecordType) ', Optional ByVal useTODO_DB As Boolean = False) 'use for recurring records  'new 9/5/2015
    Function initBirthdayRecords() As BirthdayRecordType
        Dim rec As BirthdayRecordType
        With rec
            .DateStr = ""
            .Day = 0
            .DeleteFlag = 0
            .Desc = ""
            .DescriptionNumber = 0
            .dTics = 0
            .dTicsOriginal = 0
            .ID = 0
            .Month = 0
            .Msg = ""
        End With
        Return rec
    End Function
    Sub updateBirthdayRecord(ByRef rec As BirthdayRecordType)
        Dim ff As Integer = FreeFile()

        Dim filename As String = xBuildFullFileName(file_Birthday.FileName) 'file_ListMgt.FileName
        FileOpen(ff, filename, OpenMode.Random, , , file_Birthday.FileLen)

        Dim xID As Integer = rec.ID
        FilePut(ff, rec, xID)
        FileClose(ff)
    End Sub
    Sub addBirthdayRecord(ByRef rec As BirthdayRecordType)
        '===========NEW
        Dim ff As Integer = FreeFile()

        'Dim filename As String = xBuildFullFileName(f_Appt.FileName)
        'FileOpen(ff, filename, OpenMode.Random, , , f_Appt.FileLen)

        Dim filename As String = xBuildFullFileName(file_Birthday.FileName) 'file_ListMgt.FileName
        FileOpen(ff, filename, OpenMode.Random, , , file_Birthday.FileLen)


        'Dim xCount As Integer = LOF(ff) / f_Appt.FileLen + 1
        Dim xCount As Integer = LOF(ff) / file_Birthday.FileLen + 1



        '[ID],[dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag],[DeleteAbsolute],[scUserID],[apptDateStr],[apptDate],[dTicsOriginal],[dTicsCompleted]

        'Dim xMsg As String = rec.Msg
        'SQLreplace(xMsg)

        With rec
            .ID = xCount

            '.DateStr 'ALREADY SET

            .Day = CInt(Mid(.DateStr, 4, 2))
            .DeleteFlag = 0
            .Desc = ""

            If .DescriptionNumber = 0 Then
                .DescriptionNumber = 1
            End If
            '.DescriptionNumber 'ALREADY SET

            .dTics = 0
            .dTicsOriginal = Date.Now.Ticks + xCount
            .Month = CInt(Mid(.DateStr, 1, 2))

            'Debug.Print(.ID & " " & .Msg & " " & formatApptDateStr(.dTics) & " " & .DateStr)
            Debug.Print(String.Format("{0:0000}", .ID) & " " & Trim(.DateStr) & " desc#=" & .DescriptionNumber & " del=" & .DeleteFlag & " day=" & String.Format("{0:00}", .Day) & " mth=" & String.Format("{0:00}", .Month) & " " & Trim(.Msg))
            '.Msg 'ALREADY SET

            '==============================

            '.dTics = L + cnt
            ''.Msg 'already set
            '.AcctNum = 0
            '.StrikeOut = 0
            '.DeleteFlag = 0
            '.DeleteAbsolute = 0
            '.scUserID = 1
            '.ApptDateStr = formatApptTodoDateStr(.dTics)
            '.apptDate = createPureDate_Ticks_FromTicks(.dTics)
            '.dTicsOriginal = Date.Now.Ticks

            'If .dTicsOriginal <= ApptGV.Highest_dTicsOriginal Then
            '    .dTicsOriginal = ApptGV.Highest_dTicsOriginal + 1
            'End If
            'ApptGV.Highest_dTicsOriginal = .dTicsOriginal

            '.dTicsCompleted = 0
        End With
        FilePut(ff, rec, xCount)
        FileClose(ff)

        '===========

    End Sub
    Sub AddApptTodoRecord(ByRef rec As ApptTodoRecordType, Optional ExitAfterAdd As Boolean = False) ', Optional ByVal useTODO_DB As Boolean = False) 'use for recurring records  'new 9/5/2015
        'change rec to ByRef

        'pass:
        '1. Msg
        '2. dtics

        Dim L As Long
        Dim cnt As Integer

        'Dim xMsg As String = rec.Msg
        'SQLreplace(xMsg)

        'If noDate = True Then
        L = (rec.dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute 'New Date(5, 1, 1, 0, 0, 0).Ticks
        cnt = getNewHighCountFromApptTodoRecords(L)

        '==
        'Dim L As Long = Date.Now.Ticks

        '===========NEW
        Dim ff As Integer = FreeFile()

        'Dim filename As String = xBuildFullFileName(f_Appt.FileName)
        'FileOpen(ff, filename, OpenMode.Random, , , f_Appt.FileLen)

        Dim filename As String = xBuildFullFileName(file_ListMgt.FileName) 'file_ListMgt.FileName
        FileOpen(ff, filename, OpenMode.Random, , , file_ListMgt.FileLen)


        'Dim xCount As Integer = LOF(ff) / f_Appt.FileLen + 1
        Dim xCount As Integer = LOF(ff) / file_ListMgt.FileLen + 1


        If xCount <> ApptGV.All_ApptTodoRecords_Cnt + 1 Then
            MsgBox("in: AddApptTodoRecord")
            'error
            End
        End If

        '[ID],[dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag],[DeleteAbsolute],[scUserID],[apptDateStr],[apptDate],[dTicsOriginal],[dTicsCompleted]

        'Dim xMsg As String = rec.Msg
        'SQLreplace(xMsg)

        With rec
            .ID = xCount
            .dTics = L + cnt
            '.Msg 'already set
            .AcctNum = 0
            .StrikeOut = 0
            .DeleteFlag = 0
            .DeleteAbsolute = 0
            .scUserID = 1
            .ApptDateStr = formatApptTodoDateStr(.dTics)
            .apptDate = createPureDate_Ticks_FromTicks(.dTics)
            .dTicsOriginal = Date.Now.Ticks

            If .dTicsOriginal <= ApptGV.Highest_dTicsOriginal Then
                .dTicsOriginal = ApptGV.Highest_dTicsOriginal + 1
            End If
            ApptGV.Highest_dTicsOriginal = .dTicsOriginal

            .dTicsCompleted = 0
        End With
        FilePut(ff, rec, xCount)
        FileClose(ff)

        '===========
        ApptGV.All_ApptTodoRecords_Cnt += 1
        ReDim Preserve ApptGV.All_ApptTodoRecords(xCount)
        ApptGV.All_ApptTodoRecords(xCount) = rec

        If ExitAfterAdd = True Then Exit Sub


        QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt)
        set_CreateDateArray() '0.007 seconds






    End Sub
    ''Sub AddTodoRecord(ByVal rec As ApptTodoRecordType, Optional ByVal noDate As Boolean = False, Optional xSelectedDate As Date = Nothing)
    'Sub AddTodoRecord(ByVal rec As ApptTodoRecordType) ', Optional ByVal TodoRec As Boolean = False, Optional xSelectedDate As Date = Nothing)
    '    'Public Sub AddTodoRecord(ByVal rec As ApptTodoRecordType, Optional TodoType As Boolean = False, Optional xSelectedDate As Date = Nothing) '(ByVal rec As ApptTodoRecordType)

    '    'pass:
    '    '1. Msg
    '    '2. dtics

    '    Dim L As Long
    '    Dim cnt As Integer

    '    'Dim xMsg As String = rec.Msg
    '    'SQLreplace(xMsg)

    '    'If noDate = True Then
    '    L = (rec.dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute 'New Date(5, 1, 1, 0, 0, 0).Ticks
    '    cnt = getNewHighCountFromApptTodoRecords(L)

    '    'If L = ApptGV.todoTicks Then
    '    '    'L = New Date(5, 1, 1, 0, 0, 0).Ticks
    '    '    'createPureDateFromTicks(L)
    '    '    'cnt = QretInt("select isnull(max(dtics) % 10000000,0) from [dbo].[scAppt] where apptdate='0005/01/01' and scUserID =0")

    '    '    'cnt = QInt("select cast(isnull(max(dtics) % 10000000,0) as int) from [dbo].[Appt] where apptdate='0005/01/01' and scUserID =" & ApptGV.UserID)
    '    '    'cnt = QInt("select cast(isnull(max(dtics % 10000000),0) as int) from [dbo].[Appt] where apptdate='0005/01/01' and scUserID =" & ApptGV.UserID) 'new 5/2/2019

    '    '    cnt = getNewHighCountFromApptTodoRecords(L)
    '    'Else

    '    '    L = xSelectedDate.Ticks
    '    '    'createPureDateFromTicks(L)
    '    '    'cnt = QretInt("select isnull(max(dtics) % 10000000,0) from [dbo].[scAppt] where apptdate='" & SelectedDate & "' and scUserID =0")

    '    '    'cnt = QInt("select cast(isnull(max(dtics) % 10000000,0) as int) from [dbo].[Appt] where apptdate='" & SelectedDate & "' and scUserID =0")

    '    '    'cnt = QInt("select cast(isnull(max(dtics) % 10000000,0) as int) from [dbo].[Appt] where apptdate='" & SelectedDate & "' and scUserID =" & ApptGV.UserID) 'changed 5/1/2019
    '    '    cnt = QInt("select cast(isnull(max(dtics % 10000000),0) as int) from [dbo].[Appt] where apptdate='" & xSelectedDate & "' and scUserID =" & ApptGV.UserID) 'new 5/2/2019

    '    'End If
    '    'L = L + cnt + 1
    '    'create one
    '    'Debug.Print(TimeSpan.TicksPerDay)

    '    With rec
    '        .AcctNum = 0
    '        .DeleteFlag = 0
    '        .dTics = L + cnt

    '        'don't dod this here as it's done in the Add function
    '        'Dim xmsg As String = s
    '        'SQLreplace(xmsg)
    '        '.msg = xmsg 's

    '        '.Msg = xMsg
    '        .StrikeOut = 0
    '        .dTicsOriginal = Date.Now.Ticks '.dTics
    '        .dTicsCompleted = 0
    '        .scUserID = ApptGV.UserID
    '    End With
    '    AddApptTodoRecord(rec)

    '    'AddApptRecord(rec)

    'End Sub
    Sub test_getHighCountFromApptTodoRecords_TodoCnt()
        Set_ApptTodoRecordType_UsingReader()

        Dim uniqueDateTimeRecords() As Long 'ApptTodoRecordType
        startTime()
        uniqueDateTimeRecords = getAllUniqueDateTimeRecords()
        endTime()

        Dim nItems As Integer = UBound(uniqueDateTimeRecords)
        Debug.Print(nItems)

        'uniqueDateTimeRecords(0) = ApptGV.todoTicks
        uniqueDateTimeRecords(0) = #3/19/2019#.Ticks

        Dim i As Integer
        Dim x As Integer = 0
        startTime()

        For i = 0 To nItems
            x = getNewHighCountFromApptTodoRecords(uniqueDateTimeRecords(i))
            'Debug.Print(i & " " & New Date(uniqueDateTimeRecords(i)) & " " & getNewHighCountFromApptTodoRecords(uniqueDateTimeRecords(i)))
        Next
        endTime()

    End Sub
    Function getAllUniqueDateTimeRecords() As Long()
        Dim sPoint As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
        Dim ePoint As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerDay) 'TimeSpan.TicksPerMinute)
        Dim nItems As Integer = ePoint - sPoint
        Dim cnt As Integer = 0
        'ReDim ApptGV.TodoCreateDate(nItems)
        Dim L(1000000) As Long
        Dim LL As Long
        Dim i As Integer
        'For i = sPoint To ePoint - 1
        For i = ePoint To ApptGV.All_ApptTodoRecords_Cnt
            With ApptGV.All_ApptTodoRecords(i)
                LL = (.dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute
                If LL <> L(cnt) Then
                    cnt += 1
                    L(cnt) = LL
                End If
            End With

            'cnt += 1
            'With ApptGV.TodoCreateDate(cnt)
            '    .L = (ApptGV.All_ApptTodoRecords(i).dTicsOriginal \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay + (ApptGV.All_ApptTodoRecords(i).dTics - ApptGV.All_ApptTodoRecords(i).apptDate)
            '    .i = i 'position in array
            'End With
            'Debug.Print(cnt & " " & formatApptTodoDateStr(ApptGV.TodoCreateDate(i).L) & " " & ApptGV.TodoCreateDate(i).i & " " & ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i).DeleteFlag & " " & ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i).Msg)
        Next
        ReDim Preserve L(cnt)
        Return L

        'QuickSort_LongInteger_Long(ApptGV.TodoCreateDate, 1, cnt)
    End Function
    Function getNewHighCountFromApptTodoRecords(xsearch As Long) As Integer
        Dim search As Long = (xsearch \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute
        Dim x As Long = 0
        Dim RecCnt As Integer = 0
        'xSearch = 1262304000000000 + TimeSpan.TicksPerMinute 'or TicksPerSecond?' DateSerial(5, 1, 2).Ticks

        Dim LocInFile As Integer
        Dim L As Long = search + TimeSpan.TicksPerMinute
        LocInFile = biSearch_GTE_ApptTodoRecords(L) - 1


        If LocInFile > 0 Then
            x = (ApptGV.All_ApptTodoRecords(LocInFile).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute
            If x = search Then
                RecCnt = ApptGV.All_ApptTodoRecords(LocInFile).dTics Mod 10000000 + 1
            Else
                RecCnt = 1
            End If
        Else
            RecCnt = 1 'new 11/25/2019
        End If


        Return RecCnt

    End Function

    Function getHighCountFromApptTodoRecords_TodoCnt(Optional search As Long = 0) As Integer
        Dim xSearch As Long


        'setAll_ApptTodoRecords() '0.2 seconds
        'QuickSort_LongInteger_Long(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt))' UBound(ApptGV.AllAppts_LongInteger))
        'QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt)

        Dim RecCnt As Integer = 0
        If search = 0 Then
            xSearch = 1262304000000000 + TimeSpan.TicksPerMinute 'or TicksPerSecond?' DateSerial(5, 1, 2).Ticks
        Else
            xSearch = search
        End If

        'Debug.Print(New Date(xSearch))

        Dim LocInFile As Integer
        LocInFile = biSearch_GTE_ApptTodoRecords(xSearch) - 1

        ''================test
        'Dim xx As Date = CDate(#1/1/0005#) '1262304000000000
        'Debug.Print(xx.Ticks & " " & New Date(xx.Ticks))

        'xx = #4/27/2019#
        'Debug.Print(xx.Ticks & " " & New Date(xx.Ticks))
        ''636919200000000000 4/27/2019
        ''636919524068563400 4/27/2019 9:00:06
        ''      324068563400

        ''Debug.Print(DateSerial(2019, 3, 18).Ticks)
        ''636884640000000000
        ''632401344000000000

        ''1262304000000000
        ''1262304 '138561302
        'Dim i As Integer
        'For i = LocInFile + 10000 To LocInFile + 10010 'LocInFile To LocInFile + 1000 ' ApptGV.All_ApptTodoRecords_Cnt
        '    With ApptGV.All_ApptTodoRecords(i)
        '        'Debug.Print(temp & " " & .ID & " " & temp & "=cnt " & .ApptDateStr & " " & .dTics & " " & New Date(.dTics) & " " & .Msg)
        '        'Debug.Print(ApptGV.All_ApptTodoRecords(LocInFile).dTics Mod 10000000 & " ----" & formatApptTodoDateStr(.dTics) & "--" & .ID & " " & .ApptDateStr & " " & .dTics & " " & New Date(.dTics) & " " & .Msg)
        '        Debug.Print(ApptGV.All_ApptTodoRecords(i).dTics Mod TimeSpan.TicksPerDay & " ----" & formatApptTodoDateStr(.dTics) & "--" & .ID & " " & .ApptDateStr & " " & .dTics & " " & New Date(.dTics) & " " & .Msg)
        '    End With

        'Next

        ''================end test

        'Dim cnt As Integer = 0
        If LocInFile > 0 Then
            'If Mid(ApptGV.All_ApptTodoRecords(LocInFile).ApptDateStr, 1, 10) = "0005/01/01" Then
            'If (xSearch - TimeSpan.TicksPerMinute) = ApptGV.All_ApptTodoRecords(LocInFile).apptDate Then
            If (xSearch - TimeSpan.TicksPerMinute) = (ApptGV.All_ApptTodoRecords(LocInFile).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute Then
                RecCnt = ApptGV.All_ApptTodoRecords(LocInFile).dTics Mod 10000000 'Then '8561302
            End If
        End If


        Return RecCnt



    End Function
    Sub UpdateApptTodoRecord(ByVal rec As ApptTodoRecordType)
        Dim xLocation As Integer = rec.ID

        Dim ff As Integer = FreeFile()
        Dim filename As String = xBuildFullFileName(file_ListMgt.FileName)
        FileOpen(ff, filename, OpenMode.Random, , , file_ListMgt.FileLen)

        FilePut(ff, rec, xLocation)
        FileClose(ff)

        'new
        'QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))
        'set_AllAppts_LongInteger()
    End Sub
    Function getFirstOfMonthTicks(ByVal xx As Long) As Long
        Dim x As Long = xx
        Dim y As Integer

        '==
        x = (x \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay
        y = New Date(x).Day - 1
        x = x - y * TimeSpan.TicksPerDay
        Return x

    End Function

    Sub Test_fixTodoIndex(Optional ByVal do_set As Boolean = True)
        Dim OK As Boolean = True

        If do_set = True Then
            Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
        End If


        'startTime()
        Dim startoftodosPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
        Dim endoftodosPt As Integer = biSearch_GTE_ApptTodoRecords(#0005/2/1#.Ticks) - 1
        Dim i As Integer
        Dim cnt As Integer = 0

        'If ApptGV.All_ApptTodoRecords(startoftodosPt).dTics Mod TimeSpan.TicksPerSecond <> 1 Then
        '    MsgBox("test failed - 1st rec should be 1")
        'End If
        '==


        'For i = startoftodosPt To startoftodosPt + 10
        '    Debug.Print(ApptGV.All_ApptTodoRecords(i).dTics)
        'Next

        '==0.001 seconds
        For i = startoftodosPt To endoftodosPt - 1

            cnt += 1
            If ApptGV.All_ApptTodoRecords(1).dTics + cnt <> ApptGV.All_ApptTodoRecords(i + 1).dTics Then
                'Debug.Print(ApptGV.All_ApptTodoRecords(i).dTics + cnt & " " & ApptGV.All_ApptTodoRecords(i + 1).dTics)
                ' i = i
                'MsgBox("test failed")
                OK = False
                MsgBox("failed")
                Exit For
            End If
        Next
        'endTime()
        'Debug.Print("OK" & " cnt=" & cnt)
        'MsgBox("OK" & " cnt=" & cnt)
        If OK = False Then
            fixTodoIndex()
        End If

    End Sub
    Sub fixTodoIndex() 'about 20 seconds (17.8 for 24000 records) more than 1,000 records per second
        startTime() '1.1239967 seconds by putting the fileput inside this routine vrs calling "update" routine!
        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
        'set_CreateDateArray() '0.007 seconds


        Dim startoftodosPt As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
        Dim endoftodosPt As Integer = biSearch_GTE_ApptTodoRecords(#0005/2/2#.Ticks) - 1
        Dim i As Integer
        Dim cnt As Integer = 0

        '==
        Dim xLocation As Integer ' = rec.ID

        Dim ff As Integer = FreeFile()
        Dim filename As String = xBuildFullFileName(file_ListMgt.FileName)
        FileOpen(ff, filename, OpenMode.Random, , , file_ListMgt.FileLen)



        '==
        For i = startoftodosPt To endoftodosPt
            With ApptGV.All_ApptTodoRecords(i)
                'Debug.Print(.dTics & " " & .ID & " id  " & .ApptDateStr & " " & .DeleteFlag & " " & .Msg)
                cnt += 1
                'If ApptGV.todoTicks + cnt <> .dTics Then
                '    i = i
                'End If
                .dTics = ApptGV.todoTicks + cnt
                .ApptDateStr = formatApptDateStr(.dTics)

                xLocation = .ID
            End With
            'UpdateApptTodoRecord(ApptGV.All_ApptTodoRecords(i))
            FilePut(ff, ApptGV.All_ApptTodoRecords(i), xLocation)
        Next
        FileClose(ff)

        'With ApptGV.All_ApptTodoRecords(i - 1)
        '    Debug.Print(.dTics & " " & .ID & " id  " & .ApptDateStr & " " & .DeleteFlag & " " & .Msg)
        'End With
        endTime()

        'now reset arrays
        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
        set_CreateDateArray() '0.007 seconds

    End Sub
    Sub update_RangeOfApptTodoRecords(ByVal startPT As Integer, ByVal endPT As Integer, ByRef xApptTodoRecords() As ApptTodoRecordType, Optional ByVal doTesting As Boolean = False)
        Dim i As Integer

        Dim ff As Integer = FreeFile()
        Dim filename As String = xBuildFullFileName(file_ListMgt.FileName)
        FileOpen(ff, filename, OpenMode.Random, , , file_ListMgt.FileLen)

        Dim xLocation As Integer
        For i = startPT To endPT
            'xLocation = ApptGV.All_ApptTodoRecords(i).ID
            'FilePut(ff, ApptGV.All_ApptTodoRecords(i), xLocation)

            xLocation = xApptTodoRecords(i).ID
            FilePut(ff, xApptTodoRecords(i), xLocation)
        Next

        FileClose(ff)
        If doTesting = True Then
            Test_fixTodoIndex() '(True)
        End If

    End Sub
    Function getXorder(ByRef xxtodoRecords() As ApptTodoRecordType, ByRef xxNewOrder() As ApptTodoRecordType) As Integer()
        Dim nRecs As Integer = UBound(xxtodoRecords)
        Dim xOrder(nRecs) As Integer
        Dim i As Integer
        Dim j As Integer
        For i = 1 To nRecs

            For j = 1 To nRecs
                With xxtodoRecords(j)
                    If .dTics = xxNewOrder(i).dTics Then
                        xOrder(i) = j
                        Exit For
                    End If
                End With
            Next
        Next
        Return xOrder
    End Function

#Region "PutTodoOnTop"
    'Private Sub DoUpdate_ApptTodo()
    'Sub PutTodoOnTop(ByVal xSelectedYearMonthLong As Long)
    '    Dim i As Integer
    '    Dim xTodos() As ApptTodoRecordType = getApptRecords_AllFields_ApptTodo_NEW(xSelectedYearMonthLong)
    '    Dim nRecs As Integer = UBound(xTodos)
    '    'For i = 1 To nRecs
    '    '    With xTodos(i)
    '    '        Debug.Print(.dTics & " " & .ID & " id  " & .ApptDateStr & " " & .DeleteFlag & " " & .Msg)
    '    '    End With

    '    'Next
    '    For i = 1 To 30000
    '        With ApptGV.All_ApptTodoRecords(i)
    '            If Mid(.ApptDateStr, 23) > "24300" AndAlso Mid(.ApptDateStr, 23) < "24400" Then
    '                Debug.Print(.dTics & " " & .ID & " id  " & .ApptDateStr & " " & .DeleteFlag & " " & .Msg)
    '            End If
    '        End With
    '    Next


    '    'setNewOrder()

    '    'Dim x As Long = getFirstOfMonthTicks(_SelDate.Ticks)
    '    ''TodoRecs_ApptTodo = getApptRecords_AllFields_ApptTodo_NEW(x, True) '_SelDate.Ticks, True)

    '    'Dim todoRecords() As ApptTodoRecordType = getApptRecords_AllFields_ApptTodo_NEW(x, True) '_SelDate.Ticks, True)
    '    ''TodoRecs_ApptTodo = getApptRecords_AllFields_ApptTodo_NEW(_SelDate.Ticks)
    '    'Dim activeDtics(UBound(todoRecords)) As Long '= QgetLongs(Q)
    '    'For i = 1 To UBound(todoRecords)
    '    '    activeDtics(i) = todoRecords(i).dTics
    '    'Next


    '    ''Dim activeDtics() As Long = QgetLongs(Q)

    '    'If UBound(activeDtics) <> UBound(TodoRecsNewOrder_ApptTodo) Then
    '    '    Debug.Print(UBound(activeDtics) & " " & UBound(TodoRecsNewOrder_ApptTodo))
    '    '    MsgBox("In - DoUpdate_ApptTodo - UBound(activeDtics) <> UBound(TodoRecsNewOrder_ApptTodo) - will END")
    '    '    End
    '    'End If
    '    'Dim cnt As Integer = 0


    '    'For i = 1 To UBound(TodoRecsNewOrder_ApptTodo)
    '    '    With TodoRecsNewOrder_ApptTodo(i)
    '    '        cnt += 1
    '    '        .dTics = activeDtics(cnt) '_SelDate.Ticks + 23 * TimeSpan.TicksPerHour + 59 * TimeSpan.TicksPerMinute + i * TimeSpan.TicksPerSecond
    '    '        .ApptDateStr = formatApptDateStr(.dTics)


    '    '        UpdateApptTodoRecord(TodoRecsNewOrder_ApptTodo(i))


    '    '    End With
    '    'Next
    '    'endTime() ' < .01 seconds


    'End Sub
#End Region
End Module
