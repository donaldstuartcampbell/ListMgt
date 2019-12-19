Option Explicit On
Module mDateFunc
    Structure HolidaysType
        Dim HolidayDate As Date
        Dim HolidayName As String
    End Structure
    Sub text_datefunc()
        Dim i As Integer
        Dim holidays() As HolidaysType
        'holidays = Holidays_GivenYear(2014)
        'holidays = Holidays_Next12Months_fromGivenDate(Today)



        holidays = Holidays_Next12Months_fromGivenDate(DateSerial(2014, 11, 27))
        For i = 1 To UBound(holidays)
            Debug.Print(holidays(i).HolidayName & " " & holidays(i).HolidayDate)
        Next
        End


        Dim L As List(Of Date)
        L = getHolidayList(2014)


        For i = 0 To L.Count - 1
            Debug.Print(L(i).ToString)
        Next
        End

        For i = 2013 To 2017
            'Debug.Print(thanksgivingDate(i))
            Debug.Print(CType(F_easter_date(i), String))
        Next

        End
        Dim d, dd As Date
        d = New Date(2014, 11, 1)
        'dd = GetFirstSundayInMonth(d)
        'Debug.Print(dd)

        dd = GetFirstDOWInMonth(d, DayOfWeek.Thursday)
        Debug.Print(CType(dd, String))
        End
    End Sub
    Sub dateStuff()
        'Dim DayOfWeek As FirstDayOfWeek = FirstDayOfWeek.Sunday
        Dim d As Date
        d = New DateTime(2014, 8, 4) 'this is a monday
        Debug.Print(CType(d.DayOfWeek, String)) '1  0 1 sun mon
        Debug.Print(CType(Weekday(d), String))  '2  1 2 sun mon
        End

    End Sub
    Public Function GetFirstSundayInMonth(ByVal aDate As DateTime) As DateTime
        Dim day1 As DateTime = New DateTime(aDate.Year, aDate.Month, 1) 'first day of month   
        If day1.DayOfWeek = DayOfWeek.Sunday Then 'if sunday return   
            Return day1
        Else 'calc first sunday   
            'DayOfWeek enum   
            '0 = Sunday   
            '1 = Monday   
            '2 = Tuesday   
            '3 = Wednesday   
            '4 = Thursday   
            '5 = Friday   
            '6 = Saturday   
            Return day1.AddDays(7 - day1.DayOfWeek)
        End If
    End Function
    Public Function get_1stDayOfAllWeeksOfGivenYear_Ticks(ByVal xYear As Integer) As Long() '7/22/2016
        Dim d As Date = New Date(xYear, 1, 1)
        d = GetFirstSundayInMonth(d) 'GetFirstDOWofGivenYear(xYear)
        Dim wkDays(53) As Long
        'Dim cnt As Integer = 0
        Dim i As Integer
        Dim daysToAdd As Long = TimeSpan.TicksPerDay * 7
        Dim x As Long = d.Ticks
        For i = 1 To 53
            wkDays(i) = x
            x += daysToAdd
        Next
        'Do Until d.Year > xYear
        '    cnt += 1
        '    wkDays(cnt) = d.Ticks '.Day
        '    d = DateAdd(DateInterval.Day, 7, d)
        'Loop
        'ReDim Preserve wkDays(cnt)
        Return wkDays
    End Function
    'Public Function GetFirstDOWofGivenYear(ByVal xYear As Integer) As DateTime
    '    Dim d As Date = New Date(xYear, 1, 1)
    '    Return d
    'End Function
    Public Function GetFirstDOWInMonth(ByVal aDate As DateTime, ByVal whDayOWeek As DayOfWeek) As DateTime

        'returns the date of the first Sun - Sat(whDayOWeek) of a month.   

        'aDate = a date you wish to know the first whDayOWeek of.   
        'whDayOWeek = DayOfWeek.   
        '   Sunday   
        '   Monday   
        '   Tuesday   
        '   Wednesday   
        '   Thursday   
        '   Friday   
        '   Saturday   

        'create a date that is the first day of the month / year in aDate   
        Dim dayFirst As DateTime = New DateTime(aDate.Year, aDate.Month, 1) 'first day of month   

        If whDayOWeek < dayFirst.DayOfWeek Then dayFirst = dayFirst.AddDays(7) 'so we don't get the last whDayOWeek of prev month   

        dayFirst = dayFirst.AddDays(whDayOWeek - dayFirst.DayOfWeek) 'calculate   
        Return dayFirst
    End Function
    Function thanksgivingDate(ByVal year As Integer) As Date
        Dim d As Date = New Date(year, 11, 1)
        Return DateAdd(DateInterval.Day, 3 * 7, GetFirstDOWInMonth(d, DayOfWeek.Thursday))
    End Function

    Function F_easter_date(ByVal Year_of_easter As Integer) As DateTime
        Dim y As Integer = Year_of_easter
        Dim d As Integer = (((255 - 11 * (y Mod 19)) - 21) Mod 30) + 21
        Dim easter_date = New DateTime(y, 3, 1)
        easter_date = easter_date.AddDays(+d + CInt((d > 48)) + 6 - ((y + y \ 4 + d + CInt((d > 48)) + 1) Mod 7))
        Return (easter_date)
    End Function
    Function F_GoodFriday_date(ByVal Year As Integer) As DateTime
        Dim d As DateTime = DateAdd(DateInterval.Day, -2, F_easter_date(Year))
        Return d
    End Function
    Sub test_easterSunday()
        Dim y As Integer
        For y = 1900 To 2100
            'If F_easter_date(y) <> EasterSunday(y) Then
            ' Debug.Print(y & " " & F_easter_date(y) & " " & EasterSunday(y))
            'End If
        Next
    End Sub
    '============================
    Public Function getHolidayList(ByVal vYear As Integer) As List(Of Date)

        Dim FirstWeek As Integer = 1
        Dim SecondWeek As Integer = 2
        Dim ThirdWeek As Integer = 3
        Dim FourthWeek As Integer = 4
        Dim LastWeek As Integer = 5

        Dim HolidayList As New List(Of Date)

        '   http://www.usa.gov/citizens/holidays.shtml      
        '   http://archive.opm.gov/operating_status_schedules/fedhol/2013.asp

        ' New Year's Day            Jan 1
        HolidayList.Add(DateSerial(vYear, 1, 1))

        ' Martin Luther King, Jr. third Mon in Jan
        HolidayList.Add(GetNthDayOfNthWeek(DateSerial(vYear, 1, 1), DayOfWeek.Monday, ThirdWeek))

        ' Washington's Birthday third Mon in Feb
        HolidayList.Add(GetNthDayOfNthWeek(DateSerial(vYear, 2, 1), DayOfWeek.Monday, ThirdWeek))

        'St. Patrick's Day
        HolidayList.Add(DateSerial(vYear, 3, 17))

        ' Memorial Day          last Mon in May
        HolidayList.Add(GetNthDayOfNthWeek(DateSerial(vYear, 5, 1), DayOfWeek.Monday, LastWeek))

        ' Independence Day      July 4
        HolidayList.Add(DateSerial(vYear, 7, 4))

        ' Labor Day             first Mon in Sept
        HolidayList.Add(GetNthDayOfNthWeek(DateSerial(vYear, 9, 1), DayOfWeek.Monday, FirstWeek))

        ' Columbus Day          second Mon in Oct
        HolidayList.Add(GetNthDayOfNthWeek(DateSerial(vYear, 10, 1), DayOfWeek.Monday, SecondWeek))

        ' Veterans Day          Nov 11
        HolidayList.Add(DateSerial(vYear, 11, 11))

        ' Thanksgiving Day      fourth Thur in Nov
        HolidayList.Add(GetNthDayOfNthWeek(DateSerial(vYear, 11, 1), DayOfWeek.Thursday, FourthWeek))

        ' Christmas Day         Dec 25
        HolidayList.Add(DateSerial(vYear, 12, 25))

        'saturday holidays are moved to Fri; Sun to Mon
        For i As Integer = 0 To HolidayList.Count - 1
            Dim dt As Date = HolidayList(i)
            If dt.DayOfWeek = DayOfWeek.Saturday Then
                HolidayList(i) = dt.AddDays(-1)
            End If
            If dt.DayOfWeek = DayOfWeek.Sunday Then
                HolidayList(i) = dt.AddDays(1)
            End If
        Next

        'return
        Return HolidayList

    End Function

    'Private Function GetNthDayOfNthWeek(ByVal dt As Date, ByVal DayofWeek As Integer, ByVal WhichWeek As Integer) As Date
    Function GetNthDayOfNthWeek(ByVal dt As Date, ByVal DayofWeek As Integer, ByVal WhichWeek As Integer) As Date
        'specify which day of which week of a month and this function will get the date
        'this function uses the month and year of the date provided

        'get first day of the given date
        Dim dtFirst As Date = DateSerial(dt.Year, dt.Month, 1)

        'get first DayOfWeek of the month
        Dim dtRet As Date = dtFirst.AddDays(6 - dtFirst.AddDays(-(DayofWeek + 1)).DayOfWeek)

        'get which week
        dtRet = dtRet.AddDays((WhichWeek - 1) * 7)

        'if day is past end of month then adjust backwards a week
        If dtRet >= dtFirst.AddMonths(1) Then
            dtRet = dtRet.AddDays(-7)
        End If

        'return
        Return dtRet

    End Function
    Private Function ReturnDay_GetNthDayOfNthWeek(ByVal dt As Date, ByVal DayofWeek As Integer, ByVal WhichWeek As Integer) As Integer
        'specify which day of which week of a month and this function will get the date
        'this function uses the month and year of the date provided

        'get first day of the given date
        Dim dtFirst As Date = DateSerial(dt.Year, dt.Month, 1)

        'get first DayOfWeek of the month
        Dim dtRet As Date = dtFirst.AddDays(6 - dtFirst.AddDays(-(DayofWeek + 1)).DayOfWeek)

        'get which week
        dtRet = dtRet.AddDays((WhichWeek - 1) * 7)

        'if day is past end of month then adjust backwards a week
        If dtRet >= dtFirst.AddMonths(1) Then
            dtRet = dtRet.AddDays(-7)
        End If

        'return
        Return dtRet.Day

    End Function
    Sub test_Holidays_GivenDate()
        Dim d As Date = Date.Today
        Dim i As Integer
        Dim Hol As String
        startTime()
        For i = 0 To 500 '0.001
            d = DateAdd(DateInterval.Day, 1, d)
            Hol = Holidays_GivenDate(d)
            'If Hol <> "" Then Debug.Print(d & " " & d.DayOfWeek.ToString & "  " & Hol)
        Next
        endTime()
        End
    End Sub
    Function Holidays_GivenDate(ByVal xDate As Date) As String
        Dim m As Integer = xDate.Month
        Dim d As Integer = xDate.Day
        Dim x As Date
        Dim ret As String = ""
        Select Case m
            Case 1
                If d = 1 Then
                    ret = "New Year's Day"
                Else
                    ' Martin Luther King, Jr. third Mon in Jan
                    x = ThirdThur_ReturnDate(3, 1, xDate)
                    If x.Day = d Then
                        ret = "Martin Luther King Day"
                    End If
                End If
            Case 2
                If d = 12 Then
                    ret = "Lincoln's Birthday"
                ElseIf d = 14 Then
                    ret = "Valentine's Day"
                ElseIf d = 22 Then
                    ret = "Washington's Birthday"
                Else
                    'Washington 's Birthday third Mon in Feb
                    'changed to President's Day on 1/8/2018

                    x = ThirdThur_ReturnDate(3, 1, xDate)
                    If x.Day = d Then
                        ret = "President's Day" '"Washington's Birthday"
                    End If
                End If
            Case 3
                ' DayLight Time Begins Second Sun in March
                x = ThirdThur_ReturnDate(2, 0, xDate)
                If x.Day = d Then
                    ret = "Daylight Time Begins"
                ElseIf d = 17 Then
                    ret = "St. Patrick's Day"
                Else
                    Dim dOfWeek As Integer = xDate.DayOfWeek
                    'Debug.Print(dOfWeek)
                    If dOfWeek = 0 Then 'Sunday
                        'Easter Range March 22 - April 25
                        x = F_easter_date(xDate.Year)
                        If x.Day = d AndAlso x.Month = m Then ret = "Easter"
                    ElseIf dOfWeek = 5 Then 'Friday
                        'Good Friday Range March 22 - April 25
                        x = F_GoodFriday_date(xDate.Year)
                        If x.Day = d AndAlso x.Month = m Then ret = "Good Friday"
                    End If
                End If
            Case 4
                'Dim dOfWeek As Integer = xDate.DayOfWeek
                'Debug.Print(dOfWeek)

                Dim dOfWeek As Integer = xDate.DayOfWeek
                'Debug.Print(dOfWeek)
                If dOfWeek = 0 Then 'Sunday
                    'Easter Range March 22 - April 25
                    x = F_easter_date(xDate.Year)
                    If x.Day = d AndAlso x.Month = m Then ret = "Easter"
                ElseIf dOfWeek = 5 Then 'Friday
                    'Good Friday Range March 22 - April 25
                    x = F_GoodFriday_date(xDate.Year)
                    If x.Day = d AndAlso x.Month = m Then ret = "Good Friday"
                End If

                '==

                'Easter Range March 22 - April 25
                'x = F_easter_date(xDate.Year)
                'If x.Day = d AndAlso x.Month = m Then ret = "Easter"
            Case 5
                ' Memorial Day          last Mon in May
                x = ThirdThur_ReturnDate(5, 1, xDate)
                If x.Day = d Then ret = "Memorial Day"
            Case 6
            Case 7
                ' Independence Day      July 4
                If d = 4 Then ret = "Independence Day"
            Case 8
            Case 9
                ' Labor Day             first Mon in Sept
                x = ThirdThur_ReturnDate(1, 1, xDate)
                If x.Day = d Then ret = "Labor Day"
            Case 10
                If d = 31 Then
                    ret = "Halloween"
                Else
                    ' Columbus Day          second Mon in Oct
                    x = ThirdThur_ReturnDate(2, 1, xDate)
                    If x.Day = d Then ret = "Columbus Day"
                End If
            Case 11
                If d = 11 Then
                    ' Veterans Day          Nov 11
                    ret = "Veterans Day"
                Else
                    ' DayLight Time Ends First Sun in November
                    x = ThirdThur_ReturnDate(1, 0, xDate)
                    If x.Day = d Then
                        ret = "Daylight Time Ends"
                    Else
                        ' Thanksgiving Day      fourth Thur in Nov
                        x = ThirdThur_ReturnDate(4, 4, xDate)
                        If x.Day = d Then ret = "Thanksgiving"
                    End If
                End If

            Case 12
                If d = 25 Then ret = "Christmas"
        End Select
        Return ret
    End Function
    Function Holidays_GivenMonthYear_ReturnDates(ByVal xDate As Date) As Date()
        Dim m As Integer = xDate.Month
        Dim y As Integer = xDate.Year

        Dim xDaysInMonth As Integer = DateTime.DaysInMonth(y, m)
        Dim Holidays(xDaysInMonth) As Date
        Dim cnt As Integer = 0
        Dim i As Integer

        Dim d As Date = New Date(y, m, 1)
        d = DateAdd(DateInterval.Day, -1, d)
        Dim s As String
        For i = 1 To xDaysInMonth
            d = DateAdd(DateInterval.Day, 1, d)
            s = Holidays_GivenDate(d)
            If s <> "" Then
                cnt += 1
                Holidays(cnt) = d
            End If
        Next
        ReDim Preserve Holidays(cnt)
        Return Holidays

    End Function
    Function Holidays_GivenMonthYear(ByVal xDate As Date) As HolidaysType()
        Dim m As Integer = xDate.Month
        Dim y As Integer = xDate.Year

        Dim xDaysInMonth As Integer = DateTime.DaysInMonth(y, m)
        Dim Holidays(xDaysInMonth) As HolidaysType
        Dim cnt As Integer = 0
        Dim i As Integer

        Dim d As Date = New Date(y, m, 1)
        d = DateAdd(DateInterval.Day, -1, d)
        Dim s As String
        For i = 1 To xDaysInMonth
            d = DateAdd(DateInterval.Day, 1, d)
            s = Holidays_GivenDate(d)
            If s <> "" Then
                cnt += 1
                Holidays(cnt).HolidayDate = d
                Holidays(cnt).HolidayName = s
            End If
        Next
        ReDim Preserve Holidays(cnt)
        Return Holidays
    End Function
    Function Holidays_GivenYear(ByVal year As Integer) As HolidaysType() 'reliable
        Dim x(100) As HolidaysType
        Dim i As Integer = 0
        '==
        Dim FirstWeek As Integer = 1
        Dim SecondWeek As Integer = 2
        Dim ThirdWeek As Integer = 3
        Dim FourthWeek As Integer = 4
        Dim LastWeek As Integer = 5

        Dim HolidayList As New List(Of Date)


        ' New Year's Day            Jan 1
        i += 1

        x(i).HolidayName = "New Year's Day"
        x(i).HolidayDate = DateSerial(year, 1, 1)

        i += 1
        ' Martin Luther King, Jr. third Mon in Jan
        x(i).HolidayName = "Martin Luther King Day" ' "Birthday of Martin Luther King, Jr."
        x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 1, 1), DayOfWeek.Monday, ThirdWeek)

        'Lincoln's Birthday 'added 1/8/2018
        i += 1
        x(i).HolidayName = "Lincoln's Birthday"
        x(i).HolidayDate = DateSerial(year, 2, 12)

        'Valentine's Day 
        i += 1
        x(i).HolidayName = "Valentine's Day"
        x(i).HolidayDate = DateSerial(year, 2, 14)


        ' Washington's Birthday third Mon in Feb
        'changed to President's Day on 1/8/2018 -Range 2-15 to 2-21 
        i += 1
        x(i).HolidayName = "President's Day" '"Washington's Birthday" ' "Now is the time for all good men to come to the rescue" '"WwwwwwWwwwwwWww" '"Washington's Birthday"
        x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 2, 1), DayOfWeek.Monday, ThirdWeek)


        'Washington's Birthday 'added 1/8/2018
        i += 1
        x(i).HolidayName = "Washington's Birthday"
        x(i).HolidayDate = DateSerial(year, 2, 22)


        ' DayLight Time Begins Second Sun in March
        i += 1
        x(i).HolidayName = "Daylight Time Begins"
        x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 3, 1), DayOfWeek.Sunday, SecondWeek)

        'St. Patrick's Day
        i += 1
        x(i).HolidayName = "St. Patrick's Day"
        x(i).HolidayDate = DateSerial(year, 3, 17)


        'Good Friday Range March 20 - April 23
        i += 1
        x(i).HolidayName = "Good Friday"
        x(i).HolidayDate = F_GoodFriday_date(year)


        'Easter Range March 22 - April 25
        i += 1
        x(i).HolidayName = "Easter"
        x(i).HolidayDate = F_easter_date(year)


        ' Memorial Day          last Mon in May
        i += 1
        x(i).HolidayName = "Memorial Day"
        x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 5, 1), DayOfWeek.Monday, LastWeek)


        ' Independence Day      July 4
        i += 1
        x(i).HolidayName = "Independence Day"
        x(i).HolidayDate = DateSerial(year, 7, 4)


        ' Labor Day             first Mon in Sept
        i += 1
        x(i).HolidayName = "Labor Day"
        x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 9, 1), DayOfWeek.Monday, FirstWeek)


        ' Columbus Day          second Mon in Oct
        i += 1
        x(i).HolidayName = "Columbus Day"
        x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 10, 1), DayOfWeek.Monday, SecondWeek)


        'Halloween Oct 31
        i += 1
        x(i).HolidayName = "Halloween"
        x(i).HolidayDate = DateSerial(year, 10, 31)

        ' DayLight Time Ends First Sun in November
        i += 1
        x(i).HolidayName = "Daylight Time Ends"
        x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Sunday, FirstWeek)

        ' Veterans Day          Nov 11
        i += 1
        x(i).HolidayName = "Veterans Day"
        x(i).HolidayDate = DateSerial(year, 11, 11)

        ' Thanksgiving Day      fourth Thur in Nov
        i += 1
        x(i).HolidayName = "Thanksgiving" '"Thanksgiving Day"
        x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Thursday, FourthWeek)

        ' Christmas Day         Dec 25
        i += 1
        x(i).HolidayName = "Christmas"
        x(i).HolidayDate = DateSerial(year, 12, 25)

        ''saturday holidays are moved to Fri; Sun to Mon
        'For i As Integer = 0 To HolidayList.Count - 1
        '    Dim dt As Date = HolidayList(i)
        '    If dt.DayOfWeek = DayOfWeek.Saturday Then
        '        HolidayList(i) = dt.AddDays(-1)
        '    End If
        '    If dt.DayOfWeek = DayOfWeek.Sunday Then
        '        HolidayList(i) = dt.AddDays(1)
        '    End If
        'Next

        '==
        ReDim Preserve x(i)
        Return x
    End Function
    Sub test_holidays_givenMonthYear()

        startTime()
        Dim H() As Date = Holidays_GivenMonthYear_ReturnDates(Date.Today)
        endTime()
        Dim nItems As Integer = UBound(H)
        Dim i As Integer
        For i = 1 To nItems
            Debug.Print(CType(H(i), String))
        Next
        '==

        'startTime()
        'Dim H() As HolidaysType = Holidays_GivenMonthYear(Date.Today)
        'endTime()
        'Dim nItems As Integer = UBound(H)
        'Dim i As Integer
        'For i = 1 To nItems
        '    Debug.Print(H(i).HolidayName & " " & H(i).HolidayDate)
        'Next
    End Sub
    Sub test_Holidays_2years()

        setHolidaysFor2Years() '0.005 seconds

        Dim i As Integer
        Dim j As Integer
        Dim nRecs As Integer = UBound(ApptGV.Holidays)
        'For i = 1 To nRecs
        '    With ApptGV.Holidays(i)
        '        Debug.Print(i.ToString & " " & .HolidayName & " " & .HolidayDate)
        '    End With
        'Next

        Dim d As Date = Date.Today
        Dim hol() As HolidaysType
        For i = 1 To 12
            hol = Holidays_GivenMonthYear(d)
            For j = 1 To UBound(hol)
                With hol(j)
                    Debug.Print(Format(d, "MMMM yyyy ") & .HolidayName & " " & .HolidayDate)
                End With

            Next
            d = DateAdd(DateInterval.Month, 1, d)
        Next
    End Sub
    Function initBirthdayRecord()
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
    Function getBirthdayRecord(ByVal LocInFile As Integer)
        Dim rec As BirthdayRecordType = initBirthdayRecord()


        Dim i As Integer
        Dim nRecs As Integer = UBound(ApptGV.Birthdays)
        ReDim ApptGV.BirthdayDates(nRecs * 2)

        Dim currentYear As Integer = Date.Today.Year
        Dim nextYear As Integer = currentYear + 1

        Dim cnt As Integer = 0

        Dim d As Date
        For i = 1 To nRecs
            With ApptGV.Birthdays(i)
                If .ID = LocInFile AndAlso .DeleteFlag = 0 Then
                    rec = ApptGV.Birthdays(i)
                    Exit For
                End If
            End With
        Next
        Return rec
    End Function
    Sub setBirthdayDates_next2years()
        Dim i As Integer
        Dim nRecs As Integer = UBound(ApptGV.Birthdays)
        ReDim ApptGV.BirthdayDates(nRecs * 2)

        Dim currentYear As Integer = Date.Today.Year
        Dim nextYear As Integer = currentYear + 1

        Dim cnt As Integer = 0

        Dim d As Date
        For i = 1 To nRecs
            With ApptGV.Birthdays(i)
                If .DeleteFlag = 0 Then
                    d = DateSerial(currentYear, .Month, .Day)
                    cnt += 1
                    ApptGV.BirthdayDates(cnt) = d
                End If
            End With
        Next
        For i = 1 To nRecs
            With ApptGV.Birthdays(i)
                If .DeleteFlag = 0 Then
                    d = DateSerial(nextYear, .Month, .Day)
                    cnt += 1
                    ApptGV.BirthdayDates(cnt) = d
                End If
            End With
        Next
        ReDim Preserve ApptGV.BirthdayDates(cnt)
        '====
        'For i = 1 To UBound(ApptGV.BirthdayDates)
        '    'With ApptGV.Birthdays(i)
        '    Debug.Print(i & " " & ApptGV.BirthdayDates(i))
        'Next
    End Sub
    Sub setHolidayDates_next2years()

        Dim i As Integer
        Dim nRecs As Integer = UBound(ApptGV.Holidays)
        ReDim ApptGV.HolidayDates(nRecs)
        For i = 1 To nRecs
            ApptGV.HolidayDates(i) = ApptGV.Holidays(i).HolidayDate
        Next
    End Sub
    Sub setHolidaysFor2Years()
        ReDim ApptGV.Holidays(1000)
        Dim cnt As Integer = 0
        Dim i As Integer
        Dim currentYear As Integer = Date.Today.Year
        Dim nextYear As Integer = currentYear + 1

        Dim y() As HolidaysType = Holidays_GivenYear(currentYear)
        For i = 1 To UBound(y)
            cnt += 1
            ApptGV.Holidays(cnt) = y(i)
        Next
        y = Holidays_GivenYear(nextYear)
        For i = 1 To UBound(y)
            cnt += 1
            ApptGV.Holidays(cnt) = y(i)
        Next
        ReDim Preserve ApptGV.Holidays(cnt)
    End Sub
    Sub test_Holidays_GivenYear_GivenDate(ByVal Year As Integer)
        startTime()

        Dim y() As HolidaysType = Holidays_GivenYear(Year)
        endTime()
        Dim nItems As Integer = UBound(y)

        Dim cnt As Integer = 0
        Dim i As Integer

        For i = 1 To UBound(y)
            Debug.Print(i.ToString & " " & y(i).HolidayName & " " & y(i).HolidayDate)
        Next

        Dim adj As Integer
        If DateTime.IsLeapYear(Year) Then adj = 1 Else adj = 0
        Dim d As Date = New Date(Year - 1, 12, 31)
        Dim dd As Date
        Dim s As String
        '==
        'For i = 1 To UBound(y)
        '    With y(i)
        '        Debug.Print(.HolidayDate & " " & .HolidayName)
        '    End With
        'Next
        '===
        startTime()

        For i = 1 To 365 + adj
            dd = DateAdd(DateInterval.Day, i, d)
            s = Holidays_GivenDate(dd)
            If s <> "" Then
                'Debug.Print(dd & " " & s)

                cnt += 1
                'Debug.Print(s & " " & y(cnt).HolidayName & " " & dd & " " & y(cnt).HolidayDate & " " & cnt)

                If s <> y(cnt).HolidayName Then
                    Debug.Print(s & " " & y(cnt).HolidayName & " " & dd & " " & y(cnt).HolidayDate & " " & cnt)
                End If
            End If
        Next
        endTime()
    End Sub
    Function Holidays_Next12Months_fromGivenDate(ByVal givenDate As Date) As HolidaysType()
        Dim i As Integer = 0
        Dim c As Integer = 0

        Dim year As Integer = DatePart(DateInterval.Year, givenDate)
        Dim y() As HolidaysType = Holidays_GivenYear(year)
        Dim nItems As Integer = UBound(y)
        Dim x(nItems) As HolidaysType
        Dim z() As HolidaysType = Holidays_GivenYear(year + 1)
        For i = 1 To nItems
            If y(i).HolidayDate >= givenDate Then
                c += 1
                x(c) = y(i)
            End If
        Next
        'Dim givenDatePlus1 = DateAdd(DateInterval.Year, 1, givenDate)
        For i = 1 To nItems - c
            'If z(i).HolidayDate < givenDatePlus1 Then
            c += 1
            x(c) = z(i)
            'End If
        Next
        'ReDim Preserve x(c)
        Return x
    End Function
    Function IsHoliday(ByVal givenDate As Date) As Boolean
        'on 4/5/2015 easter added easter logic
        '==
        Dim FirstWeek As Integer = 1
        Dim SecondWeek As Integer = 2
        Dim ThirdWeek As Integer = 3
        Dim FourthWeek As Integer = 4
        Dim LastWeek As Integer = 5
        '==
        Dim m As Integer = DatePart(DateInterval.Month, givenDate)
        Dim d As Integer = DatePart(DateInterval.Day, givenDate)
        Dim year As Integer = DatePart(DateInterval.Year, givenDate)
        Dim b As Boolean = False
        Select Case m
            Case 1 'newYear MartinLuther
                If d = 1 Then
                    b = True
                Else
                    'givenDate = GetNthDayOfNthWeek(DateSerial(year, 1, 1), DayOfWeek.Monday, ThirdWeek)
                    If d = ReturnDay_GetNthDayOfNthWeek(DateSerial(year, 1, 1), DayOfWeek.Monday, ThirdWeek) Then b = True
                End If
            Case 2 'Washingtion's Bday
                If d = 14 Or d = 12 Or d = 22 Then
                    b = True
                Else
                    If d = ReturnDay_GetNthDayOfNthWeek(DateSerial(year, 2, 1), DayOfWeek.Monday, ThirdWeek) Then b = True
                End If
                'Case 3
                'Case 4
            Case 3
                'x(i).HolidayName = "Daylight Time Begins"
                'x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 3, 1), DayOfWeek.Sunday, SecondWeek)

                If d = 17 Then
                    b = True
                Else
                    If d = ReturnDay_GetNthDayOfNthWeek(DateSerial(year, 3, 1), DayOfWeek.Sunday, SecondWeek) Then
                        b = True
                    Else
                        Dim xdate As Date = F_easter_date(year)
                        If DatePart(DateInterval.Month, xdate) = 3 Then
                            If DatePart(DateInterval.Day, xdate) = d Then
                                b = True
                            End If
                        End If
                    End If
                End If
            Case 4
                Dim xdate As Date = F_easter_date(year)
                If DatePart(DateInterval.Month, xdate) = 4 Then
                    If DatePart(DateInterval.Day, xdate) = d Then
                        b = True
                    End If
                End If
            Case 5 'mem day
                If d = ReturnDay_GetNthDayOfNthWeek(DateSerial(year, 5, 1), DayOfWeek.Monday, LastWeek) Then b = True
                'Case 6
            Case 7 '4th of July
                If d = 4 Then b = True
                'Case 8
            Case 9 'labor day
                If d = ReturnDay_GetNthDayOfNthWeek(DateSerial(year, 9, 1), DayOfWeek.Monday, FirstWeek) Then b = True
            Case 10 'columbus day
                If d = 31 Then
                    If d = 31 Then b = True 'halloween
                Else
                    If d = ReturnDay_GetNthDayOfNthWeek(DateSerial(year, 10, 1), DayOfWeek.Monday, SecondWeek) Then b = True
                End If
            Case 11 'vetDay and Thanksgiving
                'x(i).HolidayName = "Daylight Time Ends"
                'x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Sunday, FirstWeek)

                If d = 11 Then
                    b = True
                Else
                    If d = ReturnDay_GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Sunday, FirstWeek) Then
                        b = True
                    Else
                        'thankgiving
                        If d = ReturnDay_GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Thursday, FourthWeek) Then b = True
                    End If
                End If
            Case 12 'Christmas
                If d = 25 Then b = True
        End Select

        Return b
    End Function
    'Function ReturnHoliday(ByVal givenDate As Date) As HolidaysType
    '    '==
    '    Dim FirstWeek As Integer = 1
    '    Dim SecondWeek As Integer = 2
    '    Dim ThirdWeek As Integer = 3
    '    Dim FourthWeek As Integer = 4
    '    Dim LastWeek As Integer = 5
    '    '==
    '    Dim m As Integer = DatePart(DateInterval.Month, givenDate)
    '    Dim d As Integer = DatePart(DateInterval.Day, givenDate)
    '    Dim year As Integer = DatePart(DateInterval.Year, givenDate)
    '    Dim x As HolidaysType
    '    x.HolidayDate = New Date(1, 1, 1, 0, 0, 0)
    '    x.HolidayName = ""
    '    Select Case m
    '        Case 1 'newYear MartinLuther
    '            If d = 1 Then
    '                x.HolidayName = "New Year's Day"
    '                x.HolidayDate = DateSerial(year, 1, 1)
    '            Else
    '                ' Martin Luther King, Jr. third Mon in Jan
    '                x.HolidayName = "Martin Luther King Day" '"Birthday of Martin Luther King, Jr." changed 12/28/2016
    '                x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 1, 1), DayOfWeek.Monday, ThirdWeek)
    '            End If
    '        Case 2 ' Washington's Birthday third Mon in Feb
    '            x.HolidayName = "Washington's Birthday"
    '            x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 2, 1), DayOfWeek.Monday, ThirdWeek)
    '            'Case 3
    '            'Case 4
    '        Case 3
    '            If d = 17 Then
    '                x.HolidayName = "St. Patrick's Day"
    '                x.HolidayDate = DateSerial(year, 3, 17)
    '            Else
    '                Dim tempDate As Date = GetNthDayOfNthWeek(DateSerial(year, 3, 1), DayOfWeek.Sunday, SecondWeek)
    '                If d = tempDate.Day Then
    '                    x.HolidayName = "Daylight Time Begins"
    '                    x.HolidayDate = tempDate
    '                Else
    '                    Dim xdate As Date = F_easter_date(year)
    '                    If DatePart(DateInterval.Month, xdate) = 3 Then
    '                        If DatePart(DateInterval.Day, xdate) = d Then
    '                            x.HolidayName = "Easter"
    '                            x.HolidayDate = DateSerial(year, 3, d)
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Case 4
    '            Dim xdate As Date = F_easter_date(year)
    '            If DatePart(DateInterval.Month, xdate) = 4 Then
    '                If DatePart(DateInterval.Day, xdate) = d Then
    '                    x.HolidayName = "Easter"
    '                    x.HolidayDate = DateSerial(year, 4, d)
    '                End If
    '            End If

    '        Case 5 ' Memorial Day          last Mon in May
    '            x.HolidayName = "Memorial Day"
    '            x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 5, 1), DayOfWeek.Monday, LastWeek)
    '            'Case 6
    '        Case 7 ' Independence Day      July 4
    '            x.HolidayName = "Independence Day"
    '            x.HolidayDate = DateSerial(year, 7, 4)
    '            'Case 8
    '        Case 9 ' Labor Day             first Mon in Sept
    '            x.HolidayName = "Labor Day"
    '            x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 9, 1), DayOfWeek.Monday, FirstWeek)
    '        Case 10 ' Columbus Day          second Mon in Oct
    '            If d = 31 Then
    '                x.HolidayName = "Halloween"
    '                x.HolidayDate = DateSerial(year, 10, 31)

    '            Else
    '                x.HolidayName = "Columbus Day"
    '                x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 10, 1), DayOfWeek.Monday, SecondWeek)

    '            End If
    '        Case 11 'vetDay and Thanksgiving
    '            Dim TempDate As Date = GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Sunday, FirstWeek)

    '            ' Veterans Day          Nov 11
    '            If d = 11 Then
    '                x.HolidayName = "Veterans Day"
    '                x.HolidayDate = DateSerial(year, 11, 11)
    '            ElseIf d = TempDate.Day Then
    '                x.HolidayName = "Daylight Time Ends"
    '                x.HolidayDate = TempDate
    '            Else
    '                ' Thanksgiving Day      fourth Thur in Nov
    '                x.HolidayName = "Thanksgiving Day"
    '                x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Thursday, FourthWeek)
    '            End If

    '        Case 12 ' Christmas Day         Dec 25
    '            x.HolidayName = "Christmas"
    '            x.HolidayDate = DateSerial(year, 12, 25)
    '    End Select
    '    Return x
    'End Function
    Function ReturnHolidayName(ByVal givenDate As Date) As String
        '==
        Dim FirstWeek As Integer = 1
        Dim SecondWeek As Integer = 2
        Dim ThirdWeek As Integer = 3
        Dim FourthWeek As Integer = 4
        Dim LastWeek As Integer = 5
        '==
        Dim m As Integer = DatePart(DateInterval.Month, givenDate)
        Dim d As Integer = DatePart(DateInterval.Day, givenDate)
        Dim year As Integer = DatePart(DateInterval.Year, givenDate)
        Dim x As String = ""
        Select Case m
            Case 1 'newYear MartinLuther
                If d = 1 Then
                    x = "New Year's Day"
                    'x.HolidayDate = DateSerial(year, 1, 1)
                Else
                    ' Martin Luther King, Jr. third Mon in Jan
                    x = "Martin Luther King Day" '"Birthday of Martin Luther King, Jr." 'changed 12/28/2016
                    'x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 1, 1), DayOfWeek.Monday, ThirdWeek)
                End If
            Case 2 ' Washington's Birthday third Mon in Feb
                If d = 12 Then
                    x = "Lincoln's Birthday"
                ElseIf d = 14 Then
                    x = "St. Valentine's Day"
                ElseIf d = 22 Then
                    x = "Washington's Birthday"
                Else
                    x = "President's Day" '"Washington's Birthday"
                End If

                'x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 2, 1), DayOfWeek.Monday, ThirdWeek)
                'Case 3
                'Case 4
            Case 3
                'x(i).HolidayName = "Daylight Time Begins"
                'x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 3, 1), DayOfWeek.Sunday, SecondWeek)

                If d = 17 Then
                    x = "St. Patrick's Day"
                ElseIf givenDate = GetNthDayOfNthWeek(DateSerial(year, 3, 1), DayOfWeek.Sunday, SecondWeek) Then
                    x = "Daylight Time Begins"
                Else
                    x = "Easter"
                End If

            Case 4
                x = "Easter"
            Case 5 ' Memorial Day          last Mon in May
                x = "Memorial Day"
                'x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 5, 1), DayOfWeek.Monday, LastWeek)
                'Case 6
            Case 7 ' Independence Day      July 4
                x = "Independence Day"
                'x.HolidayDate = DateSerial(year, 7, 4)
                'Case 8
            Case 9 ' Labor Day             first Mon in Sept
                x = "Labor Day"
                'x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 9, 1), DayOfWeek.Monday, FirstWeek)
            Case 10 ' Columbus Day          second Mon in Oct
                If d = 31 Then
                    x = "Halloween"
                Else
                    'x(i).HolidayName = "Daylight Time Ends"
                    'x(i).HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Sunday, FirstWeek)

                    x = "Columbus Day"
                End If
                'x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 10, 1), DayOfWeek.Monday, SecPondWeek)
            Case 11 'vetDay and Thanksgiving
                ' Veterans Day          Nov 11
                If d = 11 Then
                    x = "Veterans Day"
                    'x.HolidayDate = DateSerial(year, 11, 11)
                ElseIf givenDate = GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Sunday, FirstWeek) Then
                    x = "Daylight Time Ends"
                Else
                    ' Thanksgiving Day      fourth Thur in Nov
                    x = "Thanksgiving Day"
                    'x.HolidayDate = GetNthDayOfNthWeek(DateSerial(year, 11, 1), DayOfWeek.Thursday, FourthWeek)
                End If
            Case 12 ' Christmas Day         Dec 25
                x = "Christmas"
                'x.HolidayDate = DateSerial(year, 12, 25)
        End Select
        Return x
    End Function

    Sub test_Washington() 'Washington's Birthday third Mon in Feb -
        Dim d As Date
        Dim year As Integer
        Dim lowest As Integer = 50
        Dim day As Integer
        For year = 2000 To 3000 '2100
            d = GetNthDayOfNthWeek(DateSerial(year, 2, 1), DayOfWeek.Monday, 3)
            day = DatePart(DateInterval.Day, d)
            If day < lowest Then lowest = day : Debug.Print(CType(year, String))
        Next
        Debug.Print(CType(lowest, String))

    End Sub
    Function getOneHoliday(ByVal xDate As DateTime) As String
        Dim year As Integer = DatePart(DateInterval.Year, xDate)
        'Dim month As Integer = DatePart(DateInterval.Month, xDate)
        'Dim day As Integer = DatePart(DateInterval.Day, xDate)
        Dim x() As HolidaysType = Holidays_GivenYear(year)
        Dim i As Integer
        Dim h As String = ""

        For i = 1 To UBound(x)
            If x(i).HolidayDate = xDate Then
                h = Trim(x(i).HolidayName)
                Exit For
            End If
        Next
        Return h
    End Function
    Sub test_getOneHoliday() '365 day - 0 to .01 seconds
        Dim y, d As Integer
        Dim dd As Date = #12/31/2013#
        Dim h As String
        startTime()
        For d = 1 To 366

            dd = DateAdd(DateInterval.Day, 1, dd)
            h = getOneHoliday(dd)
            'If h <> "" Then Debug.Print(h)
        Next

        endTime()
        ' DiffTime()
    End Sub
    Sub replace3charMonthWith4char(ByRef s As String)
        'Dim xinstr As Integer
        'xinstr(s, "Jun") : If xinstr Then s = Replace(s)
        s = Replace(s, "Jun", "June")
        s = Replace(s, "Jul", "July")
        s = Replace(s, "Sep", "Sept")
    End Sub
    Function getDaysFromTodayString(ByVal selDate As Date) As String
        Dim retString As String = ""
        Dim d As Date = Date.Today
        Dim xdays As String
        Dim dd As String
        If selDate = d Then
            retString = "Today"
        Else
            Dim dateDif As Integer = CInt(DateDiff(DateInterval.Day, d, selDate))
            If Math.Abs(dateDif) = 1 Then
                dd = "Day"
            Else
                dd = "Days"
            End If
            If dateDif > 0 Then
                xdays = dd & " Hence"
            Else
                xdays = dd & " Ago"
            End If
            retString = LTrim(Math.Abs(dateDif).ToString & " " & xdays)
        End If
        Return retString
    End Function
    Function arrayOfFuture() As Integer
        MsgBox("arrayOfFuture - not active")
        Return 0
        End

        ''from: getAllFutureAppointments_Time
        ''from: Function getNextAppt_Time


        ''Dim retDates As String = "No future appointments were found"

        'Dim ApptRec As ApptRecType = initApptRecType()
        'Dim ReadApptRec As ApptRecType = initApptRecType()
        'Dim a() As LongIntegerType = set_AllAppts_LongInteger_ReturnLongIntergerArray()


        ''new
        'Dim nRecs As Integer = UBound(a)
        'Dim cnt As Integer = 0
        'Dim retDates(0) As Date


        ''Dim xL As Long = Date.Now.Ticks + 1 ' add 1 in event ticks are an even minute to assure getting NEXT
        ''8/3/2016 changed to include entire date
        'Dim xL As Long = Date.Today.Ticks ' add 1 in event ticks are an even minute to assure getting NEXT

        'Dim xStart = biSearch_GTE_LongInteger(xL, a)
        'Dim xEnd As Integer = UBound(a)
        'If xStart > xEnd Then
        '    ReDim a(0)
        '    Return retDates 'retDates 'ApptRec 'no appt
        'End If
        'ReDim retDates(nRecs)

        ''--------------------------------------------------------
        'Dim highestApptTicks As Long = 86340 * TimeSpan.TicksPerSecond

        'Dim k As Integer
        'Dim ff As Integer = FreeFile()
        'FileOpen(ff, xBuildFullFileName(f_Appt.FileName), OpenMode.Random, , , f_Appt.FileLen)
        'For k = xStart To xEnd
        '    If (a(k).L Mod TimeSpan.TicksPerDay) <= highestApptTicks Then
        '        FileGet(ff, ReadApptRec, a(k).i) 'gets apptRec
        '        If Trim(ReadApptRec.msg) <> "" AndAlso Mid(ReadApptRec.msg, 1, 1) <> Chr(0) Then
        '            cnt += 1
        '            retDates(cnt) = New Date((ReadApptRec.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay) ' formatNextApptTimeMsg(ApptRec)
        '        End If
        '    End If
        'Next
        'ReDim Preserve retDates(cnt)
        'FileClose(ff)
        'ReDim a(0)
        'Array.Sort(retDates)
        'Dim i As Integer

        'cnt = 1
        'For i = 2 To UBound(retDates)
        '    If retDates(cnt) <> retDates(i) Then
        '        cnt += 1
        '        retDates(cnt) = retDates(i)
        '        'Else
        '        '   cnt = cnt
        '    End If
        'Next
        'For i = 1 To cnt
        '    retDates(i - 1) = retDates(i)
        'Next
        'ReDim Preserve retDates(cnt - 1)
        'Return retDates 'ApptRec

    End Function
    'Sub test_arrayOfFuture()
    '    'Dim d(0) As Date
    '    'Debug.Print(d(0).Ticks)
    '    'Debug.Print(d(0))
    '    Dim d() As Date
    '    d = arrayOfFuture()
    '    Dim i As Integer
    '    For i = 0 To UBound(d)
    '        Debug.Print(i.ToString & " " & d(i))
    '    Next
    '    End
    'End Sub
    Sub MonthCalendar_putCurrentMonthOnTop_DateToday(ByRef passSkip As Boolean, ByRef MthCal As MonthCalendar)
        '============================solves the problem of current date not in top calendar

        '1. first get the starting and ending dates of the displayRange (if 2 months are shown ie mar apr then mar 1 - apr 30
        Dim startDate As Date = MthCal.GetDisplayRange(True).Start
        Dim endDate As Date = MthCal.GetDisplayRange(True).End
        'Debug.Print(startDate)
        'Debug.Print(endDate)

        Dim currentDate As Date = Date.Today.Date '(could also do: currentDate = MthCal.SelectionStart)
        Dim currentMonth = currentDate.Month 'Date.Today.Month

        If currentMonth = startDate.Month Then
            'ok - already on top
        ElseIf currentMonth = endDate.Month Then
            'the purpose of skip to to eliminate changing the appt and todo display to a different date
            Dim xSkip As Boolean = passSkip
            passSkip = True

            MthCal.SelectionStart = DateAdd(DateInterval.Month, 1, Date.Today.Date) 'force display to increment by one month
            'increment by 1 month then
            'decrement by 1 month
            MthCal.SelectionStart = currentDate 'Date.Today.Date 'now put date back to where it should be

            passSkip = xSkip
        Else
            'problem
        End If
        '============================
    End Sub
    Function getDaysHence(ByVal dd As Date) As Integer
        Dim d As Date = Date.Today
        Dim dif As Integer = CInt(DateDiff(DateInterval.Day, d, dd))
        Return dif
    End Function
    Sub formatDaysHence_PassLabel(ByVal dd As Date, ByRef LBL As Label)
        LBL.Text = formatDaysHence(dd)
        Dim days As Integer = getDaysHence(dd)
        If days < 0 Then
            LBL.ForeColor = Color.Red
        Else
            LBL.ForeColor = Color.Black
        End If

    End Sub

    Function formatDaysHence(ByVal dd As Date) As String
        Dim days As Integer = getDaysHence(dd)
        Dim x As String = ""
        'Return days

        Select Case Math.Sign(days)
            Case 1 : If days = 1 Then x = days.ToString & " Day Hence" Else x = days.ToString & " Days Hence"
            Case 0 : x = "Today"
            Case -1 : If days = -1 Then x = Math.Abs(days).ToString & " Day Ago" Else x = Math.Abs(days).ToString & " Days Ago"
        End Select
        Return x

        'Select Case Math.Sign(days)
        '    Case 1
        '        Select Case Math.Abs(days)
        '            Case 1
        '            Case Else
        '        End Select
        '    Case 0
        '        Select Case Math.Abs(days)
        '            Case 1
        '            Case Else
        '        End Select

        'End Select
        'Return x
    End Function
    Private Function ThirdThur_ReturnDate(ByVal WkNum As Integer, ByVal xDayOfWeek As Integer, ByVal passDate As Date) As Date
        'wkNum=1 to 5 xDayofWeek 0 to 6
        Dim xdaysInMonth As Integer = DateTime.DaysInMonth(passDate.Year, passDate.Month)
        Dim workingDate As Date = New Date(passDate.Year, passDate.Month, 1)
        Dim retDate As Date
        Dim DayOfWeekOfFirstOfMonth As Integer = workingDate.DayOfWeek
        'Dim adj As Integer = ((DayOfWeekOfFirstOfMonth - xDayOfWeek + 7) Mod 7) + 1
        Dim adj As Integer = ((xDayOfWeek - DayOfWeekOfFirstOfMonth + 7) Mod 7) + 1
        Dim dayNum As Integer = adj + (WkNum - 1) * 7
        If dayNum > xdaysInMonth Then dayNum = dayNum - 7
        retDate = New Date(passDate.Year, passDate.Month, dayNum)
        Return retDate
    End Function
    Function replaceNewLineWithForwardSlash(ByVal x As String)
        Dim y As String
        y = Replace(x, vbNewLine, "/")
        Return y
    End Function
End Module
