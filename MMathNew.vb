Option Explicit On
Module MMathNew
    Public ValueE As Decimal = Math.Exp(1)
    'exp(1) is equivilant to (1 + 1/1000000)^1000000 or as many zeros that exist in the universe!
    'or fv=1.05^10 = 1.628894626777442 or valueE^(.05*10) = 1.6487212707001266 or valueE^-(.05*10) =0.606530659712634

    Function mp(ByVal intRate As Decimal, Optional mths As Decimal = 360) As Decimal
        Dim n As Decimal = mths
        Dim i As Decimal = intRate / 1200
        Dim p As Decimal = (1 + i) ^ -n '-360


        Dim ans As Decimal = i / (1 - p)
        Return ans
    End Function
    Function pvAnnuityMonthly(ByVal intRate As Decimal, Optional mths As Decimal = 360) As Decimal
        Dim n As Decimal = mths 'yrs * 12
        Dim i As Decimal = intRate / 1200
        Dim p As Decimal = (1 + i) ^ -n '360

        Dim ans As Decimal = (1 - p) / i
        Return ans
    End Function

    Function pvAnnuity(ByVal intRate As Decimal, Optional mths As Decimal = 360) As Decimal
        Dim n As Decimal = mths / 12
        Dim i As Decimal = intRate / 100
        Dim p As Decimal = (1 + i) ^ -n

        Dim ans As Decimal = (1 - p) / i
        Return ans
    End Function
    Function computeInterestRate(ByVal paymentPer1dollarLoan As Decimal, ByVal mths As Decimal) As Decimal
        'Dim i As Decimal
        If 1 / mths > paymentPer1dollarLoan Then
            MsgBox("Duration is not long enough as 1/Mths > PaymentPer1dollarLoan")
            Return 0
        End If

        Const b As Long = 1000000000
        Dim xmp As Long
        Dim low As Long = 1
        Dim high As Long = 1000 * b
        Dim increment As Long = 1
        Dim half As Long
        Dim cnt As Integer = 0

        Dim ppd As Long = paymentPer1dollarLoan * b
        Do
            half = (high + low) \ 2
            xmp = mpActual(CDec(half / b), mths) * b ' mp(CDec(half / b), yrs) * b
            If xmp >= ppd Then high = half - increment Else low = half + increment
            cnt += 1
            If cnt > 1000000 Then Exit Do
            Debug.Print(cnt & " " & xmp / b & " " & ppd / b & " " & CDec(half / b) / 12)
            'If cnt = 16 Then
            '    cnt = cnt
            'End If
        Loop Until low > high

        'Debug.Print(cnt)


        Dim ans As Decimal = low / b '* 12
        Return ans
    End Function
    Function mpActual(ByVal intRate As Decimal, ByVal mths As Decimal)
        Dim n As Decimal = mths 'yrs * 12
        Dim i As Decimal = intRate / 12
        Dim p As Decimal = (1 + i) ^ -n


        Dim ans As Decimal = i / (1 - p)
        Return ans
    End Function
    Function computeInterestRate2(ByVal paymentPer1dollarLoan As Decimal, ByVal mths As Decimal) As Decimal
        'Dim i As Decimal
        If 1 / mths > paymentPer1dollarLoan Then
            MsgBox("Duration is not DECIMAL enough as 1/Mths > PaymentPer1dollarLoan")
            Return 0
        End If

        Dim xmp As Decimal
        Dim low As Decimal = 0.0000001
        Dim high As Decimal = 10
        Dim increment As Decimal = 0.0000001
        Dim half As Decimal
        Dim cnt As Integer = 0

        Dim ppd As Decimal = paymentPer1dollarLoan
        Do
            half = (high + low) / 2
            xmp = mpActual(half, mths)
            If xmp >= ppd Then high = half - increment Else low = half + increment
            cnt += 1
            If cnt > 1000 Then Exit Do
            'Debug.Print(cnt & " " & xmp & " " & ppd & " " & half / 12)
            'If cnt = 16 Then
            '    cnt = cnt
            'End If
        Loop Until low > high

        'Debug.Print(cnt)


        Dim ans As Decimal = low
        'Debug.Print(low)
        Return ans
    End Function
    Function computeMonths(ByVal paymentPer1dollarLoan As Decimal, ByVal Rate As Decimal) As Integer
        'Dim i As Decimal

        Dim xmp As Decimal
        Dim low As Long = 1
        Dim high As Long = 10000
        Dim increment As Long = 1
        Dim half As Long
        Dim cnt As Integer = 0

        Dim ppd As Decimal = paymentPer1dollarLoan
        Do
            half = (high + low) \ 2
            xmp = mpActual(Rate, half)
            If xmp < ppd Then high = half - increment Else low = half + increment
            cnt += 1
            If cnt > 1000 Then Exit Do
            Debug.Print(cnt & " " & xmp & " " & ppd & " " & half)
            'If cnt = 16 Then
            '    cnt = cnt
            'End If
        Loop Until low > high

        'Debug.Print(cnt)


        Dim ans As Integer = low
        'Debug.Print(low)
        Return ans
    End Function
    Function FirstYearPayment2Equal1after30YearsIncreasing5percentPerYear_compute30years(ByVal NumberOfYears As Integer, ByVal startingAmount As Decimal, Optional ByVal AnnualIncrease As Decimal = 0.05) As Decimal
        'startingAmount = 0.015051435
        'Dim y As Decimal = 6246345.56
        Dim x As Decimal = 0
        'Dim xx As Decimal = 0
        Dim i As Integer
        For i = 0 To (NumberOfYears - 1) '29
            'If i = 0 Then
            '    xx = 0.015051435
            '    Debug.Print(xx)
            'Else
            '    xx = xx * 1.05
            '    Debug.Print(xx)
            'End If

            'Debug.Print(0.015051435 * (1.05 ^ i))
            x += startingAmount * ((1 + AnnualIncrease) ^ i)
            'Debug.Print(i & " " & x & " " & startingAmount * (1.05 ^ i))
        Next
        Return x
    End Function
    Function FirstYearPayment2Equal1after30YearsIncreasing5percentPerYear(Optional ByVal NumberOfYears As Integer = 30, Optional ByVal AnnualIncrease As Decimal = 0.05) As Decimal
        '0.01505144
        '0.01505143509
        Dim x As Integer = 100 / NumberOfYears + 1
        Dim startingAmt As Decimal = CDec(x) / 100 '1 / NumberOfYears ' 0.04
        'note: 100/30 = 3 + 1 = 4 /100 =.04

        Dim i As Decimal 'Long
        Dim decrement As Decimal = 0.01
        Dim result As Decimal
        'AnnualIncrease = 0.03
        For i = 1 To 100 '71 iterations
            result = FirstYearPayment2Equal1after30YearsIncreasing5percentPerYear_compute30years(NumberOfYears, startingAmt, AnnualIncrease)
            If Math.Abs(result - 1) < 0.000000001 Then Exit For
            If result > 1 Then
                startingAmt = startingAmt - decrement
                'Debug.Print(startingAmt)
            ElseIf result < 1 Then
                startingAmt = startingAmt + decrement
                decrement = decrement / 10
                'startingAmt = startingAmt + decrement
                'Debug.Print(startingAmt)
            Else
                Exit For
            End If
        Next
        'Debug.Print(i)


        'compute pv
        Dim xPV As Decimal = 0
        Dim intRate As Decimal = 0.0306
        For i = 0 To 29
            xPV += startingAmt * (1.05 ^ i) * (1 + intRate) ^ -i
            'Debug.Print(i & " " & x & " " & startingAmount * (1.05 ^ i))
        Next
        Debug.Print(xPV)

        Return startingAmt
    End Function
    Function FirstYearPayment2Equal1after30YearsIncreasing5percentPerYear_pv(Optional ByVal Years As Integer = 30, Optional ByVal AnnualIncrease As Decimal = 0.05, Optional ByVal interestRate As Decimal = 0.0303) As Decimal
        'Dim Y As Integer = 30
        'Dim AnnualIncrease As Decimal = 0.05 ' 5%
        'Dim interestRate As Decimal = 0.0303

        Dim StartingAmount As Decimal = 0
        StartingAmount = FirstYearPayment2Equal1after30YearsIncreasing5percentPerYear(Years, AnnualIncrease)

        Dim xPV As Decimal = 0

        Dim i As Integer
        For i = 0 To 29
            xPV += StartingAmount * ((1 + AnnualIncrease) ^ i) * (1 + interestRate) ^ -i
            'Debug.Print(i & " " & x & " " & startingAmount * (1.05 ^ i))
        Next
        Return xPV

    End Function
    'Function QString_Ticks2Date()
    '    Select DateAdd(Day, cast(dTicsCompleted / 864000000000 As int), cast('0001-01-01 00:00:00.000000' as datetime2)) from scAppt
    '    select dateAdd(day, cast(dTicsCompleted / 864000000000 As int), cast('0001-01-01' as date)) from scAppt

    '    Dim t As String
    '        t = "cast(dateAdd(day,dTicsCompleted/864000000000,cast('0001-01-01 00:00:00.000000' as datetime2)) as datetime2"

    '        Declare @xDate date
    '    Declare @TicksPerDay bigint =864000000000
    '    Declare @datetime2 datetime2 = '0001-01-01 00:00:00.000000'

    '    Set @xDate= cast(dateadd(day,@Ticks/@TicksPerDay,@datetime2) As Date)
    '    Return @xDate
    'End Function
    Function DaysToToday() As Integer
        Dim x As Integer = DateDiff(DateInterval.Day, #1/1/0001#, Date.Today.Date)
        Return x
    End Function

End Module
