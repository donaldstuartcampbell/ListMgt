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

'Imports System.Data
'Imports System.Data.Sql

''''Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Module Mod_Todo_FileRoutines 'in Module M_ListMgt_Files
    '    Structure ApptTodoRecordType
    '        '  [ID]
    '        ',[dTics]
    '        ',[Msg]
    '        ',[AcctNum]
    '        ',[StrikeOut]
    '        ',[DeleteFlag]
    '        ',[IntegerDate]
    '        ',[scUserID]
    '        ',[apptDateStr]
    '        ',[apptDate]
    '        ',[dTicsOriginal]
    '        ',[dTicsCompleted]


    '        '[ID],[dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag],[DeleteAbsolute],[scUserID],[apptDateStr],[apptDate],[dTicsOriginal],[dTicsCompleted]

    '        '       Dim a As Integer
    '        'a=4+8+84+4+4+4+4+4+27+8+8+8

    '        Dim ID As Integer
    '        Dim dTics As Long
    '        <VBFixedString(84)> Dim Msg As String
    '        Dim AcctNum As Integer
    '        Dim StrikeOut As Integer
    '        Dim DeleteFlag As Integer
    '        Dim DeleteAbsolute As Integer 'IntegerDAte As Integer
    '        Dim scUserID As Integer
    '        <VBFixedString(27)> Dim ApptDateStr As String
    '        Dim apptDate As Long 'Date
    '        Dim dTicsOriginal As Long
    '        Dim dTicsCompleted As Long

    '    End Structure
    'Structure BirthdayRecordType ''in Module M_ListMgt_Files
    '    'id,dtics,msg,desc,descriptionnumber,deleteflag,datestr,month,day,dticsoriginal
    '    Dim ID As Integer
    '    Dim dTics As Long
    '    <VBFixedString(70)> Dim Msg As String
    '    <VBFixedString(14)> Dim Desc As String
    '    Dim DescriptionNumber As Integer
    '    Dim DeleteFlag As Integer
    '    <VBFixedString(10)> Dim DateStr As String
    '    Dim Month As Integer
    '    Dim Day As Integer
    '    Dim dTicsOriginal As Long

    'End Structure
    'Structure file_ApptTodo
    '    Public Shared FileName = "ApptTodo.Fil"
    '    Public Shared FileLen = 167
    'End Structure
    'Structure file_Birthday
    '        Public Shared FileName = "Birthday.Fil"
    '        Public Shared FileLen = 130
    '    End Structure
    '    Structure f_Appt
    '        Public Shared FileName = "Appt.Fil"
    '        Public Shared FileLen = 100
    '    End Structure
    '    Structure timeMsgType 'from: mPrintingGridBW
    '        Dim time As String
    '        Dim msg As String
    '        'Dim deleteFlag As Integer
    '        Dim StrikeOut As Integer
    '    End Structure

    '    Structure LongIntegerType
    '        Dim L As Long
    '        Dim i As Integer
    '    End Structure
    'Structure ApptRecTypeOLD '100 bybytes
    '    'Dim dTics As Date '8
    '    Dim dTics As Int64 'Date '8
    '    <VBFixedString(84)> Dim msg As String
    '    Dim DeleteFlag As Int32 '4
    '    Dim AcctNum As Int32 '4
    'End Structure
    '    Structure ApptRec2Type '100 bybytes
    '        Dim dTics As Int64 'Date '8
    '        <VBFixedString(84)> Dim msg As String
    '        Dim AcctNum As Int32 '4
    '        Dim StrikeOut As Int16 '2 bytes 0 or 1
    '        Dim DeleteFlag As Int16 '2 bytes 0 or 1
    '    End Structure


    '    Function xBuildFullFileName(ByVal filename As String) As String '800 times faster than Combine
    '        'Return ApptGV.ccPath & filename
    '        Dim x As String
    '        'BuildFullFileName
    '        'Dim x As String = NotesLogicGV.filePath & filename

    '        'x = ApptGV.ccPath & filename
    '        'Return x

    '        'Even through this is slower it is the right way to do it!
    '        'Try
    '        '    x = My.Computer.FileSystem.CombinePath(ApptGV.ccPath, filename)
    '        'Catch ex As Exception
    '        '    x = ""
    '        'End Try

    '        x = My.Computer.FileSystem.CombinePath(ApptGV.ccPath, filename)

    '        'Debug.Print(x)
    '        Return x


    '    End Function
    '    Sub setLongType()
    '        'from: getLongInteger_Apptfil 10/23/2015
    '        Dim filename As String = xBuildFullFileName(f_Appt.FileName)
    '        Dim NumRecs As Integer = 0
    '        Dim rec As ApptRec2Type
    '        Using reader As New BinaryReader(File.Open(filename, FileMode.Open)) ' Loop through length of file.

    '            Dim length As Integer = reader.BaseStream.Length
    '            NumRecs = length / f_Appt.FileLen
    '            ReDim ApptGV.LongType(NumRecs)
    '            ''''Dim rec As ApptRec2Type = initApptRec2Type()

    '            Dim i As Integer
    '            Dim cnt As Integer = 0
    '            For i = 1 To NumRecs

    '                ApptGV.LongType(i) = reader.ReadInt64
    '                rec.msg = reader.ReadChars(84) '.ReadInt32()
    '                rec.AcctNum = reader.ReadInt32
    '                rec.StrikeOut = reader.ReadInt16 '20140821
    '                rec.DeleteFlag = reader.ReadInt16
    '                'ApptGV.LongType(i) += i

    '                'If ApptGV.LongType(i) <= ApptGV.HighBDayTics Then
    '                '    If ApptGV.LongType(i) >= ApptGV.LowBDayTics AndAlso rec.DeleteFlag = 0 Then
    '                '        cnt += 1
    '                '        ApptGV.LongType(cnt) = ApptGV.LongType(i) + i 'the i = location in file
    '                '    End If
    '                'ElseIf Mid(rec.msg, 1, 1) <> Chr(0) AndAlso rec.DeleteFlag = 0 AndAlso Trim(rec.msg) <> "" Then
    '                '    cnt += 1
    '                '    ApptGV.LongType(cnt) = ApptGV.LongType(i) + i 'the i = location in file
    '                'End If
    '            Next
    '            If NumRecs <> cnt Then ReDim Preserve ApptGV.LongType(cnt)
    '        End Using
    '        Array.Sort(ApptGV.LongType)
    '    End Sub
    '    Sub test_allAppt()
    '        Dim cnt As Integer = 0
    '        Dim i As Integer
    '        Dim recs(0) As ApptRec2Type
    '        Dim NumRecs As Integer

    '        'Dim FullFileName As String = "c:\todo\Appt.fil" '"d:\todo\appt.fil" ' "f:\bu\Appt.fil" ' xBuildFullFileName(f_Appt.FileName)
    '        Dim FullFileName As String = xBuildFullFileName(f_Appt.FileName)
    '        'Debug.Print(FullFileName)

    '        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
    '            Dim length As Integer = reader.BaseStream.Length
    '            NumRecs = length / f_Appt.FileLen
    '            'Debug.Print(NumRecs)
    '            ReDim recs(NumRecs)
    '            ReDim ApptGV.AllAppts_LongInteger(NumRecs)
    '            'startTime()
    '            For i = 1 To NumRecs
    '                With recs(i)
    '                    .dTics = reader.ReadInt64
    '                    .msg = reader.ReadChars(84)
    '                    .AcctNum = reader.ReadInt32
    '                    .StrikeOut = reader.ReadInt16
    '                    .DeleteFlag = reader.ReadInt16

    '                    ApptGV.AllAppts_LongInteger(i).L = .dTics + i
    '                    ApptGV.AllAppts_LongInteger(i).i = i

    '                    'If .dTics <= ApptGV.HighBDayTics Then
    '                    '    If .DeleteFlag = 0 Then cnt += 1 : recs(cnt) = recs(i)
    '                    'ElseIf Mid(.msg, 1, 1) <> Chr(0) Then
    '                    '    cnt += 1 : recs(cnt) = recs(i)
    '                    'End If
    '                End With
    '            Next
    '            'endTime()
    '            'End
    '            'ReDim Preserve recs(cnt)
    '            'Return recs
    '        End Using
    '        'End
    '        'For i = 1 To NumRecs
    '        '    With recs(i)
    '        '        Debug.Print("{0}-{1}-{2}-{3}-{4}-{5}", i, .AcctNum, .DeleteFlag, .StrikeOut, New Date(.dTics), .msg)
    '        '        Debug.Print("")

    '        '    End With


    '        'Next

    '        'End

    '        ''''ApptGV.AllAppts_LongInteger = getLongInteger_Apptfil() 'as of 11/11/2015
    '        'new 11/26/2016 usage
    '        QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))

    '        For i = 1 To NumRecs
    '            With recs(i)
    '                'Debug.Print(("{0}-{1}-{2}-{3}-{4}-{5}", i, .AcctNum, .DeleteFlag, .StrikeOut, New Date(.dTics), .msg))
    '                'Debug.Print("{0}-{1}-{2}-{3}-{4}-{5}", i, .AcctNum, .DeleteFlag, .StrikeOut, New Date(.dTics), .msg)
    '                'Debug.Print("")
    '                Debug.Print(recs(i).dTics & " - " & New Date(recs(i).dTics) & " --- " & New Date(ApptGV.AllAppts_LongInteger(i).L) & " - " & ApptGV.AllAppts_LongInteger(i).i &
    '                            " - " & ApptGV.AllAppts_LongInteger(i).L & " - " & ApptGV.AllAppts_LongInteger(i).L Mod TimeSpan.TicksPerSecond)

    '            End With


    '        Next
    '    End Sub
    '    Sub Set_ApptGvAllAppts_LongInteger()
    '        Dim cnt As Integer = 0
    '        Dim i As Integer
    '        Dim rec As ApptRec2Type
    '        Dim NumRecs As Integer

    '        Dim FullFileName As String = xBuildFullFileName(f_Appt.FileName)

    '        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
    '            Dim length As Integer = reader.BaseStream.Length
    '            NumRecs = length / f_Appt.FileLen

    '            ReDim ApptGV.AllAppts_LongInteger(NumRecs)
    '            For i = 1 To NumRecs
    '                With rec
    '                    '.dTics = reader.ReadInt64

    '                    ApptGV.AllAppts_LongInteger(i).L = reader.ReadInt64

    '                    .msg = reader.ReadChars(84)
    '                    .AcctNum = reader.ReadInt32
    '                    .StrikeOut = reader.ReadInt16
    '                    .DeleteFlag = reader.ReadInt16

    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics
    '                    ApptGV.AllAppts_LongInteger(i).i = i

    '                End With
    '            Next
    '        End Using

    '    End Sub

    '    Sub test_allAppt2()
    '        Dim cnt As Integer = 0
    '        Dim i As Integer
    '        Dim rec As ApptRec2Type
    '        Dim NumRecs As Integer

    '        'Dim FullFileName As String = "c:\todo\Appt.fil" '"d:\todo\appt.fil" ' "f:\bu\Appt.fil" ' xBuildFullFileName(f_Appt.FileName)
    '        Dim FullFileName As String = xBuildFullFileName(f_Appt.FileName)
    '        'Debug.Print(FullFileName)

    '        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
    '            Dim length As Integer = reader.BaseStream.Length
    '            NumRecs = length / f_Appt.FileLen
    '            'Debug.Print(NumRecs)

    '            ReDim ApptGV.AllAppts_LongInteger(NumRecs)
    '            'startTime()
    '            For i = 1 To NumRecs
    '                With rec
    '                    '.dTics = reader.ReadInt64

    '                    ApptGV.AllAppts_LongInteger(i).L = reader.ReadInt64

    '                    .msg = reader.ReadChars(84)
    '                    .AcctNum = reader.ReadInt32
    '                    .StrikeOut = reader.ReadInt16
    '                    .DeleteFlag = reader.ReadInt16

    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics
    '                    ApptGV.AllAppts_LongInteger(i).i = i

    '                    'If .dTics <= ApptGV.HighBDayTics Then
    '                    '    If .DeleteFlag = 0 Then cnt += 1 : recs(cnt) = recs(i)
    '                    'ElseIf Mid(.msg, 1, 1) <> Chr(0) Then
    '                    '    cnt += 1 : recs(cnt) = recs(i)
    '                    'End If
    '                End With
    '            Next
    '            'endTime()
    '            'End
    '            'ReDim Preserve recs(cnt)
    '            'Return recs
    '        End Using
    '        'End
    '        'For i = 1 To NumRecs
    '        '    With recs(i)
    '        '        Debug.Print("{0}-{1}-{2}-{3}-{4}-{5}", i, .AcctNum, .DeleteFlag, .StrikeOut, New Date(.dTics), .msg)
    '        '        Debug.Print("")

    '        '    End With


    '        'Next

    '        'End

    '        ''''ApptGV.AllAppts_LongInteger = getLongInteger_Apptfil() 'as of 11/11/2015
    '        'new 11/26/2016 usage
    '        'QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))
    '    End Sub
    '    Sub QuickSort_BirthdayRecords(ByRef c() As BirthdayRecordType, ByVal First As Integer, ByVal Last As Integer)
    '        If First >= Last Then Exit Sub

    '        Dim Low, High As Integer 'Always Integer or Long
    '        Dim MidValue As String 'Long 'change
    '        Dim T As BirthdayRecordType 'change

    '        Low = First : High = Last : MidValue = c((Low + High) \ 2).DateStr '.dTics
    '        Do
    '            While (c(Low).DateStr < MidValue) : Low += 1 : End While
    '            While (c(High).DateStr > MidValue) : High -= 1 : End While
    '            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '        Loop While Low <= High
    '        If First < High Then QuickSort_BirthdayRecords(c, First, High)
    '        If Low < Last Then QuickSort_BirthdayRecords(c, Low, Last)

    '    End Sub
    '    Sub QuickSort_ApptTodoRecords(ByRef c() As ApptTodoRecordType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '        'new Oct 1 2019
    '        If First >= Last Then Exit Sub

    '        Dim Low, High As Integer 'Long 'Always Integer or Long
    '        Dim MidValue As Long 'Integer
    '        Dim T As ApptTodoRecordType 'LongIntegerType

    '        Low = First : High = Last : MidValue = c((Low + High) \ 2).dTics
    '        Do
    '            While (c(Low).dTics < MidValue) : Low += 1 : End While
    '            While (c(High).dTics > MidValue) : High -= 1 : End While
    '            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '        Loop While Low <= High
    '        If First < High Then QuickSort_ApptTodoRecords(c, First, High)
    '        If Low < Last Then QuickSort_ApptTodoRecords(c, Low, Last)
    '    End Sub
    '    Sub test_QuickSort_ApptTodoRecords_bydTicsCompleted()
    '        Dim L As Long = #10/1/2019#.Ticks
    '        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
    '        Dim a(1000000) As ApptTodoRecordType
    '        Dim cnt As Integer = 0
    '        Dim i As Integer
    '        For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
    '            If ApptGV.All_ApptTodoRecords(i).dTicsCompleted > 0 AndAlso ApptGV.All_ApptTodoRecords(i).dTicsOriginal > L Then
    '                cnt += 1
    '                a(cnt) = ApptGV.All_ApptTodoRecords(i)
    '            End If
    '        Next
    '        ReDim Preserve a(cnt)
    '        Dim b() As ApptTodoRecordType = a.Clone


    '        For i = 1 To Math.Min(cnt, 10)
    '            'With ApptGV.All_ApptTodoRecords(i)
    '            With a(i)
    '                Debug.Print(.ApptDateStr & " " & " " & .DeleteFlag & " " & New Date(.dTicsCompleted) & " " & .Msg)
    '            End With

    '        Next
    '        Debug.Print("--------")

    '        QuickSort_ApptTodoRecords_bydTicsCompleted(a, 1, cnt)


    '        For i = 1 To Math.Min(cnt, 10)
    '            'With ApptGV.All_ApptTodoRecords(i)
    '            With a(i)
    '                Debug.Print(.ApptDateStr & " " & " " & .DeleteFlag & " " & New Date(.dTicsCompleted) & " " & .Msg)
    '            End With

    '        Next
    '        Debug.Print("--------")
    '        QuickSort_ApptTodoRecords_descending_bydTicsCompleted(b, 1, cnt)
    '        For i = 1 To Math.Min(cnt, 10)
    '            'With ApptGV.All_ApptTodoRecords(i)
    '            With b(i)
    '                Debug.Print(.ApptDateStr & " " & " " & .DeleteFlag & " " & New Date(.dTicsCompleted) & " " & .Msg)
    '            End With

    '        Next

    '    End Sub
    '    Sub QuickSort_ApptTodoRecords_bydTicsCompleted(ByRef c() As ApptTodoRecordType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '        'new Oct 1 2019
    '        If First >= Last Then Exit Sub

    '        Dim Low, High As Integer 'Long 'Always Integer or Long
    '        Dim MidValue As Long 'Integer
    '        Dim T As ApptTodoRecordType 'LongIntegerType

    '        'Low = First : High = Last : MidValue = c((Low + High) \ 2).dTics
    '        Low = First : High = Last : MidValue = c((Low + High) \ 2).dTicsCompleted 'CHANGE 1
    '        Do
    '            While (c(Low).dTicsCompleted < MidValue) : Low += 1 : End While 'CHANGE 2
    '            While (c(High).dTicsCompleted > MidValue) : High -= 1 : End While 'CHANGE 3
    '            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '        Loop While Low <= High
    '        If First < High Then QuickSort_ApptTodoRecords_bydTicsCompleted(c, First, High) 'CHANGE 4
    '        If Low < Last Then QuickSort_ApptTodoRecords_bydTicsCompleted(c, Low, Last) 'CHANGE 5
    '    End Sub
    '    Sub QuickSort_ApptTodoRecords_descending_bydTicsCompleted(ByRef c() As ApptTodoRecordType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '        'new Oct 1 2019
    '        If First >= Last Then Exit Sub

    '        Dim Low, High As Integer 'Long 'Always Integer or Long
    '        Dim MidValue As Long 'Integer
    '        Dim T As ApptTodoRecordType 'LongIntegerType

    '        'Low = First : High = Last : MidValue = c((Low + High) \ 2).dTics
    '        Low = First : High = Last : MidValue = c((Low + High) \ 2).dTicsCompleted
    '        Do
    '            'While (c(Low).dTicsCompleted < MidValue) : Low += 1 : End While ' CHANGE 1
    '            'While (c(High).dTicsCompleted > MidValue) : High -= 1 : End While ' CHANGE 2
    '            'If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1

    '            While (c(Low).dTicsCompleted > MidValue) : Low += 1 : End While ' CHANGE 1
    '            While (c(High).dTicsCompleted < MidValue) : High -= 1 : End While 'CHANGE 2
    '            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '        Loop While Low <= High
    '        If First < High Then QuickSort_ApptTodoRecords_descending_bydTicsCompleted(c, First, High)
    '        If Low < Last Then QuickSort_ApptTodoRecords_descending_bydTicsCompleted(c, Low, Last)
    '    End Sub

    '    Sub QuickSort_LongInteger_Long(ByRef c() As LongIntegerType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '        'new 11/26/2016
    '        If First >= Last Then Exit Sub

    '        Dim Low, High As Integer 'Long 'Always Integer or Long
    '        Dim MidValue As Long 'Integer
    '        Dim T As LongIntegerType

    '        Low = First : High = Last : MidValue = c((Low + High) \ 2).L
    '        Do
    '            While (c(Low).L < MidValue) : Low += 1 : End While
    '            While (c(High).L > MidValue) : High -= 1 : End While
    '            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '        Loop While Low <= High
    '        If First < High Then QuickSort_LongInteger_Long(c, First, High)
    '        If Low < Last Then QuickSort_LongInteger_Long(c, Low, Last)
    '    End Sub
    '    Sub test_QuickSort_Integer()
    '        Dim i As Integer
    '        Dim nItems As Integer = 1000000 ' 1000000 '1000000 '00 '1000000
    '        Dim SampleArray(nItems) As Integer
    '        Randomize() 'adding the 1 cause the same numbers to be generated each time!
    '        For i = 1 To nItems
    '            SampleArray(i) = Rnd() * nItems
    '            'Debug.Print(SampleArray(i))
    '        Next
    '        'Debug.Print("====")
    '        Dim b() As Integer = SampleArray.Clone

    '        startTime()
    '        'QuickSort_IntegerDESC(SampleArray, 1, nItems)
    '        QuickSort_Integer(SampleArray, 1, nItems)
    '        Exit Sub
    '        endTime()

    '        startTime()
    '        'QuickSort_IntegerDesc2(b, 1, nItems)
    '        QuickSort_Integer2(b, 1, nItems)
    '        endTime()

    '        For i = 1 To nItems
    '            If SampleArray(i) <> b(i) Then
    '                i = i
    '            End If
    '        Next
    '        Exit Sub
    '        'For i = 1 To 8
    '        '    Debug.Print(i & " " & SampleArray(i))
    '        'Next
    '        'Debug.Print("==============")

    '        startTime()
    '        QuickSort_Integer(SampleArray, 1, nItems) '0.15 seconds
    '        endTime()

    '        'startTime()
    '        'MergeSort(b, 1, nItems)
    '        'endTime()

    '        'For i = 1 To nItems
    '        '    If SampleArray(i) <> b(i) Then
    '        '        i = i
    '        '    End If
    '        'Next

    '        QuickSort_IntegerDESC(b, 1, nItems)

    '        Exit Sub


    '        'For i = 1 To 8
    '        '    Debug.Print(i & " " & SampleArray(i))
    '        'Next
    '        'Debug.Print("==============")

    '        startTime()
    '        QuickSort_IntegerDESC(b, 1, nItems)
    '        endTime()

    '        'For i = 993 To nItems ' 1 To 8
    '        '    Debug.Print(i & " " & b(i))
    '        'Next

    '        For i = 2 To nItems
    '            If SampleArray(i) < SampleArray(i - 1) Then
    '                i = i
    '            End If
    '        Next
    '        For i = 2 To nItems
    '            If b(i) > b(i - 1) Then
    '                i = i
    '            End If
    '        Next

    '    End Sub
    '    Sub QuickSort_Integer2(ByRef numbers As Integer(), ByVal low As Integer, ByVal high As Integer) 'from quickSortInAscendingOrder

    '        Dim i As Integer = low
    '        Dim j As Integer = high
    '        Dim temp As Integer
    '        Dim middle As Integer = numbers((low + high) / 2)

    '        While i < j
    '            'While numbers(i) > middle : i += 1 : End While
    '            'While numbers(j) < middle : j -= 1 : End While
    '            'If j >= i Then temp = numbers(i) : numbers(i) = numbers(j) : numbers(j) = temp : i += 1 : j -= 1

    '            'ascending order (ONLY CHANGES - next 3 lines)
    '            While numbers(i) < middle : i += 1 : End While
    '            While numbers(j) > middle : j -= 1 : End While
    '            If i <= j Then temp = numbers(i) : numbers(i) = numbers(j) : numbers(j) = temp : i += 1 : j -= 1
    '            'above is the way the code was written but BELOW results in the same result and represents NO CHANGE code in Descending Order
    '            'If j >= i Then temp = numbers(i) : numbers(i) = numbers(j) : numbers(j) = temp : i += 1 : j -= 1

    '        End While

    '        If low < j Then QuickSort_Integer2(numbers, low, j)
    '        If i < high Then QuickSort_Integer2(numbers, i, high)
    '    End Sub
    '    Sub QuickSort_Integer(ByRef c() As Integer, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '        'new 11/26/2016
    '        If First >= Last Then Exit Sub

    '        Dim Low, High As Integer 'Long 'Always Integer or Long
    '        Dim MidValue As Integer
    '        Dim T As Integer

    '        Low = First : High = Last : MidValue = c((Low + High) \ 2)

    '        'in testing Low is never > High
    '        'If Low > High Then
    '        '    Low = Low
    '        'End If

    '        Do
    '            While (c(Low) < MidValue) : Low += 1 : End While
    '            While (c(High) > MidValue) : High -= 1 : End While
    '            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '        Loop While Low <= High
    '        If First < High Then QuickSort_Integer(c, First, High)
    '        If Low < Last Then QuickSort_Integer(c, Low, Last)
    '    End Sub
    '    Sub QuickSort_IntegerDesc2(ByVal numbers As Integer(), ByVal low As Integer, ByVal high As Integer) 'from quickSortInDescendingOrder
    '        'Public Shared Sub quickSortInDescendingOrder(ByVal numbers As Integer(), ByVal low As Integer, ByVal high As Integer)
    '        'https://www.careerbless.com/samplecodes/java/beginners/sorting/QuickSortDescending.php

    '        Dim i As Integer = low
    '        Dim j As Integer = high
    '        Dim temp As Integer
    '        Dim middle As Integer = numbers((low + high) / 2)

    '        While i < j
    '            While numbers(i) > middle : i += 1 : End While
    '            While numbers(j) < middle : j -= 1 : End While
    '            If j >= i Then temp = numbers(i) : numbers(i) = numbers(j) : numbers(j) = temp : i += 1 : j -= 1
    '        End While

    '        If low < j Then QuickSort_IntegerDesc2(numbers, low, j)
    '        If i < high Then QuickSort_IntegerDesc2(numbers, i, high)
    '    End Sub
    '    Sub QuickSort_IntegerDESC(ByRef c() As Integer, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '        'to make it work in descending order only 2 lines (as indicated below) need to be changed! amazing!
    '        'new 11/26/2016
    '        If First >= Last Then Exit Sub

    '        'Static cnt As Integer
    '        'cnt += 1
    '        'Debug.Print(cnt)

    '        '868 for 1000 element array
    '        '8685 for 10000 element array
    '        '86293 for 100000 element array


    '        Dim Low, High As Integer 'Long 'Always Integer or Long
    '        Dim MidValue As Integer 'Long 
    '        Dim T As Integer

    '        Low = First : High = Last : MidValue = c((Low + High) \ 2)
    '        Do
    '            While (c(Low) > MidValue) : Low += 1 : End While 'change from < to >
    '            While (c(High) < MidValue) : High -= 1 : End While 'change from > to <
    '            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '        Loop While Low <= High

    '        'Debug.Print(First & " " & Low & " " & Last & " " & High)

    '        If First < High Then QuickSort_IntegerDESC(c, First, High)
    '        If Low < Last Then QuickSort_IntegerDESC(c, Low, Last)
    '    End Sub
    '    Sub MergeSort(ByVal array() As Integer, lowIndex As Integer, highIndex As Integer)

    '        If (lowIndex < highIndex) Then
    '            Dim midIndex = Math.Floor((lowIndex + highIndex) / 2)    'ensure that we get integer result i.e. 5/2 yields 2
    '            'Recursively break apart original array until the tree bottoms out
    '            MergeSort(array, lowIndex, midIndex)
    '            MergeSort(array, midIndex + 1, highIndex)
    '            'Then merge our single element arrays
    '            Merge(array, lowIndex, midIndex, highIndex)
    '        End If

    '    End Sub

    '    Sub Merge(ByVal array() As Integer, lowIndex As Integer, midIndex As Integer, highIndex As Integer)
    '        'creating 2 sub arrays for left hand and right hand part of merge

    '        Dim n1 = midIndex - lowIndex + 1
    '        Dim n2 = highIndex - midIndex

    '        Dim L(n1) As Integer
    '        Dim R(n2) As Integer

    '        'creating index variable to keep track of final answer
    '        Dim k As Integer = lowIndex

    '        Dim counterI = 0
    '        Dim counterJ = 0
    '        'Fill each of the two arrays declared

    '        While (counterI < n1)
    '            L(counterI) = array(lowIndex + counterI)
    '            counterI = counterI + 1
    '        End While


    '        While (counterJ < n2)
    '            R(counterJ) = array(midIndex + 1 + counterJ)
    '            counterJ = counterJ + 1
    '        End While

    '        'Reset index variables
    '        k = lowIndex
    '        Dim i As Integer = 0
    '        Dim j As Integer = 0

    '        'Go through and compare the two subarrays and fill index k of our answer
    '        'with the lower value until one of the subarrays is empty
    '        While (i < n1 And j < n2)

    '            If (L(i) <= R(j)) Then
    '                array(k) = L(i)
    '                i = i + 1
    '            Else
    '                array(k) = R(j)
    '                j = j + 1
    '            End If
    '            k = k + 1
    '        End While

    '        'If one array is empty we go ahead and fill our answer with remaining array
    '        'this removes the sentinels from example (I was struggling with index bounds)

    '        While (i < n1)
    '            array(k) = L(i)
    '            i = i + 1
    '            k = k + 1
    '        End While

    '        While (j < n2)
    '            array(k) = R(j)
    '            j = j + 1
    '            k = k + 1
    '        End While

    '        ''Print our answer
    '        'Console.WriteLine("The sorted array using Merge Sort is: ")
    '        'For index2 As Integer = 0 To array.Length - 1
    '        '    Console.Write(array(index2) & " ")
    '        'Next
    '    End Sub
    '    Sub sort_LongInteger(ByRef a() As LongIntegerType)
    '        'System.Array.Sort(a, Function(x As LongIntegerType, y As LongIntegerType) Number.Compare(x.L, y.L)) 'lambda function
    '        'Array.Sort(a, Function(x As LongIntegerType, y As LongIntegerType) Number.Compare(x.L, y.L)) 'lambda function

    '        'Array.Sort(a, (Function(x As LongIntegerType, y As LongIntegerType) y.L > x.L)) 'lambda
    '        'Array.Sort(a, (Function(x As LongIntegerType, y As LongIntegerType) y.L >= x.L)) 'lambda

    '        'Array.Sort(a, (Function(x As LongIntegerType, y As LongIntegerType) CompareMethod.Binary y.)) 'lambda

    '        'Array.Sort(a, Function(
    '        '    x As LongIntegerType,
    '        '    y As LongIntegerType) Compare(x.L, y.L)) 'lambda

    '        'a = a.OrderBy(Function(c) c.L)

    '        'see: Imports System.Linq.Expressions

    '        'cars.Sort(Function(c1,c2) c1.MPH.CompareTo(c2.MPH))

    '        'a.Sort(Function(x As LongIntegerType, y As LongIntegerType) x.L.CompareTo(y.L))
    '        'Array.Sort(a(Function(x, y) x.L.CompareTo(y.L))
    '        Array.Sort(a, (Function(x As LongIntegerType, y As LongIntegerType) x.L.CompareTo(y.L))) 'lambda
    '    End Sub

    '    Sub test_Standalone_Todo2()
    '        Set_ApptGvAllAppts_LongInteger()

    '        Dim i As Integer
    '        Dim nrecs As Integer = UBound(ApptGV.AllAppts_LongInteger)

    '        Dim find1 As Integer
    '        Dim find2 As Integer

    '        For i = 1 To nrecs

    '        Next
    '    End Sub
    '    Sub test_Standalone_Todo()
    '        Dim r As ApptTodoRecordType
    '        Debug.Print(Len(r))


    '        Dim i As Integer = 0
    '        startTime()

    '        'setLongType()
    '        test_allAppt()
    '        'test_allAppt2()
    '        endTime()

    '        Dim b() As LongIntegerType = ApptGV.AllAppts_LongInteger.Clone

    '        '===
    '        'For i = 1 To UBound(b)
    '        '    'If b(i).i <> ApptGV.AllAppts_LongInteger(i).i Then

    '        '    'Debug.Print(i & " -: " & New Date(b(i).L) & " --- " & New Date(ApptGV.AllAppts_LongInteger(i).L))
    '        '    Debug.Print(i & " -: " & New Date(b(i).L) & ":" & b(i).i.ToString & " --- " & New Date(ApptGV.AllAppts_LongInteger(i).L) & ":" & ApptGV.AllAppts_LongInteger(i).i.ToString)

    '        '    'End If

    '        'Next
    '        '===

    '        startTime()
    '        QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))
    '        endTime()

    '        'startTime()
    '        'sort_LongInteger(b)
    '        'endTime()


    '        'For i = 1 To UBound(b)
    '        '    'If b(i).i <> ApptGV.AllAppts_LongInteger(i).i Then

    '        '    'Debug.Print(i & " -: " & New Date(b(i).L) & " --- " & New Date(ApptGV.AllAppts_LongInteger(i).L))
    '        '    Debug.Print(i & " -: " & New Date(b(i).L) & ":" & b(i).i.ToString & " --- " & New Date(ApptGV.AllAppts_LongInteger(i).L) & ":" & ApptGV.AllAppts_LongInteger(i).i.ToString)

    '        '    'End If

    '        'Next
    '    End Sub
    '#Region "Add or Update"
    '    Sub AddApptRecord2(ByVal rec As ApptRec2Type) 'use for recurring records  'new 9/5/2015
    '        Dim ff As Integer = FreeFile()
    '        Dim filename As String = xBuildFullFileName(f_Appt.FileName)
    '        FileOpen(ff, filename, OpenMode.Random, , , f_Appt.FileLen)

    '        Dim xCount As Integer = LOF(ff) / f_Appt.FileLen + 1

    '        FilePut(ff, rec, xCount)
    '        FileClose(ff)

    '        'new
    '        ReDim Preserve ApptGV.AllAppts_LongInteger(xCount)
    '        With ApptGV.AllAppts_LongInteger(xCount)
    '            .L = rec.dTics
    '            .i = xCount
    '        End With
    '        QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))

    '        'set_AllAppts_LongInteger()
    '    End Sub
    '    Sub UpdateApptRecord(ByVal rec As ApptRec2Type) 'use for recurring records 'new 9/5/2015
    '        Dim xLocation As Integer = getLocationOfRecurRecord(rec)
    '        If xLocation = 0 Then
    '            MsgBox("Could not update record as it could not be found", MsgBoxStyle.Critical, "Update Failed")
    '            Exit Sub
    '        End If
    '        Dim ff As Integer = FreeFile()
    '        Dim filename As String = xBuildFullFileName(f_Appt.FileName)
    '        FileOpen(ff, filename, OpenMode.Random, , , f_Appt.FileLen)

    '        'Dim xCount As Integer = LOF(ff) / f_Appt.FileLen + 1

    '        FilePut(ff, rec, xLocation)
    '        FileClose(ff)

    '        'new
    '        QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))
    '        'set_AllAppts_LongInteger()
    '    End Sub
    '    Function getLocationOfRecurRecord(ByVal Rec2Find As ApptRec2Type) As Integer 'new 9/5/2015
    '        'from: getHighestDailyRecurTicks
    '        '============================
    '        'uses: FastRead_Apptfil logic
    '        Dim xLocation As Integer = 0

    '        Dim i As Integer
    '        Dim NumRecs As Integer
    '        Dim FullFileName As String = xBuildFullFileName(f_Appt.FileName)
    '        Dim rec As ApptRec2Type
    '        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open), System.Text.Encoding.ASCII)
    '            Dim length As Long = reader.BaseStream.Length
    '            NumRecs = length / f_Appt.FileLen
    '            For i = 1 To NumRecs
    '                With rec
    '                    .dTics = reader.ReadInt64
    '                    .msg = reader.ReadChars(84)
    '                    .AcctNum = reader.ReadInt32
    '                    .StrikeOut = reader.ReadInt16
    '                    .DeleteFlag = reader.ReadInt16

    '                    If .dTics = Rec2Find.dTics Then
    '                        xLocation = i
    '                        Exit For
    '                    End If
    '                End With
    '            Next
    '        End Using
    '        Return xLocation
    '    End Function
    '    Function biSearch_E_AllAppt(ByVal search As Long) As Integer
    '        'from: biSearch_E_iMeanAll
    '        Dim LocInFile As Integer = 0

    '        Dim HH As Integer = UBound(ApptGV.AllAppts_LongInteger)
    '        Dim high As Integer = HH
    '        Dim low As Integer = 1
    '        Dim half As Integer
    '        Do
    '            half = (high + low) \ 2
    '            If search > ApptGV.AllAppts_LongInteger(half).L Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '        Loop Until low > high
    '        If low <= HH AndAlso search = ApptGV.AllAppts_LongInteger(low).L Then LocInFile = ApptGV.AllAppts_LongInteger(low).i
    '        Return LocInFile
    '    End Function
    '    Function FileGetApptRec(ByVal LocInFile As Integer) As ApptRec2Type
    '        If LocInFile = 0 Then
    '            MsgBox("In: FileGetApptRec - LocInFile = 0 - continue to see where call was made from")
    '            LocInFile = LocInFile
    '            Exit Function
    '        End If
    '        Dim rec As ApptRec2Type = initApptRec2Type()
    '        Dim ff As Integer = FreeFile()
    '        Dim FileName As String = xBuildFullFileName(f_Appt.FileName)
    '        FileOpen(ff, FileName, OpenMode.Random, , , f_Appt.FileLen)
    '        FileGet(ff, rec, LocInFile)
    '        FileClose(ff)
    '        Return rec
    '    End Function
    '    Function initApptRec2Type() As ApptRec2Type
    '        Dim x As ApptRec2Type
    '        With x
    '            .dTics = 0
    '            .msg = ""
    '            .AcctNum = 0
    '            .StrikeOut = 0
    '            .DeleteFlag = 0
    '        End With
    '        Return x
    '    End Function
    '#End Region
    '#Region "NewRecord"
    '    Sub Create_ApptTodo_File_ifDoesNotExists()
    '        Dim x As String = "C:\todo\ApptTodo.Fil"
    '        Kill(x)
    '        Dim filepath As String = x
    '        If Not System.IO.File.Exists(filepath) Then
    '            System.IO.File.Create(filepath).Dispose()
    '        End If
    '    End Sub

    '    Sub addnewapptTodoRecord()

    '        'Dim rec As ApptTodoRecordType
    '        Dim rec As ApptTodoRecordType = initApptTodoRecord()
    '        With rec
    '            '.dTics = addnewapptTodoRecord_getNewDTick(ByVal xTicks As Long) 'sets the counter element

    '            .ID = ApptGV.All_ApptTodoRecords_Cnt + 1
    '            '.dTics = .dTics

    '            .Msg = .Msg
    '            .AcctNum = .AcctNum
    '            .StrikeOut = .StrikeOut
    '            .DeleteFlag = .DeleteFlag
    '            .DeleteAbsolute = 0
    '            .scUserID = .scUserID
    '            .ApptDateStr = .ApptDateStr
    '            .apptDate = (.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay 'xDate 'CDate(Mid(.apptDate, 1, 10)) 'CDate(.apptDate) 'Date.Today.Date '.apptDate
    '            .dTicsOriginal = .dTicsOriginal
    '            .dTicsCompleted = .dTicsCompleted
    '        End With
    '        'biSearch_GTE_AllAppts_Long
    '    End Sub
    '    Function addnewapptTodoRecord_getNewDTick(ByVal xTicks As Long) As Long ', ByRef c() As ApptTodoRecordType, ByVal high As Integer) As Long
    '        Dim xDTic As Long = 0
    '        Dim yyyymmddhhmm As Long = (xTicks \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute
    '        Dim NextMinute As Long = yyyymmddhhmm + TimeSpan.TicksPerMinute
    '        Dim LocInFile As Integer

    '        'LocInFile = biSearch_GTE_ApptTodoRecords_test(c, NextMinute, high) - 1
    '        LocInFile = biSearch_GTE_AllAppts_Long(NextMinute) - 1

    '        Dim dif As Long = ApptGV.All_ApptTodoRecords(LocInFile).dTics - yyyymmddhhmm
    '        If dif > 0 Then
    '            xDTic = yyyymmddhhmm + CInt(dif + 1)
    '        Else
    '            xDTic = yyyymmddhhmm + 1
    '        End If
    '        Return xDTic
    '    End Function
    '    Sub moveExisting2New_ApptTodoFile2()
    '        startTime() '524.751714

    '        Create_ApptTodo_File_ifDoesNotExists()
    '        Dim Q As String
    '        Q = "SELECT [ID],[dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag],[IntegerDate],[scUserID],[apptDateStr],[apptDate],[dTicsOriginal],[dTicsCompleted] FROM [Todo].[dbo].[Appt] Order by id"
    '        'where substring(apptdatestr,1,10)<>'0005/01/01' ORDER BY dtics"
    '        Dim recs() As ApptRecAllType 'ApptRecType
    '        recs = getApptRecords_AllFields(Q)
    '        Dim i As Integer
    '        Dim nRecs As Integer = UBound(recs)

    '        'For i = 1 To nRecs
    '        '    Debug.Print(i & " " & recs(i).dTics)
    '        'Next


    '        Dim NewRecs(nRecs) As ApptTodoRecordType

    '        For i = 1 To nRecs
    '            With recs(i)
    '                'Debug.Print(.apptDate & " - " & i & " - " & .ID & " " & .msg)
    '                Dim xDate As Date
    '                xDate = New Date((.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay).Date
    '                'xDate = Switch([DateTime], .dTics, New Date(.dTics).Date)

    '                NewRecs(i).ID = i
    '                NewRecs(i).dTics = .dTics

    '                NewRecs(i).Msg = .msg
    '                NewRecs(i).AcctNum = .AcctNum
    '                NewRecs(i).StrikeOut = .StrikeOut
    '                NewRecs(i).DeleteFlag = .DeleteFlag
    '                NewRecs(i).DeleteAbsolute = .IntegerDate
    '                NewRecs(i).scUserID = .scUserID
    '                NewRecs(i).ApptDateStr = .apptDateStr
    '                NewRecs(i).apptDate = (.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay 'xDate 'CDate(Mid(.apptDate, 1, 10)) 'CDate(.apptDate) 'Date.Today.Date '.apptDate
    '                NewRecs(i).dTicsOriginal = .dTicsOriginal
    '                NewRecs(i).dTicsCompleted = .dTicsCompleted
    '                'AddApptTodoRecord(NewRecs(i))

    '                'Dim rrec As ApptTodoRecordType
    '                'rrec = getApptTodoRecord(i)
    '                'If rrec.dTics <> .dTics Then
    '                '    Debug.Print(NewRecs(i).dTics & " orig=" & .dTics & " new=" & rrec.dTics)
    '                '    i = i
    '                'End If



    '                'Debug.Print(NewRecs(i).apptDate)
    '            End With


    '        Next

    '        'build big file 2700 time above
    '        'Dim j As Integer
    '        'For j = 1 To 2700
    '        '    For i = 1 To nRecs
    '        '        AddApptTodoRecord(NewRecs(i))
    '        '    Next
    '        'Next


    '        '=============3.6 seconds
    '        'Dim ff As Integer = FreeFile()
    '        'Dim filename As String = xBuildFullFileName(file_ApptTodo.FileName)
    '        'FileOpen(ff, filename, OpenMode.Random, , , file_ApptTodo.FileLen)

    '        Dim NewRecsArray(200000) As ApptTodoRecordType

    '        Dim xCount As Integer = 0
    '        Dim highCnt As Integer = 0
    '        Dim j As Integer
    '        For j = 1 To 2700
    '            For i = 1 To nRecs
    '                xCount += 1
    '                If xCount Mod 1000 = 0 Then Debug.Print(xCount)
    '                With NewRecsArray(xCount)
    '                    If xCount > 2 Then
    '                        .dTics = getNewDTick(NewRecs(i).dTics, NewRecsArray, xCount - 1)
    '                    ElseIf xCount = 2 Then
    '                        .dTics = (NewRecs(2).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute ' = NewRecs(1).dTics
    '                        If (NewRecs(2).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute = (NewRecs(1).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute Then
    '                            .dTics = .dTics + 2
    '                        Else
    '                            .dTics = .dTics + 1
    '                        End If
    '                    Else
    '                        .dTics = (NewRecs(1).dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute + 1
    '                    End If


    '                    '.dTics = NewRecs(i).dTics
    '                    .ApptDateStr = formatApptTodoDateStr(.dTics)

    '                    .AcctNum = 0
    '                    .apptDate = NewRecs(i).apptDate
    '                    .DeleteAbsolute = 0
    '                    .DeleteFlag = NewRecs(i).DeleteFlag
    '                    .dTicsCompleted = NewRecs(i).dTicsCompleted
    '                    .dTicsOriginal = NewRecs(i).dTicsOriginal
    '                    .ID = xCount
    '                    .Msg = NewRecs(i).Msg
    '                    .scUserID = NewRecs(i).scUserID
    '                    .StrikeOut = NewRecs(i).StrikeOut

    '                End With
    '                If xCount > 1 Then QuickSort_ApptTodoRecords(NewRecsArray, 1, xCount)

    '                'NewRecs(i).ID = xCount
    '                'NewRecs(i).dTics += xCount
    '                'FilePut(ff, NewRecs(i), xCount)
    '            Next
    '        Next
    '        ReDim Preserve NewRecsArray(xCount)

    '        endTime()

    '        'Dim i As Integer
    '        'Dim FullFileName As String = "C:\todo\ApptTodo.Fil"
    '        'Using writer As New BinaryWriter(File.Open(FullFileName, FileMode.Open), System.Text.Encoding.ASCII)
    '        '    'Using writer As BinaryWriter = New BinaryWriter(File.Open(FullFileName, FileMode.Create), System.Text.Encoding.ASCII)
    '        '    For i = 0 To NewRecsArray.Count - 1
    '        '        writer.Write(a(i))
    '        '    Next
    '        'End Using


    '        Array.Sort(NewRecsArray, (Function(x As ApptTodoRecordType, y As ApptTodoRecordType) x.ID.CompareTo(y.ID)))
    '        startTime()

    '        '=============3.6 seconds - 3.2 seconds
    '        Dim ff As Integer = FreeFile()
    '        Dim filename As String = xBuildFullFileName(file_ApptTodo.FileName)
    '        FileOpen(ff, filename, OpenMode.Random, , , file_ApptTodo.FileLen)

    '        For i = 1 To xCount
    '            xCount += 1
    '            FilePut(ff, NewRecsArray(i), i)
    '        Next
    '        FileClose(ff)

    '        endTime()


    '    End Sub
    '    'Function getHighCnt(ByVal xTicks As Long, ByRef c() As ApptTodoRecordType, ByVal high As Integer) As Integer
    '    Function getNewDTick(ByVal xTicks As Long, ByRef c() As ApptTodoRecordType, ByVal high As Integer) As Long
    '        Dim xDTic As Long = 0
    '        Dim yyyymmddhhmm As Long = (xTicks \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute
    '        Dim NextMinute As Long = yyyymmddhhmm + TimeSpan.TicksPerMinute
    '        Dim LocInFile As Integer
    '        LocInFile = biSearch_GTE_ApptTodoRecords_test(c, NextMinute, high) - 1
    '        Dim dif As Long = c(LocInFile).dTics - yyyymmddhhmm
    '        If dif > 0 Then
    '            xDTic = yyyymmddhhmm + CInt(dif + 1)
    '        Else
    '            xDTic = yyyymmddhhmm + 1
    '        End If
    '        Return xDTic
    '    End Function
    '    Function biSearch_GTE_ApptTodoRecords_test(ByRef c() As ApptTodoRecordType, search As Long, ByVal high As Integer) As Integer ' pass date desired in Long Format
    '        Dim low As Integer = 1
    '        Dim half As Integer
    '        Do
    '            half = (high + low) \ 2
    '            If search > c(half).dTics Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '        Loop Until low > high
    '        Return low
    '    End Function
    '    Sub moveExisting2New_ApptTodoFile()
    '        startTime()
    '        Create_ApptTodo_File_ifDoesNotExists()
    '        Dim Q As String
    '        Q = "SELECT [ID],[dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag],[IntegerDate],[scUserID],[apptDateStr],[apptDate],[dTicsOriginal],[dTicsCompleted] FROM [Todo].[dbo].[Appt] Order by id"
    '        'where substring(apptdatestr,1,10)<>'0005/01/01' ORDER BY dtics"
    '        Dim recs() As ApptRecAllType 'ApptRecType
    '        recs = getApptRecords_AllFields(Q)
    '        Dim i As Integer
    '        Dim nRecs As Integer = UBound(recs)

    '        'For i = 1 To nRecs
    '        '    Debug.Print(i & " " & recs(i).dTics)
    '        'Next


    '        Dim NewRecs(nRecs) As ApptTodoRecordType

    '        For i = 1 To nRecs
    '            With recs(i)
    '                'Debug.Print(.apptDate & " - " & i & " - " & .ID & " " & .msg)
    '                Dim xDate As Date
    '                xDate = New Date((.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay).Date
    '                'xDate = Switch([DateTime], .dTics, New Date(.dTics).Date)

    '                NewRecs(i).ID = i
    '                NewRecs(i).dTics = .dTics

    '                NewRecs(i).Msg = .msg
    '                NewRecs(i).AcctNum = .AcctNum
    '                NewRecs(i).StrikeOut = .StrikeOut
    '                NewRecs(i).DeleteFlag = .DeleteFlag
    '                NewRecs(i).DeleteAbsolute = .IntegerDate
    '                NewRecs(i).scUserID = .scUserID
    '                NewRecs(i).ApptDateStr = .apptDateStr
    '                NewRecs(i).apptDate = (.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay 'xDate 'CDate(Mid(.apptDate, 1, 10)) 'CDate(.apptDate) 'Date.Today.Date '.apptDate
    '                NewRecs(i).dTicsOriginal = .dTicsOriginal
    '                NewRecs(i).dTicsCompleted = .dTicsCompleted
    '                'AddApptTodoRecord(NewRecs(i))

    '                'Dim rrec As ApptTodoRecordType
    '                'rrec = getApptTodoRecord(i)
    '                'If rrec.dTics <> .dTics Then
    '                '    Debug.Print(NewRecs(i).dTics & " orig=" & .dTics & " new=" & rrec.dTics)
    '                '    i = i
    '                'End If



    '                'Debug.Print(NewRecs(i).apptDate)
    '            End With


    '        Next

    '        'build big file 2700 time above
    '        'Dim j As Integer
    '        'For j = 1 To 2700
    '        '    For i = 1 To nRecs
    '        '        AddApptTodoRecord(NewRecs(i))
    '        '    Next
    '        'Next


    '        '=============3.6 seconds
    '        Dim ff As Integer = FreeFile()
    '        Dim filename As String = xBuildFullFileName(file_ApptTodo.FileName)
    '        FileOpen(ff, filename, OpenMode.Random, , , file_ApptTodo.FileLen)

    '        Dim xCount As Integer = 0
    '        Dim j As Integer
    '        For j = 1 To 2700
    '            For i = 1 To nRecs
    '                xCount += 1
    '                NewRecs(i).ID = xCount
    '                NewRecs(i).dTics += xCount
    '                FilePut(ff, NewRecs(i), xCount)
    '            Next
    '        Next

    '        'Debug.Print(rec.apptDate)



    '        FileClose(ff)
    '        '=============
    '        endTime()
    '    End Sub
    '    Function test_GetApptsForGivenDate(Optional ByVal xDate As Date = Nothing) As ApptTodoRecordType()
    '        If xDate = Nothing Then xDate = Date.Today
    '        Dim search As Long = xDate.Ticks
    '        Dim searchE As Long = search + TimeSpan.TicksPerDay


    '        Dim appts(1000000) As ApptTodoRecordType
    '        Dim cnt As Integer = 0
    '        Dim startPt As Integer
    '        Dim endPt As Integer
    '        startPt = biSearch_GTE_AllAppts_Long(search) ' biSearch_E_AllAppt(search)
    '        endPt = biSearch_GTE_AllAppts_Long(searchE) 'startPt ' BiSearch_Long_LTE(searchE)
    '        '637047936000000001
    '        'For i = 1 To UBound(ApptGV.AllAppts_LongInteger)
    '        '    Debug.Print(i & " " & ApptGV.AllAppts_LongInteger(i).L)
    '        'Next
    '        '=========================1.3394177 secords
    '        'For i = startPt To endPt - 1
    '        '    cnt += 1
    '        '    appts(cnt) = getApptTodoRecord(ApptGV.AllAppts_LongInteger(i).i)
    '        'Next
    '        '=========================

    '        '=========0.0957441 seconds
    '        Dim ff As Integer = FreeFile()
    '        Dim filename As String = xBuildFullFileName(file_ApptTodo.FileName)
    '        FileOpen(ff, filename, OpenMode.Random, , , file_ApptTodo.FileLen)
    '        '==
    '        For i = startPt To endPt - 1
    '            cnt += 1
    '            FileGet(ff, appts(cnt), ApptGV.AllAppts_LongInteger(i).i) 'posInFile)
    '        Next
    '        '==
    '        FileClose(ff)
    '        '=========

    '        ReDim Preserve appts(cnt)
    '        Return appts
    '    End Function

    '    Sub test_getPat()
    '        Dim appts() As ApptTodoRecordType
    '        Dim xDate As Date = #9/23/2019#
    '        'appts = GetApptsForGivenDate(xDate)
    '        appts = test_GetApptsForGivenDate(xDate)
    '        Dim i As Integer
    '        Dim nItems As Integer = UBound(appts)
    '        For i = 1 To nItems
    '            With appts(i)
    '                'Debug.Print(.ID & " " & .ApptDateStr & " " & .Msg)
    '            End With
    '        Next
    '    End Sub
    '    Sub test_getPat2()
    '        Dim appts(1000000) As ApptTodoRecordType
    '        Dim cnt As Integer = 0
    '        Dim xDate As Date = #9/23/2019#
    '        'appts = GetApptsForGivenDate(xDate)
    '        ''''appts = test_GetApptsForGivenDate(xDate)
    '        Dim i As Integer

    '        For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
    '            If Mid(ApptGV.All_ApptTodoRecords(i).ApptDateStr, 1, 10) = "2019/09/23" Then
    '                cnt += 1
    '                appts(cnt) = ApptGV.All_ApptTodoRecords(i)
    '            End If
    '        Next
    '        ReDim Preserve appts(cnt)

    '        Dim nItems As Integer = UBound(appts)
    '        For i = 1 To nItems
    '            With appts(i)
    '                ' Debug.Print(.ID & " " & .ApptDateStr & " " & .Msg)
    '            End With
    '        Next
    '    End Sub
    '    Sub test_setAll_ApptTodoRecords()
    '        startTime()
    '        setAll_ApptTodoRecords() '0.2 seconds
    '        endTime()

    '        startTime()
    '        QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger)) '0.012 seconds for 100,000 records
    '        endTime()

    '        startTime()

    '        'test_getPat()
    '        test_getPat2()
    '        endTime()
    '        Exit Sub

    '        'Debug.Print(ApptGV.All_ApptTodoRecords(1).ApptDateStr)'0005/01/01 00:00:00.0000003

    '        Dim i As Integer
    '        startTime()

    '        For i = 1 To ApptGV.All_ApptTodoRecords_Cnt '0.02 seconds
    '            If Mid(ApptGV.All_ApptTodoRecords(i).ApptDateStr, 1, 10) = "0005/01/01" Then '0.16 seconds
    '                'If InStr(ApptGV.All_ApptTodoRecords(i).Msg, "Pat") Then
    '                'Debug.Print(i & ApptGV.All_ApptTodoRecords(i).Msg)
    '            End If
    '        Next
    '        endTime()
    '    End Sub
    '    Sub setAll_ApptTodoRecords()
    '        ApptGV.All_ApptTodoRecords = getAll_ApptTodoRecordType_UsingReader() '0.2 secs
    '        ApptGV.All_ApptTodoRecords_Cnt = UBound(ApptGV.All_ApptTodoRecords)
    '    End Sub
    '    Sub test_getAll_ApptTodoRecordType_Read1by1()
    '        startTime()

    '        Dim rRecs() As ApptTodoRecordType = getAll_ApptTodoRecordType_UsingReader() '0.2 secs
    '        endTime()
    '        Debug.Print(UBound(rRecs))
    '        Exit Sub

    '        Dim i As Integer

    '        startTime()
    '        Dim recs() As ApptTodoRecordType = getAll_ApptTodoRecordType_Read1by1() '2.3 secs
    '        endTime()
    '        Debug.Print(UBound(recs))
    '        Exit Sub

    '        Dim nRecs As Integer = UBound(recs)
    '        For i = 1 To nRecs
    '            'If recs(i).dTics <> rRecs(i).dTics Then
    '            '    Debug.Print(i & " " & recs(i).dTics & " " & rRecs(i).dTics)
    '            'End If
    '            'If recs(i).Msg <> rRecs(i).Msg Then
    '            '    Debug.Print(i & " " & recs(i).Msg & " " & rRecs(i).Msg)
    '            'End If

    '        Next
    '    End Sub
    '    Function getAll_ApptTodoRecordType_Read1by1() As ApptTodoRecordType() '2.3 secs for read 102600 recs
    '        Dim info As New FileInfo(xBuildFullFileName(file_ApptTodo.FileName))
    '        Dim nRecs As Integer = info.Length / file_ApptTodo.FileLen

    '        ' Get length of the file.
    '        Dim length As Long = info.Length
    '        Dim Recs(nRecs) As ApptTodoRecordType
    '        Dim i As Integer

    '        '===
    '        Dim ff As Integer = FreeFile()
    '        Dim filename As String = xBuildFullFileName(file_ApptTodo.FileName)
    '        FileOpen(ff, filename, OpenMode.Random, , , file_ApptTodo.FileLen)
    '        '===

    '        For i = 1 To nRecs
    '            'Recs(i) = getApptTodoRecord(i)


    '            FileGet(ff, Recs(i), i)

    '        Next
    '        '===
    '        FileClose(ff)
    '        '===
    '        Return Recs
    '    End Function
    '    Sub Set_BirthdayRecordType_UsingReader()
    '        Dim cnt As Integer = 0
    '        Dim i As Integer

    '        Dim NumRecs As Integer
    '        Dim cntTodo As Integer = 0

    '        Dim FullFileName As String = xBuildFullFileName(file_Birthday.FileName)

    '        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
    '            Dim length As Integer = reader.BaseStream.Length
    '            NumRecs = length / file_Birthday.FileLen

    '            ReDim ApptGV.Birthdays(NumRecs)

    '            'ReDim Recs(NumRecs)
    '            'ReDim ApptGV.All_ApptTodoRecords(NumRecs) 'ApptGV.AllAppts_LongInteger(NumRecs)
    '            'ReDim ApptGV.AllAppts_LongInteger(NumRecs)

    '            'new

    '            Dim r As BirthdayRecordType
    '            For i = 1 To NumRecs
    '                'id, dtics, msg, desc, descriptionnumber, deleteflag, datestr, Month, Day, dticsoriginal
    '                With ApptGV.Birthdays(i)
    '                    .ID = reader.ReadInt32
    '                    .dTics = reader.ReadInt64

    '                    .Msg = reader.ReadChars(70)
    '                    .Desc = reader.ReadChars(14)
    '                    .DescriptionNumber = reader.ReadInt32
    '                    .DeleteFlag = reader.ReadInt32
    '                    .DateStr = reader.ReadChars(10)
    '                    .Month = reader.ReadInt32
    '                    .Day = reader.ReadInt32
    '                    .dTicsOriginal = reader.ReadInt64

    '                    'Debug.Print(.ID & " " & Trim(.Msg) & " " & formatApptDateStr(.dTics) & " " & Trim(.DateStr) & " " & Trim(.Desc) & " " & .DescriptionNumber & " " & .DeleteFlag & " " & .Day & " " & .Month & " " & formatApptDateStr(.dTicsOriginal))
    '                    'Debug.Print(.ID & " " & Trim(.Msg) & " " & Trim(.DateStr) & " desc#=" & .DescriptionNumber & " del=" & .DeleteFlag & " day=" & .Day & " mth=" & .Month & " org=" & formatApptDateStr(.dTicsOriginal))

    '                    'below is GOOD!
    '                    'Debug.Print(String.Format("{0:0000}", .ID) & " " & Trim(.DateStr) & " desc#=" & .DescriptionNumber & " del=" & .DeleteFlag & " day=" & String.Format("{0:00}", .Day) & " mth=" & String.Format("{0:00}", .Month) & " " & Trim(.Msg))
    '                    'Debug.Print("--")

    '                    'If .dTicsOriginal > ApptGV.Highest_dTicsOriginal Then ApptGV.Highest_dTicsOriginal = .dTicsOriginal
    '                    '===
    '                    'If .Msg = "test adding todo for noDate //" Then
    '                    'If InStr(.Msg, "test adding todo for noDate //") > 0 Then
    '                    '    Debug.Print(Trim(.Msg) & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & New Date(.dTicsOriginal))
    '                    '    i = i
    '                    'End If
    '                    '===

    '                    '.dTics = reader.ReadInt64
    '                    '=========================
    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
    '                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

    '                    'Debug.Print(.dTics)

    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
    '                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

    '                    'If (.dTicsOriginal \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay = ApptGV.todoTicks Then
    '                    'If .apptDate = ApptGV.todoTicks Then
    '                    '    cntTodo += 1
    '                    '    ApptGV.TodoCreateDate(cntTodo).L = .dTicsOriginal
    '                    '    ApptGV.TodoCreateDate(cntTodo).i = .ID
    '                    'End If

    '                End With
    '            Next
    '        End Using
    '        'ApptGV.All_ApptTodoRecords_Cnt = NumRecs

    '        'QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt) '0.0408907 seconds (This sort is Twice as fast as the Array.sort lambda function below))

    '        QuickSort_BirthdayRecords(ApptGV.Birthdays, 1, NumRecs)

    '    End Sub
    '    Sub doReorgBirthdays()
    '        startTime()
    '        Set_BirthdayRecordType_UsingReader()
    '        Build_BirthdayRecordFile_UsingWriter()
    '        Set_BirthdayRecordType_UsingReader()
    '        endTime()
    '        'End
    '        Debug.Print("delete done!")
    '    End Sub
    '    Sub Build_BirthdayRecordFile_UsingWriter()
    '        Dim NumRecs As Integer = UBound(ApptGV.Birthdays)
    '        Dim i As Integer
    '        Dim cnt As Integer = 0
    '        Dim FullFileName As String = xBuildFullFileName(file_Birthday.FileName)

    '        If System.IO.File.Exists(FullFileName) = True Then Kill(FullFileName)
    '        System.IO.File.Create(FullFileName).Dispose()

    '        'If System.IO.File.Exists(FullFileName) = False Then System.IO.File.Create(FullFileName).Dispose()
    '        'Dim i As Integer
    '        'Using writer As New BinaryWriter(File.Open(FullFileName, FileMode.Open), System.Text.Encoding.ASCII)
    '        Using writer As New BinaryWriter(File.Open(FullFileName, FileMode.Open))
    '            'Using writer As BinaryWriter = New BinaryWriter(File.Open(FullFileName, FileMode.Create), System.Text.Encoding.ASCII)
    '            For i = 1 To NumRecs
    '                With ApptGV.Birthdays(i)
    '                    If .DeleteFlag = 0 Then
    '                        cnt += 1
    '                        .ID = cnt
    '                        'writer.Write(ApptGV.Birthdays(i))
    '                        writer.Write(.ID)
    '                        writer.Write(.dTics) ' = reader.ReadInt64
    '                        writer.Write(.Msg, 0, 70) ' = reader.ReadChars(70)
    '                        writer.Write(.Desc, 0, 14) ' = reader.ReadChars(14)
    '                        writer.Write(.DescriptionNumber) ' = reader.ReadInt32
    '                        writer.Write(.DeleteFlag) ' = reader.ReadInt32
    '                        writer.Write(.DateStr, 0, 10) ' = reader.ReadChars(10)
    '                        writer.Write(.Month) ' = reader.ReadInt32
    '                        writer.Write(.Day) ' = reader.ReadInt32
    '                        writer.Write(.dTicsOriginal) ' = reader.ReadInt64
    '                    End If
    '                End With

    '            Next
    '        End Using
    '        Exit Sub

    '        '========



    '        'Dim NumRecs As Integer = UBound(ApptGV.Birthdays)
    '        Dim cntTodo As Integer = 0

    '        'Dim FullFileName As String = xBuildFullFileName(file_Birthday.FileName)

    '        Using writer As New BinaryWriter(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
    '            'Dim length As Integer = reader.BaseStream.Length
    '            'NumRecs = length / file_Birthday.FileLen

    '            'ReDim ApptGV.Birthdays(NumRecs)

    '            ''ReDim Recs(NumRecs)
    '            'ReDim ApptGV.All_ApptTodoRecords(NumRecs) 'ApptGV.AllAppts_LongInteger(NumRecs)
    '            ''ReDim ApptGV.AllAppts_LongInteger(NumRecs)

    '            ''new

    '            Dim r As BirthdayRecordType
    '            For i = 1 To NumRecs
    '                'id, dtics, msg, desc, descriptionnumber, deleteflag, datestr, Month, Day, dticsoriginal
    '                With ApptGV.Birthdays(i)
    '                    '.ID = reader.ReadInt32
    '                    '.dTics = reader.ReadInt64

    '                    '.Msg = reader.ReadChars(70)
    '                    '.Desc = reader.ReadChars(14)
    '                    '.DescriptionNumber = reader.ReadInt32
    '                    '.DeleteFlag = reader.ReadInt32
    '                    '.DateStr = reader.ReadChars(10)
    '                    '.Month = reader.ReadInt32
    '                    '.Day = reader.ReadInt32
    '                    '.dTicsOriginal = reader.ReadInt64

    '                    'If .dTicsOriginal > ApptGV.Highest_dTicsOriginal Then ApptGV.Highest_dTicsOriginal = .dTicsOriginal
    '                    '===
    '                    'If .Msg = "test adding todo for noDate //" Then
    '                    'If InStr(.Msg, "test adding todo for noDate //") > 0 Then
    '                    '    Debug.Print(Trim(.Msg) & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & New Date(.dTicsOriginal))
    '                    '    i = i
    '                    'End If
    '                    '===

    '                    '.dTics = reader.ReadInt64
    '                    '=========================
    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
    '                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

    '                    'Debug.Print(.dTics)

    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
    '                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

    '                    'If (.dTicsOriginal \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay = ApptGV.todoTicks Then
    '                    'If .apptDate = ApptGV.todoTicks Then
    '                    '    cntTodo += 1
    '                    '    ApptGV.TodoCreateDate(cntTodo).L = .dTicsOriginal
    '                    '    ApptGV.TodoCreateDate(cntTodo).i = .ID
    '                    'End If

    '                End With
    '            Next
    '        End Using
    '        'ApptGV.All_ApptTodoRecords_Cnt = NumRecs

    '        'QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt) '0.0408907 seconds (This sort is Twice as fast as the Array.sort lambda function below))

    '        QuickSort_BirthdayRecords(ApptGV.Birthdays, 1, NumRecs)

    '    End Sub

    '    Sub Set_ApptTodoRecordType_UsingReader() ' As ApptTodoRecordType()

    '        Dim cnt As Integer = 0
    '        Dim i As Integer

    '        'new
    '        'Dim distinctDates(0) As Long

    '        'Dim Recs() As ApptTodoRecordType

    '        Dim NumRecs As Integer
    '        Dim cntTodo As Integer = 0

    '        Dim FullFileName As String = xBuildFullFileName(file_ApptTodo.FileName)

    '        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
    '            Dim length As Integer = reader.BaseStream.Length
    '            NumRecs = length / file_ApptTodo.FileLen

    '            ReDim ApptGV.TodoCreateDate(NumRecs)

    '            'ReDim Recs(NumRecs)
    '            ReDim ApptGV.All_ApptTodoRecords(NumRecs) 'ApptGV.AllAppts_LongInteger(NumRecs)
    '            'ReDim ApptGV.AllAppts_LongInteger(NumRecs)

    '            'new
    '            ReDim ApptGV.distinctDates(NumRecs)

    '            For i = 1 To NumRecs
    '                With ApptGV.All_ApptTodoRecords(i)
    '                    .ID = reader.ReadInt32
    '                    .dTics = reader.ReadInt64

    '                    .Msg = reader.ReadChars(84)


    '                    .AcctNum = reader.ReadInt32
    '                    .StrikeOut = reader.ReadInt32
    '                    .DeleteFlag = reader.ReadInt32
    '                    .DeleteAbsolute = reader.ReadInt32
    '                    .scUserID = reader.ReadInt32
    '                    .ApptDateStr = reader.ReadChars(27)



    '                    '.apptDate = (.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay 'xDate 'CDate(Mid(.apptDate, 1, 10)) 'CDate(.apptDate) 'Date.Today.Date '.apptDate
    '                    .apptDate = reader.ReadInt64

    '                    'new
    '                    'ApptGV.distinctDates(i) = (.dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '.apptDate

    '                    .dTicsOriginal = reader.ReadInt64
    '                    .dTicsCompleted = reader.ReadInt64


    '                    If .dTicsOriginal > ApptGV.Highest_dTicsOriginal Then ApptGV.Highest_dTicsOriginal = .dTicsOriginal
    '                    '===
    '                    'If .Msg = "test adding todo for noDate //" Then
    '                    'If InStr(.Msg, "test adding todo for noDate //") > 0 Then
    '                    '    Debug.Print(Trim(.Msg) & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & New Date(.dTicsOriginal))
    '                    '    i = i
    '                    'End If
    '                    '===

    '                    '.dTics = reader.ReadInt64
    '                    '=========================
    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
    '                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

    '                    'Debug.Print(.dTics)

    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
    '                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

    '                    'If (.dTicsOriginal \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay = ApptGV.todoTicks Then
    '                    'If .apptDate = ApptGV.todoTicks Then
    '                    '    cntTodo += 1
    '                    '    ApptGV.TodoCreateDate(cntTodo).L = .dTicsOriginal
    '                    '    ApptGV.TodoCreateDate(cntTodo).i = .ID
    '                    'End If

    '                End With
    '            Next
    '        End Using
    '        ApptGV.All_ApptTodoRecords_Cnt = NumRecs
    '        'For i = ApptGV.All_ApptTodoRecords_Cnt To ApptGV.All_ApptTodoRecords_Cnt - 8 Step -1
    '        '    With ApptGV.All_ApptTodoRecords(i)
    '        '        'Debug.Print(Trim(.Msg) & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & New Date(.dTicsOriginal))
    '        '        Debug.Print(.ID & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & .dTics & " " & New Date(.dTicsOriginal) & " :::" & Trim(.Msg))
    '        '    End With

    '        'Next
    '        QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt) '0.0408907 seconds (This sort is Twice as fast as the Array.sort lambda function below))


    '        '===

    '        ''===========================fix origDate
    '        'startTime()
    '        'Dim sPoint As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
    '        'Dim ePoint As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerMinute)
    '        'Dim nItems As Integer = ePoint - sPoint
    '        'Dim xCount As Integer = 0
    '        'Debug.Print(sPoint & " " & ePoint & " " & nItems)
    '        ''=============3.6 seconds - 3.2 seconds
    '        'Dim ff As Integer = FreeFile()
    '        'Dim filename As String = xBuildFullFileName(file_ApptTodo.FileName)
    '        'FileOpen(ff, filename, OpenMode.Random, , , file_ApptTodo.FileLen)

    '        'For i = sPoint To ePoint - 1
    '        '    xCount += 1
    '        '    ApptGV.All_ApptTodoRecords(i).dTicsOriginal = ApptGV.All_ApptTodoRecords(i).dTicsOriginal + xCount
    '        '    FilePut(ff, ApptGV.All_ApptTodoRecords(i), ApptGV.All_ApptTodoRecords(i).ID)
    '        'Next
    '        'FileClose(ff)

    '        'endTime()
    '        ''===========================end

    '        'ReDim Preserve ApptGV.TodoCreateDate(cntTodo)
    '        'ReDim ApptGV.TodoDistinctDates(cntTodo)
    '        'For i = 1 To cntTodo
    '        '    ApptGV.TodoDistinctDates(i) = ApptGV.TodoCreateDate(i).L
    '        'Next

    '        'Debug.Print(cntTodo)
    '        ''Debug.Print(UBound(ApptGV.TodoDistinctDates))
    '        ''startTime()
    '        'ApptGV.TodoDistinctDates = ApptGV.TodoDistinctDates.Distinct.ToArray '0.002 seconds for 5400 records
    '        'Debug.Print(UBound(ApptGV.TodoDistinctDates))
    '        ''===


    '        'QuickSort_LongInteger_Long(ApptGV.TodoCreateDate, 1, cntTodo)
    '        'For i = 1 To cntTodo
    '        '    'Debug.Print(ApptGV.TodoCreateDate(i).i & " " & ApptGV.All_ApptTodoRecords(i).ID)
    '        '    'ApptGV.TodoCreateDate(cntTodo).i = ApptGV.All_ApptTodoRecords(i).ID
    '        '    ApptGV.TodoCreateDate(i).i = i
    '        'Next

    '        'new
    '        'startTime()
    '        ''''ApptGV.distinctDates = ApptGV.distinctDates.Distinct.ToArray '0.004 seconds for 5400 records
    '        'endTime()


    '        ''''Array.Sort(ApptGV.distinctDates) '25 items
    '        'Debug.print(UBound(ApptGV.distinctDates))
    '        'For i = 1 To UBound(ApptGV.distinctDates)
    '        '    'Debug.Print(distinctDates(i) & " - " & New Date(distinctDates(i)))
    '        '    Debug.Print(ApptGV.distinctDates(i))
    '        'Next
    '    End Sub
    '    Sub set_CreateDateArray() 'to be run after Set_ApptTodoRecordType_UsingReader - and sorted
    '        'startTime()

    '        Dim sPoint As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks)
    '        Dim ePoint As Integer = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + TimeSpan.TicksPerDay) 'TimeSpan.TicksPerMinute)
    '        Dim nItems As Integer = ePoint - sPoint
    '        Dim cnt As Integer = 0
    '        ReDim ApptGV.TodoCreateDate(nItems)
    '        Dim i As Integer
    '        'Dim LL As Long
    '        For i = sPoint To ePoint - 1
    '            cnt += 1
    '            With ApptGV.TodoCreateDate(cnt)
    '                '.L = (ApptGV.All_ApptTodoRecords(i).dTicsOriginal \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay + (ApptGV.All_ApptTodoRecords(i).dTics - ApptGV.All_ApptTodoRecords(i).apptDate)
    '                .L = getFirstOfMonthTicks(ApptGV.All_ApptTodoRecords(i).dTicsOriginal) + (ApptGV.All_ApptTodoRecords(i).dTics - ApptGV.All_ApptTodoRecords(i).apptDate)

    '                ''''.L = getFirstOfMonthTicks(ApptGV.All_ApptTodoRecords(i).dTicsOriginal) + (ApptGV.All_ApptTodoRecords(i).dTics Mod TimeSpan.TicksPerSecond)
    '                'If .L <> LL Then
    '                '    MsgBox(i & "stop")
    '                'End If
    '                'or
    '                '.L = getFirstOfMonthTicks(ApptGV.All_ApptTodoRecords(i).dTicsOriginal) + (ApptGV.All_ApptTodoRecords(i).dTics Mod TimeSpan.TicksPerSecond)

    '                'TimeSpan.

    '                .i = i 'position in array
    '            End With
    '            'Debug.Print(cnt & " " & formatApptTodoDateStr(ApptGV.TodoCreateDate(i).L) & " " & ApptGV.TodoCreateDate(i).i & " " & ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i).DeleteFlag & " " & ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i).Msg)
    '            'MsgBox(i & "stop")
    '        Next
    '        QuickSort_LongInteger_Long(ApptGV.TodoCreateDate, 1, cnt)

    '        'For i = cnt To cnt - 8 Step -1
    '        '    Debug.Print(i & " " & formatApptTodoDateStr(ApptGV.TodoCreateDate(i).L) & " " & ApptGV.TodoCreateDate(i).i & " " & ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i).DeleteFlag & " " & ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i).Msg)
    '        'Next

    '        'For i = 1 To cnt
    '        'For i = 8090 To 8110
    '        '    'Debug.Print(New Date(ApptGV.TodoCreateDate(i).L) & " " & ApptGV.TodoCreateDate(i).i & " " & ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i).Msg)
    '        '    Debug.Print(i & " " & formatApptTodoDateStr(ApptGV.TodoCreateDate(i).L) & " " & ApptGV.TodoCreateDate(i).i & " " & ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i).DeleteFlag & " " & ApptGV.All_ApptTodoRecords(ApptGV.TodoCreateDate(i).i).Msg)
    '        'Next
    '        'endTime()
    '    End Sub
    '    Function getAll_ApptTodoRecordType_UsingReader() As ApptTodoRecordType()

    '        Dim cnt As Integer = 0
    '        Dim i As Integer

    '        Dim Recs() As ApptTodoRecordType

    '        Dim NumRecs As Integer

    '        Dim FullFileName As String = xBuildFullFileName(file_ApptTodo.FileName)

    '        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
    '            Dim length As Integer = reader.BaseStream.Length
    '            NumRecs = length / file_ApptTodo.FileLen

    '            ReDim Recs(NumRecs)
    '            ReDim ApptGV.AllAppts_LongInteger(NumRecs)
    '            'ReDim ApptGV.AllAppts_LongInteger(NumRecs)

    '            For i = 1 To NumRecs
    '                With Recs(i)
    '                    .ID = reader.ReadInt32
    '                    .dTics = reader.ReadInt64

    '                    .Msg = reader.ReadChars(84)
    '                    .AcctNum = reader.ReadInt32
    '                    .StrikeOut = reader.ReadInt32
    '                    .DeleteFlag = reader.ReadInt32
    '                    .DeleteAbsolute = reader.ReadInt32
    '                    .scUserID = reader.ReadInt32
    '                    .ApptDateStr = reader.ReadChars(27)

    '                    '.apptDate = (.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay 'xDate 'CDate(Mid(.apptDate, 1, 10)) 'CDate(.apptDate) 'Date.Today.Date '.apptDate
    '                    .apptDate = reader.ReadInt64

    '                    .dTicsOriginal = reader.ReadInt64
    '                    .dTicsCompleted = reader.ReadInt64

    '                    '.dTics = reader.ReadInt64
    '                    '=========================
    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
    '                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

    '                    'Debug.Print(.dTics)

    '                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
    '                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

    '                End With
    '            Next
    '        End Using
    '        Return Recs
    '    End Function
    '    'Sub AddApptTodoRecord(ByVal rec As ApptTodoRecordType) 'use for recurring records  'new 9/5/2015

    '    '    Dim ff As Integer = FreeFile()
    '    '    Dim filename As String = xBuildFullFileName(file_ApptTodo.FileName)
    '    '    FileOpen(ff, filename, OpenMode.Random, , , file_ApptTodo.FileLen)

    '    '    Dim xCount As Integer = LOF(ff) / file_ApptTodo.FileLen + 1
    '    '    'Debug.Print(rec.apptDate)
    '    '    FilePut(ff, rec, xCount)
    '    '    FileClose(ff)

    '    '    'new
    '    '    'ReDim Preserve ApptGV.AllAppts_LongInteger(xCount)
    '    '    'With ApptGV.AllAppts_LongInteger(xCount)
    '    '    '    .L = rec.dTics
    '    '    '    .i = xCount
    '    '    'End With
    '    '    'QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))

    '    '    'set_AllAppts_LongInteger()
    '    'End Sub
    '    Function getApptTodoRecord(ByVal posInFile As Integer) As ApptTodoRecordType
    '        'Dim rec As ApptTodoRecordType
    '        Dim rec As ApptTodoRecordType = initApptTodoRecord()
    '        Dim ff As Integer = FreeFile()
    '        Dim filename As String = xBuildFullFileName(file_ApptTodo.FileName)
    '        FileOpen(ff, filename, OpenMode.Random, , , file_ApptTodo.FileLen)


    '        FileGet(ff, rec, posInFile)
    '        FileClose(ff)

    '        Return rec
    '    End Function
    '    Sub test_ReadAllRecordsAndBuildExtermalArray()
    '        startTime()
    '        ReadAllRecordsAndBuildExtermalArray()


    '        Dim appts() As ApptTodoRecordType
    '        Dim xDate As Date = #9/23/2019#
    '        appts = GetApptsForGivenDate(xDate)

    '        Dim i As Integer
    '        Dim nItems As Integer = UBound(appts)
    '        For i = 1 To nItems
    '            With appts(i)
    '                Debug.Print(.ApptDateStr & " " & .Msg)
    '            End With
    '        Next
    '        endTime()
    '    End Sub
    '    Sub ReadAllRecordsAndBuildExtermalArray()

    '        Dim cnt As Integer = 0
    '        Dim i As Integer
    '        'Dim rec As ApptTodoRecordType
    '        Dim rec As ApptTodoRecordType = initApptTodoRecord()
    '        Dim NumRecs As Integer

    '        'Dim FullFileName As String = "c:\todo\Appt.fil" '"d:\todo\appt.fil" ' "f:\bu\Appt.fil" ' xBuildFullFileName(f_Appt.FileName)
    '        Dim FullFileName As String = xBuildFullFileName(file_ApptTodo.FileName)
    '        'Debug.Print(FullFileName)

    '        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
    '            Dim length As Integer = reader.BaseStream.Length
    '            NumRecs = length / file_ApptTodo.FileLen
    '            'Debug.Print(NumRecs)

    '            ReDim ApptGV.AllAppts_LongInteger(NumRecs)
    '            'startTime()
    '            For i = 1 To NumRecs
    '                With rec
    '                    '            Dim ID As Integer
    '                    '            Dim dTics As Long
    '                    '<VBFixedString(84)> Dim Msg As String
    '                    '            Dim AcctNum As Integer
    '                    '            Dim StrikeOut As Integer
    '                    '            Dim DeleteFlag As Integer
    '                    '            Dim DeleteAbsolute As Integer 'IntegerDAte As Integer
    '                    '            Dim scUserID As Integer
    '                    '<VBFixedString(27)> Dim ApptDateStr As String
    '                    '            Dim apptDate As Long 'Date
    '                    '            Dim dTicsOriginal As Long
    '                    '            Dim dTicsCompleted As Long


    '                    .ID = reader.ReadInt32
    '                    .dTics = reader.ReadInt64

    '                    .Msg = reader.ReadChars(84)
    '                    .AcctNum = reader.ReadInt32
    '                    .StrikeOut = reader.ReadInt32
    '                    .DeleteFlag = reader.ReadInt32
    '                    .DeleteAbsolute = reader.ReadInt32
    '                    .scUserID = reader.ReadInt32
    '                    .ApptDateStr = reader.ReadChars(27)

    '                    '.apptDate = (.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay 'xDate 'CDate(Mid(.apptDate, 1, 10)) 'CDate(.apptDate) 'Date.Today.Date '.apptDate
    '                    .apptDate = reader.ReadInt64

    '                    .dTicsOriginal = reader.ReadInt64
    '                    .dTicsCompleted = reader.ReadInt64
    '                    '.dTics = reader.ReadInt64
    '                    '=========================
    '                    ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
    '                    ApptGV.AllAppts_LongInteger(i).i = .ID 'i

    '                    Debug.Print(.dTics)

    '                    'If .dTics <= ApptGV.HighBDayTics Then
    '                    '    If .DeleteFlag = 0 Then cnt += 1 : recs(cnt) = recs(i)
    '                    'ElseIf Mid(.msg, 1, 1) <> Chr(0) Then
    '                    '    cnt += 1 : recs(cnt) = recs(i)
    '                    'End If
    '                End With
    '            Next
    '            'endTime()
    '            'End
    '            'ReDim Preserve recs(cnt)
    '            'Return recs
    '        End Using
    '        'End
    '        'For i = 1 To NumRecs
    '        '    With recs(i)
    '        '        Debug.Print("{0}-{1}-{2}-{3}-{4}-{5}", i, .AcctNum, .DeleteFlag, .StrikeOut, New Date(.dTics), .msg)
    '        '        Debug.Print("")

    '        '    End With


    '        'Next

    '        'End

    '        ''''ApptGV.AllAppts_LongInteger = getLongInteger_Apptfil() 'as of 11/11/2015
    '        'new 11/26/2016 usage
    '        QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))
    '    End Sub
    '    Function GetApptsForGivenDate(Optional ByVal xDate As Date = Nothing) As ApptTodoRecordType()
    '        If xDate = Nothing Then xDate = Date.Today
    '        Dim search As Long = xDate.Ticks
    '        Dim searchE As Long = search + TimeSpan.TicksPerDay


    '        Dim appts(1000000) As ApptTodoRecordType
    '        Dim cnt As Integer = 0
    '        Dim startPt As Integer
    '        Dim endPt As Integer
    '        startPt = biSearch_GTE_AllAppts_Long(search) ' biSearch_E_AllAppt(search)
    '        endPt = biSearch_GTE_AllAppts_Long(searchE) 'startPt ' BiSearch_Long_LTE(searchE)
    '        '637047936000000001
    '        For i = 1 To UBound(ApptGV.AllAppts_LongInteger)
    '            Debug.Print(i & " " & ApptGV.AllAppts_LongInteger(i).L)
    '        Next
    '        For i = startPt To endPt - 1
    '            cnt += 1
    '            appts(cnt) = getApptTodoRecord(ApptGV.AllAppts_LongInteger(i).i)
    '        Next
    '        ReDim Preserve appts(Math.Min(ApptGV.MaxApptTodos, cnt))
    '        Return appts
    '    End Function
    '    Function biSearch_GTE_AllAppts_Long(ByVal search As Long, Optional ByVal USEi As Boolean = False) As Integer
    '        'This function only works on Appt and Todos and bdays (not recur)
    '        'from: biSearch_GTE_LongInteger 10/23/2015

    '        'If ApptGV.AllAppts_LongInteger.Count = 1 Then
    '        '    setLongType() '0.002 seconds
    '        '    'startTime()
    '        '    ''WriteLongType_Base1("c:\testfiles\LongType.fil", ApptGV.LongType) '0.002 sec
    '        '    'ApptGV.LongType = ReadLongType_Base1("c:\testfiles\LongType.fil") '0.04 secs
    '        '    'endTime()
    '        'End If
    '        'If ApptGV.AllAppts_LongInteger.Count = 1 Then Return 0

    '        'Dim HH As Integer = UBound(a)
    '        Dim high As Integer = UBound(ApptGV.AllAppts_LongInteger) 'HH
    '        Dim low As Integer = 1
    '        Dim half As Integer

    '        If USEi = False Then
    '            Do
    '                half = (high + low) \ 2
    '                If search > ApptGV.AllAppts_LongInteger(half).L Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '            Loop Until low > high
    '        Else
    '            Do
    '                half = (high + low) \ 2
    '                If search > ApptGV.AllAppts_LongInteger(half).L Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '            Loop Until low > high
    '        End If
    '        Return low
    '    End Function
    '    Function biSearch_GTE_ApptTodoRecords(ByVal search As Long) As Integer ' pass date desired in Long Format
    '        'ApptGV.All_ApptTodoRecords_Cnt

    '        'Dim cnt As Integer = 0

    '        Dim high As Integer = ApptGV.All_ApptTodoRecords_Cnt
    '        Dim low As Integer = 1
    '        Dim half As Integer
    '        Do
    '            'cnt += 1

    '            half = (high + low) \ 2
    '            If search > ApptGV.All_ApptTodoRecords(half).dTics Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '        Loop Until low > high
    '        'Debug.Print("biSearch_GTE_ApptTodoRecords - cnt=" & cnt) '17
    '        Return low
    '    End Function
    '    Function biSearch_GTE_TodoCreateDates(ByVal search As Long) As Integer ', Optional ByVal USEi As Boolean = False) As Integer
    '        If ApptGV.TodoCreateDate.Count = 1 Then Return 0

    '        'Dim HH As Integer = UBound(a)
    '        Dim high As Integer = UBound(ApptGV.TodoCreateDate) 'HH
    '        Dim low As Integer = 1
    '        Dim half As Integer

    '        'If USEi = False Then
    '        Do
    '            half = (high + low) \ 2
    '            'Debug.Print(New Date(ApptGV.TodoCreateDate(half).L))
    '            If search > ApptGV.TodoCreateDate(half).L Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '        Loop Until low > high
    '        'Else
    '        '    Do
    '        '        half = (high + low) \ 2
    '        '        If search > (ApptGV.LongType(half) Mod TimeSpan.TicksPerSecond) Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '        '    Loop Until low > high
    '        'End If
    '        Return low

    '    End Function
    '    Function biSearch_GTE_LongType(ByVal search As Long, Optional ByVal USEi As Boolean = False) As Integer
    '        'This function only works on Appt and Todos and bdays (not recur)
    '        'from: biSearch_GTE_LongInteger 10/23/2015

    '        If ApptGV.LongType.Count = 1 Then
    '            setLongType() '0.002 seconds
    '            'startTime()
    '            ''WriteLongType_Base1("c:\testfiles\LongType.fil", ApptGV.LongType) '0.002 sec
    '            'ApptGV.LongType = ReadLongType_Base1("c:\testfiles\LongType.fil") '0.04 secs
    '            'endTime()
    '        End If
    '        If ApptGV.LongType.Count = 1 Then Return 0

    '        'Dim HH As Integer = UBound(a)
    '        Dim high As Integer = UBound(ApptGV.LongType) 'HH
    '        Dim low As Integer = 1
    '        Dim half As Integer

    '        If USEi = False Then
    '            Do
    '                half = (high + low) \ 2
    '                If search > ApptGV.LongType(half) Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '            Loop Until low > high
    '        Else
    '            Do
    '                half = (high + low) \ 2
    '                If search > (ApptGV.LongType(half) Mod TimeSpan.TicksPerSecond) Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '            Loop Until low > high
    '        End If
    '        Return low
    '    End Function
    '    Function BiSearch_Long_GT(ByRef a() As Long, ByVal Find As Long, Optional ByVal base As Long = 1) As Long 'GT means GreaterThan
    '        Dim high As Integer = UBound(a)
    '        Dim low As Integer = base
    '        Dim half As Integer
    '        Do
    '            half = (high + low) \ 2
    '            If Find < a(half) Then high = half - 1 Else low = half + 1 'find greaterThan
    '        Loop Until low > high
    '        Return low
    '    End Function
    '    'Function biSearch_E_AllAppt(ByVal search As Long) As Integer
    '    '    'from: biSearch_E_iMeanAll
    '    '    Dim LocInFile As Integer = 0

    '    '    Dim HH As Integer = UBound(ApptGV.AllAppts_LongInteger)
    '    '    Dim high As Integer = HH
    '    '    Dim low As Integer = 1
    '    '    Dim half As Integer
    '    '    Do
    '    '        half = (high + low) \ 2
    '    '        If search > ApptGV.AllAppts_LongInteger(half).L Then low = half + 1 Else high = half - 1 ' if f > arrayItem 
    '    '    Loop Until low > high
    '    '    If low <= HH AndAlso search = ApptGV.AllAppts_LongInteger(low).L Then LocInFile = ApptGV.AllAppts_LongInteger(low).i
    '    '    Return LocInFile
    '    'End Function
    '    'Private Function BiSearch_Int_LTE(ByVal Find As Integer) As Integer 'LTE means LessThanOrEqualTo -Return of -1 means COULDN'T FIND
    '    '    Dim HH As Integer = UBound(posOfDates)
    '    '    Dim high As Integer = HH
    '    '    Dim low As Integer = 1
    '    '    Dim half As Integer
    '    '    Do
    '    '        half = (high + low) \ 2
    '    '        If Find > posOfDates(half) Then low = half + 1 Else high = half - 1 'GTE
    '    '    Loop Until low > high
    '    '    If low > HH Then 'must be done this way because LOW must be with array limits
    '    '        Return low - 1
    '    '    ElseIf Find = posOfDates(low) Then
    '    '        Return low
    '    '    ElseIf low > 1 Then
    '    '        Return low - 1
    '    '    Else
    '    '        Return -1 'error
    '    '    End If
    '    'End Function
    '    Private Function BiSearch_Long_LTE(ByVal Find As Long) As Integer 'LTE means LessThanOrEqualTo -Return of -1 means COULDN'T FIND
    '        Dim HH As Integer = UBound(ApptGV.AllAppts_LongInteger)
    '        Dim high As Integer = HH
    '        Dim low As Integer = 1
    '        Dim half As Integer
    '        Do
    '            half = (high + low) \ 2
    '            If Find > ApptGV.AllAppts_LongInteger(half).L Then low = half + 1 Else high = half - 1 'GTE
    '        Loop Until low > high
    '        If low > HH Then 'must be done this way because LOW must be with array limits
    '            Return low - 1
    '        ElseIf Find = ApptGV.AllAppts_LongInteger(low).L Then 'posOfDates(low) Then
    '            Return low
    '        ElseIf low > 1 Then
    '            Return low - 1
    '        Else
    '            Return -1 'error
    '        End If


    '    End Function
    '#End Region
    '    Private Sub WriteWordsTwoStrs(ByVal fileName As String, ByVal a() As TwoStringsType)
    '        If System.IO.File.Exists(fileName) = True Then Kill(fileName)
    '        System.IO.File.Create(fileName).Dispose()

    '        Dim i As Integer
    '        Using writer As BinaryWriter = New BinaryWriter(File.Open(fileName, FileMode.Open), System.Text.Encoding.ASCII)
    '            For i = 0 To UBound(a)
    '                writer.Write(a(i).s1)
    '                writer.Write(a(i).s2)
    '            Next
    '        End Using
    '    End Sub
    '    Function getBirthdaysForGivenDate(ByVal xdate As Date) 'ByVal mthDay As String) '03/18
    '        Dim mthDay As String
    '        mthDay = Format(xdate, "MM/dd").ToString

    '        Dim i As Integer
    '        Dim nRecs As Integer = UBound(ApptGV.Birthdays)
    '        Dim a(1000) As StringPointerType 'BirthdayRecordType
    '        Dim cnt As Integer = 0
    '        For i = 1 To nRecs

    '            If Mid(ApptGV.Birthdays(i).DateStr, 1, 5) = mthDay AndAlso ApptGV.Birthdays(i).DeleteFlag = 0 Then
    '                cnt += 1
    '                a(cnt).str = ApptGV.Birthdays(i).Msg
    '                a(cnt).ptr = ApptGV.Birthdays(i).ID
    '            End If
    '        Next
    '        ReDim Preserve a(cnt)
    '        Return a
    '    End Function
    '    Function getBirthdaysforGivenMonth(ByVal mth As Integer)
    '        Dim i As Integer
    '        Dim nRecs As Integer = UBound(ApptGV.Birthdays)
    '        Dim a(1000) As BirthdayRecordType
    '        Dim cnt As Integer = 0
    '        For i = 1 To nRecs
    '            If ApptGV.Birthdays(i).Month = mth AndAlso ApptGV.Birthdays(i).DeleteFlag = 0 Then
    '                cnt += 1
    '                a(cnt) = ApptGV.Birthdays(i)
    '            End If
    '        Next
    '        ReDim Preserve a(cnt)
    '        Return a
    '    End Function
End Module
