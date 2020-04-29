
Option Explicit On
Module M_ListMgt
    Structure tempType
        Dim RecNumber As Integer
        Dim Completed As Integer
        Dim Deleted As Integer
        Dim msg As String
    End Structure
    Function rTrimF(ByVal s As String)
        Dim r As String = RTrim(s)
        Dim chars2RemoveAtEndOfLine() As Char = {vbCr, vbLf}
        r = RTrim(r.TrimEnd(chars2RemoveAtEndOfLine))
        Return r
    End Function
    Function TrimF(ByVal s As String)
        Dim r As String = Trim(s)
        Dim chars2RemoveAtEndOfLine() As Char = {vbCr, vbLf}
        r = Trim(r.TrimEnd(chars2RemoveAtEndOfLine))
        Return r
    End Function
    Sub test_trimF()
        Dim s As String
        s = "now is the time.   " & vbNewLine & vbNewLine & vbNewLine & vbNewLine & "  "
        s = TrimF(s)
        Dim a() As String = Split(s, vbNewLine)
        Dim i As Integer
        For i = 0 To UBound(a)
            Debug.Print(i.ToString & " " & a(i))
        Next
    End Sub
    Function getRecs()
        Dim nRecs As Integer = 20
        Dim i As Integer
        Dim r(nRecs) As tempType
        Dim s As String = "now is the time for all good men to come the rescue of their countryment. Now is the time for all good men to come the the rescue of their countrymen."
        For i = 1 To nRecs
            With r(i)
                .RecNumber = i
                .Completed = i Mod 2
                .Deleted = 0
                .msg = s
            End With
        Next
        Dim ss As String = r(4).msg & r(4).msg
        r(4).msg = ss

        ss = Mid(ss, 1, 50)
        r(3).msg = ss

        ss = "1. Number One Item"
        ss &= vbNewLine
        ss &= "2. Number Two Item"
        r(7).msg = ss
        Return r
    End Function
    Sub testing_DeletingRecord()
        'startTime()

        setPath2use()
        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
        set_DeleteDateArray() 'should be called CompletedDateArray
        'endTime() '.028 seconds

        'startTime()
        '================
        Dim i As Integer
        Dim completedDate As Date
        Dim msg As String
        For i = 1 To ApptGV.All_ApptTodoRecords_Cnt
            With ApptGV.All_ApptTodoRecords(i)
                completedDate = New Date(.dTicsCompleted)
                msg = Replace(.Msg, vbNewLine, "^")
                'Debug.Print(completedDate)
                If .DeleteAbsolute = 1 Then
                    Debug.Print(i & " " & .DeleteAbsolute & "=abs - " & .DeleteFlag & " del - comDate=" & completedDate & " -- " & .ApptDateStr & " " & msg)
                Else
                    Debug.Print(i & "         " & .DeleteFlag & " del - comDate=" & completedDate & " -- " & .ApptDateStr & " " & msg)
                End If

            End With
        Next
        'endTime()'.001 second
    End Sub
End Module
