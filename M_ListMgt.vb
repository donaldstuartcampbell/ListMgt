
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
End Module
