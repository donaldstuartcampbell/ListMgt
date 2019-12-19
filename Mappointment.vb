Option Explicit On
Imports System.Data.SqlClient
Module Mappointment
    Public Const BLACK_RIGHT_POINTING_TRIANGLE As Char = ChrW(&H25B6) '▶
    Public Const BLACK_LEFT_POINTING_TRIANGLE As Char = ChrW(&H25C0) '◀
    Structure ApptRecordsTypeNew 'changed to Public 8/18/2018 NOT DONE
        Dim ApptRec As ApptTodoRecordType
        Dim Minutes As Integer
        'Dim Seconds As Integer
        Dim ApptDate As Date 'convert dtics to date
        'Dim LocationInApptFil As Integer
        Dim LocationInApptRecordArray As Integer
    End Structure
    Structure ApptRecordsType 'changed to Public 8/18/2018 NOT DONE
        Dim ApptRec As ApptRecAllType
        Dim Minutes As Integer
        'Dim Seconds As Integer
        Dim ApptDate As Date 'convert dtics to date
        'Dim LocationInApptFil As Integer
        Dim LocationInApptRecordArray As Integer
    End Structure
    Structure ApptRecAllType
        'SP_RENAME 'scAppt.IsRecurring' , 'scUserID', 'COLUMN'

        '[ID] ,[dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag],[IntegerDate],[scUserID],[apptDateStr]
        Dim ID As Integer

        Dim dTics As Int64 'Date '8
        <VBFixedString(84)> Dim msg As String
        Dim AcctNum As Int32 '4
        Dim StrikeOut As Int16 '2 bytes 0 or 1
        Dim DeleteFlag As Int16 '2 bytes 0 or 1

        Dim IntegerDate As Integer
        Dim scUserID As Integer
        Dim apptDateStr As String
        Dim apptDate As Date 'new 10/27/2017

        Dim dTicsOriginal As Long
        Dim dTicsCompleted As Long

        'Public Sub New(i As Integer, d As Int64, m As String, a As Int32, s As Int16, df As Int16, idate As Integer, IsRecur As Integer, adatestr As String, adate As Date)
        '    'Public Sub New(i As Integer, d As Int64, m As <VBFixedString(84), a As Int32, s As Int16, df As Int16, idate As Integer, IsRecur As Integer, adatestr As String, adate As Date)
        '    'Public Sub New(i As Integer, d As Integer, m As String, a As Integer, s As Integer, df As Integer, idate As Integer, IsRecur As Integer, adatestr As String, adate As Date)
        '    ID = i
        '    dTics = d
        '    msg = m
        '    AcctNum = a
        '    StrikeOut = s
        '    DeleteFlag = df
        '    IntegerDate = idate
        '    scUserID = IsRecur
        '    apptDateStr = adatestr
        '    apptDate = adate

        'End Sub
    End Structure
    Structure ApptRecType '100 bybytes
        Dim dTics As Int64 'Date '8
        <VBFixedString(84)> Dim msg As String
        Dim AcctNum As Int32 '4
        Dim StrikeOut As Int16 '2 bytes 0 or 1
        Dim DeleteFlag As Int16 '2 bytes 0 or 1
    End Structure
    Function initApptRecAllType() As ApptRecAllType
        Dim rec As ApptRecAllType
        With rec
            .AcctNum = 0
            .apptDate = "01/01/0001"
            .apptDateStr = ""
            .DeleteFlag = 0
            .dTics = 0
            .dTicsCompleted = 0
            .dTicsOriginal = 0
            .ID = 0
            .IntegerDate = 0
            .msg = ""
            .scUserID = 0
            .StrikeOut = 0
        End With
        Return rec
    End Function
    Function initApptTodoRecord() As ApptTodoRecordType
        Dim rec As ApptTodoRecordType
        With rec
            .AcctNum = 0
            .apptDate = 0
            .ApptDateStr = ""
            .DeleteFlag = 0
            .dTics = 0
            .dTicsCompleted = 0
            .dTicsOriginal = 0
            .ID = 0
            .DeleteAbsolute = 0
            .Msg = ""
            .scUserID = 0
            .StrikeOut = 0
        End With
        Return rec
    End Function
    Function daysAgoSmallString(ByRef xDate As Date, Optional ByVal UseTriangle As Boolean = False) As String

        '0x25A0	9632	BLACK SQUARE	■
        Dim BLACK_SQUARE As Char = ChrW(&H25A0) '■

        'Public Const BLACK_RIGHT_POINTING_TRIANGLE As Char = ChrW(&H25B6) '▶
        'Public Const BLACK_LEFT_POINTING_TRIANGLE As Char = ChrW(&H25C0) '◀


        Dim dToday As Date = Date.Today

        Dim dDiff As Integer = DateDiff(DateInterval.Day, dToday, xDate)
        Dim dDiffStr As String
        If dDiff = 0 Then
            dDiffStr = "Today"
        ElseIf dDiff < 0 Then
            If dDiff = -1 Then
                dDiffStr = "1"
            Else
                dDiffStr = Math.Abs(dDiff).ToString ' & " days ago"
            End If
        Else
            If dDiff = 1 Then
                dDiffStr = "1" ' day hence"
            Else
                dDiffStr = dDiff.ToString ' & " days hence"
            End If
        End If
        If UseTriangle Then
            Select Case dDiff
                'Case 0 : dDiffStr = BLACK_RIGHT_POINTING_TRIANGLE & dDiffStr & BLACK_LEFT_POINTING_TRIANGLE
                'Case 0 : dDiffStr = BLACK_RIGHT_POINTING_TRIANGLE & " " & dDiffStr & " " & BLACK_LEFT_POINTING_TRIANGLE
                Case 0 : dDiffStr = BLACK_SQUARE & " " & dDiffStr ' & " " & BLACK_SQUARE

                Case Is > 0 : dDiffStr = BLACK_RIGHT_POINTING_TRIANGLE & " " & dDiffStr
                Case Is < 0 : dDiffStr = BLACK_LEFT_POINTING_TRIANGLE & " " & dDiffStr

                    'Case Is > 0 : dDiffStr &= BLACK_RIGHT_POINTING_TRIANGLE
                    'Case Is < 0 : dDiffStr &= BLACK_LEFT_POINTING_TRIANGLE

            End Select
        End If

        Return dDiffStr
    End Function
    Function daysAgoString(ByRef xDate As Date, Optional ByVal UseTriangle As Boolean = False) As String

        '0x25A0	9632	BLACK SQUARE	■
        Dim BLACK_SQUARE As Char = ChrW(&H25A0) '■

        'Public Const BLACK_RIGHT_POINTING_TRIANGLE As Char = ChrW(&H25B6) '▶
        'Public Const BLACK_LEFT_POINTING_TRIANGLE As Char = ChrW(&H25C0) '◀


        Dim dToday As Date = Date.Today

        Dim dDiff As Integer = DateDiff(DateInterval.Day, dToday, xDate)
        Dim dDiffStr As String
        If dDiff = 0 Then
            dDiffStr = "Today"
        ElseIf dDiff < 0 Then
            If dDiff = -1 Then
                dDiffStr = "1 day ago"
            Else
                dDiffStr = Math.Abs(dDiff).ToString & " days ago"
            End If
        Else
            If dDiff = 1 Then
                dDiffStr = "1 day hence"
            Else
                dDiffStr = dDiff.ToString & " days hence"
            End If
        End If
        ''''If UseTriangle Then
        ''''    Select Case dDiff
        ''''        'Case 0 : dDiffStr = BLACK_RIGHT_POINTING_TRIANGLE & dDiffStr & BLACK_LEFT_POINTING_TRIANGLE
        ''''        'Case 0 : dDiffStr = BLACK_RIGHT_POINTING_TRIANGLE & " " & dDiffStr & " " & BLACK_LEFT_POINTING_TRIANGLE
        ''''        Case 0 : dDiffStr = BLACK_SQUARE & " " & dDiffStr ' & " " & BLACK_SQUARE

        ''''        Case Is > 0 : dDiffStr = BLACK_RIGHT_POINTING_TRIANGLE & " " & dDiffStr
        ''''        Case Is < 0 : dDiffStr = BLACK_LEFT_POINTING_TRIANGLE & " " & dDiffStr

        ''''            'Case Is > 0 : dDiffStr &= BLACK_RIGHT_POINTING_TRIANGLE
        ''''            'Case Is < 0 : dDiffStr &= BLACK_LEFT_POINTING_TRIANGLE

        ''''    End Select
        ''''End If

        Return dDiffStr
    End Function

    Function getApptForSelectedDateNew(ByVal sDate As Date) As ApptRecordsTypeNew() 'ApptRecordsType() 'added 9/14/2016
        'Dim apptRecs() As ApptRecAllType = getApptRecords_GivenDate_retApptRecType(sDate.Ticks)
        Dim apptRecs() As ApptTodoRecordType = getApptRecords_GivenDate_ApptTodoType(sDate)

        Dim nRecs As Integer = UBound(apptRecs)
        Dim s(nRecs) As ApptRecordsTypeNew 'ApptRecordsType
        Dim i As Integer
        Dim c As Integer = 0
        For i = 1 To nRecs
            'Debug.Print(apptRecs(i).dTics Mod TimeSpan.TicksPerMinute & " " & apptRecs(i).dTics & " " & apptRecs(i).Msg)
            'If apptRecs(i).dTics Mod TimeSpan.TicksPerMinute = 0 Then '???'remmed out 11/24/2019
            If apptRecs(i).dTics Mod TimeSpan.TicksPerMinute = 0 Then
                MsgBox("error in: getApptForSelectedDateNew (counter=0) - will end")
                End
            End If
            c += 1
                With s(c)
                    .ApptRec = apptRecs(i)
                    'Now set Actual DATE and MINUTES
                    .ApptDate = New Date(apptRecs(i).dTics)
                    .Minutes = DatePart(DateInterval.Hour, .ApptDate) * 60 + DatePart(DateInterval.Minute, .ApptDate)
                    '.LocationInApptFil = 0 ' ApptGV.AllAppts_LongInteger(k).i
                    '.LocationInApptRecordArray = i
                End With
            'End If
        Next
        ReDim Preserve s(c)
        Return s

    End Function
    'Function getApptForSelectedDate(ByVal sDate As Date) As ApptRecordsType() 'added 9/14/2016

    '    Dim apptRecs() As ApptRecAllType = getApptRecords_GivenDate_retApptRecType(sDate.Ticks)
    '    Dim nRecs As Integer = UBound(apptRecs)
    '    Dim s(nRecs) As ApptRecordsType
    '    Dim i As Integer
    '    Dim c As Integer = 0
    '    For i = 1 To nRecs
    '        If apptRecs(i).dTics Mod TimeSpan.TicksPerMinute = 0 Then
    '            c += 1
    '            With s(c)
    '                .ApptRec = apptRecs(i)
    '                'Now set Actual DATE and MINUTES
    '                .ApptDate = New Date(apptRecs(i).dTics)
    '                .Minutes = DatePart(DateInterval.Hour, .ApptDate) * 60 + DatePart(DateInterval.Minute, .ApptDate)
    '                '.LocationInApptFil = 0 ' ApptGV.AllAppts_LongInteger(k).i
    '                '.LocationInApptRecordArray = i
    '            End With
    '        End If
    '    Next
    '    ReDim Preserve s(c)
    '    Return s


    'End Function
    'Function getApptRecords_GivenDate_retApptRecType(ByVal selectedDateTicks As Long, Optional ByVal AppointmentsOnly As Boolean = False) As ApptRecAllType() 'getApptRecs(ByVal xDate As Date)
    '    Dim sTicks As Long = selectedDateTicks ' xDate.Ticks
    '    Dim nextDayTicks As Long = selectedDateTicks + TimeSpan.TicksPerDay  'DateAdd(DateInterval.Day, 1, xDate).Ticks

    '    Dim Q As String
    '    Q = "select * " ' [dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag]" 'match ApptRecType"
    '    If ApptGV.DBbeingUsed = "SC" Then
    '        Q &= " from [scAppt]"
    '    Else
    '        Q &= " from [Appt]"
    '    End If

    '    Q &= " where scUserId=" & ApptGV.UserID & " and dtics>=" & sTicks & " and dtics<" & nextDayTicks & " and msg<>'' and deleteflag=0"
    '    If AppointmentsOnly = True Then
    '        Q &= " and substring(apptdatestr,18,2)='00'"
    '    End If
    '    Q &= " order by dtics"

    '    Return getApptRecords_AllFields(Q) 'getApptRecords(Q)

    'End Function
    ''Function getApptRecords_AllFields(ByVal Q As String) As ApptRecAllType()
    ''    'ALTER table [scAppt] ALTER COLUMN IsRecurring INT use this on 8/16/2017 at 9:46am to change
    ''    '.scUserID = (reader.GetBoolean(7)) from bit field to int field so now use .scUserID = (reader.GetInt32(7))
    ''    Dim cnt As Integer = 0
    ''    Dim apptRecs(1000000) As ApptRecAllType

    ''    Dim connection As New SqlConnection(ApptGV.SQLconnString)
    ''    Using connection
    ''        Dim command As SqlCommand = New SqlCommand(Q, connection)
    ''        connection.Open()
    ''        Dim reader As SqlDataReader = command.ExecuteReader()

    ''        If reader.HasRows Then
    ''            Do While reader.Read()
    ''                cnt += 1
    ''                '[ID] ,[dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag],[IntegerDate],[scUserID],[apptDateStr]
    ''                With apptRecs(cnt)
    ''                    '[dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag]

    ''                    .ID = (reader.GetInt32(0))

    ''                    .dTics = (reader.GetInt64(1))
    ''                    .msg = (reader.GetString(2))
    ''                    .AcctNum = (reader.GetInt32(3))
    ''                    .StrikeOut = (reader.GetInt16(4))
    ''                    .DeleteFlag = (reader.GetInt16(5))

    ''                    .IntegerDate = (reader.GetInt32(6))
    ''                    .scUserID = (reader.GetInt32(7))
    ''                    .apptDateStr = (reader.GetString(8))
    ''                    .apptDate = (reader.GetDateTime(9))

    ''                    ' Debug.Print(.apptDateStr & " " & .apptDate)
    ''                End With
    ''            Loop
    ''        Else
    ''            'Console.WriteLine("No rows found.")
    ''        End If
    ''        reader.Close()
    ''    End Using
    ''    ReDim Preserve apptRecs(cnt)
    ''    Return apptRecs

    ''End Function

    'Sub AddApptRecord(ByVal rec As ApptRecAllType) ', Optional ByVal useTODO_DB As Boolean = False) 'use for recurring records  'new 9/5/2015
    '    ',[dTicsOriginal] ,[dTicsCompleted]
    '    Dim xDate As String = formatApptDateStr(rec.dTics)
    '    Dim aDate As Date = createPureDateFromTicks(rec.dTics)
    '    Dim L As Long = Date.Now.Ticks
    '    Dim Q As String = ""
    '    'If ApptGV.DBbeingUsed = "SC" Then Q = "INSERT INTO [scAppt]" Else Q = "INSERT INTO [Appt]"
    '    Q = "INSERT INTO [Appt] "
    '    Q &= "([dTics], [Msg], [AcctNum], [StrikeOut], [DeleteFlag], [IntegerDate], [scUserID], [apptDateStr], [apptDate], [dTicsOriginal], [dTicsCompleted])"
    '    Q &= " VALUES("

    '    With rec
    '        Q &= .dTics
    '        Dim xMsg As String = .msg
    '        SQLreplace(xMsg)
    '        Q &= ", '" & xMsg & "'"
    '        Q &= ", " & .AcctNum
    '        Q &= ", " & .StrikeOut
    '        Q &= ", " & .DeleteFlag
    '        Q &= ",0"
    '        Q &= ", " & .scUserID
    '        Q &= ",'" & xDate & "'"
    '        Q &= ",'" & aDate & "'"
    '        Q &= ", " & L '.dTics
    '        Q &= ", " & 0
    '        Q &= ")"


    '    End With
    '    'Dim csNew As String = ""
    '    'If useTODO_DB = False Then csNew = ApptGV.SQLconnString Else csNew = ApptGV.SQLconnStringTODO
    '    'Debug.Print(Q)
    '    Try

    '        Using sqlCon = New SqlClient.SqlConnection(ApptGV.SQLconnString) 'New SqlConnection(connStr)
    '            'Using sqlCon = New SqlClient.SqlConnection(csNew) 'New SqlConnection(connStr)

    '            Try
    '                sqlCon.Open()
    '            Catch ex As Exception
    '                MsgBox("Add Appointment Failed when Opening Database: see message below" & vbNewLine & ex.Message)
    '                Exit Sub
    '            End Try

    '            Dim cmd = New SqlClient.SqlCommand(Q, sqlCon) 'New SqlCommand(sqlText, sqlCon)
    '            cmd.ExecuteNonQuery()
    '            sqlCon.Close()
    '        End Using
    '    Catch ex As Exception
    '        MsgBox("Add Appointment Failed: see message below" & vbNewLine & ex.Message)
    '    End Try
    'End Sub
    Function formatApptDateStr(ByVal dTics As Long) As String
        Dim x As String = Format(New DateTime(dTics), "yyyy/MM/dd HH:mm:ss.fffffff")
        Return x
    End Function
    Function createPureDateFromTicks(ByVal x As Long) As Date
        'createPureDateFromTicks
        Dim xDate As Date = New Date((x \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay)
        Return xDate
    End Function
    'Function getApptRecords_GivenDate_retApptRecType(ByVal selectedDateTicks As Long, Optional ByVal AppointmentsOnly As Boolean = False) As ApptRecAllType() 'getApptRecs(ByVal xDate As Date)
    '    Dim sTicks As Long = selectedDateTicks ' xDate.Ticks
    '    Dim nextDayTicks As Long = selectedDateTicks + TimeSpan.TicksPerDay  'DateAdd(DateInterval.Day, 1, xDate).Ticks

    '    Dim Q As String
    '    Q = "select * " ' [dTics],[Msg],[AcctNum],[StrikeOut],[DeleteFlag]" 'match ApptRecType"
    '    Q &= " from [appt]"
    '    Q &= " where dtics>=" & sTicks & " and dtics<" & nextDayTicks & " and msg<>'' and deleteflag=0"
    '    If AppointmentsOnly = True Then
    '        Q &= " and substring(apptdatestr,18,2)='00'"
    '    End If
    '    Q &= " order by dtics"

    '    Return getApptRecords_AllFields(Q) 'getApptRecords(Q)

    'End Function
    'Private Function setAddAppointmentPoint() As Point
    '    'Dim boundWidth As Integer = Screen.PrimaryScreen.Bounds.Width
    '    Dim boundWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width

    '    'Dim LeftPos As Integer = Me.Left + PanelTodo.Left - 150
    '    'Dim LeftPos As Integer = Me.Left + PanelTodo.Left + PanelTodo.Width - fAddAppointment.Width ' - 150
    '    Dim LeftPos As Integer = Form.Left + Panel.Left + apptT.StartOfApptMsg + CInt(apptT.V) + 165 ' + PanelTodo.Width - fAddAppointment.Width ' - 150
    '    'apptT.StartOfApptMsg - apptT.V

    '    If LeftPos + fAddAppointment.Width > boundWidth Then
    '        LeftPos = boundWidth - fAddAppointment.Width 'fMinutePicker.Width
    '    End If

    '    Dim boundHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
    '    'Dim boundHeight As Integer = Screen.PrimaryScreen.Bounds.Height

    '    Dim TopPos As Integer = Form.Top + Panel.Top + 70 ' + PanelTop.Height '+ 60 '70 '90
    '    'Dim TopPos As Integer = Me.Top + apptT.V 'apptT.WnoV '.V

    '    If (TopPos + fAddAppointment.Height) > boundHeight Then
    '        'If (TopPos + fAddAppointment.Height + fAddAppointment.PanelBottom.Height) > boundHeight Then
    '        TopPos = boundHeight - fAddAppointment.Height 'fMinutePicker.Height
    '    End If
    '    Dim p As New Point(LeftPos, TopPos)
    '    Return p
    'End Function
    '   Sub UpdateApptRecord_usingID(ByVal id As Integer, ByVal rec As ApptRecAllType) 'use for recurring records 'new 9/5/2015
    '       'usingID this is where the dtics is changed and therefore necessary to update using ID
    '       Dim Q As String
    '       Dim xDate As String = formatApptDateStr(rec.dTics)
    '       Dim xMsg As String = rec.msg
    '       SQLreplace(xMsg)
    '       Dim aDate As Date = createPureDateFromTicks(rec.dTics)

    '       With rec
    '           If ApptGV.DBbeingUsed = "SC" Then
    '               Q = "UPDATE [scAppt]"
    '           Else
    '               Q = "UPDATE [Appt]"
    '           End If

    '           Q &= " SET [dTics] = " & .dTics
    '           Q &= " ,[Msg] = '" & xMsg & "'"
    '           Q &= " ,[AcctNum] = " & .AcctNum
    '           Q &= " ,[StrikeOut] = " & .StrikeOut
    '           Q &= " ,[DeleteFlag] = " & .DeleteFlag
    '           Q &= " ,[IntegerDate] =" & .IntegerDate
    '           Q &= " ,[scUserID] =" & .scUserID
    '           Q &= " ,[apptDateStr] ='" & xDate & "'"
    '           Q &= " ,[apptDate] ='" & aDate & "'"

    '           ' Q &= " ,[dTicsOriginal] = " & .dTics

    '           Q &= " WHERE ID = " & id

    '       End With
    '       'Debug.Print(Q)
    '       Try
    '           passQ_return(Q)
    '       Catch ex As Exception
    '           MsgBox(ex.Message)
    '           End
    '       End Try

    '   End Sub
    '   Sub FixTotoDticsOriginal()
    '       'Date.Now.Ticks
    '       Dim recs() As ApptRecAllType = getApptRecords_AllFields("select * from scAppt where apptdate='1/1/0005'")
    '       Dim nItems As Integer = UBound(recs)
    '       Dim i As Integer
    '       For i = 1 To nItems
    '           'Dim Q As String
    '           'Q = "update dbo.scAppt set dticsOriginal=" & Date.Now.Ticks & " where id=" & recs(i).ID
    '           'Qexe(Q)
    '           'Debug.Print(recs(i).msg & " " & recs(i).dTicsOriginal)
    '           Debug.Print(recs(i).dTicsOriginal)
    '       Next
    '   End Sub
    '   Function getTodoRecs(ByVal UserID As Integer) As ApptRecAllType()
    '       Dim r() As ApptRecAllType
    '       Dim Q As String
    '       If ApptGV.DBbeingUsed = "SC" Then
    '           Q = "select * from scAppt where apptDate='0005/1/1' and scUserID=" & UserID & " and deleteFlag=0 order by dtics"
    '       Else
    '           Q = "select * from Appt where apptDate='0005/1/1' and scUserID=" & UserID & " and deleteFlag=0 order by dtics"
    '       End If
    '       r = getApptRecords_AllFields(Q)
    '       Return r
    '   End Function
    '   Function getTodoRecs_IncludingDeletedToday(ByVal UserID As Integer, Optional ByVal DeletesAtBottom As Boolean = False) As ApptRecAllType()
    '       Dim r() As ApptRecAllType
    '       'Note: (deleteflag ^ 1) has the affect to turning (0 to 1) and (1 to 0)
    '       '*(dtics % 10000000) 
    '       '1. @days2today gets the number of days from 1/1/0001 to today
    '       '2. cast(dticsCompleted/864000000000 as int) computes number of days from 1/1/0001 to dticsCompleted date
    '       '3. if the diff is 0 then it's as TODAY completed record
    '       'Order by
    '       'order by deleteflag, dtics - puts deleted or completed records at the bottom
    '       'order by dtics - keeps existing order (deletes are intermingled with actives)



    '       Dim Q As String
    '       'difference is in the 'order by' statement
    '       If DeletesAtBottom = True Then
    '           Q = "declare @days2today int =datediff(day,'00010101',getdate())
    'select * from appt where apptDate='0005/1/1' and scUserID=" & UserID & "  
    'and (deleteFlag=0 or (deleteflag=1 and @days2today-cast(dticsCompleted/864000000000 as int)=0)) 
    'order by deleteflag,dtics*(deleteflag ^1),dTicsCompleted"

    '           'order by deleteflag,dtics 'sorts inactives at bottom but retains the ORDER number so that the BOTTOM items will not be sorted by date deleted!

    '           'comments
    '           'deleteflag,dtics*(deleteflag ^1),dTicsCompleted
    '           '1. deleteflag separates active (top) from inactive (bottom)
    '           '2. dtics*(deleteflag ^1)
    '           '   has effect on TOP to order by order dtics % 10000000
    '           '   since (deleteflag ^1) is ZERO if deleteFlag is 1 - for Bottom items (inactives) this is a constant 0
    '           '3. finally, dticsCompleted has no effect on TOP items as dticsCompleted field is ZERO for actives but causes Bottom items to be sorted by dateTime deleted
    '           'NOTE: Top  Bottom
    '           '        0       1
    '           'dtics * 1  ...* 0
    '           '        0  dticsCompleted

    '           'NOTE: if i want to order by existing order ORDER but at the BOTTOM then use: 'order by deleteflag,dtics'
    '           'NOTE: if i want to order by existing order ORDER then use: order by deleteflag,dtics (Not at bottome)


    '           ' order by deleteflag, dtics"
    '       Else
    '           Q = "declare @days2today int =datediff(day,'00010101',getdate())
    'select * from appt where apptDate='0005/1/1' and scUserID=" & UserID & "  
    'and (deleteFlag=0 or (deleteflag=1 and @days2today-cast(dticsCompleted/864000000000 as int)=0))
    'order by dtics"
    '       End If

    '       '==================
    '       'Dim fontReg As Font = DGV_Todo.Font 'DGV_Printing.Font
    '       'Dim fontForStrikeOut As Font = New Font("Microsoft Sans Serif", 8, FontStyle.Strikeout)



    '       'Debug.Print(Q)
    '       'Oct 1 2018
    '       r = getApptRecords_AllFields(Q) ', useTODO_DB)
    '       Return r
    '   End Function





    '   Function getDeletedTodos_withingDateRange(ByVal sDate As Date, ByVal eDate As Date, Optional ByVal useTODO_DB As Boolean = False) As ApptRecAllType()
    '       Dim Q As String
    '       If ApptGV.DBbeingUsed = "SC" Then
    '           Q = "declare @NumDays2Today int= datediff(day,'0001/01/01',getdate())
    '       declare @today datetime=cast(getdate() as date)
    '       select * FROM [SealcoatCorp2].[dbo].[scAppt] where scUserID=" & ApptGV.UserID & "
    '       and dtics % 10000000>0 and dTicsCompleted >0
    '       and dateadd(day,cast(dTicsCompleted/864000000000 as int)-@numdays2today,@today)>='" & sDate & "' 
    '       and dateadd(day,cast(dTicsCompleted/864000000000 as int)-@numdays2today,@today)<='" & eDate & "' order by dTicsCompleted"

    '       Else

    '           Q = "declare @NumDays2Today int= datediff(day,'0001/01/01',getdate())
    ' declare @today datetime=cast(getdate() as date)
    ' select * FROM [Appt] where scUserID=" & ApptGV.UserID & "
    ' and dtics % 10000000>0 and dTicsCompleted >0
    ' and dateadd(day,cast(dTicsCompleted/864000000000 as int)-@numdays2today,@today)>='" & sDate & "' 
    ' and dateadd(day,cast(dTicsCompleted/864000000000 as int)-@numdays2today,@today)<='" & eDate & "' order by dTicsCompleted"
    '       End If

    '       Dim r() As ApptRecAllType
    '       r = getApptRecords_AllFields(Q) ', useTODO_DB)
    '       Return r
    '   End Function
    '   Function findApptRecords(ByVal find As String, Optional ByVal includeTodos As Boolean = False) As ApptRecAllType()

    '       Dim r(0) As ApptRecAllType
    '       If find = "" Then
    '           Return r
    '       End If
    '       Dim Q As String
    '       'Q = "select * from appt where apptdate<>'1/1/0005' and charindex('t',msg)>0 order by dtics"
    '       'select * from appt where apptdate<>'1/1/0005' and charindex('t',msg)>0 order by dtics
    '       'Select * From appt Where apptdate ='1/1/0005' and deleteflag=0 and charindex('t',msg)>0 order by dtics
    '       Q = "declare @find varchar(100)='" & find & "'"
    '       If includeTodos = False Then
    '           Q &= " Select * from appt where apptdate<>'1/1/0005' and charindex(@find,msg)>0 order by dtics"
    '       Else
    '           Q &= " Select * from appt where deleteflag=0 and charindex(@find,msg)>0 order by substring(apptdatestr,1,1) desc, dtics"
    '       End If



    '       r = getApptRecords_AllFields(Q) ', useTODO_DB)
    '       Return r

    '   End Function
End Module
