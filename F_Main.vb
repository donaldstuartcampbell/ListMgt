Imports System.Data.DataTable
Public Class F_Main

    'Dim SourceRowIndex As Integer = -1
    'Dim TargetRowIndex As Integer = -1
    'Dim dtSource As DataTable
    'Dim setSelectedRow As Boolean = False

    Dim formOrigWidth As Single
    Dim skip As Boolean = True
    Dim msgNum As Integer = 2

    'Dim Recs() As tempType 'never changes except when new rec added
    'Dim Recs2Use() As tempType 'holds only what is displayed
    Dim Rec2Edit As ApptTodoRecordType  'tempType
    Private Sub F_Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load




        'dtSource .Rows.Add()




        'set file system to use
        setPath2use()


        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
        set_CreateDateArray()


        ApptGV.Use_ApptTodoRecords = ApptGV.All_ApptTodoRecords.Clone
        'refreshDataGridView1()
        '===

        'Dim rec As ApptTodoRecordType
        'Debug.Print("len Of rec =" & Len(rec))



        'Recs = getRecs()
        'Recs2Use = Recs.Clone

        skip = True
        initSetUp()

        refreshDataGridView1()

        skip = False
    End Sub
    Private Sub initSetUp()
        'AddHandler DataGridView1.CurrentCellDirtyStateChanged, AddressOf DataGridView1_CurrentCellDirtyStateChanged
        'AddHandler DataGridView1.CellValueChanged, AddressOf DataGridView1_CellValueChanged

        RadioButton1.Checked = True

        Me.Top = 0
        Me.Left = 0
        Me.MinimumSize = Me.Size

        formOrigWidth = Me.Width



        With DataGridView1
            .SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            '.TabIndex = 0

            '.ColumnHeadersVisible = False
            '.EnableHeadersVisualStyles = False
            '.RowHeadersDefaultCellStyle.NullValue = 0

            Dim i As Integer
            For i = 0 To .Columns.Count - 1
                .Columns.Item(i).SortMode = DataGridViewColumnSortMode.Programmatic
            Next i



            .MultiSelect = False ' this allows only 1 items to be selected at a time!


            'To make the completed column editable do the following:
            '1. set DGV readonly to false
            '2. set the individual columns that are not editable to readonly.true
            'note: doesn't work the other way aroung!
            '------------------------
            .ReadOnly = False
            .Columns(0).ReadOnly = True
            .Columns(msgNum).ReadOnly = True

            '.ReadOnly = True



            '.Columns("Completed").ReadOnly = False
            '.Columns(2).ReadOnly = False
            '------------------------

            .AllowUserToAddRows = False
            .AllowUserToOrderColumns = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False



            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige


            'Dim newPadding As Padding = New Padding(0, 5, 0, 5)
            '.RowTemplate.DefaultCellStyle.Padding = newPadding
            '.Columns(1).HeaderCell.Style.Padding = New Padding(20, 20, 20, 20) '0, 5, 0, 0)
            .Columns(msgNum).DefaultCellStyle.Padding = New Padding(5, 5, 5, 5) '(20, 20, 20, 20) '0, 5, 0, 0)

            .Columns(msgNum).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders




            '.Columns(0).Font = New Font("microsoft sans serif", 20, FontStyle.Regular)
            .Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", 10, FontStyle.Regular)
            '.Columns(1).DefaultCellStyle.Font.Size = 20 'dg.font.size / 2
            '.Columns(0).DefaultCellStyle.Font.Size = 20 'dg.font.size / 2

            '.Columns(5).DefaultCellStyle.Font = New Font(dg.font, emSize:=(dg.font.size / 2))

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter '.MiddleRight

            .RowHeadersVisible = False 'this gets rid of left column!


        End With

        'test_sampleItems()
    End Sub
    'Private Sub test_sampleItems()
    '    'TextBox1.Text = "now is the time for all good men to come the rescue of their countryment. Now is the time for all good men to come the the rescue of their countrymen."
    '    'Debug.Print(Len(TextBox1.Text))

    '    Dim i As Integer
    '    DataGridView1.Rows.Clear()
    '    For i = 1 To UBound(Recs2Use)
    '        'DataGridView1.Rows.Add(i.ToString("0000"), i Mod 2, TextBox1.Text, "Edit", "Delete")
    '        With Recs2Use(i)
    '            'DataGridView1.Rows.Add(.RecNumber.ToString, .Completed, .msg, "Edit", "Delete")
    '            DataGridView1.Rows.Add(i.ToString, .Completed, .msg, "Edit", "Delete")
    '        End With
    '        'DataGridView1.Rows.Add(i.ToString, i Mod 2, TextBox1.Text, "Edit", "Delete")

    '        'DataGridView1.Rows(i).Cells(0).Value = i.ToString("0000")
    '        'DataGridView1.Rows(i).Cells(1).Value = TextBox1.Text

    '        'If i Mod 2 = 0 Then
    '        '    DataGridView1.Rows(i - 1).DefaultCellStyle.BackColor = Color.Gainsboro '.LightGray
    '        'Else

    '        '    DataGridView1.Rows(i - 1).DefaultCellStyle.BackColor = Color.White
    '        'End If
    '    Next


    'End Sub
    Private Sub F_Main_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If skip Then Exit Sub
        DataGridView1.Columns(msgNum).Width = 500 + (Me.Width - formOrigWidth)



        Dim s As Integer = 10
        If Me.Height > 600 Then s = 10 + (Me.Height - 600) / 200

        Debug.Print(Me.Height & " " & s)
        DataGridView1.Columns(0).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
        DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
        DataGridView1.Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If skip Then Exit Sub
        DataGridView1.Rows(0).Cells(0).Value = 1.ToString '("0000")
        DataGridView1.Rows(0).Cells(msgNum).Value = rTrimF(TextBox1.Text)

        'Dim MyNewString As String = MyString.Replace(Environment.NewLine,String.Empty)
    End Sub


    'This will fire off after CurrentCellDirtyStateChanged occured...
    'You can get row or column index from e as well here...
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        'Do what you need to do...
        If skip Then Exit Sub
        If e.ColumnIndex <> 1 Then Exit Sub
        If DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = False Then '"True" Then
            'Checked condition'
            'Debug.Print("UnChecked" & " " & ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).Msg)

            'two things
            '1. update record (save)
            '2. find rec in all... and update that rec
            '3. update rec in use...

            ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTicsCompleted = 0
            ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).DeleteFlag = 0
            UpdateApptTodoRecord(ApptGV.Use_ApptTodoRecords(e.RowIndex + 1))

        Else
            'Unchecked Condition'
            'Debug.Print("Checked" & " " & ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).Msg)

            ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTicsCompleted = Now.Ticks
            ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).DeleteFlag = 1
            UpdateApptTodoRecord(ApptGV.Use_ApptTodoRecords(e.RowIndex + 1))

        End If

    End Sub

    'This will fire immediately when you click in the cell...
    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If skip Then Exit Sub
        'If e.ColumnIndex <> 1 Then Exit Sub

        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Select Case e.ColumnIndex
            Case 3 'edit
                Debug.Print("edit " & " " & e.RowIndex & " len=" & Len(ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).Msg))
                TextBox1.Text = ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).Msg

                Rec2Edit = ApptGV.Use_ApptTodoRecords(e.RowIndex + 1)
            Case 4 'delete
                Debug.Print("delete " & e.RowIndex)
        End Select
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If skip Then Exit Sub
        'Recs2Use = Recs.Clone
        'test_sampleItems()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            'Dim nRecs As Integer = UBound(Recs)
            'Dim i As Integer
            'ReDim Recs2Use(nRecs)
            'Dim cnt As Integer = 0
            'For i = 1 To nRecs
            '    With Recs(i)
            '        If .Completed = 0 Then
            '            cnt += 1
            '            Recs2Use(cnt) = Recs(i)
            '        End If
            '    End With
            'Next
            'ReDim Preserve Recs2Use(cnt)
            'test_sampleItems()
        End If


    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            'Dim nRecs As Integer = UBound(Recs)
            'Dim i As Integer
            'ReDim Recs2Use(nRecs)
            'Dim cnt As Integer = 0
            'For i = 1 To nRecs
            '    With Recs(i)
            '        If .Completed = 1 Then
            '            cnt += 1
            '            Recs2Use(cnt) = Recs(i)
            '        End If
            '    End With
            'Next
            'ReDim Preserve Recs2Use(cnt)
            'test_sampleItems()
        End If
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


    '    Dim nRecs As Integer = UBound(Recs2Use)
    '    Dim i As Integer
    '    For i = 1 To nRecs
    '        If Recs2Use(i).RecNumber = Rec2Edit.RecNumber Then
    '            Recs2Use(i).msg = rTrimF(TextBox1.Text)

    '            DataGridView1.Rows(i - 1).Cells(msgNum).Value = Recs2Use(i).msg
    '            Exit For
    '        End If
    '    Next

    '    nRecs = UBound(Recs)

    '    For i = 1 To nRecs
    '        If Recs(i).RecNumber = Rec2Edit.RecNumber Then
    '            Recs(i).msg = rTrimF(TextBox1.Text)
    '            Exit For
    '        End If
    '    Next


    'End Sub
    Private Sub Move2Row(ByVal recNumber As Integer)
        Dim nRecs As Integer = UBound(ApptGV.Use_ApptTodoRecords)
        Dim i As Integer
        For i = 1 To nRecs 'ApptGV.All_ApptTodoRecords_Cnt
            If ApptGV.Use_ApptTodoRecords(i).ID = recNumber Then
                DataGridView1.Rows(i - 1).Cells(2).Selected = True
                DataGridView1.Select()

                'DataGridView1.VerticalScrollingOffset(i - 1) '.ScrollIntoView(i - 1)
                DataGridView1.FirstDisplayedScrollingRowIndex = i - 1 'this.dataGridView2.FirstDisplayedScrollingRowIndex
                Exit For
            End If
        Next

    End Sub
    Private Sub B_AddTodo_Click(sender As Object, e As EventArgs) Handles B_AddTodo.Click


        'see: Function ModifyTodoMsg

        'first step - set todo list to today month and year
        'ApptGV.Todo_SelDate = adjDateToFirstDayOfMonth(Date.Today)
        'setTodo_CalendarInfo()
        '------------------------
        F_AddTodo.PassCanRollover = False

        F_AddTodo.AddingNewToDo = True
        F_AddTodo.PassMessage = ""
        F_AddTodo.ShowDialog()
        '============

        'Debug.Print(Format(Now, "o") & " start")
        'Dim dd As Long = Now.Ticks
        'startTime()

        If F_AddTodo.RetMessage = "" Then
            Exit Sub
        End If
        'Else
        Dim rec As ApptTodoRecordType = initApptTodoRecord()

        If F_AddTodo.ReturnPutOnTop = False Then
            'Debug.Print("put on BOTTOM")
            'Exit Sub

            rec.Msg = F_AddTodo.RetMessage
            rec.dTics = ApptGV.todoTicks
            AddApptTodoRecord(rec) 'added to ApptGV.All_ApptTodoRecords
            'rec added - now do the following:
            '1. refresh dataGridView1

            ApptGV.Use_ApptTodoRecords = ApptGV.All_ApptTodoRecords.Clone

            refreshDataGridView1()

            Move2Row(rec.ID)
            Exit Sub

            '=============

            'Dim nRecs As Integer = UBound(Recs) + 1

            'ReDim Preserve Recs(nRecs)
            'With Recs(nRecs)
            '    .RecNumber = nRecs
            '    .Completed = 0
            '    .Deleted = 0
            '    .msg = F_AddTodo.RetMessage

            '    DataGridView1.Rows.Add(.RecNumber.ToString, .Completed, .msg, "Edit", "Delete")
            'End With
            'Recs2Use = Recs.Clone

            'Move2Row(nRecs)



            'SetDGtodo()

            'Set_appt_StartDate_Records_DGprinting()
            '===================
            'Else
            '    Dim xtodos() As ApptTodoRecordType = getApptRecords_AllFields_ApptTodo_NEW(SelectedYearMonthLong, Not CheckBoxShowTodosDeletedToday.Checked) '12/1/2019
            '    Dim nRecs As Integer = UBound(xtodos)
            '    '==
            '    If nRecs = 0 Then
            '        rec.Msg = F_AddTodo.RetMessage
            '        rec.dTics = ApptGV.todoTicks
            '        AddApptTodoRecord(rec)

            '        SetDGtodo()

            '        Set_appt_StartDate_Records_DGprinting()

            '    Else
            '        rec.Msg = F_AddTodo.RetMessage
            '        rec.dTics = ApptGV.todoTicks
            '        AddApptTodoRecord(rec, True)
            '        '==
            '        'startTime()
            '        Dim bottomDics As Long = rec.dTics
            '        Dim i As Integer

            '        xtodos(0) = rec
            '        For i = 0 To nRecs
            '            If i < nRecs Then
            '                xtodos(i).dTics = xtodos(i + 1).dTics
            '                'xTodos(i).ApptDateStr = formatApptDateStr(xTodos(i + 1).dTics)

            '                'xTodos(i).ApptDateStr = xTodos(i + 1).ApptDateStr
            '                xtodos(i).ApptDateStr = formatApptDateStr(bottomDics)

            '                ''''UpdateApptTodoRecord(xTodos(i))
            '            Else
            '                xtodos(i).dTics = bottomDics
            '                xtodos(i).ApptDateStr = formatApptDateStr(bottomDics)
            '                ''''UpdateApptTodoRecord(xTodos(i))
            '            End If
            '        Next
            '        'dd = Now.Ticks
            '        'startTime()

            '        Dim doThisOrThat As Boolean = True ' True 'False
            '        'means either DO the NEW or the OLD
            '        If doThisOrThat = True Then
            '            'this is faster
            '            '.111 vrs (almost 3 times faster
            '            '.3
            '            '===============NEW 11/17/2019
            '            ApptGV.CreateDate_LOCATIONarray(0) = ApptGV.All_ApptTodoRecords_Cnt
            '            For i = 0 To nRecs
            '                ApptGV.All_ApptTodoRecords(ApptGV.CreateDate_LOCATIONarray(i)) = xtodos(i)
            '            Next
            '            update_RangeOfApptTodoRecords(0, nRecs, xtodos) 'NEW

            '            QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt)
            '            set_CreateDateArray()

            '            SetDGtodo()
            '            Set_appt_StartDate_Records_DGprinting()
            '            'endTime() '0.0668216
            '            'i = i
            '            '==============
            '        Else


            '            Dim doTesting As Boolean = True
            '            update_RangeOfApptTodoRecords(0, nRecs, xtodos, doTesting) 'NEW
            '            ''''Set_ApptTodoRecordType_UsingReader()
            '            'endTime() '0.2 seconds
            '            '==
            '            'do following after update dtics
            '            'QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt)
            '            'set_CreateDateArray() '0.007 seconds

            '            'the follow is done in UPDATE

            '            If doTesting = False Then 'skip this step is doTesting is set to True
            '                Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
            '            End If

            '            'endTime() '0.2 seconds if doTesting = True else 0.2024454 therefore DO TESTING!
            '            set_CreateDateArray()
            '            'setTodoDistinctDates()

            '            SetDGtodo()
            '            Set_appt_StartDate_Records_DGprinting()

            '            'endTime() '0.0448798 vrs 0.0668216
            '            'i = i
            '        End If
            '        'endTime()
            '    End If

            'End If
            'endTime()
            'Debug.Print("added rec")
            'Debug.Print(Format(Now, "o") & " end")
            'Debug.Print((Now.Ticks - dd) / TimeSpan.TicksPerSecond & " end")
        End If
        'Return F_AddTodo.RetMessage
        '=============
        '=====================

    End Sub
    Sub refreshDataGridView1()
        Dim i As Integer
        Dim msg As String
        DataGridView1.Rows.Clear()

        For i = 1 To ApptGV.All_ApptTodoRecords_Cnt

            'DataGridView1.Rows.Add(i.ToString("0000"), i Mod 2, TextBox1.Text, "Edit", "Delete")
            'With ApptGV.All_ApptTodoRecords(i)
            With ApptGV.Use_ApptTodoRecords(i)
                'DataGridView1.Rows.Add(.RecNumber.ToString, .Completed, .msg, "Edit", "Delete")
                msg = TrimF(.Msg)
                DataGridView1.Rows.Add(i.ToString, .DeleteFlag, msg, "Edit", "Delete")
            End With
        Next

    End Sub
    '======
    '======


    'Private Shared Sub EnsureVisibleRow(ByVal view As DataGridView, ByVal rowToShow As Integer)
    '    If rowToShow >= 0 AndAlso rowToShow < view.RowCount Then
    '        Dim countVisible = view.DisplayedRowCount(False)
    '        Dim firstVisible = view.FirstDisplayedScrollingRowIndex

    '        If rowToShow < firstVisible Then
    '            view.FirstDisplayedScrollingRowIndex = rowToShow
    '        ElseIf rowToShow >= firstVisible + countVisible Then
    '            view.FirstDisplayedScrollingRowIndex = rowToShow - countVisible + 1
    '        End If
    '    End If
    'End Sub





    Private Sub B_ShowReorg_Click(sender As Object, e As EventArgs) Handles B_ShowReorg.Click
        F_Reorg.ShowDialog()
    End Sub

    Private Sub B_showCompleted_Click(sender As Object, e As EventArgs) Handles B_showCompleted.Click
        F_Completed.ShowDialog()
    End Sub
End Class
'