'Imports System.Data.DataTable
Imports System.IO
Public Class F_Main
#Region "PrintingVariables"
    Dim DGV_Printing As DataGridView '= DataGridView1
    Private bmpDirections1 As System.Drawing.Bitmap

    Dim BirthdayRecs() As BirthdayRecordType
    Dim BirthdayRecsDeleted() As BirthdayRecordType


    Dim ApptsSearchNEW() As ApptTodoRecordType

    Dim SelectedYearMonth As String
    Dim SelectedYearMonthLong As Long

    Dim heightDGV As Integer
    Dim SelectedDate As Date

    Dim whatToDisplay() As StringPointerType
    Dim PrintPage1ofTwoPagePrintout As Boolean 'new 4/2/2017 
    Dim canUseDuplexPrinting As Boolean 'new 4/2/2017 'used for printing phone number list

    '===Variables for Printing
    Dim Appt_FirstTimeIn_PrintingCalendar As Boolean = False

    Dim appt_mRow As Integer 'NEW
    Dim appt_newPage As Boolean 'NEW
    Dim appt_ColNum As Integer 'NEW
    Dim appt_PageNum As Integer 'NEW
    Dim appt_LeftMarginAdj As Single 'NEW
    Dim appt_LastDateString As String 'NEW
    Dim appt_Date As String
    Dim todo_Header As String
    Dim DeletedTodos_Header As String


    'Dim appt_TodayDate As String
    Dim appt_TodayDateFlag As Boolean
    '=

    Dim appt_distinctDates() As Date
    Dim appt_distinctDates_for_MonthCalendar1() As Date


    'Dim appt_ApptRecs() As ApptRecAllType 'used for print screen showing appts and (optionally) todos on second appt screen
    Dim appt_ApptRecs() As ApptTodoRecordType 'NEW 'used for print screen showing appts and (optionally) todos on second appt screen

    'Dim appt_TodoDeletedRecs() As ApptRecAllType
    Dim appt_TodoDeletedRecs() As ApptTodoRecordType 'NEW


    Dim appt_startDate As Date
    Dim appt_EndDate As Date
    ' Dim appt_UserID As Integer 'use ApptGV.UserID instead
    Dim appt_Header As String

    'Dim appt_UseTodos As Boolean

    '==Variables for Printing - END

    Dim ApptMode1Day As Boolean = True

    'Dim appts() As ApptRecAllType
    Dim appts() As ApptTodoRecordType 'NEW

    'Dim firstTimeIn As Boolean = True
    ''''Dim skip As Boolean

    'Dim todos() As ApptRecAllType
    Dim todos() As ApptTodoRecordType 'NEW

    Const BallotBoxWithCheck As String = ChrW(&H2611)
    Const BallotBoxWithoutCheck As String = ChrW(&H2610)
    Const NoNumber As String = ChrW(&H2116)
    Const noTimeSymbol As String = ChrW(&H2297)
    Const appt_RightArrow As String = ChrW(&H25B6)

#End Region'PrintingVariables

    'Dim SourceRowIndex As Integer = -1
    'Dim TargetRowIndex As Integer = -1
    'Dim dtSource As DataTable
    'Dim setSelectedRow As Boolean = False

    Dim formOrigWidth As Single
    Dim skip As Boolean = True
    Dim msgNum As Integer = 2
    'Dim currentIndexG As Integer = -1
    Dim Msg2clipboard As String = ""

    'Dim Recs() As tempType 'never changes except when new rec added
    'Dim Recs2Use() As tempType 'holds only what is displayed
    Dim Rec2Edit As ApptTodoRecordType  'tempType
    Private Sub F_Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        B_Utilities.Visible = False
        PanelHide.Visible = False


        Dim toolTip1 As New ToolTip()
        toolTip1.ShowAlways = True
        'toolTip1.SetToolTip(B_AddTodo, "Click me to execute.")
        toolTip1.SetToolTip(B_AddTodo, "or press Ctrl+A")
        'B_AddTodo.tooltip1 = "hello"

        Dim toolTipZ As New ToolTip()
        toolTipZ.ShowAlways = True
        'toolTip1.SetToolTip(B_AddTodo, "Click me to execute.")
        toolTipZ.SetToolTip(B_Copy2Clipboard, "or press Ctrl+Z")


        MonthCalendarForPrinting.Visible = False
        MonthCalendarForPrinting.ShowToday = False 'True 'False
        'note: false looks better

        DGV_Printing = DataGridView1
        If System.IO.Directory.Exists("C:\Donald Stuart Campbell Personal Folder") Then
            PanelBackUpRestore.Visible = True
            'B_Backup.Visible = True
            'B_Restore.Visible = True
        Else
            PanelBackUpRestore.Visible = False
            'B_Backup.Visible = False
            'B_Restore.Visible = False
        End If


        'set file system to use
        setPath2use()
        AddUtilityDirectoryIfDoesNotExist()


        ApptGV.fontSize = get_FontSizeInUse()
        'Debug.Print(ApptGV.fontSize)
        'If System.IO.Directory.Exists("C:\Donald Stuart Campbell Personal Folder") Then
        '    'ApptGV.ccPath = "C:\todoProgramCOM2020ListMgt"
        '    ApptGV.ccPath = "C:\Donald Stuart Campbell Personal Folder\BackupListMgt"
        'End If



        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
        'set_CreateDateArray()
        set_DeleteDateArray() 'should be called CompletedDateArray



        'ApptGV.Use_ApptTodoRecords = ApptGV.All_ApptTodoRecords.Clone

        ''''setRecsToBeDisplayed()
        'refreshDataGridView1()
        '===

        'Dim rec As ApptTodoRecordType
        'Debug.Print("len Of rec =" & Len(rec))



        'Recs = getRecs()
        'Recs2Use = Recs.Clone

        skip = True


        initSetUp()

        CheckBoxIncludeItemsCompletedToday.Checked = False
        'RadioButton1.Checked = True

        setRecsToBeDisplayed()
        refreshDataGridView1()

        ToolStripComboBox1.SelectedIndex = (ApptGV.fontSize - 8)
        ComboBoxFontSize.SelectedIndex = (ApptGV.fontSize - 8)

        skip = False

        setFileSetName_BeingUsed() 'sets: ApptGV.nameOfFileSet_BeingUsed
        Lbl_FileNameBeingUsed.Text = ApptGV.nameOfFileSet_BeingUsed

        'Debug.Print(DataGridView1.Columns(2).DefaultCellStyle.Font.Name)
        'Debug.Print(DataGridView1.Columns(2).DefaultCellStyle.Font.Size)
        'Debug.Print(DataGridView1.Columns(2).DefaultCellStyle.Font.SystemFontName)


        DataGridView1.Columns(0).DefaultCellStyle.Font = New Font("microsoft sans serif", ApptGV.fontSize, FontStyle.Regular)
        DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", ApptGV.fontSize, FontStyle.Regular)
        DataGridView1.Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", ApptGV.fontSize, FontStyle.Regular)


        Msg2clipboard = ""
        setClipBoard()


    End Sub
    Private Sub initSetUp()
        'AddHandler DataGridView1.CurrentCellDirtyStateChanged, AddressOf DataGridView1_CurrentCellDirtyStateChanged
        'AddHandler DataGridView1.CellValueChanged, AddressOf DataGridView1_CellValueChanged

        CheckBoxIncludeItemsCompletedToday.Checked = False
        'RadioButton1.Checked = True

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
            '.Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", 10, FontStyle.Regular)
            .Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", ApptGV.fontSize, FontStyle.Regular)
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
        'DataGridView1.Columns(msgNum).Width = 500 + (Me.Width - formOrigWidth)
        DataGridView1.Columns(msgNum).Width = 540 + (Me.Width - formOrigWidth)



        'Dim s As Integer = 10
        '' If Me.Height > 600 Then s = 10 + (Me.Height - 600) / 200 'as of 1/7/2020 turn off auto changing of font size

        ''Debug.Print(Me.Height & " " & s)
        'DataGridView1.Columns(0).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
        'DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
        'DataGridView1.Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
    End Sub

    Private Sub absDelete(ByVal eRowIndex As Integer)
        Dim LocIn_All_ApptTodoRecords As Integer
        Dim recID As Integer = 0
        Dim absDeleteTicks As Long = 1577664000000000 'new date(6,1,1).ticks
        'ApptGV.todoTicks
        'new date(6,1,1).ticks
        '1577664000000000

        '=============
        Dim recLoc As Integer = eRowIndex + 1
        Dim nRecs As Integer = UBound(ApptGV.Use_ApptTodoRecords)
        If recLoc > 1 Then
            recID = ApptGV.Use_ApptTodoRecords(recLoc - 1).ID
        ElseIf nRecs > 1 Then
            recID = ApptGV.Use_ApptTodoRecords(recLoc + 1).ID
        Else
            recID = 0
        End If
        'Dim d As Date = New Date(5, 1, 1)
        '=============

        'recID = ApptGV.Use_ApptTodoRecords(eRowIndex + 1).ID
        'new
        ApptGV.Use_ApptTodoRecords(eRowIndex + 1).dTicsCompleted = absDeleteTicks 'Now.Ticks

        ApptGV.Use_ApptTodoRecords(eRowIndex + 1).DeleteFlag = 1

        'new
        ApptGV.Use_ApptTodoRecords(eRowIndex + 1).DeleteAbsolute = 1 'new 1/17/2020
        ApptGV.Use_ApptTodoRecords(eRowIndex + 1).dTicsOriginal = Now.Ticks 'new 1/17/2020

        UpdateApptTodoRecord(ApptGV.Use_ApptTodoRecords(eRowIndex + 1))

        'now update record in ApptGV.All_ApptTodoRecords
        'LocIn_All_ApptTodoRecords = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTics Mod 10000000)
        LocIn_All_ApptTodoRecords = biSearch_GTE_ApptTodoRecords(ApptGV.Use_ApptTodoRecords(eRowIndex + 1).dTics)
        ApptGV.All_ApptTodoRecords(LocIn_All_ApptTodoRecords).DeleteFlag = 1

        'new
        ApptGV.All_ApptTodoRecords(LocIn_All_ApptTodoRecords).DeleteAbsolute = 1

        ApptGV.All_ApptTodoRecords(LocIn_All_ApptTodoRecords).dTicsCompleted = ApptGV.Use_ApptTodoRecords(eRowIndex + 1).dTicsCompleted

        set_DeleteDateArray()

        setRecsToBeDisplayed()
        refreshDataGridView1()

        If recID > 0 Then
            Move2Row(recID)
        End If
    End Sub


    'This will fire off after CurrentCellDirtyStateChanged occured...
    'You can get row or column index from e as well here...
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        'Do what you need to do...
        If skip Then Exit Sub
        If e.ColumnIndex <> 1 Then Exit Sub

        Dim LocIn_All_ApptTodoRecords As Integer

        Dim recID As Integer = 0

        If DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = False Then '"True" Then
            'Checked condition'
            'Debug.Print("UnChecked" & " " & ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).Msg)

            'two things
            '1. update record (save)
            '2. find rec in all... and update that rec
            '3. update rec in use...

            recID = ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).ID

            ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTicsCompleted = 0
            ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).DeleteFlag = 0
            UpdateApptTodoRecord(ApptGV.Use_ApptTodoRecords(e.RowIndex + 1))

            'now update record in ApptGV.All_ApptTodoRecords
            'LocIn_All_ApptTodoRecords = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTics Mod 10000000)
            LocIn_All_ApptTodoRecords = biSearch_GTE_ApptTodoRecords(ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTics)
            ApptGV.All_ApptTodoRecords(LocIn_All_ApptTodoRecords).DeleteFlag = 0
            ApptGV.All_ApptTodoRecords(LocIn_All_ApptTodoRecords).dTicsCompleted = 0

            set_DeleteDateArray()


            setRecsToBeDisplayed()
            refreshDataGridView1()

            Move2Row(recID)
        Else
            'Unchecked Condition'
            'Debug.Print("Checked" & " " & ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).Msg)

            recID = ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).ID

            ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTicsCompleted = Now.Ticks
            ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).DeleteFlag = 1
            UpdateApptTodoRecord(ApptGV.Use_ApptTodoRecords(e.RowIndex + 1))

            'now update record in ApptGV.All_ApptTodoRecords
            'LocIn_All_ApptTodoRecords = biSearch_GTE_ApptTodoRecords(ApptGV.todoTicks + ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTics Mod 10000000)
            LocIn_All_ApptTodoRecords = biSearch_GTE_ApptTodoRecords(ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTics)
            ApptGV.All_ApptTodoRecords(LocIn_All_ApptTodoRecords).DeleteFlag = 1
            ApptGV.All_ApptTodoRecords(LocIn_All_ApptTodoRecords).dTicsCompleted = ApptGV.Use_ApptTodoRecords(e.RowIndex + 1).dTicsCompleted

            set_DeleteDateArray()

            setRecsToBeDisplayed()
            refreshDataGridView1()

            Move2Row(recID)
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
        'Debug.Print(e.RowIndex & "  i am here")
        Dim xLoc As Integer = 0
        Select Case e.ColumnIndex
            Case 3 'edit
                ' Dim tempIndex As Integer = currentIndexG

                Rec2Edit = ApptGV.Use_ApptTodoRecords(e.RowIndex + 1)
                xLoc = biSearch_GTE_ApptTodoRecords(Rec2Edit.dTics)
                'If Trim(ApptGV.All_ApptTodoRecords(xLoc).Msg) <> Rec2Edit.Msg Then
                '    If ApptGV.All_ApptTodoRecords(xLoc).dTics <> Rec2Edit.dTics Then
                '        xLoc = xLoc
                '    End If
                '    Debug.Print(ApptGV.All_ApptTodoRecords(xLoc).Msg & " " & Len(ApptGV.All_ApptTodoRecords(xLoc).Msg))
                '    Debug.Print(Rec2Edit.Msg & " " & Len(Rec2Edit.Msg))
                '    xLoc = xLoc
                'End If
                Dim Msg As String = TrimF(Rec2Edit.Msg)

                F_AddTodo.PassCanRollover = False

                F_AddTodo.PassMessage = Msg
                F_AddTodo.AddingNewToDo = False
                F_AddTodo.ShowDialog()
                '============
                'Debug.Print(F_AddTodo.RetMessage)
                Msg = TrimF(F_AddTodo.RetMessage)
                ' Msg = Msg
                'Return F_AddTodo.RetMessage
                '==
                ''''Msg = ModifyTodoMsg(Trim(existingTodoRecord.Msg))
                'Debug.Print(msg)
                'If Msg <> "" AndAlso Msg <> Trim(ApptGV.All_ApptTodoRecords(xLoc).Msg) Then
                If Msg <> "" Then
                    ApptGV.All_ApptTodoRecords(xLoc).Msg = Msg
                    UpdateApptTodoRecord(ApptGV.All_ApptTodoRecords(xLoc))

                    setRecsToBeDisplayed()
                    refreshDataGridView1()
                    'SetDGtodo()
                    'Set_appt_StartDate_Records_DGprinting()

                End If
                'DataGridView1.Rows(e.RowIndex).Selected = True
                Move2Row(Rec2Edit.ID)
                '==
                '=============
                'Debug.Print(currentIndexG & " G   T " & tempIndex)
                'currentIndexG = tempIndex
            Case 4 'delete
                'Dim xmsg As String = "Are you sure you want to Permanently" & vbNewLine & "Delete this Todo Record?"
                'Dim xmsg As String = "Confirm you want to Permanently" & vbNewLine & "Delete this Todo Record?"
                Dim xmsg As String = "Press OK to Permanently Delete" & vbNewLine & vbNewLine & "o/w Press Cancel"
                Dim xTitle As String = "Please Confirm"
                'Dim result As DialogResult = MessageBox.Show(xmsg, xTitle, MessageBoxButtons.YesNoCancel)
                'Dim result As DialogResult = MessageBox.Show(xmsg, xTitle, MessageBoxButtons.YesNo)
                'Dim result As DialogResult = MessageBox.Show(xmsg, xTitle, MessageBoxButtons.OKCancel)
                'MessageBoxButtons.OKCancel,MessageBoxIcon.Hand,MessageBoxDefaultButton.Button2)

                'Dim result As DialogResult = MessageBox.Show(xmsg, xTitle, MessageBoxButtons.OKCancel, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button2)
                'Dim result As DialogResult = MessageBox.Show(xmsg, xTitle, MessageBoxButtons.OKCancel + MessageBoxDefaultButton.Button2)

                Dim result As DialogResult = MsgBox(xmsg, MsgBoxStyle.OkCancel + MsgBoxStyle.DefaultButton2, xTitle)
                'Dim result As DialogResult = MsgBox(xmsg, MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, xTitle)

                If result = DialogResult.OK Then 'DialogResult.Yes Then
                    'Debug.Print("delete " & e.RowIndex & " " & Now)
                    absDelete(e.RowIndex)
                End If
                'If MessageBox.Show(xmsg, xTitle, MessageBoxButtons.YesNo) = vbYes Then
                '    ' Other Code goes here
                '    Debug.Print("delete " & e.RowIndex)
                'End If
                'Debug.Print("delete " & e.RowIndex)
                'absDelete(e.RowIndex)
        End Select
    End Sub

    'Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
    '    If skip Then Exit Sub
    '    Dim x As Integer = currentrow() - 1
    '    'Debug.Print(x)
    '    'Recs2Use = Recs.Clone
    '    'test_sampleItems()
    '    setRecsToBeDisplayed()
    '    refreshDataGridView1()

    '    If x >= 0 AndAlso x <= DataGridView1.Rows.Count - 1 Then
    '        DataGridView1.Rows(x).Selected = True
    '    End If

    '    'Try
    '    '    DataGridView1.Rows(x).Selected = True
    '    'Catch ex As Exception
    '    '    Debug.Print(x)
    '    'End Try

    'End Sub
    'Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
    '    If skip Then Exit Sub
    '    Dim x As Integer = currentrow() - 1

    '    If x >= 0 AndAlso x <= DataGridView1.Rows.Count - 1 Then
    '        DataGridView1.Rows(x).Selected = True
    '    End If

    'End Sub
    Private Function currentrow() As Integer
        Dim x As Integer = 0
        If DataGridView1.Rows.Count = 0 Then
            'If IsNothing(Me.DataGridView1.CurrentRow.Cells(0).Value) OrElse Not IsNumeric((Me.DataGridView1.CurrentRow.Cells(0).Value)) Then
            x = 0
        Else
            x = Me.DataGridView1.CurrentRow.Cells(0).Value
        End If

        'DataGridView1.CurrentRow '.SelectedRow.RowIndex
        Return x

    End Function
    'Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
    '    If skip Then Exit Sub
    '    setRecsToBeDisplayed()
    '    refreshDataGridView1()

    '    'If RadioButton2.Checked = True Then
    '    '    'Dim nRecs As Integer = UBound(Recs)
    '    '    'Dim i As Integer
    '    '    'ReDim Recs2Use(nRecs)
    '    '    'Dim cnt As Integer = 0
    '    '    'For i = 1 To nRecs
    '    '    '    With Recs(i)
    '    '    '        If .Completed = 0 Then
    '    '    '            cnt += 1
    '    '    '            Recs2Use(cnt) = Recs(i)
    '    '    '        End If
    '    '    '    End With
    '    '    'Next
    '    '    'ReDim Preserve Recs2Use(cnt)
    '    '    'test_sampleItems()
    '    'End If


    'End Sub



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
        AddTodoRecord()
    End Sub
    Private Sub AddTodoRecord()

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
            'Debug.Print("put On BOTTOM")
            'Exit Sub

            rec.Msg = TrimF(F_AddTodo.RetMessage)
            rec.dTics = ApptGV.todoTicks
            AddApptTodoRecord(rec) 'added to ApptGV.All_ApptTodoRecords
            'rec added - now do the following:
            '1. refresh dataGridView1

            setRecsToBeDisplayed()
            'ApptGV.Use_ApptTodoRecords = ApptGV.All_ApptTodoRecords.Clone

            refreshDataGridView1()

            Move2Row(rec.ID)

            'setRecsToBeDisplayed()
            'refreshDataGridView1()

            '=============================
            ''''Exit Sub

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
            'Debug.Print(Format(Now, "o") & " End")
            'Debug.Print((Now.Ticks - dd) / TimeSpan.TicksPerSecond & " End")
        Else
            'put on Top
            '==========
            'new 1/5/2020
            Dim a() As ApptTodoRecordType = getDeletedTodos_withinDateRange_NEW(New Date(0))
            QuickSort_ApptTodoRecords(a, 1, UBound(a))

            Dim nRecs As Integer = UBound(a)
            Dim b(nRecs) As Integer
            Dim i As Integer
            For i = 1 To nRecs
                b(i) = biSearch_GTE_ApptTodoRecords(a(i).dTics)
                If ApptGV.All_ApptTodoRecords(b(i)).ID <> a(i).ID Then
                    'error
                    'i = i
                    QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt)
                    set_DeleteDateArray()

                    setRecsToBeDisplayed()
                    refreshDataGridView1()
                    Exit Sub

                End If
            Next

            '==================================================
            rec.Msg = TrimF(F_AddTodo.RetMessage)
            rec.dTics = ApptGV.todoTicks
            AddApptTodoRecord(rec, True) 'True skips QuckSort... and set_DeledDateArray'added to ApptGV.All_ApptTodoRecords
            '==================================================
            'rec added - now do the following:
            '1.
            Dim tempL As Long = ApptGV.All_ApptTodoRecords(ApptGV.All_ApptTodoRecords_Cnt).dTics  'get first items
            a(0) = ApptGV.All_ApptTodoRecords(ApptGV.All_ApptTodoRecords_Cnt)
            For i = 1 To nRecs
                a(i - 1).dTics = a(i).dTics
                a(i - 1).ApptDateStr = formatApptTodoDateStr(a(i - 1).dTics) 'yyyy/MM/dd HH:mm:ss.fffffff
            Next
            a(nRecs).dTics = tempL
            a(nRecs).ApptDateStr = formatApptTodoDateStr(tempL)
            For i = 1 To nRecs
                ApptGV.All_ApptTodoRecords(b(i)) = a(i)
            Next
            ApptGV.All_ApptTodoRecords(ApptGV.All_ApptTodoRecords_Cnt) = a(0)
            '2. now save to disk
            SaveTotoRecords2Disk(a)
            '========================
            '1. refresh dataGridView1

            'Move2Row(rec.ID)


            QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt)
            set_DeleteDateArray()

            setRecsToBeDisplayed()
            refreshDataGridView1()

            Move2Row(rec.ID)

        End If
        'Return F_AddTodo.RetMessage
        '=============
        '=====================

    End Sub
    Sub refreshDataGridView1()
        Dim i As Integer
        Dim msg As String
        DataGridView1.Rows.Clear()

        Dim nRows As Integer = UBound(ApptGV.Use_ApptTodoRecords)
        For i = 1 To nRows 'ApptGV.All_ApptTodoRecords_Cnt

            'DataGridView1.Rows.Add(i.ToString("0000"), i Mod 2, TextBox1.Text, "Edit", "Delete")
            'With ApptGV.All_ApptTodoRecords(i)
            With ApptGV.Use_ApptTodoRecords(i)
                'DataGridView1.Rows.Add(.RecNumber.ToString, .Completed, .msg, "Edit", "Delete")
                'msg = .dTics Mod 10000000 & " " & TrimF(.Msg)
                'msg = .dTics Mod 10000000 & " " & TrimF(.Msg)
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
        ReorderList()
    End Sub
    Private Sub ReorderList()

        F_Reorg.ShowDialog()

        setRecsToBeDisplayed()
        refreshDataGridView1()
    End Sub

    Private Sub B_showCompleted_Click(sender As Object, e As EventArgs) Handles B_showCompleted.Click
        F_Completed.ShowDialog()

        setRecsToBeDisplayed()
        refreshDataGridView1()
    End Sub
    Private Sub setRecsToBeDisplayed()
        ApptGV.Use_ApptTodoRecords = getDeletedTodos_withinDateRange_NEW(New Date(0)) '(New Date(ApptGV.todoTicks)) '(selDate, True)
        QuickSort_ApptTodoRecords(ApptGV.Use_ApptTodoRecords, 1, UBound(ApptGV.Use_ApptTodoRecords))
        'QuickSort_ApptTodoRecords_desc(ApptGV.Use_ApptTodoRecords, 1, UBound(ApptGV.Use_ApptTodoRecords))

        'If RadioButton1.Checked = False Then
        If CheckBoxIncludeItemsCompletedToday.Checked = True Then

            Dim delToday() As ApptTodoRecordType = getDeletedTodos_DeletedToday()
            Dim nRecs As Integer = UBound(delToday)

            Dim i As Integer
            Dim cnt = UBound(ApptGV.Use_ApptTodoRecords)
            ReDim Preserve ApptGV.Use_ApptTodoRecords(cnt + nRecs)
            For i = 1 To UBound(delToday)
                cnt += 1
                'ReDim Preserve ApptGV.Use_ApptTodoRecords(cnt)
                ApptGV.Use_ApptTodoRecords(cnt) = delToday(i)
                'Debug.Print(Trim(delToday(i).Msg))
            Next
        End If
    End Sub

    Private Sub B_Utilities_Click(sender As Object, e As EventArgs) Handles B_Utilities.Click
        ' Debug.Print(ApptGV.ccPath)
        F_Utilities.ShowDialog()
        Restart()

        setFileSetName_BeingUsed() 'sets: ApptGV.nameOfFileSet_BeingUsed
        'Lbl_FileNameBeingUsed.Text = "Todo List Name: " & ApptGV.nameOfFileSet_BeingUsed
        Lbl_FileNameBeingUsed.Text = ApptGV.nameOfFileSet_BeingUsed

    End Sub

    Private Sub B_Backup_Click(sender As Object, e As EventArgs) Handles B_Backup.Click

        BackUp_WorkingFiles()
        MsgBox("Backup Done!")
        End
    End Sub

    Private Sub B_Restore_Click(sender As Object, e As EventArgs) Handles B_Restore.Click
        Restore_WorkingFiles()
        MsgBox("Restore Done!")
        End
    End Sub
    Private Sub Restart()
        'setPath2use()
        'AddUtilityDirectoryIfDoesNotExist()


        Set_ApptTodoRecordType_UsingReader() 'sets: 1. array and 2. ApptGV.All_ApptTodoRecords_Cnt '0.2 seconds 3. does Sort based on dTics '.2323784 seconds
        'set_CreateDateArray()
        set_DeleteDateArray() 'should be called CompletedDateArray

        skip = True


        'initSetUp()

        'RadioButton1.Checked = True
        CheckBoxIncludeItemsCompletedToday.Checked = False

        setRecsToBeDisplayed()
        refreshDataGridView1()

        skip = False



    End Sub

    Private Sub B_EraseAllFiles_Click(sender As Object, e As EventArgs) Handles B_EraseAllFiles.Click
        EraseAllFiles()
    End Sub
#Region "Printing"


    Private Sub B_Print_Click(sender As Object, e As EventArgs) Handles B_Print.Click
        PrintTodoList()
    End Sub
    Private Sub PrintTodoList()

        Appt_FirstTimeIn_PrintingCalendar = True

        'DGV_Printing.Columns(msgNum).DefaultCellStyle.Padding = New Padding(5, 5, 5, 5)
        DGV_Printing.Columns(msgNum).DefaultCellStyle.Padding = New Padding(0, 0, 0, 0)

        'Dim DGV_Printing As DataGridView = DataGridView1
        appt_ApptRecs = ApptGV.Use_ApptTodoRecords.Clone

        'SurroundingSub()
#Region "Print Appt - Start of Routine"
        '===========================appt printing
        Dim existingFont As Font = DGV_Printing.Font

        'DGV_Printing.Font = New Font("Microsoft SansSerif", 8, FontStyle.Regular)


        DGV_Printing.Font = New Font("Microsoft SansSerif", ApptGV.fontSize, FontStyle.Regular) 'new 1/7/2020
        'DGV_Printing.Font = New Font("Segoe Print", ApptGV.fontSize, FontStyle.Regular) 'new 1/7/2020


        'DGV_Printing.Width = PanelApptPrinting.Width - 20
        ''''SetUpDataGridView_printing()
        ''''SetDG_printing()

        ''''PanelApptPrinting.Width = DGV_Printing.Width + 30
#End Region


        'MonthCalendarForPrinting.Visible = True
        ''''SetStart_End_DatesETC()
        'above also sets the following:
        'appt_distinctDates
        'appt_ApptRecs

        '--

        ' PrintDocument1.DefaultPageSettings.Landscape = True
        Dim ret As DialogResult


        'SelectAllDates()



        'set the following:
        appt_mRow = 0
        appt_newPage = True
        appt_ColNum = 1
        appt_PageNum = 1
        appt_LeftMarginAdj = 0
        appt_LastDateString = ""
        appt_Date = "Printed: " & Format(Now, "dddd, MMMM d, yyyy h:mm tt")
        'appt_Date = "Printed: " & Format(New Date(2019, 9, 25), "dddd, MMMM d, yyyy h:mm tt")

        '''' MonthCalendarForPrinting.ShowToday = False 'True 'False
        'note: false looks better

        'moved to where regular calendar is updated with distinctDates
        'updateBoldedDates_MonthCalendarForPrinting_Appt(MonthCalendarForPrinting, appt_distinctDates)


        ''''updateBoldedDates_MonthCalendarForPrinting_Appt(MonthCalendarForPrinting, appt_distinctDates_for_MonthCalendar1)


        'MonthCalendarForPrinting.BackColor = Color.Azure


        ''''Set3monthcalendar_whatToShow()
        ' MonthCalendarForPrinting.SelectionStart = DateTimePickerFrom.Value
        'Debug.Print(MonthCalendarForPrinting.SelectionStart & " " & MonthCalendarForPrinting.SelectionEnd)
        'Debug.Print(MonthCalendarForPrinting.SelectionRange.ToString)

        appt_TodayDateFlag = False

        '==
        ''''Set_appt_StartDate_Records_DGprinting()
        'SetStart_End_DatesETC() 'testing
        'SetDG_printing() 'put here for testing - uses appt_ApptRecs()
        'Exit Sub

        Try
            ret = PrintDialog1.ShowDialog()
        Catch ex As Exception
            ' Debug.Print("after printDialog1.showDialog" & vbNewLine & ex.Message)
        End Try

        'If ret = Windows.Forms.DialogResult.OK Then
        If ret = DialogResult.OK Then

            PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings

            PrintDocument1.OriginAtMargins = False 'true = soft margins, false = hard margins
            PrintDocument1.DefaultPageSettings.Landscape = False 'True ' False'!!!!
            PrintDocument1.OriginAtMargins = False
            '==================
            'get selected printer name
            Dim printerName As String = PrintDialog1.PrinterSettings.PrinterName
            If InStr(printerName, "XPS") OrElse InStr(printerName, "PDF") Then
                'leftMarginAdj_XPS_PDF_type_width = 25
                PrintDocument1.DefaultPageSettings.Margins.Left = 26.5 + 25 ' 29 ' 31.5 '50 '0
            Else
                'leftMarginAdj_XPS_PDF_type_width = 0
                PrintDocument1.DefaultPageSettings.Margins.Left = 26.5 ' 29 ' 31.5 '50 '0
            End If
            '=========

            '==================
            'Debug.Print(MonthCalendarForPrinting.Height)
            PrintDocument1.DefaultPageSettings.Margins.Top = MonthCalendarForPrinting.Height + 95  '100 '70 '70 ' 0'90 95! note: 88 perfectly fits
            PrintDocument1.DefaultPageSettings.Margins.Right = 0 '50 '0
            PrintDocument1.DefaultPageSettings.Margins.Bottom = 70 '105 ' 70 ' 0


            Try
                If PrintDocument1.PrinterSettings.CanDuplex Then
                    PrintDocument1.PrinterSettings.Duplex = Drawing.Printing.Duplex.Vertical
                    canUseDuplexPrinting = True
                Else
                    canUseDuplexPrinting = False
                End If



            Catch ex As Exception
                canUseDuplexPrinting = False
            End Try


            MonthCalendarForPrinting.Visible = True
            PrintPage1ofTwoPagePrintout = True 'canUseDuplexPrinting
            Try
                PrintDocument1.Print()
            Catch ex As Exception
                'Debug.Print(ex.Message)

                'this.Close()

                MonthCalendarForPrinting.Visible = False
                MsgBox("The process cannot access the file because it is being used by another process" & vbNewLine & "Printout not successful - close File named and try again!")

                Exit Sub
            End Try

        End If
        MonthCalendarForPrinting.Visible = False
#Region "Print Appt - End of Routine"
        '===========================appt printing
        DGV_Printing.Font = New Font("Microsoft SansSerif", existingFont.Size, FontStyle.Regular)
        'DGV_Printing.Width = PanelApptPrinting.Width - 20
        ''''SetUpDataGridView_printing()
        ''''SetDG_printing()

        ''''PanelApptPrinting.Width = DGV_Printing.Width + 30

#End Region

        'MonthCalendarForPrinting.Visible = True

        DGV_Printing.Columns(msgNum).DefaultCellStyle.Padding = New Padding(5, 5, 5, 5)
    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'Dim font_ballotBox = New Font("Microsoft Sans Serif", 10, FontStyle.Regular)
        Dim font_ballotBox = New Font("Microsoft Sans Serif", ApptGV.fontSize, FontStyle.Regular)

        '9/30/2018 changed all references from appts to appt_apptRecs

        'Dim mRow As Integer
        'Static mRow As Integer = 0

        'Dim newpage As Boolean


        'Dim DateToday As Date = Date.Today
        Dim DateToday As Long = Date.Today.Ticks 'changed in ApptTodo

        'NEW 12/20/2019
        'The following solves the issue of Not printing the Calendar after the first page
        Dim y As Single '= e.MarginBounds.Top '257'
        If Appt_FirstTimeIn_PrintingCalendar = True Then
            PrintCalendarOnApptPage(e)
            Appt_FirstTimeIn_PrintingCalendar = False
            y = e.MarginBounds.Top '257'
        Else
            y = e.MarginBounds.Top - 190
        End If

        'Exit Sub
        'Dim msg As String
        'Dim backColor1 As Color = Color.Azure
        'Dim backColor2 As Color = Color.LightGray
        'Dim myBrush1 As System.Drawing.SolidBrush = New System.Drawing.SolidBrush(System.Drawing.Color.Red)
        Dim myBrush1 As System.Drawing.SolidBrush = New System.Drawing.SolidBrush(System.Drawing.Color.Azure)
        Dim myBrush2 As System.Drawing.SolidBrush = New System.Drawing.SolidBrush(System.Drawing.Color.Gainsboro) '.WhiteSmoke) 'LightGreen) ' .MintCream) '.LightGray)

        Dim myBrushTODO1 As System.Drawing.SolidBrush = New System.Drawing.SolidBrush(System.Drawing.Color.LightSalmon) '.LightYellow)
        Dim myBrushTODO2 As System.Drawing.SolidBrush = New System.Drawing.SolidBrush(System.Drawing.Color.SeaShell) '.LightGoldenrodYellow) '.Linen) 'MintCream)
        Dim myBrushTODO As System.Drawing.SolidBrush = myBrushTODO2 'want darker color to be the 1st TODO to be displaced for contrast purposes so that the eye it drawn to it

        Dim myBrush As System.Drawing.SolidBrush = myBrush1

        Dim totalWidth As Integer = 750 'NEW 12/20/2019 'solve problem of width of each line printed
        Dim UsedWidth As Integer = 0

        'With DGV_appt 'DataGridView3
        With DGV_Printing
            Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
            Dim text As String

            'new 5/4/2019
            '=========================================
            Dim firstColFormat As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
            With firstColFormat
                ''''.Alignment = StringAlignment.Far
                .Alignment = StringAlignment.Center
                .LineAlignment = StringAlignment.Center
            End With
            Dim SecondColFormat As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
            With SecondColFormat
                .Alignment = StringAlignment.Center
                .LineAlignment = StringAlignment.Center
            End With
            '==========================================

            ''''text = "Appointments - " & appt_Header
            'text = "Todos"


            text = "Todo List: " & Lbl_FileNameBeingUsed.Text ' & " - as of " & asOf


            fmt.LineAlignment = StringAlignment.Center
            fmt.Trimming = StringTrimming.EllipsisCharacter
            'e.Graphics.DrawString(text, New Font("Verdana", 10, FontStyle.Bold), Brushes.Black, 350, 30)
            e.Graphics.DrawString(text, New Font("Verdana", 10, FontStyle.Bold), Brushes.Black, e.MarginBounds.Left, 30) ' 35, 30)

            Dim noTimeSymbolFont As New Font("Microsoft SansSerif", 15, FontStyle.Regular)

            Dim unused As New SizeF()
            Dim f As Font = New Font("Verdana", 8, FontStyle.Regular)

            'print date
            Dim asOf As String = "As of " & Format(Now, "MMMM d, yyyy")
            e.Graphics.DrawString(asOf, f, Brushes.Black, e.MarginBounds.Left, 45) '20)

            Dim w As SizeF = e.Graphics.MeasureString(appt_Date, f, 300)
            'e.Graphics.DrawString(appt_Date, f, Brushes.Black, 789 - w.Width, 55) '30)

            e.Graphics.DrawString(appt_Date, f, Brushes.Black, (e.MarginBounds.Left + 749) - w.Width, 1035) '30)

            'new 12/30/2019 deleted printing of "The first Date..." and "No Time Appointment" - next 4 lines -NOT NEEDED!
            'e.Graphics.DrawString(appt_RightArrow & " = The first Date at or after the print date.", f, Brushes.Black, e.MarginBounds.Left, 1035) ' 35, 1035) '30)
            ''5/5/2019
            'e.Graphics.DrawString(noTimeSymbol, noTimeSymbolFont, Brushes.DarkBlue, e.MarginBounds.Left + 273, 1030)
            'e.Graphics.DrawString(" = No Time Appointment", f, Brushes.Black, e.MarginBounds.Left + 283, 1035)

            '==

            text = "Page Number: " & appt_PageNum.ToString
            w = e.Graphics.MeasureString(text, f, 300)
            ' e.Graphics.DrawString(text, f, Brushes.Black, 789 - w.Width, 30) ' 45)
            '789-35 =  754 so (e.MarginBounds.Left+754)
            e.Graphics.DrawString(text, f, Brushes.Black, (e.MarginBounds.Left + 754) - w.Width, 30) ' 45)


            'Dim y As Single = e.MarginBounds.Top '257'NEW moved to firsttimeIn_Calendar 12/30/2019
            'Debug.Print(y)
            'outside loop

            Do While appt_mRow < DGV_Printing.RowCount ' DGV_appt.RowCount

                Dim row As DataGridViewRow = .Rows(appt_mRow)
                If row.Height + y <= e.MarginBounds.Bottom Then 'new 12/30/2019

                    'Debug.Print(appt_mRow & " " & row.Height & " y=" & y)
                    Dim x As Single = e.MarginBounds.Left + appt_LeftMarginAdj
                    Dim h As Single = 0
                    Dim cnt As Integer = 0
                    UsedWidth = 0
                    For Each cell As DataGridViewCell In row.Cells
                        cnt += 1
                        If cnt > 3 Then Exit For

                        'NEW 12/20/2019
                        Dim cellSizeWidth As Integer
                        'Dim cellSizeHeight As Integer
                        If cnt < 3 Then
                            UsedWidth += cell.Size.Width
                            cellSizeWidth = cell.Size.Width
                        Else
                            cellSizeWidth = totalWidth - UsedWidth
                        End If
                        'Dim rc As RectangleF = New RectangleF(x, y, cell.Size.Width, cell.Size.Height)
                        Dim rc As RectangleF = New RectangleF(x, y, cellSizeWidth, cell.Size.Height)


                        'If (appt_newPage) Then
                        '    e.Graphics.DrawRectangle(Pens.Black, rc.Left, rc.Top, rc.Width, rc.Height)
                        '    e.Graphics.DrawString(DGV_appt.Columns(cell.ColumnIndex).HeaderText, .Font, Brushes.Black, rc, fmt)
                        'End If

                        'Else
                        If cnt = 1 Then
                            If DGV_Printing.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString <> appt_LastDateString Then
                                appt_LastDateString = DGV_Printing.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString

                                'If DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString <> appt_LastDateString Then
                                '    appt_LastDateString = DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString
                                If InStr(DGV_Printing.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString, NoNumber) Then
                                    If myBrushTODO.Color = myBrushTODO1.Color Then
                                        myBrushTODO = myBrushTODO2
                                    Else
                                        myBrushTODO = myBrushTODO1
                                    End If
                                    myBrush = myBrushTODO
                                Else
                                    If myBrush.Color = myBrush1.Color Then
                                        myBrush = myBrush2
                                    Else
                                        myBrush = myBrush1
                                    End If

                                End If
                                'If myBrush.Color = myBrush1.Color Then
                                '    myBrush = myBrush2
                                'Else
                                '    myBrush = myBrush1
                                'End If
                            End If
                        End If
                        e.Graphics.FillRectangle(myBrush, New Rectangle(rc.Left, rc.Top, rc.Width, rc.Height))
                        e.Graphics.DrawRectangle(Pens.Black, rc.Left, rc.Top, rc.Width, rc.Height)

                        '==IF At or after Today
                        'NEW 12/30/2019 remmed out the following
                        'If appt_TodayDateFlag = False Then
                        '    ' Debug.Print(DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString & " " & appt_TodayDate)
                        '    ' If DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString >= appt_TodayDate Then

                        '    'Debug.Print(DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString & " " & appt_apptRecs(appt_mRow + 1).apptDate)
                        '    If appt_ApptRecs(appt_mRow + 1).apptDate >= DateToday Then
                        '        appt_TodayDateFlag = True
                        '        e.Graphics.DrawString(appt_RightArrow, .Font, Brushes.Black, rc.Left - 10, rc.Top)
                        '    End If


                        'End If

                        'If cnt = 2 Then
                        '    If InStr(DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), "PM") Then 'If Mid(appt_apptRecs(appt_mRow).apptDateStr, 12, 2) >= "12" Then
                        '        e.Graphics.DrawString(DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.DarkRed, rc, fmt)
                        '    Else
                        '        e.Graphics.DrawString(DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.DarkBlue, rc, fmt)
                        '    End If
                        'Else
                        '    e.Graphics.DrawString(DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, fmt)
                        'End If

                        'msg = DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString()

                        '=========================================
                        'If cnt = 2 Then
                        '    If InStr(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), "PM") Then 'If Mid(appt_apptRecs(appt_mRow).apptDateStr, 12, 2) >= "12" Then
                        '        e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.DarkRed, rc, fmt)
                        '    Else
                        '        e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.DarkBlue, rc, fmt)
                        '    End If
                        'Else
                        '    e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, fmt)
                        'End If

                        'new 5/4/2019 -replace above with below
                        '==========================================
                        If cnt = 2 Then
                            'BallotBoxWithCheck
                            'e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), font_ballotBox, Brushes.Black, rc, fmt)

                            'If ApptGV.Use_ApptTodoRecords(appt_mRow).DeleteFlag = 1 Then

                            If appt_ApptRecs(appt_mRow + 1).DeleteFlag = 1 Then
                                e.Graphics.DrawString(BallotBoxWithCheck, font_ballotBox, Brushes.Black, rc, SecondColFormat) 'fmt)
                            Else
                                e.Graphics.DrawString(BallotBoxWithoutCheck, font_ballotBox, Brushes.Black, rc, SecondColFormat) 'fmt)
                            End If




                            'Dim CellCheckBox = New CheckBox
                            '''''CellCheckBox.Size = New Size(14, 14)
                            'CellCheckBox.Size = New Size(cell.Size.Width, cell.Size.Height)

                            ''CellCheckBox.Checked = CType(Cel.Value, Boolean)
                            ''Dim bmp As New Bitmap(ColumnWidths(i), CellHeight)

                            'CellCheckBox.Checked = CType(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).Value, Boolean)

                            ''Dim s As SizeF = New Point(.Columns(1).Width, .Columns(1).Width)
                            ''rc.Left, rc.Top, rc.Width, rc.Height
                            ''cell.Size.Width, cell.Size.Height
                            'Dim bmp As New Bitmap(cell.Size.Width, cell.Size.Height)

                            'Dim tmpGraphics As Graphics = Graphics.FromImage(bmp)
                            'tmpGraphics.FillRectangle(Brushes.White, New Rectangle(0, 0, bmp.Width, bmp.Height))
                            'CellCheckBox.DrawToBitmap(bmp, New Rectangle(CType((bmp.Width - CellCheckBox.Width) / 2, Int32), CType((bmp.Height - CellCheckBox.Height) / 2, Int32), CellCheckBox.Width, CellCheckBox.Height))
                            '''''e.Graphics.DrawImage(bmp, New Point(ColumnLefts(i), tmpTop))
                            'e.Graphics.DrawImage(bmp, New Point(CInt(x), CInt(y)))
                            ''''e.Graphics.DrawImage(bmp, New Point(CInt(x) + cell.Size.Width / 2, CInt(y)))

                            '===========================================





                            'e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), noTimeSymbolFont, Brushes.DarkBlue, rc, SecondColFormat) ' 12/29/2019
                            ''''e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.DarkBlue, rc, SecondColFormat) ' 12/29/2019

                            'If InStr(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), "PM") Then 'If Mid(appt_TodoDeletedRecs(appt_mRow).apptDateStr, 12, 2) >= "12" Then
                            '    'e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.DarkRed, rc, fmt)
                            '    e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.DarkRed, rc, SecondColFormat)
                            'Else
                            '    If InStr(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), noTimeSymbol) Then
                            '        e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), noTimeSymbolFont, Brushes.DarkBlue, rc, SecondColFormat) ' 5/5/2019
                            '    Else
                            '        e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.DarkBlue, rc, SecondColFormat) 'fmt)
                            '    End If

                            'End If
                        ElseIf cnt = 1 Then

                            e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, firstColFormat)
                        Else

                            e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, fmt)
                        End If
                        '===========================================


                        'msg = .Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString()

                        'End If

                        x += rc.Width
                        h = Math.Max(h, rc.Height)
                    Next
                    appt_newPage = False
                    y += h
                    appt_mRow += 1
                    'Debug.Print("appt_mRow=" & appt_mRow)
                    'Debug.Print("Y=" & y)
                    'Debug.Print("h=" & h)
                    'Debug.Print("e.MarginBounds.Bottom=" & e.MarginBounds.Bottom)
                    'Debug.Print("msg=" & msg)

                    'OLD
                    'If y + h > e.MarginBounds.Bottom Then
                    '    e.HasMorePages = True
                    '    appt_mRow -= 0
                    '    appt_newPage = True
                    '    Exit Sub
                    'End If

                    'If y + h > e.MarginBounds.Bottom Then
                    '    If appt_ColNum = 2 Then 'new actual new page is required
                    '        appt_LeftMarginAdj = 0
                    '        PrintDocument1.DefaultPageSettings.Margins.Left = 35 '50 '0
                    '        appt_ColNum = 1

                    '        'if appt_mRow < DGV_appt.RowCount
                    '        'Debug.Print(appt_mRow & " " & DGV_appt.RowCount)
                    '        If appt_mRow >= .RowCount Then ' DGV_appt.RowCount Then
                    '            e.HasMorePages = False
                    '        Else
                    '            e.HasMorePages = True
                    '        End If

                    '        appt_mRow -= 0
                    '        appt_newPage = True
                    '        appt_PageNum += 1
                    '        Exit Sub
                    '    Else
                    '        appt_LeftMarginAdj = 390 '425
                    '        y = e.MarginBounds.Top 'reset TOP
                    '        appt_ColNum = 2 'now print second column
                    '        'PrintDocument1.DefaultPageSettings.Margins.Left = 460 '(425+35) 35 '50 '0
                    '        appt_mRow -= 0
                    '        appt_newPage = True
                    '    End If
                    'End If

                    'new 12/30/2019 move to ELSE below
                    'If y + h > e.MarginBounds.Bottom Then
                    '    appt_LeftMarginAdj = 0
                    '    PrintDocument1.DefaultPageSettings.Margins.Left = 35 '50 '0
                    '    appt_ColNum = 1

                    '    'if appt_mRow < DGV_appt.RowCount
                    '    'Debug.Print(appt_mRow & " " & DGV_appt.RowCount)
                    '    If appt_mRow >= .RowCount Then ' DGV_appt.RowCount Then
                    '        e.HasMorePages = False
                    '    Else
                    '        e.HasMorePages = True
                    '    End If

                    '    appt_mRow -= 0
                    '    appt_newPage = True
                    '    appt_PageNum += 1
                    '    Exit Sub
                    'End If
                Else
                    'If y + h > e.MarginBounds.Bottom Then
                    appt_LeftMarginAdj = 0
                    PrintDocument1.DefaultPageSettings.Margins.Left = 35 '50 '0
                    appt_ColNum = 1

                    'if appt_mRow < DGV_appt.RowCount
                    'Debug.Print(appt_mRow & " " & DGV_appt.RowCount)
                    If appt_mRow >= .RowCount Then ' DGV_appt.RowCount Then
                        e.HasMorePages = False
                    Else
                        e.HasMorePages = True
                    End If

                    appt_mRow -= 0
                    appt_newPage = True
                    appt_PageNum += 1
                    Exit Sub
                    'End If

                End If
            Loop
            appt_mRow = 0

        End With
        'End Sub
    End Sub
    Private Sub PrintCalendarOnApptPage(ByRef e As Printing.PrintPageEventArgs)
        ' Debug.Print(MonthCalendarForPrinting.Visible & " is visible or not")
        MonthCalendarForPrinting.Visible = True
        Dim adjHeight As Integer = 0
        If MonthCalendarForPrinting.ShowToday = False Then
            adjHeight = 9 '10 '15
        End If
        '=================print calendar
        'MonthCalendarForPrinting.SendToBack()
        'print calendar
        'https://www.daniweb.com/programming/software-development/threads/350417/vb-net-printing-print-preview-multiple-tabpages
        'https://stackoverflow.com/questions/4664217/how-do-i-print-just-one-or-two-controls-instead-of-the-whole-form-in-visual-basi

        Dim myBrushCalendar As System.Drawing.SolidBrush = New System.Drawing.SolidBrush(System.Drawing.Color.DarkBlue)
        'e.Graphics.FillRectangle(myBrushCalendar, New Rectangle(rc.Left, rc.Top, rc.Width, rc.Height))
        'e.Graphics.FillRectangle(myBrushCalendar, New Rectangle(35, 65, 750, 200)) '150
        'e.Graphics.FillRectangle(myBrushCalendar, New Rectangle(35, 55, 750, 200)) '150
        'e.Graphics.FillRectangle(myBrushCalendar, New Rectangle(35, 60, 750, 190)) 'up 5 dn 10

        '11/2/2019
        e.Graphics.FillRectangle(myBrushCalendar, New Rectangle(e.MarginBounds.Left, 60, 750, 190)) 'up 5 dn 10
        'now paint 3 boxes with different colors

        ' Debug.Print("e.MarginBounds.Left=" & e.MarginBounds.Left & " MonthCalendarForPrinting.Width=" & MonthCalendarForPrinting.Width)
        'MonthCalendarForPrinting.Width=" & MonthCalendarForPrinting.Width = 689 - not 542
        'Debug.Print(" MonthCalendarForPrinting.Height=" & MonthCalendarForPrinting.Height) ' first time in 172 - then 162


        Dim calendarW As Single = (MonthCalendarForPrinting.Width - 50) / 3 + 3
        Dim k As Integer
        'For k = 0 To 2
        '    Select Case k
        '        Case 0 : myBrushCalendar = New System.Drawing.SolidBrush(System.Drawing.Color.Azure)
        '        Case 1 : myBrushCalendar = New System.Drawing.SolidBrush(System.Drawing.Color.Cornsilk)
        '        Case 2 : myBrushCalendar = New System.Drawing.SolidBrush(System.Drawing.Color.OrangeRed)

        '    End Select
        '    'e.Graphics.FillRectangle(myBrushCalendar, New Rectangle(k * 233 + 70, 60, 233, 10)) 'up 5 dn 10
        '    'e.Graphics.FillRectangle(myBrushCalendar, New Rectangle(k * 233 + 70, 63, calendarW, 5)) 'up 5 dn 10
        '    e.Graphics.FillRectangle(myBrushCalendar, New Rectangle(k * 233 + 70, 63, calendarW, 184)) 'up 5 dn 10
        'Next


        Dim xBitMap As System.Drawing.Bitmap
        xBitMap = New Bitmap(MonthCalendarForPrinting.Width, MonthCalendarForPrinting.Height)

        'xBitMap.SetResolution(900, 900) 'new 11/1/2019
        ' Me.MonthCalendarForPrinting.Font.SizeInPoints= 50 '== New Font("microsoft sans serif", 30, FontStyle.Regular)
        'Using G = Graphics.FromImage(xBitMap)
        '    'Paint the canvas white
        '    G.Clear(Color.White)
        '    'Set various modes to higher quality
        '    G.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        '    G.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        '    G.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        'End Using



        Me.MonthCalendarForPrinting.DrawToBitmap(xBitMap, Me.MonthCalendarForPrinting.ClientRectangle)
        'g.DrawImage(xBitMap, New Point(leftSideMargin + 6, startTop + 5))
        'g.DrawImage(xBitMap, New Point(35, 70))
        'e.Graphics.DrawImage(xBitMap, New Point(50, 70)) '(35, 70))
        'e.Graphics.DrawImage(xBitMap, New Point(51, 70)) '(35, 70))'e.MarginBounds.Left (50-35=15 now 16)
        e.Graphics.DrawImage(xBitMap, New Point(e.MarginBounds.Left + 16, 70)) '(35, 70))'e.MarginBounds.Left (50-35=15 now 16)

        'Debug.Print(MonthCalendarForPrinting.Width)

        'ControlPaint.DrawBorder(e.Graphics, Me.myControl1.ClientRectangle, Color.Black, ButtonBorderStyle.Solid)
        'ControlPaint.DrawBorder(e.Graphics, MonthCalendarForPrinting.ClientRectangle, Color.Black, ButtonBorderStyle.Solid)

        'Dim Pen As Pen = New Pen(Brushes.LightGray) '(Color.FromArgb(255, 0, 0, 0))
        Dim Pen As Pen = New Pen(Brushes.SlateGray, 1)

        Dim x233 As Integer = 230 '227
        'Debug.Print(MonthCalendarForPrinting.Width)

        Dim ww As Single = MonthCalendarForPrinting.Bounds.Width
        'Debug.Print(ww)


        Dim PenThick As Pen = New Pen(Brushes.Black, 5)
        'e.Graphics.DrawLine(PenThick, 63, 75, 63 + 230, 75)


        Dim PenOrange As Pen = New Pen(Brushes.OrangeRed, 7)
        Dim PenDarkBlue As Pen = New Pen(Brushes.SlateGray, 6) '.DarkSlateBlue, 6)
        For k = 1 To 2 '0 To 3

            'e.Graphics.DrawLine(PenDarkBlue, 70 + (k * 230), 102, (k * 230) + 70 + calendarW, 102)
            '63-35=28 so it's e.MarginBounds.Left + 28
            Dim xxLeft As Integer = k * x233 + (e.MarginBounds.Left + 28) '63
            Dim xxTop As Integer = 73 + adjHeight
            e.Graphics.DrawLine(Pen, xxLeft, xxTop, xxLeft, xxTop + MonthCalendarForPrinting.Height - 2 * adjHeight)
        Next
        For k = 0 To 2

            'e.Graphics.DrawLine(PenDarkBlue, 70 + (k * 230), 102 + adjHeight, (k * 230) + 70 + calendarW, 102 + adjHeight)
            '70-35=35 so it's (e.MarginBounds.Left + 35)
            e.Graphics.DrawLine(PenDarkBlue, (e.MarginBounds.Left + 35) + (k * 230), 102 + adjHeight, (k * 230) + (e.MarginBounds.Left + 35) + calendarW, 102 + adjHeight)

        Next

        Using myBrushWhite As SolidBrush = New System.Drawing.SolidBrush(Color.White) '.Color.Red)        'Dim xLeft As Integer = 68 '70 (68-35 = 33 so (e.MarginBounds.Left + 33)
            Dim xLeft As Integer = (e.MarginBounds.Left + 33) '70 (68-35 = 33 so (e.MarginBounds.Left + 33)

            Dim xTop As Integer = 82 + adjHeight
            Dim xWidth As Integer = 20
            Dim xHeight As Integer = 15 '20
            e.Graphics.FillRectangle(myBrushWhite, New Rectangle(xLeft, xTop, xWidth, xHeight))

            'e.Graphics.FillRectangle(myBrushWhite, New Rectangle(730, xTop, xWidth, xHeight))
            '730-35 =695 so (e.MarginBounds.Left + 695)
            e.Graphics.FillRectangle(myBrushWhite, New Rectangle((e.MarginBounds.Left + 695), xTop, xWidth, xHeight))
        End Using

        '=================END print calendar
        xBitMap.Dispose()
    End Sub

    Private Sub B_BackUpPermanent_Click(sender As Object, e As EventArgs) Handles B_BackUpPermanent.Click
        BackUp_Permanent_DonaldStuartCampbellPersonalFolder()
        MsgBox("Permanent Backup Done!")
        End
    End Sub

    Private Sub CheckBoxIncludeItemsCompletedToday_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxIncludeItemsCompletedToday.CheckedChanged
        If skip Then Exit Sub
        Dim x As Integer = currentrow() - 1
        'Debug.Print(x)
        'Recs2Use = Recs.Clone
        'test_sampleItems()
        setRecsToBeDisplayed()
        refreshDataGridView1()

        If x >= 0 AndAlso x <= DataGridView1.Rows.Count - 1 Then
            DataGridView1.Rows(x).Selected = True
        End If

        'Try
        '    DataGridView1.Rows(x).Selected = True
        'Catch ex As Exception
        '    Debug.Print(x)
        'End Try
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    FontDialog1.ShowDialog()
    '    'DataGridView1.Font = FontDialog1.Font
    '    'DataGridView1.Appearance.Row.FontSizeDelta = 5
    '    DataGridView1.Columns(2).DefaultCellStyle.Font = FontDialog1.Font


    '    'DataGridView1.Columns(0).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
    '    'DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
    '    'DataGridView1.Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
    'End Sub

    'Private Sub ToolStripComboBox1_Click(sender As Object, e As EventArgs) Handles ToolStripComboBox1.Click
    '    'Debug.Print(ToolStripComboBox1.ComboBox.SelectedIndex)
    'End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        If skip Then Exit Sub
        'Debug.Print(ToolStripComboBox1.ComboBox.SelectedIndex)
        'Debug.Print(ToolStripComboBox1.SelectedIndex)

        'Debug.Print(ToolStripComboBox1.SelectedIndex)
        'Debug.Print(ToolStripComboBox1.Text)

        Dim s As Integer = CInt(ToolStripComboBox1.Text)
        put_FontSizeInUse(s)
        ApptGV.fontSize = s
        changeFontSize() '(s)

        'Debug.Print(sender.ToString)
        'lStripComboBox1.


        'DataGridView1.Focus()

        FontSizeToolStripMenuItem.HideDropDown()


    End Sub

    Private Sub changeFontSize() '(ByVal s As Integer)
        DataGridView1.Columns(0).DefaultCellStyle.Font = New Font("microsoft sans serif", ApptGV.fontSize, FontStyle.Regular)
        DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", ApptGV.fontSize, FontStyle.Regular)
        DataGridView1.Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", ApptGV.fontSize, FontStyle.Regular)
    End Sub

    Private Sub ComboBoxFontSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxFontSize.SelectedIndexChanged
        If skip Then Exit Sub
        Dim s As Integer = CInt(ComboBoxFontSize.Text)
        put_FontSizeInUse(s)
        ApptGV.fontSize = s
        'changeFontSize(s)
        changeFontSize()
    End Sub

    Private Sub B_ListSelector_Click(sender As Object, e As EventArgs) Handles B_ListSelector.Click
        ListSelector()
    End Sub
    Private Sub ListSelector()
        F_ListSelector.ShowDialog()

        Restart()

        setFileSetName_BeingUsed() 'sets: ApptGV.nameOfFileSet_BeingUsed
        'Lbl_FileNameBeingUsed.Text = "Todo List Name: " & ApptGV.nameOfFileSet_BeingUsed
        Lbl_FileNameBeingUsed.Text = ApptGV.nameOfFileSet_BeingUsed

    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintTodoList()
    End Sub

    Private Sub ReorderListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReorderListToolStripMenuItem.Click
        ReorderList()
    End Sub

    Private Sub ListSelectorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListSelectorToolStripMenuItem.Click
        ListSelector()
    End Sub

    Private Sub F_Main_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Control And e.KeyCode = Keys.A Then 'this works!
            'If e.KeyCode = Keys.D AndAlso e.Modifiers.ControlKey = Keys.Control Then
            'insert date into text
            'MsgBox("CTRL + A Pressed !")
            AddTodoRecord()
        ElseIf e.Control And e.KeyCode = Keys.Z Then

            clipboardInsert()

            'ElseIf e.Control And e.KeyCode = Keys.c Then
            '    clipboardInsert()
        End If
    End Sub

    Private Sub MultiListDisplayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MultiListDisplayToolStripMenuItem.Click
        Form3RecsPerList.ShowDialog()
    End Sub

    'Private Sub ToolStripComboBox1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStripComboBox1.MouseLeave
    '    FontSizeToolStripMenuItem.HideDropDown()
    'End Sub





    'Private Sub F_Main_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
    '    'Debug.Print(e.KeyCode)
    '    'If e.KeyCode = Keys.ControlKey And e.KeyCode = Keys.D Then
    '    If e.Control And e.KeyCode = Keys.D Then 'this works!
    '        'If e.KeyCode = Keys.D AndAlso e.Modifiers.ControlKey = Keys.Control Then
    '        'insert date into text
    '        MsgBox("CTRL + D Pressed !")
    '    End If
    'End Sub
#End Region'Printing
#Region "copy2clipboard"

    'Private Sub DataGridView1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
    '    currentIndexG = DataGridView1.CurrentRow.Index
    '    'Debug.Print(DataGridView1.CurrentRow.Index & "index")
    '    'If DataGridView1.SelectedRows.Count = 0 Then Return
    '    'DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.SelectedRows(0).Index
    'End Sub

    Private Sub B_Copy2Clipboard_Click(sender As Object, e As EventArgs) Handles B_Copy2Clipboard.Click
        InsertIntoClipBoard()
        'If currentIndexG < 0 Then Exit Sub

        'Clipboard.Clear()
        'Rec2Edit = ApptGV.Use_ApptTodoRecords(currentIndexG + 1)
        'Dim Msg As String = TrimF(Rec2Edit.Msg)

        '' Debug.Print(Msg)
        'Clipboard.SetText(Msg)
    End Sub
    Private Sub clipboardInsert()
        'If Trim(Msg2clipboard) = "" Then Exit Sub
        InsertIntoClipBoard()
        'If currentIndexG < 0 Then Exit Sub

        'Clipboard.Clear()
        'Rec2Edit = ApptGV.Use_ApptTodoRecords(currentIndexG + 1)
        'Dim Msg As String = TrimF(Rec2Edit.Msg)

        '' Debug.Print(Msg)
        'Clipboard.SetText(Msg)

    End Sub
    Private Sub DataGridView1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        If skip Then Exit Sub
        If DataGridView1.CurrentRow.Index < 0 Then Exit Sub


        If (DataGridView1.CurrentRow.Index + 1) > UBound(ApptGV.Use_ApptTodoRecords) Then
            'Debug.Print("I am here")
            'happens when I turn off the show completed items
            Exit Sub
        End If

        'blew up here
        'DataGridView1.CurrentRow.Index value was 2 therefore the index +1 exceeded the ubound of ApptGV.Use_ApptTodoRecords array
        'added above line to solve problem

        Dim xRec As ApptTodoRecordType = ApptGV.Use_ApptTodoRecords(DataGridView1.CurrentRow.Index + 1)
        Msg2clipboard = TrimF(xRec.Msg)
    End Sub
    Private Sub InsertIntoClipBoard()
        If Trim(Msg2clipboard) = "" Then
            MsgBox("No text in Clipboard",, "Clipboard Status")
            Exit Sub
        End If
        Clipboard.Clear()
        'Clipboard.SetText(Msg2clipboard & vbNewLine & "---------------" & vbNewLine & vbNewLine)
        Clipboard.SetText(Msg2clipboard)
    End Sub
    Private Sub setClipBoard()
        If DataGridView1.RowCount < 1 OrElse DataGridView1.CurrentRow.Index < 0 Then
            Msg2clipboard = ""
            Exit Sub
        End If

        Dim xRec As ApptTodoRecordType = ApptGV.Use_ApptTodoRecords(DataGridView1.CurrentRow.Index + 1)
        Msg2clipboard = TrimF(xRec.Msg)


    End Sub

#End Region
End Class
'