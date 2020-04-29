Public Class F_ListSelector
    Dim skip As Boolean

    Dim NamesOfFileSets() As FileSetType
    Dim FileSetNumberInUse As Integer
    Dim FileSetNumber2Use As Integer
    Dim recs() As ApptTodoRecordType
    Private Sub F_ListSelector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = 0
        Me.Top = 0
        Dim i As Integer
        Static firstTimeIn As Boolean = True

        Lbl_showInfo.Visible = False
        If firstTimeIn = True Then
            'this.dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(this.dataGridView1.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold);
            DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 11, FontStyle.Regular) '(this.dataGridView1.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
            ' DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AliceBlue



            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter


            'Me.Width = Screen.PrimaryScreen.WorkingArea.Width
            'Me.Height = 325 'Screen.PrimaryScreen.WorkingArea.Height
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height

            Me.MinimumSize = Me.Size
            Me.MaximumSize = Me.Size

            ''''origSize = New Point(Me.Size)

            firstTimeIn = False
            skip = True
            initGrid()

            ''''formOrigWidth = Me.Width
            ''''formOrigHeight = Me.Width
            DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", 11, FontStyle.Regular)

            Me.MaximizeBox = False
            Me.MinimizeBox = False


            skip = False
        End If
        '===================
        setupUtilities()
        FileSetNumberInUse = get_FileSetNumberInUse()
        'Debug.Print(FileSetNumberInUse)
        'put_FileSetNumberInUse(1)
        'End

        ''''Label4.Text = getUtilityPath()

        ''''Tb_NewFileSetName.Text = ""
        ''''Lbl_FileSetNumber.Text = ""

        ''''Lbl_FileSetCurrentlyInUse.Text = ""

        setListBox_FileSetNames()


        '''''setNamesOfFileSets()

        ''''Lbl_SetName.Text = ""



        ''''With NamesOfFileSets(FileSetNumberInUse)
        ''''    Lbl_FileSetCurrentlyInUse.Text = "FileSet Currently In Use: " & .FileSetNumber.ToString("000") & " - " & Trim(.FileSetName) 'get from file
        ''''End With
        '===================
        Lbl_FileNameBeingUsed.Text = TrimF(NamesOfFileSets(FileSetNumberInUse).FileSetName)
    End Sub
    Private Sub setListBox_FileSetNames(Optional skipMove2Row As Boolean = False)
        'Debug.Print(FileSetNumberInUse)
        NamesOfFileSets = get_FileSets_UsingReader()
        'Debug.Print(FileSetNumberInUse)
        Dim nRecs As Integer = UBound(NamesOfFileSets)
        Dim i As Integer
        Dim x As String
        Dim y As Integer = 0
        Dim z As String = ""
        DataGridView1.Rows.Clear()
        For i = 1 To nRecs
            With NamesOfFileSets(i)
                'x = .FileSetNumber.ToString("000") & " - " & Trim(.FileSetName)
                x = Trim(.FileSetName)
                'If i = FileSetNumberInUse Then 'FileSetNumber2Use Then
                If NamesOfFileSets(i).FileSetNumber = FileSetNumberInUse Then 'recNumber Then
                    y = 1
                    z = "In Use"
                    'DataGridView1.Columns(1).DefaultCellStyle.BackColor = Color.AliceBlue
                    'DataGridView1.Columns(1).DefaultCellStyle.ForeColor = Color.Red
                Else
                    y = 0
                    z = "Click to Select"
                    'DataGridView1.Columns(1).DefaultCellStyle.BackColor = Color.Green
                    'DataGridView1.Columns(1).DefaultCellStyle.ForeColor = Color.Black
                End If
                'DataGridView1.Rows.Add(y, x, z)
                DataGridView1.Rows.Add(i, y, "Edit", x, z)
                If NamesOfFileSets(i).FileSetNumber = FileSetNumberInUse Then 'recNumber Then
                    DataGridView1.Rows(i - 1).DefaultCellStyle.BackColor = Color.AliceBlue '.Columns(1).DefaultCellStyle.BackColor = Color.AliceBlue
                    'DataGridView1.Rows(i - 1).Cells(2).Style.BackColor = Color.Red

                    'DataGridView1.Rows(i - 1).Columns(1).DefaultCellStyle.ForeColor = Color.Red
                    Dim bc As DataGridViewButtonCell = New DataGridViewButtonCell()
                    'Me.DataGridView1.AdvancedCellBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None
                    'Me.DataGridView1.AdvancedCellBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None



                    bc.FlatStyle = FlatStyle.Flat
                    bc.Style.BackColor = Color.AliceBlue
                    bc.Value = "List In Use"
                    DataGridView1.Rows(i - 1).Cells(4) = bc

                End If
            End With
        Next
        If skipMove2Row = False Then
            Move2Row(FileSetNumberInUse)
        End If

        '===


    End Sub
    Private Sub initGrid()
        With DataGridView1
            Dim i As Integer
            'For i = 0 To 4
            '    DataGridView1.Columns(i).HeaderCell.Style.BackColor = Color.Black '.AliceBlue
            'Next

            .AllowDrop = True

            .SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            ''''.Size = New System.Drawing.Size(707, 268)
            .TabIndex = 0
            '==================================================
            '==
            .AllowDrop = True
            .MultiSelect = False ' this allows only 1 items to be selected at a time!
            .AllowUserToAddRows = False

            '.Columns.SortMode = DataGridViewColumnSortMode.NotSortable
            '.SortOrder = SortOrder.None
            '.SortedColumn (

            'Columns(SortOrder) = SortOrder.None

            'To make the completed column editable do the following:
            '1. set DGV readonly to false
            '2. set the individual columns that are not editable to readonly.true
            'note: doesn't work the other way aroung!
            '------------------------
            .ReadOnly = True 'False
            '.Columns(0).ReadOnly = True
            '.Columns(messageNumber).ReadOnly = True


            'DataGridView1.Columns("Completed").ReadOnly = False
            'DataGridView1.Columns(2).ReadOnly = False
            '------------------------

            'DataGridView1.AllowUserToAddRows = False

            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige


            'Dim newPadding As Padding = New Padding(0, 5, 0, 5)
            '.RowTemplate.DefaultCellStyle.Padding = newPadding
            '.Columns(1).HeaderCell.Style.Padding = New Padding(20, 20, 20, 20) '0, 5, 0, 0)
            ''''.Columns(messageNumber).DefaultCellStyle.Padding = New Padding(5, 5, 5, 5) '(20, 20, 20, 20) '0, 5, 0, 0)

            ''''.Columns(messageNumber).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            '.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
            '.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            .AllowUserToResizeRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeColumns = False


            'Dim i As Integer
            'For i = 0 To .ColumnCount - 1
            '    .AutoResizeColumnHeadersHeight.columns(i) = False
            'Next





            '.Columns(0).Font = New Font("microsoft sans serif", 20, FontStyle.Regular)
            '''' .Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", 10, FontStyle.Regular)
            '.Columns(1).DefaultCellStyle.Font.Size = 20 'dg.font.size / 2
            '.Columns(0).DefaultCellStyle.Font.Size = 20 'dg.font.size / 2

            '.Columns(5).DefaultCellStyle.Font = New Font(dg.font, emSize:=(dg.font.size / 2))

            ''''.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter '.MiddleRight

            .RowHeadersVisible = False 'this gets rid of left column!


            For i = 0 To .ColumnCount - 1
                .Columns.Item(i).SortMode = DataGridViewColumnSortMode.NotSortable '.Programmatic
            Next

        End With

    End Sub
    Private Sub Move2Row(ByVal recNumber As Integer)
        Dim nRecs As Integer = UBound(NamesOfFileSets) 'new 1/9/2020' UBound(ApptGV.Use_ApptTodoRecords)
        Dim i As Integer
        For i = 1 To nRecs 'ApptGV.All_ApptTodoRecords_Cnt
            'If ApptGV.Use_ApptTodoRecords(i).ID = recNumber Then
            If NamesOfFileSets(i).FileSetNumber = recNumber Then
                DataGridView1.Rows(i - 1).Cells(1).Selected = True
                DataGridView1.Select()

                'DataGridView1.VerticalScrollingOffset(i - 1) '.ScrollIntoView(i - 1)
                DataGridView1.FirstDisplayedScrollingRowIndex = i - 1 'this.dataGridView2.FirstDisplayedScrollingRowIndex
                Exit For
            End If
        Next
        '==
        'For i = 1 To nRecs
        '    With NamesOfFileSets(i)
        '        x = .FileSetNumber.ToString("000") & " - " & Trim(.FileSetName)
        '        If i = FileSetNumberInUse Then 'FileSetNumber2Use Then

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If skip Then Exit Sub
        If e.RowIndex = -1 Then Exit Sub 'very important to have the test! because it represents a click to the HEADER area not an actual ROW

        Dim rowIndex As Integer = e.RowIndex + 1
        Dim xLoc As Integer = 0

        Select Case e.ColumnIndex
            Case 2 'edit
                'Debug.Print("Edit" & " " & NamesOfFileSets(rowIndex).FileSetName)
                EditListName(rowIndex)
            Case 4 'select
                'If NamesOfFileSets(rowIndex).FileSetNumber = 0 Then
                '    Exit Sub
                'End If

                If NamesOfFileSets(rowIndex).FileSetNumber = FileSetNumberInUse Then

                    'MsgBox("InUse and 2Use Numbers are the same - must be different")
                    MsgBox("This LIST is already selected!" & vbNewLine & "Choose another LIST or Exit")
                    Exit Sub
                    'can't select what already exits
                End If
                Debug.Print("Selelect" & " " & NamesOfFileSets(rowIndex).FileSetName & " " & NamesOfFileSets(rowIndex).FileSetNumber & " " & FileSetNumberInUse)

                FileSetNumber2Use = NamesOfFileSets(rowIndex).FileSetNumber
                writeFileSetFiles(FileSetNumberInUse, FileSetNumber2Use)
                Me.Close() 'new added 12/28/2019
        End Select

    End Sub
    Private Sub EditListName(ByVal RowNumber As Integer)
        '=======================
        'NamesOfFileSets(rowIndex).FileSetName)

        F_ListAdd.AddingNewListName = False
        F_ListAdd.PassMessage = TrimF(NamesOfFileSets(RowNumber).FileSetName)
        F_ListAdd.ShowDialog()

        'F_AddTodo.PassCanRollover = False
        'F_AddTodo.AddingNewToDo = True
        'F_AddTodo.PassMessage = ""
        'F_AddTodo.ShowDialog()
        '============

        'Debug.Print(Format(Now, "o") & " start")
        'Dim dd As Long = Now.Ticks
        'startTime()

        If F_ListAdd.RetMessage = "" Then
            'no action
        Else
            Dim xname As String = TrimF(F_ListAdd.RetMessage)
            If xname = "" Then Exit Sub

            'addFileSet(xname)
            Dim rec As FileSetType = NamesOfFileSets(RowNumber)
            rec.FileSetName = xname
            UpdateFileSet(rec)

            setListBox_FileSetNames(True)

            'Dim nRecs As Integer = UBound(NamesOfFileSets)
            Move2Row(rec.FileSetNumber)
        End If
    End Sub
    Private Sub B_AddNewListName_Click(sender As Object, e As EventArgs) Handles B_AddNewListName.Click
        'addFileSet(XName)

        'setListBox_FileSetNames()

        'Tb_NewFileSetName.Text = ""
        '=======================
        F_ListAdd.AddingNewListName = True
        F_ListAdd.PassMessage = ""
        F_ListAdd.ShowDialog()

        'F_AddTodo.PassCanRollover = False
        'F_AddTodo.AddingNewToDo = True
        'F_AddTodo.PassMessage = ""
        'F_AddTodo.ShowDialog()
        '============

        'Debug.Print(Format(Now, "o") & " start")
        'Dim dd As Long = Now.Ticks
        'startTime()

        If F_ListAdd.RetMessage = "" Then
            'no action
        Else
            Dim xname As String = TrimF(F_ListAdd.RetMessage)
            If xname = "" Then Exit Sub

            addFileSet(xname)

            setListBox_FileSetNames(True)

            Dim nRecs As Integer = UBound(NamesOfFileSets)
            Move2Row(nRecs)
        End If

    End Sub

    Private Sub B_Exit_Click(sender As Object, e As EventArgs) Handles B_Exit.Click
        Me.Close()
    End Sub

    'Private Sub DataGridView1_MouseEnter(sender As Object, e As EventArgs) Handles DataGridView1.MouseEnter
    '    Debug.Print(sender.GetType().ToString)
    'End Sub
    'Private Sub DataGridView1_CellMouseEnter(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseEnter
    '    If e.RowIndex = -1 Then Exit Sub

    '    Debug.Print(e.RowIndex & " " & NamesOfFileSets(e.RowIndex + 1).FileSetName)

    '    recs = GetActiveTodoFrom00Xfiles(NamesOfFileSets(e.RowIndex + 1).FileSetNumber)

    '    Dim i As Integer
    '    Dim nRecs As Integer = UBound(recs)
    '    For i = 1 To nRecs
    '        Debug.Print(recs(i).dTics Mod 10000000 & " i=" & i & " msg=" & TrimF(recs(i).Msg))
    '    Next
    '    'If e.RowIndex = HighlightedRowIndex Then Return

    '    'If HighlightedRowIndex >= 0 Then
    '    '    SetRowStyle(dgvValues.Rows(HighlightedRowIndex), Nothing)
    '    'End If

    '    'HighlightedRowIndex = e.RowIndex

    '    'If HighlightedRowIndex >= 0 Then
    '    '    SetRowStyle(dgvValues.Rows(HighlightedRowIndex), HighlightStyle)
    '    'End If
    'End Sub
    'Private Sub setRecsToBeDisplayed()
    '    recs = getDeletedTodos_withinDateRange_NEW(New Date(0)) '(New Date(ApptGV.todoTicks)) '(selDate, True)
    '    QuickSort_ApptTodoRecords(ApptGV.Use_ApptTodoRecords, 1, UBound(ApptGV.Use_ApptTodoRecords))
    '    'QuickSort_ApptTodoRecords_desc(ApptGV.Use_ApptTodoRecords, 1, UBound(ApptGV.Use_ApptTodoRecords))


    '    Dim i As Integer
    '    Dim nrecs As Integer = UBound(recs)
    '    For i = 1 To nrecs
    '        Debug.Print(TrimF(recs(i).Msg))
    '    Next


    'End Sub
    Private Sub DataGridView1_CellMouseDown(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown

        If e.RowIndex = -1 Then Exit Sub
        If e.ColumnIndex <> 0 Then Exit Sub
        If e.Button = MouseButtons.Right Then
            'DataGridView1.CurrentCell = DataGridView1(e.ColumnIndex, e.RowIndex)
            show3records(e.RowIndex + 1)
            Lbl_showInfo.Visible = True
        End If
    End Sub
    Private Sub DataGridView1_CellMouseUp(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseUp
        Lbl_showInfo.Visible = False
        'If e.RowIndex = -1 Then Exit Sub
        'If e.ColumnIndex <> 0 Then Exit Sub
        'If e.Button = MouseButtons.Right Then
        '    DataGridView1.CurrentCell = DataGridView1(e.ColumnIndex, e.RowIndex)
        '    show3records(e.RowIndex + 1)
        '    Lbl_showInfo.Visible = True
        'End If
    End Sub

    Private Sub show3records(ByVal fileNumber As Integer)
        recs = GetActiveTodoFrom00XfilesNEW(NamesOfFileSets(fileNumber).FileSetNumber)
        Dim LineStr As String
        LineStr = New String("_"c, 60)

        Dim i As Integer
        Dim nRecs As Integer = UBound(recs)
        Dim xStr As String = ""
        'xStr = "List: " & Trim(NamesOfFileSets(fileNumber).FileSetName) & " - show 1st 3 Todos" & vbNewLine
        xStr = "List Name:   " & Trim(NamesOfFileSets(fileNumber).FileSetName) & vbNewLine

        For i = 1 To Math.Min(nRecs, 3)
            'Debug.Print(recs(i).dTics Mod 10000000 & " i=" & i & " msg=" & TrimF(recs(i).Msg))
            'xStr &= "Todo Number: " & i.ToString & LineStr
            xStr &= LineStr
            xStr &= vbNewLine & TrimF(recs(i).Msg)
            xStr &= vbNewLine
            xStr &= vbNewLine
        Next
        xStr &= LineStr
        If nRecs > 3 Then
            xStr &= vbNewLine
            'xStr &= "There are more Todos ... - only showing the first three!" '"..."
            xStr &= "There are more Todos for this LIST ..."
            xStr &= vbNewLine
        End If


        'Debug.Print(xStr)
        'MsgBox(TrimF(xStr))
        Lbl_showInfo.Text = TrimF(xStr)
    End Sub

    'Private Sub B_Print3RecsPerList_Click(sender As Object, e As EventArgs) Handles B_Print3RecsPerList.Click
    '    Form3RecsPerList.ShowDialog()
    '    'Dim recs() As TwoStringsAndOneIntegerType = get_3TodosFromEachList()
    '    'Dim nRecs As Integer = UBound(recs)
    '    'Dim i As Integer
    '    'For i = 1 To nRecs
    '    '    If recs(i).s1 <> recs(i - 1).s1 Then
    '    '        With recs(i)
    '    '            'Debug.Print(.i1 & " " & TrimF(.s1) & ": " & TrimF(.s2))
    '    '            Debug.Print(TrimF(.s1) & ": " & .i1 & " - " & TrimF(.s2))
    '    '        End With
    '    '    Else
    '    '        With recs(i)
    '    '            'Debug.Print(.i1 & " " & TrimF(.s1) & ": " & TrimF(.s2))
    '    '            Debug.Print("          " & .i1 & " - " & TrimF(.s2))
    '    '        End With
    '    '    End If

    '    'Next

    'End Sub
End Class