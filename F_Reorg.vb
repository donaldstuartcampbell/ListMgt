'Imports System.Data.DataTable
'Imports System.Drawing
'Imports System.Threading

Public Class F_Reorg
    Private dragBoxFromMouseDown As Rectangle
    Private rowIndexFromMouseDown As Integer
    Private rowIndexOfItemUnderMouseToDrop As Integer

    Dim skip As Boolean = True
    Dim formOrigWidth As Single
    Dim formOrigHeight As Single
    Dim origSize As Point
    Dim dticsArray() As Long
    Dim nRecords As Integer = 0

    'Dim PreviousWindowState As FormWindowState
    Private Sub F_Reorg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = 0
        Me.Top = 0
        Static firstTimeIn As Boolean = True
        If firstTimeIn = True Then
            'Me.Width = Screen.PrimaryScreen.WorkingArea.Width
            'Me.Height = 325 'Screen.PrimaryScreen.WorkingArea.Height
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height

            Me.MinimumSize = Me.Size

            origSize = New Point(Me.Size)

            firstTimeIn = False
            skip = True
            initGrid()

            formOrigWidth = Me.Width
            formOrigHeight = Me.Width
            DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", 10, FontStyle.Regular)

            'Me.MaximizeBox = False
            Me.MinimizeBox = False


            skip = False
        Else
            'Me.Width = formOrigWidth
            'Me.WindowState = FormWindowState.Normal
            'Me.Size = Me.MinimumSize

            'Me.Size = origSize 'New Size(800, 200)
            'Me.Size = New Size(formOrigWidth, formOrigHeight)
            'Me.WindowState = FormWindowState.Normal
        End If
        ApptGV.Use_ApptTodoRecords = getDeletedTodos_withinDateRange_NEW(New Date(0)) '(New Date(ApptGV.todoTicks)) '(selDate, True)
        nRecords = UBound(ApptGV.Use_ApptTodoRecords)
        QuickSort_ApptTodoRecords(ApptGV.Use_ApptTodoRecords, 1, nRecords)

        ReDim dticsArray(nRecords)

        setValues()

        setFileSetName_BeingUsed() 'sets: ApptGV.nameOfFileSet_BeingUsed
        Lbl_FileNameBeingUsed.Text = ApptGV.nameOfFileSet_BeingUsed
    End Sub
    'Class SurroundingClass
    Private Sub dataGridView1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles DataGridView1.MouseMove

        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then

            If dragBoxFromMouseDown <> Rectangle.Empty AndAlso Not dragBoxFromMouseDown.Contains(e.X, e.Y) Then
                Dim dropEffect As DragDropEffects = DataGridView1.DoDragDrop(DataGridView1.Rows(rowIndexFromMouseDown), DragDropEffects.Move)
            End If
        End If
    End Sub

    Private Sub dataGridView1_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles DataGridView1.MouseDown
        rowIndexFromMouseDown = DataGridView1.HitTest(e.X, e.Y).RowIndex

        If rowIndexFromMouseDown <> -1 Then
            Dim dragSize As Size = SystemInformation.DragSize
            dragBoxFromMouseDown = New Rectangle(New Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize)
        Else
            dragBoxFromMouseDown = Rectangle.Empty
        End If
    End Sub

    Private Sub dataGridView1_DragOver(ByVal sender As Object, ByVal e As DragEventArgs) Handles DataGridView1.DragOver
        e.Effect = DragDropEffects.Move

        '==
        If rowIndexFromMouseDown >= 0 Then
            e.Effect = DragDropEffects.Move

            Dim pt As Point = DataGridView1.PointToClient(New Point(e.X, e.Y))
            Dim TargetRowIndex As Integer = DataGridView1.HitTest(pt.X, pt.Y).RowIndex
            'TargetRowIndex = DataGridView1.HitTest(pt.X, pt.Y).RowIndex
            If TargetRowIndex <> -1 Then
                If TargetRowIndex = rowIndexFromMouseDown Then Exit Sub

                'Debug.Print(TargetRowIndex & " drag over " & rowIndexFromMouseDown)
                EnsureVisibleRow(DataGridView1, TargetRowIndex)

                'Debug.Print(TargetRowIndex & " drag over " & rowIndexFromMouseDown)
                If TargetRowIndex > rowIndexFromMouseDown Then
                    If DataGridView1.FirstDisplayedScrollingRowIndex < (DataGridView1.Rows.Count - 1) Then
                        'DataGridView2.FirstDisplayedScrollingRowIndex += 1
                        EnsureVisibleRow(DataGridView1, TargetRowIndex + 1)
                    End If
                    'TargetRowIndex -= 1
                ElseIf TargetRowIndex < rowIndexFromMouseDown Then
                    'TargetRowIndex += 1
                    If DataGridView1.FirstDisplayedScrollingRowIndex > 0 Then
                        'DataGridView1.FirstDisplayedScrollingRowIndex -= 1
                        EnsureVisibleRow(DataGridView1, TargetRowIndex - 1)
                    End If
                End If
            End If
        End If
        '==

        '    Debug.Print("drawOver")
        'Debug.Print(e.Effect)
        'DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.CurrentRow.Index
        'EnsureVisibleRow(DataGridView1, DataGridView1.CurrentRow.Index)
        ' DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Visible)
    End Sub
    'Private Sub datagridview2_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView2.DragOver
    '    'Debug.Print("DragOver")
    '    If SourceRowIndex >= 0 Then
    '        e.Effect = DragDropEffects.Move

    '        Dim pt As Point = DataGridView2.PointToClient(New Point(e.X, e.Y))
    '        'Dim TargetRowIndex As Integer = DataGridView2.HitTest(pt.X, pt.Y).RowIndex
    '        TargetRowIndex = DataGridView2.HitTest(pt.X, pt.Y).RowIndex
    '        If TargetRowIndex <> -1 Then
    '            EnsureVisibleRow(DataGridView2, TargetRowIndex)

    '            If TargetRowIndex > SourceRowIndex Then
    '                If DataGridView2.FirstDisplayedScrollingRowIndex < (DataGridView2.Rows.Count - 1) Then
    '                    'DataGridView2.FirstDisplayedScrollingRowIndex += 1
    '                    EnsureVisibleRow(DataGridView2, TargetRowIndex + 1)
    '                End If
    '                'TargetRowIndex -= 1
    '            ElseIf TargetRowIndex < SourceRowIndex Then
    '                'TargetRowIndex += 1
    '                If DataGridView2.FirstDisplayedScrollingRowIndex > 0 Then
    '                    'DataGridView2.FirstDisplayedScrollingRowIndex -= 1
    '                    EnsureVisibleRow(DataGridView2, TargetRowIndex - 1)
    '                End If
    '            End If
    '        End If
    '    End If
    'End Sub
    Private Sub dataGridView1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles DataGridView1.DragDrop
        'Dim clientPoint As Point = DataGridView1.PointToClient(New Point(e.X, e.Y)) '.PointToClient(New Point(e.X, e.Y))

        Dim clientPoint As Point = DataGridView1.PointToClient(New Point(e.X, e.Y))
        rowIndexOfItemUnderMouseToDrop = DataGridView1.HitTest(clientPoint.X, clientPoint.Y).RowIndex

        If e.Effect = DragDropEffects.Move Then
            'GetGridViewDisplayedCells(DataGridView1)
            'EnsureVisibleRow(DataGridView1, rowIndexOfItemUnderMouseToDrop)

            Dim rowToMove As DataGridViewRow = TryCast(e.Data.GetData(GetType(DataGridViewRow)), DataGridViewRow)

            'below 1 lines NEW
            If rowIndexOfItemUnderMouseToDrop = -1 Then Exit Sub



            DataGridView1.Rows.RemoveAt(rowIndexFromMouseDown)
            DataGridView1.Rows.Insert(rowIndexOfItemUnderMouseToDrop, rowToMove)

            'NEW
            DataGridView1.Rows(rowIndexOfItemUnderMouseToDrop).Selected = True
        End If
    End Sub
    'End Class
    Private Sub initGrid()
        With DataGridView1
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

            Dim i As Integer
            For i = 0 To .ColumnCount - 1
                .Columns.Item(i).SortMode = DataGridViewColumnSortMode.NotSortable '.Programmatic
            Next

        End With

    End Sub
    Private Sub setValues()
        Dim i As Integer
        Dim msg As String
        DataGridView1.Rows.Clear()

        Dim nRecs As Integer = UBound(ApptGV.Use_ApptTodoRecords)
        For i = 1 To nRecs 'ApptGV.All_ApptTodoRecords_Cnt

            'DataGridView1.Rows.Add(i.ToString("0000"), i Mod 2, TextBox1.Text, "Edit", "Delete")
            'With ApptGV.All_ApptTodoRecords(i)
            With ApptGV.Use_ApptTodoRecords(i)
                'DataGridView1.Rows.Add(.RecNumber.ToString, .Completed, .msg, "Edit", "Delete")
                msg = TrimF(.Msg)
                'DataGridView1.Rows.Add(i.ToString, .DeleteFlag, msg, "Edit", "Delete")
                'DataGridView1.Rows.Add(i.ToString, .DeleteFlag, msg)
                DataGridView1.Rows.Add(i.ToString, msg)
                'formatApptTodoDateStr()
            End With
        Next
        'setValues_LB1()
    End Sub



    Private Shared Sub EnsureVisibleRow(ByVal view As DataGridView, ByVal rowToShow As Integer)

        'GetVisibleCells(view)
        ''==
        'Dim visibleRowsCount As Integer = view.DisplayedRowCount(True)
        'Dim firstDisplayedRowIndex As Integer = view.FirstDisplayedCell.RowIndex
        'Dim lastvisibleRowIndex As Integer = (firstDisplayedRowIndex + visibleRowsCount) - 1
        ''3 rows showing
        ''first displayed =4
        ''last displayed=6
        'Debug.Print(visibleRowsCount & "<---rowCnt   1st=" & firstDisplayedRowIndex & " last=" & lastvisibleRowIndex)
        '==

        If rowToShow >= 0 AndAlso rowToShow < view.RowCount Then
            Dim countVisible = view.DisplayedRowCount(False)
            Dim firstVisible = view.FirstDisplayedScrollingRowIndex
            ' Debug.Print(countVisible & " first_visible_ScrollingRowIndex=" & firstVisible & " rowtoSHOW= " & rowToShow)
            If rowToShow < firstVisible Then
                'Debug.Print("<")
                view.FirstDisplayedScrollingRowIndex = rowToShow
            ElseIf rowToShow >= firstVisible + countVisible Then
                'Debug.Print(">=")
                'view.FirstDisplayedScrollingRowIndex = rowToShow - countVisible + 1
                view.FirstDisplayedScrollingRowIndex = rowToShow - countVisible 'changed to this
            End If
        End If
    End Sub
    Private Sub F_Main_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If skip Then Exit Sub
        DataGridView1.Columns(1).Width = 600 + (Me.Width - formOrigWidth)



        Dim s As Integer = 10
        If Me.Height > 600 Then s = 10 + (Me.Height - 600) / 200

        Debug.Print(Me.Height & " " & s)
        DataGridView1.Columns(0).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
        'DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
        DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
    End Sub

    'Private Sub DataGridView1_MouseHover(sender As Object, e As EventArgs) Handles DataGridView1.MouseHover
    '    'EnsureVisibleRow(DataGridView1, DataGridView1.CurrentRow.Index)
    '    ' DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Visible)

    '    Debug.Print(sender.ToString & " " & e.ToString)

    '    'Dim clientPoint As Point = DataGridView1.PointToClient(New Point(e.X, e.Y))
    '    'rowIndexOfItemUnderMouseToDrop = DataGridView1.HitTest(clientPoint.X, clientPoint.Y).RowIndex
    '    'Debug.Print(rowIndexOfItemUnderMouseToDrop)

    'End Sub
    Sub GetGridViewDisplayedCells(ByVal GridView As Windows.Forms.DataGridView, Optional ByVal PartialRow As Boolean = False, Optional ByVal Partialcolumn As Boolean = False)
        Dim intRowIndex As Integer = GridView.FirstDisplayedScrollingRowIndex

        While intRowIndex < GridView.DisplayedRowCount(PartialRow)
            Dim intColumnIndex As Integer = 0
            While intColumnIndex < GridView.DisplayedColumnCount(Partialcolumn)
                Dim CurrentCell As DataGridViewCell =
                GridView.Rows(intRowIndex).Cells(intColumnIndex)
                If Not CurrentCell.Value Is DBNull.Value Or Nothing Then
                    ' do what we came for
                End If
                System.Math.Min(System.Threading.Interlocked.Increment(intColumnIndex), intColumnIndex - 1)
            End While
            System.Math.Min(System.Threading.Interlocked.Increment(intRowIndex), intRowIndex - 1)
        End While
    End Sub
    'Private Sub DataGridView1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseWheel
    '    'Happening when the User used the mouse wheel 
    '    If e.Delta > 0 Then
    '        SendKeys.Send("{UP}") 'If the wheel is going up 
    '    Else
    '        SendKeys.Send("{DOWN}") 'If the wheel is going down 
    '    End If
    'End Sub
    Private Shared Sub GetVisibleCells(ByVal dgv As DataGridView)
        Dim visibleRowsCount = dgv.DisplayedRowCount(True)
        Dim firstDisplayedRowIndex = dgv.FirstDisplayedCell.RowIndex
        Dim lastvisibleRowIndex = (firstDisplayedRowIndex + visibleRowsCount) - 1

        For rowIndex As Integer = firstDisplayedRowIndex To lastvisibleRowIndex
            Debug.Print(rowIndex)
            'Dim cells = dgv.Rows(rowIndex).Cells

            'For Each cell As DataGridViewCell In cells

            '    If cell.Displayed Then
            '    End If
            'Next
        Next
    End Sub

    Private Sub B_Update_Click(sender As Object, e As EventArgs) Handles B_Update.Click
        Dim i As Integer
        Dim LocInArray(nRecords) As Integer
        For i = 1 To nRecords
            dticsArray(i) = ApptGV.Use_ApptTodoRecords(i).dTics

            LocInArray(i) = biSearch_GTE_ApptTodoRecords(dticsArray(i))

            'dticsArray(i) = ApptGV.Use_ApptTodoRecords(DataGridView1.Rows(i - 1).Cells(0).Value).dTics
            'With ApptGV.Use_ApptTodoRecords(i)
            '    Debug.Print(.Msg & " " & .dTics & " " & .ApptDateStr)
            'End With
            'Debug.Print(DataGridView1.Rows(i - 1).Cells(0).Value & " " & ApptGV.Use_ApptTodoRecords(i).dTics & " " & ApptGV.Use_ApptTodoRecords(DataGridView1.Rows(i - 1).Cells(0).Value).dTics)
        Next

        Dim xLoc As Integer = 0
        For i = 1 To nRecords
            xLoc = DataGridView1.Rows(i - 1).Cells(0).Value
            'Debug.Print(i & " " & xLoc)
            With ApptGV.All_ApptTodoRecords(LocInArray(xLoc))
                .dTics = dticsArray(i)
                .ApptDateStr = formatApptTodoDateStr(.dTics)
            End With
        Next


        For i = 1 To nRecords
            'With ApptGV.Use_ApptTodoRecords(i)
            With ApptGV.Use_ApptTodoRecords(DataGridView1.Rows(i - 1).Cells(0).Value)

                'LocInArray = biSearch_GTE_ApptTodoRecords(.dTics)
                'ApptGV.All_ApptTodoRecords(LocInArray).dTics = dticsArray(i)
                'ApptGV.All_ApptTodoRecords(LocInArray).ApptDateStr = formatApptTodoDateStr(.dTics)

                .dTics = dticsArray(i)
                .ApptDateStr = formatApptTodoDateStr(.dTics)
                'Debug.Print(.Msg & " " & .dTics & " " & .ApptDateStr)
            End With
        Next

        'For i = 1 To nRecords
        '    LocInArray = biSearch_GTE_ApptTodoRecords(dticsArray(i))
        '    With ApptGV.All_ApptTodoRecords(LocInArray)
        '        '.dTics = dticsArray(i)
        '        '.ApptDateStr = formatApptTodoDateStr(.dTics)
        '        'Debug.Print(.dTics)
        '        .dTics = ApptGV.Use_ApptTodoRecords(i).dTics
        '        .ApptDateStr = formatApptTodoDateStr(.dTics)
        '        Debug.Print(.dTics)
        '        'Debug.Print(.Msg)
        '    End With
        'Next


        'Now save records in file
        '===========NEW
        Dim ff As Integer = FreeFile()
        Dim filename As String = xBuildFullFileName(file_ListMgt.FileName)
        FileOpen(ff, filename, OpenMode.Random, , , file_ListMgt.FileLen)

        'Dim xCount As Integer = LOF(ff) / file_ListMgt.FileLen + 1

        Dim id As Integer
        For i = 1 To nRecords
            With ApptGV.Use_ApptTodoRecords(i)
                id = .ID
                FilePut(ff, ApptGV.Use_ApptTodoRecords(i), id)
            End With
        Next
        FileClose(ff)
        '===========


        'Set_ApptTodoRecordType_UsingReader()
        'Note: if not sure of "ApptGV.All_ApptTodoRecords" being populated correctly use SET... above
        'o/w use QuickSort... below

        QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt)

        set_DeleteDateArray() 'should be called CompletedDateArray

        Me.Close()


        'For i = 1 To ApptGV.All_ApptTodoRecords_Cnt

        '    With ApptGV.Use_ApptTodoRecords(DataGridView1.Rows(i - 1).Cells(0).Value)
        '        .dTics = ApptGV.All_ApptTodoRecords(i).dTics
        '        '(.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay + i
        '    End With
        'Next
        ''QuickSort_ApptRecType_dTics((ApptGV.Use_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt)
        'QuickSort_ApptTodoRecords(ApptGV.Use_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt)

        'ApptGV.All_ApptTodoRecords = ApptGV.Use_ApptTodoRecords.Clone

        'setValues()
        'setValues_LB2()
    End Sub

    Private Sub B_ExitWithoutUpdate_Click(sender As Object, e As EventArgs) Handles B_ExitWithoutUpdate.Click
        Me.Close()

    End Sub

    Private Sub F_Reorg_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Size = origSize
        'Me.WindowState = FormWindowState.Minimized
        ' e.Cancel = True
        'Me.formOrigHeight = FormWindowState.Normal
        'Me.formOrigWidth = FormWindowState.Normal

        Me.WindowState = FormWindowState.Normal

    End Sub
    'Private Sub MotionManagerDialog_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
    '    If skip Then Exit Sub


    '    If Me.Bounds.Equals(Me.RestoreBounds) And Not PreviousWindowState.Equals(Nothing) And
    '                        PreviousWindowState.Equals(System.Windows.Forms.FormWindowState.Maximized) Then
    '        ' Captures restore down action
    '    End If
    '    PreviousWindowState = Me.WindowState
    'End Sub
End Class