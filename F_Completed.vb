Public Class F_Completed
    Dim skip As Boolean = True
    Dim formOrigWidth As Single
    Dim delRecs() As ApptTodoRecordType
    Private Sub F_Completed_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.Left = 0
        Me.Top = 0
        Static firstTimeIn As Boolean = True
        If firstTimeIn = True Then


            'Me.Width = Screen.PrimaryScreen.WorkingArea.Width
            'Me.Height = 300 '325 'Screen.PrimaryScreen.WorkingArea.Height
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height

            Me.MinimumSize = Me.Size

            firstTimeIn = False
            skip = True
            initGrid()

            formOrigWidth = Me.Width

            DataGridView1.Columns(3).DefaultCellStyle.Font = New Font("microsoft sans serif", 10, FontStyle.Regular)
            skip = False
        Else
            Me.WindowState = FormWindowState.Normal
            Me.Size = Me.MinimumSize
        End If
        setValues()
    End Sub
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
    Private Sub setValues() 'last 500 completed records
        delRecs = getDeletedTodos(True)
        Dim nRecs As Integer = UBound(delRecs)


        Dim i As Integer
        Dim msg As String
        DataGridView1.Rows.Clear()

        For i = 1 To Math.Min(nRecs, 500)

            'DataGridView1.Rows.Add(i.ToString("0000"), i Mod 2, TextBox1.Text, "Edit", "Delete")
            'With ApptGV.All_ApptTodoRecords(i)
            With delRecs(i) 'ApptGV.Use_ApptTodoRecords(i)
                'DataGridView1.Rows.Add(.RecNumber.ToString, .Completed, .msg, "Edit", "Delete")
                'msg = TrimF(.Msg) '& New Date(.dTicsCompleted)
                'msg = Replace(TrimF(.Msg), vbNewLine, "/\")
                msg = Trim(.Msg)

                '& New Date(.dTicsCompleted)
                'DataGridView1.Rows.Add(i.ToString, .DeleteFlag, msg, "Edit", "Delete")
                DataGridView1.Rows.Add(i.ToString, .DeleteFlag, New Date(.dTicsCompleted), msg, "Edit", "Delete")

            End With
        Next

    End Sub
    Private Sub F_Completed_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If skip Then Exit Sub
        DataGridView1.Columns(3).Width = 430 + (Me.Width - formOrigWidth)



        'Dim s As Integer = 10
        'If Me.Height > 600 Then s = 10 + (Me.Height - 600) / 200

        'Debug.Print(Me.Height & " " & s)
        ''DataGridView1.Columns(0).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
        ''DataGridView1.Columns(1).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
        'DataGridView1.Columns(3).DefaultCellStyle.Font = New Font("microsoft sans serif", s, FontStyle.Regular)
    End Sub


End Class