Public Class Form3RecsPerList
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

    Dim recs() As TwoStringsAndOneIntegerType
    Dim formOrigWidth As Integer
    Dim skip As Boolean = True
    Dim msgNum As Integer

    Dim FileSets() As FileSetType
    Dim nFileSets As Integer
    Dim FileSetArray() As Integer '0 or 1 if 0 then OK 

    Private Sub Form3RecsPerList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        skip = True

        Text = "Multi-List Display"

        lbl_NumItems.Text = ""

        CheckedListBox1.CheckOnClick = True

        FileSets = get_FileSets_UsingReader() ' As FileSetType()
        nFileSets = UBound(FileSets)
        ReDim FileSetArray(nFileSets)

        ComboBoxNumber2Show.SelectedIndex = 1


        MonthCalendarForPrinting.Visible = False

        setCheckedListBox() '1. set Name 2. turn ckbox=1 3. set FileSetArray items to 1

        msgNum = 2

        recs = get_3TodosFromEachList(FileSetArray, ComboBoxNumber2Show.SelectedIndex)

        Me.Height = Screen.PrimaryScreen.WorkingArea.Height


        initSetUp()

        showRecs()
        skip = False
    End Sub
    Private Sub setCheckedListBox()

        Dim i As Integer

        CheckedListBox1.Items.Clear()

        For i = 1 To nFileSets
            With CheckedListBox1
                .Items.Add(TrimF(FileSets(i).FileSetName))
                .SetItemChecked(i - 1, True)
            End With
            FileSetArray(i) = 1
        Next

    End Sub
    Private Sub showRecs()
        Dim nRecs As Integer = UBound(recs)
        Dim i As Integer
        'For i = 1 To nRecs
        '    If recs(i).s1 <> recs(i - 1).s1 Then
        '        With recs(i)
        '            'Debug.Print(.i1 & " " & TrimF(.s1) & ": " & TrimF(.s2))
        '            Debug.Print(TrimF(.s1) & ": " & .i1 & " - " & TrimF(.s2))
        '        End With
        '    Else
        '        With recs(i)
        '            'Debug.Print(.i1 & " " & TrimF(.s1) & ": " & TrimF(.s2))
        '            Debug.Print("          " & .i1 & " - " & TrimF(.s2))
        '        End With
        '    End If

        'Next
        '==============
        'Sub refreshDataGridView1()
        'Dim i As Integer
        Dim msg As String
        DataGridView1.Rows.Clear()


        lbl_NumItems.Text = "Number Items: " & nRecs.ToString
        For i = 1 To nRecs 'nRows 'ApptGV.All_ApptTodoRecords_Cnt
            With recs(i) 'ApptGV.Use_ApptTodoRecords(i)
                msg = TrimF(.s2)
                DataGridView1.Rows.Add(TrimF(.s1), .i1.ToString, msg)
            End With

            'If recs(i).s1 <> recs(i - 1).s1 Then
            '    With recs(i) 'ApptGV.Use_ApptTodoRecords(i)
            '        msg = TrimF(.s2)
            '        DataGridView1.Rows.Add(TrimF(.s1), .i1.ToString, msg)
            '    End With

            'Else
            '    With recs(i) 'ApptGV.Use_ApptTodoRecords(i)
            '        msg = TrimF(.s2)
            '        DataGridView1.Rows.Add("", .i1.ToString, msg)
            '    End With
            'End If
        Next

    End Sub
    Private Sub initSetUp()
        'AddHandler DataGridView1.CurrentCellDirtyStateChanged, AddressOf DataGridView1_CurrentCellDirtyStateChanged
        'AddHandler DataGridView1.CellValueChanged, AddressOf DataGridView1_CellValueChanged

        ''''CheckBoxIncludeItemsCompletedToday.Checked = False
        'RadioButton1.Checked = True

        Me.Top = 0
        Me.Left = 0
        Me.MinimumSize = Me.Size

        formOrigWidth = Me.Width

        Dim msgNum As Integer = 2

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
            ''''.Columns(msgNum).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True


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

            .Columns(0).DefaultCellStyle.WrapMode = DataGridViewTriState.True

            .Columns(msgNum).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders




            '.Columns(0).Font = New Font("microsoft sans serif", 20, FontStyle.Regular)
            '.Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", 10, FontStyle.Regular)
            .Columns(msgNum).DefaultCellStyle.Font = New Font("microsoft sans serif", ApptGV.fontSize, FontStyle.Regular)
            '.Columns(1).DefaultCellStyle.Font.Size = 20 'dg.font.size / 2
            '.Columns(0).DefaultCellStyle.Font.Size = 20 'dg.font.size / 2

            '.Columns(5).DefaultCellStyle.Font = New Font(dg.font, emSize:=(dg.font.size / 2))

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter '.MiddleRight
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False 'this gets rid of left column!


        End With

        'test_sampleItems()
    End Sub

    Private Sub Form3RecsPerList_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If skip Then Exit Sub
        'DataGridView1.Columns(msgNum).Width = 500 + (Me.Width - formOrigWidth)
        DataGridView1.Columns(2).Width = 500 + (Me.Width - formOrigWidth)

    End Sub
#Region "Print"
    Private Sub B_Print_Click(sender As Object, e As EventArgs) Handles B_Print.Click
        PrintTodoList()
    End Sub
    Private Sub PrintTodoList()

        Appt_FirstTimeIn_PrintingCalendar = True

        'DataGridView1.Columns(msgNum).DefaultCellStyle.Padding = New Padding(5, 5, 5, 5)
        DataGridView1.Columns(msgNum).DefaultCellStyle.Padding = New Padding(0, 0, 0, 0)

        'Dim DataGridView1 As DataGridView = DataGridView1
        appt_ApptRecs = ApptGV.Use_ApptTodoRecords.Clone

        'SurroundingSub()
#Region "Print Appt - Start of Routine"
        '===========================appt printing
        Dim existingFont As Font = DataGridView1.Font

        'DataGridView1.Font = New Font("Microsoft SansSerif", 8, FontStyle.Regular)


        DataGridView1.Font = New Font("Microsoft SansSerif", ApptGV.fontSize, FontStyle.Regular) 'new 1/7/2020
        'DataGridView1.Font = New Font("Segoe Print", ApptGV.fontSize, FontStyle.Regular) 'new 1/7/2020


        'DataGridView1.Width = PanelApptPrinting.Width - 20
        ''''SetUpDataGridView_printing()
        ''''SetDG_printing()

        ''''PanelApptPrinting.Width = DataGridView1.Width + 30
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
        DataGridView1.Font = New Font("Microsoft SansSerif", existingFont.Size, FontStyle.Regular)
        'DataGridView1.Width = PanelApptPrinting.Width - 20
        ''''SetUpDataGridView_printing()
        ''''SetDG_printing()

        ''''PanelApptPrinting.Width = DataGridView1.Width + 30

#End Region

        'MonthCalendarForPrinting.Visible = True

        DataGridView1.Columns(msgNum).DefaultCellStyle.Padding = New Padding(5, 5, 5, 5)
    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'Dim font_ballotBox = New Font("Microsoft Sans Serif", 10, FontStyle.Regular)
        Dim font_ballotBox = New Font("Microsoft Sans Serif", ApptGV.fontSize, FontStyle.Regular)
        Dim font_regular = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)

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
        With DataGridView1
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
                .Alignment = StringAlignment.Near
                '.Center
                .LineAlignment = StringAlignment.Center
                '.Center
            End With
            '==========================================

            ''''text = "Appointments - " & appt_Header
            'text = "Todos"


            'text = "Todo List: " & Lbl_FileNameBeingUsed.Text ' & " - as of " & asOf
            'text = "Top Three Todos from Each List in System"
            text = "Todos from Multiple Lists"


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

            Do While appt_mRow < DataGridView1.RowCount ' DGV_appt.RowCount

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
                            If DataGridView1.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString <> appt_LastDateString Then
                                appt_LastDateString = DataGridView1.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString

                                'If DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString <> appt_LastDateString Then
                                '    appt_LastDateString = DGV_appt.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString
                                If InStr(DataGridView1.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString, NoNumber) Then
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

                        Select Case cnt
                            Case 1
                                e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), font_regular, Brushes.Black, rc, firstColFormat)
                            Case 2
                                e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), font_regular, Brushes.Black, rc, firstColFormat)
                            Case 3
                                e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, SecondColFormat)

                        End Select
                        'e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, firstColFormat)


                        ''''If cnt = 2 Then
                        ''''    'BallotBoxWithCheck
                        ''''    'e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), font_ballotBox, Brushes.Black, rc, fmt)

                        ''''    'If ApptGV.Use_ApptTodoRecords(appt_mRow).DeleteFlag = 1 Then

                        ''''    If appt_ApptRecs(appt_mRow + 1).DeleteFlag = 1 Then
                        ''''        e.Graphics.DrawString(BallotBoxWithCheck, font_ballotBox, Brushes.Black, rc, SecondColFormat) 'fmt)
                        ''''    Else
                        ''''        e.Graphics.DrawString(BallotBoxWithoutCheck, font_ballotBox, Brushes.Black, rc, SecondColFormat) 'fmt)
                        ''''    End If




                        ''''ElseIf cnt = 1 Then

                        ''''    e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, firstColFormat)
                        ''''Else

                        ''''    e.Graphics.DrawString(.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, fmt)
                        ''''End If
                        '===========================================


                        'msg = .Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString()

                        'End If

                        x += rc.Width
                        h = Math.Max(h, rc.Height)
                    Next
                    appt_newPage = False
                    y += h
                    appt_mRow += 1

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



#End Region 'Print
    Private Sub CheckedListBox1_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck
        If skip Then Exit Sub
        If e.Index = -1 Then Exit Sub

        Dim index As Integer = e.Index + 1
        'Debug.Print(e.Index & " " & CheckedListBox1.GetItemChecked(index) & " " & CType(CheckedListBox1.Items(e.Index).selected))
        'FileSetArray(index) = 1

        If FileSetArray(index) = 0 Then
            FileSetArray(index) = 1
        Else
            FileSetArray(index) = 0
        End If

        recs = get_3TodosFromEachList(FileSetArray, ComboBoxNumber2Show.SelectedIndex)
        showRecs()


        'Dim x As Integer = CType(CheckedListBox1.Items(e.Index))

        ''''Debug.Print(FileSets(index).FileSetName & " " & index & " " & CheckedListBox1.GetItemChecked(index) & " " & CheckedListBox1.Items(e.Index))

        'CType(CheckedListBox1.Items(e.Index), Product).Selected = Not CheckedListBox1.GetItemChecked(e.Index)

        'TextBox1.Text = String.Join(Environment.NewLine, CType(CheckedListBox1.DataSource, List(Of Product)).
        '                               Where(Function(product) product.Selected).
        '                               Select(Function(product) product.DisplayData).ToArray())
    End Sub

    Private Sub B_CkAll_Click(sender As Object, e As EventArgs) Handles B_CkAll.Click
        skip = True
        Dim i As Integer
        For i = 1 To nFileSets
            FileSetArray(i) = 1
            CheckedListBox1.SetItemChecked(i - 1, True)
        Next
        skip = False
        recs = get_3TodosFromEachList(FileSetArray, ComboBoxNumber2Show.SelectedIndex)
        showRecs()

    End Sub

    Private Sub B_UnCkAll_Click(sender As Object, e As EventArgs) Handles B_UnCkAll.Click
        skip = True
        Dim i As Integer
        For i = 1 To nFileSets
            FileSetArray(i) = 0
            CheckedListBox1.SetItemChecked(i - 1, False)
        Next
        skip = False
        recs = get_3TodosFromEachList(FileSetArray, ComboBoxNumber2Show.SelectedIndex)
        showRecs()

    End Sub

    Private Sub ComboBoxNumber2Show_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxNumber2Show.SelectedIndexChanged
        If skip = True Then Exit Sub

        recs = get_3TodosFromEachList(FileSetArray, ComboBoxNumber2Show.SelectedIndex)
        showRecs()
    End Sub
End Class