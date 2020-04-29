<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class F_Completed
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.MonthCalendarForPrinting = New System.Windows.Forms.MonthCalendar()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ItemNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CompletedColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.CompletedDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Msg = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Lbl_DisplayMonth = New System.Windows.Forms.Label()
        Me.B_Current = New System.Windows.Forms.Button()
        Me.B_Next = New System.Windows.Forms.Button()
        Me.B_Prev = New System.Windows.Forms.Button()
        Me.B_Print = New System.Windows.Forms.Button()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Lbl_FileNameBeingUsed = New System.Windows.Forms.Label()
        Me.B_Exit = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.MonthCalendarForPrinting)
        Me.Panel1.Controls.Add(Me.DataGridView1)
        Me.Panel1.Location = New System.Drawing.Point(12, 80)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(776, 302)
        Me.Panel1.TabIndex = 2
        '
        'MonthCalendarForPrinting
        '
        Me.MonthCalendarForPrinting.CalendarDimensions = New System.Drawing.Size(3, 1)
        Me.MonthCalendarForPrinting.Location = New System.Drawing.Point(21, 73)
        Me.MonthCalendarForPrinting.Name = "MonthCalendarForPrinting"
        Me.MonthCalendarForPrinting.TabIndex = 17
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNumber, Me.CompletedColumn, Me.CompletedDate, Me.Msg})
        Me.DataGridView1.Location = New System.Drawing.Point(10, 3)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(739, 259)
        Me.DataGridView1.TabIndex = 1
        '
        'ItemNumber
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.ItemNumber.DefaultCellStyle = DataGridViewCellStyle3
        Me.ItemNumber.HeaderText = "Item #"
        Me.ItemNumber.MinimumWidth = 50
        Me.ItemNumber.Name = "ItemNumber"
        Me.ItemNumber.Width = 50
        '
        'CompletedColumn
        '
        Me.CompletedColumn.HeaderText = "Completed"
        Me.CompletedColumn.MinimumWidth = 70
        Me.CompletedColumn.Name = "CompletedColumn"
        Me.CompletedColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CompletedColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.CompletedColumn.Width = 70
        '
        'CompletedDate
        '
        Me.CompletedDate.HeaderText = "Completed Date"
        Me.CompletedDate.MinimumWidth = 150
        Me.CompletedDate.Name = "CompletedDate"
        Me.CompletedDate.Width = 150
        '
        'Msg
        '
        Me.Msg.HeaderText = "Message"
        Me.Msg.MinimumWidth = 430
        Me.Msg.Name = "Msg"
        Me.Msg.Width = 430
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Azure
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Lbl_DisplayMonth)
        Me.Panel2.Controls.Add(Me.B_Current)
        Me.Panel2.Controls.Add(Me.B_Next)
        Me.Panel2.Controls.Add(Me.B_Prev)
        Me.Panel2.Location = New System.Drawing.Point(22, 11)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(268, 63)
        Me.Panel2.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 26)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Choose Month" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "to be displayed"
        '
        'Lbl_DisplayMonth
        '
        Me.Lbl_DisplayMonth.AutoSize = True
        Me.Lbl_DisplayMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_DisplayMonth.Location = New System.Drawing.Point(87, 9)
        Me.Lbl_DisplayMonth.Name = "Lbl_DisplayMonth"
        Me.Lbl_DisplayMonth.Size = New System.Drawing.Size(109, 20)
        Me.Lbl_DisplayMonth.TabIndex = 4
        Me.Lbl_DisplayMonth.Text = "Display Month"
        '
        'B_Current
        '
        Me.B_Current.Location = New System.Drawing.Point(202, 32)
        Me.B_Current.Name = "B_Current"
        Me.B_Current.Size = New System.Drawing.Size(54, 24)
        Me.B_Current.TabIndex = 6
        Me.B_Current.Text = "Current"
        Me.B_Current.UseVisualStyleBackColor = True
        '
        'B_Next
        '
        Me.B_Next.Location = New System.Drawing.Point(157, 32)
        Me.B_Next.Name = "B_Next"
        Me.B_Next.Size = New System.Drawing.Size(39, 24)
        Me.B_Next.TabIndex = 5
        Me.B_Next.Text = "Next"
        Me.B_Next.UseVisualStyleBackColor = True
        '
        'B_Prev
        '
        Me.B_Prev.Location = New System.Drawing.Point(91, 32)
        Me.B_Prev.Name = "B_Prev"
        Me.B_Prev.Size = New System.Drawing.Size(59, 24)
        Me.B_Prev.TabIndex = 4
        Me.B_Prev.Text = "Previous"
        Me.B_Prev.UseVisualStyleBackColor = True
        '
        'B_Print
        '
        Me.B_Print.Location = New System.Drawing.Point(645, 45)
        Me.B_Print.Name = "B_Print"
        Me.B_Print.Size = New System.Drawing.Size(116, 32)
        Me.B_Print.TabIndex = 16
        Me.B_Print.Text = "Print"
        Me.B_Print.UseVisualStyleBackColor = True
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'PrintDocument1
        '
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(313, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Todo List Name"
        '
        'Lbl_FileNameBeingUsed
        '
        Me.Lbl_FileNameBeingUsed.AutoSize = True
        Me.Lbl_FileNameBeingUsed.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Lbl_FileNameBeingUsed.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lbl_FileNameBeingUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_FileNameBeingUsed.ForeColor = System.Drawing.Color.Cyan
        Me.Lbl_FileNameBeingUsed.Location = New System.Drawing.Point(313, 36)
        Me.Lbl_FileNameBeingUsed.Name = "Lbl_FileNameBeingUsed"
        Me.Lbl_FileNameBeingUsed.Padding = New System.Windows.Forms.Padding(5)
        Me.Lbl_FileNameBeingUsed.Size = New System.Drawing.Size(172, 28)
        Me.Lbl_FileNameBeingUsed.TabIndex = 18
        Me.Lbl_FileNameBeingUsed.Text = "Lbl_FileNameBeingUsed"
        '
        'B_Exit
        '
        Me.B_Exit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.B_Exit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.B_Exit.Location = New System.Drawing.Point(653, 385)
        Me.B_Exit.Name = "B_Exit"
        Me.B_Exit.Size = New System.Drawing.Size(69, 40)
        Me.B_Exit.TabIndex = 20
        Me.B_Exit.Text = "Exit"
        Me.B_Exit.UseVisualStyleBackColor = True
        '
        'F_Completed
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.B_Exit)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Lbl_FileNameBeingUsed)
        Me.Controls.Add(Me.B_Print)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "F_Completed"
        Me.Text = "Completed Todos"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ItemNumber As DataGridViewTextBoxColumn
    Friend WithEvents CompletedColumn As DataGridViewCheckBoxColumn
    Friend WithEvents CompletedDate As DataGridViewTextBoxColumn
    Friend WithEvents Msg As DataGridViewTextBoxColumn
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label1 As Label
    Friend WithEvents Lbl_DisplayMonth As Label
    Friend WithEvents B_Current As Button
    Friend WithEvents B_Next As Button
    Friend WithEvents B_Prev As Button
    Friend WithEvents MonthCalendarForPrinting As MonthCalendar
    Friend WithEvents B_Print As Button
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
    Friend WithEvents Label2 As Label
    Friend WithEvents Lbl_FileNameBeingUsed As Label
    Friend WithEvents B_Exit As Button
End Class
