<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3RecsPerList
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.B_Print = New System.Windows.Forms.Button()
        Me.MonthCalendarForPrinting = New System.Windows.Forms.MonthCalendar()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.B_CkAll = New System.Windows.Forms.Button()
        Me.B_UnCkAll = New System.Windows.Forms.Button()
        Me.ComboBoxNumber2Show = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lbl_NumItems = New System.Windows.Forms.Label()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3})
        Me.DataGridView1.Location = New System.Drawing.Point(36, 188)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(713, 315)
        Me.DataGridView1.TabIndex = 0
        '
        'Column1
        '
        Me.Column1.HeaderText = "List Name"
        Me.Column1.MinimumWidth = 100
        Me.Column1.Name = "Column1"
        '
        'Column2
        '
        Me.Column2.HeaderText = "Todo #"
        Me.Column2.MinimumWidth = 40
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 40
        '
        'Column3
        '
        Me.Column3.HeaderText = "Todo Message"
        Me.Column3.MinimumWidth = 500
        Me.Column3.Name = "Column3"
        Me.Column3.Width = 500
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'PrintDocument1
        '
        '
        'B_Print
        '
        Me.B_Print.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.B_Print.Location = New System.Drawing.Point(684, 144)
        Me.B_Print.Name = "B_Print"
        Me.B_Print.Size = New System.Drawing.Size(65, 37)
        Me.B_Print.TabIndex = 1
        Me.B_Print.Text = "Print"
        Me.B_Print.UseVisualStyleBackColor = True
        '
        'MonthCalendarForPrinting
        '
        Me.MonthCalendarForPrinting.CalendarDimensions = New System.Drawing.Size(3, 1)
        Me.MonthCalendarForPrinting.Location = New System.Drawing.Point(36, 309)
        Me.MonthCalendarForPrinting.Name = "MonthCalendarForPrinting"
        Me.MonthCalendarForPrinting.TabIndex = 17
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(140, 12)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(223, 154)
        Me.CheckedListBox1.TabIndex = 18
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Azure
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Padding = New System.Windows.Forms.Padding(5)
        Me.Label1.Size = New System.Drawing.Size(121, 44)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Check the LISTS" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "to be INCLUDED"
        '
        'B_CkAll
        '
        Me.B_CkAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.B_CkAll.Location = New System.Drawing.Point(10, 79)
        Me.B_CkAll.Name = "B_CkAll"
        Me.B_CkAll.Size = New System.Drawing.Size(121, 28)
        Me.B_CkAll.TabIndex = 20
        Me.B_CkAll.Text = "Check All Boxes"
        Me.B_CkAll.UseVisualStyleBackColor = True
        '
        'B_UnCkAll
        '
        Me.B_UnCkAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.B_UnCkAll.Location = New System.Drawing.Point(13, 136)
        Me.B_UnCkAll.Name = "B_UnCkAll"
        Me.B_UnCkAll.Size = New System.Drawing.Size(121, 28)
        Me.B_UnCkAll.TabIndex = 21
        Me.B_UnCkAll.Text = "Un-Check All Boxes"
        Me.B_UnCkAll.UseVisualStyleBackColor = True
        '
        'ComboBoxNumber2Show
        '
        Me.ComboBoxNumber2Show.FormattingEnabled = True
        Me.ComboBoxNumber2Show.Items.AddRange(New Object() {"All Todos", "1", "2", "3", "4", "5"})
        Me.ComboBoxNumber2Show.Location = New System.Drawing.Point(374, 60)
        Me.ComboBoxNumber2Show.Name = "ComboBoxNumber2Show"
        Me.ComboBoxNumber2Show.Size = New System.Drawing.Size(89, 21)
        Me.ComboBoxNumber2Show.TabIndex = 22
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(371, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 45)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Number of" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Todos To Show" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "for eash List"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightBlue
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ComboBoxNumber2Show)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.CheckedListBox1)
        Me.Panel1.Controls.Add(Me.B_CkAll)
        Me.Panel1.Controls.Add(Me.B_UnCkAll)
        Me.Panel1.Location = New System.Drawing.Point(36, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(472, 180)
        Me.Panel1.TabIndex = 24
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.OrangeRed
        Me.Label3.Location = New System.Drawing.Point(604, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(145, 20)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Multi-List Display"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(527, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(231, 90)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "This Form enables you to display Todos" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "from Multiple Lists - not just one List." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "You can show all LISTS or selected ones." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "You can also select the number to T" &
    "odos" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "to be displayed from each list."
        '
        'lbl_NumItems
        '
        Me.lbl_NumItems.AutoSize = True
        Me.lbl_NumItems.Location = New System.Drawing.Point(527, 156)
        Me.lbl_NumItems.Name = "lbl_NumItems"
        Me.lbl_NumItems.Size = New System.Drawing.Size(54, 13)
        Me.lbl_NumItems.TabIndex = 27
        Me.lbl_NumItems.Text = "NumItems"
        '
        'Form3RecsPerList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 515)
        Me.Controls.Add(Me.lbl_NumItems)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.MonthCalendarForPrinting)
        Me.Controls.Add(Me.B_Print)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "Form3RecsPerList"
        Me.Text = "Form3RecsPerList"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents Column3 As DataGridViewTextBoxColumn
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
    Friend WithEvents B_Print As Button
    Friend WithEvents MonthCalendarForPrinting As MonthCalendar
    Friend WithEvents CheckedListBox1 As CheckedListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents B_CkAll As Button
    Friend WithEvents B_UnCkAll As Button
    Friend WithEvents ComboBoxNumber2Show As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents lbl_NumItems As Label
End Class
