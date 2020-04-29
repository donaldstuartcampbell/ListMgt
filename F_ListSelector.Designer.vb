<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class F_ListSelector
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.B_AddNewListName = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Lbl_FileNameBeingUsed = New System.Windows.Forms.Label()
        Me.B_Exit = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Lbl_showInfo = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column4, Me.Column1, Me.Column5, Me.Column2, Me.Column3})
        Me.DataGridView1.Location = New System.Drawing.Point(51, 80)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(753, 449)
        Me.DataGridView1.TabIndex = 0
        '
        'Column4
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column4.DefaultCellStyle = DataGridViewCellStyle1
        Me.Column4.HeaderText = "Item #"
        Me.Column4.MinimumWidth = 50
        Me.Column4.Name = "Column4"
        Me.Column4.Width = 50
        '
        'Column1
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.NullValue = False
        Me.Column1.DefaultCellStyle = DataGridViewCellStyle2
        Me.Column1.HeaderText = "In Use"
        Me.Column1.MinimumWidth = 60
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 60
        '
        'Column5
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Column5.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column5.HeaderText = "Edit Name"
        Me.Column5.MinimumWidth = 50
        Me.Column5.Name = "Column5"
        Me.Column5.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Column5.Width = 50
        '
        'Column2
        '
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Column2.DefaultCellStyle = DataGridViewCellStyle4
        Me.Column2.HeaderText = "List Name"
        Me.Column2.MinimumWidth = 400
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 400
        '
        'Column3
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Column3.DefaultCellStyle = DataGridViewCellStyle5
        Me.Column3.HeaderText = "Select List"
        Me.Column3.MinimumWidth = 125
        Me.Column3.Name = "Column3"
        Me.Column3.Width = 125
        '
        'B_AddNewListName
        '
        Me.B_AddNewListName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.B_AddNewListName.Location = New System.Drawing.Point(51, 560)
        Me.B_AddNewListName.Name = "B_AddNewListName"
        Me.B_AddNewListName.Size = New System.Drawing.Size(120, 45)
        Me.B_AddNewListName.TabIndex = 1
        Me.B_AddNewListName.Text = "Add New List Name"
        Me.B_AddNewListName.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(48, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(159, 13)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Todo List Name Currently in Use"
        '
        'Lbl_FileNameBeingUsed
        '
        Me.Lbl_FileNameBeingUsed.AutoSize = True
        Me.Lbl_FileNameBeingUsed.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Lbl_FileNameBeingUsed.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lbl_FileNameBeingUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_FileNameBeingUsed.ForeColor = System.Drawing.Color.Cyan
        Me.Lbl_FileNameBeingUsed.Location = New System.Drawing.Point(48, 40)
        Me.Lbl_FileNameBeingUsed.Name = "Lbl_FileNameBeingUsed"
        Me.Lbl_FileNameBeingUsed.Padding = New System.Windows.Forms.Padding(5)
        Me.Lbl_FileNameBeingUsed.Size = New System.Drawing.Size(172, 28)
        Me.Lbl_FileNameBeingUsed.TabIndex = 18
        Me.Lbl_FileNameBeingUsed.Text = "Lbl_FileNameBeingUsed"
        '
        'B_Exit
        '
        Me.B_Exit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.B_Exit.Location = New System.Drawing.Point(748, 560)
        Me.B_Exit.Name = "B_Exit"
        Me.B_Exit.Size = New System.Drawing.Size(56, 45)
        Me.B_Exit.TabIndex = 20
        Me.B_Exit.Text = "Exit"
        Me.B_Exit.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Azure
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(282, 543)
        Me.Label2.Name = "Label2"
        Me.Label2.Padding = New System.Windows.Forms.Padding(5)
        Me.Label2.Size = New System.Drawing.Size(344, 76)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "You can have any number of Todo Lists. Here you can:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "1. Add a new LIST name" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "2. " &
    "Select the List name you want to use (display)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "3. Change ( EDIT ) the name of a" &
    "n existing LIST"
        '
        'Lbl_showInfo
        '
        Me.Lbl_showInfo.AutoSize = True
        Me.Lbl_showInfo.BackColor = System.Drawing.Color.Azure
        Me.Lbl_showInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lbl_showInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_showInfo.Location = New System.Drawing.Point(273, 51)
        Me.Lbl_showInfo.Name = "Lbl_showInfo"
        Me.Lbl_showInfo.Padding = New System.Windows.Forms.Padding(5)
        Me.Lbl_showInfo.Size = New System.Drawing.Size(97, 28)
        Me.Lbl_showInfo.TabIndex = 22
        Me.Lbl_showInfo.Text = "Lbl_showInfo"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Azure
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(468, 4)
        Me.Label3.Name = "Label3"
        Me.Label3.Padding = New System.Windows.Forms.Padding(5)
        Me.Label3.Size = New System.Drawing.Size(335, 44)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "To view some to the Todos items in each LIST shown" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Right Click your mouse on the" &
    " Item # and hold down"
        '
        'F_ListSelector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(857, 627)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Lbl_showInfo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.B_Exit)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Lbl_FileNameBeingUsed)
        Me.Controls.Add(Me.B_AddNewListName)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "F_ListSelector"
        Me.Text = "List Selector"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents B_AddNewListName As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Lbl_FileNameBeingUsed As Label
    Friend WithEvents B_Exit As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Column4 As DataGridViewTextBoxColumn
    Friend WithEvents Column1 As DataGridViewCheckBoxColumn
    Friend WithEvents Column5 As DataGridViewButtonColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents Column3 As DataGridViewButtonColumn
    Friend WithEvents Lbl_showInfo As Label
    Friend WithEvents Label3 As Label
End Class
