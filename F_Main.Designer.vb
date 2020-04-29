<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class F_Main
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.MonthCalendarForPrinting = New System.Windows.Forms.MonthCalendar()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ItemNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Completed = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Msg = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Edit = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Delete = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxIncludeItemsCompletedToday = New System.Windows.Forms.CheckBox()
        Me.B_AddTodo = New System.Windows.Forms.Button()
        Me.B_ShowReorg = New System.Windows.Forms.Button()
        Me.B_showCompleted = New System.Windows.Forms.Button()
        Me.B_Utilities = New System.Windows.Forms.Button()
        Me.PanelBackUpRestore = New System.Windows.Forms.Panel()
        Me.B_BackUpPermanent = New System.Windows.Forms.Button()
        Me.B_EraseAllFiles = New System.Windows.Forms.Button()
        Me.B_Restore = New System.Windows.Forms.Button()
        Me.B_Backup = New System.Windows.Forms.Button()
        Me.B_Print = New System.Windows.Forms.Button()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.Lbl_FileNameBeingUsed = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FontSizeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripComboBox1 = New System.Windows.Forms.ToolStripComboBox()
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReorderListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListSelectorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MultiListDisplayToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FontDialog1 = New System.Windows.Forms.FontDialog()
        Me.ComboBoxFontSize = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.B_ListSelector = New System.Windows.Forms.Button()
        Me.PanelHide = New System.Windows.Forms.Panel()
        Me.B_Copy2Clipboard = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.PanelBackUpRestore.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.PanelHide.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = True
        Me.Panel1.Controls.Add(Me.MonthCalendarForPrinting)
        Me.Panel1.Controls.Add(Me.DataGridView1)
        Me.Panel1.Location = New System.Drawing.Point(12, 97)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(776, 620)
        Me.Panel1.TabIndex = 0
        '
        'MonthCalendarForPrinting
        '
        Me.MonthCalendarForPrinting.CalendarDimensions = New System.Drawing.Size(3, 1)
        Me.MonthCalendarForPrinting.Location = New System.Drawing.Point(6, 146)
        Me.MonthCalendarForPrinting.Name = "MonthCalendarForPrinting"
        Me.MonthCalendarForPrinting.TabIndex = 16
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNumber, Me.Completed, Me.Msg, Me.Edit, Me.Delete})
        Me.DataGridView1.Location = New System.Drawing.Point(3, 26)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(770, 591)
        Me.DataGridView1.TabIndex = 0
        '
        'ItemNumber
        '
        Me.ItemNumber.HeaderText = "Item #"
        Me.ItemNumber.MinimumWidth = 50
        Me.ItemNumber.Name = "ItemNumber"
        Me.ItemNumber.Width = 50
        '
        'Completed
        '
        Me.Completed.HeaderText = "Completed"
        Me.Completed.MinimumWidth = 60
        Me.Completed.Name = "Completed"
        Me.Completed.Width = 60
        '
        'Msg
        '
        Me.Msg.HeaderText = "Message"
        Me.Msg.MinimumWidth = 540
        Me.Msg.Name = "Msg"
        Me.Msg.Width = 540
        '
        'Edit
        '
        Me.Edit.HeaderText = "Edit"
        Me.Edit.MinimumWidth = 35
        Me.Edit.Name = "Edit"
        Me.Edit.Width = 35
        '
        'Delete
        '
        Me.Delete.HeaderText = "Delete"
        Me.Delete.MinimumWidth = 50
        Me.Delete.Name = "Delete"
        Me.Delete.Width = 50
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.Color.Azure
        Me.GroupBox1.Controls.Add(Me.CheckBoxIncludeItemsCompletedToday)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 732)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(203, 74)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "What to Display"
        '
        'CheckBoxIncludeItemsCompletedToday
        '
        Me.CheckBoxIncludeItemsCompletedToday.AutoSize = True
        Me.CheckBoxIncludeItemsCompletedToday.BackColor = System.Drawing.Color.PowderBlue
        Me.CheckBoxIncludeItemsCompletedToday.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxIncludeItemsCompletedToday.Location = New System.Drawing.Point(11, 19)
        Me.CheckBoxIncludeItemsCompletedToday.Name = "CheckBoxIncludeItemsCompletedToday"
        Me.CheckBoxIncludeItemsCompletedToday.Padding = New System.Windows.Forms.Padding(5)
        Me.CheckBoxIncludeItemsCompletedToday.Size = New System.Drawing.Size(181, 46)
        Me.CheckBoxIncludeItemsCompletedToday.TabIndex = 18
        Me.CheckBoxIncludeItemsCompletedToday.TabStop = False
        Me.CheckBoxIncludeItemsCompletedToday.Text = "Check to Include" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Items Completed Today"
        Me.CheckBoxIncludeItemsCompletedToday.UseVisualStyleBackColor = False
        '
        'B_AddTodo
        '
        Me.B_AddTodo.Location = New System.Drawing.Point(639, 62)
        Me.B_AddTodo.Name = "B_AddTodo"
        Me.B_AddTodo.Size = New System.Drawing.Size(105, 32)
        Me.B_AddTodo.TabIndex = 4
        Me.B_AddTodo.TabStop = False
        Me.B_AddTodo.Text = "Add Todo"
        Me.B_AddTodo.UseVisualStyleBackColor = True
        '
        'B_ShowReorg
        '
        Me.B_ShowReorg.Location = New System.Drawing.Point(106, 44)
        Me.B_ShowReorg.Name = "B_ShowReorg"
        Me.B_ShowReorg.Size = New System.Drawing.Size(102, 25)
        Me.B_ShowReorg.TabIndex = 8
        Me.B_ShowReorg.TabStop = False
        Me.B_ShowReorg.Text = "Reorder Todo List"
        Me.B_ShowReorg.UseVisualStyleBackColor = True
        '
        'B_showCompleted
        '
        Me.B_showCompleted.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.B_showCompleted.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.B_showCompleted.Location = New System.Drawing.Point(255, 732)
        Me.B_showCompleted.Name = "B_showCompleted"
        Me.B_showCompleted.Size = New System.Drawing.Size(173, 35)
        Me.B_showCompleted.TabIndex = 9
        Me.B_showCompleted.Text = "Show Completed Records"
        Me.B_showCompleted.UseVisualStyleBackColor = True
        '
        'B_Utilities
        '
        Me.B_Utilities.Location = New System.Drawing.Point(128, 10)
        Me.B_Utilities.Name = "B_Utilities"
        Me.B_Utilities.Size = New System.Drawing.Size(53, 23)
        Me.B_Utilities.TabIndex = 10
        Me.B_Utilities.TabStop = False
        Me.B_Utilities.Text = "Utilities"
        Me.B_Utilities.UseVisualStyleBackColor = True
        '
        'PanelBackUpRestore
        '
        Me.PanelBackUpRestore.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PanelBackUpRestore.Controls.Add(Me.B_BackUpPermanent)
        Me.PanelBackUpRestore.Controls.Add(Me.B_EraseAllFiles)
        Me.PanelBackUpRestore.Controls.Add(Me.B_Restore)
        Me.PanelBackUpRestore.Controls.Add(Me.B_Backup)
        Me.PanelBackUpRestore.Location = New System.Drawing.Point(551, 731)
        Me.PanelBackUpRestore.Name = "PanelBackUpRestore"
        Me.PanelBackUpRestore.Size = New System.Drawing.Size(237, 76)
        Me.PanelBackUpRestore.TabIndex = 14
        '
        'B_BackUpPermanent
        '
        Me.B_BackUpPermanent.Location = New System.Drawing.Point(115, 42)
        Me.B_BackUpPermanent.Name = "B_BackUpPermanent"
        Me.B_BackUpPermanent.Size = New System.Drawing.Size(109, 29)
        Me.B_BackUpPermanent.TabIndex = 17
        Me.B_BackUpPermanent.Text = "Permanent BU"
        Me.B_BackUpPermanent.UseVisualStyleBackColor = True
        '
        'B_EraseAllFiles
        '
        Me.B_EraseAllFiles.Location = New System.Drawing.Point(115, 9)
        Me.B_EraseAllFiles.Name = "B_EraseAllFiles"
        Me.B_EraseAllFiles.Size = New System.Drawing.Size(109, 29)
        Me.B_EraseAllFiles.TabIndex = 16
        Me.B_EraseAllFiles.Text = "Erase All Files"
        Me.B_EraseAllFiles.UseVisualStyleBackColor = True
        '
        'B_Restore
        '
        Me.B_Restore.Location = New System.Drawing.Point(0, 42)
        Me.B_Restore.Name = "B_Restore"
        Me.B_Restore.Size = New System.Drawing.Size(109, 29)
        Me.B_Restore.TabIndex = 15
        Me.B_Restore.Text = "Restore"
        Me.B_Restore.UseVisualStyleBackColor = True
        '
        'B_Backup
        '
        Me.B_Backup.Location = New System.Drawing.Point(0, 7)
        Me.B_Backup.Name = "B_Backup"
        Me.B_Backup.Size = New System.Drawing.Size(109, 29)
        Me.B_Backup.TabIndex = 14
        Me.B_Backup.Text = "Backup"
        Me.B_Backup.UseVisualStyleBackColor = True
        '
        'B_Print
        '
        Me.B_Print.Location = New System.Drawing.Point(3, 4)
        Me.B_Print.Name = "B_Print"
        Me.B_Print.Size = New System.Drawing.Size(48, 25)
        Me.B_Print.TabIndex = 15
        Me.B_Print.TabStop = False
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
        'Lbl_FileNameBeingUsed
        '
        Me.Lbl_FileNameBeingUsed.AutoSize = True
        Me.Lbl_FileNameBeingUsed.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Lbl_FileNameBeingUsed.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lbl_FileNameBeingUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_FileNameBeingUsed.ForeColor = System.Drawing.Color.Cyan
        Me.Lbl_FileNameBeingUsed.Location = New System.Drawing.Point(15, 70)
        Me.Lbl_FileNameBeingUsed.Name = "Lbl_FileNameBeingUsed"
        Me.Lbl_FileNameBeingUsed.Padding = New System.Windows.Forms.Padding(5)
        Me.Lbl_FileNameBeingUsed.Size = New System.Drawing.Size(172, 28)
        Me.Lbl_FileNameBeingUsed.TabIndex = 16
        Me.Lbl_FileNameBeingUsed.Text = "Lbl_FileNameBeingUsed"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 13)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Todo List Name"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.LightSkyBlue
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FontSizeToolStripMenuItem, Me.PrintToolStripMenuItem, Me.ReorderListToolStripMenuItem, Me.ListSelectorToolStripMenuItem, Me.MultiListDisplayToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.ShowItemToolTips = True
        Me.MenuStrip1.Size = New System.Drawing.Size(800, 24)
        Me.MenuStrip1.TabIndex = 18
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FontSizeToolStripMenuItem
        '
        Me.FontSizeToolStripMenuItem.AutoToolTip = True
        Me.FontSizeToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripComboBox1})
        Me.FontSizeToolStripMenuItem.Name = "FontSizeToolStripMenuItem"
        Me.FontSizeToolStripMenuItem.Size = New System.Drawing.Size(66, 20)
        Me.FontSizeToolStripMenuItem.Text = "Font Size"
        Me.FontSizeToolStripMenuItem.ToolTipText = "Change the Size of the Font" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "in the range of 8 to 14"
        '
        'ToolStripComboBox1
        '
        Me.ToolStripComboBox1.AutoSize = False
        Me.ToolStripComboBox1.Items.AddRange(New Object() {"8", "9", "10", "11", "12", "13", "14"})
        Me.ToolStripComboBox1.Name = "ToolStripComboBox1"
        Me.ToolStripComboBox1.Size = New System.Drawing.Size(40, 23)
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.PrintToolStripMenuItem.Text = "Print"
        '
        'ReorderListToolStripMenuItem
        '
        Me.ReorderListToolStripMenuItem.AutoToolTip = True
        Me.ReorderListToolStripMenuItem.Name = "ReorderListToolStripMenuItem"
        Me.ReorderListToolStripMenuItem.Size = New System.Drawing.Size(81, 20)
        Me.ReorderListToolStripMenuItem.Text = "Reorder List"
        Me.ReorderListToolStripMenuItem.ToolTipText = "Change the Order" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "of Todo List Items"
        '
        'ListSelectorToolStripMenuItem
        '
        Me.ListSelectorToolStripMenuItem.AutoToolTip = True
        Me.ListSelectorToolStripMenuItem.Name = "ListSelectorToolStripMenuItem"
        Me.ListSelectorToolStripMenuItem.Size = New System.Drawing.Size(82, 20)
        Me.ListSelectorToolStripMenuItem.Text = "List Selector"
        Me.ListSelectorToolStripMenuItem.ToolTipText = "Use to Add a New List" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "or Select a List"
        '
        'MultiListDisplayToolStripMenuItem
        '
        Me.MultiListDisplayToolStripMenuItem.Name = "MultiListDisplayToolStripMenuItem"
        Me.MultiListDisplayToolStripMenuItem.Size = New System.Drawing.Size(111, 20)
        Me.MultiListDisplayToolStripMenuItem.Text = "Multi-List Display"
        '
        'ComboBoxFontSize
        '
        Me.ComboBoxFontSize.FormattingEnabled = True
        Me.ComboBoxFontSize.Items.AddRange(New Object() {"8", "9", "10", "11", "12", "13", "14"})
        Me.ComboBoxFontSize.Location = New System.Drawing.Point(56, 22)
        Me.ComboBoxFontSize.Name = "ComboBoxFontSize"
        Me.ComboBoxFontSize.Size = New System.Drawing.Size(52, 21)
        Me.ComboBoxFontSize.TabIndex = 20
        Me.ComboBoxFontSize.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(57, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 13)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Font Size"
        '
        'B_ListSelector
        '
        Me.B_ListSelector.Location = New System.Drawing.Point(0, 41)
        Me.B_ListSelector.Name = "B_ListSelector"
        Me.B_ListSelector.Size = New System.Drawing.Size(77, 25)
        Me.B_ListSelector.TabIndex = 22
        Me.B_ListSelector.Text = "List Selector"
        Me.B_ListSelector.UseVisualStyleBackColor = True
        '
        'PanelHide
        '
        Me.PanelHide.BackColor = System.Drawing.Color.Azure
        Me.PanelHide.Controls.Add(Me.B_Print)
        Me.PanelHide.Controls.Add(Me.B_ListSelector)
        Me.PanelHide.Controls.Add(Me.Label2)
        Me.PanelHide.Controls.Add(Me.ComboBoxFontSize)
        Me.PanelHide.Controls.Add(Me.B_Utilities)
        Me.PanelHide.Controls.Add(Me.B_ShowReorg)
        Me.PanelHide.Location = New System.Drawing.Point(245, 25)
        Me.PanelHide.Name = "PanelHide"
        Me.PanelHide.Size = New System.Drawing.Size(211, 72)
        Me.PanelHide.TabIndex = 23
        '
        'B_Copy2Clipboard
        '
        Me.B_Copy2Clipboard.Location = New System.Drawing.Point(493, 61)
        Me.B_Copy2Clipboard.Name = "B_Copy2Clipboard"
        Me.B_Copy2Clipboard.Size = New System.Drawing.Size(114, 32)
        Me.B_Copy2Clipboard.TabIndex = 24
        Me.B_Copy2Clipboard.Text = "Copy to Clipboard"
        Me.B_Copy2Clipboard.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.DarkCyan
        Me.Label3.Location = New System.Drawing.Point(490, 31)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 26)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Copy highlighted TEXT" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "to the clipboard"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'F_Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 818)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.B_Copy2Clipboard)
        Me.Controls.Add(Me.PanelHide)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Lbl_FileNameBeingUsed)
        Me.Controls.Add(Me.PanelBackUpRestore)
        Me.Controls.Add(Me.B_showCompleted)
        Me.Controls.Add(Me.B_AddTodo)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "F_Main"
        Me.Text = "List Management"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.PanelBackUpRestore.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.PanelHide.ResumeLayout(False)
        Me.PanelHide.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents B_AddTodo As Button
    Friend WithEvents B_ShowReorg As Button
    Friend WithEvents B_showCompleted As Button
    Friend WithEvents B_Utilities As Button
    Friend WithEvents PanelBackUpRestore As Panel
    Friend WithEvents B_EraseAllFiles As Button
    Friend WithEvents B_Restore As Button
    Friend WithEvents B_Backup As Button
    Friend WithEvents B_Print As Button
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
    Friend WithEvents MonthCalendarForPrinting As MonthCalendar
    Friend WithEvents Lbl_FileNameBeingUsed As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents B_BackUpPermanent As Button
    Friend WithEvents CheckBoxIncludeItemsCompletedToday As CheckBox
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FontSizeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FontDialog1 As FontDialog
    Friend WithEvents ToolStripComboBox1 As ToolStripComboBox
    Friend WithEvents ComboBoxFontSize As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents ItemNumber As DataGridViewTextBoxColumn
    Friend WithEvents Completed As DataGridViewCheckBoxColumn
    Friend WithEvents Msg As DataGridViewTextBoxColumn
    Friend WithEvents Edit As DataGridViewButtonColumn
    Friend WithEvents Delete As DataGridViewButtonColumn
    Friend WithEvents B_ListSelector As Button
    Friend WithEvents PrintToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PanelHide As Panel
    Friend WithEvents ReorderListToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ListSelectorToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MultiListDisplayToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents B_Copy2Clipboard As Button
    Friend WithEvents Label3 As Label
End Class
