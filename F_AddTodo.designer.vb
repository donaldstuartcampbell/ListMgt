<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class F_AddTodo
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
        Me.Lbl_AddingEditing = New System.Windows.Forms.Label()
        Me.Tb_Msg = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.B_OK = New System.Windows.Forms.Button()
        Me.B_Cancel = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Lbl_FileNameBeingUsed = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InsertTodaysDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InsertTodaysDateAndTimeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddOrUpdateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Lbl_AddingEditing
        '
        Me.Lbl_AddingEditing.AutoSize = True
        Me.Lbl_AddingEditing.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_AddingEditing.Location = New System.Drawing.Point(12, 53)
        Me.Lbl_AddingEditing.Name = "Lbl_AddingEditing"
        Me.Lbl_AddingEditing.Size = New System.Drawing.Size(117, 16)
        Me.Lbl_AddingEditing.TabIndex = 0
        Me.Lbl_AddingEditing.Text = "Lbl_AddingEditing"
        '
        'Tb_Msg
        '
        Me.Tb_Msg.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tb_Msg.Location = New System.Drawing.Point(12, 116)
        Me.Tb_Msg.MaxLength = 557
        Me.Tb_Msg.Multiline = True
        Me.Tb_Msg.Name = "Tb_Msg"
        Me.Tb_Msg.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Tb_Msg.Size = New System.Drawing.Size(600, 404)
        Me.Tb_Msg.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 532)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(147, 26)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Press the F2 key to complete." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Or, Chick the OK button."
        '
        'B_OK
        '
        Me.B_OK.Location = New System.Drawing.Point(536, 36)
        Me.B_OK.Name = "B_OK"
        Me.B_OK.Size = New System.Drawing.Size(76, 32)
        Me.B_OK.TabIndex = 3
        Me.B_OK.TabStop = False
        Me.B_OK.Text = "OK"
        Me.B_OK.UseVisualStyleBackColor = True
        '
        'B_Cancel
        '
        Me.B_Cancel.Location = New System.Drawing.Point(536, 74)
        Me.B_Cancel.Name = "B_Cancel"
        Me.B_Cancel.Size = New System.Drawing.Size(76, 32)
        Me.B_Cancel.TabIndex = 4
        Me.B_Cancel.TabStop = False
        Me.B_Cancel.Text = "Cancel"
        Me.B_Cancel.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Azure
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.CheckBox1)
        Me.Panel1.Location = New System.Drawing.Point(285, 37)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(224, 69)
        Me.Panel1.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(203, 26)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Note: Todo will be added to the bottom of" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "the List unless you check the box belo" &
    "w"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(15, 47)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(184, 17)
        Me.CheckBox1.TabIndex = 2
        Me.CheckBox1.TabStop = False
        Me.CheckBox1.Text = "Check to Add Todo to Top of List"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(186, 532)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(82, 13)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "Todo List Name"
        '
        'Lbl_FileNameBeingUsed
        '
        Me.Lbl_FileNameBeingUsed.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Lbl_FileNameBeingUsed.AutoSize = True
        Me.Lbl_FileNameBeingUsed.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Lbl_FileNameBeingUsed.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lbl_FileNameBeingUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_FileNameBeingUsed.ForeColor = System.Drawing.Color.Cyan
        Me.Lbl_FileNameBeingUsed.Location = New System.Drawing.Point(186, 549)
        Me.Lbl_FileNameBeingUsed.Name = "Lbl_FileNameBeingUsed"
        Me.Lbl_FileNameBeingUsed.Padding = New System.Windows.Forms.Padding(5)
        Me.Lbl_FileNameBeingUsed.Size = New System.Drawing.Size(172, 28)
        Me.Lbl_FileNameBeingUsed.TabIndex = 22
        Me.Lbl_FileNameBeingUsed.Text = "Lbl_FileNameBeingUsed"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EditToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(629, 24)
        Me.MenuStrip1.TabIndex = 24
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.InsertTodaysDateToolStripMenuItem, Me.InsertTodaysDateAndTimeToolStripMenuItem, Me.AddOrUpdateToolStripMenuItem, Me.ExitToolStripMenuItem1})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.T), System.Windows.Forms.Keys)
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "&Edit"
        '
        'InsertTodaysDateToolStripMenuItem
        '
        Me.InsertTodaysDateToolStripMenuItem.Name = "InsertTodaysDateToolStripMenuItem"
        Me.InsertTodaysDateToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
        Me.InsertTodaysDateToolStripMenuItem.Size = New System.Drawing.Size(264, 22)
        Me.InsertTodaysDateToolStripMenuItem.Text = "Insert Today's Date"
        '
        'InsertTodaysDateAndTimeToolStripMenuItem
        '
        Me.InsertTodaysDateAndTimeToolStripMenuItem.Name = "InsertTodaysDateAndTimeToolStripMenuItem"
        Me.InsertTodaysDateAndTimeToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.T), System.Windows.Forms.Keys)
        Me.InsertTodaysDateAndTimeToolStripMenuItem.Size = New System.Drawing.Size(264, 22)
        Me.InsertTodaysDateAndTimeToolStripMenuItem.Text = "Insert Today's Date and Time"
        '
        'AddOrUpdateToolStripMenuItem
        '
        Me.AddOrUpdateToolStripMenuItem.Name = "AddOrUpdateToolStripMenuItem"
        Me.AddOrUpdateToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F2
        Me.AddOrUpdateToolStripMenuItem.Size = New System.Drawing.Size(264, 22)
        Me.AddOrUpdateToolStripMenuItem.Text = "Add or Update"
        '
        'ExitToolStripMenuItem1
        '
        Me.ExitToolStripMenuItem1.Name = "ExitToolStripMenuItem1"
        Me.ExitToolStripMenuItem1.Size = New System.Drawing.Size(264, 22)
        Me.ExitToolStripMenuItem1.Text = "Exit / Cancel"
        '
        'F_AddTodo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(629, 580)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Lbl_FileNameBeingUsed)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.B_Cancel)
        Me.Controls.Add(Me.B_OK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Tb_Msg)
        Me.Controls.Add(Me.Lbl_AddingEditing)
        Me.Controls.Add(Me.MenuStrip1)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "F_AddTodo"
        Me.Text = "F_AddTodo"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Lbl_AddingEditing As Label
    Friend WithEvents Tb_Msg As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents B_OK As Button
    Friend WithEvents B_Cancel As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Lbl_FileNameBeingUsed As Label
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents EditToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents InsertTodaysDateToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents InsertTodaysDateAndTimeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AddOrUpdateToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem1 As ToolStripMenuItem
End Class
