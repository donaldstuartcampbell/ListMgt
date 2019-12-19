<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_AddTodo
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
        Me.Lbl_AddingEditing = New System.Windows.Forms.Label()
        Me.Tb_Msg = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.B_OK = New System.Windows.Forms.Button()
        Me.B_Cancel = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Lbl_AddingEditing
        '
        Me.Lbl_AddingEditing.AutoSize = True
        Me.Lbl_AddingEditing.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_AddingEditing.Location = New System.Drawing.Point(12, 27)
        Me.Lbl_AddingEditing.Name = "Lbl_AddingEditing"
        Me.Lbl_AddingEditing.Size = New System.Drawing.Size(117, 16)
        Me.Lbl_AddingEditing.TabIndex = 0
        Me.Lbl_AddingEditing.Text = "Lbl_AddingEditing"
        '
        'Tb_Msg
        '
        Me.Tb_Msg.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tb_Msg.Location = New System.Drawing.Point(12, 97)
        Me.Tb_Msg.MaxLength = 237
        Me.Tb_Msg.Multiline = True
        Me.Tb_Msg.Name = "Tb_Msg"
        Me.Tb_Msg.Size = New System.Drawing.Size(600, 114)
        Me.Tb_Msg.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 225)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(147, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Press the F2 key to complete."
        '
        'B_OK
        '
        Me.B_OK.Location = New System.Drawing.Point(536, 12)
        Me.B_OK.Name = "B_OK"
        Me.B_OK.Size = New System.Drawing.Size(76, 32)
        Me.B_OK.TabIndex = 3
        Me.B_OK.TabStop = False
        Me.B_OK.Text = "OK"
        Me.B_OK.UseVisualStyleBackColor = True
        '
        'B_Cancel
        '
        Me.B_Cancel.Location = New System.Drawing.Point(536, 50)
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
        Me.Panel1.Location = New System.Drawing.Point(288, 12)
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
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.CheckBox2)
        Me.Panel2.Location = New System.Drawing.Point(48, 13)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(224, 69)
        Me.Panel2.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(163, 26)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "This is not a current month Todo." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Click below to make it current."
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(11, 47)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(193, 17)
        Me.CheckBox2.TabIndex = 3
        Me.CheckBox2.TabStop = False
        Me.CheckBox2.Text = "Check to Rollover to Current Month"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'F_AddTodo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(629, 266)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.B_Cancel)
        Me.Controls.Add(Me.B_OK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Tb_Msg)
        Me.Controls.Add(Me.Lbl_AddingEditing)
        Me.Name = "F_AddTodo"
        Me.Text = "F_AddTodo"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
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
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label3 As Label
    Friend WithEvents CheckBox2 As CheckBox
End Class
