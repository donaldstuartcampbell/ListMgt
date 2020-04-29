<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_ListAdd
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.B_Cancel = New System.Windows.Forms.Button()
        Me.B_OK = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Tb_Msg = New System.Windows.Forms.TextBox()
        Me.Lbl_AddingEditing = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Azure
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.CheckBox1)
        Me.Panel1.Location = New System.Drawing.Point(155, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(57, 69)
        Me.Panel1.TabIndex = 12
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
        'B_Cancel
        '
        Me.B_Cancel.Location = New System.Drawing.Point(262, 73)
        Me.B_Cancel.Name = "B_Cancel"
        Me.B_Cancel.Size = New System.Drawing.Size(76, 32)
        Me.B_Cancel.TabIndex = 11
        Me.B_Cancel.TabStop = False
        Me.B_Cancel.Text = "Cancel"
        Me.B_Cancel.UseVisualStyleBackColor = True
        '
        'B_OK
        '
        Me.B_OK.Location = New System.Drawing.Point(262, 35)
        Me.B_OK.Name = "B_OK"
        Me.B_OK.Size = New System.Drawing.Size(76, 32)
        Me.B_OK.TabIndex = 10
        Me.B_OK.TabStop = False
        Me.B_OK.Text = "OK"
        Me.B_OK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 159)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(172, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Press the ENTER key to complete."
        '
        'Tb_Msg
        '
        Me.Tb_Msg.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tb_Msg.Location = New System.Drawing.Point(21, 122)
        Me.Tb_Msg.MaxLength = 28
        Me.Tb_Msg.Name = "Tb_Msg"
        Me.Tb_Msg.Size = New System.Drawing.Size(317, 22)
        Me.Tb_Msg.TabIndex = 8
        '
        'Lbl_AddingEditing
        '
        Me.Lbl_AddingEditing.AutoSize = True
        Me.Lbl_AddingEditing.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_AddingEditing.Location = New System.Drawing.Point(21, 52)
        Me.Lbl_AddingEditing.Name = "Lbl_AddingEditing"
        Me.Lbl_AddingEditing.Size = New System.Drawing.Size(117, 16)
        Me.Lbl_AddingEditing.TabIndex = 7
        Me.Lbl_AddingEditing.Text = "Lbl_AddingEditing"
        '
        'F_ListAdd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(365, 211)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.B_Cancel)
        Me.Controls.Add(Me.B_OK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Tb_Msg)
        Me.Controls.Add(Me.Lbl_AddingEditing)
        Me.Name = "F_ListAdd"
        Me.Text = "F_ListAdd"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents B_Cancel As Button
    Friend WithEvents B_OK As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Tb_Msg As TextBox
    Friend WithEvents Lbl_AddingEditing As Label
End Class
