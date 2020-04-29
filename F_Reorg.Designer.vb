<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_Reorg
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(F_Reorg))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ItemNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Msg = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.B_Update = New System.Windows.Forms.Button()
        Me.B_ExitWithoutUpdate = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Lbl_FileNameBeingUsed = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.DataGridView1)
        Me.Panel1.Location = New System.Drawing.Point(12, 64)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(725, 356)
        Me.Panel1.TabIndex = 1
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNumber, Me.Msg})
        Me.DataGridView1.Location = New System.Drawing.Point(14, 24)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(688, 313)
        Me.DataGridView1.TabIndex = 1
        '
        'ItemNumber
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.ItemNumber.DefaultCellStyle = DataGridViewCellStyle2
        Me.ItemNumber.HeaderText = "Item #"
        Me.ItemNumber.MinimumWidth = 50
        Me.ItemNumber.Name = "ItemNumber"
        Me.ItemNumber.Width = 50
        '
        'Msg
        '
        Me.Msg.HeaderText = "Message"
        Me.Msg.MinimumWidth = 600
        Me.Msg.Name = "Msg"
        Me.Msg.Width = 600
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Azure
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(26, 423)
        Me.Label1.Name = "Label1"
        Me.Label1.Padding = New System.Windows.Forms.Padding(5)
        Me.Label1.Size = New System.Drawing.Size(381, 76)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = resources.GetString("Label1.Text")
        '
        'B_Update
        '
        Me.B_Update.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.B_Update.Location = New System.Drawing.Point(471, 426)
        Me.B_Update.Name = "B_Update"
        Me.B_Update.Size = New System.Drawing.Size(85, 34)
        Me.B_Update.TabIndex = 5
        Me.B_Update.Text = "Update"
        Me.B_Update.UseVisualStyleBackColor = True
        '
        'B_ExitWithoutUpdate
        '
        Me.B_ExitWithoutUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.B_ExitWithoutUpdate.Location = New System.Drawing.Point(577, 426)
        Me.B_ExitWithoutUpdate.Name = "B_ExitWithoutUpdate"
        Me.B_ExitWithoutUpdate.Size = New System.Drawing.Size(137, 34)
        Me.B_ExitWithoutUpdate.TabIndex = 6
        Me.B_ExitWithoutUpdate.Text = "Exit Without Update"
        Me.B_ExitWithoutUpdate.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(26, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Todo List Name"
        '
        'Lbl_FileNameBeingUsed
        '
        Me.Lbl_FileNameBeingUsed.AutoSize = True
        Me.Lbl_FileNameBeingUsed.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Lbl_FileNameBeingUsed.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lbl_FileNameBeingUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_FileNameBeingUsed.ForeColor = System.Drawing.Color.Cyan
        Me.Lbl_FileNameBeingUsed.Location = New System.Drawing.Point(26, 35)
        Me.Lbl_FileNameBeingUsed.Name = "Lbl_FileNameBeingUsed"
        Me.Lbl_FileNameBeingUsed.Padding = New System.Windows.Forms.Padding(5)
        Me.Lbl_FileNameBeingUsed.Size = New System.Drawing.Size(172, 28)
        Me.Lbl_FileNameBeingUsed.TabIndex = 20
        Me.Lbl_FileNameBeingUsed.Text = "Lbl_FileNameBeingUsed"
        '
        'F_Reorg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(749, 514)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Lbl_FileNameBeingUsed)
        Me.Controls.Add(Me.B_ExitWithoutUpdate)
        Me.Controls.Add(Me.B_Update)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "F_Reorg"
        Me.Text = "Change Order of Todos"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents B_Update As Button
    Friend WithEvents ItemNumber As DataGridViewTextBoxColumn
    Friend WithEvents Msg As DataGridViewTextBoxColumn
    Friend WithEvents B_ExitWithoutUpdate As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Lbl_FileNameBeingUsed As Label
End Class
