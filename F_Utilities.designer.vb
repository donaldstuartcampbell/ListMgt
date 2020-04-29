<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_Utilities
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(F_Utilities))
        Me.LB_NamesOfFileSets = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Lbl_SetName = New System.Windows.Forms.Label()
        Me.Lbl_FileSetCurrentlyInUse = New System.Windows.Forms.Label()
        Me.Tb_NewFileSetName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.B_AddFileSetName = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.B_UseNewFileSet = New System.Windows.Forms.Button()
        Me.Lbl_FileSetNumber = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'LB_NamesOfFileSets
        '
        Me.LB_NamesOfFileSets.FormattingEnabled = True
        Me.LB_NamesOfFileSets.Location = New System.Drawing.Point(32, 120)
        Me.LB_NamesOfFileSets.Name = "LB_NamesOfFileSets"
        Me.LB_NamesOfFileSets.Size = New System.Drawing.Size(164, 251)
        Me.LB_NamesOfFileSets.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 104)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Names of File Sets"
        '
        'Lbl_SetName
        '
        Me.Lbl_SetName.AutoSize = True
        Me.Lbl_SetName.Location = New System.Drawing.Point(29, 374)
        Me.Lbl_SetName.Name = "Lbl_SetName"
        Me.Lbl_SetName.Size = New System.Drawing.Size(71, 13)
        Me.Lbl_SetName.TabIndex = 2
        Me.Lbl_SetName.Text = "Lbl_SetName"
        '
        'Lbl_FileSetCurrentlyInUse
        '
        Me.Lbl_FileSetCurrentlyInUse.AutoSize = True
        Me.Lbl_FileSetCurrentlyInUse.Location = New System.Drawing.Point(29, 414)
        Me.Lbl_FileSetCurrentlyInUse.Name = "Lbl_FileSetCurrentlyInUse"
        Me.Lbl_FileSetCurrentlyInUse.Size = New System.Drawing.Size(128, 13)
        Me.Lbl_FileSetCurrentlyInUse.TabIndex = 3
        Me.Lbl_FileSetCurrentlyInUse.Text = "Lbl_FileSetCurrentlyInUse"
        '
        'Tb_NewFileSetName
        '
        Me.Tb_NewFileSetName.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tb_NewFileSetName.Location = New System.Drawing.Point(32, 25)
        Me.Tb_NewFileSetName.Name = "Tb_NewFileSetName"
        Me.Tb_NewFileSetName.Size = New System.Drawing.Size(164, 24)
        Me.Tb_NewFileSetName.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(29, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "New FileSet Name"
        '
        'B_AddFileSetName
        '
        Me.B_AddFileSetName.Location = New System.Drawing.Point(32, 55)
        Me.B_AddFileSetName.Name = "B_AddFileSetName"
        Me.B_AddFileSetName.Size = New System.Drawing.Size(122, 22)
        Me.B_AddFileSetName.TabIndex = 6
        Me.B_AddFileSetName.Text = "Add FileSet Name"
        Me.B_AddFileSetName.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(220, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(198, 154)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = resources.GetString("Label3.Text")
        '
        'B_UseNewFileSet
        '
        Me.B_UseNewFileSet.Location = New System.Drawing.Point(271, 290)
        Me.B_UseNewFileSet.Name = "B_UseNewFileSet"
        Me.B_UseNewFileSet.Size = New System.Drawing.Size(134, 25)
        Me.B_UseNewFileSet.TabIndex = 8
        Me.B_UseNewFileSet.Text = "Use New FileSet "
        Me.B_UseNewFileSet.UseVisualStyleBackColor = True
        '
        'Lbl_FileSetNumber
        '
        Me.Lbl_FileSetNumber.AutoSize = True
        Me.Lbl_FileSetNumber.Location = New System.Drawing.Point(220, 296)
        Me.Lbl_FileSetNumber.Name = "Lbl_FileSetNumber"
        Me.Lbl_FileSetNumber.Size = New System.Drawing.Size(96, 13)
        Me.Lbl_FileSetNumber.TabIndex = 9
        Me.Lbl_FileSetNumber.Text = "Lbl_FileSetNumber"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(329, 15)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Label4"
        '
        'F_Utilities
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Lbl_FileSetNumber)
        Me.Controls.Add(Me.B_UseNewFileSet)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.B_AddFileSetName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Tb_NewFileSetName)
        Me.Controls.Add(Me.Lbl_FileSetCurrentlyInUse)
        Me.Controls.Add(Me.Lbl_SetName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LB_NamesOfFileSets)
        Me.Name = "F_Utilities"
        Me.Text = "F_Utilities"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LB_NamesOfFileSets As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Lbl_SetName As Label
    Friend WithEvents Lbl_FileSetCurrentlyInUse As Label
    Friend WithEvents Tb_NewFileSetName As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents B_AddFileSetName As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents B_UseNewFileSet As Button
    Friend WithEvents Lbl_FileSetNumber As Label
    Friend WithEvents Label4 As Label
End Class
