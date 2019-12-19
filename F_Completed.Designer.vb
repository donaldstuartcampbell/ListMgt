<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_Completed
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ItemNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CompletedColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.CompletedDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Msg = New System.Windows.Forms.DataGridViewTextBoxColumn()
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
        Me.Panel1.Location = New System.Drawing.Point(12, 69)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(776, 313)
        Me.Panel1.TabIndex = 2
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNumber, Me.CompletedColumn, Me.CompletedDate, Me.Msg})
        Me.DataGridView1.Location = New System.Drawing.Point(14, 24)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(739, 270)
        Me.DataGridView1.TabIndex = 1
        '
        'ItemNumber
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.ItemNumber.DefaultCellStyle = DataGridViewCellStyle1
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
        'F_Completed
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "F_Completed"
        Me.Text = "Completed Todos"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ItemNumber As DataGridViewTextBoxColumn
    Friend WithEvents CompletedColumn As DataGridViewCheckBoxColumn
    Friend WithEvents CompletedDate As DataGridViewTextBoxColumn
    Friend WithEvents Msg As DataGridViewTextBoxColumn
End Class
