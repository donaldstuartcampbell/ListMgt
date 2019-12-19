Public Class F_AddTodo
    Private _AddingNewToDo As Boolean
    Public Property AddingNewToDo As Boolean
        Get
            Return _AddingNewToDo
        End Get
        Set(ByVal value As Boolean)
            _AddingNewToDo = value
        End Set
    End Property
    Private _ReturnPutOnTop As Boolean
    Public Property ReturnPutOnTop As Boolean
        Get
            Return _ReturnPutOnTop
        End Get
        Set(ByVal value As Boolean)
            _ReturnPutOnTop = value
        End Set
    End Property


    'Private _ReturnDelete As Boolean
    'Public Property ReturnDelete As Boolean
    '    Get
    '        Return _ReturnDelete
    '    End Get
    '    Set(ByVal value As Boolean)
    '        _ReturnDelete = value
    '    End Set
    'End Property

    'Private _AcctNum As Integer
    'Public Property AcctNum As Integer
    '    Get
    '        Return _AcctNum
    '    End Get
    '    Set(ByVal value As Integer)
    '        _AcctNum = value
    '    End Set
    'End Property

    'Private _passLocation As Point
    'Public Property passLocation As Point
    '    Get
    '        Return _passLocation
    '    End Get
    '    Set(ByVal value As Point)
    '        _passLocation = value
    '    End Set
    'End Property
    'Private _passDate As DateTime
    'Public Property passDate As DateTime
    '    Get
    '        Return _passDate
    '    End Get
    '    Set(ByVal value As DateTime)
    '        _passDate = value
    '    End Set
    'End Property
    'Private _retDate As DateTime
    'Public Property retDate As DateTime
    '    Get
    '        Return _retDate
    '    End Get
    '    Set(ByVal value As DateTime)
    '        _retDate = value
    '    End Set
    'End Property
    ''Private _passReturn_RescheduleDTics As Long
    ''Public Property passReturn_RescheduleDTics As Long
    ''    Get
    ''        Return _passReturn_RescheduleDTics
    ''    End Get
    ''    Set(ByVal value As Long)
    ''        _passReturn_RescheduleDTics = value
    ''    End Set
    ''End Property

    'Private _passRescheduleFlag As Boolean
    'Public Property passRescheduleFlag As Boolean
    '    Get
    '        Return _passRescheduleFlag
    '    End Get
    '    Set(ByVal value As Boolean)
    '        _passRescheduleFlag = value
    '    End Set
    'End Property
    Private _PassMessage As String
    Public Property PassMessage As String
        Get
            Return _PassMessage
        End Get
        Set(ByVal value As String)
            _PassMessage = value
        End Set
    End Property

    Private _RetMessage As String
    Public Property RetMessage As String
        Get
            Return _RetMessage
        End Get
        Set(ByVal value As String)
            _RetMessage = value
        End Set
    End Property

    Private _PassCanRollover As Boolean
    Public Property PassCanRollover As Boolean
        Get
            Return _PassCanRollover
        End Get
        Set(ByVal value As Boolean)
            _PassCanRollover = value
        End Set
    End Property

    Private _RetRolloverToCurrentMonth As Boolean
    Public Property RetRolloverToCurrentMonth As Boolean
        Get
            Return _RetRolloverToCurrentMonth
        End Get
        Set(ByVal value As Boolean)
            _RetRolloverToCurrentMonth = value
        End Set
    End Property

    Private Sub F_AddTodo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True


        Tb_Msg.MaxLength = 237 '84
        _RetMessage = "" '"Nothing"

        '==
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MaximumSize = Me.Size
        Me.MinimumSize = Me.Size
        '==
        CheckBox1.Checked = False

        _RetRolloverToCurrentMonth = False
        CheckBox2.Checked = False


        If _AddingNewToDo = True Then
            Panel2.Visible = False
            Me.Text = "Adding New Todo"
            Lbl_AddingEditing.Text = "Enter Todo Message"
            _ReturnPutOnTop = False
            ' CheckBox1.Checked = False
            '_ReturnPutOnTop = CheckBox1.Checked
            Panel1.Visible = True
            Tb_Msg.Text = ""
        Else

            If _PassCanRollover = True Then
                Panel2.Visible = True
                Panel2.Location = Panel1.Location
            Else
                Panel2.Visible = False
            End If

            Me.Text = "Editing Todo"
            Lbl_AddingEditing.Text = "Edit Todo Message"
            Tb_Msg.Text = Mid(Trim(_PassMessage), 1, 237) ' 84)

            Panel1.Visible = False
            With Tb_Msg
                .SelectionStart = 0
                .SelectionStart = .Text.Length 'put cursor at the end - but not highlighted
                '.SelectionLength = .Text.Length
            End With
        End If

        'Tb_Msg.Text = "0123456789012345678901234567890123456789012345678901234567890123456789012345678901234wWWWW"
        'Tb_Msg.Text = "1WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW1234"
        'Tb_Msg.Text = Mid(Trim(Tb_Msg.Text), 1, 84)

        Me.Show()
        Tb_Msg.Focus()

    End Sub

    'Private Sub Tb_Msg_KeyDown(sender As Object, e As KeyEventArgs) Handles Tb_Msg.KeyDown
    '    If e.KeyCode = Keys.Return Then
    '        retText()
    '    End If
    'End Sub

    Private Sub B_OK_Click(sender As Object, e As EventArgs) Handles B_OK.Click
        retText()
    End Sub

    Private Sub B_Cancel_Click(sender As Object, e As EventArgs) Handles B_Cancel.Click
        _RetMessage = ""
        Me.Close()
    End Sub
    Private Sub retText()
        _RetMessage = Trim(Mid(Trim(Tb_Msg.Text), 1, 237)) '84))
        _ReturnPutOnTop = CheckBox1.Checked
        _RetRolloverToCurrentMonth = CheckBox2.Checked
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Tb_Msg.Focus()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Tb_Msg.Focus()
    End Sub

    Private Sub F_AddTodo_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.F2 Then
            retText()
        End If
    End Sub
End Class