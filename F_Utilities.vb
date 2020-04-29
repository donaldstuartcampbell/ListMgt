Public Class F_Utilities
    Dim NamesOfFileSets() As FileSetType
    Dim FileSetNumberInUse As Integer
    Dim FileSetNumber2Use As Integer
    'Dim FileSetNamesNumbers() As FileSetType
    Private Sub F_Utilities_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'AddUtilityDirectoryIfDoesNotExist()
        'End
        'put_FileSetNumberInUse(1)
        'End
        setupUtilities()

        'put_FileSetNumberInUse(1)
        'End

        Label4.Text = getUtilityPath()

        Tb_NewFileSetName.Text = ""
        Lbl_FileSetNumber.Text = ""

        Lbl_FileSetCurrentlyInUse.Text = ""

        setListBox_FileSetNames()


        'setNamesOfFileSets()

        Lbl_SetName.Text = ""

        FileSetNumberInUse = get_FileSetNumberInUse()

        With NamesOfFileSets(FileSetNumberInUse)
            Lbl_FileSetCurrentlyInUse.Text = "FileSet Currently In Use: " & .FileSetNumber.ToString("000") & " - " & Trim(.FileSetName) 'get from file
        End With


        '1. determine current fileset being used ie 5
        '2. determine new fileset to be used - say 3
        '3.a copy ApptTodo.fil to ApptTodo.005  'in the donald... directory
        '3.b copy Birthday.fil to Birthday.005
        '4. now copy ApptTodo.003 to ApptTodo.fil (working directory) and the same for Birthday
        '   and set the FileSEt in use to 3

        'note: if files.003 don't exist them simple erase ApptTodo.fil and Birthday.fil (figure out how to create them!)

    End Sub
    Private Sub setListBox_FileSetNames()
        NamesOfFileSets = get_FileSets_UsingReader()
        Dim nRecs As Integer = UBound(NamesOfFileSets)
        Dim i As Integer
        LB_NamesOfFileSets.Items.Clear()

        For i = 1 To nRecs
            With NamesOfFileSets(i)
                LB_NamesOfFileSets.Items.Add(.FileSetNumber.ToString("000") & " - " & Trim(.FileSetName))
            End With
        Next
    End Sub
    'Private Sub setNamesOfFileSets()
    '    ReDim NamesOfFileSets(1000)
    '    Dim i As Integer
    '    'these names can never change

    '    'i += 1 : NamesOfFileSets(i) = "Main"
    '    'i += 1 : NamesOfFileSets(i) = "Clean"
    '    'i += 1 : NamesOfFileSets(i) = "Advertising"
    '    'i += 1 : NamesOfFileSets(i) = "DTC"
    '    'i += 1 : NamesOfFileSets(i) = "Don"
    '    ''===
    '    ''i += 1 : NamesOfFileSets(i) = "Main"
    '    ''i += 1 : NamesOfFileSets(i) = "Main"
    '    ''i += 1 : NamesOfFileSets(i) = "Main"
    '    ''i += 1 : NamesOfFileSets(i) = "Main"
    '    ''i += 1 : NamesOfFileSets(i) = "Main"

    '    'ReDim Preserve NamesOfFileSets(i)
    '    'With LB_NamesOfFileSets
    '    '    .Items.Clear()
    '    '    Dim j As Integer
    '    '    For j = 1 To i
    '    '        .Items.Add(j.ToString("000") & " - " & NamesOfFileSets(j))
    '    '    Next
    '    'End With
    '    'Lbl_FileSetCurrentlyInUse.Text = "FileSet Currently In Use: 0005 - Don" 'get from file
    'End Sub

    Private Sub LB_NamesOfFileSets_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LB_NamesOfFileSets.SelectedIndexChanged
        Dim index As Integer = LB_NamesOfFileSets.SelectedIndex + 1
        FileSetNumber2Use = NamesOfFileSets(index).FileSetNumber
        Lbl_FileSetNumber.Text = FileSetNumber2Use.ToString("000")
        Lbl_SetName.Text = FileSetNumber2Use.ToString("000") & " - " & NamesOfFileSets(index).FileSetName
    End Sub

    Private Sub B_AddFileSetName_Click(sender As Object, e As EventArgs) Handles B_AddFileSetName.Click
        Dim xname As String = Trim(Tb_NewFileSetName.Text)
        If xname = "" Then Exit Sub
        addFileSet(xname)

        setListBox_FileSetNames()

        Tb_NewFileSetName.Text = ""
    End Sub

    Private Sub B_UseNewFileSet_Click(sender As Object, e As EventArgs) Handles B_UseNewFileSet.Click
        'File.Copy("file-a.txt", "file-b.txt", True)
        'first write existing .fil's to xxx location
        If FileSetNumberInUse = 0 OrElse FileSetNumber2Use = 0 Then Exit Sub
        If FileSetNumberInUse = FileSetNumber2Use Then
            MsgBox("InUse and 2Use Numbers are the same - must be different")
            Exit Sub
        End If

        writeFileSetFiles(FileSetNumberInUse, FileSetNumber2Use)
        Me.Close() 'new added 12/28/2019

    End Sub

End Class