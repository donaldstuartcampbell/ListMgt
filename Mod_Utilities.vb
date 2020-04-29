Option Explicit On
Imports System.IO
Module Mod_Utilities
    'ApptGV.ccPath="C:\todoProgramCOM2020ListMgt"
    'Dim x As String = ApptGV.ccPath

    'Dim donaldPath As String = "c:\TodoProgramCOM2020\DonaldStuartCampbell03181940Utilities\"
    Dim UtilityPath As String = ApptGV.ccPath & "\Utilities\"

    Dim UtilityFileSet_Name As String = "UtilityFileSet.fil"
    Dim LenOfUtilityFileSet_Name As Integer = 32
    Dim FileSetNumberInUse_FileName As String = "FileSetNumberInUse.fil"
    Structure FileSetType
        <VBFixedString(28)> Dim FileSetName As String
        Dim FileSetNumber As Integer
    End Structure
    Function getUtilityPath() As String
        Return UtilityPath
    End Function
    Sub setFileSetName_BeingUsed() 'sets: ApptGV.nameOfFileSet_BeingUsed
        ApptGV.nameOfFileSet_BeingUsed = ""
        setupUtilities()

        Dim xNamesOfFileSets() As FileSetType = get_FileSets_UsingReader()
        Dim xFileSetNumberInUse As Integer = get_FileSetNumberInUse()
        ApptGV.nameOfFileSet_BeingUsed = Trim(xNamesOfFileSets(xFileSetNumberInUse).FileSetName)
    End Sub
    Sub setupUtilities()
        UtilityPath = ApptGV.ccPath & "\Utilities\"
        UtilityFileSet_Name = "UtilityFileSet.fil"
        LenOfUtilityFileSet_Name = 32
        FileSetNumberInUse_FileName = "FileSetNumberInUse.fil"
        'Debug.Print(UtilityPath)
    End Sub
    Sub AddUtilityDirectoryIfDoesNotExist()
        '1. create directory if it does not exist
        'createDirectoryTDP(UtilityPath) '(ApptGV.ccPath)
        If System.IO.Directory.Exists(UtilityPath) Then
            ' add_ApptTodoXXXAndBirthdayXXX(1)
            'addFileSet("MAIN")
            'put_FileSetNumberInUse(1)
        Else
            'create One
            'Dim di As Directory = System.IO.Directory.CreateDirectory(DirName)
            System.IO.Directory.CreateDirectory(UtilityPath)

            addFileSet("MAIN - ToDo List") 'this call automatically sets the next line
            'add_ApptTodoXXXAndBirthdayXXX(1)

            Dim ToUse As Integer = 1
            put_FileSetNumberInUse(ToUse)
        End If
    End Sub
    Sub add_ApptTodoXXXAndBirthdayXXX(ByVal xNumber As Integer)
        'Dim FullFileName As String = UtilityPath & "ListMgt." & xNumber.ToString("000")

        'Dim FullFileName As String = UtilityPath & "ListMgt." & xNumber.ToString("000")
        Dim FullFileName As String = $"{UtilityPath}ListMgt.{xNumber.ToString("000")}"

        If System.IO.File.Exists(FullFileName) Then
        Else
            System.IO.File.Create(FullFileName).Dispose()
        End If

        FullFileName = UtilityPath & "Birthday." & xNumber.ToString("000")
        If System.IO.File.Exists(FullFileName) Then
        Else
            System.IO.File.Create(FullFileName).Dispose()
        End If
    End Sub
    Sub addFileSet(ByVal fileSetName As String)
        Dim FullFileName As String = setFullUtilityFileSetName()

        If System.IO.File.Exists(FullFileName) Then
        Else
            System.IO.File.Create(FullFileName).Dispose()
        End If

        'add
        '==
        Dim ff As Integer = FreeFile()

        FileOpen(ff, FullFileName, OpenMode.Random, , , LenOfUtilityFileSet_Name)

        Dim xCount As Integer = LOF(ff) / LenOfUtilityFileSet_Name + 1

        Dim rec As FileSetType
        With rec
            .FileSetNumber = xCount
            .FileSetName = fileSetName
        End With
        FilePut(ff, rec, xCount)
        FileClose(ff)
        '==
        add_ApptTodoXXXAndBirthdayXXX(rec.FileSetNumber)
    End Sub
    Sub UpdateFileSet(ByVal rec As FileSetType) '(ByVal fileSetName As String)
        Dim FullFileName As String = setFullUtilityFileSetName()
        Dim recordNumber As Integer = rec.FileSetNumber

        'If System.IO.File.Exists(FullFileName) Then
        'Else
        '    System.IO.File.Create(FullFileName).Dispose()
        'End If

        'update
        '==
        Dim ff As Integer = FreeFile()

        FileOpen(ff, FullFileName, OpenMode.Random, , , LenOfUtilityFileSet_Name)

        'Dim xCount As Integer = LOF(ff) / LenOfUtilityFileSet_Name + 1

        'Dim rec As FileSetType
        'With rec
        '    .FileSetNumber = xCount
        '    .FileSetName = fileSetName
        'End With
        FilePut(ff, rec, recordNumber)
        FileClose(ff)
        '==
        ' add_ApptTodoXXXAndBirthdayXXX(rec.FileSetNumber)
    End Sub
    Function setFullUtilityFileSetName() As String
        Return UtilityPath & UtilityFileSet_Name
    End Function
    Function setFullFileNameUsingUtilityPath(ByVal xFileName As String) As String
        Return UtilityPath & xFileName
    End Function


    Function get_FileSets_UsingReader() As FileSetType()
        Dim cnt As Integer = 0
        Dim i As Integer

        Dim recs() As FileSetType

        Dim NumRecs As Integer

        Dim FullFileName As String = setFullUtilityFileSetName()
        If System.IO.File.Exists(FullFileName) Then
        Else
            System.IO.File.Create(FullFileName).Dispose()
        End If

        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
            Dim length As Integer = reader.BaseStream.Length
            NumRecs = length / LenOfUtilityFileSet_Name

            ReDim recs(NumRecs)

            For i = 1 To NumRecs
                With recs(i)
                    .FileSetName = reader.ReadChars(28)
                    .FileSetNumber = reader.ReadInt32
                End With
            Next
        End Using

        'QuickSort_BirthdayRecords(ApptGV.Birthdays, 1, NumRecs)

        Return recs
    End Function
    Sub put_FileSetNumberInUse(ByVal xNumber As Integer)
        Dim FullFileName As String = setFullFileNameUsingUtilityPath(FileSetNumberInUse_FileName)
        If System.IO.File.Exists(FullFileName) Then
        Else
            init_FileSetNumberInUse()
            'System.IO.File.Create(FullFileName).Dispose()
        End If
        '==
        Dim ff As Integer = FreeFile()
        FileOpen(ff, FullFileName, OpenMode.Random, , , 4)

        FilePut(ff, xNumber, 1)
        FileClose(ff)
        '==

    End Sub
    Private Sub init_FileSetNumberInUse() 'new 1/7/2020 'NOTE: added fontsize as second element
        Dim FullFileName As String = setFullFileNameUsingUtilityPath(FileSetNumberInUse_FileName)
        System.IO.File.Create(FullFileName).Dispose()
        '==
        Dim ff As Integer = FreeFile()
        FileOpen(ff, FullFileName, OpenMode.Random, , , 4)
        FilePut(ff, 1, 1)
        FilePut(ff, 10, 2) 'fixed on 1/9/2010 was ff,2,10
        FileClose(ff)
        '==
    End Sub
    Sub put_FontSizeInUse(ByVal xNumber As Integer)
        Dim FullFileName As String = setFullFileNameUsingUtilityPath(FileSetNumberInUse_FileName)
        If System.IO.File.Exists(FullFileName) Then
        Else
            init_FileSetNumberInUse()
            'System.IO.File.Create(FullFileName).Dispose()
        End If
        '==
        Dim ff As Integer = FreeFile()
        FileOpen(ff, FullFileName, OpenMode.Random, , , 4)

        FilePut(ff, xNumber, 2)
        FileClose(ff)
        '==

    End Sub

    Function get_FontSizeInUse() As Integer
        Dim FullFileName As String = setFullFileNameUsingUtilityPath(FileSetNumberInUse_FileName)
        If System.IO.File.Exists(FullFileName) Then
        Else
            'System.IO.File.Create(FullFileName).Dispose()
            init_FileSetNumberInUse()
        End If
        '==
        Dim FontSize As Integer = 0
        Dim ff As Integer = FreeFile()
        FileOpen(ff, FullFileName, OpenMode.Random, , , 4)
        FileGet(ff, FontSize, 2)
        FileClose(ff)

        'MsgBox(LOF(ff))
        'If LOF(ff) = 4 Then
        '    FontSize = 10
        '    FilePut(ff, FontSize, 2)
        'Else
        'FileGet(ff, FontSize, 2)
        'End If
        'FileClose(ff)
        '==
        'If FontSize = 0 Then FontSize = 10
        Return FontSize
    End Function
    Function get_FileSetNumberInUse() As Integer
        Dim FullFileName As String = setFullFileNameUsingUtilityPath(FileSetNumberInUse_FileName)
        If System.IO.File.Exists(FullFileName) Then
        Else
            'System.IO.File.Create(FullFileName).Dispose()
            init_FileSetNumberInUse()
        End If
        '==
        Dim NumInUse As Integer = 0
        Dim ff As Integer = FreeFile()
        FileOpen(ff, FullFileName, OpenMode.Random, , , 4)

        FileGet(ff, NumInUse, 1)
        FileClose(ff)
        '==
        Return NumInUse
    End Function
    Function get_NumberAndName_FileSetInUse() As String
        Dim NumberInUse As Integer = get_FileSetNumberInUse()
        '---
        Dim FullFileName As String = setFullUtilityFileSetName()


        'open and get FileSet
        '==
        Dim ff As Integer = FreeFile()
        FileOpen(ff, FullFileName, OpenMode.Random, , , LenOfUtilityFileSet_Name)
        Dim rec As FileSetType
        FileGet(ff, rec, NumberInUse)
        FileClose(ff)
        '==
        Dim x As String = ""
        With rec
            x = .FileSetNumber.ToString("000")
            x &= "-" & Trim(.FileSetName)
        End With
        Return x
    End Function
    Sub writeFileSetFiles_OLD(ByVal copy2xxxNumber As Integer, copy2mainNumber As Integer)
        'File.Copy("file-a.txt", "file-b.txt", True)
        'first write existing .fil's to xxx location
        Dim mainPath As String = "c:\TodoProgramCOM2020\"
        'donaldPath = "c:\TodoProgramCOM2020\DonaldStuartCampbell03181940Utilities\"

        Dim tempApptTodoXXX As String = UtilityPath & "ListMgt.tmp"
        Dim tempBirthdayXXX As String = UtilityPath & "Birthday.tmp"

        Dim fileFrom_copy As String = mainPath & "ListMgt.fil"
        Dim file2 As String = UtilityPath & "ListMgt." & copy2xxxNumber.ToString("000")

        'first to tmp
        Dim Y As String
        Y = UtilityPath & "ListMgt." & copy2mainNumber.ToString("000")
        File.Copy(Y, tempApptTodoXXX, True)


        'Debug.Print("fileFrom_copy-" & fileFrom_copy)
        'Debug.Print("file2-" & file2)
        'File.Copy(fileFrom_copy, file2, True)
        File.Copy(fileFrom_copy, file2, True)
        '==
        'now reverse
        Dim x As String = fileFrom_copy
        fileFrom_copy = file2
        file2 = x
        'Debug.Print("fileFrom_copy-" & fileFrom_copy)
        'Debug.Print("file2-" & file2)
        File.Copy(tempApptTodoXXX, file2, True)
        '==
        '================
        fileFrom_copy = mainPath & "Birthday.fil"
        file2 = UtilityPath & "Birthday." & copy2xxxNumber.ToString("000")

        'first to tmp
        'File.Copy(file2, tempBirthdayXXX, True)
        Y = UtilityPath & "Birthday." & copy2mainNumber.ToString("000")
        File.Copy(Y, tempBirthdayXXX, True)



        'Debug.Print("fileFrom_copy-" & fileFrom_copy)
        'Debug.Print("file2-" & file2)
        File.Copy(fileFrom_copy, file2, True)
        '========


        fileFrom_copy = mainPath & "Birthday.fil"
        file2 = UtilityPath & "Birthday." & copy2xxxNumber.ToString("000")
        '==
        x = fileFrom_copy
        fileFrom_copy = file2
        file2 = x
        'Debug.Print("fileFrom_copy-" & fileFrom_copy)
        'Debug.Print("file2-" & file2)
        'File.Copy(fileFrom_copy, file2, True)
        File.Copy(tempBirthdayXXX, file2, True)
        '==

        put_FileSetNumberInUse(copy2mainNumber)
        End
    End Sub
    Sub writeFileSetFiles(ByVal InUse As Integer, ByVal ToUse As Integer) 'inUse ToUse
        If InUse = 0 OrElse ToUse = 0 Then Exit Sub

        If InUse = ToUse Then
            MsgBox("InUse and ToUse Numbers can not be the same - will exit sub")
            Exit Sub
        End If

        'Copy Working Files to xxxInUse

        'Dim mainPath As String = "c:\TodoProgramCOM2020\"
        Dim mainPath As String = ApptGV.ccPath & "\"

        Dim Source As String = ""
        Dim Target As String = ""

        '1st copy Working Files To Vault
        Source = mainPath & "ListMgt.fil"
        Target = UtilityPath & "ListMgt." & InUse.ToString("000")
        File.Copy(Source, Target, True)
        '-
        Source = mainPath & "Birthday.fil"
        Target = UtilityPath & "Birthday." & InUse.ToString("000")
        File.Copy(Source, Target, True)

        'NOW copy new Vault to Working Files
        Source = UtilityPath & "ListMgt." & ToUse.ToString("000")
        Target = mainPath & "ListMgt.fil"
        File.Copy(Source, Target, True)
        '-
        Source = UtilityPath & "Birthday." & ToUse.ToString("000")
        Target = mainPath & "Birthday.fil"
        File.Copy(Source, Target, True)
        '=================================================================
        put_FileSetNumberInUse(ToUse)
        '=================================================================
        'End

        '=================================================================
    End Sub
#Region "BackUpWorkingFiles"
    Sub BackUp_WorkingFiles()
        Debug.Print("Starting")
        ApptGV.ccPath = "C:\todoProgramCOM2020ListMgt"
        Dim SourcePath As String = ApptGV.ccPath
        Dim DestinationPath As String = ApptGV.ccPath & "_BackUp" 'C:\todoProgramCOM2020ListMgt_BackUp




        'Dim newDirectory As String = System.IO.Path.Combine(DestinationPath, Path.GetFileName(Path.GetDirectoryName(SourcePath)))
        ''Debug.Print(Path.GetDirectoryName(SourcePath))'C:\
        'Debug.Print(Path.GetFileName(Path.GetDirectoryName(SourcePath)))
        'Debug.Print(newDirectory)
        'If Not (Directory.Exists(newDirectory)) Then
        '    Debug.Print("creating new direcory")
        '    'Directory.CreateDirectory(newDirectory)
        'End If

        'If Not (Directory.Exists(DestinationPath)) Then
        '    Directory.CreateDirectory(DestinationPath)
        'End If
        System.IO.Directory.Delete(DestinationPath, True)
        Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(SourcePath, DestinationPath, True) 'adding True causes existing files to be overwritten 
        Debug.Print("Done!")

    End Sub
    Sub BackUp_Permanent_DonaldStuartCampbellPersonalFolder()
        'C:\Donald Stuart Campbell Personal Folder\BackupListMgt
        Debug.Print("Starting")
        'ApptGV.ccPath = "C:\todoProgramCOM2020ListMgt"
        Dim SourcePath As String = "C:\todoProgramCOM2020ListMgt"
        Dim DestinationPath As String = "C:\Donald Stuart Campbell Personal Folder\BackupListMgt"




        System.IO.Directory.Delete(DestinationPath, True)
        Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(SourcePath, DestinationPath, True) 'adding True causes existing files to be overwritten 
        Debug.Print("Done!")


    End Sub
    Sub Restore_WorkingFiles()
        Debug.Print("Starting")
        ApptGV.ccPath = "C:\todoProgramCOM2020ListMgt"
        'Dim SourcePath As String = ApptGV.ccPath
        'Dim DestinationPath As String = ApptGV.ccPath & "_BackUp" 'C:\todoProgramCOM2020ListMgt_BackUp
        Dim SourcePath As String = ApptGV.ccPath & "_BackUp"
        Dim DestinationPath As String = ApptGV.ccPath

        System.IO.Directory.Delete(DestinationPath, True)

        'If Not (Directory.Exists(DestinationPath)) Then
        '    Directory.CreateDirectory(DestinationPath)
        'End If
        Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(SourcePath, DestinationPath, True) 'adding True causes existing files to be overwritten 
        Debug.Print("Done!")
    End Sub
    Sub EraseAllFiles()
        Dim SourcePath As String = ApptGV.ccPath


        System.IO.Directory.Delete(SourcePath, True)

        Debug.Print("Done!")
        End

    End Sub
    Sub compare_filesIn2Directories()
        ApptGV.ccPath = "C:\todoProgramCOM2020ListMgt"
        Dim SourcePath As String = ApptGV.ccPath
        Dim DestinationPath As String = ApptGV.ccPath & "_BackUp" 'C:\todoProgramCOM2020ListMgt_BackUp
        'Dim SourcePath As String = ApptGV.ccPath & "_BackUp"
        'Dim DestinationPath As String = ApptGV.ccPath


        Dim PathA As String = SourcePath 'path for folder A
        Dim PathB As String = DestinationPath 'path for folder B
        Dim Dir1 As New System.IO.DirectoryInfo(PathA)
        Dim Dir2 As New System.IO.DirectoryInfo(PathB)
        Dim List1 = Dir1.GetFiles("*.*")
        Dim List2 = Dir2.GetFiles("*.*")

        'For Each file In List1
        '    Console.WriteLine(file.Name)
        'Next
        'Console.WriteLine("-----------")
        'For Each file In List2
        '    Console.WriteLine(file.Name)
        'Next

        Dim List3 = List1.Intersect(List2, New FileComparer)

        For Each info In List3
            'Console.WriteLine(info.Name)
            Debug.Print(info.Name)
        Next
    End Sub
    Public Class FileComparer
        Implements IEqualityComparer(Of FileInfo)

        Public Function Equals1(ByVal x As FileInfo, ByVal y As FileInfo) As Boolean _
                        Implements IEqualityComparer(Of FileInfo).Equals


            If x Is y Then Return True
            If x Is Nothing OrElse y Is Nothing Then Return False

            Return (x.Name = y.Name) AndAlso (x.Length = y.Length) AndAlso (x.LastWriteTime = y.LastWriteTime)
        End Function

        Public Function GetHashCode1(ByVal info As FileInfo) As Integer _
                         Implements IEqualityComparer(Of FileInfo).GetHashCode

            If info Is Nothing Then Return 0
            Return info.Name.GetHashCode()
        End Function
    End Class

    Sub Test_CheckFiles()
        Dim file1 As String = ApptGV.ccPath & "\ListMgt.fil"
        Dim file2 As String
        'file2 = "C:\Donald Stuart Campbell Personal Folder\BackupListMgt\listmgt.fil"
        'file2 = "C:\todoProgramCOM2020ListMgt\Utilities\listmgt.001"
        file2 = "C:\todoProgramCOM2020ListMgt_BackUp\Utilities\listmgt.001"
        'CheckFiles(file1, file2)

        If (FileCompare(file1, file2)) Then
            MessageBox.Show("Files are equal.")
        Else
            MessageBox.Show("Files are not equal.")
        End If
    End Sub
    Private Function FileCompare(ByVal file1 As String, ByVal file2 As String) As Boolean
        Dim file1byte As Integer
        Dim file2byte As Integer
        Dim fs1 As FileStream
        Dim fs2 As FileStream

        If (file1 = file2) Then
            Return True
        End If

        fs1 = New FileStream(file1, FileMode.Open)
        fs2 = New FileStream(file2, FileMode.Open)

        If (fs1.Length <> fs2.Length) Then
            fs1.Close()
            fs2.Close()
            Return False
        End If

        Do
            file1byte = fs1.ReadByte()
            file2byte = fs2.ReadByte()
        Loop While ((file1byte = file2byte) And (file1byte <> -1))

        fs1.Close()
        fs2.Close()

        Return ((file1byte - file2byte) = 0)
    End Function


#End Region
#Region "read00Xfiles"
    Function GetActiveTodoFrom00Xfiles(ByVal XXXfileNumber As Integer)
        'from Set_ApptTodoRecordType_UsingReader() ' As ApptTodoRecordType()
        Dim TodoRecs() As ApptTodoRecordType

        ''''Dim fileNameToReadFrom As String = UtilityPath & XXXfile.ToString("000").fil
        'Dim FullFileName As String = $"{UtilityPath}ListMgt.{XXXfileNumber.ToString("000")}"
        Dim FullFileName As String = UtilityPath & "ListMgt." & XXXfileNumber.ToString("000")

        Dim cnt As Integer = 0
        Dim i As Integer

        'new
        'Dim distinctDates(0) As Long

        'Dim Recs() As ApptTodoRecordType

        Dim NumRecs As Integer
        Dim cntTodo As Integer = 0

        ''''Dim FullFileName As String = xBuildFullFileName(file_ListMgt.FileName)
        'Dim FullFileName As String = xBuildFullFileName(file_ListMgt640.FileName)

        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
            Dim length As Integer = reader.BaseStream.Length
            NumRecs = length / file_ListMgt.FileLen
            'NumRecs = length / file_ListMgt640.FileLen

            ''''ReDim ApptGV.TodoCreateDate(NumRecs)

            'ReDim Recs(NumRecs)
            ''''ReDim ApptGV.All_ApptTodoRecords(NumRecs) 'ApptGV.AllAppts_LongInteger(NumRecs)
            ReDim TodoRecs(NumRecs)
            'ReDim ApptGV.AllAppts_LongInteger(NumRecs)

            'new
            ''''ReDim ApptGV.distinctDates(NumRecs)
            'startTime()

            For i = 1 To NumRecs
                With TodoRecs(i) 'ApptGV.All_ApptTodoRecords(i)
                    .ID = reader.ReadInt32
                    .dTics = reader.ReadInt64

                    '.Msg = reader.ReadChars(84)
                    '/ListMgt/
                    '.Msg = reader.ReadChars(237) '557
                    .Msg = reader.ReadChars(557)

                    Dim msg As String
                    msg = TrimF(.Msg)
                    ' Debug.Print(Len(msg) & vbNewLine & msg)


                    .AcctNum = reader.ReadInt32
                    .StrikeOut = reader.ReadInt32
                    .DeleteFlag = reader.ReadInt32
                    .DeleteAbsolute = reader.ReadInt32
                    .scUserID = reader.ReadInt32
                    .ApptDateStr = reader.ReadChars(27)



                    '.apptDate = (.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay 'xDate 'CDate(Mid(.apptDate, 1, 10)) 'CDate(.apptDate) 'Date.Today.Date '.apptDate
                    .apptDate = reader.ReadInt64

                    'new
                    'ApptGV.distinctDates(i) = (.dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '.apptDate

                    .dTicsOriginal = reader.ReadInt64
                    .dTicsCompleted = reader.ReadInt64


                    If .dTicsOriginal > ApptGV.Highest_dTicsOriginal Then ApptGV.Highest_dTicsOriginal = .dTicsOriginal
                    '===
                    'If .Msg = "test adding todo for noDate //" Then
                    'If InStr(.Msg, "test adding todo for noDate //") > 0 Then
                    'Dim xMsg As String = Replace(TrimF(.Msg), vbNewLine, "/\")
                    'Debug.Print(Trim(.Msg) & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & New Date(.dTicsOriginal))
                    'Debug.Print(Trim(.ID & " " & .dTics & " " & xMsg & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & New Date(.dTicsOriginal)))
                    '    i = i
                    'End If
                    '===

                    '.dTics = reader.ReadInt64
                    '=========================
                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

                    'Debug.Print(.dTics)

                    'ApptGV.AllAppts_LongInteger(i).L = .dTics ' reader.ReadInt64
                    'ApptGV.AllAppts_LongInteger(i).i = .ID 'i

                    'If (.dTicsOriginal \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay = ApptGV.todoTicks Then
                    'If .apptDate = ApptGV.todoTicks Then
                    '    cntTodo += 1
                    '    ApptGV.TodoCreateDate(cntTodo).L = .dTicsOriginal
                    '    ApptGV.TodoCreateDate(cntTodo).i = .ID
                    'End If

                End With
            Next
        End Using
        cnt = 0
        For i = 1 To NumRecs
            If TodoRecs(i).DeleteFlag = 0 Then
                cnt += 1
                TodoRecs(cnt) = TodoRecs(i)
            End If
        Next
        ReDim Preserve TodoRecs(cnt)
        QuickSort_ApptTodoRecords(TodoRecs, 1, cnt)
        Return TodoRecs

        'endTime()

        ''++create new larger file
        '=========0.0957441 seconds
        'Dim ff As Integer = FreeFile()

        'Dim filename As String = xBuildFullFileName(file_ListMgt640.FileName)
        'Kill(filename)
        'createFileTDP(FullFileName)
        'FileOpen(ff, filename, OpenMode.Random, , , file_ListMgt640.FileLen)
        ''==
        'cnt = 0
        ''Dim j As Integer

        'For i = 1 To NumRecs
        '    cnt += 1
        '    ApptGV.All_ApptTodoRecords(i).Msg = TrimF(ApptGV.All_ApptTodoRecords(i).Msg)
        '    FilePut(ff, ApptGV.All_ApptTodoRecords(i), cnt) 'posInFile)
        'Next


        'FileClose(ff)
        'Debug.Print("done!")
        'End
        '=========
        'FullFileName = xBuildFullFileName(file_ListMgt640.FileName)
        'createFileTDP(FullFileName)
        'Using writer As New BinaryWriter(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
        '    'Dim length As Integer = reader.BaseStream.Length
        '    'NumRecs = length / file_ListMgt640.FileLen

        '    For i = 1 To NumRecs
        '        With ApptGV.All_ApptTodoRecords(i)
        '            writer.Write(.ID)
        '            writer.Write(.dTics)
        '            writer.Write(.Msg)
        '            writer.Write(.AcctNum)

        '            writer.Write(.StrikeOut)
        '            writer.Write(.DeleteFlag)

        '            writer.Write(.DeleteAbsolute)
        '            writer.Write(.scUserID)
        '            writer.Write(.ApptDateStr)
        '            writer.Write(.apptDate)
        '            writer.Write(.dTicsOriginal)
        '            writer.Write(.dTicsCompleted)

        '            '.ID = reader.ReadInt32
        '            '.dTics = reader.ReadInt64

        '            ''.Msg = reader.ReadChars(84)
        '            ''/ListMgt/
        '            '.Msg = reader.ReadChars(237) '557

        '            '.AcctNum = reader.ReadInt32
        '            '.StrikeOut = reader.ReadInt32
        '            '.DeleteFlag = reader.ReadInt32
        '            '.DeleteAbsolute = reader.ReadInt32
        '            '.scUserID = reader.ReadInt32
        '            '.ApptDateStr = reader.ReadChars(27)



        '            ''.apptDate = (.dTics \ TimeSpan.TicksPerDay) * TimeSpan.TicksPerDay 'xDate 'CDate(Mid(.apptDate, 1, 10)) 'CDate(.apptDate) 'Date.Today.Date '.apptDate
        '            '.apptDate = reader.ReadInt64

        '            ''new
        '            ''ApptGV.distinctDates(i) = (.dTics \ TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute '.apptDate

        '            '.dTicsOriginal = reader.ReadInt64
        '            '.dTicsCompleted = reader.ReadInt64
        '        End With
        '    Next

        'End Using
        'End
        ''================
        'Fix===
        'For i = 1 To NumRecs
        '    With ApptGV.All_ApptTodoRecords(i)
        '        If (.dTics \ TimeSpan.TicksPerDay) = 0 Then

        '            .dTics = ApptGV.todoTicks + i
        '            .ApptDateStr = formatApptTodoDateStr(.dTics)
        '            .apptDate = createPureDate_Ticks_FromTicks(.dTics)

        '            Dim xMsg As String = Replace(TrimF(.Msg), vbNewLine, "/\")

        '            'Debug.Print(.ID & " " & Trim(.dTics & " " & xMsg & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & New Date(.dTicsOriginal)))
        '            putApptTodoRecord(ApptGV.All_ApptTodoRecords(i))
        '        End If

        '    End With
        'Next
        'Fix===End

        ''''ApptGV.All_ApptTodoRecords_Cnt = NumRecs
        'For i = ApptGV.All_ApptTodoRecords_Cnt To ApptGV.All_ApptTodoRecords_Cnt - 8 Step -1
        '    With ApptGV.All_ApptTodoRecords(i)
        '        'Debug.Print(Trim(.Msg) & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & New Date(.dTicsOriginal))
        '        Debug.Print(.ID & " " & i & " " & .ApptDateStr & " " & .DeleteFlag & " " & .dTics & " " & New Date(.dTicsOriginal) & " :::" & Trim(.Msg))
        '    End With

        'Next
        ''''QuickSort_ApptTodoRecords(ApptGV.All_ApptTodoRecords, 1, ApptGV.All_ApptTodoRecords_Cnt) '0.0408907 seconds (This sort is Twice as fast as the Array.sort lambda function below))


    End Function
    Function GetActiveTodoFrom00XfilesNEW(ByVal XXXfileNumber As Integer)
        Dim TodoRecs() As ApptTodoRecordType

        Dim FullFileName As String = UtilityPath & "ListMgt." & XXXfileNumber.ToString("000")

        Dim cnt As Integer = 0
        Dim i As Integer

        Dim Rec As ApptTodoRecordType = initApptTodoRecord()

        Dim NumRecs As Integer

        'NEW 1/12/2020
        Dim InUse As Integer = get_FileSetNumberInUse()
        If InUse = XXXfileNumber Then
            TodoRecs = getDeletedTodos_withinDateRange_NEW(New Date(0)) '(New Date(ApptGV.todoTicks))
            cnt = UBound(TodoRecs)
            QuickSort_ApptTodoRecords(TodoRecs, 1, cnt)
            'NumRecs = UBound(ApptGV.All_ApptTodoRecords)
            'ReDim TodoRecs(NumRecs)
            'For i = 1 To NumRecs
            '    If ApptGV.All_ApptTodoRecords(i).DeleteFlag = 0 Then
            '        cnt += 1
            '        TodoRecs(cnt) = Rec
            '    End If
            'Next
            Return TodoRecs
        End If
        '================='NEW 1/12/2020

        Using reader As New BinaryReader(File.Open(FullFileName, FileMode.Open)) ' Loop through length of file.
            Dim length As Integer = reader.BaseStream.Length
            NumRecs = length / file_ListMgt.FileLen

            ReDim TodoRecs(NumRecs)

            For i = 1 To NumRecs
                'With TodoRecs(i) 'ApptGV.All_ApptTodoRecords(i)
                With Rec 'ApptGV.All_ApptTodoRecords(i)
                    .ID = reader.ReadInt32
                    .dTics = reader.ReadInt64

                    .Msg = reader.ReadChars(557)

                    'Dim msg As String
                    'msg = TrimF(.Msg)
                    '' Debug.Print(Len(msg) & vbNewLine & msg)


                    .AcctNum = reader.ReadInt32
                    .StrikeOut = reader.ReadInt32
                    .DeleteFlag = reader.ReadInt32
                    .DeleteAbsolute = reader.ReadInt32
                    .scUserID = reader.ReadInt32
                    .ApptDateStr = reader.ReadChars(27)

                    .apptDate = reader.ReadInt64

                    .dTicsOriginal = reader.ReadInt64
                    .dTicsCompleted = reader.ReadInt64

                End With
                If Rec.DeleteFlag = 0 Then
                    'If XXXfileNumber = 1 Then
                    '    Debug.Print(TrimF(Rec.Msg) & " del=" & Rec.DeleteFlag)
                    'End If
                    cnt += 1
                    TodoRecs(cnt) = Rec
                End If
            Next

        End Using
        ReDim Preserve TodoRecs(cnt)
        QuickSort_ApptTodoRecords(TodoRecs, 1, cnt)
        Return TodoRecs

    End Function
    Function get_3TodosFromEachList(ByRef FileSetArray() As Integer, Optional ByVal NumTodosToShow As Integer = 3) As TwoStringsAndOneIntegerType()
        Dim FileSets() As FileSetType = get_FileSets_UsingReader() ' As FileSetType()
        Dim nFileSets As Integer = UBound(FileSets)
        Dim i As Integer
        Dim j As Integer
        Dim nRecs As Integer = 3 * nFileSets
        Dim recs(nRecs) As TwoStringsAndOneIntegerType
        Dim todoRecs() As ApptTodoRecordType
        Dim cnt As Integer = 0
        If NumTodosToShow = 0 Then
            nRecs = 1000000
            ReDim recs(nRecs)

            For i = 1 To nFileSets
                If FileSetArray(i) = 1 Then
                    todoRecs = GetActiveTodoFrom00XfilesNEW(FileSets(i).FileSetNumber)
                    For j = 1 To UBound(todoRecs)
                        cnt += 1
                        With recs(cnt)
                            .i1 = j
                            .s1 = FileSets(i).FileSetName
                            .s2 = todoRecs(j).Msg
                        End With
                    Next
                End If
            Next
        Else
            For i = 1 To nFileSets
                If FileSetArray(i) = 1 Then
                    todoRecs = GetActiveTodoFrom00XfilesNEW(FileSets(i).FileSetNumber)
                    For j = 1 To Math.Min(NumTodosToShow, UBound(todoRecs))
                        cnt += 1
                        With recs(cnt)
                            .i1 = j
                            .s1 = FileSets(i).FileSetName
                            .s2 = todoRecs(j).Msg
                        End With
                    Next
                End If
            Next
        End If
        ReDim Preserve recs(cnt)
        Return recs
    End Function

#End Region'read00Xfiles

End Module
