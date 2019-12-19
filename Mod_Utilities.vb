Option Explicit On
Imports System.IO
Module Mod_Utilities
    Dim donaldPath As String = "c:\TodoProgramCOM2020\DonaldStuartCampbell03181940Utilities\"
    Dim UtilityFileSet_Name As String = "UtilityFileSet.fil"
    Dim LenOfUtilityFileSet_Name = 32
    Dim FileSetNumberInUse_FileName As String = "FileSetNumberInUse.fil"
    Structure FileSetType
        <VBFixedString(28)> Dim FileSetName As String
        Dim FileSetNumber As Integer
    End Structure
    Sub add_ApptTodoXXXAndBirthdayXXX(ByVal xNumber As Integer)
        Dim FullFileName As String = donaldPath & "ApptTodo." & xNumber.ToString("000")

        If System.IO.File.Exists(FullFileName) Then
        Else
            System.IO.File.Create(FullFileName).Dispose()
        End If

        FullFileName = donaldPath & "Birthday." & xNumber.ToString("000")
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
    Function setFullUtilityFileSetName() As String
        Return donaldPath & UtilityFileSet_Name
    End Function
    Function setFullFileNameUsingDonaldPath(ByVal xFileName As String) As String
        Return donaldPath & xFileName
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
        Dim FullFileName As String = setFullFileNameUsingDonaldPath(FileSetNumberInUse_FileName)
        If System.IO.File.Exists(FullFileName) Then
        Else
            System.IO.File.Create(FullFileName).Dispose()
        End If
        '==
        Dim ff As Integer = FreeFile()
        FileOpen(ff, FullFileName, OpenMode.Random, , , 4)

        FilePut(ff, xNumber, 1)
        FileClose(ff)
        '==

    End Sub
    Function get_FileSetNumberInUse() As Integer
        Dim FullFileName As String = setFullFileNameUsingDonaldPath(FileSetNumberInUse_FileName)
        If System.IO.File.Exists(FullFileName) Then
        Else
            System.IO.File.Create(FullFileName).Dispose()
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

        Dim tempApptTodoXXX As String = donaldPath & "ApptTodo.tmp"
        Dim tempBirthdayXXX As String = donaldPath & "Birthday.tmp"

        Dim fileFrom_copy As String = mainPath & "ApptTodo.fil"
        Dim file2 As String = donaldPath & "ApptTodo." & copy2xxxNumber.ToString("000")

        'first to tmp
        Dim Y As String
        Y = donaldPath & "ApptTodo." & copy2mainNumber.ToString("000")
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
        file2 = donaldPath & "Birthday." & copy2xxxNumber.ToString("000")

        'first to tmp
        'File.Copy(file2, tempBirthdayXXX, True)
        Y = donaldPath & "Birthday." & copy2mainNumber.ToString("000")
        File.Copy(Y, tempBirthdayXXX, True)



        'Debug.Print("fileFrom_copy-" & fileFrom_copy)
        'Debug.Print("file2-" & file2)
        File.Copy(fileFrom_copy, file2, True)
        '========


        fileFrom_copy = mainPath & "Birthday.fil"
        file2 = donaldPath & "Birthday." & copy2xxxNumber.ToString("000")
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
        If InUse = ToUse Then
            MsgBox("InUse and ToUse Numbers can not be the same - will exit sub")
            Exit Sub
        End If

        'Copy Working Files to xxxInUse

        Dim mainPath As String = "c:\TodoProgramCOM2020\"

        Dim Source As String = ""
        Dim Target As String = ""

        '1st copy Working Files To Vault
        Source = mainPath & "ApptTodo.fil"
        Target = donaldPath & "ApptTodo." & InUse.ToString("000")
        File.Copy(Source, Target, True)
        '-
        Source = mainPath & "Birthday.fil"
        Target = donaldPath & "Birthday." & InUse.ToString("000")
        File.Copy(Source, Target, True)

        'NOW copy new Vault to Working Files
        Source = donaldPath & "ApptTodo." & ToUse.ToString("000")
        Target = mainPath & "ApptTodo.fil"
        File.Copy(Source, Target, True)
        '-
        Source = donaldPath & "Birthday." & ToUse.ToString("000")
        Target = mainPath & "Birthday.fil"
        File.Copy(Source, Target, True)
        '=================================================================
        put_FileSetNumberInUse(ToUse)
        '=================================================================
        End
        '=================================================================
    End Sub

End Module
