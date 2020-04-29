Imports System.IO
Module Mod_FileSystem_ListMgt
    'Note: from: Mod_FileSystem_TodoProgram in ApptTodo Program
    Sub setPath2use()
        Dim path2use As String
        path2use = "main"
        'path2use = "f"
        'path2use = "orig"
        Select Case path2use
            Case "main"

                'ApptGV.ccPath = "c:\todoProgramCOM2020"

                '/ListMgt/
                ApptGV.ccPath = "C:\todoProgramCOM2020ListMgt"
                '=========

                '1. create directory if it does not exist
                createDirectoryTDP(ApptGV.ccPath)


                Dim fullFileName As String = xBuildFullFileName(file_ListMgt.FileName)
                '2. create file if does not exist
                createFileTDP(fullFileName)

                '3. create birthdayFile id does not exist
                fullFileName = xBuildFullFileName("Birthday.fil")
                createFileTDP(fullFileName)

                'Dim rec As ApptTodoRecordType
                'Dim rec As BirthdayRecordType
                'Debug.Print(Len(rec))'130


                'Case "f"
                '    ApptGV.ccPath = "F:\todo\" '"c:\testFiles\"
                'Case "orig"
                '    ApptGV.ccPath = "c:\todo\" '"c:\testFiles\"
        End Select


    End Sub
    Sub createDirectoryTDP(ByVal DirName As String)
        If System.IO.Directory.Exists(DirName) Then
        Else
            'create One
            'Dim di As Directory = System.IO.Directory.CreateDirectory(DirName)
            System.IO.Directory.CreateDirectory(DirName)


        End If

    End Sub
    Sub createFileTDP(ByVal FullFileName As String)
        'createFile("c:\todo\","ApptTodo.fil")

        If System.IO.File.Exists(FullFileName) Then
            'MsgBox(FullFileName & " - Does exists")
        Else
            'create One
            'Dim di As Directory = System.IO.Directory.CreateDirectory(DirName)
            'System.IO.Directory.CreateDirectory(DirName)

            'If Not System.IO.File.Exists(FullFileName) Then
            System.IO.File.Create(FullFileName).Dispose()
            'End If

        End If

        '==
        'Dim x As String = "C:\todo\ApptTodo.Fil"
        'Kill(x)
        'Dim filepath As String = x
        'If Not System.IO.File.Exists(filepath) Then
        '    System.IO.File.Create(filepath).Dispose()
        'End If

    End Sub
    Function UtilityDirectoryExists() As Boolean
        Dim DirName As String = "c:\TodoProgramCOM2020\DonaldStuartCampbell03181940Utilities"
        '                                              DonaldStuartCampbell03181940Utilities
        Dim Ret As Boolean = False
        If System.IO.Directory.Exists(DirName) Then
            Ret = True
        End If
        Return Ret
    End Function
    Sub SaveTotoRecords2Disk(ByRef a() As ApptTodoRecordType)
        Dim nRecs As Integer = UBound(a)
        Dim i As Integer

        Dim ff As Integer = FreeFile()
        Dim filename As String = xBuildFullFileName(file_ListMgt.FileName)
        FileOpen(ff, filename, OpenMode.Random, , , file_ListMgt.FileLen)

        Dim xLocation As Integer
        For i = 0 To nRecs
            xLocation = a(i).ID
            FilePut(ff, a(i), xLocation)
        Next

        FileClose(ff)
    End Sub

End Module
