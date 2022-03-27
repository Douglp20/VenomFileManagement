Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Public Class FileManagement

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
    End Sub
    Public Enum FileInfo
        FullName
        Name
        Extension
        Length
        DirectoryName
    End Enum
    Public Function WriteProperty(ByRef strValue As String, strfile As String)

        On Error GoTo Err

        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream

        ts = fso.CreateTextFile(strfile)
        ts.WriteLine(strValue)
        ts.Close()


        Exit Function

Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function

    Public Function ReadProperty(ByRef strValue As String, strfile As String) As String

        On Error GoTo Err



        Dim curRow As String
        Dim curLine As String
        Dim curSplit As Array
        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream

        If fso.FileExists(strfile) Then


            ts = fso.OpenTextFile(strfile)

            Do While Not ts.AtEndOfLine
                curLine = ts.ReadLine
                curSplit = curLine.Split(":")
                If curSplit(0).ToString = strValue Then
                    ReadProperty = curSplit(1).ToString
                    Exit Do
                End If

            Loop

        End If

        Exit Function
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

    Private Function DeleteProperty(ByRef strValue As String, strfile As String)

        On Error GoTo Err

        Dim curLineNo As Long
        Dim lctr As Long
        Dim curSplit As Array
        Dim curLine As String
        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim bolCur As Boolean = False
        Dim strTemp As String = ""

        'get the line no of the value
        ts = fso.OpenTextFile(strfile)
        lctr = 1
        Do While Not ts.AtEndOfStream
            curLine = ts.ReadLine
            curSplit = curLine.Split(":")
            If curSplit(0).ToString = strValue Then
                bolCur = True
                curLineNo = lctr
                Exit Do
            End If
            lctr = lctr = 1
        Loop
        ts.Close()

        'Rebuilding the text without the line
        If bolCur Then
            ts = fso.OpenTextFile(strfile)
            curLineNo = 1
            Do While Not ts.AtEndOfStream
                curLine = ts.ReadLine
                If curLineNo <> lctr Then
                    strTemp = strTemp & curLine & vbCrLf
                End If
            Loop
            ts.Close()
            ts.Write(strfile)
        End If

        Exit Function

Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
    Public Function ImportCSVItem(lvw As System.Windows.Forms.ListView, ByRef FilePathName As String, NoOfColumns As Integer, withHeaders As Boolean)
        On Error GoTo Err

        Dim curLine As String
        Dim curSplit As Array
        Dim curSplitColNo As Integer = 0
        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim row As Integer = 0
        Dim StartRow As Integer

        Dim lViewItem As System.Windows.Forms.ListViewItem
        If withHeaders = True Then
            StartRow = 1
        Else
            StartRow = 0

        End If
        row = 0

        If fso.FileExists(FilePathName) Then

            ts = fso.OpenTextFile(FilePathName)

            Do While Not ts.AtEndOfStream
                curLine = ts.ReadLine
                ' If Len(curLine >= 0) Then End
                curSplit = curLine.Split(",")
                If curSplit.Length > 0 Then
                    If row >= StartRow Then
                        lViewItem = New System.Windows.Forms.ListViewItem("")
                        For c As Integer = 0 To NoOfColumns - 1
                            lViewItem.SubItems.Add(curSplit(c).ToString)
                        Next
                        lvw.Items.Add(lViewItem)
                    End If
                End If

                row = row + 1
            Loop




        End If





        Exit Function

Err:

        Dim rtn As String = "The error occur within the module in line " + curLine + " : " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Sub SaveByteToFile(ByVal filePath As String, ByVal Image As Byte())
        On Error GoTo Err


        System.IO.File.WriteAllBytes(filePath, Image)

        Exit Sub

Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub DeleteAllFilesInFolder(ByVal FolderPath As String)
        On Error GoTo Err

        For Each deleteFile In System.IO.Directory.GetFiles(FolderPath, "*.*", SearchOption.TopDirectoryOnly)
            System.IO.File.Delete(deleteFile)
        Next


        Exit Sub

Err:

            Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
            RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub DeleteSingleFileInFolder(ByVal filePath As String)
        On Error GoTo Err


        System.IO.File.Delete(filePath)

        Exit Sub

Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub


    Public Function GetStreamByteFromFile(filePath As String, ext As String) As Byte()


        Dim ReturnfileBytes
        On Error GoTo Err
        Dim shareOption As System.IO.FileShare
        Dim fs = New System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, shareOption)


        If ext = ".pdf" Then
            Dim numBytes As Long = fs.Length
            Dim br As New BinaryReader(fs)
            Dim NewByte As Byte() = br.ReadBytes(CInt(numBytes))
            br.Close()
            ReturnfileBytes = NewByte
        Else
            Dim fileBytes(CInt(fs.Length - 1)) As Byte
            fs.Read(fileBytes, 0, fileBytes.Length)
            ReturnfileBytes = fileBytes
        End If



        Return ReturnfileBytes




        Exit Function
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function GetStreamByteFromFile(filePath As String) As Byte()

        Dim valueArray As New ArrayList
        Dim ReturnfileBytes
        On Error GoTo Err
        Dim shareOption As System.IO.FileShare
        Dim fs = New System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, shareOption)

        valueArray = GetFileInfo(filePath)

        If valueArray(modFile.FileInfo.Extension) = ".pdf" Then
            Dim numBytes As Long = fs.Length
            Dim br As New BinaryReader(fs)
            Dim NewByte As Byte() = br.ReadBytes(CInt(numBytes))
            br.Close()
            ReturnfileBytes = NewByte
        Else
            Dim fileBytes(CInt(fs.Length - 1)) As Byte
            fs.Read(fileBytes, 0, fileBytes.Length)
            ReturnfileBytes = fileBytes
        End If



        Return ReturnfileBytes




        Exit Function
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function GetFileInfo(ByVal File As String) As ArrayList


        Dim ReturnValue As New ArrayList
        On Error GoTo Err

        Dim fi As New IO.FileInfo(File)
        ReturnValue.Add(fi.FullName)
        ReturnValue.Add(fi.Name)
        ReturnValue.Add(fi.Extension)
        ReturnValue.Add(fi.Length)
        ReturnValue.Add(fi.DirectoryName)


        GetFileInfo = ReturnValue


        Exit Function
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function GetFile(Title As String, Filter As String, InitialDirectory As String) As String


        Dim dialog As New OpenFileDialog()
        Dim ReturnValue As String = String.Empty

        On Error GoTo Err


        dialog.Filter = Filter
        dialog.InitialDirectory = InitialDirectory
        dialog.Title = Title

        If dialog.ShowDialog() = DialogResult.OK Then

            ReturnValue = dialog.FileName

        End If

        Return ReturnValue

        Exit Function
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

    Public Function GetFolder(Title As String, Filter As String, InitialDirectory As String) As String


        Dim dialog As New OpenFileDialog()
        Dim ReturnValue As String = String.Empty

        On Error GoTo Err


        dialog.Filter = Filter
        dialog.InitialDirectory = InitialDirectory
        dialog.Title = Title

        If dialog.ShowDialog() = DialogResult.OK Then

            Dim fileName As String = dialog.SafeFileName
            Dim filePath As String = dialog.FileName
            Dim fileFolder As String = filePath.Replace(dialog.SafeFileName, "")
            ReturnValue = fileFolder
        End If

        Return ReturnValue

        Exit Function
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

    Public Function GetFilesFromFolder(pathName As String) As ArrayList

        On Error GoTo Err

        Dim arr As New ArrayList

        For Each file As String In System.IO.Directory.GetFiles(pathName)
            Dim information = New System.IO.FileInfo(file)
            arr.Add(information)
        Next

        Return arr

        Exit Function
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Sub OpenExplore(PathName As String)

        On Error GoTo Err

        Process.Start("explorer.exe", String.Format("/n, /e, {0}", pathName))

        Exit Sub

Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
    Public Function CheckFolderExists(FolderName As String) As Boolean

        Dim fso As New Scripting.FileSystemObject
        On Error GoTo Err

        If fso.FolderExists(FolderName) Then
            Return True
        Else
            Return False
        End If

        Exit Function

Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function CreateNewFolder(NewFolder As String) As Boolean

        Dim fso As New Scripting.FileSystemObject

        Dim creationOK As Boolean = False
        On Error GoTo Err


        System.IO.Directory.CreateDirectory(NewFolder)
        creationOK = CheckFolderExists(NewFolder)
        If creationOK Then
            Dim fs As FileStream = File.Create(NewFolder & "\Default.text")
            Dim info As Byte() = New UTF8Encoding(True).GetBytes("This is a default text. Not to be use")
            fs.Write(info, 0, info.Length)
            fs.Close()
        End If


        Return creationOK



        Exit Function

Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

End Class
