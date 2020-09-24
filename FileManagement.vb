
Public Class FileManagement

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
    End Sub

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
    Public Function ReadProperty(ByRef strValue As String, strfile As String, strDefault As String) As String

        On Error GoTo Err



        Dim curRow As String
        Dim curLine As String
        Dim curSplit As Array
        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim ReturnValues As String = strDefault
        If fso.FileExists(strfile) Then


            ts = fso.OpenTextFile(strfile)

            Do While Not ts.AtEndOfLine
                curLine = ts.ReadLine
                curSplit = curLine.Split(":")
                If curSplit(0).ToString = strValue Then
                    ReturnValues = curSplit(1).ToString
                    Exit Do
                End If

            Loop

        End If
        ReadProperty = ReturnValues

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


End Class
