Imports System.IO
Public Class Form1

    '// GTA IMG Viewer Source for VB.NET
    '// by Idabgsram
    '// Based from gta126705@hotmail.com vb6 source with bunch of adjustment

#Region "Variable Declarations"
    Private lngCurrentOffset As Long
    Private imgFileName As String
    Private imgDirPath As String
    Private imgListPath As String
    Private imgFileLocation As String
    Private fileTypeTemp As String = ""
    Private fileNameTemp As String = ""
    Private fileSizeTemp As String = ""
    Private fileOffsetTemp As String = ""
    Private DirFileOpened As Boolean = False
#End Region

#Region "Function"
    Private Function ReverseByte(strData As String) As String
        Dim strTemp() As String
        Dim strReturn As String
        Dim i As Long
        Dim y As Long

        ReDim strTemp(Len(strData) / 2)

        y = 1
        strReturn = ""

        For i = 0 To UBound(strTemp) - 1
            strTemp(i) = Mid(strData, y, 2)
            y = y + 2
        Next i

        For i = (UBound(strTemp) - 1) To 0 Step -1
            strReturn = strReturn & strTemp(i)
        Next i

        ReverseByte = strReturn
    End Function
    Private Function ReadByteHex(strData As String) As String
        Dim i As Long
        Dim strTemp As String

        strTemp = ""

        For i = 1 To Len(strData)
            If Len(Hex(Asc(Mid(strData, i, 1)))) = 1 Then
                strTemp = strTemp & "0" & Hex(Asc(Mid(strData, i, 1)))
            Else
                strTemp = strTemp & Hex(Asc(Mid(strData, i, 1)))
            End If
        Next i

        strTemp = ReverseByte(strTemp)

        ReadByteHex = strTemp
    End Function
    Private Function GetValue(lngOffset As Long, lngSizeInByte As Integer, Optional intOption As Integer = 2) As String
        Dim strTemp As String
        strTemp = StrDup(lngSizeInByte, Str(&H0))
        FileGet(1, strTemp, lngOffset + 1)

        Select Case intOption
            Case 0
                GetValue = strTemp
            Case 1
                GetValue = ReadByteHex(strTemp)
            Case 2
                GetValue = CDbl("&H" & ReadByteHex(strTemp))
            Case 3
                If ReadByteHex(strTemp) = True Then GetValue = "True" Else GetValue = "False"
            Case Else
                GetValue = "Unknown"
        End Select
    End Function
    Private Function GetIMGVersion() As Integer
        Dim i As Integer 'counter

        FileOpen(1, imgDirPath & "\" & imgFileName, OpenMode.Binary, OpenAccess.Read, OpenShare.LockRead)
        lngCurrentOffset = 0

        If GetValue(lngCurrentOffset, 4, 0) = "VER2" Then
            GetIMGVersion = 2
            FileClose(1)
            Exit Function
        End If
        FileClose(1)

        i = 0

        imgListPath = imgDirPath & "\gta3.dir"
        If DirFileOpened = True Then imgFileLocation = imgDirPath & "\gta3.img"

        If File.Exists(imgListPath) = True Then
            GetIMGVersion = 1
            Exit Function
        End If

        GetIMGVersion = -1
    End Function
    Private Sub ReadIMGVersion2()
        lngCurrentOffset = 0
        ToolStripStatusLabel2.Visible = False
        ToolStripStatusLabel1.Text = "Reading v2 Archive"
        lngCurrentOffset = lngCurrentOffset + &H8
        ToolStripProgressBar1.Value = 0
        ListView1.BeginUpdate()
        AddToList(GetValue(lngCurrentOffset + &H4, 4), lngCurrentOffset, True)
        ListView1.EndUpdate()
        ToolStripStatusLabel2.Visible = True
        ToolStripStatusLabel2.Text = "IMG 2 - GTA SA"
    End Sub
    Private Sub ReadIMGVersion1()
        lngCurrentOffset = 0
        ToolStripStatusLabel2.Visible = False
        ToolStripStatusLabel1.Text = "Reading v1 Archive"
        ToolStripProgressBar1.Value = 0
        ListView1.BeginUpdate()
        AddToList(FileLen(imgListPath) / 32, lngCurrentOffset, True)
        ListView1.EndUpdate()
        ToolStripStatusLabel2.Visible = True
        ToolStripStatusLabel2.Text = "IMG 1 - GTA 3 / VC"
    End Sub
    Private Sub AddToList(intFileNumber As Integer, offsetNumber As String, isVersionOne As Boolean)
        On Error Resume Next
        ToolStripProgressBar1.Visible = True
        Dim i As Integer
        Select Case isVersionOne
            Case True
                For i = 0 To (intFileNumber - 1)
                    fileNameTemp = GetValue(lngCurrentOffset + &H8, 24, 0)
                    fileTypeTemp = Strings.Right(fileTypeTemp, 3)
                    fileSizeTemp = ((CDbl(GetValue(lngCurrentOffset + &H4, 4)) * 2048) / 1024) & " kb"
                    offsetNumber = Hex(CDbl(GetValue(lngCurrentOffset, 4)) * 2048)
                    Select Case offsetNumber.Length
                        Case 1
                            fileOffsetTemp = StrDup(7, "0") & offsetNumber
                        Case 2
                            fileOffsetTemp = StrDup(6, "0") & offsetNumber
                        Case 3
                            fileOffsetTemp = StrDup(5, "0") & offsetNumber
                        Case 4
                            fileOffsetTemp = StrDup(4, "0") & offsetNumber
                        Case 5
                            fileOffsetTemp = StrDup(3, "0") & offsetNumber
                        Case 6
                            fileOffsetTemp = StrDup(2, "0") & offsetNumber
                        Case 7
                            fileOffsetTemp = StrDup(1, "0") & offsetNumber
                        Case 8
                            fileOffsetTemp = offsetNumber
                        Case Else
                            fileOffsetTemp = offsetNumber
                    End Select

                    With ListView1.Items.Add(fileNameTemp)
                        .SubItems.Add(fileSizeTemp)
                        .SubItems.Add(fileOffsetTemp)
                        lngCurrentOffset = lngCurrentOffset + 32
                    End With
                    Application.DoEvents()
                    ToolStripStatusLabel1.Text = "Reading file(s)..."
                    ToolStripProgressBar1.Value = 100 / (intFileNumber) * (i + 1)
                Next i
            Case False
                For i = 0 To (intFileNumber - 1)
                    fileNameTemp = GetValue(lngCurrentOffset + &H8, 24, 0)
                    fileTypeTemp = Strings.Right(fileTypeTemp, 3)
                    fileSizeTemp = ((CDbl(GetValue(lngCurrentOffset + &H4, 4)) * 2048) / 1024) & " kb"
                    fileOffsetTemp = Hex(CDbl(GetValue(lngCurrentOffset, 4)) * 2048)
                    With ListView1.Items.Add(fileNameTemp)
                        .SubItems.Add(fileSizeTemp)
                        .SubItems.Add(fileOffsetTemp)
                        lngCurrentOffset = lngCurrentOffset + 32
                    End With
                    Application.DoEvents()
                    ToolStripStatusLabel1.Text = "Reading file(s)..."
                    ToolStripProgressBar1.Value = 100 / (intFileNumber) * (i + 1)
                Next i
        End Select
        If intFileNumber <= 1 Then
            ToolStripStatusLabel1.Text = "Loaded " & intFileNumber & " file"
        Else
            ToolStripStatusLabel1.Text = "Loaded " & intFileNumber & " files"
        End If
        ToolStripProgressBar1.Visible = False
    End Sub
    Private Sub OpenIMG(imgFilePath As String)
        DirFileOpened = False
        imgDirPath = Path.GetDirectoryName(imgFilePath)
        Select Case Strings.Right(imgFilePath, 3).ToLower
            Case "dir"
                If File.Exists(imgDirPath & "\gta3.img") = True Then
                    DirFileOpened = True

                Else
                    MsgBox("gta3.img not found, cannot process gta3.dir", vbExclamation, "Error")
                    Exit Sub
                End If
            Case Else : DirFileOpened = False
        End Select
        imgFileName = Path.GetFileName(imgFilePath)

        ListView1.Items.Clear()

        Select Case GetIMGVersion()
            Case 2 'GTA SA
                FileOpen(1, imgFilePath, OpenMode.Binary, OpenAccess.Read)
                ReadIMGVersion2()
                FileClose(1)
            Case 1 'GTA 3 and VC
                FileOpen(1, imgListPath, OpenMode.Binary, OpenAccess.Read)
                ReadIMGVersion1()
                FileClose(1)
            Case Else
                MsgBox("Not a valid GTA-IMG file", vbExclamation, "Error")
        End Select
    End Sub
#End Region

#Region "Button"
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        MsgBox("GTA IMG Viewer" & vbCrLf & Application.ProductVersion & vbCrLf & vbCrLf & "by Idabgsram")
    End Sub
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Select Case OpenFileDialog1.ShowDialog()
            Case DialogResult.Cancel
                Exit Sub
            Case DialogResult.No
                Exit Sub
        End Select
        OpenIMG(OpenFileDialog1.FileName.ToString)
    End Sub
#End Region
End Class
