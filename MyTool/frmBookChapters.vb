Imports System.Data.SqlClient

Public Class frmBookChapters
    Dim strLOGINSFilePath As String
    Dim strLOGINSFileName As String
    Dim strOutputFilePath As String
    Dim ConDestSql As New SqlConnection
    Dim xFile As System.IO.File
    Dim xWrite As System.IO.StreamWriter
    Dim isEverythingOK As Boolean

    Dim intHLinkValueLength As Int16 = 0
    Dim strItemValue As String = ""
    Dim strItemChapter As String = ""
    Dim sNo As Long = 0

    Dim MyCmd As SqlCommand
    Dim MyRs As SqlDataReader
    Dim SelectQ As String = ""

    Private Sub frmBookChapters_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            txtSourceLoc.Text = "D:\Narender\Projects\ASP.NET\2005\EBooks\temp\Books"
            strLOGINSFilePath = My.Application.Info.DirectoryPath()
            strLOGINSFileName = "Chapter"

            '***GET DATABASE SETTINGS***
            Dim strDBDetailsSplit() As String = MyCLS.clsCOMMON.GetSettings()
            'txtFile.Text = strDBDetailsSplit(0)
            txtServer.Text = strDBDetailsSplit(1)
            txtUID.Text = strDBDetailsSplit(2)
            txtPassword.Text = strDBDetailsSplit(3)
            txtDatabase.Text = strDBDetailsSplit(4)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdSelectSource_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectSource.Click
        Try
            If Len(txtSourceLoc.Text) > 0 Then
                FolderBrowserDialog1.SelectedPath = txtSourceLoc.Text
            Else
                FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.Desktop
            End If

            FolderBrowserDialog1.ShowDialog(Me)
            '.xls
            'If Len(FolderBrowserDialog1.SelectedPath) > 0 And (FolderBrowserDialog1.SelectedPath <> "*.xls" Or FolderBrowserDialog1.SelectedPath <> "*.xlsx") Then
            txtSourceLoc.Text = FolderBrowserDialog1.SelectedPath.ToString
            strLOGINSFilePath = Mid(FolderBrowserDialog1.SelectedPath, 1, Len(FolderBrowserDialog1.SelectedPath) - InStr(StrReverse(FolderBrowserDialog1.SelectedPath), "\"))

            strLOGINSFileName = Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath) - InStr(StrReverse(FolderBrowserDialog1.SelectedPath), "\") + 2, InStr(StrReverse(FolderBrowserDialog1.SelectedPath), "\") - 1)
            strLOGINSFileName = Mid(strLOGINSFileName, 1, Len(strLOGINSFileName) - 4)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try
            '***SAVE DATABASE SETTINGS***
            MyCLS.clsCOMMON.SaveSettings("", txtServer.Text, txtUID.Text, txtPassword.Text, txtDatabase.Text)

            OpenConnectionDest()

            ''''''''FillListWithTablesFromDest()


        Catch ex As Exception

        End Try
    End Sub
    Private Sub OpenConnectionDest()
        Try
            Try
                ConDestSql.Close()
            Catch ex As Exception

            End Try
            'Con.ConnectionString = "Server=.;Initial Catalog=jpbrothers;uid=sa;password=sa123;" 'Integrated Security=SSPI"
            Try
                ConDestSql.ConnectionString = CreateConnString("SQL")
                ConDestSql.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Function CreateConnString(Optional ByVal isSQL As String = "") As String
        Dim strConnStr As String
        If Len(isSQL) > 0 Then
            strConnStr = "UID=" & txtUID.Text & ";Password=" & txtPassword.Text & ";Data Source=" & txtServer.Text & ";Initial Catalog=" & txtDatabase.Text & ";"
        Else
            strConnStr = "UID=" & txtUID.Text & ";Password=" & txtPassword.Text & ";Data Source=" & txtServer.Text & ";Initial Catalog=" & txtDatabase.Text & ";Provider=SQLOLEDB.1;"
        End If
        If Len(txtPassword.Text) = 0 Then
            strConnStr = strConnStr.Replace("Password=", "Integrated Security=SSPI")
        End If
        Return strConnStr
    End Function

    Function fnValidate() As Boolean
        If System.IO.Directory.Exists(txtSourceLoc.Text) = False Then
            MsgBox("Please Check Output Path!", MsgBoxStyle.Critical)
            fnValidate = False
        ElseIf ConDestSql.State <> ConnectionState.Open Then
            MsgBox("Please Connect to Database!", MsgBoxStyle.Critical)
            fnValidate = False
        Else
            fnValidate = True
        End If
    End Function

    Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click
        Try
            If fnValidate() Then
                cmdStart.Text = "Running..."
                cmdStart.Enabled = False
                'CREATE FILES FOR LOG & INSERT QUERIES FOR MSSQL
                OpenFile()
                prcRecursiveCopyFiles(txtSourceLoc.Text, txtSourceLoc.Text, True)
                CloseFile()
                cmdStart.Text = "Start"
                cmdStart.Enabled = True
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub prcRecursiveCopyFiles(ByVal sourceDir As String, ByVal destDir As String, ByVal fRecursive As Boolean)
        'On Error GoTo errHand
        Dim i As Integer
        Dim posSep As Integer
        Dim sDir As String
        Dim aDirs() As String
        Dim sFile As String
        Dim aFiles() As String
        Try

            ' Add trailing separators to the supplied paths if they don't exist.
            If Not sourceDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
                sourceDir &= System.IO.Path.DirectorySeparatorChar
            End If

            If Not destDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
                destDir &= System.IO.Path.DirectorySeparatorChar
            End If

            ' Recursive switch to continue drilling down into dir structure.
            If fRecursive Then
                If InStr(sourceDir, "Source") > 0 Then Exit Sub
                ' Get a list of directories from the current parent.
                'MsgBox(sourceDir)
                aDirs = System.IO.Directory.GetDirectories(sourceDir)

                '//PBExtract.Maximum = IIf(aDirs.GetUpperBound(0) > 0, aDirs.GetUpperBound(0), 100)

                For i = 0 To aDirs.GetUpperBound(0)
                    ' Get the position of the last separator in the current path.
                    posSep = aDirs(i).LastIndexOf("\")
                    ' Get the path of the source directory.
                    sDir = aDirs(i).Substring((posSep + 1), aDirs(i).Length - (posSep + 1))

                    'lblCPath.Text = aDirs(i)
                    'lblDirName.Text = sDir
                    lblMSG.Text = sDir
                    ' //PBExtract.Value = i

                    ' Since we are in recursive mode, copy the children also
                    prcRecursiveCopyFiles(aDirs(i), (destDir + sDir), fRecursive)
                    'MsgBox("")
                    System.Windows.Forms.Application.DoEvents()
                Next
            End If
            ' Get the files from the current parent.
            aFiles = System.IO.Directory.GetFiles(sourceDir)

            ' Copy all files.
            For i = 0 To aFiles.GetUpperBound(0)
                ' Get the position of the trailing separator.
                posSep = aFiles(i).LastIndexOf("\")
                ' Get the full path of the source file.
                sFile = aFiles(i).Substring((posSep + 1), aFiles(i).Length - (posSep + 1))

                'MOVE the file.
                Try
                    'MsgBox(InStr(destDir + aFiles(i), "Source"))
                    If InStr(destDir + aFiles(i), "Source") = 0 Then
                        If InStr(sFile, ".pdf") > 0 Then
                            'System.IO.File.Copy(aFiles(i), destDir + sFile, True)
                            'MsgBox(destDir.Split("\")(destDir.Split("\").Length - 2))
                            If InStr(destDir, "Chapter wise Pdf") > 0 And destDir.Split("\")(destDir.Split("\").Length - 2) = "Chapter wise Pdf" Then
                                strItemValue = aFiles(i).ToString()
                                intHLinkValueLength = strItemValue.IndexOf("\Books")
                                If intHLinkValueLength > 0 Then
                                    strItemValue = Mid(strItemValue, intHLinkValueLength + 8)
                                    strItemChapter = sFile
                                    strItemValue = Mid(strItemValue, 1, strItemValue.IndexOf("\"))

                                    '***FETCH THE CHAPTER SNO FROM TABLE_CHAPTER***   
                                    Try
                                        'Replace(strItemChapter, "’", "'")
                                        strItemChapter = Replace(Replace(strItemChapter, "'", "''"), ".pdf", "")
                                        If Len(strItemChapter) >= 20 Then
                                            SelectQ = "Select SNO,substring(fullName,1,20) from Table_Chapter WHERE ISBN='" & strItemValue & "' AND substring(fullName,1,20)='" & strItemChapter.Substring(0, 20) & "'"
                                        Else
                                            SelectQ = "Select SNO,fullName from Table_Chapter WHERE ISBN='" & strItemValue & "' AND fullName='" & strItemChapter & "'"
                                        End If
                                        MyCmd = New SqlCommand(SelectQ, ConDestSql)
                                        MyRs = MyCmd.ExecuteReader
                                        MyRs.Read()
                                        If MyRs.HasRows Then
                                            sNo = MyRs(0).ToString
                                            'MsgBox(destDir)
                                            'MsgBox(aFiles(i))
                                            'MsgBox("From : " & aFiles(i))
                                            'MsgBox("To : " & destDir + sNo.ToString() + "\" + sFile)
                                            System.IO.Directory.CreateDirectory(destDir + sNo.ToString())
                                            lblMSG.Text = destDir + sNo.ToString() + "\" + sFile
                                            System.IO.File.Move(aFiles(i), destDir + sNo.ToString() + "\" + sFile)
                                            'WriteFile("***MOVED : " + destDir + sNo.ToString() + "\" + sFile)
                                        Else
                                            sNo = 0
                                            WriteFile("--NOT FOUND : " + destDir + sNo.ToString() + "\" + sFile)
                                            WriteFile(SelectQ)
                                        End If
                                        MyCmd = Nothing
                                        MyRs.Close()
                                    Catch ex As Exception
                                        WriteFile("--NOT DONE : " + destDir + sNo.ToString() + "\" + sFile)
                                        WriteFile(SelectQ)
                                        MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
                                        MyRs.Close()
                                        If MsgBox("Stop the App?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                            End
                                        End If
                                    Finally
                                        MyCmd = Nothing
                                        MyRs = Nothing
                                    End Try
                                End If

                            End If
                            'WriteFile("***  " & destDir + sFile)
                        End If
                    End If
                Catch ex As Exception
                    WriteFile("ERR : " & txtSourceLoc.Text & " :   " & Err.Description)
                    isEverythingOK = False
                End Try

                System.Windows.Forms.Application.DoEvents()
                'MsgBox("")
            Next i
        Catch ex As Exception
            WriteFile("ERR : " & txtSourceLoc.Text & " :   " & Err.Description)
            isEverythingOK = False
        End Try
        'errHand:
        '        If Err.Description <> "" And Err.Description <> "Resume without error." Then
        '            MsgBox(Err.Description & vbCrLf & Err.Source)
        '            WriteFile("ERR : " & txtSourceLoc.Text & " :   " & Err.Description)
        '        End If
        'Resume Next
    End Sub



    Sub OpenFile()
        'strOutputFilePath = txtOutput.Text & "\" & CboFormat.SelectedItem.ToString & "_" & Format(Date.Now, "ddMMMyyyy_hhmmsstt") & ".txt"
        'strOutputFilePath = txtSourceLoc.Text & "\LOG_" & Format(Date.Now, "ddMMMyyyy_hhmmsstt") & ".txt"
        strOutputFilePath = strLOGINSFilePath & "\_" & strLOGINSFileName & "_LOG_" & Format(Date.Now, "ddMMMyyyy_hhmmsstt") & ".log"

        xWrite = xFile.CreateText(strOutputFilePath)
    End Sub
    Sub WriteFile(ByVal Str As String)
        xWrite.WriteLine(Str)
    End Sub
    Sub CloseFile()
        xWrite.Close()
        xFile = Nothing
        System.Windows.Forms.Application.DoEvents()
        Shell("Notepad.exe " & strOutputFilePath, AppWinStyle.MaximizedFocus)
    End Sub
End Class