Public Class frmRenameFiles
    Dim strLOGINSFilePath As String
    Dim strLOGINSFileName As String

    Private Sub frmRenameFiles_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = MDI
    End Sub

    Private Sub chkRenameAtSameLoc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRenameAtSameLoc.CheckedChanged
        If chkRenameAtSameLoc.Checked = True Then
            gbDest.Enabled = False
        Else
            gbDest.Enabled = True
        End If
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

    Private Sub cmdSelectDest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectDest.Click
        Try
            If Len(txtDestLoc.Text) > 0 Then
                FolderBrowserDialog1.SelectedPath = txtDestLoc.Text
            Else
                FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.Desktop
            End If

            FolderBrowserDialog1.ShowDialog(Me)
            '.xls
            'If Len(FolderBrowserDialog1.SelectedPath) > 0 And (FolderBrowserDialog1.SelectedPath <> "*.xls" Or FolderBrowserDialog1.SelectedPath <> "*.xlsx") Then
            txtDestLoc.Text = FolderBrowserDialog1.SelectedPath.ToString
            strLOGINSFilePath = Mid(FolderBrowserDialog1.SelectedPath, 1, Len(FolderBrowserDialog1.SelectedPath) - InStr(StrReverse(FolderBrowserDialog1.SelectedPath), "\"))

            strLOGINSFileName = Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath) - InStr(StrReverse(FolderBrowserDialog1.SelectedPath), "\") + 2, InStr(StrReverse(FolderBrowserDialog1.SelectedPath), "\") - 1)
            strLOGINSFileName = Mid(strLOGINSFileName, 1, Len(strLOGINSFileName) - 4)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRename.Click
        Try
            If System.IO.Directory.Exists(txtSourceLoc.Text) = False Then
                MsgBox("Check Source Path!")
                Exit Sub
            End If
            If Len(txtDestLoc.Text) > 0 Then
                If System.IO.Directory.Exists(txtDestLoc.Text) = False Then
                    MsgBox("Check Destination Path!")
                    Exit Sub
                End If
            End If

            If chkRenameAtSameLoc.Checked = True Then
                Call RenameFilesAtSameLoc()
            Else
                Call CopyFilesToDestLoc()
            End If
            lblMSG.Text = "Done!"            
        Catch ex As Exception

        End Try
    End Sub

    Sub RenameFilesAtSameLoc()
        Try
            RecursiveRenameFiles(txtSourceLoc.Text, True, txtRemoveChars.Text, txtReplaceChars.Text)
        Catch ex As Exception

        End Try
    End Sub

    Sub CopyFilesToDestLoc()
        Try
            RecursiveCopyFiles(txtSourceLoc.Text, txtDestLoc.Text, True, txtRemoveChars.Text, txtReplaceChars.Text)
        Catch ex As Exception

        End Try
    End Sub


    Public Sub RecursiveCopyFiles(ByVal sourceDir As String, ByVal destDir As String, ByVal fRecursive As Boolean, Optional ByVal CharsToRemoveFromFileNames As String = "", Optional ByVal CharsToReplaceWithInFileNames As String = "")
        Try
            Dim i As Integer
            Dim posSep As Integer
            Dim sDir As String
            Dim aDirs() As String
            Dim sFile As String
            Dim aFiles() As String

            'cmdNext.Enabled = False
            'cmdBack.Enabled = False

            ' Add trailing separators to the supplied paths if they don't exist.
            If Not sourceDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
                sourceDir &= System.IO.Path.DirectorySeparatorChar
            End If

            If Not destDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
                destDir &= System.IO.Path.DirectorySeparatorChar
            End If

            ' Recursive switch to continue drilling down into dir structure.
            If fRecursive Then
                ' Get a list of directories from the current parent.
                aDirs = System.IO.Directory.GetDirectories(sourceDir)

                '//PBExtract.Maximum = IIf(aDirs.GetUpperBound(0) > 0, aDirs.GetUpperBound(0), 100)

                For i = 0 To aDirs.GetUpperBound(0)
                    ' Get the position of the last separator in the current path.
                    posSep = aDirs(i).LastIndexOf("\")
                    ' Get the path of the source directory.
                    sDir = aDirs(i).Substring((posSep + 1), aDirs(i).Length - (posSep + 1))
                    ' Create the new directory in the destination directory.
                    System.IO.Directory.CreateDirectory(destDir + sDir)

                    ' //PBExtract.Value = i

                    ' Since we are in recursive mode, copy the children also
                    RecursiveCopyFiles(aDirs(i), (destDir + sDir), fRecursive, CharsToRemoveFromFileNames, CharsToReplaceWithInFileNames)
                Next
            End If
            ' Get the files from the current parent.
            aFiles = System.IO.Directory.GetFiles(sourceDir)
            'PBExtract.Maximum = IIf(aFiles.GetUpperBound(0) > 0, aFiles.GetUpperBound(0), 100)
            ' Copy all files.
            For i = 0 To aFiles.GetUpperBound(0)
                ' Get the position of the trailing separator.
                posSep = aFiles(i).LastIndexOf("\")
                ' Get the full path of the source file.

                sFile = aFiles(i).Substring((posSep + 1), aFiles(i).Length - (posSep + 1))

                lblMSG.text = sFile

                ' Copy the file.        
                Try
                    If Len(CharsToRemoveFromFileNames) > 0 Then
                        If Len(CharsToReplaceWithInFileNames) > 0 Then
                            System.IO.File.Copy(aFiles(i), destDir + Replace(sFile, CharsToRemoveFromFileNames, CharsToReplaceWithInFileNames))
                        Else
                            System.IO.File.Copy(aFiles(i), destDir + Replace(sFile, CharsToRemoveFromFileNames, ""))
                        End If
                    Else
                        System.IO.File.Copy(aFiles(i), destDir + sFile)
                    End If
                Catch ex As Exception
                    lblMSG.Text = ex.Message
                    If InStr(ex.Message, "already exists") > 0 Then
                        System.IO.File.Copy(aFiles(i), destDir + "already exists_" + sFile)
                    End If
                    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
                End Try

                'lstExtractFiles.Items.Add("Extract :" & sFile & "......")
                'lstExtractFiles.SelectedIndex = lstExtractFiles.Items.Count - 1
                'lstExtractFiles.EndUpdate()
                System.Windows.Forms.Application.DoEvents()
                'PBExtract.Value = i
            Next i
        'PBExtract.Value = 0
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
        End Try
    End Sub

    Public Sub RecursiveRenameFiles(ByVal sourceDir As String, ByVal fRecursive As Boolean, Optional ByVal CharsToRemoveFromFileNames As String = "", Optional ByVal CharsToReplaceWithInFileNames As String = "")
        Try
            Dim i As Integer
            Dim posSep As Integer
            Dim sDir As String
            Dim aDirs() As String
            Dim sFile As String
            Dim aFiles() As String

            'cmdNext.Enabled = False
            'cmdBack.Enabled = False

            ' Add trailing separators to the supplied paths if they don't exist.
            If Not sourceDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
                sourceDir &= System.IO.Path.DirectorySeparatorChar
            End If

            '' ''If Not destDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
            '' ''    destDir &= System.IO.Path.DirectorySeparatorChar
            '' ''End If

            ' Recursive switch to continue drilling down into dir structure.
            If fRecursive Then
                ' Get a list of directories from the current parent.
                aDirs = System.IO.Directory.GetDirectories(sourceDir)

                '//PBExtract.Maximum = IIf(aDirs.GetUpperBound(0) > 0, aDirs.GetUpperBound(0), 100)

                For i = 0 To aDirs.GetUpperBound(0)
                    ' Get the position of the last separator in the current path.
                    posSep = aDirs(i).LastIndexOf("\")
                    ' Get the path of the source directory.
                    sDir = aDirs(i).Substring((posSep + 1), aDirs(i).Length - (posSep + 1))
                    ' Create the new directory in the destination directory.
                    '' ''System.IO.Directory.CreateDirectory(destDir + sDir)

                    ' //PBExtract.Value = i

                    ' Since we are in recursive mode, copy the children also
                    RecursiveCopyFiles(aDirs(i), fRecursive, CharsToRemoveFromFileNames, CharsToReplaceWithInFileNames)
                Next
            End If
            ' Get the files from the current parent.
            aFiles = System.IO.Directory.GetFiles(sourceDir)
            'PBExtract.Maximum = IIf(aFiles.GetUpperBound(0) > 0, aFiles.GetUpperBound(0), 100)
            ' Copy all files.
            For i = 0 To aFiles.GetUpperBound(0)
                ' Get the position of the trailing separator.
                posSep = aFiles(i).LastIndexOf("\")
                ' Get the full path of the source file.

                sFile = aFiles(i).Substring((posSep + 1), aFiles(i).Length - (posSep + 1))

                lblMSG.Text = sFile

                ' Copy the file.
                Try
                    If Len(CharsToRemoveFromFileNames) > 0 Then
                        If Len(CharsToReplaceWithInFileNames) > 0 Then
                            System.IO.File.Move(aFiles(i), sourceDir + Replace(sFile, CharsToRemoveFromFileNames, CharsToReplaceWithInFileNames))
                        Else
                            System.IO.File.Move(aFiles(i), sourceDir + Replace(sFile, CharsToRemoveFromFileNames, ""))
                        End If
                    Else
                        System.IO.File.Move(aFiles(i), sourceDir + "Copy_" + sFile)
                    End If
                Catch ex As Exception
                    lblMSG.Text = ex.Message
                    If InStr(ex.Message, "already exists") > 0 Then
                        System.IO.File.Copy(aFiles(i), sourceDir + "already exists_" + sFile)
                        System.IO.File.Delete(aFiles(i))
                    End If
                    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
                End Try

                'lstExtractFiles.Items.Add("Extract :" & sFile & "......")
                'lstExtractFiles.SelectedIndex = lstExtractFiles.Items.Count - 1
                'lstExtractFiles.EndUpdate()
                System.Windows.Forms.Application.DoEvents()
                'PBExtract.Value = i
            Next i
            'PBExtract.Value = 0
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
        End Try
    End Sub
End Class