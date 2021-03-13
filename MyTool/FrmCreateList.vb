Imports System.Data.OleDb
Imports System.IO

Public Class FrmCreateList

    Dim xFile As System.IO.File
    Dim xWrite As System.IO.StreamWriter
    Dim strOutputFilePath As String = "C:\_List.txt"

    Private Sub FrmCreateList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.MdiParent = MDI
            txtOutput.Text = "C:\_CODE"
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CmdExportFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdExportFolder.Click
        If Len(txtOutput.Text) > 0 Then
            If System.IO.Directory.Exists(txtOutput.Text) = True Then
                FBDialog1.SelectedPath = txtOutput.Text
            Else
                FBDialog1.RootFolder = Environment.SpecialFolder.Desktop
            End If
        Else
            FBDialog1.RootFolder = Environment.SpecialFolder.Desktop
        End If

        FBDialog1.ShowDialog(Me)
        txtOutput.Text = FBDialog1.SelectedPath.ToString
    End Sub

    ''Private Sub CmdCreateList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCreateList.Click
    ''    If fnValidate() Then
    ''        MyCLS.clsFileHandling.OpenFile(strOutputFilePath)
    ''        If chkListInXLS.Checked = True Then
    ''            'Call prcCreateListInXLS(txtOutput.Text, 0)                
    ''            'Call prcCreateList(txtOutput.Text, 0)
    ''            MyCLS.clsFileHandling.prcCreateFileListInTxt(txtOutput.Text, 0, True, ".")
    ''            MyCLS.clsFileHandling.CloseFile()
    ''            Dim TxtXls As New MyCLS.clsXLSOperations
    ''            TxtXls.Txt2Xls(strOutputFilePath, True)
    ''            'Call Txt2Xls(strOutputFilePath)
    ''        Else
    ''            'Call prcCreateList(txtOutput.Text, 0)
    ''            MyCLS.clsFileHandling.prcCreateFileListInTxt(txtOutput.Text, 0, True, ".")
    ''            MyCLS.clsFileHandling.CloseFile(True)
    ''        End If

    ''        lblCPath.Text = ""
    ''        lblDirName.Text = ""
    ''        lblFileName.Text = ""
    ''    End If
    ''End Sub


    Function fnValidate() As Boolean
        If System.IO.Directory.Exists(txtOutput.Text) = False Then
            MsgBox("Please Check Output Path!", MsgBoxStyle.Critical)
            fnValidate = False
        Else
            fnValidate = True
        End If
    End Function

    Sub prcCreateList(ByVal DirLoc As String, ByVal FillTab As Integer)
        On Error Resume Next
        Dim i As Integer
        Dim posSep As Integer
        Dim sDir As String
        Dim aDirs() As String
        Dim sFile As String
        Dim aFiles() As String

        aDirs = System.IO.Directory.GetDirectories(DirLoc)

        '//PBExtract.Maximum = IIf(aDirs.GetUpperBound(0) > 0, aDirs.GetUpperBound(0), 100)
        For i = 0 To aDirs.GetUpperBound(0)
            ' Get the position of the last separator in the current path.
            posSep = aDirs(i).LastIndexOf("\")
            ' Get the path of the source directory.
            sDir = aDirs(i).Substring((posSep + 1), aDirs(i).Length - (posSep + 1))
            lblCPath.Text = aDirs(i)
            lblDirName.Text = sDir
            Debug.Print(Space(FillTab * 5) & sDir)
            'WriteFile(Space(FillTab * 5) & sDir)
            MyCLS.clsFileHandling.WriteFile(MyCLS.clsCOMMON.fnTABs(FillTab) & sDir)

            ' Since we are in recursive mode, copy the children also
            FillTab = FillTab + 1
            prcCreateList(aDirs(i), FillTab)
            FillTab = FillTab - 1
        Next

        ' Get the files from the current parent.
        aFiles = System.IO.Directory.GetFiles(DirLoc)

        ' Copy all files.
        For i = 0 To aFiles.GetUpperBound(0)
            ' Get the position of the trailing separator.
            posSep = aFiles(i).LastIndexOf("\")

            ' Get the full path of the source file.
            sFile = aFiles(i).Substring((posSep + 1), aFiles(i).Length - (posSep + 1))
            lblFileName.Text = sFile
            Debug.Print(Space(FillTab * 5) & sFile & " - (" & System.IO.File.ReadAllBytes(DirLoc & "\" & sFile).Length & ")")
            MyCLS.clsFileHandling.WriteFile(MyCLS.clsCOMMON.fnTABs(FillTab) & sFile & " - (" & System.IO.File.ReadAllBytes(DirLoc & "\" & sFile).Length & ")")

            System.Windows.Forms.Application.DoEvents()
        Next i    
    End Sub

    'Sub Txt2Xls(ByVal TxtFilePath As String)
    '    Const xlFixedWidth = 2
    '    Const xlNormal = -4143
    '    Const xlLastCell = 11

    '    Dim sFiNa
    '    Dim oFS
    '    Dim oExcel
    '    Dim oWBook
    '    Dim sTmp

    '    ' you'll have to change this 
    '    sFiNa = TxtFilePath

    '    oFS = CreateObject("Scripting.FileSystemObject")
    '    sFiNa = oFS.GetAbsolutePathName(sFiNa)
    '    oExcel = CreateObject("Excel.Application")

    '    oExcel.Visible = True   ' while testing 
    '    ' oExcel.Whatever = False  ' todo: what property to set to suppress silly question 

    '    sTmp = "Working with MS Excel Vers. " & oExcel.Version _
    '           & " (" & oExcel.Workbooks.Count & " Workbooks)"

    '    'WScript.Echo(sTmp)

    '    'oExcel.Workbooks.Open(sFiNa, xlFixedWidth)
    '    oExcel.Workbooks.Open(sFiNa, , , 1, , , , , vbTab)
    '    oWBook = oExcel.Workbooks(1)
    '    'WScript.Echo("Open:   ", oWBook.Name)

    '    ' magic from Giovanni Cenati 
    '    ' http://www.codecomments.com/archive299-2005-2-401145.html 
    '    ' oExcel.Range(oExcel.cells(1,1),oExcel.cells(100,1)).Select 

    '    'oExcel.Range(oExcel.cells(1, 1), oExcel.cells(oWBook.Sheets(1).UsedRange.SpecialCells(xlLastCell).Row, 1)).Select()
    '    'oExcel.Selection.TextToColumns(oExcel.Range("A1"), xlFixedWidth)

    '    ' save as XLS 
    '    oWBook.SaveAs(sFiNa + ".xls", xlNormal)
    '    'WScript.Echo("SaveAs: ", oWBook.Name)

    '    oWBook.Close()
    '    oExcel.Quit()
    'End Sub

    'Function fnTABs(ByVal FillTab As Int16) As String
    '    Dim strTABs As String = ""
    '    For i As Int16 = 0 To FillTab - 1
    '        strTABs = strTABs & vbTab
    '    Next
    '    Return strTABs
    'End Function





    'Sub GetDataFromXLS(ByVal vFile As String, ByVal strSheetName As String)
    '    Try
    '        'Dim Oleda As OleDbDataAdapter
    '        Dim Olecn As OleDbConnection
    '        'Dim dt1 As DataTable            
    '        Olecn = New OleDbConnection( _
    '            "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    '            "Data Source=" & vFile & ";" & _
    '            "Extended Properties=Excel 8.0;")
    '        Olecn.Open()

    '        'Dim Rss As New ADODB.Recordset
    '        'Rss.Open("[" & strSheetName & "$]", Olecn)
    '        'Rss.AddNew()
    '        Dim ds As New DataSet
    '        MyCLS.clsCOMMON.SetCon(Olecn)
    '        MyCLS.clsCOMMON.prcQuerySelectDS(ds, "Select * From [" & strSheetName & "$]", "[" & strSheetName & "$]")

    '        MsgBox(ds.Tables(0).Columns.Count)

    '        Dim DRCons As DataRow
    '        DRCons = ds.Tables(0).NewRow()
    '        DRCons.Item(0) = 0
    '        DRCons(1) = 1
    '        ds.Tables(0).Rows.Add(DRCons)
    '        ds.AcceptChanges()


    '        Olecn.Close()
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '        'MYCLS.strGlobalErrorInfo = "Query is : " & TruncateTable
    '        MyCLS.strGlobalErrorInfo = MyCLS.strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
    '        MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
    '        MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.Data)
    '        MyCLS.clsCOMMON.fnWrite2LOG(MyCLS.strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
    '    End Try
    'End Sub
End Class