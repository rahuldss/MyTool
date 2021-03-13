Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class frmImportExport
    Dim strLOGINSFilePath As String
    Dim strLOGINSFileName As String

    Dim ConSource As New OleDbConnection
    Dim ConDest As New OleDbConnection
    Dim ConDestSql As New SqlConnection
    Dim strSourceFile As String

    Dim sqlBulk As SqlBulkCopy

    Dim strColumnsSOURCEMapped As String()
    Dim strColumnsDESTMapped As String()
    Dim intColumnsSOURCEMappedIndex As Integer = 0
    Dim intColumnsDESTMappedIndex As Integer = 0
    Dim Resetting As Boolean = False

    Private Sub ImportExport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.MdiParent = MDI
        Try
            '***GET DATABASE SETTINGS***
            Dim strDBDetailsSplit() As String = MyCLS.clsCOMMON.GetSettings()
            'txtFile.Text = strDBDetailsSplit(0)
            txtServer.Text = strDBDetailsSplit(1)
            txtUID.Text = strDBDetailsSplit(2)
            txtPassword.Text = strDBDetailsSplit(3)
            txtDatabase.Text = strDBDetailsSplit(4)
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub

    Private Sub cmdSelectExcelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectExcelFile.Click
        Try
            If Len(txtExcelFile.Text) > 0 Then
                OFDialog1.FileName = txtExcelFile.Text
            Else
                OFDialog1.InitialDirectory = Environment.SpecialFolder.Desktop
            End If

            OFDialog1.ShowDialog(Me)
            '.xls
            'If Len(OFDialog1.FileName) > 0 And (OFDialog1.FileName <> "*.xls" Or OFDialog1.FileName <> "*.xlsx") Then
            txtExcelFile.Text = OFDialog1.FileName.ToString
            strLOGINSFilePath = Mid(OFDialog1.FileName, 1, Len(OFDialog1.FileName) - InStr(StrReverse(OFDialog1.FileName), "\"))

            strLOGINSFileName = Mid(OFDialog1.FileName, Len(OFDialog1.FileName) - InStr(StrReverse(OFDialog1.FileName), "\") + 2, InStr(StrReverse(OFDialog1.FileName), "\") - 1)
            strLOGINSFileName = Mid(strLOGINSFileName, 1, Len(strLOGINSFileName) - 4)

            strSourceFile = txtExcelFile.Text
            Call FillSheets()
            Call MyCLS.clsControls.ComboBox_AdjustWidth(cboTableSource)
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub

    Private Sub cboTableSource_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTableSource.SelectedIndexChanged
        Try
            FillListWithColumnsFromSource()
            ReDim Preserve strColumnsSOURCEMapped(LstColumnsSource.Items.Count)
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub
    Private Sub cboTableDest_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTableDest.SelectedIndexChanged
        Try
            FillListWithColumnsFromDest()
            ReDim Preserve strColumnsDESTMapped(LstColumnsDest.Items.Count)
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try
            '***SAVE DATABASE SETTINGS***
            MyCLS.clsCOMMON.SaveSettings("", txtServer.Text, txtUID.Text, txtPassword.Text, txtDatabase.Text)

            OpenConnectionDest()

            FillListWithTablesFromDest()
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub

    Private Sub cmdImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport.Click
        Try

            'For i As Int16 = 0 To lstColumnsMapped.Items.Count - 1                
            '    MsgBox(strColumnsSOURCEMapped(i).ToString & " : " & strColumnsDESTMapped(i).ToString)
            'Next

            'Exit Sub

            Call GetDataFromSource()
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub




    Function CreateConnString(Optional ByVal isSQL As String = "") As String
        Dim strConnStr As String
        If chkDBFile.Checked = True Then
            strConnStr = txtServer.Text
        Else
            If Len(isSQL) > 0 Then
                strConnStr = "UID=" & txtUID.Text & ";Password=" & txtPassword.Text & ";Data Source=" & txtServer.Text & ";Initial Catalog=" & txtDatabase.Text & ";"
            Else
                strConnStr = "UID=" & txtUID.Text & ";Password=" & txtPassword.Text & ";Data Source=" & txtServer.Text & ";Initial Catalog=" & txtDatabase.Text & ";Provider=SQLOLEDB.1;"
            End If

            If Len(txtPassword.Text) = 0 Then
                strConnStr = strConnStr.Replace("Password=", "Integrated Security=SSPI")
            End If
        End If
        Return strConnStr
    End Function
    Private Sub OpenConnectionDest()
        Try
            Try
                ConDest.Close()
            Catch ex As Exception

            End Try
            Try
                ConDestSql.Close()
            Catch ex As Exception

            End Try            
            'Con.ConnectionString = "Server=.;Initial Catalog=jpbrothers;uid=sa;password=sa123;" 'Integrated Security=SSPI"
            ConDest.ConnectionString = CreateConnString()
            ConDest.Open()

            ConDestSql.ConnectionString = CreateConnString("SQL")
            ConDestSql.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub OpenConnectionSrc()
        Try
            Try
                ConSource.Close()
            Catch ex As Exception

            End Try
            If InStr(strSourceFile, ".xls") > 0 Then
                ConSource.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & strSourceFile & ";" & _
                   "Extended Properties=Excel 8.0;"
            Else
                ConSource.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & strSourceFile & ";" & _
                   "Persist Security Info=True;"
            End If
            ConSource.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub FillListWithTablesFromSource()
        Try
            MyCLS.clsCOMMON.SetCon(ConSource)

            Dim strSheets As String() = MyCLS.clsDBOperations.GetTables()

            cboTableSource.Items.Clear()
            For i As Int16 = 0 To strSheets.Length - 1
                cboTableSource.Items.Add(strSheets(i))
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub
    Sub FillListWithTablesFromDest()
        Try
            MyCLS.clsCOMMON.SetCon(ConDest)

            Dim strSheets As String() = MyCLS.clsDBOperations.GetTables()

            cboTableDest.Items.Clear()
            For i As Int16 = 0 To strSheets.Length - 1
                cboTableDest.Items.Add(strSheets(i))
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub
    Sub FillListWithColumnsFromSource()
        Try
            MyCLS.clsCOMMON.SetCon(ConSource)
            Dim strColumns As String(,) = MyCLS.clsDBOperations.GetColumns("[" & Replace(cboTableSource.Text, "'", "") & "]")

            LstColumnsSource.Items.Clear()
            For i As Int16 = 0 To strColumns.Length - 1
                LstColumnsSource.Items.Add(strColumns(i, 0))
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub
    Sub FillListWithColumnsFromDest()
        Try
            MyCLS.clsCOMMON.SetCon(ConDest)
            Dim strColumns As String(,) = MyCLS.clsDBOperations.GetColumns("[" & Replace(cboTableDest.Text, "'", "") & "]")

            LstColumnsDest.Items.Clear()
            For i As Int16 = 0 To strColumns.Length - 1
                LstColumnsDest.Items.Add(strColumns(i, 0))
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub

    Sub FillSheets()
        Try
            Call OpenConnectionSrc()
            Call FillListWithTablesFromSource()
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub

    Public Sub GetDataFromSource()
        Try
            'Dim ExcelCommand As New System.Data.OleDb.OleDbCommand("SELECT INTO [ODBC Driver={SQL Server};Server=tsi_dev_02;Database=ndhhs_updated;uid=sa;pwd=sa123].[tblOutstanding] FROM [Sheet1$]", Olecn)            
            'Dim ExcelCommand As New OleDbCommand("SELECT * FROM [" & strSheetName & "$]", Olecn)
            Dim sQ As String = CreateXlsQry(Replace(cboTableSource.Text, "'", ""))
            Dim ExcelCommand As New OleDbCommand(sQ, ConSource)

            Dim Rs As OleDbDataReader = ExcelCommand.ExecuteReader()

            '***From Mala Mam - SQLBULKCOPY ************
            sqlBulk = New SqlBulkCopy(ConDestSql)
            sqlBulk.DestinationTableName = cboTableDest.Text
            'sqlBulk.ColumnMappings.Add("ISBN", "ISBN")
            'sqlBulk.ColumnMappings.Add("TITLE", "TITLE")

            For i As Int16 = 0 To lstColumnsMapped.Items.Count - 1
                'MsgBox(strColumnsSOURCEMapped(i).ToString & " : " & strColumnsDESTMapped(i).ToString)
                sqlBulk.ColumnMappings.Add(strColumnsSOURCEMapped(i).ToString, strColumnsDESTMapped(i).ToString)
            Next

            '*
            'sqlBulk.BatchSize = 2
            'MyCLS.clsCOMMON.SetCon(ConSource)            

            'While Rs.Read()
            '    'For i As Long = 0 To MyCLS.clsCOMMON.fnQuerySelect1Value("Select Count(*) From [" & Replace(cboTableSource.Text, "'", "") & "]", "Number")
            '    Debug.Print(Rs("ISBN") & " : " & Rs(1))
            '    'Next
            'End While
            '********************
            'Dim ds As New DataSet
            'MyCLS.clsCOMMON.prcQuerySelectDS(ds, sQ, Replace(cboTableSource.Text, "'", ""))
            'For i As Long = 0 To ds.Tables(0).Rows.Count - 1
            '    Debug.Print(ds.Tables(0).Rows(i)("ISBN").ToString() & " : " & ds.Tables(0).Rows(i)(1).ToString())
            'Next

            'MyCLS.clsCOMMON.SetCon(ConDestSql)
            '**

            sqlBulk.WriteToServer(Rs)
            MsgBox("Data transfer to sql database successfully")
        Catch ex As Exception
            MsgBox(ex.Message)
            'MYCLS.strGlobalErrorInfo = "Query is : " & TruncateTable
            MyCLS.strGlobalErrorInfo = MyCLS.strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
            MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
            MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.Data)
            MyCLS.clsCOMMON.fnWrite2LOG(MyCLS.strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
        End Try
    End Sub

    Function CreateXlsQry(ByVal TableName As String) As String
        Dim strQ As String = ""
        Try
            MyCLS.clsCOMMON.SetCon(ConSource)
            strQ = "SELECT Top " & MyCLS.clsCOMMON.fnQuerySelect1Value("Select Count(*) From [" & TableName & "]", "Number") & " "
            For i As Int16 = 0 To LstColumnsSource.Items.Count - 1
                If LstColumnsSource.GetItemChecked(i) = True Then
                    strQ = strQ & "[" & LstColumnsSource.Items(i) & "],"
                    '"[ISBN],[TITLE] FROM [" & strSheetName & "$]"
                End If
            Next
            strQ = Mid(strQ, 1, Len(strQ) - 1) & " From [" & TableName & "]"
            MyCLS.clsCOMMON.SetCon(ConDest)
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
        Return strQ
    End Function

    Private Sub LstColumnsSource_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles LstColumnsSource.ItemCheck        
        Try
            If Resetting = True Then Exit Sub
            If LstColumnsSource.GetItemCheckState(e.Index) = CheckState.Checked Then
                ''strColumnsSOURCEMapped(intColumnsSOURCEMappedIndex - 1).Remove(0)
                'strColumnsSOURCEMapped(intColumnsSOURCEMappedIndex - 1) = Nothing
                'intColumnsSOURCEMappedIndex -= 1
                LstColumnsSource.SetSelected(e.Index, True)
                MsgBox("Please Click on Reset Mapping and Then Start Again!", MsgBoxStyle.Critical, "Can't Remove")
            Else
                strColumnsSOURCEMapped(intColumnsSOURCEMappedIndex) = LstColumnsSource.Items(e.Index).ToString()
                'intColumnsSOURCEMappedIndex += 1
            End If
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub

    Private Sub LstColumnsDest_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles LstColumnsDest.ItemCheck
        Try
            If Resetting = True Then Exit Sub
            If LstColumnsDest.GetItemCheckState(e.Index) = CheckState.Checked Then
                ''strColumnsDESTMapped(intColumnsSOURCEMappedIndex - 1).Remove(0)
                'strColumnsDESTMapped(intColumnsSOURCEMappedIndex - 1) = Nothing
                ''intColumnsSOURCEMappedIndex -= 1
                LstColumnsDest.SetSelected(e.Index, True)
                MsgBox("Please Click on Reset Mapping and Then Start Again!", MsgBoxStyle.Critical, "Can't Remove")
            Else
                strColumnsDESTMapped(intColumnsSOURCEMappedIndex) = LstColumnsDest.Items(e.Index).ToString()

                lstColumnsMapped.Items.Add(strColumnsSOURCEMapped(intColumnsSOURCEMappedIndex) & " : " & strColumnsDESTMapped(intColumnsSOURCEMappedIndex))

                intColumnsSOURCEMappedIndex += 1
            End If
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub

    Private Sub cmdResetMapping_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResetMapping.Click
        Try
            ReDim Preserve strColumnsSOURCEMapped(LstColumnsSource.Items.Count)
            ReDim Preserve strColumnsDESTMapped(LstColumnsDest.Items.Count)

            intColumnsSOURCEMappedIndex = 0

            lstColumnsMapped.Items.Clear()

            Resetting = True
            MyCLS.clsControls.prcListUnCheckAll(LstColumnsDest)
            MyCLS.clsControls.prcListUnCheckAll(LstColumnsSource)
            Resetting = False
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try        
    End Sub

    Private Sub chkSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectAll.CheckedChanged
        Try
            For i As Int16 = 0 To LstColumnsSource.Items.Count - 1
                LstColumnsSource.SetItemChecked(i, chkSelectAll.Checked)
                LstColumnsDest.SetItemChecked(i, chkSelectAll.Checked)
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Sub
End Class