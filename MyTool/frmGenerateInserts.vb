Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class frmGenerateInserts
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

    Private Sub GenerateInserts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.MdiParent = MDI
    End Sub

    Private Sub frmGenerateInserts_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Try
            '***GET DATABASE SETTINGS***
            'Dim strDBDetailsSplit() As String = MyCLS.clsCOMMON.GetSettings()
            ''txtFile.Text = strDBDetailsSplit(0)
            'txtServer.Text = strDBDetailsSplit(1)
            'txtUID.Text = strDBDetailsSplit(2)
            'txtPassword.Text = strDBDetailsSplit(3)
            'txtDatabase.Text = strDBDetailsSplit(4)

            '***GET DATABASE SETTINGS SOURCE***
            Dim strDBDetailsSplitSRC() As String = MyCLS.clsCOMMON.GetSettings(True)
            'txtFile.Text = strDBDetailsSplit(0)
            txtServer.Text = strDBDetailsSplitSRC(1)
            txtUID.Text = strDBDetailsSplitSRC(2)
            txtPassword.Text = strDBDetailsSplitSRC(3)
            txtDatabase.Text = strDBDetailsSplitSRC(4)
        Catch ex As Exception

        End Try
    End Sub

    'Private Sub cmdSelectExcelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If Len(txtExcelFile.Text) > 0 Then
    '            OFDialog1.FileName = txtExcelFile.Text
    '        Else
    '            OFDialog1.InitialDirectory = Environment.SpecialFolder.Desktop
    '        End If

    '        OFDialog1.ShowDialog(Me)
    '        '.xls
    '        'If Len(OFDialog1.FileName) > 0 And (OFDialog1.FileName <> "*.xls" Or OFDialog1.FileName <> "*.xlsx") Then
    '        txtExcelFile.Text = OFDialog1.FileName.ToString
    '        strLOGINSFilePath = Mid(OFDialog1.FileName, 1, Len(OFDialog1.FileName) - InStr(StrReverse(OFDialog1.FileName), "\"))

    '        strLOGINSFileName = Mid(OFDialog1.FileName, Len(OFDialog1.FileName) - InStr(StrReverse(OFDialog1.FileName), "\") + 2, InStr(StrReverse(OFDialog1.FileName), "\") - 1)
    '        strLOGINSFileName = Mid(strLOGINSFileName, 1, Len(strLOGINSFileName) - 4)

    '        strSourceFile = txtExcelFile.Text
    '        Call FillSheets()
    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub cboTableSource_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        FillListWithColumnsFromSource()
    '        ReDim Preserve strColumnsSOURCEMapped(LstColumnsSource.Items.Count)
    '    Catch ex As Exception

    '    End Try
    'End Sub
    Private Sub cboTableDest_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTableDest.SelectedIndexChanged
        Try
            FillListWithColumnsFromDest()
            ReDim Preserve strColumnsDESTMapped(LstColumnsDest.Items.Count)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try
            '***SAVE DATABASE SETTINGS***
            MyCLS.clsCOMMON.SaveSettings("", txtServer.Text, txtUID.Text, txtPassword.Text, txtDatabase.Text, True)

            OpenConnectionDest()

            FillListWithTablesFromDest()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdGenerateInserts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerateInserts.Click
        Try
            Dim ds As New DataSet
            Dim objParamList As New List(Of SqlParameter)()
            Dim sTable_Name As String = ""
            Try
                OpenConnectionDest()
                MyCLS.clsCOMMON.SetCon(ConDest)
                MyCLS.clsCOMMON.SetCon(ConDestSql)

                IO.Directory.CreateDirectory(My.Application.Info.DirectoryPath() & "\_Inserts")
            Catch ex As Exception

            End Try

            Try
                If chkForAllTables.Checked = False Then 'GENERATE FOR ONE TABLE
                    sTable_Name = cboTableDest.Text
                    objParamList.Add(New SqlParameter("@table_name", sTable_Name))

                    '***To Generate for Selected Columns***
                    If MyCLS.clsControls.fnListIsChecked(LstColumnsDest) = True Then
                        Dim sColsSelected As String = ""
                        For iCol As Int16 = 0 To LstColumnsDest.Items.Count - 1
                            If LstColumnsDest.GetItemChecked(iCol) = True Then
                                sColsSelected = sColsSelected & "'" & LstColumnsDest.Items(iCol).ToString() & "',"
                            End If
                        Next
                        If Len(sColsSelected) > 0 Then
                            sColsSelected = Mid(sColsSelected, 1, Len(sColsSelected) - 1)
                        End If
                        objParamList.Add(New SqlParameter("@cols_to_include", sColsSelected))
                    End If
                    '**************************************
                    '***Top Rows***
                    If Len(txtTop.Text) > 0 And IsNumeric(txtTop.Text) Then
                        objParamList.Add(New SqlParameter("@top", txtTop.Text))
                    End If
                    '**************
                    '***Where Clause***
                    If Len(txtFrom_Where.Text) > 0 Then                        
                        objParamList.Add(New SqlParameter("@from", "From " & cboTableDest.Text & " " & txtFrom_Where.Text))
                    End If
                    '******************
                    ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_Generate_Inserts", objParamList)
                    If ds IsNot Nothing Then
                        If ds.Tables IsNot Nothing Then
                            If ds.Tables(0).Rows IsNot Nothing Then
                                If ds.Tables(0).Rows.Count > 0 Then
                                    MyCLS.clsFileHandling.OpenFile(My.Application.Info.DirectoryPath() & "\_Inserts\" & sTable_Name & ".sql")
                                    MyCLS.clsFileHandling.WriteFile("Set Identity_insert " & sTable_Name & " On")
                                    pb2.Minimum = 0
                                    pb2.Maximum = ds.Tables(0).Rows.Count - 1
                                    For intRow As Integer = 0 To ds.Tables(0).Rows.Count - 1
                                        pb2.Value = intRow
                                        System.Windows.Forms.Application.DoEvents()
                                        MyCLS.clsFileHandling.WriteFile(ds.Tables(0).Rows(intRow)(0).ToString())
                                    Next
                                    MyCLS.clsFileHandling.WriteFile("Set Identity_insert " & sTable_Name & " Off")
                                    MyCLS.clsFileHandling.CloseFile()
                                End If
                            End If
                        End If
                    End If
                Else 'GENERATE FOR ALL TABLE
                    Pb1.Minimum = 0
                    Pb1.Maximum = cboTableDest.Items.Count - 1
                    For i As Int16 = 0 To cboTableDest.Items.Count - 1
                        Try
                            Pb1.Value = i
                            sTable_Name = cboTableDest.Items(i).ToString()
                            Try
                                If sTable_Name = "sysdiagrams" Then Continue For
                                objParamList.RemoveAt(0)
                            Catch ex As Exception

                            End Try
                            objParamList.Add(New SqlParameter("@table_name", sTable_Name))

                            ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_Generate_Inserts", objParamList)
                            If ds IsNot Nothing Then
                                If ds.Tables IsNot Nothing Then
                                    If ds.Tables(0).Rows IsNot Nothing Then
                                        If ds.Tables(0).Rows.Count > 0 Then
                                            MyCLS.clsFileHandling.OpenFile(My.Application.Info.DirectoryPath() & "\_Inserts\" & sTable_Name & ".sql")
                                            MyCLS.clsFileHandling.WriteFile("Set Identity_insert " & sTable_Name & " On")
                                            pb2.Minimum = 0
                                            pb2.Maximum = ds.Tables(0).Rows.Count - 1
                                            For intRow As Integer = 0 To ds.Tables(0).Rows.Count - 1
                                                pb2.Value = intRow
                                                pb2.Update()
                                                System.Windows.Forms.Application.DoEvents()
                                                MyCLS.clsFileHandling.WriteFile(ds.Tables(0).Rows(intRow)(0).ToString())
                                            Next
                                            MyCLS.clsFileHandling.WriteFile("Set Identity_insert " & sTable_Name & " Off")
                                            MyCLS.clsFileHandling.CloseFile()
                                        End If
                                    End If
                                End If
                            End If
                            System.Windows.Forms.Application.DoEvents()
                        Catch ex As Exception
                            MsgBox(sTable_Name & vbCrLf & ex.Message & vbCrLf & "PB1 : " & Pb1.Value & vbCrLf & "PB2 : " & pb2.Value)
                        End Try
                    Next
                    Pb1.Value = cboTableDest.Items.Count - 1
                End If
                MsgBox("Done!")
                Process.Start(My.Application.Info.DirectoryPath() & "\_Inserts")
            Catch ex As Exception
                MsgBox(sTable_Name & vbCrLf & ex.Message & vbCrLf & "PB1 : " & Pb1.Value & vbCrLf & "PB2 : " & pb2.Value)
            End Try

            '***create SP from here***
            'SET ANSI_NULLS ON
            'GO
            'SET QUOTED_IDENTIFIER ON
            'GO
            'IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[Sp_GetMailingList4CPanel]') AND type in (N'P', N'PC'))
            'BEGIN
            'EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [Sp_GetMailingList4CPanel] 
            '
            'AS
            'BEGIN
            '
            '	select b.Isbn,b.Title, b.SPECIALITY, ea.Email,ea.Name  from Book as b
            '		inner join EmailAlert as ea on b.SPECIALITY = ea.Speciality
            '		inner join NEWRELEASED as nr on b.ISBN  = nr.NR_ISBN 
            '		where (nr.NR_ISBN Not in (select NR_ISBN from NEWRELEASED_Temp))
            '
            'END
            '' 
            'END
            'GO

            'Call GetDataFromSource()

            'MsgBox(MyCLS.clsCOMMON.fnQueryInsert(""))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdDisplayRecords_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDisplayRecords.Click
        Try
            Dim objclsXLSOperations As New MyCLS.clsXLSOperations

            Dim ds As New DataSet
            Dim strQry As String = ""
            If chkForAllTables.Checked = False Then
                strQry = CreateQry(cboTableDest.Text)
                'MsgBox(strQry)

                'Added for Gurgaon Rajat
                'strQry = <s>SELECT * FROM 
                '     (Select row_number() over (order by c.id,sd.id) AS line_no,
                '                C.id, c.first_name, c.Middle_name, c.last_name, c.pia, c.personal_email, c.work_email, c.primary_phone, c.secondary_phone, c.primary_phone_type, c.secondary_phone_type, c.state, c.city, c.zipcode, c.address1, c.address2, c.dob,
                '                sd.subscription_id, sd.customer_id, sd.product_id, sd.registration_date, sd.activation_date, sd.expiry_date,
                '                p.product_name, p.product_description,
                '                cc.cc_name_on_card, cc.cc_card_type, cc.amount_charged
                '     from Customer C
                '      Left Outer Join Subscription_Details sd
                '       On sd.customer_id=c.id
                '      Left Outer Join Product P
                '       On P.Id=sd.product_id
                '      Left Outer Join cc_details cc
                '       On cc.sid=sd.id) as cust
                '     Where cust.line_no between <%= txtRowFrom.Text %> and <%= txtRowTo.Text %></s>

                'rtbSelectQuery.Text = strQry
                'Added for Gurgaon Rajat
                'strQry = <s>SELECT * 
                '            FROM (SELECT row_number() over (order by UserID) AS line_no, * 
                '                  FROM dbo.Userinfo) as users
                '            WHERE users.line_no BETWEEN <%= txtTop.Text %> and <%= txtFrom_Where.Text %></s>

                'MyCLS.clsCOMMON.prcQuerySelectDS(ds, strQry, cboTableDest.Text)
                If (rtbSelectQuery.TextLength > 0) Then
                    MyCLS.clsCOMMON.prcQuerySelectDS(ds, rtbSelectQuery.Text, "records")
                Else
                    MyCLS.clsCOMMON.prcQuerySelectDS(ds, strQry, "records")
                End If

                'MyCLS.clsXLSOperations.DataSetToExcel("Records.xls", ds)
                'System.Diagnostics.Process.Start("Records.xls")
                'Added for Gurgaon Rajat
                If (txtFrom_Where.Text.Length > 0) Then
                    'MyCLS.clsXLSOperations.DataSetToExcel(txtDatabase.Text + "_" + txtRowTo.Text & ".xls", ds)
                    objclsXLSOperations.generateExcel(ds.Tables(0), "Records", "", False, False, IIf(cboTableDest.Text.Length > 0, txtDatabase.Text + "_" + cboTableDest.Text, txtDatabase.Text + "_" + "Records") + ".xls")
                    System.Diagnostics.Process.Start(My.Application.Info.DirectoryPath() & "\Export\" & txtDatabase.Text + "_" + txtRowTo.Text & ".xlsx")
                Else
                    'MyCLS.clsXLSOperations.DataSetToExcel(IIf(cboTableDest.Text.Length > 0, txtDatabase.Text + "_" + cboTableDest.Text, txtDatabase.Text + "_" + "Records") + ".xlsx", ds)
                    objclsXLSOperations.generateExcel(ds.Tables(0), "Records", "", False, False, IIf(cboTableDest.Text.Length > 0, txtDatabase.Text + "_" + cboTableDest.Text, txtDatabase.Text + "_" + "Records") + ".xlsx")
                    System.Diagnostics.Process.Start(My.Application.Info.DirectoryPath() & "\Export\" & IIf(cboTableDest.Text.Length > 0, txtDatabase.Text + "_" + cboTableDest.Text, txtDatabase.Text + "_" + "Records") + ".xlsx")
                End If
                'Added for Gurgaon Rajat
            Else    '####FOR ALL THE TABLES#####
                pb2.Minimum = 0
                pb2.Maximum = cboTableDest.Items.Count - 1
                For i As Int16 = 0 To cboTableDest.Items.Count - 1
                    lblTableName.Text = cboTableDest.Items(i).ToString()
                    pb2.Value = i
                    pb2.Update()
                    System.Windows.Forms.Application.DoEvents()

                    Dim newFile As New IO.FileInfo(My.Application.Info.DirectoryPath() & "\Export\" & IIf(cboTableDest.Items(i).ToString().Length > 0, txtDatabase.Text + "_" + cboTableDest.Items(i).ToString(), txtDatabase.Text + "_" + "Records") + ".xlsx")
                    If newFile.Exists Then
                        Continue For
                    End If

                    Dim dsAll As New DataSet
                    strQry = CreateQryAllColumns(cboTableDest.Items(i).ToString())

                    If (rtbSelectQuery.TextLength > 0) Then
                        MyCLS.clsCOMMON.prcQuerySelectDS(dsAll, rtbSelectQuery.Text, "records")
                    Else
                        MyCLS.clsCOMMON.prcQuerySelectDS(dsAll, strQry, "records")
                    End If

                    If (dsAll.Tables(0).Rows.Count > 400000) Then
                        Continue For
                    End If

                    If (txtFrom_Where.Text.Length > 0) Then
                        'MyCLS.clsXLSOperations.DataSetToExcel(txtDatabase.Text + "_" + txtRowTo.Text & ".xls", dsAll)
                        objclsXLSOperations.generateExcel(dsAll.Tables(0), "Records", "", False, False, txtDatabase.Text + "_" + txtRowTo.Text + ".xlsx")
                        'System.Diagnostics.Process.Start(txtRowTo.Text & ".xls")
                    Else
                        'MyCLS.clsXLSOperations.DataSetToExcel(txtDatabase.Text + "_" + cboTableDest.Items(i).ToString() + ".xls", dsAll)
                        objclsXLSOperations.generateExcel(dsAll.Tables(0), cboTableDest.Items(i).ToString(), "", False, False, IIf(cboTableDest.Items(i).ToString().Length > 0, txtDatabase.Text + "_" + cboTableDest.Items(i).ToString(), txtDatabase.Text + "_" + "Records") + ".xlsx")
                        'System.Diagnostics.Process.Start(My.Application.Info.DirectoryPath() & "\Export\" & cboTableDest.Text + ".xls")
                    End If
                    dsAll.Dispose()
                    dsAll = Nothing
                Next
            End If
            objclsXLSOperations = Nothing
            MsgBox("Done")
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex.ToString(), "DisplayRecords")
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
            Try
                ConDest.ConnectionString = CreateConnString()
                ConDest.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
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

    'Sub FillListWithTablesFromSource()
    '    Try
    '        MyCLS.clsCOMMON.SetCon(ConSource)

    '        Dim strSheets As String() = MyCLS.clsDBOperations.GetTables()

    '        cboTableSource.Items.Clear()
    '        For i As Int16 = 0 To strSheets.Length - 1
    '            cboTableSource.Items.Add(strSheets(i))
    '        Next
    '    Catch ex As Exception

    '    End Try
    'End Sub
    Sub FillListWithTablesFromDest()
        Try
            MyCLS.clsCOMMON.SetCon(ConDest)

            Dim strSheets As String() = MyCLS.clsDBOperations.GetTables()

            cboTableDest.Items.Clear()
            For i As Int16 = 0 To strSheets.Length - 1
                cboTableDest.Items.Add(strSheets(i))
            Next
        Catch ex As Exception

        End Try
    End Sub
    'Sub FillListWithColumnsFromSource()
    '    Try
    '        MyCLS.clsCOMMON.SetCon(ConSource)
    '        Dim strColumns As String(,) = MyCLS.clsDBOperations.GetColumns("[" & Replace(cboTableSource.Text, "'", "") & "]")

    '        LstColumnsSource.Items.Clear()
    '        For i As Int16 = 0 To strColumns.Length - 1
    '            LstColumnsSource.Items.Add(strColumns(i, 0))
    '        Next
    '    Catch ex As Exception

    '    End Try
    'End Sub
    Sub FillListWithColumnsFromDest()
        Try
            MyCLS.clsCOMMON.SetCon(ConDest)
            Dim strColumns As String(,) = MyCLS.clsDBOperations.GetColumns("[" & Replace(cboTableDest.Text, "'", "") & "]")

            LstColumnsDest.Items.Clear()
            For i As Int16 = 0 To strColumns.Length - 1
                LstColumnsDest.Items.Add(strColumns(i, 0))
            Next
        Catch ex As Exception

        End Try
    End Sub

    'Sub FillSheets()
    '    Try
    '        Call OpenConnectionSrc()
    '        Call FillListWithTablesFromSource()
    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Public Sub GetDataFromSource()
    '    Try
    '        'Dim ExcelCommand As New System.Data.OleDb.OleDbCommand("SELECT INTO [ODBC Driver={SQL Server};Server=tsi_dev_02;Database=ndhhs_updated;uid=sa;pwd=sa123].[tblOutstanding] FROM [Sheet1$]", Olecn)            
    '        'Dim ExcelCommand As New OleDbCommand("SELECT * FROM [" & strSheetName & "$]", Olecn)
    '        Dim sQ As String = CreateXlsQry(Replace(cboTableSource.Text, "'", ""))
    '        Dim ExcelCommand As New OleDbCommand(sQ, ConSource)

    '        Dim Rs As OleDbDataReader = ExcelCommand.ExecuteReader()

    '        '***From Mala Mam - SQLBULKCOPY ************
    '        sqlBulk = New SqlBulkCopy(ConDestSql)
    '        sqlBulk.DestinationTableName = cboTableDest.Text
    '        'sqlBulk.ColumnMappings.Add("ISBN", "ISBN")
    '        'sqlBulk.ColumnMappings.Add("TITLE", "TITLE")

    '        For i As Int16 = 0 To lstColumnsToExclude.Items.Count - 1
    '            'MsgBox(strColumnsSOURCEMapped(i).ToString & " : " & strColumnsDESTMapped(i).ToString)
    '            sqlBulk.ColumnMappings.Add(strColumnsSOURCEMapped(i).ToString, strColumnsDESTMapped(i).ToString)
    '        Next

    '        '*
    '        'sqlBulk.BatchSize = 2
    '        'MyCLS.clsCOMMON.SetCon(ConSource)            

    '        'While Rs.Read()
    '        '    'For i As Long = 0 To MyCLS.clsCOMMON.fnQuerySelect1Value("Select Count(*) From [" & Replace(cboTableSource.Text, "'", "") & "]", "Number")
    '        '    Debug.Print(Rs("ISBN") & " : " & Rs(1))
    '        '    'Next
    '        'End While
    '        '********************
    '        'Dim ds As New DataSet
    '        'MyCLS.clsCOMMON.prcQuerySelectDS(ds, sQ, Replace(cboTableSource.Text, "'", ""))
    '        'For i As Long = 0 To ds.Tables(0).Rows.Count - 1
    '        '    Debug.Print(ds.Tables(0).Rows(i)("ISBN").ToString() & " : " & ds.Tables(0).Rows(i)(1).ToString())
    '        'Next

    '        'MyCLS.clsCOMMON.SetCon(ConDestSql)
    '        '**

    '        sqlBulk.WriteToServer(Rs)
    '        MsgBox("Data transfer to sql database successfully")
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '        'MYCLS.strGlobalErrorInfo = "Query is : " & TruncateTable
    '        MyCLS.strGlobalErrorInfo = MyCLS.strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
    '        MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
    '        MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.Data)
    '        MyCLS.clsCOMMON.fnWrite2LOG(MyCLS.strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
    '    End Try
    'End Sub

    Function CreateQry(ByVal TableName As String) As String
        Dim strQ As String = ""
        Try
            'MyCLS.clsCOMMON.SetCon(ConSource)
            ''strQ = "SELECT Top " & MyCLS.clsCOMMON.fnQuerySelect1Value("Select Count(*) From [" & TableName & "]", "Number") & " "
            strQ = "SELECT "
            If (txtRowFrom.Text.Length > 0) Then
                strQ = strQ & " * FROM (Select row_number() over (order by " + LstColumnsDest.Items(0) + "," + LstColumnsDest.Items(0) + ") AS line_no,"
                For i As Int16 = 0 To LstColumnsDest.Items.Count - 1
                    If LstColumnsDest.GetItemChecked(i) = True Then
                        strQ = strQ & "[" & LstColumnsDest.Items(i) & "],"
                        '"[ISBN],[TITLE] FROM [" & strSheetName & "$]"
                    End If
                Next
                strQ = Mid(strQ, 1, Len(strQ) - 1) & " From [" & TableName & "]" & " " & txtFrom_Where.Text & ") as x Where x.line_no between " + txtRowFrom.Text + " and " + txtRowTo.Text
            Else
                For i As Int16 = 0 To LstColumnsDest.Items.Count - 1
                    If LstColumnsDest.GetItemChecked(i) = True Then
                        strQ = strQ & "[" & LstColumnsDest.Items(i) & "],"
                        '"[ISBN],[TITLE] FROM [" & strSheetName & "$]"
                    End If
                Next
                strQ = Mid(strQ, 1, Len(strQ) - 1) & " From [" & TableName & "]" & " " & txtFrom_Where.Text
            End If
            'MyCLS.clsCOMMON.SetCon(ConDest)
        Catch ex As Exception

        End Try
        Return strQ
    End Function
    Function CreateQryAllColumns(ByVal TableName As String) As String
        Dim strQ As String = ""
        Try
            'MyCLS.clsCOMMON.SetCon(ConSource)
            ''strQ = "SELECT Top " & MyCLS.clsCOMMON.fnQuerySelect1Value("Select Count(*) From [" & TableName & "]", "Number") & " "
            strQ = "SELECT * From [" & TableName & "]" & " " & txtFrom_Where.Text
            'MyCLS.clsCOMMON.SetCon(ConDest)
        Catch ex As Exception

        End Try
        Return strQ
    End Function

    Private Sub LstColumnsDest_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles LstColumnsDest.ItemCheck
        Try
            If Resetting = True Then Exit Sub
            If LstColumnsDest.GetItemCheckState(e.Index) = CheckState.Checked Then
                ''strColumnsDESTMapped(intColumnsSOURCEMappedIndex - 1).Remove(0)
                'strColumnsDESTMapped(intColumnsSOURCEMappedIndex - 1) = Nothing
                ''intColumnsSOURCEMappedIndex -= 1
                LstColumnsDest.SetSelected(e.Index, True)
                'MsgBox("Please Click on Reset Mapping and Then Start Again!", MsgBoxStyle.Critical, "Can't Remove")                
            Else
                strColumnsDESTMapped(intColumnsSOURCEMappedIndex) = LstColumnsDest.Items(e.Index).ToString()

                lstColumnsToExclude.Items.Add(strColumnsSOURCEMapped(intColumnsSOURCEMappedIndex) & " : " & strColumnsDESTMapped(intColumnsSOURCEMappedIndex))

                intColumnsSOURCEMappedIndex += 1
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdResetMapping_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Try
            'ReDim Preserve strColumnsSOURCEMapped(LstColumnsSource.Items.Count)
            ReDim Preserve strColumnsDESTMapped(LstColumnsDest.Items.Count)

            intColumnsSOURCEMappedIndex = 0

            lstColumnsToExclude.Items.Clear()

            Resetting = True
            MyCLS.clsControls.prcListUnCheckAll(LstColumnsDest)
            'MyCLS.clsCOMMON.prcListUnCheckAll(LstColumnsSource)
            Resetting = False
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectAll.CheckedChanged
        If chkSelectAll.Checked = True Then
            MyCLS.clsControls.prcListCheckAll(LstColumnsDest)
        Else
            MyCLS.clsControls.prcListUnCheckAll(LstColumnsDest)
        End If
    End Sub
End Class