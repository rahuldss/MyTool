Imports System.Data.OleDb

Public Class frmPropertCreater
    Dim strLOGINSFilePath As String
    Dim strLOGINSFileName As String

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try            
            'Me.MdiParent = MDI

            gbDBOptionsAccess.Top = 94
            gbDBOptionsSql.Top = 94
            gbDBOptionsOracle.Top = 94


            ''txtFile.Text = My.Application.Info.DirectoryPath() & "\SkyP.mdb"
            'txtFile.Text = "D:\Narender\Projects\VB.NET\2008\InTakeApp\InTakeApp\bin\Debug\Data\db.mdb"
            'txtServer.Text = "tsi_dev_02"
            'txtUID.Text = ""
            'txtPassword.Text = ""
            'txtDatabase.Text = ""

            '***GET DATABASE SETTINGS***
            Dim strDBDetailsSplit() As String = MyCLS.clsCOMMON.GetSettings()
            txtFile.Text = strDBDetailsSplit(0)
            txtServer.Text = strDBDetailsSplit(1)
            txtUID.Text = strDBDetailsSplit(2)
            txtPassword.Text = strDBDetailsSplit(3)
            txtDatabase.Text = strDBDetailsSplit(4)
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    Private Sub cmdSelectFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectFile.Click
        Try
            If Len(txtFile.Text) > 0 Then
                OFDialog1.FileName = txtFile.Text
            Else
                OFDialog1.InitialDirectory = Environment.SpecialFolder.Desktop
            End If

            OFDialog1.ShowDialog(Me)
            '.xls
            'If Len(OFDialog1.FileName) > 0 And (OFDialog1.FileName <> "*.xls" Or OFDialog1.FileName <> "*.xlsx") Then
            txtFile.Text = OFDialog1.FileName.ToString
            strLOGINSFilePath = Mid(OFDialog1.FileName, 1, Len(OFDialog1.FileName) - InStr(StrReverse(OFDialog1.FileName), "\"))

            strLOGINSFileName = Mid(OFDialog1.FileName, Len(OFDialog1.FileName) - InStr(StrReverse(OFDialog1.FileName), "\") + 2, InStr(StrReverse(OFDialog1.FileName), "\") - 1)
            strLOGINSFileName = Mid(strLOGINSFileName, 1, Len(strLOGINSFileName) - 4)
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    Private Sub cmdCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreate.Click
        Try
            lblMSG.Text = "Validating..."
            If fnValidate() = True Then
                '***SAVE DATABASE SETTINGS***
                MyCLS.clsCOMMON.SaveSettings(txtFile.Text, txtServer.Text, txtUID.Text, txtPassword.Text, txtDatabase.Text)

                cmdCreate.Enabled = False
                cmdCancel.Enabled = False
                'DELETE DIR
                Try
                    IO.Directory.Delete("C:\_CODE", True)
                    'IO.Directory.Delete("C:\_CODE\LIB", True)
                    'IO.Directory.Delete("C:\_CODE\DAL", True)
                    'IO.Directory.Delete("C:\_CODE\SQL", True)
                Catch ex As Exception

                End Try
                'CREATE DIR
                Try
                    lblMSG.Text = "Dir Creation..."
                    IO.Directory.CreateDirectory("C:\_CODE")
                    IO.Directory.CreateDirectory("C:\_CODE\LIB")
                    IO.Directory.CreateDirectory("C:\_CODE\DAL")
                    IO.Directory.CreateDirectory("C:\_CODE\SQL")
                    IO.Directory.CreateDirectory("C:\_CODE\Columns")
                Catch ex As Exception

                End Try

                'CREATE CONNECTION STRING
                lblMSG.Text = "Creating Conn String..."
                MyCLS.strConnStringOLEDB = CreateConnString()
                MyCLS.strConnStringSQLCLIENT = CreateConnString().Replace("Provider=SQLOLEDB;", "")

                lblMSG.Text = "Opening Connection..."
                MyCLS.clsCOMMON.ConOpen(False)

                If Len(MyCLS.strGlobalErrorInfo) > 0 Then
                    cmdCreate.Enabled = True
                    cmdCancel.Enabled = True
                    lblMSG.Text = "Finished!"
                    'MsgBox("Not Done!", MsgBoxStyle.Information, "Not Completed")
                    Exit Sub
                End If

                ''GET ALL THE TABLES
                'Dim str As String() = MyCLS.clsDBOperations.GetTables()

                'GET DETAILED DATABASE IN A CLASS
                Dim dbInfo As New List(Of MyCLS.clsTables)
                lblMSG.Text = "Fetching Data From Database..."
                dbInfo = MyCLS.clsDBOperations.FillDetails()

                'CREATE PROPERTIES
                If optCSMVCAPI.Checked = True Then
                    lblMSG.Text = "Writing CS MVC Files..."
                    Call WritePropertyInCS_MVCAPI(dbInfo)
                    Call WriteDALInCSSqlClient_UPDATED_MVC(dbInfo)
                    lblMSG.Text = "Writing SP Files..."
                    Call WriteStoredProc_UPDATED(dbInfo)
                ElseIf optCSMvc.Checked = True Then
                    lblMSG.Text = "Writing CS MVC Files..."
                    Call WritePropertyInCS_MVC(dbInfo)
                    Call WriteDALInCSSqlClient_UPDATED_MVC(dbInfo)
                    lblMSG.Text = "Writing SP Files..."
                    Call WriteStoredProc_UPDATED(dbInfo)
                ElseIf optCS.Checked = True Then
                    lblMSG.Text = "Writing CS Files..."
                    Call WritePropertyInCS_UPDATED(dbInfo)
                    'Call WriteDALInCS(dbInfo)
                    If optMSAccess.Checked = True Then
                        Call WriteDALInCS4Access(dbInfo)
                    Else
                        If optOleDb.Checked = True Then
                            Call WriteDALInCS(dbInfo)
                        Else
                            Call WriteDALInCSSqlClient_UPDATED(dbInfo)
                        End If
                        lblMSG.Text = "Writing SP Files..."
                        Call WriteStoredProc_UPDATED(dbInfo)
                    End If
                ElseIf optVB6.Checked = True Then
                    lblMSG.Text = "Writing VB6.0 Files..."
                    Call WritePropertyInVB6(dbInfo)
                    Call WriteDALInVB6(dbInfo)
                    lblMSG.Text = "Writing SP Files..."
                    Call WriteStoredProc_UPDATED(dbInfo)
                Else
                    lblMSG.Text = "Writing VB Files..."
                    Call WritePropertyInVB_UPDATED(dbInfo)
                    If optMSAccess.Checked = True Then
                        Call WriteDALInVB4Access(dbInfo)
                    Else
                        If optOleDb.Checked = True Then
                            Call WriteDALInVB(dbInfo)
                        Else
                            Call WriteDALInVBSqlClient_UPDATED(dbInfo)
                        End If
                        lblMSG.Text = "Writing SP Files..."
                        Call WriteStoredProc_UPDATED(dbInfo)
                    End If
                End If

                If chkColList.Checked = True Then
                    Call CreateColList(dbInfo)
                End If

                MyCLS.clsCOMMON.ConClose()
                lblMSG.Text = "Opening Files..."
                Shell("Explorer C:\_CODE", AppWinStyle.MaximizedFocus)

                cmdCreate.Enabled = True
                cmdCancel.Enabled = True
                lblMSG.Text = "Finished!"
                'MsgBox("Done!", MsgBoxStyle.Information, "Completed")
            End If
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    Private Sub optMSSql_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles optMSSql.Click
        Try
            gbDBOptionsAccess.Visible = False
            gbDBOptionsSql.Visible = True
            gbDBOptionsOracle.Visible = False
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    Private Sub optMSAccess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles optMSAccess.Click
        Try
            gbDBOptionsSql.Visible = False
            gbDBOptionsAccess.Visible = True
            gbDBOptionsOracle.Visible = False
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    Private Sub optOracle_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles optOracle.Click
        Try
            gbDBOptionsAccess.Visible = False
            gbDBOptionsSql.Visible = False
            gbDBOptionsOracle.Visible = True
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    Private Sub optMSSql_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMSSql.CheckedChanged
        If optMSSql.Checked = True Then
            gbDBConnType.Visible = True
        Else
            gbDBConnType.Visible = False
        End If
    End Sub

    Private Sub chkConnStr_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkConnStr.CheckedChanged
        If chkConnStr.Checked = True Then
            lblServerOra.Text = "Connection String"
        Else
            lblServerOra.Text = "Server"
        End If
    End Sub

    Private Sub chkDBFile_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDBFile.CheckedChanged
        'txtServer.Text = "Provider=SQLOLEDB;AttachDbFilename=D:\Narender\All_DataBases\TMPDATABASE.mdf;Initial Catalog=TMPDATABASE;Trusted_Connection=Yes"
        'txtServer.Text = "Driver={SQL Native Client};Server=.\SQLExpress;AttachDbFilename=D:\Narender\All_DataBases\TMPDATABASE.mdf;Database=TMPDATABASE;Trusted_Connection=Yes;"
        txtServer.Text = "Server=TSI_DEV_02\SQLExpress;AttachDbFilename=D:\Narender\All_DataBases\TMPDATABASE.mdf;Database=TMPDATABASE;Trusted_Connection=Yes;Integrated Security=True;Provider=SQLOLEDB;"
    End Sub

    Private Sub lnkConnStr_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkConnStr.LinkClicked
        frmConnStr.Show()
    End Sub

    Private Sub cmdSelectServer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectServer.Click
        Try
            frmSelectServer.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub


    Function fnValidate()
        If optMSAccess.Checked = True Then
            If txtFile.Text = "" Then
                MsgBox("Please Select DB Location", MsgBoxStyle.Information, "DB Required")
                txtFile.Focus()
                fnValidate = False
                Exit Function
            Else
                If System.IO.File.Exists(txtFile.Text) = False Then
                    MsgBox("File does not exists!", MsgBoxStyle.Critical)
                    cmdSelectFile.Focus()
                    fnValidate = False
                    Exit Function
                ElseIf UCase(MyCLS.clsCOMMON.fnGetExtension(txtFile.Text)) <> "MDB" Then
                    MsgBox("Invalid DB File!", MsgBoxStyle.Critical)
                    cmdSelectFile.Focus()
                    fnValidate = False
                    Exit Function
                End If
            End If
        ElseIf optMSSql.Checked = True Then
            If txtServer.Text = "" Then
                MsgBox("Please Enter Server", MsgBoxStyle.Information, "Server Required")
                txtServer.Focus()
                fnValidate = False
                Exit Function
            ElseIf txtUID.Text = "" Then
                MsgBox("Please Enter UID", MsgBoxStyle.Information, "UID Required")
                txtUID.Focus()
                fnValidate = False
                Exit Function
            ElseIf txtPassword.Text = "" Then
                MsgBox("Please Enter Password", MsgBoxStyle.Information, "Password Required")
                txtPassword.Focus()
                fnValidate = False
                Exit Function
            ElseIf txtDatabase.Text = "" Then
                MsgBox("Please Enter Database", MsgBoxStyle.Information, "Database Required")
                txtDatabase.Focus()
                fnValidate = False
                Exit Function
            End If
        Else

        End If
        fnValidate = True
    End Function

    Function CreateConnString() As String
        Dim strConnStr As String
        If optMSAccess.Checked = True Then
            strConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtFile.Text & ";Persist Security Info=True"
        ElseIf optMSSql.Checked = True Then
            If chkDBFile.Checked = True Then
                strConnStr = txtServer.Text
            Else
                strConnStr = "UID=" & txtUID.Text & ";Password=" & txtPassword.Text & ";Data Source=" & txtServer.Text & ";Initial Catalog=" & txtDatabase.Text & ";Provider=SQLOLEDB.1;"
            End If
        Else
            If chkConnStr.Checked = True Then
                strConnStr = txtServerOra.Text
            Else
                strConnStr = "Provider=" & txtProviderOra.Text & ";Data Source=" & txtServerOra.Text & ";User Id=" & txtUIDOra.Text & ";Password=" & txtPasswordOra.Text & ";"
            End If
        End If
        Return strConnStr
    End Function

    ''' <summary>
    ''' CREATE PROPERTIES LIBRARY IN VB
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks></remarks>
    Sub WritePropertyInVB(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassLIB As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassLIB = "Namespace NDS.LIB" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "Public Class LIB" & dbInfo(i).TABLENAME & vbCrLf & vbCrLf
                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    'MsgBox(dbInfo(i).TABLENAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE)
                    strClassLIB = strClassLIB & "Private _" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " as " & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE & vbCrLf
                    strClassLIB = strClassLIB & "Public Property " & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " as " & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE & vbCrLf
                    strClassLIB = strClassLIB & "Get" & vbCrLf
                    strClassLIB = strClassLIB & "Return _" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf
                    strClassLIB = strClassLIB & "End Get" & vbCrLf
                    strClassLIB = strClassLIB & "Set(ByVal value As " & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE & ")" & vbCrLf
                    strClassLIB = strClassLIB & "_" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = value" & vbCrLf
                    strClassLIB = strClassLIB & "End Set" & vbCrLf
                    strClassLIB = strClassLIB & "End Property" & vbCrLf & vbCrLf

                    Application.DoEvents()
                Next
                strClassLIB = strClassLIB & "End Class" & vbCrLf & vbCrLf & vbCrLf

                strClassLIB = strClassLIB & "<Serializable()> _" & vbCrLf
                strClassLIB = strClassLIB & "Public Class LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassLIB = strClassLIB & " Inherits List(Of LIB" & dbInfo(i).TABLENAME & ")" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "End Class" & vbCrLf

                strClassLIB = strClassLIB & "End Namespace"

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\LIB\LIB" & dbInfo(i).TABLENAME & ".vb")
                MyCLS.clsFileHandling.WriteFile(strClassLIB)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    Sub WritePropertyInVB_UPDATED(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassLIB As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassLIB = "Namespace NDS.LIB" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "Public Class LIB" & dbInfo(i).TABLENAME & vbCrLf & vbCrLf
                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    'MsgBox(dbInfo(i).TABLENAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE)
                    strClassLIB = strClassLIB & "Private _" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " as " & dbInfo(i).COLUMNDETAILS(j).COLDataType.Replace("[]", "()") & vbCrLf
                    strClassLIB = strClassLIB & "Public Property " & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " as " & dbInfo(i).COLUMNDETAILS(j).COLDataType.Replace("[]", "()") & vbCrLf
                    strClassLIB = strClassLIB & "Get" & vbCrLf
                    strClassLIB = strClassLIB & "Return _" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf
                    strClassLIB = strClassLIB & "End Get" & vbCrLf
                    strClassLIB = strClassLIB & "Set(ByVal value As " & dbInfo(i).COLUMNDETAILS(j).COLDataType.Replace("[]", "()") & ")" & vbCrLf
                    strClassLIB = strClassLIB & "_" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = value" & vbCrLf
                    strClassLIB = strClassLIB & "End Set" & vbCrLf
                    strClassLIB = strClassLIB & "End Property" & vbCrLf & vbCrLf

                    Application.DoEvents()
                Next
                strClassLIB = strClassLIB & "End Class" & vbCrLf & vbCrLf & vbCrLf

                strClassLIB = strClassLIB & "<Serializable()> _" & vbCrLf
                strClassLIB = strClassLIB & "Public Class LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassLIB = strClassLIB & " Inherits List(Of LIB" & dbInfo(i).TABLENAME & ")" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "End Class" & vbCrLf

                strClassLIB = strClassLIB & "End Namespace"

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\LIB\LIB" & dbInfo(i).TABLENAME & ".vb")
                MyCLS.clsFileHandling.WriteFile(strClassLIB)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    ''' <summary>
    ''' CREATE DATA ACCESS LAYER IN VB
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks>CREATES SELECT, SELECT BY VALUE AND INSERT FUNCTIONs</remarks>
    Sub WriteDALInVB(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassDAL As String = ""
            Dim strClassDALStart As String = ""
            Dim strClassDALEnd As String = ""
            Dim strClassDALSelect As String = ""
            Dim strClassDALSelectByValue As String = ""
            Dim strClassDALInsert As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassDAL = ""
                strClassDALStart = ""
                strClassDALEnd = ""
                strClassDALSelect = ""
                strClassDALSelectByValue = ""
                strClassDALInsert = ""

                strClassDALStart = "Imports NDS.LIB" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data.OleDb" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Namespace NDS.DAL" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Public Class DAL" & dbInfo(i).TABLENAME & vbCrLf & vbCrLf



                strClassDALSelectByValue = strClassDALSelectByValue & "Public Function Get" & dbInfo(i).TABLENAME & "Details(ByVal Id As Int16) as LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim ds As New DataSet" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim objParamList As New List(Of OleDbParameter)()" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim  clsESPSql as New MyCLS.clsExecuteStoredProcSql" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Try" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "objParamList.Add(New OleDbParameter(""@Id"", Id))" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList)" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf

                strClassDALSelect = strClassDALSelect & "Public Function Get" & dbInfo(i).TABLENAME & "Details() as LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim ds As New DataSet" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim  clsESPSql as New MyCLS.clsExecuteStoredProcSql" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Try" & vbCrLf
                strClassDALSelect = strClassDALSelect & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf

                strClassDALInsert = strClassDALInsert & "Public Function Insert" & dbInfo(i).TABLENAME & "(ByVal objLIB" & dbInfo(i).TABLENAME & " As LIB" & dbInfo(i).TABLENAME & ", ByRef Result As Int16) As String()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim strOutParamValues As String()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim objParamList As New List(Of OleDbParameter)()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim objParamListOut As New List(Of OleDbParameter)()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Try" & vbCrLf

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    strClassDALInsert = strClassDALInsert & "objParamList.Add(New OleDbParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "))" & vbCrLf


                    Application.DoEvents()
                Next

                strClassDALSelectByValue = strClassDALSelectByValue & "Next" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Catch ex As Exception" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Try" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Return objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Function" & vbCrLf & vbCrLf

                strClassDALSelect = strClassDALSelect & "Next" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Catch ex As Exception" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Try" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Return objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Function" & vbCrLf & vbCrLf

                strClassDALInsert = strClassDALInsert & "objParamListOut.Add(New OleDbParameter(""@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & """, OleDbType." & dbInfo(i).COLUMNDETAILS(0).COLUMNTYPE & "))" & vbCrLf
                strClassDALInsert = strClassDALInsert & "'MyCLS.clsCOMMON.ConOpen(true)" & vbCrLf
                strClassDALInsert = strClassDALInsert & "strOutParamValues = MyCLS.clsExecuteStoredProc.ExecuteSPNonQueryOutPut(""SP_Insert" & dbInfo(i).TABLENAME & """, objParamList, objParamListOut, Result)" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Result = 1" & vbCrLf
                strClassDALInsert = strClassDALInsert & "'MyCLS.clsCOMMON.ConClose()" & vbCrLf & vbCrLf
                strClassDALInsert = strClassDALInsert & "Catch ex As Exception" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Result = -1" & vbCrLf
                strClassDALInsert = strClassDALInsert & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Try" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Return strOutParamValues" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Function" & vbCrLf & vbCrLf


                strClassDALEnd = strClassDALEnd & "End Class" & vbCrLf
                strClassDALEnd = strClassDALEnd & "End Namespace"

                strClassDAL = strClassDALStart & strClassDALSelect & vbCrLf & strClassDALSelectByValue & vbCrLf & strClassDALInsert & strClassDALEnd

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".vb")
                MyCLS.clsFileHandling.WriteFile(strClassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    ''' <summary>
    ''' CREATE DATA ACCESS LAYER IN VB WITH TP AND SQL CLIENT
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks>CREATES SELECT, SELECT BY VALUE AND INSERT FUNCTIONs</remarks>
    Sub WriteDALInVBSqlClient(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassDAL As String = ""
            Dim strClassDALStart As String = ""
            Dim strClassDALEnd As String = ""
            Dim strClassDALSelect As String = ""
            Dim strClassDALSelectByValue As String = ""
            Dim strClassDALInsert As String = ""
            Dim strClassDALInsert2Use As String = ""
            Dim strClassDALSelect2Use As String = ""
            Dim strClassDALSelectByValue2Use As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassDAL = ""
                strClassDALStart = ""
                strClassDALEnd = ""
                strClassDALSelect = ""
                strClassDALSelectByValue = ""
                strClassDALInsert = ""
                strClassDALInsert2Use = ""
                strClassDALSelect2Use = ""
                strClassDALSelectByValue2Use = ""

                strClassDALStart = "Imports NDS.LIB" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data.OleDb" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data.SqlClient" & vbCrLf & vbCrLf
                'strClassDALStart = strClassDALStart & "Imports MyClsWin" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Namespace NDS.DAL" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Public Class DAL" & dbInfo(i).TABLENAME & vbCrLf & vbCrLf


                strClassDALSelectByValue = strClassDALSelectByValue & "''' <summary>" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "''' Accepts=TransportationPacket, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "''' </summary>" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "''' <returns></returns>" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "''' <remarks></remarks>" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Public Function Get" & dbInfo(i).TABLENAME & "Details(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim ds As New DataSet" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim objParamList As New List(Of SqlParameter)()" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim  clsESPSql as New MyCLS.clsExecuteStoredProcSql" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Try" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "objParamList.Add(New SqlParameter(""@Id"", Packet.MessagePacket))" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList)" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf
                '***Connection String***
                strClassDALSelect2Use = strClassDALSelect2Use & "'PUT IT IN LOAD EVENTS" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'MyCLS.strConnStringOLEDB = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;Provider=SQLOLEDB.1""" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'MyCLS.strConnStringSQLCLIENT = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;""" & vbCrLf & vbCrLf
                '***Connection String***
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'*******COPY IT TO USE BELOW FUNCTION - SELECT BY ID************" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'Try" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    Dim objDAL" & dbInfo(i).TABLENAME & " As New DAL" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    Dim tp As New MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    Dim ds As New DataSet" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    tp.MessagePacket = 1    'ID to be Passed" & vbCrLf & vbCrLf

                strClassDALSelect = strClassDALSelect & "''' <summary>" & vbCrLf
                strClassDALSelect = strClassDALSelect & "''' Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strClassDALSelect = strClassDALSelect & "''' </summary>" & vbCrLf
                strClassDALSelect = strClassDALSelect & "''' <returns></returns>" & vbCrLf
                strClassDALSelect = strClassDALSelect & "''' <remarks></remarks>" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Public Function Get" & dbInfo(i).TABLENAME & "Details() as MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim ds As New DataSet" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim Packet As New MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim clsESPSql As New MyCLS.clsExecuteStoredProcSql" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Try" & vbCrLf
                strClassDALSelect = strClassDALSelect & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'*******COPY IT TO USE BELOW FUNCTION - SELECT ALL************" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'Try" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    Dim objDAL" & dbInfo(i).TABLENAME & " As New DAL" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    Dim tp As New MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    Dim ds As New DataSet" & vbCrLf & vbCrLf

                strClassDALInsert = strClassDALInsert & "''' <summary>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' Accepts=Packet, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' </summary>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' <param name=""Packet""></param>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' <returns></returns>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' <remarks></remarks>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Public Function Insert" & dbInfo(i).TABLENAME & "(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim strOutParamValues As String()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim objParamList As New List(Of SqlParameter)()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim objParamListOut As New List(Of SqlParameter)()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim Result As Int16 = 0" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Try" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim objLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALInsert = strClassDALInsert & "objLIB" & dbInfo(i).TABLENAME & " = Packet.MessagePacket" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'*******COPY IT TO USE BELOW FUNCTION - INSERT************" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'Try" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    Dim objLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    Dim objDAL" & dbInfo(i).TABLENAME & " As New DAL" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    Dim tp As New MyCLS.TransportationPacket" & vbCrLf & vbCrLf


                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    strClassDALInsert = strClassDALInsert & "objParamList.Add(New SqlParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "))" & vbCrLf

                    'strClassDALInsert2Use = strClassDALInsert2Use & "'    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & IIf(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE = "byte()", "MyCLS.clsImaging.PictureBoxToByteArray()", """") & vbCrLf
                    'strClassDALSelect2Use = strClassDALSelect2Use & "''    objLIB" & dbInfo(i).TABLENAME & "Listing(" & j & ")." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & vbCrLf
                    'strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "''    objLIB" & dbInfo(i).TABLENAME & "Listing(" & j & ")." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & vbCrLf
                    strClassDALInsert2Use = strClassDALInsert2Use & "'    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & IIf(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE = "byte()", "MyCLS.clsImaging.PictureBoxToByteArray()", "txt.Text") & vbCrLf
                    strClassDALSelect2Use = strClassDALSelect2Use & "'' txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing(0)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf
                    strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'' txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing(0)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf

                    Application.DoEvents()
                Next

                strClassDALSelectByValue = strClassDALSelectByValue & "Next" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageId = 1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Else" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageId = -1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageResultsetDS = ds" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Catch ex As Exception" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageId = -1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Try" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Return Packet" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Function" & vbCrLf & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details(tp)" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    If tp.MessageId = 1 Then" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'        objLIB" & dbInfo(i).TABLENAME & "Listing = tp.MessageResultset" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'        ds = tp.MessageResultsetDS" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'        MyCLS.clsImaging.ByteArray2Image(,)" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'        MsgBox(objLIB" & dbInfo(i).TABLENAME & "Listing(0))" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    End If" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'Catch ex As Exception" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    MsgBox(ex.Message)" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'End Try" & vbCrLf

                strClassDALSelect = strClassDALSelect & "Next" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageId = 1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Else" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageId = -1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageResultsetDS = ds" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Catch ex As Exception" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageId = -1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Try" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Return Packet" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Function" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details()" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    If tp.MessageId = 1 Then" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'        objLIB" & dbInfo(i).TABLENAME & "Listing = tp.MessageResultset" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'        ds = tp.MessageResultsetDS" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'        MyCLS.clsImaging.ByteArray2Image(,)" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'        MsgBox(objLIB" & dbInfo(i).TABLENAME & "Listing(0))" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    End If" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'Catch ex As Exception" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    MsgBox(ex.Message)" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'End Try" & vbCrLf

                strClassDALInsert = strClassDALInsert & "objParamListOut.Add(New SqlParameter(""@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & """, SqlDbType." & MyCLS.clsDBOperations.GetDBTypeValue4SqlDbTypes(dbInfo(i).COLUMNDETAILS(0).COLUMNTYPE) & "))" & vbCrLf
                'strClassDALInsert = strClassDALInsert & "'MyCLS.clsCOMMON.ConOpen(true)" & vbCrLf
                strClassDALInsert = strClassDALInsert & "strOutParamValues = MyCLS.clsExecuteStoredProcSql.ExecuteSPNonQueryOutPut(""SP_Insert" & dbInfo(i).TABLENAME & """, objParamList, objParamListOut, Packet.MessageId)" & vbCrLf
                'strClassDALInsert = strClassDALInsert & "Packet.MessageId = Result" & vbCrLf
                'strClassDALInsert = strClassDALInsert & "'MyCLS.clsCOMMON.ConClose()" & vbCrLf & vbCrLf
                strClassDALInsert = strClassDALInsert & "Packet.MessageResultset = strOutParamValues" & vbCrLf & vbCrLf
                strClassDALInsert = strClassDALInsert & "Catch ex As Exception" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Packet.MessageId = -1" & vbCrLf
                strClassDALInsert = strClassDALInsert & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Try" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Return Packet" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Function" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    tp.MessagePacket = objLIB" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    tp = objDAL" & dbInfo(i).TABLENAME & ".Insert" & dbInfo(i).TABLENAME & "(tp)" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    If tp.MessageId > -1 Then" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'        Dim strOutParamValues As String() = tp.MessageResultset" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'        MsgBox(strOutParamValues(0))" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    End If" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'Catch ex As Exception" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    MsgBox(ex.Message)" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'End Try" & vbCrLf


                strClassDALEnd = strClassDALEnd & "End Class" & vbCrLf
                strClassDALEnd = strClassDALEnd & "End Namespace"

                strClassDAL = strClassDALStart & strClassDALSelect2Use & vbCrLf & strClassDALSelect & vbCrLf & strClassDALSelectByValue2Use & vbCrLf & strClassDALSelectByValue & vbCrLf & strClassDALInsert2Use & vbCrLf & strClassDALInsert & strClassDALEnd

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".vb")
                MyCLS.clsFileHandling.WriteFile(strClassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    Sub WriteDALInVBSqlClient_UPDATED(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassDAL As String = ""
            Dim strClassDALStart As String = ""
            Dim strClassDALEnd As String = ""
            Dim strClassDALSelect As String = ""
            Dim strClassDALSelectByValue As String = ""
            Dim strClassDALInsert As String = ""
            Dim strClassDALInsert2Use As String = ""
            Dim strClassDALSelect2Use As String = ""
            Dim strClassDALSelectByValue2Use As String = ""
            Dim strClassDALDeleteByValue As String = ""
            Dim strClassDALDeleteByValue2Use As String = ""

            For i As Integer = 0 To dbInfo.Count - 1
                strClassDAL = ""
                strClassDALStart = ""
                strClassDALEnd = ""
                strClassDALSelect = ""
                strClassDALSelectByValue = ""
                strClassDALInsert = ""
                strClassDALInsert2Use = ""
                strClassDALSelect2Use = ""
                strClassDALSelectByValue2Use = ""
                strClassDALDeleteByValue = ""
                strClassDALDeleteByValue2Use = ""

                strClassDALStart = "Imports NDS.LIB" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data.OleDb" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data.SqlClient" & vbCrLf & vbCrLf
                'strClassDALStart = strClassDALStart & "Imports MyClsWin" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Namespace NDS.DAL" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Public Class DAL" & dbInfo(i).TABLENAME & vbCrLf & vbCrLf


                strClassDALSelectByValue = strClassDALSelectByValue & "''' <summary>" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "''' Accepts=TransportationPacket, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "''' </summary>" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "''' <returns></returns>" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "''' <remarks></remarks>" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Public Function Get" & dbInfo(i).TABLENAME & "Details(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim ds As New DataSet" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim objParamList As New List(Of SqlParameter)()" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim  clsESPSql as New MyCLS.clsExecuteStoredProcSql" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Try" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "objParamList.Add(New SqlParameter(""@Id"", Packet.MessagePacket))" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList)" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf

                strClassDALDeleteByValue = strClassDALDeleteByValue & "''' <summary>" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "''' Deletes Row By ID " & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "''' </summary>" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "''' <param name=""Packet""></param> " & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "''' <returns></returns>" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "''' <remarks></remarks>" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Public Function Delete" & dbInfo(i).TABLENAME & "(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Dim objParamList As New List(Of SqlParameter)()" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Dim Result As Int16 = 0" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Dim objLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Dim  clsESPSql as New MyCLS.clsExecuteStoredProcSql" & vbCrLf & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Try" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "objLIB" & dbInfo(i).TABLENAME & " = Packet.MessagePacket" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "objParamList.Add(New SqlParameter(""@Id"", Packet.MessagePacket))" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Result = clsESPSql.ExecuteSPNonQuery(""SP_Delete" & dbInfo(i).TABLENAME & """, objParamList)" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Packet.MessageId = Result" & vbCrLf

                '***Connection String***
                strClassDALSelect2Use = strClassDALSelect2Use & "'PUT IT IN LOAD EVENTS" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'MyCLS.strConnStringOLEDB = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;Provider=SQLOLEDB.1""" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'MyCLS.strConnStringSQLCLIENT = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;""" & vbCrLf & vbCrLf
                '***Connection String***
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'*******COPY IT TO USE BELOW FUNCTION - SELECT BY ID************" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'Try" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    Dim objDAL" & dbInfo(i).TABLENAME & " As New DAL" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    Dim tp As New MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    Dim ds As New DataSet" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    tp.MessagePacket = 1    'ID to be Passed" & vbCrLf & vbCrLf

                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'*******COPY IT TO USE BELOW FUNCTION - SELECT BY ID************" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'Try" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'    Dim objLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'    Dim objDAL" & dbInfo(i).TABLENAME & " As New DAL" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'    Dim tp As New MyCLS.TransportationPacket" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'    tp.MessagePacket = 1    'ID to be Passed" & vbCrLf & vbCrLf

                strClassDALSelect = strClassDALSelect & "''' <summary>" & vbCrLf
                strClassDALSelect = strClassDALSelect & "''' Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strClassDALSelect = strClassDALSelect & "''' </summary>" & vbCrLf
                strClassDALSelect = strClassDALSelect & "''' <returns></returns>" & vbCrLf
                strClassDALSelect = strClassDALSelect & "''' <remarks></remarks>" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Public Function Get" & dbInfo(i).TABLENAME & "Details() as MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim ds As New DataSet" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim Packet As New MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim clsESPSql As New MyCLS.clsExecuteStoredProcSql" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Try" & vbCrLf
                strClassDALSelect = strClassDALSelect & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'*******COPY IT TO USE BELOW FUNCTION - SELECT ALL************" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'Try" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    Dim objDAL" & dbInfo(i).TABLENAME & " As New DAL" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    Dim tp As New MyCLS.TransportationPacket" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    Dim ds As New DataSet" & vbCrLf & vbCrLf

                strClassDALInsert = strClassDALInsert & "''' <summary>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' Accepts=Packet, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' </summary>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' <param name=""Packet""></param>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' <returns></returns>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "''' <remarks></remarks>" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Public Function Insert" & dbInfo(i).TABLENAME & "(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim strOutParamValues As String()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim objParamList As New List(Of SqlParameter)()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim objParamListOut As New List(Of SqlParameter)()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim Result As Int16 = 0" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Try" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim objLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALInsert = strClassDALInsert & "objLIB" & dbInfo(i).TABLENAME & " = Packet.MessagePacket" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim clsESPSql As New MyCLS.clsExecuteStoredProcSql" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'*******COPY IT TO USE BELOW FUNCTION - INSERT************" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'Try" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    Dim objLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    Dim objDAL" & dbInfo(i).TABLENAME & " As New DAL" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    Dim tp As New MyCLS.TransportationPacket" & vbCrLf & vbCrLf


                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    strClassDALInsert = strClassDALInsert & "objParamList.Add(New SqlParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "))" & vbCrLf

                    'strClassDALInsert2Use = strClassDALInsert2Use & "'    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & IIf(dbInfo(i).COLUMNDETAILS(j).COLDataType = "byte()", "MyCLS.clsImaging.PictureBoxToByteArray()", """") & vbCrLf
                    'strClassDALSelect2Use = strClassDALSelect2Use & "''    objLIB" & dbInfo(i).TABLENAME & "Listing(" & j & ")." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & vbCrLf
                    'strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "''    objLIB" & dbInfo(i).TABLENAME & "Listing(" & j & ")." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & vbCrLf
                    strClassDALInsert2Use = strClassDALInsert2Use & "'    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & IIf(dbInfo(i).COLUMNDETAILS(j).COLDataType = "Byte[]", "MyCLS.clsImaging.PictureBoxToByteArray()", "txt.Text") & vbCrLf
                    strClassDALSelect2Use = strClassDALSelect2Use & "'' txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing(0)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf
                    strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'' txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing(0)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf

                    Application.DoEvents()
                Next

                strClassDALSelectByValue = strClassDALSelectByValue & "Next" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageId = 1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Else" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageId = -1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageResultsetDS = ds" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Catch ex As Exception" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Packet.MessageId = -1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Try" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Return Packet" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Function" & vbCrLf & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details(tp)" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    If tp.MessageId = 1 Then" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'        objLIB" & dbInfo(i).TABLENAME & "Listing = tp.MessageResultset" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'        ds = tp.MessageResultsetDS" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'        MyCLS.clsImaging.ByteArray2Image(,)" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'        MsgBox(objLIB" & dbInfo(i).TABLENAME & "Listing(0))" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    End If" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'Catch ex As Exception" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'    MsgBox(ex.Message)" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "'End Try" & vbCrLf

                strClassDALDeleteByValue = strClassDALDeleteByValue & "Catch ex As Exception" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Packet.MessageId = -1" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "End Try" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "Return Packet" & vbCrLf
                strClassDALDeleteByValue = strClassDALDeleteByValue & "End Function" & vbCrLf & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'    tp = objDAL" & dbInfo(i).TABLENAME & ".Delete" & dbInfo(i).TABLENAME & "(tp)" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'    If tp.MessageId > 0 Then" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'        Dim strOutParamValues As String() = tp.MessageResultset" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'    End If" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'Catch ex As Exception" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'    MsgBox(ex.Message)" & vbCrLf
                strClassDALDeleteByValue2Use = strClassDALDeleteByValue2Use & "'End Try" & vbCrLf

                strClassDALSelect = strClassDALSelect & "Next" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageId = 1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Else" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageId = -1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageResultsetDS = ds" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Catch ex As Exception" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Packet.MessageId = -1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Try" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Return Packet" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Function" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details()" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    If tp.MessageId = 1 Then" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'        objLIB" & dbInfo(i).TABLENAME & "Listing = tp.MessageResultset" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'        ds = tp.MessageResultsetDS" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'        MyCLS.clsImaging.ByteArray2Image(,)" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'        MsgBox(objLIB" & dbInfo(i).TABLENAME & "Listing(0))" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    End If" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'Catch ex As Exception" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'    MsgBox(ex.Message)" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "'End Try" & vbCrLf

                If dbInfo(i).COLUMNDETAILS(0).COLSizeChar > 0 Then
                    strClassDALInsert = strClassDALInsert & "objParamListOut.Add(New SqlParameter(""@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & """, SqlDbType." & MyCLS.clsDBOperations.GetDBTypeValue4SqlDbTypes_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & "," & dbInfo(i).COLUMNDETAILS(0).COLSizeChar & "))" & vbCrLf
                Else
                    strClassDALInsert = strClassDALInsert & "objParamListOut.Add(New SqlParameter(""@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & """, SqlDbType." & MyCLS.clsDBOperations.GetDBTypeValue4SqlDbTypes_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & "))" & vbCrLf
                End If

                'strClassDALInsert = strClassDALInsert & "'MyCLS.clsCOMMON.ConOpen(true)" & vbCrLf
                strClassDALInsert = strClassDALInsert & "strOutParamValues = clsESPSql.ExecuteSPNonQueryOutPut(""SP_Insert" & dbInfo(i).TABLENAME & """, objParamList, objParamListOut, Packet.MessageId)" & vbCrLf
                'strClassDALInsert = strClassDALInsert & "Packet.MessageId = Result" & vbCrLf
                'strClassDALInsert = strClassDALInsert & "'MyCLS.clsCOMMON.ConClose()" & vbCrLf & vbCrLf
                strClassDALInsert = strClassDALInsert & "Packet.MessageResultset = strOutParamValues" & vbCrLf & vbCrLf
                strClassDALInsert = strClassDALInsert & "Catch ex As Exception" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Packet.MessageId = -1" & vbCrLf
                strClassDALInsert = strClassDALInsert & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Try" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Return Packet" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Function" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    tp.MessagePacket = objLIB" & dbInfo(i).TABLENAME & "" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    tp = objDAL" & dbInfo(i).TABLENAME & ".Insert" & dbInfo(i).TABLENAME & "(tp)" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    If tp.MessageId > -1 Then" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'        Dim strOutParamValues As String() = tp.MessageResultset" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'        MsgBox(strOutParamValues(0))" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    End If" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'Catch ex As Exception" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'    MsgBox(ex.Message)" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "'End Try" & vbCrLf


                strClassDALEnd = strClassDALEnd & "End Class" & vbCrLf
                strClassDALEnd = strClassDALEnd & "End Namespace"

                strClassDAL = strClassDALStart & strClassDALSelect2Use & vbCrLf & strClassDALSelect & vbCrLf & strClassDALSelectByValue2Use & vbCrLf & strClassDALSelectByValue & vbCrLf & strClassDALInsert2Use & vbCrLf & strClassDALInsert & vbCrLf & strClassDALDeleteByValue2Use & vbCrLf & strClassDALDeleteByValue & strClassDALEnd

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".vb")
                MyCLS.clsFileHandling.WriteFile(strClassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    ''' <summary>
    ''' CREATE DATA ACCESS LAYER IN VB
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks>CREATES SELECT, SELECT BY VALUE AND INSERT FUNCTIONs</remarks>
    Sub WriteDALInVB4Access(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassDAL As String = ""
            Dim strClassDALStart As String = ""
            Dim strClassDALEnd As String = ""
            Dim strClassDALSelect As String = ""
            Dim strClassDALSelectByValue As String = ""
            Dim strClassDALSelectByKey As String = ""
            Dim strClassDALInsert As String = ""
            Dim strClassDALInsertFields As String = ""
            Dim strClassDALInsertFieldsValues As String = ""
            Dim strClassDALUpdateFields As String = ""
            Dim strClassDALUpdateFieldsValues As String = ""
            '*
            Dim strClassDALInsert2Use As String = ""
            Dim strClassDALSelect2Use As String = ""
            Dim strClassDALSelectByValue2Use As String = ""
            '**
            For i As Integer = 0 To dbInfo.Count - 1
                strClassDAL = ""
                strClassDALStart = ""
                strClassDALEnd = ""
                strClassDALSelect = ""
                strClassDALSelectByValue = ""
                strClassDALSelectByKey = ""
                strClassDALInsert = ""
                strClassDALInsertFields = ""
                strClassDALInsertFieldsValues = ""
                strClassDALUpdateFields = ""
                strClassDALUpdateFieldsValues = ""

                strClassDALInsert2Use = ""
                strClassDALSelect2Use = ""
                strClassDALSelectByValue2Use = ""

                strClassDALStart = "Imports NDS.LIB" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data.OleDb" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Namespace NDS.DAL" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Public Class DAL" & dbInfo(i).TABLENAME & vbCrLf & vbCrLf



                strClassDALSelectByValue = strClassDALSelectByValue & "Public Function Get" & dbInfo(i).TABLENAME & "Details(ByVal Id As Int16) as LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim ds As New DataSet" & vbCrLf
                'strClassDALSelectByValue = strClassDALSelectByValue & "Dim objParamList As New List(Of OleDbParameter)()" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Try" & vbCrLf
                'strClassDALSelectByValue = strClassDALSelectByValue & "objParamList.Add(New OleDbParameter(""@Id"", Id))" & vbCrLf
                'strClassDALSelectByValue = strClassDALSelectByValue & "ds = MyCLS.clsExecuteStoredProc.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList)" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim strQ As String" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "strQ = ""Select * From " & dbInfo(i).TABLENAME & " Where " & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=" & """ & Id" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsCOMMON.ConOpen()" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsCOMMON.prcQuerySelectDS(ds, strQ,""" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsCOMMON.ConClose()" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf

                strClassDALSelectByKey = strClassDALSelectByKey & "Public Function Get" & dbInfo(i).TABLENAME & "DetailsByKeyword(ByVal Key As String) as LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Dim ds As New DataSet" & vbCrLf
                'strClassDALSelectByKey = strClassDALSelectByKey & "Dim objParamList As New List(Of OleDbParameter)()" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Try" & vbCrLf
                'strClassDALSelectByKey = strClassDALSelectByKey & "objParamList.Add(New OleDbParameter(""@Id"", Id))" & vbCrLf
                'strClassDALSelectByKey = strClassDALSelectByKey & "ds = MyCLS.clsExecuteStoredProc.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList)" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Dim strQ As String" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "strQ = ""Select * From " & dbInfo(i).TABLENAME & " Where "" & _" & vbCrLf & vbTab & vbTab & """"


                strClassDALSelect = strClassDALSelect & "Public Function Get" & dbInfo(i).TABLENAME & "Details() as LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim ds As New DataSet" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Try" & vbCrLf
                'strClassDALSelect = strClassDALSelect & "ds = MyCLS.clsExecuteStoredProc.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim strQ As String" & vbCrLf
                strClassDALSelect = strClassDALSelect & "strQ = ""Select * From " & dbInfo(i).TABLENAME & """" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsCOMMON.ConOpen()" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsCOMMON.prcQuerySelectDS(ds, strQ,""" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsCOMMON.ConClose()" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf

                strClassDALInsert = strClassDALInsert & "Public Function Insert" & dbInfo(i).TABLENAME & "(ByVal objLIB" & dbInfo(i).TABLENAME & " As LIB" & dbInfo(i).TABLENAME & ") As Int16" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim strQ As String" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Try" & vbCrLf
                strClassDALInsert = strClassDALInsert & "MyCLS.clsCOMMON.ConOpen()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "If objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "= -1 Then" & vbCrLf
                strClassDALInsert = strClassDALInsert & "strQ = ""INSERT INTO " & dbInfo(i).TABLENAME & """ & _" & vbCrLf
                strClassDALInsertFields = """ ("
                strClassDALInsertFieldsValues = """ Values("
                strClassDALUpdateFields = "strQ = ""UPDATE " & dbInfo(i).TABLENAME & """ & _" & vbCrLf & """ SET "

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    strClassDALSelectByKey = strClassDALSelectByKey & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " Like '%"" & Key & ""%' Or "

                    strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    'strClassDALInsert = strClassDALInsert & "objParamList.Add(New OleDbParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "))" & vbCrLf
                    strClassDALInsertFields = strClassDALInsertFields & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ","
                    strClassDALInsertFieldsValues = strClassDALInsertFieldsValues & "'"" & objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " & ""',"

                    strClassDALUpdateFieldsValues = strClassDALUpdateFieldsValues & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & "'"" & objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " & ""',"


                    '*
                    strClassDALInsert2Use = strClassDALInsert2Use & "'    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & IIf(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE = "byte()", "MyCLS.clsImaging.PictureBoxToByteArray()", """") & vbCrLf
                    strClassDALSelect2Use = strClassDALSelect2Use & "''    objLIB" & dbInfo(i).TABLENAME & "Listing(" & j & ")." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & vbCrLf
                    strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "''    objLIB" & dbInfo(i).TABLENAME & "Listing(" & j & ")." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & vbCrLf
                    '**
                    Application.DoEvents()
                Next
                strClassDALInsertFields = Mid(strClassDALInsertFields, 1, Len(strClassDALInsertFields) - 1) & ")"""
                strClassDALInsertFieldsValues = Mid(strClassDALInsertFieldsValues, 1, Len(strClassDALInsertFieldsValues) - 1) & ")"""
                strClassDALUpdateFields = strClassDALUpdateFields & Mid(strClassDALUpdateFieldsValues, 1, Len(strClassDALUpdateFieldsValues) - 1) & """ & _" & vbCrLf & """ Where " & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " = " & """ & objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " & """
                strClassDALInsert = strClassDALInsert & strClassDALInsertFields & " & _" & vbCrLf
                strClassDALInsert = strClassDALInsert & strClassDALInsertFieldsValues & vbCrLf
                strClassDALInsert = strClassDALInsert & "Return MyCLS.fnQueryInsert(strQ)" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Else" & vbCrLf
                strClassDALInsert = strClassDALInsert & strClassDALUpdateFields & vbCrLf
                strClassDALInsert = strClassDALInsert & "Return MyCLS.fnQueryUpdate(strQ)" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End If" & vbCrLf
                strClassDALInsert = strClassDALInsert & "MyCLS.clsCOMMON.ConClose()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Catch ex As Exception" & vbCrLf
                strClassDALInsert = strClassDALInsert & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Return -1" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Try" & vbCrLf
                'strClassDALInsert = strClassDALInsert & "Return strOutParamValues" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Function" & vbCrLf & vbCrLf


                strClassDALSelectByValue = strClassDALSelectByValue & "Next" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Catch ex As Exception" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Try" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Return objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Function" & vbCrLf & vbCrLf


                strClassDALSelectByKey = Mid(strClassDALSelectByKey, 1, Len(strClassDALSelectByKey) - 4) & """" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "MyCLS.clsCOMMON.ConOpen()" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "MyCLS.clsCOMMON.prcQuerySelectDS(ds, strQ,""" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "MyCLS.clsCOMMON.ConClose()" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Next" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End If" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End If" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End If" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End If" & vbCrLf & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Catch ex As Exception" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End Try" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Return objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End Function" & vbCrLf & vbCrLf


                strClassDALSelect = strClassDALSelect & "Next" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Catch ex As Exception" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Try" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Return objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Function" & vbCrLf & vbCrLf



                strClassDALEnd = strClassDALEnd & "End Class" & vbCrLf
                strClassDALEnd = strClassDALEnd & "End Namespace"

                '*
                strClassDAL = strClassDALStart & strClassDALSelect2Use & vbCrLf & strClassDALSelect & vbCrLf & strClassDALSelectByValue2Use & vbCrLf & strClassDALSelectByValue & vbCrLf & strClassDALInsert2Use & vbCrLf & strClassDALInsert & strClassDALEnd
                'strClassDAL = strClassDALStart & strClassDALSelect & vbCrLf & strClassDALSelectByValue & vbCrLf & strClassDALInsert & vbCrLf & strClassDALSelectByKey & strClassDALEnd
                '**

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".vb")
                MyCLS.clsFileHandling.WriteFile(strClassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    ''' <summary>
    ''' CREATE PROPERTIES LIBRARY IN CS
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks></remarks>
    Sub WritePropertyInCS(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassLIB As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassLIB = "using System;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Collections.Generic;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Linq;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Text;" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "namespace NDS.LIB" & vbCrLf & "{" & vbCrLf & "[Serializable]" & vbCrLf
                strClassLIB = strClassLIB & "public class LIB" & dbInfo(i).TABLENAME & vbCrLf & "{" & vbCrLf
                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    'MsgBox(dbInfo(i).TABLENAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE)
                    strClassLIB = strClassLIB & "private " & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE & " _" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ";" & vbCrLf
                    strClassLIB = strClassLIB & "public " & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE & " " & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf & "{" & vbCrLf
                    strClassLIB = strClassLIB & "get;" & vbCrLf
                    strClassLIB = strClassLIB & "set;" & vbCrLf
                    strClassLIB = strClassLIB & "}" & vbCrLf

                    Application.DoEvents()
                Next
                strClassLIB = strClassLIB & "}" & vbCrLf & vbCrLf & vbCrLf

                strClassLIB = strClassLIB & "[Serializable]" & vbCrLf
                strClassLIB = strClassLIB & "public class LIB" & dbInfo(i).TABLENAME & "Listing : List<LIB" & dbInfo(i).TABLENAME & ">" & vbCrLf
                strClassLIB = strClassLIB & "{" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "}" & vbCrLf

                strClassLIB = strClassLIB & "}"

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\LIB\LIB" & dbInfo(i).TABLENAME & ".cs")
                MyCLS.clsFileHandling.WriteFile(strClassLIB)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    Sub WritePropertyInCS_UPDATED(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassLIB As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassLIB = "using System;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Collections.Generic;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Linq;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Text;" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "namespace NDS.LIB" & vbCrLf & "{" & vbCrLf & "[Serializable]" & vbCrLf
                strClassLIB = strClassLIB & "public class LIB" & dbInfo(i).TABLENAME & vbCrLf & "{" & vbCrLf
                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    'MsgBox(dbInfo(i).TABLENAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLDataType)
                    strClassLIB = strClassLIB & "private " & dbInfo(i).COLUMNDETAILS(j).COLDataType & " _" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ";" & vbCrLf
                    strClassLIB = strClassLIB & "public " & dbInfo(i).COLUMNDETAILS(j).COLDataType & " " & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf & "{" & vbCrLf
                    strClassLIB = strClassLIB & "get;" & vbCrLf
                    strClassLIB = strClassLIB & "set;" & vbCrLf
                    strClassLIB = strClassLIB & "}" & vbCrLf

                    Application.DoEvents()
                Next
                strClassLIB = strClassLIB & "}" & vbCrLf & vbCrLf & vbCrLf

                strClassLIB = strClassLIB & "[Serializable]" & vbCrLf
                strClassLIB = strClassLIB & "public class LIB" & dbInfo(i).TABLENAME & "Listing : List<LIB" & dbInfo(i).TABLENAME & ">" & vbCrLf
                strClassLIB = strClassLIB & "{" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "}" & vbCrLf

                strClassLIB = strClassLIB & "}"

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\LIB\LIB" & dbInfo(i).TABLENAME & ".cs")
                MyCLS.clsFileHandling.WriteFile(strClassLIB)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    Sub WritePropertyInCS_MVC(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassLIB As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassLIB = "using System;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Collections.Generic;" & vbCrLf
                strClassLIB = strClassLIB & "using System.ComponentModel.DataAnnotations;" & vbCrLf
                strClassLIB = strClassLIB & "using System.ComponentModel.DataAnnotations.Schema;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Linq;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Web;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Web.Mvc;" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "namespace NDS.Models" & vbCrLf & "{" & vbCrLf
                strClassLIB = strClassLIB & "[Serializable()]" & vbCrLf
                strClassLIB = strClassLIB & "public class LIB" & dbInfo(i).TABLENAME & vbCrLf & "{" & vbCrLf
                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    'MsgBox(dbInfo(i).TABLENAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLDataType)
                    'strClassLIB = strClassLIB & "private " & dbInfo(i).COLUMNDETAILS(j).COLDataType & " _" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ";" & vbCrLf
                    strClassLIB = strClassLIB & "[Display(Name = """ & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """)]" & vbCrLf
                    strClassLIB = strClassLIB & "public " & dbInfo(i).COLUMNDETAILS(j).COLDataType & " " & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "{" '& vbCrLf
                    strClassLIB = strClassLIB & "get;" '& vbCrLf
                    strClassLIB = strClassLIB & "set;" '& vbCrLf
                    strClassLIB = strClassLIB & "}" & vbCrLf

                    Application.DoEvents()
                Next
                strClassLIB = strClassLIB & "}" & vbCrLf & vbCrLf & vbCrLf

                strClassLIB = strClassLIB & "[Serializable]" & vbCrLf
                strClassLIB = strClassLIB & "public class LIB" & dbInfo(i).TABLENAME & "Listing : List<LIB" & dbInfo(i).TABLENAME & ">" & vbCrLf
                strClassLIB = strClassLIB & "{" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "}" & vbCrLf

                strClassLIB = strClassLIB & "}"

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\LIB\LIB" & dbInfo(i).TABLENAME & ".cs")
                MyCLS.clsFileHandling.WriteFile(strClassLIB)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    Sub WritePropertyInCS_MVCAPI(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassLIB As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassLIB = "using System;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Collections.Generic;" & vbCrLf
                strClassLIB = strClassLIB & "using System.ComponentModel.DataAnnotations;" & vbCrLf
                strClassLIB = strClassLIB & "using System.ComponentModel.DataAnnotations.Schema;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Linq;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Web;" & vbCrLf
                strClassLIB = strClassLIB & "using System.Web.Mvc;" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "namespace NDS.Models" & vbCrLf & "{" & vbCrLf
                strClassLIB = strClassLIB & "//[Serializable()]" & vbCrLf
                strClassLIB = strClassLIB & "public class " & dbInfo(i).TABLENAME & vbCrLf & "{" & vbCrLf
                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    'MsgBox(dbInfo(i).TABLENAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLDataType)
                    'strClassLIB = strClassLIB & "private " & dbInfo(i).COLUMNDETAILS(j).COLDataType & " _" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ";" & vbCrLf
                    'strClassLIB = strClassLIB & "[Display(Name = """ & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """)]" & vbCrLf
                    strClassLIB = strClassLIB & "public " & dbInfo(i).COLUMNDETAILS(j).COLDataType & " " & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "{" '& vbCrLf
                    strClassLIB = strClassLIB & "get;" '& vbCrLf
                    strClassLIB = strClassLIB & "set;" '& vbCrLf
                    strClassLIB = strClassLIB & "}" & vbCrLf

                    Application.DoEvents()
                Next
                strClassLIB = strClassLIB & "}" & vbCrLf & vbCrLf & vbCrLf

                strClassLIB = strClassLIB & "//[Serializable]" & vbCrLf
                strClassLIB = strClassLIB & "//public class LIB" & dbInfo(i).TABLENAME & "Listing : List<LIB" & dbInfo(i).TABLENAME & ">" & vbCrLf
                strClassLIB = strClassLIB & "//{" & vbCrLf & vbCrLf
                strClassLIB = strClassLIB & "//}" & vbCrLf

                strClassLIB = strClassLIB & "}"

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\LIB\" & dbInfo(i).TABLENAME & ".cs")
                MyCLS.clsFileHandling.WriteFile(strClassLIB)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    ''' <summary>
    ''' CREATE DATA ACCESS LAYER IN CS
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks>CREATES SELECT, SELECT BY VALUE AND INSERT FUNCTIONs</remarks>
    Sub WriteDALInCS(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strclassDAL As String = ""
            Dim strclassDALStart As String = ""
            Dim strclassDALEnd As String = ""
            Dim strclassDALSelect As String = ""
            Dim strclassDALSelectByValue As String = ""
            Dim strclassDALInsert As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strclassDAL = ""
                strclassDALStart = ""
                strclassDALEnd = ""
                strclassDALSelect = ""
                strclassDALSelectByValue = ""
                strclassDALInsert = ""

                strclassDALStart = "using NDS.LIB;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data.OleDb;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Collections.Generic;" & vbCrLf
                strclassDALStart = strclassDALStart & "using MyCLSWin;" & vbCrLf & vbCrLf
                strclassDALStart = strclassDALStart & "namespace NDS.DAL" & vbCrLf & "{" & vbCrLf
                strclassDALStart = strclassDALStart & "public class DAL" & dbInfo(i).TABLENAME & vbCrLf & "{" & vbCrLf



                strclassDALSelectByValue = strclassDALSelectByValue & "public LIB" & dbInfo(i).TABLENAME & "Listing Get" & dbInfo(i).TABLENAME & "Details(int Id)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "DataSet ds = new DataSet();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "List<OleDbParameter> objParamList = new List<OleDbParameter>();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "MyCLS.clsExecuteStoredProcSql clsESPSql = New MyCLS.clsExecuteStoredProcSql();" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "try" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "objParamList.Add(new OleDbParameter(""@Id"", Id));" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList);" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables[0].Rows != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables[0].Rows.Count > 0)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "for (int i = 0; i < ds.Tables[0].Rows.Count; i++)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "LIB" & dbInfo(i).TABLENAME & " oLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ");" & vbCrLf

                strclassDALSelect = strclassDALSelect & "public LIB" & dbInfo(i).TABLENAME & "Listing" & " Get" & dbInfo(i).TABLENAME & "Details()" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "DataSet ds = new DataSet();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "MyCLS.clsExecuteStoredProcSql clsESPSql = New MyCLS.clsExecuteStoredProcSql();" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "try" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """);" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables[0].Rows != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables[0].Rows.Count > 0)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "for (int i = 0; i < ds.Tables[0].Rows.Count; i++)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "LIB" & dbInfo(i).TABLENAME & " oLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ");" & vbCrLf

                strclassDALInsert = strclassDALInsert & "public string[] Insert" & dbInfo(i).TABLENAME & "(LIB" & dbInfo(i).TABLENAME & " objLIB" & dbInfo(i).TABLENAME & ", short Result)" & vbCrLf & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "String[] strOutParamValues = new String[10];" & vbCrLf
                strclassDALInsert = strclassDALInsert & "List<OleDbParameter> objParamList = new List<OleDbParameter>();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "List<OleDbParameter> objParamListOut = new List<OleDbParameter>();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "try" & vbCrLf & "{" & vbCrLf

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """].ToString();" & vbCrLf

                    strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """].ToString();" & vbCrLf

                    strclassDALInsert = strclassDALInsert & "objParamList.Add(new OleDbParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "));" & vbCrLf


                    Application.DoEvents()
                Next

                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "catch (Exception ex)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString());" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "return objLIB" & dbInfo(i).TABLENAME & "Listing;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf & vbCrLf

                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "catch (Exception ex)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString());" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "return objLIB" & dbInfo(i).TABLENAME & "Listing;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf & vbCrLf

                strclassDALInsert = strclassDALInsert & "objParamListOut.Add(new OleDbParameter(""@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & """, OleDbType." & dbInfo(i).COLUMNDETAILS(0).COLUMNTYPE & "));" & vbCrLf
                strclassDALInsert = strclassDALInsert & "//MyCLS.clsCOMMON.ConOpen(true);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "strOutParamValues = MyCLS.clsExecuteStoredProc.ExecuteSPNonQueryOutPut(""SP_Insert" & dbInfo(i).TABLENAME & """, objParamList, objParamListOut, ref Result);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "Result = 1;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "//MyCLS.clsCOMMON.ConClose();" & vbCrLf & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf
                strclassDALInsert = strclassDALInsert & "catch (Exception ex)" & vbCrLf
                strclassDALInsert = strclassDALInsert & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "Result = -1;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString());" & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf
                strclassDALInsert = strclassDALInsert & "return strOutParamValues;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf & vbCrLf


                strclassDALEnd = strclassDALEnd & "}" & vbCrLf
                strclassDALEnd = strclassDALEnd & "}"

                strclassDAL = strclassDALStart & strclassDALSelect & vbCrLf & strclassDALSelectByValue & vbCrLf & strclassDALInsert & strclassDALEnd

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".cs")
                MyCLS.clsFileHandling.WriteFile(strclassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    ''' <summary>
    ''' CREATE DATA ACCESS LAYER IN VB WITH TP AND SQL CLIENT
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks>CREATES SELECT, SELECT BY VALUE AND INSERT FUNCTIONs</remarks>
    Sub WriteDALInCSSqlClient(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strclassDAL As String = ""
            Dim strclassDALStart As String = ""
            Dim strclassDALEnd As String = ""
            Dim strclassDALSelect As String = ""
            Dim strclassDALSelectByValue As String = ""
            Dim strclassDALInsert As String = ""
            Dim strClassDALInsert2Use As String = ""
            Dim strClassDALSelect2Use As String = ""
            Dim strClassDALSelectByValue2Use As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strclassDAL = ""
                strclassDALStart = ""
                strclassDALEnd = ""
                strclassDALSelect = ""
                strclassDALSelectByValue = ""
                strclassDALInsert = ""
                strClassDALInsert2Use = ""
                strClassDALSelect2Use = ""
                strClassDALSelectByValue2Use = ""

                strclassDALStart = "using NDS.LIB;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data.OleDb;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data.SqlClient;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Collections.Generic;" & vbCrLf
                'strclassDALStart = strclassDALStart & "using MyCLSWin;" & vbCrLf & vbCrLf
                strclassDALStart = strclassDALStart & "namespace NDS.DAL" & vbCrLf & "{" & vbCrLf
                strclassDALStart = strclassDALStart & "public class DAL" & dbInfo(i).TABLENAME & vbCrLf & "{" & vbCrLf


                strclassDALSelectByValue = strclassDALSelectByValue & "/// <summary>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// Accepts=TransportationPacket, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// </summary>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// <returns></returns>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// <remarks></remarks>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "public MyCLS.TransportationPacket Get" & dbInfo(i).TABLENAME & "Details(MyCLS.TransportationPacket Packet)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "DataSet ds = new DataSet();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "List<SqlParameter> objParamList = new List<SqlParameter>();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "MyCLS.clsExecuteStoredProcSql clsESPSql = New MyCLS.clsExecuteStoredProcSql();" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "try" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "objParamList.Add(new SqlParameter(""@Id"", Packet.MessagePacket));" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList);" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables[0].Rows != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables[0].Rows.Count > 0)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "for (int i = 0; i < ds.Tables[0].Rows.Count; i++)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "LIB" & dbInfo(i).TABLENAME & " oLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ");" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//*******COPY IT TO USE BELOW FUNCTION - SELECT BY ID************" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//try" & vbCrLf & "//{" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    DAL" & dbInfo(i).TABLENAME & " objDAL" & dbInfo(i).TABLENAME & " = new DAL" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    MyCLS.TransportationPacket tp = new MyCLS.TransportationPacket();" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    DataSet ds = new DataSet();" & vbCrLf & vbCrLf
                'strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    tp.MessagePacket = 1;    //ID to be Passed" & vbCrLf & vbCrLf

                strclassDALSelect = strclassDALSelect & "/// <summary>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// </summary>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// <returns></returns>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// <remarks></remarks>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "public MyCLS.TransportationPacket Get" & dbInfo(i).TABLENAME & "Details()" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "DataSet ds = new DataSet();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "MyCLS.clsExecuteStoredProcSql clsESPSql = New MyCLS.clsExecuteStoredProcSql();" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "MyCLS.TransportationPacket Packet = new MyCLS.TransportationPacket();" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "try" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """);" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables[0].Rows != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables[0].Rows.Count > 0)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "for (int i = 0; i < ds.Tables[0].Rows.Count; i++)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "LIB" & dbInfo(i).TABLENAME & " oLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ");" & vbCrLf
                '***Connection String***
                strClassDALSelect2Use = strClassDALSelect2Use & "//PUT IT IN LOAD EVENTS" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//MyCLS.strConnStringOLEDB = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"";" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//MyCLS.strConnStringSQLCLIENT = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;"";" & vbCrLf & vbCrLf
                '***Connection String***
                strClassDALSelect2Use = strClassDALSelect2Use & "//*******COPY IT TO USE BELOW FUNCTION - SELECT ALL************" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//try" & vbCrLf & "//{" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    DAL" & dbInfo(i).TABLENAME & " objDAL" & dbInfo(i).TABLENAME & " = new DAL" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    MyCLS.TransportationPacket tp = new MyCLS.TransportationPacket();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    DataSet ds = new DataSet();" & vbCrLf & vbCrLf

                strclassDALInsert = strclassDALInsert & "/// <summary>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// Accepts=Packet, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// </summary>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// <param name=""Packet""></param>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// <returns></returns>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// <remarks></remarks>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "public MyCLS.TransportationPacket Insert" & dbInfo(i).TABLENAME & "(MyCLS.TransportationPacket Packet)" & vbCrLf & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "String[] strOutParamValues = new String[10];" & vbCrLf
                strclassDALInsert = strclassDALInsert & "List<SqlParameter> objParamList = new List<SqlParameter>();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "List<SqlParameter> objParamListOut = new List<SqlParameter>();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "int Result=0;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "try" & vbCrLf & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "LIB" & dbInfo(i).TABLENAME & " objLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "objLIB" & dbInfo(i).TABLENAME & " = (LIB" & dbInfo(i).TABLENAME & ")Packet.MessagePacket;" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//*******COPY IT TO USE BELOW FUNCTION - INSERT************" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//try" & vbCrLf & "//{" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    LIB" & dbInfo(i).TABLENAME & " objLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    DAL" & dbInfo(i).TABLENAME & " objDAL" & dbInfo(i).TABLENAME & " = new DAL" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    MyCLS.TransportationPacket tp = new MyCLS.TransportationPacket();" & vbCrLf & vbCrLf

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """].ToString();" & vbCrLf

                    strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """].ToString();" & vbCrLf

                    strclassDALInsert = strclassDALInsert & "objParamList.Add(new SqlParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "));" & vbCrLf

                    'strClassDALInsert2Use = strClassDALInsert2Use & "//    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf
                    'strClassDALSelect2Use = strClassDALSelect2Use & "////    objLIB" & dbInfo(i).TABLENAME & "Listing[" & j & "]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf
                    'strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "////    objLIB" & dbInfo(i).TABLENAME & "Listing[" & j & "]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf
                    strClassDALInsert2Use = strClassDALInsert2Use & "//    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & IIf(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE = "byte()", "MyCLS.clsImaging.PictureBoxToByteArray()", "txt.Text") & ";" & vbCrLf
                    strClassDALSelect2Use = strClassDALSelect2Use & "////  txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing[0]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ";" & vbCrLf
                    strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "////  txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing[0]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf

                    Application.DoEvents()
                Next

                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageId = 1;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "else" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageId = -1;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageResultsetDS = ds;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing;" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "catch (Exception ex)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageId = -1;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString());" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "return Packet;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    tp.MessagePacket = 1;    //ID to be Passed" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details(tp);" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    if(tp.MessageId == 1)" & vbCrLf & "//{" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//        objLIB" & dbInfo(i).TABLENAME & "Listing = (LIB" & dbInfo(i).TABLENAME & "Listing)tp.MessageResultset;" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//        ds = (DataSet)tp.MessageResultsetDS;" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//        MessageBox.Show(objLIB" & dbInfo(i).TABLENAME & "Listing[0].ToString());" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    }" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    }" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//catch(Exception ex)" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    {" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    }" & vbCrLf

                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageId = 1;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "else" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageId = -1;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageResultsetDS = ds;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing;" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "catch (Exception ex)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageId = -1;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "return Packet;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    if(tp.MessageId == 1)" & vbCrLf & "//{" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//        objLIB" & dbInfo(i).TABLENAME & "Listing = (LIB" & dbInfo(i).TABLENAME & "Listing)tp.MessageResultset;" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//        ds = (DataSet)tp.MessageResultsetDS;" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//        MessageBox.Show(objLIB" & dbInfo(i).TABLENAME & "Listing[0].ToString());" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    }" & vbCrLf & "//}" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//catch(Exception ex)" & vbCrLf & "//{" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//}" & vbCrLf

                strclassDALInsert = strclassDALInsert & "objParamListOut.Add(new SqlParameter(""@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & """, SqlDbType." & MyCLS.clsDBOperations.GetDBTypeValue4SqlDbTypes(dbInfo(i).COLUMNDETAILS(0).COLUMNTYPE) & "));" & vbCrLf
                'strclassDALInsert = strclassDALInsert & "//MyCLS.clsCOMMON.ConOpen(true);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "strOutParamValues = clsESPSql.ExecuteSPNonQueryOutPut(""SP_Insert" & dbInfo(i).TABLENAME & """, objParamList, objParamListOut, ref Result);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "Packet.MessageId = Result;" & vbCrLf
                'strclassDALInsert = strclassDALInsert & "//MyCLS.clsCOMMON.ConClose();" & vbCrLf & vbCrLf
                strclassDALInsert = strclassDALInsert & "Packet.MessageResultset = strOutParamValues;" & vbCrLf & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf
                strclassDALInsert = strclassDALInsert & "catch (Exception ex)" & vbCrLf
                strclassDALInsert = strclassDALInsert & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "Packet.MessageId = -1;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString());" & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf
                strclassDALInsert = strclassDALInsert & "return Packet;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    tp.MessagePacket = objLIB" & dbInfo(i).TABLENAME & ";" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    tp = objDAL" & dbInfo(i).TABLENAME & ".Insert" & dbInfo(i).TABLENAME & "(tp);" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    if(tp.MessageId > -1)" & vbCrLf & "//{" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//        string[] strOutParamValues = (string[])tp.MessageResultset;" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//        MessageBox.Show(strOutParamValues[0].ToString());" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    }" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    }" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//catch(Exception ex)" & vbCrLf & "//{" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//}" & vbCrLf

                strclassDALEnd = strclassDALEnd & "}" & vbCrLf
                strclassDALEnd = strclassDALEnd & "}"

                strclassDAL = strclassDALStart & strClassDALSelect2Use & vbCrLf & strclassDALSelect & vbCrLf & strClassDALSelectByValue2Use & vbCrLf & strclassDALSelectByValue & vbCrLf & strClassDALInsert2Use & vbCrLf & strclassDALInsert & strclassDALEnd

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".cs")
                MyCLS.clsFileHandling.WriteFile(strclassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    Sub WriteDALInCSSqlClient_UPDATED(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strclassDAL As String = ""
            Dim strclassDALStart As String = ""
            Dim strclassDALEnd As String = ""
            Dim strclassDALSelect As String = ""
            Dim strclassDALSelectByValue As String = ""
            Dim strclassDALInsert As String = ""
            Dim strClassDALInsert2Use As String = ""
            Dim strClassDALSelect2Use As String = ""
            Dim strClassDALSelectByValue2Use As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strclassDAL = ""
                strclassDALStart = ""
                strclassDALEnd = ""
                strclassDALSelect = ""
                strclassDALSelectByValue = ""
                strclassDALInsert = ""
                strClassDALInsert2Use = ""
                strClassDALSelect2Use = ""
                strClassDALSelectByValue2Use = ""

                strclassDALStart = "using NDS.LIB;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data.OleDb;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data.SqlClient;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Collections.Generic;" & vbCrLf
                strclassDALStart = strclassDALStart & "using MyCLS;" & vbCrLf & vbCrLf
                strclassDALStart = strclassDALStart & "namespace NDS.DAL" & vbCrLf & "{" & vbCrLf
                strclassDALStart = strclassDALStart & "public class DAL" & dbInfo(i).TABLENAME & vbCrLf & "{" & vbCrLf


                strclassDALSelectByValue = strclassDALSelectByValue & "/// <summary>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// Accepts=TransportationPacket, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// </summary>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// <returns></returns>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// <remarks></remarks>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "public MyCLS.TransportationPacket Get" & dbInfo(i).TABLENAME & "Details(MyCLS.TransportationPacket Packet)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "DataSet ds = new DataSet();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "List<SqlParameter> objParamList = new List<SqlParameter>();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "MyCLS.clsExecuteStoredProcSql clsESPSql = new MyCLS.clsExecuteStoredProcSql();" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "try" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "objParamList.Add(new SqlParameter(""@Id"", Packet.MessagePacket));" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList);" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables[0].Rows != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables[0].Rows.Count > 0)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "for (int i = 0; i < ds.Tables[0].Rows.Count; i++)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "LIB" & dbInfo(i).TABLENAME & " oLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ");" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//*******COPY IT TO USE BELOW FUNCTION - SELECT BY ID************" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//try" & vbCrLf & "//{" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    DAL" & dbInfo(i).TABLENAME & " objDAL" & dbInfo(i).TABLENAME & " = new DAL" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    MyCLS.TransportationPacket tp = new MyCLS.TransportationPacket();" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    DataSet ds = new DataSet();" & vbCrLf & vbCrLf
                'strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    tp.MessagePacket = 1;    //ID to be Passed" & vbCrLf & vbCrLf

                strclassDALSelect = strclassDALSelect & "/// <summary>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// </summary>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// <returns></returns>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// <remarks></remarks>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "public MyCLS.TransportationPacket Get" & dbInfo(i).TABLENAME & "Details()" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "DataSet ds = new DataSet();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "MyCLS.TransportationPacket Packet = new MyCLS.TransportationPacket();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "MyCLS.clsExecuteStoredProcSql clsESPSql = new MyCLS.clsExecuteStoredProcSql();" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "try" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """);" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables[0].Rows != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables[0].Rows.Count > 0)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "for (int i = 0; i < ds.Tables[0].Rows.Count; i++)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "LIB" & dbInfo(i).TABLENAME & " oLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ");" & vbCrLf
                '***Connection String***
                strClassDALSelect2Use = strClassDALSelect2Use & "//PUT IT IN LOAD EVENTS" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//MyCLS.strConnStringOLEDB = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"";" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//MyCLS.strConnStringSQLCLIENT = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;"";" & vbCrLf & vbCrLf
                '***Connection String***
                strClassDALSelect2Use = strClassDALSelect2Use & "//*******COPY IT TO USE BELOW FUNCTION - SELECT ALL************" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//try" & vbCrLf & "//{" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    DAL" & dbInfo(i).TABLENAME & " objDAL" & dbInfo(i).TABLENAME & " = new DAL" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    MyCLS.TransportationPacket tp = new MyCLS.TransportationPacket();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    DataSet ds = new DataSet();" & vbCrLf & vbCrLf

                strclassDALInsert = strclassDALInsert & "/// <summary>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// Accepts=Packet, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// </summary>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// <param name=""Packet""></param>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// <returns></returns>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// <remarks></remarks>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "public MyCLS.TransportationPacket Insert" & dbInfo(i).TABLENAME & "(MyCLS.TransportationPacket Packet)" & vbCrLf & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "String[] strOutParamValues = new String[10];" & vbCrLf
                strclassDALInsert = strclassDALInsert & "List<SqlParameter> objParamList = new List<SqlParameter>();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "List<SqlParameter> objParamListOut = new List<SqlParameter>();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "MyCLS.clsExecuteStoredProcSql clsESPSql = new MyCLS.clsExecuteStoredProcSql();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "int Result=0;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "try" & vbCrLf & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "LIB" & dbInfo(i).TABLENAME & " objLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "objLIB" & dbInfo(i).TABLENAME & " = (LIB" & dbInfo(i).TABLENAME & ")Packet.MessagePacket;" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//*******COPY IT TO USE BELOW FUNCTION - INSERT************" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//try" & vbCrLf & "//{" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    LIB" & dbInfo(i).TABLENAME & " objLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    DAL" & dbInfo(i).TABLENAME & " objDAL" & dbInfo(i).TABLENAME & " = new DAL" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    MyCLS.TransportationPacket tp = new MyCLS.TransportationPacket();" & vbCrLf & vbCrLf

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    'MsgBox(dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " : " & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE & " : " & dbInfo(i).COLUMNDETAILS(j).COLDataTypeSQL)
                    If dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE = "string" Then
                        strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """].ToString();" & vbCrLf
                        strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """].ToString();" & vbCrLf
                        strclassDALInsert = strclassDALInsert & "objParamList.Add(new SqlParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "));" & vbCrLf
                    Else
                        strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " =(" & fnTypeCasting(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE) & ")ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """];" & vbCrLf
                        strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " =(" & fnTypeCasting(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE) & ")ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """];" & vbCrLf
                        strclassDALInsert = strclassDALInsert & "objParamList.Add(new SqlParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "));" & vbCrLf
                    End If
                    'strClassDALInsert2Use = strClassDALInsert2Use & "//    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf
                    'strClassDALSelect2Use = strClassDALSelect2Use & "////    objLIB" & dbInfo(i).TABLENAME & "Listing[" & j & "]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf
                    'strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "////    objLIB" & dbInfo(i).TABLENAME & "Listing[" & j & "]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf
                    strClassDALInsert2Use = strClassDALInsert2Use & "//    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & IIf(dbInfo(i).COLUMNDETAILS(j).COLDataType = "byte()", "MyCLS.clsImaging.PictureBoxToByteArray()", "txt.Text") & ";" & vbCrLf
                    strClassDALSelect2Use = strClassDALSelect2Use & "////  txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing[0]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ";" & vbCrLf
                    strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "////  txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing[0]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf

                    Application.DoEvents()
                Next

                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageId = 1;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "else" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageId = -1;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageResultsetDS = ds;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing;" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "catch (Exception ex)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageId = -1;" & vbCrLf

                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.ex = ex;" & vbCrLf

                strclassDALSelectByValue = strclassDALSelectByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "return Packet;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    tp.MessagePacket = 1;    //ID to be Passed" & vbCrLf & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details(tp);" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    if(tp.MessageId == 1)" & vbCrLf & "//{" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//        objLIB" & dbInfo(i).TABLENAME & "Listing = (LIB" & dbInfo(i).TABLENAME & "Listing)tp.MessageResultset;" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//        ds = (DataSet)tp.MessageResultsetDS;" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//        MessageBox.Show(objLIB" & dbInfo(i).TABLENAME & "Listing[0].ToString());" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    }" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    }" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//catch(Exception ex)" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    {" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    }" & vbCrLf

                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageId = 1;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "else" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageId = -1;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageResultsetDS = ds;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing;" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "catch (Exception ex)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageId = -1;" & vbCrLf

                strclassDALSelect = strclassDALSelect & "Packet.ex = ex;" & vbCrLf

                strclassDALSelect = strclassDALSelect & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "return Packet;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    if(tp.MessageId == 1)" & vbCrLf & "//{" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//        objLIB" & dbInfo(i).TABLENAME & "Listing = (LIB" & dbInfo(i).TABLENAME & "Listing)tp.MessageResultset;" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//        ds = (DataSet)tp.MessageResultsetDS;" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//        MessageBox.Show(objLIB" & dbInfo(i).TABLENAME & "Listing[0].ToString());" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    }" & vbCrLf & "//}" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//catch(Exception ex)" & vbCrLf & "//{" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//}" & vbCrLf

                strclassDALInsert = strclassDALInsert & "objParamListOut.Add(new SqlParameter(""@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & """, SqlDbType." & MyCLS.clsDBOperations.GetDBTypeValue4SqlDbTypes_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & "));" & vbCrLf
                'strclassDALInsert = strclassDALInsert & "//MyCLS.clsCOMMON.ConOpen(true);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "strOutParamValues = clsESPSql.ExecuteSPNonQueryOutPut(""SP_Insert" & dbInfo(i).TABLENAME & """, objParamList, objParamListOut, ref Result);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "Packet.MessageId = Result;" & vbCrLf
                'strclassDALInsert = strclassDALInsert & "//MyCLS.clsCOMMON.ConClose();" & vbCrLf & vbCrLf
                strclassDALInsert = strclassDALInsert & "Packet.MessageResultset = strOutParamValues;" & vbCrLf & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf
                strclassDALInsert = strclassDALInsert & "catch (Exception ex)" & vbCrLf
                strclassDALInsert = strclassDALInsert & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "Packet.MessageId = -1;" & vbCrLf

                strclassDALInsert = strclassDALInsert & "Packet.ex = ex;" & vbCrLf

                strclassDALInsert = strclassDALInsert & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf
                strclassDALInsert = strclassDALInsert & "return Packet;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    tp.MessagePacket = objLIB" & dbInfo(i).TABLENAME & ";" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    tp = objDAL" & dbInfo(i).TABLENAME & ".Insert" & dbInfo(i).TABLENAME & "(tp);" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    if(tp.MessageId > -1)" & vbCrLf & "//{" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//        string[] strOutParamValues = (string[])tp.MessageResultset;" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//        MessageBox.Show(strOutParamValues[0].ToString());" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    }" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    }" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//catch(Exception ex)" & vbCrLf & "//{" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//}" & vbCrLf

                strclassDALEnd = strclassDALEnd & "}" & vbCrLf
                strclassDALEnd = strclassDALEnd & "}"

                strclassDAL = strclassDALStart & strClassDALSelect2Use & vbCrLf & strclassDALSelect & vbCrLf & strClassDALSelectByValue2Use & vbCrLf & strclassDALSelectByValue & vbCrLf & strClassDALInsert2Use & vbCrLf & strclassDALInsert & strclassDALEnd

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".cs")
                MyCLS.clsFileHandling.WriteFile(strclassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    Sub WriteDALInCSSqlClient_UPDATED_MVC(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strclassDAL As String = ""
            Dim strclassDALStart As String = ""
            Dim strclassDALEnd As String = ""
            Dim strclassDALSelect As String = ""
            Dim strclassDALSelectByValue As String = ""
            Dim strclassDALInsert As String = ""
            Dim strClassDALInsert2Use As String = ""
            Dim strClassDALSelect2Use As String = ""
            Dim strClassDALSelectByValue2Use As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strclassDAL = ""
                strclassDALStart = ""
                strclassDALEnd = ""
                strclassDALSelect = ""
                strclassDALSelectByValue = ""
                strclassDALInsert = ""
                strClassDALInsert2Use = ""
                strClassDALSelect2Use = ""
                strClassDALSelectByValue2Use = ""

                strclassDALStart = "using NDS.Models;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data.OleDb;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Data.SqlClient;" & vbCrLf
                strclassDALStart = strclassDALStart & "using System.Collections.Generic;" & vbCrLf
                strclassDALStart = strclassDALStart & "using MyCLS;" & vbCrLf & vbCrLf
                strclassDALStart = strclassDALStart & "namespace NDS.DAL" & vbCrLf & "{" & vbCrLf
                strclassDALStart = strclassDALStart & "public class DAL" & dbInfo(i).TABLENAME & vbCrLf & "{" & vbCrLf


                strclassDALSelectByValue = strclassDALSelectByValue & "/// <summary>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// Accepts=TransportationPacket, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// </summary>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// <returns></returns>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "/// <remarks></remarks>" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "public MyCLS.TransportationPacket Get" & dbInfo(i).TABLENAME & "Details(MyCLS.TransportationPacket Packet)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "DataSet ds = new DataSet();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "List<SqlParameter> objParamList = new List<SqlParameter>();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "MyCLS.clsExecuteStoredProcSql clsESPSql = new MyCLS.clsExecuteStoredProcSql();" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "try" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "objParamList.Add(new SqlParameter(""@Id"", Packet.MessagePacket));" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList);" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables[0].Rows != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "if (ds.Tables[0].Rows.Count > 0)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "for (int i = 0; i < ds.Tables[0].Rows.Count; i++)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "LIB" & dbInfo(i).TABLENAME & " oLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ");" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//*******COPY IT TO USE BELOW FUNCTION - SELECT BY ID************" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//try" & vbCrLf & "//{" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    DAL" & dbInfo(i).TABLENAME & " objDAL" & dbInfo(i).TABLENAME & " = new DAL" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    MyCLS.TransportationPacket tp = new MyCLS.TransportationPacket();" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    DataSet ds = new DataSet();" & vbCrLf & vbCrLf

                strclassDALSelect = strclassDALSelect & "/// <summary>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// </summary>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// <returns></returns>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "/// <remarks></remarks>" & vbCrLf
                strclassDALSelect = strclassDALSelect & "public MyCLS.TransportationPacket Get" & dbInfo(i).TABLENAME & "Details()" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "DataSet ds = new DataSet();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "MyCLS.TransportationPacket Packet = new MyCLS.TransportationPacket();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "MyCLS.clsExecuteStoredProcSql clsESPSql = new MyCLS.clsExecuteStoredProcSql();" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "try" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "ds = clsESPSql.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """);" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables[0].Rows != null)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "if (ds.Tables[0].Rows.Count > 0)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "for (int i = 0; i < ds.Tables[0].Rows.Count; i++)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "LIB" & dbInfo(i).TABLENAME & " oLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ");" & vbCrLf
                '***Connection String***
                strClassDALSelect2Use = strClassDALSelect2Use & "//PUT IT IN LOAD EVENTS" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//MyCLS.strConnStringOLEDB = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"";" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//MyCLS.strConnStringSQLCLIENT = ""Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;"";" & vbCrLf & vbCrLf
                '***Connection String***
                strClassDALSelect2Use = strClassDALSelect2Use & "//*******COPY IT TO USE BELOW FUNCTION - SELECT ALL************" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//try" & vbCrLf & "//{" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    LIB" & dbInfo(i).TABLENAME & "Listing objLIB" & dbInfo(i).TABLENAME & "Listing = new LIB" & dbInfo(i).TABLENAME & "Listing();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    DAL" & dbInfo(i).TABLENAME & " objDAL" & dbInfo(i).TABLENAME & " = new DAL" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    MyCLS.TransportationPacket tp = new MyCLS.TransportationPacket();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    DataSet ds = new DataSet();" & vbCrLf & vbCrLf

                strclassDALInsert = strclassDALInsert & "/// <summary>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// Accepts=Packet, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// </summary>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// <param name=""Packet""></param>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// <returns></returns>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "/// <remarks></remarks>" & vbCrLf
                strclassDALInsert = strclassDALInsert & "public MyCLS.TransportationPacket Insert" & dbInfo(i).TABLENAME & "(MyCLS.TransportationPacket Packet)" & vbCrLf & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "String[] strOutParamValues = new String[10];" & vbCrLf
                strclassDALInsert = strclassDALInsert & "List<SqlParameter> objParamList = new List<SqlParameter>();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "List<SqlParameter> objParamListOut = new List<SqlParameter>();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "MyCLS.clsExecuteStoredProcSql clsESPSql = new MyCLS.clsExecuteStoredProcSql();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "int Result=0;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "try" & vbCrLf & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "LIB" & dbInfo(i).TABLENAME & " objLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strclassDALInsert = strclassDALInsert & "objLIB" & dbInfo(i).TABLENAME & " = (LIB" & dbInfo(i).TABLENAME & ")Packet.MessagePacket;" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//*******COPY IT TO USE BELOW FUNCTION - INSERT************" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//try" & vbCrLf & "//{" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    LIB" & dbInfo(i).TABLENAME & " objLIB" & dbInfo(i).TABLENAME & " = new LIB" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    DAL" & dbInfo(i).TABLENAME & " objDAL" & dbInfo(i).TABLENAME & " = new DAL" & dbInfo(i).TABLENAME & "();" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    MyCLS.TransportationPacket tp = new MyCLS.TransportationPacket();" & vbCrLf & vbCrLf

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    'MsgBox(dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " : " & dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE & " : " & dbInfo(i).COLUMNDETAILS(j).COLDataTypeSQL)
                    If dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE = "string" Then
                        strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """].ToString();" & vbCrLf
                        strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """].ToString();" & vbCrLf
                        strclassDALInsert = strclassDALInsert & "objParamList.Add(new SqlParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "));" & vbCrLf
                    Else
                        strclassDALSelectByValue = strclassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " =(" & fnTypeCasting(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE) & ")ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """];" & vbCrLf
                        strclassDALSelect = strclassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing[i]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " =(" & fnTypeCasting(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE) & ")ds.Tables[0].Rows[i][""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """];" & vbCrLf
                        strclassDALInsert = strclassDALInsert & "objParamList.Add(new SqlParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "));" & vbCrLf
                    End If
                    'strClassDALInsert2Use = strClassDALInsert2Use & "//    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf
                    'strClassDALSelect2Use = strClassDALSelect2Use & "////    objLIB" & dbInfo(i).TABLENAME & "Listing[" & j & "]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf
                    'strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "////    objLIB" & dbInfo(i).TABLENAME & "Listing[" & j & "]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf

                    'strClassDALInsert2Use = strClassDALInsert2Use & "//    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & IIf(dbInfo(i).COLUMNDETAILS(j).COLDataType = "byte()", "MyCLS.clsImaging.PictureBoxToByteArray()", "txt.Text") & ";" & vbCrLf
                    strClassDALInsert2Use = strClassDALInsert2Use & "//    objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & IIf(dbInfo(i).COLUMNDETAILS(j).COLDataType = "byte()", "MyCLS.clsImaging.PictureBoxToByteArray()", "model." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME) & ";" & vbCrLf
                    strClassDALSelect2Use = strClassDALSelect2Use & "////  txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing[0]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ";" & vbCrLf
                    strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "////  txt.Text = " & "objLIB" & dbInfo(i).TABLENAME & "Listing[0]." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & ";" & vbCrLf

                    Application.DoEvents()
                Next

                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageId = 1;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "else" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageId = -1;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageResultsetDS = ds;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing;" & vbCrLf & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "catch (Exception ex)" & vbCrLf & "{" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.MessageId = -1;" & vbCrLf

                strclassDALSelectByValue = strclassDALSelectByValue & "Packet.ex = ex;" & vbCrLf

                strclassDALSelectByValue = strclassDALSelectByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "return Packet;" & vbCrLf
                strclassDALSelectByValue = strclassDALSelectByValue & "}" & vbCrLf & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    tp.MessagePacket = 1;    //ID to be Passed" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details(tp);" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    if(tp.MessageId == 1)" & vbCrLf & "//{" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//        objLIB" & dbInfo(i).TABLENAME & "Listing = (LIB" & dbInfo(i).TABLENAME & "Listing)tp.MessageResultset;" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//        ds = (DataSet)tp.MessageResultsetDS;" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//        MessageBox.Show(objLIB" & dbInfo(i).TABLENAME & "Listing[0].ToString());" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    }" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    }" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//catch(Exception ex)" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    {" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strClassDALSelectByValue2Use = strClassDALSelectByValue2Use & "//    }" & vbCrLf

                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageId = 1;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "else" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageId = -1;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageResultsetDS = ds;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageResultset = objLIB" & dbInfo(i).TABLENAME & "Listing;" & vbCrLf & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "catch (Exception ex)" & vbCrLf & "{" & vbCrLf
                strclassDALSelect = strclassDALSelect & "Packet.MessageId = -1;" & vbCrLf

                strclassDALSelect = strclassDALSelect & "Packet.ex = ex;" & vbCrLf

                strclassDALSelect = strclassDALSelect & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf
                strclassDALSelect = strclassDALSelect & "return Packet;" & vbCrLf
                strclassDALSelect = strclassDALSelect & "}" & vbCrLf & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    tp = objDAL" & dbInfo(i).TABLENAME & ".Get" & dbInfo(i).TABLENAME & "Details();" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    if(tp.MessageId == 1)" & vbCrLf & "//{" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//        objLIB" & dbInfo(i).TABLENAME & "Listing = (LIB" & dbInfo(i).TABLENAME & "Listing)tp.MessageResultset;" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//        ds = (DataSet)tp.MessageResultsetDS;" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//        MessageBox.Show(objLIB" & dbInfo(i).TABLENAME & "Listing[0].ToString());" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    }" & vbCrLf & "//}" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//catch(Exception ex)" & vbCrLf & "//{" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strClassDALSelect2Use = strClassDALSelect2Use & "//}" & vbCrLf

                strclassDALInsert = strclassDALInsert & "objParamListOut.Add(new SqlParameter(""@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & """, SqlDbType." & MyCLS.clsDBOperations.GetDBTypeValue4SqlDbTypes_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & "));" & vbCrLf
                'strclassDALInsert = strclassDALInsert & "//MyCLS.clsCOMMON.ConOpen(true);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "strOutParamValues = clsESPSql.ExecuteSPNonQueryOutPut(""SP_Insert" & dbInfo(i).TABLENAME & """, objParamList, objParamListOut, ref Result);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "Packet.MessageId = Result;" & vbCrLf
                'strclassDALInsert = strclassDALInsert & "//MyCLS.clsCOMMON.ConClose();" & vbCrLf & vbCrLf
                strclassDALInsert = strclassDALInsert & "Packet.MessageResultset = strOutParamValues;" & vbCrLf & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf
                strclassDALInsert = strclassDALInsert & "catch (Exception ex)" & vbCrLf
                strclassDALInsert = strclassDALInsert & "{" & vbCrLf
                strclassDALInsert = strclassDALInsert & "Packet.MessageId = -1;" & vbCrLf

                strclassDALInsert = strclassDALInsert & "Packet.ex = ex;" & vbCrLf

                strclassDALInsert = strclassDALInsert & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf
                strclassDALInsert = strclassDALInsert & "return Packet;" & vbCrLf
                strclassDALInsert = strclassDALInsert & "}" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    tp.MessagePacket = objLIB" & dbInfo(i).TABLENAME & ";" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    tp = objDAL" & dbInfo(i).TABLENAME & ".Insert" & dbInfo(i).TABLENAME & "(tp);" & vbCrLf & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    if(tp.MessageId > -1)" & vbCrLf & "//{" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//        string[] strOutParamValues = (string[])tp.MessageResultset;" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//        MessageBox.Show(strOutParamValues[0].ToString());" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    }" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    }" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//catch(Exception ex)" & vbCrLf & "//{" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//    MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString(), true);" & vbCrLf
                strClassDALInsert2Use = strClassDALInsert2Use & "//}" & vbCrLf

                strclassDALEnd = strclassDALEnd & "}" & vbCrLf
                strclassDALEnd = strclassDALEnd & "}"

                strclassDAL = strclassDALStart & strClassDALSelect2Use & vbCrLf & strclassDALSelect & vbCrLf & strClassDALSelectByValue2Use & vbCrLf & strclassDALSelectByValue & vbCrLf & strClassDALInsert2Use & vbCrLf & strclassDALInsert & strclassDALEnd

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".cs")
                MyCLS.clsFileHandling.WriteFile(strclassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub


    Function fnTypeCasting(ByVal DataType As String) As String
        Try
            Dim RetDataType As String = DataType
            If DataType = "integer" Then
                RetDataType = "int"
            ElseIf DataType = "decimal" Or DataType = "double" Then
                RetDataType = "decimal"
            ElseIf DataType = "date" Then
                RetDataType = "DateTime"
            ElseIf DataType = "boolean" Then
                RetDataType = "Boolean"
            End If
            Return RetDataType
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
        End Try
    End Function

    ''' <summary>
    ''' CREATE DATA ACCESS LAYER IN VB
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks>CREATES SELECT, SELECT BY VALUE AND INSERT FUNCTIONs</remarks>
    Sub WriteDALInCS4Access(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassDAL As String = ""
            Dim strClassDALStart As String = ""
            Dim strClassDALEnd As String = ""
            Dim strClassDALSelect As String = ""
            Dim strClassDALSelectByValue As String = ""
            Dim strClassDALSelectByKey As String = ""
            Dim strClassDALInsert As String = ""
            Dim strClassDALInsertFields As String = ""
            Dim strClassDALInsertFieldsValues As String = ""
            Dim strClassDALUpdateFields As String = ""
            Dim strClassDALUpdateFieldsValues As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassDAL = ""
                strClassDALStart = ""
                strClassDALEnd = ""
                strClassDALSelect = ""
                strClassDALSelectByValue = ""
                strClassDALSelectByKey = ""
                strClassDALInsert = ""
                strClassDALInsertFields = ""
                strClassDALInsertFieldsValues = ""
                strClassDALUpdateFields = ""
                strClassDALUpdateFieldsValues = ""

                strClassDALStart = "Imports NDS.LIB" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data" & vbCrLf
                strClassDALStart = strClassDALStart & "Imports System.Data.OleDb" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Namespace NDS.DAL" & vbCrLf & vbCrLf
                strClassDALStart = strClassDALStart & "Public Class DAL" & dbInfo(i).TABLENAME & vbCrLf & vbCrLf



                strClassDALSelectByValue = strClassDALSelectByValue & "Public Function Get" & dbInfo(i).TABLENAME & "Details(ByVal Id As Int16) as LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim ds As New DataSet" & vbCrLf
                'strClassDALSelectByValue = strClassDALSelectByValue & "Dim objParamList As New List(Of OleDbParameter)()" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Try" & vbCrLf
                'strClassDALSelectByValue = strClassDALSelectByValue & "objParamList.Add(New OleDbParameter(""@Id"", Id))" & vbCrLf
                'strClassDALSelectByValue = strClassDALSelectByValue & "ds = MyCLS.clsExecuteStoredProc.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList)" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim strQ As String" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "strQ = ""Select * From " & dbInfo(i).TABLENAME & " Where " & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=" & """ & Id" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsCOMMON.ConOpen()" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsCOMMON.prcQuerySelectDS(ds, strQ,""" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsCOMMON.ConClose()" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf

                strClassDALSelectByKey = strClassDALSelectByKey & "Public Function Get" & dbInfo(i).TABLENAME & "DetailsByKeyword(ByVal Key As String) as LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Dim ds As New DataSet" & vbCrLf
                'strClassDALSelectByKey = strClassDALSelectByKey & "Dim objParamList As New List(Of OleDbParameter)()" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Try" & vbCrLf
                'strClassDALSelectByKey = strClassDALSelectByKey & "objParamList.Add(New OleDbParameter(""@Id"", Id))" & vbCrLf
                'strClassDALSelectByKey = strClassDALSelectByKey & "ds = MyCLS.clsExecuteStoredProc.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById"", objParamList)" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Dim strQ As String" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "strQ = ""Select * From " & dbInfo(i).TABLENAME & " Where "" & _" & vbCrLf & vbTab & vbTab & """"


                strClassDALSelect = strClassDALSelect & "Public Function Get" & dbInfo(i).TABLENAME & "Details() as LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim ds As New DataSet" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim objLIB" & dbInfo(i).TABLENAME & "Listing As New LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Try" & vbCrLf
                'strClassDALSelect = strClassDALSelect & "ds = MyCLS.clsExecuteStoredProc.ExecuteSPDataSet(""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim strQ As String" & vbCrLf
                strClassDALSelect = strClassDALSelect & "strQ = ""Select * From " & dbInfo(i).TABLENAME & """" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsCOMMON.ConOpen()" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsCOMMON.prcQuerySelectDS(ds, strQ,""" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsCOMMON.ConClose()" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelect = strClassDALSelect & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf

                strClassDALInsert = strClassDALInsert & "Public Function Insert" & dbInfo(i).TABLENAME & "(ByVal objLIB" & dbInfo(i).TABLENAME & " As LIB" & dbInfo(i).TABLENAME & ") As Int16" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Dim strQ As String" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Try" & vbCrLf
                strClassDALInsert = strClassDALInsert & "MyCLS.clsCOMMON.ConOpen()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "If objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "= -1 Then" & vbCrLf
                strClassDALInsert = strClassDALInsert & "strQ = ""INSERT INTO " & dbInfo(i).TABLENAME & """ & _" & vbCrLf
                strClassDALInsertFields = """ ("
                strClassDALInsertFieldsValues = """ Values("
                strClassDALUpdateFields = "strQ = ""UPDATE " & dbInfo(i).TABLENAME & """ & _" & vbCrLf & """ SET "

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strClassDALSelectByValue = strClassDALSelectByValue & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    strClassDALSelectByKey = strClassDALSelectByKey & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " Like '%"" & Key & ""%' Or "

                    strClassDALSelect = strClassDALSelect & "objLIB" & dbInfo(i).TABLENAME & "Listing(i)." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = ds.Tables(0).Rows(i)(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).ToString" & vbCrLf

                    'strClassDALInsert = strClassDALInsert & "objParamList.Add(New OleDbParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "))" & vbCrLf
                    strClassDALInsertFields = strClassDALInsertFields & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ","
                    strClassDALInsertFieldsValues = strClassDALInsertFieldsValues & "'"" & objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " & ""',"

                    strClassDALUpdateFieldsValues = strClassDALUpdateFieldsValues & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = " & "'"" & objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " & ""',"

                    Application.DoEvents()
                Next
                strClassDALInsertFields = Mid(strClassDALInsertFields, 1, Len(strClassDALInsertFields) - 1) & ")"""
                strClassDALInsertFieldsValues = Mid(strClassDALInsertFieldsValues, 1, Len(strClassDALInsertFieldsValues) - 1) & ")"""
                strClassDALUpdateFields = strClassDALUpdateFields & Mid(strClassDALUpdateFieldsValues, 1, Len(strClassDALUpdateFieldsValues) - 1) & """ & _" & vbCrLf & """ Where " & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " = " & """ & objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " & """
                strClassDALInsert = strClassDALInsert & strClassDALInsertFields & " & _" & vbCrLf
                strClassDALInsert = strClassDALInsert & strClassDALInsertFieldsValues & vbCrLf
                strClassDALInsert = strClassDALInsert & "Return MyCLS.fnQueryInsert(strQ)" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Else" & vbCrLf
                strClassDALInsert = strClassDALInsert & strClassDALUpdateFields & vbCrLf
                strClassDALInsert = strClassDALInsert & "Return MyCLS.fnQueryUpdate(strQ)" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End If" & vbCrLf
                strClassDALInsert = strClassDALInsert & "MyCLS.clsCOMMON.ConClose()" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Catch ex As Exception" & vbCrLf
                strClassDALInsert = strClassDALInsert & "Return -1" & vbCrLf
                strClassDALInsert = strClassDALInsert & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Try" & vbCrLf
                'strClassDALInsert = strClassDALInsert & "Return strOutParamValues" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Function" & vbCrLf & vbCrLf


                strClassDALSelectByValue = strClassDALSelectByValue & "Next" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End If" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Catch ex As Exception" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Try" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "Return objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Function" & vbCrLf & vbCrLf


                strClassDALSelectByKey = Mid(strClassDALSelectByKey, 1, Len(strClassDALSelectByKey) - 4) & """" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "MyCLS.clsCOMMON.ConOpen()" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "MyCLS.clsCOMMON.prcQuerySelectDS(ds, strQ,""" & dbInfo(i).TABLENAME & """)" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "MyCLS.clsCOMMON.ConClose()" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "If ds IsNot Nothing Then" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "If ds.Tables IsNot Nothing Then" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "If ds.Tables(0).Rows IsNot Nothing Then" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "If ds.Tables(0).Rows.Count > 0 Then" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Dim oLIB" & dbInfo(i).TABLENAME & " As New LIB" & dbInfo(i).TABLENAME & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "objLIB" & dbInfo(i).TABLENAME & "Listing.Add(oLIB" & dbInfo(i).TABLENAME & ")" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Next" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End If" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End If" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End If" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End If" & vbCrLf & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Catch ex As Exception" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End Try" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "Return objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelectByKey = strClassDALSelectByKey & "End Function" & vbCrLf & vbCrLf


                strClassDALSelect = strClassDALSelect & "Next" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End If" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & "Catch ex As Exception" & vbCrLf
                strClassDALSelect = strClassDALSelect & "MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Try" & vbCrLf
                strClassDALSelect = strClassDALSelect & "Return objLIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Function" & vbCrLf & vbCrLf



                strClassDALEnd = strClassDALEnd & "End Class" & vbCrLf
                strClassDALEnd = strClassDALEnd & "End Namespace"

                strClassDAL = strClassDALStart & strClassDALSelect & vbCrLf & strClassDALSelectByValue & vbCrLf & strClassDALInsert & vbCrLf & strClassDALSelectByKey & strClassDALEnd

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".vb")
                MyCLS.clsFileHandling.WriteFile(strClassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    ''' <summary>
    ''' CREATE PROPERTIES LIBRARY IN VB6.0
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks></remarks>
    Sub WritePropertyInVB6(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassLIB As String = ""
            Dim strClassLIBStart As String = ""
            Dim strClassLIBProp As String = ""
            Dim strClassLIBVars As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                'strClassLIB = "Namespace NDS.LIB" & vbCrLf & vbCrLf
                'strClassLIB = strClassLIB & "Public Class LIB" & dbInfo(i).TABLENAME & vbCrLf & vbCrLf
                strClassLIBStart = "VERSION 1.0 CLASS" & vbCrLf
                strClassLIBStart = strClassLIBStart & "BEGIN" & vbCrLf
                strClassLIBStart = strClassLIBStart & "MultiUse = -1 'True" & vbCrLf
                strClassLIBStart = strClassLIBStart & "Persistable = 0 'NotPersistable" & vbCrLf
                strClassLIBStart = strClassLIBStart & "DataBindingBehavior = 0 'vbNone" & vbCrLf
                strClassLIBStart = strClassLIBStart & "DataSourceBehavior = 0 'vbNone" & vbCrLf
                strClassLIBStart = strClassLIBStart & "MTSTransactionMode = 0 'NotAnMTSObject" & vbCrLf
                strClassLIBStart = strClassLIBStart & "END" & vbCrLf
                strClassLIBStart = strClassLIBStart & "Attribute VB_Name = ""LIB" & dbInfo(i).TABLENAME & """" & vbCrLf
                strClassLIBStart = strClassLIBStart & "Attribute VB_GlobalNameSpace = False" & vbCrLf
                strClassLIBStart = strClassLIBStart & "Attribute VB_Creatable = True" & vbCrLf
                strClassLIBStart = strClassLIBStart & "Attribute VB_PredeclaredId = False" & vbCrLf
                strClassLIBStart = strClassLIBStart & "Attribute VB_Exposed = False" & vbCrLf

                strClassLIBStart = strClassLIBStart & "Option Explicit" & vbCrLf & vbCrLf

                strClassLIBVars = ""
                strClassLIBProp = ""

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    'MsgBox(dbInfo(i).TABLENAME & vbCrLf & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf & MyCLS.clsDBOperations.GetDBTypeValue4VB6(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE))

                    strClassLIBVars = strClassLIBVars & "Private p_" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " as " & MyCLS.clsDBOperations.GetDBTypeValue4VB6(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE) & vbCrLf
                    strClassLIBProp = strClassLIBProp & "Public Property Get " & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "() as " & MyCLS.clsDBOperations.GetDBTypeValue4VB6(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE) & vbCrLf
                    strClassLIBProp = strClassLIBProp & vbTab & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = p_" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf
                    strClassLIBProp = strClassLIBProp & "End Property" & vbCrLf

                    strClassLIBProp = strClassLIBProp & "Public Property Let " & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "(ByVal vNewValue as " & MyCLS.clsDBOperations.GetDBTypeValue4VB6(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE) & ")" & vbCrLf
                    strClassLIBProp = strClassLIBProp & vbTab & "p_" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "=vNewValue" & vbCrLf
                    strClassLIBProp = strClassLIBProp & "End Property" & vbCrLf & vbCrLf

                    Application.DoEvents()
                Next

                'strClassLIB = strClassLIB & "<Serializable()> _" & vbCrLf
                'strClassLIB = strClassLIB & "Public Class LIB" & dbInfo(i).TABLENAME & "Listing" & vbCrLf
                'strClassLIB = strClassLIB & " Inherits List(Of LIB" & dbInfo(i).TABLENAME & ")" & vbCrLf & vbCrLf

                '***vb6.0**********
                'Dim ObjectList As New Collection
                'Dim op As Object
                'op = New ObjectOfAnyKind
                'ObjectList.Add op [, "KeyForThisObject"]

                strClassLIB = strClassLIBStart & strClassLIBVars & vbCrLf & strClassLIBProp

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\LIB\LIB" & dbInfo(i).TABLENAME & ".cls")
                MyCLS.clsFileHandling.WriteFile(strClassLIB)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    ''' <summary>
    ''' CREATE DATA ACCESS LAYER IN VB6
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks>CREATES SELECT, SELECT BY VALUE AND INSERT FUNCTIONs</remarks>
    Sub WriteDALInVB6(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strClassDAL As String = ""
            Dim strClassDALStart As String = ""
            Dim strClassDALEnd As String = ""
            Dim strClassDALSelect As String = ""
            Dim strClassDALSelectByValue As String = ""
            Dim strClassDALSelectByValueParam As String = ""
            Dim strClassDALSelectByValueParamUse As String = ""
            Dim strClassDALInsert As String = ""
            Dim strClassDALInsertUse As String = ""
            For i As Integer = 0 To dbInfo.Count - 1
                strClassDAL = ""
                strClassDALStart = ""
                strClassDALEnd = ""
                strClassDALSelect = ""
                strClassDALSelectByValue = ""
                strClassDALInsert = ""
                strClassDALSelectByValueParam = ""
                strClassDALSelectByValueParamUse = ""
                strClassDALInsertUse = ""

                strClassDALStart = "VERSION 1.0 CLASS" & vbCrLf
                strClassDALStart = strClassDALStart & "BEGIN" & vbCrLf
                strClassDALStart = strClassDALStart & "MultiUse = -1 'True" & vbCrLf
                strClassDALStart = strClassDALStart & "Persistable = 0 'NotPersistable" & vbCrLf
                strClassDALStart = strClassDALStart & "DataBindingBehavior = 0 'vbNone" & vbCrLf
                strClassDALStart = strClassDALStart & "DataSourceBehavior = 0 'vbNone" & vbCrLf
                strClassDALStart = strClassDALStart & "MTSTransactionMode = 0 'NotAnMTSObject" & vbCrLf
                strClassDALStart = strClassDALStart & "END" & vbCrLf
                strClassDALStart = strClassDALStart & "Attribute VB_Name = ""DAL" & dbInfo(i).TABLENAME & """" & vbCrLf
                strClassDALStart = strClassDALStart & "Attribute VB_GlobalNameSpace = False" & vbCrLf
                strClassDALStart = strClassDALStart & "Attribute VB_Creatable = True" & vbCrLf
                strClassDALStart = strClassDALStart & "Attribute VB_PredeclaredId = False" & vbCrLf
                strClassDALStart = strClassDALStart & "Attribute VB_Exposed = False" & vbCrLf

                strClassDALStart = strClassDALStart & "Option Explicit" & vbCrLf & vbCrLf


                strClassDALSelectByValue = strClassDALSelectByValue & "Public Sub Get" & dbInfo(i).TABLENAME & "DetailsById(ByVal Con As ADODB.Connection, ByRef Rs As ADODB.Recordset, objLIB" & dbInfo(i).TABLENAME & " As LIB" & dbInfo(i).TABLENAME & ")" & vbCrLf
                'strClassDALSelect = strClassDALSelect & vbTab & "Dim Rs As New Recordset" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Dim Cmd As Command" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Set Cmd = New ADODB.Command" & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Cmd.ActiveConnection = Con" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Cmd.CommandType = adCmdStoredProc" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Cmd.CommandText = ""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById""" & vbCrLf & vbCrLf



                strClassDALSelect = strClassDALSelect & "Public Sub Get" & dbInfo(i).TABLENAME & "Details(ByVal Con As ADODB.Connection, ByRef Rs As ADODB.Recordset)" & vbCrLf
                'strClassDALSelect = strClassDALSelect & vbTab & "Dim Rs As New Recordset" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "Dim Cmd As Command" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "Set Cmd = New ADODB.Command" & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "Cmd.ActiveConnection = Con" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "Cmd.CommandType = adCmdStoredProc" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "Cmd.CommandText = ""SP_GetDetailsFrom" & dbInfo(i).TABLENAME & """" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "Set Rs = Cmd.Execute" & vbCrLf & vbCrLf & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "Set Cmd.ActiveConnection = Nothing" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "'***PICK IF FROM HERE AND USE TO RETRIVE DATA***" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "Rs.MoveFirst" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "While Not Rs.EOF" & vbCrLf



                strClassDALInsert = strClassDALInsert & "Public Sub Insert" & dbInfo(i).TABLENAME & "(ByVal Con As ADODB.Connection, ByRef Result As Integer, ByVal objLIB" & dbInfo(i).TABLENAME & " As LIB" & dbInfo(i).TABLENAME & ")" & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "Dim Rs As New Recordset" & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "Dim Cmd As Command" & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "Set Cmd = New ADODB.Command" & vbCrLf & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "Cmd.ActiveConnection = Con" & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "Cmd.CommandType = adCmdStoredProc" & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "Cmd.CommandText = ""SP_Insert" & dbInfo(i).TABLENAME & """" & vbCrLf & vbCrLf



                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strClassDALSelectByValueParamUse = strClassDALSelectByValueParamUse & vbTab & vbTab & "Debug.Print Rs(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).Value" & vbCrLf

                    strClassDALSelect = strClassDALSelect & vbTab & vbTab & "Debug.Print Rs(""" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """).Value" & vbCrLf

                    strClassDALInsert = strClassDALInsert & vbTab & "Cmd.Parameters.Append Cmd.CreateParameter(""@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & """, " & MyCLS.clsDBOperations.GetDBTypeValue4VB6SPParam(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE) & ", adParamInput, 100, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ")" & vbCrLf
                    strClassDALInsertUse = strClassDALInsertUse & "objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " = """"" & vbCrLf

                    Application.DoEvents()
                Next


                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Cmd.Parameters.Append Cmd.CreateParameter(""@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & """, " & MyCLS.clsDBOperations.GetDBTypeValue4VB6SPParam(dbInfo(i).COLUMNDETAILS(0).COLUMNTYPE) & ", adParamInput, 100, objLIB" & dbInfo(i).TABLENAME & "." & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & ")" & vbCrLf & vbCrLf

                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Set Rs = Cmd.Execute" & vbCrLf & vbCrLf & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Set Cmd.ActiveConnection = Nothing" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "'***PICK IF FROM HERE AND USE TO RETRIVE DATA***" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Rs.MoveFirst" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "While Not Rs.EOF" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & strClassDALSelectByValueParamUse & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & vbTab & "Rs.MoveNext" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "Wend" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & vbTab & "'***PICK IF FROM HERE AND USE TO RETRIVE DATA***" & vbCrLf
                strClassDALSelectByValue = strClassDALSelectByValue & "End Sub" & vbCrLf & vbCrLf




                strClassDALSelect = strClassDALSelect & vbTab & vbTab & "Rs.MoveNext" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "Wend" & vbCrLf
                strClassDALSelect = strClassDALSelect & vbTab & "'***PICK IF FROM HERE AND USE TO RETRIVE DATA***" & vbCrLf
                strClassDALSelect = strClassDALSelect & "End Sub" & vbCrLf & vbCrLf




                strClassDALInsert = strClassDALInsert & vbCrLf & vbTab & "Set Rs = Cmd.Execute" & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "Result = 1" & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "Set Cmd.ActiveConnection = Nothing" & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "'***PICK IF FROM HERE AND USE TO RETRIVE DATA***" & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & strClassDALInsertUse & vbCrLf
                strClassDALInsert = strClassDALInsert & vbTab & "'***PICK IF FROM HERE AND USE TO RETRIVE DATA***" & vbCrLf
                strClassDALInsert = strClassDALInsert & "End Sub" & vbCrLf & vbCrLf


                'strClassDALEnd = strClassDALEnd & "End Class" & vbCrLf
                'strClassDALEnd = strClassDALEnd & "End Namespace"

                strClassDAL = strClassDALStart & strClassDALSelect & vbCrLf & strClassDALSelectByValue & vbCrLf & strClassDALInsert & strClassDALEnd

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\DAL\DAL" & dbInfo(i).TABLENAME & ".cls")
                MyCLS.clsFileHandling.WriteFile(strClassDAL)
                MyCLS.clsFileHandling.CloseFile()
            Next
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    ''' <summary>
    ''' CREATE STORED PROCEDURE
    ''' </summary>
    ''' <param name="dbInfo"></param>
    ''' <remarks>CREATES STORED PROCEDUREs</remarks>
    Sub WriteStoredProc(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strSP As String = ""

            Dim strSPSelect As String = ""
            Dim strSPSelectByValue As String = ""
            Dim strSPInsert As String = ""

            Dim strSPInsertCreate As String = ""
            Dim strSPInsertParamListDef As String = ""
            Dim strSPInsertBegin As String = ""
            Dim strSPInsertParamListUpdate As String = ""
            Dim strSPInsertUpdateWhere As String = ""
            Dim strSPInsertReturnUpdate As String = ""
            Dim strSPInsertParamListIns As String = ""
            Dim strSPInsertValues As String = ""
            Dim strSPInsertParamListInsValues As String = ""
            Dim strSPInsertReturnInsert As String = ""

            For i As Integer = 0 To dbInfo.Count - 1
                strSPSelect = ""
                strSPSelectByValue = ""

                strSPInsertCreate = ""
                strSPInsertParamListDef = ""
                strSPInsertBegin = ""
                strSPInsertParamListUpdate = ""
                strSPInsertUpdateWhere = ""
                strSPInsertReturnUpdate = ""
                strSPInsertParamListIns = ""
                strSPInsertValues = ""
                strSPInsertParamListInsValues = ""
                strSPInsertReturnInsert = ""

                strSPSelectByValue = strSPSelectByValue & "DROP PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById]" & vbCrLf & "Go" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "CREATE PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById] " & vbCrLf & "@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP(dbInfo(i).COLUMNDETAILS(0).COLUMNTYPE) & vbCrLf & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "AS" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "BEGIN" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & vbTab & "SELECT" & vbCrLf & vbTab & vbTab

                strSPSelect = strSPSelect & "DROP PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "]" & vbCrLf & "Go" & vbCrLf
                strSPSelect = strSPSelect & "CREATE PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "]" & vbCrLf & vbCrLf
                strSPSelect = strSPSelect & "AS" & vbCrLf
                strSPSelect = strSPSelect & "BEGIN" & vbCrLf
                strSPSelect = strSPSelect & vbTab & "SELECT" & vbCrLf & vbTab & vbTab

                strSPInsertCreate = "DROP proc [SP_Insert" & dbInfo(i).TABLENAME & "]" & vbCrLf & "Go" & vbCrLf
                strSPInsertCreate = strSPInsertCreate & "CREATE PROCEDURE [SP_Insert" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & "as" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & "BEGIN" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & "IF(EXISTS(SELECT TOP 1 1 FROM " & dbInfo(i).TABLENAME & " WHERE " & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "))" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & vbTab & vbTab & "Update [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & vbTab & vbTab & vbTab & "SET" & vbCrLf

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strSPSelectByValue = strSPSelectByValue & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "],"

                    strSPSelect = strSPSelect & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "],"

                    strSPInsertParamListDef = strSPInsertParamListDef & vbTab & "@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP(dbInfo(i).COLUMNDETAILS(j).COLUMNTYPE) & "," & vbCrLf
                    strSPInsertParamListUpdate = strSPInsertParamListUpdate & vbTab & vbTab & vbTab & vbTab & vbTab & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "] = @" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ", " & vbCrLf
                    strSPInsertParamListIns = strSPInsertParamListIns & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "],"
                    strSPInsertParamListInsValues = strSPInsertParamListInsValues & "@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ","

                    Application.DoEvents()
                Next

                strSPInsertParamListDef = strSPInsertParamListDef & vbTab & "@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP(dbInfo(i).COLUMNDETAILS(0).COLUMNTYPE) & " out," & vbCrLf
                strSPSelectByValue = Mid(strSPSelectByValue, 1, Len(strSPSelectByValue) - 1)
                strSPSelectByValue = strSPSelectByValue & vbCrLf & vbTab & vbTab & "FROM [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & vbTab & "WHERE [" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "]=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "End" & vbCrLf

                strSPSelect = Mid(strSPSelect, 1, Len(strSPSelect) - 1)
                strSPSelect = strSPSelect & vbCrLf & vbTab & vbTab & "FROM [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPSelect = strSPSelect & "End" & vbCrLf

                strSPInsertParamListDef = Mid(strSPInsertParamListDef, 1, Len(strSPInsertParamListDef) - 3)
                'strSPInsertParamListDef = Mid(strSPInsertParamListDef, 1, Len(strSPInsertParamListDef) - 1)
                strSPInsertParamListDef = strSPInsertParamListDef & vbCrLf

                strSPInsertParamListUpdate = Mid(strSPInsertParamListUpdate, 1, Len(strSPInsertParamListUpdate) - 4)
                'strSPInsertParamListUpdate = Mid(strSPInsertParamListUpdate, 1, Len(strSPInsertParamListUpdate) - 1)
                strSPInsertParamListUpdate = strSPInsertParamListUpdate & vbCrLf

                strSPInsertParamListIns = Mid(strSPInsertParamListIns, 1, Len(strSPInsertParamListIns) - 1)
                'strSPInsertParamListIns = Mid(strSPInsertParamListIns, 1, Len(strSPInsertParamListIns) - 1)
                strSPInsertParamListIns = strSPInsertParamListIns & vbCrLf

                strSPInsertParamListInsValues = Mid(strSPInsertParamListInsValues, 1, Len(strSPInsertParamListInsValues) - 1)
                'strSPInsertParamListInsValues = Mid(strSPInsertParamListInsValues, 1, Len(strSPInsertParamListInsValues) - 1)
                strSPInsertParamListInsValues = strSPInsertParamListInsValues & vbCrLf



                strSPInsertUpdateWhere = strSPInsertUpdateWhere & vbTab & vbTab & vbTab & "WHERE [" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "] = @" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf
                strSPInsertUpdateWhere = strSPInsertUpdateWhere & vbTab & vbTab & vbTab & "Set @@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & "IF @@ERROR <> 0" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & vbTab & "Return -2" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & "Return 2" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & "Else" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & "INSERT INTO [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & "(" & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
                strSPInsertValues = strSPInsertValues & vbTab & vbTab & vbTab & vbTab & ")" & vbCrLf
                strSPInsertValues = strSPInsertValues & vbTab & vbTab & vbTab & "VALUES (" & vbCrLf & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & ")" & vbCrLf & vbCrLf
                'strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & "Set @@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & "Set @@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=(Select @@Identity)" & vbCrLf & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & "IF @@ERROR <> 0" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & vbTab & "Return -1" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & "Return 1" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & "End"

                strSPInsert = strSPInsertCreate & strSPInsertParamListDef & strSPInsertBegin & strSPInsertParamListUpdate & strSPInsertUpdateWhere & strSPInsertReturnUpdate & strSPInsertParamListIns & strSPInsertValues & strSPInsertParamListInsValues & strSPInsertReturnInsert

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_GetDetailsFrom" & dbInfo(i).TABLENAME & ".sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPSelect)
                MyCLS.clsFileHandling.CloseFile()

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById.sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPSelectByValue)
                MyCLS.clsFileHandling.CloseFile()

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_Insert" & dbInfo(i).TABLENAME & ".sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPInsert)
                MyCLS.clsFileHandling.CloseFile()

                strSP = strSP & vbCrLf & strSPSelect & vbCrLf & "Go" & vbCrLf & strSPSelectByValue & vbCrLf & "Go" & vbCrLf & strSPInsert & vbCrLf & "Go" & vbCrLf
            Next

            MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\_FullScript.sql")
            MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSP)
            MyCLS.clsFileHandling.CloseFile()
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
    Sub WriteStoredProc_UPDATED(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strSP As String = ""            
            Dim strSPSelect As String = ""
            Dim strSPSelectByValue As String = ""
            Dim strSPInsert As String = ""
            Dim strSPDeleteByValue As String = ""

            Dim strSPInsertCreate As String = ""
            Dim strSPInsertParamListDef As String = ""
            Dim strSPInsertBegin As String = ""
            Dim strSPInsertParamListUpdate As String = ""
            Dim strSPInsertUpdateWhere As String = ""
            Dim strSPInsertReturnUpdate As String = ""
            Dim strSPInsertParamListIns As String = ""
            Dim strSPInsertValues As String = ""
            Dim strSPInsertParamListInsValues As String = ""
            Dim strSPInsertReturnInsert As String = ""

            For i As Integer = 0 To dbInfo.Count - 1
                strSPSelect = ""
                strSPSelectByValue = ""

                strSPDeleteByValue = ""

                strSPInsertCreate = ""
                strSPInsertParamListDef = ""
                strSPInsertBegin = ""
                strSPInsertParamListUpdate = ""
                strSPInsertUpdateWhere = ""
                strSPInsertReturnUpdate = ""
                strSPInsertParamListIns = ""
                strSPInsertValues = ""
                strSPInsertParamListInsValues = ""
                strSPInsertReturnInsert = ""

                strSPSelectByValue = strSPSelectByValue & "DROP PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById]" & vbCrLf & "Go" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "CREATE PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById] " & vbCrLf & "@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & vbCrLf & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "AS" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "BEGIN" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & vbTab & "SELECT" & vbCrLf & vbTab & vbTab

                strSPDeleteByValue = strSPDeleteByValue & "DROP PROCEDURE [SP_Delete" & dbInfo(i).TABLENAME & "]" & vbCrLf & "Go" & vbCrLf
                strSPDeleteByValue = strSPDeleteByValue & "CREATE PROCEDURE [SP_Delete" & dbInfo(i).TABLENAME & "] " & vbCrLf & vbTab & "@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & vbCrLf & vbCrLf
                strSPDeleteByValue = strSPDeleteByValue & "AS" & vbCrLf
                strSPDeleteByValue = strSPDeleteByValue & "BEGIN" & vbCrLf
                strSPDeleteByValue = strSPDeleteByValue & vbTab & "Delete From " & dbInfo(i).TABLENAME & " Where " & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf

                strSPSelect = strSPSelect & "DROP PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "]" & vbCrLf & "Go" & vbCrLf
                strSPSelect = strSPSelect & "CREATE PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "]" & vbCrLf & vbCrLf
                strSPSelect = strSPSelect & "AS" & vbCrLf
                strSPSelect = strSPSelect & "BEGIN" & vbCrLf
                strSPSelect = strSPSelect & vbTab & "SELECT" & vbCrLf & vbTab & vbTab

                strSPInsertCreate = "DROP proc [SP_Insert" & dbInfo(i).TABLENAME & "]" & vbCrLf & "Go" & vbCrLf
                strSPInsertCreate = strSPInsertCreate & "CREATE PROCEDURE [SP_Insert" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & "as" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & "BEGIN" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & "IF(EXISTS(SELECT TOP 1 1 FROM " & dbInfo(i).TABLENAME & " WHERE " & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "))" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & vbTab & vbTab & "Update [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & vbTab & vbTab & vbTab & "SET" & vbCrLf

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strSPSelectByValue = strSPSelectByValue & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "],"

                    strSPSelect = strSPSelect & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "],"

                    strSPInsertParamListDef = strSPInsertParamListDef & vbTab & "@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP_UPDATED(dbInfo(i).COLUMNDETAILS(j)) & "," & vbCrLf
                    If dbInfo(i).COLUMNDETAILS(j).COLDataTypeSQL.ToUpper <> "TIMESTAMP" And dbInfo(i).COLUMNDETAILS(j).COLUMNNAME.ToUpper <> "ID" Then ' dbInfo(i).COLUMNDETAILS(j).COLIsAutoIncrement <> True Then
                        strSPInsertParamListUpdate = strSPInsertParamListUpdate & vbTab & vbTab & vbTab & vbTab & vbTab & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "] = @" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ", " & vbCrLf
                        strSPInsertParamListIns = strSPInsertParamListIns & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "],"
                        strSPInsertParamListInsValues = strSPInsertParamListInsValues & "@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ","
                    End If

                    Application.DoEvents()
                Next

                strSPInsertParamListDef = strSPInsertParamListDef & vbTab & "@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & " out," & vbCrLf
                strSPSelectByValue = Mid(strSPSelectByValue, 1, Len(strSPSelectByValue) - 1)
                strSPSelectByValue = strSPSelectByValue & vbCrLf & vbTab & vbTab & "FROM [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & vbTab & "WHERE [" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "]=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "End" & vbCrLf

                strSPDeleteByValue = strSPDeleteByValue & "End" & vbCrLf

                strSPSelect = Mid(strSPSelect, 1, Len(strSPSelect) - 1)
                strSPSelect = strSPSelect & vbCrLf & vbTab & vbTab & "FROM [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPSelect = strSPSelect & "End" & vbCrLf

                strSPInsertParamListDef = Mid(strSPInsertParamListDef, 1, Len(strSPInsertParamListDef) - 3)
                'strSPInsertParamListDef = Mid(strSPInsertParamListDef, 1, Len(strSPInsertParamListDef) - 1)
                strSPInsertParamListDef = strSPInsertParamListDef & vbCrLf

                strSPInsertParamListUpdate = Mid(strSPInsertParamListUpdate, 1, Len(strSPInsertParamListUpdate) - 4)
                'strSPInsertParamListUpdate = Mid(strSPInsertParamListUpdate, 1, Len(strSPInsertParamListUpdate) - 1)
                strSPInsertParamListUpdate = strSPInsertParamListUpdate & vbCrLf

                strSPInsertParamListIns = Mid(strSPInsertParamListIns, 1, Len(strSPInsertParamListIns) - 1)
                'strSPInsertParamListIns = Mid(strSPInsertParamListIns, 1, Len(strSPInsertParamListIns) - 1)
                strSPInsertParamListIns = strSPInsertParamListIns & vbCrLf

                strSPInsertParamListInsValues = Mid(strSPInsertParamListInsValues, 1, Len(strSPInsertParamListInsValues) - 1)
                'strSPInsertParamListInsValues = Mid(strSPInsertParamListInsValues, 1, Len(strSPInsertParamListInsValues) - 1)
                strSPInsertParamListInsValues = strSPInsertParamListInsValues & vbCrLf



                strSPInsertUpdateWhere = strSPInsertUpdateWhere & vbTab & vbTab & vbTab & "WHERE [" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "] = @" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf
                strSPInsertUpdateWhere = strSPInsertUpdateWhere & vbTab & vbTab & vbTab & "Set @@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & "IF @@ERROR <> 0" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & vbTab & "Return -2" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & "Return 2" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & "Else" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & "INSERT INTO [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & "(" & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
                strSPInsertValues = strSPInsertValues & vbTab & vbTab & vbTab & vbTab & ")" & vbCrLf
                strSPInsertValues = strSPInsertValues & vbTab & vbTab & vbTab & "VALUES (" & vbCrLf & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & ")" & vbCrLf & vbCrLf
                'strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & "Set @@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & "Set @@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=(Select @@Identity)" & vbCrLf & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & "IF @@ERROR <> 0" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & vbTab & "Return -1" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & "Return 1" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & "End"

                strSPInsert = strSPInsertCreate & strSPInsertParamListDef & strSPInsertBegin & strSPInsertParamListUpdate & strSPInsertUpdateWhere & strSPInsertReturnUpdate & strSPInsertParamListIns & strSPInsertValues & strSPInsertParamListInsValues & strSPInsertReturnInsert

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_GetDetailsFrom" & dbInfo(i).TABLENAME & ".sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPSelect)
                MyCLS.clsFileHandling.CloseFile()

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById.sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPSelectByValue)
                MyCLS.clsFileHandling.CloseFile()

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_Delete" & dbInfo(i).TABLENAME & ".sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPDeleteByValue)
                MyCLS.clsFileHandling.CloseFile()

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_Insert" & dbInfo(i).TABLENAME & ".sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPInsert)
                MyCLS.clsFileHandling.CloseFile()

                strSP = strSP & vbCrLf & strSPSelect & vbCrLf & "Go" & vbCrLf & strSPSelectByValue & vbCrLf & "Go" & vbCrLf & strSPInsert & vbCrLf & "Go" & vbCrLf & strSPDeleteByValue & vbCrLf & "Go" & vbCrLf
            Next

            MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\_FullScript.sql")
            MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSP)
            MyCLS.clsFileHandling.CloseFile()
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    Sub CreateColList(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strCols As String = ""
            Dim strColsAll As String = ""

            For i As Integer = 0 To dbInfo.Count - 1
                strCols = dbInfo(i).TABLENAME & vbCrLf
                strColsAll = strColsAll & dbInfo(i).TABLENAME & vbCrLf

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strCols = strCols & vbTab & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf
                    strColsAll = strColsAll & vbTab & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & vbCrLf

                    Application.DoEvents()
                Next

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\Columns\_Columns" & dbInfo(i).TABLENAME & ".txt")
                MyCLS.clsFileHandling.WriteFile(strCols)
                MyCLS.clsFileHandling.CloseFile()
            Next

            Dim strAllColumnsFileName As String = ""
            If optMSAccess.Checked = True Then
                strAllColumnsFileName = MyCLS.clsCOMMON.fnGetFileName(txtFile.Text)
            ElseIf optMSSql.Checked = True Then
                strAllColumnsFileName = txtDatabase.Text
            ElseIf optOracle.Checked = True Then
                strAllColumnsFileName = txtServerOra.Text
            End If

            MyCLS.clsFileHandling.OpenFile("C:\_CODE\Columns\_AllColumnsOf_" & strAllColumnsFileName & ".txt")
            MyCLS.clsFileHandling.WriteFile(strColsAll)
            MyCLS.clsFileHandling.CloseFile()
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub


    '***********************************************************************************************************************************
    '***************ALTER TABLE COLLATE*************************************************************************************************
    '***********************************************************************************************************************************
    Private Sub cmdAlterTableCollate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAlterTableCollate.Click
        Try
            lblMSG.Text = "Validating..."
            If fnValidate() = True Then
                '***SAVE DATABASE SETTINGS***
                MyCLS.clsCOMMON.SaveSettings(txtFile.Text, txtServer.Text, txtUID.Text, txtPassword.Text, txtDatabase.Text)

                cmdCreate.Enabled = False
                cmdCancel.Enabled = False
                'DELETE DIR
                Try
                    IO.Directory.Delete("C:\_CODE", True)
                    'IO.Directory.Delete("C:\_CODE\LIB", True)
                    'IO.Directory.Delete("C:\_CODE\DAL", True)
                    'IO.Directory.Delete("C:\_CODE\SQL", True)
                Catch ex As Exception

                End Try
                'CREATE DIR
                Try
                    lblMSG.Text = "Dir Creation..."
                    IO.Directory.CreateDirectory("C:\_CODE")
                    IO.Directory.CreateDirectory("C:\_CODE\LIB")
                    IO.Directory.CreateDirectory("C:\_CODE\DAL")
                    IO.Directory.CreateDirectory("C:\_CODE\SQL")
                    IO.Directory.CreateDirectory("C:\_CODE\Columns")
                Catch ex As Exception

                End Try

                'CREATE CONNECTION STRING
                lblMSG.Text = "Creating Conn String..."
                MyCLS.strConnStringOLEDB = CreateConnString()
                MyCLS.strConnStringSQLCLIENT = CreateConnString().Replace("Provider=SQLOLEDB;", "")

                lblMSG.Text = "Opening Connection..."
                MyCLS.clsCOMMON.ConOpen(False)

                If Len(MyCLS.strGlobalErrorInfo) > 0 Then
                    cmdCreate.Enabled = True
                    cmdCancel.Enabled = True
                    cmdAlterTableCollate.Enabled = True '***FOR COLLATE
                    lblMSG.Text = "Finished!"
                    'MsgBox("Not Done!", MsgBoxStyle.Information, "Not Completed")
                    Exit Sub
                End If

                ''GET ALL THE TABLES
                'Dim str As String() = MyCLS.clsDBOperations.GetTables()

                'GET DETAILED DATABASE IN A CLASS
                Dim dbInfo As New List(Of MyCLS.clsTables)
                lblMSG.Text = "Fetching Data From Database..."
                dbInfo = MyCLS.clsDBOperations.FillDetails()



                '***********************************************************************************************
                '******CREATE ALTER TABLE COLLATE***************************************************************
                lblMSG.Text = "Writing COLLATE Files..."                            
                Call WriteAlterTableCollate(dbInfo)
                '***********************************************************************************************
                '******CREATE ALTER TABLE COLLATE***************************************************************

                MyCLS.clsCOMMON.ConClose()
                lblMSG.Text = "Opening Files..."
                Shell("Explorer C:\_CODE", AppWinStyle.MaximizedFocus)

                cmdCreate.Enabled = True
                cmdCancel.Enabled = True
                lblMSG.Text = "Finished!"
                'MsgBox("Done!", MsgBoxStyle.Information, "Completed")
            End If
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub

    Sub WriteAlterTableCollate(ByVal dbInfo As List(Of MyCLS.clsTables))
        Try
            Dim strSP As String = ""
            Dim strSPSelect As String = ""
            Dim strSPSelectByValue As String = ""
            Dim strSPInsert As String = ""
            Dim strSPDeleteByValue As String = ""

            Dim strSPInsertCreate As String = ""
            Dim strSPInsertParamListDef As String = ""
            Dim strSPInsertBegin As String = ""
            Dim strSPInsertParamListUpdate As String = ""
            Dim strSPInsertUpdateWhere As String = ""
            Dim strSPInsertReturnUpdate As String = ""
            Dim strSPInsertParamListIns As String = ""
            Dim strSPInsertValues As String = ""
            Dim strSPInsertParamListInsValues As String = ""
            Dim strSPInsertReturnInsert As String = ""

            For i As Integer = 0 To dbInfo.Count - 1
                strSPSelect = ""
                strSPSelectByValue = ""

                strSPDeleteByValue = ""

                strSPInsertCreate = ""
                strSPInsertParamListDef = ""
                strSPInsertBegin = ""
                strSPInsertParamListUpdate = ""
                strSPInsertUpdateWhere = ""
                strSPInsertReturnUpdate = ""
                strSPInsertParamListIns = ""
                strSPInsertValues = ""
                strSPInsertParamListInsValues = ""
                strSPInsertReturnInsert = ""

                strSPSelectByValue = strSPSelectByValue & "DROP PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById]" & vbCrLf & "Go" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "CREATE PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById] " & vbCrLf & "@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & vbCrLf & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "AS" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "BEGIN" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & vbTab & "SELECT" & vbCrLf & vbTab & vbTab

                strSPDeleteByValue = strSPDeleteByValue & "DROP PROCEDURE [SP_Delete" & dbInfo(i).TABLENAME & "]" & vbCrLf & "Go" & vbCrLf
                strSPDeleteByValue = strSPDeleteByValue & "CREATE PROCEDURE [SP_Delete" & dbInfo(i).TABLENAME & "] " & vbCrLf & vbTab & "@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & vbCrLf & vbCrLf
                strSPDeleteByValue = strSPDeleteByValue & "AS" & vbCrLf
                strSPDeleteByValue = strSPDeleteByValue & "BEGIN" & vbCrLf
                strSPDeleteByValue = strSPDeleteByValue & vbTab & "Delete From " & dbInfo(i).TABLENAME & " Where " & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf

                strSPSelect = strSPSelect & "DROP PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "]" & vbCrLf & "Go" & vbCrLf
                strSPSelect = strSPSelect & "CREATE PROCEDURE [SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "]" & vbCrLf & vbCrLf
                strSPSelect = strSPSelect & "AS" & vbCrLf
                strSPSelect = strSPSelect & "BEGIN" & vbCrLf
                strSPSelect = strSPSelect & vbTab & "SELECT" & vbCrLf & vbTab & vbTab

                strSPInsertCreate = "DROP proc [SP_Insert" & dbInfo(i).TABLENAME & "]" & vbCrLf & "Go" & vbCrLf
                strSPInsertCreate = strSPInsertCreate & "CREATE PROCEDURE [SP_Insert" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & "as" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & "BEGIN" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & "IF(EXISTS(SELECT TOP 1 1 FROM " & dbInfo(i).TABLENAME & " WHERE " & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "))" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & vbTab & vbTab & "Update [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPInsertBegin = strSPInsertBegin & vbTab & vbTab & vbTab & vbTab & "SET" & vbCrLf

                For j As Int16 = 0 To dbInfo(i).COLUMNDETAILS.Count - 1
                    strSPSelectByValue = strSPSelectByValue & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "],"

                    strSPSelect = strSPSelect & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "],"

                    strSPInsertParamListDef = strSPInsertParamListDef & vbTab & "@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP_UPDATED(dbInfo(i).COLUMNDETAILS(j)) & "," & vbCrLf
                    If dbInfo(i).COLUMNDETAILS(j).COLDataTypeSQL.ToUpper <> "TIMESTAMP" And dbInfo(i).COLUMNDETAILS(j).COLUMNNAME.ToUpper <> "ID" Then ' dbInfo(i).COLUMNDETAILS(j).COLIsAutoIncrement <> True Then
                        strSPInsertParamListUpdate = strSPInsertParamListUpdate & vbTab & vbTab & vbTab & vbTab & vbTab & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "] = @" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ", " & vbCrLf
                        strSPInsertParamListIns = strSPInsertParamListIns & "[" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & "],"
                        strSPInsertParamListInsValues = strSPInsertParamListInsValues & "@" & dbInfo(i).COLUMNDETAILS(j).COLUMNNAME & ","
                    End If

                    Application.DoEvents()
                Next

                strSPInsertParamListDef = strSPInsertParamListDef & vbTab & "@@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & " " & MyCLS.clsDBOperations.GetDBTypeValue4SP_UPDATED(dbInfo(i).COLUMNDETAILS(0)) & " out," & vbCrLf
                strSPSelectByValue = Mid(strSPSelectByValue, 1, Len(strSPSelectByValue) - 1)
                strSPSelectByValue = strSPSelectByValue & vbCrLf & vbTab & vbTab & "FROM [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPSelectByValue = strSPSelectByValue & vbTab & "WHERE [" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "]=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & vbCrLf
                strSPSelectByValue = strSPSelectByValue & "End" & vbCrLf

                strSPDeleteByValue = strSPDeleteByValue & "End" & vbCrLf

                strSPSelect = Mid(strSPSelect, 1, Len(strSPSelect) - 1)
                strSPSelect = strSPSelect & vbCrLf & vbTab & vbTab & "FROM [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPSelect = strSPSelect & "End" & vbCrLf

                strSPInsertParamListDef = Mid(strSPInsertParamListDef, 1, Len(strSPInsertParamListDef) - 3)
                'strSPInsertParamListDef = Mid(strSPInsertParamListDef, 1, Len(strSPInsertParamListDef) - 1)
                strSPInsertParamListDef = strSPInsertParamListDef & vbCrLf

                strSPInsertParamListUpdate = Mid(strSPInsertParamListUpdate, 1, Len(strSPInsertParamListUpdate) - 4)
                'strSPInsertParamListUpdate = Mid(strSPInsertParamListUpdate, 1, Len(strSPInsertParamListUpdate) - 1)
                strSPInsertParamListUpdate = strSPInsertParamListUpdate & vbCrLf

                strSPInsertParamListIns = Mid(strSPInsertParamListIns, 1, Len(strSPInsertParamListIns) - 1)
                'strSPInsertParamListIns = Mid(strSPInsertParamListIns, 1, Len(strSPInsertParamListIns) - 1)
                strSPInsertParamListIns = strSPInsertParamListIns & vbCrLf

                strSPInsertParamListInsValues = Mid(strSPInsertParamListInsValues, 1, Len(strSPInsertParamListInsValues) - 1)
                'strSPInsertParamListInsValues = Mid(strSPInsertParamListInsValues, 1, Len(strSPInsertParamListInsValues) - 1)
                strSPInsertParamListInsValues = strSPInsertParamListInsValues & vbCrLf



                strSPInsertUpdateWhere = strSPInsertUpdateWhere & vbTab & vbTab & vbTab & "WHERE [" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "] = @" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf
                strSPInsertUpdateWhere = strSPInsertUpdateWhere & vbTab & vbTab & vbTab & "Set @@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & "IF @@ERROR <> 0" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & vbTab & "Return -2" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & "Return 2" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & "Else" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & "INSERT INTO [" & dbInfo(i).TABLENAME & "]" & vbCrLf
                strSPInsertReturnUpdate = strSPInsertReturnUpdate & vbTab & vbTab & vbTab & vbTab & "(" & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
                strSPInsertValues = strSPInsertValues & vbTab & vbTab & vbTab & vbTab & ")" & vbCrLf
                strSPInsertValues = strSPInsertValues & vbTab & vbTab & vbTab & "VALUES (" & vbCrLf & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & ")" & vbCrLf & vbCrLf
                'strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & "Set @@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & "Set @@" & dbInfo(i).COLUMNDETAILS(0).COLUMNNAME & "=(Select @@Identity)" & vbCrLf & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & "IF @@ERROR <> 0" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & vbTab & "Return -1" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & vbTab & "Return 1" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & vbTab & vbTab & "End" & vbCrLf
                strSPInsertReturnInsert = strSPInsertReturnInsert & "End"

                strSPInsert = strSPInsertCreate & strSPInsertParamListDef & strSPInsertBegin & strSPInsertParamListUpdate & strSPInsertUpdateWhere & strSPInsertReturnUpdate & strSPInsertParamListIns & strSPInsertValues & strSPInsertParamListInsValues & strSPInsertReturnInsert

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_GetDetailsFrom" & dbInfo(i).TABLENAME & ".sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPSelect)
                MyCLS.clsFileHandling.CloseFile()

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_GetDetailsFrom" & dbInfo(i).TABLENAME & "ById.sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPSelectByValue)
                MyCLS.clsFileHandling.CloseFile()

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_Delete" & dbInfo(i).TABLENAME & ".sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPDeleteByValue)
                MyCLS.clsFileHandling.CloseFile()

                MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\SP_Insert" & dbInfo(i).TABLENAME & ".sql")
                MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSPInsert)
                MyCLS.clsFileHandling.CloseFile()

                strSP = strSP & vbCrLf & strSPSelect & vbCrLf & "Go" & vbCrLf & strSPSelectByValue & vbCrLf & "Go" & vbCrLf & strSPInsert & vbCrLf & "Go" & vbCrLf & strSPDeleteByValue & vbCrLf & "Go" & vbCrLf
            Next

            MyCLS.clsFileHandling.OpenFile("C:\_CODE\SQL\_FullScript.sql")
            MyCLS.clsFileHandling.WriteFile("Use " & txtDatabase.Text & vbCrLf & "Go" & vbCrLf & strSP)
            MyCLS.clsFileHandling.CloseFile()
        Catch ex As Exception
            MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            cmdCreate.Enabled = True
            cmdCancel.Enabled = True
            lblMSG.Text = "Finished!"
        End Try
    End Sub
End Class