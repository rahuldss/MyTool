'********************************************************************************'
'               Name        :       MyCLS.vb                                     '
'               Created on  :       5th Mar, 2008                                '
'               Description :       Contains some specific fuctions to be used   '
'                                   Database Operation, Emailing, Uploading,     '
'                                   For Operations like PDF,DOC,Currency,File,   '
'                                   Multiple Queries,Windows,APIs, etc.          '
'               Created By  :       Narender Sharma (Netsoft)                    '
'               Modified On :       07-Mar-2011                                  '
'********************************************************************************'

'TO MAKE DEFAULT FOR ENTER KEY
'Page.RegisterHiddenField("__EVENTTARGET", "CmdOK")

'============================================
'**********ACTUAL CODE TO OPEN PDF FILE******
'Dim client As New System.Net.WebClient
'Dim buffer(10) As Byte
'   buffer = client.DownloadData(strFilePath)''
'
'   If buffer.ToString <> "" Then
'       Response.ContentType = "application/pdf"
'       Response.AddHeader("content-length", buffer.Length.ToString())
'       Response.BinaryWrite(buffer)
'   End If
'============================================                    

'===================================================
'**********CODE TO USE AFTER ALL FUNCTION CALL******

'   If Len(strGlobalErrorInfo) > 0 Then
'       Response.Write(strGlobalErrorInfo)
'       Exit Sub
'   End If

'==========PLEASE COPY THESE LINE====================                    


Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports Microsoft.Win32
Imports System.Xml.Serialization
Imports System.Xml
Imports Acrobat
Imports Microsoft.Office.Interop.word
Imports System.Xml.XPath
Imports System.Xml.Xsl
Imports iTextSharp.text.pdf
Imports System.Text
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class MyCLS
    Public Shared strGlobalErrorInfo As String
    Public Shared strDBErrorInfo As String
    Public Shared strGlobalErrorDB As String = "if @@error <> 0 begin select -1 end else begin select @@Identity end"

    'Shared strConnStringOLEDB As String = "UID=sa;Password=sa123;Data Source=127.0.0.1;Initial Catalog=ndhhs_rms;Provider=SQLOLEDB.1;"
    'Public Shared strConnStringOLEDB As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\Data\SkyP.mdb;Persist Security Info=True"
    Public Shared strConnStringOLEDB As String = "Data Source=TSIDEV02;Initial Catalog=AB;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"
    Public Shared strConnStringSQLCLIENT As String = "Data Source=TSIDEV02;Initial Catalog=AB;UID=sa;PWD=sa123;"

    Shared MyCon As OleDbConnection
    Shared MyConSql As SqlConnection
    Shared MyStr As String
    Shared MyCmd As OleDbCommand
    Shared MyRs As OleDbDataReader
    Shared MyDa As OleDbDataAdapter
    Shared MyDs As DataSet
    Shared MyTrans As OleDbTransaction

    'TO STORE SEARCHED ID VALUE
    Public Shared intSearchedID As Int32

    'TO STORE LOCATION VALUE
    Public Shared strLOCATION As String

    'TO Get Currently Executing Method Name
    Public Shared st As New StackTrace()


    Public Class clsCOMMON

        ''' <summary>
        ''' This is the Internal Function to be invoked from the DBConnection
        ''' The purpose of the function 'OpenDatabase' is to open the XML File and have all the values in the class
        ''' object through Serialization. This is static function of the class and can be invoked without any instance.
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function ConOpenFromXMLFile() As String 'OleDbConnection
            Dim strConnectionString As String = ""
            Dim objConnectionConfig As New ConnectionInfo()
            'Dim objConnection As OleDbConnection = Nothing
            Dim objXml As SerializeXML = Nothing
            Try
                Dim strXMLFilePath As String = ""
                strXMLFilePath = AppDomain.CurrentDomain.BaseDirectory & "startup.xml"
                objXml = New SerializeXML()
                objConnectionConfig = DirectCast(objXml.ConvertXML(strXMLFilePath, False, Nothing), ConnectionInfo)

                If Len(objConnectionConfig.UserID) > 0 And Len(objConnectionConfig.Password) > 0 Then
                    strConnectionString = ((("Data Source=" & objConnectionConfig.ServerName & ";Initial Catalog=") + objConnectionConfig.Database & ";uid=") + objConnectionConfig.UserID & ";pwd=") + objConnectionConfig.Password & ";Provider=SQLOLEDB.1;"
                Else
                    strConnectionString = ("Data Source=" & objConnectionConfig.ServerName & ";Initial Catalog=") + objConnectionConfig.Database & ";Integrated Security=SSPI;Provider=SQLOLEDB.1;"
                End If
                'If strConnectionString <> String.Empty Then
                '    objConnection = New OleDbConnection(strConnectionString)
                '    objConnection.Open()
                'End If
            Catch exSQL As SqlException
                clsHandleException.HandleEx(exSQL, System.Reflection.MethodBase.GetCurrentMethod.ToString())
                'BugsHandler.BugLogging("Exception Thrown while Opening Connection: " + WebDBErrorTypes.ConnectionError.ToString(), exSQL.Message, true);
            Catch ex As Exception
                'BugsHandler.BugLogging(ex.StackTrace, ex.Message, true);

                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            Finally
                'MyCon = objConnection
                'objConnection.Close()
            End Try

            Return strConnectionString 'objConnection
        End Function
        ''' <summary>
        ''' This is the Internal Function to be invoked from the DBConnection
        ''' The purpose of the function 'OpenDatabase' is to open the XML File and have all the values in the class
        ''' object through Serialization. This is static function of the class and can be invoked without any instance.
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function ConOpenFromXMLFile(ByVal isReturnConnectionInfo As Boolean) As ConnectionInfo
            Dim strConnectionString As String = ""
            Dim objConnectionConfig As New ConnectionInfo()
            'Dim objConnection As OleDbConnection = Nothing
            Dim objXml As SerializeXML = Nothing
            Try
                Dim strXMLFilePath As String = ""
                strXMLFilePath = AppDomain.CurrentDomain.BaseDirectory & "startup.xml"
                objXml = New SerializeXML()
                objConnectionConfig = DirectCast(objXml.ConvertXML(strXMLFilePath, False, Nothing), ConnectionInfo)

                If Len(objConnectionConfig.UserID) > 0 And Len(objConnectionConfig.Password) > 0 Then
                    strConnectionString = ((("Data Source=" & objConnectionConfig.ServerName & ";Initial Catalog=") + objConnectionConfig.Database & ";uid=") + objConnectionConfig.UserID & ";pwd=") + objConnectionConfig.Password & ";Provider=SQLOLEDB.1;"
                Else
                    strConnectionString = ("Data Source=" & objConnectionConfig.ServerName & ";Initial Catalog=") + objConnectionConfig.Database & ";Integrated Security=SSPI;Provider=SQLOLEDB.1;"
                End If
                'If strConnectionString <> String.Empty Then
                '    objConnection = New OleDbConnection(strConnectionString)
                '    objConnection.Open()
                'End If
            Catch exSQL As SqlException
                clsHandleException.HandleEx(exSQL, System.Reflection.MethodBase.GetCurrentMethod.ToString())
                'BugsHandler.BugLogging("Exception Thrown while Opening Connection: " + WebDBErrorTypes.ConnectionError.ToString(), exSQL.Message, true);
            Catch ex As Exception
                'BugsHandler.BugLogging(ex.StackTrace, ex.Message, true);

                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            Finally
                'MyCon = objConnection
                'objConnection.Close()
            End Try

            Return objConnectionConfig 'objConnection            
        End Function

        Public Shared Sub ConOpen(Optional ByVal FromXMLFile As Boolean = False)
            Try
                strGlobalErrorInfo = ""
                'Dim strconn As String = ConfigurationSettings.AppSettings("strconn1") & path & ConfigurationSettings.AppSettings("strconn3")
                'Dim Strconn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fnGetDBString()

                'LOCAL
                'Dim Strconn As String = "Integrated Security=SSPI;Data Source=Node04;Initial Catalog=ndhhs_online;Provider=SQLOLEDB.1"
                Dim StrconnOleDb As String = strConnStringOLEDB
                Dim StrconnSqlClient As String = strConnStringSQLCLIENT
                'Dim Strconn As String = "Initial Catalog=ndhhs_Updated;Data Source=TSI_DEV_02;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"
                'ONLINE
                'Dim Strconn As String = "Data Source=202.71.129.59;Initial Catalog=netsofthr;UID=netsofthr;PWD=Pa$$w0rd;Provider=SQLOLEDB.1"
                'ONLINE - NO REMOTE CONECTIVITY YET
                'Dim Strconn As String = "Data Source=p3swhsql-v21.shr.phx3.secureserver.net;Initial Catalog=tsisw;UID=tsisw;PWD=Netsoft12;Provider=SQLOLEDB.1"
                Try
                    If FromXMLFile Then
                        MyCon = New OleDbConnection(ConOpenFromXMLFile())
                        MyConSql = New SqlConnection(ConOpenFromXMLFile().Replace("Provider=SQLOLEDB.1;", ""))

                        Dim MyConnInfo As New ConnectionInfo
                        MyConnInfo = ConOpenFromXMLFile(True)
                        MyCLS.DataBaseCredentials.ServerName = MyConnInfo.ServerName
                        MyCLS.DataBaseCredentials.DatabaseName = MyConnInfo.Database
                        MyCLS.DataBaseCredentials.UserName = MyConnInfo.UserID
                        MyCLS.DataBaseCredentials.Password = MyConnInfo.Password
                    Else
                        Try
                            MyCon = New OleDbConnection(StrconnOleDb)
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                        Try
                            MyConSql = New SqlConnection(StrconnSqlClient.Replace("Provider=SQLOLEDB.1;", ""))
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If
                Catch ex As Exception

                End Try
                Try
                    If Not MyCon.State = ConnectionState.Open Then
                        MyCon.Open()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                Try
                    If Not MyConSql.State = ConnectionState.Open Then
                        MyConSql.Open()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())

                MsgBox("Database Connection Could not be open!", MsgBoxStyle.Critical, "Database Connection Error")
                MsgBox(strGlobalErrorInfo)
                'End
            End Try
        End Sub
        'Public Shared Sub ConOpen(Optional ByVal FromXMLFile As Boolean = False)
        '    Try
        '        strGlobalErrorInfo = ""
        '        'Dim strconn As String = ConfigurationSettings.AppSettings("strconn1") & path & ConfigurationSettings.AppSettings("strconn3")
        '        'Dim Strconn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fnGetDBString()

        '        'LOCAL
        '        'Dim Strconn As String = "Integrated Security=SSPI;Data Source=Node04;Initial Catalog=ndhhs_online;Provider=SQLOLEDB.1"
        '        Dim StrconnOleDb As String = strConnStringOLEDB
        '        Dim StrconnSqlClient As String = strConnStringSQLCLIENT
        '        'Dim Strconn As String = "Initial Catalog=ndhhs_Updated;Data Source=TSI_DEV_02;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"
        '        'ONLINE
        '        'Dim Strconn As String = "Data Source=202.71.129.59;Initial Catalog=netsofthr;UID=netsofthr;PWD=Pa$$w0rd;Provider=SQLOLEDB.1"
        '        'ONLINE - NO REMOTE CONECTIVITY YET
        '        'Dim Strconn As String = "Data Source=p3swhsql-v21.shr.phx3.secureserver.net;Initial Catalog=tsisw;UID=tsisw;PWD=Netsoft12;Provider=SQLOLEDB.1"
        '        Try
        '            If FromXMLFile Then
        '                MyCon = New OleDbConnection(ConOpenFromXMLFile())
        '                MyConSql = New SqlConnection(ConOpenFromXMLFile())
        '            Else
        '                MyCon = New OleDbConnection(StrconnOleDb)
        '                MyConSql = New SqlConnection(StrconnSqlClient)
        '            End If
        '        Catch ex As Exception

        '        End Try
        '        Try
        '            If Not MyCon.State = ConnectionState.Open Then
        '                MyCon.Open()
        '            End If
        '        Catch ex As Exception
        '            strGlobalErrorInfo = ex.Message
        '            MsgBox(ex.Message)
        '        End Try
        '        Try
        '            If Not MyConSql.State = ConnectionState.Open Then
        '                MyConSql.Open()
        '            End If
        '        Catch ex As Exception

        '        End Try
        '    Catch ex As Exception
        '        clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())

        '        MsgBox("Database Connection Could not be open!", MsgBoxStyle.Critical, "Database Connection Error")
        '        MsgBox(strGlobalErrorInfo)
        '        'End
        '    End Try
        'End Sub
        Public Shared Sub GetCon(ByRef NewCon As OleDbConnection)
            Try
                If MyCon Is Nothing Then
                    ConOpen(True)
                End If
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                NewCon = MyCon
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub GetCon(ByRef NewCon As SqlConnection)
            Try
                If MyConSql Is Nothing Then
                    ConOpen(True)
                End If
                If Not MyConSql.State = ConnectionState.Open Then
                    MyConSql.Open()
                End If
                NewCon = MyConSql
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub SetCon(ByRef NewCon As OleDbConnection)
            Try
                MyCon = NewCon
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub SetCon(ByRef NewCon As SqlConnection)
            Try
                MyConSql = NewCon
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub ConClose()
            Try
                strGlobalErrorInfo = ""
                'Dim strconn As String = ConfigurationSettings.AppSettings("strconn1") & path & ConfigurationSettings.AppSettings("strconn3")
                If MyCon.State Then
                    MyCon.Close()
                End If
                If MyConSql.State Then
                    MyConSql.Close()
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Function fnIsDuplicate(ByVal TblName As String, ByVal WhereCol1 As String, ByVal Col1Value As String, Optional ByVal strLogicalOperator As String = "AND", Optional ByRef Col1Type As String = "String", Optional ByVal WhereCol2 As String = "", Optional ByVal Col2Value As String = "", Optional ByVal Col2Type As String = "String") As Boolean
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim QChr1 As Char
                Dim QChr2 As Char

                If Col1Type = "String" Then
                    QChr1 = "'"
                ElseIf Col1Type = "Date" Then
                    QChr1 = "#"
                Else
                    QChr1 = " "
                End If
                If Col2Type = "String" Then
                    QChr2 = "'"
                ElseIf Col2Type = "Date" Then
                    QChr2 = "#"
                Else
                    QChr2 = " "
                End If

                If WhereCol2.Length > 0 Then
                    'If Col2Value.Length > 0 Then
                    MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2
                    'End If
                Else
                    MyStr = "Select * From " & TblName & " Where " & WhereCol1 & " = " & QChr1 & Col1Value & QChr1
                End If
                MyCmd = New OleDbCommand(MyStr, MyCon)
                MyRs = MyCmd.ExecuteReader
                MyRs.Read()
                If MyRs.HasRows Then
                    fnIsDuplicate = True
                Else
                    fnIsDuplicate = False
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Function fnIsDuplicateW(ByVal TblName As String, ByVal WhereCol1 As String, ByVal Col1Value As String, Optional ByVal strLogicalOperator As String = "AND", Optional ByRef Col1Type As String = "String", Optional ByVal WhereCol2 As String = "", Optional ByVal Col2Value As String = "", Optional ByVal Col2Type As String = "String", Optional ByVal WhereCol3 As String = "", Optional ByVal Col3Value As String = "", Optional ByVal Col3Type As String = "String", Optional ByVal WhereCol4 As String = "", Optional ByVal Col4Value As String = "", Optional ByVal Col4Type As String = "String", Optional ByVal WhereCol5 As String = "", Optional ByVal Col5Value As String = "", Optional ByVal Col5Type As String = "String", Optional ByVal WhereCol6 As String = "", Optional ByVal Col6Value As String = "", Optional ByVal Col6Type As String = "String", Optional ByVal WhereCol7 As String = "", Optional ByVal Col7Value As String = "", Optional ByVal Col7Type As String = "String", Optional ByVal WhereCol8 As String = "", Optional ByVal Col8Value As String = "", Optional ByVal Col8Type As String = "String", Optional ByVal WhereCol9 As String = "", Optional ByVal Col9Value As String = "", Optional ByVal Col9Type As String = "String", Optional ByVal WhereCol0 As String = "", Optional ByVal Col0Value As String = "", Optional ByVal Col0Type As String = "String") As Boolean
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim QChr1 As Char
                Dim QChr2 As Char
                Dim QChr3 As Char
                Dim QChr4 As Char
                Dim QChr5 As Char
                Dim QChr6 As Char
                Dim QChr7 As Char
                Dim QChr8 As Char
                Dim QChr9 As Char
                Dim QChr0 As Char

                If Col1Type = "String" Then
                    QChr1 = "'"
                ElseIf Col1Type = "Date" Then
                    QChr1 = "#"
                Else
                    QChr1 = " "
                End If
                If Col2Type = "String" Then
                    QChr2 = "'"
                ElseIf Col2Type = "Date" Then
                    QChr2 = "#"
                Else
                    QChr2 = " "
                End If
                'If Col1Type = "String" Then
                '    QChr1 = "'"
                'ElseIf Col1Type = "Date" Then
                '    QChr1 = "#"
                'Else
                '    QChr1 = " "
                'End If
                If Col3Type = "String" Then
                    QChr3 = "'"
                ElseIf Col3Type = "Date" Then
                    QChr3 = "#"
                Else
                    QChr3 = " "
                End If
                If Col4Type = "String" Then
                    QChr4 = "'"
                ElseIf Col4Type = "Date" Then
                    QChr4 = "#"
                Else
                    QChr4 = " "
                End If
                If Col5Type = "String" Then
                    QChr5 = "'"
                ElseIf Col5Type = "Date" Then
                    QChr5 = "#"
                Else
                    QChr5 = " "
                End If
                If Col6Type = "String" Then
                    QChr6 = "'"
                ElseIf Col6Type = "Date" Then
                    QChr6 = "#"
                Else
                    QChr6 = " "
                End If
                If Col7Type = "String" Then
                    QChr7 = "'"
                ElseIf Col7Type = "Date" Then
                    QChr7 = "#"
                Else
                    QChr7 = " "
                End If
                If Col8Type = "String" Then
                    QChr8 = "'"
                ElseIf Col8Type = "Date" Then
                    QChr8 = "#"
                Else
                    QChr8 = " "
                End If
                If Col9Type = "String" Then
                    QChr9 = "'"
                ElseIf Col9Type = "Date" Then
                    QChr9 = "#"
                Else
                    QChr9 = " "
                End If
                If Col0Type = "String" Then
                    QChr0 = "'"
                ElseIf Col0Type = "Date" Then
                    QChr0 = "#"
                Else
                    QChr0 = " "
                End If
                If WhereCol1.Length > 0 Then
                    MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1
                    If WhereCol2.Length > 0 Then
                        MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2
                        If WhereCol3.Length > 0 Then
                            MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2 & " " & strLogicalOperator & " " & WhereCol3 & "=" & QChr3 & Col3Value & QChr3
                            If WhereCol4.Length > 0 Then
                                MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2 & " " & strLogicalOperator & " " & WhereCol3 & "=" & QChr3 & Col3Value & QChr3 & " " & strLogicalOperator & " " & WhereCol4 & "=" & QChr4 & Col4Value & QChr4
                                If WhereCol5.Length > 0 Then
                                    MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2 & " " & strLogicalOperator & " " & WhereCol3 & "=" & QChr3 & Col3Value & QChr3 & " " & strLogicalOperator & " " & WhereCol4 & "=" & QChr4 & Col4Value & QChr4 & " " & strLogicalOperator & " " & WhereCol5 & "=" & QChr5 & Col5Value & QChr5
                                    If WhereCol6.Length > 0 Then
                                        MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2 & " " & strLogicalOperator & " " & WhereCol3 & "=" & QChr3 & Col3Value & QChr3 & " " & strLogicalOperator & " " & WhereCol4 & "=" & QChr4 & Col4Value & QChr4 & " " & strLogicalOperator & " " & WhereCol5 & "=" & QChr5 & Col5Value & QChr5 & " " & strLogicalOperator & " " & WhereCol6 & "=" & QChr6 & Col6Value & QChr6
                                        If WhereCol7.Length > 0 Then
                                            MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2 & " " & strLogicalOperator & " " & WhereCol3 & "=" & QChr3 & Col3Value & QChr3 & " " & strLogicalOperator & " " & WhereCol4 & "=" & QChr4 & Col4Value & QChr4 & " " & strLogicalOperator & " " & WhereCol5 & "=" & QChr5 & Col5Value & QChr5 & " " & strLogicalOperator & " " & WhereCol6 & "=" & QChr6 & Col6Value & QChr6 & " " & strLogicalOperator & " " & WhereCol7 & "=" & QChr7 & Col7Value & QChr7
                                            If WhereCol8.Length > 0 Then
                                                MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2 & " " & strLogicalOperator & " " & WhereCol3 & "=" & QChr3 & Col3Value & QChr3 & " " & strLogicalOperator & " " & WhereCol4 & "=" & QChr4 & Col4Value & QChr4 & " " & strLogicalOperator & " " & WhereCol5 & "=" & QChr5 & Col5Value & QChr5 & " " & strLogicalOperator & " " & WhereCol6 & "=" & QChr6 & Col6Value & QChr6 & " " & strLogicalOperator & " " & WhereCol7 & "=" & QChr7 & Col7Value & QChr7 & " " & strLogicalOperator & " " & WhereCol8 & "=" & QChr8 & Col8Value & QChr8
                                                If WhereCol9.Length > 0 Then
                                                    MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2 & " " & strLogicalOperator & " " & WhereCol3 & "=" & QChr3 & Col3Value & QChr3 & " " & strLogicalOperator & " " & WhereCol4 & "=" & QChr4 & Col4Value & QChr4 & " " & strLogicalOperator & " " & WhereCol5 & "=" & QChr5 & Col5Value & QChr5 & " " & strLogicalOperator & " " & WhereCol6 & "=" & QChr6 & Col6Value & QChr6 & " " & strLogicalOperator & " " & WhereCol7 & "=" & QChr7 & Col7Value & QChr7 & " " & strLogicalOperator & " " & WhereCol8 & "=" & QChr8 & Col8Value & QChr8 & " " & strLogicalOperator & " " & WhereCol9 & "=" & QChr9 & Col9Value & QChr9
                                                    If WhereCol0.Length > 0 Then
                                                        MyStr = "Select * From " & TblName & " Where " & WhereCol1 & "=" & QChr1 & Col1Value & QChr1 & " " & strLogicalOperator & " " & WhereCol2 & "=" & QChr2 & Col2Value & QChr2 & " " & strLogicalOperator & " " & WhereCol3 & "=" & QChr3 & Col3Value & QChr3 & " " & strLogicalOperator & " " & WhereCol4 & "=" & QChr4 & Col4Value & QChr4 & " " & strLogicalOperator & " " & WhereCol5 & "=" & QChr5 & Col5Value & QChr5 & " " & strLogicalOperator & " " & WhereCol6 & "=" & QChr6 & Col6Value & QChr6 & " " & strLogicalOperator & " " & WhereCol7 & "=" & QChr7 & Col7Value & QChr7 & " " & strLogicalOperator & " " & WhereCol8 & "=" & QChr8 & Col8Value & QChr8 & " " & strLogicalOperator & " " & WhereCol9 & "=" & QChr9 & Col9Value & QChr9 & " " & strLogicalOperator & " " & WhereCol0 & "=" & QChr0 & Col0Value & QChr0
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                'msgbox(MyStr)
                'Exit Function

                MyCmd = New OleDbCommand(MyStr, MyCon)
                MyRs = MyCmd.ExecuteReader
                MyRs.Read()
                If MyRs.HasRows Then
                    fnIsDuplicateW = True
                Else
                    fnIsDuplicateW = False
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnValidateEmail(ByVal Obj As Object, ByRef LblObj2ShowError As Object) As Boolean
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                If Obj.Text = "" Then
                    'LblObj2ShowError = "Enter Value for E-MAIL!"
                    'Obj.Focus()
                    fnValidateEmail = True
                    'Exit Function
                ElseIf InStr(1, Obj.Text, "@") = 0 Then
                    LblObj2ShowError = "Enter Valid E-MAIL!"
                    Obj.Focus()
                    fnValidateEmail = False
                    Exit Function
                ElseIf InStr(1, Obj.Text, ".") = 0 Then
                    LblObj2ShowError = "Enter Valid E-MAIL!"
                    Obj.Focus()
                    fnValidateEmail = False
                    Exit Function
                ElseIf InStr(1, Obj.Text, "@") = 1 Then
                    LblObj2ShowError = "Enter Valid E-MAIL!"
                    Obj.Focus()
                    fnValidateEmail = False
                    Exit Function
                ElseIf InStr(1, Obj.Text, ".") = Len(Obj.Text) Or InStr(1, Obj.Text, ".") + 1 = Len(Obj.Text) Then
                    LblObj2ShowError = "Enter Valid E-MAIL!"
                    Obj.Focus()
                    fnValidateEmail = False
                    Exit Function
                ElseIf InStr(1, Obj.Text, "@") + 1 = InStr(1, Obj.Text, ".") Or InStr(1, Obj.Text, ".") + 1 = InStr(1, Obj.Text, "@") Then
                    LblObj2ShowError = "Enter Valid E-MAIL!"
                    Obj.Focus()
                    fnValidateEmail = False
                    Exit Function
                ElseIf IsNumeric(Mid(Obj.Text, 1, 1)) Then
                    LblObj2ShowError = "Enter Valid E-MAIL!"
                    Obj.Focus()
                    fnValidateEmail = False
                    Exit Function
                End If
                LblObj2ShowError = ""
                fnValidateEmail = True
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnGetIncSNO(ByVal TblName As String, ByVal ColName As String) As Long
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                'FOR MSSQL
                'MyStr = "Select isNull(Max(SNO)+1,1) From AddBook"
                'FOR MSACCESS
                'MyStr = "SELECT iif(isnull(max(sno)),1,max(sno)+1) FROM Book"

                MyStr = "SELECT ISNULL(MAX(" & ColName & ")+1, 1) FROM " & TblName
                MyCmd = New OleDbCommand(MyStr, MyCon)
                MyRs = MyCmd.ExecuteReader
                MyRs.Read()

                fnGetIncSNO = MyRs.Item(0)
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnQuerySelect1Value(ByVal SelectQ As String, ByVal ReturnType As String) As String
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                MyCmd = New OleDbCommand(SelectQ, MyCon)
                MyRs = MyCmd.ExecuteReader
                MyRs.Read()
                If MyRs.HasRows Then
                    fnQuerySelect1Value = MyRs(0).ToString
                Else
                    If ReturnType = "String" Then
                        fnQuerySelect1Value = ""
                    Else
                        fnQuerySelect1Value = 0
                    End If
                End If
            Catch ex As Exception
                'strGlobalErrorInfo = "Qurery is : " & SelectQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Function fnQuerySelect1Value(ByVal TblName As String, ByVal ColName As String, ByVal ReturnType As String) As String
            Dim SelectQ As String
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                SelectQ = "Select " & ColName & " from " & TblName
                MyCmd = New OleDbCommand(SelectQ, MyCon)
                MyRs = MyCmd.ExecuteReader
                MyRs.Read()
                If MyRs.HasRows Then
                    fnQuerySelect1Value = MyRs(ColName).ToString
                Else
                    If ReturnType = "String" Then
                        fnQuerySelect1Value = ""
                    Else
                        fnQuerySelect1Value = 0
                    End If
                End If
            Catch ex As Exception
                'strGlobalErrorInfo = "Qurery is : " & SelectQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Function fnQuerySelect1Value(ByVal TblName As String, ByVal ColName As String, ByVal WhereCol As String, ByVal ColValue As String, ByVal ReturnType As String) As String
            Dim SelectQ As String
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                SelectQ = "Select " & ColName & " from " & TblName & " Where " & WhereCol & "='" & ColValue & "'"
                MyCmd = New OleDbCommand(SelectQ, MyCon)
                MyRs = MyCmd.ExecuteReader
                MyRs.Read()
                If MyRs.HasRows Then
                    fnQuerySelect1Value = MyRs(ColName).ToString
                Else
                    If ReturnType = "String" Then
                        fnQuerySelect1Value = ""
                    Else
                        fnQuerySelect1Value = 0
                    End If
                End If
            Catch ex As Exception
                'strGlobalErrorInfo = "Qurery is : " & SelectQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnQuerySelectRS(ByVal SelectQ As String) As OleDbDataReader
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""

                MyCmd = New OleDbCommand(SelectQ, MyCon)
                MyRs = MyCmd.ExecuteReader
                'If MyRs.HasRows = True Then
                fnQuerySelectRS = MyRs
                'End If
            Catch ex As Exception
                strGlobalErrorInfo = "Query is : " & SelectQ
                strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.Data)
                '            Dim a As New ErrorLog
                '           a.ErrorLogS(strGlobalErrorInfo)
                fnWrite2LOG(strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
        End Function
        Public Shared Function fnQuerySelectDA(ByVal SelectQ As String) As OleDbDataAdapter
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                Else
                    MyCon.ResetState()
                End If
                strGlobalErrorInfo = ""

                'MyCmd = New OleDbCommand(SelectQ, MyCon)
                MyDa = New OleDbDataAdapter(SelectQ, MyCon)
                'MyDa = MyCmd.ExecuteReader            
                fnQuerySelectDA = MyDa
            Catch ex As Exception
                'strGlobalErrorInfo = "Qurery is : " & SelectQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Sub prcQuerySelectDS(ByRef MyDataset As DataSet, ByVal SelectQ As String, ByVal TableName As String)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""

                MyDa = New OleDbDataAdapter(SelectQ, MyCon)
                MyDa.SelectCommand.CommandTimeout = 10000
                MyDa.Fill(MyDataset, TableName)
            Catch ex As Exception
                'strGlobalErrorInfo = "Qurery is : " & SelectQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Function fnQueryExecuter(ByVal InsertQ As String, ByVal Trans As OleDbTransaction) As Long
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                strDBErrorInfo = ""

                MyCmd = New OleDbCommand(InsertQ, MyCon, Trans)
                MyRs = MyCmd.ExecuteReader
                fnQueryExecuter = MyRs.RecordsAffected
                Return 1
            Catch ex As Exception
                strDBErrorInfo = ex.Message
                'strGlobalErrorInfo = "Qurery is : " & InsertQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
                Return 0
            End Try
        End Function
        Public Shared Function fnQueryExecuter(ByVal InsertQ As String) As Long
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                strDBErrorInfo = ""

                MyTrans = MyCon.BeginTransaction(IsolationLevel.ReadUncommitted)

                MyCmd = New OleDbCommand(InsertQ, MyCon, MyTrans)
                MyRs = MyCmd.ExecuteReader
                fnQueryExecuter = MyRs.RecordsAffected
                MyTrans.Commit()
            Catch ex As Exception
                MyTrans.Rollback()
                strDBErrorInfo = ex.Message
                'strGlobalErrorInfo = "Qurery is : " & InsertQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnQueryInsert(ByVal InsertQ As String) As Long
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                strDBErrorInfo = ""

                MyCmd = New OleDbCommand(InsertQ, MyCon)
                MyRs = MyCmd.ExecuteReader
                fnQueryInsert = MyRs.RecordsAffected
            Catch ex As Exception
                strDBErrorInfo = ex.Message
                'strGlobalErrorInfo = "Qurery is : " & InsertQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnQueryUpdate(ByVal UpdateQ As String) As Long
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                strDBErrorInfo = ""

                MyCmd = New OleDbCommand(UpdateQ, MyCon)
                MyRs = MyCmd.ExecuteReader
                fnQueryUpdate = MyRs.RecordsAffected
            Catch ex As Exception
                strDBErrorInfo = ex.Message
                'strGlobalErrorInfo = "Qurery is : " & UpdateQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnQueryDelete(ByVal DeleteQ As String) As Long
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""

                MyCmd = New OleDbCommand(DeleteQ, MyCon)
                MyRs = MyCmd.ExecuteReader
                fnQueryDelete = MyRs.RecordsAffected
            Catch ex As Exception
                'strGlobalErrorInfo = "Qurery is : " & DeleteQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnQuoteRemove(ByVal StrV As Object) As String
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                fnQuoteRemove = Replace(StrV, "'", "+ 39 +")
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Function fnQuoteRetrieve(ByVal StrV As Object) As String
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                fnQuoteRetrieve = Replace(StrV, "+ 39 +", "'")
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Function fnQuoteConvert4Query(ByVal StrV As Object) As String
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                fnQuoteConvert4Query = Replace(StrV, "'", "''")
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnGeneratePassword() As String
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim rndValue As New Random
                fnGeneratePassword = Chr(rndValue.Next(65, 90)) & Chr(rndValue.Next(65, 90)) & Chr(rndValue.Next(97, 122)) & Chr(rndValue.Next(97, 122)) & Chr(rndValue.Next(65, 90)) & Chr(rndValue.Next(97, 122)) & Chr(rndValue.Next(97, 122)) & Chr(rndValue.Next(97, 122))
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnCountRows(ByVal StrQ As String) As Long
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim TRows As Long

                MyCmd = New OleDbCommand(StrQ, MyCon)
                MyRs = MyCmd.ExecuteReader()
                TRows = 0

                While MyRs.Read
                    TRows = TRows + 1
                End While
                fnCountRows = TRows
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnGetExtension(ByVal FileName As String) As String
            Try
                strGlobalErrorInfo = ""
                Dim intLoc As Integer
                Dim strExt As String = ""
                Dim strExtRet As String = ""

                'strExt = FileName
                'If InStr(strExt, ".") > 0 Then
                '    intLoc = InStr(strExt, ".") + 1
                '    strExt = Mid(strExt, intLoc)
                '    fnGetExtension(strExt)
                'Else
                '    Exit Function
                'End If
                'fnGetExtension = strExt

                FileName = Strings.StrReverse(FileName)
                FileName = Mid(FileName, 1, InStr(FileName, ".") - 1)
                strExt = Strings.StrReverse(FileName)
                fnGetExtension = strExt
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnGetFileName(ByVal strFileName As String) As String
            Try
                strGlobalErrorInfo = ""
                'strFileName = FUCLetter.PostedFile.FileName
                'MsgBox(strFileName)
                strFileName = Strings.StrReverse(strFileName)
                'MsgBox(strFileName)

                'MsgBox(InStr(strFileName, "\"))
                strFileName = Mid(strFileName, 1, InStr(strFileName, "\") - 1)
                strFileName = Strings.StrReverse(strFileName)
                'MsgBox(strFileName)
            Catch ex As Exception
                strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.Data)
            End Try
            Return strFileName
        End Function

        Public Shared Function fnShowException(ByVal ex As Exception) As String
            strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
            strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
            strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.Data)

            fnShowException = strGlobalErrorInfo
        End Function

        Public Shared Function fnConvertedPathString(ByVal strOldPath As String) As String
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                Dim str12s() As String = strOldPath.Split("\")
                Dim i As Integer
                Dim strNewPath As String
                For i = 0 To str12s.Length - 1
                    '   'msgbox(str12s(i))
                    If str12s(i).Contains(" ") = True Then
                        str12s(i).Replace(" ", "")
                        If (str12s(i).Length > 6) Then
                            str12s(i) = str12s(i).Substring(0, 6) & "~1"
                            '            'msgbox(str12s(i))
                        Else
                            str12s(i) = str12s(i) & "~1"
                            'msgbox(str12s(i))
                        End If
                        strNewPath = strNewPath & str12s(i) & "\"
                    Else
                        If (str12s(i).Length > 8) Then
                            str12s(i) = str12s(i).Substring(0, 6) & "~1"
                            'msgbox(str12s(i))
                        End If
                        strNewPath = strNewPath & str12s(i) & "\"
                    End If
                Next
                fnConvertedPathString = strNewPath
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Sub RecursiveCopyFiles(ByVal sourceDir As String, ByVal destDir As String, ByVal fRecursive As Boolean)
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
                        RecursiveCopyFiles(aDirs(i), (destDir + sDir), fRecursive)
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
                    ' Copy the file.            
                    System.IO.File.Copy(aFiles(i), destDir + sFile)

                    'lstExtractFiles.Items.Add("Extract :" & sFile & "......")
                    'lstExtractFiles.SelectedIndex = lstExtractFiles.Items.Count - 1
                    'lstExtractFiles.EndUpdate()
                    'System.Windows.Forms.Application.DoEvents()
                    'PBExtract.Value = i
                Next i
                'PBExtract.Value = 0
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Function fnDeleteVirtualDirectory(ByVal VirtualDirName As String) As String
            Try
                Dim IIsWebVDirRootObj As Object
                Dim IIsWebVDirObj As Object
                ' Create an instance of the virtual directory object 
                ' that represents the default Web site.
                IIsWebVDirRootObj = GetObject("IIS://localhost/W3SVC/1/Root")
                ' Use the Windows ADSI container object "Create" method to create 
                ' a new virtual directory.
                Try
                    IIsWebVDirObj = IIsWebVDirRootObj.Delete("IIsWebVirtualDir", VirtualDirName)
                    fnDeleteVirtualDirectory = "Deleted!"
                Catch ex As Exception
                    clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
                End Try
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Sub prcDeleteFolder(ByVal Path As String)
            Try
                Directory.Delete(Path, True)
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Sub ModifyWebConfig(ByVal VirtualDirPath As String, ByVal strIPAddress As String, ByVal strSqlUserName As String, ByVal strSqlPassword As String)
            Try
                Dim FILENAME As String = VirtualDirPath & "\web1.config"
                Dim filename1 As String = VirtualDirPath & "\web.config"
                Dim objStreamReader As StreamReader
                objStreamReader = File.OpenText(FILENAME)
                Dim objStreamWriter As StreamWriter

                Dim contents As String

                objStreamWriter = File.CreateText(filename1)
                While objStreamReader.Peek() > -1
                    contents = objStreamReader.ReadLine()
                    contents = Replace(contents, "IPADD", strIPAddress)
                    contents = Replace(contents, "USERNAME", strSqlUserName)
                    contents = Replace(contents, "PASSWORD", strSqlPassword)
                    objStreamWriter.WriteLine(contents)
                End While
                objStreamReader.Close()
                objStreamWriter.Close()
                'File.Delete(VirtualDirPath & "\PublishedData\web1.config")
                '========================================================
                'msgbox(strWebConfig)
                'msgbox(FileSystem.ReadAllText(strWebConfig))
                'msgbox(strIPAddress)
                'msgbox(strSqlUserName)
                'msgbox(strSqlPassword)

                'FileSystem.ReadAllText(strWebConfig).Replace("IPADD", strIPAddress)
                'FileSystem.ReadAllText(strWebConfig).Replace("USERNAME", strSqlUserName)
                'FileSystem.ReadAllText(strWebConfig).Replace("PASSWORD", strSqlPassword)

                'msgbox(FileSystem.ReadAllText(strWebConfig))
                '========================================================

            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Private Function fnReplaceSpecialChars(ByVal strHEAD As String) As String
            On Error Resume Next
            'SPECIAL CHARS CAN'T BE USED IN FILE NAMES
            '\ / : * ? " < > | 
            strHEAD = Replace(strHEAD, "\", "")
            strHEAD = Replace(strHEAD, "/", "")
            strHEAD = Replace(strHEAD, ":", "")
            strHEAD = Replace(strHEAD, "*", "")
            strHEAD = Replace(strHEAD, "?", "")
            strHEAD = Replace(strHEAD, """", "")
            strHEAD = Replace(strHEAD, "<", "")
            strHEAD = Replace(strHEAD, ">", "")
            strHEAD = Replace(strHEAD, "|", "")
            fnReplaceSpecialChars = strHEAD
        End Function






        '******TryING TO ASSIGN MULTIPLE VALUES TO A PARAMETER IN FUNCTION DEFINITION******
        ''Dim PArray() As String
        'Private Shared Sub FillArray(ByRef PArray())
        '    PArray(0) = "a"
        '    PArray(1) = "b"
        '    PArray(2) = "c"
        'End Sub
        'Public Shared Sub Abc(ByVal PArray() As String)
        '    Call FillArray(PArray)
        '    Dim i
        '    For i = LBound(PArray) To UBound(PArray)
        '        'msgbox(PArray(i))
        '    Next
        'End Sub
        'Public Sub AbcCall()
        '    Abc("ffgdf")
        'End Sub
        Public Shared Function fnGetLastSNO(ByVal TblName As String, ByVal ColName As String) As Long
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""

                'FOR MSSQL
                'MyStr = "Select isNull(Max(SNO),1) From AddBook"
                'FOR MSACCESS
                'MyStr = "SELECT iif(isnull(max(sno)),1,max(sno)+1) FROM Book"
                MyStr = "SELECT MAX(" & ColName & ") FROM " & TblName
                MyCmd = New OleDbCommand(MyStr, MyCon)
                MyRs = MyCmd.ExecuteReader
                MyRs.Read()

                fnGetLastSNO = MyRs.Item(0)
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function fnQueryTruncate(ByVal TruncateTable As String) As Long
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim trCmd As String = "Truncate table " & TruncateTable
                MyCmd = New OleDbCommand(trCmd, MyCon)
                MyRs = MyCmd.ExecuteReader
                fnQueryTruncate = MyRs.RecordsAffected
            Catch ex As Exception
                strGlobalErrorInfo = "Query is : " & TruncateTable
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Function Encrypt(ByVal Value As String) As String
            Try
                If Len(Value) > 0 Then
                    Dim i As Integer
                    Dim NValue As String
                    For i = 1 To Value.Length
                        NValue = NValue & Chr(Asc(Mid(Value, i, 1)) + 1)
                    Next
                    Encrypt = NValue
                End If
            Catch ex As Exception
                clsCOMMON.fnWrite2LOG(strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
        End Function

        Public Shared Function Decrypt(ByVal Value As String) As String
            Try
                If Len(Value) > 0 Then
                    Dim i As Integer
                    Dim NValue As String
                    For i = 1 To Value.Length
                        NValue = NValue & Chr(Asc(Mid(Value, i, 1)) - 1)
                    Next
                    Decrypt = NValue
                End If
            Catch ex As Exception
                clsCOMMON.fnWrite2LOG(strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
        End Function

        Public Shared Function fnGetDBString() As String
            Try
                Dim oFile As System.IO.File
                Dim oRead As System.IO.StreamReader

                oRead = oFile.OpenText(My.Application.Info.DirectoryPath() & "\db.dat")
                fnGetDBString = Decrypt(oRead.ReadLine())


                oFile = Nothing
            Catch ex As Exception
                clsCOMMON.fnWrite2LOG(strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Database error")
            End Try
        End Function


        Public Shared Sub fnWrite2LOG(ByVal ErrMSG As String, ByVal MethodName As String)
            Try
                Dim oFile As System.IO.File
                Dim oWrite As System.IO.StreamWriter

                oWrite = oFile.AppendText(My.Application.Info.DirectoryPath() & "\Err.dat")

                oWrite.WriteLine(vbCrLf & "***" & Format(Now(), "MM/dd/yyyy hh:mm:ss tt") & "********************" & vbCrLf & MethodName & vbCrLf & ErrMSG)

                oWrite.Close()
                oFile = Nothing
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Information)
            End Try
        End Sub

        Public Shared Sub fnWrite2LOG4Import(ByVal strMSG As String)
            Try
                Dim oFile As System.IO.File
                Dim oWrite As System.IO.StreamWriter

                oWrite = oFile.AppendText(My.Application.Info.DirectoryPath() & "\Import_" & Format(Now(), "yyyyMMdd") & ".log")

                oWrite.WriteLine(strMSG)

                oWrite.Close()
                oFile = Nothing
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Information)
            End Try
        End Sub

        Public Shared Sub prcAlignObjects(ByRef Frm As Form, ByRef Ctrl As Object)
            'Dim Ctrl As Control
            'For Each Ctrl In Frm.Controls
            'If TypeOf Ctrl Is GroupBox Then        
            Ctrl.Left = (Frm.Width - Ctrl.Width) / 2
            Ctrl.Top = (Frm.Height - Ctrl.Height) / 2
            'End If
            'Next
        End Sub

        Public Shared Sub prcMaximizeObjects(ByRef Frm As Form, ByRef Ctrl As Object)
            Ctrl.Top = 5
            Ctrl.Left = 5
            Ctrl.Width = Frm.Width - 10
            Ctrl.Height = Frm.Height - 10
        End Sub

        Public Shared Sub prcInitialize(ByRef Obj As Object)
            Dim Ctrl As Control
            For Each Ctrl In Obj.Controls
                If TypeOf Ctrl Is GroupBox Or TypeOf Ctrl Is TabControl Then
                    prcInitialize(Ctrl)
                Else
                    If UCase(Ctrl.Name) <> "CMDADD" And UCase(Ctrl.Name) <> "CMDCLOSE" And UCase(Ctrl.Name) <> "CMDFIND" Then
                        Ctrl.Enabled = False
                    Else
                        Ctrl.Enabled = True
                    End If
                End If
            Next

            'Me.cmdAdd.Parent.Parent.Enabled = True
            'Me.cmdAdd.Parent.Enabled = True
            'Me.cmdAdd.Enabled = True
            'Me.cmdFind.Enabled = True
            'Me.CmdClose.Enabled = True
        End Sub
        Public Shared Sub prcInitAdd(ByRef Obj As Object)
            Dim Ctrl As Control
            For Each Ctrl In Obj.Controls
                If TypeOf Ctrl Is GroupBox Or TypeOf Ctrl Is TabControl Then
                    prcInitAdd(Ctrl)
                Else
                    If UCase(Ctrl.Name) <> "CMDMODIFY" And UCase(Ctrl.Name) <> "CMDDELETE" Then
                        Ctrl.Enabled = True
                    Else
                        Ctrl.Enabled = False
                    End If
                End If
            Next
        End Sub
        Public Shared Sub prcInitFind(ByRef Obj As Object)
            Dim Ctrl As Control
            For Each Ctrl In Obj.Controls
                If TypeOf Ctrl Is GroupBox Or TypeOf Ctrl Is TabControl Then
                    prcInitFind(Ctrl)
                Else
                    If UCase(Ctrl.Name) = "CMDMODIFY" Or UCase(Ctrl.Name) = "CMDDELETE" Then
                        Ctrl.Enabled = True
                    End If
                    If UCase(Ctrl.Name) = "CMDSAVE" Then
                        Ctrl.Enabled = False
                    End If
                End If
            Next
        End Sub
        Public Shared Sub prcClear(ByRef Obj As Object, Optional ByVal Obj2Leave As String = "")
            Dim Ctrl As Control
            For Each Ctrl In Obj.Controls
                'If InStr(Ctrl.Name, "txt") Then
                'MsgBox(Ctrl.Name)
                'End If
                If TypeOf Ctrl Is GroupBox Then
                    prcClear(Ctrl, Obj2Leave)
                ElseIf TypeOf Ctrl Is TabControl Then
                    prcClear(Ctrl, Obj2Leave)
                ElseIf TypeOf Ctrl Is TabPage Then
                    prcClear(Ctrl, Obj2Leave)
                ElseIf Ctrl.Name = Obj2Leave Then
                    Continue For
                ElseIf TypeOf Ctrl Is TextBox Then
                    Ctrl.Text = ""
                ElseIf TypeOf Ctrl Is ComboBox Then
                    Ctrl.Text = ""
                ElseIf TypeOf Ctrl Is System.Windows.Forms.CheckBox Then
                    CType(Ctrl, System.Windows.Forms.CheckBox).Checked = False
                ElseIf TypeOf Ctrl Is DataGridView Then
                    CType(Ctrl, DataGridView).DataSource = Nothing
                End If
            Next
        End Sub

        Public Shared Function fnValidateDATA(ByRef txt As TextBox, ByVal strPattern As String, ByVal Type As String) As Boolean
            'Dim pattern As New Regex(Validation)
            Dim pattern As New Regex(strPattern)
            Dim patternMatch As Match = pattern.Match(txt.Text)

            If Len(txt.Text) = 0 Then
                Return True
            End If
            If Not patternMatch.Success Then
                If UCase(Type) = "FAX" Then
                    MsgBox("Invalid Fax Number!")
                ElseIf UCase(Type) = "EMAIL" Then
                    MsgBox("Invalid Email ID!")
                ElseIf UCase(Type) = "ZIP" Then
                    MsgBox("Invalid Zipcode!")
                ElseIf UCase(Type) = "PHONE" Then
                    MsgBox("Invalid Phone Number!")
                End If
                txt.Focus()
                Return False
            Else
                Return True
            End If
        End Function

        Public Shared Sub prcUpdateWithOleDbDataAdapter(ByRef ADP As OleDbDataAdapter, ByRef DS As DataSet, ByVal strAllQueries() As String, ByVal TableName As String)
            For i As Int16 = 1 To strAllQueries.GetUpperBound(0)
                MsgBox(strAllQueries(i))
                ADP.InsertCommand = New OleDb.OleDbCommand
                ADP.InsertCommand.CommandText = strAllQueries(i)
                ADP.InsertCommand.Connection = MyCon
                ADP.Update(DS, TableName)
                DS.AcceptChanges()
            Next
        End Sub


        Public Shared Function fnRemoveHTML(ByVal strText)
            Dim TAGLIST
            TAGLIST = ";!--;!DOCTYPE;A;ACRONYM;ADDRESS;APPLET;AREA;B;BASE;BASEFONT;" & _
                      "BGSOUND;BIG;BLOCKQUOTE;BODY;BR;BUTTON;CAPTION;CENTER;CITE;CODE;" & _
                      "COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EM;EMBED;FIELDSET;" & _
                      "FONT;FORM;FRAME;FRAMESET;HEAD;H1;H2;H3;H4;H5;H6;HR;HTML;I;IFRAME;IMG;" & _
                      "INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LI;LINK;LISTING;MAP;MARQUEE;" & _
                      "MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OL;OPTION;P;PARAM;PLAINTEXT;" & _
                      "PRE;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STRONG;STYLE;SUB;SUP;" & _
                      "TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;U;UL;VAR;WBR;XMP;"

            Const BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"

            Dim nPos1
            Dim nPos2
            Dim nPos3
            Dim strResult
            Dim strTagName
            Dim bRemove
            Dim bSearchForBlock

            nPos1 = InStr(strText, "<")
            Do While nPos1 > 0
                nPos2 = InStr(nPos1 + 1, strText, ">")
                If nPos2 > 0 Then
                    strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1)
                    strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ")

                    nPos3 = InStr(strTagName, " ")
                    If nPos3 > 0 Then
                        strTagName = Left(strTagName, nPos3 - 1)
                    End If

                    If Left(strTagName, 1) = "/" Then
                        strTagName = Mid(strTagName, 2)
                        bSearchForBlock = False
                    Else
                        bSearchForBlock = True
                    End If

                    If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                        bRemove = True
                        If bSearchForBlock Then
                            If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                                nPos2 = Len(strText)
                                nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare)
                                If nPos3 > 0 Then
                                    nPos3 = InStr(nPos3 + 1, strText, ">")
                                End If

                                If nPos3 > 0 Then
                                    nPos2 = nPos3
                                End If
                            End If
                        End If
                    Else
                        bRemove = False
                    End If

                    If bRemove Then
                        strResult = strResult & Left(strText, nPos1 - 1)
                        strText = Mid(strText, nPos2 + 1)
                    Else
                        strResult = strResult & Left(strText, nPos1)
                        strText = Mid(strText, nPos1 + 1)
                    End If
                Else
                    strResult = strResult & strText
                    strText = ""
                End If

                nPos1 = InStr(strText, "<")
            Loop
            strResult = strResult & strText

            fnRemoveHTML = strResult
        End Function

        Public Shared Function fnTABs(ByVal FillTab As Int16) As String
            Dim strTABs As String = ""
            For i As Int16 = 0 To FillTab - 1
                strTABs = strTABs & vbTab
            Next
            Return strTABs
        End Function

#Region "Emailing"
        Public Shared Sub prcEMail(ByVal strTo As String, ByVal strFrom As String, ByVal strSubject As String, ByVal strBody As String, Optional ByVal strCC As String = "", Optional ByVal strBCC As String = "")
            Try
                '*********2.0********
                strGlobalErrorInfo = ""
                Dim Mail As New System.Net.Mail.MailMessage()
                Mail.To.Add(strTo)         'User Email ID
                If strCC.Length > 0 Then Mail.CC.Add(strCC)
                If strBCC.Length > 0 Then Mail.Bcc.Add(strBCC)
                Mail.From = New MailAddress(strFrom)
                Mail.Subject = strSubject
                Mail.Body = strBody
                'SmtpMail.SmtpServer.Insert(0, "smtp.jaypeebrothers.com")
                Dim smtp As New SmtpClient("mail.jaypeebrothers.com")
                smtp.Send(Mail)
                strGlobalErrorInfo = "MAIL SENT SUCCESSFULLY!"



                '*********1.0********
                'strGlobalErrorInfo = ""
                'Dim Mail As New MailMessage()
                'Mail.To = strTo         'User Email ID
                'If strCC.Length > 0 Then Mail.Cc = strCC
                'If strBCC.Length > 0 Then Mail.Cc = strBCC
                'Mail.From = strFrom
                'Mail.Subject = strSubject
                'Mail.Body = strBody
                'SmtpMail.SmtpServer = "localhost"   'your real server goes here

                'Mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusing", 2) 'basic authentication
                ''Mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserver", "202.71.148.84") 'set your username here
                'Mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserver", "202.71.129.68") 'set your username here
                'Mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", 25)


                'Dim SmtpMail As New SmtpClient("mail.jaypeebrothers.com")
                'SmtpMail.Send(Mail)
                'strGlobalErrorInfo = "MAIL SENT SUCCESSFULLY!"
            Catch ex As Exception
                Throw ex 'MyCLS.prcEMailOnError(ex, , , , System.Reflection.MethodBase.GetCurrentMethod.ToString())
                strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.Data)
            End Try
        End Sub
        Public Shared Sub prcEMailCDO(ByVal strTo As String, ByVal strFrom As String, ByVal strSubject As String, ByVal strBody As String, Optional ByVal strCC As String = "", Optional ByVal strBCC As String = "")
            Dim mailMessage As New System.Net.Mail.MailMessage(strFrom, strTo, strSubject, strBody)
            Dim mailClient As New System.Net.Mail.SmtpClient("209.44.115.202", 25)
            mailClient.UseDefaultCredentials = True
            'mailClient.Credentials = New System.Net.NetworkCredential("10.90.2.143\in.itsupport", "123456")
            mailClient.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network

            Try
                If strCC.Length > 0 Then mailMessage.CC.Add(strCC)
                If strBCC.Length > 0 Then mailMessage.Bcc.Add(strBCC)
                mailClient.Send(mailMessage)
            Catch ex As Exception
                Throw ex 'MyCLS.prcEMailOnError(ex, , , , System.Reflection.MethodBase.GetCurrentMethod.ToString())
                strGlobalErrorInfo = ex.ToString
                Exit Sub
            End Try
        End Sub
        Public Shared Sub prcEMailOnError(ByVal ex As Exception, Optional ByVal strTo As String = "narender.sharma@netsoftit.com", Optional ByVal strFrom As String = "error@jaypeebrothers.com", Optional ByVal strSubject As String = "Error On Jaypeebrothers.com", Optional ByVal strBody As String = "", Optional ByVal strCC As String = "", Optional ByVal strBCC As String = "")
            Try
                If ex.Message.ToString().Contains("Thread was being aborted") = True Then Exit Sub

                strGlobalErrorInfo = ""
                strGlobalErrorInfo = "Error MSG : " & "<BR>" & "<BR>" & String.Concat(ex.Message & "<BR>", ex.Source & "<BR>", ex.StackTrace & "<BR>")
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & "<BR>", ex.TargetSite, ex.InnerException)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & "<BR>", ex.Data) & "<BR>" & "<BR>" & "<BR>" & "***" & "<BR>"

                Dim Mail As New System.Net.Mail.MailMessage()
                Mail.To.Add(strTo)         'User Email ID
                If strCC.Length > 0 Then Mail.CC.Add(strCC)
                If strBCC.Length > 0 Then Mail.Bcc.Add(strBCC)
                Mail.From = New MailAddress(strFrom)
                Mail.Subject = strSubject
                Mail.Body = strGlobalErrorInfo & strBody
                Mail.IsBodyHtml = True
                'SmtpMail.SmtpServer.Insert(0, "smtp.jaypeebrothers.com")
                Dim smtp As New SmtpClient("mail.jaypeebrothers.com")
                smtp.Send(Mail)
                strGlobalErrorInfo = "MAIL SENT SUCCESSFULLY!"
            Catch ex1 As Exception
                strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex1.Message & vbCrLf, ex1.Source & vbCrLf, ex1.StackTrace & vbCrLf)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex1.TargetSite, ex1.InnerException)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex1.Data)
            End Try
        End Sub
#End Region

        Public Shared Sub SaveSettings(ByVal DBFilePath As String, ByVal Server As String, ByVal UID As String, ByVal Password As String, ByVal Database As String, Optional ByVal isSource As Boolean = False)
            Try
                If isSource Then
                    MyCLS.clsFileHandling.OpenFile("DBSettingsSRC.txt")
                Else
                    MyCLS.clsFileHandling.OpenFile("DBSettings.txt")
                End If
                MyCLS.clsFileHandling.WriteFile(DBFilePath & "~" & Server & "~" & UID & "~" & Password & "~" & Database)
                MyCLS.clsFileHandling.CloseFile()
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Function GetSettings(Optional ByVal isSource As Boolean = False) As String()
            Dim strDBDetails As String
            Dim strDBDetailsSplit() As String = {""}
            Try
                If isSource Then
                    strDBDetails = MyCLS.clsFileHandling.ReadFile("DBSettingsSRC.txt")
                Else
                    strDBDetails = MyCLS.clsFileHandling.ReadFile("DBSettings.txt")
                End If
                'MsgBox(Asc(Mid(strDBDetails, Len(strDBDetails))))
                strDBDetails = strDBDetails.Replace(Chr(10), "").Replace(Chr(13), "")
                strDBDetailsSplit = strDBDetails.Split("~")
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return strDBDetailsSplit
        End Function

        Public Shared Sub prcFillTextAutoComplete(ByRef TxtBox As TextBox, ByVal Table As String, ByVal Column As String)
            Try
                Dim filterVals As New AutoCompleteStringCollection() '= New AutoCompleteStringCollection()

                MyRs = fnQuerySelectRS("Select " & Column & " From " & Table)

                While MyRs.Read
                    filterVals.Add(MyRs.GetValue(0).ToString)
                End While

                TxtBox.AutoCompleteSource = filterVals.ToString

            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
    End Class


#Region "Class to be used for List/CheckList/Grid/Grid With Checkbox"
    Public Class clsControls
        Public Shared Sub ComboBox_AdjustWidth(ByVal cbo As ComboBox)
            Try
                Dim g As Graphics = cbo.CreateGraphics()
                Dim maxWidth As Single = 0.0F

                For Each o As Object In cbo.Items
                    Dim w As Single = g.MeasureString(o.ToString(), cbo.Font).Width
                    If w > maxWidth Then
                        maxWidth = w
                    End If
                Next
                g.Dispose()
                cbo.Width = maxWidth
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Sub prcFillGrid(ByRef GridObj As Object, ByVal SelectQ As String)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                MyCmd = New OleDbCommand(SelectQ, MyCon)
                MyRs = MyCmd.ExecuteReader

                GridObj.DataSource = MyRs
                'GridObj.DataBind()
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Sub prcFillGridWin(ByRef objGrid As Object, ByVal SelectQ As String, ByVal TableName As String, ByVal isWithCheckBox As Boolean, Optional ByVal ChkColName As String = "", Optional ByVal ChkColText As String = "")
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                MyCmd = New OleDbCommand(SelectQ, MyCon)

                MyDa = New OleDbDataAdapter(MyCmd)
                MyDs = New DataSet

                MyDa.Fill(MyDs, TableName)

                objGrid.DataSource = MyDs
                objGrid.DataMember = TableName

                '****ADD CHECK BOX****
                If isWithCheckBox = True Then
                    Try
                        objGrid.Columns.Remove(ChkColName)
                    Catch ex As Exception

                    End Try


                    Dim myCheck As New DataGridViewCheckBoxColumn
                    myCheck.Name = ChkColName
                    myCheck.HeaderText = ChkColText
                    myCheck.DataPropertyName = ChkColName
                    myCheck.FalseValue = "0"
                    myCheck.TrueValue = "1"
                    objGrid.AutoGenerateColumns = False
                    objGrid.Columns.Insert(0, myCheck)
                    '********
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcFillGridWin(ByRef objGrid As Object, ByVal ds As DataSet, ByVal TableName As String, ByVal isWithCheckBox As Boolean, Optional ByVal ChkColName As String = "", Optional ByVal ChkColText As String = "")
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If

                objGrid.DataSource = ds
                objGrid.DataMember = TableName

                '****ADD CHECK BOX****
                If isWithCheckBox = True Then
                    Try
                        objGrid.Columns.Remove(ChkColName)
                    Catch ex As Exception

                    End Try


                    Dim myCheck As New DataGridViewCheckBoxColumn
                    myCheck.Name = ChkColName
                    myCheck.HeaderText = ChkColText
                    myCheck.DataPropertyName = ChkColName
                    myCheck.FalseValue = "0"
                    myCheck.TrueValue = "1"
                    objGrid.AutoGenerateColumns = False
                    objGrid.Columns.Insert(0, myCheck)
                    '********
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        'Public Shared Sub prcFillGridWin(ByRef objGrid As Object, ByVal objListing As Object, ByVal TableName As String)
        '    Try
        '        If Not MyCon.State = ConnectionState.Open Then
        '            MyCon.Open()
        '        End If

        '        objGrid.DataSource = objListing
        '        objGrid.DataMember = TableName
        '    Catch ex As Exception
        '        strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
        '        strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
        '        strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.Data)
        '        clsCOMMON.fnWrite2LOG(strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
        '    End Try
        'End Sub
        Public Shared Sub prcFillGridWinWithCHKBOX(ByRef objGrid As Object, ByVal SelectQ As String, ByVal TableName As String, ByVal ChkColName As String, ByVal ChkColText As String)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                MyCmd = New OleDbCommand(SelectQ, MyCon)

                MyDa = New OleDbDataAdapter(MyCmd)
                MyDs = New DataSet

                MyDa.Fill(MyDs, TableName)

                'MyDs.Tables.Add.Columns.Add(dtcCheck)
                'objGrid.Columns.Add(chk)


                objGrid.DataSource = MyDs
                objGrid.DataMember = TableName

                '****ADD CHECK BOX****
                Try
                    objGrid.Columns.Remove(ChkColName)
                Catch ex As Exception

                End Try


                Dim myCheck As New DataGridViewCheckBoxColumn
                myCheck.Name = ChkColName
                myCheck.HeaderText = ChkColText
                myCheck.DataPropertyName = ChkColName
                myCheck.FalseValue = "0"
                myCheck.TrueValue = "1"
                objGrid.AutoGenerateColumns = False
                objGrid.Columns.Insert(0, myCheck)
                '********
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Sub prcFillCombo(ByRef Cbo As ComboBox, ByVal SelectQ As String, ByVal TabName As String, ByVal Col1 As String, ByVal Col1Type As String, Optional ByVal Col2 As String = "", Optional ByVal Col2Type As String = "")
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                Dim DS As New DataSet
                MyCLS.clsCOMMON.fnQuerySelectDA(SelectQ).Fill(DS, TabName)

                Dim dt As New DataTable
                If Col1Type = "String" Then
                    dt.Columns.Add(Col1, GetType(System.String))
                End If
                If Len(Col2) > 0 And Col2Type = "String" Then
                    dt.Columns.Add(Col2, GetType(System.String))
                End If


                '
                ' Populate the DataTable to bind to the Combobox.
                '

                Dim drDSRow As DataRow
                Dim drNewRow As DataRow

                For Each drDSRow In DS.Tables(TabName).Rows()
                    drNewRow = dt.NewRow()
                    drNewRow(Col1) = drDSRow(Col1)
                    If Len(Col2) > 0 Then
                        drNewRow(Col2) = drDSRow(Col2)
                    End If
                    dt.Rows.Add(drNewRow)
                Next


                Cbo.DropDownStyle = ComboBoxStyle.DropDownList

                With Cbo
                    .DataSource = dt
                    .DisplayMember = Col1
                    If Len(Col2) > 0 Then
                        .ValueMember = Col2
                    End If
                    '.SelectedIndex = 0
                End With
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcFillCombo(ByRef Cbo As DataGridViewComboBoxColumn, ByVal SelectQ As String, ByVal TabName As String, ByVal Col1 As String, ByVal Col1Type As String, Optional ByVal Col2 As String = "", Optional ByVal Col2Type As String = "")
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                Dim DS As New DataSet
                MyCLS.clsCOMMON.fnQuerySelectDA(SelectQ).Fill(DS, TabName)

                Dim dt As New DataTable
                If Col1Type = "String" Then
                    dt.Columns.Add(Col1, GetType(System.String))
                End If
                If Len(Col2) > 0 And Col2Type = "String" Then
                    dt.Columns.Add(Col2, GetType(System.String))
                End If


                '
                ' Populate the DataTable to bind to the Combobox.
                '

                Dim drDSRow As DataRow
                Dim drNewRow As DataRow

                For Each drDSRow In DS.Tables(TabName).Rows()
                    drNewRow = dt.NewRow()
                    drNewRow(Col1) = drDSRow(Col1)
                    If Len(Col2) > 0 Then
                        drNewRow(Col2) = drDSRow(Col2)
                    End If
                    dt.Rows.Add(drNewRow)
                Next


                'Cbo.DropDownStyle = ComboBoxStyle.DropDownList

                With Cbo
                    .DataSource = dt
                    .DisplayMember = Col1
                    If Len(Col2) > 0 Then
                        .ValueMember = Col2
                    End If
                    '.SelectedIndex = 0
                End With
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Sub prcFillList(ByRef Lst As ListBox, ByVal SelectQ As String, ByVal TabName As String, ByVal Col1 As String, ByVal Col1Type As String, Optional ByVal Col2 As String = "", Optional ByVal Col2Type As String = "")
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If

                Dim DR As OleDbDataReader
                DR = MyCLS.clsCOMMON.fnQuerySelectRS(SelectQ)

                Lst.Items.Clear()
                While DR.Read()
                    Lst.Items.Add(DR(Col1).ToString)
                End While
                DR = Nothing


                'Dim DS As New DataSet
                'MyCLS.clsCOMMON.fnQuerySelectDA(SelectQ).Fill(DS, TabName)

                'Dim dt As New DataTable
                'If Col1Type = "String" Then
                '    dt.Columns.Add(Col1, GetType(System.String))
                'End If
                'If Len(Col2) > 0 And Col2Type = "String" Then
                '    dt.Columns.Add(Col2, GetType(System.String))
                'End If


                ''
                '' Populate the DataTable to bind to the Combobox.
                ''

                'Dim drDSRow As DataRow
                'Dim drNewRow As DataRow

                'For Each drDSRow In DS.Tables(TabName).Rows()
                '    drNewRow = dt.NewRow()
                '    drNewRow(Col1) = drDSRow(Col1)
                '    If Len(Col2) > 0 Then
                '        drNewRow(Col2) = drDSRow(Col2)
                '    End If
                '    dt.Rows.Add(drNewRow)
                'Next

                'With Lst              
                '    .DataSource = dt
                '    .DisplayMember = Col1
                '    If Len(Col2) > 0 Then
                '        .ValueMember = Col2
                '    End If
                '    '.SelectedIndex = 0
                'End With
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Sub prcFillListChecked(ByRef Lst As CheckedListBox, ByVal SelectQ As String, ByVal TabName As String, ByVal Col1 As String, ByVal Col1Type As String, Optional ByVal Col2 As String = "", Optional ByVal Col2Type As String = "")
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If

                Dim DR As OleDbDataReader
                DR = MyCLS.clsCOMMON.fnQuerySelectRS(SelectQ)

                Lst.Items.Clear()
                While DR.Read()
                    Lst.Items.Add(DR(Col1).ToString, False)
                End While
                DR = Nothing


                'Dim DS As New DataSet
                'MyCLS.clsCOMMON.fnQuerySelectDA(SelectQ).Fill(DS, TabName)

                'Dim dt As New DataTable
                'If Col1Type = "String" Then
                '    dt.Columns.Add(Col1, GetType(System.String))
                'End If
                'If Len(Col2) > 0 And Col2Type = "String" Then
                '    dt.Columns.Add(Col2, GetType(System.String))
                'End If


                ''
                '' Populate the DataTable to bind to the Combobox.
                ''

                'Dim drDSRow As DataRow
                'Dim drNewRow As DataRow

                'For Each drDSRow In DS.Tables(TabName).Rows()
                '    drNewRow = dt.NewRow()
                '    drNewRow(Col1) = drDSRow(Col1)
                '    If Len(Col2) > 0 Then
                '        drNewRow(Col2) = drDSRow(Col2)
                '    End If
                '    dt.Rows.Add(drNewRow)
                'Next

                'With Lst              
                '    .DataSource = dt
                '    .DisplayMember = Col1
                '    If Len(Col2) > 0 Then
                '        .ValueMember = Col2
                '    End If
                '    '.SelectedIndex = 0
                'End With
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcFillListChecked(ByRef Lst As CheckedListBox, ByVal SelectQ As String, ByVal TabName As String, ByVal Col1 As String, ByVal Col1Type As String, ByVal NoofCols As Int16, Optional ByVal Col2 As String = "", Optional ByVal Col2Type As String = "")
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If

                If NoofCols = 1 Then
                    Dim DR As OleDbDataReader
                    DR = MyCLS.clsCOMMON.fnQuerySelectRS(SelectQ)

                    While DR.Read()
                        Lst.Items.Add(DR(Col1).ToString, False)
                    End While
                    DR = Nothing
                ElseIf NoofCols = 2 Then
                    Lst.MultiColumn = True
                    Dim DR As OleDbDataReader
                    DR = MyCLS.clsCOMMON.fnQuerySelectRS(SelectQ)

                    '*********ANOTHER WAY**********
                    'Dim DS As New DataSet
                    'MyCLS.clsCOMMON.fnQuerySelectDA(SelectQ).Fill(DS, TabName)

                    'Dim dt As New DataTable
                    'If Col1Type = "String" Then
                    '    dt.Columns.Add(Col1, GetType(System.String))
                    'End If
                    'If Len(Col2) > 0 And Col2Type = "String" Then
                    '    dt.Columns.Add(Col2, GetType(System.String))
                    'End If



                    'Dim drDSRow As DataRow
                    'Dim drNewRow As DataRow

                    'For Each drDSRow In DS.Tables(TabName).Rows()
                    '    drNewRow = dt.NewRow()
                    '    drNewRow(Col1) = drDSRow(Col1)
                    '    If Len(Col2) > 0 Then
                    '        drNewRow(Col2) = drDSRow(Col2)
                    '    End If
                    '    dt.Rows.Add(drNewRow)
                    'Next
                    'Lst.Items.Add(dt)                
                    '*******************************

                    While DR.Read()
                        Lst.Items.Add(DR(Col1).ToString, False)
                    End While
                    DR = Nothing
                Else
                    MsgBox("Col No Must be 1 or 2!", MsgBoxStyle.Critical, "Invalid No of Cols")
                    Exit Sub
                End If





                ''
                '' Populate the DataTable to bind to the Combobox.
                ''


                'With Lst              
                '    .DataSource = dt
                '    .DisplayMember = Col1
                '    If Len(Col2) > 0 Then
                '        .ValueMember = Col2
                '    End If
                '    '.SelectedIndex = 0
                'End With
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Function fnListIsChecked(ByRef ListObj As Object) As Boolean
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                For intCounter = 0 To ListObj.Items.Count - 1
                    If ListObj.Items(intCounter).Selected = True Then
                        fnListIsChecked = True
                        Exit Function
                    End If
                Next
                fnListIsChecked = False
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Function fnListIsChecked(ByRef ListObj As CheckedListBox) As Boolean
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                For intCounter = 0 To ListObj.Items.Count - 1
                    If ListObj.GetItemChecked(intCounter) = True Then
                        fnListIsChecked = True
                        Exit Function
                    End If
                Next
                fnListIsChecked = False
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Sub prcListCheckAll(ByRef ListObj As Object)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                For intCounter = 0 To ListObj.Items.Count - 1
                    ListObj.Items(intCounter).Selected = True
                Next
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcListUnCheckAll(ByRef ListObj As Object)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                For intCounter = 0 To ListObj.Items.Count - 1
                    ListObj.Items(intCounter).Selected = False
                Next
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcListCheckAll(ByRef ListObj As CheckedListBox)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                For intCounter = 0 To ListObj.Items.Count - 1
                    ListObj.SetItemCheckState(intCounter, CheckState.Checked)
                Next
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcListUnCheckAll(ByRef ListObj As CheckedListBox)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                For intCounter = 0 To ListObj.Items.Count - 1
                    ListObj.SetItemCheckState(intCounter, CheckState.Unchecked)
                Next
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcListCheckSelected(ByRef ListObj As CheckedListBox, ByVal strSelectedValues As String())
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                prcListUnCheckAll(ListObj)
                For intCounter = 0 To strSelectedValues.Count - 1
                    Try
                        ListObj.SetItemChecked(ListObj.FindStringExact(strSelectedValues(intCounter).ToString()), True)
                    Catch ex As Exception
                    End Try
                Next
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcListCheckSelected(ByRef ListObj As ListBox, ByVal strSelectedValues As String())
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                prcListUnCheckAll(ListObj)
                For intCounter = 0 To strSelectedValues.Count - 1
                    Try
                        ListObj.SetSelected(ListObj.FindStringExact(strSelectedValues(intCounter).ToString()), True)
                    Catch ex As Exception
                    End Try
                Next
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Function fnGridIsChecked(ByRef GridObj As DataGridView, ByVal isWithCheckBox As Boolean) As Boolean
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                If isWithCheckBox = True Then
                    For intCounter = 0 To GridObj.RowCount - 1
                        If GridObj.Rows(intCounter).Cells(0).Value = True Then
                            fnGridIsChecked = True
                            Exit Function
                        End If
                    Next
                    fnGridIsChecked = False
                Else
                    For intCounter = 0 To GridObj.RowCount - 1
                        If GridObj.Rows(intCounter).Selected = True Then
                            fnGridIsChecked = True
                            Exit Function
                        End If
                    Next
                    fnGridIsChecked = False
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Sub prcGridCheckAll(ByRef GridObj As DataGridView, ByVal isWithCheckBox As Boolean)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                '****IF CHECK BOX****
                If isWithCheckBox = True Then
                    For intCounter = 0 To GridObj.RowCount - 1
                        GridObj.Rows(intCounter).Cells(0).Value = True
                    Next
                Else
                    For intCounter = 0 To GridObj.RowCount - 1
                        GridObj.Rows(intCounter).Selected = True
                    Next
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcGridUnCheckAll(ByRef GridObj As DataGridView, ByVal isWithCheckBox As Boolean)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intCounter As Integer
                '****IF CHECK BOX****
                If isWithCheckBox = True Then
                    For intCounter = 0 To GridObj.RowCount - 1
                        GridObj.Rows(intCounter).Cells(0).Value = False
                    Next
                Else
                    For intCounter = 0 To GridObj.RowCount - 1
                        GridObj.Rows(intCounter).Selected = False
                    Next
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Sub prcGridCheckSelected(ByRef GridObj As DataGridView, ByVal isWithCheckBox As Boolean, ByVal strSelectedValues As String(), ByVal intColumnNo As Int16)
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                Dim intGCounter As Integer
                Dim intSCounter As Integer
                prcGridUnCheckAll(GridObj, isWithCheckBox)
                '****IF CHECK BOX****
                If isWithCheckBox = True Then
                    For intGCounter = 0 To GridObj.RowCount - 1
                        For intSCounter = 0 To strSelectedValues.Count - 1
                            Try
                                If GridObj.Rows(intGCounter).Cells(intColumnNo).Value() = strSelectedValues(intSCounter).ToString() Then
                                    GridObj.Rows(intGCounter).Cells(0).Value = True
                                End If
                            Catch ex As Exception
                            End Try
                        Next
                    Next
                Else
                    For intGCounter = 0 To GridObj.RowCount - 1
                        For intSCounter = 0 To strSelectedValues.Count - 1
                            Try
                                If GridObj.Rows(intGCounter).Cells(intColumnNo).Value() = strSelectedValues(intSCounter).ToString() Then
                                    GridObj.Rows(intGCounter).Selected = True
                                End If
                            Catch ex As Exception
                            End Try
                        Next
                    Next
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

#Region "FILL DDL Using DS"
        Enum SortOrder
            NOSORT = 0
            ASC = 1
            DESC = 2
        End Enum
        Public Shared Sub prcFillDDLUsingDS(ByVal DDLObj As ComboBox, ByVal TblName As String, ByVal IDCol As String, ByVal ValueCol As String, ByVal WhereClause As String, ByVal OrderByCol As String, ByVal SortOrder As SortOrder, ByVal FirstRowBlank As Boolean)
            Try
                strGlobalErrorInfo = ""

                MyStr = "SELECT " & ValueCol & ", " & IDCol & " FROM " & TblName

                If Len(WhereClause) > 0 Then
                    MyStr = MyStr & " " & WhereClause & " "
                End If

                If SortOrder.ASC Then
                    MyStr = MyStr & " ORDER BY " & OrderByCol & " ASC"
                ElseIf SortOrder.DESC Then
                    MyStr = MyStr & " ORDER BY " & OrderByCol & " DESC"
                Else
                    MyStr = MyStr & " ORDER BY " & OrderByCol
                End If

                Dim ds As New DataSet
                MyCLS.clsCOMMON.prcQuerySelectDS(ds, MyStr, TblName)
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables.Count > 0 Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                DDLObj.DataSource = ds.Tables(0)
                                DDLObj.DisplayMember = Replace(ValueCol, "distinct ", "", , , CompareMethod.Text)
                                DDLObj.ValueMember = IDCol
                            End If
                        End If
                    End If
                End If

                If FirstRowBlank Then
                    DDLObj.Items.Insert(0, "")
                End If

            Catch ex As Exception
                strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
                strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.Data)
            End Try
        End Sub
#End Region
    End Class
#End Region


    '**************START - TO MAIL*****************************************************
    Public Class clsMail
        Public Shared strFrom As String
        Public Shared strPassword As String
        Public Shared strSMTPClient As String
        Public Shared intPort As Int16
        Public Shared isBodyHTML As Boolean

        Public Shared Sub prcEMail(ByVal strTo As String, ByVal strSub As String, ByVal strBody As String, Optional ByVal strCC As String = "", Optional ByVal strBCC As String = "", Optional ByVal isRequireSSL As Boolean = False)
            Try
                '************SEND MAIL USING GMAIL******************************       
                'intPort=587    -   IF GMAIL
                'intPort=0      -   IF LOCALHOST

                Dim Mail As New System.Net.Mail.MailMessage()

                Mail.To.Add(strTo)
                Mail.From = New MailAddress(strFrom)
                If strCC.Length > 0 Then Mail.CC.Add(strCC)
                If strBCC.Length > 0 Then Mail.Bcc.Add(strBCC)
                Mail.Subject = strSub
                Mail.Body = strBody
                Mail.IsBodyHtml = isBodyHTML

                'SMTP CLIENT IS LOCALHOST OR NOT
                Dim Smtp As SmtpClient
                If intPort = 0 Then
                    Smtp = New SmtpClient(strSMTPClient)
                Else
                    Smtp = New SmtpClient(strSMTPClient, intPort)
                End If

                'USER CREDENTIALS REQUIRED OR NOT
                If Len(strFrom) > 0 And Len(strPassword) > 0 Then
                    Smtp.Credentials = New Net.NetworkCredential(strFrom, strPassword)
                Else
                    Smtp.UseDefaultCredentials = True
                End If

                'ENABLE SSL OR NOT
                'If isRequireSSL = False Then
                Smtp.EnableSsl = isRequireSSL
                'Else
                'Smtp.EnableSsl = True
                'End If                

                'SEND MAIL
                '*System.Windows.Forms.Application.DoEvents()
                Smtp.Send(Mail)
                '*System.Windows.Forms.Application.DoEvents()
            Catch ex As Exception
                'MsgBox(ex.Message)
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try


            '************START - SEND MAIL USING MICROSOFT OUTLOOK******************************
            '---'ByVal AppOL As Outlook.Application, ByVal OLNameSpace As Outlook.NameSpace
            '' ''Dim msg As Outlook.MailItem
            ' '' ''dim SM as Outlook.
            ' '' ''creating newblan mail message 
            '' ''msg = AppOL.CreateItem(Outlook.OlItemType.olMailItem)
            ' '' ''msg.Permission = Outlook.OlPermission.olUnrestricted

            ' '' ''AppOL.COMAddIns
            ' '' ''Adding recipeints to the mail message 
            '' ''msg.Recipients.Add("rahul.dss@gmail.com")
            ' '' ''adding subject information to the mail message 
            '' ''msg.Subject = strSub
            ' '' ''adding body message information to the mail message 
            '' ''msg.Body = strBody
            ' '' ''sending message 
            '' ''msg.Send()
            '************END - SEND MAIL USING MICROSOFT OUTLOOK******************************
        End Sub

        '***** FOR DESKTOP *****
        Public Shared Sub prcEMailOL(ByVal strTo As String, ByVal strSub As String, ByVal strBody As String, Optional ByVal strCC As String = "", Optional ByVal strBCC As String = "")
            Dim AppOL As New Outlook.Application
            'creating new blank mail message
            Dim Msg As Outlook.MailItem
            Msg = AppOL.CreateItem(Outlook.OlItemType.olMailItem)
            Try
                '************START - SEND MAIL USING MICROSOFT OUTLOOK******************************                
                Dim OLNameSpace As Outlook.NameSpace
                'dim SM as Outlook.

                'msg.Permission = Outlook.OlPermission.olUnrestricted

                'AppOL.COMAddIns
                'Adding recipeints to the mail message 
                Msg.To = strTo
                Msg.CC = strCC
                Msg.BCC = strBCC
                'Msg.Recipients.Add(strTo)

                'adding subject information to the mail message 
                Msg.Subject = strSub
                'adding body message information to the mail message 
                'Msg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
                Msg.HTMLBody = strBody
                'Msg.Body = strBody
                System.Windows.Forms.Application.DoEvents()
                'Show message 
                Msg.Display(1)
                System.Windows.Forms.Application.DoEvents()
                'sending message 
                'Msg.Send()
                AppOL = Nothing
                'AppOL.Quit()
                '************END - SEND MAIL USING MICROSOFT OUTLOOK******************************
            Catch ex As Exception
                MsgBox(ex.Message)
                Msg.Save()
                'AppOL.Quit()
                AppOL = Nothing
            End Try
        End Sub

        Public Shared Function prcFindInFile(ByVal strFilePath As String, ByVal strFind As String, ByVal strReplace As String) As String
            Dim strFile As String = ""
            Try
                Dim FR As IO.StreamReader
                FR = File.OpenText(strFilePath)

                While Not FR.EndOfStream
                    strFile = FR.ReadToEnd()
                End While
                'MsgBox(Len(strFile) & vbCrLf & strFile)
                'MsgBox("<a href=""#"">")            
                strFile = strFile.Replace(strFind, strReplace)
                'MsgBox(Len(strFile) & vbCrLf & strFile)
                FR.Close()
                Return strFile
            Catch ex As Exception
                'MsgBox(ex.Message)
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
                Return strFile
            End Try
        End Function

        Public Shared Function prcFindInString(ByVal strString As String, ByVal strFind As String, ByVal strReplace As String) As String
            Dim strModified As String = ""
            Try
                Dim FR As New StringReader(strString)

                While FR.Peek() <> -1
                    'MsgBox(FR.Peek & vbCrLf & strModified)
                    strModified = strModified & FR.ReadLine()
                End While
                'MsgBox(Len(strModified) & vbCrLf & strModified)
                'MsgBox("<a href=""#"">")            

                strModified = strModified.Replace(strFind, strReplace)
                'MsgBox(Len(strModified) & vbCrLf & strModified)
                FR.Close()
                Return strModified
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
                'MsgBox(ex.Message)
                Return strModified
            End Try
        End Function
    End Class
    '**************END - TO MAIL*****************************************************




#Region "DataBase Operation using MVC Model"
    '**************START - TO STORE SP ERROR FLAGS*****************************************************
#Region "SPStatus of the Action Perform on SP"
    Public Enum SPStatus As Integer
        Err_Updating = -2
        Err_Inserting = -1
        Failure = -1
        Success_Inserting = 1
        Success_Updating = 2
        RecordNotFound = 3
        Duplicate = 4
        Exception = 5
    End Enum
    'end SPStatus
#End Region
    '**************END - TO STORE SP ERROR FLAGS*****************************************************

    '**************START - TO STORE ERROR FLAGS*****************************************************
#Region "EStatus of the Action Perform"
    Public Enum EStatus As Integer
        Err = -1
        Failure
        Success
        DatabaseNotFound
        Exception
        RecordNotFound
        Duplicate
    End Enum
    'end EStatus
#End Region
    '**************END - TO STORE ERROR FLAGS*****************************************************

    '**************START - TO STORE MESSAGE TYPES*****************************************************
#Region "EMessageType of the Action Perform"
    Public Enum EMessageType As Integer
        Overtime = 1
        payperiodtype = 2
        Payperiodday = 3
        Payperiodtimeslot = 4
        payperioddays = 5
        ColorCode = 6
        PayType = 7
        YesNo = 8
    End Enum
#End Region
    '**************END - TO STORE MESSAGE TYPES*****************************************************


    '*************START - TRANSPORTPACKET CLASS*************************************************************
    <Serializable()> _
    Public Class TransportationPacket

        Public Sub New()
            'Default Constructor 
        End Sub

        Protected Overrides Sub Finalize()
            Try
                'Default Destructor 
            Finally
                MyBase.Finalize()
            End Try
        End Sub

        Public Overridable Sub Dispose()

        End Sub

        Private _CommandText As String
        Public Property CommandText() As String
            Get
                Return _CommandText
            End Get
            Set(ByVal value As String)
                _CommandText = value
            End Set
        End Property

        Private _MessageId As Integer
        Public Property MessageId() As Integer
            Get
                Return _MessageId
            End Get
            Set(ByVal value As Integer)
                _MessageId = value
            End Set
        End Property

        Private _MessagePacket As Object
        Public Property MessagePacket() As Object
            Get
                Return _MessagePacket
            End Get
            Set(ByVal value As Object)
                _MessagePacket = value
            End Set
        End Property

        Private _MessageResultset As Object
        Public Property MessageResultset() As Object
            Get
                Return _MessageResultset
            End Get
            Set(ByVal value As Object)
                _MessageResultset = value
            End Set
        End Property

        Private _MessageResultsetDS As Object
        Public Property MessageResultsetDS() As Object
            Get
                Return _MessageResultsetDS
            End Get
            Set(ByVal value As Object)
                _MessageResultsetDS = value
            End Set
        End Property

        Private _MessageResultsetS As Object
        Public Property MessageResultsetS() As Object
            Get
                Return _MessageResultsetS
            End Get
            Set(ByVal value As Object)
                _MessageResultsetS = value
            End Set
        End Property

        Private _MessageStatus As EStatus
        Public Property MessageStatus() As EStatus
            Get
                Return _MessageStatus
            End Get
            Set(ByVal value As EStatus)
                _MessageStatus = value
            End Set
        End Property

        Private _MessageType As EMessageType
        Public Property MessageType() As EMessageType
            Get
                Return _MessageType
            End Get
            Set(ByVal value As EMessageType)
                _MessageType = value
            End Set
        End Property
    End Class
    '*************END - TRANSPORTPACKET CLASS*************************************************************

    '**************START - TO EXECUTE STORED PROCEDURES*****************************************************
#Region "Execute using OleDb"
    Public Class clsExecuteStoredProcOleDb

        Shared m_OleDbConnection As OleDbConnection

#Region "Open And Close Database Connection"
        ''' <summary>
        ''' This function Opens Connection to be used within this class
        ''' And Closes after operation is completed
        ''' </summary>
        ''' <returns></returns>
        Shared Function OpenDatabase() As OleDbConnection
            Try
                strGlobalErrorInfo = ""
                'MyCLS.clsCOMMON.ConOpen()
                MyCLS.clsCOMMON.GetCon(m_OleDbConnection)
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_OleDbConnection
        End Function

        Shared Sub CloseDatabase()
            Try
                strGlobalErrorInfo = ""
                If m_OleDbConnection.State Then
                    m_OleDbConnection.Close()
                End If
                'MyCLS.clsCOMMON.ConClose()
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
        End Sub
#End Region

#Region "My Function for Stored Procedure"
        '' ''Public Shared Sub ExecuteStoredProc(ByRef ds As DataSet, ByVal CMD As OleDbCommand, ByVal CommandText As String)
        '' ''    Try
        '' ''        strGlobalErrorInfo = ""
        '' ''        CMD.CommandText = CommandText
        '' ''        CMD.CommandType = CommandType.StoredProcedure
        '' ''        CMD.Connection = MyCon

        '' ''        MyDa = New OleDbDataAdapter
        '' ''        MyDa.SelectCommand = CMD

        '' ''        MyDa.Fill(ds)
        '' ''    Catch ex As Exception
        '' ''        strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
        '' ''        strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
        '' ''        strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.Data)
        '' ''    End Try
        '' ''End Sub
#End Region

#Region "ExecuteNonQuery"
        ''' <summary>
        ''' This function shall Execute the Stored Procedure on the Database, this is a replica of using DOTNET ExecuteNonQuery
        ''' This shall reduce lot of Development time in invoking the database properties.
        ''' Input Parameters: String SPName -> Name of the Stored Procedures
        ''' ParameterList -> List of Type SQLParameter
        ''' The function is responsible for database connectivity and shall open and close the connection on it's own.
        ''' </summary>
        ''' <param name="SPName"></param>
        ''' <param name="ParameterList"></param>
        ''' <returns></returns>
        Public Shared Function ExecuteSPNonQuery(ByVal SPName As String, ByVal ParameterList As List(Of OleDbParameter)) As Integer
            Dim m_intReturnValue As Integer = 0
            Try
                Dim m_cmdStoredProcedure As New OleDbCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_OleDbConnection Is Nothing Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection.State <> ConnectionState.Open Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_OleDbConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()
                    CloseDatabase()
                End If
            Catch exObj As Exception
                m_intReturnValue = -1
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_intReturnValue
        End Function
#End Region

#Region "ExecuteSPDataSet"
        ''' <summary>
        ''' This function shall Execute the Stored Procedure on the Database, this is a replica of using DOTNET ExecuteReader 
        ''' or the method of filling up the DataSet.
        ''' This shall reduce lot of Development time in invoking the database properties.
        ''' Input Parameters: String SPName -> Name of the Stored Procedures
        ''' ParameterList -> List of Type SQLParameter
        ''' The function is responsible for database connectivity and shall open and close the connection on it's own.
        ''' </summary>
        ''' <param name="SPName"></param>
        ''' <param name="ParameterList"></param>
        ''' <returns></returns>
        Public Shared Function ExecuteSPDataSet(ByVal SPName As String, ByVal ParameterList As List(Of OleDbParameter)) As DataSet
            Dim m_dsReturnValue As New DataSet()
            Try
                Dim m_cmdStoredProcedure As New OleDbCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_OleDbConnection Is Nothing Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection.State <> ConnectionState.Open Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_OleDbConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    Dim daAdapater As New OleDbDataAdapter(m_cmdStoredProcedure)
                    daAdapater.Fill(m_dsReturnValue)
                    CloseDatabase()
                End If
            Catch exObj As Exception
                m_dsReturnValue = Nothing
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_dsReturnValue
        End Function

        Public Shared Function ExecuteSPDataSet(ByVal SPName As String) As DataSet
            Dim m_dsReturnValue As New DataSet()
            Try
                Dim m_cmdStoredProcedure As New OleDbCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_OleDbConnection Is Nothing Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection.State <> ConnectionState.Open Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_OleDbConnection
                    Dim daAdapater As New OleDbDataAdapter(m_cmdStoredProcedure)
                    daAdapater.Fill(m_dsReturnValue)
                    CloseDatabase()
                End If
            Catch exObj As Exception
                m_dsReturnValue = Nothing
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_dsReturnValue
        End Function

        Public Shared Function ExecuteSPDataAdapter(ByVal SPName As String, ByVal ParameterList As List(Of OleDbParameter)) As OleDbDataAdapter
            Dim daAdapater As OleDbDataAdapter
            Try
                Dim m_cmdStoredProcedure As New OleDbCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_OleDbConnection Is Nothing Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection.State <> ConnectionState.Open Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_OleDbConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    daAdapater = New OleDbDataAdapter(m_cmdStoredProcedure)
                    CloseDatabase()
                End If
            Catch exObj As Exception
                daAdapater = Nothing
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return daAdapater
        End Function
#End Region

#Region "ExecuteSPScalar"
        Public Shared Function ExecuteSPScalar(ByVal SPName As String, ByVal ParameterList As List(Of OleDbParameter)) As Integer
            Dim m_intReturnValue As Integer = 0
            Try
                Dim m_cmdStoredProcedure As New OleDbCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_OleDbConnection Is Nothing Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection.State <> ConnectionState.Open Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_OleDbConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    m_intReturnValue = Convert.ToInt32(m_cmdStoredProcedure.ExecuteScalar())
                    CloseDatabase()
                End If
            Catch exObj As Exception
                m_intReturnValue = -1
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_intReturnValue
        End Function
#End Region

#Region "ExecuteNonQuery OutPut"
        ''' <summary> 
        ''' This function shall Execute the Stored Procedure on the Database, this is a replica of using DOTNET ExecuteNonQuery 
        ''' This shall reduce lot of Development time in invoking the database properties. 
        ''' Input Parameters: String SPName -> Name of the Stored Procedures 
        ''' ParameterList -> List of Type SQLParameter 
        ''' The function is responsible for database connectivity and shall open and close the connection on it's own. 
        ''' </summary> 
        ''' <param name="SPName"></param> 
        ''' <param name="ParameterList"></param> 
        ''' <returns></returns> 
        Public Shared Function ExecuteSPNonQueryOutPut(ByVal SPName As String, ByVal ParameterList As List(Of OleDbParameter), ByVal OutParameterList As List(Of OleDbParameter), ByRef m_intReturnValue As Int16) As String()
            m_intReturnValue = 0
            Dim OutParameterArray As String() = New String(OutParameterList.Count - 1) {}
            Try
                Dim m_cmdStoredProcedure As New OleDbCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_OleDbConnection Is Nothing Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection.State <> ConnectionState.Open Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_OleDbConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next

                    For intLoop As Integer = 0 To OutParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(OutParameterList(intLoop))
                        OutParameterList(intLoop).Direction = ParameterDirection.Output
                    Next

                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()

                    CloseDatabase()

                    For intLoop As Integer = 0 To OutParameterList.Count - 1
                        OutParameterArray(intLoop) = m_cmdStoredProcedure.Parameters(OutParameterList(intLoop).ParameterName).Value.ToString()
                    Next
                End If
            Catch exObj As Exception
                m_intReturnValue = -1
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return OutParameterArray
        End Function

        Public Shared Function ExecuteSPNonQueryReturnValue(ByVal SPName As String, ByVal ParameterList As List(Of OleDbParameter)) As Integer
            Dim m_intReturnValue As Integer = 0
            Dim intReturnValues As Integer = 0

            Try
                Dim m_cmdStoredProcedure As New OleDbCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_OleDbConnection Is Nothing Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection.State <> ConnectionState.Open Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_OleDbConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    Dim objReturnParameter As New OleDbParameter("RETURNVALUE", OleDbType.Integer)
                    objReturnParameter.Direction = ParameterDirection.ReturnValue

                    m_cmdStoredProcedure.Parameters.Add(objReturnParameter)


                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()
                    intReturnValues = Convert.ToInt32(m_cmdStoredProcedure.Parameters("RETURNVALUE").Value)

                    CloseDatabase()
                End If
            Catch exObj As Exception
                intReturnValues = -1
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return intReturnValues + m_intReturnValue
        End Function

        Public Shared Function ExecuteNonQueryReturnValueWithoutAdd(ByVal SPName As String, ByVal ParameterList As List(Of OleDbParameter)) As Integer
            Dim m_intReturnValue As Integer = 0
            Dim intReturnValues As Integer = 0

            Try
                Dim m_cmdStoredProcedure As New OleDbCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_OleDbConnection Is Nothing Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection.State <> ConnectionState.Open Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_OleDbConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    Dim objReturnParameter As New OleDbParameter("RETURNVALUE", OleDbType.Integer)
                    objReturnParameter.Direction = ParameterDirection.ReturnValue

                    m_cmdStoredProcedure.Parameters.Add(objReturnParameter)


                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()
                    intReturnValues = Convert.ToInt32(m_cmdStoredProcedure.Parameters("RETURNVALUE").Value)

                    CloseDatabase()
                End If
            Catch exObj As Exception
                intReturnValues = -1
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return intReturnValues
        End Function


        Public Shared Function ExecuteNonQueryDecimal(ByVal SPName As String, ByVal ParameterList As List(Of OleDbParameter)) As Decimal
            Dim m_intReturnValue As Integer = 0
            Dim intReturnValues As Decimal = 0

            Try
                Dim m_cmdStoredProcedure As New OleDbCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_OleDbConnection Is Nothing Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection.State <> ConnectionState.Open Then
                    m_OleDbConnection = OpenDatabase()
                End If
                If m_OleDbConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_OleDbConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    Dim objReturnParameter As New OleDbParameter("RETURNVALUE", OleDbType.[Decimal])
                    objReturnParameter.Direction = ParameterDirection.ReturnValue

                    m_cmdStoredProcedure.Parameters.Add(objReturnParameter)


                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()
                    intReturnValues = Convert.ToDecimal(m_cmdStoredProcedure.Parameters("RETURNVALUE").Value)

                    CloseDatabase()
                End If
            Catch exObj As Exception
                intReturnValues = -1
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return intReturnValues
        End Function
#End Region
    End Class
#End Region

#Region "Execute using Sql"
    Public Class clsExecuteStoredProcSql

        Shared m_SqlConnection As SqlConnection

#Region "Open And Close Database Connection"
        ''' <summary>
        ''' This function Opens Connection to be used within this class
        ''' And Closes after operation is completed
        ''' </summary>
        ''' <returns></returns>
        Shared Function OpenDatabase() As SqlConnection
            Try
                strGlobalErrorInfo = ""
                'MyCLS.clsCOMMON.ConOpen()
                MyCLS.clsCOMMON.GetCon(m_SqlConnection)
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_SqlConnection
        End Function

        Shared Sub CloseDatabase()
            Try
                strGlobalErrorInfo = ""
                If m_SqlConnection.State Then
                    m_SqlConnection.Close()
                End If
                'MyCLS.clsCOMMON.ConClose()
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
        End Sub
#End Region

#Region "My Function for Stored Procedure"
        '' ''Public Shared Sub ExecuteStoredProc(ByRef ds As DataSet, ByVal CMD As SqlCommand, ByVal CommandText As String)
        '' ''    Try
        '' ''        strGlobalErrorInfo = ""
        '' ''        CMD.CommandText = CommandText
        '' ''        CMD.CommandType = CommandType.StoredProcedure
        '' ''        CMD.Connection = MyCon

        '' ''        MyDa = New SqlDataAdapter
        '' ''        MyDa.SelectCommand = CMD

        '' ''        MyDa.Fill(ds)
        '' ''    Catch ex As Exception
        '' ''        strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
        '' ''        strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
        '' ''        strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.Data)
        '' ''    End Try
        '' ''End Sub
#End Region

#Region "ExecuteNonQuery"
        ''' <summary>
        ''' This function shall Execute the Stored Procedure on the Database, this is a replica of using DOTNET ExecuteNonQuery
        ''' This shall reduce lot of Development time in invoking the database properties.
        ''' Input Parameters: String SPName -> Name of the Stored Procedures
        ''' ParameterList -> List of Type SQLParameter
        ''' The function is responsible for database connectivity and shall open and close the connection on it's own.
        ''' </summary>
        ''' <param name="SPName"></param>
        ''' <param name="ParameterList"></param>
        ''' <returns></returns>
        Public Shared Function ExecuteSPNonQuery(ByVal SPName As String, ByVal ParameterList As List(Of SqlParameter)) As Integer
            Dim m_intReturnValue As Integer = 0
            Try
                Dim m_cmdStoredProcedure As New SqlCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()
                    CloseDatabase()
                End If
            Catch exObj As Exception
                m_intReturnValue = -1
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_intReturnValue
        End Function
#End Region

#Region "ExecuteSPDataSet"
        ''' <summary>
        ''' This function shall Execute the Stored Procedure on the Database, this is a replica of using DOTNET ExecuteReader 
        ''' or the method of filling up the DataSet.
        ''' This shall reduce lot of Development time in invoking the database properties.
        ''' Input Parameters: String SPName -> Name of the Stored Procedures
        ''' ParameterList -> List of Type SQLParameter
        ''' The function is responsible for database connectivity and shall open and close the connection on it's own.
        ''' </summary>
        ''' <param name="SPName"></param>
        ''' <param name="ParameterList"></param>
        ''' <returns></returns>
        Public Shared Function ExecuteSPDataSet(ByVal SPName As String, ByVal ParameterList As List(Of SqlParameter)) As DataSet
            Dim m_dsReturnValue As New DataSet()
            Try
                Dim m_cmdStoredProcedure As New SqlCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    Dim daAdapater As New SqlDataAdapter(m_cmdStoredProcedure)
                    daAdapater.Fill(m_dsReturnValue)
                    CloseDatabase()
                End If
            Catch exObj As Exception
                m_dsReturnValue = Nothing
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_dsReturnValue
        End Function

        Public Shared Function ExecuteSPDataSet(ByVal SPName As String, Optional ByVal TableName As String = "") As DataSet
            Dim m_dsReturnValue As New DataSet()
            Try
                Dim m_cmdStoredProcedure As New SqlCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    Dim daAdapater As New SqlDataAdapter(m_cmdStoredProcedure)
                    If Len(TableName) > 0 Then
                        daAdapater.Fill(m_dsReturnValue, TableName)
                    Else
                        daAdapater.Fill(m_dsReturnValue)
                    End If
                    CloseDatabase()
                End If
            Catch exObj As Exception
                m_dsReturnValue = Nothing
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_dsReturnValue
        End Function

        Public Shared Function ExecuteSPDataAdapter(ByVal SPName As String, ByVal ParameterList As List(Of SqlParameter)) As SqlDataAdapter
            Dim daAdapater As SqlDataAdapter
            Try
                Dim m_cmdStoredProcedure As New SqlCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    daAdapater = New SqlDataAdapter(m_cmdStoredProcedure)
                    CloseDatabase()
                End If
            Catch exObj As Exception
                daAdapater = Nothing
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return daAdapater
        End Function
#End Region

#Region "ExecuteSPScalar"
        Public Shared Function ExecuteSPScalar(ByVal SPName As String, ByVal ParameterList As List(Of SqlParameter)) As Integer
            Dim m_intReturnValue As Integer = 0
            Try
                Dim m_cmdStoredProcedure As New SqlCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    m_intReturnValue = Convert.ToInt32(m_cmdStoredProcedure.ExecuteScalar())
                    CloseDatabase()
                End If
            Catch exObj As Exception
                m_intReturnValue = -1
                'HandleException.ExceptionLogging(exObj.Source, exObj.Message, True)
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return m_intReturnValue
        End Function
#End Region

#Region "ExecuteNonQuery OutPut"
        ''' <summary> 
        ''' This function shall Execute the Stored Procedure on the Database, this is a replica of using DOTNET ExecuteNonQuery 
        ''' This shall reduce lot of Development time in invoking the database properties. 
        ''' Input Parameters: String SPName -> Name of the Stored Procedures 
        ''' ParameterList -> List of Type SQLParameter 
        ''' The function is responsible for database connectivity and shall open and close the connection on it's own. 
        ''' </summary> 
        ''' <param name="SPName"></param> 
        ''' <param name="ParameterList"></param> 
        ''' <returns></returns> 
        Public Shared Function ExecuteSPNonQueryOutPut(ByVal SPName As String, ByVal ParameterList As List(Of SqlParameter), ByVal OutParameterList As List(Of SqlParameter), ByRef m_intReturnValue As Int16) As String()
            m_intReturnValue = 0
            Dim OutParameterArray As String() = New String(OutParameterList.Count - 1) {}
            Try
                Dim m_cmdStoredProcedure As New SqlCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next

                    For intLoop As Integer = 0 To OutParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(OutParameterList(intLoop))
                        OutParameterList(intLoop).Direction = ParameterDirection.Output
                    Next

                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()

                    CloseDatabase()

                    For intLoop As Integer = 0 To OutParameterList.Count - 1
                        OutParameterArray(intLoop) = m_cmdStoredProcedure.Parameters(OutParameterList(intLoop).ParameterName).Value.ToString()
                    Next
                End If
            Catch exObj As Exception
                m_intReturnValue = -1
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return OutParameterArray
        End Function
        Public Function ExecuteSPNonQueryOutPutCMDText(ByVal SPText As String, ByVal ParameterList As List(Of SqlParameter), ByVal OutParameterList As List(Of SqlParameter), ByRef m_intReturnValue As Int16) As String()
            m_intReturnValue = 0
            Dim OutParameterArray As String() = New String(OutParameterList.Count - 1) {}
            Dim m_cmdStoredProcedure As New SqlCommand()
            Try
                m_cmdStoredProcedure.CommandText = SPText
                'm_cmdStoredProcedure.CommandType = CommandType.Text
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next

                    For intLoop As Integer = 0 To OutParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(OutParameterList(intLoop))
                        OutParameterList(intLoop).Direction = ParameterDirection.Output
                    Next

                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()
                    'MsgBox(m_cmdStoredProcedure.Parameters.ToString)
                    CloseDatabase()

                    For intLoop As Integer = 0 To OutParameterList.Count - 1
                        OutParameterArray(intLoop) = m_cmdStoredProcedure.Parameters(OutParameterList(intLoop).ParameterName).Value.ToString()
                    Next
                End If
            Catch exObj As Exception
                'MsgBox(exObj.Message)
                m_intReturnValue = -1
                clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
                clsHandleException.LogINSERTQuery(ParameterList, System.Reflection.MethodBase.GetCurrentMethod.ToString())
                Throw exObj
            End Try
            Return OutParameterArray
        End Function

        Public Shared Function ExecuteSPNonQueryReturnValue(ByVal SPName As String, ByVal ParameterList As List(Of SqlParameter)) As Integer
            Dim m_intReturnValue As Integer = 0
            Dim intReturnValues As Integer = 0

            Try
                Dim m_cmdStoredProcedure As New SqlCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    Dim objReturnParameter As New SqlParameter("RETURNVALUE", SqlDbType.BigInt)
                    objReturnParameter.Direction = ParameterDirection.ReturnValue

                    m_cmdStoredProcedure.Parameters.Add(objReturnParameter)


                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()
                    intReturnValues = Convert.ToInt32(m_cmdStoredProcedure.Parameters("RETURNVALUE").Value)

                    CloseDatabase()
                End If
            Catch exObj As Exception
                intReturnValues = -1
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return intReturnValues + m_intReturnValue
        End Function

        Public Shared Function ExecuteNonQueryReturnValueWithoutAdd(ByVal SPName As String, ByVal ParameterList As List(Of SqlParameter)) As Integer
            Dim m_intReturnValue As Integer = 0
            Dim intReturnValues As Integer = 0

            Try
                Dim m_cmdStoredProcedure As New SqlCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    Dim objReturnParameter As New SqlParameter("RETURNVALUE", SqlDbType.BigInt)
                    objReturnParameter.Direction = ParameterDirection.ReturnValue

                    m_cmdStoredProcedure.Parameters.Add(objReturnParameter)


                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()
                    intReturnValues = Convert.ToInt32(m_cmdStoredProcedure.Parameters("RETURNVALUE").Value)

                    CloseDatabase()
                End If
            Catch exObj As Exception
                intReturnValues = -1
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return intReturnValues
        End Function


        Public Shared Function ExecuteNonQueryDecimal(ByVal SPName As String, ByVal ParameterList As List(Of SqlParameter)) As Decimal
            Dim m_intReturnValue As Integer = 0
            Dim intReturnValues As Decimal = 0

            Try
                Dim m_cmdStoredProcedure As New SqlCommand()
                m_cmdStoredProcedure.CommandText = SPName
                m_cmdStoredProcedure.CommandType = CommandType.StoredProcedure
                If m_SqlConnection Is Nothing Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection.State <> ConnectionState.Open Then
                    m_SqlConnection = OpenDatabase()
                End If
                If m_SqlConnection IsNot Nothing Then
                    m_cmdStoredProcedure.Connection = m_SqlConnection
                    For intLoop As Integer = 0 To ParameterList.Count - 1
                        m_cmdStoredProcedure.Parameters.Add(ParameterList(intLoop))
                    Next
                    Dim objReturnParameter As New SqlParameter("RETURNVALUE", SqlDbType.Decimal)
                    objReturnParameter.Direction = ParameterDirection.ReturnValue

                    m_cmdStoredProcedure.Parameters.Add(objReturnParameter)


                    m_intReturnValue = m_cmdStoredProcedure.ExecuteNonQuery()
                    intReturnValues = Convert.ToDecimal(m_cmdStoredProcedure.Parameters("RETURNVALUE").Value)

                    CloseDatabase()
                End If
            Catch exObj As Exception
                intReturnValues = -1
                Throw exObj  'clsHandleException.HandleEx(exObj, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return intReturnValues
        End Function
#End Region
    End Class
#End Region
    '**************END - TO EXECUTE STORED PROCEDURES*****************************************************
#End Region

    '**************START - TO HANDLE EXCEPTION*****************************************************
    Public Class clsHandleException

        Public Sub New()
            MsgBox(Err.Description)
        End Sub
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        Public Shared Sub HandleEx(ByVal ex As Exception, ByVal MethodName As String)
            strGlobalErrorInfo = strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
            strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
            strGlobalErrorInfo = String.Concat(strGlobalErrorInfo & vbCrLf, ex.Data)
            clsCOMMON.fnWrite2LOG(strGlobalErrorInfo, MethodName)
        End Sub

        Public Shared Sub HandleEx(ByVal exStr As String, ByVal MethodName As String)
            strGlobalErrorInfo = exStr
            clsCOMMON.fnWrite2LOG(strGlobalErrorInfo, MethodName)
        End Sub

        Public Shared Sub LogINSERTQuery(ByVal ParamList As List(Of SqlParameter), ByVal MethodName As String)
            Try
                Dim strLOG As String = ""
                For i As Int16 = 0 To ParamList.Count - 1
                    strLOG = strLOG & ParamList(i).ParameterName & " : " & ParamList(i).Value & vbCrLf
                Next
                clsCOMMON.fnWrite2LOG(strLOG, MethodName)
            Catch ex As Exception

            End Try
        End Sub
    End Class
    '**************END - TO HANDLE EXCEPTION*****************************************************


    '***********DATABASE CREDENTIALS FOR CRYSTAL REPORT***********************************
    Public Class DataBaseCredentials
        Public Shared ServerName As String
        Public Shared DatabaseName As String
        Public Shared UserName As String
        Public Shared Password As String
    End Class
    '***********END DATABASE CREDENTIALS FOR CRYSTAL REPORT*******************************

    '*********Validation Patterns*************************************************************************
#Region "Validate Patterns"
    Public Class clsPatterns
        Public Const MAILPattern As String = "\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
        Public Const USPHONEPattern As String = "^[01]?[- .]?(\([2-9]\d{2}\)|[2-9]\d{2})[- .]?\d{3}[- .]?\d{4}$"
        Public Const USFAXPattern As String = "^[01]?[- .]?(\([2-9]\d{2}\)|[2-9]\d{2})[- .]?\d{3}[- .]?\d{4}$"
        Public Const USZIPPattern As String = "^(\d{5}-\d{4}|\d{5}|\d{9})$|^([a-zA-Z]\d[a-zA-Z] \d[a-zA-Z]\d)$"
        Public Const INTEGERPattern As String = "^\d+(10)?$"
        'Public Const DECIMALPattern As String = "^\d+(\.\d\d)?$"
        Public Const DATEPattern As String = "^[0-2]?[1-9](/|-)[0-3]?[0-9](/|-)[1-2][0-9][0-9][0-9]$"
        Public Const DECIMALPattern As String = "^\d*[0-9](|.\d*[0-9]|,\d*[0-9])?$"
        'Public Const WEBSTITEPattern As String = "(http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?"
        Public Const WEBSTITEPattern As String = "[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?"
        Public Const TimePattern As String = "^([0]?[1-9]|[1][0-2])[:](0[0-9]|[1-5][0-9])[:]([A][M]))$"
        Public Const TimePattern24 As String = "^(20|21|22|23|[01]\d|\d)(([:][0-5]\d){1,2})$"
        Public Const StrongPasswordPattern As String = "(?=.{8,})[a-zA-Z]+[^a-zA-Z]+|[^a-zA-Z]+[a-zA-Z]+"
    End Class

    Public Class clsPatternTypes
        Public Const EMAILType As String = "EMAIL"
        Public Const FAXType As String = "FAX"
        Public Const PHONEType As String = "PHONE"
        Public Const ZIPType As String = "ZIP"
        Public Const INTEGERType As String = "INTEGER"
        Public Const DECIMALType As String = "DECIMAL"
        Public Const DATEType As String = "DATE"
        Public Const TimeType As String = "Time"
        Public Const StrongPasswordType As String = "STRONGPASSWORD"
    End Class

    ''' <summary>
    ''' CHECKs FOR STRONG PASSWORD
    ''' </summary>
    ''' <param name="strPassword"></param>
    ''' <param name="IsContainCaps"></param>
    ''' <param name="IsContainSmall"></param>
    ''' <param name="IsContainNumber"></param>
    ''' <param name="IsContainSpecialChars"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isPasswordStrong(ByVal strPassword As String, Optional ByVal IsContainCaps As Boolean = True, Optional ByVal IsContainSmall As Boolean = True, Optional ByVal IsContainNumber As Boolean = True, Optional ByVal IsContainSpecialChars As Boolean = False) As Boolean
        Dim isStrong As Boolean = False
        Try
            'CHECK FOR DOUBLEQUOTES,SPECES
            If strPassword.LastIndexOfAny("""") > -1 Then Exit Function
            If strPassword.LastIndexOfAny(" ") > -1 Then Exit Function

            'CHECK FOR PASSWORD LENGTH
            If Len(strPassword) >= 8 Then
                isStrong = True
            End If
            If isStrong = False Then Exit Function

            'CHECK FOR A CAPS LETTER
            If IsContainCaps Then
                isStrong = False
                For i As Int16 = 0 To strPassword.Length - 1
                    If Asc(strPassword(i)) >= 65 And Asc(strPassword(i)) <= 90 Then
                        isStrong = True
                        Exit For
                    End If
                Next
                If isStrong = False Then Exit Function
            End If

            'CHECK FOR A SMALL LETTER
            If IsContainSmall Then
                isStrong = False
                For i As Int16 = 0 To strPassword.Length - 1
                    If Asc(strPassword(i)) >= 97 And Asc(strPassword(i)) <= 122 Then
                        isStrong = True
                        Exit For
                    End If
                Next
                If isStrong = False Then Exit Function
            End If

            'CHECK FOR A NUMBER
            If IsContainNumber Then
                isStrong = False
                For i As Int16 = 0 To strPassword.Length - 1
                    If Asc(strPassword(i)) >= 48 And Asc(strPassword(i)) <= 57 Then
                        isStrong = True
                        Exit For
                    End If
                Next
                If isStrong = False Then Exit Function
            End If

            'CHECK FOR A SPECIAL CHAR
            If IsContainSpecialChars Then
                isStrong = False
                If strPassword.LastIndexOfAny("~`!@#$%^&*()+-=[]\;',./{}|:<>?") > -1 Then
                    isStrong = True
                End If
            Else
                isStrong = True
                If strPassword.LastIndexOfAny("~`!@#$%^&*()+-=[]\;',./{}|:<>?") > -1 Then
                    isStrong = False
                End If
            End If
        Catch ex As Exception

        End Try
        Return isStrong
    End Function

    ''' <summary>
    ''' FUNCTION TO MATCH VALID PATTERNS USING REGEX
    ''' </summary>
    ''' <param name="toMachString"></param>
    ''' <param name="patternString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fnValidatePatterns(ByVal toMachString As String, ByVal patternString As String) As Boolean
        Dim pattern As String = patternString
        Dim Test As New Regex(pattern)
        Dim valid As Boolean = False
        valid = Test.IsMatch(toMachString, 0)
        Return valid
    End Function

    ''' <summary>
    ''' FUNCTION TO MATCH VALID PATTERNS 
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <param name="strPattern"></param>
    ''' <param name="Type"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fnValidatePatterns(ByRef txt As TextBox, ByVal strPattern As String, ByVal Type As String) As Boolean
        'Dim pattern As New Regex(Validation)
        Dim pattern As New Regex(strPattern)
        Dim patternMatch As Match = pattern.Match(txt.Text)

        If Len(txt.Text) = 0 Then
            Return True
        End If
        If Not patternMatch.Success Then
            If UCase(Type) = "FAX" Then
                MsgBox("Invalid Fax Number!")
            ElseIf UCase(Type) = "EMAIL" Then
                MsgBox("Invalid Email ID!")
            ElseIf UCase(Type) = "ZIP" Then
                MsgBox("Invalid Zipcode!")
            ElseIf UCase(Type) = "PHONE" Then
                MsgBox("Invalid Phone Number!")
            End If
            txt.Focus()
            Return False
        Else
            Return True
        End If
    End Function
#End Region
    '*********End Validation Patterns**********************************************************************





    '**************TO STORE QUERIES TO BE PROCESSED LATER*****************************************************
    Public Class clsProcessQueries
        Private Shared strQueries() As String
        Private Shared intNoofQueries As Int16 = 0
        Private Shared Trans As OleDbTransaction
        Private Shared TransNew As OleDbTransaction

        Public Shared Sub AddNewQuery(ByVal NewQuery As String)
            Try
                If Not strQueries Is Nothing Then
                    Dim Found As Boolean
                    'For Each Item As String In strQueries
                    For i As Int16 = 0 To strQueries.GetUpperBound(0)
                        If Mid(strQueries(i), 1, 25) = Mid(NewQuery, 1, 25) Then
                            strQueries(i) = NewQuery
                            Found = True
                            Exit For
                        End If
                    Next

                    If Not Found Then
                        intNoofQueries += 1
                        ReDim Preserve strQueries(intNoofQueries)
                        strQueries(intNoofQueries) = NewQuery
                    End If
                Else
                    intNoofQueries += 1
                    ReDim Preserve strQueries(intNoofQueries)
                    strQueries(intNoofQueries) = NewQuery
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
        Public Shared Function AddNewQuery(ByVal NewQuery As String, ByVal isMatchedFully As Boolean) As Boolean
            Try
                If Not strQueries Is Nothing Then
                    If NewQuery.Length = 0 Then Exit Function
                    Dim Found As Boolean
                    'For Each Item As String In strQueries
                    For i As Int16 = 0 To strQueries.GetUpperBound(0)
                        If isMatchedFully = False Then
                            If Mid(strQueries(i), 1, 25) = Mid(NewQuery, 1, 25) Then
                                strQueries(i) = NewQuery
                                Found = True
                                AddNewQuery = False
                                Exit For
                            End If
                        Else
                            If strQueries(i) = NewQuery Then
                                strQueries(i) = NewQuery
                                Found = True
                                AddNewQuery = False
                                Exit For
                            End If
                        End If
                    Next

                    If Not Found Then
                        intNoofQueries += 1
                        ReDim Preserve strQueries(intNoofQueries)
                        strQueries(intNoofQueries) = NewQuery
                        AddNewQuery = True
                    End If
                Else
                    intNoofQueries += 1
                    ReDim Preserve strQueries(intNoofQueries)
                    strQueries(intNoofQueries) = NewQuery
                    AddNewQuery = True
                End If
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Sub AddNewQueryAtLast(ByVal NewQuery As String)
            Try
                If strQueries Is Nothing Then
                    ReDim Preserve strQueries(0)
                End If
                strQueries(0) = NewQuery
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        ''' <summary>
        ''' Process All Queries With its Own Transaction
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ProcessAllQueries() As Boolean
            Try
                If strQueries Is Nothing Then
                    Exit Function
                End If
                Dim i As Int16
                Trans = MyCon.BeginTransaction(IsolationLevel.Chaos)
                'Trans.Begin()
                For i = intNoofQueries To 0 Step -1
                    If Len(strQueries(i)) > 0 Then
                        If MyCLS.clsCOMMON.fnQueryExecuter(strQueries(i), Trans) = 0 Then
                            Trans.Rollback()
                            Return False
                        End If
                        'Else
                        '    Trans.Commit()
                        'MyCLS.clsCOMMON.fnQueryExecuter(strQueries(i), Trans)
                    End If
                Next
                intNoofQueries = 0
                Trans.Commit()
                Return True
            Catch ex As Exception
                Trans.Rollback()
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

#Region "Process Queries with New Transaction"
        ''' <summary>
        ''' Begin Transaction to be used by ProcessAllQueries(Transaction) as Boolean
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub BeginTransaction()
            Try
                TransNew = MyCon.BeginTransaction(IsolationLevel.Chaos)
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        ''' <summary>
        ''' Commit Transaction to be used by ProcessAllQueries(Transaction) as Boolean
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CommitTransaction()
            Try
                TransNew.Commit()
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        ''' <summary>
        ''' Commit Transaction to be used by ProcessAllQueries(Transaction) as Boolean
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub RollbackTransaction()
            Try
                TransNew.Rollback()
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        ''' <summary>
        ''' Process All Queries With New Transaction Begin Seperately With BeginTransaction
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ProcessRestQueries() As Boolean
            Try
                If strQueries Is Nothing Then Exit Function

                Dim i As Int16

                For i = intNoofQueries To 0 Step -1
                    If Len(strQueries(i)) > 0 Then
                        If InStr(strQueries(i), "@@***@@") = 0 Then
                            If MyCLS.clsCOMMON.fnQueryExecuter(strQueries(i), TransNew) = 0 Then
                                Return False
                            End If
                        End If
                    End If
                Next
                intNoofQueries = 0
                Return True
            Catch ex As Exception
                Return False
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        ''' <summary>
        ''' Process Query at Specified Index(Starts with 1) With New Transaction Begin Seperately With BeginTransaction
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ProcessQueryAt(ByVal Index As Int16) As Boolean
            Try
                If strQueries Is Nothing Then Exit Function

                Index = strQueries.Length - Index

                If Len(strQueries(Index)) > 0 Then
                    If MyCLS.clsCOMMON.fnQueryExecuter(strQueries(Index), TransNew) = 0 Then
                        Return False
                    Else
                        strQueries(Index) = "@@***@@" & strQueries(Index)
                    End If
                End If
                Return True
            Catch ex As Exception
                Return False
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        ''' <summary>
        ''' Process Single Query With New Transaction Begin Seperately With BeginTransaction
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ProcessSingleQuery(ByVal Qry As String) As Boolean
            Try
                If Len(Qry) > 0 Then
                    If MyCLS.clsCOMMON.fnQueryExecuter(Qry, TransNew) = 0 Then
                        Return False
                    Else
                        Return True
                    End If
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        ''' <summary>
        ''' Select Value From Query With New Transaction Begin Seperately With BeginTransaction
        ''' </summary>
        ''' <param name="SelectQ"></param>
        ''' <param name="ReturnType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function fnQuerySelect1Value(ByVal SelectQ As String, ByVal ReturnType As String) As String
            Try
                If Not MyCon.State = ConnectionState.Open Then
                    MyCon.Open()
                End If
                strGlobalErrorInfo = ""
                MyCmd = New OleDbCommand(SelectQ, MyCon, TransNew)
                MyRs = MyCmd.ExecuteReader
                MyRs.Read()
                If MyRs.HasRows Then
                    fnQuerySelect1Value = MyRs(0).ToString
                Else
                    If ReturnType = "String" Then
                        fnQuerySelect1Value = ""
                    Else
                        fnQuerySelect1Value = 0
                    End If
                End If
            Catch ex As Exception
                'strGlobalErrorInfo = "Qurery is : " & SelectQ
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function


        Public Shared Sub ReplaceTextInQueries(ByVal OldValue As String, ByVal NewValue As String)
            Try
                If strQueries Is Nothing Then Exit Sub

                Dim i As Int16
                For i = intNoofQueries To 0 Step -1
                    If Len(strQueries(i)) > 0 Then
                        strQueries(i) = Replace(strQueries(i), OldValue, NewValue)
                    End If
                Next
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
#End Region

        Public Shared Sub ClearAllQueries()
            Try
                If strQueries Is Nothing Then
                    Exit Sub
                End If
                For i As Int16 = 0 To strQueries.GetUpperBound(0)
                    strQueries(i) = ""
                Next
                ReDim strQueries(0)
                intNoofQueries = 0
            Catch ex As Exception
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Shared Function GetAllQueries() As String()
            GetAllQueries = strQueries
        End Function
        Public Shared Sub SetAllQueries(ByVal strQ() As String)
            strQueries = strQ
        End Sub
    End Class
    '**************END - TO STORE QUERIES TO BE PROCESSED LATER*****************************************************


    '**************TO UPDATE DATABASE USIGN DATASET*****************************************************
    ' TO BE USED IN DESKTOP APPLICATION

    Public Class clsUpdateWithAdapter
        Shared oCon As New SqlConnection(strConnStringSQLCLIENT)

        Public Shared Sub ConOpen()
            If oCon.State <> ConnectionState.Open Then
                oCon.Open()
            End If
        End Sub
        Public Shared Sub ConClose()
            If oCon.State <> ConnectionState.Closed Then
                oCon.Close()
            End If
        End Sub

        Public Shared Function InitDataset(ByRef oAdapter As SqlDataAdapter, ByRef oDataset As DataSet, ByRef oDataGridView As DataGridView, ByVal oQry As String, ByVal oTableName As String) As Boolean
            If oCon.State <> ConnectionState.Open Then
                oCon.Open()
            End If
            Try
                oAdapter = New SqlDataAdapter(oQry, oCon)
                Dim myDataRowsCommandBuilder As New SqlCommandBuilder(oAdapter)

                'oAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                oDataset.Clear()
                oAdapter.Fill(oDataset, oTableName)
                oDataGridView.DataSource = oDataset
                oDataGridView.DataMember = oTableName
                'oDataGridView.Columns(0).Visible = False
                If oCon.State <> ConnectionState.Closed Then
                    oCon.Close()
                End If
                Return True
            Catch ex As Exception
                MsgBox(ex.Message)
                'MyCLS.strGlobalErrorInfo = MyCLS.strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
                'MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
                'MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.Data)
                'MyCLS.clsCOMMON.fnWrite2LOG(MyCLS.strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
                If oCon.State <> ConnectionState.Closed Then
                    oCon.Close()
                End If
                Return False
            End Try
        End Function

        Public Shared Function CheckDuplicateInDataset(ByRef oDataset As DataSet, ByVal oColNumber As Int16, ByVal oValue As String) As Boolean
            Dim hasValue As Boolean
            For Each r As DataRow In oDataset.Tables(0).Rows
                For Each item As String In r(oColNumber).ToString
                    If item = oValue Then
                        hasValue = True
                        Exit For
                    End If
                Next
                If hasValue Then
                    Exit For
                End If
            Next
            Return hasValue
        End Function

        Public Shared Function UpdateDataset(ByRef oAdapter As SqlDataAdapter, ByRef oDataset As DataSet, ByVal oTableName As String) As Boolean
            'If oCon.State <> ConnectionState.Open Then
            '    oCon.Open()
            'End If
            Try
                MyCLS.clsCOMMON.ConClose()
            Catch ex As Exception
                MsgBox("Connection Could not Close")
            End Try
            'Try                
            '    oCon.Close()
            'Catch ex As Exception
            '    MsgBox("Connection Could not Close")
            'End Try

            'Try
            '    oCon.Open()
            'Catch ex As Exception
            '    MsgBox("Connection Could not Open")
            'End Try

            '--------------'Dim myDataRowsCommandBuilder As New SqlCommandBuilder(oAdapter)
            '--------------'oAdapter.InsertCommand = myDataRowsCommandBuilder.GetUpdateCommand()

            oAdapter.Update(oDataset, oTableName)


            Try
                MyCLS.clsCOMMON.ConOpen(False)
            Catch ex As Exception
                MsgBox("Connection Could not Open")
            End Try

            'If oCon.State <> ConnectionState.Closed Then
            '    oCon.Close()
            'End If
        End Function
    End Class
    '**************END - TO UPDATE DATABASE USIGN DATASET*****************************************************


    '*****START - CLASS FOR WINDOWS FUNCTION LIKE - SERVICES,REGISTRY,APIs ETC.*************
    Public Class clsWindows

        '*******START - API AREA*************
        Public Class clsWinAPI
            Private Const SW_HIDE As Integer = 0
            Private Const SW_RESTORE As Integer = 9
            Shared hWnd As Integer

            <DllImport("User32")> _
            Private Shared Function ShowWindow(ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
            End Function

            Public Shared Sub HideApplication(ByVal AppName As String)
                Dim p As Process() = Process.GetProcessesByName(AppName)
                hWnd = CType(p(0).MainWindowHandle, Integer)

                ShowWindow(hWnd, SW_HIDE)
            End Sub

            Public Shared Sub RestoreApplication(ByVal AppName As String)
                Dim p As Process() = Process.GetProcessesByName(AppName)
                hWnd = CType(p(0).MainWindowHandle, Integer)

                ShowWindow(hWnd, SW_RESTORE)
            End Sub


            Const HWND_TOPMOST = -1
            Const SWP_NOMOVE = &H2
            Const SWP_NOSIZE = &H1
            Private Const HWND_NOTOPMOST = -2
            Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
            Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
            Public Shared Sub MakeTopMost(ByVal hWnd As Object)
                SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
            End Sub




            Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
            Const SW_SHOWNORMAL = 1
            Public Shared Sub Execute(ByVal hWnd As Object, ByVal FilePath As String)
                ShellExecute(hWnd, vbNullString, FilePath, vbNullString, vbNullString, SW_SHOWNORMAL)
            End Sub



            Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Object, ByVal fuWinIni As Long) As Long
            Public Shared Sub Disablekeys(ByVal Key As Long, ByVal disable As Boolean)
                SystemParametersInfo(Key, disable, CStr(1), 0)
            End Sub


            '#Region "CreateParams"
            '            ' Thank you goes to uavfun for this awesome finding to disable the alt tab form. 
            '            Private Shared WS_POPUP As UInteger = &H80000000
            '            Private Shared WS_EX_TOPMOST As UInteger = &H8
            '            Private Shared WS_EX_TOOLWINDOW As UInteger = &H80
            '            Private Shared WS_MINIMIZE As UInteger = &H20000000

            '            Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
            '                Get
            '                    Dim cp As CreateParams = MyBase.CreateParams
            '                    cp.Style = CInt((WS_POPUP Or WS_MINIMIZE))
            '                    cp.ExStyle = CInt(WS_EX_TOPMOST) + CInt(WS_EX_TOOLWINDOW)

            '                    ' Set location 
            '                    cp.X = 100
            '                    cp.Y = 100

            '                    Return cp
            '                End Get
            '            End Property
            '#End Region

#Region "Print Screen the Desired Object/Control"
            Private Declare Function BitBlt Lib "GDI32" (ByVal srchDC As Integer, ByVal srcX As Integer, ByVal srcY As Integer, ByVal srcW As Integer, ByVal srcH As Integer, ByVal desthDC As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal op As Integer) As Integer
            Const SRCCOPY As Integer = &HCC0020

            Private Function GetImageFrom(ByVal Obj As Object) As Bitmap
                ' Get this Object's Graphics object.
                Dim me_gr As Graphics = Obj.CreateGraphics

                ' Make a Bitmap to hold the image.
                Dim bm As New Bitmap(Obj.ClientSize.Width, Obj.ClientSize.Height, me_gr)
                Dim bm_gr As Graphics = me_gr.FromImage(bm)
                Dim bm_hdc As IntPtr = bm_gr.GetHdc

                ' Get the Object's hDC. We must do this after 
                ' creating the new Bitmap, which uses me_gr.
                Dim me_hdc As IntPtr = me_gr.GetHdc

                ' BitBlt the Object's image onto the Bitmap.
                BitBlt(bm_hdc, 0, 0, Obj.ClientSize.Width, Obj.ClientSize.Height, me_hdc, 0, 0, SRCCOPY)
                me_gr.ReleaseHdc(me_hdc)
                bm_gr.ReleaseHdc(bm_hdc)

                ' Return the result.
                Return bm
            End Function
#End Region
        End Class
        '*******END - API AREA*************

        Public Shared MyAppPath As String = "C:\Program Files\KenApp\ShellApp\KenApp.exe"

        Public Shared Sub Shutdown()
            System.Diagnostics.Process.Start("shutdown", "-s -f -t 00")
        End Sub

        Public Shared Sub RunWindowsShell()
            System.Diagnostics.Process.Start("Explorer.exe")
        End Sub

        Public Shared Sub GetCurrentShellValues()
            MsgBox(My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", ""))
            MsgBox(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", ""))
        End Sub

        Public Shared Sub ChangeShell2Windows()
            My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe")
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe")
        End Sub
        Public Shared Sub ChangeShell2MyApp()
            My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", MyAppPath)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", MyAppPath)
        End Sub

        Public Shared Sub DisableTaskManager()
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 1)
        End Sub
        Public Shared Sub EnableTaskManager()
            Dim regKey As RegistryKey
            regKey = Registry.CurrentUser
            regKey.DeleteSubKey("Software\Microsoft\Windows\CurrentVersion\Policies\System")
        End Sub

    End Class
    '*****START - CLASS FOR WINDOWS FUNCTION LIKE - SERVICES,REGISTRY,APIs ETC.*************

    '*******START - CLASS FOR WRITING TEXT FILES**************************
    ''' <summary>
    ''' CLASS TO HANDEL TEXT FILE OPERATIONS LIKE - OPEN,CLOSE AND WRITE
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsFileHandling
        Shared xWrite As System.IO.StreamWriter
        Shared xFile As System.IO.File
        Shared strOutputFilePath As String

        ''' <summary>
        ''' TO OPEN A FILE TO WRITE USING WRITEFILE FUNCTION
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub OpenFile(ByVal strFile As String)
            strOutputFilePath = strFile
            xWrite = xFile.CreateText(strOutputFilePath)
        End Sub

        ''' <summary>
        ''' TO WRITE A FILE. AFTER WRITING THE FILE USE CLOSEFILE!
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <remarks></remarks>
        Public Shared Sub WriteFile(ByVal Str As String)
            xWrite.WriteLine(Str)
        End Sub

        ''' <summary>
        ''' TO CLOSE THE FILE OPENED BY OPENFILE FUNCTION
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CloseFile(Optional ByVal RunFile As Boolean = False)
            xWrite.Close()
            xFile = Nothing
            System.Windows.Forms.Application.DoEvents()
            If RunFile Then
                Shell("Notepad.exe " & strOutputFilePath, AppWinStyle.MaximizedFocus)
            End If
        End Sub

        ''' <summary>
        ''' TO READ A FILE
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ReadFile(ByVal Str As String) As String
            Try
                Return xFile.ReadAllText(Str)
            Catch ex As Exception
                Return ""
            End Try
        End Function


        ''' <summary>
        ''' Create File-Folder List in Txt File
        ''' </summary>
        ''' <param name="DirLoc"></param>
        ''' <param name="FillTab"></param>
        ''' <remarks></remarks>
        Public Shared Sub prcCreateFileListInTxt(ByVal DirLoc As String, ByVal FillTab As Integer, Optional ByVal ShowFileSize As Boolean = False, Optional ByVal FileNameSplitter As String = "")
            On Error Resume Next
            Dim i As Integer
            Dim posSep As Integer
            Dim sDir As String
            Dim aDirs() As String
            Dim sFile As String
            Dim aFiles() As String
            Dim FileParts() As String  'USED TO SPLIT FILENAMEs
            Dim strString As String = ""

            aDirs = System.IO.Directory.GetDirectories(DirLoc)

            '//PBExtract.Maximum = IIf(aDirs.GetUpperBound(0) > 0, aDirs.GetUpperBound(0), 100)
            For i = 0 To aDirs.GetUpperBound(0)
                ' Get the position of the last separator in the current path.
                posSep = aDirs(i).LastIndexOf("\")
                ' Get the path of the source directory.
                sDir = aDirs(i).Substring((posSep + 1), aDirs(i).Length - (posSep + 1))
                'lblCPath.Text = aDirs(i)
                'lblDirName.Text = sDir
                'Debug.Print(Space(FillTab * 5) & sDir)
                'WriteFile(Space(FillTab * 5) & sDir)
                If FillTab = 0 Then
                    MyCLS.clsFileHandling.WriteFile(MyCLS.clsCOMMON.fnTABs(FillTab) & sDir)
                End If
                ' Since we are in recursive mode, copy the children also
                FillTab = FillTab + 1
                'If InStr(aDirs(i), "Chapter wise Pdf") = 0 Then
                prcCreateFileListInTxt(aDirs(i), FillTab, ShowFileSize, FileNameSplitter)
                'End If
                FillTab = FillTab - 1
            Next
            '***###OPEN IF LIST SPECIFIED FOLDERS ONLY###*************************
            'If InStr(DirLoc, "Chapter wise Pdf") > 0 Then
            ' Get the files from the current parent.
            aFiles = System.IO.Directory.GetFiles(DirLoc)

            ' Copy all files.
            For i = 0 To aFiles.GetUpperBound(0)
                ' Get the position of the trailing separator.
                posSep = aFiles(i).LastIndexOf("\")

                ' Get the full path of the source file.
                sFile = aFiles(i).Substring((posSep + 1), aFiles(i).Length - (posSep + 1))
                'lblFileName.Text = sFile
                'Debug.Print(Space(FillTab * 5) & sFile & " - (" & System.IO.File.ReadAllBytes(DirLoc & "\" & sFile).Length & ")")

                If Len(FileNameSplitter) = 0 Then
                    strString = MyCLS.clsCOMMON.fnTABs(FillTab) & sFile
                Else
                    '***###USED TO SPLIT FILENAMEs###*****************************
                    FileParts = Split(sFile, FileNameSplitter)
                    If FileParts.Length > 1 Then
                        strString = MyCLS.clsCOMMON.fnTABs(FillTab) & FileParts(0) & vbTab & FileParts(1)
                    Else
                        strString = MyCLS.clsCOMMON.fnTABs(FillTab) & FileParts(0)
                    End If
                    '***###USED TO SPLIT FILENAMEs###*****************************
                End If
                If ShowFileSize = True Then
                    strString = strString & vbTab & System.IO.File.ReadAllBytes(DirLoc & "\" & sFile).Length
                End If

                MyCLS.clsFileHandling.WriteFile(strString)

                System.Windows.Forms.Application.DoEvents()
            Next i
            'End If
            '***###OPEN IF LIST SPECIFIED FOLDERS ONLY###*************************
        End Sub
    End Class
    '*******START - CLASS FOR WRITING TEXT FILES**************************

    '***********START - CONVERT AMOUNT IN WORDS************************************
    Public Class clsConvertCurrencyInUS
        Dim mOnesArray(8) As String
        Dim mOneTensArray(9) As String
        Dim mTensArray(7) As String
        Dim mPlaceValues(4) As String

        Public Sub New()
            mOnesArray(0) = "One"
            mOnesArray(1) = "Two"
            mOnesArray(2) = "Three"
            mOnesArray(3) = "Four"
            mOnesArray(4) = "Five"
            mOnesArray(5) = "Six"
            mOnesArray(6) = "Seven"
            mOnesArray(7) = "Eight"
            mOnesArray(8) = "Nine"

            mOneTensArray(0) = "Ten"
            mOneTensArray(1) = "Eleven"
            mOneTensArray(2) = "Twelve"
            mOneTensArray(3) = "Thirteen"
            mOneTensArray(4) = "Fourteen"
            mOneTensArray(5) = "Fifteen"
            mOneTensArray(6) = "Sixteen"
            mOneTensArray(7) = "Seventeen"
            mOneTensArray(8) = "Eightteen"
            mOneTensArray(9) = "Nineteen"

            mTensArray(0) = "Twenty"
            mTensArray(1) = "Thirty"
            mTensArray(2) = "Forty"
            mTensArray(3) = "Fifty"
            mTensArray(4) = "Sixty"
            mTensArray(5) = "Seventy"
            mTensArray(6) = "Eighty"
            mTensArray(7) = "Ninety"

            mPlaceValues(0) = "Hundred"
            mPlaceValues(1) = "Thousand"
            mPlaceValues(2) = "Million"
            mPlaceValues(3) = "Billion"
            mPlaceValues(4) = "Trillion"

            'mPlaceValues(0) = "lac"
            'mPlaceValues(1) = "cr"
            'mPlaceValues(2) = "cr1"
            'mPlaceValues(3) = "cr2"
            'mPlaceValues(4) = "cr3"
        End Sub

        Protected Function GetOnes(ByVal OneDigit As Integer) As String
            GetOnes = ""

            If OneDigit = 0 Then
                Exit Function
            End If

            GetOnes = mOnesArray(OneDigit - 1)
        End Function

        Protected Function GetTens(ByVal TensDigit As Integer) As String
            GetTens = ""

            If TensDigit = 0 Or TensDigit = 1 Then
                Exit Function
            End If

            GetTens = mTensArray(TensDigit - 2)
        End Function

        Public Function ConvertNumberToWords(ByVal NumberValue As String) As String
            Dim Delimiter As String = " "
            Dim TensDelimiter As String = "-"
            Dim mNumberValue As String = ""
            Dim mNumbers As String = ""
            Dim mNumWord As String = ""
            Dim mFraction As String = ""
            Dim mNumberStack() As String
            Dim j As Integer = 0
            Dim i As Integer = 0
            Dim mOneTens As Boolean = False

            ConvertNumberToWords = ""

            ' validate input
            Try
                j = CDbl(NumberValue)
            Catch ex As Exception
                ConvertNumberToWords = "Invalid input."
                Exit Function
            End Try

            ' get fractional part {if any}
            If InStr(NumberValue, ".") = 0 Then
                ' no fraction
                mNumberValue = NumberValue
            Else
                mNumberValue = Microsoft.VisualBasic.Left(NumberValue, InStr(NumberValue, ".") - 1)
                mFraction = Mid(NumberValue, InStr(NumberValue, ".")) ' + 1)
                mFraction = Math.Round(CSng(mFraction), 2) * 100

                If CInt(mFraction) = 0 Then
                    mFraction = ""
                Else
                    mFraction = "&& " & mFraction & "/100"
                End If
            End If
            mNumbers = mNumberValue.ToCharArray

            ' move numbers to stack/array backwards
            For j = mNumbers.Length - 1 To 0 Step -1
                ReDim Preserve mNumberStack(i)

                mNumberStack(i) = mNumbers(j)
                i += 1
            Next

            For j = mNumbers.Length - 1 To 0 Step -1
                Select Case j
                    Case 0, 3, 6, 9, 12
                        ' ones  value
                        If Not mOneTens Then
                            mNumWord &= GetOnes(Val(mNumberStack(j))) & Delimiter
                        End If

                        Select Case j
                            Case 3
                                ' thousands
                                mNumWord &= mPlaceValues(1) & Delimiter

                            Case 6
                                ' millions
                                mNumWord &= mPlaceValues(2) & Delimiter

                            Case 9
                                ' billions
                                mNumWord &= mPlaceValues(3) & Delimiter

                            Case 12
                                ' trillions
                                mNumWord &= mPlaceValues(4) & Delimiter
                        End Select


                    Case Is = 1, 4, 7, 10, 13
                        ' tens value
                        If Val(mNumberStack(j)) = 0 Then
                            mNumWord &= GetOnes(Val(mNumberStack(j - 1))) & Delimiter
                            mOneTens = True
                            Exit Select
                        End If

                        If Val(mNumberStack(j)) = 1 Then
                            mNumWord &= mOneTensArray(Val(mNumberStack(j - 1))) & Delimiter
                            mOneTens = True
                            Exit Select
                        End If

                        mNumWord &= GetTens(Val(mNumberStack(j)))

                        ' this places the tensdelimiter; check for succeeding 0
                        If Val(mNumberStack(j - 1)) <> 0 Then
                            mNumWord &= TensDelimiter
                        End If
                        mOneTens = False

                    Case Else
                        ' hundreds value 
                        mNumWord &= GetOnes(Val(mNumberStack(j))) & Delimiter

                        If Val(mNumberStack(j)) <> 0 Then
                            mNumWord &= mPlaceValues(0) & Delimiter
                        End If
                End Select
            Next
            Return mNumWord & mFraction
        End Function
    End Class
    '*********************************************************************************************************************************
    '*********************************************************************************************************************************
    Public Class clsConvertCurrencyInRupees

        Public Shared Function ConvertNumberToWords(ByVal src_num As String) As String
            Dim SNUM As Double
            SNUM = Val(src_num)
            'SNUM = Convert.ChangeType(src_num, TypeCode.Double)
            If SNUM > 9999999999999.0# Then
                ConvertNumberToWords = "Error: To much number."
                Exit Function
            End If
            Dim WHOLE As String
            Dim EXTRA As String
            Dim WORD As String
            Dim WORDDECIMAL As String = ""
            Dim NWHOLE As Double

            If InStr(1, Str$(SNUM), ".", vbTextCompare) <> 0 Then
                WHOLE = Split(Str$(SNUM), ".")(0)
                EXTRA = Split(src_num, ".")(1)
            Else
                WHOLE = SNUM
                EXTRA = 0
            End If

            If SNUM < 1 Then WORD = "Zero"

            NWHOLE = Val(WHOLE)
            'Check for One and Tens
            If Val(Right(NWHOLE, 2)) > 0 And Val(Right(NWHOLE, 2)) < 21 Or Val(Right(NWHOLE, 2)) = 30 Or Val(Right(NWHOLE, 2)) = 40 Or Val(Right(NWHOLE, 2)) = 50 Or Val(Right(NWHOLE, 2)) = 60 Or Val(Right(NWHOLE, 2)) = 70 Or Val(Right(NWHOLE, 2)) = 80 Or Val(Right(NWHOLE, 2)) = 90 Then
                WORD = WORD & WordTens(Val(Right(NWHOLE, 2)))
            ElseIf Val(Right(NWHOLE, 2)) > 20 Then
                WORD = WORD & WordTens(Left(Right(NWHOLE, 2), 1) & "0")
                WORD = WORD & WordTens(Right(NWHOLE, 1))
            End If
            'Check for Hundred
            If NWHOLE > 99 Then
                If Left(Right(NWHOLE, 3), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 3), 1)) & " Hundred" & WORD
            End If
            'Check for Thousand
            If NWHOLE > 999 Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 3))) & " Thousand" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 3), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 2, 1)) & " Thousand" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 99 Then
                    '            If Left(Right(NWHOLE, 6), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 6), 1)) & " Hundred" & WORD
                    '        End If
                    If WORD = " Thousand" Then WORD = ""
                End If
            End If

            'Check for Lakh
            If NWHOLE > 99999 Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 5))) & " Lakh" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 5), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 5)), 2), 2, 1)) & " Lakh" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 5)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 99 Then
                    '            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
                    '        End If
                    If WORD = " Lakh" Then WORD = ""
                End If
            End If

            'Check for crore
            If NWHOLE > 9999999 Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 7))) & " Crore" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 7), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 7)), 2), 2, 1)) & " Crore" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 7)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 99 Then
                    '            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
                    '        End If
                    If WORD = " Crore" Then WORD = ""
                End If
            End If

            'Check for billion
            If NWHOLE > 999999999 Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 9))) & " Arab" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 9), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 2, 1)) & " Arab" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 99 Then
                    '            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
                    '        End If
                    If WORD = " Arab" Then WORD = ""
                End If
            End If

            'Check for trillion
            If NWHOLE > 99999999999.0# Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 11))) & " Kharab" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 11), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 11)), 2), 2, 1)) & " Kharab" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 11)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 99 Then
                    '            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
                    '        End If
                    If WORD = " Kharab" Then WORD = ""
                End If
            End If


            If EXTRA > 0 Then
                WORDDECIMAL = " And " & ConvertDecimalsToWords(src_num) & " Paisa Only"
            End If

            If Len(WORDDECIMAL) > 16 Then
                ConvertNumberToWords = WORD & WORDDECIMAL
            Else
                ConvertNumberToWords = WORD
            End If

            NWHOLE = 0
            WORD = ""
            EXTRA = ""
            WHOLE = ""
        End Function

        Private Shared Function ConvertDecimalsToWords(ByVal src_num As String) As String
            Dim SNUM As Double
            SNUM = Val(src_num)
            If SNUM > 999999999999999.0# Then
                ConvertDecimalsToWords = "Error: To much number."
                Exit Function
            End If
            Dim WHOLE As String
            Dim EXTRA As String
            Dim WORD As String
            Dim NWHOLE As Double

            If InStr(1, Str$(SNUM), ".", vbTextCompare) <> 0 Then
                EXTRA = Split(Str$(SNUM), ".")(0)
                WHOLE = Split(src_num, ".")(1)
            Else
                WHOLE = SNUM
            End If

            If SNUM < 1 Then WORD = "Zero"

            NWHOLE = Val(WHOLE)
            'Check for One and Tens
            If Val(Right(NWHOLE, 2)) > 0 And Val(Right(NWHOLE, 2)) < 21 Or Val(Right(NWHOLE, 2)) = 30 Or Val(Right(NWHOLE, 2)) = 40 Or Val(Right(NWHOLE, 2)) = 50 Or Val(Right(NWHOLE, 2)) = 60 Or Val(Right(NWHOLE, 2)) = 70 Or Val(Right(NWHOLE, 2)) = 80 Or Val(Right(NWHOLE, 2)) = 90 Then
                WORD = WORD & WordTens(Val(Right(NWHOLE, 2)))
            ElseIf Val(Right(NWHOLE, 2)) > 20 Then
                WORD = WORD & WordTens(Left(Right(NWHOLE, 2), 1) & "0")
                WORD = WORD & WordTens(Right(NWHOLE, 1))
            End If
            'Check for Hundred
            If NWHOLE > 99 Then
                If Left(Right(NWHOLE, 3), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 3), 1)) & " Hundred" & WORD
            End If
            'Check for Thousand
            If NWHOLE > 999 Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 3))) & " Thousand" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 3), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 2, 1)) & " Thousand" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 99 Then
                    '            If Left(Right(NWHOLE, 6), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 6), 1)) & " Hundred" & WORD
                    '        End If
                End If
            End If

            'Check for Lakh
            If NWHOLE > 99999 Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 5))) & " Lakh" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 5), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 5)), 2), 2, 1)) & " Lakh" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 5)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 99 Then
                    '            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
                    '        End If
                End If
            End If

            'Check for crore
            If NWHOLE > 9999999 Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 7))) & " crore" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 7)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 7), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 7)), 2), 2, 1)) & " crore" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 7)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 99 Then
                    '            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
                    '        End If
                End If
            End If

            'Check for billion
            If NWHOLE > 999999999 Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 9))) & " billion" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 9), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 2, 1)) & " billion" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 99 Then
                    '            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
                    '        End If
                End If
            End If

            'Check for trillion
            If NWHOLE > 99999999999.0# Then
                If Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) = 90 Then
                    WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 11))) & " trillion" & WORD
                ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 11)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 11), 3) <> "000" Then
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 11)), 2), 2, 1)) & " trillion" & WORD
                    WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 11)), 2), 1, 1) & "0") & WORD
                    '        If Val(Left(NWHOLE, Len("" & NWHOLE) - 5)) > 99 Then
                    '            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
                    '        End If
                End If
            End If


            ConvertDecimalsToWords = WORD

            NWHOLE = 0
            WORD = ""
            EXTRA = ""
            WHOLE = ""
        End Function

        Private Shared Function WordTens(ByVal SNUM As Long) As String
            Select Case SNUM
                Case 1
                    WordTens = " One"
                Case 2
                    WordTens = " Two"
                Case 3
                    WordTens = " Three"
                Case 4
                    WordTens = " Four"
                Case 5
                    WordTens = " Five"
                Case 6
                    WordTens = " Six"
                Case 7
                    WordTens = " Seven"
                Case 8
                    WordTens = " Eight"
                Case 9
                    WordTens = " Nine"
                Case 10
                    WordTens = " Ten"
                Case 11
                    WordTens = " Eleven"
                Case 12
                    WordTens = " Twelve"
                Case 13
                    WordTens = " Thirteen"
                Case 14
                    WordTens = " Fourteen"
                Case 15
                    WordTens = " Fifteen"
                Case 16
                    WordTens = " Sixteen"
                Case 17
                    WordTens = " Seventeen"
                Case 18
                    WordTens = " Eighteen"
                Case 19
                    WordTens = " Nineteen"
                Case 20
                    WordTens = " Twenty"
                Case 30
                    WordTens = " Thirty"
                Case 40
                    WordTens = " Fourty"
                Case 50
                    WordTens = " Fifty"
                Case 60
                    WordTens = " Sixty"
                Case 70
                    WordTens = " Seventy"
                Case 80
                    WordTens = " Eighty"
                Case 90
                    WordTens = " Ninety"
            End Select
        End Function

        Public Function cDecToWord(ByVal src_num As String) As String
            Dim SNUM As Double
            SNUM = Val(src_num)
            If SNUM > 999999999999999.0# Then
                cDecToWord = "Error: To much number."
                Exit Function
            End If
            Dim WHOLE As String
            Dim EXTRA As String
            Dim WORD As String
            Dim NWHOLE As Double

            If InStr(1, Str$(SNUM), ".", vbTextCompare) <> 0 Then
                WHOLE = Split(Str$(SNUM), ".")(0)
                EXTRA = Split(src_num, ".")(1)
            Else
                WHOLE = SNUM
            End If

            If SNUM < 1 Then WORD = "Zero"

            NWHOLE = Val(WHOLE)
            'Check for One and Tens
            If Val(Right(NWHOLE, 2)) > 0 And Val(Right(NWHOLE, 2)) < 21 Or Val(Right(NWHOLE, 2)) = 30 Or Val(Right(NWHOLE, 2)) = 40 Or Val(Right(NWHOLE, 2)) = 50 Or Val(Right(NWHOLE, 2)) = 60 Or Val(Right(NWHOLE, 2)) = 70 Or Val(Right(NWHOLE, 2)) = 80 Or Val(Right(NWHOLE, 2)) = 90 Then
                WORD = WORD & WordTens(Val(Right(NWHOLE, 2)))
            ElseIf Val(Right(NWHOLE, 2)) > 20 Then
                WORD = WORD & WordTens(Left(Right(NWHOLE, 2), 1) & "0")
                WORD = WORD & WordTens(Right(NWHOLE, 1))
            End If

            cDecToWord = WORD
        End Function
    End Class
    '***********END - CONVERT AMOUNT IN WORDS************************************

    '****************START - TO CHANGE NUMBER FORMAT WITH COMMAS***********
    Public Class clsUtility
        Public Shared Function ToDouble(ByVal originalValue As String, ByVal roundPlaces As Integer) As Double
            Dim returnValue As Double = 0
            Try
                originalValue = originalValue.Trim()
                If originalValue.Length < 1 Then
                    Return 0
                End If
                If roundPlaces < 100 Then
                    returnValue = System.Math.Round(System.[Double].Parse(originalValue), roundPlaces)
                Else
                    returnValue = System.[Double].Parse(originalValue)
                End If
            Catch
                Return 0
            End Try
            Return returnValue
        End Function

        Public Shared Function FormatEmptyNumber(ByVal decimalDelimiter As String, ByVal decimalPlaces As Integer) As String
            Dim preDecimal As String = "0"
            Dim postDecimal As String = ""
            For i As Integer = 0 To decimalPlaces - 1
                If i = 0 Then
                    postDecimal += decimalDelimiter
                End If
                postDecimal += "0"
            Next
            Return preDecimal + postDecimal
        End Function
        Public Shared Function FormatNumber(ByVal value As String, ByVal commaDelimiter As String, ByVal decimalDelimiter As String, ByVal decimalPlaces As Integer) As String
            Dim minus As String = ""
            Dim preDecimal As String = ""
            Dim postDecimal As String = ""
            Dim regex__1 As Regex = Nothing
            Dim returnValue As String = ""
            Dim pattern As String = "(-?[0-9]+)([0-9]{3})"
            'Dim pattern As String = "([0-9]{2})([0-9]{3})" 
            Try
                value = value.Trim()
                If decimalPlaces < 1 Then
                    decimalDelimiter = ""
                End If
                If value.LastIndexOf("-") = 0 Then
                    minus = "-"
                    value = value.Replace("-", "")
                End If
                preDecimal = value
                ' preDecimal doesn't contain a number at all. 
                ' Return formatted zero representation. 
                If preDecimal.Length < 1 Then
                    Return minus + FormatEmptyNumber(decimalDelimiter, decimalPlaces)
                End If
                ' preDecimal is 0 or a series of 0's. 
                ' Return formatted zero representation. 
                If commaDelimiter.Length > 0 Then
                    preDecimal = preDecimal.Replace(commaDelimiter, "")
                End If
                If decimalDelimiter.Length > 0 Then
                    preDecimal = preDecimal.Replace(decimalDelimiter, "")
                End If
                If ToDouble(preDecimal, 0) < 1 Then
                    Return minus + FormatEmptyNumber(decimalDelimiter, decimalPlaces)
                End If
                ' predecimal has no numbers to the left. 
                ' Return formatted zero representation. 
                If preDecimal.Length = decimalPlaces Then
                    Return (minus & "0") + decimalDelimiter + preDecimal
                End If
                ' predecimal has fewer characters than the 
                ' specified number of decimal places. 
                ' Return formatted leading zero representation. 
                If preDecimal.Length < decimalPlaces Then
                    If decimalPlaces = 2 Then
                        Return minus + FormatEmptyNumber(decimalDelimiter, decimalPlaces - 1) + preDecimal
                    End If
                    Return minus + FormatEmptyNumber(decimalDelimiter, decimalPlaces - 2) + preDecimal
                End If
                ' predecimal contains enough characters to 
                ' qualify to need decimal points rendered. 
                ' Parse out the pre and post decimal values 
                ' for future formatting. 
                If preDecimal.Length > decimalPlaces Then
                    postDecimal = decimalDelimiter + preDecimal.Substring(preDecimal.Length - decimalPlaces)
                    preDecimal = preDecimal.Substring(0, preDecimal.Length - decimalPlaces)
                End If
                ' Place comma oriented delimiter every 3 characters 
                ' against the numeric represenation of the "left" side 
                ' of the decimal representation. When finished, return 
                ' both the left side comma formatted value together with 
                ' the right side decimal formatted value. 
                regex__1 = New Regex(pattern)
                While Regex.IsMatch(preDecimal, pattern)
                    preDecimal = regex__1.Replace(preDecimal, "$1" & commaDelimiter & "$2")
                End While
            Catch generatedExceptionName As Exception
                returnValue = ToDouble("0", 0).ToString()
            End Try
            Return minus + preDecimal + postDecimal
        End Function
    End Class
    '****************END - TO CHANGE NUMBER FORMAT WITH COMMAS***********



#Region "Property Creation"
    '******START - CLASS CONTAINS TABLES AND IT'S PROPERTIES**************************************
    Public Class clsTables
        Public TABLENAME As String
        Public TABLESCHEMA As String
        Public COLUMNDETAILS As List(Of clsColumns)

    End Class

    Public Class clsColumns
        Public COLUMNNAME As String
        Public COLUMNTYPE As String
        Public COLName As String
        Public COLSize As String
        Public COLSizeChar As String
        Public COLNumericPrecision As String
        Public COLNumericScale As String
        Public COLDataType As String
        Public COLDataTypeSQL As String
        Public COLProviderType As String
        Public COLIsAutoIncrement As String

    End Class
    '******END - CLASS CONTAINS TABLES AND IT'S PROPERTIES**************************************

    '*************START - DATABASE OPERATIONS***************
    Public Class clsDBOperations
        'TO STORE ALL THE DB TYPES IN DATABASE
        Public Shared strDBTypeValueALL As String = ""

        ''' <summary>
        ''' TO GET THE LIST OF ALL TABLES WITH IN A DATABASE
        ''' </summary>
        ''' <returns>SINGLE DIMENTION ARRAY OF STRING</returns>
        ''' <remarks>CAN BE USED DIRECTLY</remarks>
        Public Shared Function GetTables() As String()
            Dim strTables(1) As String
            Try
                Dim dt As DataTable = MyCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
                ReDim strTables(dt.Rows.Count - 1)
                For i As Integer = 0 To dt.Rows.Count - 1
                    strTables(i) = dt.Rows(i)("TABLE_NAME").ToString()
                Next
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return strTables
        End Function

        ''' <summary>
        ''' TO GET THE LIST OF COLUMN NAMES AND COLUMN DATA TYPE WITH IN A TABLE
        ''' </summary>
        ''' <param name="Table"></param>
        ''' <returns>TWO DIMENTION ARRAY OF STRING</returns>
        ''' <remarks>CAN BE USED DIRECTLY</remarks>
        Public Shared Function GetColumns(ByVal Table As String) As String(,)
            Dim strColumns(1, 1) As String
            Try
                Dim dr As OleDbDataReader
                dr = MyCLS.clsCOMMON.fnQuerySelectRS("Select * From " & Table & " Where 1>2")
                If dr Is Nothing Then
                    dr = MyCLS.clsCOMMON.fnQuerySelectRS("Select * From " & Table & "")
                End If

                ReDim strColumns(dr.FieldCount - 1, 1)
                For i As Integer = 0 To dr.FieldCount - 1
                    strColumns(i, 0) = dr.GetName(i).ToString
                    strColumns(i, 1) = GetDBTypeValue(dr.GetDataTypeName(i).ToString)
                    
                    System.Windows.Forms.Application.DoEvents()
                Next
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return strColumns
        End Function
        Public Shared Function GetColumns_UPDATED(ByVal Table As String, ByVal Schema As String) As String(,)
            Dim strColumns(1, 1) As String
            Try
                Dim drLIB As OleDbDataReader    'FOR LIB PROPERTY
                Dim dsSQL As New DataSet            'FOR SQL SP
                If Schema.Length > 0 Then
                    drLIB = MyCLS.clsCOMMON.fnQuerySelectRS("Select * From [" & Schema & "].[" & Table & "] Where 1>2")
                    If drLIB Is Nothing Then
                        drLIB = MyCLS.clsCOMMON.fnQuerySelectRS("Select * From [" & Schema & "].[" & Table & "]")
                    End If
                Else
                    drLIB = MyCLS.clsCOMMON.fnQuerySelectRS("Select * From [" & Table & "] Where 1>2")
                    If drLIB Is Nothing Then
                        drLIB = MyCLS.clsCOMMON.fnQuerySelectRS("Select * From [" & Table & "]")
                    End If
                End If
                '***CHANGED FOR MVC API***
                'MyCLS.clsCOMMON.prcQuerySelectDS(dsSQL, "SELECT ORDINAL_POSITION,COLUMN_NAME,DATA_TYPE,ISNULL(CHARACTER_MAXIMUM_LENGTH,'') AS CHARACTER_MAXIMUM_LENGTH,ISNULL(NUMERIC_PRECISION,'') AS NUMERIC_PRECISION,ISNULL(NUMERIC_SCALE,'') AS NUMERIC_SCALE FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & Table & "' ORDER BY ORDINAL_POSITION", Table)
                MyCLS.clsCOMMON.prcQuerySelectDS(dsSQL, "SELECT DISTINCT C.ORDINAL_POSITION,C.COLUMN_NAME,DATA_TYPE,ISNULL(CHARACTER_MAXIMUM_LENGTH,'') AS CHARACTER_MAXIMUM_LENGTH,ISNULL(NUMERIC_PRECISION,'') AS NUMERIC_PRECISION,ISNULL(NUMERIC_SCALE,'') AS NUMERIC_SCALE,isNULL(IS_NULLABLE,'') as IS_NULLABLE,TC.CONSTRAINT_TYPE FROM  INFORMATION_SCHEMA.COLUMNS C LEFT OUTER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE CU ON CU.COLUMN_NAME=C.COLUMN_NAME LEFT OUTER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS TC ON TC.CONSTRAINT_NAME=CU.CONSTRAINT_NAME WHERE C.TABLE_NAME = '" & Table & "' ORDER BY C.ORDINAL_POSITION", Table)

                '***TO GET MORE DETAILS FOR COLUMNS***
                Dim dtCols As DataTable = drLIB.GetSchemaTable()

                ReDim strColumns(drLIB.FieldCount - 1, 11)
                For i As Integer = 0 To drLIB.FieldCount - 1
                    strColumns(i, 0) = drLIB.GetName(i).ToString
                    strColumns(i, 1) = GetDBTypeValue(drLIB.GetDataTypeName(i).ToString)

                    strColumns(i, 2) = dtCols.Rows(i).Item("ColumnName").ToString
                    strColumns(i, 3) = dtCols.Rows(i).Item("ColumnSize").ToString
                    strColumns(i, 4) = dsSQL.Tables(0).Rows(i)("NUMERIC_PRECISION").ToString
                    strColumns(i, 5) = dsSQL.Tables(0).Rows(i)("NUMERIC_SCALE").ToString
                    strColumns(i, 6) = dtCols.Rows(i).Item("DataType").ToString.Replace("System.", "")
                    strColumns(i, 7) = dtCols.Rows(i).Item("ProviderType").ToString
                    strColumns(i, 8) = dtCols.Rows(i).Item("IsAutoIncrement").ToString
                    strColumns(i, 9) = dsSQL.Tables(0).Rows(i)("CHARACTER_MAXIMUM_LENGTH").ToString
                    strColumns(i, 10) = dsSQL.Tables(0).Rows(i)("DATA_TYPE").ToString
                    strColumns(i, 11) = dsSQL.Tables(0).Rows(i)("IS_NULLABLE").ToString

                    System.Windows.Forms.Application.DoEvents()
                Next
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return strColumns
        End Function
        'Public Shared Function GetColumns_UPDATED(ByVal Table As String) As String(,)
        '    Dim strColumns(1, 1) As String
        '    Try
        '        Dim dr As OleDbDataReader
        '        dr = MyCLS.clsCOMMON.fnQuerySelectRS("Select * From [" & Table & "] Where 1>2")
        '        If dr Is Nothing Then
        '            dr = MyCLS.clsCOMMON.fnQuerySelectRS("Select * From [" & Table & "]")
        '        End If

        '        '***TO GET MORE DETAILS FOR COLUMNS***
        '        Dim dtCols As DataTable = dr.GetSchemaTable()

        '        ReDim strColumns(dr.FieldCount - 1, 8)
        '        For i As Integer = 0 To dr.FieldCount - 1
        '            strColumns(i, 0) = dr.GetName(i).ToString
        '            strColumns(i, 1) = GetDBTypeValue(dr.GetDataTypeName(i).ToString)

        '            strColumns(i, 2) = dtCols.Rows(i).Item("ColumnName").ToString
        '            strColumns(i, 3) = dtCols.Rows(i).Item("ColumnSize").ToString
        '            strColumns(i, 4) = dtCols.Rows(i).Item("NumericPrecision").ToString
        '            strColumns(i, 5) = dtCols.Rows(i).Item("NumericScale").ToString
        '            strColumns(i, 6) = dtCols.Rows(i).Item("DataType").ToString.Replace("System.", "")
        '            strColumns(i, 7) = dtCols.Rows(i).Item("ProviderType").ToString
        '            strColumns(i, 8) = dtCols.Rows(i).Item("IsAutoIncrement").ToString

        '            System.Windows.Forms.Application.DoEvents()
        '        Next
        '    Catch ex As Exception
        '        MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
        '    End Try
        '    Return strColumns
        'End Function

        ''' <summary>
        ''' TO SET CLSTABLES CLASS WITH TABLE NAMES, COLUMN NAMES AND COLUMN DATA TYPES
        ''' </summary>
        ''' <returns>LIST OF clsTables CLASS WITH DETAILS</returns>
        ''' <remarks>USE THIS FUNCTION ONLY TO CAPTURE FULL DETAILS OF DATABASE</remarks>
        Public Shared Function FillDetails() As List(Of clsTables)
            Dim objclsTablesListing As New List(Of clsTables)

            Try
                Dim dt As DataTable = MyCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
                For i As Integer = 0 To dt.Rows.Count - 1
                    Dim objclsTables As New clsTables
                    objclsTables.TABLENAME = dt.Rows(i)("TABLE_NAME").ToString()
                    objclsTables.TABLESCHEMA = dt.Rows(i)("TABLE_SCHEMA").ToString()

                    Dim objclsColumnsListing As New List(Of clsColumns)

                    Dim strColumns(,) As String = GetColumns_UPDATED(objclsTables.TABLENAME, objclsTables.TABLESCHEMA)
                    'MsgBox(MyCLS.clsDBOperations.strDBTypeValueALL)
                    For j As Int16 = 0 To strColumns.GetLength(0) - 1
                        Dim objclsColumns As New clsColumns

                        objclsColumns.COLUMNNAME = strColumns(j, 0).ToString
                        objclsColumns.COLUMNTYPE = strColumns(j, 1).ToString

                        objclsColumns.COLName = strColumns(j, 2).ToString
                        objclsColumns.COLSize = strColumns(j, 3).ToString
                        objclsColumns.COLNumericPrecision = strColumns(j, 4).ToString
                        objclsColumns.COLNumericScale = strColumns(j, 5).ToString
                        objclsColumns.COLDataType = strColumns(j, 6).ToString
                        objclsColumns.COLProviderType = strColumns(j, 7).ToString
                        objclsColumns.COLIsAutoIncrement = strColumns(j, 8).ToString
                        objclsColumns.COLSizeChar = strColumns(j, 9).ToString
                        objclsColumns.COLDataTypeSQL = strColumns(j, 10).ToString

                        objclsColumnsListing.Add(objclsColumns)

                        System.Windows.Forms.Application.DoEvents()
                    Next
                    objclsTables.COLUMNDETAILS = objclsColumnsListing
                    objclsTablesListing.Add(objclsTables)

                    System.Windows.Forms.Application.DoEvents()
                Next
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return objclsTablesListing
        End Function

        Private Shared Function GetDBTypeValue(ByVal DBTypeValue As String) As String
            Dim strDBTypeValue As String = ""
            Try
                '*****TO CAPTURE ALL DBTYPE VALUES IN THIS STRING TO BE USED FURTHER**********
                '************NOT NECESSORY IN THIS FUNCTION
                strDBTypeValueALL = strDBTypeValueALL & DBTypeValue & vbCrLf
                '*****
                '***FOR MS ACCESS********************
                If DBTypeValue = "DBTYPE_I4" Then strDBTypeValue = "Integer"
                If DBTypeValue = "DBTYPE_WVARCHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_WLONGVARCHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_DATE" Then strDBTypeValue = "Date"
                If DBTypeValue = "DBTYPE_CY" Then strDBTypeValue = "Double"
                If DBTypeValue = "DBTYPE_BOOL" Then strDBTypeValue = "Boolean"
                If DBTypeValue = "DBTYPE_LONGVARBINARY" Then strDBTypeValue = "Byte()"
                '***FOR MS SQL********************
                If DBTypeValue = "DBTYPE_I8" Then strDBTypeValue = "Long"
                If DBTypeValue = "DBTYPE_BINARY" Then strDBTypeValue = "Object"
                If DBTypeValue = "DBTYPE_BOOL" Then strDBTypeValue = "Boolean"
                If DBTypeValue = "DBTYPE_CHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_DBTIMESTAMP" Then strDBTypeValue = "Date"
                If DBTypeValue = "DBTYPE_NUMERIC" Then strDBTypeValue = "Double"
                If DBTypeValue = "DBTYPE_NUMERIC" Then strDBTypeValue = "Double"
                If DBTypeValue = "DBTYPE_R8" Then strDBTypeValue = "Single"
                If DBTypeValue = "DBTYPE_LONGVARBINARY" Then strDBTypeValue = "Byte()"
                If DBTypeValue = "DBTYPE_I4" Then strDBTypeValue = "Integer"
                If DBTypeValue = "DBTYPE_CY" Then strDBTypeValue = "Double"
                If DBTypeValue = "DBTYPE_WCHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_WLONGVARCHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_NUMERIC" Then strDBTypeValue = "Double"
                If DBTypeValue = "DBTYPE_NUMERIC" Then strDBTypeValue = "Double"
                If DBTypeValue = "DBTYPE_WVARCHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_WLONGVARCHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_R4" Then strDBTypeValue = "Single"
                If DBTypeValue = "DBTYPE_DBTIMESTAMP" Then strDBTypeValue = "Date"
                If DBTypeValue = "DBTYPE_I2" Then strDBTypeValue = "Int16"
                If DBTypeValue = "DBTYPE_CY" Then strDBTypeValue = "Single"
                If DBTypeValue = "DBTYPE_VARIANT" Then strDBTypeValue = "Object"
                If DBTypeValue = "DBTYPE_LONGVARCHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_BINARY" Then strDBTypeValue = "Time"
                If DBTypeValue = "DBTYPE_UI1" Then strDBTypeValue = "Int16"
                If DBTypeValue = "DBTYPE_GUID" Then strDBTypeValue = "Long"
                If DBTypeValue = "DBTYPE_VARBINARY" Then strDBTypeValue = "Object"
                If DBTypeValue = "DBTYPE_LONGVARBINARY" Then strDBTypeValue = "Byte()"
                If DBTypeValue = "DBTYPE_VARCHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_LONGVARCHAR" Then strDBTypeValue = "String"
                If DBTypeValue = "DBTYPE_WLONGVARCHAR" Then strDBTypeValue = "String"

                'strDBTypeValue = DBTypeValue
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return LCase(strDBTypeValue)
        End Function
        'Private Shared Function GetDBTypeValue_UPDATED(ByVal DBTypeValue As String) As String
        '    Dim strDBTypeValue As String = ""
        '    Try
        '        '*****TO CAPTURE ALL DBTYPE VALUES IN THIS STRING TO BE USED FURTHER**********
        '        '************NOT NECESSORY IN THIS FUNCTION
        '        strDBTypeValueALL = strDBTypeValueALL & DBTypeValue & vbCrLf
        '        '*****
        '        '***FOR MS ACCESS********************
        '        If DBTypeValue = "DBTYPE_I4" Then strDBTypeValue = "Integer"
        '        If DBTypeValue = "DBTYPE_WVARCHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_WLONGVARCHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_DATE" Then strDBTypeValue = "Date"
        '        If DBTypeValue = "DBTYPE_CY" Then strDBTypeValue = "Double"
        '        If DBTypeValue = "DBTYPE_BOOL" Then strDBTypeValue = "Boolean"
        '        If DBTypeValue = "DBTYPE_LONGVARBINARY" Then strDBTypeValue = "Byte()"
        '        '***FOR MS SQL********************
        '        If DBTypeValue = "DBTYPE_I8" Then strDBTypeValue = "Long"
        '        If DBTypeValue = "DBTYPE_BINARY" Then strDBTypeValue = "Object"
        '        If DBTypeValue = "DBTYPE_BOOL" Then strDBTypeValue = "Boolean"
        '        If DBTypeValue = "DBTYPE_CHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_DBTIMESTAMP" Then strDBTypeValue = "Date"
        '        If DBTypeValue = "DBTYPE_NUMERIC" Then strDBTypeValue = "Double"
        '        If DBTypeValue = "DBTYPE_NUMERIC" Then strDBTypeValue = "Double"
        '        If DBTypeValue = "DBTYPE_R8" Then strDBTypeValue = "Single"
        '        If DBTypeValue = "DBTYPE_LONGVARBINARY" Then strDBTypeValue = "Byte()"
        '        If DBTypeValue = "DBTYPE_I4" Then strDBTypeValue = "Integer"
        '        If DBTypeValue = "DBTYPE_CY" Then strDBTypeValue = "Double"
        '        If DBTypeValue = "DBTYPE_WCHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_WLONGVARCHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_NUMERIC" Then strDBTypeValue = "Double"
        '        If DBTypeValue = "DBTYPE_NUMERIC" Then strDBTypeValue = "Double"
        '        If DBTypeValue = "DBTYPE_WVARCHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_WLONGVARCHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_R4" Then strDBTypeValue = "Single"
        '        If DBTypeValue = "DBTYPE_DBTIMESTAMP" Then strDBTypeValue = "Date"
        '        If DBTypeValue = "DBTYPE_I2" Then strDBTypeValue = "Int16"
        '        If DBTypeValue = "DBTYPE_CY" Then strDBTypeValue = "Single"
        '        If DBTypeValue = "DBTYPE_VARIANT" Then strDBTypeValue = "Object"
        '        If DBTypeValue = "DBTYPE_LONGVARCHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_BINARY" Then strDBTypeValue = "Time"
        '        If DBTypeValue = "DBTYPE_UI1" Then strDBTypeValue = "Int16"
        '        If DBTypeValue = "DBTYPE_GUID" Then strDBTypeValue = "Long"
        '        If DBTypeValue = "DBTYPE_VARBINARY" Then strDBTypeValue = "Object"
        '        If DBTypeValue = "DBTYPE_LONGVARBINARY" Then strDBTypeValue = "Byte()"
        '        If DBTypeValue = "DBTYPE_VARCHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_LONGVARCHAR" Then strDBTypeValue = "String"
        '        If DBTypeValue = "DBTYPE_WLONGVARCHAR" Then strDBTypeValue = "String"

        '        'strDBTypeValue = DBTypeValue
        '    Catch ex As Exception
        '        MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
        '    End Try
        '    Return LCase(strDBTypeValue)
        'End Function

        Public Shared Function GetDBTypeValue4SP(ByVal DBTypeValue As String) As String
            Dim strDBTypeValue As String = ""
            Try
                '*****TO CAPTURE ALL DBTYPE VALUES IN THIS STRING TO BE USED FURTHER**********
                '************NOT NECESSORY IN THIS FUNCTION
                strDBTypeValueALL = strDBTypeValueALL & DBTypeValue & vbCrLf
                '*****

                '***FOR MS SQL********************
                If DBTypeValue = "long" Then strDBTypeValue = "int"
                If DBTypeValue = "object" Then strDBTypeValue = "binary"
                If DBTypeValue = "boolean" Then strDBTypeValue = "bit"
                If DBTypeValue = "string" Then strDBTypeValue = "varchar(100)"
                If DBTypeValue = "date" Then strDBTypeValue = "datetime"
                If DBTypeValue = "double" Then strDBTypeValue = "numeric"
                If DBTypeValue = "double" Then strDBTypeValue = "numeric"
                If DBTypeValue = "single" Then strDBTypeValue = "float"
                If DBTypeValue = "object" Then strDBTypeValue = "varbinary(100)"
                If DBTypeValue = "integer" Then strDBTypeValue = "int"
                If DBTypeValue = "double" Then strDBTypeValue = "double"
                If DBTypeValue = "string" Then strDBTypeValue = "varchar(100)"
                If DBTypeValue = "string" Then strDBTypeValue = "varchar(100)"
                If DBTypeValue = "double" Then strDBTypeValue = "numeric"
                If DBTypeValue = "double" Then strDBTypeValue = "numeric"
                If DBTypeValue = "string" Then strDBTypeValue = "varchar(100)"
                If DBTypeValue = "string" Then strDBTypeValue = "varchar(100)"
                If DBTypeValue = "single" Then strDBTypeValue = "float"
                If DBTypeValue = "date" Then strDBTypeValue = "datetime"
                If DBTypeValue = "int16" Then strDBTypeValue = "int"
                If DBTypeValue = "single" Then strDBTypeValue = "float"
                If DBTypeValue = "object" Then strDBTypeValue = "variant"
                If DBTypeValue = "string" Then strDBTypeValue = "varchar(100)"
                If DBTypeValue = "time" Then strDBTypeValue = "binary"
                If DBTypeValue = "int16" Then strDBTypeValue = "int"
                If DBTypeValue = "long" Then strDBTypeValue = "int"
                If DBTypeValue = "object" Then strDBTypeValue = "varbinary"
                If DBTypeValue = "byte()" Then strDBTypeValue = "varbinary(max)"
                If DBTypeValue = "string" Then strDBTypeValue = "varchar(100)"
                If DBTypeValue = "string" Then strDBTypeValue = "varchar(100)"
                If DBTypeValue = "string" Then strDBTypeValue = "varchar(100)"



                'strDBTypeValue = DBTypeValue
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return LCase(strDBTypeValue)
        End Function
        Public Shared Function GetDBTypeValue4SP_UPDATED(ByVal clsColumn As MyCLS.clsColumns) As String
            Dim strDBTypeValue As String = ""
            Try
                Dim DBTypeValue As String = clsColumn.COLDataTypeSQL
                Dim DBTypeSize As String = ""

                '*****TO CAPTURE ALL DBTYPE VALUES IN THIS STRING TO BE USED FURTHER**********
                '************NOT NECESSORY IN THIS FUNCTION
                strDBTypeValueALL = strDBTypeValueALL & DBTypeValue & vbCrLf
                '*****

                '***FOR MS SQL********************
                If clsColumn.COLSizeChar = -1 And clsColumn.COLDataTypeSQL <> "xml" Then
                    DBTypeSize = "(max)"
                ElseIf clsColumn.COLSizeChar = 0 And (clsColumn.COLDataTypeSQL = "numeric" Or clsColumn.COLDataTypeSQL = "decimal") Then
                    DBTypeSize = "(" & clsColumn.COLNumericPrecision & ", " & clsColumn.COLNumericScale & ")"
                ElseIf clsColumn.COLSizeChar > 0 And clsColumn.COLSizeChar <= 8000 Then
                    DBTypeSize = "(" & clsColumn.COLSizeChar & ")"
                End If

                strDBTypeValue = DBTypeValue & DBTypeSize

                'strDBTypeValue = DBTypeValue
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return LCase(strDBTypeValue)
        End Function

        Public Shared Function GetDBTypeValue4VB6(ByVal DBTypeValue As String) As String
            Dim strDBTypeValue As String = ""
            Try
                '*****TO CAPTURE ALL DBTYPE VALUES IN THIS STRING TO BE USED FURTHER**********
                '************NOT NECESSORY IN THIS FUNCTION
                strDBTypeValueALL = strDBTypeValueALL & DBTypeValue & vbCrLf
                '*****

                '***FOR MS SQL********************
                If DBTypeValue = "time" Then
                    strDBTypeValue = "Date"
                ElseIf DBTypeValue = "int16" Then
                    strDBTypeValue = "Integer"
                Else
                    strDBTypeValue = DBTypeValue
                End If

                'strDBTypeValue = DBTypeValue
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return LCase(strDBTypeValue)
        End Function

        Public Shared Function GetDBTypeValue4VB6SPParam(ByVal DBTypeValue As String) As String
            Dim strDBTypeValue As String = ""
            Try
                '*****TO CAPTURE ALL DBTYPE VALUES IN THIS STRING TO BE USED FURTHER**********
                '************NOT NECESSORY IN THIS FUNCTION
                strDBTypeValueALL = strDBTypeValueALL & DBTypeValue & vbCrLf
                '*****
                '***FOR MS SQL********************
                If DBTypeValue = "long" Then strDBTypeValue = "adNumeric"
                If DBTypeValue = "object" Then strDBTypeValue = "adobject"
                If DBTypeValue = "boolean" Then strDBTypeValue = "adboolean"
                If DBTypeValue = "string" Then strDBTypeValue = "adVarChar"
                If DBTypeValue = "date" Then strDBTypeValue = "addate"
                If DBTypeValue = "double" Then strDBTypeValue = "addouble"
                If DBTypeValue = "double" Then strDBTypeValue = "addouble"
                If DBTypeValue = "single" Then strDBTypeValue = "adsingle"
                If DBTypeValue = "object" Then strDBTypeValue = "adobject"
                If DBTypeValue = "integer" Then strDBTypeValue = "adinteger"
                If DBTypeValue = "double" Then strDBTypeValue = "addouble"
                If DBTypeValue = "string" Then strDBTypeValue = "adVarChar"
                If DBTypeValue = "string" Then strDBTypeValue = "adVarChar"
                If DBTypeValue = "double" Then strDBTypeValue = "addouble"
                If DBTypeValue = "double" Then strDBTypeValue = "addouble"
                If DBTypeValue = "string" Then strDBTypeValue = "adVarChar"
                If DBTypeValue = "string" Then strDBTypeValue = "adVarChar"
                If DBTypeValue = "single" Then strDBTypeValue = "adsingle"
                If DBTypeValue = "date" Then strDBTypeValue = "addate"
                If DBTypeValue = "Integer" Then strDBTypeValue = "adInteger"
                If DBTypeValue = "single" Then strDBTypeValue = "adsingle"
                If DBTypeValue = "object" Then strDBTypeValue = "adobject"
                If DBTypeValue = "string" Then strDBTypeValue = "adVarChar"
                If DBTypeValue = "time" Then strDBTypeValue = "adtime"
                If DBTypeValue = "Integer" Then strDBTypeValue = "adInteger"
                If DBTypeValue = "long" Then strDBTypeValue = "adlong"
                If DBTypeValue = "object" Then strDBTypeValue = "adobject"
                If DBTypeValue = "object" Then strDBTypeValue = "adobject"
                If DBTypeValue = "string" Then strDBTypeValue = "adVarChar"
                If DBTypeValue = "string" Then strDBTypeValue = "adVarChar"
                If DBTypeValue = "string" Then strDBTypeValue = "adVarChar"


                'strDBTypeValue = DBTypeValue
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return LCase(strDBTypeValue)
        End Function

        Public Shared Function GetDBTypeValue4SqlDbTypes(ByVal DBTypeValue As String) As String
            Dim strDBTypeValue As String = ""
            Try
                '*****TO CAPTURE ALL DBTYPE VALUES IN THIS STRING TO BE USED FURTHER**********
                '************NOT NECESSORY IN THIS FUNCTION
                strDBTypeValueALL = strDBTypeValueALL & DBTypeValue & vbCrLf
                '*****

                '***FOR MS SQL********************
                If DBTypeValue = "long" Then strDBTypeValue = "BigInt"
                If DBTypeValue = "object" Then strDBTypeValue = "Binary"
                If DBTypeValue = "boolean" Then strDBTypeValue = "Bit"
                If DBTypeValue = "string" Then strDBTypeValue = "VarChar , 100"
                If DBTypeValue = "date" Then strDBTypeValue = "DateTime"
                If DBTypeValue = "double" Then strDBTypeValue = "Decimal"
                If DBTypeValue = "single" Then strDBTypeValue = "Float"
                If DBTypeValue = "integer" Then strDBTypeValue = "Int"
                If DBTypeValue = "int16" Then strDBTypeValue = "int"
                If DBTypeValue = "time" Then strDBTypeValue = "Timestamp"

            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return strDBTypeValue
        End Function
        Public Shared Function GetDBTypeValue4SqlDbTypes_UPDATED(ByVal clsColumn As MyCLS.clsColumns) As String
            Dim strDBTypeValue As String = ""
            Try
                Dim DBTypeValue As String = clsColumn.COLDataTypeSQL
                '*****TO CAPTURE ALL DBTYPE VALUES IN THIS STRING TO BE USED FURTHER**********
                '************NOT NECESSORY IN THIS FUNCTION
                strDBTypeValueALL = strDBTypeValueALL & DBTypeValue & vbCrLf
                '*****

                '***FOR MS SQL********************
                strDBTypeValue = DBTypeValue

                If DBTypeValue = "binary" Then strDBTypeValue = "varbinary"
                If DBTypeValue = "smalldatetime" Then strDBTypeValue = "datetime"
                If DBTypeValue = "numeric" Then strDBTypeValue = "decimal"
                If DBTypeValue = "numeric" Then strDBTypeValue = "decimal"
                If DBTypeValue = "sql_variant" Then strDBTypeValue = "variant"
                If DBTypeValue = "int" Then strDBTypeValue = "Int"

            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
            Return strDBTypeValue
        End Function
    End Class
    '*************END - DATABASE OPERATIONS***************
#End Region


    '***********START - OPEN CONNECTION FROM XML FILE******************
#Region "ConnectionInfo"
    Public Class ConnectionInfo
        Private m_strServerName As String
        Private m_strPassword As String
        Private m_strDatabase As String
        Private m_strUserID As String

        Public Property ServerName() As String
            Get
                Return m_strServerName
            End Get
            Set(ByVal value As String)
                m_strServerName = value
            End Set
        End Property


        Public Property Password() As String
            Get
                Return m_strPassword
            End Get
            Set(ByVal value As String)
                m_strPassword = value
            End Set
        End Property

        Public Property Database() As String
            Get
                Return m_strDatabase
            End Get
            Set(ByVal value As String)
                m_strDatabase = value
            End Set
        End Property

        Public Property UserID() As String
            Get
                Return m_strUserID
            End Get
            Set(ByVal value As String)
                m_strUserID = value
            End Set
        End Property
    End Class
#End Region

#Region "XML Related"
    Public Class SerializeXML
        'Public Shared Function ConvertXML(ByVal FilePath As String, ByVal XMLContext As EXMLContextTypes, ByVal Encrypted As Boolean, ByVal EncryptedString As String) As Object
        Public Function ConvertXML(ByVal FilePath As String, ByVal Encrypted As Boolean, ByVal EncryptedString As String) As Object
            Dim objSerialClassObject As Object = Nothing
            Dim objXMLSerializer As XmlSerializer = Nothing
            Dim objStream As FileStream = Nothing
            Dim objXMLReader As XmlReader = Nothing
            Try
                'Select Case XMLContext
                '    Case EXMLContextTypes.ClientDBConnection
                If True Then
                    objXMLSerializer = New XmlSerializer(GetType(ConnectionInfo))
                    '    Exit Select
                End If
                '    Case Else

                'Exit Select
                'End Select

                If objXMLSerializer Is Nothing Then
                    Return objSerialClassObject
                End If

                objStream = New FileStream(FilePath, FileMode.Open)
                objXMLReader = New XmlTextReader(objStream)

                objSerialClassObject = objXMLSerializer.Deserialize(objXMLReader)
            Catch exObj As Exception
                'BugsHandler.BugLogging(exObj.StackTrace, exObj.Message, true);
            Finally
                If objStream IsNot Nothing Then
                    objStream.Close()
                End If
                objStream = Nothing
                If objXMLReader IsNot Nothing Then
                    objXMLReader.Close()
                End If
                objXMLReader = Nothing
            End Try
            Return objSerialClassObject
        End Function
    End Class
#End Region
    '***********END - OPEN CONNECTION FROM XML FILE******************



    '***********START - EXCEL FILE OPERATIONS******************
    Public Class clsXLSOperations
        Dim app As Excel.Application
        Dim WB As Excel.Workbook

        Public Function GetDataFromXLS(ByVal vFile As String, ByVal strTableName As String, ByVal strQuery As String) As DataSet
            Try
                'Dim Oleda As OleDbDataAdapter
                Dim Olecn As OleDbConnection
                'Dim dt1 As DataTable
                'lblMSG.Text = "Opening Excel File..."
                Olecn = New OleDbConnection( _
                    "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & vFile & ";" & _
                    "Extended Properties=Excel 8.0;")
                Olecn.Open()

                'Dim ExcelCommand As New System.Data.OleDb.OleDbCommand("SELECT INTO [ODBC Driver={SQL Server};Server=tsi_dev_02;Database=ndhhs_updated;uid=sa;pwd=sa123].[tblOutstanding] FROM [Sheet1$]", Olecn)
                'lblMSG.Text = "Reading Excel File..."
                'Dim ExcelCommand As New OleDbCommand("SELECT * FROM [" & strSheetName & "$]", Olecn)
                'Dim ExcelCommand As New OleDbCommand(strQuery, Olecn)

                Dim ds As New DataSet
                MyCLS.clsCOMMON.SetCon(Olecn)
                MyCLS.clsCOMMON.prcQuerySelectDS(ds, strQuery, strTableName)

                'Dim Rs As OleDbDataReader = ExcelCommand.ExecuteReader()

                Olecn.Close()

                Return ds
            Catch ex As Exception
                MsgBox(ex.Message)
                'MYCLS.strGlobalErrorInfo = "Query is : " & TruncateTable
                MyCLS.strGlobalErrorInfo = MyCLS.strGlobalErrorInfo & vbCrLf & vbCrLf & String.Concat(ex.Message & vbCrLf, ex.Source & vbCrLf, ex.StackTrace & vbCrLf)
                MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.TargetSite, ex.InnerException)
                MyCLS.strGlobalErrorInfo = String.Concat(MyCLS.strGlobalErrorInfo & vbCrLf, ex.Data)
                MyCLS.clsCOMMON.fnWrite2LOG(MyCLS.strGlobalErrorInfo, System.Reflection.MethodBase.GetCurrentMethod().ToString())
            End Try
        End Function

        Public Function GetDataFromXLS(ByVal strExcelFile As String) As DataTable
            Dim dt As New DataTable
            Try
                'EXCEL DETAIL FILE
                'WB = app.Workbooks.Open(txtSource.Text & "\EXLSheet.xls")
                app = New Excel.Application
                WB = app.Workbooks.Open(strExcelFile)
                Dim UsedRng As Excel.Range
                Dim Cell As Excel.Range

                If WB IsNot Nothing Then
                    Dim ws As Excel.Worksheet = WB.Worksheets.Item(1)

                    'CREATE FILES FOR LOG & INSERT QUERIES FOR MSSQL
                    'OpenFile()

                    UsedRng = ws.UsedRange

                    Dim MaxCols As Long = UsedRng.Columns.Count
                    Dim iCol As Int16 = 0
                    For iCol = 0 To MaxCols - 1
                        dt.Columns.Add(New DataColumn("Col" & iCol))
                    Next
                    Dim dtR As DataRow

                    Dim iRow As Int16 = 0
                    iCol = 0
                    For Each Cell In UsedRng.Cells
                        'START FROM 2nd ROW
                        If Cell.Row.ToString() > 1 Then
                            Try
                                If Cell.Column = 1 Then
                                    dtR = dt.NewRow()
                                    iRow += 1
                                    iCol = 0
                                End If

                                'Debug.Print("R" & Cell.Row & "C" & Cell.Column & " : " & Cell.Text)
                                'dtR("Col" & iCol) = Cell(iRow, iCol + 1).text
                                dtR("Col" & iCol) = Cell.Text
                                iCol += 1

                                If iCol = MaxCols Then
                                    dt.Rows.Add(dtR)
                                    'Debug.Print(dtR(0) & ", " & dtR(1) & ", " & dtR(2) & ", " & dtR(3) & ", " & dtR(4) & ", " & dtR(5) & ", " & dtR(6) & ", " & dtR(7) & ", " & dtR(8) & ", " & dtR(9) & ", " & dtR(10) & ", " & dtR(11))
                                    'iRow += 1
                                    'iCol = 0
                                End If
                            Catch ex As Exception
                                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
                            End Try
                        End If
                        System.Windows.Forms.Application.DoEvents()
                    Next
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message & vbCrLf & ex.StackTrace)
                'WriteFile("ERR : " & strISBN & " -- " & ex.Message)
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            Finally
                Try
                    WB.Save()
                    WB.Close()
                Catch
                Finally
                    Try
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(WB)
                        WB = Nothing
                        app.Quit()
                    Finally
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app)
                        app = Nothing
                    End Try
                End Try
            End Try
            Return dt
        End Function

        Public Sub Txt2Xls(ByVal TxtFilePath As String, Optional ByVal RunFile As Boolean = False)
            Const xlFixedWidth = 2
            Const xlNormal = -4143
            Const xlLastCell = 11

            Dim sFiNa
            Dim oFS
            Dim oExcel
            Dim oWBook
            Dim sTmp

            ' you'll have to change this 
            sFiNa = TxtFilePath

            oFS = CreateObject("Scripting.FileSystemObject")
            sFiNa = oFS.GetAbsolutePathName(sFiNa)
            oExcel = CreateObject("Excel.Application")

            oExcel.Visible = True   ' while testing 
            ' oExcel.Whatever = False  ' todo: what property to set to suppress silly question 

            sTmp = "Working with MS Excel Vers. " & oExcel.Version _
                   & " (" & oExcel.Workbooks.Count & " Workbooks)"

            'WScript.Echo(sTmp)

            'oExcel.Workbooks.Open(sFiNa, xlFixedWidth)
            oExcel.Workbooks.Open(sFiNa, , , 1, , , , , vbTab)
            oWBook = oExcel.Workbooks(1)
            'WScript.Echo("Open:   ", oWBook.Name)

            ' magic from Giovanni Cenati 
            ' http://www.codecomments.com/archive299-2005-2-401145.html 
            ' oExcel.Range(oExcel.cells(1,1),oExcel.cells(100,1)).Select 

            'oExcel.Range(oExcel.cells(1, 1), oExcel.cells(oWBook.Sheets(1).UsedRange.SpecialCells(xlLastCell).Row, 1)).Select()
            'oExcel.Selection.TextToColumns(oExcel.Range("A1"), xlFixedWidth)

            ' save as XLS 
            oWBook.SaveAs(Replace(sFiNa, ".txt", "") + ".xls", xlNormal)
            'WScript.Echo("SaveAs: ", oWBook.Name)

            oWBook.Close()
            oExcel.Quit()
            If RunFile Then
                Process.Start("Excel.exe", Replace(sFiNa, ".txt", "") + ".xls")
                'Shell("Excel.exe " & Replace(sFiNa, ".txt", "") + ".xls", AppWinStyle.MaximizedFocus)
            End If
        End Sub

        Public Shared Sub DataGridToExcel(ByVal fileName As String, ByVal dgv As System.Windows.Forms.DataGridView)
            Try
                Dim fs As New IO.StreamWriter(fileName, False)
                fs.WriteLine("<?xml version=""1.0""?>")
                fs.WriteLine("<?mso-application progid=""Excel.Sheet""?>")
                fs.WriteLine("<ss:Workbook xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">")
                fs.WriteLine(" <ss:Styles>")
                fs.WriteLine(" <ss:Style ss:ID=""1"">")
                fs.WriteLine(" <ss:Font ss:Bold=""1""/>")
                fs.WriteLine(" </ss:Style>")
                fs.WriteLine(" </ss:Styles>")
                fs.WriteLine(" <ss:Worksheet ss:Name=""Sheet1"">")
                fs.WriteLine(" <ss:Table>")
                For x As Integer = 0 To dgv.Columns.Count - 1
                    'If dtgEntities.Rows(x).Visible = True Then
                    fs.WriteLine(" <ss:Column ss:Width=""{0}""/>", dgv.Columns.Item(x).Width)
                    'End If
                Next
                fs.WriteLine(" <ss:Row ss:StyleID=""1"">")
                For i As Integer = 0 To dgv.Columns.Count - 1
                    'If dtgEntities.Rows(i).Visible = True Then
                    fs.WriteLine(" <ss:Cell>")
                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", dgv.Columns.Item(i).HeaderText))
                    fs.WriteLine(" </ss:Cell>")
                    'End If
                Next
                fs.WriteLine(" </ss:Row>")
                'For intRow As Integer = 0 To dgv.RowCount - 1
                For intRow As Integer = 0 To dgv.RowCount - 2
                    Try
                        If dgv.Rows(intRow).Visible = True Then
                            Try
                                fs.WriteLine(String.Format(" <ss:Row ss:Height =""{0}"">", dgv.Rows(intRow).Height))
                                For intCol As Integer = 0 To dgv.Columns.Count - 1
                                    fs.WriteLine(" <ss:Cell>")
                                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", dgv.Item(intCol, intRow).Value.ToString))
                                    fs.WriteLine(" </ss:Cell>")
                                Next
                                fs.WriteLine(" </ss:Row>")
                            Catch ex As Exception

                            End Try
                        End If
                    Catch ex As Exception

                    End Try
                Next
                fs.WriteLine(" </ss:Table>")
                fs.WriteLine(" </ss:Worksheet>")
                fs.WriteLine("</ss:Workbook>")
                fs.Close()
                'MsgBox("Export Complete!", MsgBoxStyle.OkOnly, "Export Complete")
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Sub DataSetToExcel(ByVal fileName As String, ByVal ds As DataSet)
            Try
                Dim fs As New IO.StreamWriter(My.Application.Info.DirectoryPath() & "\Export\" & fileName, False)
                fs.WriteLine("<?xml version=""1.0""?>")
                fs.WriteLine("<?mso-application progid=""Excel.Sheet""?>")
                fs.WriteLine("<ss:Workbook xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">")
                fs.WriteLine(" <ss:Styles>")
                fs.WriteLine(" <ss:Style ss:ID=""1"">")
                fs.WriteLine(" <ss:Font ss:Bold=""1""/>")
                fs.WriteLine(" </ss:Style>")
                fs.WriteLine(" </ss:Styles>")
                fs.WriteLine(" <ss:Worksheet ss:Name=""Sheet1"">")
                fs.WriteLine(" <ss:Table>")
                For x As Integer = 0 To ds.Tables(0).Columns.Count - 1
                    'If dtgEntities.Rows(x).Visible = True Then
                    fs.WriteLine(" <ss:Column ss:Width=""{0}""/>", 70)
                    'End If
                Next
                fs.WriteLine(" <ss:Row ss:StyleID=""1"">")
                For i As Integer = 0 To ds.Tables(0).Columns.Count - 1
                    'If dtgEntities.Rows(i).Visible = True Then
                    fs.WriteLine(" <ss:Cell>")
                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", ds.Tables(0).Columns(i).ColumnName))
                    fs.WriteLine(" </ss:Cell>")
                    'End If
                Next
                fs.WriteLine(" </ss:Row>")
                'For intRow As Integer = 0 To ds.Tables(0).RowCount - 1
                For intRow As Integer = 0 To ds.Tables(0).Rows.Count - 1
                    Try
                        'If ds.Tables(0).Rows(intRow).Visible = True Then
                        Try
                            fs.WriteLine(String.Format(" <ss:Row ss:Height =""{0}"">", 15))
                            For intCol As Integer = 0 To ds.Tables(0).Columns.Count - 1
                                fs.WriteLine(" <ss:Cell>")

                                '1. OLD One
                                'fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", ds.Tables(0).Rows(intRow)(intCol).ToString()))

                                '2. Generic to be run in all default conditions
                                If (ds.Tables(0).Columns(intCol).DataType.FullName = "System.Int32" Or ds.Tables(0).Columns(intCol).DataType.FullName = "System.Int64" Or ds.Tables(0).Columns(intCol).DataType.FullName = "System.Decimal") Then
                                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""Number"">{0}</ss:Data>", ds.Tables(0).Rows(intRow)(intCol)))
                                ElseIf (ds.Tables(0).Columns(intCol).DataType.FullName = "System.DateTime") Then
                                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", ds.Tables(0).Rows(intRow)(intCol).ToString().Replace(" 12:00:00 AM", "")))
                                Else
                                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", ReplaceXMLSpecialChars(ds.Tables(0).Rows(intRow)(intCol).ToString())))
                                End If

                                fs.WriteLine(" </ss:Cell>")
                            Next
                            fs.WriteLine(" </ss:Row>")
                        Catch ex As Exception

                        End Try
                        'End If
                    Catch ex As Exception

                    End Try
                Next
                fs.WriteLine(" </ss:Table>")
                fs.WriteLine(" </ss:Worksheet>")
                fs.WriteLine("</ss:Workbook>")
                fs.Close()
                'MsgBox("Export Complete!", MsgBoxStyle.OkOnly, "Export Complete")
            Catch ex As Exception

            End Try
        End Sub
        ''' <summary>
        ''' Exports Data to Excel...
        ''' </summary>
        ''' <param name="fileName"></param>
        ''' <param name="ds"></param>
        ''' <remarks></remarks>
        Public Shared Sub DataTableToExcel(ByVal fileName As String, ByVal dt As DataTable)
            Try
                Dim fs As New IO.StreamWriter(fileName, False)
                fs.WriteLine("<?xml version=""1.0""?>")
                fs.WriteLine("<?mso-application progid=""Excel.Sheet""?>")
                fs.WriteLine("<ss:Workbook xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">")
                fs.WriteLine(" <ss:Styles>")
                fs.WriteLine(" <ss:Style ss:ID=""1"">")
                fs.WriteLine(" <ss:Font ss:Bold=""1""/>")
                fs.WriteLine(" </ss:Style>")
                fs.WriteLine(" </ss:Styles>")
                fs.WriteLine(" <ss:Worksheet ss:Name=""Sheet1"">")
                fs.WriteLine(" <ss:Table>")
                For x As Integer = 0 To dt.Columns.Count - 1
                    'If dtgEntities.Rows(x).Visible = True Then
                    fs.WriteLine(" <ss:Column ss:Width=""{0}""/>", 70)
                    'End If
                Next
                fs.WriteLine(" <ss:Row ss:StyleID=""1"">")
                For i As Integer = 0 To dt.Columns.Count - 1
                    'If dtgEntities.Rows(i).Visible = True Then
                    fs.WriteLine(" <ss:Cell>")
                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", dt.Columns(i).ColumnName))
                    fs.WriteLine(" </ss:Cell>")
                    'End If
                Next
                fs.WriteLine(" </ss:Row>")
                'For intRow As Integer = 0 To dt.RowCount - 1
                For intRow As Integer = 0 To dt.Rows.Count - 1
                    Try
                        'If dt.Rows(intRow).Visible = True Then
                        Try
                            fs.WriteLine(String.Format(" <ss:Row ss:Height =""{0}"">", 15))
                            For intCol As Integer = 0 To dt.Columns.Count - 1
                                fs.WriteLine(" <ss:Cell>")
                                '1.fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", dt.Rows(intRow)(intCol).ToString()))
                                '2. Generic to be run in all default conditions
                                If (dt.Columns(intCol).DataType.FullName = "System.Int32" Or dt.Columns(intCol).DataType.FullName = "System.Int64" Or dt.Columns(intCol).DataType.FullName = "System.Decimal") Then
                                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""Number"">{0}</ss:Data>", dt.Rows(intRow)(intCol)))
                                ElseIf (dt.Columns(intCol).DataType.FullName = "System.DateTime") Then
                                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", dt.Rows(intRow)(intCol).ToString().Replace(" 12:00:00 AM", "")))
                                Else
                                    fs.WriteLine(String.Format(" <ss:Data ss:Type=""String"">{0}</ss:Data>", ReplaceXMLSpecialChars(dt.Rows(intRow)(intCol).ToString())))
                                End If
                                fs.WriteLine(" </ss:Cell>")
                            Next
                            fs.WriteLine(" </ss:Row>")
                        Catch ex As Exception

                        End Try
                        'End If
                    Catch ex As Exception

                    End Try
                Next
                fs.WriteLine(" </ss:Table>")
                fs.WriteLine(" </ss:Worksheet>")
                fs.WriteLine("</ss:Workbook>")
                fs.Close()
                'MsgBox("Export Complete!", MsgBoxStyle.OkOnly, "Export Complete")
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        Public Sub generateExcel(ByVal dt As System.Data.DataTable, ByVal Heading As String, ByVal filterText As String, ByVal requiredFilterText As Boolean, ByVal requiredSerialNumber As Boolean, ByVal fileName As String)
            Dim newFile As FileInfo = New FileInfo(My.Application.Info.DirectoryPath() & "\Export\" & fileName)

            If newFile.Exists Then
                Return
                newFile.Delete()
                newFile = New FileInfo(fileName)
            End If

            Dim totalColumn As Integer = dt.Columns.Count
            Dim totalRow As Integer = dt.Rows.Count
            Dim row As Integer = 1
            Dim column As Integer = 1

            If requiredSerialNumber Then
                totalColumn += 1
            End If

            Using package As ExcelPackage = New ExcelPackage(newFile)
                Dim ws As ExcelWorksheet = package.Workbook.Worksheets.Add("Sheet1")
                ws.Row(row).Height = 20
                ws.Cells(row, column, row, totalColumn).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                ws.Cells(row, column, row, totalColumn).Style.VerticalAlignment = ExcelVerticalAlignment.Center
                ws.Cells(row, column, row, totalColumn).Style.Font.Bold = True
                ws.Cells(row, column, row, totalColumn).Style.Font.Size = 20
                ws.Cells(row, column, row, totalColumn).Style.Fill.PatternType = ExcelFillStyle.Solid
                ws.Cells(row, column, row, totalColumn).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffffff"))
                ws.Cells(row, column, row, totalColumn).Style.Font.Name = "Calibri"
                ws.Cells(row, column, row, totalColumn).Style.Font.Color.SetColor(Color.Black)
                ws.Cells(row, column, row, totalColumn).Merge = True
                ws.Cells(row, column, row, totalColumn).Value = "MyTool"
                row += 1
                ws.Row(row).Height = 20
                ws.Cells(row, column, row, totalColumn).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                ws.Cells(row, column, row, totalColumn).Style.VerticalAlignment = ExcelVerticalAlignment.Center
                ws.Cells(row, column, row, totalColumn).Style.Font.Bold = True
                ws.Cells(row, column, row, totalColumn).Style.Font.Size = 17
                ws.Cells(row, column, row, totalColumn).Style.Fill.PatternType = ExcelFillStyle.Solid
                ws.Cells(row, column, row, totalColumn).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#3c8dbc"))
                ws.Cells(row, column, row, totalColumn).Style.Font.Name = "Calibri"
                ws.Cells(row, column, row, totalColumn).Style.Font.Color.SetColor(Color.White)
                ws.Cells(row, column, row, totalColumn).Merge = True
                ws.Cells(row, column, row, totalColumn).Value = Heading
                row += 1

                If requiredFilterText Then
                    ws.Row(row).Height = 20
                    ws.Cells(row, column, row, totalColumn).Style.Font.Size = 12
                    ws.Cells(row, column, row, totalColumn).Style.Font.Bold = True
                    ws.Cells(row, column, row, totalColumn).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    ws.Cells(row, column, row, totalColumn).Style.VerticalAlignment = ExcelVerticalAlignment.Top
                    ws.Cells(row, column, row, totalColumn).Style.WrapText = True
                    ws.Cells(row, column, row, totalColumn).Style.Font.Name = "Calibri"
                    ws.Cells(row, column, row, totalColumn).Merge = True
                    ws.Cells(row, column, row, totalColumn).Value = filterText
                    row += 1
                End If

                ws.Cells(row, column, row, totalColumn).Style.Font.Size = 12
                ws.Cells(row, column, row, totalColumn).Style.Font.Bold = True
                ws.Cells(row, column, row, totalColumn).Style.Font.Name = "Calibri"
                ws.Cells(row, column, row, totalColumn).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                ws.Cells(row, column, row, totalColumn).Style.Fill.PatternType = ExcelFillStyle.Solid
                ws.Cells(row, column, row, totalColumn).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ecf0f5"))
                ws.Cells(row, column, row, totalColumn).Style.WrapText = True
                column = 1

                If requiredSerialNumber Then
                    ws.Cells(row, column).Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                    ws.Cells(row, column).Value = "S. No."
                    ws.Column(column).Width = 5
                    column += 1
                End If

                For Each dc As DataColumn In dt.Columns
                    ws.Cells(row, column).Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                    ws.Cells(row, column).Value = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToUpper(dc.ColumnName.ToString())
                    ws.Column(column).Width = 17
                    column += 1
                    System.Windows.Forms.Application.DoEvents()
                Next

                column = 1
                row += 1
                Dim serialNumber As Integer = 1
                ws.Cells(row, column, totalRow + row, totalColumn).Style.Font.Size = 12
                ws.Cells(row, column, totalRow + row, totalColumn).Style.Font.Name = "Calibri"
                ws.Cells(row, column, totalRow + row, totalColumn).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                ws.Cells(row, column, totalRow + row, totalColumn).Style.VerticalAlignment = ExcelVerticalAlignment.Top
                ws.Cells(row, column, totalRow + row, totalColumn).Style.WrapText = True

                For Each dr As DataRow In dt.Rows

                    If requiredSerialNumber Then
                        ws.Cells(row, Math.Min(System.Threading.Interlocked.Increment(column), column - 1)).Value = Math.Min(System.Threading.Interlocked.Increment(serialNumber), serialNumber - 1)
                    End If

                    For Each dc As DataColumn In dt.Columns
                        ws.Cells(row, Math.Min(System.Threading.Interlocked.Increment(column), column - 1)).Value = dr(dc.ColumnName).ToString()
                        System.Windows.Forms.Application.DoEvents()
                    Next

                    column = 1
                    row += 1
                    System.Windows.Forms.Application.DoEvents()

                    frmGenerateInserts.lblTableName.Text = Heading & " : " & row
                Next

                package.Save()

                ws.Dispose()
            End Using

            newFile = Nothing
        End Sub

        Public Shared Function ReplaceXMLSpecialChars(strColValue As String) As String
            '.Replace("", "")
            'Return strColValue.Replace(Chr(10), "").Replace(Chr(13), "").Replace("&", "&amp;").Replace("'", "&apos;").Replace("""", "&quot;").Replace("  ", "").Replace("<", "&lt;").Replace(">", "&gt;")
            Return System.Security.SecurityElement.Escape(strColValue)
        End Function


#Region "Open Multiple Excel Files by Passing PASSWORD"
        Shared objExcel

        Public Shared Sub OpenXLSObject(ByVal ShowExcel As Boolean)
            objExcel = CreateObject("Excel.Application")
            If ShowExcel Then
                objExcel.Visible = True
            End If
        End Sub
        Public Shared Sub CloseXLSObject()
            objExcel.Quit()
        End Sub
        Public Shared Sub OpenXLSFile(ByVal xlsFilePath As String, ByVal xlsPassword As String, ByVal RunFile As Boolean)
            Dim sTmp

            sTmp = "Working with MS Excel Vers. " & objExcel.Version _
                   & " (" & objExcel.Workbooks.Count & " Workbooks)"

            Try
                objExcel.Workbooks.Open(FileName:=xlsFilePath, Password:=xlsPassword)
            Catch ex As Exception
                Exit Sub
            End Try

            '***save as XLS***
            ''''oWBook = objExcel.Workbooks(1)
            ''''oWBook.SaveAs(Replace(sFiNa, ".txt", "") + ".xls", xlNormal)
            ''''oWBook.Close()            

            If RunFile Then
                Process.Start("Excel.exe", xlsFilePath)
                MsgBox(xlsPassword)
            End If
        End Sub
#End Region
    End Class



    Public Class clsXLSImportXML

        Private Structure ColumnType
            Public type As Type
            Private name As String

            Public Sub New(ByVal type As Type)
                Me.type = type
                Me.name = type.ToString().ToLower()
            End Sub

            Public Function ParseString(ByVal input As String) As Object
                If [String].IsNullOrEmpty(input) Then
                    Return DBNull.Value
                End If
                Select Case type.ToString()
                    Case "system.datetime"
                        Return DateTime.Parse(input)
                    Case "system.decimal"
                        Return Decimal.Parse(input)
                    Case "system.boolean"
                        Return Boolean.Parse(input)
                    Case Else
                        Return input
                End Select
            End Function
        End Structure

        Public Function ImportExcelXML(ByVal inputFileStream As Stream, ByVal hasHeaders As Boolean, ByVal autoDetectColumnType As Boolean) As DataSet
            Dim doc As New XmlDocument()
            doc.Load(New XmlTextReader(inputFileStream))
            Dim nsmgr As New XmlNamespaceManager(doc.NameTable)

            nsmgr.AddNamespace("o", "urn:schemas-microsoft-com:office:office")
            nsmgr.AddNamespace("x", "urn:schemas-microsoft-com:office:excel")
            nsmgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet")

            Dim ds As New DataSet()

            For Each node As System.Xml.XmlNode In doc.DocumentElement.SelectNodes("//ss:Worksheet", nsmgr)
                Dim dt As New DataTable(node.Attributes("ss:Name").Value)
                ds.Tables.Add(dt)
                Dim rows As XmlNodeList = node.SelectNodes("ss:Table/ss:Row", nsmgr)
                If rows.Count > 0 Then
                    Dim columns As New List(Of ColumnType)()
                    Dim startIndex As Integer = 0
                    If hasHeaders Then
                        For Each data As System.Xml.XmlNode In rows(0).SelectNodes("ss:Cell/ss:Data", nsmgr)
                            columns.Add(New ColumnType(GetType(String)))
                            'default to text
                            dt.Columns.Add(data.InnerText, GetType(String))
                        Next
                        startIndex += 1
                    End If
                    If autoDetectColumnType AndAlso rows.Count > 0 Then
                        Dim cells As XmlNodeList = rows(startIndex).SelectNodes("ss:Cell", nsmgr)
                        Dim actualCellIndex As Integer = 0
                        For cellIndex As Integer = 0 To cells.Count - 1
                            Dim cell As System.Xml.XmlNode = cells(cellIndex)
                            If cell.Attributes("ss:Index") IsNot Nothing Then
                                actualCellIndex = Integer.Parse(cell.Attributes("ss:Index").Value) - 1
                            End If

                            Dim autoDetectType As ColumnType = [getType](cell.SelectSingleNode("ss:Data", nsmgr))

                            If actualCellIndex >= dt.Columns.Count Then
                                dt.Columns.Add("Column" & actualCellIndex.ToString(), autoDetectType.type)
                                columns.Add(autoDetectType)
                            Else
                                dt.Columns(actualCellIndex).DataType = autoDetectType.type
                                columns(actualCellIndex) = autoDetectType
                            End If

                            actualCellIndex += 1
                        Next
                    End If
                    For i As Integer = startIndex To rows.Count - 1
                        Dim row As DataRow = dt.NewRow()
                        Dim cells As XmlNodeList = rows(i).SelectNodes("ss:Cell", nsmgr)
                        Dim actualCellIndex As Integer = 0
                        For cellIndex As Integer = 0 To cells.Count - 1
                            Dim cell As System.Xml.XmlNode = cells(cellIndex)
                            If cell.Attributes("ss:Index") IsNot Nothing Then
                                actualCellIndex = Integer.Parse(cell.Attributes("ss:Index").Value) - 1
                            End If

                            Dim data As System.Xml.XmlNode = cell.SelectSingleNode("ss:Data", nsmgr)

                            If actualCellIndex >= dt.Columns.Count Then
                                For j As Integer = dt.Columns.Count To actualCellIndex - 1
                                    dt.Columns.Add("Column" & actualCellIndex.ToString(), GetType(String))
                                    columns.Add(getDefaultType())
                                Next
                                Dim autoDetectType As ColumnType = [getType](cell.SelectSingleNode("ss:Data", nsmgr))
                                dt.Columns.Add("Column" & actualCellIndex.ToString(), GetType(String))
                                columns.Add(autoDetectType)
                            End If
                            If data IsNot Nothing Then
                                row(actualCellIndex) = data.InnerText
                            End If

                            actualCellIndex += 1
                        Next

                        dt.Rows.Add(row)
                    Next
                End If
            Next
            Return ds
        End Function

        Private Shared Function getDefaultType() As ColumnType
            Return New ColumnType(GetType([String]))
        End Function

        Private Overloads Shared Function [getType](ByVal data As System.Xml.XmlNode) As ColumnType
            Dim type As String = Nothing
            If data.Attributes("ss:Type") Is Nothing OrElse data.Attributes("ss:Type").Value Is Nothing Then
                type = ""
            Else
                type = data.Attributes("ss:Type").Value
            End If

            Select Case type
                Case "DateTime"
                    Return New ColumnType(GetType(DateTime))
                Case "Boolean"
                    Return New ColumnType(GetType([Boolean]))
                Case "Number"
                    Return New ColumnType(GetType([Decimal]))
                Case ""
                    Dim test2 As Decimal
                    If data Is Nothing OrElse [String].IsNullOrEmpty(data.InnerText) OrElse Decimal.TryParse(data.InnerText, test2) Then
                        Return New ColumnType(GetType([Decimal]))
                    Else
                        Return New ColumnType(GetType([String]))
                    End If
                Case Else
                    '"String"
                    Return New ColumnType(GetType([String]))
            End Select
        End Function
    End Class
    '***********END - EXCEL FILE OPERATIONS******************

#Region "PDF Operations"
    '***********START - PDF FILE OPERATIONS******************
    Public Class clsPDFOperations
        Shared AcroApp As CAcroApp
        Shared PDDoc As CAcroPDDoc
        Shared AVDoc As CAcroAVDoc

        Public Shared Sub OpenPDF(Optional ByVal Visible As Boolean = False)
            Try
                AcroApp = CreateObject("AcroExch.App", "")
                PDDoc = CreateObject("AcroExch.PDDoc", "")
                AVDoc = CreateObject("AcroExch.AVDoc", "")
                If Visible Then
                    AcroApp.Show()
                Else
                    AcroApp.Hide()
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Sub ClosePDF()
            Try
                For Each P As Process In Process.GetProcessesByName("Acrobat")
                    P.Kill()
                Next
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Function PDFPageCount(ByVal PDFFilePath As String) As Long
            Dim TotalPages As Long = 0
            Try
                Dim PDDoc As CAcroPDDoc

                PDDoc = CreateObject("AcroExch.PDDoc", "")

                Dim StrTemp As String = PDFFilePath

                If PDDoc.Open(StrTemp) Then
                    TotalPages = PDDoc.GetNumPages
                    PDDoc.Close()
                    PDDoc = Nothing
                End If
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return TotalPages
        End Function

        Public Shared Function PDFWordsCount(ByVal PDFFilePath As String, ByVal PageNo As Int16) As Long
            Dim TotalWords As Double = 0
            Dim jso
            Try
                Dim PDDoc As CAcroPDDoc

                PDDoc = CreateObject("AcroExch.PDDoc", "")

                Dim StrTemp As String = PDFFilePath

                If PDDoc.Open(StrTemp) Then
                    jso = PDDoc.GetJSObject
                    If Not jso Is Nothing Then
                        TotalWords = jso.getPageNumWords(PageNo)
                    End If
                    jso = Nothing
                    PDDoc.Close()
                    PDDoc = Nothing
                End If
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return TotalWords
        End Function

        ''' <summary>
        ''' Write PDF Data into TXT File
        ''' </summary>
        ''' <param name="PDFFilePath"></param>
        ''' <param name="RunFile"></param>
        ''' <remarks></remarks>
        Public Shared Sub WritePDFinTXT(ByVal PDFFilePath As String, Optional ByVal RunFile As Boolean = False)
            Try
                'On Error Resume Next
                'Dim AcroApp As PdfLib.Pdf
                Dim a As AcroPDPage

                'Dim AcroApp As CAcroApp
                'Dim PDDoc As CAcroPDDoc
                'Dim AVDoc As CAcroAVDoc

                Dim X As Double
                Dim Pg As Double
                Dim TempStr
                'acroapp.

                'AcroApp = CreateObject("AcroExch.App", "")
                'PDDoc = CreateObject("AcroExch.PDDoc", "")
                'AVDoc = CreateObject("AcroExch.AVDoc", "")

                'AcroApp.Hide()

                Dim StrTemp As String = PDFFilePath

                Dim bFileOpen = AVDoc.Open(StrTemp, "FILE NAME") 'Boolean

                If PDDoc.Open(StrTemp) Then
                    Dim JSO = PDDoc.GetJSObject

                    JSO.Console.Hide()
                    JSO.Console.Clear()

                    Dim W As System.IO.StreamWriter
                    W = System.IO.File.CreateText(My.Application.Info.DirectoryPath() & "\PDFRead.txt")

                    'MsgBox(AVDoc.FindText("XYY", 0, 0, 0)) 'TO SEARCH WITHIN PDF FILE

                    For Pg = 1 To PDDoc.GetNumPages
                        'JSO = PDDoc.GetJSObject
                        'For X = 0 To JSO.GetPageNumWords 'Get the total number of words found
                        'Label1.Text = Pg
                        W.WriteLine("Page No : " & Pg)
                        Try
                            For X = 0 To 1000 'Get the 10000 of words
                                'Try

                                'Label2.Text = X

                                TempStr = JSO.GetPageNthWord(Pg, X) '(page,word)
                                'Debug.Print(TempStr)
                                If Len(TempStr) > 0 Then
                                    W.Write(TempStr & " ")
                                End If
                                'Catch ex As Exception

                                System.Windows.Forms.Application.DoEvents()

                                'End Try
                            Next X 'Next word
                            W.WriteLine("")
                            System.Windows.Forms.Application.DoEvents()
                        Catch ex As Exception

                        End Try
                    Next
                    W.Close()
                    PDDoc.Close()
                    If RunFile Then
                        Process.Start(My.Application.Info.DirectoryPath() & "\PDFRead.txt")
                    End If
                End If
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub


        ''' <summary>
        ''' Return PDF Data as String When Found "FindWhat" After "AfterWords" Words
        ''' </summary>
        ''' <param name="PDFFilePath"></param>
        ''' <param name="FindWhat"></param>       
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FindLocationInPDF(ByVal PDFFilePath As String, ByVal FindWhat As Byte, Optional ByVal AtPageNo As Long = 0) As String
            Dim Pr As New PdfReader(PDFFilePath)
            Try
                Dim bText() As Byte
                Dim sStr As String = ""

                If AtPageNo <> 0 Then
                    bText = Pr.GetPageContent(AtPageNo)
                    For i As Long = 0 To bText.Length - 1
                        If bText(i) = FindWhat Then
                            'sStr += Chr(bText(i))
                            sStr = "0," & i + 1
                            Return sStr
                        End If
                    Next
                Else
                    For p As Long = 1 To Pr.NumberOfPages
                        bText = Pr.GetPageContent(p)
                        For i As Long = 0 To bText.Length - 1
                            If bText(i) = FindWhat Then
                                'sStr += Chr(bText(i))
                                sStr = p & "," & i + 1
                                Return sStr
                            End If
                        Next
                    Next
                End If
                Pr.Close()
                Pr = Nothing
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            Finally
                Pr.Close()
                Pr = Nothing
            End Try
        End Function

        Public Shared Function ReadALLPDF(ByVal PDFFilePath As String, Optional ByVal AtPageNo As Long = 0) As String
            Dim Pr As New PdfReader(PDFFilePath)
            Try
                Dim bText() As Byte
                Dim sStr As String = ""
                'Dim sChar
                If AtPageNo <> 0 Then
                    bText = Pr.GetPageContent(AtPageNo)
                    sStr = System.Text.ASCIIEncoding.ASCII.GetString(bText)
                    'For i As Long = 0 To bText.Length - 1
                    '    'sChar = bText(i)
                    '    'If VarType(sChar) = vbString Then
                    '    sStr += Chr(bText(i))
                    '    'End If
                    'Next
                Else
                    For p As Long = 1 To Pr.NumberOfPages
                        bText = Pr.GetPageContent(p)
                        sStr += System.Text.ASCIIEncoding.ASCII.GetString(bText)
                        'For i As Long = 0 To bText.Length - 1
                        '    'sChar = Chr(bText(i))
                        '    'If VarType(sChar) = VariantType.Char Then
                        '    sStr += Chr(bText(i))
                        '    'End If
                        'Next
                    Next
                End If
                Pr.Close()
                Pr = Nothing
                Return sStr
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
        Public Shared Function ReadALLPDFPg(ByVal PDFFilePath As String, Optional ByVal AtPageNo As Long = 0) As String
            Dim Pr As New PdfReader(PDFFilePath)
            Try
                Dim bText() As Byte
                Dim sStr As String = ""
                'Dim sChar
                If AtPageNo <> 0 Then
                    bText = Pr.GetPageContent(AtPageNo)
                    'sStr.Insert(0, System.Text.ASCIIEncoding.ASCII.GetString(bText))
                    'sStr = System.Text.ASCIIEncoding.ASCII.GetString(bText)
                    sStr = System.Text.Encoding.GetEncoding("utf-8").GetString(bText)
                    'For i As Long = 0 To bText.Length - 1
                    '    'sChar = bText(i)
                    '    'If VarType(sChar) = vbString Then
                    '    sStr += Chr(bText(i))
                    '    'End If
                    'Next
                Else
                    For p As Int32 = 1 To Pr.NumberOfPages
                        bText = Pr.GetPageContent(p)
                        sStr += System.Text.Encoding.GetEncoding("utf-8").GetString(bText)
                        'sStr.Insert(p - 1, System.Text.ASCIIEncoding.ASCII.GetString(bText))
                        'ReDim Preserve sStr(p)
                        'sStr(p - 1) = System.Text.ASCIIEncoding.ASCII.GetString(bText)
                        'For i As Long = 0 To bText.Length - 1
                        '    'sChar = Chr(bText(i))
                        '    'If VarType(sChar) = VariantType.Char Then
                        '    sStr += Chr(bText(i))
                        '    'End If
                        'Next
                    Next
                End If
                Pr.Close()
                Pr = Nothing
                Return sStr
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Dim app, oDialog
        Dim fname, gPDFPath, numOpenPDFs, bExisting
        Dim InputText, dateTimeStart, dateTimeEnd, timeSpan
        Public Shared Function FindWordJSO(ByVal PDFFilePath, ByVal InputText) As Object
            Dim bStop
            Dim jso, nCount, i, j
            Dim word, result, foundErr, nPages, nWords
            Dim rc, str_Renamed
            Try
                ' ** get JavaScript Object
                ' ** note jso is related to PDDoc of a PDF,

                If PDDoc.Open(PDFFilePath) Then
                    jso = PDDoc.GetJSObject
                    nCount = 0
                    bStop = False
                    ' ** search for the text
                    If Not jso Is Nothing Then
                        MsgBox("Searching ... ")
                        ' ** total number of pages
                        nPages = jso.numPages
                        ' ** Go through pages
                        For i = 0 To nPages - 1
                            ' ** check each word in a page
                            nWords = jso.getPageNumWords(i)
                            For j = 0 To nWords - 1
                                ' ** get a word
                                word = jso.getPageNthWord(i, j)
                                Debug.Print(word)
                                If VarType(word) = vbString Then
                                    ' ** compare the word with what the user wants
                                    result = StrComp(word, InputText, vbTextCompare)

                                    ' ** if same
                                    If result = 0 Then
                                        nCount = nCount + 1
                                        rc = jso.selectPageNthWord(i, j)
                                        MsgBox("# " & nCount & " found in page " & (i + 1))
                                    End If
                                End If
                            Next
                        Next
                    End If
                    FindWordJSO = nCount
                    jso = Nothing
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Function
        Public Shared Function FindWordsJSO(ByVal PDFFilePath As String, ByVal PageNo As Int16, ByVal WordNo As Int16) As Object
            Dim bStop
            Dim jso, nCount, i, j
            Dim word, result, foundErr, nPages, nWords
            Dim rc, str_Renamed
            Try
                ' ** get JavaScript Object
                ' ** note jso is related to PDDoc of a PDF,

                If PDDoc.Open(PDFFilePath) Then
                    jso = PDDoc.GetJSObject
                    nCount = 0
                    bStop = False
                    ' ** search for the text
                    If Not jso Is Nothing Then
                        'MsgBox("Searching ... ")
                        ' ** total number of pages
                        nPages = jso.numPages
                        ' ** Go through pages
                        'For i = 0 To nPages - 1
                        ' ** check each word in a page                        
                        nWords = jso.getPageNumWords(PageNo - 1)
                        If Len(nWords) >= WordNo Then
                            nWords = WordNo
                        End If
                        word = ""
                        For j = 0 To nWords - 1
                            ' ** get a word
                            word += jso.getPageNthWord(i, j) & " "
                            ''Debug.Print(word)
                            'If VarType(word) = vbString Then
                            '    ' ** compare the word with what the user wants
                            '    result = StrComp(word, InputText, vbTextCompare)

                            '    ' ** if same
                            '    If result = 0 Then
                            '        nCount = nCount + 1
                            '        rc = jso.selectPageNthWord(i, j)
                            '        MsgBox("# " & nCount & " found in page " & (i + 1))
                            '    End If
                            'End If
                        Next
                        'Next
                    End If
                    FindWordsJSO = word
                    jso = Nothing
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Function

        Public Shared Sub ListFieldNames(ByVal pdfTemplate As String)
            'Dim pdfTemplate As String = "D:\Narender\Projects\ASP.NET\2005\EbookPdfDisplay\Books\817179808X\Chapter wise Pdf\Chapter-01_History Taking.pdf"

            ' title the form
            'Me.Text += " - " + pdfTemplate

            ' create a new PDF reader based on the PDF template document
            Dim pdfReader As PdfReader = New PdfReader(pdfTemplate)

            ' create and populate a string builder with each of the 
            ' field names available in the subject PDF
            Dim sb As New StringBuilder()

            Dim de As New DictionaryEntry
            For Each de In pdfReader.AcroFields.Fields
                sb.Append(de.Key.ToString() + Environment.NewLine)
            Next

            ' Write the string builder's content to the form's textbox
            'textBox1.Text = sb.ToString()
            'textBox1.SelectionStart = 0
            MsgBox(sb.ToString())
        End Sub

#Region "PDF Extraction Area"


        Public Shared Sub RemovePassword(ByVal sourcePdf As String, ByVal fromPageNum As Integer, ByVal toPageNum As Integer, ByVal outPdf As String)
            Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
            Dim doc As iTextSharp.text.Document = Nothing
            Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
            Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
            If fromPageNum = 0 Then
                fromPageNum = 1
            End If
            Try
                Dim enc As System.Text.Encoding = System.Text.Encoding.ASCII
                Dim myPasswordArray As Byte() = enc.GetBytes("ashishsameer")
                reader = New iTextSharp.text.pdf.PdfReader(sourcePdf, myPasswordArray)
                If toPageNum = 0 Then
                    toPageNum = reader.NumberOfPages
                End If
                doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
                Try
                    IO.Directory.CreateDirectory(Mid(outPdf, 1, outPdf.LastIndexOf("\")))
                Catch ex As Exception

                End Try
                pdfCpy = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outPdf, IO.FileMode.OpenOrCreate))
                doc.Open()
                For i As Integer = fromPageNum To toPageNum Step 1
                    page = pdfCpy.GetImportedPage(reader, i)
                    pdfCpy.AddPage(page)
                Next
                doc.Close()
                reader.Close()
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        Public Shared Sub AssignPassword(ByVal sourcePdf As String, ByVal fromPageNum As Integer, ByVal toPageNum As Integer, ByVal outPdf As String)
            Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
            Dim doc As iTextSharp.text.Document = Nothing
            Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
            Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
            If fromPageNum = 0 Then
                fromPageNum = 1
            End If
            Try
                'Dim enc As System.Text.Encoding = System.Text.Encoding.ASCII
                'Dim myPasswordArray As Byte() = enc.GetBytes("ashishsameer")
                reader = New iTextSharp.text.pdf.PdfReader(sourcePdf) ', myPasswordArray)
                If toPageNum = 0 Then
                    toPageNum = reader.NumberOfPages
                End If
                doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
                Try
                    IO.Directory.CreateDirectory(Mid(outPdf, 1, outPdf.LastIndexOf("\")))
                Catch ex As Exception

                End Try
                pdfCpy = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outPdf, IO.FileMode.OpenOrCreate))
                doc.Open()
                For i As Integer = fromPageNum To toPageNum Step 1
                    page = pdfCpy.GetImportedPage(reader, i)
                    pdfCpy.AddPage(page)
                Next
                doc.Close()
                reader.Close()
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        Public Shared Function ParsePdfText(ByVal sourcePDF As String, Optional ByVal PageNum As Integer = 0) As String
            Dim sb As New System.Text.StringBuilder()
            Try
                Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePDF)
                Dim pageBytes() As Byte = Nothing
                Dim token As iTextSharp.text.pdf.PRTokeniser = Nothing
                Dim tknType As Integer = -1
                Dim tknValue As String = String.Empty
                ''Dim objRegx As Regex = Nothing
                ' ''objRegx = New Regex("@Payments")
                ''objRegx = New Regex("@")
                Dim fromPageNum As Int16 = 0
                Dim toPageNum As Int16 = 0

                If PageNum = 0 Then
                    fromPageNum = 1
                    toPageNum = reader.NumberOfPages
                Else
                    fromPageNum = PageNum
                    toPageNum = PageNum
                End If

                If fromPageNum > toPageNum Then
                    Throw New ApplicationException("Parameter error: The value of fromPageNum can " & _
                                               "not be larger than the value of toPageNum")
                End If
                Dim intPagr As Integer = 0
                Dim pages As New StringBuilder()

                For i As Integer = fromPageNum To toPageNum Step 1
                    pageBytes = reader.GetPageContent(i)

                    If Not IsNothing(pageBytes) Then
                        token = New iTextSharp.text.pdf.PRTokeniser(pageBytes)
                        Dim strInBldr As New StringBuilder()
                        While token.NextToken()
                            tknType = token.TokenType()
                            tknValue = token.StringValue
                            If tknType = iTextSharp.text.pdf.PRTokeniser.TK_STRING Then
                                strInBldr.Append(token.StringValue)
                                'MyCLS.clsFileHandling.WriteFile(tknType & vbTab & ": " & tknValue)
                                'ElseIf tknType = 1 AndAlso tknValue = "-600" Then
                                '    strInBldr.Append(" ")
                                'ElseIf tknType = 10 AndAlso tknValue = "TJ" Then
                                '    strInBldr.Append(" ")
                            End If
                        End While

                        ''If strInBldr.ToString().ToLower().Contains(Keyword.Trim(" ").ToLower()) And Keyword.ToLower().Length > 2 Then
                        pages.Append(strInBldr.ToString())
                        ''End If
                    End If
                Next i
                sb.Append(pages.ToString())
            Catch ex As Exception
                Return String.Empty
            End Try

            Return sb.ToString()
        End Function

        Public Shared Function ExtractText(ByVal sStr As String) As String
            Dim strEx As String = ""
            Dim strExAll As String = ""
            'Try
            Dim ST As Int32 = 0
            Dim ET As Int32 = 0
            'REPLACE THE TEXT
            ReplaceText(sStr)

            '*1
            'ST = InStr(1, sStr, "TD (", CompareMethod.Binary)
            'ET = InStr(ST + 1, sStr, ")Tj 0.0016", CompareMethod.Binary)
            'MsgBox(Mid(sStr, ST + 4, ET - ST - 4))
            '*2
            For i As Int32 = 1 To sStr.Length
                ST = InStr(i, sStr, "Tm (", CompareMethod.Binary)
                If ST = 0 Then
                    ST = IIf(InStr(i, sStr, "T* (", CompareMethod.Binary) > InStr(i, sStr, "TD (", CompareMethod.Binary), InStr(i, sStr, "TD (", CompareMethod.Binary), InStr(i, sStr, "T* (", CompareMethod.Binary))
                    If ST = 0 Then
                        ST = InStr(i, sStr, "Td (", CompareMethod.Binary)
                    Else
                        ST = IIf(ST > InStr(i, sStr, "Td (", CompareMethod.Binary), InStr(i, sStr, "Td (", CompareMethod.Binary), ST)
                    End If
                    If ST = 0 Then
                        Exit For
                    End If
                End If
                ET = InStr(ST + 1, sStr, ")Tj", CompareMethod.Binary)

                strEx = Mid(sStr, ST + 4, ET - ST - 4) '& vbCrLf & "***" & vbCrLf
                strExAll += Mid(sStr, ST + 4, ET - ST - 4) '& vbCrLf & "***" & vbCrLf
                'MsgBox(strEx)                
                i = ET
            Next
            '*3

            '*
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            'End Try
            Return strExAll
        End Function

        Private Structure sStrFindReplace
            Public Index As Int32
            Public ToFind As String
            Public ToReplace As String
        End Structure

        Public Shared Sub ReplaceText(ByRef sStr As String)
            'Try
            Dim sFR As New sStrFindReplace
            Dim sFRList As New List(Of sStrFindReplace)

            sFR.Index = 0
            sFR.ToFind = "\221"
            sFR.ToReplace = "'"
            sFRList.Add(sFR)

            sFR.Index = 1
            sFR.ToFind = "\222"
            sFR.ToReplace = "'"
            sFRList.Add(sFR)

            sFR.Index = 1
            sFR.ToFind = "\" & Chr(13)
            sFR.ToReplace = ""
            sFRList.Add(sFR)

            sFR.Index = 1
            sFR.ToFind = "CHAPTER"
            sFR.ToReplace = ""
            sFRList.Add(sFR)

            For i As Int32 = 0 To sFRList.Count - 1
                sStr = Replace(sStr, sFRList.Item(i).ToFind, sFRList.Item(i).ToReplace)
            Next
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            'End Try
        End Sub
#End Region
    End Class
    '***********END - PDF FILE OPERATIONS******************
#End Region

#Region "DOC Operations"
    '***********START - DOC FILE OPERATIONS******************
    Public Class clsDOCOperations
        Shared w As Application
        Shared doc As Document
        Public Shared MatchCase As Object
        Public Shared MatchWholeWord As Object
        Public Shared MatchWildCard As Object
        Public Shared MatchSoundLike As Object
        Public Shared MatchAllWordsForm As Object
        Public Shared Forward As Object
        Public Shared Wrap As Object
        Public Shared Format As Object


        ''' <summary>
        ''' Open Doc File
        ''' </summary>
        ''' <param name="DocFile"></param>
        ''' <param name="Visible"></param>
        ''' <param name="SaveAs"></param>
        ''' <remarks></remarks>
        Public Shared Sub OpenDOC(ByVal DocFile As String, Optional ByVal Visible As Boolean = False, Optional ByVal SaveAs As Boolean = True)
            Try
                w = New Application
                doc = New Document

                doc = w.Documents.Open(DocFile)
                If SaveAs Then
                    doc.SaveAs("CL_" & Date.Now.Year & Date.Now.Month & Date.Now.Day & Date.Now.Hour & Date.Now.Minute)
                End If

                If Visible Then
                    w.Visible = True
                    w.WindowState = WdWindowState.wdWindowStateMaximize
                Else
                    w.Visible = False
                End If
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' Close Doc File
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CloseDOC()
            Try
                doc = Nothing
                w = Nothing
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' Replace text within Doc File
        ''' </summary>
        ''' <param name="strFind"></param>
        ''' <param name="strReplace"></param>
        ''' <remarks></remarks>
        Public Shared Sub DocReplace(ByVal strFind As String, ByVal strReplace As String)
            Try
                doc.Content.Find.Execute(strFind, , , , , , , , , strReplace)
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub

        ''' <summary>
        ''' Get Complete Text From Doc File
        ''' </summary>       
        ''' <remarks></remarks>
        Public Shared Function GetDocData() As String
            Try
                Return doc.Content.Text.ToString()
            Catch ex As Exception
                Return ""
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function
    End Class
    '***********END - DOC FILE OPERATIONS******************
#End Region

#Region "Quotes Fixing in Query"
    '**********START - QUOTE FIXING IN SQL QUERIES********************
    Public Class clsQuotesFixInQuery
        Shared strQBlank As String = ""

        ''' <summary>
        ''' Fixes Quotes problems in Sql Queries
        ''' </summary>
        ''' <param name="strQ"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FixQuotesInQuery(ByVal strQ As String) As String
            Try
                Dim QryType As String = ""
                strQBlank = strQ
                fnQuotesBlanksReplace(strQ)
                strQ = strQBlank
                Dim strQWithQuotes As String = strQ
                Dim strQWithoutQuotes As String = strQ
                Dim Ext As String = ""
                Dim Ext2Replace As String = ""
                Dim i As Int16

                QryType = UCase(Mid(strQ, 1, 6))

                If QryType = "INSERT" Then
                    For i = 1 To strQWithQuotes.Length
                        Ext = Mid(strQWithQuotes, i, 1)
                        If InStr(Ext, "'") > 0 Then
                            Ext = Mid(strQWithQuotes, i - 1, 3)
                            If InStr(Ext, "',") > 0 Or InStr(Ext, ",'") > 0 Or InStr(Ext, "('") > 0 Or InStr(Ext, "')") > 0 Or InStr(Ext, "'' ") Then

                            Else
                                Ext2Replace = Replace(Ext, "'", "''")
                                strQWithoutQuotes = Replace(strQWithoutQuotes, Ext, Ext2Replace)
                                Ext2Replace = ""
                            End If
                        End If
                    Next
                    strQ = strQWithoutQuotes
                ElseIf QryType = "UPDATE" Then
                    For i = 1 To strQWithQuotes.Length
                        Ext = Mid(strQWithQuotes, i, 1)
                        If InStr(Ext, "'") > 0 Then
                            Ext = Mid(strQWithQuotes, i - 1, 3)
                            If InStr(Ext, "',") > 0 Or InStr(Ext, ",'") > 0 Or InStr(Ext, "='") > 0 Or InStr(Ext, "' ") > 0 Or InStr(Ext, " '") > 0 Then

                            Else
                                If Len(Ext) > 2 Then
                                    Ext2Replace = Replace(Ext, "'", "''")
                                    strQWithoutQuotes = Replace(strQWithoutQuotes, Ext, Ext2Replace)
                                    Ext2Replace = ""
                                End If
                            End If
                        End If
                    Next
                    strQ = strQWithoutQuotes
                Else
                    strQ = strQWithoutQuotes
                End If
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            'Clipboard.SetText(strQ)
            'MsgBox(strQ)
            Return strQ
        End Function

        Private Shared Sub fnQuotesBlanksReplace(ByVal strQ As String)
            Try
                'If (InStr(strQ, "' ") > 0 Or InStr(strQ, " '") > 0) And (InStr(strQ, "' Where") = 0) Then
                '    strQBlank = Replace(Replace(Replace(strQ, "' ", "'"), " '", ""), "'Where", "' Where")
                '    'strQBlank = Replace(strQBlank, " '", "'")
                '    fnQuotesBlanksReplace(strQBlank)
                'End If

                'If InStr(strQ, " ,") > 0 Or InStr(strQ, ", ") > 0 Then
                '    strQBlank = Replace(Replace(strQ, " ,", ","), ", ", "")
                '    fnQuotesBlanksReplace(strQBlank)
                'End If
            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
    End Class
    '**********END - QUOTE FIXING IN SQL QUERIES********************
#End Region

    '***********HTML to RTF Converter***********************************
    Public Class clsRTFtoHTML

#Region "Private Members"

        ' A RichTextBox control to use to help with parsing.
        Private _rtfSource As New System.Windows.Forms.RichTextBox

#End Region

#Region "Read/Write Properties"

        ''' <summary>
        ''' Returns/Sets The RTF formatted text to parse
        ''' </summary>
        Public Property rtf() As String
            Get
                Return _rtfSource.Rtf
            End Get
            Set(ByVal value As String)
                _rtfSource.Rtf = value
            End Set
        End Property

#End Region

#Region "ReadOnly Properties"

        ''' <summary>
        ''' Returns the HTML code for the provided RTF
        ''' </summary>
        Public ReadOnly Property html() As String
            Get
                Return GetHtml()
            End Get
        End Property

#End Region

#Region "Private Functions"

        ''' <summary>
        ''' Returns an HTML Formated Color string for the style from a system.drawing.color
        ''' </summary>
        ''' <param name="clr">The color you wish to convert</param>
        Private Function HtmlColorFromColor(ByRef clr As System.Drawing.Color) As String
            Dim strReturn As String = ""
            If clr.IsNamedColor Then
                strReturn = clr.Name.ToLower
            Else
                strReturn = clr.Name
                If strReturn.Length > 6 Then
                    strReturn = strReturn.Substring(strReturn.Length - 6, 6)
                End If
                strReturn = "#" & strReturn
            End If
            Return strReturn
        End Function

        ''' <summary>
        ''' Provides the font style per given font
        ''' </summary>
        ''' <param name="fnt">The font you wish to convert</param>
        Private Function HtmlFontStyleFromFont(ByRef fnt As System.Drawing.Font) As String
            Dim strReturn As String = ""
            'style
            If fnt.Italic Then
                strReturn &= "italic "
            Else
                strReturn &= "normal "
            End If
            'variant
            strReturn &= "normal "
            'weight
            If fnt.Bold Then
                strReturn &= "bold "
            Else
                strReturn &= "normal "
            End If
            'size
            strReturn &= fnt.SizeInPoints & "pt/normal "
            'family
            strReturn &= fnt.FontFamily.Name
            Return strReturn
        End Function

        ''' <summary>
        ''' Parses the given rich text and returns the html.
        ''' </summary>
        Private Function GetHtml() As String
            Dim strReturn As String = "<div>"
            Dim clrForeColor As System.Drawing.Color = Color.Black
            Dim clrBackColor As System.Drawing.Color = Color.Black
            Dim fntCurrentFont As System.Drawing.Font = _rtfSource.Font
            Dim altCurrent As System.Windows.Forms.HorizontalAlignment = HorizontalAlignment.Left
            Dim intPos As Integer = 0
            For intPos = 0 To _rtfSource.Text.Length - 1
                _rtfSource.Select(intPos, 1)
                'Forecolor
                If intPos = 0 Then
                    strReturn &= "<span style=""color:" & HtmlColorFromColor(_rtfSource.SelectionColor) & """>"
                    clrForeColor = _rtfSource.SelectionColor
                Else
                    If _rtfSource.SelectionColor <> clrForeColor Then
                        strReturn &= "</span>"
                        strReturn &= "<span style=""color:" & HtmlColorFromColor(_rtfSource.SelectionColor) & """>"
                        clrForeColor = _rtfSource.SelectionColor
                    End If
                End If
                'Background color
                If intPos = 0 Then
                    strReturn &= "<span style=""background-color:" & HtmlColorFromColor(_rtfSource.SelectionBackColor) & """>"
                    clrBackColor = _rtfSource.SelectionBackColor
                Else
                    If _rtfSource.SelectionBackColor <> clrBackColor Then
                        strReturn &= "</span>"
                        strReturn &= "<span style=""background-color:" & HtmlColorFromColor(_rtfSource.SelectionBackColor) & """>"
                        clrBackColor = _rtfSource.SelectionBackColor
                    End If
                End If
                'Font
                If intPos = 0 Then
                    strReturn &= "<span style=""font:" & HtmlFontStyleFromFont(_rtfSource.SelectionFont) & """>"
                    fntCurrentFont = _rtfSource.SelectionFont
                Else
                    If _rtfSource.SelectionFont.GetHashCode <> fntCurrentFont.GetHashCode Then
                        strReturn &= "</span>"
                        strReturn &= "<span style=""font:" & HtmlFontStyleFromFont(_rtfSource.SelectionFont) & """>"
                        fntCurrentFont = _rtfSource.SelectionFont
                    End If
                End If
                'Alignment
                If intPos = 0 Then
                    strReturn &= "<p style=""text-align:" & _rtfSource.SelectionAlignment.ToString & """>"
                    altCurrent = _rtfSource.SelectionAlignment
                Else
                    If _rtfSource.SelectionAlignment <> altCurrent Then
                        strReturn &= "</p>"
                        strReturn &= "<p style=""text-align:" & _rtfSource.SelectionAlignment.ToString & """>"
                        altCurrent = _rtfSource.SelectionAlignment
                    End If
                End If
                strReturn &= _rtfSource.Text.Substring(intPos, 1)
            Next
            'close all the spans
            strReturn &= "</span>"
            strReturn &= "</span>"
            strReturn &= "</span>"
            strReturn &= "</p>"
            strReturn &= "</div>"
            strReturn = strReturn.Replace(Convert.ToChar(10), "<br />")
            Return strReturn
        End Function

#End Region

    End Class
    '***********END HTML to RTF Converter***********************************

    '***********START XML OPERATIONS****************************************
    Public Class clsXMLOperations

        Public Enum XMLFormat
            ThroughXmlReader = 1
            ThroughSqlDataAdapter = 2
        End Enum

        Public Shared Sub WriteXMLFromSQL(ByVal SQLQuery As String, ByVal XMLFilePath As String, ByVal Format As XMLFormat)
            Try
                Dim mySqlCommand As SqlCommand = New SqlCommand(SQLQuery, MyConSql)
                mySqlCommand.CommandTimeout = 15

                If Format = XMLFormat.ThroughXmlReader Then
                    ' Now create the DataSet and fill it with xml data.
                    MyDs = New DataSet()
                    MyDs.ReadXml(mySqlCommand.ExecuteXmlReader(), XmlReadMode.Fragment)

                    '' Modify to match the other dataset
                    'MyDs.DataSetName = "NewDataSet"
                ElseIf Format = XMLFormat.ThroughSqlDataAdapter Then
                    ' Get the same data through the provider.
                    Dim mySqlDataAdapter As SqlDataAdapter = New SqlDataAdapter(SQLQuery, MyConSql)
                    MyDs = New DataSet()
                    mySqlDataAdapter.Fill(MyDs)
                End If
                ' Write data to files: data1.xml and data2.xml.
                MyDs.WriteXml(XMLFilePath)
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Sub ReadXML(ByVal XMLFilePath As String)
            Try
                ' Create an isntance of XmlTextReader and call Read method to read the(file)
                Dim textReader As XmlTextReader = New XmlTextReader(XMLFilePath)
                textReader.Read()
                ' If the node has value
                If textReader.HasValue Then
                    ' Move to fist element
                    textReader.MoveToElement()
                    Console.WriteLine("XmlTextReader Properties Test")
                    Console.WriteLine("===================")
                    ' Read this element's properties and display them on console
                    Console.WriteLine("Name:" + textReader.Name)
                    Console.WriteLine("Base URI:" + textReader.BaseURI)
                    Console.WriteLine("Local Name:" + textReader.LocalName)
                    Console.WriteLine("Attribute Count:" + textReader.AttributeCount.ToString())
                    Console.WriteLine("Depth:" + textReader.Depth.ToString())
                    Console.WriteLine("Line Number:" + textReader.LineNumber.ToString())
                    Console.WriteLine("Node Type:" + textReader.NodeType.ToString())
                    Console.WriteLine("Attribute Count:" + textReader.Value.ToString())

                    ' Move to fist element
                    textReader.MoveToNextAttribute()

                    ' Read this element's properties and display them on console
                    Console.WriteLine("Name:" + textReader.Name)
                    Console.WriteLine("Base URI:" + textReader.BaseURI)
                    Console.WriteLine("Local Name:" + textReader.LocalName)
                    Console.WriteLine("Attribute Count:" + textReader.AttributeCount.ToString())
                    Console.WriteLine("Depth:" + textReader.Depth.ToString())
                    Console.WriteLine("Line Number:" + textReader.LineNumber.ToString())
                    Console.WriteLine("Node Type:" + textReader.NodeType.ToString())
                    Console.WriteLine("Attribute Count:" + textReader.Value.ToString())
                End If
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' Transform Xml into a XML Style Document as Defined in XSL file
        ''' </summary>
        ''' <param name="sXmlPath"></param>
        ''' <param name="sXslPath"></param>
        ''' <remarks></remarks>
        Public Shared Sub TransformXML(ByVal sXmlPath As String, ByVal sXslPath As String, ByVal WriteXMLPath As String)
            Try

                'load the Xml doc 
                Dim myXPathDoc As New XPathDocument(sXmlPath)

                Dim myXslTrans As New System.Xml.Xsl.XslTransform()

                'load the Xsl 
                myXslTrans.Load(sXslPath)

                'create the output stream 
                Dim myWriter As New XmlTextWriter(WriteXMLPath, System.Text.Encoding.UTF8)

                'do the actual transform of Xml 
                myXslTrans.Transform(myXPathDoc, Nothing, myWriter)

                myWriter.Close()
            Catch e As Exception
                Console.WriteLine("Exception: {0}", e.ToString())
            End Try
        End Sub
    End Class
    '***********END XML OPERATIONS******************************************

#Region "Image Processing"
    Public Class clsImaging

        Public Shared imgBitmap_MEMORY As Bitmap

        Public Shared Function PictureBoxToByteArray(ByVal oPictureBox As PictureBox) As Byte()
            Dim oStream As New MemoryStream
            Try
                Dim bmp As New Bitmap(oPictureBox.Image)

                bmp.Save(oStream, Imaging.ImageFormat.Bmp)
                PictureBoxToByteArray = oStream.ToArray
                bmp.Dispose()
                oStream.Close()
            Catch ex As Exception
                PictureBoxToByteArray = oStream.ToArray
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Function

        Public Shared Sub ByteArray2Image(ByRef NewImage As System.Drawing.Image, ByVal ByteArr() As Byte)
            Dim ImageStream As MemoryStream
            Try
                If ByteArr.GetUpperBound(0) > 0 Then
                    ImageStream = New MemoryStream(ByteArr)
                    NewImage = System.Drawing.Image.FromStream(ImageStream)
                Else
                    NewImage = Nothing
                End If
            Catch ex As Exception
                NewImage = Nothing
                clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
        End Sub
    End Class
#End Region


#Region "Insert Multiple Rows"

    '*****STORED PROCEDURE CODE*********MODIFY ACCORDINGLY*************
    '    ALTER PROC [dbo].[SP_Insert] (@x xml) 
    'AS BEGIN 
    '    INSERT INTO TblName (Name, Color) 
    '    SELECT  
    '        row.col.value( 'Name[1]', 'VARCHAR(20)' ) AS Name, 
    '        row.col.value( 'Color[1]', 'VARCHAR(20)' ) AS Color 
    '    FROM @x.nodes('cars/car') row(col)
    'END 
    '******KIND OF STRING TO BE PASSED TO SP*****
    '<cars><car><Name>BMW</Name><Color>Red</Color></car><car><Name>Audi</Name><Color>Green</Color></car></cars>
    '*******************************************************************

    Public Class XMLItemList
        Private sb As System.Text.StringBuilder

        Public Sub New()
            sb = New System.Text.StringBuilder
            sb.Append("<items>" & vbCrLf)
        End Sub

        Public Sub AddItem(ByVal Item As String)
            sb.AppendFormat("<item id={0}{1}{2}></item>{3}", Chr(34), Item, Chr(34), vbCrLf)
        End Sub

        Public Overrides Function ToString() As String
            sb.Append("</items>" & vbCrLf)
            Return sb.ToString
        End Function
    End Class


    Public Class XMLItemListManyCols
        Private sb As System.Text.StringBuilder

        Public Sub New()
            sb = New System.Text.StringBuilder
            sb.Append("<rows>" & vbCrLf)
        End Sub

        Public Sub RowBegin(ByVal RowName As String)
            '<cars><car><Name>BMW</Name><Color>Red</Color></car><car><Name>Audi</Name><Color>Green</Color></car></cars>
            sb.AppendFormat("<{0}>", RowName, vbCrLf)
        End Sub

        Public Sub AddItem(ByVal ColName As String, ByVal Item As String)
            '<cars><car><Name>BMW</Name><Color>Red</Color></car><car><Name>Audi</Name><Color>Green</Color></car></cars>
            sb.AppendFormat("<{0}>{1}</{0}>", ColName, Item, vbCrLf)            
        End Sub

        Public Sub RowEnd(ByVal RowName As String)
            '<cars><car><Name>BMW</Name><Color>Red</Color></car><car><Name>Audi</Name><Color>Green</Color></car></cars>
            sb.AppendFormat("</{0}>", RowName, vbCrLf)
        End Sub

        Public Overrides Function ToString() As String
            sb.Append("</rows>" & vbCrLf)
            Return sb.ToString
        End Function
    End Class
#End Region
End Class