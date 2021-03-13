Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports Acrobat
Imports MyTool.NDS.LIB
Imports MyTool.NDS.DAL
Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Xsl

Imports iTextSharp
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.xml
Imports System.Text

Imports System.IO
Imports System.Drawing


'Imports System.Management '.SqlServer.Management.Common
Imports Microsoft.SqlServer '.Management .Smo


Public Class frmRnD

    Private Sub cmdExecuteScripts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExecuteScripts.Click
        'Dim sqlConnectionString As String = "Data Source=(local);Initial Catalog=AdventureWorks;Integrated Security=True"
        'Dim file As New FileInfo("C:\myscript.sql")
        'Dim script As String = file.OpenText().ReadToEnd()
        'Dim conn As New SqlConnection(sqlConnectionString)
        'Dim server As New Server(New ServerConnection(conn))
        'server.ConnectionContext.ExecuteNonQuery(script)
    End Sub

    Private Sub frmRnD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load        
        'Me.MdiParent = MDI

        '' ''MyCLS.strConnStringSQLCLIENT = "Initial Catalog=JPData_MyShop;Data Source=tsi_dev_02;UID=sa;PWD=sa123;"
        '' ''MyCLS.strConnStringOLEDB = "Initial Catalog=JPData_MyShop;Data Source=tsi_dev_02;UID=sa;PWD=sa123;Provider=SQLOLEDB.1;"
        '' ''MyCLS.clsCOMMON.ConOpen(False)


        '**JP JOURNALS USERNAME - WHILE PURCHASING**
        '' ''Dim sBody As String = "1    User : Narender    Transaction ID : 26    Issue : Journal of Ultrasound in Obstetrics and Gynecology    Topic : The Role of 3D Ultrasound and 3D Power Doppler Imaging in the Diagnosis and Evaluation of Ovarian Cancer: New Perspectives    Topic Heading : The Role of 3D Ultrasound and 3D Power Doppler Imaging in the Diagnosis and Evaluation of Ovarian Cancer    Year : 2007    Month : April-June    Volume : 1             Number : 2    Pages : 38-41    Total Pages : 88    Author : MT Redondo, I Orensanz, FJ Salazar, S Iniesta, B Bueno, T Perez-Medina , JM Bajo    Price : 300    Status : Error    Date & Time : 5/5/2009 4:26:17 AM    "
        ' '' ''sBody = sBody & "    User : " & "UserName"
        ' '' ''sBody = sBody & vbCrLf & "  Transaction ID : " & "TID"

        '' ''Dim UN As String = (Mid(sBody, InStr(sBody, "User :"), InStr(sBody, "ID")))
        '' ''UN = Replace(UN, "User : ", "")
        '' ''UN = (Mid(UN, 1, InStr(UN, "Transaction") - 1))
        ' '' ''UN = Replace(UN, "Transaction ID : ", "")
        '' ''UN = Replace(UN, vbCrLf, "")
        '' ''UN = Trim(UN)
        '' ''MsgBox(UN)


        '*****
        Dim intL As Integer = Convert.ToInt32(100 * 100)
        Dim pBuff As Byte() = New Byte(intL) {}
        Dim pBptr As System.IntPtr = pBuff(0)


        Call cmdInsertAllTypes_Click(sender, e)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '*******COPY IT TO USE ABOVE FUNCTION - SELECT ALL************
        Try
            Dim objLIBTable_1Listing As New LIBTable_1Listing
            Dim objDALTable_1 As New DALTable_1
            Dim tp As New MyCLS.TransportationPacket

            '    objLIBTable_1Listing(0).a = ""
            '    objLIBTable_1Listing(1).b = ""
            tp = objDALTable_1.GetTable_1Details()
            If tp.MessageId = 1 Then
                objLIBTable_1Listing = tp.MessageResultset
                MsgBox(objLIBTable_1Listing(0))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



        '*******COPY IT TO USE ABOVE FUNCTION - SELECT BY ID************
        Try
            Dim objLIBTable_1Listing As New LIBTable_1Listing
            Dim objDALTable_1 As New DALTable_1
            Dim tp As New MyCLS.TransportationPacket
            tp.MessagePacket = 1    'ID to be Passed

            '    objLIBTable_1Listing(0).a = ""
            '    objLIBTable_1Listing(1).b = ""
            tp = objDALTable_1.GetTable_1Details(tp)
            If tp.MessageId = 1 Then
                objLIBTable_1Listing = tp.MessageResultset
                MsgBox(objLIBTable_1Listing(0))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        '*******COPY IT TO USE ABOVE FUNCTION - INSERT************
        Try
            Dim objLIBTable_1 As New LIBTable_1
            Dim objDALTable_1 As New DALTable_1
            Dim tp As New MyCLS.TransportationPacket

            objLIBTable_1.a = "mala"
            objLIBTable_1.b = "mala"
            tp.MessagePacket = objLIBTable_1
            tp = objDALTable_1.InsertTable_1(tp)

            If tp.MessageId = 1 Then
                Dim strOutParamValues As String() = tp.MessageResultset
                MsgBox(strOutParamValues(0))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MyCLS.clsDOCOperations.OpenDOC("C:\Documents and Settings\DEV2\Desktop\satyam_history.doc", False, False)
        Label1.Text = MyCLS.clsDOCOperations.GetDocData()
        MyCLS.clsDOCOperations.CloseDOC()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            'Call XML1()
            'Call XML2()
            'Call XML3()
            'Call Transform("C:\Documents and Settings\DEV2\Desktop\DOI\XMLData2.xml", "C:\Documents and Settings\DEV2\Desktop\DOI\Book.xsl")            
            'Call CreateTheNavigator()

            'MyCLS.clsXMLOperations.TransformXML("C:\Documents and Settings\DEV2\Desktop\DOI\XMLData2.xml", "C:\Documents and Settings\DEV2\Desktop\DOI\Book.xsl", "C:\Documents and Settings\DEV2\Desktop\DOI\Result.xml")

            '***START - CREATE PATH***
            Dim Path As String
            Path = Replace(System.Windows.Forms.Application.StartupPath, "bin\Debug", "")
            '***END - CREATE PATH***

            '***START - CREATE XMLDATAFILE FROM SQL********************************************************************
            '' ''Dim Qstr As String
            '' ''Qstr = "SELECT [SNO],[TITLE],[ISBN],LTRIM(LEFT([AUTHOR],LEN([AUTHOR])-CHARINDEX(' ',reverse([AUTHOR]),1))) AS AUTHORNAME,LTRIM(RIGHT([AUTHOR],CHARINDEX(' ',reverse([AUTHOR]),1))) AS SURNAME,[NEWRELEASES] " & _
            '' ''        "      ,[FEATUREDTITLE],[BESTSELLER],[eBOOKTYPE],[WITH_CD],[SUBCATEGORY],[PRICE] " & _
            '' ''        "      ,[USPRICE],[DISCOUNT],[SUBTITLE] " & _
            '' ''        "      ,[EDITION],[PUBYEAR],[PUBLISHER],[DISPLAYTEXT],[LANGUAGE],[TYPE],[VOLUME] " & _
            '' ''        "      ,[CATEGORY],[Abstract] " & _
            '' ''        "      ,[authorAf1],[authorAf2],[authorAf3],[authorAf4],[authorAf5],[authorAf6],[authorAf7] " & _
            '' ''        "	  ,[authorAf8],[authorAf9],[authorAf10],[authorAf11],[video] " & _
            '' ''        "  FROM [BOOK]"

            '' ''MyCLS.strConnStringSQLCLIENT = "Initial Catalog=EBook_NEW;Data Source=tsi_dev_02;UID=sa;PWD=sa123;"
            '' ''MyCLS.clsCOMMON.ConOpen(False)

            '' ''MyCLS.clsXMLOperations.WriteXMLFromSQL(Qstr, Path & "XML\XMLDataFromSQL.xml", MyCLS.clsXMLOperations.XMLFormat.ThroughSqlDataAdapter)
            '***END - CREATE XMLDATAFILE FROM SQL********************************************************************


            '***START - CREATE FINAL XML FROM XMLDATAFILE********************************************************************
            MyCLS.clsXMLOperations.TransformXML(Path & "XML\XMLDataFromSQL.xml", Path & "XML\FormatStyle.xsl", Path & "XML\ResultXML.xml")
            '***END - CREATE FINAL XML FROM XMLDATAFILE********************************************************************


            End
        Catch ex As Exception

        End Try
    End Sub

    Public Sub CreateTheNavigator()
        Dim xml As String = ""
        Dim xslt As String = ""
        'Initialize the xml and xslt variable so we can show how to read a string
        xml = ReadFileAsString("C:\Documents and Settings\DEV2\Desktop\DOI\XMLData2.xml")
        xslt = ReadFileAsString("C:\Documents and Settings\DEV2\Desktop\DOI\Book.xsl")
        'Load the String into a TextReader
        Dim tr As System.IO.TextReader = New System.IO.StringReader(xml)
        'Using the TextReader, load it into an XPathDocument
        Dim xp As System.Xml.XPath.XPathDocument = New System.Xml.XPath.XPathDocument(tr)
        'Now do the same with the Xslt document
        Dim xsltSR As System.IO.TextReader = New System.IO.StringReader(xslt)
        Dim xsltXR As System.Xml.XmlReader = New System.Xml.XmlTextReader(xsltSR)
        Dim trans As System.Xml.Xsl.XslTransform = New System.Xml.Xsl.XslTransform()
        'Load the XmlReader StyleSheet into the Transformation
        trans.Load(xsltXR)
        'Create the Stream to place the output.            
        Dim str As System.IO.Stream = New System.IO.MemoryStream()

        'Create the XPathNavigator
        Dim xpn As System.Xml.XPath.XPathNavigator = xp.CreateNavigator()
        'Transform the file.
        trans.Transform(xpn, Nothing, str)
        'You may return the information as seen in example 8.
    End Sub
    Private Function ReadFileAsString(ByVal path As String) As String
        Dim s As System.IO.Stream = New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        'Create a byte array the size of the file - 1. VB.NET has issues with the zero based stuff
        Dim by(Convert.ToInt32(s.Length - 1)) As Byte
        'read the file in as a byte array
        s.Read(by, 0, by.Length)
        'Close the Stream
        s.Close()
        'Use the GetString Method to return a String
        Return System.Text.Encoding.UTF8.GetString(by)
    End Function

    Public Shared Sub Transform(ByVal sXmlPath As String, ByVal sXslPath As String)
        Try

            'load the Xml doc 
            Dim myXPathDoc As New XPathDocument(sXmlPath)

            Dim myXslTrans As New XslTransform()

            'load the Xsl 
            myXslTrans.Load(sXslPath)

            'create the output stream 
            Dim myWriter As New XmlTextWriter("C:\Documents and Settings\DEV2\Desktop\DOI\Result.xml", System.Text.Encoding.UTF8)

            'do the actual transform of Xml 
            myXslTrans.Transform(myXPathDoc, Nothing, myWriter)

            myWriter.Close()
        Catch e As Exception
            Console.WriteLine("Exception: {0}", e.ToString())
        End Try
    End Sub

    Sub XML3()
        Dim textWriter As XmlTextWriter
        Try
            Dim FilePath1 As String = "C:\Documents and Settings\DEV2\Desktop\DOI\Book.xml"


            'Create a new file in C:\\ dir
            textWriter = New XmlTextWriter(FilePath1, Nothing)
            'After creating an instance, first thing you call us WriterStartDocument. When you're done writing, you call WriteEndDocument and TextWriter's Close method.  
            textWriter.WriteStartDocument()

            textWriter.WriteStartElement("Books")
            textWriter.WriteEndElement()

            'Write the ProcessingInstruction node
            Dim PI As String = "type='text/xsl' href='C:\Documents and Settings\DEV2\Desktop\DOI\Book.xsl'"
            textWriter.WriteProcessingInstruction("xml-stylesheet", PI)

            'Write the DocumentType node
            textWriter.WriteDocType("book", Nothing, Nothing, "<!ENTITY h'softcover'>")


            textWriter.WriteEndDocument()
            textWriter.Close()
        Catch ex As Exception

        Finally
            textWriter.Close()
        End Try
    End Sub

    Sub XML2()
        Try
            Dim FilePath2 As String = "C:\Documents and Settings\DEV2\Desktop\DOI\XMLData2.xml"
            MyCLS.clsXMLOperations.ReadXML(FilePath2)
        Catch ex As Exception

        End Try
    End Sub

    Sub XML1()
        Try
            MyCLS.strConnStringSQLCLIENT = "Initial Catalog=EBook_NEW;Data Source=127.0.0.1;UID=sa;PWD=sa123;"
            MyCLS.clsCOMMON.ConOpen(False)

            Dim FilePath1 As String = "C:\Documents and Settings\DEV2\Desktop\DOI\XMLData1.xml"
            Dim FilePath2 As String = "C:\Documents and Settings\DEV2\Desktop\DOI\XMLData2.xml"
            'Dim Query1 As String = "select * from Referrals FOR XML AUTO, XMLDATA"

            Dim Query2 As String '= "Select * From Referrals R " & _
            '"	Inner join Ref_Loc RL " & _
            '"		On RL.RefeNmbr=R.RefeNmbr "
            '**********************************************************************************
            'Query2 = "SELECT cast( " & _
            '        "		( " & _
            '        "			( " & _
            '        "				SELECT * " & _
            '        "				FROM Referrals R " & _
            '        "				WHERE R.RefeNmbr = 1 " & _
            '        "				FOR xml path, " & _
            '        "				root('Referrals') " & _
            '        "			)+ " & _
            '        "			isnull( " & _
            '        "					( " & _
            '        "						SELECT * " & _
            '        "						FROM Ref_Loc RL " & _
            '        "						WHERE RL.RefeNmbr = 1 " & _
            '        "						FOR xml path,root('Ref_Loc') " & _
            '        "					),'' " & _
            '        "				   ) " & _
            '        "		) AS xml) "
            '**********************************************************************************
            Query2 = "Select * From Book"
            '**********************************************************************************
            'Query2 = "SELECT " & _
            '        "	(  " & _
            '        "		SELECT ISBN From Book Where ISBN='9788184484243' " & _
            '        "		FOR	XML PATH('ISBN'), TYPE " & _
            '        "	), " & _
            '        "	( " & _
            '        "		SELECT Title From Book Where ISBN='9788184484243' " & _
            '        "		FOR XML PATH('TITLE'),	TYPE " & _
            '        "	), " & _
            '        "	( " & _
            '        "		SELECT " & _
            '        "			( " & _
            '        "				SELECT Author From Book Where ISBN='9788184484243' " & _
            '        "				FOR	XML PATH('Author'), TYPE " & _
            '        "			), " & _
            '        "			( " & _
            '        "				SELECT SubCategory From Book Where ISBN='9788184484243' " & _
            '        "				FOR XML PATH('SubCategory'), TYPE " & _
            '        "			) " & _
            '        "			FOR XML PATH('Author-SubCategory'), TYPE " & _
            '        "	) " & _
            '        "FOR XML PATH(''), ROOT('Books') "



            'MyCLS.clsXMLOperations.WriteXMLFromSQL(Query2, FilePath1, MyCLS.clsXMLOperations.XMLFormat.ThroughXmlReader)
            MyCLS.clsXMLOperations.WriteXMLFromSQL(Query2, FilePath2, MyCLS.clsXMLOperations.XMLFormat.ThroughSqlDataAdapter)

            MyCLS.clsCOMMON.ConClose()
        Catch ex As Exception

        End Try
    End Sub

    'Sub PDFOperations()
    '    Dim PDFFile As String = "D:\Narender\VBJavaScript.pdf"
    '    MyCLS.clsPDFOperations.OpenPDF()
    '    MsgBox(MyCLS.clsPDFOperations.PDFPageCount(PDFFile))
    '    MyCLS.clsPDFOperations.ClosePDF()
    '    MyCLS.clsPDFOperations.OpenPDF()
    '    MyCLS.clsPDFOperations.WritePDFinTXT(PDFFile, True)
    '    MyCLS.clsPDFOperations.ClosePDF()
    'End Sub

    'Sub PDF(ByVal PDFFilePath As String)
    '    'On Error Resume Next
    '    'Dim AcroApp As PdfLib.Pdf
    '    Dim a As AcroPDPage

    '    Dim AcroApp As CAcroApp
    '    Dim PDDoc As CAcroPDDoc
    '    Dim AVDoc As CAcroAVDoc

    '    Dim X As Double
    '    Dim Pg As Double
    '    Dim TempStr
    '    'acroapp.

    '    AcroApp = CreateObject("AcroExch.App", "")
    '    PDDoc = CreateObject("AcroExch.PDDoc", "")
    '    AVDoc = CreateObject("AcroExch.AVDoc", "")

    '    AcroApp.Hide()

    '    Dim StrTemp As String = PDFFilePath

    '    Dim bFileOpen = AVDoc.Open(StrTemp, "FILE NAME") 'Boolean

    '    If PDDoc.Open(StrTemp) Then
    '        Dim JSO = PDDoc.GetJSObject

    '        JSO.Console.Hide()
    '        JSO.Console.Clear()

    '        Dim W As System.IO.StreamWriter
    '        W = System.IO.File.CreateText(My.Application.Info.DirectoryPath() & "\PDFRead.txt")

    '        'MsgBox(AVDoc.FindText("XYY", 0, 0, 0)) 'TO SEARCH WITHIN PDF FILE

    '        For Pg = 1 To PDDoc.GetNumPages
    '            'JSO = PDDoc.GetJSObject
    '            'For X = 0 To JSO.GetPageNumWords 'Get the total number of words found
    '            'Label1.Text = Pg
    '            W.WriteLine("Page No : " & Pg)
    '            Try
    '                For X = 0 To 500 'Get the 10000 of words
    '                    'Try

    '                    'Label2.Text = X

    '                    TempStr = JSO.GetPageNthWord(Pg, X) '(page,word)
    '                    'Debug.Print(TempStr)    
    '                    If Len(TempStr) > 0 Then
    '                        W.Write(TempStr & " ")
    '                    End If
    '                    'Catch ex As Exception

    '                    System.Windows.Forms.Application.DoEvents()

    '                    'End Try
    '                Next X 'Next word
    '                W.WriteLine("")
    '                System.Windows.Forms.Application.DoEvents()
    '            Catch ex As Exception

    '            End Try
    '        Next
    '        W.Close()
    '        Process.Start(My.Application.Info.DirectoryPath() & "\PDFRead.txt")
    '    End If
    'End Sub

    'Sub CreateXMLFromSQL()
    'Try
    '    Dim sConnection As String = "Initial Catalog=ndhhs_rms;Data Source=127.0.0.1;UID=sa;PWD=sa123;"
    '    Dim mySqlConnection As SqlConnection = New SqlConnection(sConnection)
    '    Dim mySqlCommand As SqlCommand = New SqlCommand("select * from customer FOR XML AUTO, XMLDATA", mySqlConnection)
    '    mySqlCommand.CommandTimeout = 15
    '    '...
    '    mySqlConnection.Open()


    '    ' Now create the DataSet and fill it with xml data.
    '    Dim myDataSet1 As DataSet = New DataSet()
    '    myDataSet1.ReadXml(mySqlCommand.ExecuteXmlReader(), XmlReadMode.Fragment)

    '    ' Modify to match the other dataset
    '    myDataSet1.DataSetName = "NewDataSet"


    '    ' Get the same data through the provider.
    '    Dim mySqlDataAdapter As SqlDataAdapter = New SqlDataAdapter("select * from customer", sConnection)
    '    Dim myDataSet2 As DataSet = New DataSet()
    '    mySqlDataAdapter.Fill(myDataSet2)

    '    ' Write data to files: data1.xml and data2.xml.
    '    myDataSet1.WriteXml("c:\data1.xml")
    '    myDataSet2.WriteXml("c:\data2.xml")
    '    'Console.WriteLine("Data has been written to the output files: data1.xml and data2.xml")
    '    'Console.WriteLine()
    '    'Console.WriteLine("********************data1.xml********************")
    '    'Console.WriteLine(myDataSet1.GetXml())
    '    'Console.WriteLine()
    '    'Console.WriteLine("********************data2.xml********************")
    '    'Console.WriteLine(myDataSet2.GetXml())
    'Catch ex As Exception

    'End Try
    'End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        MyCLS.clsControls.prcFillDDLUsingDS(ComboBox1, "Book", "ISBN", "Distinct Title", "", "Title", MyCLS.clsControls.SortOrder.ASC, True)
    End Sub

    Private Sub cmdReadPDF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReadPDF.Click
        'Try
        Dim pPDF = "D:\Narender\Projects\ASP.NET\2005\EbookPdfDisplay\Books\817179808X\Chapter wise Pdf\Chapter-06_Postnatal and Neonatal History.pdf"

        'Dim Pr As New PdfReader(pPDF)            
        'Dim bText() As Byte
        'bText = Pr.GetPageContent(3)

        ''Dim Po As New PdfString(bText)
        ''bText = Po.GetOriginalBytes()

        'Dim p1 As New PdfEncodings
        'MsgBox(p1.ConvertToString(bText, "ISO-8859-1").Replace("128", "&euro;"))
        ''bText = Pr.LZWDecode(bText)
        ''bText = Pr.FlateDcode(bText)

        ''MsgBox()

        'End
        'MsgBox(Asc("."))   =   46
        'Dim p1 As PdfStream
        'Dim Pn As PdfName

        '' ''WORKING FINE
        '' ''Dim Pr As New PdfReader("D:\Narender\Projects\ASP.NET\2005\EbookPdfDisplay\Books\817179808X\Chapter wise Pdf\Chapter-01_History Taking.pdf")
        '' ''Dim bText() As Byte = Pr.GetPageContent(3)
        '' ''Dim sStr As String = ""
        '' ''For i As Long = 0 To bText.Length - 1
        '' ''    'If VarType(bText(i)) = vbString Then
        '' ''    If bText(i) = 46 Then
        '' ''        sStr += Chr(bText(i))
        '' ''    End If
        '' ''Next
        '' ''MsgBox(sStr)

        'Pn.fil()
        'p1.GetAsString("D:\Narender\Projects\ASP.NET\2005\EbookPdfDisplay\Books\817179808X\Chapter wise Pdf\Chapter-01_History Taking.pdf")
        'ListFieldNames()
        'Exit Sub


        '*
        'Dim Pr As New iTextSharp.text.pdf.
        'Dim Prs As pdf.PRStream
        'Dim x As pdf.RandomAccessFileOrArray

        'Prs.GetBytes()
        'Pr.GetStreamBytes(Prs)


        'Dim MyProcess As New Process
        'MyProcess.StartInfo.CreateNoWindow = False
        'MyProcess.StartInfo.Verb = "print"
        'MyProcess.StartInfo.FileName = pPDF
        'MyProcess.Start()
        'MyProcess.WaitForExit(10000)
        'MyProcess.CloseMainWindow()
        'MyProcess.Close()
        '**

        Dim sStr As String = ""
        Dim sStrPg As String
        MyCLS.clsPDFOperations.OpenPDF(True)
        'MyCLS.clsPDFOperations.ReadPDF(pPDF, ".", 50)
        'MsgBox(Convert.ToByte(Convert.ToString(".")))

        'sStr = MyCLS.clsPDFOperations.FindLocationInPDF(pPDF, 49)
        'MsgBox(sStr.Split(",")(0) & vbCrLf & sStr.Split(",")(1))
        'MsgBox(MyCLS.clsPDFOperations.FindWordsJSO(pPDF, sStr.Split(",")(0), sStr.Split(",")(1)).ToString())


        'MsgBox(Convert.ToByte(Convert.ToString(49)))
        'sStr = MyCLS.clsPDFOperations.ReadPDF(pPDF, Convert.ToByte(Convert.ToString(49)), 50, 0)
        'sStr = MyCLS.clsPDFOperations.ReadALLPDF(pPDF, 0)
        sStrPg = MyCLS.clsPDFOperations.ReadALLPDFPg(pPDF, 1)
        ''Debug.Print(sStr)
        'For i As Int16 = 0 To sStrPg.Length - 1
        '    Debug.Print(sStrPg(i))
        'Next
        'MyCLS.clsPDFOperations.ClosePDF()
        '***WRITE ALL DATA***
        MyCLS.clsFileHandling.OpenFile("c:\_ALL.txt")
        MyCLS.clsFileHandling.WriteFile(sStrPg)
        MyCLS.clsFileHandling.CloseFile(True)
        '***WRITE ALL DATA***

        '***WRITE EXTRACTED DATA***
        MyCLS.clsFileHandling.OpenFile("c:\_EXTRACTED.txt")
        MsgBox(MyCLS.clsPDFOperations.PDFWordsCount(pPDF, 3))
        MyCLS.clsPDFOperations.ExtractText(sStrPg)
        MyCLS.clsFileHandling.CloseFile(True)
        '***WRITE EXTRACTED DATA***

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub

    Private Sub cmdPDFExtract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPDFExtract.Click        
        'Dim pPDF = "D:\Narender\Projects\ASP.NET\2005\EbookPdfDisplay\Books\817179808X\Chapter wise Pdf\Chapter-05_Natal History.pdf"
        Dim pPDF = "C:\_9788180610363_SECURED\TestAbstract1\9788180613240\Combined Books\CB.pdf"
        '***WORKING FINE****
        '' ''Dim pg As Int16 = 0
        '' ''Dim sStrPg As String
        '' ''Dim sStrPgEX As String
        ' '' ''Try
        ' '' ''***Find Page of More Thank 100 Words***
        '' ''While (1)
        '' ''    If MyCLS.clsPDFOperations.PDFWordsCount(pPDF, pg) > 100 Then
        '' ''        Exit While
        '' ''    Else
        '' ''        pg += 1
        '' ''    End If
        '' ''End While

        ' '' ''***Extract All From the Page***
        '' ''sStrPg = MyCLS.clsPDFOperations.ReadALLPDFPg(pPDF, pg + 1)
        ' '' ''***WRITE ALL DATA***
        '' ''MyCLS.clsFileHandling.OpenFile("c:\_ALL.txt")
        '' ''MyCLS.clsFileHandling.WriteFile(sStrPg)
        '' ''MyCLS.clsFileHandling.CloseFile(True)
        ' '' ''***WRITE ALL DATA***

        ' '' ''***Extract Text From All Page Contents***
        '' ''MyCLS.clsFileHandling.OpenFile("c:\_EXTRACTED.txt")
        '' ''sStrPgEX = MyCLS.clsPDFOperations.ExtractText(sStrPg)
        '' ''MyCLS.clsFileHandling.CloseFile(True)

        ' '' ''***Find The Paragraph & Extract It***
        '' ''MyCLS.clsFileHandling.OpenFile("c:\_Para_EX.txt")
        '' ''MsgBox(Mid(sStrPgEX, 1, InStr(250, sStrPgEX, ".", CompareMethod.Binary)))
        '' ''MyCLS.clsFileHandling.WriteFile(Mid(sStrPgEX, 1, InStr(250, sStrPgEX, ".", CompareMethod.Binary)))
        '' ''MyCLS.clsFileHandling.CloseFile(True)

        ' '' ''MyCLS.clsFileHandling.OpenFile("c:\_ASCII_EXTRACTED.txt")
        ' '' ''For i As Int16 = 0 To sStrPgEX.Length - 1
        ' '' ''    MyCLS.clsFileHandling.WriteFile(Asc(sStrPgEX(i)))
        ' '' ''Next
        ' '' ''MyCLS.clsFileHandling.CloseFile(True)
        '***WORKING FINE****

        '******USING FUNCTION ABOVE CODE IN THIS PROCEDURE
        MyCLS.clsFileHandling.OpenFile("c:\_Para_EX.txt")
        MyCLS.clsFileHandling.WriteFile(ExtractPARA(pPDF))
        MyCLS.clsFileHandling.CloseFile(True)

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub

    Function ExtractPARA(ByVal pPDF As String) As String
        Dim pg As Int16 = 0
        Dim sStrPg As String
        Dim sStrPgEX As String
        'Try
        '***Find Page of More Thank 100 Words***
        ''While (1)
        ''    If MyCLS.clsPDFOperations.PDFWordsCount(pPDF, pg) > 100 Then
        ''        Exit While
        ''    Else
        ''        pg += 1
        ''    End If
        ''End While

        '***Extract All From the Page***
        ''''''''''''''''sStrPg = MyCLS.clsPDFOperations.ReadALLPDFPg(pPDF, pg + 1)
        ' '' ''***WRITE ALL DATA***
        '' ''MyCLS.clsFileHandling.OpenFile("c:\_ALL.txt")
        '' ''MyCLS.clsFileHandling.WriteFile(sStrPg)
        '' ''MyCLS.clsFileHandling.CloseFile(True)
        ' '' ''***WRITE ALL DATA***

        '***Extract Text From All Page Contents***
        '' ''MyCLS.clsFileHandling.OpenFile("c:\_EXTRACTED.txt")
        ''''''''''''''''sStrPgEX = MyCLS.clsPDFOperations.ExtractText(sStrPg)
        '' ''MyCLS.clsFileHandling.WriteFile(sStrPgEX)
        '' ''MyCLS.clsFileHandling.CloseFile(True)

        '***Find The Paragraph & Extract It***
        'MyCLS.clsFileHandling.OpenFile("c:\_Para_EX.txt")
        'MsgBox(Mid(sStrPgEX, 1, InStr(250, sStrPgEX, ".", CompareMethod.Binary)))
        'MyCLS.clsFileHandling.WriteFile(Mid(sStrPgEX, 1, InStr(250, sStrPgEX, ".", CompareMethod.Binary)))
        'MyCLS.clsFileHandling.CloseFile(True)

        ''''''''''''''''Return Mid(sStrPgEX, 1, InStr(250, sStrPgEX, ".", CompareMethod.Binary))


        'sStrPgEX = MyCLS.clsPDFOperations.ParsePdfText(pPDF, pg + 1) '.Replace("?", "")
        sStrPgEX = MyCLS.clsPDFOperations.ParsePdfText(pPDF, 13)

        Return Mid(sStrPgEX, 1, sStrPgEX.LastIndexOf(".") + 1)


        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Function













    'Private Const oldchar = 15
    '  Private Sub  ProcessOutput(FILE* file, char* output, size_t len)
    '    'Are we currently inside a text object?
    '    Dim intextobject As Boolean = False
    '    'Is the next character literal 
    '    '(e.g. \\ to get a \ character or \( to get ( ):
    '    Dim nextliteral As Boolean = False

    '    '() Bracket nesting level. Text appears inside ()
    '    Dim rbdepth As Int16 = 0

    '    'Keep previous chars to extract numbers etc.:
    '    Dim oc(oldchar) As Char
    '    Dim j As Int16 = 0
    '    For j = 0 To j < oldchar
    '        oc(j) = " "
    '        j += 1
    '    Next


    'for (size_t i=0; i<len; i++)
    '{
    '  char c = output[i];
    '        If (intextobject) Then
    '  {
    '    if (rbdepth==0 && seen2("TD", oc))
    '    {
    '                'Positioning.
    '                'See if a new line has to start or just a tab:
    '      float num = ExtractNumber(oc,oldchar-5);
    '                If (num > 1.0) Then
    '      {
    '        fputc(0x0d, file);
    '        fputc(0x0a, file);
    '      }
    '                    If (num < 1.0) Then
    '      {
    '        fputc('\t', file);
    '      }
    '    }
    '    if (rbdepth==0 && seen2("ET", oc))
    '    {
    '                            'End of a text object, also go to a new line.
    '      intextobject = false;
    '      fputc(0x0d, file);
    '      fputc(0x0a, file);
    '    }
    '    else if (c=='(' && rbdepth==0 && !nextliteral) 
    '    {
    '                            'Start outputting text!
    '      rbdepth=1;
    '                            'See if a space or tab (>1000) is called for by looking
    '                            'at the number in front of (
    '      int num = ExtractNumber(oc,oldchar-1);
    '                            If (num > 0) Then
    '      {
    '                                If (num > 1000.0) Then
    '        {
    '          fputc('\t', file);
    '        }
    '                                ElseIf (num > 100.0) Then
    '        {
    '          fputc(' ', file);
    '        }
    '      }
    '    }
    '    else if (c==')' && rbdepth==1 && !nextliteral) 
    '    {
    '                                    'Stop outputting text
    '      rbdepth=0;
    '    }
    '    else if (rbdepth==1) 
    '    {
    '                                    'Just a normal text character:
    '      if (c=='\\' && !nextliteral)
    '      {
    '                                        'Only print out next character 
    '                                        'no matter what. Do not interpret.
    '        nextliteral = true;
    '      }
    '                                    Else
    '      {
    '        nextliteral = false;
    '        if ( ((c>=' ') && (c<='~')) || ((c>=128) && (c<255)) )
    '        {
    '          fputc(c, file);
    '        }
    '      }
    '    }
    '  }

    '                                            'Store the recent characters for 
    '                                            'when we have to go back for a number:
    '  for (j=0; j<oldchar-1; j++) oc[j]=oc[j+1];
    '    oc[oldchar-1]=c;
    '                                                If (!intextobject) Then
    '  {
    '                                                    If (seen2("BT", oc)) Then
    '    {
    '                                                        'Start of a text object:
    '      intextobject = true;
    '    }
    '  }
    '}
    'End Sub

    ''Private Sub cmdCombination_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCombination.Click
    ''    Dim strNationality() As String = {"N", "B", "S", "D", "G"}
    ''    Dim strHouse() As String = {"G", "W", "R", "Y", "B"}
    ''    Dim strDrink() As String = {"T", "C", "W", "M", "B"}
    ''    Dim strSmoke() As String = {"P", "P", "B", "D", "B"}
    ''    Dim strPet() As String = {"D", "C", "B", "H", "F"}
    ''    Dim strPass As String
    ''    Dim i As Int16 = 0

    ''    MyCLS.clsFileHandling.OpenFile("C:\_Passes.txt")
    ''    MyCLS.clsXLSOperations.OpenXLSObject(True)
    ''    For n As Int16 = 0 To 4
    ''        For h As Int16 = 0 To 4
    ''            For d As Int16 = 0 To 4
    ''                For s As Int16 = 0 To 4
    ''                    For p As Int16 = 0 To 4
    ''                        If strPet(p) = "F" Then
    ''                            strPass = strNationality(n) & strHouse(h) & strDrink(d) & strSmoke(s) & strPet(p)
    ''                            strPass = strPass.ToLower()
    ''                            MyCLS.clsFileHandling.WriteFile(strPass)
    ''                            Label1.Text = i
    ''                            i += 1
    ''                            MyCLS.clsXLSOperations.OpenXLSFile("C:\Documents and Settings\DEV02\Desktop\EINSTEIN_CRACKED_(2).xlsx", strPass, True)
    ''                            System.Windows.Forms.Application.DoEvents()
    ''                        End If
    ''                    Next
    ''                Next
    ''            Next
    ''        Next
    ''    Next
    ''    MyCLS.clsXLSOperations.CloseXLSObject()
    ''    MyCLS.clsFileHandling.CloseFile(True)
    ''End Sub


    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        PictureBox1.ImageLocation = "C:\_FP\PBLIndex.jpg"
    End Sub

    Private Sub cmdInsertImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsertImage.Click
        Try
            Dim objLIBTable_2 As New LIBTable_2
            Dim objDALTable_2 As New DALTable_2
            Dim tp As New MyCLS.TransportationPacket

            objLIBTable_2.c = "c"
            objLIBTable_2.d = "d"
            objLIBTable_2.ImgVarBinary = MyCLS.clsImaging.PictureBoxToByteArray(PictureBox1)
            tp.MessagePacket = objLIBTable_2
            tp = objDALTable_2.InsertTable_2(tp)

            If tp.MessageId = 1 Then
                Dim strOutParamValues As String() = tp.MessageResultset
                MsgBox(strOutParamValues(0))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub cmdGetImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetImage.Click
        Try
            Dim objLIBTable_2Listing As New LIBTable_2Listing
            Dim objDALTable_2 As New DALTable_2
            Dim tp As New MyCLS.TransportationPacket
            tp.MessagePacket = "c"    'ID to be Passed

            '    objLIBTable_2Listing(0).c = ""
            '    objLIBTable_2Listing(1).d = ""
            '    objLIBTable_2Listing(2).ImgVarBinary = ""

            tp = objDALTable_2.GetTable_2Details(tp)
            If tp.MessageId = 1 Then
                objLIBTable_2Listing = tp.MessageResultset
                MyCLS.clsImaging.ByteArray2Image(PictureBox1.Image, objLIBTable_2Listing(0).ImgVarBinary)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



    '' ''***WORKING FINE***
    '' ''Private Function PictureBoxToByteArray(ByVal oPictureBox As PictureBox) As Byte()
    '' ''    Dim oStream As New MemoryStream
    '' ''    Dim bmp As New Bitmap(oPictureBox.Image)
    '' ''    Try
    '' ''        bmp.Save(oStream, Imaging.ImageFormat.Bmp)
    '' ''        PictureBoxToByteArray = oStream.ToArray
    '' ''        bmp.Dispose()
    '' ''        oStream.Close()
    '' ''    Catch ex As Exception
    '' ''        '//--Catch Msg 
    '' ''    End Try
    '' ''End Function
    '' ''Public Sub ByteArray2Image(ByRef NewImage As System.Drawing.Image, ByVal ByteArr() As Byte)
    '' ''    Dim ImageStream As MemoryStream
    '' ''    Try
    '' ''        If ByteArr.GetUpperBound(0) > 0 Then
    '' ''            ImageStream = New MemoryStream(ByteArr)
    '' ''            NewImage = System.Drawing.Image.FromStream(ImageStream)
    '' ''        Else
    '' ''            NewImage = Nothing
    '' ''        End If
    '' ''    Catch ex As Exception
    '' ''        NewImage = Nothing
    '' ''    End Try
    '' ''End Sub

    Private Sub CmdPDFRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdPDFRead.Click
        ' ''MyCLS.clsFileHandling.OpenFile("C:\_Token.txt")
        ' ''PdfManipulation.ParsePdfText("D:\Narender\Projects\ASP.NET\2005\EbookPdfDisplay\Books\817179808X\Chapter wise Pdf\Chapter-05_Natal History.pdf", 0, 0)
        ' ''MyCLS.clsFileHandling.CloseFile(True)
        '
        PdfManipulation.ExtractPdfPageOpen("C:\_9788180610363\Chapter wise Pdf\Chapter-01_Pain Bane or Beneficial.pdf", 0, 0, "C:\qqq.pdf")
    End Sub

    Private Sub cmdInsertAllTypes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsertAllTypes.Click
        Try
            Dim objLIBAllDBTypesSQLTable As New LIBAllDBTypesSQLTable
            Dim objDALAllDBTypesSQLTable As New DALAllDBTypesSQLTable
            Dim tp As New MyCLS.TransportationPacket

            Dim objLIBAllDBTypesSQLTable_toPass As New LIBAllDBTypesSQLTable

            Dim intL As Integer = Convert.ToInt32(100 * 100)
            Dim pBuff As Byte() = New Byte(intL) {}
            pBuff(0) = "11"

            objLIBAllDBTypesSQLTable.ID = -1
            objLIBAllDBTypesSQLTable.A = 65000
            objLIBAllDBTypesSQLTable.B = pBuff
            'objLIBAllDBTypesSQLTable.B = MyCLS.clsImaging.PictureBoxToByteArray(New PictureBox) 'pBuff
            objLIBAllDBTypesSQLTable.C = True
            objLIBAllDBTypesSQLTable.D = "Char"
            objLIBAllDBTypesSQLTable.E = Date.Now()
            objLIBAllDBTypesSQLTable.F = 14.0
            objLIBAllDBTypesSQLTable.G = 14.253
            objLIBAllDBTypesSQLTable.H = 9.5251
            objLIBAllDBTypesSQLTable.I = MyCLS.clsImaging.PictureBoxToByteArray(New PictureBox)
            objLIBAllDBTypesSQLTable.J = 25
            objLIBAllDBTypesSQLTable.K = 125466008.25
            objLIBAllDBTypesSQLTable.L = "NChar"
            objLIBAllDBTypesSQLTable.M = "NText"
            objLIBAllDBTypesSQLTable.N = "26.25"
            objLIBAllDBTypesSQLTable.O = 12643545334.235
            objLIBAllDBTypesSQLTable.P = "NVarChar 50"
            objLIBAllDBTypesSQLTable.Q = "NVarChar MAX"
            objLIBAllDBTypesSQLTable.R = 10.36
            objLIBAllDBTypesSQLTable.S = Date.Now().ToShortDateString
            objLIBAllDBTypesSQLTable.T = 23768
            objLIBAllDBTypesSQLTable.U = 50.25
            objLIBAllDBTypesSQLTable.V = "Object"
            objLIBAllDBTypesSQLTable.W = "Text"
            objLIBAllDBTypesSQLTable.X = pBuff
            objLIBAllDBTypesSQLTable.Y = 2
            objLIBAllDBTypesSQLTable.Z = Guid.NewGuid
            objLIBAllDBTypesSQLTable.A1 = MyCLS.clsImaging.PictureBoxToByteArray(New PictureBox)
            objLIBAllDBTypesSQLTable.B1 = MyCLS.clsImaging.PictureBoxToByteArray(New PictureBox)
            objLIBAllDBTypesSQLTable.C1 = "varchar"
            objLIBAllDBTypesSQLTable.D1 = "varchar max"
            objLIBAllDBTypesSQLTable.E1 = "xml"

            tp.MessagePacket = objLIBAllDBTypesSQLTable
            tp = objDALAllDBTypesSQLTable.InsertAllDBTypesSQLTable(tp)

            If tp.MessageId = 1 Then
                Dim strOutParamValues As String() = tp.MessageResultset
                MsgBox(strOutParamValues(0))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdInsertAllTypesNEW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsertAllTypesNEW.Click
        'Try
        '    Dim objLIBAllDBTypesSQLTable_NEW As New LIBAllDBTypesSQLTable_NEW
        '    Dim objDALAllDBTypesSQLTable_NEW As New DALAllDBTypesSQLTable_NEW
        '    Dim tp As New MyCLS.TransportationPacket

        '    Dim intL As Integer = Convert.ToInt32(100 * 100)
        '    Dim pBuff As Byte() = New Byte(intL) {}
        '    pBuff(0) = "11"

        '    objLIBAllDBTypesSQLTable_NEW.ID = -1
        '    objLIBAllDBTypesSQLTable_NEW.bigint_A = 65000
        '    objLIBAllDBTypesSQLTable_NEW.binary_B = pBuff
        '    objLIBAllDBTypesSQLTable_NEW.bit_C = True
        '    objLIBAllDBTypesSQLTable_NEW.char_D = "Char"
        '    objLIBAllDBTypesSQLTable_NEW.datetime_E = Date.Now()
        '    objLIBAllDBTypesSQLTable_NEW.decimal_F = 14
        '    objLIBAllDBTypesSQLTable_NEW.decimal_G = 14.253
        '    objLIBAllDBTypesSQLTable_NEW.float_H = 9.5251
        '    objLIBAllDBTypesSQLTable_NEW.image_I = MyCLS.clsImaging.PictureBoxToByteArray(New PictureBox)
        '    objLIBAllDBTypesSQLTable_NEW.int_J = 25
        '    objLIBAllDBTypesSQLTable_NEW.money_K = 125466008.3
        '    objLIBAllDBTypesSQLTable_NEW.nchar_L = "NChar"
        '    objLIBAllDBTypesSQLTable_NEW.ntext_M = "NText"
        '    objLIBAllDBTypesSQLTable_NEW.numeric_N = 26.25
        '    objLIBAllDBTypesSQLTable_NEW.numeric_O = 12643545334
        '    objLIBAllDBTypesSQLTable_NEW.nvarchar_P = "NVarChar 50"
        '    objLIBAllDBTypesSQLTable_NEW.nvarchar_Max_Q = "NVarChar MAX"
        '    objLIBAllDBTypesSQLTable_NEW.real_R = 10.36
        '    objLIBAllDBTypesSQLTable_NEW.smalldatetime_S = Date.Now().ToShortDateString
        '    objLIBAllDBTypesSQLTable_NEW.smallint_T = 23768
        '    objLIBAllDBTypesSQLTable_NEW.smallmoney_U = 50.25
        '    objLIBAllDBTypesSQLTable_NEW.sql_variant_V = "Object"
        '    objLIBAllDBTypesSQLTable_NEW.text_W = "Text"
        '    objLIBAllDBTypesSQLTable_NEW.timestamp_X = pBuff
        '    objLIBAllDBTypesSQLTable_NEW.tinyint_Y = 2
        '    objLIBAllDBTypesSQLTable_NEW.uniqueidentifier_Z = Guid.NewGuid
        '    objLIBAllDBTypesSQLTable_NEW.varbinary_A1 = MyCLS.clsImaging.PictureBoxToByteArray(New PictureBox)
        '    objLIBAllDBTypesSQLTable_NEW.varbinary_Max_B1 = MyCLS.clsImaging.PictureBoxToByteArray(New PictureBox)
        '    objLIBAllDBTypesSQLTable_NEW.varchar_C1 = "varchar"
        '    objLIBAllDBTypesSQLTable_NEW.varchar_Max_D1 = "varchar max"
        '    objLIBAllDBTypesSQLTable_NEW.xml_E1 = "xml"

        '    tp.MessagePacket = objLIBAllDBTypesSQLTable_NEW
        '    tp = objDALAllDBTypesSQLTable_NEW.InsertAllDBTypesSQLTable_NEW(tp)

        '    If tp.MessageId > -1 Then
        '        Dim strOutParamValues As String() = tp.MessageResultset
        '        MsgBox(strOutParamValues(0))
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
        '*************************************************************
        '*************************************************************
        '*************************************************************
        '*************************************************************
        Try
            Dim objLIBTBL1 As New LIBTBL1
            Dim objDALTBL1 As New DALTBL1
            Dim tp As New MyCLS.TransportationPacket

            objLIBTBL1.ID = -1
            objLIBTBL1.Img = MyCLS.clsImaging.PictureBoxToByteArray(New PictureBox)
            tp.MessagePacket = objLIBTBL1
            tp = objDALTBL1.InsertTBL1(tp)

            If tp.MessageId > -1 Then
                Dim strOutParamValues As String() = tp.MessageResultset
                MsgBox(strOutParamValues(0))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdFill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFill.Click
        Try
            MyCLS.clsControls.prcFillListChecked(ChkLst1, "Select ISBN From Book", "Book", "ISBN", "String")

            MyCLS.clsControls.prcFillGridWin(DGVChk1, "Select ISBN,Title From Book", "Book", True, "SelectToPrint", "Select To Print")
            MyCLS.clsControls.prcFillGridWin(DGV1, "Select ISBN,Title From Book", "Book", False, "SelectToPrint", "Select To Print")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdTestChecked_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestChecked.Click
        Try
            MsgBox(MyCLS.clsControls.fnListIsChecked(ChkLst1))
            MsgBox(MyCLS.clsControls.fnGridIsChecked(DGVChk1, True))
            MsgBox(MyCLS.clsControls.fnGridIsChecked(DGV1, False))

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub chkSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectAll.CheckedChanged
        If chkSelectAll.Checked Then
            MyCLS.clsControls.prcListCheckAll(ChkLst1)
            MyCLS.clsControls.prcGridCheckAll(DGVChk1, True)
            MyCLS.clsControls.prcGridCheckAll(DGV1, False)
        Else
            MyCLS.clsControls.prcListUnCheckAll(ChkLst1)
            MyCLS.clsControls.prcGridUnCheckAll(DGVChk1, True)
            MyCLS.clsControls.prcGridUnCheckAll(DGV1, False)
        End If
    End Sub

    Private Sub cmdCheckSelected_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCheckSelected.Click
        Try
            Dim s As String() = {"a", "81-8061-036-5"}
            MyCLS.clsControls.prcListCheckSelected(ChkLst1, s)
            MyCLS.clsControls.prcGridCheckSelected(DGVChk1, True, s, 1)
            MyCLS.clsControls.prcGridCheckSelected(DGV1, False, s, 0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdError_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdError.Click
        Try
            Dim i As Int16 = 10
            i = i / 0
        Catch ex As Exception
            Dim e1 As MyCLS.clsHandleException
        End Try
    End Sub

    Private Sub cmdToInt16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToInt16.Click
        'MsgBox(Convert.ToInt16(TextBox1.Text))        
        Dim pData As String = "00A4040007D4100000030001"
        Dim iDataLen As Integer = pData.Trim().Length / 2
        Dim byData(iDataLen) As Byte

        For i As Integer = 0 To iDataLen - 1
            byData(i) = Convert.ToByte(Convert.ToInt16(pData.Substring(i * 2, 2), 16))
        Next

        'Dim strHex2StrData As String = "10000048440110100000010000170000"
        'iDataLen = strHex2StrData.Length
        'Dim bsData() As Byte = {"1", "0", "000048440110100000010000170000"}
        'Dim strHex2StrData1 As String = ""
        'pData = ""

        'For i As Integer = 0 To iDataLen - 1
        '    bsData(i) = Convert.ToByte(strHex2StrData(i))
        'Next

        'For i As Integer = 0 To iDataLen - 1
        '    strHex2StrData1 += String.Format("{0:x2}", Convert.ToInt16(bsData(i)))
        '    strResult = strResult & Right("00" & Hex(pData(i)), 2)
        'Next

        ''MsgBox(Val(TextBox2.Text) ^ Val(TextBox1.Text))
        'MsgBox(System.Convert.ToInt32("00490C01", 16))

        'Command(&H21, "01")        
    End Sub

    Private Sub Command(ByVal pCmd As Byte, ByVal pData As String)        
        Dim iDataLen As Int16 = pData.Trim().Length / 2
        Dim byData(iDataLen) As Byte

        For i As Int16 = 0 To iDataLen - 1
            byData(i) = Convert.ToByte(Convert.ToInt16(pData.Substring(i * 2, 2), 16))
        Next
        MsgBox("cmd : " & pCmd & vbCrLf & "Data : " & Convert.ToBase64String(byData) & vbCrLf & "DataLen : " & iDataLen)
    End Sub

#Region "Insert Multiple Rows Single Column"
    Private Sub cmdInsertMultipleRows1Col_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsertMultipleRows1Col.Click
        Dim itemList As Int16() = {1, 2, 3, 4}
        Dim MyXMLItemList As New MyCLS.XMLItemList
        For i As Integer = 0 To itemList.Count - 1
            MyXMLItemList.AddItem(itemList(i))
        Next

        Dim Packet As New MyCLS.TransportationPacket
        'Dim objLIBInsertManyRows As New LIBInsertManyRows
        'Dim objDALInsertManyRows As New DALInsertManyRows

        'objLIBInsertManyRows.InsertedIDString = MyXMLItemList.ToString
        'Packet.MessagePacket = objLIBInsertManyRows
        'objDALInsertManyRows.InsertInsertManyRows(Packet)
    End Sub

    'Public Class XMLItemList
    '    Private sb As System.Text.StringBuilder

    '    Public Sub New()
    '        sb = New System.Text.StringBuilder
    '        sb.Append("<items>" & vbCrLf)
    '    End Sub

    '    Public Sub AddItem(ByVal Item As String)
    '        sb.AppendFormat("<item id={0}{1}{2}></item>{3}", Chr(34), Item, Chr(34), vbCrLf)
    '    End Sub

    '    Public Overrides Function ToString() As String
    '        sb.Append("</items>" & vbCrLf)
    '        Return sb.ToString
    '    End Function
    'End Class
#End Region




#Region "Insert Multiple Rows Many Columns"
    Private Sub cmdInsertMultipleRows2Col_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsertMultipleRows2Col.Click
        Dim itemList As String() = {"a", "b", "c", "d"}
        Dim MyXMLItemListManyCols As New MyCLS.XMLItemListManyCols
        For i As Integer = 0 To itemList.Count - 1
            MyXMLItemListManyCols.RowBegin("Location")
            MyXMLItemListManyCols.AddItem("Col" & i, itemList(i))
            MyXMLItemListManyCols.RowEnd("Location")
        Next

        MsgBox(MyXMLItemListManyCols.ToString)
        'Dim Packet As New MyCLS.TransportationPacket
        'Dim objLIBInsertManyRows As New LIBInsertManyRows
        'Dim objDALInsertManyRows As New DALInsertManyRows

        'objLIBInsertManyRows.InsertedIDString = MyXMLItemListManyCols.ToString
        'Packet.MessagePacket = objLIBInsertManyRows
        'objDALInsertManyRows.InsertInsertManyRows(Packet)
    End Sub



    'Public Class XMLItemListManyCols
    '    Private sb As System.Text.StringBuilder

    '    Public Sub New()
    '        sb = New System.Text.StringBuilder
    '        sb.Append("<rows>" & vbCrLf)
    '    End Sub

    '    Public Sub RowBegin(ByVal RowName As String)
    '        '<cars><car><Name>BMW</Name><Color>Red</Color></car><car><Name>Audi</Name><Color>Green</Color></car></cars>
    '        sb.AppendFormat("<{0}>", RowName, vbCrLf)
    '    End Sub

    '    Public Sub AddItem(ByVal ColName As String, ByVal Item As String)
    '        '<cars><car><Name>BMW</Name><Color>Red</Color></car><car><Name>Audi</Name><Color>Green</Color></car></cars>
    '        sb.AppendFormat("<{0}>{1}</{0}>", ColName, Item, vbCrLf)
    '    End Sub

    '    Public Sub RowEnd(ByVal RowName As String)
    '        '<cars><car><Name>BMW</Name><Color>Red</Color></car><car><Name>Audi</Name><Color>Green</Color></car></cars>
    '        sb.AppendFormat("</{0}>", RowName, vbCrLf)
    '    End Sub

    '    Public Overrides Function ToString() As String
    '        sb.Append("</rows>" & vbCrLf)
    '        Return sb.ToString
    '    End Function
    'End Class
#End Region

    Private Sub cmdLamda_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLamda.Click
        Dim numbers() As Integer = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}
        Dim lastIndex = Function(intArray() As Integer) intArray.Length - 1
        For i = 0 To lastIndex(numbers)
            numbers(i) += 1
        Next



        Dim notNothing =Function(num? As Integer) num IsNot Nothing
        Dim arg As Integer = 14
        Console.WriteLine("Does the argument have an assigned value?")
        MsgBox(notNothing(arg))
        MsgBox(notNothing(Nothing))


    End Sub

    Private Sub cmdAttachDBFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAttachDBFile.Click
        Dim objCon As New SqlClient.SqlConnection

        objCon.ConnectionString = "Server=.\SQLExpress;AttachDbFilename=D:\Narender\All_DataBases\TMPDATABASE.mdf;Database=TMPDATABASE; Trusted_Connection=Yes;"
        objCon.Open()
        MsgBox("Opened")
        objCon.Close()
    End Sub
End Class