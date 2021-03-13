Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.Text


Public Class PdfManipulation
    Public Shared Function ParsePdfText(ByVal sourcePDF As String, _
                                 Optional ByVal fromPageNum As Integer = 0, _
                                 Optional ByVal toPageNum As Integer = 0, Optional ByVal keywordText As String = "") As String

        Dim sb As New System.Text.StringBuilder()
        Try
            Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePDF)
            Dim pageBytes() As Byte = Nothing
            Dim token As iTextSharp.text.pdf.PRTokeniser = Nothing
            Dim tknType As Integer = -1
            Dim tknValue As String = String.Empty
            ''Dim Keyword As String = ClassPhraseSearch.GetSplitSearchPhrase(keywordText).ToLower
            Dim objRegx As Regex = Nothing
            'objRegx = New Regex("@Payments")
            objRegx = New Regex("@")

            If fromPageNum = 0 Then
                fromPageNum = 1
            End If
            If toPageNum = 0 Then
                toPageNum = reader.NumberOfPages
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
                            'I need to add these additional tests to properly add whitespace to the output string
                            MyCLS.clsFileHandling.WriteFile(tknType & vbTab & ": " & tknValue)
                        ElseIf tknType = 1 AndAlso tknValue = "-600" Then
                            strInBldr.Append(" ")
                        ElseIf tknType = 10 AndAlso tknValue = "TJ" Then
                            strInBldr.Append(" ")
                        End If                        
                    End While

                    ''If strInBldr.ToString().ToLower().Contains(Keyword.Trim(" ").ToLower()) And Keyword.ToLower().Length > 2 Then
                    pages.Append(i.ToString("#000") + ",")
                    ''End If
                End If
            Next i
            sb.Append(reader.NumberOfPages.ToString() + "," + pages.ToString())
        Catch ex As Exception
            Return String.Empty
        End Try

        Return sb.ToString()
    End Function

    ''Public Shared Function ParsePdfPhraseSearch(ByVal sourcePDF As String, _
    ''                           Optional ByVal fromPageNum As Integer = 0, _
    ''                           Optional ByVal toPageNum As Integer = 0, Optional ByVal keyword As String = "") As String

    ''    Dim sb As New System.Text.StringBuilder()
    ''    Try
    ''        Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePDF)
    ''        Dim pageBytes() As Byte = Nothing
    ''        Dim token As iTextSharp.text.pdf.PRTokeniser = Nothing
    ''        Dim tknType As Integer = -1
    ''        Dim tknValue As String = String.Empty

    ''        If fromPageNum = 0 Then
    ''            fromPageNum = 1
    ''        End If
    ''        If toPageNum = 0 Then
    ''            toPageNum = reader.NumberOfPages
    ''        End If

    ''        If fromPageNum > toPageNum Then
    ''            Throw New ApplicationException("Parameter error: The value of fromPageNum can " & _
    ''                                       "not be larger than the value of toPageNum")
    ''        End If
    ''        Dim strSerachPhrase As String = ClassPhraseSearch.GetSplitSearchPhrase(keyword.TrimEnd.TrimStart).ToLower()
    ''        Dim intPagr As Integer = 0
    ''        Dim pages As New StringBuilder()
    ''        Dim objPdfParser As New PDFParser()
    ''        Dim strInBldr1 As New StringBuilder()
    ''        'For i As Integer = fromPageNum To toPageNum Step 1
    ''        '    pageBytes = reader.GetPageContent(i)

    ''        '    If Not IsNothing(pageBytes) Then
    ''        '        token = New iTextSharp.text.pdf.PRTokeniser(pageBytes)

    ''        '        strInBldr1.Append(objPdfParser.ExtractTextFromPDFBytes(reader.GetPageContent(i)) + " ").Replace(ControlChars.Lf, " ")
    ''        '    End If
    ''        'Next

    ''        For i As Integer = fromPageNum To toPageNum Step 1
    ''            pageBytes = reader.GetPageContent(i)

    ''            If Not IsNothing(pageBytes) Then
    ''                token = New iTextSharp.text.pdf.PRTokeniser(pageBytes)
    ''                Dim strInBldr As New StringBuilder()
    ''                strInBldr.Append(ReplaceNewLine(objPdfParser.ExtractTextFromPDFBytes(reader.GetPageContent(i)) + " "))

    ''                If strSerachPhrase.IndexOf((" JPPHR ").ToLower) > 0 And strInBldr.ToString().ToLower().Length > 0 Then

    ''                    If Regex.IsMatch(strInBldr.ToString().ToLower(), "\S*(" + strSerachPhrase.ToLower.Trim(" ").Replace((" JPPHR ").ToLower, " ") + ")\S*", RegexOptions.IgnoreCase) And strSerachPhrase.Length > 2 Then
    ''                        pages.Append(i.ToString("#000") + ",")
    ''                    End If
    ''                Else
    ''                    If strSerachPhrase.Length > 2 And strInBldr.ToString().ToLower().Length > 2 Then

    ''                        If strSerachPhrase.IndexOf((" or ").ToLower) > -1 Then
    ''                            Dim strSerachPhraseList As String() = strSerachPhrase.Replace((" or ").ToLower, " ").Split(" ".ToCharArray())
    ''                            Dim FlagOr As Boolean = False
    ''                            For Each Item As String In strSerachPhraseList

    ''                                If Array.IndexOf(strInBldr.ToString().ToLower().TrimEnd.TrimStart.Replace(".", " ").Split(" ".ToCharArray()), Item.TrimStart.TrimEnd) > -1 Then
    ''                                    'If Array.IndexOf(strInBldr.ToString().ToLower().TrimEnd.TrimStart.Split(" ".ToCharArray()), Item.TrimStart.TrimEnd) > -1 Then
    ''                                    FlagOr = True

    ''                                End If
    ''                            Next

    ''                            If FlagOr Then
    ''                                pages.Append(i.ToString("#000") + ",")
    ''                            End If
    ''                        End If

    ''                        If strSerachPhrase.IndexOf((" and ").ToLower) > -1 Then
    ''                            Dim strSerachPhraseList As String() = strSerachPhrase.Replace((" and ").ToLower, " ").Split(" ".ToCharArray())
    ''                            Dim FlagAnd As Boolean = False
    ''                            Dim wordcount As Integer = 0
    ''                            For Each Item As String In strSerachPhraseList

    ''                                If Array.IndexOf(strInBldr.ToString().ToLower().TrimEnd.TrimStart.Split(" ".ToCharArray()), Item) > -1 Then
    ''                                    FlagAnd = True
    ''                                    wordcount += 1

    ''                                Else
    ''                                    FlagAnd = False
    ''                                End If

    ''                            Next

    ''                            If FlagAnd And strSerachPhraseList.Length = wordcount Then
    ''                                pages.Append(i.ToString("#000") + ",")
    ''                            End If

    ''                        End If

    ''                        If strSerachPhrase.TrimEnd.TrimStart.IndexOf((" ").ToLower) = -1 Then
    ''                            If strInBldr.ToString().ToLower().TrimEnd.TrimStart.IndexOf(strSerachPhrase.TrimEnd.TrimStart) > -1 Then
    ''                                pages.Append(i.ToString("#000") + ",")
    ''                            End If
    ''                        End If
    ''                    End If
    ''                End If

    ''            End If

    ''        Next i
    ''        sb.Append(reader.NumberOfPages.ToString() + "," + pages.ToString())
    ''    Catch ex As Exception

    ''        Return String.Empty
    ''    End Try

    ''    Return sb.ToString()
    ''End Function


    Public Shared Sub ExtractPdfPage(ByVal sourcePdf As String, ByVal fromPageNum As Integer, ByVal toPageNum As Integer, ByVal outPdf As String)
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        If fromPageNum = 0 Then
            fromPageNum = 1
        End If
        Try

            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            If toPageNum = 0 Then
                toPageNum = reader.NumberOfPages
            End If

            doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
            pdfCpy = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outPdf, IO.FileMode.Create))
            doc.Open()

            For i As Integer = fromPageNum To toPageNum Step 1
                page = pdfCpy.GetImportedPage(reader, i)
                pdfCpy.AddPage(page)
            Next
            ''For Each pageNum As Integer In pageNumbersToExtract
            ''    page = pdfCpy.GetImportedPage(reader, pageNum)
            ''    pdfCpy.AddPage(page)
            ''Next
            doc.Close()
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub ExtractPdfPageOpen(ByVal sourcePdf As String, ByVal fromPageNum As Integer, ByVal toPageNum As Integer, ByVal outPdf As String)
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

    ''Public Shared Sub PDFHighlitePhrase(ByVal PDFFilePath As String, ByVal KeywordPhrase As String, ByVal Nthpage As Int16)
    ''    'Dim pAcroPDPage As Acrobat.AcroPDPage
    ''    'Dim avobj As Acrobat.AcroAVDoc
    ''    'Dim myPDFPageHiliteObj As Acrobat.AcroHiliteList
    ''    'Dim wordToHilite As Acrobat.AcroHiliteList

    ''    Dim pAcroPDPage As Object
    ''    Dim avobj As Object
    ''    Dim myPDFPageHiliteObj As Object
    ''    Dim wordToHilite As Object

    ''    Dim acroAppObj As Object, PDFDocObj As Object, PDFJScriptObj As Object
    ''    Dim wordHilite As Object, annot As Object, props As Object, RectDim As Object
    ''    Dim colorObject(3) As Object, objRect(3) As Object
    ''    Dim iword As Integer, popupRect(0 To 3) As Integer, iTotalWords As Integer
    ''    Dim numOfPage As Integer
    ''    Dim strPDFText As New StringBuilder()
    ''    Dim word As String, sPath As String
    ''    Try
    ''        acroAppObj = CreateObject("AcroExch.App")
    ''        avobj = CreateObject("AcroExch.AVDoc")
    ''        myPDFPageHiliteObj = CreateObject("AcroExch.HiliteList")
    ''        wordToHilite = CreateObject("AcroExch.HiliteList")
    ''        myPDFPageHiliteObj.Add(0, 32767)
    ''        sPath = PDFFilePath
    ''        PDFDocObj = CreateObject("AcroExch.PDDoc")
    ''        PDFDocObj.Open(sPath)
    ''        acroAppObj.Hide()
    ''        numOfPage = PDFDocObj.GetNumPAges
    ''        word = vbNullString
    ''        PDFJScriptObj = Nothing
    ''        pAcroPDPage = PDFDocObj.AcquirePage(Nthpage)
    ''        wordHilite = pAcroPDPage.CreateWordHilite(myPDFPageHiliteObj)
    ''        PDFJScriptObj = PDFDocObj.GetJSObject
    ''        iTotalWords = wordHilite.GetNumText
    ''        iTotalWords = PDFJScriptObj.getPageNumWords(Nthpage)

    ''        For index As Integer = 0 To iTotalWords - 1
    ''            strPDFText.Append(PDFJScriptObj.getPageNthWord(Nthpage, index).ToString.TrimEnd().TrimStart())
    ''            strPDFText.Append(" ")
    ''        Next

    ''        Dim Keyword As String = ""
    ''        Dim Flage As Boolean = False

    ''        If KeywordPhrase.IndexOf((" JPPHR ").ToLower) > 0 Then

    ''            KeywordPhrase = KeywordPhrase.Replace((" JPPHR ").ToLower, " ").ToLower().TrimStart.TrimEnd

    ''            Dim strKeyWordList As String() = KeywordPhrase.ToLower().TrimStart.TrimEnd.Split(" ".ToCharArray())

    ''            Dim AryKeyword(strKeyWordList.Length - 1) As Integer

    ''            Dim wordIndex As Integer = strPDFText.ToString().ToLower().IndexOf(KeywordPhrase)

    ''            If wordIndex > -1 Then
    ''              Dim intMatchWord As Integer = 0
    ''                Dim intWordCount As Integer = 0
    ''                Dim FlageMatch As Boolean = False

    ''                For Each matchstr As String In strPDFText.ToString().ToLower().Split(" ".ToCharArray())

    ''                    If Array.IndexOf(strKeyWordList, matchstr) = intMatchWord Then
    ''                        AryKeyword(intMatchWord) = intWordCount
    ''                        intMatchWord += 1
    ''                        FlageMatch = True
    ''                    Else
    ''                        intMatchWord = 0
    ''                        FlageMatch = False
    ''                    End If


    ''                    If intMatchWord = AryKeyword.Length Then
    ''                        intMatchWord = 0

    ''                        For WordNumber As Integer = AryKeyword(0) To AryKeyword(AryKeyword.Length - 1)

    ''                            wordToHilite = CreateObject("AcroExch.HiliteList")
    ''                            wordHilite = pAcroPDPage.CreateWordHilite(myPDFPageHiliteObj)
    ''                            word = PDFJScriptObj.getPageNthWord(Nthpage, WordNumber).ToString.Trim
    ''                            wordToHilite.Add(WordNumber, 1)
    ''                            wordHilite = pAcroPDPage.CreateWordHilite(wordToHilite)
    ''                            RectDim = wordHilite.GetBoundingRect

    ''                            If Not PDFJScriptObj Is Nothing Then
    ''                                popupRect(0) = RectDim.Left
    ''                                popupRect(1) = RectDim.Top
    ''                                popupRect(2) = RectDim.Right
    ''                                popupRect(3) = RectDim.bottom
    ''                                annot = PDFJScriptObj.AddAnnot
    ''                                props = annot.getProps
    ''                                props.Type = "Square"
    ''                                annot.setProps(props)
    ''                                props = annot.getProps
    ''                                props.page = Nthpage
    ''                                props.Hidden = False
    ''                                props.Lock = True 'False
    ''                                props.Name = word
    ''                                props.noView = False
    ''                                props.opacity = 0.5
    ''                                props.ReadOnly = True ' False
    ''                                props.Style = "S"
    ''                                props.toggleNoView = False
    ''                                props.popupOpen = False
    ''                                props.rect = popupRect
    ''                                props.popupRect = popupRect
    ''                                props.strokeColor = PDFJScriptObj.Color.Gray
    ''                                props.fillColor = PDFJScriptObj.Color.Gray

    ''                                annot.setProps(props)
    ''                                wordToHilite = Nothing
    ''                            End If

    ''                            System.Windows.Forms.Application.DoEvents()

    ''                        Next
    ''                        System.Windows.Forms.Application.DoEvents()
    ''                        Array.Clear(AryKeyword, 0, AryKeyword.Length - 1)
    ''                    End If

    ''                    intWordCount += 1
    ''                Next
    ''            End If

    ''        End If

    ''        KeywordPhrase = ClassPhraseSearch.GetSplitSearchPhrase(KeywordPhrase.Trim(" ").TrimStart.TrimEnd).ToLower().ToLower().TrimStart.TrimEnd
    ''        If KeywordPhrase.IndexOf(" or ") > 0 Or KeywordPhrase.IndexOf(" not ") > 0 Or KeywordPhrase.IndexOf(" and ") > 0 Then
    ''            If KeywordPhrase.IndexOf(" and ".ToLower) > 0 Then
    ''                KeywordPhrase = KeywordPhrase.Replace(" and ", " ")
    ''            End If
    ''            If KeywordPhrase.IndexOf(" or ".ToLower) > 0 Then
    ''                KeywordPhrase = KeywordPhrase.Replace(" or ", " ")
    ''            End If
    ''            Dim strKeyWordList As String() = KeywordPhrase.ToLower().TrimStart.TrimEnd.Split(" ".ToCharArray())
    ''            Dim WordNumber As Integer = 0
    ''            For Each matchstr As String In strPDFText.ToString().ToLower().Split(" ".ToCharArray())

    ''                If Array.IndexOf(strKeyWordList, matchstr) > -1 Then

    ''                    wordToHilite = CreateObject("AcroExch.HiliteList")
    ''                    wordHilite = pAcroPDPage.CreateWordHilite(myPDFPageHiliteObj)
    ''                    word = PDFJScriptObj.getPageNthWord(Nthpage, WordNumber).ToString.Trim
    ''                    wordToHilite.Add(WordNumber, 1)
    ''                    wordHilite = pAcroPDPage.CreateWordHilite(wordToHilite)
    ''                    RectDim = wordHilite.GetBoundingRect
    ''                    If Not PDFJScriptObj Is Nothing Then
    ''                        popupRect(0) = RectDim.Left
    ''                        popupRect(1) = RectDim.Top
    ''                        popupRect(2) = RectDim.Right
    ''                        popupRect(3) = RectDim.bottom
    ''                        annot = PDFJScriptObj.AddAnnot
    ''                        props = annot.getProps
    ''                        props.Type = "Square"
    ''                        annot.setProps(props)
    ''                        props = annot.getProps
    ''                        ' props.fillColor = PDFJScriptObj.Color.red
    ''                        props.page = Nthpage
    ''                        props.Hidden = False
    ''                        props.Lock = True 'False
    ''                        props.Name = word
    ''                        props.noView = False
    ''                        props.opacity = 0.5
    ''                        props.ReadOnly = True ' False
    ''                        props.Style = "S"
    ''                        props.toggleNoView = False
    ''                        props.popupOpen = False
    ''                        props.rect = popupRect
    ''                        props.popupRect = popupRect
    ''                        props.strokeColor = PDFJScriptObj.Color.Gray
    ''                        props.fillColor = PDFJScriptObj.Color.Gray

    ''                        annot.setProps(props)
    ''                        wordToHilite = Nothing
    ''                    End If

    ''                    System.Windows.Forms.Application.DoEvents()
    ''                End If
    ''                System.Windows.Forms.Application.DoEvents()
    ''                WordNumber += 1
    ''            Next


    ''        End If

    ''    Catch ex As Exception
    ''        acroAppObj = Nothing
    ''        avobj = Nothing
    ''        myPDFPageHiliteObj = Nothing
    ''        wordToHilite = Nothing
    ''        PDFDocObj = Nothing

    ''    Finally
    ''        acroAppObj = Nothing
    ''        avobj = Nothing
    ''        myPDFPageHiliteObj = Nothing
    ''        wordToHilite = Nothing
    ''        PDFDocObj = Nothing
    ''    End Try


    ''    'Next Nthpage
    ''End Sub

    Public Shared Sub PDFHighlitePgWithWord(ByVal PDFFilePath As String, ByVal KeywordText As String, ByVal Nthpage As Int16)

        Dim pAcroPDPage As Object
        Dim avobj As Object
        Dim myPDFPageHiliteObj As Object
        Dim wordToHilite As Object

        Dim acroAppObj As Object, PDFDocObj As Object, PDFJScriptObj As Object
        Dim wordHilite As Object, annot As Object, props As Object, RectDim As Object
        Dim colorObject(3) As Object, objRect(3) As Object
        Dim iword As Integer, popupRect(0 To 3) As Integer, iTotalWords As Integer
        Dim numOfPage As Integer
        Dim word As String, sPath As String

        Try
            acroAppObj = CreateObject("AcroExch.App")
            avobj = CreateObject("AcroExch.AVDoc")
            myPDFPageHiliteObj = CreateObject("AcroExch.HiliteList")
            wordToHilite = CreateObject("AcroExch.HiliteList")
            myPDFPageHiliteObj.Add(0, 32767)
            sPath = PDFFilePath
            PDFDocObj = CreateObject("AcroExch.PDDoc")

            PDFDocObj.Open(sPath)

            acroAppObj.Hide()
            numOfPage = PDFDocObj.GetNumPAges
            word = vbNullString
            PDFJScriptObj = Nothing
            pAcroPDPage = PDFDocObj.AcquirePage(Nthpage)
            wordHilite = pAcroPDPage.CreateWordHilite(myPDFPageHiliteObj)
            PDFJScriptObj = PDFDocObj.GetJSObject
            iTotalWords = wordHilite.GetNumText
            iTotalWords = PDFJScriptObj.getPageNumWords(Nthpage)

            Dim Keyword() As String = KeywordText.Split("c")

            For ind As Integer = 0 To Keyword.Length - 1

                If Not Keyword(ind).ToString() = "" Then

                    iword = CType(Keyword(ind).ToString(), Integer)
                    wordToHilite = CreateObject("AcroExch.HiliteList")
                    wordHilite = pAcroPDPage.CreateWordHilite(myPDFPageHiliteObj)
                    word = PDFJScriptObj.getPageNthWord(Nthpage, iword).ToString.Trim
                    wordToHilite.Add(iword, 1)
                    wordHilite = pAcroPDPage.CreateWordHilite(wordToHilite)
                    RectDim = wordHilite.GetBoundingRect
                    If Not PDFJScriptObj Is Nothing Then
                        popupRect(0) = RectDim.Left
                        popupRect(1) = RectDim.Top
                        popupRect(2) = RectDim.Right
                        popupRect(3) = RectDim.bottom
                        annot = PDFJScriptObj.AddAnnot
                        props = annot.getProps
                        props.Type = "Square"
                        annot.setProps(props)
                        props = annot.getProps
                        ' props.fillColor = PDFJScriptObj.Color.red
                        props.page = Nthpage
                        props.Hidden = False
                        props.Lock = True 'False
                        props.Name = word
                        props.noView = False
                        props.opacity = 0.5
                        props.ReadOnly = True ' False
                        props.Style = "S"
                        props.toggleNoView = False
                        props.popupOpen = False
                        props.rect = popupRect
                        props.popupRect = popupRect
                        props.strokeColor = PDFJScriptObj.Color.Gray
                        props.fillColor = PDFJScriptObj.Color.Gray
                        annot.setProps(props)
                        wordToHilite = Nothing
                    End If
                End If
            Next

            'System.Windows.Forms.Application.DoEvents()

        Catch ex As Exception
            acroAppObj = Nothing
            avobj = Nothing
            myPDFPageHiliteObj = Nothing
            wordToHilite = Nothing
            PDFDocObj = Nothing

        Finally
            acroAppObj = Nothing
            avobj = Nothing
            myPDFPageHiliteObj = Nothing
            wordToHilite = Nothing
            PDFDocObj = Nothing
        End Try


        'Next Nthpage
    End Sub

   
    ''Public Shared Sub PDFHighlitePg(ByVal PDFFilePath As String, ByVal KeywordText As String, ByVal Nthpage As Int16)
    ''    'Dim pAcroPDPage As Acrobat.AcroPDPage
    ''    'Dim avobj As Acrobat.AcroAVDoc
    ''    'Dim myPDFPageHiliteObj As Acrobat.AcroHiliteList
    ''    'Dim wordToHilite As Acrobat.AcroHiliteList

    ''    Dim pAcroPDPage As Object
    ''    Dim avobj As Object
    ''    Dim myPDFPageHiliteObj As Object
    ''    Dim wordToHilite As Object

    ''    Dim acroAppObj As Object, PDFDocObj As Object, PDFJScriptObj As Object
    ''    Dim wordHilite As Object, annot As Object, props As Object, RectDim As Object
    ''    Dim colorObject(3) As Object, objRect(3) As Object
    ''    Dim iword As Integer, popupRect(0 To 3) As Integer, iTotalWords As Integer
    ''    Dim numOfPage As Integer
    ''    'Nthpage As Integer, 

    ''    Dim word As String, sPath As String

    ''    Try
    ''        acroAppObj = CreateObject("AcroExch.App")
    ''        'PDFDocObj = CreateObject("AcroExch.PDDoc")
    ''        avobj = CreateObject("AcroExch.AVDoc")
    ''        myPDFPageHiliteObj = CreateObject("AcroExch.HiliteList")
    ''        wordToHilite = CreateObject("AcroExch.HiliteList")
    ''        myPDFPageHiliteObj.Add(0, 32767)
    ''        ' RectDim = CreateObject("AcroExch.Rect")
    ''        sPath = PDFFilePath
    ''        PDFDocObj = CreateObject("AcroExch.PDDoc")

    ''        PDFDocObj.Open(sPath)
    ''        ' Hide Acrobat application so everything is done in silent
    ''        acroAppObj.Hide()
    ''        numOfPage = PDFDocObj.GetNumPAges
    ''        word = vbNullString
    ''        PDFJScriptObj = Nothing
    ''        pAcroPDPage = PDFDocObj.AcquirePage(Nthpage)
    ''        wordHilite = pAcroPDPage.CreateWordHilite(myPDFPageHiliteObj)
    ''        PDFJScriptObj = PDFDocObj.GetJSObject
    ''        iTotalWords = wordHilite.GetNumText
    ''        iTotalWords = PDFJScriptObj.getPageNumWords(Nthpage)
    ''        ''check the each word
    ''        Dim Keyword As String = ClassPhraseSearch.GetSplitSearchPhrase(KeywordText)

    ''        For iword = 0 To iTotalWords - 1
    ''            wordToHilite = CreateObject("AcroExch.HiliteList")
    ''            wordHilite = pAcroPDPage.CreateWordHilite(myPDFPageHiliteObj)

    ''            Dim flage As Boolean = False
    ''            If Keyword.ToLower.Trim(" ").Split(" ").Length > 1 Then
    ''                If Regex.IsMatch(Keyword.ToLower.Trim(" "), "\S*(" + PDFJScriptObj.getPageNthWord(Nthpage, iword).ToString.Trim(" ") + ")\S*", RegexOptions.IgnoreCase) And PDFJScriptObj.getPageNthWord(Nthpage, iword).ToString.Trim.Length > 2 Then
    ''                    flage = True
    ''                End If
    ''            Else

    ''                If Regex.IsMatch(PDFJScriptObj.getPageNthWord(Nthpage, iword).ToString.Trim(" "), "\S*(" + Keyword.ToLower.Trim(" ") + ")\S*", RegexOptions.IgnoreCase) And PDFJScriptObj.getPageNthWord(Nthpage, iword).ToString.Trim.Length > 2 Then
    ''                    flage = True
    ''                End If
    ''            End If

    ''            'If PDFJScriptObj.getPageNthWord(Nthpage, iword).ToString.Trim.ToLower.Contains(Keyword.ToLower) Then
    ''            If flage Then

    ''                word = PDFJScriptObj.getPageNthWord(Nthpage, iword).ToString.Trim

    ''                ''create obj to highlight word
    ''                wordToHilite.Add(iword, 1)
    ''                wordHilite = pAcroPDPage.CreateWordHilite(wordToHilite)
    ''                RectDim = wordHilite.GetBoundingRect

    ''                If Not PDFJScriptObj Is Nothing Then

    ''                    popupRect(0) = RectDim.Left
    ''                    popupRect(1) = RectDim.Top
    ''                    popupRect(2) = RectDim.Right
    ''                    popupRect(3) = RectDim.bottom
    ''                    annot = PDFJScriptObj.AddAnnot
    ''                    props = annot.getProps
    ''                    props.Type = "Square"
    ''                    annot.setProps(props)
    ''                    props = annot.getProps
    ''                    ' props.fillColor = PDFJScriptObj.Color.red
    ''                    props.page = Nthpage
    ''                    props.Hidden = False
    ''                    props.Lock = True 'False
    ''                    props.Name = word
    ''                    props.noView = False
    ''                    props.opacity = 0.5
    ''                    props.ReadOnly = True ' False
    ''                    props.Style = "S"
    ''                    props.toggleNoView = False
    ''                    props.popupOpen = False
    ''                    props.rect = popupRect
    ''                    props.popupRect = popupRect

    ''                    props.strokeColor = PDFJScriptObj.Color.Gray
    ''                    props.fillColor = PDFJScriptObj.Color.Gray

    ''                    annot.setProps(props)
    ''                    wordToHilite = Nothing
    ''                End If

    ''                'System.Windows.Forms.Application.DoEvents()
    ''            End If
    ''        Next iword

    ''        'System.Windows.Forms.Application.DoEvents()

    ''    Catch ex As Exception
    ''        acroAppObj = Nothing
    ''        avobj = Nothing
    ''        myPDFPageHiliteObj = Nothing
    ''        wordToHilite = Nothing
    ''        PDFDocObj = Nothing

    ''    Finally
    ''        acroAppObj = Nothing
    ''        avobj = Nothing
    ''        myPDFPageHiliteObj = Nothing
    ''        wordToHilite = Nothing
    ''        PDFDocObj = Nothing
    ''    End Try


    ''    'Next Nthpage
    ''End Sub
    Public Shared Function ReplaceNewLine(ByVal des As String) As String
        Dim strResult As String

        Try
            Dim sPattern, sReplaceText As String

            sPattern = "\r\n"
            sReplaceText = String.Empty
            strResult = Regex.Replace(RemoveDelimiter(des), sPattern, sReplaceText)

        Catch ex As Exception

        End Try
        Return strResult
    End Function

    Public Shared Function RemoveDelimiter(ByVal des As String) As String
        Dim strorigFileName As String
        Dim intCounter As Integer
        Dim arrSpecialChar() As String = {".", ",", "<", ">", ":", "?", """", "/", "{", "[", "}", "]", "`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", " ", "\"}
        strorigFileName = des
        intCounter = 0
        Dim i As Integer
        For i = 0 To arrSpecialChar.Length - 1

            Do Until intCounter = 29
                des = Replace(strorigFileName, arrSpecialChar(i), " ")
                intCounter = intCounter + 1
                strorigFileName = des
            Loop
            intCounter = 0
        Next
        Return strorigFileName

    End Function
    ''Public Shared Function GetFileDetails(ByVal ISBN As String, Optional ByVal SNMBR As Integer = -1, Optional ByVal filename As String = "") As DataSet
    ''    Dim objDataSet As New DataSet()
    ''    Dim strConnection As String = ConfigurationManager.ConnectionStrings("EBookConnectionStr").ToString()

    ''    Dim objSQLConnection As New SqlConnection(strConnection)
    ''    Try
    ''        objSQLConnection.Open()
    ''        Dim objSQLCommand As New SqlCommand()
    ''        objSQLCommand.Connection = objSQLConnection
    ''        objSQLCommand.CommandText = "GetChapterList"
    ''        objSQLCommand.CommandType = CommandType.StoredProcedure
    ''        Dim objSQLParameter As New SqlParameter("@ISBN", ISBN)
    ''        Dim objSQLParameterSN As New SqlParameter("@SNMBR", SNMBR)
    ''        Dim objSQLParameterfilename As New SqlParameter("@FILENAME", filename)
    ''        objSQLCommand.Parameters.Add(objSQLParameter)
    ''        objSQLCommand.Parameters.Add(objSQLParameterSN)
    ''        objSQLCommand.Parameters.Add(objSQLParameterfilename)
    ''        Dim objSQLDataAdaptor As New SqlDataAdapter(objSQLCommand)
    ''        objSQLDataAdaptor.Fill(objDataSet)


    ''    Catch ex As Exception
    ''    Finally

    ''    End Try
    ''    objSQLConnection.Close()

    ''    Return objDataSet
    ''End Function


End Class
