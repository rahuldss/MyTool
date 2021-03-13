Imports iTextSharp
Imports iTextSharp.text
Imports iTextSharp.text.pdf


Public Class frmBarcode

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            ''1.
            'Dim f As New Font("IDAutomationUPCEAN", 20, FontStyle.Regular, GraphicsUnit.Pixel, 100)
            ''Call PrintEANBarCode("12345678", pbBarcode, , , , , f)

            'Dim FE As New FontEncoder
            'MsgBox(FE.EAN8("12345678"))
            'Call PrintEANBarCode(FE.EAN8("12345678"), pbBarcode, , , , , f)

            '2.
            Dim BCode As New BarcodeCodabar
            BCode.Code = "a1234567890123d"
            Dim bc As System.Drawing.Image = BCode.CreateDrawingImage(Drawing.Color.Black, Drawing.Color.White)

            pbBarcode.Image = bc
            pbBarcode.BackColor = Drawing.Color.Beige
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ''#Region "Barcode"

    ''    Private Function EAN2Bin(ByVal strEANCode As String) As String
    ''        Dim K As Integer
    ''        Dim strAux As String
    ''        Dim strExit As String
    ''        Dim strCode As String

    ''        strEANCode = Trim(strEANCode)
    ''        strAux = strEANCode
    ''        If (strAux.Length <> 13) And (strAux.Length <> 8) Then
    ''            Err.Raise(5, "EAN2Bin", "Invalid EAN Code")
    ''        End If
    ''        For K = 0 To strEANCode.Length - 1
    ''            Select Case (strAux.Chars(K).ToString)
    ''                Case Is < "0", Is > "9"
    ''                    Err.Raise(5, "EAN2Bin", "Invalid char on EAN Code")
    ''            End Select
    ''        Next
    ''        If (strAux.Length = 13) Then
    ''            strAux = Mid(strAux, 2)
    ''            Select Case CInt(Strings.Left(strEANCode, 1))
    ''                Case 0
    ''                    strCode = "000000"
    ''                Case 1
    ''                    strCode = "001011"
    ''                Case 2
    ''                    strCode = "001101"
    ''                Case 3
    ''                    strCode = "001110"
    ''                Case 4
    ''                    strCode = "010011"
    ''                Case 5
    ''                    strCode = "011001"
    ''                Case 6
    ''                    strCode = "011100"
    ''                Case 7
    ''                    strCode = "010101"
    ''                Case 8
    ''                    strCode = "010110"
    ''                Case 9
    ''                    strCode = "011010"
    ''            End Select
    ''        Else
    ''            strCode = "0000"
    ''        End If
    ''        '* The EAN BarCode starts with a "guardian" 
    ''        strExit = "000101"
    ''        '* First half of the code
    ''        For K = 1 To Len(strAux) \ 2
    ''            Select Case CInt(Mid(strAux, K, 1))
    ''                Case 0
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0001101", "0100111")
    ''                Case 1
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0011001", "0110011")
    ''                Case 2
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0010011", "0011011")
    ''                Case 3
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0111101", "0100001")
    ''                Case 4
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0100011", "0011101")
    ''                Case 5
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0110001", "0111001")
    ''                Case 6
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0101111", "0000101")
    ''                Case 7
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0111011", "0010001")
    ''                Case 8
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0110111", "0001001")
    ''                Case 9
    ''                    strExit &= IIf(Mid(strCode, K, 1) = "0", "0001011", "0010111")
    ''            End Select
    ''        Next K
    ''        '* Middle "guardian" separator
    ''        strExit &= "01010"
    ''        '* Second half of the code
    ''        For K = Len(strAux) \ 2 + 1 To Len(strAux)
    ''            Select Case CInt(Mid(strAux, K, 1))
    ''                Case 0
    ''                    strExit &= "1110010"
    ''                Case 1
    ''                    strExit &= "1100110"
    ''                Case 2
    ''                    strExit &= "1101100"
    ''                Case 3
    ''                    strExit &= "1000010"
    ''                Case 4
    ''                    strExit &= "1011100"
    ''                Case 5
    ''                    strExit &= "1001110"
    ''                Case 6
    ''                    strExit &= "1010000"
    ''                Case 7
    ''                    strExit &= "1000100"
    ''                Case 8
    ''                    strExit &= "1001000"
    ''                Case 9
    ''                    strExit &= "1110100"
    ''            End Select
    ''        Next K
    ''        strExit &= "101000"
    ''        EAN2Bin = strExit
    ''    End Function

    ''    Public Sub PrintEANBarCode(ByVal strEANCode As String, ByVal objPicBox As PictureBox, _
    ''                                    Optional ByVal sngX1 As Single = (-1), _
    ''                                    Optional ByVal sngY1 As Single = (-1), _
    ''                                    Optional ByVal sngX2 As Single = (-1), _
    ''                                    Optional ByVal sngY2 As Single = (-1), _
    ''                                    Optional ByVal FontForText As Font = Nothing)
    ''        Dim K As Single
    ''        Dim sngPosX As Single
    ''        Dim sngPosY As Single
    ''        Dim sngScaleX As Single
    ''        Dim strEANBin As String
    ''        Dim strFormat As New StringFormat
    ''        strEANBin = EAN2Bin(strEANCode)
    ''        If (FontForText Is Nothing) Then
    ''            FontForText = New Font("Courier New", 10)
    ''        End If
    ''        If sngX1 = (-1) Then sngX1 = 0
    ''        If sngY1 = (-1) Then sngY1 = 0
    ''        If sngX2 = (-1) Then sngX2 = objPicBox.Width
    ''        If sngY2 = (-1) Then sngY2 = objPicBox.Height
    ''        sngPosX = sngX1
    ''        sngPosY = sngY2 - CSng(1.5 * FontForText.Height)
    ''        objPicBox.CreateGraphics.FillRectangle(New System.Drawing.SolidBrush(objPicBox.BackColor.Blue), sngX1, sngY1, sngX2 - sngX1, sngY2 - sngY1)
    ''        For K = 1 To Len(strEANBin)
    ''            If Mid(strEANBin, K, 1) = "1" Then
    ''                objPicBox.CreateGraphics.FillRectangle(New System.Drawing.SolidBrush(objPicBox.ForeColor.Red), sngPosX, sngY1, sngScaleX, sngPosY)
    ''            End If
    ''            sngPosX = sngX1 + (K * sngScaleX)
    ''        Next K
    ''        strFormat.Alignment = StringAlignment.Center
    ''        strFormat.FormatFlags = StringFormatFlags.NoWrap
    ''        objPicBox.CreateGraphics.DrawString(strEANCode, FontForText, New System.Drawing.SolidBrush(objPicBox.ForeColor.Orange), CSng((sngX2 - sngX1) / 2), CSng(sngY2 - FontForText.Height), strFormat)
    ''        objPicBox.Update()
    ''    End Sub
    ''#End Region

End Class