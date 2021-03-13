<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRnD
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRnD))
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.cmdReadPDF = New System.Windows.Forms.Button
        Me.cmdPDFExtract = New System.Windows.Forms.Button
        Me.cmdCombination = New System.Windows.Forms.Button
        Me.cmdInsertImage = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.cmdGetImage = New System.Windows.Forms.Button
        Me.CmdPDFRead = New System.Windows.Forms.Button
        Me.cmdInsertAllTypes = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.ChkLst1 = New System.Windows.Forms.CheckedListBox
        Me.DGV1 = New System.Windows.Forms.DataGridView
        Me.cmdTestChecked = New System.Windows.Forms.Button
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
        Me.LineShape1 = New Microsoft.VisualBasic.PowerPacks.LineShape
        Me.chkSelectAll = New System.Windows.Forms.CheckBox
        Me.cmdFill = New System.Windows.Forms.Button
        Me.DGVChk1 = New System.Windows.Forms.DataGridView
        Me.cmdCheckSelected = New System.Windows.Forms.Button
        Me.cmdError = New System.Windows.Forms.Button
        Me.txtSqlScripts = New System.Windows.Forms.TextBox
        Me.cmdExecuteScripts = New System.Windows.Forms.Button
        Me.cmdToInt16 = New System.Windows.Forms.Button
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdInsertMultipleRows2Col = New System.Windows.Forms.Button
        Me.cmdInsertMultipleRows1Col = New System.Windows.Forms.Button
        Me.cmdLamda = New System.Windows.Forms.Button
        Me.cmdAttachDBFile = New System.Windows.Forms.Button
        Me.cmdInsertAllTypesNEW = New System.Windows.Forms.Button
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGVChk1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(135, 89)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 107)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(137, 87)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(308, 57)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Label1"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(12, 200)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(137, 87)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(12, 293)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(135, 48)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(153, 293)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(315, 21)
        Me.ComboBox1.TabIndex = 5
        '
        'cmdReadPDF
        '
        Me.cmdReadPDF.Location = New System.Drawing.Point(12, 347)
        Me.cmdReadPDF.Name = "cmdReadPDF"
        Me.cmdReadPDF.Size = New System.Drawing.Size(135, 48)
        Me.cmdReadPDF.TabIndex = 6
        Me.cmdReadPDF.Text = "PDF Read"
        Me.cmdReadPDF.UseVisualStyleBackColor = True
        '
        'cmdPDFExtract
        '
        Me.cmdPDFExtract.Location = New System.Drawing.Point(12, 400)
        Me.cmdPDFExtract.Name = "cmdPDFExtract"
        Me.cmdPDFExtract.Size = New System.Drawing.Size(135, 48)
        Me.cmdPDFExtract.TabIndex = 7
        Me.cmdPDFExtract.Text = "PDF Extract"
        Me.cmdPDFExtract.UseVisualStyleBackColor = True
        '
        'cmdCombination
        '
        Me.cmdCombination.Location = New System.Drawing.Point(249, 377)
        Me.cmdCombination.Name = "cmdCombination"
        Me.cmdCombination.Size = New System.Drawing.Size(98, 34)
        Me.cmdCombination.TabIndex = 8
        Me.cmdCombination.Text = "Combination"
        Me.cmdCombination.UseVisualStyleBackColor = True
        '
        'cmdInsertImage
        '
        Me.cmdInsertImage.Location = New System.Drawing.Point(249, 419)
        Me.cmdInsertImage.Name = "cmdInsertImage"
        Me.cmdInsertImage.Size = New System.Drawing.Size(98, 34)
        Me.cmdInsertImage.TabIndex = 9
        Me.cmdInsertImage.Text = "&Insert Image"
        Me.cmdInsertImage.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(319, 115)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(124, 131)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'cmdGetImage
        '
        Me.cmdGetImage.Location = New System.Drawing.Point(362, 419)
        Me.cmdGetImage.Name = "cmdGetImage"
        Me.cmdGetImage.Size = New System.Drawing.Size(98, 34)
        Me.cmdGetImage.TabIndex = 11
        Me.cmdGetImage.Text = "&Get Image"
        Me.cmdGetImage.UseVisualStyleBackColor = True
        '
        'CmdPDFRead
        '
        Me.CmdPDFRead.Location = New System.Drawing.Point(362, 379)
        Me.CmdPDFRead.Name = "CmdPDFRead"
        Me.CmdPDFRead.Size = New System.Drawing.Size(98, 31)
        Me.CmdPDFRead.TabIndex = 12
        Me.CmdPDFRead.Text = "Read PDF"
        Me.CmdPDFRead.UseVisualStyleBackColor = True
        '
        'cmdInsertAllTypes
        '
        Me.cmdInsertAllTypes.Location = New System.Drawing.Point(303, 11)
        Me.cmdInsertAllTypes.Name = "cmdInsertAllTypes"
        Me.cmdInsertAllTypes.Size = New System.Drawing.Size(82, 31)
        Me.cmdInsertAllTypes.TabIndex = 13
        Me.cmdInsertAllTypes.Text = "InsertAllTypes"
        Me.cmdInsertAllTypes.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(403, 51)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(77, 20)
        Me.TextBox1.TabIndex = 14
        '
        'ChkLst1
        '
        Me.ChkLst1.FormattingEnabled = True
        Me.ChkLst1.Location = New System.Drawing.Point(559, 12)
        Me.ChkLst1.Name = "ChkLst1"
        Me.ChkLst1.Size = New System.Drawing.Size(235, 94)
        Me.ChkLst1.TabIndex = 15
        '
        'DGV1
        '
        Me.DGV1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV1.Location = New System.Drawing.Point(545, 138)
        Me.DGV1.Name = "DGV1"
        Me.DGV1.Size = New System.Drawing.Size(285, 122)
        Me.DGV1.TabIndex = 16
        '
        'cmdTestChecked
        '
        Me.cmdTestChecked.Location = New System.Drawing.Point(638, 412)
        Me.cmdTestChecked.Name = "cmdTestChecked"
        Me.cmdTestChecked.Size = New System.Drawing.Size(87, 26)
        Me.cmdTestChecked.TabIndex = 17
        Me.cmdTestChecked.Text = "Test Checked"
        Me.cmdTestChecked.UseVisualStyleBackColor = True
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape1})
        Me.ShapeContainer1.Size = New System.Drawing.Size(860, 624)
        Me.ShapeContainer1.TabIndex = 18
        Me.ShapeContainer1.TabStop = False
        '
        'LineShape1
        '
        Me.LineShape1.Name = "LineShape1"
        Me.LineShape1.X1 = 508
        Me.LineShape1.X2 = 508
        Me.LineShape1.Y1 = 20
        Me.LineShape1.Y2 = 443
        '
        'chkSelectAll
        '
        Me.chkSelectAll.AutoSize = True
        Me.chkSelectAll.Location = New System.Drawing.Point(710, 115)
        Me.chkSelectAll.Name = "chkSelectAll"
        Me.chkSelectAll.Size = New System.Drawing.Size(70, 17)
        Me.chkSelectAll.TabIndex = 19
        Me.chkSelectAll.Text = "Select All"
        Me.chkSelectAll.UseVisualStyleBackColor = True
        '
        'cmdFill
        '
        Me.cmdFill.Location = New System.Drawing.Point(545, 412)
        Me.cmdFill.Name = "cmdFill"
        Me.cmdFill.Size = New System.Drawing.Size(87, 26)
        Me.cmdFill.TabIndex = 20
        Me.cmdFill.Text = "Fill"
        Me.cmdFill.UseVisualStyleBackColor = True
        '
        'DGVChk1
        '
        Me.DGVChk1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVChk1.Location = New System.Drawing.Point(545, 271)
        Me.DGVChk1.Name = "DGVChk1"
        Me.DGVChk1.Size = New System.Drawing.Size(284, 123)
        Me.DGVChk1.TabIndex = 21
        '
        'cmdCheckSelected
        '
        Me.cmdCheckSelected.Location = New System.Drawing.Point(733, 412)
        Me.cmdCheckSelected.Name = "cmdCheckSelected"
        Me.cmdCheckSelected.Size = New System.Drawing.Size(66, 26)
        Me.cmdCheckSelected.TabIndex = 22
        Me.cmdCheckSelected.Text = "Check Selected"
        Me.cmdCheckSelected.UseVisualStyleBackColor = True
        '
        'cmdError
        '
        Me.cmdError.Location = New System.Drawing.Point(196, 11)
        Me.cmdError.Name = "cmdError"
        Me.cmdError.Size = New System.Drawing.Size(90, 31)
        Me.cmdError.TabIndex = 23
        Me.cmdError.Text = "Error"
        Me.cmdError.UseVisualStyleBackColor = True
        '
        'txtSqlScripts
        '
        Me.txtSqlScripts.Location = New System.Drawing.Point(12, 459)
        Me.txtSqlScripts.Multiline = True
        Me.txtSqlScripts.Name = "txtSqlScripts"
        Me.txtSqlScripts.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSqlScripts.Size = New System.Drawing.Size(299, 153)
        Me.txtSqlScripts.TabIndex = 24
        '
        'cmdExecuteScripts
        '
        Me.cmdExecuteScripts.Location = New System.Drawing.Point(316, 459)
        Me.cmdExecuteScripts.Name = "cmdExecuteScripts"
        Me.cmdExecuteScripts.Size = New System.Drawing.Size(104, 36)
        Me.cmdExecuteScripts.TabIndex = 25
        Me.cmdExecuteScripts.Text = "Execute Scripts"
        Me.cmdExecuteScripts.UseVisualStyleBackColor = True
        '
        'cmdToInt16
        '
        Me.cmdToInt16.Location = New System.Drawing.Point(404, 81)
        Me.cmdToInt16.Name = "cmdToInt16"
        Me.cmdToInt16.Size = New System.Drawing.Size(75, 25)
        Me.cmdToInt16.TabIndex = 26
        Me.cmdToInt16.Text = "ToInt16"
        Me.cmdToInt16.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(360, 54)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(39, 20)
        Me.TextBox2.TabIndex = 27
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdInsertMultipleRows2Col)
        Me.GroupBox1.Controls.Add(Me.cmdInsertMultipleRows1Col)
        Me.GroupBox1.Location = New System.Drawing.Point(169, 75)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(131, 171)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "DB ""AB"" Used"
        '
        'cmdInsertMultipleRows2Col
        '
        Me.cmdInsertMultipleRows2Col.Location = New System.Drawing.Point(4, 107)
        Me.cmdInsertMultipleRows2Col.Name = "cmdInsertMultipleRows2Col"
        Me.cmdInsertMultipleRows2Col.Size = New System.Drawing.Size(122, 55)
        Me.cmdInsertMultipleRows2Col.TabIndex = 31
        Me.cmdInsertMultipleRows2Col.Text = "Insert Multiple Rows 2 Col Using XML"
        Me.cmdInsertMultipleRows2Col.UseVisualStyleBackColor = True
        '
        'cmdInsertMultipleRows1Col
        '
        Me.cmdInsertMultipleRows1Col.Location = New System.Drawing.Point(4, 46)
        Me.cmdInsertMultipleRows1Col.Name = "cmdInsertMultipleRows1Col"
        Me.cmdInsertMultipleRows1Col.Size = New System.Drawing.Size(122, 55)
        Me.cmdInsertMultipleRows1Col.TabIndex = 30
        Me.cmdInsertMultipleRows1Col.Text = "Insert Multiple Rows 1 Col Using XML"
        Me.cmdInsertMultipleRows1Col.UseVisualStyleBackColor = True
        '
        'cmdLamda
        '
        Me.cmdLamda.Location = New System.Drawing.Point(475, 477)
        Me.cmdLamda.Name = "cmdLamda"
        Me.cmdLamda.Size = New System.Drawing.Size(137, 36)
        Me.cmdLamda.TabIndex = 31
        Me.cmdLamda.Text = "Lamda Expression"
        Me.cmdLamda.UseVisualStyleBackColor = True
        '
        'cmdAttachDBFile
        '
        Me.cmdAttachDBFile.Location = New System.Drawing.Point(646, 467)
        Me.cmdAttachDBFile.Name = "cmdAttachDBFile"
        Me.cmdAttachDBFile.Size = New System.Drawing.Size(133, 27)
        Me.cmdAttachDBFile.TabIndex = 32
        Me.cmdAttachDBFile.Text = "Attach DB File"
        Me.cmdAttachDBFile.UseVisualStyleBackColor = True
        '
        'cmdInsertAllTypesNEW
        '
        Me.cmdInsertAllTypesNEW.Location = New System.Drawing.Point(391, 12)
        Me.cmdInsertAllTypesNEW.Name = "cmdInsertAllTypesNEW"
        Me.cmdInsertAllTypesNEW.Size = New System.Drawing.Size(111, 31)
        Me.cmdInsertAllTypesNEW.TabIndex = 33
        Me.cmdInsertAllTypesNEW.Text = "InsertAllTypes NEW"
        Me.cmdInsertAllTypesNEW.UseVisualStyleBackColor = True
        '
        'frmRnD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(860, 624)
        Me.Controls.Add(Me.cmdInsertAllTypesNEW)
        Me.Controls.Add(Me.cmdAttachDBFile)
        Me.Controls.Add(Me.cmdLamda)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.cmdToInt16)
        Me.Controls.Add(Me.cmdExecuteScripts)
        Me.Controls.Add(Me.txtSqlScripts)
        Me.Controls.Add(Me.cmdError)
        Me.Controls.Add(Me.cmdCheckSelected)
        Me.Controls.Add(Me.DGVChk1)
        Me.Controls.Add(Me.cmdFill)
        Me.Controls.Add(Me.chkSelectAll)
        Me.Controls.Add(Me.cmdTestChecked)
        Me.Controls.Add(Me.DGV1)
        Me.Controls.Add(Me.ChkLst1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.cmdInsertAllTypes)
        Me.Controls.Add(Me.CmdPDFRead)
        Me.Controls.Add(Me.cmdGetImage)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdInsertImage)
        Me.Controls.Add(Me.cmdCombination)
        Me.Controls.Add(Me.cmdPDFExtract)
        Me.Controls.Add(Me.cmdReadPDF)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmRnD"
        Me.Text = "R & D"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGVChk1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmdReadPDF As System.Windows.Forms.Button
    Friend WithEvents cmdPDFExtract As System.Windows.Forms.Button
    Friend WithEvents cmdCombination As System.Windows.Forms.Button
    Friend WithEvents cmdInsertImage As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdGetImage As System.Windows.Forms.Button
    Friend WithEvents CmdPDFRead As System.Windows.Forms.Button
    Friend WithEvents cmdInsertAllTypes As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ChkLst1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents DGV1 As System.Windows.Forms.DataGridView
    Friend WithEvents cmdTestChecked As System.Windows.Forms.Button
    Friend WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents LineShape1 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents chkSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents cmdFill As System.Windows.Forms.Button
    Friend WithEvents DGVChk1 As System.Windows.Forms.DataGridView
    Friend WithEvents cmdCheckSelected As System.Windows.Forms.Button
    Friend WithEvents cmdError As System.Windows.Forms.Button
    Friend WithEvents txtSqlScripts As System.Windows.Forms.TextBox
    Friend WithEvents cmdExecuteScripts As System.Windows.Forms.Button
    Friend WithEvents cmdToInt16 As System.Windows.Forms.Button
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdInsertMultipleRows2Col As System.Windows.Forms.Button
    Friend WithEvents cmdInsertMultipleRows1Col As System.Windows.Forms.Button
    Friend WithEvents cmdLamda As System.Windows.Forms.Button
    Friend WithEvents cmdAttachDBFile As System.Windows.Forms.Button
    Friend WithEvents cmdInsertAllTypesNEW As System.Windows.Forms.Button

End Class
