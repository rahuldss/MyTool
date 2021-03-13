<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenerateInserts
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGenerateInserts))
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.rtbSelectQuery = New System.Windows.Forms.RichTextBox()
        Me.gbDBOptionsSql = New System.Windows.Forms.GroupBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cboTableDest = New System.Windows.Forms.ComboBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.chkDBFile = New System.Windows.Forms.CheckBox()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtUID = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblMSG = New System.Windows.Forms.Label()
        Me.cmdGenerateInserts = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.chkSelectAll = New System.Windows.Forms.CheckBox()
        Me.cmdReset = New System.Windows.Forms.Button()
        Me.lstColumnsToExclude = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LstColumnsDest = New System.Windows.Forms.CheckedListBox()
        Me.OFDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.cmdDisplayRecords = New System.Windows.Forms.Button()
        Me.chkForAllTables = New System.Windows.Forms.CheckBox()
        Me.pb2 = New System.Windows.Forms.ProgressBar()
        Me.Pb1 = New System.Windows.Forms.ProgressBar()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtFrom_Where = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtTop = New System.Windows.Forms.TextBox()
        Me.lblTop = New System.Windows.Forms.Label()
        Me.txtRowFrom = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtRowTo = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblTableName = New System.Windows.Forms.Label()
        Me.txtIDColumnName = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.GroupBox2.SuspendLayout()
        Me.gbDBOptionsSql.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.rtbSelectQuery)
        Me.GroupBox2.Controls.Add(Me.gbDBOptionsSql)
        Me.GroupBox2.Controls.Add(Me.lblMSG)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(489, 357)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(16, 200)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(77, 13)
        Me.Label10.TabIndex = 40
        Me.Label10.Text = "Select Query : "
        '
        'rtbSelectQuery
        '
        Me.rtbSelectQuery.Location = New System.Drawing.Point(18, 216)
        Me.rtbSelectQuery.Name = "rtbSelectQuery"
        Me.rtbSelectQuery.Size = New System.Drawing.Size(433, 135)
        Me.rtbSelectQuery.TabIndex = 32
        Me.rtbSelectQuery.Text = ""
        '
        'gbDBOptionsSql
        '
        Me.gbDBOptionsSql.Controls.Add(Me.Label9)
        Me.gbDBOptionsSql.Controls.Add(Me.cboTableDest)
        Me.gbDBOptionsSql.Controls.Add(Me.cmdOK)
        Me.gbDBOptionsSql.Controls.Add(Me.chkDBFile)
        Me.gbDBOptionsSql.Controls.Add(Me.txtDatabase)
        Me.gbDBOptionsSql.Controls.Add(Me.Label5)
        Me.gbDBOptionsSql.Controls.Add(Me.txtPassword)
        Me.gbDBOptionsSql.Controls.Add(Me.Label3)
        Me.gbDBOptionsSql.Controls.Add(Me.txtUID)
        Me.gbDBOptionsSql.Controls.Add(Me.Label6)
        Me.gbDBOptionsSql.Controls.Add(Me.txtServer)
        Me.gbDBOptionsSql.Controls.Add(Me.Label7)
        Me.gbDBOptionsSql.Location = New System.Drawing.Point(9, 15)
        Me.gbDBOptionsSql.Name = "gbDBOptionsSql"
        Me.gbDBOptionsSql.Size = New System.Drawing.Size(457, 180)
        Me.gbDBOptionsSql.TabIndex = 31
        Me.gbDBOptionsSql.TabStop = False
        Me.gbDBOptionsSql.Text = "Select Destination"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 156)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(67, 13)
        Me.Label9.TabIndex = 39
        Me.Label9.Text = "Select Table"
        '
        'cboTableDest
        '
        Me.cboTableDest.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTableDest.FormattingEnabled = True
        Me.cboTableDest.Location = New System.Drawing.Point(109, 153)
        Me.cboTableDest.Name = "cboTableDest"
        Me.cboTableDest.Size = New System.Drawing.Size(240, 21)
        Me.cboTableDest.TabIndex = 38
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(399, 116)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(52, 26)
        Me.cmdOK.TabIndex = 37
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'chkDBFile
        '
        Me.chkDBFile.Location = New System.Drawing.Point(109, 112)
        Me.chkDBFile.Name = "chkDBFile"
        Me.chkDBFile.Size = New System.Drawing.Size(282, 18)
        Me.chkDBFile.TabIndex = 36
        Me.chkDBFile.Text = "Use Connection String From Server Option (.MDF File)"
        Me.chkDBFile.UseVisualStyleBackColor = True
        '
        'txtDatabase
        '
        Me.txtDatabase.Location = New System.Drawing.Point(109, 82)
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(271, 20)
        Me.txtDatabase.TabIndex = 34
        Me.txtDatabase.Text = "JpData_MyShop"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 85)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(53, 13)
        Me.Label5.TabIndex = 33
        Me.Label5.Text = "Database"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(109, 60)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(271, 20)
        Me.txtPassword.TabIndex = 32
        Me.txtPassword.Text = "sa123"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 63)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 13)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "Password"
        '
        'txtUID
        '
        Me.txtUID.Location = New System.Drawing.Point(109, 37)
        Me.txtUID.Name = "txtUID"
        Me.txtUID.Size = New System.Drawing.Size(271, 20)
        Me.txtUID.TabIndex = 30
        Me.txtUID.Text = "sa"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(26, 13)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "UID"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(109, 14)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(271, 20)
        Me.txtServer.TabIndex = 28
        Me.txtServer.Text = "TSI_DEV_02"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 17)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(38, 13)
        Me.Label7.TabIndex = 27
        Me.Label7.Text = "Server"
        '
        'lblMSG
        '
        Me.lblMSG.Location = New System.Drawing.Point(10, 288)
        Me.lblMSG.Name = "lblMSG"
        Me.lblMSG.Size = New System.Drawing.Size(456, 57)
        Me.lblMSG.TabIndex = 2
        Me.lblMSG.Text = "."
        '
        'cmdGenerateInserts
        '
        Me.cmdGenerateInserts.Location = New System.Drawing.Point(336, 12)
        Me.cmdGenerateInserts.Name = "cmdGenerateInserts"
        Me.cmdGenerateInserts.Size = New System.Drawing.Size(150, 23)
        Me.cmdGenerateInserts.TabIndex = 1
        Me.cmdGenerateInserts.Text = "&Generate Inserts"
        Me.cmdGenerateInserts.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.chkSelectAll)
        Me.GroupBox3.Controls.Add(Me.cmdReset)
        Me.GroupBox3.Controls.Add(Me.lstColumnsToExclude)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.LstColumnsDest)
        Me.GroupBox3.Location = New System.Drawing.Point(511, 12)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(362, 599)
        Me.GroupBox3.TabIndex = 6
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Map Columns"
        '
        'chkSelectAll
        '
        Me.chkSelectAll.AutoSize = True
        Me.chkSelectAll.Location = New System.Drawing.Point(10, 421)
        Me.chkSelectAll.Name = "chkSelectAll"
        Me.chkSelectAll.Size = New System.Drawing.Size(70, 17)
        Me.chkSelectAll.TabIndex = 32
        Me.chkSelectAll.Text = "Select &All"
        Me.chkSelectAll.UseVisualStyleBackColor = True
        '
        'cmdReset
        '
        Me.cmdReset.Location = New System.Drawing.Point(90, 420)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(183, 20)
        Me.cmdReset.TabIndex = 31
        Me.cmdReset.Text = "Reset"
        Me.cmdReset.UseVisualStyleBackColor = True
        '
        'lstColumnsToExclude
        '
        Me.lstColumnsToExclude.FormattingEnabled = True
        Me.lstColumnsToExclude.Location = New System.Drawing.Point(10, 445)
        Me.lstColumnsToExclude.Name = "lstColumnsToExclude"
        Me.lstColumnsToExclude.Size = New System.Drawing.Size(345, 147)
        Me.lstColumnsToExclude.TabIndex = 30
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(160, 13)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "List of Columns From Destination"
        '
        'LstColumnsDest
        '
        Me.LstColumnsDest.CheckOnClick = True
        Me.LstColumnsDest.FormattingEnabled = True
        Me.LstColumnsDest.Location = New System.Drawing.Point(10, 31)
        Me.LstColumnsDest.Name = "LstColumnsDest"
        Me.LstColumnsDest.Size = New System.Drawing.Size(345, 379)
        Me.LstColumnsDest.TabIndex = 28
        '
        'OFDialog1
        '
        Me.OFDialog1.AddExtension = False
        Me.OFDialog1.DefaultExt = "*.xls"
        Me.OFDialog1.Filter = "Excel Files (*.xls)|*.xls|New Excel Files (*.xlsx)|*.xlsx"
        '
        'cmdDisplayRecords
        '
        Me.cmdDisplayRecords.Location = New System.Drawing.Point(350, 391)
        Me.cmdDisplayRecords.Name = "cmdDisplayRecords"
        Me.cmdDisplayRecords.Size = New System.Drawing.Size(150, 23)
        Me.cmdDisplayRecords.TabIndex = 7
        Me.cmdDisplayRecords.Text = "&Display Records"
        Me.cmdDisplayRecords.UseVisualStyleBackColor = True
        '
        'chkForAllTables
        '
        Me.chkForAllTables.AutoSize = True
        Me.chkForAllTables.Location = New System.Drawing.Point(240, 16)
        Me.chkForAllTables.Name = "chkForAllTables"
        Me.chkForAllTables.Size = New System.Drawing.Size(90, 17)
        Me.chkForAllTables.TabIndex = 8
        Me.chkForAllTables.Text = "For All Tables"
        Me.chkForAllTables.UseVisualStyleBackColor = True
        '
        'pb2
        '
        Me.pb2.Location = New System.Drawing.Point(18, 517)
        Me.pb2.Name = "pb2"
        Me.pb2.Size = New System.Drawing.Size(481, 17)
        Me.pb2.TabIndex = 9
        '
        'Pb1
        '
        Me.Pb1.Location = New System.Drawing.Point(18, 540)
        Me.Pb1.Name = "Pb1"
        Me.Pb1.Size = New System.Drawing.Size(481, 17)
        Me.Pb1.TabIndex = 10
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtFrom_Where)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtTop)
        Me.GroupBox1.Controls.Add(Me.lblTop)
        Me.GroupBox1.Controls.Add(Me.cmdGenerateInserts)
        Me.GroupBox1.Controls.Add(Me.chkForAllTables)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 415)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(489, 64)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Generate Inserts"
        '
        'txtFrom_Where
        '
        Me.txtFrom_Where.Location = New System.Drawing.Point(130, 39)
        Me.txtFrom_Where.Name = "txtFrom_Where"
        Me.txtFrom_Where.Size = New System.Drawing.Size(355, 20)
        Me.txtFrom_Where.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "From (Where Clause) :"
        '
        'txtTop
        '
        Me.txtTop.Location = New System.Drawing.Point(62, 16)
        Me.txtTop.Name = "txtTop"
        Me.txtTop.Size = New System.Drawing.Size(160, 20)
        Me.txtTop.TabIndex = 10
        '
        'lblTop
        '
        Me.lblTop.AutoSize = True
        Me.lblTop.Location = New System.Drawing.Point(16, 17)
        Me.lblTop.Name = "lblTop"
        Me.lblTop.Size = New System.Drawing.Size(35, 13)
        Me.lblTop.TabIndex = 9
        Me.lblTop.Text = "Top : "
        '
        'txtRowFrom
        '
        Me.txtRowFrom.Location = New System.Drawing.Point(71, 390)
        Me.txtRowFrom.Name = "txtRowFrom"
        Me.txtRowFrom.Size = New System.Drawing.Size(85, 20)
        Me.txtRowFrom.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 391)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Row From : "
        '
        'txtRowTo
        '
        Me.txtRowTo.Location = New System.Drawing.Point(190, 391)
        Me.txtRowTo.Name = "txtRowTo"
        Me.txtRowTo.Size = New System.Drawing.Size(85, 20)
        Me.txtRowTo.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(162, 392)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(29, 13)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "To : "
        '
        'lblTableName
        '
        Me.lblTableName.AutoSize = True
        Me.lblTableName.Location = New System.Drawing.Point(20, 490)
        Me.lblTableName.Name = "lblTableName"
        Me.lblTableName.Size = New System.Drawing.Size(16, 13)
        Me.lblTableName.TabIndex = 16
        Me.lblTableName.Text = "..."
        '
        'txtIDColumnName
        '
        Me.txtIDColumnName.Location = New System.Drawing.Point(281, 392)
        Me.txtIDColumnName.Name = "txtIDColumnName"
        Me.txtIDColumnName.Size = New System.Drawing.Size(60, 20)
        Me.txtIDColumnName.TabIndex = 18
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(281, 373)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(90, 13)
        Me.Label11.TabIndex = 17
        Me.Label11.Text = "ID Column Name:"
        '
        'frmGenerateInserts
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(889, 623)
        Me.Controls.Add(Me.txtIDColumnName)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.lblTableName)
        Me.Controls.Add(Me.txtRowTo)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtRowFrom)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Pb1)
        Me.Controls.Add(Me.pb2)
        Me.Controls.Add(Me.cmdDisplayRecords)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmGenerateInserts"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Generate Inserts - Narender Sharma"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.gbDBOptionsSql.ResumeLayout(False)
        Me.gbDBOptionsSql.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblMSG As System.Windows.Forms.Label
    Friend WithEvents cmdGenerateInserts As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LstColumnsDest As System.Windows.Forms.CheckedListBox
    Friend WithEvents OFDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents gbDBOptionsSql As System.Windows.Forms.GroupBox
    Friend WithEvents chkDBFile As System.Windows.Forms.CheckBox
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtUID As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lstColumnsToExclude As System.Windows.Forms.ListBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboTableDest As System.Windows.Forms.ComboBox
    Friend WithEvents cmdReset As System.Windows.Forms.Button
    Friend WithEvents cmdDisplayRecords As System.Windows.Forms.Button
    Friend WithEvents chkSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents chkForAllTables As System.Windows.Forms.CheckBox
    Friend WithEvents pb2 As System.Windows.Forms.ProgressBar
    Friend WithEvents Pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtFrom_Where As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTop As System.Windows.Forms.TextBox
    Friend WithEvents lblTop As System.Windows.Forms.Label
    Friend WithEvents rtbSelectQuery As System.Windows.Forms.RichTextBox
    Friend WithEvents txtRowFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtRowTo As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblTableName As System.Windows.Forms.Label
    Friend WithEvents txtIDColumnName As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
End Class
