﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportExportSQL
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportExportSQL))
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblRecords = New System.Windows.Forms.Label()
        Me.gbDBOptionsSqlSRC = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboTableSource = New System.Windows.Forms.ComboBox()
        Me.cmdOKSRC = New System.Windows.Forms.Button()
        Me.txtDatabaseSRC = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtPasswordSRC = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtUIDSRC = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtServerSRC = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmdImport = New System.Windows.Forms.Button()
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
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.chkSelectAll = New System.Windows.Forms.CheckBox()
        Me.cmdResetMapping = New System.Windows.Forms.Button()
        Me.lstColumnsMapped = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LstColumnsDest = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LstColumnsSource = New System.Windows.Forms.CheckedListBox()
        Me.OFDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.chkIdentityRequired = New System.Windows.Forms.CheckBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtMaxID = New System.Windows.Forms.TextBox()
        Me.GroupBox2.SuspendLayout()
        Me.gbDBOptionsSqlSRC.SuspendLayout()
        Me.gbDBOptionsSql.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtMaxID)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.chkIdentityRequired)
        Me.GroupBox2.Controls.Add(Me.lblRecords)
        Me.GroupBox2.Controls.Add(Me.gbDBOptionsSqlSRC)
        Me.GroupBox2.Controls.Add(Me.cmdImport)
        Me.GroupBox2.Controls.Add(Me.gbDBOptionsSql)
        Me.GroupBox2.Controls.Add(Me.lblMSG)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(489, 592)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        '
        'lblRecords
        '
        Me.lblRecords.Location = New System.Drawing.Point(10, 169)
        Me.lblRecords.Name = "lblRecords"
        Me.lblRecords.Size = New System.Drawing.Size(453, 30)
        Me.lblRecords.TabIndex = 34
        Me.lblRecords.Text = "."
        '
        'gbDBOptionsSqlSRC
        '
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.Label4)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.cboTableSource)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.cmdOKSRC)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.txtDatabaseSRC)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.Label8)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.txtPasswordSRC)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.Label10)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.txtUIDSRC)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.Label11)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.txtServerSRC)
        Me.gbDBOptionsSqlSRC.Controls.Add(Me.Label12)
        Me.gbDBOptionsSqlSRC.Location = New System.Drawing.Point(9, 12)
        Me.gbDBOptionsSqlSRC.Name = "gbDBOptionsSqlSRC"
        Me.gbDBOptionsSqlSRC.Size = New System.Drawing.Size(457, 152)
        Me.gbDBOptionsSqlSRC.TabIndex = 33
        Me.gbDBOptionsSqlSRC.TabStop = False
        Me.gbDBOptionsSqlSRC.Text = "Select Source"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 123)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 13)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Select Table"
        '
        'cboTableSource
        '
        Me.cboTableSource.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTableSource.FormattingEnabled = True
        Me.cboTableSource.Location = New System.Drawing.Point(109, 120)
        Me.cboTableSource.Name = "cboTableSource"
        Me.cboTableSource.Size = New System.Drawing.Size(240, 21)
        Me.cboTableSource.TabIndex = 38
        '
        'cmdOKSRC
        '
        Me.cmdOKSRC.Location = New System.Drawing.Point(399, 116)
        Me.cmdOKSRC.Name = "cmdOKSRC"
        Me.cmdOKSRC.Size = New System.Drawing.Size(52, 26)
        Me.cmdOKSRC.TabIndex = 37
        Me.cmdOKSRC.Text = "&OK"
        Me.cmdOKSRC.UseVisualStyleBackColor = True
        '
        'txtDatabaseSRC
        '
        Me.txtDatabaseSRC.Location = New System.Drawing.Point(109, 82)
        Me.txtDatabaseSRC.Name = "txtDatabaseSRC"
        Me.txtDatabaseSRC.Size = New System.Drawing.Size(271, 20)
        Me.txtDatabaseSRC.TabIndex = 34
        Me.txtDatabaseSRC.Text = "JpData_MyShop"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 85)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(53, 13)
        Me.Label8.TabIndex = 33
        Me.Label8.Text = "Database"
        '
        'txtPasswordSRC
        '
        Me.txtPasswordSRC.Location = New System.Drawing.Point(109, 60)
        Me.txtPasswordSRC.Name = "txtPasswordSRC"
        Me.txtPasswordSRC.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPasswordSRC.Size = New System.Drawing.Size(271, 20)
        Me.txtPasswordSRC.TabIndex = 32
        Me.txtPasswordSRC.Text = "sa123"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 63)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(53, 13)
        Me.Label10.TabIndex = 31
        Me.Label10.Text = "Password"
        '
        'txtUIDSRC
        '
        Me.txtUIDSRC.Location = New System.Drawing.Point(109, 37)
        Me.txtUIDSRC.Name = "txtUIDSRC"
        Me.txtUIDSRC.Size = New System.Drawing.Size(271, 20)
        Me.txtUIDSRC.TabIndex = 30
        Me.txtUIDSRC.Text = "sa"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 40)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(26, 13)
        Me.Label11.TabIndex = 29
        Me.Label11.Text = "UID"
        '
        'txtServerSRC
        '
        Me.txtServerSRC.Location = New System.Drawing.Point(109, 14)
        Me.txtServerSRC.Name = "txtServerSRC"
        Me.txtServerSRC.Size = New System.Drawing.Size(271, 20)
        Me.txtServerSRC.TabIndex = 28
        Me.txtServerSRC.Text = "TSI_DEV_02"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 17)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(38, 13)
        Me.Label12.TabIndex = 27
        Me.Label12.Text = "Server"
        '
        'cmdImport
        '
        Me.cmdImport.Location = New System.Drawing.Point(408, 555)
        Me.cmdImport.Name = "cmdImport"
        Me.cmdImport.Size = New System.Drawing.Size(75, 26)
        Me.cmdImport.TabIndex = 32
        Me.cmdImport.Text = "&Import"
        Me.cmdImport.UseVisualStyleBackColor = True
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
        Me.gbDBOptionsSql.Location = New System.Drawing.Point(9, 203)
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
        Me.lblMSG.Location = New System.Drawing.Point(10, 388)
        Me.lblMSG.Name = "lblMSG"
        Me.lblMSG.Size = New System.Drawing.Size(453, 159)
        Me.lblMSG.TabIndex = 2
        Me.lblMSG.Text = "."
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.chkSelectAll)
        Me.GroupBox3.Controls.Add(Me.cmdResetMapping)
        Me.GroupBox3.Controls.Add(Me.lstColumnsMapped)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.LstColumnsDest)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.LstColumnsSource)
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
        Me.chkSelectAll.Location = New System.Drawing.Point(10, 420)
        Me.chkSelectAll.Name = "chkSelectAll"
        Me.chkSelectAll.Size = New System.Drawing.Size(70, 17)
        Me.chkSelectAll.TabIndex = 33
        Me.chkSelectAll.Text = "Select &All"
        Me.chkSelectAll.UseVisualStyleBackColor = True
        '
        'cmdResetMapping
        '
        Me.cmdResetMapping.Location = New System.Drawing.Point(90, 420)
        Me.cmdResetMapping.Name = "cmdResetMapping"
        Me.cmdResetMapping.Size = New System.Drawing.Size(183, 20)
        Me.cmdResetMapping.TabIndex = 31
        Me.cmdResetMapping.Text = "Reset Mapping"
        Me.cmdResetMapping.UseVisualStyleBackColor = True
        '
        'lstColumnsMapped
        '
        Me.lstColumnsMapped.FormattingEnabled = True
        Me.lstColumnsMapped.Location = New System.Drawing.Point(10, 445)
        Me.lstColumnsMapped.Name = "lstColumnsMapped"
        Me.lstColumnsMapped.Size = New System.Drawing.Size(345, 147)
        Me.lstColumnsMapped.TabIndex = 30
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(188, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(160, 13)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "List of Columns From Destination"
        '
        'LstColumnsDest
        '
        Me.LstColumnsDest.CheckOnClick = True
        Me.LstColumnsDest.FormattingEnabled = True
        Me.LstColumnsDest.Location = New System.Drawing.Point(191, 31)
        Me.LstColumnsDest.Name = "LstColumnsDest"
        Me.LstColumnsDest.Size = New System.Drawing.Size(164, 379)
        Me.LstColumnsDest.TabIndex = 28
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(141, 13)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "List of Columns From Source"
        '
        'LstColumnsSource
        '
        Me.LstColumnsSource.CheckOnClick = True
        Me.LstColumnsSource.FormattingEnabled = True
        Me.LstColumnsSource.Location = New System.Drawing.Point(10, 31)
        Me.LstColumnsSource.Name = "LstColumnsSource"
        Me.LstColumnsSource.Size = New System.Drawing.Size(164, 379)
        Me.LstColumnsSource.TabIndex = 26
        '
        'OFDialog1
        '
        Me.OFDialog1.AddExtension = False
        Me.OFDialog1.DefaultExt = "*.xls"
        Me.OFDialog1.Filter = "Excel Files (*.xls)|*.xls|New Excel Files (*.xlsx)|*.xlsx"
        '
        'chkIdentityRequired
        '
        Me.chkIdentityRequired.AutoSize = True
        Me.chkIdentityRequired.Checked = True
        Me.chkIdentityRequired.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIdentityRequired.Location = New System.Drawing.Point(294, 559)
        Me.chkIdentityRequired.Name = "chkIdentityRequired"
        Me.chkIdentityRequired.Size = New System.Drawing.Size(106, 17)
        Me.chkIdentityRequired.TabIndex = 35
        Me.chkIdentityRequired.Text = "&Identity Required"
        Me.chkIdentityRequired.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(45, 559)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(128, 13)
        Me.Label13.TabIndex = 36
        Me.Label13.Text = "Max ID (in Source Table):"
        '
        'txtMaxID
        '
        Me.txtMaxID.Location = New System.Drawing.Point(175, 559)
        Me.txtMaxID.Name = "txtMaxID"
        Me.txtMaxID.Size = New System.Drawing.Size(104, 20)
        Me.txtMaxID.TabIndex = 37
        Me.txtMaxID.Text = "0"
        '
        'frmImportExportSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(889, 623)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmImportExportSQL"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ImportExportSQL - NDS"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.gbDBOptionsSqlSRC.ResumeLayout(False)
        Me.gbDBOptionsSqlSRC.PerformLayout()
        Me.gbDBOptionsSql.ResumeLayout(False)
        Me.gbDBOptionsSql.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblMSG As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LstColumnsDest As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LstColumnsSource As System.Windows.Forms.CheckedListBox
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
    Friend WithEvents lstColumnsMapped As System.Windows.Forms.ListBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboTableDest As System.Windows.Forms.ComboBox
    Friend WithEvents cmdResetMapping As System.Windows.Forms.Button
    Friend WithEvents cmdImport As System.Windows.Forms.Button
    Friend WithEvents gbDBOptionsSqlSRC As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboTableSource As System.Windows.Forms.ComboBox
    Friend WithEvents cmdOKSRC As System.Windows.Forms.Button
    Friend WithEvents txtDatabaseSRC As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtPasswordSRC As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtUIDSRC As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtServerSRC As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents chkSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents lblRecords As System.Windows.Forms.Label
    Friend WithEvents chkIdentityRequired As CheckBox
    Friend WithEvents txtMaxID As TextBox
    Friend WithEvents Label13 As Label
End Class
