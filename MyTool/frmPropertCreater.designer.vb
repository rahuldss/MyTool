<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPropertCreater
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPropertCreater))
        Me.gbDBType = New System.Windows.Forms.GroupBox()
        Me.optOracle = New System.Windows.Forms.RadioButton()
        Me.optMSSql = New System.Windows.Forms.RadioButton()
        Me.optMSAccess = New System.Windows.Forms.RadioButton()
        Me.gbLanguage = New System.Windows.Forms.GroupBox()
        Me.optCSMVCAPI = New System.Windows.Forms.RadioButton()
        Me.optCSMvc = New System.Windows.Forms.RadioButton()
        Me.optVB6 = New System.Windows.Forms.RadioButton()
        Me.optCS = New System.Windows.Forms.RadioButton()
        Me.optVB = New System.Windows.Forms.RadioButton()
        Me.gbDBOptionsAccess = New System.Windows.Forms.GroupBox()
        Me.txtFile = New System.Windows.Forms.TextBox()
        Me.cmdSelectFile = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.gbDBOptionsSql = New System.Windows.Forms.GroupBox()
        Me.cmdSelectServer = New System.Windows.Forms.Button()
        Me.chkDBFile = New System.Windows.Forms.CheckBox()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtUID = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.gbDBOptionsOracle = New System.Windows.Forms.GroupBox()
        Me.chkConnStr = New System.Windows.Forms.CheckBox()
        Me.txtProviderOra = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtPasswordOra = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtUIDOra = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtServerOra = New System.Windows.Forms.TextBox()
        Me.lblServerOra = New System.Windows.Forms.Label()
        Me.cmdCreate = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.OFDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.lblMSG = New System.Windows.Forms.Label()
        Me.lnkConnStr = New System.Windows.Forms.LinkLabel()
        Me.chkColList = New System.Windows.Forms.CheckBox()
        Me.gbDBConnType = New System.Windows.Forms.GroupBox()
        Me.optSql = New System.Windows.Forms.RadioButton()
        Me.optOleDb = New System.Windows.Forms.RadioButton()
        Me.cmdAlterTableCollate = New System.Windows.Forms.Button()
        Me.gbDBType.SuspendLayout()
        Me.gbLanguage.SuspendLayout()
        Me.gbDBOptionsAccess.SuspendLayout()
        Me.gbDBOptionsSql.SuspendLayout()
        Me.gbDBOptionsOracle.SuspendLayout()
        Me.gbDBConnType.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbDBType
        '
        Me.gbDBType.Controls.Add(Me.optOracle)
        Me.gbDBType.Controls.Add(Me.optMSSql)
        Me.gbDBType.Controls.Add(Me.optMSAccess)
        Me.gbDBType.Location = New System.Drawing.Point(12, 9)
        Me.gbDBType.Name = "gbDBType"
        Me.gbDBType.Size = New System.Drawing.Size(179, 77)
        Me.gbDBType.TabIndex = 0
        Me.gbDBType.TabStop = False
        Me.gbDBType.Text = "Select Database Type"
        '
        'optOracle
        '
        Me.optOracle.AutoSize = True
        Me.optOracle.Location = New System.Drawing.Point(19, 54)
        Me.optOracle.Name = "optOracle"
        Me.optOracle.Size = New System.Drawing.Size(68, 17)
        Me.optOracle.TabIndex = 2
        Me.optOracle.TabStop = True
        Me.optOracle.Text = "ORACLE"
        Me.optOracle.UseVisualStyleBackColor = True
        '
        'optMSSql
        '
        Me.optMSSql.AutoSize = True
        Me.optMSSql.Checked = True
        Me.optMSSql.Location = New System.Drawing.Point(19, 35)
        Me.optMSSql.Name = "optMSSql"
        Me.optMSSql.Size = New System.Drawing.Size(62, 17)
        Me.optMSSql.TabIndex = 1
        Me.optMSSql.TabStop = True
        Me.optMSSql.Text = "MSSQL"
        Me.optMSSql.UseVisualStyleBackColor = True
        '
        'optMSAccess
        '
        Me.optMSAccess.AutoSize = True
        Me.optMSAccess.Location = New System.Drawing.Point(19, 17)
        Me.optMSAccess.Name = "optMSAccess"
        Me.optMSAccess.Size = New System.Drawing.Size(83, 17)
        Me.optMSAccess.TabIndex = 0
        Me.optMSAccess.Text = "MSACCESS"
        Me.optMSAccess.UseVisualStyleBackColor = True
        '
        'gbLanguage
        '
        Me.gbLanguage.Controls.Add(Me.optCSMVCAPI)
        Me.gbLanguage.Controls.Add(Me.optCSMvc)
        Me.gbLanguage.Controls.Add(Me.optVB6)
        Me.gbLanguage.Controls.Add(Me.optCS)
        Me.gbLanguage.Controls.Add(Me.optVB)
        Me.gbLanguage.Location = New System.Drawing.Point(197, 9)
        Me.gbLanguage.Name = "gbLanguage"
        Me.gbLanguage.Size = New System.Drawing.Size(182, 77)
        Me.gbLanguage.TabIndex = 2
        Me.gbLanguage.TabStop = False
        Me.gbLanguage.Text = "Select Language"
        '
        'optCSMVCAPI
        '
        Me.optCSMVCAPI.AutoSize = True
        Me.optCSMVCAPI.Checked = True
        Me.optCSMVCAPI.Location = New System.Drawing.Point(78, 36)
        Me.optCSMVCAPI.Name = "optCSMVCAPI"
        Me.optCSMVCAPI.Size = New System.Drawing.Size(85, 17)
        Me.optCSMVCAPI.TabIndex = 5
        Me.optCSMVCAPI.TabStop = True
        Me.optCSMVCAPI.Text = "CS MVC API"
        Me.optCSMVCAPI.UseVisualStyleBackColor = True
        '
        'optCSMvc
        '
        Me.optCSMvc.AutoSize = True
        Me.optCSMvc.Checked = True
        Me.optCSMvc.Location = New System.Drawing.Point(78, 16)
        Me.optCSMvc.Name = "optCSMvc"
        Me.optCSMvc.Size = New System.Drawing.Size(65, 17)
        Me.optCSMvc.TabIndex = 4
        Me.optCSMvc.TabStop = True
        Me.optCSMvc.Text = "CS MVC"
        Me.optCSMvc.UseVisualStyleBackColor = True
        '
        'optVB6
        '
        Me.optVB6.AutoSize = True
        Me.optVB6.Location = New System.Drawing.Point(12, 55)
        Me.optVB6.Name = "optVB6"
        Me.optVB6.Size = New System.Drawing.Size(54, 17)
        Me.optVB6.TabIndex = 3
        Me.optVB6.Text = "VB6.0"
        Me.optVB6.UseVisualStyleBackColor = True
        '
        'optCS
        '
        Me.optCS.AutoSize = True
        Me.optCS.Location = New System.Drawing.Point(12, 36)
        Me.optCS.Name = "optCS"
        Me.optCS.Size = New System.Drawing.Size(39, 17)
        Me.optCS.TabIndex = 1
        Me.optCS.Text = "CS"
        Me.optCS.UseVisualStyleBackColor = True
        '
        'optVB
        '
        Me.optVB.AutoSize = True
        Me.optVB.Location = New System.Drawing.Point(12, 17)
        Me.optVB.Name = "optVB"
        Me.optVB.Size = New System.Drawing.Size(39, 17)
        Me.optVB.TabIndex = 0
        Me.optVB.Text = "VB"
        Me.optVB.UseVisualStyleBackColor = True
        '
        'gbDBOptionsAccess
        '
        Me.gbDBOptionsAccess.Controls.Add(Me.txtFile)
        Me.gbDBOptionsAccess.Controls.Add(Me.cmdSelectFile)
        Me.gbDBOptionsAccess.Controls.Add(Me.Label2)
        Me.gbDBOptionsAccess.Location = New System.Drawing.Point(12, 94)
        Me.gbDBOptionsAccess.Name = "gbDBOptionsAccess"
        Me.gbDBOptionsAccess.Size = New System.Drawing.Size(515, 58)
        Me.gbDBOptionsAccess.TabIndex = 3
        Me.gbDBOptionsAccess.TabStop = False
        Me.gbDBOptionsAccess.Text = "Access DB Options"
        Me.gbDBOptionsAccess.Visible = False
        '
        'txtFile
        '
        Me.txtFile.Location = New System.Drawing.Point(109, 20)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.Size = New System.Drawing.Size(374, 20)
        Me.txtFile.TabIndex = 28
        '
        'cmdSelectFile
        '
        Me.cmdSelectFile.Location = New System.Drawing.Point(485, 20)
        Me.cmdSelectFile.Name = "cmdSelectFile"
        Me.cmdSelectFile.Size = New System.Drawing.Size(24, 21)
        Me.cmdSelectFile.TabIndex = 29
        Me.cmdSelectFile.Text = "..."
        Me.cmdSelectFile.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 13)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "Select DB Location"
        '
        'gbDBOptionsSql
        '
        Me.gbDBOptionsSql.Controls.Add(Me.cmdSelectServer)
        Me.gbDBOptionsSql.Controls.Add(Me.chkDBFile)
        Me.gbDBOptionsSql.Controls.Add(Me.txtDatabase)
        Me.gbDBOptionsSql.Controls.Add(Me.Label5)
        Me.gbDBOptionsSql.Controls.Add(Me.txtPassword)
        Me.gbDBOptionsSql.Controls.Add(Me.Label4)
        Me.gbDBOptionsSql.Controls.Add(Me.txtUID)
        Me.gbDBOptionsSql.Controls.Add(Me.Label3)
        Me.gbDBOptionsSql.Controls.Add(Me.txtServer)
        Me.gbDBOptionsSql.Controls.Add(Me.Label1)
        Me.gbDBOptionsSql.Location = New System.Drawing.Point(12, 158)
        Me.gbDBOptionsSql.Name = "gbDBOptionsSql"
        Me.gbDBOptionsSql.Size = New System.Drawing.Size(515, 113)
        Me.gbDBOptionsSql.TabIndex = 30
        Me.gbDBOptionsSql.TabStop = False
        Me.gbDBOptionsSql.Text = "SQL DB Options"
        Me.gbDBOptionsSql.Visible = False
        '
        'cmdSelectServer
        '
        Me.cmdSelectServer.Location = New System.Drawing.Point(381, 14)
        Me.cmdSelectServer.Name = "cmdSelectServer"
        Me.cmdSelectServer.Size = New System.Drawing.Size(24, 21)
        Me.cmdSelectServer.TabIndex = 37
        Me.cmdSelectServer.Text = "..."
        Me.cmdSelectServer.UseVisualStyleBackColor = True
        '
        'chkDBFile
        '
        Me.chkDBFile.Location = New System.Drawing.Point(392, 56)
        Me.chkDBFile.Name = "chkDBFile"
        Me.chkDBFile.Size = New System.Drawing.Size(117, 46)
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
        Me.txtDatabase.Text = "AB"
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
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 63)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 31
        Me.Label4.Text = "Password"
        '
        'txtUID
        '
        Me.txtUID.Location = New System.Drawing.Point(109, 37)
        Me.txtUID.Name = "txtUID"
        Me.txtUID.Size = New System.Drawing.Size(271, 20)
        Me.txtUID.TabIndex = 30
        Me.txtUID.Text = "sa"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(26, 13)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "UID"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(109, 14)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(271, 20)
        Me.txtServer.TabIndex = 28
        Me.txtServer.Text = "."
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Server"
        '
        'gbDBOptionsOracle
        '
        Me.gbDBOptionsOracle.Controls.Add(Me.chkConnStr)
        Me.gbDBOptionsOracle.Controls.Add(Me.txtProviderOra)
        Me.gbDBOptionsOracle.Controls.Add(Me.Label6)
        Me.gbDBOptionsOracle.Controls.Add(Me.txtPasswordOra)
        Me.gbDBOptionsOracle.Controls.Add(Me.Label7)
        Me.gbDBOptionsOracle.Controls.Add(Me.txtUIDOra)
        Me.gbDBOptionsOracle.Controls.Add(Me.Label8)
        Me.gbDBOptionsOracle.Controls.Add(Me.txtServerOra)
        Me.gbDBOptionsOracle.Controls.Add(Me.lblServerOra)
        Me.gbDBOptionsOracle.Location = New System.Drawing.Point(12, 121)
        Me.gbDBOptionsOracle.Name = "gbDBOptionsOracle"
        Me.gbDBOptionsOracle.Size = New System.Drawing.Size(515, 113)
        Me.gbDBOptionsOracle.TabIndex = 35
        Me.gbDBOptionsOracle.TabStop = False
        Me.gbDBOptionsOracle.Text = "Oracle DB Options"
        Me.gbDBOptionsOracle.Visible = False
        '
        'chkConnStr
        '
        Me.chkConnStr.Location = New System.Drawing.Point(392, 16)
        Me.chkConnStr.Name = "chkConnStr"
        Me.chkConnStr.Size = New System.Drawing.Size(117, 78)
        Me.chkConnStr.TabIndex = 35
        Me.chkConnStr.Text = "Use Connection String From Server Option"
        Me.chkConnStr.UseVisualStyleBackColor = True
        '
        'txtProviderOra
        '
        Me.txtProviderOra.Location = New System.Drawing.Point(109, 82)
        Me.txtProviderOra.Name = "txtProviderOra"
        Me.txtProviderOra.Size = New System.Drawing.Size(271, 20)
        Me.txtProviderOra.TabIndex = 34
        Me.txtProviderOra.Text = "msdaora"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 85)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(46, 13)
        Me.Label6.TabIndex = 33
        Me.Label6.Text = "Provider"
        '
        'txtPasswordOra
        '
        Me.txtPasswordOra.Location = New System.Drawing.Point(109, 60)
        Me.txtPasswordOra.Name = "txtPasswordOra"
        Me.txtPasswordOra.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPasswordOra.Size = New System.Drawing.Size(271, 20)
        Me.txtPasswordOra.TabIndex = 32
        Me.txtPasswordOra.Text = "tiger"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 63)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(53, 13)
        Me.Label7.TabIndex = 31
        Me.Label7.Text = "Password"
        '
        'txtUIDOra
        '
        Me.txtUIDOra.Location = New System.Drawing.Point(109, 37)
        Me.txtUIDOra.Name = "txtUIDOra"
        Me.txtUIDOra.Size = New System.Drawing.Size(271, 20)
        Me.txtUIDOra.TabIndex = 30
        Me.txtUIDOra.Text = "scott"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 40)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(26, 13)
        Me.Label8.TabIndex = 29
        Me.Label8.Text = "UID"
        '
        'txtServerOra
        '
        Me.txtServerOra.Location = New System.Drawing.Point(109, 14)
        Me.txtServerOra.Name = "txtServerOra"
        Me.txtServerOra.Size = New System.Drawing.Size(271, 20)
        Me.txtServerOra.TabIndex = 28
        Me.txtServerOra.Text = "ORA"
        '
        'lblServerOra
        '
        Me.lblServerOra.AutoSize = True
        Me.lblServerOra.Location = New System.Drawing.Point(6, 17)
        Me.lblServerOra.Name = "lblServerOra"
        Me.lblServerOra.Size = New System.Drawing.Size(38, 13)
        Me.lblServerOra.TabIndex = 27
        Me.lblServerOra.Text = "Server"
        '
        'cmdCreate
        '
        Me.cmdCreate.Location = New System.Drawing.Point(172, 281)
        Me.cmdCreate.Name = "cmdCreate"
        Me.cmdCreate.Size = New System.Drawing.Size(86, 30)
        Me.cmdCreate.TabIndex = 31
        Me.cmdCreate.Text = "&Create"
        Me.cmdCreate.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(276, 281)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(86, 30)
        Me.cmdCancel.TabIndex = 32
        Me.cmdCancel.Text = "Cance&l"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'OFDialog1
        '
        Me.OFDialog1.AddExtension = False
        Me.OFDialog1.DefaultExt = "*.xls"
        Me.OFDialog1.Filter = "All Files(*.*)|*.*"
        '
        'lblMSG
        '
        Me.lblMSG.Location = New System.Drawing.Point(12, 322)
        Me.lblMSG.Name = "lblMSG"
        Me.lblMSG.Size = New System.Drawing.Size(515, 49)
        Me.lblMSG.TabIndex = 33
        Me.lblMSG.Text = "..."
        '
        'lnkConnStr
        '
        Me.lnkConnStr.AutoSize = True
        Me.lnkConnStr.Location = New System.Drawing.Point(395, 278)
        Me.lnkConnStr.Name = "lnkConnStr"
        Me.lnkConnStr.Size = New System.Drawing.Size(122, 13)
        Me.lnkConnStr.TabIndex = 36
        Me.lnkConnStr.TabStop = True
        Me.lnkConnStr.Text = "View Connection Strings"
        '
        'chkColList
        '
        Me.chkColList.AutoSize = True
        Me.chkColList.Location = New System.Drawing.Point(12, 281)
        Me.chkColList.Name = "chkColList"
        Me.chkColList.Size = New System.Drawing.Size(119, 17)
        Me.chkColList.TabIndex = 37
        Me.chkColList.Text = "Create Columns List"
        Me.chkColList.UseVisualStyleBackColor = True
        '
        'gbDBConnType
        '
        Me.gbDBConnType.Controls.Add(Me.optSql)
        Me.gbDBConnType.Controls.Add(Me.optOleDb)
        Me.gbDBConnType.Location = New System.Drawing.Point(384, 9)
        Me.gbDBConnType.Name = "gbDBConnType"
        Me.gbDBConnType.Size = New System.Drawing.Size(143, 77)
        Me.gbDBConnType.TabIndex = 38
        Me.gbDBConnType.TabStop = False
        Me.gbDBConnType.Text = "DB Conn Type"
        '
        'optSql
        '
        Me.optSql.AutoSize = True
        Me.optSql.Checked = True
        Me.optSql.Location = New System.Drawing.Point(9, 38)
        Me.optSql.Name = "optSql"
        Me.optSql.Size = New System.Drawing.Size(96, 17)
        Me.optSql.TabIndex = 1
        Me.optSql.TabStop = True
        Me.optSql.Text = "Using SqlClient"
        Me.optSql.UseVisualStyleBackColor = True
        '
        'optOleDb
        '
        Me.optOleDb.AutoSize = True
        Me.optOleDb.Location = New System.Drawing.Point(9, 16)
        Me.optOleDb.Name = "optOleDb"
        Me.optOleDb.Size = New System.Drawing.Size(85, 17)
        Me.optOleDb.TabIndex = 0
        Me.optOleDb.Text = "Using OleDb"
        Me.optOleDb.UseVisualStyleBackColor = True
        '
        'cmdAlterTableCollate
        '
        Me.cmdAlterTableCollate.Location = New System.Drawing.Point(398, 294)
        Me.cmdAlterTableCollate.Name = "cmdAlterTableCollate"
        Me.cmdAlterTableCollate.Size = New System.Drawing.Size(113, 20)
        Me.cmdAlterTableCollate.TabIndex = 39
        Me.cmdAlterTableCollate.Text = "Alter Table Collate"
        Me.cmdAlterTableCollate.UseVisualStyleBackColor = True
        '
        'frmPropertCreater
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(535, 383)
        Me.Controls.Add(Me.cmdAlterTableCollate)
        Me.Controls.Add(Me.gbDBConnType)
        Me.Controls.Add(Me.chkColList)
        Me.Controls.Add(Me.lnkConnStr)
        Me.Controls.Add(Me.lblMSG)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdCreate)
        Me.Controls.Add(Me.gbDBOptionsSql)
        Me.Controls.Add(Me.gbDBOptionsAccess)
        Me.Controls.Add(Me.gbLanguage)
        Me.Controls.Add(Me.gbDBType)
        Me.Controls.Add(Me.gbDBOptionsOracle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPropertCreater"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create Property Class - Narender Sharma"
        Me.gbDBType.ResumeLayout(False)
        Me.gbDBType.PerformLayout()
        Me.gbLanguage.ResumeLayout(False)
        Me.gbLanguage.PerformLayout()
        Me.gbDBOptionsAccess.ResumeLayout(False)
        Me.gbDBOptionsAccess.PerformLayout()
        Me.gbDBOptionsSql.ResumeLayout(False)
        Me.gbDBOptionsSql.PerformLayout()
        Me.gbDBOptionsOracle.ResumeLayout(False)
        Me.gbDBOptionsOracle.PerformLayout()
        Me.gbDBConnType.ResumeLayout(False)
        Me.gbDBConnType.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbDBType As System.Windows.Forms.GroupBox
    Friend WithEvents optMSSql As System.Windows.Forms.RadioButton
    Friend WithEvents optMSAccess As System.Windows.Forms.RadioButton
    Friend WithEvents gbLanguage As System.Windows.Forms.GroupBox
    Friend WithEvents optCS As System.Windows.Forms.RadioButton
    Friend WithEvents optVB As System.Windows.Forms.RadioButton
    Friend WithEvents gbDBOptionsAccess As System.Windows.Forms.GroupBox
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents cmdSelectFile As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents gbDBOptionsSql As System.Windows.Forms.GroupBox
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtUID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdCreate As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents OFDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblMSG As System.Windows.Forms.Label
    Friend WithEvents optOracle As System.Windows.Forms.RadioButton
    Friend WithEvents gbDBOptionsOracle As System.Windows.Forms.GroupBox
    Friend WithEvents txtProviderOra As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPasswordOra As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtUIDOra As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtServerOra As System.Windows.Forms.TextBox
    Friend WithEvents lblServerOra As System.Windows.Forms.Label
    Friend WithEvents chkConnStr As System.Windows.Forms.CheckBox
    Friend WithEvents chkDBFile As System.Windows.Forms.CheckBox
    Friend WithEvents optVB6 As System.Windows.Forms.RadioButton
    Friend WithEvents lnkConnStr As System.Windows.Forms.LinkLabel
    Friend WithEvents chkColList As System.Windows.Forms.CheckBox
    Friend WithEvents gbDBConnType As System.Windows.Forms.GroupBox
    Friend WithEvents optSql As System.Windows.Forms.RadioButton
    Friend WithEvents optOleDb As System.Windows.Forms.RadioButton
    Friend WithEvents cmdSelectServer As System.Windows.Forms.Button
    Friend WithEvents cmdAlterTableCollate As System.Windows.Forms.Button
    Friend WithEvents optCSMvc As System.Windows.Forms.RadioButton
    Friend WithEvents optCSMVCAPI As System.Windows.Forms.RadioButton

End Class
