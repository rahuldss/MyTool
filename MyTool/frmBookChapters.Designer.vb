<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBookChapters
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.gbDBOptionsSql = New System.Windows.Forms.GroupBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtUID = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblMSG = New System.Windows.Forms.Label()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.gbSource = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSourceLoc = New System.Windows.Forms.TextBox()
        Me.cmdSelectSource = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.GroupBox1.SuspendLayout()
        Me.gbDBOptionsSql.SuspendLayout()
        Me.gbSource.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.gbDBOptionsSql)
        Me.GroupBox1.Controls.Add(Me.lblMSG)
        Me.GroupBox1.Controls.Add(Me.cmdStart)
        Me.GroupBox1.Controls.Add(Me.gbSource)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(506, 456)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Manage Book Chapter PDF Files"
        '
        'gbDBOptionsSql
        '
        Me.gbDBOptionsSql.Controls.Add(Me.cmdOK)
        Me.gbDBOptionsSql.Controls.Add(Me.txtDatabase)
        Me.gbDBOptionsSql.Controls.Add(Me.Label5)
        Me.gbDBOptionsSql.Controls.Add(Me.txtPassword)
        Me.gbDBOptionsSql.Controls.Add(Me.Label3)
        Me.gbDBOptionsSql.Controls.Add(Me.txtUID)
        Me.gbDBOptionsSql.Controls.Add(Me.Label6)
        Me.gbDBOptionsSql.Controls.Add(Me.txtServer)
        Me.gbDBOptionsSql.Controls.Add(Me.Label7)
        Me.gbDBOptionsSql.Location = New System.Drawing.Point(21, 83)
        Me.gbDBOptionsSql.Name = "gbDBOptionsSql"
        Me.gbDBOptionsSql.Size = New System.Drawing.Size(457, 110)
        Me.gbDBOptionsSql.TabIndex = 32
        Me.gbDBOptionsSql.TabStop = False
        Me.gbDBOptionsSql.Text = "Select Destination"
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(399, 78)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(52, 26)
        Me.cmdOK.TabIndex = 37
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = True
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
        Me.lblMSG.Location = New System.Drawing.Point(10, 200)
        Me.lblMSG.Name = "lblMSG"
        Me.lblMSG.Size = New System.Drawing.Size(489, 201)
        Me.lblMSG.TabIndex = 5
        Me.lblMSG.Text = "..."
        '
        'cmdStart
        '
        Me.cmdStart.Location = New System.Drawing.Point(203, 410)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(97, 33)
        Me.cmdStart.TabIndex = 4
        Me.cmdStart.Text = "Start"
        Me.cmdStart.UseVisualStyleBackColor = True
        '
        'gbSource
        '
        Me.gbSource.Controls.Add(Me.Label4)
        Me.gbSource.Controls.Add(Me.txtSourceLoc)
        Me.gbSource.Controls.Add(Me.cmdSelectSource)
        Me.gbSource.Location = New System.Drawing.Point(21, 19)
        Me.gbSource.Name = "gbSource"
        Me.gbSource.Size = New System.Drawing.Size(457, 52)
        Me.gbSource.TabIndex = 1
        Me.gbSource.TabStop = False
        Me.gbSource.Text = "Select PDF Files Source"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(85, 13)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "Source Location"
        '
        'txtSourceLoc
        '
        Me.txtSourceLoc.Location = New System.Drawing.Point(97, 13)
        Me.txtSourceLoc.Name = "txtSourceLoc"
        Me.txtSourceLoc.Size = New System.Drawing.Size(324, 20)
        Me.txtSourceLoc.TabIndex = 23
        '
        'cmdSelectSource
        '
        Me.cmdSelectSource.Location = New System.Drawing.Point(427, 13)
        Me.cmdSelectSource.Name = "cmdSelectSource"
        Me.cmdSelectSource.Size = New System.Drawing.Size(24, 19)
        Me.cmdSelectSource.TabIndex = 24
        Me.cmdSelectSource.Text = "..."
        Me.cmdSelectSource.UseVisualStyleBackColor = True
        '
        'frmBookChapters
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(515, 467)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmBookChapters"
        Me.Text = "frmBookChapters - NDS"
        Me.GroupBox1.ResumeLayout(False)
        Me.gbDBOptionsSql.ResumeLayout(False)
        Me.gbDBOptionsSql.PerformLayout()
        Me.gbSource.ResumeLayout(False)
        Me.gbSource.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblMSG As System.Windows.Forms.Label
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents gbSource As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSourceLoc As System.Windows.Forms.TextBox
    Friend WithEvents cmdSelectSource As System.Windows.Forms.Button
    Friend WithEvents gbDBOptionsSql As System.Windows.Forms.GroupBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtUID As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
End Class
