<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRenameFiles
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRenameFiles))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtReplaceChars = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtRemoveChars = New System.Windows.Forms.TextBox()
        Me.lblMSG = New System.Windows.Forms.Label()
        Me.cmdRename = New System.Windows.Forms.Button()
        Me.chkRenameAtSameLoc = New System.Windows.Forms.CheckBox()
        Me.gbDest = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDestLoc = New System.Windows.Forms.TextBox()
        Me.cmdSelectDest = New System.Windows.Forms.Button()
        Me.gbSource = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSourceLoc = New System.Windows.Forms.TextBox()
        Me.cmdSelectSource = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.GroupBox1.SuspendLayout()
        Me.gbDest.SuspendLayout()
        Me.gbSource.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtReplaceChars)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtRemoveChars)
        Me.GroupBox1.Controls.Add(Me.lblMSG)
        Me.GroupBox1.Controls.Add(Me.cmdRename)
        Me.GroupBox1.Controls.Add(Me.chkRenameAtSameLoc)
        Me.GroupBox1.Controls.Add(Me.gbDest)
        Me.GroupBox1.Controls.Add(Me.gbSource)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(467, 388)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Rename Files"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 199)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(144, 13)
        Me.Label3.TabIndex = 26
        Me.Label3.Text = "Replace Chars In File Names"
        '
        'txtReplaceChars
        '
        Me.txtReplaceChars.Location = New System.Drawing.Point(171, 196)
        Me.txtReplaceChars.Name = "txtReplaceChars"
        Me.txtReplaceChars.Size = New System.Drawing.Size(286, 20)
        Me.txtReplaceChars.TabIndex = 27
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 175)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(158, 13)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "Remove Chars From File Names"
        '
        'txtRemoveChars
        '
        Me.txtRemoveChars.Location = New System.Drawing.Point(171, 172)
        Me.txtRemoveChars.Name = "txtRemoveChars"
        Me.txtRemoveChars.Size = New System.Drawing.Size(286, 20)
        Me.txtRemoveChars.TabIndex = 25
        '
        'lblMSG
        '
        Me.lblMSG.Location = New System.Drawing.Point(6, 283)
        Me.lblMSG.Name = "lblMSG"
        Me.lblMSG.Size = New System.Drawing.Size(455, 63)
        Me.lblMSG.TabIndex = 5
        Me.lblMSG.Text = "..."
        '
        'cmdRename
        '
        Me.cmdRename.Location = New System.Drawing.Point(185, 349)
        Me.cmdRename.Name = "cmdRename"
        Me.cmdRename.Size = New System.Drawing.Size(97, 33)
        Me.cmdRename.TabIndex = 4
        Me.cmdRename.Text = "Start Renaming"
        Me.cmdRename.UseVisualStyleBackColor = True
        '
        'chkRenameAtSameLoc
        '
        Me.chkRenameAtSameLoc.AutoSize = True
        Me.chkRenameAtSameLoc.Location = New System.Drawing.Point(14, 81)
        Me.chkRenameAtSameLoc.Name = "chkRenameAtSameLoc"
        Me.chkRenameAtSameLoc.Size = New System.Drawing.Size(152, 17)
        Me.chkRenameAtSameLoc.TabIndex = 3
        Me.chkRenameAtSameLoc.Text = "Rename at Same Location"
        Me.chkRenameAtSameLoc.UseVisualStyleBackColor = True
        '
        'gbDest
        '
        Me.gbDest.Controls.Add(Me.Label1)
        Me.gbDest.Controls.Add(Me.txtDestLoc)
        Me.gbDest.Controls.Add(Me.cmdSelectDest)
        Me.gbDest.Location = New System.Drawing.Point(5, 106)
        Me.gbDest.Name = "gbDest"
        Me.gbDest.Size = New System.Drawing.Size(457, 52)
        Me.gbDest.TabIndex = 2
        Me.gbDest.TabStop = False
        Me.gbDest.Text = "Select Destination"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 13)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Destination Location"
        '
        'txtDestLoc
        '
        Me.txtDestLoc.Location = New System.Drawing.Point(116, 13)
        Me.txtDestLoc.Name = "txtDestLoc"
        Me.txtDestLoc.Size = New System.Drawing.Size(305, 20)
        Me.txtDestLoc.TabIndex = 23
        '
        'cmdSelectDest
        '
        Me.cmdSelectDest.Location = New System.Drawing.Point(427, 13)
        Me.cmdSelectDest.Name = "cmdSelectDest"
        Me.cmdSelectDest.Size = New System.Drawing.Size(24, 19)
        Me.cmdSelectDest.TabIndex = 24
        Me.cmdSelectDest.Text = "..."
        Me.cmdSelectDest.UseVisualStyleBackColor = True
        '
        'gbSource
        '
        Me.gbSource.Controls.Add(Me.Label4)
        Me.gbSource.Controls.Add(Me.txtSourceLoc)
        Me.gbSource.Controls.Add(Me.cmdSelectSource)
        Me.gbSource.Location = New System.Drawing.Point(5, 19)
        Me.gbSource.Name = "gbSource"
        Me.gbSource.Size = New System.Drawing.Size(457, 52)
        Me.gbSource.TabIndex = 1
        Me.gbSource.TabStop = False
        Me.gbSource.Text = "Select Source"
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
        'frmRenameFiles
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(491, 412)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmRenameFiles"
        Me.Text = "Rename Files - NDS"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.gbDest.ResumeLayout(False)
        Me.gbDest.PerformLayout()
        Me.gbSource.ResumeLayout(False)
        Me.gbSource.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents gbSource As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSourceLoc As System.Windows.Forms.TextBox
    Friend WithEvents cmdSelectSource As System.Windows.Forms.Button
    Friend WithEvents gbDest As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDestLoc As System.Windows.Forms.TextBox
    Friend WithEvents cmdSelectDest As System.Windows.Forms.Button
    Friend WithEvents chkRenameAtSameLoc As System.Windows.Forms.CheckBox
    Friend WithEvents lblMSG As System.Windows.Forms.Label
    Friend WithEvents cmdRename As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtRemoveChars As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtReplaceChars As System.Windows.Forms.TextBox
End Class
