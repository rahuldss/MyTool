<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCreateList
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmCreateList))
        Me.CmdCreateList = New System.Windows.Forms.Button()
        Me.GBPaths = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.CmdExportFolder = New System.Windows.Forms.Button()
        Me.FBDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblDirName = New System.Windows.Forms.Label()
        Me.lblFileName = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblCPath = New System.Windows.Forms.Label()
        Me.chkListInXLS = New System.Windows.Forms.CheckBox()
        Me.GBPaths.SuspendLayout()
        Me.SuspendLayout()
        '
        'CmdCreateList
        '
        Me.CmdCreateList.Location = New System.Drawing.Point(211, 235)
        Me.CmdCreateList.Name = "CmdCreateList"
        Me.CmdCreateList.Size = New System.Drawing.Size(75, 23)
        Me.CmdCreateList.TabIndex = 0
        Me.CmdCreateList.Text = "Create List"
        Me.CmdCreateList.UseVisualStyleBackColor = True
        '
        'GBPaths
        '
        Me.GBPaths.Controls.Add(Me.Label2)
        Me.GBPaths.Controls.Add(Me.txtOutput)
        Me.GBPaths.Controls.Add(Me.CmdExportFolder)
        Me.GBPaths.Location = New System.Drawing.Point(8, 6)
        Me.GBPaths.Name = "GBPaths"
        Me.GBPaths.Size = New System.Drawing.Size(480, 52)
        Me.GBPaths.TabIndex = 12
        Me.GBPaths.TabStop = False
        Me.GBPaths.Text = "Select Paths..."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Output Folder :"
        '
        'txtOutput
        '
        Me.txtOutput.Location = New System.Drawing.Point(88, 16)
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.Size = New System.Drawing.Size(355, 20)
        Me.txtOutput.TabIndex = 4
        '
        'CmdExportFolder
        '
        Me.CmdExportFolder.Location = New System.Drawing.Point(446, 16)
        Me.CmdExportFolder.Name = "CmdExportFolder"
        Me.CmdExportFolder.Size = New System.Drawing.Size(28, 19)
        Me.CmdExportFolder.TabIndex = 5
        Me.CmdExportFolder.Text = "..."
        Me.CmdExportFolder.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 126)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Current Directory : "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(69, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Current File : "
        '
        'lblDirName
        '
        Me.lblDirName.AutoSize = True
        Me.lblDirName.Location = New System.Drawing.Point(109, 126)
        Me.lblDirName.Name = "lblDirName"
        Me.lblDirName.Size = New System.Drawing.Size(0, 13)
        Me.lblDirName.TabIndex = 15
        '
        'lblFileName
        '
        Me.lblFileName.AutoSize = True
        Me.lblFileName.Location = New System.Drawing.Point(109, 152)
        Me.lblFileName.Name = "lblFileName"
        Me.lblFileName.Size = New System.Drawing.Size(0, 13)
        Me.lblFileName.TabIndex = 16
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(14, 71)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 13)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Current Path : "
        '
        'lblCPath
        '
        Me.lblCPath.Location = New System.Drawing.Point(93, 71)
        Me.lblCPath.Name = "lblCPath"
        Me.lblCPath.Size = New System.Drawing.Size(390, 50)
        Me.lblCPath.TabIndex = 18
        Me.lblCPath.Text = "..."
        '
        'chkListInXLS
        '
        Me.chkListInXLS.AutoSize = True
        Me.chkListInXLS.Location = New System.Drawing.Point(17, 202)
        Me.chkListInXLS.Name = "chkListInXLS"
        Me.chkListInXLS.Size = New System.Drawing.Size(116, 17)
        Me.chkListInXLS.TabIndex = 19
        Me.chkListInXLS.Text = "Create List in Excel"
        Me.chkListInXLS.UseVisualStyleBackColor = True
        '
        'FrmCreateList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(497, 265)
        Me.Controls.Add(Me.chkListInXLS)
        Me.Controls.Add(Me.lblCPath)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblFileName)
        Me.Controls.Add(Me.lblDirName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GBPaths)
        Me.Controls.Add(Me.CmdCreateList)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmCreateList"
        Me.Text = "Create List - NDS"
        Me.GBPaths.ResumeLayout(False)
        Me.GBPaths.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CmdCreateList As System.Windows.Forms.Button
    Friend WithEvents GBPaths As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtOutput As System.Windows.Forms.TextBox
    Friend WithEvents CmdExportFolder As System.Windows.Forms.Button
    Friend WithEvents FBDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblDirName As System.Windows.Forms.Label
    Friend WithEvents lblFileName As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblCPath As System.Windows.Forms.Label
    Friend WithEvents chkListInXLS As System.Windows.Forms.CheckBox

End Class
