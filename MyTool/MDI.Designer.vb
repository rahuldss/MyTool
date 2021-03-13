<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MDI
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDI))
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.mnuTools = New System.Windows.Forms.ToolStripMenuItem()
        Me.PropertyCreaterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuGenerateInserts = New System.Windows.Forms.ToolStripMenuItem()
        Me.GetFileNamesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RenameFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConvertVideosInFLVToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportExportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuImportExportSQL = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRnD = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTest = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuShowCrystalReport = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuShowErrorLog = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.MenuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuTools, Me.mnuRnD, Me.mnuShowCrystalReport, Me.mnuShowErrorLog, Me.mnuExit})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(632, 24)
        Me.MenuStrip.TabIndex = 5
        Me.MenuStrip.Text = "MenuStrip"
        '
        'mnuTools
        '
        Me.mnuTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PropertyCreaterToolStripMenuItem, Me.mnuGenerateInserts, Me.GetFileNamesToolStripMenuItem, Me.RenameFilesToolStripMenuItem, Me.ConvertVideosInFLVToolStripMenuItem, Me.ImportExportToolStripMenuItem, Me.mnuImportExportSQL})
        Me.mnuTools.Name = "mnuTools"
        Me.mnuTools.Size = New System.Drawing.Size(48, 20)
        Me.mnuTools.Text = "&Tools"
        '
        'PropertyCreaterToolStripMenuItem
        '
        Me.PropertyCreaterToolStripMenuItem.Name = "PropertyCreaterToolStripMenuItem"
        Me.PropertyCreaterToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.PropertyCreaterToolStripMenuItem.Text = "&Property Creater"
        '
        'mnuGenerateInserts
        '
        Me.mnuGenerateInserts.Name = "mnuGenerateInserts"
        Me.mnuGenerateInserts.Size = New System.Drawing.Size(189, 22)
        Me.mnuGenerateInserts.Text = "Generate &Inserts"
        '
        'GetFileNamesToolStripMenuItem
        '
        Me.GetFileNamesToolStripMenuItem.Name = "GetFileNamesToolStripMenuItem"
        Me.GetFileNamesToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.GetFileNamesToolStripMenuItem.Text = "Get &File Names"
        '
        'RenameFilesToolStripMenuItem
        '
        Me.RenameFilesToolStripMenuItem.Name = "RenameFilesToolStripMenuItem"
        Me.RenameFilesToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.RenameFilesToolStripMenuItem.Text = "&Rename Files"
        '
        'ConvertVideosInFLVToolStripMenuItem
        '
        Me.ConvertVideosInFLVToolStripMenuItem.Name = "ConvertVideosInFLVToolStripMenuItem"
        Me.ConvertVideosInFLVToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.ConvertVideosInFLVToolStripMenuItem.Text = "Convert &Videos in FLV"
        '
        'ImportExportToolStripMenuItem
        '
        Me.ImportExportToolStripMenuItem.Name = "ImportExportToolStripMenuItem"
        Me.ImportExportToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.ImportExportToolStripMenuItem.Text = "I&mport Export"
        '
        'mnuImportExportSQL
        '
        Me.mnuImportExportSQL.Name = "mnuImportExportSQL"
        Me.mnuImportExportSQL.Size = New System.Drawing.Size(189, 22)
        Me.mnuImportExportSQL.Text = "Import Export S&QL"
        '
        'mnuRnD
        '
        Me.mnuRnD.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuTest})
        Me.mnuRnD.Name = "mnuRnD"
        Me.mnuRnD.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.mnuRnD.Size = New System.Drawing.Size(41, 20)
        Me.mnuRnD.Text = "&RnD"
        '
        'mnuTest
        '
        Me.mnuTest.Name = "mnuTest"
        Me.mnuTest.Size = New System.Drawing.Size(96, 22)
        Me.mnuTest.Text = "&Test"
        '
        'mnuShowCrystalReport
        '
        Me.mnuShowCrystalReport.Name = "mnuShowCrystalReport"
        Me.mnuShowCrystalReport.Size = New System.Drawing.Size(125, 20)
        Me.mnuShowCrystalReport.Text = "Show Crystal Report"
        '
        'mnuShowErrorLog
        '
        Me.mnuShowErrorLog.Name = "mnuShowErrorLog"
        Me.mnuShowErrorLog.Size = New System.Drawing.Size(99, 20)
        Me.mnuShowErrorLog.Text = "Show Error &Log"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.mnuExit.Size = New System.Drawing.Size(37, 20)
        Me.mnuExit.Text = "E&xit"
        '
        'MDI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(632, 453)
        Me.Controls.Add(Me.MenuStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "MDI"
        Me.Text = "My Tool - Narender Sharma"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuTools As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PropertyCreaterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuGenerateInserts As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GetFileNamesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RenameFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ConvertVideosInFLVToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRnD As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportExportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuShowErrorLog As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuShowCrystalReport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuImportExportSQL As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuTest As System.Windows.Forms.ToolStripMenuItem

End Class
