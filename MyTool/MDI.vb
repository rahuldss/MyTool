Imports System.Windows.Forms

Public Class MDI

    Private Sub MDI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '_frmTest.Show()
        'FrmCreateList.Show()
        'frmPropertCreater.Show()
        'frmRnD.Show()
        'frmImportExport.Show()
        'frmImportExportSQL.Show()
        'frmGenerateInserts.Show()

        'Me.Icon = System.Drawing.SystemIcons.WinLogo

        'Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click

        End
    End Sub

    Private Sub PropertyCreaterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PropertyCreaterToolStripMenuItem.Click
        frmPropertCreater.Show()
    End Sub

    Private Sub mnuRnD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRnD.Click
        'frmRnD.Show()
    End Sub

    Private Sub ImportExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportExportToolStripMenuItem.Click
        frmImportExport.Show()
    End Sub

    Private Sub RenameFilesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RenameFilesToolStripMenuItem.Click
        frmRenameFiles.Show()
    End Sub

    Private Sub GetFileNamesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GetFileNamesToolStripMenuItem.Click
        FrmCreateList.Show()
    End Sub

    Private Sub mnuShowErrorLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShowErrorLog.Click
        Try
            Process.Start("Notepad.exe", My.Application.Info.DirectoryPath() & "\Err.dat")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub mnuGenerateInserts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGenerateInserts.Click
        frmGenerateInserts.Show()
    End Sub

    Private Sub mnuShowCrystalReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShowCrystalReport.Click
        'Dim objForm As New FrmCrystalReport
        ''objForm.ViewReport("D:\Narender\Projects\VB.NET\2008\MyTool\MyTool\CrystalReport\rptAddressCodeList.rpt", , "@parameter1=IN000001&parameter2=IN000001")
        ''objForm.ViewReport("C:\Program Files\Sage Software\Sage Accpac\OE55A\ENG\Report1.rpt", , )
        'objForm.ViewReport("D:\Narender\Projects\VB.NET\2008\JP_ERP_Sage\JP_ERP_Sage\bin\Debug\Reports\RPTApprovalMemorpt.rpt", , "Memono=25")
        'objForm.show()
    End Sub

    Private Sub mnuImportExportSQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImportExportSQL.Click
        frmImportExportSQL.Show()
    End Sub

    Private Sub mnuTest_Click(sender As System.Object, e As System.EventArgs) Handles mnuTest.Click
        _frmTest.Show()
    End Sub
End Class