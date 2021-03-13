<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConnStr
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConnStr))
        Me.txtConnStr = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtConnStr
        '
        Me.txtConnStr.Location = New System.Drawing.Point(6, 10)
        Me.txtConnStr.Multiline = True
        Me.txtConnStr.Name = "txtConnStr"
        Me.txtConnStr.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtConnStr.Size = New System.Drawing.Size(644, 495)
        Me.txtConnStr.TabIndex = 0
        Me.txtConnStr.Text = resources.GetString("txtConnStr.Text")
        '
        'frmConnStr
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(662, 517)
        Me.Controls.Add(Me.txtConnStr)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmConnStr"
        Me.Text = "frmConnStr - NDS"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtConnStr As System.Windows.Forms.TextBox
End Class
