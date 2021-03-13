<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmTest
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdWave = New System.Windows.Forms.Button()
        Me.cmdBeep = New System.Windows.Forms.Button()
        Me.btn_LCM = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(66, 36)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(142, 68)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'cmdWave
        '
        Me.cmdWave.Location = New System.Drawing.Point(66, 116)
        Me.cmdWave.Name = "cmdWave"
        Me.cmdWave.Size = New System.Drawing.Size(142, 29)
        Me.cmdWave.TabIndex = 1
        Me.cmdWave.Text = "Wave"
        Me.cmdWave.UseVisualStyleBackColor = True
        '
        'cmdBeep
        '
        Me.cmdBeep.Location = New System.Drawing.Point(66, 151)
        Me.cmdBeep.Name = "cmdBeep"
        Me.cmdBeep.Size = New System.Drawing.Size(142, 29)
        Me.cmdBeep.TabIndex = 2
        Me.cmdBeep.Text = "Beep"
        Me.cmdBeep.UseVisualStyleBackColor = True
        '
        'btn_LCM
        '
        Me.btn_LCM.Location = New System.Drawing.Point(68, 194)
        Me.btn_LCM.Name = "btn_LCM"
        Me.btn_LCM.Size = New System.Drawing.Size(139, 31)
        Me.btn_LCM.TabIndex = 3
        Me.btn_LCM.Text = "7 Theif - LCM"
        Me.btn_LCM.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(61, 238)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(186, 20)
        Me.TextBox1.TabIndex = 4
        '
        '_frmTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.btn_LCM)
        Me.Controls.Add(Me.cmdBeep)
        Me.Controls.Add(Me.cmdWave)
        Me.Controls.Add(Me.Button1)
        Me.Name = "_frmTest"
        Me.Text = "_frmTest"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdWave As System.Windows.Forms.Button
    Friend WithEvents cmdBeep As System.Windows.Forms.Button
    Friend WithEvents btn_LCM As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
