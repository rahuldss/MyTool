Imports System.Data.SqlClient

Public Class _frmTest

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim conn As New SqlConnection("Server=(local);Data Source=;Integrated Security=SSPI")
        Dim cmd As New SqlCommand("", conn)

        'cmd.CommandText = "CREATE DATABASE ncspl2 ON " & _
        '    "PRIMARY ( FILENAME =  'c:\inetpub\wwwroot\IntranetApplication\DB\ncspl2.mdf' ) " & _
        '      "FILEGROUP MyDatabase_Log ( FILENAME = 'c:\inetpub\wwwroot\IntranetApplication\DB\ncspl2_1.ldf')" & _
        '    "FOR ATTACH"

        cmd.CommandText = "exec sys.sp_attach_db    Intranet,    'c:\inetpub\wwwroot\IntranetApplication\DB\intranet_Data.mdf'"

        conn.Open()

        cmd.ExecuteNonQuery()

        cmd.Dispose()
        conn.Dispose()
    End Sub

    Private Sub cmdWave_Click(sender As System.Object, e As System.EventArgs) Handles cmdWave.Click
        CSharpCodes.Wave.MakeSound()        
        End
    End Sub

    Private Sub cmdBeep_Click(sender As System.Object, e As System.EventArgs) Handles cmdBeep.Click
        CSharpCodes.Wave.Beep()
        End
    End Sub

    Private Sub btn_LCM_Click(sender As System.Object, e As System.EventArgs) Handles btn_LCM.Click
        Dim k As Long = 60

        Dim LArr As Long() = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28}
        'Dim LArr As Long() = {2, 3, 4, 5, 6}
        k = LCM(LArr)
        Dim i As Long = k
        Dim j As Long = 1
        MsgBox(LArr.GetValue(LArr.Length - 1) + 1)
        'Exit Sub

        While (1)
            'If (i Mod 2 = 1) And (i Mod 3 = 1) And (i Mod 4 = 1) And (i Mod 5 = 1) And (i Mod 6 = 1) And (i Mod 7 = 1) And (i Mod 8 = 1) And (i Mod 9 = 1) And (i Mod 10 = 1) And (i Mod 11 = 1) And (i Mod 12 = 1) And (i Mod 13 = 1) And (i Mod 14 = 1) And (i Mod 15 = 1) And (i Mod 16 = 1) And (i Mod 17 = 1) And (i Mod 18 = 1) And (i Mod 19 = 1) And (i Mod 20 = 1) And (i Mod 21 = 1) And (i Mod 22 = 1) And (i Mod 23 = 1) And (i Mod 24 = 1) And (i Mod 25 = 1) And (i Mod 26 = 1) And (i Mod 27 = 1) And (i Mod 28 = 1) And (i Mod 29 = 0) Then
            If ((i + 1) Mod (LArr.GetValue(LArr.Length - 1) + 1) = 0) Then
                MsgBox(i + 1)
                Exit While
            End If

            i = i + k
            j = j + 1
            TextBox1.Text = j
            System.Windows.Forms.Application.DoEvents()
        End While
    End Sub


    Private Shared Function LCM(numbers As Long()) As Long
        Return numbers.Aggregate(AddressOf lcm)
    End Function
    Private Shared Function lcm(a As Long, b As Long) As Long
        Return Math.Abs(a * b) / GCD(a, b)
    End Function
    Private Shared Function GCD(a As Long, b As Long) As Long
        Return If(b = 0, a, GCD(b, a Mod b))
    End Function

End Class