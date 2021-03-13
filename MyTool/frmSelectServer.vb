'Imports System.Runtime.InteropServices
'Imports SQLDMO
Imports IntranetSetup1.CustomClass


Imports System.Collections
Imports System.Text

Public Class frmSelectServer
    Public Shared availServer As String

    Private Sub frmSelectServer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Left = (DesktopBounds.Width) / 2
        Me.Top = (DesktopBounds.Height) / 2
        'Try
        '    Dim sqlApp As New SQLDMO.Application()
        '    Dim NL As SQLDMO.NameList
        '    Dim index As Int32 'Use INT32 instead ofLONG.

        '    NL = sqlApp.ListAvailableSQLServers
        '    For index = 1 To NL.Count
        '        lstAvailableServer.Items.Add(NL.Item(index))
        '    Next
        '    sqlApp = Nothing
        '    NL = Nothing
        'Catch ex As Exception
        '    MsgBox("err : " & ex.Message)
        '    MsgBox("err : " & ex.InnerException.ToString)
        'End Try


        Dim dmo As Object
        'Dim dmo As New SQLDMO.Application
        Dim nameList As Object
        Try
            dmo = CreateObject("SQLDMO.Application")
            nameList = dmo.ListAvailableSQLServers()

            'Dim serverName As ArrayList = New ArrayList
            Dim sname As String
            lstAvailableServer.Items.Clear()
            For Each sname In nameList 'serverList
                'serverName.Add(sname)
                lstAvailableServer.Items.Add(sname)
            Next sname
            sname = ""

        Catch e1 As Exception
            MessageBox.Show(e1.Message.ToString)
        End Try
    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click        
        availServer = IIf(Len(lstAvailableServer.SelectedItem.ToString) > 0, lstAvailableServer.SelectedItem.ToString, "tsi_dev_02")
        ' MsgBox(availserver)
        frmPropertCreater.txtServer.Text = availServer
        Me.Hide()        
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        'CancelInstallation.Show()
        Me.Close()
    End Sub

    Private Sub frmSelectServer_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'CancelInstallation.Show()
    End Sub

    Private Function IPAddresses(ByVal server As String) As String
        Dim objArray As New ArrayList
        Dim ip As String = String.Empty
        Try
            Dim curAdd As System.Net.IPAddress
            If (server = "(local)") Then
                Dim sam As System.Net.IPAddress
                Dim sam1 As String
                With System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName())
                    sam = New System.Net.IPAddress(.AddressList(0).Address)
                    sam1 = sam.ToString
                End With
                ip = sam1
                Return ip
            End If

            ' Get server related information.
            Dim heserver As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(server)
            ' Loop on the AddressList
            'Dim curAdd As System.Net.IPAddress
            For Each curAdd In heserver.AddressList
                ip = curAdd.ToString
                Exit For
            Next curAdd
        Catch e As Exception
            Debug.WriteLine("Exception: " + e.ToString())
        End Try
        Return ip
    End Function

    Private Sub lstAvailableServer_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstAvailableServer.MouseDoubleClick
        Call cmdOk_Click(sender, e)
    End Sub
End Class