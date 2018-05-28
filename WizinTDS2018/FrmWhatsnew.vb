Public Class FrmWhatsnew

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Me.WindowState = FormWindowState.Maximized
        'Me.Visible = True
        Me.BringToFront()
        Me.Location = New Point(50, 90)
        WebBrowser1.Navigate(Application.StartupPath & "\Whatsnew.htm")
        ' Me.Show()
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

    End Sub

    Private Sub cmdok_Click(sender As Object, e As EventArgs) Handles cmdok.Click
        Dim myprd As String
        Dim appcap As String
        appcap = Application.StartupPath
        myprd = Application.ProductName
        If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
            With System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
                myprd = Application.ProductName
                appcap = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name
            End With
        End If
        Dim aapcap1 As String = appcap & myprd
        SaveSetting("Wizin-TDS", Application.ProductName & "\", "StartUp", 0)
        Me.Close()
        frmCoMst.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myprd As String
        Dim appcap As String

        If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
            With System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
                myprd = Application.ProductName
                appcap = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name
            End With
        End If
        Dim aapcap1 As String = appcap & myprd
        SaveSetting("Wizin-TDS", Application.ProductName & "\", "StartUp", 1)
        Me.Close()
        frmCoMst.Show()
    End Sub
End Class