Public Class SelectSheet

    Private Sub SelectSheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        WebBrowser1.Navigate(Application.StartupPath & "\Help.htm")
    End Sub
End Class