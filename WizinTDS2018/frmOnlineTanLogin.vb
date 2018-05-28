

Public Class frmOnlineTanLogin
    Dim MyOptBtn As HtmlElement
    Dim MyCombo As HtmlElementCollection
    Dim StartingAddress As String
    Dim LoginURL As String
    Dim myObj As Object
    Dim AllElements As HtmlElementCollection
    Dim LogName As String '= "DDFLTAN"
    Dim LogPwd As String '= "DDFL7894"
    Dim LogTAN As String '= "NGPD00368E"

    Private Sub frmOnlineTanLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        StartingAddress = "https://www.tdscpc.gov.in/app/login.xhtml"
        If Len(StartingAddress) > 0 Then
            timTimer.Enabled = True
            brwWebBrowser.Navigate(StartingAddress)
        End If

    End Sub

    Private Sub cboFormNo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormNo.Leave
        cboFormNo.BackColor = Color.White
    End Sub

    Private Sub cboFormNo_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormNo.Enter
        cboFormNo.BackColor = Color.LightYellow
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub brwWebBrowser_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles brwWebBrowser.DocumentCompleted
        Dim HTMLDoc As HtmlDocument = brwWebBrowser.Document
        Dim HTMLElements As HtmlElementCollection = brwWebBrowser.Document.All
        Dim CurElement As HtmlElement
        Dim LoginURL As String = "https://www.tdscpc.gov.in/app/"
        Dim newds As New DataSet
        Dim logname As String
        Dim logpwd As String
        Dim logTan As String
        Dim sql As String = "SELECT TanuserID,TanPassword,coTAN From CoMst WHERE CoMst.CoID = " & selectedcoid
        newds = FetchDataSet(sql)
        logname = newds.Tables(0).Rows(0)(0).ToString()
        logpwd = newds.Tables(0).Rows(0)(1).ToString()
        logTan = newds.Tables(0).Rows(0)(2).ToString()
        newds.Dispose()
        If e.Url.ToString = "https://www.tdscpc.gov.in/app/login.xhtml" Then
            'CODE FOR LOGIN PAGE
            For Each CurElement In HTMLElements
                Select Case CurElement.Name
                    Case "username"
                        CurElement.SetAttribute("value", LogName)
                        CurElement.InvokeMember("onchange")
                        'System.Windows.Forms.SendKeys.Send(Keys.Tab)
                    Case "j_password"
                        'CurElement.InvokeMember("click")
                        'CurElement.RaiseEvent("Onclick")
                        CurElement.SetAttribute("value", LogPwd)
                    Case "j_tanPan"
                        CurElement.SetAttribute("value", LogTAN)
                End Select

            Next

        End If
    End Sub

    Private Sub frmOnlineTanLogin_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        brwWebBrowser.Left = 0
        brwWebBrowser.Top = 0
        brwWebBrowser.Width = Me.Width - 10
        brwWebBrowser.Height = Me.Height - 10
    End Sub
End Class