Imports System.ComponentModel

Public Class frmRegister
    Dim yr1, yr2, yr3 As String
    Public reguser As Boolean
    'Dim lockCtrl As AxJAKLock.AxActiveLock

    Private Sub frmRegister_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub txtCDKey_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCDKey.Enter
        txtCDKey.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCDKey_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCDKey.Leave
        txtCDKey.BackColor = Color.White
    End Sub

    Private Sub txtSiteKey_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSiteKey.Leave
        txtSiteKey.BackColor = Color.White
    End Sub

    Private Sub txtSiteKey_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSiteKey.Enter
        txtSiteKey.BackColor = Color.LightYellow
    End Sub

    Private Sub txtRegKey_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRegKey.Leave
        txtRegKey.BackColor = Color.White
    End Sub

    Private Sub cmdregister_Click(sender As Object, e As EventArgs) Handles cmdregister.Click

        Dim Random As String
        Dim libration As String
        'set the liberation key
        '******************************
        'For i = 1 To 4
        'j = 0
        'If i = 4 Then


        MyLock.RegistryPath = (Application.ProductName) & "\"
        If Len(Trim(txtRegKey.Text)) <> 0 Then
            MyLock.LiberationKey = txtRegKey.Text
        End If


        If Not (MyLock.RegisteredUser) Then
            MsgBox("Invalid Registration key!", vbExclamation, "Wizin-TDS")
        Else

            yr1 = Strings.Left(Application.ProductName, 8) & (Strings.Right(Application.ProductName, 4)) - 1 & "\"
            yr2 = Strings.Left(Application.ProductName, 8) & (Strings.Right(Application.ProductName, 4)) - 2 & "\"
            yr3 = Strings.Left(Application.ProductName, 8) & (Strings.Right(Application.ProductName, 4)) - 3 & "\"
            Call SaveSetting("Wizin-TDS", MyLock.RegistryPath, "CDKey", txtCDKey.Text)
            Call SaveSetting("Wizin-TDS", yr1, "CDKey", txtCDKey.Text)
            Call SaveSetting("Wizin-TDS", yr2, "CDKey", txtCDKey.Text)
            Call SaveSetting("Wizin-TDS", yr3, "CDKey", txtCDKey.Text)
            ' for random key copy paste from current year ro previous year
            Random = GetSetting("Wizin-TDS", MyLock.RegistryPath, "RandomKey")
            SaveSetting("Wizin-TDS", yr1, "RandomKey", Random)
            SaveSetting("Wizin-TDS", yr2, "RandomKey", Random)
            SaveSetting("Wizin-TDS", yr3, "RandomKey", Random)
            'for libration Key
            libration = GetSetting("Wizin-TDS", MyLock.RegistryPath, "LiberationKey")
            SaveSetting("Wizin-TDS", yr1, "LiberationKey", libration)
            SaveSetting("Wizin-TDS", yr2, "LiberationKey", libration)
            SaveSetting("Wizin-TDS", yr3, "LiberationKey", libration)

            MsgBox("Thank you for registering!", vbInformation, "Wizin-TDS")
            Me.Dispose()
            frmCoMst.Show()
        End If
    End Sub

    Private Sub txtCDKey_TextChanged(sender As Object, e As EventArgs) Handles txtCDKey.TextChanged
        If Len(txtCDKey.Text) = txtCDKey.MaxLength Then
            MyLock.SoftwareName = "Wizin-TDS" & txtCDKey.Text
            txtSiteKey.Text = MyLock.SoftwareCode
            txtReadSite.Text = ConvertSiteKey2Words(txtSiteKey.Text)
        Else
            txtSiteKey.Text = vbNullString
        End If
        txtReadCD.Text = ConvertSiteKey2Words(txtCDKey.Text)
    End Sub

    Private Sub txtRegKey_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRegKey.Enter
        txtRegKey.BackColor = Color.LightYellow
    End Sub

    Private Sub txtRegKey_TextChanged(sender As Object, e As EventArgs) Handles txtRegKey.TextChanged
        cmdregister.Enabled = Not (Len(txtRegKey.Text) = 0)
        cmdregister.Enabled = Not (Len(txtCDKey.Text) = 0)
    End Sub

    Private Sub cmdRegisterLater_Click(sender As Object, e As EventArgs) Handles cmdRegisterLater.Click
        frmCoMst.Show()
        'FrmWhatsnew.Show()
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
        frmLogin.Dispose()
    End Sub

    Private Function ConvertSiteKey2Words(Str As String) As String
        Dim i As Long, str2Retn As String
        str2Retn = vbNullString
        For i = 1 To Len(Str)
            str2Retn = str2Retn + GetWord(Mid(Str, i, 1)) & " - "
        Next i
        ConvertSiteKey2Words = str2Retn
    End Function
    Private Function GetWord(Alpha As String) As String
        If Asc(UCase(Alpha)) >= Asc("A") And Asc(UCase(Alpha)) <= Asc("Z") Then
            GetWord = Choose(Asc(UCase(Alpha)) - 64, "AISHWARYA", "BACHHAN", "CHALLAN", "DOCTOR", "ELEPHANT", "FARDEEN", "GOA", "HERO", "INCOME", "JAK", "KARISHMA", "LONDON", "MUMBAI", "NAGPUR", "ORANGE", "PRIETY", "QUESTION", "RAVEENA", "SHAHRUKH", "TUTI-FRUITY", "URMILA", "VANILLA", "WORD", "XMAS", "YAK", "ZEBRA")
        Else
            GetWord = Alpha
        End If
    End Function


    Private Sub frmRegister_Load(sender As Object, e As EventArgs) Handles Me.Load

        MyLock.RegistryPath = Application.ProductName & "\" 'Left(App.ProductName, 8) & Right(App.ProductName, 4) & "\"
        If MyLock.SoftwareName = "Wizin-TDS" Then
            'Assign the name from the registry for the first time..
            MyLock.SoftwareName = SoftName
        End If
        'Dim strHash As String, strKey As String
        'strHash = Hash(MyLock.SoftwareCode & MyLock.SoftwareName) ' 123 'Nitin
        'strKey = modNMB.DecryptText(GetSetting(Strings.Left(MyLock.SoftwareName, 9), MyLock.RegistryPath, "LiberationKey"), "Apr01Apr29Aug23")
        cmdregister.Enabled = False
        'reguser = MyLock.RegisteredUser

    End Sub
    'Private Function Hash(strHashThis As String) As String
    '    ' Allow different hash types

    '    Select Case MyLock.HashAlgorithm
    '        Case "SHA1AA1Hash" : Hash = SHA1AA1Hash(strHashThis)
    '            'Case "SHA1AA2Hash" : Hash = SHA1AA2Hash(strHashThis)
    '            'Case "MD5AA1Hash" : Hash = MD5AA1Hash(strHashThis)
    '            'Case "MD5AA2Hash" : Hash = MD5AA2Hash(strHashThis)
    '            'Case "MD5AB1Hash" : Hash = MD5AB1Hash(strHashThis)
    '            'Case "MD5AB2Hash" : Hash = MD5AB2Hash(strHashThis)
    '        Case Else : Hash = SHA1AA1Hash(strHashThis) ' Default type
    '    End Select

    'End Function

    Private Sub MyLock_Registration(sender As Object, e As AxJAKLock.__ActiveLock_RegistrationEvent)

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub cmdregister_AutoSizeChanged(sender As Object, e As EventArgs) Handles cmdregister.AutoSizeChanged

    End Sub

    Private Sub frmRegister_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Me.Dispose()
        frmLogin.Dispose()
    End Sub

    Private Sub txtRegKey_LostFocus(sender As Object, e As EventArgs) Handles txtRegKey.LostFocus
        cmdRegisterLater.Focus()
    End Sub
End Class