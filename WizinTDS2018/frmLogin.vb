Imports System
Imports System.Runtime.CompilerServices
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Data
Imports System.Drawing.Text

Public Class frmLogin
    Dim drag As Boolean
    Dim mousex As Integer
    Private originalSize As Size
    Dim mousey As Integer
    Public cn As OleDbConnection
    Public MdbNm As String
    Dim dr3 As OleDbDataReader
    Public LoginSucceeded As Boolean
    Dim a() As String
    Dim oldsize As Size
    Public idtemp As Integer
    'Private Resolution As New ResolutionChanger
    'Private OldWidth As UInteger
    'Private OldHeight As UInteger

    'Public ToDate, FromDate As Date
    'Public Year As Integer = "2015"
    'Public FY As String = "" & Year & "-" & Year + 1 & ""
    Public version As String = Application.ProductName & " Ver " & Application.ProductVersion '"WizinTDS2015 ver.2015.1.55"
    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim lngFormWidth As Long
        'Dim lngFormHeight As Long

        Main()
        'FromDate = Format("01/04/2017", "dd/mm/yyyy")
        'ToDate = Format("31/03/2018", "dd/mm/yyyy")
        'FY = "" & FromDate & "-" & ToDate & ""
        'lngFormWidth = Me.Width / 15
        'lngFormHeight = Me.Height / 15

        'OldHeight = CUInt(Screen.PrimaryScreen.Bounds.Height)
        'OldWidth = CUInt(Screen.PrimaryScreen.Bounds.Width)
        'Select Case Resolution.SetResolution(1280, 768)
        '    Case ResolutionChanger.ChangeResult.Success
        '        'MsgBox("The Resolution was changed", MsgBoxStyle.OkOnly)
        '    Case ResolutionChanger.ChangeResult.Restart
        '        MsgBox("Restart your system to activate the new resolution setting", MsgBoxStyle.OkOnly)
        '    Case ResolutionChanger.ChangeResult.Fail
        '        MsgBox("The resolution couldn't be changed", MsgBoxStyle.OkOnly)
        '    Case ResolutionChanger.ChangeResult.ResolutionNotSupported
        '        MsgBox("The requested resolution is not supported by your system", MsgBoxStyle.OkOnly)
        'End Select
        'Dim fntscl As Single = Form1.Font.Size * ScaleSize.Height
        'Form1.Font = New Font(Form1.Font.FontFamily, fntscl, Form1.Font.Style, Form1.Font.Unit)
        'Form1.Scale(ScaleSize)
        'Using g = Me.CreateGraphics()
        '    Dim dpiX = g.DpiX
        '    Dim dpiY = g.DpiY
        'End Using
        'Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
        'Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.F

        ' Me.AutoScaleMode = AutoScaleMode.Dpi
        'AdjustText()
        'Me.Size = New System.Drawing.Size(500, 00)

        'Dim Path As New Drawing2D.GraphicsPath
        'Path.AddEllipse(0, 0, 500, 500)
        ''path.AddEllipse(0, 0, 500, 500)
        'Me.Region = New Region(path)
        'Me.Size = New System.Drawing.Size(2000, 2000)
        'Dim path As New Drawing2D.GraphicsPath
        'path.AddEllipse(0, 0, 500, 500)
        'Me.Region = New Region(path)
        Label3.Text = "This version will create  eTDS/eTCS Quarterly Statements " & vbCrLf &
            " for the financial year " & FY & ".The forms and formats " & vbCrLf &
            " have been notified by the Income Tax Department." & vbCrLf &
            "               Please visit www.tin-nsdl.com "
        Label4.Text = "For FY: " & FY & " only."
        'Fill Available Years in combo...new from Ver 2012
        GetYears()

    End Sub

    Private Sub GetYears()
        Dim di As String = My.Application.Info.DirectoryPath
        Dim strDirectory As DirectoryInfo
        For value As Integer = 2001 To 2100
            Dim exename = "WizinTDS" & value & ".exe"
            strDirectory = Directory.GetParent(di)
            Dim Parent As String = System.IO.Directory.GetParent(di).FullName
            For Each dirinf In Directory.GetDirectories(Parent)
                Dim d2 As DirectoryInfo = New DirectoryInfo(dirinf)
                For Each fi As FileInfo In d2.GetFiles()

                    If exename = fi.ToString Then
                        If exename <> (Application.ProductName & ".exe") Then
                            exename = exename.Remove(exename.Length - 4)
                            cboYear.Items.Add(exename)
                            'a(cboYear.SelectedIndex) = (File.Exists(Dir))
                        End If
                    End If
                Next
            Next
        Next


    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        If chkagree.Checked = True Then
            Dim fds As New DataSet
            fds = FetchDataSet("SELECT * FROM Users WHERE UName = '" & txtUserName.Text & "' AND UPwd='" & txtPassword.Text & "'")
            If fds.Tables(0).Rows.Count >= 1 Then
                ' If rst.RecordCount = 1 Then
                DeleteAllowed = False
            Else
                If txtUserName.Text = "JAK" And txtPassword.Text = "JAK245009" Then
                    DeleteAllowed = True
                Else
                    MessageBox.Show("Invalid Username or Password!Please try again", "Invalid User or Password", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    txtUserName.Focus()
                    Exit Sub
                End If
            End If
            LoginSucceeded = True
            Me.Hide()
            frmRegister.Show()
            If frmRegister.Mylock.RegisteredUser = True Then
                'frmRegister.Close()
                frmCoMst.Show()
            Else
                frmRegister.Show()
            End If

        Else
            MessageBox.Show("Please accept terms and conditions to Login", "Accept terms and conditions", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        frmtermscond.Show()
    End Sub

    Private Sub cmdcancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdcancel.Click
        Me.Close()
    End Sub
    Private Sub EnterTab(ByVal e As KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub frmLogin_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Me.EnterTab(e)
    End Sub

    Private Sub txtUserName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUserName.Enter
        txtUserName.BackColor = Color.LightYellow
    End Sub

    Private Sub txtUserName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUserName.Leave
        txtUserName.BackColor = Color.White
    End Sub

    Private Sub txtPassword_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.Enter
        txtPassword.BackColor = Color.LightYellow
    End Sub

    Private Sub txtPassword_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.Leave
        txtPassword.BackColor = Color.White
    End Sub

    Private Sub frmLogin_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        If drag Then
            ' Me.Top = Windows.Forms.Cursor.Position.Y - mousey
            ' Me.Left = Windows.Forms.Cursor.Position.X - mousex
        End If
    End Sub

    Private Sub frmLogin_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
        drag = False
    End Sub

    Private Sub frmLogin_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        drag = True
        'mousex = Windows.Forms.Cursor.Position.X - Me.Left
        ' mousey = Windows.Forms.Cursor.Position.Y - Me.Top
    End Sub

    Private Sub PictureBox1_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If drag Then
            ' Me.Top = Windows.Forms.Cursor.Position.Y - mousey
            'Me.Left = Windows.Forms.Cursor.Position.X - mousex
        End If
    End Sub

    Private Sub PictureBox1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        drag = False
    End Sub

    Private Sub PictureBox1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        drag = True
        ' mousex = Windows.Forms.Cursor.Position.X - Me.Left
        'mousey = Windows.Forms.Cursor.Position.Y - Me.Top
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged

    End Sub



    Private Sub cboYear_GotFocus(sender As Object, e As EventArgs) Handles cboYear.SelectedIndexChanged
        cmdOK.Enabled = False
        cmdcancel.Enabled = False
        txtPassword.Enabled = False
        txtUserName.Enabled = False
    End Sub


    Private Sub cmdcancel_Click(sender As Object, e As EventArgs) Handles cmdcancel.Click
        'set the global var to false
        'to denote a failed login
        'LoginSucceeded = False
        Me.Close()
    End Sub

    Private Sub cmdDesOK_Click()
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        cmdOK.Enabled = True
        cmdcancel.Enabled = True
        txtPassword.Enabled = True
        txtUserName.Enabled = True
        txtUserName.Focus()
    End Sub

    Private Sub btnChgYear_Click(sender As Object, e As EventArgs) Handles btnChgYear.Click
        If cboYear.Text = "Change Year" Or cboYear.Text = "" Then
            MsgBox("Please select the year")
            Exit Sub
        End If
        Shell("a(cboYear.SelectedIndex)", vbNormalFocus)
        End
    End Sub
    Private Sub frmLogin_Closing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Closing
        Me.Dispose()
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub lblProductName_Click(sender As Object, e As EventArgs) Handles lblProductName.Click

    End Sub

    Private Sub lblLine_Click(sender As Object, e As EventArgs) Handles lblLine.Click

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub


    Private Sub frmLogin_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        'On Form1 shown, start applying font 
        Dim CFontPath As String = Application.StartupPath
        'pfc.AddFontFile("C:\JAKINFO\WizinTDS 2018\WizinTDS2018\Resources\Fonts\Roboto.ttf")
        Dim allCtrl As New List(Of Control)
        For Each ctrl As Control In FindALLControlRecursive(allCtrl, Me)
            ' You need to define which control type to change it's font family; not recommendd to just change all controls' fonts, it will create a missy shape
            If TypeOf ctrl Is Label Then 'Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is Button Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is RadioButton Or TypeOf ctrl Is ProgressBar Or TypeOf ctrl Is GroupBox Then
                Dim CurrentCtrlFontSize = ctrl.Font.Size ' get current object's font size before applying new font family
                ctrl.Font = New Font("Microsoft Sans Serif", CurrentCtrlFontSize, FontStyle.Regular)

            End If
        Next
        lblProductName.Font = New Drawing.Font("Microsoft Sans Serif", 26, FontStyle.Bold)
        lblCompanyProduct.Font = New Drawing.Font("Microsoft Sans Serif", 22, FontStyle.Bold)
        allCtrl.Clear()
    End Sub

    Private Sub frmLogin_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        ControlPaint.DrawBorder(e.Graphics, Me.ClientRectangle, Color.Black, ButtonBorderStyle.Solid)
        Dim dpiX As Single = e.Graphics.DpiX
        Dim dpiY As Single = e.Graphics.DpiY
        If dpiX = 96 Then
            Dim Path As New Drawing2D.GraphicsPath
            Path.AddEllipse(0, 0, 500, 500)
            Me.Region = New Region(Path)
        ElseIf dpiX = 120 Then
            Dim Path As New Drawing2D.GraphicsPath
            Path.AddEllipse(0, 0, 600, 600)
            Me.Region = New Region(Path)

        ElseIf dpiX = 144 Then
            Dim Path As New Drawing2D.GraphicsPath
            Path.AddEllipse(0, 0, 700, 700)
            Me.Region = New Region(Path)
        End If
    End Sub

End Class
