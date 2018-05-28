Imports System.Text.RegularExpressions
Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Data
Public Class frmAltDetails


    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim cmd As New OleDbCommand
        Dim transaction As OleDbTransaction
        Dim sql As String

        sql = "Update CoMSt Set CoEmailAlt = " & IIf(txtCoEmailAlt.Text = vbNullString, "Null", Chr(34) & txtCoEmailAlt.Text & Chr(34)) & "," _
    & " CoSTDAlt = " & IIf(txtCoSTDAlt.Text = vbNullString, "Null", Chr(34) & txtCoSTDAlt.Text & Chr(34)) & "," _
    & " CoPhoneAlt = " & IIf(txtCoPhoneAlt.Text = vbNullString, "Null", Chr(34) & txtCoPhoneAlt.Text & Chr(34)) & "," _
    & " PR24EmailAlt = " & IIf(txtPR24EmailAlt.Text = vbNullString, "Null", Chr(34) & txtPR24EmailAlt.Text & Chr(34)) & "," _
    & " PR24STDAlt = " & IIf(txtPR24STDAlt.Text = vbNullString, "Null", Chr(34) & txtPR24STDAlt.Text & Chr(34)) & "," _
    & " PR24PhoneAlt = " & IIf(txtPR24PhoneAlt.Text = vbNullString, "Null", Chr(34) & txtPR24PhoneAlt.Text & Chr(34)) & "," _
    & " PR26EmailAlt = " & IIf(txtPR26EmailAlt.Text = vbNullString, "Null", Chr(34) & txtPR26EmailAlt.Text & Chr(34)) & "," _
    & " PR26STDAlt = " & IIf(txtPR26STDAlt.Text = vbNullString, "Null", Chr(34) & txtPR26STDAlt.Text & Chr(34)) & "," _
    & " PR26PhoneAlt = " & IIf(txtPR26PhoneAlt.Text = vbNullString, "Null", Chr(34) & txtPR26PhoneAlt.Text & Chr(34)) & "," _
    & " PR27EmailAlt = " & IIf(txtPR27EmailAlt.Text = vbNullString, "Null", Chr(34) & txtPR27EmailAlt.Text & Chr(34)) & "," _
    & " PR27STDAlt = " & IIf(txtPR27STDAlt.Text = vbNullString, "Null", Chr(34) & txtPR27STDAlt.Text & Chr(34)) & "," _
    & " PR27PhoneAlt = " & IIf(txtPR27PhoneAlt.Text = vbNullString, "Null", Chr(34) & txtPR27PhoneAlt.Text & Chr(34)) & "," _
    & " PR27EEmailAlt = " & IIf(txtPR27EEmailAlt.Text = vbNullString, "Null", Chr(34) & txtPR27EEmailAlt.Text & Chr(34)) & "," _
    & " PR27ESTDAlt = " & IIf(txtPR27ESTDAlt.Text = vbNullString, "Null", Chr(34) & txtPR27ESTDAlt.Text & Chr(34)) & "," _
    & " PR27EPhoneAlt = " & IIf(txtPR27EPhoneAlt.Text = vbNullString, "Null", Chr(34) & txtPR27EPhoneAlt.Text & Chr(34)) & "" _
    & " Where CoId = " & selectedcoid 'frmCoMst.SelectedId

        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction
        cmd.CommandText = sql

        Try

            cmd.ExecuteNonQuery()
            transaction.Commit()
        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show(ex.Message) 'Error MEssage

        End Try
        cmd.Dispose()
        transaction.Dispose()
        ' Me.Close()
    End Sub

    Private Sub frmAltDetails_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Me.EnterTab(e)

    End Sub
    Public Sub EnterTab(ByVal e As KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub frmAltDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load



        Dim nds As New DataSet
        Dim sql As String = "SELECT CoStdAlt,CoPhoneAlt,CoEmailAlt,PR24StdAlt,PR24PhoneAlt,PR24EmailAlt,PR26StdAlt,PR26PhoneAlt,PR26EmailAlt,PR27StdAlt,PR27PhoneAlt,PR27EmailAlt,PR27EStdAlt,PR27EPhoneAlt,PR27EEmailAlt From CoMst WHERE CoMst.CoID = " & selectedcoid 'frmCoMst.SelectedId

        nds = FetchDataSet(sql)
        If nds.Tables(0).Rows.Count > 0 Then

            txtCoSTDAlt.Text = nds.Tables(0).Rows(0)(0) & ""
            txtCoPhoneAlt.Text = nds.Tables(0).Rows(0)(1) & ""
            txtCoEmailAlt.Text = nds.Tables(0).Rows(0)(2) & ""

            txtPR24STDAlt.Text = nds.Tables(0).Rows(0)(3) & ""
            txtPR24PhoneAlt.Text = nds.Tables(0).Rows(0)(4) & ""
            txtPR24EmailAlt.Text = nds.Tables(0).Rows(0)(5) & ""

            txtPR26STDAlt.Text = nds.Tables(0).Rows(0)(6) & ""
            txtPR26PhoneAlt.Text = nds.Tables(0).Rows(0)(7) & ""
            txtPR26EmailAlt.Text = nds.Tables(0).Rows(0)(8) & ""

            txtPR27STDAlt.Text = nds.Tables(0).Rows(0)(9) & ""
            txtPR27PhoneAlt.Text = nds.Tables(0).Rows(0)(10) & ""
            txtPR27EmailAlt.Text = nds.Tables(0).Rows(0)(11) & ""

            txtPR27ESTDAlt.Text = nds.Tables(0).Rows(0)(12) & ""
            txtPR27EPhoneAlt.Text = nds.Tables(0).Rows(0)(13) & ""
            txtPR27EEmailAlt.Text = nds.Tables(0).Rows(0)(14) & ""
        End If
        nds.Dispose()
        txtCoEmailAlt.Focus()
    End Sub

    Private Sub txtCoEmailAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoEmailAlt.Enter

        CtrlGotFocus(txtCoEmailAlt)
    End Sub

    Private Sub txtCoEmailAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoEmailAlt.Leave
        txtCoEmailAlt.BackColor = Color.White
    End Sub

    Private Sub txtCoSTDAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoSTDAlt.Leave
        txtCoSTDAlt.BackColor = Color.White
    End Sub

    Private Sub txtCoSTDAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoSTDAlt.Enter

        CtrlGotFocus(txtCoSTDAlt)
    End Sub

    Private Sub txtCoPhoneAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoPhoneAlt.Leave
        txtCoPhoneAlt.BackColor = Color.White
    End Sub

    Private Sub txtCoPhoneAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoPhoneAlt.Enter
        'txtCoPhoneAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtCoPhoneAlt)
    End Sub

    Private Sub txtPR24EmailAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR24EmailAlt.BackColor = Color.White

    End Sub

    Private Sub txtPR24EmailAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'txtPR24EmailAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR24EmailAlt)
    End Sub

    Private Sub txtPR24STDAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR24STDAlt.BackColor = Color.White
    End Sub

    Private Sub txtPR24STDAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '   txtPR24STDAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR24STDAlt)
    End Sub

    Private Sub txtPR24PhoneAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR24PhoneAlt.BackColor = Color.White

    End Sub

    Private Sub txtPR24PhoneAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'txtPR24PhoneAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR24PhoneAlt)
    End Sub

    Private Sub txtPR26EmailAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'txtPR24PhoneAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR24PhoneAlt)
    End Sub
    Private Sub txtPR26STDAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR26STDAlt.BackColor = Color.White
    End Sub

    Private Sub txtPR26STDAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'txtPR26STDAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR26STDAlt)
    End Sub

    Private Sub txtPR26PhoneAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '  txtPR26PhoneAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR26PhoneAlt)
    End Sub

    Private Sub txtPR26PhoneAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR26PhoneAlt.BackColor = Color.White
    End Sub

    Private Sub txtPR27EmailAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR27EmailAlt.BackColor = Color.White
    End Sub

    Private Sub txtPR27EmailAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'txtPR27EmailAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR27EmailAlt)
    End Sub

    Private Sub txtPR27STDAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR27STDAlt.BackColor = Color.White
    End Sub

    Private Sub txtPR27STDAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'txtPR27STDAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR27STDAlt)
    End Sub

    Private Sub txtPR27PhoneAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' txtPR27PhoneAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR27PhoneAlt)

    End Sub

    Private Sub txtPR27PhoneAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR27PhoneAlt.BackColor = Color.White
    End Sub

    Private Sub txtPR27EEmailAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR27EEmailAlt.BackColor = Color.White
    End Sub

    Private Sub txtPR27EEmailAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '  txtPR27EEmailAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR27EEmailAlt)

    End Sub

    Private Sub txtPR27ESTDAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR27ESTDAlt.BackColor = Color.White
    End Sub

    Private Sub txtPR27ESTDAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'txtPR27ESTDAlt.BackColor = Color.LightYellow
        CtrlGotFocus(txtPR27ESTDAlt)
    End Sub

    Private Sub txtPR27EPhoneAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR27EPhoneAlt.BackColor = Color.White
    End Sub

    Private Sub txtPR27EPhoneAlt_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '   txtPR27EPhoneAlt.BackColor = Color.LightYellow

        CtrlGotFocus(txtPR27EPhoneAlt)
    End Sub

    Private Sub txtPR26EmailAlt_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtPR26EmailAlt.BackColor = Color.White
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Dispose()

    End Sub

    Private Sub txtCoSTDAlt_TextChanged(sender As Object, e As EventArgs) Handles txtCoSTDAlt.TextChanged

    End Sub

    Private Sub txtCoSTDAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCoSTDAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtCoPhoneAlt_TextChanged(sender As Object, e As EventArgs) Handles txtCoPhoneAlt.TextChanged

    End Sub

    Private Sub txtCoPhoneAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCoPhoneAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPR24PhoneAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR24PhoneAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPR24STDAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR24STDAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPR26PhoneAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR26PhoneAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPR26STDAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR26STDAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPR27EPhoneAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR27EPhoneAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPR27ESTDAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR27ESTDAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPR27PhoneAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR27PhoneAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPR27STDAlt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR27STDAlt.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub Tabmain_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Tabmain.SelectedIndexChanged
        Select Case Tabmain.SelectedIndex
            Case 0
                txtCoEmailAlt.Focus()

            Case 1
                txtPR24EmailAlt.Focus()
            Case 2
                txtPR26EmailAlt.Focus()
            Case 3
                txtPR27EmailAlt.Focus()
            Case 4
                txtPR27EEmailAlt.Focus()
        End Select
    End Sub
End Class