Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Data
Imports System.Text.RegularExpressions
Public Class frmCIT
    Dim dr As OleDbDataReader
    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        'Main()
        Dim cmd As New OleDbCommand
        Dim transaction As OleDbTransaction
        If txtTDSAddress.Text = " " Or txtTDSCity.Text = " " Or txtTDSPin.Text = " " Then
            MessageBox.Show("Please Enter Proper CIT(TDS) Details", "CIT(TDS) Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        ElseIf txtFatherName.Text = " " Then
            MessageBox.Show("Please Enter Father Name", "Enter Father Name", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        Else
            'frmLogin.cn.Open()
            Dim sql As String = "UPDATE CoMst SET CoMst.CITTDSAddtess =" & IIf(txtTDSAddress.Text = vbNullString, "Null", Chr(34) & txtTDSAddress.Text & Chr(34)) & "," _
                & "CoMst.CITTDSCity = " & IIf(txtTDSCity.Text = vbNullString, "Null", Chr(34) & txtTDSCity.Text & Chr(34)) & "," _
              & " CoMst.CITTDSPin = " & IIf(txtTDSPin.Text = vbNullString, "Null", Chr(34) & txtTDSPin.Text & Chr(34)) & "," _
              & " CoMst.FatherName = " & IIf(txtFatherName.Text = vbNullString, "Null", Chr(34) & txtFatherName.Text & Chr(34)) & "" _
              & " WHERE COMST.COID = 1"
            'Main()
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
            Me.Close()
        End If
    End Sub
    Private Sub frmCIT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'frmLogin.cn.Open()
        Dim sql As String = "SELECT CoMst.CITTDSAddtess, CoMst.CITTDSCity, CoMst.CITTDSPin,comst.fathername From CoMst WHERE CoMst.CoID =1"
        Dim cmd As New OleDbCommand
        Dim transaction As OleDbTransaction
        ' Main()
        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction
        cmd.CommandText = sql
        Try
            dr = cmd.ExecuteReader()
            transaction.Commit()
        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show(ex.Message) 'Error MEssage
        End Try
        cmd.Dispose()
        transaction.Dispose()
        While dr.Read()
            txtTDSAddress.Text = dr(0) & ""
            txtTDSCity.Text = dr(1) & ""
            txtTDSPin.Text = dr(2) & ""
            txtFatherName.Text = dr(3) & ""
        End While
        'frmLogin.cn.Close()
    End Sub

    Private Sub frmCIT_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        ' frmCoMst.EnterTab(e)
    End Sub

    Private Sub txtadress1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTDSAddress.Leave
        txtTDSAddress.BackColor = Color.White
    End Sub

    Private Sub txtadress1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTDSAddress.Enter
        txtTDSAddress.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCity_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTDSCity.Enter
        txtTDSCity.BackColor = Color.LightYellow

    End Sub

    Private Sub txtCity_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTDSCity.Leave
        txtTDSCity.BackColor = Color.White

    End Sub

    Private Sub txtPcode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTDSPin.Leave
        txtTDSPin.BackColor = Color.White

    End Sub

    Private Sub txtPcode_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTDSPin.Enter
        txtTDSPin.BackColor = Color.LightYellow

    End Sub

    Private Sub txtFatherName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFatherName.Enter
        txtFatherName.BackColor = Color.LightYellow

    End Sub

    Private Sub txtFatherName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFatherName.Leave
        txtFatherName.BackColor = Color.White

    End Sub

    Private Sub txtFatherName_TextChanged(sender As Object, e As EventArgs) Handles txtFatherName.TextChanged

    End Sub

    Private Sub txtTDSAddress_TextChanged(sender As Object, e As EventArgs) Handles txtTDSAddress.TextChanged

    End Sub
    Public Sub EnterTab(ByVal e As KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub txtTDSAddress_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTDSAddress.KeyPress
        Me.EnterTab(e)
    End Sub

    Private Sub txtTDSCity_TextChanged(sender As Object, e As EventArgs) Handles txtTDSCity.TextChanged

    End Sub

    Private Sub txtTDSCity_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTDSCity.KeyPress
        Me.EnterTab(e)
    End Sub

    Private Sub txtTDSPin_TextChanged(sender As Object, e As EventArgs) Handles txtTDSPin.TextChanged

    End Sub

    Private Sub txtFatherName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtFatherName.KeyPress
        Me.EnterTab(e)
    End Sub

    Private Sub txtTDSPin_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTDSPin.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
        Me.EnterTab(e)
    End Sub
End Class