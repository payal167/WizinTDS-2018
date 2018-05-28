Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Data
Imports System.Text.RegularExpressions
Public Class frmTanReg
    Dim dr As OleDbDataReader
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If txtTanUserId.Text = "" Or txtTanPass.Text = "" Then
            MessageBox.Show("Please Enter Proper Tan Registration Details", "Tan Registration Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        Else
            Dim cmd As New OleDbCommand
            Dim transaction As OleDbTransaction
            Dim sql As String = "UPDATE CoMst SET CoMst.TanUserID =" & Chr(34) & txtTanUserId.Text & Chr(34) & ", CoMst.TANPAssword = " & Chr(34) & txtTanPass.Text & Chr(34) & ", comst.TANRegNo=" & Chr(34) & txtTANreg.Text & Chr(34) & " WHERE COMST.COID = " & selectedcoid
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

    Private Sub frmTanReg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim transaction As OleDbTransaction
        Dim cmd As New OleDbCommand
        Dim sql As String = "SELECT CoMst.TanUserID,CoMst.TANPassword, comst.TANRegNo From CoMst WHERE CoMst.CoID= " & selectedcoid
        'Main()
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
        'dr = cmd.ExecuteReader
        While dr.Read()
            txtTanUserId.Text = dr(0) & ""
            txtTanPass.Text = dr(1) & ""
            txtTANreg.Text = dr(2) & ""
        End While
    End Sub

    Private Sub frmTanReg_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub txtId_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTanUserId.Enter
        txtTanUserId.BackColor = Color.LightYellow
    End Sub

    Private Sub txtId_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTanUserId.Leave
        txtTanUserId.BackColor = Color.White
    End Sub

    Private Sub txtPw_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTanPass.Leave
        txtTanPass.BackColor = Color.White
    End Sub

    Private Sub txtPw_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTanPass.Enter
        txtTanPass.BackColor = Color.LightYellow
    End Sub

    Private Sub txtTanRegNo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTANreg.Leave
        txtTANreg.BackColor = Color.White
    End Sub

    Private Sub txtTanRegNo_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTANreg.Enter
        txtTANreg.BackColor = Color.LightYellow
    End Sub

    Private Sub txtTanUserId_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTanUserId.KeyPress

    End Sub
End Class