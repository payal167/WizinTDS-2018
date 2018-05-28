Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Data
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class frmBankMst
    Dim cmd As New OleDbCommand
    Dim WithEvents oBank As New ClsBankMstObj

    Private Sub frmBankMst_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        oBank = New ClsBankMstObj
        oBank.Fill_cmbState(Me)
        oBank.Fill_CmbBrnchCode(Me)
    End Sub

    Private Sub frmBankMst_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Me.EnterTab(e)
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click


        cbobanbrcode.Text = SetFormat("0000000", cbobanbrcode.Text) 'String.Format("{0:0000000}", cbobanbrcode.Text)

        If cmdSave.Text = "&Add" Then
            'Add item..

            If oBank.Insert(oBank) = False Then
                'MsgBox "Unable to Insert Bank Detail in Database" & vbCrLf & "Call JAK Infosolutions", vbCritical, "CANNOT ADD NOW"
            Else
                Call ClearBankCtrls()
                cbobanbrcode.Focus()
                'Call NormalMode(Me)
                cmdSave.Text = "&Add"
            End If
        Else

            'Edit Item..
            If oBank.Update(oBank) = False Then
                MessageBox.Show("Unable to update Bank details in database" & vbCrLf & "Call JAK Infosolutions", "CANNOT UPDATE NOW")
            Else
                Call ClearBankCtrls()

                'Call NormalMode(Me)
                cmdSave.Text = "&Add"
                cbobanbrcode.Focus()
            End If
        End If

        oBank.Fill_CmbBrnchCode(Me)
    End Sub


    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub frmBankMst_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        cbobanbrcode.Focus()
    End Sub

    Private Sub txtBankName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBankName.Enter
        txtBankName.BackColor = Color.LightYellow
    End Sub

    Private Sub txtBankName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBankName.Leave
        txtBankName.BackColor = Color.White
    End Sub

    Private Sub txtBranch_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranch.Enter
        txtBranch.BackColor = Color.LightYellow
    End Sub

    Private Sub txtBranch_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranch.Leave
        txtBranch.BackColor = Color.White
    End Sub

    Private Sub txtCity_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCity.Leave
        txtCity.BackColor = Color.White

    End Sub

    Private Sub txtCity_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCity.Enter
        txtCity.BackColor = Color.LightYellow
    End Sub

    Private Sub cbobanbrcode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbobanbrcode.Leave
        cbobanbrcode.BackColor = Color.White
        If cbobanbrcode.Text <> "" And cmdSave.Text = "&Add" Then
            cbobanbrcode.Text = SetFormat("0000000", cbobanbrcode.Text)
        End If
    End Sub

    Private Sub cbobanbrcode_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbobanbrcode.Enter
        cbobanbrcode.BackColor = Color.LightYellow
    End Sub

    Private Sub cboDState_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDState.Leave
        cboDState.BackColor = Color.White
    End Sub

    Private Sub cboDState_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDState.Enter
        cboDState.BackColor = Color.LightYellow
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cbobanbrcode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbobanbrcode.SelectedIndexChanged

        If cbobanbrcode.SelectedIndex >= 0 And cbobanbrcode.Text <> "System.Data.DataRowView" Then
            oBank = oBank.Fetch(cbobanbrcode.Text)
            txtBankName.Text = oBank.BankName & ""
            txtBranch.Text = oBank.Branch
            txtCity.Text = oBank.City
            cboDState.Text = oBank.State
            cmdSave.Text = "&Save"
        Else
            ClearBankCtrls()
        End If
    End Sub

    Private Sub ClearBankCtrls()
        cbobanbrcode.Text = ""
        txtBankName.Text = ""
        txtBranch.Text = ""
        txtCity.Text = ""
        cboDState.Text = ""
        cmdSave.Text = "&Add"
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        If cbobanbrcode.Text = vbNullString Then
            Exit Sub
        End If
        If cbobanbrcode.SelectedIndex = -1 Then
            Exit Sub
        End If


        Select Case MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion + vbDefaultButton1, "Delete")
            Case vbYes
                If cbobanbrcode.Text <> vbNullString Then

                    If oBank.Delete(cbobanbrcode.SelectedValue) = False Then
                        'Exit Sub
                    End If
                End If
            Case vbNo
                Exit Sub
        End Select
        Call NormalMode(Me)
        ClearBankCtrls()
        oBank.Fill_CmbBrnchCode(Me)
        cbobanbrcode.Focus()
    End Sub

    Private Sub oBank_BeforeSave(ByRef Cancel As Boolean) Handles oBank.BeforeSave
        If Trim(cbobanbrcode.Text) = vbNullString Then
            Call MsgBox("Bank Branch Code Cannot Be Blank!", vbInformation, "Caution")
            Cancel = True
            cbobanbrcode.Focus()
            Exit Sub
        End If
        If Len(cbobanbrcode.Text) > 7 Then
            Call MsgBox("Bank Branch Code Cannot greater than 7 digit!", vbInformation, "Caution")
            Cancel = True
            cbobanbrcode.Focus()
            Exit Sub
        End If
        If Trim(txtBankName.Text) = vbNullString Then
            Call MsgBox("Bank Name Cannot Be Blank!", vbInformation, "Caution")
            Cancel = True
            txtBankName.Focus()
            Exit Sub
        End If


        Cancel = False


    End Sub

    Private Sub oBank_BeforeDelete(ByRef Cancel As Boolean) Handles oBank.BeforeDelete
        Dim nds As New DataSet
        Dim sql As String
        sql = " Select BankBrCode FROM Challan24Q where BankBrCode=" & cbobanbrcode.Text _
                           & " union all Select BankBrCode FROM Challan26Q where BankBrCode=" & cbobanbrcode.Text
        '& " union all Select BankBrCode FROM Challan27EQ where BankBrCode=" & cbobanbrcode.Text _
        '& " union all Select BankBrCode FROM Challan27Q where BankBrCode=" & cbobanbrcode.Text
        nds = FetchDataSet(sql)
        If nds.Tables(0).Rows.Count > 0 Then
            Cancel = True
            MessageBox.Show("This Record is further used, Cannot Delete")

        Else
            Cancel = False
        End If
    End Sub

    Private Sub oBank_PrepareDataForSave(ByRef Cancel As Boolean) Handles oBank.PrepareDataForSave
        With oBank

            .BankBrCode = cbobanbrcode.Text
            .BankName = txtBankName.Text
            .Branch = txtBranch.Text
            .Address = vbNullString
            .City = txtCity.Text
            .State = cboDState.Text
            .coid = selectedcoid
            .BType = vbNullString
            .Region = vbNullString
        End With
    End Sub

    Private Sub oBank_AfterSave() Handles oBank.AfterSave

    End Sub

    Private Sub cbobanbrcode_DragDrop(sender As Object, e As DragEventArgs) Handles cbobanbrcode.DragDrop

    End Sub

    Private Sub cbobanbrcode_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cbobanbrcode.KeyPress

    End Sub

    Public Sub EnterTab(ByVal e As KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtCity_TextChanged(sender As Object, e As EventArgs) Handles txtCity.TextChanged

    End Sub

    Private Sub cbobanbrcode_KeyDown(sender As Object, e As KeyEventArgs) Handles cbobanbrcode.KeyDown

    End Sub

    Private Sub cbobanbrcode_Validating(sender As Object, e As CancelEventArgs) Handles cbobanbrcode.Validating

    End Sub

    Private Sub cbobanbrcode_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbobanbrcode.SelectedValueChanged

    End Sub

    Private Sub cboDState_LostFocus(sender As Object, e As EventArgs) Handles cboDState.LostFocus
        cmdSave.Focus()
    End Sub
End Class