Imports System.Data.OleDb

Public Class frmGovDetails
    Private Sub cmdBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBack.Click
        Me.Visible = False
    End Sub

    Private Sub frmGovDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Main()
        Fill_Ministry()
        Fill_cmbState()
    End Sub
    Private Sub Fill_Ministry()

        Dim nds As New DataSet

        Dim QueSt As String = "Select MinistryCode,MinistryName from MinistryMst"
        nds = FetchDataSet(QueSt)

        If nds.Tables(0).Rows.Count > 0 Then
            cboMinistry.DataSource = nds.Tables(0)
            cboMinistry.ValueMember = "MinistryCode"
            cboMinistry.DisplayMember = "MinistryName"
            cboMinistry.SelectedIndex = -1
        End If
        nds.Dispose()

    End Sub
    Private Sub Fill_cmbState()

        Dim nds As New DataSet

        Dim QueSt As String = "Select StateID,StateName from StateMst"
        nds = FetchDataSet(QueSt)
        If nds.Tables(0).Rows.Count > 0 Then
            cboGovtState.DataSource = nds.Tables(0)
            cboGovtState.ValueMember = "StateID"
            cboGovtState.DisplayMember = "StateName"
            cboGovtState.SelectedIndex = -1
        End If
        nds.Dispose()
    End Sub

    Private Sub frmGovDetails_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub cboGovtState_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGovtState.Enter
        cboGovtState.BackColor = Color.LightYellow
    End Sub

    Private Sub cboGovtState_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGovtState.Leave
        cboGovtState.BackColor = Color.White
    End Sub

    Private Sub txtPAOCode_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPAOCode.Enter
        txtPAOCode.BackColor = Color.LightYellow
    End Sub

    Private Sub txtPAOCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPAOCode.Leave
        txtPAOCode.BackColor = Color.White
    End Sub

    Private Sub txtPAORegNo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPAORegNo.Leave
        txtPAORegNo.BackColor = Color.White
    End Sub

    Private Sub txtPAORegNo_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPAORegNo.Enter
        txtPAORegNo.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDDOCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDDOCode.Leave
        txtDDOCode.BackColor = Color.White
    End Sub

    Private Sub txtDDOCode_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDDOCode.Enter
        txtDDOCode.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDDORegNo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDDORegNo.Leave
        txtDDORegNo.BackColor = Color.White
    End Sub

    Private Sub txtDDORegNo_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDDORegNo.Enter
        txtDDORegNo.BackColor = Color.LightYellow
    End Sub

    Private Sub cboMinistry_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMinistry.Leave
        cboMinistry.BackColor = Color.White
    End Sub

    Private Sub cboMinistry_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMinistry.Enter
        cboMinistry.BackColor = Color.LightYellow
    End Sub

    Private Sub txtMinistryName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMinistryName.Leave
        txtMinistryName.BackColor = Color.White
    End Sub

    Private Sub txtMinistryName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMinistryName.Enter
        txtMinistryName.BackColor = Color.LightYellow
    End Sub

    Private Sub txtAIN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAIN.Enter
        txtAIN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtAIN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAIN.Leave
        txtAIN.BackColor = Color.White
    End Sub

    Private Sub txtPAORegNo_TextChanged(sender As Object, e As EventArgs) Handles txtPAORegNo.TextChanged

    End Sub

    Private Sub txtPAORegNo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPAORegNo.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub cboGovtState_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboGovtState.SelectedIndexChanged

    End Sub

    Private Sub cboGovtState_KeyDown(sender As Object, e As KeyEventArgs) Handles cboGovtState.KeyDown
        If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub cboGovtState_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboGovtState.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboGovtState.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub cboMinistry_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMinistry.SelectedIndexChanged

    End Sub

    Private Sub cboMinistry_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboMinistry.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboMinistry.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub
End Class