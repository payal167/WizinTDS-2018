Public Class frm_Find
    Public obj As Cls_Find
    Public findflag As Boolean
    Public markst As String
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_cancel.Click
        findflag = False
        Me.Close()
    End Sub

    Public Sub lvw_resultshead()
        With lvw_results
            .Columns.Clear()
            .Columns.Add("DESCRIPTION", 200, HorizontalAlignment.Left)
            .Columns.Add("ID", 50, HorizontalAlignment.Left)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True

        End With
    End Sub

    'Private Sub frm_Find_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
    '    frmCoMst.EnterTab(e)
    'End Sub

    Private Sub txtSearch_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_find.Leave
        txt_find.BackColor = Color.White
    End Sub

    Private Sub txtSearch_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_find.Enter
        txt_find.BackColor = Color.LightYellow
    End Sub

    Private Sub Cmd_OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_OK.Click
        findflag = True
        If Me.lvw_results.Items.Count > 0 Then
            If markst = "m" Then
                obj.mrkgid = lvw_results.SelectedItems(0).SubItems(1).Text
                obj.gdesc = lvw_results.SelectedItems(0).SubItems(0).Text
                'ElseIf lvw_results.Text = "" Then

                '    Me.Close()
                '     Exit Sub
            Else
                obj.gid = lvw_results.SelectedItems(0).SubItems(1).Text
                obj.gdesc = lvw_results.SelectedItems(0).SubItems(0).Text
            End If
        End If
        Me.Close()
    End Sub

    Private Sub lvw_results_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub frm_Find_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.Left = 2600
        'Me.Top = 1450
        lvw_resultshead()
        obj = New Cls_Find
        obj.FillList(findsql1 & txtfind & findsql2, Me.lvw_results)
        Me.txt_find.Text = txtfind
    End Sub

    Private Sub lvw_results_DoubleClick(sender As Object, e As EventArgs) Handles lvw_results.DoubleClick
        findflag = True
        If Me.lvw_results.Items.Count > 0 Then
            If markst = "m" Then
                obj.mrkgid = lvw_results.SelectedItems(0).SubItems(1).Text
                obj.gdesc = lvw_results.SelectedItems(0).SubItems(0).Text
            Else
                obj.gid = lvw_results.SelectedItems(0).SubItems(1).Text
                obj.gdesc = lvw_results.SelectedItems(0).SubItems(0).Text
            End If
        End If
        Me.Close()
        'frmdeduteeTDSMST.list1()
        'If findflag = True Then
        '    frmdeduteeTDSMST.cboDName.SelectedIndex = frmdeduteeTDSMST.cboDName.FindString(lvw_results.SelectedItems(0).SubItems(0).Text)

        '    'frmdeduteeTDSMST.DDid = obj.gid
        '    'For i = 0 To frmdeduteeTDSMST.cboDName.Items.Count - 1
        '    '    If frmdeduteeTDSMST.cboDName.SelectedItem(i) = obj.gid Then
        '    '        frmdeduteeTDSMST.cboDName.SelectedIndex = i
        '    '        Exit For
        '    '    End If
        '    'Next i
        'End If
    End Sub

    Private Sub txt_find_TextChanged(sender As Object, e As EventArgs) Handles txt_find.TextChanged
        Me.lvw_results.Items.Clear()
        obj.FillList(findsql1 & Me.txt_find.Text & findsql2, Me.lvw_results)
    End Sub

    Private Sub txt_find_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_find.KeyUp
        If e.KeyCode = 40 Then
            Me.lvw_results.Focus()
        End If
    End Sub

    Private Sub lvw_results_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles lvw_results.SelectedIndexChanged

    End Sub
End Class