Imports System.Data.OleDb

Public Class frmMulSelDed
    Public sql1 As String
    Public flg As Integer
    Public strflg As Boolean
    Dim chk As New DataGridViewCheckBoxColumn()
    Dim allchecked As Boolean = True
    Dim btbr As Integer
    Private Sub frmMulSelDed_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        'frmCoMst.EnterTab(e)
    End Sub

    Private Sub cmbcancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbcancel.Click
        'Dim i As Integer
        'sql1 = ""
        If allchecked Then
            For Each row As DataGridViewRow In grdbtb.Rows
                row.Cells(1).Value = False
                allchecked = False
            Next
        End If
        'With grdbtb
        '    For i = 1 To .Rows - 1
        '        .TextMatrix(i, 2) = .ValueMatrix(i, 3)
        '        If .Cell(flexcpChecked, i, 2) = 1 Then
        '            If strflg = True Then
        '                Sql = Sql & "'" & .TextMatrix(i, 1) & "'" & ","
        '            Else
        '                Sql = Sql & .TextMatrix(i, 1) & ","
        '            End If
        '        End If
        '    Next i
        'End With
        'If Len(Trim(Sql)) > 1 Then
        '    Sql = "(" & Left(Sql, Len(Trim(Sql)) - 1) & ")"
        'End If
        Close()
        'Me.Hide()
    End Sub

    Private Sub frmMulSelDed_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'grdbtb.Editable = False
        'Me.Left = 3800
        'Me.Top = 2000
        'Main()
        fillbtb()
        sql1 = ""
    End Sub

    Private Sub cmdback_Click(sender As Object, e As EventArgs) Handles cmdback.Click
        'Dim i As Integer
        sql1 = ""
        With grdbtb

            For i = 0 To .Rows.Count - 1
                ' .Rows(i).Cells(2) = .Rows(i).Cells(1).Value
                If .Rows(i).Cells(1).Value = True Then
                    If strflg = True Then
                        sql1 = sql1 & "'" & .Rows(i).Cells(0).Value.ToString & "'" & ","
                    Else
                        sql1 = sql1 & .Rows(i).Cells(0).Value.ToString & ","
                    End If
                End If
            Next i
        End With
        If Len(Trim(sql1)) > 1 Then
            sql1 = "(" & Strings.Left(sql1, Len(Trim(sql1)) - 1) & ")"
            sql1 = sql1
        End If
        Me.Hide()
    End Sub

    Private Sub cmbcl_Click(sender As Object, e As EventArgs) Handles cmbcl.Click
        Dim i As Integer
        ' Dim val1 As DataGridViewCheckBoxCell = DGVList(grdbtb.CurrentCell.ColumnIndex)
        If allchecked Then
            For Each row As DataGridViewRow In grdbtb.Rows
                row.Cells(1).Value = False
                allchecked = False
            Next
        End If
        'With grdbtb
        '    For i = 1 To .Rows.Count - 1
        '        .SelectedCells(1).Value = False
        '    Next i
        'End With
        grdbtb.Focus()
    End Sub

    Private Sub cmbsl_Click(sender As Object, e As EventArgs) Handles cmbsl.Click
        If allchecked Then
            For Each row As DataGridViewRow In grdbtb.Rows
                row.Cells(1).Value = True
                allchecked = True
            Next
        End If
        grdbtb.Focus()
    End Sub

    Private Sub fillbtb()
        'cn.Close()
        'Main()
        Dim ds As New DataSet
        Dim btbr As Integer
        ds = FetchDataSet(sql1)
        grdbtb.DataSource = ds.Tables(0)

        'grdbtb.Columns.Add(chk)
        chk.HeaderText = "Select"
        'chk.Name = "chk"
        grdbtb.Columns.Insert(1, chk)
        'grdbtb.Rows(1).Cells(2).Value = True
        btbr = 1
    End Sub

    Private Sub grdbtb_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdbtb.CellContentClick

    End Sub

    Private Sub grdbtb_Validated(sender As Object, e As EventArgs) Handles grdbtb.Validated

    End Sub
End Class