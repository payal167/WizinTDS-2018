Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class frmchlnlist
    Dim sqld As String
    Dim cmd As OleDbCommand
    Private Sub frmchlnlist_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub cmbDesti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbdest.Enter
        cmbdest.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbDesti_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbdest.Leave
        cmbdest.BackColor = Color.White

    End Sub

    Private Sub cmbCompName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCNm.Leave
        cmbCNm.BackColor = Color.White

    End Sub

    Private Sub cmbCompName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCNm.Enter
        cmbCNm.BackColor = Color.LightYellow

    End Sub

    Private Sub cmbFrmTyp_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbQuarter.Leave
        cmbQuarter.BackColor = Color.White

    End Sub

    Private Sub cmbFrmTyp_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbQuarter.Enter
        cmbQuarter.BackColor = Color.LightYellow

    End Sub

    Private Sub cmbQuar_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbtyp.Leave
        cmbtyp.BackColor = Color.White

    End Sub

    Private Sub cmbQuar_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbtyp.Enter
        cmbtyp.BackColor = Color.LightYellow
    End Sub

    Private Sub frmchlnlist_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmbdest.Items.Clear()
        'cmbdest.AddItem "Screen"
        'cmbdest.AddItem "Printer"
        cmbdest.Items.Add("Export")
        Me.cmbdest.SelectedIndex = 0
        'cmbtyp.Clear()
        cmbtyp.Items.Add("24Q")
        cmbtyp.Items.Add("26Q")
        cmbtyp.Items.Add("27EQ")
        cmbtyp.Items.Add("27Q")
        cmbtyp.SelectedIndex = 0
        FillCName()
        'With cmbQuarter
        '    .Clear
        '    .AddItem "Quarter 1"
        '    .AddItem "Quarter 2"
        '    .AddItem "Quarter 3"
        '    .AddItem "Quarter 4"
        '    .ListIndex = 0
        'End With

        'FillCName()
        'If cmbCNm.ListCount <> 0 Then
        '    For i = 0 To cmbCNm.ListCount - 1
        '        If cmbCNm.ItemData(i) = selectedcoid Then
        '            cmbCNm.ListIndex = i
        '            Exit For
        '        End If
        '    Next
        'End If
    End Sub

    Private Sub FillCName()
        Dim ds As DataSet
        ds = FetchDataSet("Select CoId,CoName From CoMst")
        cmbCNm.DataSource = ds.Tables(0)
        cmbCNm.ValueMember = "CoId"
        cmbCNm.DisplayMember = "CoName"
        ds.Dispose()
        '        Dim rs As New ADODB.Recordset
        '        Dim sql As String
        '        sql = "Select CoId,CoName From CoMst Order By CoName"
        '        rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
        '    If rs.RecordCount > 0 Then
        '            cmbCNm.Clear
        '            While Not rs.EOF
        '                cmbCNm.AddItem rs!CoName & ""
        '            cmbCNm.ItemData(cmbCNm.NewIndex) = rs!coid
        '                rs.MoveNext
        '        Wend
        '        cmbCNm.ListIndex = 0
        '        End If
        '        If rs.State = adStateOpen Then rs.Close
        'Set rs = Nothing
    End Sub

    Private Sub cmdgen_Click(sender As Object, e As EventArgs) Handles cmdgen.Click
        ' Dim ctl As Control
        Dim sql As String
        Dim ftyp As String
        Dim rs As New DataSet

        ' Report.DiscardSavedData
        Select Case (Me.cmbtyp.SelectedIndex)
            Case 0
                sql = "SELECT CO.CoID, CO.CoName, R.RetnID, R.AYear, R.FrmType, R.TxtFileName, " _
                & " C.ChallanID, C.Sec, C.TaxAmt, C.BankChallanNo, C.DtOfChallan, FORMAT(C.BankBrCode,'0000000'), C.Surcharge, C.ECess, C.Interest, C.Others,C.ChqDDNo , C.TranVouNo,c.AFees,c.MinorHead" _
                & " FROM CoMst AS CO INNER JOIN (RetnMst AS R INNER JOIN Challan24Q AS C ON R.RetnID = C.RetnID) ON CO.CoID = R.CoID "
               ' ftyp = Right(Trim(cmbtyp.Text), 3) + Right(cmbQuarter.Text, 1)
            Case 1
                sql = "SELECT CO.CoID, CO.CoName, R.RetnID, R.AYear, R.FrmType, R.TxtFileName, " _
                & " C.ChallanID, C.Sec, C.TaxAmt, C.BankChallanNo, C.DtOfChallan, FORMAT(C.BankBrCode,'0000000'), C.Surcharge, C.ECess, C.Interest, C.Others,C.ChqDDNo , C.TranVouNo,c.AFees,c.MinorHead" _
                & " FROM CoMst AS CO INNER JOIN (RetnMst AS R INNER JOIN Challan26Q AS C ON R.RetnID = C.RetnID) ON CO.CoID = R.CoID "
               ' ftyp = Right(Trim(cmbtyp.Text), 3) + Right(cmbQuarter.Text, 1)

            Case 2
                sql = "SELECT CO.CoID, CO.CoName, R.RetnID, R.AYear, R.FrmType, R.TxtFileName, " _
                & " C.ChallanID, C.Sec, C.TaxAmt, C.BankChallanNo, C.DtOfChallan, FORMAT(C.BankBrCode,'0000000'), C.Surcharge, C.ECess, C.Interest, C.Others ,C.ChqDDNo, C.TranVouNo,c.AFees,c.MinorHead" _
                & " FROM CoMst AS CO INNER JOIN (RetnMst AS R INNER JOIN Challan27EQ AS C ON R.RetnID = C.RetnID) ON CO.CoID = R.CoID "
               ' ftyp = Right(Trim(cmbtyp.Text), 4) + Right(cmbQuarter.Text, 1)
            Case 3
                sql = "SELECT CO.CoID, CO.CoName, R.RetnID, R.AYear, R.FrmType, R.TxtFileName, " _
                & " C.ChallanID, C.Sec, C.TaxAmt, C.BankChallanNo, C.DtOfChallan, FORMAT(C.BankBrCode,'0000000'), C.Surcharge, C.ECess, C.Interest, C.Others,C.ChqDDNo , C.TranVouNo,c.AFees,c.MinorHead" _
                & " FROM CoMst AS CO INNER JOIN (RetnMst AS R INNER JOIN Challan27Q AS C ON R.RetnID = C.RetnID) ON CO.CoID = R.CoID "
                '  ftyp = Right(Trim(cmbtyp.Text), 3) + Right(cmbQuarter.Text, 1)
        End Select
        If cmbQuarter.SelectedIndex >= 0 Then
            ' ftyp = Right(Trim(cmbtyp.Text), 3) + Right(cmbQuarter.Text, 1)
            sql = sql & " where r.frmtype='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'"

            If cmbCNm.SelectedIndex > -1 Then
                sql = sql & " and  CO.COId = " & cmbCNm.SelectedValue
            End If

            ElseIf sqld <> vbNullString Then
                ' sql = sql & " where R.FrmType in " & sqld & ""
                sql = sql & " where R.FrmType in " & "(" & sqld & ")"
            If cmbCNm.SelectedIndex > -1 Then
                sql = sql & " and  CO.COId = " & cmbCNm.SelectedValue
            End If
        End If
        sql = sql & " Order By C.ChallanId"

        Dim xlapp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        'Dim rs As New ADODB.Recordset
        ' Dim R As Long, c As Long
        xlapp = New Excel.Application
        xlBook = xlapp.Workbooks.Add
        xlSheet = xlBook.Worksheets("Sheet1")
        ' xlapp.Visible = True
        xlSheet = xlBook.Sheets("Sheet1")
        xlSheet = xlBook.ActiveSheet
        xlSheet.Name = "Export sheet"




        Dim cmd As New OleDbCommand
        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter
        cmd = New OleDbCommand(sql, cn)
        da.SelectCommand = cmd
        ds = New DataSet()
        da.Fill(ds)


        DataGridView1.DataSource = ds.Tables(0)
        xlSheet.Range(xlSheet.Cells(3, 8), xlSheet.Cells(3, 8)).Value = "Amount"
        xlSheet.Range(xlSheet.Cells(3, 9), xlSheet.Cells(3, 9)).Value = "ChallanNo/TranVouNo"
        xlSheet.Range(xlSheet.Cells(3, 10), xlSheet.Cells(3, 10)).Value = "ChallanDt"
        xlSheet.Range(xlSheet.Cells(3, 11), xlSheet.Cells(3, 11)).Value = "BankBRCode"
        'xlSheet.Range(xlSheet.Cells(3, 16), xlSheet.Cells(3, 16)).Value = "Total"
        xlSheet.Range(xlSheet.Cells(3, 12), xlSheet.Cells(3, 12)).Value = "Surcharge"
        xlSheet.Range(xlSheet.Cells(3, 13), xlSheet.Cells(3, 13)).Value = "ECess"
        xlSheet.Range(xlSheet.Cells(3, 15), xlSheet.Cells(3, 15)).Value = "Other"
        xlSheet.Range(xlSheet.Cells(3, 16), xlSheet.Cells(3, 16)).Value = "ChqDDNo"
        xlSheet.Range(xlSheet.Cells(3, 8), xlSheet.Cells(3, 16)).Font.Bold = True
        'xlSheet.Range(xlSheet.Cells(3, 13), xlSheet.Cells(3, 13)).BorderAround()

        Dim r As String
        r = (xlSheet.UsedRange.Rows.Count) + 1
        For m = 0 To DataGridView1.Rows.Count - 2
            For i = 0 To DataGridView1.Columns.Count - 2
                xlSheet.Cells(m + 4, i + 1) = DataGridView1.Rows(m).Cells(i + 1).Value
                xlSheet.Range(xlSheet.Cells(r + 1, i + 1), xlSheet.Cells(r + 1, i + 1)).BorderAround(10)
                xlSheet.Range(xlSheet.Cells(m + 4, i + 1), xlSheet.Cells(m + 4, i + 1)).BorderAround(10)
            Next
        Next

        xlSheet.UsedRange.Cells.Columns.AutoFit()
        xlSheet.Columns("L:L").Cut
        xlSheet.Columns("I:I").Insert
        xlSheet.Columns("M:M").Cut
        xlSheet.Columns("J:J").Insert
        xlSheet.Columns("O:O").Cut
        xlSheet.Columns("K:K").Insert
        'xlSheet.Columns("I:I").Cut
        'xlSheet.Columns("P:P").Insert
        'xlSheet.Range(3, 9).Copy(xlSheet.Range(3, 17))
        'For row = 3 To xlSheet.UsedRange.Rows.Count + 1
        '    'row = (xlSheet.UsedRange.Rows.Count)
        '    xlSheet.Range(xlSheet.Cells(row, 4), xlSheet.Cells(row, 8)).Merge()
        'Next
        'xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, 1)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(4, 2), xlSheet.Cells(4, 2)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(4, 3), xlSheet.Cells(3, 3)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(4, 4), xlSheet.Cells(3, 4)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(3, 5), xlSheet.Cells(3, 5)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(3, 6), xlSheet.Cells(3, 6)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(3, 7), xlSheet.Cells(3, 7)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(3, 15), xlSheet.Cells(3, 15)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(3, 17), xlSheet.Cells(3, 17)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(3, 18), xlSheet.Cells(3, 18)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(3, 19), xlSheet.Cells(3, 19)).EntireColumn.Hidden = True
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 8), xlSheet.Cells(3, 8)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 8), xlSheet.Cells(3, 8)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 8), xlSheet.Cells(3, 8)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 8), xlSheet.Cells(3, 8)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 8), xlSheet.Cells(3, 8)).EntireColumn.Delete()
        If cmbQuarter.Text = "" Then
            xlSheet.Range(xlSheet.Cells(2, 2), xlSheet.Cells(2, 7)).Merge()
            xlSheet.Range(xlSheet.Cells(2, 2), xlSheet.Cells(2, 7)).Value = "Challan Detail List Of " & cmbtyp.Text
        Else
            xlSheet.Range(xlSheet.Cells(2, 3), xlSheet.Cells(2, 7)).Merge()
            xlSheet.Range(xlSheet.Cells(2, 3), xlSheet.Cells(2, 7)).Value = "Challan Detail List Of " & cmbQuarter.Text
        End If
        xlSheet.UsedRange.Cells.Columns.Font.Size = 8
        xlSheet.UsedRange.Cells.Columns.BorderAround(10)
        xlSheet.UsedRange.Cells.Columns.AutoFit()
        xlapp.Visible = True
    End Sub
    Private Sub Qtr()
        'Main()
        Dim rs As New DataSet
        'Dim rs1 As New ADODB.Recordset
        Dim sql As String
        Dim t1 As New DataTable()
        Dim da As New OleDbDataAdapter
        Dim cmd As New OleDbCommand
        Dim transaction As OleDbTransaction
        ' Dim sql1 As String
        cmbQuarter.Items.Clear()
        ' cmbQuarter.SelectedIndex = -1
        sql = "SELECT distinct RetnMst.FrmType FROM RetnMst" ' where COId = " & cmbCNm.ItemData(cmbCNm.ListIndex)
        cmbQuarter.Items.Clear()
        If cmbtyp.SelectedIndex = 0 Then
            sql = sql & " where RetnMst.frmtype in ('24Q1','24Q2','24Q3','24Q4')"
        ElseIf cmbtyp.SelectedIndex = 1 Then
            sql = sql & " where RetnMst.frmtype in ('26Q1','26Q2','26Q3','26Q4')"
        ElseIf cmbtyp.SelectedIndex = 2 Then
            sql = sql & " where RetnMst.frmtype in ('27EQ1','27EQ2','27EQ3','27EQ4')"
        ElseIf cmbtyp.SelectedIndex = 3 Then
            sql = sql & " where RetnMst.frmtype in ('27Q1','27Q2','27Q3','27Q4')"
        End If

        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction
        cmd.CommandText = sql
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        While dr.Read
            cmbQuarter.Items.Add(dr.Item(0))
        End While
        dr.Close()
        Try
            cmd.ExecuteNonQuery()
            transaction.Commit()
        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show(ex.Message) 'Error MEssage
        End Try
        cmd.Dispose()
        transaction.Dispose()
    End Sub

    Private Sub cmbtyp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbtyp.SelectedIndexChanged
        ' cn.Close()
        ' Main()
        Qtr()
    End Sub

    Private Sub cmdexit_Click(sender As Object, e As EventArgs) Handles cmdexit.Click
        Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub cmbQuarter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbQuarter.SelectedIndexChanged

    End Sub

    Private Sub cmdc_Click(sender As Object, e As EventArgs) Handles cmdc.Click
        Dim rs As New DataSet
        Dim sql As String
        Dim frmd As New frmMulSelDed
        sql = "SELECT distinct RetnMst.FrmType FROM RetnMst" ' where COId = " & cmbCNm.ItemData(cmbCNm.ListIndex)
        If cmbtyp.SelectedIndex = 0 Then
            sql = sql & " where RetnMst.frmtype in ('24Q1','24Q2','24Q3','24Q4')"

        ElseIf cmbtyp.SelectedIndex = 1 Then
            sql = sql & " where RetnMst.frmtype in ('26Q1','26Q2','26Q3','26Q4')"
        ElseIf cmbtyp.SelectedIndex = 2 Then
            sql = sql & " where RetnMst.frmtype in ('27EQ1','27EQ2','27EQ3','27EQ4')"
        End If
        ' rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
        cmbQuarter.SelectedIndex = -1
        frmd.sql1 = sql
        frmd.strflg = True
        frmd.ShowDialog()
        sqld = frmd.sql1
        If sqld = "" Then
            cmbQuarter.Enabled = True
        Else
            cmbQuarter.Enabled = False
        End If
    End Sub
End Class