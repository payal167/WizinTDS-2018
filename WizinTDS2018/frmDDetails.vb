Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class frmDDetails
    Public rtyp1 As String
    Dim frmd As New frmMulSelDed
    Private Sub frmDDetails_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        'frmCoMst.EnterTab(e)
    End Sub

    Private Sub cmbDesti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbdest.Enter
        cmbdest.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbDesti_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbdest.Leave
        cmbdest.BackColor = Color.White

    End Sub

    Private Sub cmbCompName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCNm.Enter
        cmbCNm.BackColor = Color.LightYellow

    End Sub

    Private Sub cmbCompName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCNm.Leave
        cmbCNm.BackColor = Color.White

    End Sub

    Private Sub cmbFrmTyp_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbtyp.Leave
        cmbtyp.BackColor = Color.White
    End Sub

    Private Sub cmbFrmTyp_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbtyp.Enter
        cmbtyp.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbQuar_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbQuarter.Enter
        cmbQuarter.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbQuar_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbQuarter.Leave
        cmbQuarter.BackColor = Color.White
    End Sub

    Private Sub cmbSec_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbsec.Enter
        cmbsec.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbSec_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbsec.Leave
        cmbsec.BackColor = Color.White
    End Sub

    Private Sub frmDDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmbdest.Items.Clear()
        cmbdest.Items.Add("Export")
        'cmbdest.Items.Add("Printer")
        Me.cmbdest.SelectedIndex = 0
        cmbtyp.Items.Clear()
        cmbtyp.Items.Add("Form 26")
        cmbtyp.Items.Add("Form 27")
        cmbtyp.SelectedIndex = 0
        FillCName()
    End Sub

    Private Sub chkChallan_CheckedChanged(sender As Object, e As EventArgs) Handles chkChallan.CheckedChanged

    End Sub

    Private Sub cmdgen_Click(sender As Object, e As EventArgs) Handles cmdgen.Click
        Dim xlapp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        xlapp = New Excel.Application
        xlBook = xlapp.Workbooks.Add
        xlSheet = xlBook.Worksheets("Sheet1")
        xlapp.Visible = True
        xlSheet = xlBook.Sheets("Sheet1")
        xlSheet = xlBook.ActiveSheet
        xlSheet.Name = "Export sheet"
        Dim cmd As New OleDbCommand
        Dim sql As String
        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter

        sql = " SELECT CO.CoID, CO.CoName, R.RetnID, R.AYear, R.FrmType, DeducteeTDS.DId, DeductMst.DName, DeductMst.DPan, DeducteeTDS.Sec, DeducteeTDS.AmtOfPay, DeducteeTDS.DtOfPay, DeducteeTDS.RateOfTDS, DeducteeTDS.AmtOfTDS, DeducteeTDS.DtOfTDS, DeducteeTDS.DtOfTDSPay AS DtOfChallan, DeducteeTDS.BankBrCode, DeducteeTDS.ChallanNo," _
      & " (SELECT  Sum(Challan.Amt) From Challan Where Challan.RetnId = r.RetnId And Challan.Sec = DeducteeTDS.Sec And Challan.DtOfVoucher = DeducteeTDS.DtOfTDSPay GROUP BY Challan.RetnID, Challan.Sec, Challan.DtOfVoucher) AS ChTotAmt" _
      & " FROM CoMst AS CO INNER JOIN ((RetnMst AS R INNER JOIN DeducteeTDS ON R.RetnID = DeducteeTDS.RetnID) INNER JOIN DeductMst ON DeducteeTDS.DId = DeductMst.DId) ON CO.CoID = R.CoID "


        Select Case cmbtyp.SelectedIndex
            Case 0
                sql = sql & " where frmtype=26"

            Case 1
                sql = sql & " where frmtype=27"
        End Select

        If cmbCNm.SelectedIndex > -1 Then
            sql = sql & " and  co.COId = " & cmbCNm.SelectedIndex + 1
        End If

        sql = sql & " ORDER BY DeducteeTDS.DtOfTDSPay, DeductMst.DName, DeducteeTDS.DtOfPay"

        'Report.ReportTitle = Me.lblhead.Caption
        'Report.FormulaFields(2).Text = Chr(39) & "FY " & FY & Chr(39)
        'Report.ParameterFields(1).SetCurrentValue IIf(Check1.Value = 1, True, False)

        If chkChallan.Checked = 1 Then
            xlSheet.Cells(2, 7) = "All Deductee Details Of " & cmbtyp.Text
        Else
            xlSheet.Cells(2, 7) = "Mismatched Deductee Details Of " & cmbtyp.Text
        End If
        cmd = New OleDbCommand(sql, cn)
        da.SelectCommand = cmd
        ds = New DataSet()
        da.Fill(ds)

        xlSheet.Range(xlSheet.Cells(3, 3), xlSheet.Cells(3, 3)).Value = "Deductee Name"
        xlSheet.Range(xlSheet.Cells(3, 4), xlSheet.Cells(3, 4)).Value = "Address"
        xlSheet.Range(xlSheet.Cells(3, 9), xlSheet.Cells(3, 9)).Value = "State Name"
        xlSheet.Range(xlSheet.Cells(3, 10), xlSheet.Cells(3, 10)).Value = "Pin Code"
        xlSheet.Range(xlSheet.Cells(3, 12), xlSheet.Cells(3, 12)).Value = "PAN No."
        xlSheet.Range(xlSheet.Cells(3, 13), xlSheet.Cells(3, 13)).Value = "Type"
        xlSheet.Range(xlSheet.Cells(3, 16), xlSheet.Cells(3, 16)).Value = "Category"
        xlSheet.Range(xlSheet.Cells(3, 3), xlSheet.Cells(3, 16)).Font.Bold = True
        xlSheet.Range(xlSheet.Cells(3, 3), xlSheet.Cells(3, 16)).BorderAround()
        DataGridView1.DataSource = ds.Tables(0)
        Dim r As String
        r = (xlSheet.UsedRange.Rows.Count) + 1
        For m = 0 To DataGridView1.Rows.Count - 2
            For i = 0 To DataGridView1.Columns.Count - 2
                xlSheet.Cells(m + 4, i + 1) = DataGridView1.Rows(m).Cells(i + 1).Value.ToString()
                xlSheet.Range(xlSheet.Cells(r, i + 1), xlSheet.Cells(r, i + 1)).BorderAround(10)
                xlSheet.Range(xlSheet.Cells(m + 4, i + 1), xlSheet.Cells(m + 4, i + 1)).BorderAround(10)
            Next
        Next

        xlSheet.UsedRange.Cells.Columns.AutoFit()
    End Sub

    Private Sub cmdexit_Click(sender As Object, e As EventArgs) Handles cmdexit.Click
        Close()
    End Sub
    Private Sub FillCName()
        Dim ds As DataSet
        ds = FetchDataSet("Select CoId,CoName From CoMst")
        cmbCNm.DataSource = ds.Tables(0)
        cmbCNm.ValueMember = "CoId"
        cmbCNm.DisplayMember = "CoName"
        ds.Dispose()
    End Sub

    Private Sub cmdc_Click(sender As Object, e As EventArgs) Handles cmdc.Click
        Dim sqld As String
        Dim rs As New DataSet
        Dim sql As String
        sql = "SELECT distinct RetnMst.FrmType FROM RetnMst" ' where COId = " & cmbCNm.ItemData(cmbCNm.ListIndex)
        If cmbtyp.Text = "24Q" Then
            sql = sql & " where RetnMst.frmtype in ('24Q1','24Q2','24Q3','24Q4')"

        ElseIf cmbtyp.Text = "26Q" Then
            sql = sql & " where RetnMst.frmtype in ('26Q1','26Q2','26Q3','26Q4')"
        ElseIf cmbtyp.Text = "27EQ" Then
            sql = sql & " where RetnMst.frmtype in ('27EQ1','27EQ2','27EQ3','27EQ4')"
        End If
        ' rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
        'cmbQuarter.ListIndex = -1
        frmd.sql1 = sql
        frmd.strflg = True
        frmd.Show 'vbModal
        sqld = frmd.sql1
        If sqld = "" Then
            cmbQuarter.Enabled = True
        Else
            cmbQuarter.Enabled = False
        End If
        frmd.Hide
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class