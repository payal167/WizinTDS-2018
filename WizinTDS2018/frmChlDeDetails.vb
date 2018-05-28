Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class frmChlDeDetails
    Public rtyp As String
    Private Sub frmChlDeDetails_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        'frmCoMst.EnterTab(e)
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

    Private Sub cmbSec_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbsec.Leave
        cmbsec.BackColor = Color.White
    End Sub

    Private Sub cmbSec_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbsec.Enter
        cmbsec.BackColor = Color.LightYellow
    End Sub

    Private Sub cmdgen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdgen.Click
        Dim sql, sqld As String
        Dim xlapp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        xlapp = New Excel.Application
        xlBook = xlapp.Workbooks.Add
        xlSheet = xlBook.Worksheets("Sheet1")
        'xlapp.Visible = True
        'xlSheet = xlBook.Sheets("Sheet1")
        'xlSheet = xlBook.ActiveSheet
        'xlSheet.Name = "Export sheet"
        Dim cmd As New OleDbCommand
        Dim ds As New DataSet
        Dim da As OleDbDataAdapter
        'Select Case (Me.chkallde.CheckState)
        '    Case 0

        '        sql = " SELECT Challan26Q.ChallanID, Challan26Q.Sec, Challan26Q.TaxAmt as TaxAmt1," _
        '           & " Challan26Q.Surcharge as Surcharge1, Challan26Q.ECess as ECess1, Challan26Q.Interest," _
        '           & " Challan26Q.Others, Challan26Q.TotalTax,Challan26Q.BankChallanNo,Challan26Q.TranVouNo, Challan26Q.DtOfChallan," _
        '           & " Deductee26Q.AmtOfPayment, Deductee26Q.DtOfPayment, Deductee26Q.TaxAmt, Deductee26Q.Surcharge," _
        '           & " Deductee26Q.ECess, Deductee26Q.TotalTaxDeposited, CoMst.CoID, CoMst.CoName," _
        '           & " RetnMst.FrmType, DeductMst.DName FROM CoMst INNER JOIN (((Challan26Q INNER JOIN Deductee26Q ON Challan26Q.ChallanID = Deductee26Q.ChallanId)" _
        '           & " INNER JOIN RetnMst ON (RetnMst.RetnID = Challan26Q.RetnID) AND (Deductee26Q.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee26Q.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"

        '    Case 1
        '        sql = " SELECT Challan26Q.ChallanID, Challan26Q.Sec, Challan26Q.TaxAmt AS TaxAmt1, Challan26Q.Surcharge AS Surcharge1, Challan26Q.ECess AS ECess1, Challan26Q.Interest, Challan26Q.Others, Challan26Q.TotalTax, Challan26Q.BankChallanNo, Challan26Q.TranVouNo, Challan26Q.DtOfChallan, Deductee26Q.AmtOfPayment, Deductee26Q.DtOfPayment, Deductee26Q.TaxAmt, Deductee26Q.Surcharge, Deductee26Q.ECess, Deductee26Q.TotalTaxDeposited, CoMst.CoID, CoMst.CoName, RetnMst.FrmType, DeductMst.DName " _
        '            & " FROM CoMst INNER JOIN (((Challan26Q INNER JOIN Deductee26Q ON (Challan26Q.ChallanID = Deductee26Q.ChallanId) OR (Challan26Q.TotalTax<>Deductee26Q.TotalTaxDeposited and Challan26Q.TaxAmt<> Deductee26Q.TaxAmt and Challan26Q.Surcharge<>Deductee26Q.Surcharge and Challan26Q.ECess<>Deductee26Q.ECess)) INNER JOIN RetnMst ON (RetnMst.RetnID = Challan26Q.RetnID) AND (Deductee26Q.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee26Q.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"
        'End Select
        'sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
        'If cmbCNm.SelectedIndex > -1 Then
        '    sql = sql & " and  CoMst.COId = " & cmbCNm.SelectedIndex + 1
        'End If

        'If cmbsec.SelectedIndex = 0 Then
        '    sql = sql
        '    xlSheet.Cells(2, 7) = "All Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1

        'ElseIf cmbsec.SelectedIndex > 0 Then
        '    sql = sql & " and (((Challan26Q.Sec)='" & cmbsec.Text & "'))"
        '    xlSheet.Cells(2, 7) = "Section Wise Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1
        'End If
        ''Report.ReportTitle = Me.lblhead.Caption
        ''Report.FormulaFields(11).Text = Chr(39) & "FY " & FY & Chr(39)
        ''    Report.ParameterFields(11).SetCurrentValue IIf(Check1.Value = 1, True, False)

        'If chkallde.Enabled = 1 Then
        '    xlSheet.Cells(2, 7) = "Mismatched Deductee Details Of " & cmbtyp.Text
        'Else
        '    xlSheet.Cells(2, 7) = "All Deductee Details Of " & cmbtyp.Text
        'End If
        Dim ctl As Control
        Dim rs As New DataSet
        ' Report.DiscardSavedData

        Select Case rtyp
            Case "26Q"
                cmbtyp.Items.Add("26Q")

                Select Case (Me.chkallde.Checked)
                    Case 0

                        sql = " SELECT Challan26Q.ChallanID, Challan26Q.Sec, Challan26Q.TaxAmt as TaxAmt1," _
       & " Challan26Q.Surcharge as Surcharge1, Challan26Q.ECess as ECess1, Challan26Q.Interest," _
       & " Challan26Q.Others, Challan26Q.TotalTax,Challan26Q.BankChallanNo,Challan26Q.TranVouNo,Challan26Q.DtOfChallan," _
       & " Deductee26Q.AmtOfPayment,Deductee26Q.DtOfPayment, Deductee26Q.TaxAmt, Deductee26Q.Surcharge," _
       & " Deductee26Q.ECess, iif(isnull(Deductee26Q.TotalTaxDeposited),0,Deductee26Q.TotalTaxDeposited), CoMst.CoID, CoMst.CoName," _
       & " RetnMst.FrmType, DeductMst.DName FROM CoMst INNER JOIN (((Challan26Q INNER JOIN Deductee26Q ON Challan26Q.ChallanID = Deductee26Q.ChallanId)" _
       & " INNER JOIN RetnMst ON (RetnMst.RetnID = Challan26Q.RetnID) AND (Deductee26Q.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee26Q.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"

                    Case 1
                        sql = " SELECT Challan26Q.ChallanID, Challan26Q.Sec, Challan26Q.TaxAmt AS TaxAmt1, Challan26Q.Surcharge AS Surcharge1, Challan26Q.ECess AS ECess1, Challan26Q.Interest, Challan26Q.Others, Challan26Q.TotalTax, Challan26Q.BankChallanNo, Challan26Q.TranVouNo, Challan26Q.DtOfChallan, Deductee26Q.AmtOfPayment, Deductee26Q.DtOfPayment, Deductee26Q.TaxAmt, Deductee26Q.Surcharge, Deductee26Q.ECess, Deductee26Q.TotalTaxDeposited, CoMst.CoID, CoMst.CoName, RetnMst.FrmType, DeductMst.DName " _
       & " FROM CoMst INNER JOIN (((Challan26Q INNER JOIN Deductee26Q ON (Challan26Q.ChallanID = Deductee26Q.ChallanId) and (Challan26Q.TotalTax<>Deductee26Q.TotalTaxDeposited or Challan26Q.TaxAmt<> Deductee26Q.TaxAmt and Challan26Q.Surcharge<>Deductee26Q.Surcharge and Challan26Q.ECess<>Deductee26Q.ECess)) INNER JOIN RetnMst ON (RetnMst.RetnID = Challan26Q.RetnID) AND (Deductee26Q.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee26Q.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"
                End Select
                If cmbQuarter.SelectedIndex >= 0 Then
                    sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
                ElseIf sqld <> vbNullString Then
                    sql = sql & " where Retnmst.FrmType in " & sqld & ""
                End If
                ' sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.ListIndex + 1 & "'))"

                If cmbCNm.SelectedIndex > -1 Then
                    sql = sql & " and  CoMst.COId = " & cmbCNm.SelectedIndex + 1
                End If

                If cmbsec.SelectedIndex = 0 Then
                    sql = sql
                    xlSheet.Cells(2, 7) = "All Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1

                ElseIf cmbsec.SelectedIndex > 0 Then
                    sql = sql & " and (((Challan26Q.Sec)='" & cmbsec.Text & "'))"
                    sql = sql + "  order by Challan26Q.sec,Challan26Q.DtOfChallan,Challan26Q.BankChallanNo,Challan26Q.challanid"

                    xlSheet.Cells(2, 7) = "Section Wise Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1
                End If
            Case "27Q"
                cmbtyp.Items.Add("27Q")

                Select Case (Me.chkallde.Checked)
                    Case 0

                        sql = " SELECT Challan27Q.ChallanID, Challan27Q.Sec, Challan27Q.TaxAmt as TaxAmt1," _
       & " Challan27Q.Surcharge as Surcharge1, Challan27Q.ECess as ECess1, Challan27Q.Interest," _
       & " Challan27Q.Others, Challan27Q.TotalTax,Challan27Q.BankChallanNo,Challan27Q.TranVouNo,Challan27Q.DtOfChallan," _
       & " Deductee27Q.AmtOfPayment,Deductee27Q.DtOfPayment, Deductee27Q.TaxAmt, Deductee27Q.Surcharge," _
       & " Deductee27Q.ECess, iif(isnull(Deductee27Q.TotalTaxDeposited),0,Deductee27Q.TotalTaxDeposited), CoMst.CoID, CoMst.CoName," _
       & " RetnMst.FrmType, DeductMst.DName FROM CoMst INNER JOIN (((Challan27Q INNER JOIN Deductee27Q ON Challan27Q.ChallanID = Deductee27Q.ChallanId)" _
       & " INNER JOIN RetnMst ON (RetnMst.RetnID = Challan27Q.RetnID) AND (Deductee27Q.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee27Q.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"

                    Case 1
                        sql = " SELECT Challan27Q.ChallanID, Challan27Q.Sec, Challan27Q.TaxAmt AS TaxAmt1, Challan27Q.Surcharge AS Surcharge1, Challan27Q.ECess AS ECess1, Challan27Q.Interest, Challan27Q.Others, Challan27Q.TotalTax, Challan27Q.BankChallanNo, Challan27Q.TranVouNo, Challan27Q.DtOfChallan, Deductee27Q.AmtOfPayment, Deductee27Q.DtOfPayment, Deductee27Q.TaxAmt, Deductee27Q.Surcharge, Deductee27Q.ECess, Deductee27Q.TotalTaxDeposited, CoMst.CoID, CoMst.CoName, RetnMst.FrmType, DeductMst.DName " _
       & " FROM CoMst INNER JOIN (((Challan27Q INNER JOIN Deductee27Q ON (Challan27Q.ChallanID = Deductee27Q.ChallanId) and (Challan27Q.TotalTax<>Deductee27Q.TotalTaxDeposited or Challan27Q.TaxAmt<> Deductee27Q.TaxAmt and Challan27Q.Surcharge<>Deductee27Q.Surcharge and Challan27Q.ECess<>Deductee27Q.ECess)) INNER JOIN RetnMst ON (RetnMst.RetnID = Challan27Q.RetnID) AND (Deductee27Q.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee27Q.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"
                End Select
                If cmbQuarter.SelectedIndex >= 0 Then
                    sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
                ElseIf sqld <> vbNullString Then
                    sql = sql & " where Retnmst.FrmType in " & sqld & ""
                End If
                ' sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.ListIndex + 1 & "'))"

                If cmbCNm.SelectedIndex > -1 Then
                    sql = sql & " and  CoMst.COId = " & cmbCNm.SelectedIndex + 1
                End If

                If cmbsec.SelectedIndex = 0 Then
                    sql = sql
                    xlSheet.Cells(2, 7) = "All Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1

                ElseIf cmbsec.SelectedIndex > 0 Then
                    sql = sql & " and (((Challan27Q.Sec)='" & cmbsec.Text & "'))"
                    sql = sql + "  order by Challan27Q.sec,Challan27Q.DtOfChallan,Challan27Q.BankChallanNo,Challan27Q.challanid"

                    xlSheet.Cells(2, 7) = "Section Wise Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1
                End If

            Case "27EQ"
                cmbtyp.Items.Add("27EQ")

                Select Case (Me.chkallde.Checked)
                    Case 0

                        sql = " SELECT Challan27EQ.ChallanID, Challan27EQ.Sec, Challan27EQ.TaxAmt as TaxAmt1," _
       & " Challan27EQ.Surcharge as Surcharge1, Challan27EQ.ECess as ECess1, Challan27EQ.Interest," _
       & " Challan27EQ.Others, Challan27EQ.TotalTax,Challan27EQ.BankChallanNo,Challan27EQ.TranVouNo,Challan27EQ.DtOfChallan," _
       & " Deductee27EQ.AmtOfPayment,Deductee27EQ.DtOfPayment, Deductee27EQ.TaxAmt, Deductee27EQ.Surcharge," _
       & " Deductee27EQ.ECess, iif(isnull(Deductee27EQ.TotalTaxDeposited),0,Deductee27EQ.TotalTaxDeposited), CoMst.CoID, CoMst.CoName," _
       & " RetnMst.FrmType, DeductMst.DName FROM CoMst INNER JOIN (((Challan27EQ INNER JOIN Deductee27EQ ON Challan27EQ.ChallanID = Deductee27EQ.ChallanId)" _
       & " INNER JOIN RetnMst ON (RetnMst.RetnID = Challan27EQ.RetnID) AND (Deductee27EQ.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee27EQ.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"

                    Case 1
                        sql = " SELECT Challan27EQ.ChallanID, Challan27EQ.Sec, Challan27EQ.TaxAmt AS TaxAmt1, Challan27EQ.Surcharge AS Surcharge1, Challan27EQ.ECess AS ECess1, Challan27EQ.Interest, Challan27EQ.Others, Challan27EQ.TotalTax, Challan27EQ.BankChallanNo, Challan27EQ.TranVouNo, Challan27EQ.DtOfChallan, Deductee27EQ.AmtOfPayment, Deductee27EQ.DtOfPayment, Deductee27EQ.TaxAmt, Deductee27EQ.Surcharge, Deductee27EQ.ECess, Deductee27EQ.TotalTaxDeposited, CoMst.CoID, CoMst.CoName, RetnMst.FrmType, DeductMst.DName " _
       & " FROM CoMst INNER JOIN (((Challan27EQ INNER JOIN Deductee27EQ ON (Challan27EQ.ChallanID = Deductee27EQ.ChallanId) and (Challan27EQ.TotalTax<>Deductee27EQ.TotalTaxDeposited or Challan27EQ.TaxAmt<> Deductee27EQ.TaxAmt and Challan27EQ.Surcharge<>Deductee27EQ.Surcharge and Challan27EQ.ECess<>Deductee27EQ.ECess)) INNER JOIN RetnMst ON (RetnMst.RetnID = Challan27EQ.RetnID) AND (Deductee27EQ.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee27EQ.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"
                End Select
                If cmbQuarter.SelectedIndex >= 0 Then
                    sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
                ElseIf sqld <> vbNullString Then
                    sql = sql & " where Retnmst.FrmType in " & sqld & ""
                End If
                ' sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.ListIndex + 1 & "'))"

                If cmbCNm.SelectedIndex > -1 Then
                    sql = sql & " and  CoMst.COId = " & cmbCNm.SelectedIndex
                End If

                If cmbsec.SelectedIndex = 0 Then
                    sql = sql
                    xlSheet.Cells(2, 7) = "All Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1

                ElseIf cmbsec.SelectedIndex > 0 Then
                    sql = sql & " and (((Challan27EQ.Sec)='" & cmbsec.Text & "'))"
                    sql = sql + "  order by Challan27EQ.sec,Challan27EQ.DtOfChallan,Challan27EQ.BankChallanNo,Challan27EQ.challanid"

                    xlSheet.Cells(2, 7) = "Section Wise Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1
                End If



            Case "24Q"
                cmbtyp.Items.Add("24Q")

                Select Case (Me.chkallde.Checked)
                    Case 0

                        sql = " SELECT Challan24Q.ChallanID, Challan24Q.Sec, Challan24Q.TaxAmt as TaxAmt1," _
           & " Challan24Q.Surcharge as Surcharge1, Challan24Q.ECess as ECess1, Challan24Q.Interest," _
           & " Challan24Q.Others, Challan24Q.TotalTax,Challan24Q.BankChallanNo,Challan24Q.TranVouNo, Challan24Q.DtOfChallan," _
           & " Deductee24Q.AmtOfPayment, Deductee24Q.DtOfPayment, Deductee24Q.TaxAmt, Deductee24Q.Surcharge," _
           & " Deductee24Q.ECess, iif(isnull(Deductee24Q.TotalTaxDeposited),0,Deductee24Q.TotalTaxDeposited), CoMst.CoID, CoMst.CoName," _
           & " RetnMst.FrmType, DeductMst.DName FROM CoMst INNER JOIN (((Challan24Q INNER JOIN Deductee24Q ON Challan24Q.ChallanID = Deductee24Q.ChallanId)" _
           & " INNER JOIN RetnMst ON (RetnMst.RetnID = Challan24Q.RetnID) AND (Deductee24Q.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee24Q.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"


                    Case 1
                        sql = " SELECT Challan24Q.ChallanID, Challan24Q.Sec, Challan24Q.TaxAmt AS TaxAmt1, Challan24Q.Surcharge AS Surcharge1, Challan24Q.ECess AS ECess1, Challan24Q.Interest, Challan24Q.Others, Challan24Q.TotalTax, Challan24Q.BankChallanNo, Challan24Q.TranVouNo, Challan24Q.DtOfChallan, Deductee24Q.AmtOfPayment, Deductee24Q.DtOfPayment, Deductee24Q.TaxAmt, Deductee24Q.Surcharge, Deductee24Q.ECess, Deductee24Q.TotalTaxDeposited, CoMst.CoID, CoMst.CoName, RetnMst.FrmType, DeductMst.DName " _
               & " FROM CoMst INNER JOIN (((Challan24Q INNER JOIN Deductee24Q ON (Challan24Q.ChallanID = Deductee24Q.ChallanId) AND (Challan24Q.TotalTax<>Deductee24Q.TotalTaxDeposited or Challan24Q.TaxAmt<> Deductee24Q.TaxAmt and Challan24Q.Surcharge<>Deductee24Q.Surcharge and Challan24Q.ECess<>Deductee24Q.ECess)) INNER JOIN RetnMst ON (RetnMst.RetnID = Challan24Q.RetnID) AND (Deductee24Q.RetnID = RetnMst.RetnID)) INNER JOIN DeductMst ON Deductee24Q.DId = DeductMst.DId) ON CoMst.CoID = DeductMst.CoID"
                End Select
                If cmbQuarter.SelectedIndex >= 0 Then
                    sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
                ElseIf sqld <> vbNullString Then
                    sql = sql & " where Retnmst.FrmType in " & sqld & ""
                End If

                If cmbCNm.SelectedIndex > -1 Then
                    sql = sql & " and  CoMst.COId = " & cmbCNm.SelectedIndex
                End If

                If cmbsec.SelectedIndex = 0 Then
                    sql = sql
                    xlSheet.Cells(2, 7) = "All Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1

                ElseIf cmbsec.SelectedIndex > 0 Then
                    sql = sql & " and (((Challan24Q.Sec)='" & cmbsec.Text & "'))"
                    sql = sql + "  order by Challan24Q.sec,Challan24Q.DtOfChallan,Challan24Q.BankChallanNo,Challan24Q.challanid"

                    xlSheet.Cells(2, 7) = "Section Wise Deductee Details Of " & cmbtyp.Text & cmbQuarter.SelectedIndex + 1
                End If
        End Select
        'Report.ReportTitle = Me.lblhead.Caption
        'Report.FormulaFields(11).Text = Chr(39) & "FY " & FY & Chr(39)
        If chkallde.Checked = 1 Then
            xlSheet.Cells(2, 7) = "Mismatched Deductee Details Of " & cmbtyp.Text
        Else
            xlSheet.Cells(2, 7) = "All Deductee Details Of " & cmbtyp.Text
        End If
        ' rs.Open sql, Cnn
        '     Set obj.Cnn = Connection
        'Report.Database.SetDataSource rs

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
        'xlSheet.HideColumn(5)
        'xlSheet.HideColumn(10)


        xlSheet.UsedRange.Cells.Columns.AutoFit()
        xlapp.Visible = True
    End Sub

    Private Sub chkallde_CheckedChanged(sender As Object, e As EventArgs) Handles chkallde.CheckedChanged

    End Sub

    Private Sub cmbQuarter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbQuarter.SelectedIndexChanged
        'SecName()
        'If cmbsec.Items.Count <> 0 Then
        '    For i = 0 To cmbsec.Items.Count - 1
        '        If cmbsec.Items(i).ToString = selectedRetnID Then
        '            cmbsec.SelectedIndex = i
        '            Exit For
        '        End If
        '    Next
        'End If
    End Sub

    Private Sub cmdexit_Click(sender As Object, e As EventArgs) Handles cmdexit.Click
        Close()
    End Sub

    Private Sub frmChlDeDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Main()
        '        cmbdest.Clear
        '        cmbdest.AddItem "Screen"
        'cmbdest.AddItem "Printer"
        cmbdest.Items.Add("Export")
        ' Me.cmbdest.ListIndex = 0
        'Me.cmbsec.ListIndex = 0
        'cmbtyp.Items.Clear()
        ' cmbtyp.Items.Add("26Q")
        Select Case rtyp
            Case "26Q"
                cmbtyp.Items.Clear()
                cmbtyp.Items.Add("26Q")
                'cmbtyp.AddItem "Form 27"
                cmbtyp.SelectedIndex = 0
            Case "24Q"
                cmbtyp.Items.Clear()
                cmbtyp.Items.Add("24Q")
                'cmbtyp.AddItem "Form 27"
                cmbtyp.SelectedIndex = 0
            Case "27Q"
                cmbtyp.Items.Clear()
                cmbtyp.Items.Add("27Q")
                cmbtyp.SelectedIndex = 0
            Case "27EQ"
                cmbtyp.Items.Clear()
                cmbtyp.Items.Add("27EQ")
                cmbtyp.SelectedIndex = 0
        End Select
        FillCName()
        'cmbtyp.AddItem "Form 27"
        cmbtyp.SelectedIndex = 0
        cmbQuarter.Items.Clear()
        cmbQuarter.Items.Add("Quarter 1")
        cmbQuarter.Items.Add("Quarter 2")
        cmbQuarter.Items.Add("Quarter 3")
        cmbQuarter.Items.Add("Quarter 4")
        cmbQuarter.SelectedIndex = 0

        'If cmbCNm.ListCount <> 0 Then
        '    For i = 0 To cmbCNm.ListCount - 1
        '        If cmbCNm.ItemData(i) = selectedcoid Then
        '            cmbCNm.ListIndex = i
        '            Exit For
        '        End If
        '    Next
        'End If

        SecName()
        'If cmbsec.ListCount <> 0 Then
        '    For i = 0 To cmbsec.ListCount - 1
        '        If cmbsec.ItemData(i) = selectedChallanID Then
        '            cmbsec.ListIndex = i
        '            Exit For
        '        End If
        '    Next
        'End If
    End Sub

    Private Sub FillCName()
        Dim ds As DataSet
        ds = FetchDataSet("Select CoId,CoName From CoMst Order By CoName")
        cmbCNm.DataSource = ds.Tables(0)
        cmbCNm.ValueMember = "CoId"
        cmbCNm.DisplayMember = "CoName"
        ds.Dispose()
        'If rs.RecordCount > 0 Then
        '    cmbCNm.Clear
        '    While Not rs.EOF
        '        cmbCNm.AddItem rs!CoName & ""
        '    cmbCNm.ItemData(cmbCNm.NewIndex) = rs!coid
        '        rs.MoveNext
        'Wend
        'cmbCNm.ListIndex = 0
        'End If
    End Sub

    Private Sub SecName()
        Dim rs As New DataSet
        Dim sql As String
        Dim t1 As New DataTable()
        Dim da As New OleDbDataAdapter
        Dim cmd As New OleDbCommand
        Dim transaction As OleDbTransaction
        '    sql = " SELECT DISTINCT Challan26Q.Sec" _
        '& " FROM RetnMst INNER JOIN Challan26Q ON RetnMst.RetnID = Challan26Q.RetnID " _
        '& " WHERE (((RetnMst.RetnID)=[Challan26Q].[RetnID]))"

        'sql = "SELECT distinct Challan26Q.Sec, Challan26Q.ChallanID, RetnMst.RetnID, RetnMst.FrmType " _
        '    & "FROM (RetnMst INNER JOIN Deductee26Q ON RetnMst.RetnID = Deductee26Q.RetnID) INNER JOIN Challan26Q ON (Deductee26Q.ChallanId = Challan26Q.ChallanID) AND (RetnMst.RetnID = Challan26Q.RetnID)"
        'sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
        ''cmd.Connection = cn
        'transaction = cn.BeginTransaction()
        'cmd.Transaction = transaction
        'cmd.CommandText = sql
        'Dim dr As OleDbDataReader = cmd.ExecuteReader
        'While dr.Read
        '    cmbsec.Items.Add(dr.Item(0))
        'End While
        'dr.Close()
        'Try
        '    cmd.ExecuteNonQuery()
        '    transaction.Commit()
        'Catch ex As Exception
        '    transaction.Rollback()
        '    MessageBox.Show(ex.Message) 'Error MEssage
        'End Try
        'cmd.Dispose()
        'transaction.Dispose()
        'If rs.RecordCount > 0 Then
        '    cmbsec.Clear
        '    cmbsec.AddItem "Select All"

        'While Not rs.EOF
        '        cmbsec.AddItem rs!Sec & ""
        '    cmbsec.ItemData(cmbsec.NewIndex) = rs!RetnID
        '        rs.MoveNext
        'Wend
        'cmbsec.ListIndex = 0
        'End If
        ' Dim rs As New ADODB.Recordset
        ' Dim sql As String
        Select Case rtyp
            Case "26Q"
                cmbsec.Items.Clear()

                sql = " SELECT DISTINCT Challan26Q.Sec, RetnMst.FrmType" _
                    & " FROM RetnMst INNER JOIN Challan26Q ON RetnMst.RetnID = Challan26Q.RetnID " _
                    & " WHERE (((RetnMst.RetnID)=[Challan26Q].[RetnID]))"

            Case "24Q"
                cmbsec.Items.Clear()
                sql = " SELECT DISTINCT Challan24Q.Sec, RetnMst.FrmType" _
                    & " FROM RetnMst INNER JOIN Challan24Q ON RetnMst.RetnID = Challan24Q.RetnID " _
                    & " WHERE (((RetnMst.RetnID)=[Challan24Q].[RetnID]))"
            Case "27Q"
                cmbsec.Items.Clear()
                sql = " SELECT DISTINCT Challan27Q.Sec, RetnMst.FrmType" _
                    & " FROM RetnMst INNER JOIN Challan27Q ON RetnMst.RetnID = Challan27Q.RetnID " _
                    & " WHERE (((RetnMst.RetnID)=[Challan27Q].[RetnID]))"
            Case "27EQ"
                cmbsec.Items.Clear()
                sql = " SELECT DISTINCT Challan27EQ.Sec, RetnMst.FrmType" _
                    & " FROM RetnMst INNER JOIN Challan27EQ ON RetnMst.RetnID = Challan27EQ.RetnID " _
                    & " WHERE (((RetnMst.RetnID)=[Challan27EQ].[RetnID]))"

        End Select
        '    sql = " SELECT Distinct RetnMst.RetnID, RetnMst.FrmType, Challan26Q.Sec, Challan26Q.ChallanID " _
        '        & " FROM RetnMst INNER JOIN Challan26Q ON RetnMst.RetnID = Challan26Q.RetnID"
        sql = sql & " and (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction
        cmd.CommandText = sql
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        While dr.Read
            cmbsec.Items.Add(dr.Item(0))
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
        'rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly

        '        If rs.RecordCount > 0 Then
        '            cmbsec.Clear
        '            cmbsec.AddItem "Select All"

        '        While Not rs.EOF
        '                cmbsec.AddItem rs!Sec & ""
        '            'cmbsec.ItemData(cmbsec.NewIndex) = rs!RetnID
        '                rs.MoveNext
        '        Wend
        '        cmbsec.ListIndex = 0
        '        End If
        '        If rs.State = adStateOpen Then rs.Close
        'Set rs = Nothing
    End Sub
End Class