Imports System
Imports System.Collections
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Data
Imports System.String
Imports System.ComponentModel

Public Class frmTDS
    Public ConvertWhich As String
    Public CorrectWhich As String

    Private Sub CertificateDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CertificateDetailsToolStripMenuItem.Click
        frmCIT.Show()
    End Sub

    Private Sub TanRegistrastionDetailsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TanRegistrastionDetailsToolStripMenuItem1.Click
        frmTanReg.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
    Private Sub MaterDedcToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MaterDedcToolStripMenuItem.Click
        frmdeduteeTDSMST.Show()
    End Sub

    Private Sub BankMasterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BankMasterToolStripMenuItem.Click
        frmBankMst.Show()
    End Sub

    Private Sub VerifyDeducteePANToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VerifyDeducteePANToolStripMenuItem.Click
        frmPanVer.Show()
    End Sub

    Private Sub DeducteeListToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeducteeListToolStripMenuItem1.Click
        frmDeList.Show()
    End Sub

    Private Sub ChallanDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChallanDetailsToolStripMenuItem.Click
        frmchlnlist.Show()
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Dim RetnID As Long
        'Dim LoadForm As Boolean
        'Dim rst As New ADODB.Recordset
        RetnID = GetForm24RetnID("2")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/July/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("30/Sep/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ConvertWhich = "24Q2"
        frmTDS24Q.Tag = RetnID
        frmTDS24Q.tabMain.SelectedIndex = 0
        frmTDS24Q.quter = "24Q2"
        'frmTDS24Q.Show      'just to give a visual effect of loading...
        If CheckPRNNo(frmTDS24Q.quter) = False Then
            Call Load24QData(RetnID)
            frmTDS24Q.ShowDialog()
        End If
        'frmTDS24Q.Text = "Form No 24Q AY:" & frmLogin.finacialyr & "-For Quarter 2"
        'frmTDS24Q.Show()
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        Dim RetnID As Long
        'Dim LoadForm As Boolean
        'Dim rst As New ADODB.Recordset
        RetnID = GetForm24RetnID("3")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Oct/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("31/Dec/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ConvertWhich = "24Q3"
        frmTDS24Q.Tag = RetnID
        frmTDS24Q.tabMain.SelectedIndex = 0
        frmTDS24Q.quter = "24Q3"
        'frmTDS24Q.Show      'just to give a visual effect of loading...
        If CheckPRNNo(frmTDS24Q.quter) = False Then
            Call Load24QData(RetnID)
            frmTDS24Q.ShowDialog()
        End If
        'frmTDS24Q.Text = "Form No 24Q AY:" & frmLogin.finacialyr & "-For Quarter 3"
        'frmTDS24Q.Show()
    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        Dim RetnID As Long
        'Dim LoadForm As Boolean
        'Dim rst As New ADODB.Recordset
        RetnID = GetForm24RetnID("4")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Jan/" & Strings.Right(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("31/Mar/" & Strings.Right(FY, 4), "dd/MMM/yyyy")
        ConvertWhich = "24Q4"
        frmTDS24Q.Tag = RetnID
        frmTDS24Q.tabMain.SelectedIndex = 0
        frmTDS24Q.quter = "24Q4"
        'frmTDS24Q.Show      'just to give a visual effect of loading...
        If CheckPRNNo(frmTDS24Q.quter) = False Then
            Call Load24QData(RetnID)
            frmTDS24Q.ShowDialog()
        End If
    End Sub

        Private Sub Load24QData(rID As Long)
        Dim obj As New clsDeductee24QObj
        Dim i As Long
        Dim nds As New DataSet
        With frmTDS24Q
            .lvwChallan.Items.Clear()
            .lvwDeductee.Items.Clear()
            .cmdNext.Enabled = False
            'Check quarter and form 16 usage check and then enable/disable Salary Detail tab accordingly.
            'If .quter = "24Q4" Then
            '    .tabMain.TabEnabled(5) = True
            '    If frmCoMst.chkUseForm16.Value = vbChecked Then
            '        .tabMain.TabEnabled(3) = False
            '        .tabMain.TabEnabled(4) = True
            '    Else
            '        .tabMain.TabEnabled(3) = True
            '        .tabMain.TabEnabled(4) = False
            '    End If
            'Else
            '    .tabMain.TabEnabled(5) = False
            '    .tabMain.TabEnabled(3) = False
            '    .tabMain.TabEnabled(4) = False
            'End If
            'Filling Challan in Deductee Detail
            Dim sql As String
            sql = " SELECT challanid,BankChallanNo,DtOfChallan" _
        & " FROM Challan24Q WHERE (BankChallanNo<>Null and BankChallanNo<>0) " _
        & " UNION ALL SELECT challanid,TranVouNo,DtOfChallan" _
        & " FROM Challan24Q WHERE (TranVouNo<>Null and TranVouNo<>0) AND RetnID=" & rID & " order by ChallanID"
            nds = FetchDataSet(sql)

            .cboChallanNo.Items.Clear()
            Dim dt As Date
            For i = 0 To nds.Tables(0).Rows.Count - 1
                dt = nds.Tables(0).Rows(i)("DtOfChallan")
                dt.ToString("dd/MM/yy")
                .cboChallanNo.Items.Add(nds.Tables(0).Rows(i)(1) & " - " & dt)
                .cboChallanNo.SelectedValue = nds.Tables(0).Rows(i)("ChallanID")

            Next i


            'Filling BSR Code in Challan Detail
            nds = FetchDataSet("select BankBrCode from BankMst WHERE CoID=" & selectedcoid & " order by BankBrCode")
            .cboBankBrCode.Items.Clear()

            For i = 0 To nds.Tables(0).Rows.Count - 1
                .cboBankBrCode.Items.Add(SetFormat("0000000", nds.Tables(0).Rows(i)("BankBrCode")))
            Next i

            'Load challan Details..

            nds = FetchDataSet("SELECT * FROM Challan24Q WHERE RetnID=" & rID)

            For i = 0 To nds.Tables(0).Rows.Count - 1

                Dim tot As Double
                tot = IIf(nds.Tables(0).Rows(i)("TaxAmt") = vbNull, vbNullString, nds.Tables(0).Rows(i)("TaxAmt")) _
                    + IIf(nds.Tables(0).Rows(i)("Surcharge") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Surcharge")) _
                    + IIf(nds.Tables(0).Rows(i)("ECess") = vbNull, vbNullString, nds.Tables(0).Rows(i)("ECess")) _
                    + IIf(nds.Tables(0).Rows(i)("Interest") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Interest")) _
                    + IIf(nds.Tables(0).Rows(i)("Others") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Others")) _
                    + IIf(nds.Tables(0).Rows(i)("AFees") = vbNull, vbNullString, nds.Tables(0).Rows(i)("AFees"))


                'checking sec
                If nds.Tables(0).Rows(i)("Sec").ToString() <> "" Then
                    SectionChecked(nds.Tables(0).Rows(i)("Sec").ToString())
                End If


                Dim newitem As New ListViewItem()
                newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() 'first column

                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString()) 'second column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Interest").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Others").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AFees").ToString())
                newitem.SubItems.Add(tot)
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChqDDNo").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("IsBookEntry").ToString())
                If nds.Tables(0).Rows(i)("BankChallanNo").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else

                    newitem.SubItems.Add((SetFormat("00000", nds.Tables(0).Rows(i)("BankChallanNo").ToString())))
                End If
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TranVouNo").ToString())
                If nds.Tables(0).Rows(i)("BankBrCode").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else

                    newitem.SubItems.Add(SetFormat("0000000", (nds.Tables(0).Rows(i)("BankBrCode").ToString())))
                End If
                dt = nds.Tables(0).Rows(i)("DtOfChallan").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("MinorHead").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AInterest").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AOthers").ToString())
                .lvwChallan.Items.Add(newitem)

            Next
            'Load DEDUCTEE Details..



            nds = FetchDataSet("Select DT.*, DM.* From Deductee24Q As DT " &
                                "INNER Join DeductMst As DM On dt.DId = DM.DId Where RetnID = " & rID & " And CoId = " & selectedcoid & " Order By dt.id24Q ")

            For i = 0 To nds.Tables(0).Rows.Count - 1
                SectionChecked(nds.Tables(0).Rows(i)("Sec").ToString())
                Dim newitem As New ListViewItem()
                newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() 'first column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("DName").ToString()) 'second column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("DPan").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AmtOfPayment").ToString())
                dt = nds.Tables(0).Rows(i)("DtOfPayment").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))
                newitem.SubItems.Add("")
                newitem.SubItems.Add("")
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeducted").ToString())
                If nds.Tables(0).Rows(i)("DtOfDeduction").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else

                    dt = nds.Tables(0).Rows(i)("DtOfDeduction").ToString()

                    newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))
                End If
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeposited").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString())
                Dim chlno As String
                chlno = obj.getChallanNo(Val(nds.Tables(0).Rows(i)("ChallanID")))
                newitem.SubItems.Add(chlno)
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("CertNo").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ID24Q").ToString())
                .lvwDeductee.Items.Add(newitem)

                If nds.Tables(0).Rows(i)("PANVerified") = False Then
                    newitem.SubItems.Item(2).ForeColor = Color.Magenta
                End If


            Next

            ' .txtTDSRate.Enabled = False
        End With

        nds.Dispose()
        'With frmTDS24Q
        '    'Filling Challan in Deductee Detail
        '    .lvwChallan.Items.Clear()
        '    .lvwDeductee.Items.Clear()
        '    Dim sql As String
        '    sql = " SELECT challanid,BankChallanNo,DtOfChallan" _
        '& " FROM Challan24Q WHERE (BankChallanNo<>Null and BankChallanNo<>0) " _
        '& " UNION ALL SELECT challanid,TranVouNo,DtOfChallan " _
        '& " FROM Challan24Q WHERE (TranVouNo<>Null and TranVouNo<>0) AND RetnID=" & rID & " order by ChallanID"
        '    nds = FetchDataSet(sql)

        '    .cboChallanNo.Items.Clear()
        '    Dim dt As Date
        '    For i = 0 To nds.Tables(0).Rows.Count - 1
        '        dt = nds.Tables(0).Rows(i)("DtOfChallan")
        '        dt.ToString("dd/MM/yy")
        '        .cboChallanNo.Items.Add(nds.Tables(0).Rows(i)(1) & " - " & dt)
        '        .cboChallanNo.SelectedValue = nds.Tables(0).Rows(i)("ChallanID")

        '    Next i

        '    'Filling BSR Code in Challan Detail
        '    nds = FetchDataSet("select BankBrCode from BankMst WHERE CoID=" & selectedcoid & " order by BankBrCode")
        '    .cboBankBrCode.Items.Clear()

        '    For i = 0 To nds.Tables(0).Rows.Count - 1
        '        .cboBankBrCode.Items.Add(nds.Tables(0).Rows(i)("BankBrCode"))
        '    Next i

        '    'Load challan Details..

        '    nds = FetchDataSet("SELECT * FROM Challan24Q WHERE RetnID=" & rID)

        '    For i = 0 To nds.Tables(0).Rows.Count - 1
        '        If nds.Tables(0).Rows(i)("Sec").ToString() <> "" Then
        '            SectionChecked24(nds.Tables(0).Rows(i)("Sec").ToString())

        '        End If
        '        Dim tot As Double
        '        tot = IIf(nds.Tables(0).Rows(i)("TaxAmt") = vbNull, vbNullString, nds.Tables(0).Rows(i)("TaxAmt")) + IIf(nds.Tables(0).Rows(i)("Surcharge") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Surcharge")) + IIf(nds.Tables(0).Rows(i)("ECess") = vbNull, vbNullString, nds.Tables(0).Rows(i)("ECess")) + IIf(nds.Tables(0).Rows(i)("Interest") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Interest")) &
        '            +IIf(nds.Tables(0).Rows(i)("Others") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Others")) + IIf(nds.Tables(0).Rows(i)("AFees") = vbNull, vbNullString, nds.Tables(0).Rows(i)("AFees"))

        '        Dim newitem As New ListViewItem()
        '        newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() 'first column
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString()) 'second column
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("Interest").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("Others").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("AFees").ToString())
        '        newitem.SubItems.Add(tot)
        '        'newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChqDDNo").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("IsBookEntry").ToString())
        '        'newitem.SubItems.Add(Format(Val(nds.Tables(0).Rows(i)("BankChallanNo")), "00000"))
        '        newitem.SubItems.Add((SetFormat("00000", nds.Tables(0).Rows(i)("BankChallanNo").ToString())))
        '        'newitem.SubItems.Add(nds.Tables(0).Rows(i)("BankChallanNo").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("TranVouNo").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("BankBrCode").ToString())
        '        dt = nds.Tables(0).Rows(i)("DtOfChallan").ToString()
        '        dt.ToString("dd/mm/yy")
        '        newitem.SubItems.Add(dt)
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("MinorHead").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("AInterest").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("AOthers").ToString())
        '        .lvwChallan.Items.Add(newitem)

        '    Next
        '    nds.Dispose()

        '    'Load DEDUCTEE Details..
        '    nds = FetchDataSet("Select DT.*, DM.* From Deductee24Q As DT " &
        '                        "INNER Join DeductMst As DM On dt.DId = DM.DId Where RetnID = " & rID & " And CoId = " & selectedcoid & " Order By dt.id24Q ")

        '    For i = 0 To nds.Tables(0).Rows.Count - 1
        '        SectionChecked24(nds.Tables(0).Rows(i)("Sec").ToString())
        '        Dim newitem As New ListViewItem()
        '        newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() 'first column
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("DName").ToString()) 'second column
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("DPan").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("AmtOfPayment").ToString())
        '        dt = nds.Tables(0).Rows(i)("DtOfPayment").ToString()
        '        dt.ToString("dd/MMM/yyyy")
        '        newitem.SubItems.Add(dt)
        '        'newitem.SubItems.Add(nds.Tables(0).Rows(i)("IsBookEntry").ToString())
        '        'newitem.SubItems.Add(nds.Tables(0).Rows(i)("RateOfTDS").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeposited").ToString())
        '        dt = nds.Tables(0).Rows(i)("DtOfDeduction").ToString()
        '        dt.ToString("dd/MMM/yyyy")
        '        newitem.SubItems.Add(dt)
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeducted").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString())
        '        Dim chlno As String
        '        chlno = obj.getChallanNo(Val(nds.Tables(0).Rows(i)("ChallanID")))
        '        newitem.SubItems.Add(chlno)
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("CertNo").ToString())
        '        newitem.SubItems.Add(nds.Tables(0).Rows(i)("ID24Q").ToString())
        '        .lvwDeductee.Items.Add(newitem)

        '        If nds.Tables(0).Rows(i)("PANVerified") = False Then
        '            newitem.SubItems.Item(2).ForeColor = Color.Magenta
        '        End If

        '    Next
        '    nds.Dispose()

        '    'Load SALARY DETAILS TAB...

        '    nds = FetchDataSet("SELECT SD.*, DM.* FROM SalaryDetail24Q AS SD " &
        '"INNER JOIN DeductMst AS DM ON SD.DId = DM.DId WHERE RetnID=" & rID)

        '    For i = 0 To nds.Tables(0).Rows.Count - 1
        '        Dim dt1, dt2 As Date
        '        'SectionChecked24(nds.Tables(0).Rows(i)("Sec").ToString())
        '        Dim Itm As New ListViewItem()
        '        'Itm = .lvwSD.ListItems.Add(, , rst!DName)
        '        Itm.Text = nds.Tables(0).Rows(i)("DName").ToString()
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("DPAN").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Category").ToString())
        '        dt1 = nds.Tables(0).Rows(i)("EmpFromDt").ToString()
        '        dt1.ToString("dd/MMM/yyyy")
        '        Itm.SubItems.Add(dt1)
        '        dt2 = nds.Tables(0).Rows(i)("EmpToDt").ToString()
        '        dt2.ToString("dd/MMM/yyyy")
        '        Itm.SubItems.Add(dt2)
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("TotalSalary").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Sec16ii").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Sec16iii").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("TotalSalaryPreEmp").ToString())
        '        Itm.SubItems.Add(Val(Itm.SubItems(5).Text) - Val(Itm.SubItems(6).Text) - Val(Itm.SubItems(7).Text) + Val(Itm.SubItems(8).Text))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("OtherIncome").ToString())
        '        Itm.SubItems.Add(Val(Itm.SubItems(9).Text) + Val(Itm.SubItems(10).Text))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Sec80CCEAmt").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Sec80CCFAmt").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Sec80CCGAmt").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("OtherVIA").ToString())
        '        Itm.SubItems.Add(Val(Itm.SubItems(11).Text) - Val(Itm.SubItems(12).Text) - Val(Itm.SubItems(13).Text) - Val(Itm.SubItems(14).Text) - Val(Itm.SubItems(15).Text))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString())
        '        Itm.SubItems.Add(Val(Itm.SubItems(17).Text) + Val(Itm.SubItems(18).Text) + Val(Itm.SubItems(19).Text))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Relief89").ToString())
        '        Itm.SubItems.Add(Val(Itm.SubItems(20).Text) - Val(Itm.SubItems(21).Text))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("TDSAmt").ToString())
        '        Itm.SubItems.Add(Val(Itm.SubItems(22).Text) - Val(Itm.SubItems(23).Text))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("TDSAmtPreEmp").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("HighRatePAN").ToString())
        '        'Itm.SubItems(8) = Val(Itm.SubItems(5)) - Val(Itm.SubItems(6)) - Val(Itm.SubItems(7)) + Val(Itm.SubItems(25))
        '        '          Itm.SubItems(8) = Val(Itm.SubItems(5)) - Val(Itm.SubItems(6)) - Val(Itm.SubItems(7)) + Val(Itm.SubItems(25))
        '        '          Itm.SubItems(10) = Val(Itm.SubItems(8)) + Val(Itm.SubItems(9))


        '        Itm.SubItems.Add((nds.Tables(0).Rows(i)("TDSAmtPreEmp") + nds.Tables(0).Rows(i)("TDSAmt")))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("SDID").ToString())
        '        .lvwSD.Items.Add(Itm)
        '        'rst.MoveNext
        '    Next

        '    nds.Dispose()

        '    'Load FORM 16 DETAILS TAB...
        '    Dim rstTmp, rstTDSSum As New DataSet
        '    Dim mTotAllw As Double
        '    Dim TotalTDSFromChallan As Double, TotalF16Challan As Double

        '    nds = FetchDataSet("SELECT DeductMst.*, Form16Details.* " &
        '  "FROM DeductMst INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId " &
        '  " WHERE Form16Details.RetnID = " & rID)

        '    '.lvwForm16.ListItems.Clear
        '    For i = 0 To nds.Tables(0).Rows.Count - 1
        '        Dim dt1, dt2 As New Date
        '        Dim Itm As New ListViewItem()
        '        rstTDSSum = FetchDataSet("SELECT sum(D24.TaxAmt) as SumTax, sum(D24.Surcharge) as SumSur, Sum(D24.ECess)as SumECess " &
        '         " FROM Challan24Q AS C24 INNER JOIN Deductee24Q AS D24 ON " &
        '         " C24.ChallanID = D24.ChallanId WHERE D24.DId=" &
        '          nds.Tables(0).Rows(0)("DeductMst.DiD").Value & ";")
        '        Itm.Text = nds.Tables(0).Rows(i)("DName").ToString()
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("DPAN").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Category").ToString())
        '        dt1 = nds.Tables(0).Rows(i)("EmpFromDt").ToString()
        '        dt1.ToString("dd/MMM/yyyy")
        '        Itm.SubItems.Add(dt1)
        '        dt2 = nds.Tables(0).Rows(i)("EmpToDt").ToString()
        '        dt2.ToString("dd/MMM/yyyy")
        '        Itm.SubItems.Add(dt2)
        '        Itm.SubItems.Add((nds.Tables(0).Rows(i)("Gross1").ToString()) + (nds.Tables(0).Rows(i)("Gross2").ToString()) + (nds.Tables(0).Rows(i)("Gross3").ToString()))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Sec16ii").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(i)("Sec16iii").ToString())
        '        'Itm.SubItems(8) = Val(Itm.SubItems(5)) - Val(Itm.SubItems(6)) - Val(Itm.SubItems(7))
        '        'Commented Row
        '        mTotAllw = 0
        '        'If rstTmp.State = adStateOpen Then 
        '        rstTmp.Dispose()
        '        rstTmp = FetchDataSet("SELECT * from Form16MoreDetails WHERE F16ID=" & nds.Tables(0).Rows(0)("F16ID"))
        '        For j = 0 To rstTmp.Tables(0).Rows.Count
        '            If rstTmp.Tables(0).Rows(0)("TypeOfDetail").ToString() = "A" Then
        '                mTotAllw = mTotAllw + CDbl(rstTmp.Tables(0).Rows(0)("GrossAmt").ToString())
        '            ElseIf rstTmp.Tables(0).Rows(0)("TypeOfDetail").ToString() = "O" Then
        '                Itm.SubItems.Add(Val(Itm.SubItems(10)) + CDbl(rstTmp.Tables(0).Rows(0)("GrossAmt").ToString()))
        '            ElseIf rstTmp.Tables(0).Rows(0)("TypeOfDetail").ToString() = "E" Then
        '                Itm.SubItems.Add(Val(Itm.SubItems(12)) + CDbl(rstTmp.Tables(0).Rows(0)("GrossAmt").ToString()))
        '            ElseIf rstTmp.Tables(0).Rows(0)("TypeOfDetail").ToString() = "F" Then
        '                Itm.SubItems.Add(Val(Itm.SubItems(13)) + CDbl(rstTmp.Tables(0).Rows(0)("DeductibleAmt").ToString()))
        '            ElseIf rstTmp.Tables(0).Rows(0)("TypeOfDetail").ToString() = "V" Then
        '                Itm.SubItems.Add(Val(Itm.SubItems(14)) + CDbl(rstTmp.Tables(0).Rows(0)("DeductibleAmt").ToString()))
        '            End If
        '            'rstTmp.MoveNext
        '        Next
        '        Itm.SubItems(6) = IIf(IsNullOrEmpty(mTotAllw), vbNullString, mTotAllw)
        '        TotalF16Challan = 0
        '        'If rstTmp.State = 1 Then 
        '        rstTmp.Dispose()
        '        rstTmp = FetchDataSet("Select sum(F.TaxAmt) as SumTax, sum(F.Surcharge) as SumSur, Sum(F.ECess)as SumECess from F16Challan as F WHERE F.F16ID=" & nds.Tables(0).Rows(0)("F16ID") & "")
        '        Itm.SubItems.Add(Val(Itm.SubItems(5)) - Val(Itm.SubItems(7)) - Val(Itm.SubItems(8)) - mTotAllw)
        '        Itm.SubItems.Add(Val(Itm.SubItems(9)) + Val(Itm.SubItems(10)))
        '        Itm.SubItems.Add(Val(Itm.SubItems(11)) - Val(Itm.SubItems(12)) - Val(Itm.SubItems(13)) - Val(Itm.SubItems(14)))
        '        Itm.SubItems.Add(IsNullOrEmpty(nds.Tables(0).Rows(0)("TaxAmt")))
        '        Itm.SubItems.Add(IsNullOrEmpty(nds.Tables(0).Rows(0)("Surcharge")))
        '        Itm.SubItems.Add(IsNullOrEmpty(nds.Tables(0).Rows(0)("ECess")))
        '        Itm.SubItems.Add(Val(Itm.SubItems(16)) + Val(Itm.SubItems(17)) + Val(Itm.SubItems(18)))
        '        Itm.SubItems.Add(IsNullOrEmpty(nds.Tables(0).Rows(0)("Relief89")))
        '        Itm.SubItems.Add(Val(Itm.SubItems(19)) - Val(Itm.SubItems(20)))

        '        ' changed on: 28/04/2009; by Prakash as Changes in WizinTDS 2009 for TDS Yearly Column;
        '        '         Itm.SubItems(20) = IIf(IsNull(rstTDSSum!SumTax), 0, rstTDSSum!SumTax) + IIf(IsNull(rstTDSSum!SumSur), 0, rstTDSSum!SumSur) + _
        '        '                            IIf(IsNull(rstTDSSum!SumECess), 0, rstTDSSum!SumECess)

        '        TotalTDSFromChallan = IsNullOrEmpty(rstTDSSum.Tables(0).Rows(0)("SumTax")) + (IsNullOrEmpty(rstTDSSum.Tables(0).Rows(0)("SumSur"))) +
        '                       IsNullOrEmpty(rstTDSSum.Tables(0).Rows(0)("SumECess"))
        '        TotalF16Challan = IsNullOrEmpty(rstTDSSum.Tables(0).Rows(0)("SumTax")) + (IsNullOrEmpty(rstTDSSum.Tables(0).Rows(0)("SumSur"))) +
        '                       IsNullOrEmpty(rstTDSSum.Tables(0).Rows(0)("SumECess"))

        '        Itm.SubItems.Add(TotalTDSFromChallan + TotalF16Challan) 'IIf(IsNull(rst!TDSAmt), vbNullString, rst!TDSAmt)
        '        '         Itm.SubItems(23) = Val(Itm.SubItems(21)) - Val(Itm.SubItems(22))
        '        '         Itm.SubItems(24) = rst!F16ID

        '        Itm.SubItems.Add(nds.Tables(0).Rows(0)("F16ID"))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(0)("TotalSalaryPreEmp").ToString())
        '        Itm.SubItems.Add(nds.Tables(0).Rows(0)("TDSAmtPreEmp").ToString())
        '        Itm.SubItems.Add(Val(Itm.SubItems(21)) - (IsNullOrEmpty(nds.Tables(0).Rows(0)("TDSOnPerks")) + Val(Itm.SubItems(22))) - Val(Itm.SubItems(26)))
        '        Itm.SubItems.Add(nds.Tables(0).Rows(0)("HighRatePAN").ToString())
        '        .lvwForm16.Items.Add(Itm)
        '        'rst.MoveNext
        '        'If rstTDSSum.State = adStateOpen Then 
        '        rstTDSSum.Dispose()
        '    Next i
        '    nds.Dispose()
        '    nds = Nothing
        '    'Itm = Nothing
        'End With

    End Sub

    Private Sub ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        frmChlDeDetails.Show()
    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        frmDDetails.Show()
    End Sub

    Private Sub ToolStripMenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        frmChlDeDetails.Show()
    End Sub

    Private Sub ToolStripMenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        frmChlDeDetails.Show()
    End Sub

    Private Sub DeducteeAllocationReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        frmChlDeDetails.Show()
    End Sub

    Private Sub ToolStripMenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        frmDDetails.Show()
    End Sub

    Private Sub ToolStripMenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        frmDDetails.Show()
    End Sub

    Private Sub ChallanAllocationReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        frmDDetails.Show()
    End Sub


    Private Sub ToolStripMenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem9.Click
        Dim RetnID As Long

        RetnID = GetForm27RetnID("1")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Apr/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("30/Jun/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS27Q
        ConvertWhich = "27Q1"
        frmTDS27Q.Tag = RetnID
        frmTDS27Q.tabMain.SelectedIndex = 0
        frmTDS27Q.quter = "27Q1"
        If CheckPRNNo(frmTDS27Q.quter) = False Then
            Load27QData(RetnID)
            frmTDS27Q.ShowDialog()
        End If
    End Sub
    Private Sub Load27QData(rID As Long)

        Dim obj As New clsDeductee27QObj
        Dim i As Long
        Dim nds As New DataSet


        With frmTDS27Q

            .cmdNext.Enabled = False
            'Filling Challan in Deductee Detail
            Dim sql As String
            sql = " SELECT challanid,BankChallanNo,DtOfChallan" _
        & " FROM Challan27Q WHERE (BankChallanNo<>Null and BankChallanNo<>0) " _
        & " UNION ALL SELECT challanid,TranVouNo,DtOfChallan" _
        & " FROM Challan27Q WHERE (TranVouNo<>Null and TranVouNo<>0) AND RetnID=" & rID & " order by ChallanID"
            nds = FetchDataSet(sql)

            .cboChallanNo.Items.Clear()
            Dim dt As Date
            For i = 0 To nds.Tables(0).Rows.Count - 1
                dt = nds.Tables(0).Rows(i)("DtOfChallan")
                dt.ToString("dd/MM/yy")
                .cboChallanNo.Items.Add(nds.Tables(0).Rows(i)(1) & " - " & dt)
                .cboChallanNo.SelectedValue = nds.Tables(0).Rows(i)("ChallanID")

            Next i


            'Filling BSR Code in Challan Detail
            nds = FetchDataSet("select BankBrCode from BankMst WHERE CoID=" & selectedcoid & " order by BankBrCode")
            .cboBankBrCode.Items.Clear()


            For i = 0 To nds.Tables(0).Rows.Count - 1
                .cboBankBrCode.Items.Add(SetFormat("0000000", nds.Tables(0).Rows(i)("BankBrCode")))
            Next i

            'Filling Remit Combo in Deductee Detail
            nds = FetchDataSet("select * from Remit27Q")
            .cboRemit.DataSource = Nothing
            .cboRemit.Items.Clear()
            .cboRemit.DataSource = nds.Tables(0)
            .cboRemit.DisplayMember = "RemitDesc"
            .cboRemit.ValueMember = "RemitCode"


            'Filling Country Combo in Deductee Detail
            nds = FetchDataSet("select * from Country27Q")
            .cboCountry.Items.Clear()
            .cboCountry.DataSource = Nothing
            .cboCountry.DataSource = nds.Tables(0)
            .cboCountry.DisplayMember = "CountryName"
            .cboCountry.ValueMember = "CountryCode"


            'Load challan Details..

            nds = FetchDataSet("SELECT * FROM Challan27Q WHERE RetnID=" & rID)

            For i = 0 To nds.Tables(0).Rows.Count - 1

                Dim tot As Double
                tot = IIf(nds.Tables(0).Rows(i)("TaxAmt") = vbNull, vbNullString, nds.Tables(0).Rows(i)("TaxAmt")) _
                + IIf(nds.Tables(0).Rows(i)("Surcharge") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Surcharge")) _
                + IIf(nds.Tables(0).Rows(i)("ECess") = vbNull, vbNullString, nds.Tables(0).Rows(i)("ECess")) _
                + IIf(nds.Tables(0).Rows(i)("Interest") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Interest")) _
                    + IIf(nds.Tables(0).Rows(i)("Others") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Others")) _
                    + IIf(nds.Tables(0).Rows(i)("AFees") = vbNull, vbNullString, nds.Tables(0).Rows(i)("AFees"))


                'checking sec
                If nds.Tables(0).Rows(i)("Sec").ToString() <> "" Then
                    SectionChecked27(nds.Tables(0).Rows(i)("Sec").ToString())
                End If


                Dim newitem As New ListViewItem()
                newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() 'first column

                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString()) 'second column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Interest").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Others").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AFees").ToString())
                newitem.SubItems.Add(tot)
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChqDDNo").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("IsBookEntry").ToString())
                If nds.Tables(0).Rows(i)("BankChallanNo").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else
                    newitem.SubItems.Add(SetFormat("00000", nds.Tables(0).Rows(i)("BankChallanNo").ToString()))
                End If
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TranVouNo").ToString())
                If nds.Tables(0).Rows(i)("BankBrCode").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else
                    newitem.SubItems.Add(SetFormat("0000000", nds.Tables(0).Rows(i)("BankBrCode").ToString()))
                End If
                dt = nds.Tables(0).Rows(i)("DtOfChallan").ToString()
                'dt.ToString("dd/MMM/yyyy")
                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("MinorHead").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AInterest").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AOthers").ToString())
                .lvwChallan.Items.Add(newitem)

            Next
            'Load DEDUCTEE Details..



            nds = FetchDataSet("Select DT.*, DM.* From Deductee27Q As DT " &
                                "INNER Join DeductMst As DM On dt.DId = DM.DId Where RetnID = " & rID & " And CoId = " & selectedcoid & " Order By dt.id27Q ")


            For i = 0 To nds.Tables(0).Rows.Count - 1
                SectionChecked27(nds.Tables(0).Rows(i)("Sec").ToString())
                Dim newitem As New ListViewItem()
                newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() 'first column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("DName").ToString()) 'second column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("DPan").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AmtOfPayment").ToString())

                dt = nds.Tables(0).Rows(i)("DtOfPayment").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("IsBookEntry").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("RateOfTDS").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeducted").ToString())
                If nds.Tables(0).Rows(i)("DtOfDeduction").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else
                    dt = nds.Tables(0).Rows(i)("DtOfDeduction").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))
                End If
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeposited").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString())
                Dim chlno As String
                chlno = obj.getChallanNo(Val(nds.Tables(0).Rows(i)("ChallanID")))
                newitem.SubItems.Add(chlno)
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("CertNo").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("DTAA").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("RemitID").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("UniqueAck").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("CountryID").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ID27Q").ToString())

                .lvwDeductee.Items.Add(newitem)
                If nds.Tables(0).Rows(i)("PANVerified") = False Then
                    newitem.SubItems.Item(2).ForeColor = Color.Magenta
                End If


            Next

            .txtTDSRate.Enabled = False
        End With

        nds.Dispose()

        '        Dim obj As New clsDeductee27QObj
        '        Dim rst As New ADODB.Recordset, Itm As ListItem, i As Long
        '        With frmTDS27Q

        '            'Filling Challan in Deductee Detail
        '            Dim sql As String
        '            sql = " SELECT challanid,BankChallanNo,DtOfChallan" _
        '        & " FROM Challan27Q WHERE (BankChallanNo<>Null and BankChallanNo<>0) " _
        '        & " UNION ALL SELECT challanid,TranVouNo,DtOfChallan " _
        '        & " FROM Challan27Q WHERE (TranVouNo<>Null and TranVouNo<>0) "

        '            rst.Open sql & " AND RetnID=" & rID & " order by ChallanID", Cnn, adOpenStatic, adLockOptimistic
        '        .cboChallanNo.Clear
        '            While Not (rst.EOF Or rst.BOF)
        '                .cboChallanNo.AddItem rst(1) & " - " & Format(rst!DtOfChallan, "dd/mm/yy")
        '            .cboChallanNo.ItemData(.cboChallanNo.NewIndex) = rst!ChallanID
        '                rst.MoveNext
        '        Wend
        '        If rst.State = adStateOpen Then rst.Close

        '            'Filling BSR Code in Challan Detail
        '            rst.Open "select BankBrCode from BankMst WHERE CoID=" & selectedcoid & " order by BankBrCode", Cnn, adOpenStatic, adLockOptimistic
        '        .cboBankBrCode.Clear
        '            While Not (rst.EOF Or rst.BOF)
        '                .cboBankBrCode.AddItem Format(rst!BankBrCode, "0000000")
        '        rst.MoveNext
        '        Wend
        '        If rst.State = adStateOpen Then rst.Close
        '            'Filling Remit Combo in Deductee Detail
        '            rst.Open "select * from Remit27Q", Cnn, adOpenStatic, adLockOptimistic
        '        .cboRemit.Clear
        '            While Not (rst.EOF Or rst.BOF)
        '                .cboRemit.AddItem rst!RemitDesc
        '            .cboRemit.ItemData(.cboRemit.NewIndex) = rst!RemitCode
        '                rst.MoveNext
        '        Wend
        '        If rst.State = adStateOpen Then rst.Close
        '            'Filling Country Combo in Deductee Detail
        '            rst.Open "select * from Country27Q", Cnn, adOpenStatic, adLockOptimistic
        '        .cboCountry.Clear
        '            While Not (rst.EOF Or rst.BOF)
        '                .cboCountry.AddItem rst!CountryName
        '            .cboCountry.ItemData(.cboCountry.NewIndex) = rst!CountryCode
        '                rst.MoveNext
        '        Wend
        '        If rst.State = adStateOpen Then rst.Close


        '            'Load challan Details..
        '            rst.Open "SELECT * FROM Challan27Q WHERE RetnID=" & rID, Cnn, adOpenKeyset, adLockReadOnly
        '      Do While Not rst.EOF
        '         Set Itm = .lvwChallan.ListItems.Add(, , rst!Sec & "")
        '         For i = 0 To .chkSection.UBound
        '                    If .chkSection(i).Caption = rst!Sec Then
        '                        .chkSection(i).Value = vbChecked
        '                        Exit For
        '                    End If
        '                Next
        '                Itm.SubItems(1) = IIf(IsNull(rst!TaxAmt), vbNullString, rst!TaxAmt)
        '                Itm.SubItems(2) = IIf(IsNull(rst!Surcharge), vbNullString, rst!Surcharge)
        '                Itm.SubItems(3) = IIf(IsNull(rst!ECess), vbNullString, rst!ECess)
        '                Itm.SubItems(4) = IIf(IsNull(rst!Interest), vbNullString, rst!Interest)
        '                Itm.SubItems(5) = IIf(IsNull(rst!Others), vbNullString, rst!Others)
        '                Itm.SubItems(6) = IIf(IsNull(rst!AFees), vbNullString, rst!AFees)
        '                Itm.SubItems(7) = Val(Itm.SubItems(1)) + Val(Itm.SubItems(2)) + Val(Itm.SubItems(3)) + Val(Itm.SubItems(4)) + Val(Itm.SubItems(5) + Val(Itm.SubItems(6)))
        '                Itm.SubItems(8) = IIf(IsNull(rst!ChqDDNo), vbNullString, rst!ChqDDNo)
        '                Itm.SubItems(9) = rst!IsBookEntry
        '                Itm.SubItems(10) = Format(rst!BankChallanNo, "00000")
        '                Itm.SubItems(11) = IIf(IsNull(rst!TranVouNo), vbNullString, rst!TranVouNo)
        '                Itm.SubItems(12) = Format(rst!BankBrCode, "0000000")
        '                Itm.SubItems(13) = Format(rst!DtOfChallan, "dd/MMM/yyyy")
        '                Itm.SubItems(14) = IIf(IsNull(rst!Remark), vbNullString, rst!Remark)
        '                Itm.SubItems(15) = IIf(IsNull(rst!MinorHead), vbNullString, rst!MinorHead)
        '                Itm.SubItems(16) = rst!ChallanID
        '                Itm.SubItems(17) = IIf(IsNull(rst!AInterest), vbNullString, rst!AInterest)
        '                Itm.SubItems(18) = IIf(IsNull(rst!AOthers), vbNullString, rst!AOthers)
        '                rst.MoveNext
        '            Loop
        '            'Load DEDUCTEE Details..
        '            If rst.State = adStateOpen Then rst.Close
        '            rst.Open "SELECT DT.*, DM.* FROM Deductee27Q AS DT " &
        '      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" & rID & " AND CoId=" & selectedcoid & " order by dt.id27Q ", Cnn, adOpenKeyset, adLockReadOnly  'Deductee Records
        '            Do While Not rst.EOF
        '         Set Itm = .lvwDeductee.ListItems.Add(, , rst!Sec)
        '         For i = 0 To .chkSection.UBound
        '                    If .chkSection(i).Caption = rst!Sec Then
        '                        .chkSection(i).Value = vbChecked
        '                        Exit For
        '                    End If
        '                Next
        '                Itm.SubItems(1) = rst!DName & ""
        '                Itm.SubItems(2) = rst!DPan & ""
        '                Itm.SubItems(3) = rst!AmtOfPayment
        '                Itm.SubItems(4) = Format(rst!DtOfPayment, "dd/MMM/yyyy")
        '                Itm.SubItems(5) = rst!IsBookEntry
        '                Itm.SubItems(6) = rst!RateOfTDS
        '                Itm.SubItems(7) = rst!TaxAmt
        '                Itm.SubItems(8) = rst!Surcharge
        '                Itm.SubItems(9) = rst!ECess
        '                Itm.SubItems(10) = rst!TotalTaxDeducted
        '                Itm.SubItems(11) = Format(rst!DtOfDeduction, "dd/MMM/yyyy")
        '                Itm.SubItems(12) = rst!TotalTaxDeposited
        '                Itm.SubItems(13) = IIf(IsNull(rst!ChallanID), vbNullString, rst!ChallanID)
        '                Itm.SubItems(14) = obj.getChallanNo(Val(Itm.SubItems(13)))
        '                Itm.SubItems(15) = IIf(IsNull(rst!Remark), vbNullString, rst!Remark)
        '                Itm.SubItems(16) = rst!CertNo & ""
        '                Itm.SubItems(17) = rst!DTAA & ""
        '                Itm.SubItems(18) = rst!RemitID & ""
        '                Itm.SubItems(19) = rst!UniqueAck & ""
        '                Itm.SubItems(20) = rst!CountryID & ""
        '                Itm.SubItems(21) = rst!ID27Q

        '          Set subItm = Itm.ListSubItems(2)
        '         If rst!PANVerified = False Then
        '                    subItm.ForeColor = vbMagenta
        '                End If

        '                rst.MoveNext
        '            Loop
        '            If rst.State = adStateOpen Then rst.Close
        '            .txtTDSRate.Enabled = False
        '        End With


        'Set rst = Nothing
        'Set Itm = Nothing
    End Sub
    Private Function GetForm27RetnID(Qtr As String) As Long
        On Error GoTo canerr
        Dim oRetnMst As New ClsRetnMstObj, nds As New DataSet
        nds = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " and FrmType= '27Q" & Qtr & "'")
        If nds.Tables(0).Rows.Count <= 0 Then
            If MsgBox("Form No 27Q for Quarter " & Qtr & " for this company does not exist" & vbCrLf & "Do you want to create it?", vbYesNo + vbQuestion, "Create Form No 27Q") = vbYes Then
                oRetnMst.coid = selectedcoid
                oRetnMst.AYear = AY
                oRetnMst.FrmType = "27Q" & Qtr
                If oRetnMst.Insert(oRetnMst) = False Then
                    MsgBox("Unable to create Form No. 27Q, Call JAK Infosolutions", vbCritical, "FORM NOT CREATED")
                    GoTo cleanup
                End If
            Else
                GoTo cleanup
            End If
        End If

        nds = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " and FrmType='27Q" & Qtr & "'")
        GetForm27RetnID = nds.Tables(0).Rows(0)("RetnID")

cleanup:
        'Me.MousePointer = vbDefault
        nds.Dispose()
        oRetnMst = Nothing
        Exit Function

canerr:
        If Err.Number <> 32755 Then
            MsgBox(Err.Description, , Err.Number)
        End If
        GoTo cleanup

    End Function
    Private Sub ToolStripMenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click
        Dim RetnID As Long

        RetnID = GetForm27RetnID("2")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/July/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("30/Sep/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS27Q
        ConvertWhich = "27Q2"
        frmTDS27Q.Tag = RetnID
        frmTDS27Q.tabMain.SelectedIndex = 0
        frmTDS27Q.quter = "27Q2"
        If CheckPRNNo(frmTDS27Q.quter) = False Then
            Load27QData(RetnID)
            frmTDS27Q.ShowDialog()
        End If
    End Sub

    Private Sub ToolStripMenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem11.Click
        Dim RetnID As Long

        RetnID = GetForm27RetnID("3")
        If RetnID = 0 Then Exit Sub

        FromDateQ = Format("01/Oct/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("31/Dec/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS27Q
        ConvertWhich = "27Q3"
        frmTDS27Q.Tag = RetnID
        frmTDS27Q.tabMain.SelectedIndex = 0
        frmTDS27Q.quter = "27Q3"
        If CheckPRNNo(frmTDS27Q.quter) = False Then
            Load27QData(RetnID)
            frmTDS27Q.ShowDialog()
        End If
    End Sub

    Private Sub ToolStripMenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem12.Click
        Dim RetnID As Long

        RetnID = GetForm27RetnID("4")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Jan/" & Strings.Right(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("31/Mar/" & Strings.Right(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS27Q
        ConvertWhich = "27Q4"
        frmTDS27Q.Tag = RetnID
        frmTDS27Q.tabMain.SelectedIndex = 0
        frmTDS27Q.quter = "27Q4"
        If CheckPRNNo(frmTDS27Q.quter) = False Then
            Load27QData(RetnID)
            frmTDS27Q.ShowDialog()
        End If
    End Sub

    Private Sub ToolStripMenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem16.Click
        Dim RetnID As Long
        'Dim LoadForm As Boolean
        'Dim rst As New ADODB.Recordset
        RetnID = GetForm26RetnID("1")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Apr/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("30/Jun/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        'frmTDS26Q.Show()
        ConvertWhich = "26Q1"
        frmTDS26Q.Tag = RetnID
        frmTDS26Q.tabMain.SelectedIndex = 0
        frmTDS26Q.quter = "26Q1"

        If CheckPRNNo(frmTDS26Q.quter) = False Then
            'lvwchallanHead()

            Call Load26QData(RetnID)

            frmTDS26Q.ShowDialog()
        End If

    End Sub

    Private Sub Load26QData(rID As Long)
        Dim obj As New clsDeductee26QObj
        Dim i As Long
        Dim nds As New DataSet


        With frmTDS26Q
            .lvwChallan.Items.Clear()
            .lvwDeductee.Items.Clear()
            .cmdNext.Enabled = False
            'Filling Challan in Deductee Detail
            Dim sql As String
            sql = " SELECT challanid,BankChallanNo,DtOfChallan" _
        & " FROM Challan26Q WHERE (BankChallanNo<>Null and BankChallanNo<>0) " _
        & " UNION ALL SELECT challanid,TranVouNo,DtOfChallan" _
        & " FROM Challan26Q WHERE (TranVouNo<>Null and TranVouNo<>0) AND RetnID=" & rID & " order by ChallanID"
            nds = FetchDataSet(sql)

            .cboChallanNo.Items.Clear()
            Dim dt As Date
            For i = 0 To nds.Tables(0).Rows.Count - 1
                dt = nds.Tables(0).Rows(i)("DtOfChallan")
                dt.ToString("dd/MM/yy")
                .cboChallanNo.Items.Add(nds.Tables(0).Rows(i)(1) & " - " & dt)
                .cboChallanNo.SelectedValue = nds.Tables(0).Rows(i)("ChallanID")

            Next i


            'Filling BSR Code in Challan Detail
            nds = FetchDataSet("select BankBrCode from BankMst WHERE CoID=" & selectedcoid & " order by BankBrCode")
            .cboBankBrCode.Items.Clear()

            For i = 0 To nds.Tables(0).Rows.Count - 1
                .cboBankBrCode.Items.Add(SetFormat("0000000", nds.Tables(0).Rows(i)("BankBrCode")))
            Next i

            'Load challan Details..

            nds = FetchDataSet("SELECT * FROM Challan26Q WHERE RetnID=" & rID)

            For i = 0 To nds.Tables(0).Rows.Count - 1

                Dim tot As Double
                tot = IIf(nds.Tables(0).Rows(i)("TaxAmt") = vbNull, vbNullString, nds.Tables(0).Rows(i)("TaxAmt")) _
                    + IIf(nds.Tables(0).Rows(i)("Surcharge") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Surcharge")) _
                    + IIf(nds.Tables(0).Rows(i)("ECess") = vbNull, vbNullString, nds.Tables(0).Rows(i)("ECess")) _
                    + IIf(nds.Tables(0).Rows(i)("Interest") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Interest")) _
                    + IIf(nds.Tables(0).Rows(i)("Others") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Others")) _
                    + IIf(nds.Tables(0).Rows(i)("AFees") = vbNull, vbNullString, nds.Tables(0).Rows(i)("AFees"))


                'checking sec
                If nds.Tables(0).Rows(i)("Sec").ToString() <> "" Then
                    SectionChecked(nds.Tables(0).Rows(i)("Sec").ToString())
                End If


                Dim newitem As New ListViewItem()
                newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() 'first column

                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString()) 'second column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Interest").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Others").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AFees").ToString())
                newitem.SubItems.Add(tot)
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChqDDNo").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("IsBookEntry").ToString())
                If nds.Tables(0).Rows(i)("BankChallanNo").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else

                    newitem.SubItems.Add((SetFormat("00000", nds.Tables(0).Rows(i)("BankChallanNo").ToString())))
                End If
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TranVouNo").ToString())
                If nds.Tables(0).Rows(i)("BankBrCode").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else

                    newitem.SubItems.Add(SetFormat("0000000", (nds.Tables(0).Rows(i)("BankBrCode").ToString())))
                End If
                dt = nds.Tables(0).Rows(i)("DtOfChallan").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("MinorHead").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AInterest").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AOthers").ToString())
                .lvwChallan.Items.Add(newitem)

            Next
            'Load DEDUCTEE Details..



            nds = FetchDataSet("Select DT.*, DM.* From Deductee26Q As DT " &
                                "INNER Join DeductMst As DM On dt.DId = DM.DId Where RetnID = " & rID & " And CoId = " & selectedcoid & " Order By dt.id26Q ")

            For i = 0 To nds.Tables(0).Rows.Count - 1
                SectionChecked(nds.Tables(0).Rows(i)("Sec").ToString())
                Dim newitem As New ListViewItem()
                newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() 'first column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("DName").ToString()) 'second column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("DPan").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AmtOfPayment").ToString())
                dt = nds.Tables(0).Rows(i)("DtOfPayment").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("IsBookEntry").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("RateOfTDS").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeducted").ToString())
                If nds.Tables(0).Rows(i)("DtOfDeduction").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else

                    dt = nds.Tables(0).Rows(i)("DtOfDeduction").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))
                End If
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeposited").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString())
                Dim chlno As String
                chlno = obj.getChallanNo(Val(nds.Tables(0).Rows(i)("ChallanID")))
                newitem.SubItems.Add(chlno)
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("CertNo").ToString())
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ID26Q").ToString())
                .lvwDeductee.Items.Add(newitem)


                If nds.Tables(0).Rows(i)("Sec").ToString() = "194A" Or nds.Tables(0).Rows(i)("Sec").ToString() = "194J" Or nds.Tables(0).Rows(i)("Sec").ToString() = "194Ib" Then

                    If Val(nds.Tables(0).Rows(i)("RateOfTDS")) < 9.5 Then
                        For Each subItm In newitem.SubItems
                            'Override of color Vbmagenta to red
                            subItm.ForeColor = Color.Red        'color the sub items also

                        Next
                        '                Itm.ForeColor = vbRed
                    End If
                End If
                If nds.Tables(0).Rows(i)("Sec").ToString() = "194Ia" Then

                    If Val(nds.Tables(0).Rows(i)("RateOfTDS")) < 1.8 Then
                        For Each subItm In newitem.SubItems
                            subItm.ForeColor = Color.Red        'color the sub items also
                        Next
                    End If
                End If
                If nds.Tables(0).Rows(i)("Sec").ToString() = "194C" Then
                    If Strings.Right(Strings.Left(nds.Tables(0).Rows(i)("DPan").ToString(), 4), 1) = "P" Or Strings.Right(Strings.Left(nds.Tables(0).Rows(i)("DPan").ToString(), 4), 1) = "H" Then
                        If Val(nds.Tables(0).Rows(i)("RateOfTDS")) < 0.9 Then
                            For Each subItm In newitem.SubItems
                                subItm.ForeColor = Color.Red         'color the sub items also
                            Next
                        End If
                    Else
                        If Val(nds.Tables(0).Rows(i)("RateOfTDS")) < 1.9 Then
                            For Each subItm In newitem.SubItems
                                subItm.ForeColor = Color.Red        'color the sub items also
                            Next
                        End If
                    End If
                End If

                If nds.Tables(0).Rows(i)("Sec").ToString() = "194H" Then

                    If Val(nds.Tables(0).Rows(i)("RateOfTDS")) < 4 Then
                        For Each subItm In newitem.SubItems
                            subItm.ForeColor = Color.Red        'color the sub items also
                        Next
                    End If
                End If
                'subItm = newitem.SubItems.Item(2)
                'subItm = newitem.ListSubItems(2)
                If nds.Tables(0).Rows(i)("PANVerified") = False Then
                    newitem.SubItems.Item(2).ForeColor = Color.Magenta
                End If


            Next

            .txtTDSRate.Enabled = False
        End With

        nds.Dispose()
    End Sub


    Private Sub Load27EQData(rID As Long)
        Dim obj As New clsDeductee27EQObj
        Dim i As Long
        Dim nds As New DataSet


        With frmTDS27EQ

            .cmdNext.Enabled = False
            'Filling Challan in Deductee Detail
            Dim sql As String
            sql = " SELECT challanid,BankChallanNo,DtOfChallan" _
        & " FROM Challan27EQ WHERE (BankChallanNo<>Null and BankChallanNo<>0) " _
        & " UNION ALL SELECT challanid,TranVouNo,DtOfChallan" _
        & " FROM Challan27EQ WHERE (TranVouNo<>Null and TranVouNo<>0) AND RetnID=" & rID & " order by ChallanID"
            nds = FetchDataSet(sql)

            .cboChallanNo.Items.Clear()
            Dim dt As Date
            For i = 0 To nds.Tables(0).Rows.Count - 1
                dt = nds.Tables(0).Rows(i)("DtOfChallan")

                .cboChallanNo.Items.Add(nds.Tables(0).Rows(i)(1) & " - " & dt.ToString("dd/MM/yy"))
                .cboChallanNo.SelectedValue = nds.Tables(0).Rows(i)("ChallanID")

            Next i


            'Filling BSR Code in Challan Detail
            nds = FetchDataSet("select BankBrCode from BankMst WHERE CoID=" & selectedcoid & " order by BankBrCode")
            .cboBankBrCode.Items.Clear()

            For i = 0 To nds.Tables(0).Rows.Count - 1
                .cboBankBrCode.Items.Add(SetFormat("0000000", nds.Tables(0).Rows(i)("BankBrCode")))
            Next i

            'Load challan Details..

            nds = FetchDataSet("SELECT * FROM Challan27EQ WHERE RetnID=" & rID)

            For i = 0 To nds.Tables(0).Rows.Count - 1

                Dim tot As Double
                tot = IIf(nds.Tables(0).Rows(i)("TaxAmt") = vbNull, vbNullString, nds.Tables(0).Rows(i)("TaxAmt")) _
                    + IIf(nds.Tables(0).Rows(i)("Surcharge") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Surcharge")) _
                    + IIf(nds.Tables(0).Rows(i)("ECess") = vbNull, vbNullString, nds.Tables(0).Rows(i)("ECess")) _
                    + IIf(nds.Tables(0).Rows(i)("Interest") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Interest")) _
                    + IIf(nds.Tables(0).Rows(i)("Others") = vbNull, vbNullString, nds.Tables(0).Rows(i)("Others")) _
                    + IIf(nds.Tables(0).Rows(i)("AFees") = vbNull, vbNullString, nds.Tables(0).Rows(i)("AFees"))


                'checking sec
                'If nds.Tables(0).Rows(i)("Sec").ToString() <> "" Then
                '    SectionChecked27E(nds.Tables(0).Rows(i)("Sec").ToString())
                'End If
                .cmdNext.Enabled = True
                Dim newitem As New ListViewItem()
                newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() '0

                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString()) '1
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString()) '2
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString()) '3
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Interest").ToString()) '4
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Others").ToString()) '5
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AFees").ToString()) '6
                newitem.SubItems.Add(tot) '7
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChqDDNo").ToString()) '8
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("IsBookEntry").ToString()) '9

                If nds.Tables(0).Rows(i)("BankChallanNo").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else

                    newitem.SubItems.Add((SetFormat("00000", nds.Tables(0).Rows(i)("BankChallanNo").ToString())))
                End If
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TranVouNo").ToString())
                If nds.Tables(0).Rows(i)("BankBrCode").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else

                    newitem.SubItems.Add(SetFormat("0000000", (nds.Tables(0).Rows(i)("BankBrCode").ToString())))
                End If
                dt = nds.Tables(0).Rows(i)("DtOfChallan").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy"))


                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString()) '14
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString()) '15
                newitem.SubItems.Add(Strings.Left(nds.Tables(0).Rows(i)("Sec").ToString(), 1)) '16
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("MinorHead").ToString()) '17
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AInterest").ToString()) '18
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AOthers").ToString()) '19
                .lvwChallan.Items.Add(newitem)

            Next
            'Load DEDUCTEE Details..



            nds = FetchDataSet("Select DT.*, DM.* From Deductee27EQ As DT " &
                                "INNER Join DeductMst As DM On dt.DId = DM.DId Where RetnID = " & rID & " And CoId = " & selectedcoid & " Order By dt.id27EQ ")

            For i = 0 To nds.Tables(0).Rows.Count - 1
                'SectionChecked27E(nds.Tables(0).Rows(i)("Sec").ToString())
                Dim newitem As New ListViewItem()

                newitem.Text = nds.Tables(0).Rows(i)("Sec").ToString() '0
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("DName").ToString()) '1
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("DPan").ToString()) '2
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Purchamt").ToString()) '3
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("AmtOfPayment").ToString()) '4
                dt = nds.Tables(0).Rows(i)("DtOfPayment").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy")) '5
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("RateOfTDS").ToString()) '6
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TaxAmt").ToString()) '7
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Surcharge").ToString()) '8
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ECess").ToString()) '9
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeducted").ToString()) '10
                If nds.Tables(0).Rows(i)("DtOfDeduction").ToString() = "" Then
                    newitem.SubItems.Add("")
                Else
                    dt = nds.Tables(0).Rows(i)("DtOfDeduction").ToString()

                newitem.SubItems.Add(dt.ToString("dd/MMM/yyyy")) '11
                End If
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("TotalTaxDeposited").ToString()) '12
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ChallanID").ToString()) '13
                Dim chlno As String
                chlno = obj.getChallanNo(Val(nds.Tables(0).Rows(i)("ChallanID")))
                newitem.SubItems.Add(chlno) '14
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("Remark").ToString()) '15
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("CertNo").ToString()) '16
                newitem.SubItems.Add(nds.Tables(0).Rows(i)("ID27EQ").ToString()) '17
                .lvwDeductee.Items.Add(newitem)

                If nds.Tables(0).Rows(i)("PANVerified") = False Then
                    newitem.SubItems.Item(2).ForeColor = Color.Magenta
                End If


            Next

            .txtTDSRate.Enabled = False
        End With

        nds.Dispose()
    End Sub

    'Private Sub lvwchallanHead()
    '    With frmTDS26Q.lvwChallan
    '        .Columns.Clear()
    '        .Columns.Add("Section", 50, HorizontalAlignment.Left)
    '        .Columns.Add("Amount Deducted", 50, HorizontalAlignment.Right)
    '        .Columns.Add("Surcharge", 50, HorizontalAlignment.Right)
    '        .Columns.Add("Edu. Cess", 50, HorizontalAlignment.Right)
    '        .Columns.Add("Interest", 50, HorizontalAlignment.Right)
    '        .Columns.Add("Others", 50, HorizontalAlignment.Right)
    '        .Columns.Add("Fees", 50, HorizontalAlignment.Right)
    '        .Columns.Add("Total Tax", 50, HorizontalAlignment.Right)
    '        .Columns.Add("Chq./DD No.", 70, HorizontalAlignment.Left)
    '        .Columns.Add("By book entry", 70, HorizontalAlignment.Left)
    '        .Columns.Add("Challan No.", 70, HorizontalAlignment.Left)
    '        .Columns.Add("Tran. Vou. No. ", 70, HorizontalAlignment.Left)
    '        .Columns.Add("BSR code", 70, HorizontalAlignment.Left)
    '        .Columns.Add("Dt. of Challan", 70, HorizontalAlignment.Left)
    '        .Columns.Add("Remark", 70, HorizontalAlignment.Left)
    '        .Columns.Add("MinorHead", 70, HorizontalAlignment.Left)
    '        .Columns.Add("ChallanID", 50, HorizontalAlignment.Left)
    '        .Columns.Add("Allocated Interest", 50, HorizontalAlignment.Right)
    '        .Columns.Add("Allocated Other Amt", 50, HorizontalAlignment.Right)
    '        'Display listview in details view
    '        .View = View.Details
    '        'display grid lines
    '        .GridLines = True
    '        'allow full row selection
    '        .FullRowSelect = True
    '    End With
    'End Sub
    Private Sub SectionChecked(sect As String)

        With frmTDS26Q
            For Each chk As CheckBox In .PanelCheckBox.Controls
                If chk.Text = sect Then
                    chk.Checked = True
                    .cmdNext.Enabled = True
                    Exit Sub
                End If
            Next

        End With
    End Sub
    Private Sub SectionChecked27(sect As String)

        With frmTDS27Q
            For Each chk As CheckBox In .PanelCheckBox.Controls
                If chk.Text = sect Then
                    chk.Checked = True
                    .cmdNext.Enabled = True
                    Exit Sub
                End If
            Next

        End With
    End Sub
    Private Sub SectionChecked27E(sect As String)

        With frmTDS27EQ
            For Each chk As CheckBox In .PanelCheckBox.Controls
                If chk.Text = sect Then
                    chk.Checked = True
                    .cmdNext.Enabled = True
                    Exit Sub
                End If
            Next

        End With
    End Sub
    Private Function GetForm26RetnID(Qtr As String) As Long
        On Error GoTo canerr
        Dim nds As New DataSet

        Dim oRetnMst As New ClsRetnMstObj

        nds = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " And FrmType= '26Q" & Qtr & "'")
        If nds.Tables(0).Rows.Count <= 0 Then
            If MsgBox("Form No 26Q for Quarter " & Qtr & " for this company does not exist" & vbCrLf & "Do you want to create it?", vbYesNo + vbQuestion, "Create Form No 26Q") = vbYes Then
                oRetnMst.coid = selectedcoid
                oRetnMst.AYear = AY
                oRetnMst.FrmType = "26Q" & Qtr
                If oRetnMst.Insert(oRetnMst) = False Then
                    MsgBox("Unable to create Form No. 26Q, Call JAK Infosolutions", vbCritical, "FORM NOT CREATED")
                    GoTo cleanup
                End If
            Else
                GoTo cleanup
            End If
        End If
        nds = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " and FrmType='26Q" & Qtr & "'")

        GetForm26RetnID = nds.Tables(0).Rows(0)("RetnID")

cleanup:
        'Me.MousePointer = vbDefault
        nds.Dispose()

        oRetnMst = Nothing
        Exit Function

canerr:
        If Err.Number <> 32755 Then
            MsgBox(Err.Description, , Err.Number)
        End If
        GoTo cleanup

    End Function
    Private Function GetForm27ERetnID(Qtr As String) As Long
        On Error GoTo canerr
        Dim oRetnMst As New ClsRetnMstObj
        Dim nds As New DataSet
        nds = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " and FrmType= '27EQ" & Qtr & "'")
        If nds.Tables(0).Rows.Count <= 0 Then
            If MsgBox("Form No 27EQ for Quarter " & Qtr & " for this company does not exist" & vbCrLf & "Do you want to create it?", vbYesNo + vbQuestion, "Create Form No 27EQ") = vbYes Then
                oRetnMst.coid = selectedcoid
                oRetnMst.AYear = AY
                oRetnMst.FrmType = "27EQ" & Qtr
                If oRetnMst.Insert(oRetnMst) = False Then
                    MsgBox("Unable to create Form No. 27EQ, Call JAK Infosolutions", vbCritical, "FORM NOT CREATED")
                    GoTo cleanup
                End If
            Else
                GoTo cleanup
            End If
        End If

        nds = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " and FrmType='27EQ" & Qtr & "'")
        GetForm27ERetnID = nds.Tables(0).Rows(0)("RetnID")

cleanup:
        'Me.MousePointer = vbDefault
        nds.Dispose()
        oRetnMst = Nothing
        Exit Function

canerr:
        If Err.Number <> 32755 Then
            MsgBox(Err.Description, , Err.Number)
        End If
        GoTo cleanup

    End Function
    Private Sub ToolStripMenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem17.Click
        Dim RetnID As Long
        'Dim LoadForm As Boolean
        RetnID = GetForm26RetnID("2")
        If RetnID = 0 Then Exit Sub

        FromDateQ = Format("01/July/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("30/Sep/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        'Load26QData()
        ConvertWhich = "26Q2"
        frmTDS26Q.Tag = RetnID
        frmTDS26Q.tabMain.SelectedIndex = 0
        frmTDS26Q.quter = "26Q2"
        If CheckPRNNo(frmTDS26Q.quter) = False Then
            Call Load26QData(RetnID)
            frmTDS26Q.ShowDialog()
        End If

    End Sub

    Private Sub ToolStripMenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem18.Click
        Dim RetnID As Long
        'Dim LoadForm As Boolean

        RetnID = GetForm26RetnID("3")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Oct/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("31/Dec/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS26Q
        ConvertWhich = "26Q3"
        frmTDS26Q.Tag = RetnID
        frmTDS26Q.tabMain.SelectedIndex = 0
        frmTDS26Q.quter = "26Q3"
        If CheckPRNNo(frmTDS26Q.quter) = False Then
            Call Load26QData(RetnID)
            frmTDS26Q.ShowDialog()
        End If
    End Sub

    Private Sub ToolStripMenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem19.Click
        Dim RetnID As Long

        RetnID = GetForm26RetnID("4")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Jan/" & Strings.Right(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("31/Mar/" & Strings.Right(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS26Q
        ConvertWhich = "26Q4"
        frmTDS26Q.Tag = RetnID
        frmTDS26Q.tabMain.SelectedIndex = 0
        frmTDS26Q.quter = "26Q4"
        If CheckPRNNo(frmTDS26Q.quter) = False Then
            Call Load26QData(RetnID)
            frmTDS26Q.ShowDialog()
        End If
    End Sub

    Private Sub ReturnForQuarter1ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnForQuarter1ToolStripMenuItem.Click
        Dim RetnID As Long

        RetnID = GetForm27ERetnID("1")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Apr/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("30/Jun/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS27Q
        ConvertWhich = "27EQ1"
        frmTDS27EQ.Tag = RetnID
        frmTDS27EQ.tabMain.SelectedIndex = 0
        frmTDS27EQ.quter = "27EQ1"
        If CheckPRNNo(frmTDS27EQ.quter) = False Then
            Load27EQData(RetnID)
            frmTDS27EQ.ShowDialog()
        End If
    End Sub

    Private Sub ReturnForQuarter2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnForQuarter2ToolStripMenuItem.Click
        Dim RetnID As Long

        RetnID = GetForm27ERetnID("2")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/July/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("30/Sep/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS27Q
        ConvertWhich = "27EQ2"
        frmTDS27EQ.Tag = RetnID
        frmTDS27EQ.tabMain.SelectedIndex = 0
        frmTDS27EQ.quter = "27EQ2"
        If CheckPRNNo(frmTDS27EQ.quter) = False Then
            Load27EQData(RetnID)
            frmTDS27EQ.ShowDialog()
        End If
    End Sub

    Private Sub ReturnForQuarter3ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnForQuarter3ToolStripMenuItem.Click
        Dim RetnID As Long

        RetnID = GetForm27ERetnID("3")
        If RetnID = 0 Then Exit Sub

        FromDateQ = Format("01/Oct/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("31/Dec/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS27Q
        ConvertWhich = "27EQ3"
        frmTDS27EQ.Tag = RetnID
        frmTDS27EQ.tabMain.SelectedIndex = 0
        frmTDS27EQ.quter = "27EQ3"
        If CheckPRNNo(frmTDS27EQ.quter) = False Then
            Load27EQData(RetnID)
            frmTDS27EQ.ShowDialog()
        End If
    End Sub

    Private Sub ReturnForQuarter4ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnForQuarter4ToolStripMenuItem.Click
        Dim RetnID As Long

        RetnID = GetForm27ERetnID("4")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Jan/" & Strings.Right(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("31/Mar/" & Strings.Right(FY, 4), "dd/MMM/yyyy")
        'Load frmTDS27Q
        ConvertWhich = "27EQ4"
        frmTDS27EQ.Tag = RetnID
        frmTDS27EQ.tabMain.SelectedIndex = 0
        frmTDS27EQ.quter = "27EQ4"
        If CheckPRNNo(frmTDS27EQ.quter) = False Then
            Load27EQData(RetnID)
            frmTDS27EQ.ShowDialog()
        End If
    End Sub

    Private Sub ChallanNo280SAAdvTaxToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'frmChallan281.Show()
    End Sub

    Private Sub ChallanNo281TDSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'frmChallan281.Show()
    End Sub

    Private Sub GoToOnlineChallanWebsiteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoToOnlineChallanWebsiteToolStripMenuItem.Click
        Process.Start("https://onlineservices.tin.nsdl.com/etaxnew/tdsnontds.jsp")
    End Sub

    Private Sub GoToTRACESWbsiteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoToTRACESWbsiteToolStripMenuItem.Click
        Process.Start("https://tin-nsdl.com/")
    End Sub

    Private Sub GoToTRACESWbsiteToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoToTRACESWbsiteToolStripMenuItem1.Click
        Process.Start("https://www.tdscpc.gov.in/")
    End Sub

    Private Sub StandartToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StandartToolStripMenuItem.Click
        frmform16Parametre.Show()
    End Sub

    Private Sub ValidationUtilityFromNSDLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValidationUtilityFromNSDLToolStripMenuItem.Click
        Process.Start("" & My.Application.Info.DirectoryPath & "\e-TDS Files\TDS_STANDALONE_FVU_5.6.jar")
        'Call OpenFVUNew("" & My.Application.Info.DirectoryPath & "\e-TDS Files\")
    End Sub

    Private Sub FindBSRCodeGoOnlineToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBSRCodeGoOnlineToolStripMenuItem.Click
        Process.Start("http://tin.nsdl.com/OLTASListOfBSR.asp")
    End Sub

    Private Sub CalculateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalculateToolStripMenuItem.Click
        Process.Start("calc.exe")
    End Sub

    Private Sub NotePadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotePadToolStripMenuItem.Click
        Process.Start("NotePad")
    End Sub

    Private Sub NSDLSCorrectionUtilityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NSDLSCorrectionUtilityToolStripMenuItem.Click
        If My.Computer.FileSystem.FileExists("" & My.Application.Info.DirectoryPath & "\e-TDS Files\RPU\RPU.jar") Then
            Process.Start("" & My.Application.Info.DirectoryPath & "\e-TDS Files\RPU\RPU.jar")
        Else
            MessageBox.Show("NSDL's RPU utility file does not exist!Please call JAK InfoSolutions Pvt. Ltd.,Nagpur.")
        End If
    End Sub

    Private Sub BlankCorrectionFrom27BToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BlankCorrectionFrom27BToolStripMenuItem.Click
        Process.Start("" & My.Application.Info.DirectoryPath & "\FORM15G\ITD_EFILING_FORM15G_PR2.jar")
    End Sub

    Private Sub DisclaimerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DisclaimerToolStripMenuItem.Click
        Process.Start("" & My.Application.Info.DirectoryPath & "\FORM15H\ITD_EFILING_FORM15H_PR2.jar")
    End Sub

    Private Sub BlankCorrectionFrom27BToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BlankCorrectionFrom27BToolStripMenuItem1.Click
        frmtermscond.Show()
    End Sub

    Private Sub TANRegistrationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        'frmOnlineTanReg.Show()
    End Sub

    Private Sub TANLoginToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TANLoginToolStripMenuItem.Click
        frmOnlineTanLogin.Show()
    End Sub

    Private Sub DeductorsManualNSDLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeductorsManualNSDLToolStripMenuItem.Click
        Process.Start("" & My.Application.Info.DirectoryPath & "\Support\Deductors_Manual_Qtrly.htm")
    End Sub

    Private Sub DeductorsChecklistWizinTDSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeductorsChecklistWizinTDSToolStripMenuItem.Click
        Process.Start("" & My.Application.Info.DirectoryPath & "\Support\JAK_Checklist.htm")
    End Sub

    Private Sub ShowWhatNewWindowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowWhatNewWindowToolStripMenuItem.Click
        FrmWhatsnew.Show()
    End Sub

    Public Function MarqueeLeft(ByVal Text As String)
        Dim Str1 As String = Text.Remove(0, 1)
        Dim Str2 As String = Text(0)
        Return Str1 & Str2
    End Function
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Label3.Text = MarqueeLeft(Label3.Text)
        If Label3.Top = -Label3.Height Then
            Label3.Top = Label3.Height
        Else
            Label3.Top -= 1
        End If
    End Sub

    Private Sub frmTDS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.Top = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - Me.Height
        'Me.Left = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - Me.Width
        'Dim DesignScreenWidth As Integer = 1600
        'Dim DesignScreenHeight As Integer = 1200
        'Dim CurrentScreenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        'Dim CurrentScreenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
        'Dim RatioX As Double = CurrentScreenWidth / DesignScreenWidth
        'Dim RatioY As Double = CurrentScreenHeight / DesignScreenHeight
        'For Each iControl In Me.Controls
        '    With iControl
        '        If (.GetType.GetProperty("Width").CanRead) Then .Width = CInt(.Width * RatioX)
        '        If (.GetType.GetProperty("Height").CanRead) Then .Height = CInt(.Height * RatioY)
        '        If (.GetType.GetProperty("Top").CanRead) Then .Top = CInt(.Top * RatioX)
        '        If (.GetType.GetProperty("Left").CanRead) Then .Left = CInt(.Left * RatioY)
        '    End With
        'Next
        'Form.CenterToScreen()
        ' Me.Size = Screen.PrimaryScreen.WorkingArea.Size
        'Me.Location = New Point(180, 220)
        'Dim fso As New FileSystemObject
        Dim txtStream As StreamReader
        txtStream = File.OpenText(Application.StartupPath & "\Support\ReleaseNotes.txt")
        If Not txtStream.EndOfStream Then
            Label3.Text = txtStream.ReadToEnd()
        End If
        txtStream.Close
        PrepareMenu()
        FetchQrtDet()
        Dim s As String
        s = frmCoMst.lvwCo.SelectedItems(0).SubItems(0).Text
        Me.Text = s
        'Label3.Top = Scroll.op + picScroll.ScaleHeight


        'Me.Parent = frmCoMst
        'lblScroll.Top = picScroll.ScaleTop + picScroll.ScaleHeight
        'Label3.Text = My.Computer.FileSystem.ReadAllText(My.Application.Info.DirectoryPath & "\Support\ReleaseNotes.txt")
    End Sub
    Private Sub FetchQrtDet()
        Dim i As Integer
        Dim nds As New DataSet
        Dim st As String
        st = "Form"
        nds = FetchDataSet(" Select retnid,frmtype,dtoffiling,prn,rprn,newReceiptNo from retnmst where coid=" & selectedcoid & " order by frmtype")
        'grd.DataSource = nds.Tables(0)

        With grd
            .ColumnCount = 9

            '.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised
            '.ColumnHeadersVisible = False
            .Columns(0).Visible = False
            .Columns(1).ReadOnly = True
            .Columns(1).Width = 100
            .Columns(2).Visible = False
            .Columns(3).Visible = False
            .Columns(4).Visible = False
            .Columns(5).Visible = False
            .Columns(6).Visible = False
            .Columns(7).Visible = False
            .Columns(8).Visible = False

            .Width = 105
            .RowCount = nds.Tables(0).Rows.Count
            For i = 0 To nds.Tables(0).Rows.Count - 1

                .Rows.Item(i).Cells(0).Value = nds.Tables(0).Rows(i)(0).ToString()
                .Rows.Item(i).Cells(1).Value = "Form " & nds.Tables(0).Rows(i)(1).ToString()

                .Rows.Item(i).Cells(2).Value = nds.Tables(0).Rows(i)(2).ToString()
                .Rows.Item(i).Cells(3).Value = nds.Tables(0).Rows(i)(3).ToString()
                .Rows.Item(i).Cells(4).Value = nds.Tables(0).Rows(i)(4).ToString()
                .Rows.Item(i).Cells(5).Value = nds.Tables(0).Rows(i)(5).ToString()
                .Rows.Item(i).Cells(6).Value = ""
                .Rows.Item(i).Cells(7).Value = ""
                .Rows.Item(i).Cells(8).Value = ""

            Next
            If .RowCount > 0 Then
                TLPanel.Visible = True
                .Rows.Item(0).Cells(1).Selected = True
            Else
                TLPanel.Visible = False

            End If
            '' Get the current cell location.
            'Dim y As Integer = DataGridView1.CurrentCellAddress.Y
            'Dim x As Integer = DataGridView1.CurrentCellAddress.X



            'If Not (rst.EOF Or rst.BOF) Then
            '.Rows = rst.RecordCount

            'End If
            '        While Not (rst.EOF Or rst.BOF)
            '            .TextMatrix(i, 0) = rst(0)
            '            .TextMatrix(i, 1) = "Form " & rst(1)
            '            .TextMatrix(i, 2) = IIf(IsNull(rst(2)), vbNullString, Format(rst(2), "dd/mm/yy"))
            '            .TextMatrix(i, 3) = IIf(IsNull(rst(3)), vbNullString, rst(3))
            '            .TextMatrix(i, 4) = IIf(IsNull(rst(5)), vbNullString, rst(5))
            '            .RowHeight(i) = 400
            '            i = i + 1
            '            rst.MoveNext
            'Wend

            '.Row = 0
            '        .Col = 1

        End With

        'If grd.TextMatrix(0, 1) <> vbNullString Then
        '    grd.Visible = True
        '    lbldt.Visible = True
        '    dtpDt.Visible = True
        '    lblprn.Visible = True
        '    txtPRN.Visible = True
        '    lblrprn.Visible = True
        '    'txtRPRN.Visible = True
        '    txtRcptNo.Visible = True
        'Else
        '    grd.Visible = False
        '    lbldt.Visible = False
        '    dtpDt.Visible = False
        '    lblprn.Visible = False
        '    txtPRN.Visible = False
        '    lblrprn.Visible = False
        '    'txtRPRN.Visible = False
        '    txtRcptNo.Visible = False
        'End If
    End Sub


    Private Sub frmTDS_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Me.EnterTab(e)
    End Sub

    Private Sub frmTDS_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Timer1.Start()
    End Sub
    Public Sub EnterTab(ByVal e As KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub PrepareMenu()
        'With cPopMnuTDS
        '    .ImageList = imgMenuIcons   'attach imagelist to menu control
        '    .SubClassMenu Me      'Subclass VB menu items..
        '    .OfficeXpStyle = True
        '    .ItemIcon("mnuMasterBank") = imgMenuIcons.ListImages("data").Index - 1
        '    .ItemIcon("mnuMasterDeductee") = imgMenuIcons.ListImages("DeducteeMst").Index - 1
        '    .ItemIcon("mnuMasterExit") = imgMenuIcons.ListImages("close").Index - 1
        '    '        .ItemIcon("mnuForm24DataEntry") = imgMenuIcons.ListImages("data1").Index - 1
        '    '        .ItemIcon("mnuForm24Convert") = imgMenuIcons.ListImages("convert").Index - 1
        '    '        .ItemIcon("mnuForm26DataEntry") = imgMenuIcons.ListImages("data1").Index - 1
        '    '        .ItemIcon("mnuForm26Convert") = imgMenuIcons.ListImages("convert").Index - 1
        '    '        .ItemIcon("mnuForm27EDataEntry") = imgMenuIcons.ListImages("data1").Index - 1
        '    '        .ItemIcon("mnuForm27EConvert") = imgMenuIcons.ListImages("convert").Index - 1
        '    '        .ItemIcon("mnuHelpEmail") = imgMenuIcons.ListImages("feedback").Index - 1
        'End With
    End Sub

    Private Sub grd_SelectionChanged(sender As Object, e As EventArgs) Handles grd.SelectionChanged
        'Get the current cell location.
        Dim y As Integer = grd.CurrentCellAddress.Y
        Dim x As Integer = grd.CurrentCellAddress.X
        Dim rect As New Rectangle
        If grd.RowCount = 1 Then
            TLPanel.Top = grd.Top

        Else
            grd.GetCellDisplayRectangle(0, 0, True)
            rect = grd.GetCellDisplayRectangle(x, y, True)
            TLPanel.Top = grd.Top + rect.Location.Y
        End If
        txtdate.Text = ""
        txtPRN.Text = ""
        txtRcptNo.Text = ""
        txtRetnid.Text = ""
        txtdate.Text = grd.Rows.Item(y).Cells(2).Value
        txtPRN.Text = grd.Rows.Item(y).Cells(3).Value
        txtRcptNo.Text = grd.Rows.Item(y).Cells(5).Value
        txtRetnid.Text = grd.Rows.Item(y).Cells(0).Value
    End Sub

    Private Sub grd_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles grd.CellBeginEdit

    End Sub

    Private Sub txtdate_LostFocus(sender As Object, e As EventArgs) Handles txtdate.LostFocus
        Dim y As Integer = grd.CurrentCellAddress.Y
        Dim ftyp As String

        txtdate.SelectionLength = 0
        If txtdate.Text <> "  /  /" And Not IsDate(txtdate.Text) Then
            MessageBox.Show("Enter Valid Date of Filing!!", "Warning")
            txtdate.Focus()
            Exit Sub
        Else
            If txtdate.Text <> "  /  /" And IsDate(txtdate.Text) Then
                ftyp = grd.Rows.Item(y).Cells(1).Value.ToString()


                Select Case Strings.Right(ftyp, 2)
                    Case "Q1"
                        If (CDate(txtdate.Text) < CDate("30/06/" & Year(FromDate))) Or (CDate(txtdate.Text) > Date.Now) Then
                            MessageBox.Show("Invalid Date!! Date range for Quarter" & ftyp.Substring(1, 1) & " is " & CDate("30/06/" & Year(FromDate)) & " - " & Date.Now, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            txtdate.Text = "  /  /"
                            txtdate.Focus()
                            Exit Sub
                        End If
                    Case "Q2"
                        If (CDate(txtdate.Text) < CDate("30/09/" & Year(FromDate))) Or (CDate(txtdate.Text) > Date.Now) Then
                            MessageBox.Show("Invalid Date!! Date range for Quarter" & ftyp.Substring(1, 1) & " is " & CDate("30/09/" & Year(FromDate)) & " - " & Date.Now, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            txtdate.Text = "  /  /"
                            txtdate.Focus()
                            Exit Sub
                        End If
                    Case "Q3"
                        If (CDate(txtdate.Text) < CDate("31/12/" & Year(FromDate))) Or (CDate(txtdate.Text) > Date.Now) Then
                            MessageBox.Show("Invalid Date!! Date range for Quarter" & Strings.Right(ftyp, 2) & " is " & CDate("31/12/" & Year(FromDate)) & " - " & Date.Now, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            txtdate.Text = "  /  /"
                            txtdate.Focus()
                            Exit Sub
                        End If
                    Case "Q4"
                        If (CDate(txtdate.Text) < CDate("31/03/" & Year(ToDate))) Or (CDate(txtdate.Text) > Date.Now) Then
                            MessageBox.Show("Invalid Date!! Date range for Quarter" & ftyp.Substring(1, 1) & " is " & CDate("31/03/" & Year(ToDate)) & " - " & Date.Now, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            txtdate.Text = "  /  /"
                            txtdate.Focus()
                            Exit Sub
                        End If
                End Select
            End If
        End If

        If Len(txtPRN.Text) = 15 And Not IsDate(txtdate.Text) Then
            'MsgBox "Enter Valid Date of Filing!!", vbCritical, "Warning"
            '    dtpDt.SetFocus
            Exit Sub
        End If
        'If Len(txtPRN) = 0 And Not IsDate(txtdate.Text) Then
        '    txtRPRN.te = ""
        'End If
        If grd.Rows.Count > 0 Then
            grd.Rows.Item(y).Cells(2).Value = txtdate.Text
        End If
        'grd.TextMatrix(grd.Row, 2) = dtpDt.Text
        CallUpdFiling()
    End Sub

    Private Sub txtdate_Enter(sender As Object, e As EventArgs) Handles txtdate.Enter
        txtdate.BackColor = Color.LightYellow
    End Sub

    Private Sub txtdate_Leave(sender As Object, e As EventArgs) Handles txtdate.Leave
        txtdate.BackColor = Color.White

    End Sub

    Private Sub txtPRN_TextChanged(sender As Object, e As EventArgs) Handles txtPRN.TextChanged

    End Sub

    Private Sub txtPRN_Enter(sender As Object, e As EventArgs) Handles txtPRN.Enter
        txtPRN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtPRN_LostFocus(sender As Object, e As EventArgs) Handles txtPRN.LostFocus
        If Len(txtPRN.Text) <> 0 And Len(txtPRN.Text) < 15 And IsDate(txtdate.Text) Then
            MessageBox.Show("PRN should be 15 digit!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPRN.Focus()
            Exit Sub
        End If
        If Len(txtPRN.Text) = 15 And Not IsDate(txtdate.Text) Then
            MessageBox.Show("Date of Filing is compulsory!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtdate.Focus()
            Exit Sub
        End If
        'If Len(txtPRN) = 0 And Not IsDate(txtdate) Then
        '    txtRPRN.text = ""
        'End If

        CallUpdFiling()
        grd.Rows.Item(grd.CurrentCell.RowIndex).Cells(3).Value = txtPRN.Text
    End Sub

    Private Sub txtPRN_Leave(sender As Object, e As EventArgs) Handles txtPRN.Leave
        txtPRN.BackColor = Color.White
    End Sub

    Private Sub txtRcptNo_TextChanged(sender As Object, e As EventArgs) Handles txtRcptNo.TextChanged

    End Sub

    Private Sub txtRcptNo_Enter(sender As Object, e As EventArgs) Handles txtRcptNo.Enter
        txtRcptNo.BackColor = Color.LightYellow
    End Sub

    Private Sub txtRcptNo_Leave(sender As Object, e As EventArgs) Handles txtRcptNo.Leave
        txtRcptNo.BackColor = Color.White

    End Sub
    Private Sub CallUpdFiling()
        If (IsDate(txtdate.Text) And Len(txtPRN.Text) = 15) Or (Not IsDate(txtdate.Text) And Len(txtPRN.Text) = 0 And Len(txtRcptNo.Text) = 0) Then
            Dim upd As Boolean = UpdRetnDet()
            If upd = False Then
                MessageBox.Show("Some error has occured!! Cannot Save", "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

    End Sub


    Public Function UpdRetnDet() As Boolean
        Dim sql As String
        Dim cmd As New OleDbCommand
        Dim dt As Date
        Dim dts As String

        dts = txtdate.Text
        If dts <> "  /  /" Then
            dt = txtdate.Text
        Else
            dt = Nothing
        End If

        sql = "Update RetnMst Set dtoffiling=" & IIf(dts = "  /  /", "Null", "'" & dt.ToString("dd/MMM/yyyy") & "'") &
                 ",PRN=" & IIf(Val(txtPRN.Text) = 0, "Null", "'" & txtPRN.Text & "'") &
                 ",NewReceiptNo=" & IIf(Len(txtRcptNo.Text) = 0, "Null", "'" & txtRcptNo.Text & "'") &
                 " Where RetnId = " & Val(txtRetnid.Text)



        Try
            cmd.CommandText = sql
            cmd.Connection = cn
            cmd.ExecuteNonQuery()
            UpdRetnDet = True
        Catch ex As Exception
            UpdRetnDet = False
            'MessageBox.Show("Message:", ex.Message)
        End Try
        cmd.Dispose()
    End Function

    Private Sub txtRcptNo_LostFocus(sender As Object, e As EventArgs) Handles txtRcptNo.LostFocus
        If Len(txtRcptNo.Text) = 8 And Len(txtRcptNo.Text) > 0 And Len(txtRcptNo.Text) <> 8 Then
            MessageBox.Show("Receipt No should be 8 digits!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtRcptNo.Focus()
            Exit Sub
        End If
        CallUpdFiling()
        grd.Rows.Item(grd.CurrentCell.RowIndex).Cells(5).Value = txtRcptNo.Text


        If (Len(txtPRN.Text) = 15 And IsDate(txtdate.Text)) Or (txtPRN.Text = vbNullString And txtdate.Text = "  /  /") Then
            If grd.RowCount - 1 > grd.CurrentRow.Index Then


                grd.Rows.Item(grd.CurrentRow.Index + 1).Cells(1).Selected = True

                'Get the current cell location.
                Dim y As Integer = grd.CurrentCellAddress.Y
                Dim x As Integer = grd.CurrentCellAddress.X

                If grd.RowCount = 1 Then
                    TLPanel.Top = grd.Top

                Else
                    Dim rect As Rectangle = grd.GetCellDisplayRectangle(x, y, True)
                    TLPanel.Top = grd.Top + rect.Location.Y
                End If
                txtdate.Text = ""
                txtPRN.Text = ""
                txtRcptNo.Text = ""
                txtRetnid.Text = ""
                txtdate.Text = grd.Rows.Item(y).Cells(2).Value
                txtPRN.Text = grd.Rows.Item(y).Cells(3).Value
                txtRcptNo.Text = grd.Rows.Item(y).Cells(5).Value
                txtRetnid.Text = grd.Rows.Item(y).Cells(0).Value
            End If
        End If
    End Sub

    Private Sub txtPRN_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPRN.KeyPress
        Dim i As Integer = CtrlKeyPress(txtPRN, Asc(e.KeyChar), MyKeypressEnum.KeyPressNumberOnly, 0, False)
        If i < 0 Then
            e.KeyChar = ""
        End If
        If Len(txtPRN.Text) = 15 Then

            txtPRN_LostFocus(sender, e)
        End If
    End Sub

    Private Sub txtRcptNo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtRcptNo.KeyPress

        If Len(txtPRN.Text) = 15 Then

            txtPRN_LostFocus(sender, e)
        End If

        Dim str As String
        str = txtPRN.Text
        If Len(str.Trim) = 0 Or Not IsDate(txtdate.Text) Then
            MessageBox.Show("Please enter data for Original PRN first")
            txtPRN.Focus()
            Exit Sub
        End If
        If Len(txtRcptNo.Text) = 8 Then

            txtRcptNo_LostFocus(sender, e)
        End If

    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Dim RetnID As Long
        'Dim LoadForm As Boolean
        'Dim rst As New ADODB.Recordset
        RetnID = GetForm24RetnID("1")
        If RetnID = 0 Then Exit Sub
        FromDateQ = Format("01/Apr/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ToDateQ = Format("30/Jun/" & Strings.Left(FY, 4), "dd/MMM/yyyy")
        ConvertWhich = "24Q1"
        frmTDS24Q.Tag = RetnID
        frmTDS24Q.tabMain.SelectedIndex = 0
        frmTDS24Q.quter = "24Q1"
        'frmTDS24Q.Show()      'just to give a visual effect of loading...
        If CheckPRNNo(frmTDS24Q.quter) = False Then
            Call Load24QData(RetnID)
            frmTDS24Q.ShowDialog()
        End If
    End Sub

    Private Sub SectionChecked24(sect As String)
        With frmTDS24Q
            For Each chk As CheckBox In .PanelCheckBox.Controls
                If chk.Text = sect Then
                    chk.Checked = True
                    .cmdNext.Enabled = True
                    Exit Sub
                End If
            Next
        End With
    End Sub

    Public Function GetForm24RetnID(Qtr As String) As Long
        On Error GoTo canerr
        Dim nds As New DataSet
        Dim oRetnMst As New ClsRetnMstObj ', rst As New ADODB.Recordset
        nds = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " and FrmType= '24Q" & Qtr & "'")
        If nds.Tables(0).Rows.Count <= 0 Then
            If MessageBox.Show("Form No 24Q for Quarter " & Qtr & " for this company does not exist" & vbCrLf & "Do you want to create it?", "Create Form No 24Q", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
                oRetnMst.coid = selectedcoid
                oRetnMst.AYear = AY
                oRetnMst.FrmType = "24Q" & Qtr
                If oRetnMst.Insert(oRetnMst) = False Then
                    MessageBox.Show("Unable to create Form No. 24Q, Call JAK Infosolutions", "FORM NOT CREATED", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                    GoTo cleanup
                End If
            Else
                GoTo cleanup
            End If
        End If
        'nds = Nothing
        'nds = New DataSet
        nds = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " and FrmType='24Q" & Qtr & "'")
        GetForm24RetnID = nds.Tables(0).Rows(0)("RetnID")

cleanup:
        'Me.MousePointer = vbDefault
        'If rst.State = adStateOpen Then rst.Close
        'Set rst = Nothing
        'Set oRetnMst = Nothing
        nds.Dispose()
        oRetnMst = Nothing
        Exit Function

canerr:
        If Err.Number <> 32755 Then

            MsgBox(Err.Description,, Err.Number)
        End If
        GoTo cleanup
    End Function


    Public Function Insert(selectedcoid As Integer, AYear As String, FrmType As String) As Boolean
        Dim sql As String
        Dim cmd As New OleDbCommand
        Dim transaction As OleDbTransaction

        sql = "Insert Into RetnMst (RetnId,Coid,AYear,FrmType,TxtFileName) Values (" _
                & MaxID() + 1 & "," & IIf(selectedcoid = 0, 0, selectedcoid) & "," _
                & IIf(AYear = vbNullString, "Null", "'" & AYear & "'") & "," _
                & IIf(FrmType = vbNullString, "Null", "'" & FrmType & "'") & "," _
                & IIf(FrmType = vbNullString, "Null", "'" & FrmType & "'") & ")"
        cmd.Connection = cn
        cmd.CommandText = sql
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction

        Try
            cmd.ExecuteNonQuery()
            transaction.Commit()
            Insert = True
        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show(ex.Message)
            Insert = False
            Exit Function
        End Try
    End Function

    Public Function MaxID() As Long
        'Dim rs As New ADODB.Recordset
        Dim sql As String
        Dim nds As New DataSet
        sql = "Select Max(RetnId) as ID From RetnMst"
        nds = FetchDataSet(sql)
        If Not nds.Tables(0).Rows.Count > 0 Then
            MaxID = nds.Tables(0).Rows(0)(0)
        Else
            MaxID = 0
        End If
        nds.Dispose()
    End Function

    Private Sub frmTDS_DoubleClick(sender As Object, e As EventArgs) Handles Me.DoubleClick
        'MsgBox(Me.Location, "g")
    End Sub

    Private Sub TLPanel_Paint(sender As Object, e As PaintEventArgs) Handles TLPanel.Paint

    End Sub

    Private Sub frmTDS_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        Me.Dispose()

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub EmailUsYourSuggestionqueriesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EmailUsYourSuggestionqueriesToolStripMenuItem.Click
        Process.Start("mailto:jakinfo_ngp@sancharnet.in, jak_tds@yahoo.com")
        'OpenEmailClient("jakinfo_ngp@sancharnet.in, jak_tds@yahoo.com")
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
    '    Public Sub FillBankDetails(frm As Form)
    '        Dim nds As New DataSet
    '        'Filling BSR Code in Challan Detail
    '        nds = FetchDataSet("select BankBrCode from BankMst WHERE CoID=" & selectedcoid & " order by BankBrCode")
    '        frm.cboBankBrCode.items.clear
    '        While Not (rst.EOF Or rst.BOF)
    '            frm.cboBankBrCode.AddItem Format(rst!BankBrCode, "0000000")
    '    rst.MoveNext
    '    Wend
    '    If rst.State = adStateOpen Then rst.Close
    'Set rst = Nothing
    'End Sub

End Class