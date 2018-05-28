Imports System.Data.OleDb

Public Class FRM16Detail
    Dim transaction As OleDb.OleDbTransaction
    Dim oCoObj As New clsCoMst
    Dim oForm16 As New Form16Details
    Dim frm As New frmdeduteeTDSMST
    '  Dim ReadStream As TextStream,
    Dim AllowCboTxt As String, Sec80CcgCboTxt, OthIncCbotxt As String ' StlvwChallan.Items(0).SubItems(1).Textring, Sec80CcgCboTxt As String
    Dim Sec80CCboTxt As String, Sec80CcfCboTxt As String, Chp6aCboTxt As String, DataRead
    Private Sub frm16Details_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub cmbName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbo16DedName.Enter
        cbo16DedName.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbo16DedName.Leave
        cbo16DedName.BackColor = Color.White
    End Sub
    Private Sub txtName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoName.Enter
        txt16CoName.BackColor = Color.LightYellow
    End Sub

    Private Sub txtName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoName.Leave
        txt16CoName.BackColor = Color.White
    End Sub

    Private Sub txtadress1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd1.Enter
        txt16CoAdd1.BackColor = Color.LightYellow
    End Sub
    Private Sub txtadress1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd1.Leave
        txt16CoAdd1.BackColor = Color.White
    End Sub

    Private Sub txtadress2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd2.Enter
        txt16CoAdd2.BackColor = Color.LightYellow
    End Sub
    Private Sub txtadress2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd2.Leave
        txt16CoAdd2.BackColor = Color.White
    End Sub

    Private Sub txtadress3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd3.Enter
        txt16CoAdd3.BackColor = Color.LightYellow
    End Sub
    Private Sub txtadress3_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd3.Leave
        txt16CoAdd3.BackColor = Color.White
    End Sub

    Private Sub txtadress4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd4.Enter
        txt16CoAdd4.BackColor = Color.LightYellow
    End Sub
    Private Sub txtadress4_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd4.Leave
        txt16CoAdd4.BackColor = Color.White
    End Sub

    Private Sub txtadress5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd5.Enter
        txt16CoAdd5.BackColor = Color.LightYellow
    End Sub
    Private Sub txtadress5_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoAdd5.Leave
        txt16CoAdd5.BackColor = Color.White
    End Sub

    Private Sub cmbState_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbo16CoState.Enter
        cbo16CoState.BackColor = Color.LightYellow
    End Sub
    Private Sub cmbState_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbo16CoState.Leave
        cbo16CoState.BackColor = Color.White
    End Sub

    Private Sub txtPCode_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoPin.Enter
        txt16CoPin.BackColor = Color.LightYellow
    End Sub
    Private Sub txtPCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoPin.Leave
        txt16CoPin.BackColor = Color.White
    End Sub

    Private Sub txtDedTAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoTan.Enter
        txt16CoTan.BackColor = Color.LightYellow
    End Sub
    Private Sub txtDedTAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoTan.Leave
        txt16CoTan.BackColor = Color.White
    End Sub

    Private Sub txtDedPan_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoPan.Enter
        txt16CoPan.BackColor = Color.LightYellow
    End Sub
    Private Sub txtDedPan_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16CoPan.Leave
        txt16CoPan.BackColor = Color.White
    End Sub

    Private Sub txtDesig_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtdesig.Enter
        txtdesig.BackColor = Color.LightYellow
    End Sub
    Private Sub txtDesig_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtdesig.Leave
        txtdesig.BackColor = Color.White
    End Sub

    Private Sub txtEmpPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16pan.Enter
        txt16pan.BackColor = Color.LightYellow
    End Sub
    Private Sub txtEmpPAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16pan.Leave
        txt16pan.BackColor = Color.White
    End Sub

    Private Sub txtAssigYear_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16AY.Enter
        txt16AY.BackColor = Color.LightYellow
    End Sub
    Private Sub txtAssigYear_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16AY.Leave
        txt16AY.BackColor = Color.White
    End Sub

    Private Sub txt171_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16gross1.Enter
        txt16gross1.BackColor = Color.LightYellow
    End Sub
    Private Sub txt171_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16gross1.Leave
        txt16gross1.BackColor = Color.White
    End Sub

    Private Sub txt172_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16gross2.Enter
        txt16gross2.BackColor = Color.LightYellow
    End Sub
    Private Sub txt172_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16gross2.Leave
        txt16gross2.BackColor = Color.White
    End Sub

    Private Sub txt173_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16gross3.Enter
        txt16gross3.BackColor = Color.LightYellow
    End Sub
    Private Sub txt173_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16gross3.Leave
        txt16gross3.BackColor = Color.White
    End Sub

    Private Sub txtTotSalCur_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16grosstotCurEmp.Enter
        txt16grosstotCurEmp.BackColor = Color.LightYellow
    End Sub
    Private Sub txtTotSalCur_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16grosstotCurEmp.Leave
        txt16grosstotCurEmp.BackColor = Color.White
    End Sub

    Private Sub txtTotSalPre_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16grosstotPreEmp.Enter
        txt16grosstotPreEmp.BackColor = Color.LightYellow
    End Sub
    Private Sub txtTotSalPre_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16grosstotPreEmp.Leave
        txt16grosstotPreEmp.BackColor = Color.White
    End Sub

    Private Sub txtTotGrossSal_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16grosstotBoth.Enter
        txt16grosstotBoth.BackColor = Color.LightYellow
    End Sub
    Private Sub txtTotGrossSal_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16grosstotBoth.Leave
        txt16grosstotBoth.BackColor = Color.White
    End Sub

    Private Sub txtAllowTotal_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtallow.Enter
        txtallow.BackColor = Color.LightYellow
    End Sub
    Private Sub txtAllowTotal_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtallow.Leave
        txtallow.BackColor = Color.White
    End Sub

    Private Sub txtBal_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16bal.Enter
        txt16bal.BackColor = Color.LightYellow
    End Sub
    Private Sub txtBal_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16bal.Leave
        txt16bal.BackColor = Color.White
    End Sub

    Private Sub txtEntAllow_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16EntAllow.Enter
        txt16EntAllow.BackColor = Color.LightYellow
    End Sub
    Private Sub txtEntAllow_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16EntAllow.Leave
        txt16EntAllow.BackColor = Color.White
    End Sub

    Private Sub txtTaxEmp_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16ProfTax.Enter
        txt16ProfTax.BackColor = Color.LightYellow
    End Sub
    Private Sub txtTaxEmp_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16ProfTax.Leave
        txt16ProfTax.BackColor = Color.White
    End Sub

    Private Sub txtTotDec_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotalDedct.Enter
        txtTotalDedct.BackColor = Color.LightYellow
    End Sub
    Private Sub txtTotDec_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotalDedct.Leave
        txtTotalDedct.BackColor = Color.White
    End Sub

    Private Sub txtTaxSal_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TaxableSalary.Enter
        txt16TaxableSalary.BackColor = Color.LightYellow
    End Sub
    Private Sub txtTaxSal_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TaxableSalary.Leave
        txt16TaxableSalary.BackColor = Color.White
    End Sub

    Private Sub txtOtherInc_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16otherIncome.Enter
        txt16otherIncome.BackColor = Color.LightYellow
    End Sub
    Private Sub txtOtherInc_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16otherIncome.Leave
        txt16otherIncome.BackColor = Color.White
    End Sub

    Private Sub txtGrossIncome_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16GTI.Enter
        txt16GTI.BackColor = Color.LightYellow
    End Sub
    Private Sub txtGrossIncome_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16GTI.Leave
        txt16GTI.BackColor = Color.White
    End Sub

    Private Sub txt80cccTot_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt1680c.Enter
        txt1680c.BackColor = Color.LightYellow
    End Sub
    Private Sub txt80cccTot_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt1680c.Leave
        txt1680c.BackColor = Color.White
    End Sub

    Private Sub txt80CCFTot_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt1680CCF.Enter
        txt1680CCF.BackColor = Color.LightYellow
    End Sub
    Private Sub txt80CCFTot_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt1680CCF.Leave
        txt1680CCF.BackColor = Color.White
    End Sub

    Private Sub txt80CCGTot_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt1680CCG.Enter
        txt1680CCG.BackColor = Color.LightYellow
    End Sub
    Private Sub txt80CCGTot_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt1680CCG.Leave
        txt1680CCG.BackColor = Color.White
    End Sub

    Private Sub txtVIATot_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtT16OtherIVA.Enter
        txtT16OtherIVA.BackColor = Color.LightYellow
    End Sub
    Private Sub txtVIATot_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtT16OtherIVA.Leave
        txtT16OtherIVA.BackColor = Color.White
    End Sub


    Private Sub txtVIADeduc_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16OtherIVA.Enter
        txt16OtherIVA.BackColor = Color.LightYellow
    End Sub
    Private Sub txtVIADeduc_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16OtherIVA.Leave
        txt16OtherIVA.BackColor = Color.White
    End Sub


    Private Sub txtTotInco_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TotalTaxableIncome.Enter
        txt16TotalTaxableIncome.BackColor = Color.LightYellow
    End Sub
    Private Sub txtTotInco_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TotalTaxableIncome.Leave
        txt16TotalTaxableIncome.BackColor = Color.White
    End Sub


    Private Sub txtincome_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtincome.Leave
        txtincome.BackColor = Color.White
    End Sub

    Private Sub txtincome_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtincome.Enter
        txtincome.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16Tax_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16Tax.Enter
        txt16Tax.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16Tax_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16Tax.Leave
        txt16Tax.BackColor = Color.White
    End Sub

    Private Sub txt16Surcharge_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16Surcharge.Leave
        txt16Surcharge.BackColor = Color.White
    End Sub

    Private Sub txt16Surcharge_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16Surcharge.Enter
        txt16Surcharge.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16EduCess_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16EduCess.Leave
        txt16EduCess.BackColor = Color.White
    End Sub

    Private Sub txt16EduCess_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16EduCess.Enter
        txt16Surcharge.BackColor = Color.LightYellow
    End Sub

    Private Sub txtTax_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTax.Leave
        txtTax.BackColor = Color.White
    End Sub

    Private Sub txtTax_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTax.Enter
        txtTax.BackColor = Color.LightYellow
    End Sub

    Private Sub txtSurcharge_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurcharge.Leave
        txtSurcharge.BackColor = Color.White
    End Sub

    Private Sub txtSurcharge_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurcharge.Enter
        txtSurcharge.BackColor = Color.LightYellow
    End Sub

    Private Sub txtEd_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEd.Enter
        txtEd.BackColor = Color.LightYellow
    End Sub

    Private Sub txtEd_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEd.Leave
        txtEd.BackColor = Color.White
    End Sub

    Private Sub txt16TotalTax_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TotalTax.Leave
        txt16TotalTax.BackColor = Color.White
    End Sub

    Private Sub txt16TotalTax_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TotalTax.Enter
        txt16TotalTax.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16Relief_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16Relief.Enter
        txt16Relief.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16Relief_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16Relief.Leave
        txt16Relief.BackColor = Color.White
    End Sub

    Private Sub txt16NetTax_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16NetTax.Leave
        txt16NetTax.BackColor = Color.White
    End Sub

    Private Sub txt16NetTax_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16NetTax.Enter
        txt16NetTax.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16TDS1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TDS1.Leave
        txt16TDS1.BackColor = Color.White
    End Sub

    Private Sub txt16TDS1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TDS1.Enter
        txt16TDS1.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16TDS2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TDS2.Leave
        txt16TDS1.BackColor = Color.White
    End Sub

    Private Sub txt16TDS2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TDS2.Enter
        txt16TDS1.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16totalTDS_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16totalTDS.Leave
        txt16totalTDS.BackColor = Color.White
    End Sub

    Private Sub txt16totalTDS_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16totalTDS.Enter
        txt16totalTDS.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16TaxPreEmp_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TaxPreEmp.Enter
        txt16TaxPreEmp.BackColor = Color.LightYellow
    End Sub

    Private Sub txt16TaxPreEmp_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt16TaxPreEmp.Leave
        txt16TaxPreEmp.BackColor = Color.White

    End Sub

    Private Sub txtPayRef_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPayRef.Leave
        txtPayRef.BackColor = Color.White
    End Sub

    Private Sub txtPayRef_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPayRef.Enter
        txtPayRef.BackColor = Color.LightYellow
    End Sub

    Private Sub frm16Details_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim rst As New DataSet
        Dim R As Integer
        ' Dim oForm16 As New clsForm16Details
        txt16FrmDt.Text = Format(FromDate, "dd/MM/yy")
        txt16ToDt.Text = Format(ToDate, "dd/MM/yy")
        txt16AY.Text = AY
        CopyCoDetails()
        Fill24PRNNo()
        'tabForm16.Tab = 0
        tabForm16.Enabled = False
        'Set the grids right...

        '  grd16allow.ColHidden(2) = True
        'grd16otherIncome.ColHidden(2) = True
        'grd1680c.ColHidden(3) = True
        'grd1680CCF.ColHidden(3) = True
        'grd1680CCG.ColHidden(3) = True
        'grd16OtherIVA.ColHidden(4) = True

        grd16allow.Columns(2).Visible = False
        grd1680c.Columns(2).Visible = False
        grd1680CCF.Columns(3).Visible = False
        grd1680CCG.Columns(3).Visible = False
        grd16OtherIVA.Columns(4).Visible = False
        FillParaData()
        '   grd16allSubTotal
        txtDstatus.Text = ""

        For R = 1 To grd16ManualTax.Rows.Count - 1
            grd16ManualTax.Rows(R).Cells(0).Value = R
        Next R
        '     Autotaxcal

    End Sub
    Private Sub FillParaData()
        'read parameter data from the text file and fill the respective combos..

        oCoObj = oCoObj.FetchCo(selectedcoid)
        '* If fso.FileExists(Application.StartupPath & "\Database\Form16Parameters.txt") Then
        '*ReadStream = fso.OpenTextFile(Application.StartupPath & "\Database\Form16Parameters.txt")
        ' Do While Not ReadStream.AtEndOfStream
        'DataRead = Split(ReadStream.ReadLine, ",")
        'If DataRead(0) = "A" Then
        'If DataRead(2) = "T" Then AllowCboTxt = AllowCboTxt & "|" & DataRead(1)
        'ElseIf DataRead(0) = "O" Then
        'If DataRead(2) = "T" Then OthIncCbotxt = OthIncCbotxt & "|" & DataRead(1)
        'ElseIf DataRead(0) = "E" Then
        'If DataRead(2) = "T" Then Sec80CCboTxt = Sec80CCboTxt & "|" & DataRead(1)
        'ElseIf DataRead(0) = "V" Then
        'If DataRead(2) = "T" Then Chp6aCboTxt = Chp6aCboTxt & "|" & DataRead(1)
        'ElseIf DataRead(0) = "F" Then
        'If DataRead(2) = "T" Then Sec80CcfCboTxt = Sec80CcfCboTxt & "|" & DataRead(1)
        'ElseIf DataRead(0) = "G" Then
        'If DataRead(2) = "T" Then Sec80CcgCboTxt = Sec80CcgCboTxt & "|" & DataRead(1)
        'End If
        'Loop
        '  End If
        '*  grd16allow.ColComboList(0) = "|" & AllowCboTxt
        '  grd16otherIncome.ColComboList(0) = "|" & OthIncCbotxt
        '  grd1680c.ColComboList(0) = "|" & Sec80CCboTxt
        '  grd1680CCF.ColComboList(0) = "|" & Sec80CcfCboTxt
        ' grd1680CCG.ColComboList(0) = "|" & Sec80CcgCboTxt
        '  grd16OtherIVA.ColComboList(0) = "|" & Chp6aCboTxt
        Dim cmb As New DataGridViewComboBoxColumn()
        cmb.HeaderText = AllowCboTxt
        cmb.Name = "cmb"
        cmb.MaxDropDownItems = 1
        cmb.Items.Add("allocbotxt")
        grd16allow.Columns.Add(cmb)
    End Sub

    Dim oAllowances As New Collection_Allowances
    Dim oOthIncomes As New Collection_OtherIncomes
    Dim oSec80CDeductions As New Collection_Sec80CDed
    Dim oChapter6ADeductions As New Collection_VI_A_Deductions
    Dim oSec80CCFDeductions As New Collection_Sec80CCFDed
    Dim oSec80CCGDeductions As New Collection_Sec80CCGDed

    Public xMode As String
    ' Dim fso As New FileSystemObject
    'Public Sub grd16allSubTotal()
    'With grd16allow
    '    .AutoSizeMode = flexAutoSizeRowHeight
    '    .AutoSize 1, 1
    '    .SubtotalPosition = flexSTBelow
    ''    .Subtotal flexSTSum, -1, 1, "NN, NN, NN, NN, NNN.00", , vbRed, True
    '    .Subtotal flexSTSum, -1, 1, "###########.00", , vbRed, True
    '    .Redraw = True
    'End With
    'End Sub
    Dim taxcal As Boolean

    Public Sub FillDeducteeCombo(Mode As String)
        Dim rst As New DataSet, opnStr As String
        If Mode = "A" Then
            opnStr = "SELECT * FROM DeductMst WHERE Did Not IN (SELECT dID FROM Form16Details) AND CoID=" & selectedcoid & "  ORDER BY DName "
        Else
            opnStr = "SELECT * FROM DeductMst WHERE CoID=" & selectedcoid & "  ORDER BY DName "
        End If
        rst = FetchDataSet(opnStr)
        'cboDedName.Clear
        cbo16DedName.DataSource = Nothing
        cbo16DedName.Items.Clear()


        For I = 0 To rst.Tables(0).Rows.Count
            cbo16DedName.DataSource = rst.Tables(0)
            cbo16DedName.DisplayMember = "DName"
            cbo16DedName.ValueMember = "DId"
        Next I

        rst.Dispose()

    End Sub

    Private Sub Fill24PRNNo()
        Dim R As Integer
        Dim rst As New DataSet
        rst = FetchDataSet("select * from RetnMst Where CoId = " & selectedcoid & " and  Left(frmtype,2) = '24' ORDER BY FRMTYPE")
        R = 1
        'Do While Not rst.EOF
        For k = 0 To rst.Tables(0).Rows.Count - 1
            grdPRNdet.Rows.Add()
            grdPRNdet.Rows(k).Cells(0).Value = rst.Tables(0).Rows(0)("FrmType").ToString()
            grdPRNdet.Rows(k).Cells(1).Value = IIf(String.IsNullOrEmpty(rst.Tables(0).Rows(0)("NewreceiptNO").ToString()), "", rst.Tables(0).Rows(0)("NewreceiptNO").ToString())
            R = R + 1
            'rst.MoveNext
        Next k
        rst.Dispose()

    End Sub

    Private Sub CopyCoDetails()
        Dim oCompany As New clsCoMst
        oCompany = oCompany.FetchCo(selectedcoid)
        With oCompany
            txt16CoName.Text = .CoName
            txt16CoName.Tag = frmCoMst.txtCoName.Tag
            txt16CoAdd1.Text = .CoAdd1
            txt16CoAdd2.Text = .CoAdd2
            txt16CoAdd3.Text = .CoAdd3
            txt16CoAdd4.Text = .CoAdd4
            txt16CoAdd5.Text = .CoAdd5
            txt16CoPin.Text = .CoPin
            cbo16CoState.Text = frmCoMst.cboCoState.Text
            txt16CoTan.Text = .CoTAN
            txt16CoPan.Text = .CoPAN

        End With
    End Sub

    Private Sub cmbName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo16DedName.SelectedIndexChanged

    End Sub

    'Private Sub cmbName_Validated(cancel As Boolean)
    '    Dim rs As New DataSet
    '    If cbo16DedName.SelectedIndex = -1 Or Trim(cbo16DedName.Text) = "" Then
    '        MsgBox("Please select a deductee first..", vbInformation, "No Deductee Selected")
    '        tabForm16.Enabled = False
    '        cancel = True
    '    Else
    '        tabForm16.Enabled = True
    '        Call frmTDS24Q.FillFrm16DataUsingDID(cbo16DedName.SelectedItem(cbo16DedName.SelectedIndex))
    '    End If
    '    rs = New DataSet
    '    rs = FetchDataSet("SELECT SUM(amtofpayment) FROM Deductee24Q WHERE DID=" & cbo16DedName.SelectedItem(cbo16DedName.SelectedIndex))
    '    txt16TaxableSalary.Text = IIf((rs.Tables(0).Rows(0)("amtofpayment").ToString()), 0, rs.Tables(0).Rows(0)("amtofpayment"))

    'End Sub

    Private Sub tabBasicData_Click(sender As Object, e As EventArgs) Handles tabBasicData.Click

    End Sub

    'Private Sub cmd16delete_Click(sender As Object, e As EventArgs) Handles cmd16delete.Click
    '    Dim sql As String
    '    Dim nds As DataSet
    '    Dim cmd As New OleDbCommand
    '    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
    '        'lvwChallan.Items(0).SubItems(1).Text
    '        sql = "Delete * from F16Challan where f16ID=" & frmTDS24Q.lvwForm16.SelectedItems(0).SubItems(24).Text
    '        cmd.Connection = cn
    '        cmd.CommandText = sql
    '        transaction = cn.BeginTransaction()
    '        cmd.Transaction = transaction
    '        Try

    '            cmd.ExecuteNonQuery()
    '            transaction.Commit()
    '        Catch ex As Exception
    '            transaction.Rollback()
    '            MessageBox.Show(ex.Message)
    '        End Try
    '        cmd.Dispose()
    '        transaction.Dispose()
    '        ' If cn.Errors.Count = 0 Then
    '        sql = "Delete * from Form16MoreDetails where f16ID=" & frmTDS24Q.lvwForm16.SelectedItems(0).SubItems(24).Text
    '        cmd.Connection = cn
    '        cmd.CommandText = sql
    '        transaction = cn.BeginTransaction()
    '        cmd.Transaction = transaction
    '        Try

    '            cmd.ExecuteNonQuery()
    '            transaction.Commit()
    '        Catch ex As Exception
    '            transaction.Rollback()
    '            MsgBox("Unable to Delete right now..", vbInformation)
    '            ' MessageBox.Show(ex.Message)
    '        End Try
    '        cmd.Dispose()
    '        transaction.Dispose()
    '        ' If cn.Errors.Count = 0 Then
    '        sql = "Delete * from Form16details where f16id= " & frmTDS24Q.lvwForm16.SelectedItems(0).SubItems(24).Text
    '        cmd.Connection = cn
    '        cmd.CommandText = sql
    '        transaction = cn.BeginTransaction()
    '        cmd.Transaction = transaction
    '        Try

    '            cmd.ExecuteNonQuery()
    '            transaction.Commit()
    '        Catch ex As Exception
    '            transaction.Rollback()
    '            MsgBox("Unable to Delete right now..", vbInformation)
    '            ' MessageBox.Show(ex.Message)
    '        End Try
    '        cmd.Dispose()
    '        transaction.Dispose()
    '        'If cn..Count = 0 Then
    '        '            cn.BeginTransaction()
    '        '            Me.Close()
    '        '        Else
    '        '            transaction.Rollback()

    '        'End If
    '    Else

    '    End If
    '    ' End If
    '    '  End If
    'End Sub

    Private Sub cmd16Save_Click(sender As Object, e As EventArgs) Handles cmd16Save.Click
        If Val(txt16TaxableSalary.Text) > Val(txt16TaxableSalary.Text) Then
            MsgBox("Taxable Salary is greater than Total Taxable Salary Of All Quarters.")
        End If

        If Val(txt16TaxableSalary.Text) < Val(txt16TaxableSalary.Text) Then
            MsgBox("Taxable Salary is less than Total Taxable Salary Of All Quarters.")
        End If
        CHECKTAXCAL()

        '    If taxcal = True Then
        Check80CTotals(Val(txtgrd80CCal.Text))
        '*  Check80CCFTotals(Val(FRM16Detail.Text))
        Check80CCGTotals(Val(txtgrd80CCGCal.Text))

        If xMode = "A" Then
            If oForm16.Insert(oForm16, oAllowances, oOthIncomes, oSec80CDeductions, oSec80CCFDeductions, oSec80CCGDeductions, oChapter6ADeductions, grd16ManualTax, Val(txt16grosstotPreEmp), Val(txt16TaxPreEmp), IIf(chkHighRate.Checked = True, True, False)) = False Then
                MsgBox("Unable to save data", vbCritical, "ERROR!!")
            Else
                'Data Saved properly...exit this form...and return to main form
                Me.Close()
            End If
        ElseIf xMode = "E" Then
            If oForm16.Update(oForm16, oAllowances, oOthIncomes, oSec80CDeductions, oSec80CCFDeductions, oSec80CCGDeductions, oChapter6ADeductions, grd16ManualTax, Val(txt16grosstotPreEmp), Val(txt16TaxPreEmp), IIf(chkHighRate.Checked = True, True, False)) = False Then
                MsgBox("Unable to save data", vbCritical, "ERROR!!")
            Else
                'Data Saved properly...exit this form...and return to main form
                Me.Close()
            End If
        Else
            MsgBox("Critical Error - xMode not Set n Save Click, Call JAK and report this error!", vbCritical, "FATAL ERROR!")
        End If

        '    End If
        '    taxcal = False
    End Sub
    Private Sub CHECKTAXCAL()
        If txt16Tax.Text <> txtTax.Text Or txt16Surcharge.Text <> txtSurcharge.Text Or txt16EduCess.Text <> txtEd.Text Then
            MsgBox("Form16 Tax calculation not matched with Tax calculation asPer Law")
        End If
    End Sub

    Private Sub Check80CTotals(TotalVal As Double)
        If TotalVal > 150000 Then
            MsgBox("Aggregate amount deductible u/s 80C shall not exceed one lakh Fifty Thousand rupees", 0 + 48, "Caution")
            Exit Sub
        End If

    End Sub

    Private Sub Check80CCFTotals(TotalVal As Double)
        If TotalVal > 20000 Then
            MsgBox("Aggregate amount deductible u/s 80CCF shall not exceed Twenty Thousand", 0 + 48, "Caution")
            Exit Sub
        End If

    End Sub

    Private Sub Check80CCGTotals(TotalVal As Double)
        If TotalVal > 25000 Then
            MsgBox("Aggregate amount deductible u/s 80CCG shall not exceed Twenty five Thousand", 0 + 48, "Caution")
            Exit Sub
        End If
    End Sub

    Private Sub cmdCopyCal_Click(sender As Object, e As EventArgs) Handles cmdCopyCal.Click
        txt16Tax.Text = txtTax.Text
        txt16Surcharge.Text = txtSurcharge.Text
        txt16EduCess.Text = txtEd.Text
        txt16TotalTax.Text = txtTotTax.Text
        txt16Relief.Focus()
    End Sub

    Private Sub cmdDedMast_Click(sender As Object, e As EventArgs) Handles cmdDedMast.Click
        Dim i As Long, DFound As Boolean, OldId As Long
        If cbo16DedName.SelectedIndex < 0 Then Exit Sub
        frmdeduteeTDSMST.Show()
        With frmdeduteeTDSMST
            For i = 0 To .cboDName.SelectedIndex - 1
                If .cboDName.SelectedIndex = cbo16DedName.SelectedItem(cbo16DedName.SelectedIndex) Then
                    .cboDName.SelectedIndex = i
                    OldId = cbo16DedName.SelectedItem(cbo16DedName.SelectedIndex)
                    DFound = True
                    Exit For
                End If
            Next i
        End With
        If DFound = True Then
            frmdeduteeTDSMST.Show()
        End If
        'refill the combo with new data...
        'Call FillDeducteeCombo
        For i = 0 To cbo16DedName.SelectedIndex - 1
            If cbo16DedName.Items.Add(i) = OldId Then
                'select the selection again...
                cbo16DedName.SelectedIndex = i
                '        cbo16DedName.SetFocus
                Exit For
            End If
        Next i
        'cbo16DedName_LostFocus
    End Sub

    Private Sub grd1680c_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles grd1680c.CellEndEdit
        Call CalculateSalaryNTax()
    End Sub
    Public Sub CalculateSalaryNTax()
        Dim R As Long, TotalAllow As Double
        Dim TotalOthInc As Double
        Dim Total80C, TotalChapVIA As Integer
        Dim Total80CCG, Total80CCF As Integer
        'Calculate the Gross Total Salary
        txt16grosstotCurEmp.Text = Val(txt16gross1.Text) + Val(txt16gross2.Text) + Val(txt16gross3.Text) '+ Val(txt16grosstotCurEmp.Text) + Val(txt16grosstotPreEmp.Text)
        txt16grosstotBoth.Text = Val(txt16grosstotCurEmp.Text) + Val(txt16grosstotPreEmp.Text)
        'Calculate total allowance and deduct from Gross Total Salary...to get Balance...
        TotalAllow = 0
        For R = 1 To grd16allow.Rows.Count - 1
            TotalAllow = TotalAllow + grd16allow.Rows(R).Cells(1).Value.ToString()
        Next R
        txtallow.Text = TotalAllow
        txt16bal.Text = Val(txt16grosstotBoth.Text) - TotalAllow
        'Calculate Total Deduction of section 16
        txtTotalDedct.Text = Val(txt16EntAllow.Text) + Val(txt16ProfTax.Text)
        'now, calculate total taxable salary...
        txt16TaxableSalary.Text = Val(txt16bal.Text) - Val(txtTotalDedct.Text)
        'now calculate the sum of other incomes and to taxable salary to get gross total income
        TotalOthInc = 0
        For R = 1 To grd16otherIncome.Rows.Count - 1
            TotalOthInc = TotalOthInc + grd16otherIncome.Rows(R).Cells(1).Value.ToString()
        Next R
        txt16GTI.Text = Val(txt16TaxableSalary.Text) + TotalOthInc
        'now calculate the sum of other incomes and to taxable salary to get gross total income
        Total80C = 0
        For R = 1 To grd1680c.Rows.Count - 1
            Total80C = Total80C + grd1680c.Rows(R).Cells(2).Value.ToString()
        Next R
        Total80CCF = 0
        For R = 1 To grd1680CCF.Rows.Count - 1
            Total80CCF = Total80CCF + grd1680CCF.Rows(R).Cells(2).Value.ToString()
        Next R
        Total80CCG = 0
        For R = 1 To grd1680CCG.Rows.Count - 1
            Total80CCG = Total80CCG + grd1680CCG.Rows(R).Cells(2).Value.ToString()
        Next R
        TotalChapVIA = 0
        For R = 1 To grd16OtherIVA.Rows.Count - 1
            TotalChapVIA = TotalChapVIA + grd16OtherIVA.Rows(R).Cells(3).Value.ToString()
        Next R
        txt1680c.Text = Total80C
        txt1680CCF.Text = Total80CCF
        txt1680CCG.Text = Total80CCG
        txtT16OtherIVA.Text = TotalChapVIA
        txt16otherIncome.Text = TotalOthInc
        txt16OtherIVA.Text = Val(txt1680c.Text) + Val(txtT16OtherIVA.Text) + Val(txt1680CCF.Text) + txt1680CCG.Text
        txt16TotalTaxableIncome.Text = Val(txt16GTI.Text) - Val(txt16OtherIVA.Text)
        txtincome.Text = Val(txt16TotalTaxableIncome.Text)
        'Now Calculate the taxes...
        ' txt16Tax.Text = Val(txt16TaxCurEmp.Text) + Val(txt16TaxPreEmp.Text)
        txt16TotalTax.Text = Val(txt16Tax.Text) + Val(txt16Surcharge.Text) + Val(txt16EduCess.Text)
        txt16NetTax.Text = Val(txt16TotalTax.Text) - Val(txt16Relief.Text)
        txt16totalTDS.Text = Val(txt16TDS1.Text) + Val(txt16TDS2.Text)
        txtPayRef.Text = Val(txt16NetTax.Text) - Val(txt16totalTDS.Text) - Val(txt16TaxPreEmp)
        Check80CTotals(Total80C)
        'UpdateAutoCalcFields
        '   Autotaxcal
    End Sub

    Private Sub txtgrd80CCFCal_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel

    End Sub

    Private Sub grd1680CCF_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles grd1680CCF.CellLeave
        CalculateSalaryNTax()
    End Sub

    Private Sub grd16otherIncome_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grd16otherIncome.CellContentClick

    End Sub

    Private Sub grd1680c_MouseUp(sender As Object, e As MouseEventArgs) Handles grd1680c.MouseUp
        '    If Button = vbRightButton Then
        '        If UCase(Me.ActiveControl.Name) <> "GRD1680CCF" Then grd1680CCF.Focus
        '        grd1680CCF.Select grd1680CCF.MouseRow, grd1680CCF.MouseCol
        '    PopupMenu mnuContext
        'End If
    End Sub

    Private Sub grd1680c_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grd1680c.CellContentClick

    End Sub

    Private Sub grd1680c_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles grd1680c.CellLeave
        CalculateSalaryNTax()
    End Sub

    Private Sub grd1680CCG_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles grd1680CCG.CellLeave
        CalculateSalaryNTax()
    End Sub

    Private Sub grd16ManualTax_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles grd16ManualTax.CellBeginEdit
        Dim i As Integer, col As Integer
        With grd16ManualTax
            If grd16ManualTax.Rows.Count > 0 And grd16ManualTax.Rows(i).Cells(1).Value > 0 Then
                '        If Col = 5 Then
                '            If Len(.TextMatrix(Row, 5)) < 6 Or Len(.TextMatrix(Row, 5)) > 6 Then
                '                MsgBox "Invalid Cheque number!", 0 + 16
                '                .EditText = .TextMatrix(Row, 6)
                '                Exit Sub
                '            End If
                '        End If
                If col = 6 Then
                    If Len(.Rows(i).Cells(6).Value) < 6 Or Len(.Rows(i).Cells(6).Value) > 7 Then
                        MsgBox("Length of Bank BSR Code should be 7 Character.", 0 + 16)
                        '* .EditText = .Rows(i).Cells(6).Value
                        Exit Sub
                    End If
                End If
                If col = 7 Then
                    If Not IsDate(.Rows(i).Cells(7).Value) Then
                        MsgBox("Invalid Date!", 0 + 16)
                        '*.EditText = .Rows(i).Cells(7).Value
                        Exit Sub
                    End If
                    If .Rows(i).Cells(7).Value = "  /  /    " Or (CDate(.Rows(i).Cells(7).Value) < FromDate Or CDate(.Rows(i).Cells(7).Value) > ToDate) Then
                        MsgBox("Invalid Date!", 0 + 16)
                        '.EditText = .TextMatrix(Row, 7)
                        '.EditCell
                        Exit Sub
                    End If
                End If
                grd16ManualTax.Rows(i).Cells(4).Value = grd16ManualTax.Rows(i).Cells(1).Value + grd16ManualTax.Rows(i).Cells(2).Value + grd16ManualTax.Rows(i).Cells(3).Value
            End If
        End With
        If grd16ManualTax.Columns.Count < grd16ManualTax.Columns.Count - 1 Then
            If grd16ManualTax.Columns.Count = 3 Then
                grd16ManualTax.ColumnCount = grd16ManualTax.Columns.Count + 2
            Else
                grd16ManualTax.ColumnCount = grd16ManualTax.Columns.Count + 1
            End If
        End If
        ' *       grd16ManualTax.Redraw = True
        '        grd16ManualTax.Subtotal flexSTSum, -1, 1, "##0", , vbRed, False, "Total:"
        'grd16ManualTax.Subtotal flexSTSum, -1, 2, "##0", , vbRed, False, "Total:"
        'grd16ManualTax.Subtotal flexSTSum, -1, 3, "##0", , vbRed, False, "Total:"
        'grd16ManualTax.Subtotal flexSTSum, -1, 4, "##0", , vbRed, False, "Total:"
        Dim r As Integer = grdchallanDetails.RowCount - 1
        txt16TDS1.Text = grdchallanDetails.Rows(r).Cells(4).Value + grd16ManualTax.Rows(r).Cells(4).Value
    End Sub

    Private Sub txt16EntAllow_TextChanged(sender As Object, e As EventArgs) Handles txt16EntAllow.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub grd16OtherIVA_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles grd16OtherIVA.CellLeave
        CalculateSalaryNTax()
    End Sub

    Private Sub grd16otherIncome_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles grd16otherIncome.CellLeave
        CalculateSalaryNTax()
    End Sub

    Private Sub cbo16DedName_LostFocus(sender As Object, e As EventArgs) Handles cbo16DedName.LostFocus
        Call CtrlLostFocus(cbo16DedName)
        If cbo16DedName.Text = vbNullString Then Exit Sub
        If cbo16DedName.Text <> vbNullString Then
            '    Dim itm As ListItem
            '    Set itm = lvwDeductee.FindItem(cbo16DedName.Text, 1)
            '    If itm Is Nothing Then
            '
            '    Else
            '        MsgBox "Entry of chosen Deductee is already present"
            '        cbo16DedName.ListIndex = -1
            '        Exit Sub
            '    End If
        End If
        Dim rs As New DataSet
        Dim i As Integer, DName As String
        DName = UCase(cbo16DedName.Text)
        rs = FetchDataSet("SELECT * FROM DeductMst WHERE CoId = " & selectedcoid & " And DName='" & DName & "'") ', Cnn
        If rs.Tables(0).Rows.Count <= 0 Then
            'not found., open deductee detail form..
            frmDeducteeTDS.Show()
            frmDeducteeTDS.Frm_typ = 24
            '*frmDeducteeTDS.Move((Me.Left + cbo16DedName.Left) + 100, (Me.Top + cbo16DedName.Top + cbo16DedName.Height + 650))
            frmDeducteeTDS.txtDName.Text = cbo16DedName.Text
            frmDeducteeTDS.optCo.Visible = False
            frmDeducteeTDS.optCo.TabStop = False
            frmDeducteeTDS.Show()
            FillDeducteeCombo(xMode)
            For i = 0 To cbo16DedName.Items.Count - 1
                If cbo16DedName.SelectedItem(i) = DName Then
                    cbo16DedName.SelectedIndex = i
                    Exit Sub
                End If
            Next i
            If i = cbo16DedName.Items.Count Then cbo16DedName.SelectedIndex = -1      'Not Found
        Else
            txt16pan.Text = rs.Tables(0).Rows(0)("DPan").ToString() & ""
            txt16pan.Tag = rs.Tables(0).Rows(0)("DType").ToString()
        End If
        txtdesig.Text = rs.Tables(0).Rows(0)("DDesgn").ToString() & ""
        txtDstatus.Text = rs.Tables(0).Rows(0)("Category").ToString()
        Call Autotaxcal()
        rs.Dispose()
        rs = Nothing
    End Sub

    Private Sub txt16gross1_TextChanged(sender As Object, e As EventArgs) Handles txt16gross1.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16gross2_TextChanged(sender As Object, e As EventArgs) Handles txt16gross2.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16gross3_TextChanged(sender As Object, e As EventArgs) Handles txt16gross3.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub Autotaxcal()
        Dim tax1 As Double : Dim tax2 As Double : Dim tax3 As Double
        Dim tax, i, itax, amt, amt1 As Double
        itax = 0
        amt = 0 : amt1 = 0 : tax = 0

        If txtDstatus.Text = "W" Then
            tax1 = 250000
            tax2 = 500000
            tax3 = 1000000
        ElseIf txtDstatus.Text = "S" Then
            tax1 = 300000
            tax2 = 500000
            tax3 = 1000000
        ElseIf txtDstatus.Text = "G" Then
            tax1 = 250000
            tax2 = 500000
            tax3 = 1000000
        ElseIf txtDstatus.Text = "O" Then
            tax1 = 500000
            tax2 = 1000000
            'tax3 = 800000
        End If

        If Not txtincome.Text = "" Then
            If (txtincome.Text <= tax1) Then
                i = 1
            ElseIf (txtincome.Text <= tax2) Then
                i = 2
            ElseIf (txtincome.Text <= tax3) Then
                i = 3
            Else
                i = 4
            End If
        End If

        amt = txtincome.Text
        While Not i = 0
            Select Case i
                Case 1
                    itax = itax
                Case 2
                    If txtDstatus.Text = "O" Then
                        amt1 = amt - tax1
                        amt = amt - amt1
                        tax = (amt1 * 20) / 100
                        itax = itax + tax
                    Else
                        amt1 = amt - tax1
                        amt = amt - amt1
                        tax = (amt1 * 10) / 100
                        itax = itax + tax
                    End If
                Case 3
                    If txtDstatus.Text = "O" Then
                        amt1 = amt - tax2
                        amt = amt - amt1
                        tax = (amt1 * 30) / 100
                        itax = itax + tax
                    Else
                        amt1 = amt - tax2
                        amt = amt - amt1
                        tax = (amt1 * 20) / 100
                        itax = itax + tax
                    End If
                Case 4
                    If txtDstatus.Text <> "O" Then
                        amt1 = amt - tax3
                        amt = amt - amt1
                        tax = (amt1 * 30) / 100
                        itax = itax + tax
                    End If
            End Select
            i = i - 1
        End While

        If Val(txtincome.Text) <= 500000 Then
            If itax > 5000 Then
                itax = itax - 5000
            Else
                itax = 0
            End If
        End If
        If (txtincome.Text <> "") Then
            If (txtincome.Text > 10000000) Then
                txtSurcharge.Text = Math.Round((itax * 15) / 100, 0)
            Else
                txtSurcharge.Text = 0
            End If
        End If
        txtTax.Text = Math.Round(itax, 0)
        'txtSurcharge = 0
        txtEd.Text = Math.Round(((itax + Val(txtSurcharge)) * 3) / 100, 0)
        txtTotTax.Text = Val(txtTax) + Val(txtSurcharge) + Val(txtEd)
    End Sub

    Private Sub cbo16DedName_GotFocus(sender As Object, e As EventArgs) Handles cbo16DedName.GotFocus
    End Sub

    Private Sub tabForm16_Click(sender As Object, e As EventArgs) Handles tabForm16.Click
        Call Autotaxcal()
    End Sub

    Private Sub txt16EntAllow_Layout(sender As Object, e As LayoutEventArgs) Handles txt16EntAllow.Layout

    End Sub

    Private Sub txt16EntAllow_GotFocus(sender As Object, e As EventArgs) Handles txt16EntAllow.GotFocus
        Call CtrlGotFocus(txt16EntAllow)
    End Sub

    Private Sub txt16EntAllow_LostFocus(sender As Object, e As EventArgs) Handles txt16EntAllow.LostFocus
        Call CtrlLostFocus(txt16EntAllow)
    End Sub

    Private Sub txt16grosstotCurEmp_TextChanged(sender As Object, e As EventArgs) Handles txt16grosstotCurEmp.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16gross3_GotFocus(sender As Object, e As EventArgs) Handles txt16gross3.GotFocus
        Call CtrlGotFocus(txt16gross3)
    End Sub

    Private Sub txt16gross3_LostFocus(sender As Object, e As EventArgs) Handles txt16gross3.LostFocus
        Call CtrlLostFocus(txt16gross3)
    End Sub

    Private Sub txt16grosstotPreEmp_TextChanged(sender As Object, e As EventArgs) Handles txt16grosstotPreEmp.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16gross2_GotFocus(sender As Object, e As EventArgs) Handles txt16gross2.GotFocus
        Call CtrlGotFocus(txt16gross2)
    End Sub

    Private Sub txt16gross2_LostFocus(sender As Object, e As EventArgs) Handles txt16gross2.LostFocus
        Call CtrlLostFocus(txt16gross2)
    End Sub

    Private Sub txt16ProfTax_TextChanged(sender As Object, e As EventArgs) Handles txt16ProfTax.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16gross1_GotFocus(sender As Object, e As EventArgs) Handles txt16gross1.GotFocus
        Call CtrlGotFocus(txt16gross1)
    End Sub

    Private Sub txt16gross1_LostFocus(sender As Object, e As EventArgs) Handles txt16gross1.LostFocus
        Call CtrlLostFocus(txt16gross1)
    End Sub

    Private Sub txt16bal_TextChanged(sender As Object, e As EventArgs) Handles txt16bal.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16TotalTaxableIncome_TextChanged(sender As Object, e As EventArgs) Handles txt16TotalTaxableIncome.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16grosstotCurEmp_LostFocus(sender As Object, e As EventArgs) Handles txt16grosstotCurEmp.LostFocus
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16grosstotCurEmp_GotFocus(sender As Object, e As EventArgs) Handles txt16grosstotCurEmp.GotFocus
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16grosstotPreEmp_GotFocus(sender As Object, e As EventArgs) Handles txt16grosstotPreEmp.GotFocus
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16grosstotPreEmp_LostFocus(sender As Object, e As EventArgs) Handles txt16grosstotPreEmp.LostFocus
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16TaxableSalary_TextChanged(sender As Object, e As EventArgs) Handles txt16TaxableSalary.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16ProfTax_LostFocus(sender As Object, e As EventArgs) Handles txt16ProfTax.LostFocus
        Call CtrlLostFocus(txt16ProfTax)
    End Sub

    Private Sub txt16ProfTax_GotFocus(sender As Object, e As EventArgs) Handles txt16ProfTax.GotFocus
        Call CtrlGotFocus(txt16ProfTax)
    End Sub

    Private Sub txt16GTI_TextChanged(sender As Object, e As EventArgs) Handles txt16GTI.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16TotalTaxableIncome_GotFocus(sender As Object, e As EventArgs) Handles txt16TotalTaxableIncome.GotFocus

    End Sub

    Private Sub txt16TotalTaxableIncome_LostFocus(sender As Object, e As EventArgs) Handles txt16TotalTaxableIncome.LostFocus

    End Sub

    Private Sub txt16OtherIVA_TextChanged(sender As Object, e As EventArgs) Handles txt16OtherIVA.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txtTotalDedct_GotFocus(sender As Object, e As EventArgs) Handles txtTotalDedct.GotFocus
        Call CtrlGotFocus(txtTotalDedct)
    End Sub

    Private Sub txtTotalDedct_LostFocus(sender As Object, e As EventArgs) Handles txtTotalDedct.LostFocus
        Call CtrlLostFocus(txtTotalDedct)
    End Sub

    Private Sub txt16Tax_TextChanged(sender As Object, e As EventArgs) Handles txt16Tax.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16TaxableSalary_LostFocus(sender As Object, e As EventArgs) Handles txt16TaxableSalary.LostFocus
        Call CtrlLostFocus(txt16TaxableSalary)
    End Sub

    Private Sub txt16TaxableSalary_GotFocus(sender As Object, e As EventArgs) Handles txt16TaxableSalary.GotFocus
        Call CtrlGotFocus(txt16TaxableSalary)
    End Sub

    Private Sub txt16Surcharge_TextChanged(sender As Object, e As EventArgs) Handles txt16Surcharge.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16GTI_LostFocus(sender As Object, e As EventArgs) Handles txt16GTI.LostFocus
        Call CtrlLostFocus(txt16GTI)
    End Sub

    Private Sub txt16GTI_GotFocus(sender As Object, e As EventArgs) Handles txt16GTI.GotFocus
        Call CtrlGotFocus(txt16GTI)
    End Sub

    Private Sub txt16EduCess_TextChanged(sender As Object, e As EventArgs) Handles txt16EduCess.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16OtherIVA_LostFocus(sender As Object, e As EventArgs) Handles txt16OtherIVA.LostFocus
        Call CtrlLostFocus(txt16OtherIVA)
    End Sub

    Private Sub txt16OtherIVA_GotFocus(sender As Object, e As EventArgs) Handles txt16OtherIVA.GotFocus
        Call CtrlGotFocus(txt16OtherIVA)
    End Sub

    Private Sub txt16Relief_TextChanged(sender As Object, e As EventArgs) Handles txt16Relief.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16Tax_LostFocus(sender As Object, e As EventArgs) Handles txt16Tax.LostFocus
        Call CtrlLostFocus(txt16Tax)
    End Sub

    Private Sub txt16Tax_GotFocus(sender As Object, e As EventArgs) Handles txt16Tax.GotFocus
        Call CtrlGotFocus(txt16Tax)
    End Sub

    Private Sub txt16TotalTax_TextChanged(sender As Object, e As EventArgs) Handles txt16TotalTax.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16Surcharge_GotFocus(sender As Object, e As EventArgs) Handles txt16Surcharge.GotFocus
        Call CtrlGotFocus(txt16Surcharge)
    End Sub

    Private Sub txt16Surcharge_LostFocus(sender As Object, e As EventArgs) Handles txt16Surcharge.LostFocus
        Call CtrlLostFocus(txt16Surcharge)
    End Sub

    Private Sub txt16EduCess_LostFocus(sender As Object, e As EventArgs) Handles txt16EduCess.LostFocus
        Call CtrlLostFocus(txt16EduCess)
    End Sub

    Private Sub txt16TDS1_TextChanged(sender As Object, e As EventArgs) Handles txt16TDS1.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16EduCess_GotFocus(sender As Object, e As EventArgs) Handles txt16EduCess.GotFocus
        Call CtrlGotFocus(txt16EduCess)
    End Sub

    Private Sub txt16Relief_LostFocus(sender As Object, e As EventArgs) Handles txt16Relief.LostFocus
        Call CtrlLostFocus(txt16Relief)
    End Sub

    Private Sub txt16TDS2_TextChanged(sender As Object, e As EventArgs) Handles txt16TDS2.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16Relief_GotFocus(sender As Object, e As EventArgs) Handles txt16Relief.GotFocus
        Call CtrlGotFocus(txt16Relief)
    End Sub

    Private Sub txt16TotalTax_LostFocus(sender As Object, e As EventArgs) Handles txt16TotalTax.LostFocus
        Call CtrlLostFocus(txt16TotalTax)
    End Sub

    Private Sub txt16NetTax_TextChanged(sender As Object, e As EventArgs) Handles txt16NetTax.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16TotalTax_MouseHover(sender As Object, e As EventArgs) Handles txt16TotalTax.MouseHover

    End Sub

    Private Sub txt16TotalTax_GotFocus(sender As Object, e As EventArgs) Handles txt16TotalTax.GotFocus
        Call CtrlGotFocus(txt16TotalTax)
    End Sub

    Private Sub txt16totalTDS_TextChanged(sender As Object, e As EventArgs) Handles txt16totalTDS.TextChanged
        CalculateSalaryNTax()
    End Sub

    Private Sub txt16TDS1_GotFocus(sender As Object, e As EventArgs) Handles txt16TDS1.GotFocus
        Call CtrlGotFocus(txt16TDS1)
    End Sub

    Private Sub txt16TDS1_LostFocus(sender As Object, e As EventArgs) Handles txt16TDS1.LostFocus
        Call CtrlLostFocus(txt16TDS1)
    End Sub

    Private Sub txt16TDS2_GotFocus(sender As Object, e As EventArgs) Handles txt16TDS2.GotFocus
        Call CtrlGotFocus(txt16TDS2)
    End Sub

    Private Sub txt16TDS2_LostFocus(sender As Object, e As EventArgs) Handles txt16TDS2.LostFocus
        Call CtrlLostFocus(txt16TDS2)
    End Sub

    Private Sub txt16NetTax_LostFocus(sender As Object, e As EventArgs) Handles txt16NetTax.LostFocus
        Call CtrlLostFocus(txt16NetTax)
    End Sub

    Private Sub mnucontext_Click(sender As Object, e As EventArgs) Handles mnucontext.Click
        If UCase(Me.ActiveControl.Name) = "GRD16ALLOW" Then
            If grd16allow.Rows.Count = 0 Then Exit Sub
            grd16allow.Rows.Add()        ', grd16allow.Row
        ElseIf UCase(Me.ActiveControl.Name) = "GRD16OTHERINCOME" Then
            If grd16otherIncome.Rows.Count = 0 Then Exit Sub
            grd16otherIncome.Rows.Add()    ', grd16otherIncome.Row
        ElseIf UCase(Me.ActiveControl.Name) = "GRD16OTHERIVA" Then
            If grd16OtherIVA.Rows.Count = 0 Then Exit Sub
            grd16OtherIVA.Rows.Add()    ', grd16OtherIVA.Row
        ElseIf UCase(Me.ActiveControl.Name) = "GRD1680C" Then
            If grd1680c.Rows.Count = 0 Then Exit Sub
            grd1680c.Rows.Add()   ', grd1680c.Row
        ElseIf UCase(Me.ActiveControl.Name) = "GRD16MANUALTAX" Then
            If grd16ManualTax.Rows.Count = 0 Then Exit Sub
            grd16ManualTax.Rows.Add()  ', grd16ManualTax.Row
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If UCase(Me.ActiveControl.Name) = "GRD16ALLOW" Then
            If grd16allow.Rows.Count = 0 Then Exit Sub
            grd16allow.Rows.Remove(grd16allow.SelectedRows(0))
            If grd16allow.Rows.Count = 1 Then
                grd16allow.Rows.Add()

            End If
        ElseIf UCase(Me.ActiveControl.Name) = "GRD16OTHERINCOME" Then
            If grd16otherIncome.Rows.Count = 0 Then Exit Sub
            grd16otherIncome.Rows.Remove(grd16otherIncome.SelectedRows(0))
            If grd16otherIncome.Rows.Count = 1 Then
                grd16otherIncome.Rows.Add()
            End If
        ElseIf UCase(Me.ActiveControl.Name) = "GRD16OTHERIVA" Then
            If grd16OtherIVA.Rows.Count = 0 Then Exit Sub
            grd16OtherIVA.Rows.Remove(grd16OtherIVA.SelectedRows(0))
            If grd16OtherIVA.Rows.Count = 1 Then
                grd16OtherIVA.Rows.Add()
            End If
        ElseIf UCase(Me.ActiveControl.Name) = "GRD1680C" Then
            If grd1680c.Rows.Count = 0 Then Exit Sub
            grd1680c.Rows.Remove(grd1680c.SelectedRows(0))
            If grd1680c.Rows.Count = 1 Then
                grd1680c.Rows.Add()
            End If
        ElseIf UCase(Me.ActiveControl.Name) = "GRD16MANUALTAX" Then
            If grd16ManualTax.Rows.Count = 0 Or grd16ManualTax.Rows.Count = grd16ManualTax.Rows.Count - 1 Then Exit Sub
            grd16ManualTax.Rows.Remove(grd16ManualTax.SelectedRows(0))
            If grd16ManualTax.Rows.Count = 1 Then
                grd16ManualTax.Rows.Add()
            End If
        ElseIf UCase(Me.ActiveControl.Name) = "GRD1680CCF" Then
            If grd1680CCF.Rows.Count = 0 Then Exit Sub
            grd1680CCF.Rows.Remove(grd1680CCF.SelectedRows(0))
            If grd1680CCF.Rows.Count = 1 Then
                grd1680CCF.Rows.Add()
            End If
        End If
    End Sub

    Private Sub txt16NetTax_GotFocus(sender As Object, e As EventArgs) Handles txt16NetTax.GotFocus
        Call CtrlGotFocus(txt16NetTax)
    End Sub

    Private Sub txt16totalTDS_GotFocus(sender As Object, e As EventArgs) Handles txt16totalTDS.GotFocus
        Call CtrlGotFocus(txt16totalTDS)
    End Sub

    Private Sub txt16totalTDS_LostFocus(sender As Object, e As EventArgs) Handles txt16totalTDS.LostFocus
        Call CtrlLostFocus(txt16totalTDS)
    End Sub

    Private Sub grd16OtherIVA_KeyUp(sender As Object, e As KeyEventArgs) Handles grd16OtherIVA.KeyUp
        If e.KeyCode = Keys.Down Then
            '* If grd16OtherIVA.Row = grd16OtherIVA.Rows.Count - 1 Then
            'If Trim(grd16OtherIVA.TextMatrix(grd16OtherIVA.Rows.Count - 1, 0)) <> vbNullString And Trim(grd16OtherIVA.TextMatrix(grd16OtherIVA.Rows - 1, 3)) <> vbNullString Then
            'grd16OtherIVA.Rows = grd16OtherIVA.Rows.Count + 1
            'grd16OtherIVA.Select grd16OtherIVA.Rows - 1, 0
            '  End If
            'End If
        End If
    End Sub

    Private Sub grd16OtherIVA_MouseUp(sender As Object, e As MouseEventArgs) Handles grd16OtherIVA.MouseUp
        'If e.Button = MouseButtons.Right Then
        '    If UCase(Me.ActiveControl.Name) <> "GRD16OTHERIVA" Then grd16OtherIVA.Focus()
        '    grd16OtherIVA.Select grd16OtherIVA.MouseRow, grd16OtherIVA.MouseCol
        '         popupmenu mnucontext
        '    End If
    End Sub
    'Private Sub oForm16_PrepareDataForSave(Cancel As Boolean)
    '    With oForm16
    '        If xMode = "E" Then
    '            .F16ID = frmTDS24Q.lvwForm16.SelectedItems(0).SubItems(24).Text
    '        End If

    '        .did = cbo16DedName.Items.Add(cbo16DedName.SelectedIndex)
    '        .RetnID = frmTDS24Q.Tag
    '        .DDesgn = txtdesig.Text
    '        .EmpFromDt = txt16FrmDt.Text
    '        .EmpToDt = txt16ToDt.Text
    '        .Gross1 = IIf(Val(txt16gross1.Text) = 0, 0, txt16gross1.Text)
    '        .Gross2 = IIf(Val(txt16gross2.Text) = 0, 0, txt16gross2.Text)
    '        .Gross3 = IIf(Val(txt16gross3.Text) = 0, 0, txt16gross3.Text)
    '        .Sec16ii = IIf(Val(txt16EntAllow.Text) = 0, 0, txt16EntAllow.Text)
    '        .Sec16iii = IIf(Val(txt16ProfTax.Text) = 0, 0, txt16ProfTax.Text)
    '        .TaxAmt = IIf(Val(txt16Tax.Text) = 0, 0, txt16Tax.Text)
    '        .Surcharge = IIf(Val(txt16Surcharge.Text) = 0, 0, txt16Surcharge.Text)
    '        .ECess = IIf(Val(txt16EduCess.Text) = 0, 0, txt16EduCess.Text)
    '        .Relief89 = IIf(Val(txt16Relief.Text) = 0, 0, txt16Relief.Text)
    '        '        .SignByName = IIf(Trim(txtSignByName.Text) = vbNullString, "", txtSignByName.Text)
    '        '        .SignByFatherName = IIf(Trim(txtSignByFatherName.Text) = vbNullString, "", txtSignByFatherName.Text)
    '        '        .SignByCapacity = IIf(Trim(txtSignByCapacity.Text) = vbNullString, "", txtSignByCapacity.Text)
    '        '        .PlaceOfForm = IIf(Trim(txtPlace.Text) = vbNullString, "", txtPlace.Text)
    '        '        .DateOfForm = txtDate.Text
    '        .TDSOnPerks = IIf(Val(txt16TDS2.Text) = 0, 0, txt16TDS2.Text)
    '        .TotalSalaryPreEmp = IIf(Val(txt16grosstotPreEmp.Text) = 0, 0, txt16grosstotPreEmp.Text)
    '        .TDSAmtPreEmp = IIf(Val(txt16TaxPreEmp.Text) = 0, 0, txt16TaxPreEmp.Text)
    '        .HighRatePAN = IIf(chkHighRate.Checked = True, False, True)
    '    End With
    '    'Fill the allowances Collection..
    '    Dim R As Long, c As Long
    '    oAllowances.Clear()

    '    'txt16TaxableSalary.Text = IIf((rs.Tables(0).Rows(0)("amtofpayment").ToString()), 0, rs.Tables(0).Rows(0)("amtofpayment"))
    '    For R = 1 To grd16allow.Rows.Count - 1
    '        '* oAllowances.Add IIf(Val(grd16allow.TextMatrix(R, 2)) = 0, 0, grd16allow.TextMatrix(R, 2)), oForm16.F16ID, "A", grd16allow.TextMatrix(R, 0), Val(grd16allow.TextMatrix(R, 1)), 0, 0

    '    Next R
    '    'Fill the Other Income Collection..
    '    oOthIncomes.Clear()

    '    For R = 1 To grd16otherIncome.Rows.Count - 1
    '        '*   oOthIncomes.Add IIf(Val(grd16otherIncome.TextMatrix(R, 2)) = 0, 0, grd16otherIncome.TextMatrix(R, 2)), oForm16.F16ID, "O", grd16otherIncome.TextMatrix(R, 0), Val(grd16otherIncome.TextMatrix(R, 1)), 0, 0
    '    Next R
    '    'Fill the 80C_deduction Collection..
    '    oSec80CDeductions.Clear()

    '    For R = 1 To grd1680c.Rows.Count - 1
    '        ' *  oSec80CDeductions.Add IIf(Val(grd1680c.TextMatrix(R, 3)) = 0, 0, grd1680c.TextMatrix(R, 3)), oForm16.F16ID, "E", grd1680c.TextMatrix(R, 0), Val(grd1680c.TextMatrix(R, 1)), 0, Val(grd1680c.TextMatrix(R, 2))
    '    Next R

    '    'Fill the 80CCF_deduction Collection..
    '    oSec80CCFDeductions.Clear()

    '    For R = 1 To grd1680CCF.Rows.Count - 1
    '        ' *   oSec80CCFDeductions.Add IIf(Val(grd1680CCF.TextMatrix(R, 3)) = 0, 0, grd1680CCF.TextMatrix(R, 3)), "E", grd1680CCF.TextMatrix(R, 0), Val(grd1680CCF.TextMatrix(R, 1)), 0, Val(grd1680CCF.TextMatrix(R, 2)), oForm16.F16ID
    '    Next R

    '    'Fill the 80CCG_deduction Collection..
    '    oSec80CCGDeductions.Clear()

    '    For R = 1 To grd1680CCG.Rows.Count - 1
    '        ' *   oSec80CCGDeductions.Add IIf(Val(grd1680CCG.TextMatrix(R, 3)) = 0, 0, grd1680CCG.TextMatrix(R, 3)), oForm16.F16ID, "G", grd1680CCG.TextMatrix(R, 0), Val(grd1680CCG.TextMatrix(R, 1)), 0, Val(grd1680CCG.TextMatrix(R, 2))
    '    Next R

    '    'Fill the Other VI-A_deduction Collection..
    '    R = 1
    '    oChapter6ADeductions.Clear()

    '    For R = 1 To grd16OtherIVA.Rows.Count - 1
    '        ' * oChapter6ADeductions.Add IIf(Val(grd16OtherIVA.TextMatrix(R, 4)) = 0, 0, grd16OtherIVA.TextMatrix(R, 4)), oForm16.F16ID, "V", grd16OtherIVA.TextMatrix(R, 0), Val(grd16OtherIVA.TextMatrix(R, 1)), Val(grd16OtherIVA.TextMatrix(R, 2)), Val(grd16OtherIVA.TextMatrix(R, 3))
    '    Next R
    'End Sub

End Class