Imports System
Imports System.IO
Imports System.Globalization
Imports System.Windows.Forms

Public Class frmform16Parametre
    ' Dim fso As New , 
    Dim coMst As New clsCoMst
    Dim ParaFileName As String
    Dim ReadStream As StreamWriter
    Dim ReadStream1 As StreamReader
    Dim DefaultAllow, DefaultOthInc, Default80C, DefaultChp6A, Default80CCF, Default80CCG

    Private Sub frmform16Parametre_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub txtAllowAdd_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAllowances.Leave
        txtAllowances.BackColor = Color.White
    End Sub

    Private Sub txtAllowAdd_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAllowances.Enter
        txtAllowances.BackColor = Color.LightYellow
    End Sub

    Private Sub txtOthAdd_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOthInc.Leave
        txtOthInc.BackColor = Color.White
    End Sub

    Private Sub txtOthAdd_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOthInc.Enter
        txtOthInc.BackColor = Color.LightYellow
    End Sub

    Private Sub txt80CCE_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt80CCE.Leave
        txt80CCE.BackColor = Color.White
    End Sub

    Private Sub txt80CCE_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt80CCE.Enter
        txt80CCE.BackColor = Color.LightYellow
    End Sub

    Private Sub txt80CCF_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt80CCF.Leave
        txt80CCF.BackColor = Color.White
    End Sub

    Private Sub txt80CCF_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt80CCF.Enter
        txt80CCF.BackColor = Color.LightYellow
    End Sub

    Private Sub txt80CCG_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt80CCG.Leave
        txt80CCG.BackColor = Color.White
    End Sub

    Private Sub txt80CCG_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt80CCG.Enter
        txt80CCG.BackColor = Color.LightYellow
    End Sub

    Private Sub TextBox6_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtChap6a.Enter
        txtChap6a.BackColor = Color.LightYellow
    End Sub

    Private Sub TextBox6_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtChap6a.Leave
        txtChap6a.BackColor = Color.White
    End Sub

    Private Sub SetDefaultValues()
        DefaultAllow = Split("HRA,Travelling Allowance,Dearness Allowance,Daily Allowance,Conveyance Allowance,Uniform Allowance,Child. Edu. Allowance", ",")
        DefaultOthInc = Split("PPF Interest,NSC Interest,Bank Interest,Rental Income,Hsg. Loan Intt.", ",")
        Default80C = Split("LIP,Deferred Annuity,PF,PPF,Super Annuation Fund,ULIP,NSC,NSC Intt,ELSS(MF),Specified Units/Bonds,Housing Loan Repayment,Tution Fees", ",")
        Default80CCF = Split("Infrastructure Bonds", ",")
        Default80CCG = Split("RG-Equity Savings Scheme", ",")
        DefaultChp6A = Split("80D-Mediclaim,80DD-Dependent Disabiltiy,80E-Edu. Loan,80G-Donation,80GG-Rent Paid,80U-Disability,80CCD(1)-NPS", ",")

    End Sub

    Private Sub frmform16Parametre_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Long
        ParaFileName = Application.StartupPath & "\Database\Form16Parameters.txt"
        If File.Exists(ParaFileName) Then
            ReadStream1 = File.OpenText(ParaFileName)
            'FileOpen.TextFile(ParaFileName)
        Else
            'Get Default Values...and save it to text file..
            SetDefaultValues()
            ReadStream = File.CreateText(ParaFileName)
            For i = 0 To UBound(DefaultAllow)
                ReadStream.WriteLine("A," & DefaultAllow(i) & ",T")
            Next i
            For i = 0 To UBound(DefaultOthInc)
                ReadStream.WriteLine("O," & DefaultOthInc(i) & ",T")
            Next i
            For i = 0 To UBound(Default80C)
                ReadStream.WriteLine("E," & Default80C(i) & ",T")
            Next i
            For i = 0 To UBound(DefaultChp6A)
                ReadStream.WriteLine("V," & DefaultChp6A(i) & ",T")
            Next i
            For i = 0 To UBound(Default80CCF)
                ReadStream.WriteLine("F," & Default80CCF(i) & ", T")
            Next i
            For i = 0 To UBound(Default80CCG)
                ReadStream.WriteLine("G," & Default80CCG(i) & ", T")
            Next i
            ReadStream1.Close()
            ReadStream1 = File.OpenText(ParaFileName)
        End If
        FillDataInGrid()
    End Sub

    Private Sub FillDataInGrid()
        Dim ReadData, a1, lvwitm As New ListViewItem
        'Dim arow As DataRow
        'Dim tbl As New DataTable("mytab")
        'tbl.Columns.Add("col1", GetType(String))
        lvwAllowances.Items.Clear()
        'Dim filereader As System.IO.StreamReader
        'If Not stringreader.StartsWith("A") Then
        '    arow = tbl.NewRow
        '    arow(0) = stringreader
        '    tbl.Rows.Add(arow)
        'End If
        'lvwitm = tbl.Rows
        ' MsgBox("firstline", stringreader)
        ReadStream1 = My.Computer.FileSystem.OpenTextFileReader(ParaFileName)
        Do While Not ReadStream1.EndOfStream
            Dim stringreader() As String
            stringreader = Split(ReadStream1.ReadLine, ",")
            ReadData.Text = stringreader(0)
            ' lvwAllowances.Items(0).Text = stringreader(0)
            'lvwAllowances.Items.Add(ReadData)
            If ReadData.SubItems(0).Text = "A" Then
                lvwAllowancesHead()
                lvwitm = lvwAllowances.Items.Add(stringreader(1)) '.Selected(1))
                'lvwAllowances.Columns.Add("lvwitm")
                lvwitm.Checked = IIf(stringreader(2) = "T", True, False)
            ElseIf ReadData.SubItems(0).Text = "O" Then
                lvw16otherIncomeHead()
                lvwitm = lvw16otherIncome.Items.Add(stringreader(1)) '(1))
                lvwitm.Checked = IIf(stringreader(2) = "T", True, False)
            ElseIf ReadData.SubItems(0).Text = "E" Then
                lvw1680cHead()
                lvwitm = lvw1680c.Items.Add(stringreader(1))
                lvwitm.Checked = IIf(stringreader(2) = "T", True, False)
            ElseIf ReadData.SubItems(0).Text = "V" Then
                lvw16OtherIVAHead()
                lvwitm = lvw16OtherIVA.Items.Add(stringreader(1)) '(1))
                lvwitm.Checked = IIf(stringreader(2) = "T", True, False)
            ElseIf ReadData.SubItems(0).Text = "F" Then
                lvw1680CCFHead()
                lvwitm = lvw1680CCF.Items.Add(stringreader(1)) '(1))
                lvwitm.Checked = IIf(stringreader(2) = "T", True, False)
            ElseIf ReadData.SubItems(0).Text = "G" Then
                lvw1680CCGHead()
                lvwitm = lvw1680CCG.Items.Add(stringreader(1)) '(1))
                lvwitm.Checked = IIf(stringreader(2) = "T", True, False)
            End If
        Loop
        ReadStream1.Close()
    End Sub

    Public Sub lvwAllowancesHead()
        With lvwAllowances
            .Columns.Clear()
            .Columns.Add("Parameter Name", 250, HorizontalAlignment.Left)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
        End With
    End Sub

    Public Sub lvw16otherIncomeHead()
        With lvw16otherIncome
            .Columns.Clear()
            .Columns.Add("Parameter Name", 250, HorizontalAlignment.Left)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
        End With
    End Sub

    Public Sub lvw16OtherIVAHead()
        With lvw16OtherIVA
            .Columns.Clear()
            .Columns.Add("Parameter Name", 250, HorizontalAlignment.Left)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
        End With
    End Sub

    Public Sub lvw1680CCFHead()
        With lvw1680CCF
            .Columns.Clear()
            .Columns.Add("Parameter Name", 250, HorizontalAlignment.Left)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
        End With
    End Sub

    Public Sub lvw1680cHead()
        With lvw1680c
            .Columns.Clear()
            .Columns.Add("Parameter Name", 250, HorizontalAlignment.Left)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
        End With
    End Sub

    Public Sub lvw1680CCGHead()
        With lvw1680CCG
            .Columns.Clear()
            .Columns.Add("Parameter Name", 250, HorizontalAlignment.Left)
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
        End With
    End Sub

    Private Sub cmdAddAllowance_Click(sender As Object, e As EventArgs) Handles cmdAddAllowance.Click
        Dim lvwitm As ListViewItem
        If Not Trim(txtAllowances.Text) = vbNullString Then
            'lvwAllowances.ListItems.Add , , txtAllowances.Text
            lvwitm = lvwAllowances.Items.Add(txtAllowances.Text)
            lvwitm.Checked = True
            txtAllowances.Text = vbNullString
        End If
        lvwAllowances.Refresh()
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Close()
    End Sub

    Private Sub cmdAddOthInc_Click(sender As Object, e As EventArgs) Handles cmdAddOthInc.Click
        Dim lvwitm As ListViewItem
        If Not Trim(txtOthInc.Text) = vbNullString Then
            'lvw16otherIncome.ListItems.Add , , txtOthInc.Text
            lvwitm = lvw16otherIncome.Items.Add(txtOthInc.Text)
            lvwitm.Checked = True
            txtOthInc.Text = vbNullString
        End If
        lvw16otherIncome.Refresh()
    End Sub

    Private Sub cmdAdd80CCE_Click(sender As Object, e As EventArgs) Handles cmdAdd80CCE.Click
        Dim lvwitm As ListViewItem
        If Not Trim(txt80CCE.Text) = vbNullString Then
            lvwitm = lvw1680c.Items.Add(txt80CCE.Text)
            lvwitm.Checked = True
            txt80CCE.Text = vbNullString
        End If
        lvw1680c.Refresh()
    End Sub

    Private Sub cmdAdd80CCF_Click(sender As Object, e As EventArgs) Handles cmdAdd80CCF.Click
        Dim lvwitm As ListViewItem
        If Not Trim(txt80CCF.Text) = vbNullString Then
            lvwitm = lvw1680CCF.Items.Add(txt80CCF.Text)
            lvwitm.Checked = True
            txt80CCF.Text = vbNullString
        End If
        lvw1680CCF.Refresh()
    End Sub

    Private Sub cmdAdd80CCG_Click(sender As Object, e As EventArgs) Handles cmdAdd80CCG.Click
        Dim lvwitm As ListViewItem
        If Not Trim(txt80CCG.Text) = vbNullString Then
            lvwitm = lvw1680CCG.Items.Add(txt80CCG.Text)
            lvwitm.Checked = True
            txt80CCG.Text = vbNullString
        End If
        lvw1680CCG.Refresh()
    End Sub

    Private Sub cmdChp6aAdd_Click(sender As Object, e As EventArgs) Handles cmdChp6aAdd.Click
        Dim lvwitm As ListViewItem
        If Not Trim(txtChap6a.Text) = vbNullString Then
            lvwitm = lvw16OtherIVA.Items.Add(txtChap6a.Text)
            lvwitm.Checked = True
            txtChap6a.Text = vbNullString
        End If
        lvw16OtherIVA.Refresh()
    End Sub
End Class