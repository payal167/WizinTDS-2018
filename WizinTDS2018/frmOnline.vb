Imports System.IO
Imports System.Data.OleDb
Imports System.Data
Imports System.ComponentModel

Public Class frmOnline
    Public paths As String
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub Optdata_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Optdata.CheckedChanged
        cmbcomp.Enabled = True
    End Sub

    Private Sub Opttext_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Opttext.CheckedChanged
        cmbcomp.Enabled = False
    End Sub

    Private Sub frmOnline_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '   Dim rs As New DataSet
        'cbotype.SelectedIndex = 0
        ' cboFormNo.SelectedIndex = 0
        ' cbOqtr.ListIndex = 0

        '   cboFormNo.Items.Clear()

        cboFormNo.Items.Add("24Q")
        cboFormNo.Items.Add("26Q")
        cboFormNo.Items.Add("27EQ")
        cboFormNo.Items.Add("27Q")
        cboFormNo.SelectedIndex = 0

        cbOqtr.Items.Clear()
        cbOqtr.Items.Add("Q1")
        cbOqtr.Items.Add("Q2")
        cbOqtr.Items.Add("Q3")
        cbOqtr.Items.Add("Q4")
        cbOqtr.SelectedIndex = 0

        ' txtpathname.Text = "" & My.Application.Info.DirectoryPath & "\WizinTDS.mdb"
        txtpathname.Text = Application.StartupPath & "\Database\WizinTDS.mdb "
        cmpload(txtpathname.Text)

        cmbfinyr.Text = "2018-2019"
        txtpathname.Enabled = True
    End Sub
    Private Sub cmpload(tpath As String)
        Dim rst As New DataSet
        'rst.Dispose()
        'rst = Nothing
        'cn.Dispose()
        'cn = Nothing

        ' cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtpathname.Text & "; Persist Security Info=False;Jet OLEDB:Database Password='apr01'"
        ' cn.CursorLocation = adUseClient
        '  cn.Open()
        cmbcomp.Items.Clear()
        rst = FetchDataSet("SELECT * FROM CoMst ORDER BY CoName")
        If rst.Tables(0).Rows.Count > 0 Then
            cmbcomp.Items.Add("Select")
            cmbcomp.SelectedIndex = -1
            Dim i As Integer
            For i = 0 To rst.Tables(0).Rows.Count - 1
                cmbcomp.Items.Add(rst.Tables(0).Rows(i)(1).ToString())
            Next
        End If
        'cmbcomp.DataSource = rst.Tables(0)
        'cmbcomp.DisplayMember = "Coname"
        '    cmbcomp.ValueMember = "coid"

        'End If
        If cmbcomp.Items.Count > 0 Then
            cmbcomp.SelectedIndex = 0
            txtpathname.Enabled = False
        Else
            cmbcomp.SelectedIndex = -1
            txtpathname.Enabled = True
        End If
        rst.Dispose()
        rst = Nothing
    End Sub
    Private Sub cmdSel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsel.Click
        Dim str As String
        Dim pos As Integer
        Dim frm As New frmImport
        str = Application.StartupPath
        pos = InStr(str, "WizinTDS" & FY.Substring(FY.Length - 4))

        If Directory.Exists(str.Substring(0, pos - 1) & "WizinTDS" & (FY.Substring(FY.Length - 4)) - 1 & "\Database") Then
            OpendgTDS.InitialDirectory = str.Substring(0, pos - 1) & "WizinTDS" & (FY.Substring(FY.Length - 4)) - 1 & "\Database"
        Else
            'OpendgTDS.InitialDirectory = str.Substring(0, pos - 1) & "WizinTDS" & (FY.Substring(FY.Length - 4)) & "\Database"
            OpendgTDS.InitialDirectory = Application.StartupPath & "\Database"
        End If

        OpendgTDS.FileName = "WizinTDS.mdb"
        Dim i = OpendgTDS.ShowDialog()
        If i = 2 Then

            Exit Sub
        End If

    End Sub

    Private Sub frmOnline_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub txtSel_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpathname.Enter
        txtpathname.BackColor = Color.LightYellow
    End Sub

    Private Sub txtSel_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpathname.Leave
        txtpathname.BackColor = Color.White
    End Sub

    Private Sub cmbSelComp_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbcomp.Leave
        cmbcomp.BackColor = Color.White
    End Sub

    Private Sub cmbSelComp_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbcomp.Enter
        cmbcomp.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbtype_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbotype.Leave
        cbotype.BackColor = Color.White
    End Sub

    Private Sub cmbtype_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbotype.Enter
        cbotype.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbForm_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormNo.Leave
        cboFormNo.BackColor = Color.White
    End Sub

    Private Sub cmbForm_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormNo.Enter
        cboFormNo.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbQuat_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOqtr.Leave
        cbOqtr.BackColor = Color.White
    End Sub

    Private Sub cmbQuat_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOqtr.Enter
        cbOqtr.BackColor = Color.LightYellow
    End Sub

    Private Sub txtId_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtuserid.Leave
        txtuserid.BackColor = Color.White
    End Sub

    Private Sub txtId_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtuserid.Enter
        txtuserid.BackColor = Color.LightYellow
    End Sub

    Private Sub txtPw_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpassword.Leave
        txtpassword.BackColor = Color.White
    End Sub

    Private Sub txtPw_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpassword.Enter
        txtpassword.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDedTAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTAN.Enter
        txtTAN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDedTAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTAN.Leave
        txtTAN.BackColor = Color.White
    End Sub

    Private Sub cmbAssYear_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbfinyr.Leave
        cmbfinyr.BackColor = Color.White
    End Sub

    Private Sub cmbAssYear_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbfinyr.Enter
        cmbfinyr.BackColor = Color.LightYellow
    End Sub

    Private Sub txtPRN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPRN.Enter
        txtPRN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtPRN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPRN.Leave
        txtPRN.BackColor = Color.White
    End Sub

    Private Sub cmdGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerate.Click
        Dim Qtr As String
        Dim rs As New DataSet
        Dim typ As String
        If Optdata.Checked = True Then
            If cbotype.SelectedIndex = 0 Then
                typ = "CS"
            ElseIf cbotype.SelectedIndex = 1 Then
                typ = "CR"
            ElseIf cbotype.SelectedIndex = 2 Then
                typ = "JU"
            Else
                typ = "CA"
            End If
            Qtr = cbOqtr.Text
            rs = FetchDataSet(" Select retnId from retnmst where coid= " & cmbcomp.Items(cmbcomp.SelectedIndex) & " and frmtype='" & cboFormNo.Text & Strings.Right(cbOqtr.Text, 1) & "'")
            If rs.Tables(0).Rows.Count > 0 Then
                Shell(Application.StartupPath & "\Online.exe " & cmbcomp.Items(cmbcomp.SelectedIndex) & " " & rs.Tables(0).Rows(0)(0) & " " & typ)
            Else
                MsgBox("Form No. " & cboFormNo.Text & Strings.Right(cbOqtr.Text, 1) & " For this company Does not Exist")
            End If
        Else
            If cbotype.SelectedIndex = 0 Then
                typ = "CS" ' conso file
            ElseIf cbotype.SelectedIndex = 1 Then
                typ = "CR" ' form 16 download
            ElseIf cbotype.SelectedIndex = 2 Then
                typ = "JU" ' justification
            Else
                typ = "CA" 'form 16A download
            End If
            Shell(Application.StartupPath & "\OnlineT.exe " & txtuserid.Text & ";" & txtpassword.Text & ";" & txtPRN.Text & ";" & txtpathname.Text & ";" & typ)
        End If

    End Sub

    Private Sub OpendgTDS_FileOk(sender As Object, e As CancelEventArgs) Handles OpendgTDS.FileOk

    End Sub

    Private Sub cbOqtr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbOqtr.SelectedIndexChanged

    End Sub

    Private Sub cbOqtr_Click(sender As Object, e As EventArgs) Handles cbOqtr.Click
        If cmbcomp.SelectedIndex > 0 Then
            OthDetail()
            AYear()
        End If
    End Sub
    Private Sub AYear()
        Dim rs As New DataSet
        rs = FetchDataSet("SELECT distinct RetnMst.Ayear FROM RetnMst where (((RetnMst.CoID)=" & cmbcomp.SelectedIndex & "))")
        cmbfinyr.Items.Clear()

        If rs.Tables(0).Rows.Count > 0 Then
            Dim i As Integer
            For i = 0 To rs.Tables(0).Rows.Count - 1
                cmbfinyr.Items.Add((rs.Tables(0).Rows(i)("AYear")) & "")
            Next
            ' cmbfinyr.SelectedIndex = -1
        End If
        rs.Dispose()
        rs = Nothing
    End Sub
    Private Sub OthDetail()
        Dim rs2 As New DataSet
        Dim rs1 As New DataSet

        rs2 = FetchDataSet("SELECT CoMst.CoID, CoMst.CoTAN, CoMst.TanUserID, CoMst.TANPAssword From coMst WHERE (CoMst.CoID=" & cmbcomp.SelectedIndex & ")")

        If String.IsNullOrEmpty(rs2.Tables(0).Rows.Count) Then
            txtuserid = rs2.Tables(0).Rows(0)("TanUserID") & ""
            txtpassword = rs2.Tables(0).Rows(0)("TANPAssword") & ""
            txtTAN = rs2.Tables(0).Rows(0)("CoTAN") & ""
        End If


        rs1 = FetchDataSet("SELECT distinct RetnMst.PRN FROM RetnMst where (((RetnMst.CoID)=" & cmbcomp.SelectedIndex & ")) and RetnMst.FrmType= '" & cboFormNo.Text & Strings.Right(cbOqtr.Text, 1) & "'")

        If String.IsNullOrEmpty(rs1.Tables(0).Rows.Count) Then
            txtPRN = rs1.Tables(0).Rows(0)("prn") & ""
        End If
        rs2.Dispose()
        rs2 = Nothing
        rs1.Dispose()
        rs1 = Nothing
    End Sub

    Private Sub cbotype_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbotype.SelectedIndexChanged

    End Sub

    Private Sub cbotype_Click(sender As Object, e As EventArgs) Handles cbotype.Click
        If cmbcomp.SelectedIndex > -1 Then
            OthDetail()
            ' Qtr
            AYear()
        End If
    End Sub

    Private Sub cmbcomp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbcomp.SelectedIndexChanged

    End Sub

    Private Sub Qtr()
        Dim rs As New DataSet
        Dim sql As String

        rs = FetchDataSet("SELECT distinct RetnMst.FrmType FROM RetnMst") ' where COId = " & cmbCNm.ItemData(cmbCNm.ListIndex)
        If cboFormNo.SelectedIndex = 0 Then
            sql = sql & " where RetnMst.frmtype in ('24Q1','24Q2','24Q3','24Q4')"

        ElseIf cboFormNo.SelectedIndex = 1 Then
            sql = sql & " where RetnMst.frmtype in ('26Q1','26Q2','26Q3','26Q4')"
        ElseIf cboFormNo.SelectedIndex = 2 Then
            sql = sql & " where RetnMst.frmtype in ('27EQ1','27EQ2','27EQ3','27EQ4')"
        ElseIf cboFormNo.SelectedIndex = 3 Then
            sql = sql & " where RetnMst.frmtype in ('27Q1','27Q2','27Q3','27Q4')"
        End If
        '  rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
        If rs.Tables(0).Rows.Count > 0 Then
            cbOqtr.Items.Clear()

            While Not rs.Tables(0).Rows.Count
                cbOqtr.Items.Add(rs.Tables(0).Rows(0)("FrmType") & "")
            End While
            cbOqtr.SelectedIndex = 0
        End If
        '  If rs.State = adStateOpen Then 
        rs.Dispose()
        rs = Nothing
    End Sub

    Private Sub cmbcomp_Click(sender As Object, e As EventArgs) Handles cmbcomp.Click
        If cmbcomp.SelectedIndex > -1 Then
            OthDetail()
            ' Qtr
            AYear()
        End If
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
End Class