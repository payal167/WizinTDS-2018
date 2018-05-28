Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Data
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class frmDeducteeTDS
    Dim cmd As OleDbCommand
    Dim headadaptor As New OleDbDataAdapter
    Dim headcommand As New OleDbCommandBuilder
    Dim ds As New System.Data.DataSet
    Dim dr As OleDbDataReader
    Dim dr2 As OleDbDataReader
    Dim dr3 As OleDbDataReader
    Public Frm_typ As String
    Dim WithEvents oDed As ClsDeductMstObj


    Private Sub frmDeducteeTDS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'txtDAdd1.Select()
        oDed = New ClsDeductMstObj


        Fill_cmbState()
        cbocat.Items.Clear()
        cbocat.Items.Add("VALID PAN")
        ' cbocat.SelectedValue = 0
        cbocat.SelectedIndex = 0
        cbocat.Items.Add("PANAPPLIED")
        'cbocat.SelectedValue = 1
        cbocat.SelectedIndex = 1
        cbocat.Items.Add("PANINVALID")
        ' cbocat.SelectedValue = 2
        cbocat.SelectedIndex = 2
        cbocat.Items.Add("PANNOTAVBL")
        ' cbocat.SelectedValue = 3
        cbocat.SelectedIndex = 3
        cbocat.SelectedIndex = 0
        'Category new inserted by nitin on 07/06/2006
        cboCategory.Items.Clear()
        cboCategory.Items.Add("G - General/Other")
        cboCategory.Items.Add("W - Woman Assessee")
        cboCategory.Items.Add("S - Senior Citizen")
        cboCategory.Items.Add("O - Super Senior Citizen")
        cboCategory.SelectedIndex = 0

        'Changes for FVU 3.0 - FY 10-11 onwards - done by Nitin Betharia
        'txtref.Enabled = False

        'Changes for FVU 5.7 - FY 10-11 onwards
        CboCollNonRes.Items.Clear()
        CboCollNonRes.Items.Add("Yes")
        CboCollNonRes.Items.Add("No")
        CboPerEstInd.Items.Add("Yes")
        CboPerEstInd.Items.Add("No")
    End Sub
    Private Sub frmDeducteeTDS_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub FillDeducteeCombo26()
        Dim nds As New DataSet
        nds = FetchDataSet("select DName,DId from DeductMst Where CoId = " & selectedcoid & "  ORDER BY DName ")
        frmTDS26Q.cboDedName.DataSource = Nothing
        frmTDS26Q.cboDedName.Items().Clear()
        If nds.Tables(0).Rows.Count > 0 Then
            frmTDS26Q.cboDedName.DataSource = nds.Tables(0)
            frmTDS26Q.cboDedName.DisplayMember = "DName"
            frmTDS26Q.cboDedName.ValueMember = "DId"
        End If

        nds.Dispose()
    End Sub

    Private Sub FillDeducteeCombo27()
        Dim nds As New DataSet
        nds = FetchDataSet("select DName,DId from DeductMst Where CoId = " & selectedcoid & "  ORDER BY DName ")
        frmTDS27Q.cboDedName.DataSource = Nothing
        frmTDS27Q.cboDedName.Items().Clear()
        If nds.Tables(0).Rows.Count > 0 Then
            frmTDS27Q.cboDedName.DataSource = nds.Tables(0)
            frmTDS27Q.cboDedName.DisplayMember = "DName"
            frmTDS27Q.cboDedName.ValueMember = "DId"
        End If

        nds.Dispose()
    End Sub

    Private Sub FillDeducteeCombo27E()
        Dim nds As New DataSet
        nds = FetchDataSet("select DName,DId from DeductMst Where CoId = " & selectedcoid & "  ORDER BY DName ")
        frmTDS27EQ.cboDedName.DataSource = Nothing
        frmTDS27EQ.cboDedName.Items().Clear()
        If nds.Tables(0).Rows.Count > 0 Then
            frmTDS27EQ.cboDedName.DataSource = nds.Tables(0)
            frmTDS27EQ.cboDedName.DisplayMember = "DName"
            frmTDS27EQ.cboDedName.ValueMember = "DId"
        End If

        nds.Dispose()
    End Sub

    Private Sub FillDeducteeCombo24()
        Dim nds As New DataSet
        nds = FetchDataSet("select DName,DId from DeductMst Where CoId = " & selectedcoid & "  ORDER BY DName ")
        frmTDS24Q.cboDedName.DataSource = Nothing
        frmTDS24Q.cboDedName.Items().Clear()
        If nds.Tables(0).Rows.Count > 0 Then
            frmTDS24Q.cboDedName.DataSource = nds.Tables(0)
            frmTDS24Q.cboDedName.DisplayMember = "DName"
            frmTDS24Q.cboDedName.ValueMember = "DId"
        End If

        nds.Dispose()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        'Dim _fontJIAdieuxRegEx As String = "^([a-zA-Z]){5}([0-9]){4}([a-zA-Z]){1}?$"
        'Dim r As New Regex(_fontJIAdieuxRegEx)
        'If r.IsMatch(txtPAN.Text) = False Then
        '    If r.IsMatch(txtPAN.Text) = False Then
        '        ErrorProvider1.SetError(txtPAN, "Invalid PAN")
        '    End If
        'Else
        If BeforeSave() = False Then Exit Sub

        If cbocat.Text <> "VALID PAN" And txtDeEmail.Text = vbNullString And txtDePhone.Text = vbNullString And txtTIN.Text = vbNullString Then
            MsgBox("Record Not Saved... Please Enter EmailId, Contact Number And TIN/UIN of the Deductee")
            Exit Sub
        End If
        If oDed.Insert(oDed) = True Then
            Select Case Frm_typ
                Case "26"
                    Dim nds As New DataSet

                    frmTDS26Q.txtDedPAN.Text = txtDPAN.Text
                    frmTDS26Q.txtDedPAN.Tag = IIf(optCo.Checked = True, "C", "O")
                    'FillDeducteeCombo26()


                    frmTDS26Q.cboDedName.SelectedValue = oDed.did
                    frmTDS26Q.did = oDed.did
                Case "27"
                    frmTDS27Q.txtDedPAN.Text = txtDPAN.Text
                Case "27E"
                    frmTDS27EQ.txtDedPAN.Text = txtDPAN.Text

                Case "24"
                    frmTDS24Q.txtDedPAN.Text = txtDPAN.Text
                    frmTDS24Q.txtDedPAN.Tag = IIf(optCo.Checked = True, "C", "O")

                    'FillDeducteeCombo24()
                    frmTDS24Q.cboDedName.SelectedValue = oDed.did
                    frmTDS24Q.did = oDed.did
            End Select
            Me.Dispose()

        Else
            'stay with this form only..unless user presses cancel.
        End If

        'If BeforeSave() = False Then Exit Sub
        'Dim sql As String = "SELECT TOP 1 DId FROM DeductMst ORDER BY Did DESC"
        'Dim cmd1 As New OleDbCommand(sql, cn)
        'Dim id As Integer

        'cmd1.ExecuteNonQuery()
        'dr3 = cmd1.ExecuteReader
        'While dr3.Read()
        '    id = dr3(0) + 1
        'End While

        ''Dim sqlQry As String = "INSERT INTO [DeductMst] ([DName],[DPan],[DAdd1],[DAdd2],[DAdd3],[DAdd4],[DAdd5],[DState],[DPin],[DType],[DPANRef],[DPANCat],[Category],[DDesgn]) VALUES (@DName,@DPan,@DAdd1,@DAdd2,@DAdd3,@DAdd4,,@DAdd5,@DState,@DPin,@DType,@DPANRef,@DPANCat,@Category,@DDesgn)"
        'Dim sqlQry As String = "INSERT INTO [DeductMst] (DId,CoID,DName,DPan,DAdd1,DAdd2,DAdd3,DAdd4,DAdd5,DState,DPin,DPANRef,DPANCat,Category,DDesgn,DType,DeEmail,DePhone,DeTin,CollNonRes,PerEstInd) VALUES (@DId,@CoID,@DName,@DPan,@DAdd1,@DAdd2,@DAdd3,@DAdd4,@DAdd5,@DState,@DPin,@DPANRef,@DPANCat,@Category,@DDesgn,@DType,@DeEmail,@DePhone,@DeTin,@CollNonRes,@PerEstInd)"
        'Dim cmd As New OleDbCommand(sqlQry, cn)
        'cmd.Parameters.AddWithValue("@DId", id)
        'cmd.Parameters.AddWithValue("@CoID", selectedcoid)
        'cmd.Parameters.AddWithValue("@DName", txtDName.Text)
        'cmd.Parameters.AddWithValue("@DPan", txtDPAN.Text)
        'cmd.Parameters.AddWithValue("@DAdd1", txtDAdd1.Text)
        'cmd.Parameters.AddWithValue("@DAdd2", txtDAdd2.Text)
        'cmd.Parameters.AddWithValue("@DAdd3", txtDAdd3.Text)
        'cmd.Parameters.AddWithValue("@DAdd4", txtDAdd4.Text)
        'cmd.Parameters.AddWithValue("@DAdd5", txtDAdd5.Text)
        'cmd.Parameters.AddWithValue("@DState", cboDState.SelectedValue)
        'cmd.Parameters.AddWithValue("@DPin", txtDPin.Text)
        'cmd.Parameters.AddWithValue("@DPANRef", txtref.Text)
        'cmd.Parameters.AddWithValue("@DPANCat", 0)
        'cmd.Parameters.AddWithValue("@Category", "G")
        'cmd.Parameters.AddWithValue("@DDesgn", txtDDesgn.Text)
        'If optCo.Checked = True Then
        '    cmd.Parameters.AddWithValue("@DType", "C")
        'Else
        '    cmd.Parameters.AddWithValue("@DType", "O")
        'End If
        'cmd.Parameters.AddWithValue("@DeEmail", txtDeEmail.Text)
        'cmd.Parameters.AddWithValue("@DePhone", txtDePhone.Text)
        'cmd.Parameters.AddWithValue("@DeTin", txtTIN.Text)
        'cmd.Parameters.AddWithValue("@CollNonRes", IIf(CboCollNonRes.Text = "Yes", True, False))
        'cmd.Parameters.AddWithValue("@PerEstInd", IIf(CboPerEstInd.Text = "Yes", True, False))
        'cmd.ExecuteNonQuery()
        'cmd.Dispose()
        'cmd1.Dispose()
        'Select Case Frm_typ
        '    Case 26
        '        frmTDS26Q.txtDedPAN.Text = txtDPAN.Text
        '        frmTDS26Q.txtDedPAN.Tag = IIf(optCo.Checked = True, "C", "O")
        '        frmTDS26Q.cboDedName.Items.Add(txtDName.Text)
        '        frmTDS26Q.cboDedName.ItemData(frmTDS26Q.cboDedName.NewIndex) = oDed.did
        '        frmTDS26Q.did = oDed.did
        '    Case 27
        '        frmTDS27Q.txtDedPAN.Text = txtDPAN.Text
        '        'frmTDS27.txtDedPAN.Tag = IIf(optCo.Value = True, "C", "O")
        '    Case 24
        '        frmTDS24Q.txtDedPAN.Text = txtDPAN.Text
        '        frmTDS24Q.txtDedPAN.Tag = IIf(optCo.Value = True, "C", "O")
        '        frmTDS24Q.cboDedName.AddItem oDed.DName
        '        frmTDS24Q.cboDedName.ItemData(frmTDS24Q.cboDedName.NewIndex) = oDed.did
        '        frmTDS24Q.did = oDed.did
        'End Select
        'Me.Close()


    End Sub

    Private Sub txtdeducPin_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDPin.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub
    Private Sub cmdFacth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdlastadd.Click
        Dim nds As New DataSet

        nds = FetchDataSet("select dadd1,dadd2,dadd3,dadd4,dadd5,dstate,dpin from deductmst where coid=" & selectedcoid & " order by did desc")
        If nds.Tables(0).Rows.Count > 0 Then
            txtDAdd1.Text = nds.Tables(0).Rows(0)(0) & ""
            txtDAdd2.Text = nds.Tables(0).Rows(0)(1) & ""
            txtDAdd3.Text = nds.Tables(0).Rows(0)(2) & ""
            txtDAdd4.Text = nds.Tables(0).Rows(0)(3) & ""
            txtDAdd5.Text = nds.Tables(0).Rows(0)(4) & ""
            txtDPin.Text = nds.Tables(0).Rows(0)(6) & ""

            cboDState.SelectedValue = nds.Tables(0).Rows(0)(5)
        End If

        txtDPAN.Focus()
        nds.Dispose()


    End Sub
    Public Sub Fill_cmbState()
        Dim nds As New DataSet
        nds = FetchDataSet("Select StateID,StateName from StateMst")
        cboDState.DataSource = nds.Tables(0)
        cboDState.ValueMember = "StateID"
        cboDState.DisplayMember = "StateName"

    End Sub

    Private Sub txtDedName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDName.Enter
        txtDName.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDedName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDName.Leave
        txtDName.BackColor = Color.White
    End Sub

    Private Sub txtadress1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd1.Leave
        txtDAdd1.BackColor = Color.White
    End Sub

    Private Sub txtadress1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd1.Enter
        txtDAdd1.BackColor = Color.LightYellow
    End Sub

    Private Sub txtadress2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd2.Enter
        txtDAdd2.BackColor = Color.LightYellow
    End Sub

    Private Sub txtadress2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd2.Leave
        txtDAdd2.BackColor = Color.White
    End Sub

    Private Sub txtadress3_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd3.Leave
        txtDAdd3.BackColor = Color.White
    End Sub

    Private Sub txtadress3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd3.Enter
        txtDAdd3.BackColor = Color.LightYellow
    End Sub

    Private Sub txtAdress4_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd4.Leave
        txtDAdd4.BackColor = Color.White
    End Sub

    Private Sub txtAdress4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd4.Enter
        txtDAdd4.BackColor = Color.LightYellow
    End Sub

    Private Sub txtAdress5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd5.Enter
        txtDAdd5.BackColor = Color.LightYellow
    End Sub

    Private Sub txtAdress5_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd5.Leave
        txtDAdd5.BackColor = Color.White
    End Sub

    Private Sub cmbdeducState_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDState.Leave
        cboDState.BackColor = Color.White
    End Sub

    Private Sub cmbdeducState_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDState.Enter
        cboDState.BackColor = Color.LightYellow
    End Sub

    Private Sub txtdeducPin_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPin.Enter
        txtDPin.BackColor = Color.LightYellow
    End Sub

    Private Sub txtdeducPin_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPin.Leave
        txtDPin.BackColor = Color.White
    End Sub

    Private Sub cmbCatofPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbocat.Enter
        cbocat.BackColor = Color.LightYellow
    End Sub

    Private Sub cbocat_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbocat.Leave
        cbocat.BackColor = Color.White
        If cbocat.SelectedIndex > 0 Then
            txtDPAN.Text = cbocat.Text
        End If
    End Sub

    Private Sub txtPAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPAN.Leave
        txtDPAN.BackColor = Color.White
    End Sub

    Private Sub txtPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPAN.Enter
        txtDPAN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtRef_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtref.Leave
        txtref.BackColor = Color.White
    End Sub

    Private Sub txtRef_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtref.Enter
        txtref.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbCate_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCategory.Enter
        cboCategory.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbCate_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCategory.Leave
        cboCategory.BackColor = Color.White
    End Sub

    Private Sub txtDesig_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDDesgn.Enter
        txtDDesgn.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDesig_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDDesgn.Leave
        txtDDesgn.BackColor = Color.White
    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub TableLayoutPanel5_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel5.Paint

    End Sub

    Private Sub cbocat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbocat.SelectedIndexChanged
        'Call CtrlLostFocus(cbocat)
        If cbocat.SelectedIndex > 0 Then
            txtDPAN.Text = cbocat.Text
            txtDPAN.Enabled = False
            CboCollNonRes.Enabled = True
            Label11.Enabled = True
        ElseIf cbocat.SelectedIndex = 0 Then
            txtDPAN.Text = vbNullString
            txtDPAN.Enabled = True
            'CboCollNonRes.Locked = False
            ' CboPerEstInd.Locked = False
            CboCollNonRes.Enabled = False
            Label11.Enabled = False
        ElseIf cbocat.SelectedIndex = 0 Then
            txtref.Enabled = False
        End If

    End Sub

    Private Sub CboCollNonRes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboCollNonRes.SelectedIndexChanged
        If CboCollNonRes.SelectedIndex = 0 Then
            CboPerEstInd.Visible = True
            Label12.Visible = True
        Else
            CboPerEstInd.Visible = False
            Label12.Visible = False
            CboPerEstInd.SelectedIndex = -1
        End If
    End Sub

    Private Sub CboCollNonRes_Leave(sender As Object, e As EventArgs) Handles CboCollNonRes.Leave
        If CboCollNonRes.SelectedIndex = 0 Then
            CboPerEstInd.Visible = True
            Label12.Visible = True
        Else
            CboPerEstInd.Visible = False
            Label12.Visible = False
            CboPerEstInd.SelectedIndex = -1
        End If

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Dispose()

        If Frm_typ = "24Q" Then
            frmTDS24Q.cboDedSection.Focus()
        End If
        If Frm_typ = "26Q" Then
            If frmTDS26Q.cboDedName.Visible = True Then
                frmTDS26Q.cboDedName.Focus()
            End If
        End If
    End Sub

    Private Function BeforeSave() As Boolean
        Dim rs As New DataSet
        Dim i As Integer
        If Trim(txtDName.Text) = vbNullString Then
            Call MsgBox("Deductee's Name Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtDName.Focus()
            Exit Function
        End If
        If Trim(txtDAdd1.Text) = vbNullString Then
            Call MsgBox("Deductee's Address Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtDAdd1.Focus()
            Exit Function
        End If
        If Trim(cboDState.Text) = vbNullString Or cboDState.SelectedIndex = -1 Then
            Call MsgBox("State Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            cboDState.Focus()
            Exit Function
        End If
        If Trim(txtDPin.Text) = vbNullString Then
            Call MsgBox("Pin Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtDPin.Focus()
            Exit Function
        End If

        If Trim(txtDPAN.Text) <> vbNullString And cbocat.SelectedIndex = 0 Then
            Dim PANErr As Integer
            PANErr = IsValidPAN(txtDPAN.Text, True, True)
            If PANErr <> 0 Then
                Call MsgBox("Length of PAN is invalid, please enter 10 digit valid PAN.", vbExclamation, "Invalid PAN")

                BeforeSave = False
                txtDPAN.Focus()
                Exit Function
            End If

        Else
            'AS per new norms from NSDL, we can leave this blank..
            '    Call MsgBox("If you do not have a PAN you should submit a duly filled and signed Form 49A alongwith" _
            '            & vbCrLf & "your e-TDS return. Please enter PANAPPLIED in this case." _
            '            , vbInformation, "PAN REQUIRED")
            '    Cancel = True
            '    txtDPAN.SetFocus
            '    Exit Sub
        End If




        'If Trim(txtref.Text) <> vbNullString Then
        '    If Len(txtref.Text) <> 10 Then
        '        Call MsgBox("REF No. Should Be Of 10 Character!", vbInformation, "Caution")
        '        Cancel = True
        '        txtref.SetFocus
        '        Exit Sub
        '    End If
        'End If
        If optCo.Checked = False And optOther.Checked = False Then
            Call MsgBox("Select Deductee Type From The Given Two Option!", vbInformation, "Caution")
            BeforeSave = False
            optCo.Focus()
            Exit Function
        End If

        If cbocat.SelectedIndex < 0 Then
            Call MsgBox("Category Can not be blank!", vbInformation, "Caution")
            BeforeSave = False
            Exit Function
        ElseIf cbocat.SelectedIndex = 0 And txtDPAN.Text = vbNullString Then
            Call MsgBox("Enter the PAN No.!", vbInformation, "Caution")
            BeforeSave = False
            txtDPAN.Focus()
            Exit Function
            'ElseIf cbocat.ListIndex > 0 And txtref = vbNullString Then
            '    Call MsgBox("Enter the Reference No.!", vbInformation, "Caution")
            '    Cancel = True
            '    txtref.SetFocus
            '    Exit Sub
        End If
        'If cbocat.ListIndex > 0 Then        'check only when valid pan is not selected
        If Len(txtref.Text) <> 10 And txtref.Text <> vbNullString Then
            Call MsgBox("Reference No. should not be less than 10 characters !", vbInformation, "Caution")
            BeforeSave = False
            txtref.Focus()
            Exit Function
        End If
        'End If

        rs = FetchDataSet("Select DPANRef from Deductmst WHERE CoID=" & selectedcoid)
        For i = 0 To rs.Tables(0).Rows.Count - 1
            If txtref.Text <> vbNullString And txtref.Text = rs.Tables(0).Rows(i)(0).ToString() Then
                MsgBox("Reference code already present", vbCritical, Me.Text)
                BeforeSave = False
                txtref.Focus()
                Exit Function
            End If

        Next

        rs = FetchDataSet("Select DPAN, DName from Deductmst WHERE CoID=" & selectedcoid & " and DPAN not in ('PANNOTAVBL','PANINVALID','PANAPPLIED')")
        For i = 0 To rs.Tables(0).Rows.Count - 1
            If txtDPAN.Text = rs.Tables(0).Rows(i)(0).ToString() Then
                MsgBox("This PAN is already mentioned for another deductee" & vbCrLf & "Name:" & rs.Tables(0).Rows(i)("DName").ToString(), vbCritical, Me.Text)
                BeforeSave = False
                txtDPAN.Focus()
                Exit Function
            End If

        Next

        rs.Dispose()

        BeforeSave = True

    End Function

    'Private Function BeforeSave() As Boolean
    '    Dim rs As New DataSet
    '    Dim i As Integer
    '    If Trim(txtDName.Text) = vbNullString Then
    '        Call MsgBox("Deductee's Name Cannot Be Blank!", vbInformation, "Caution")
    '        BeforeSave = False
    '        txtDName.Focus()
    '        Exit Function
    '    End If
    '    If Trim(txtDAdd1.Text) = vbNullString Then
    '        Call MsgBox("Deductee's Address Cannot Be Blank!", vbInformation, "Caution")
    '        BeforeSave = False
    '        txtDAdd1.Focus()
    '        Exit Function
    '    End If
    '    If Trim(cboDState.Text) = vbNullString Or cboDState.SelectedIndex = -1 Then
    '        Call MsgBox("State Cannot Be Blank!", vbInformation, "Caution")
    '        BeforeSave = False
    '        cboDState.Focus()
    '        Exit Function
    '    End If
    '    If Trim(txtDPin.Text) = vbNullString Then
    '        Call MsgBox("Pin Cannot Be Blank!", vbInformation, "Caution")
    '        BeforeSave = False
    '        txtDPin.Focus()
    '        Exit Function
    '    End If

    '    If Trim(txtDPAN.Text) <> vbNullString And cbocat.SelectedIndex = 0 Then
    '        Dim PANErr As Integer
    '        PANErr = IsValidPAN(txtDPAN.Text, True, True)
    '        If PANErr <> 0 Then
    '            Call MsgBox("Length of PAN is invalid, please enter 10 digit valid PAN.", vbExclamation, "Invalid PAN")

    '            BeforeSave = False
    '            txtDPAN.Focus()
    '            Exit Function
    '        End If


    '    End If

    '    If optCo.Checked = False And optOther.Checked = False Then
    '        Call MsgBox("Select Deductee Type From The Given Two Option!", vbInformation, "Caution")
    '        BeforeSave = False
    '        optCo.Focus()
    '        Exit Function
    '    End If

    '    If cbocat.SelectedIndex < 0 Then
    '        Call MsgBox("Category Can not be blank!", vbInformation, "Caution")
    '        BeforeSave = False
    '        Exit Function
    '    ElseIf cbocat.SelectedIndex = 0 And txtDPAN.Text = vbNullString Then
    '        Call MsgBox("Enter the PAN No.!", vbInformation, "Caution")
    '        BeforeSave = False
    '        txtDPAN.Focus()
    '        Exit Function

    '    End If

    '    If Len(txtref.Text) <> 10 And txtref.Text <> vbNullString Then
    '        Call MsgBox("Reference No. should not be less than 10 characters !", vbInformation, "Caution")
    '        BeforeSave = False
    '        txtref.Focus()
    '        Exit Function
    '    End If


    '    rs = FetchDataSet("Select DPANRef from Deductmst WHERE CoID=" & selectedcoid)
    '    For i = 0 To rs.Tables(0).Rows.Count - 1
    '        If txtref.Text <> vbNullString And txtref.Text = rs.Tables(0).Rows(i)(0).ToString() Then
    '            MsgBox("Reference code already present", vbCritical, Me.Text)
    '            BeforeSave = False
    '            txtref.Focus()
    '            Exit Function
    '        End If

    '    Next

    '    rs = FetchDataSet("Select DPAN, DName from Deductmst WHERE CoID=" & selectedcoid & " and DPAN not in ('PANNOTAVBL','PANINVALID','PANAPPLIED')")
    '    For i = 0 To rs.Tables(0).Rows.Count - 1
    '        If txtDPAN.Text = rs.Tables(0).Rows(i)(0).ToString() Then
    '            MsgBox("This PAN is already mentioned for another deductee" & vbCrLf & "Name:" & rs.Tables(0).Rows(i)("DName").ToString(), vbCritical, Me.Text)
    '            BeforeSave = False
    '            txtDPAN.Focus()
    '            Exit Function
    '        End If

    '    Next

    '    rs.Dispose()
    '    BeforeSave = True

    'End Function

    Private Sub oDed_PrepareDataForSave(Cancel As Boolean) Handles oDed.PrepareDataForSave
        With oDed
            .did = 0
            .DName = Trim(txtDName.Text)
            .DAdd1 = Trim(txtDAdd1.Text)
            .DAdd2 = Trim(txtDAdd2.Text)
            .DAdd3 = Trim(txtDAdd3.Text)
            .DAdd4 = Trim(txtDAdd4.Text)
            .DAdd5 = Trim(txtDAdd5.Text)
            .DPan = Trim(txtDPAN.Text)
            .DPin = Val(txtDPin.Text)
            .coid = selectedcoid
            .DState = cboDState.SelectedValue
            .DType = IIf(optCo.Checked = True, "C", IIf(optOther.Checked = True, "O", "C"))
            .Dref = Trim(txtref.Text)
            '.Dcat = cbocat.SelectedValue
            '.Dcat = IIf(cbocat.SelectedValue <> "", cbocat.SelectedValue, 0)
            .Dcat = cbocat.SelectedIndex
            .Category = Strings.Left(cboCategory.Text, 1)
            .DDesgn = txtDDesgn.Text & ""
            .DeEmail = txtDeEmail.Text
            .DePhone = txtDePhone.Text
            .DeTin = txtTIN.Text
            .CollNonRes = IIf(CboCollNonRes.SelectedIndex = 0, "YES", "NO")
            .PerEstInd = IIf(CboPerEstInd.SelectedIndex = 0, "YES", "NO") 'CboPerEstInd.ItemData(CboPerEstInd.ListIndex)
        End With
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub txtDeEmail_TextChanged(sender As Object, e As EventArgs) Handles txtDeEmail.TextChanged

    End Sub

    Private Sub txtDeEmail_Validating(sender As Object, e As CancelEventArgs) Handles txtDeEmail.Validating
        If Len(Trim(txtDeEmail.Text)) > 0 Then
            'check only when something is written
            If ValidEmail(txtDeEmail.Text) = False Then
                MsgBox("Invalid Email ID, please correct it", vbCritical)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub frmDeducteeTDS_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        oDed = Nothing
        Me.Dispose()
    End Sub

    Private Sub cbocat_Click(sender As Object, e As EventArgs) Handles cbocat.Click

    End Sub

    Private Sub txtDPAN_TextChanged(sender As Object, e As EventArgs) Handles txtDPAN.TextChanged

    End Sub

    Private Sub txtDPAN_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDPAN.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            txtDPAN.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub txtref_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtref.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            txtref.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub cbocat_LostFocus(sender As Object, e As EventArgs) Handles cbocat.LostFocus
        If cbocat.SelectedIndex = 0 Then
            txtref.Enabled = False
        Else
            txtref.Enabled = True
        End If
    End Sub

    Private Sub txtDePhone_TextChanged(sender As Object, e As EventArgs) Handles txtDePhone.TextChanged

    End Sub

    Private Sub txtDePhone_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDePhone.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtDeEmail_Enter(sender As Object, e As EventArgs) Handles txtDeEmail.Enter
        txtDeEmail.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDeEmail_Leave(sender As Object, e As EventArgs) Handles txtDeEmail.Leave
        txtDeEmail.BackColor = Color.White
    End Sub

    Private Sub txtDePhone_Enter(sender As Object, e As EventArgs) Handles txtDePhone.Enter
        txtDePhone.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDePhone_Leave(sender As Object, e As EventArgs) Handles txtDePhone.Leave
        txtDePhone.BackColor = Color.White
    End Sub

    Private Sub txtTIN_TextChanged(sender As Object, e As EventArgs) Handles txtTIN.TextChanged

    End Sub

    Private Sub txtTIN_Leave(sender As Object, e As EventArgs) Handles txtTIN.Leave
        txtTIN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtTIN_Enter(sender As Object, e As EventArgs) Handles txtTIN.Enter
        txtTIN.BackColor = Color.White
    End Sub

    Private Sub cboCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCategory.SelectedIndexChanged

    End Sub

    Private Sub cboDState_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDState.SelectedIndexChanged

    End Sub

    Private Sub cboDState_KeyDown(sender As Object, e As KeyEventArgs) Handles cboDState.KeyDown
        If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub cboDState_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboDState.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboDState.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub
End Class