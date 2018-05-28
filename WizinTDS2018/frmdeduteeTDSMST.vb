Imports System.ComponentModel
Public Class frmdeduteeTDSMST

    Dim WithEvents oDed As ClsDeductMstObj

    Public Frm_typ As String
    Public DDid As Integer


    Private Sub frmdeduteeTDSMST_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        ' If Asc(e.KeyChar) = Keys.Enter Then
        Me.EnterTab(e)
        'End If
    End Sub

    Private Sub cboDName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDName.Enter
        cboDName.BackColor = Color.LightYellow
    End Sub

    Private Sub cboDName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDName.Leave
        cboDName.BackColor = Color.White
    End Sub

    Private Sub txtDName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDName.Enter
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

    Private Sub txtadress2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd2.Leave
        txtDAdd2.BackColor = Color.White
    End Sub

    Private Sub txtadress2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd2.Enter
        txtDAdd2.BackColor = Color.LightYellow
    End Sub

    Private Sub txtadress3_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd3.Leave
        txtDAdd3.BackColor = Color.White
    End Sub

    Private Sub txtadress3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd3.Enter
        txtDAdd3.BackColor = Color.LightYellow
    End Sub

    Private Sub txtAdress4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd4.Enter
        txtDAdd4.BackColor = Color.LightYellow
    End Sub

    Private Sub txtAdress4_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd4.Leave
        txtDAdd4.BackColor = Color.White
    End Sub

    Private Sub txtDAdd5_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd5.Leave
        txtDAdd5.BackColor = Color.White
    End Sub

    Private Sub txtDAdd5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd5.Enter
        txtDAdd5.BackColor = Color.LightYellow
    End Sub

    Private Sub cboDState_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDState.Leave
        cboDState.BackColor = Color.White
    End Sub

    Private Sub cboDState_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDState.Enter
        cboDState.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDPin_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPin.Leave
        txtDPin.BackColor = Color.White

    End Sub

    Private Sub txtDPin_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPin.Enter
        txtDPin.BackColor = Color.LightYellow
    End Sub

    Private Sub cbocat_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbocat.Enter
        cbocat.BackColor = Color.LightYellow
    End Sub

    Private Sub cbocat_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbocat.Leave
        cbocat.BackColor = Color.White
    End Sub

    Private Sub txtDPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPAN.Enter
        txtDPAN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDPAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPAN.Leave
        txtDPAN.BackColor = Color.White
    End Sub

    Private Sub txtRef_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRef.Enter
        txtRef.BackColor = Color.LightYellow
    End Sub

    Private Sub txtRef_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRef.Leave
        txtRef.BackColor = Color.White
    End Sub

    Private Sub cboCategory_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCategory.Enter
        cboCategory.BackColor = Color.LightYellow
    End Sub

    Private Sub cboCategory_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCategory.Leave
        cboCategory.BackColor = Color.White
    End Sub

    Private Sub txtDDesgn_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDDesgn.Leave
        txtDDesgn.BackColor = Color.White
    End Sub

    Private Sub txtDDesgn_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDDesgn.Enter
        txtDDesgn.BackColor = Color.LightYellow
    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub frmdeduteeTDSMST_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox("1")
        oDed = New ClsDeductMstObj
        'Main()
        'MsgBox("2")
        'Fill Category list
        cbocat.Items.Clear()
        cbocat.Items.Add("VALID PAN")
        cbocat.Items.Add("PANAPPLIED")
        cbocat.Items.Add("PANINVALID")
        cbocat.Items.Add("PANNOTAVBL")
        cbocat.SelectedIndex = 0
        cboCategory.Items.Clear()
        cboCategory.Items.Add("G - General/Other")
        cboCategory.Items.Add("W - Woman Assessee")
        cboCategory.Items.Add("S - Senior Citizen")
        cboCategory.Items.Add("O - Super Senior Citizen")
        'cboCategory.SelectedIndex = 0

        'MsgBox("3")
        'Changes for FVU 5.7 - FY 10-11 onwards
        CboCollNonRes.Items.Clear()
        CboCollNonRes.Items.Add("Yes")
        'CboCollNonRes.ItemData(CboCollNonRes.NewIndex) = 0
        CboCollNonRes.Items.Add("No")
        'CboCollNonRes.ItemData(CboCollNonRes.NewIndex) = 1
        CboPerEstInd.Items.Clear()
        CboPerEstInd.Items.Add("Yes")
        'CboPerEstInd.ItemData(CboPerEstInd.NewIndex) = 0
        CboPerEstInd.Items.Add("No")
        'CboPerEstInd.ItemData(CboPerEstInd.NewIndex) = 1

        'Changes for FVU 3.0 - FY 10-11 onwards - done by Nitin Betharia
        'txtref.Enabled = False

        'MsgBox("4")
        Call fillcombo()
        'MsgBox("5")
        Call fillDName()
        If cboDName.Items.Count < 1 Then
            cmdSave.Enabled = False
        Else
            cmdSave.Enabled = True
        End If

    End Sub
    Private Sub fillcombo()
        Dim fds As New DataSet
        fds = FetchDataSet("Select statename,stateid from StateMst Order By StateName")
        cboDState.DataSource = fds.Tables(0)
        cboDState.DisplayMember = "StateName"
        cboDState.ValueMember = "StateID"

        fds.Dispose()

    End Sub
    Private Sub fillDName()
        Dim fds As New DataSet
        Dim sql As String
        If chkinact.CheckState = CheckState.Unchecked Then
            sql = "Select DName,DId From DeductMst Where CoId = " & selectedcoid & " Order By DName"
        Else
            sql = "Select DName,DId From DeductMst Where CoId = " & selectedcoid
            sql = sql & " and did not in(select s.did from deductee26Q s,retnmst r where s.retnid=r.retnid and r.coid=" & selectedcoid & ")"
            sql = sql & " and did not in(select s.did from deductee24Q s,retnmst r where s.retnid=r.retnid and r.coid=" & selectedcoid & ")"
            sql = sql & " and did not in(select s.did from deductee27EQ s,retnmst r where s.retnid=r.retnid and r.coid=" & selectedcoid & ")"
            sql = sql & " and did not in (select s.did from SalaryDetail24Q s,retnmst r where s.retnid=r.retnid and r.coid=" & selectedcoid & ")"  'add by jayhsree on 27/07/2006

            sql = sql & " and did not in (select s.did from Form16Details s,retnmst r where s.retnid=r.retnid and r.coid=" & selectedcoid & ")"
            sql = sql & " Order By DName"
        End If
        fds = FetchDataSet(sql)
        cboDName.DataSource = fds.Tables(0)
        cboDName.DisplayMember = "DName"
        cboDName.ValueMember = "DId"

        'For i = 0 To fds.Tables(0).Rows.Count - 1

        '    cboDName.DisplayMember = fds.Tables(0).Rows(i)(0)
        '    cboDName.ValueMember = fds.Tables(0).Rows(i)(1)
        'Next i

        'cboDName.SelectedIndex = 0
        fds.Dispose()

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub frmdeduteeTDSMST_BindingContextChanged(sender As Object, e As EventArgs) Handles Me.BindingContextChanged

    End Sub

    Private Sub frmdeduteeTDSMST_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        'EndMain()
        oDed = Nothing

    End Sub

    Private Sub chkinact_CheckedChanged(sender As Object, e As EventArgs) Handles chkinact.CheckedChanged

    End Sub

    Private Sub chkinact_Click(sender As Object, e As EventArgs) Handles chkinact.Click
        cboDName.DataSource = Nothing
        cboDName.Items().Clear()
        clearDeducteeCtrls()
        fillDName()

    End Sub
    Private Sub clearDeducteeCtrls()
        txtDName.Text = vbNullString
        txtDAdd1.Text = vbNullString
        txtDAdd2.Text = vbNullString
        txtDAdd3.Text = vbNullString
        txtDAdd4.Text = vbNullString
        txtDAdd5.Text = vbNullString
        cboDState.SelectedIndex = -1
        txtDPin.Text = vbNullString
        txtDPAN.Text = vbNullString
        txtRef.Text = vbNullString
        txtDid.Text = vbNullString

        cbocat.SelectedIndex = 0
        cboDName.SelectedIndex = -1
        txtDDesgn.Text = vbNullString
        cboDName.Focus()
        txtDeEmail.Text = vbNullString
        txtDePhone.Text = vbNullString
        txtTIN.Text = vbNullString
        CboPerEstInd.SelectedIndex = -1
        CboCollNonRes.SelectedIndex = -1
    End Sub

    Public Sub EnterTab(ByVal e As KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{TAB}")
        End If
    End Sub




    Private Sub cboDName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDName.SelectedIndexChanged
        'If cboDName.SelectedIndex > -1 Then
        '    'ShowData(cboDName.ItemData(cboDName.ListIndex))
        '    'ShowData(cboDName.SelectedIndex)
        '    ShowData(cboDName.Text)
        '    txtDName.Text = cboDName.Text
        'End If

        If cboDName.Text <> "System.Data.DataRowView" And cboDName.Text <> "" Then
            ShowData(cboDName.Text)
            txtDName.Text = cboDName.Text
        End If
    End Sub
    Private Sub ShowData(nm As String)
        Dim i As Long
        Dim oDed As New ClsDeductMstObj
        oDed = oDed.Fetch(nm)
        'Fetch(nm, oDed)
        With oDed
            txtDAdd1.Text = .DAdd1.ToString()
            txtDAdd2.Text = .DAdd2.ToString()
            txtDAdd3.Text = .DAdd3.ToString()
            txtDAdd4.Text = .DAdd4.ToString()
            txtDAdd5.Text = .DAdd5.ToString()
            cboDState.Text = .DStatenm
            txtDid.Text = .did.ToString()
            'cboDName.SelectedValue = .DState
            cboDName.Text = nm
            cbocat.SelectedIndex = .Dcat
            txtDPin.Text = .DPin
            txtDPAN.Text = .DPan
            optCo.Checked = IIf(.DType = "C", True, False)
            optOther.Checked = IIf(.DType = "O", True, False)
            txtRef.Text = .Dref
            Select Case .Category
                Case "G"
                    cboCategory.SelectedIndex = 0
                Case "W"
                    cboCategory.SelectedIndex = 1
                Case "S"
                    cboCategory.SelectedIndex = 2
                Case "O"
                    cboCategory.SelectedIndex = 3
            End Select

            txtDDesgn.Text = .DDesgn
            txtDeEmail.Text = .DeEmail
            txtDePhone.Text = .DePhone
            txtTIN.Text = .DeTin
            ' CboCollNonRes.SelectedIndex = IIf(.CollNonRes = 0, "yes", "no")
            ' CboPerEstInd.SelectedIndex = IIf(.PerEstInd = 0, "yes", "no")
            Select Case .CollNonRes
                Case "Yes"
                    CboCollNonRes.SelectedIndex = 0
                Case "No"
                    CboCollNonRes.SelectedIndex = 1
            End Select

            Select Case .PerEstInd
                Case "Yes"
                    CboPerEstInd.SelectedIndex = 0
                Case "No"
                    CboPerEstInd.SelectedIndex = 1
            End Select
        End With

    End Sub


    Private Sub frmdeduteeTDSMST_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        'Unload Me
        Me.Dispose()
        If Frm_typ = "24Q" Then
            '        frmTDS24Q.cboDedName.SetFocus
        End If
        If Frm_typ = "26Q" Then
            '        frmTDS26Q.cboDedName.SetFocus
        End If
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        If cboDName.SelectedIndex < 0 Then
            Exit Sub
        End If
        If oDed.LinkD26Q(cboDName.Text) = True Then

            MsgBox("This Record is further used, Cannot Delete", MsgBoxStyle.Critical, "Caution")
            Exit Sub
        End If
        If oDed.LinkD24Q(cboDName.Text) = True Then
            MsgBox("This Record is further used, Cannot Delete", MsgBoxStyle.Critical, "Caution")
            Exit Sub
        End If
        If oDed.LinkD24Qannual(cboDName.Text) = True Then
            MsgBox("This Record is further used, Cannot Delete", MsgBoxStyle.Critical, "Caution")
            Exit Sub
        End If
        If oDed.LinkD27Q(cboDName.Text) = True Then
            MsgBox("This Record is further used, Cannot Delete", MsgBoxStyle.Critical, "Caution")
            Exit Sub
        End If
        If oDed.LinkD27EQ(cboDName.Text) = True Then
            MsgBox("This Record is further used, Cannot Delete", MsgBoxStyle.Critical, "Caution")
            Exit Sub
        End If

        Select Case MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion + vbDefaultButton1, "Delete")
            Case vbYes
                If cboDName.SelectedIndex <> -1 Then
                    Call oDed.Delete(cboDName.Text)
                    Call fillDName()
                End If
            Case vbNo
                Exit Sub
        End Select
        clearDeducteeCtrls()

    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        'Save here..
        If cbocat.Text <> "VALID PAN" And txtDeEmail.Text = vbNullString And txtDePhone.Text = vbNullString And txtTIN.Text = vbNullString Then
            MsgBox("Record Not Saved... Please Enter EmailId, Contact Number And TIN/UIN of the Deductee")
            Exit Sub
        End If
        If oDed.Update(oDed) = True Then
            Select Case Frm_typ
                Case "26Q"
                    frmTDS26Q.txtDedPAN.Text = txtDPAN.Text
                    frmTDS26Q.txtDedPAN.Tag = IIf(optCo.Checked = True, "C", "O")
                    FillDeducteeCombo26()
                    frmTDS26Q.cboDedName.SelectedValue = oDed.did
                    frmTDS26Q.did = oDed.did
                    Me.Dispose()
                Case "27Q"
                    frmTDS27Q.txtDedPAN.Text = txtDPAN.Text

                Case "24Q"
                    frmTDS24Q.txtDedPAN.Text = txtDPAN.Text
                    frmTDS24Q.txtDedPAN.Tag = IIf(optCo.Checked = True, "C", "O")
                    FillDeducteeCombo24()
                    frmTDS24Q.cboDedName.SelectedValue = oDed.did
                    frmTDS24Q.did = oDed.did
                    Me.Dispose()
            End Select
            ' MsgBox("Record saved successfully..")
            'Call fillDName()

        Else
            'stay with this form only..unless user presses cancel.
        End If

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
    Private Sub oDed_BeforeSave(Cancel As Boolean) Handles oDed.BeforeSave
        Dim nds As New DataSet

        If Trim(cboDName.Text) = vbNullString Then
            Call MsgBox("Deductee's Name Cannot Be Blank!", vbInformation, "Caution")
            Cancel = True
            cboDName.Focus()
            Exit Sub
        End If
        If Trim(txtDName.Text) = vbNullString Then
            Call MsgBox("Deductee's Name Cannot Be Blank!", vbInformation, "Caution")
            Cancel = True
            txtDName.Focus()
            Exit Sub
        End If
        If Trim(txtDAdd1.Text) = vbNullString Then
            Call MsgBox("Deductee's Address Cannot Be Blank!", vbInformation, "Caution")
            Cancel = True
            txtDAdd1.Focus()
            Exit Sub
        End If
        If Trim(cboDState.Text) = vbNullString Or cboDState.SelectedIndex = -1 Then
            Call MsgBox("State Cannot Be Blank!", vbInformation, "Caution")
            Cancel = True
            cboDState.Focus()
            Exit Sub
        End If
        If Trim(txtDPin.Text) = vbNullString Then
            Call MsgBox("Pin Cannot Be Blank!", vbInformation, "Caution")
            Cancel = True
            txtDPin.Focus()
            Exit Sub
        End If
        If Trim(txtDPAN.Text) <> vbNullString And cbocat.SelectedIndex = 0 Then
            Dim PANErr As Integer
            PANErr = IsValidPAN(txtDPAN.Text, True, True)
            If PANErr <> 0 Then
                Call MsgBox("Length of PAN is invalid, please enter 10 digit valid PAN.", vbExclamation, "Invalid PAN")
                Cancel = True
                txtDPAN.Focus()
                Exit Sub
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
            Cancel = True
            optCo.Focus()
            Exit Sub
        End If

        If cbocat.SelectedIndex < 0 Then
            Call MsgBox("Category Can not be left blank!", vbInformation, "Caution")
            Cancel = True
            Exit Sub
        ElseIf cbocat.SelectedIndex = 0 And txtDPAN.Text = vbNullString Then
            Call MsgBox("Enter the PAN No.!", vbInformation, "Caution")
            Cancel = True
            txtDPAN.Focus()
            Exit Sub
            'ElseIf cbocat.ListIndex > 0 And txtref = vbNullString Then
            '    Call MsgBox("Enter the Reference No.!", vbInformation, "Caution")
            '    Cancel = True
            '    txtref.SetFocus
            '    Exit Sub
        End If
        If cbocat.SelectedIndex > 0 Then     'check only when valid pan is not selected


            If Len(txtRef.Text) <> 10 Then
                Call MsgBox("Reference No. should not be less than 10 characters !", vbInformation, "Caution")
                Cancel = True
                txtRef.Focus()
                Exit Sub
            End If

            nds = FetchDataSet("Select DPANRef from Deductmst where DID <>" & Val(txtDid.Text) & " and CoID=" & selectedcoid)
            'rs.Open "Select DPANRef from Deductmst where DID <> " & cboDName.ItemData(cboDName.ListIndex) & " and CoID=" & selectedcoid, Cnn, adOpenForwardOnly, adLockReadOnly
            'While Not (rs.BOF Or rs.EOF)

            For j As Long = 0 To nds.Tables(0).Rows.Count - 1
                If txtRef.Text = nds.Tables(0).Rows(j)("dpanref").ToString() Then
                    MsgBox("Reference code already present", vbCritical, Me.Text)
                    Cancel = True
                    txtRef.Focus()
                    Exit Sub
                End If
            Next
        End If
        'rs.MoveNext
        'Wend
        'If rs.State = adStateOpen Then rs.Close
        nds = FetchDataSet("Select DPAN from Deductmst where DID <>" & Val(txtDid.Text) & " and CoID=" & selectedcoid & " and DPAN not in ('PANNOTAVBL','PANINVALID','PANAPPLIED') and DPAN='" & txtDPAN.Text & "'")
        'rs.Open "Select DPAN from Deductmst WHERE DID <> " & cboDName.ItemData(cboDName.ListIndex) & " and CoID=" & selectedcoid & " and DPAN not in ('PANNOTAVBL','PANINVALID','PANAPPLIED')", Cnn, adOpenForwardOnly, adLockReadOnly
        'While Not (rs.BOF Or rs.EOF)
        'For j As Long = 0 To nds.Tables(0).Rows.Count - 1
        If nds.Tables(0).Rows.Count > 0 Then
            If MsgBox("This PAN is already mentioned for another deductee" & vbCrLf &
                       "Do you want to still save with this PAN No", vbCritical + vbYesNo,
                       "WARNING!! SAVE WITH DUPLICATE PAN") = vbYes Then
                Cancel = False
            Else
                Cancel = True
                txtDPAN.Focus()
                Exit Sub
            End If
        End If
        'Next
        'rs.MoveNext
        'Wend

        'Set rs = Nothing
        nds.Dispose()
        Cancel = False
    End Sub

    Private Sub oDed_PrepareDataForSave(Cancel As Boolean) Handles oDed.PrepareDataForSave
        Dim nds As New DataSet
        With oDed
            '    If cbocat.ListIndex < 0 Then
            '        Cancel = True
            '        Exit Sub
            '    ElseIf cbocat.ListIndex = 0 And txtDPAN = "" Then
            '        Cancel = True
            '        Exit Sub
            '    ElseIf cbocat.ListIndex > 0 And txtref = "" Then
            '        Cancel = True
            '        Exit Sub
            '    End If
            'cbocat.Items.Add("VALID PAN")
            'cbocat.Items.Add("PANAPPLIED")
            'cbocat.Items.Add("PANINVALID")
            'cbocat.Items.Add("PANNOTAVBL")
            'cbocat.SelectedIndex = 0

            'cboCategory.Items.Add("G - General/Other")
            'cboCategory.Items.Add("W - Woman Assessee")
            'cboCategory.Items.Add("S - Senior Citizen")
            'cboCategory.Items.Add("O - Super Senior Citizen")
            'cboCategory.SelectedIndex = 0

            nds = FetchDataSet("select stateid from statemst where statename='" & cboDState.Text & "'")
            .did = Val(txtDid.Text) 'cboDName.ItemData(cboDName.ListIndex)
            .DName = Trim(txtDName.Text)
            .DAdd1 = Trim(txtDAdd1.Text)
            .DAdd2 = Trim(txtDAdd2.Text)
            .DAdd3 = Trim(txtDAdd3.Text)
            .DAdd4 = Trim(txtDAdd4.Text)
            .DAdd5 = Trim(txtDAdd5.Text)
            .DPan = Trim(txtDPAN.Text)
            .DPin = Val(txtDPin.Text)
            .coid = selectedcoid
            .DState = nds.Tables(0).Rows(0)(0).ToString() 'Val(txtDStateId.Text) 
            .DStatenm = cboDState.Text
            .DType = IIf(optCo.Checked = True, "C", IIf(optOther.Checked = True, "O", "C"))
            .Dref = Trim(txtRef.Text)
            Select Case cbocat.Text
                Case "VALID PAN"
                    .Dcat = 0
                Case "PANAPPLIED"
                    .Dcat = 1
                Case "PANINVALID"
                    .Dcat = 2
                Case "PANNOTAVBL"
                    .Dcat = 3
            End Select
            '.Dcat = cbocat.ItemData(cbocat.ListIndex)
            .Category = Mid(cboCategory.Text, 1, 1)
            .DDesgn = txtDDesgn.Text
            .DeEmail = txtDeEmail.Text
            .DePhone = txtDePhone.Text
            .DeTin = txtTIN.Text
            .CollNonRes = IIf(CboCollNonRes.Text = "Yes", 0, 1) 'CboCollNonRes.ItemData(CboCollNonRes.ListIndex)
            .PerEstInd = IIf(CboPerEstInd.Text = "Yes", 0, 1) 'CboPerEstInd.ItemData(CboPerEstInd.ListIndex)
        End With
        nds.Dispose()
    End Sub





    Private Sub cmdcorrection_Click(sender As Object, e As EventArgs) Handles cmdcorrection.Click
        frmCorrection.Show()
    End Sub

    Private Sub CboCollNonRes_LostFocus(sender As Object, e As EventArgs) Handles CboCollNonRes.LostFocus
        If CboCollNonRes.SelectedIndex = 0 Then
            CboPerEstInd.Enabled = True
            Label12.Enabled = True
            CboPerEstInd.Visible = True
            Label12.Visible = True
        Else
            CboPerEstInd.Visible = False
            Label12.Visible = False
        End If
    End Sub


    Private Sub txtDeEmail_Validating(sender As Object, e As CancelEventArgs) Handles txtDeEmail.Validating
        If Len(Trim(txtDeEmail.Text)) > 0 Then
            'check only when something is written
            If ValidEmail(txtDeEmail.Text) = False Then
                MessageBox.Show("Invalid Email ID, please correct it", "Error")
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub txtDePhone_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDePhone.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtDPin_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDPin.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtDPAN_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDPAN.KeyPress
        If e.KeyChar = "/" Or e.KeyChar = "\" Or e.KeyChar = "'" Or e.KeyChar = "," Or e.KeyChar = "." Or e.KeyChar = ";" Or e.KeyChar = ":" Or e.KeyChar = "<" Or e.KeyChar = ">" Or e.KeyChar = "=" Or e.KeyChar = "-" Or e.KeyChar = "[" Or e.KeyChar = "]" Or e.KeyChar = "*" Or e.KeyChar = "+" Or e.KeyChar = "!" Or e.KeyChar = "~" Or e.KeyChar = "`" Or e.KeyChar = "@" Or e.KeyChar = "#" Or e.KeyChar = "$" Or e.KeyChar = "%" Or e.KeyChar = "^" Or e.KeyChar = "(" Or e.KeyChar = ")" Or e.KeyChar = "|" Or e.KeyChar = "{" Or e.KeyChar = "}" Or e.KeyChar = "?" Or e.KeyChar = "_" Then
            e.Handled = True
        End If
    End Sub

    Private Sub oDed_BeforeDelete(Cancel As Boolean) Handles oDed.BeforeDelete

    End Sub

    Private Sub cbocat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbocat.SelectedIndexChanged
        If cbocat.SelectedIndex > 0 Then
            txtDPAN.Text = cbocat.Text
            txtDPAN.ReadOnly = True
            CboCollNonRes.Enabled = True
            Label4.Enabled = True
            txtRef.Enabled = True
            Label12.Enabled = True
            CboPerEstInd.Enabled = True
            'Label11.Enabled = True
        ElseIf cbocat.SelectedIndex = 0 Then
            txtDPAN.Text = vbNullString
            txtDPAN.ReadOnly = False
            txtRef.Enabled = False
            txtRef.Text = ""
            'Label11.Enabled = False
            'CboCollNonRes.Locked = False
            ' CboPerEstInd.Locked = False
            CboCollNonRes.Enabled = False
            Label4.Enabled = False
            Label12.Enabled = False
            CboPerEstInd.Enabled = False
            CboCollNonRes.SelectedIndex = -1
            CboPerEstInd.SelectedIndex = -1
        End If
    End Sub

    Private Sub CboCollNonRes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboCollNonRes.SelectedIndexChanged

        If CboCollNonRes.SelectedIndex = 0 Then
            CboPerEstInd.Enabled = True
            Label12.Enabled = True
            CboPerEstInd.Visible = True
            Label12.Visible = True
        Else
            CboPerEstInd.Visible = False
            Label12.Visible = False
        End If
    End Sub

    Private Sub txtDPin_TextChanged(sender As Object, e As EventArgs) Handles txtDPin.TextChanged

    End Sub

    Private Sub cbocat_GotFocus(sender As Object, e As EventArgs) Handles cbocat.GotFocus

    End Sub

    Private Sub txtDePhone_TextChanged(sender As Object, e As EventArgs) Handles txtDePhone.TextChanged

    End Sub

    Private Sub txtDePhone_Leave(sender As Object, e As EventArgs) Handles txtDePhone.Leave

    End Sub

    Private Sub cboDState_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDState.SelectedIndexChanged

    End Sub

    Private Sub txtDPin_LostFocus(sender As Object, e As EventArgs) Handles txtDPin.LostFocus

    End Sub

    Private Sub frmdeduteeTDSMST_Leave(sender As Object, e As EventArgs) Handles Me.Leave
        If sender.ToString() <> "Cmd" Then

        End If
    End Sub

    Private Sub txtDAdd1_TextChanged(sender As Object, e As EventArgs) Handles txtDAdd1.TextChanged

    End Sub

    Private Sub cboDName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboDName.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboDName.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub cboDState_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboDState.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboDState.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub cboDName_KeyDown(sender As Object, e As KeyEventArgs) Handles cboDName.KeyDown
        If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub cboDState_KeyDown(sender As Object, e As KeyEventArgs) Handles cboDState.KeyDown
        If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub cmdSearch_Click(sender As Object, e As EventArgs) Handles cmdSearch.Click
        'frm_Find.Show()
        list()
    End Sub

    Private Sub list()
        Dim i As Integer
        Dim Dname As String
        findsql1 = "SELECT deductmst.DName + ' - ' + deductmst.dpan AS Dname , deductmst.Did from deductmst where  deductmst.coid = " & selectedcoid
        findsql1 = findsql1 & " and Dname & Dpan like '"
        findsql2 = "%'  order by Dname "
        txtfind = vbNullString
        'frm_Find. = "Search Deductee"
        frm_Find.ShowDialog()
        'If frm_Find.lvw_results.SelectedItems(0).SubItems(0).Text = "" Then
        'Else
        'cboDName.Tag = cboDName.FindString(frm_Find.lvw_results.SelectedItems(0).SubItems(0).Text)
        'End If
        'cboDName.Tag = frm_Find.lvw_results.SelectedItems(0).SubItems(0).Text
        If frm_Find.findflag = True Then
            DDid = frm_Find.obj.gid
            cboDName.SelectedValue = DDid
            'cboDName.Text = frm_Find.lvw_results.SelectedItems(0).SubItems(0).Text
            For i = 0 To cboDName.Items.Count - 1
                If cboDName.SelectedValue = DDid Then 'cboDName.FindString(frm_Find.lvw_results.SelectedItems(0).SubItems(0).Text) Then '= DDid Then
                    'cboDName.SelectedValue = DDid
                    cboDName.Text = frm_Find.lvw_results.SelectedItems(0).SubItems(0).Text
                    'cboDName.SelectedItem(cboDName.SelectedValue).ToString()
                    'cboDName.SelectedValue.ToString()
                    Exit For
                End If
            Next i
        End If
    End Sub


End Class