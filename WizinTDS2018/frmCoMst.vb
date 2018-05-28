Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Data
Imports System.ComponentModel
Imports System.Reflection
Imports System.Drawing.Text

Public Class frmCoMst
    Dim drag As Boolean
    Dim mousex As Integer
    Dim mousey As Integer
    Dim dr3 As OleDbDataReader
    'Public SelectedId As Integer
    'Public Temp1 As Integer = 0
    Public tempadd As Integer = 0
    Public tempEdit As Integer = 0
    Public MaxID As Integer
    Dim frmGovDet As New frmGovDetails
    Dim frmAltDet As frmAltDetails
    Public WithEvents oCoMst As clsCoMst
    Dim frm As New frmImport
    Public isvalidGSTNo As String
    Public gstin As String
    Dim Mode As String




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'frmChangePw.Show()
    End Sub

    Private Sub cmdCngyear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'frmChangeyr.Show()
    End Sub
    Private Sub cmdDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdonline.Click
        frmOnline.Show()
    End Sub
    Private Sub cmduserGuide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdugide.Click
        Dim url As String = "" & My.Application.Info.DirectoryPath & "\Support\userguide2012-13.htm"
        If File.Exists(url) = True Then
            Process.Start(url)
        End If

    End Sub
    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCoEdit.Click
        tempEdit = 1
        controlEnabled()
        Me.MainTab.SelectedIndex = 0
        cmdSave.Enabled = True
        cmdCoAdd.Enabled = False
        cmdBackup.Enabled = False
        cmdCoEdit.Enabled = False
        cmdRestore.Enabled = False
        cmdTDS.Enabled = False
        lvwCo.Enabled = False
        cmdExit.Text = "Cancel"

    End Sub
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        If cmdExit.Text = "Exit" Then

            frmLogin.Close()
        Else
            If lvwCo.Items.Count > 0 Then
                If lvwCo.SelectedItems Is Nothing Then
                    lvwCo.Items(1).Selected = True
                End If
                ' lvwCo_Click() lvwCo.selecteditem
            Else
                'No records and still user selected to cancel, end program..
                Close()
            End If
            tempEdit = 0
            tempadd = 0
            controlDisabled()
            ConnectData()
            lvwCo.Enabled = True
            cmdSave.Enabled = False
            cmdCoAdd.Enabled = True
            cmdBackup.Enabled = True
            cmdCoEdit.Enabled = True
            cmdTDS.Enabled = True
            cmdRestore.Enabled = True
            'connectcontols()
            lvwCo.Select()
            'SelectedId = lvwCo.SelectedItems(0).SubItems(1).Text
            ConnectData()
            cmdExit.Text = "Exit"
            'NormalMode()
        End If
        cmdretnsum.Enabled = True
        cmdugide.Enabled = True
    End Sub
    Private Sub cmdReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdretnsum.Click
        'frmretsumm.Show()
    End Sub
    Private Sub cmdTDS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTDS.Click
        If lvwCo.SelectedItems.Count = 0 Then
            Exit Sub
        Else
            selectedcoid = lvwCo.SelectedItems(0).SubItems(1).Text 'Mid(lvwCo.SelectedItems.IndexOf(), 2, Len(lvwCo.SelectedItem.KEY))
            'frmTDS.Show()
            frmTDS.Text = "TDS Forms for - " & txtCoName.Text & " For Asst. Yr. " & AY
            'frmTDS. 0, 315
            frmTDS.ShowDialog()
        End If
        'frmTDS.Show()
    End Sub
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCoAdd.Click
        MainTab.SelectedIndex = 0
        txtCoName.Focus()
        ' End If

        Add()
    End Sub

    Private Sub frmCoMst_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Me.Location = New Point(0, 0)

        Dim MaxNoCompany As Integer
        Dim nds As New DataSet

        If GetSetting("Wizin-TDS", Application.ProductName & "\", "StartUp", 0) = "0" Then
            'Me.WindowState = FormWindowState.Minimized
            FrmWhatsnew.ShowDialog()
        End If
        cmdCoDel.Visible = DeleteAllowed
        cmdCoDel.Enabled = DeleteAllowed
        If frmRegister.Mylock.RegisteredUser = False Then
            NoOfCo = 2
            AllowCertificate = False
        Else
            NoOfCo = Math.Abs(Asc(Mid(frmRegister.Mylock.SoftwareName, 10, 1)) - 64) * Val(Mid(frmRegister.Mylock.SoftwareName, 11, 2))
            If UCase(Mid(frmRegister.Mylock.SoftwareName, 13, 1)) = "C" Then
                AllowCertificate = True
            Else
                AllowCertificate = False
            End If
        End If
        ToolStripStatusLabel2.Text = Strings.Mid(frmRegister.Mylock.SoftwareName, 10, Len(frmRegister.Mylock.SoftwareCode))
        frmRegister.Hide()
        'StatusBarTDS.Panels(2).Text = Mid(frmRegister.AxActiveLock1.SoftwareName, 10, Len(frmRegister.AxActiveLock1.SoftwareCode))
        Dim sql As String = "SELECT distinct count(CoID) FROM CoMst "
        nds = FetchDataSet(sql)
        If nds.Tables(0).Rows.Count > 0 Then
            MaxNoCompany = nds.Tables(0).Rows(0)(0)
        End If

        frmGovDet.Show()
        Label47.Text = "JAK's wizin-TDS For Asst.Yr. " & AY & "-" & frmLogin.version & ""
        Me.Text = "JAK's wizin-TDS For Asst.Yr. " & AY & "-" & frmLogin.version & ""
        'NormalMode()
        frmGovDet.Visible = False
        setupListView()
        connectlistview()
        Fill_StatusDed()
        Fill_cmbdedState()

        If lvwCo.Items.Count > 0 Then
            lvwCo.Items(0).Selected = True
            lvwCo.Select()
        Else
            Me.cmdCoEdit.Enabled = False
            Me.cmdTDS.Enabled = False
            Add()
        End If
        cmdGovtDetails.Enabled = False
        CheckStatus()
        '*added now
        ShowLicInfo()
        Dim CurVersion As String
        'If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
        '    With System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
        '        CurVersion = Application.ProductName & " " & .Major & "." & .Minor & "." & .Revision
        '        Shell(Application.StartupPath & "\VersionUpdate.exe " & "admin " & CurVersion & " " & Application.ProductName)
        '        ' CurVersion = "" & App.Major & "." & App.Minor & "." & App.Revision
        '    End With
        'End If
        Dim versionNumber As String
        Dim a, b
        versionNumber = Assembly.GetExecutingAssembly.FullName
        versionNumber = Strings.Left(versionNumber, 29)
        a = Split(versionNumber, ",")
        b = Strings.Right(a(1), 7)

        CurVersion = Strings.Right(a(0), 4) + "." + Strings.Left(b, 3)
        'this is required
        Shell(Application.StartupPath & "\VersionUpdate.exe " & "admin " & CurVersion & " " & Application.ProductName, AppWinStyle.NormalFocus)
        cmdretnsum.Enabled = True

    End Sub
    Private Sub setupListView()
        'add columns to the listview
        lvwCo.Columns.Add("Deductor/Seller Name", 400, HorizontalAlignment.Left)
        lvwCo.Columns.Add("Co ID", 0, HorizontalAlignment.Left)
        'Display listview in details view
        lvwCo.View = View.Details
        'display grid lines
        lvwCo.GridLines = True
        'allow full row selection
        lvwCo.FullRowSelect = True
    End Sub
    Public Sub connectlistview()

        Dim nds As New DataSet, i As Integer
        Dim sql As String = "Select CoName,CoID from CoMst order by coname"

        nds = FetchDataSet(sql)
        If nds.Tables(0).Rows.Count > 0 Then
            For i = 0 To nds.Tables(0).Rows.Count - 1
                Dim newitem As New ListViewItem()
                newitem.Text = nds.Tables(0).Rows(i)(0) 'first column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)(1)) 'second column
                lvwCo.Items.Add(newitem)
            Next
        End If
        'lvwCo.Items(i).Selected = True
        nds.Dispose()
    End Sub
    'Private Sub connectcontols() 'Connection to controls with selected Id in listview
    '    Dim D As Integer = lvwCo.SelectedItems(0).SubItems(1).Text
    'End Sub
    Private Sub lvwCo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvwCo.SelectedIndexChanged
        If lvwCo.SelectedIndices.Count > 0 Then
            'SelectedId = lvwCo.SelectedItems(0).SubItems(1).Text
            selectedcoid = lvwCo.SelectedItems(0).SubItems(1).Text
            ConnectData()
            cmdGovtDetails.Enabled = False
        End If
    End Sub

    Private Sub Fill_cmbdedState()

        Dim nds As New DataSet
        Dim nds24 As New DataSet
        Dim nds26 As New DataSet
        Dim nds27 As New DataSet
        Dim nds27E As New DataSet

        Dim QueSt As String = "Select StateID,StateName from StateMst"
        nds = FetchDataSet(QueSt)
        nds24 = FetchDataSet(QueSt)
        nds26 = FetchDataSet(QueSt)
        nds27 = FetchDataSet(QueSt)
        nds27E = FetchDataSet(QueSt)

        cboCoState.DataSource = nds.Tables(0)
        cboCoState.ValueMember = "StateID"
        cboCoState.DisplayMember = "StateName"
        cboPR27State.DataSource = nds27.Tables(0)
        cboPR27State.ValueMember = "StateID"
        cboPR27State.DisplayMember = "StateName"
        cboPR27EState.DataSource = nds27E.Tables(0)
        cboPR27EState.ValueMember = "StateID"
        cboPR27EState.DisplayMember = "StateName"
        cboPR26State.DataSource = nds26.Tables(0)
        cboPR26State.ValueMember = "StateID"
        cboPR26State.DisplayMember = "StateName"
        cboPR24State.DataSource = nds24.Tables(0)
        cboPR24State.ValueMember = "StateID"
        cboPR24State.DisplayMember = "StateName"
        nds.Dispose()
        nds24.Dispose()
        nds26.Dispose()
        nds27.Dispose()
        nds27E.Dispose()
    End Sub
    Public Sub ConnectData()
        'frmLogin.cn.Open()
        Dim nds As New DataSet
        Dim i As Integer
        Dim sql As String = "SELECT * From CoMst WHERE CoMst.CoID = " & selectedcoid

        nds = FetchDataSet(sql)
        For i = 0 To nds.Tables(0).Rows.Count - 1
            txtCoName.Text = nds.Tables(0).Rows(i)(1) & "" 'dr(1) & ""
            txtCoBrDiv.Text = nds.Tables(0).Rows(i)(2) & "" 'dr(2) & ""
            txtCoAdd1.Text = nds.Tables(0).Rows(i)(3) & "" 'dr(3) & ""
            txtCoAdd2.Text = nds.Tables(0).Rows(i)(4) & "" 'dr(4) & ""
            txtCoAdd3.Text = nds.Tables(0).Rows(i)(5) & ""  'dr(5) & ""
            txtCoAdd4.Text = nds.Tables(0).Rows(i)(6) & "" 'dr(6) & ""
            txtCoAdd5.Text = nds.Tables(0).Rows(i)(7) & "" 'dr(7) & ""
            cboCoState.SelectedValue = nds.Tables(0).Rows(i)(8) & ""  'dr(8) & ""
            txtCoPin.Text = nds.Tables(0).Rows(i)(9) & "" 'dr(9) & ""
            txtCoSTD.Text = nds.Tables(0).Rows(i)(10) & ""  'dr(10) & ""
            txtCoPhone.Text = nds.Tables(0).Rows(i)(11) & ""  'dr(11) & ""
            txtCoEmail.Text = nds.Tables(0).Rows(i)(12) & "" 'dr(12) & ""
            txtCoTAN.Text = nds.Tables(0).Rows(i)(13) & "" 'dr(13) & ""
            txtCoPAN.Text = nds.Tables(0).Rows(i)(14) & ""  'dr(14) & ""
            If nds.Tables(0).Rows(i)(15) & "" = True Then
                chkAddChg.Checked = True
            Else
                chkAddChg.Checked = False
            End If
            cboGovtDetails.SelectedValue = nds.Tables(0).Rows(i)(16) & "" 'dr(16) & ""

            txtName24.Text = nds.Tables(0).Rows(i)(17) & "" 'dr(17) & ""
            txtDesg24.Text = nds.Tables(0).Rows(i)(18) & "" 'dr(18) & ""
            txt24Add1.Text = nds.Tables(0).Rows(i)(19) & "" 'dr(19) & ""
            txt24Add2.Text = nds.Tables(0).Rows(i)(20) & "" 'dr(20) & ""
            txt24Add3.Text = nds.Tables(0).Rows(i)(21) & "" 'dr(21) & ""
            txt24Add4.Text = nds.Tables(0).Rows(i)(22) & "" 'dr(22) & ""
            txt24Add5.Text = nds.Tables(0).Rows(i)(23) & "" 'dr(23) & ""
            cboPR24State.SelectedValue = nds.Tables(0).Rows(i)(24) & "" 'dr(24) & ""
            txtPR24Pin.Text = nds.Tables(0).Rows(i)(25) & "" 'dr(25) & ""
            txt24Email.Text = nds.Tables(0).Rows(i)(26) & "" 'dr(26) & ""
            txt24STD.Text = nds.Tables(0).Rows(i)(27) & "" 'dr(27) & ""
            txt24PHONE.Text = nds.Tables(0).Rows(i)(28) & "" ' dr(28) & ""
            If nds.Tables(0).Rows(i)(29) & "" = True Then
                chk24AddChg.Checked = True
            Else
                chk24AddChg.Checked = False
            End If
            txtName26.Text = nds.Tables(0).Rows(i)(30) & "" 'dr(30) & ""
            txtDesg26.Text = nds.Tables(0).Rows(i)(31) & "" 'dr(31) & ""
            txt26add1.Text = nds.Tables(0).Rows(i)(32) & "" 'dr(32) & ""
            txt26add2.Text = nds.Tables(0).Rows(i)(33) & "" 'dr(33) & ""
            txt26add3.Text = nds.Tables(0).Rows(i)(34) & "" 'dr(34) & ""
            txt26add4.Text = nds.Tables(0).Rows(i)(35) & "" 'dr(35) & ""
            txt26add5.Text = nds.Tables(0).Rows(i)(36) & "" 'dr(36) & ""
            cboPR26State.SelectedValue = nds.Tables(0).Rows(i)(37) & "" 'dr(37) & ""
            txtPR26Pin.Text = nds.Tables(0).Rows(i)(38) & "" 'dr(38) & ""
            txt26Email.Text = nds.Tables(0).Rows(i)(39) & "" 'dr(39) & ""
            txt26STD.Text = nds.Tables(0).Rows(i)(40) & "" 'dr(40) & ""
            txt26PHONE.Text = nds.Tables(0).Rows(i)(41) & "" 'dr(41) & ""
            If nds.Tables(0).Rows(i)(42) & "" = True Then
                chk26AddChg.Checked = True
            Else
                chk26AddChg.Checked = False
            End If
            txtName27.Text = nds.Tables(0).Rows(i)(43) & "" 'dr(43) & ""
            txtDesg27.Text = nds.Tables(0).Rows(i)(44) & "" 'dr(44) & ""
            txt27Add1.Text = nds.Tables(0).Rows(i)(45) & "" 'dr(45) & ""
            txt27Add2.Text = nds.Tables(0).Rows(i)(46) & "" 'dr(46) & ""
            txt27Add3.Text = nds.Tables(0).Rows(i)(47) & "" 'dr(47) & ""
            txt27Add4.Text = nds.Tables(0).Rows(i)(48) & "" 'dr(48) & ""
            txt27Add5.Text = nds.Tables(0).Rows(i)(49) & "" 'dr(49) & ""
            cboPR27State.SelectedValue = nds.Tables(0).Rows(i)(50) & "" 'dr(50) & ""
            txtPR27Pin.Text = nds.Tables(0).Rows(i)(51) & "" 'dr(51) & ""
            txt27Email.Text = nds.Tables(0).Rows(i)(52) & "" 'dr(52) & ""
            txt27STD.Text = nds.Tables(0).Rows(i)(53) & "" 'dr(53) & ""
            txt27PHONE.Text = nds.Tables(0).Rows(i)(54) & "" 'dr(54) & ""
            If nds.Tables(0).Rows(i)(55) & "" = True Then
                chk27AddChg.Checked = True
            Else
                chk27AddChg.Checked = False
            End If
            txtName27E.Text = nds.Tables(0).Rows(i)(56) & "" 'dr(56) & ""
            txtDesg27EQ.Text = nds.Tables(0).Rows(i)(57) & "" 'dr(57) & ""
            txt27EAdd1.Text = nds.Tables(0).Rows(i)(58) & "" 'dr(58) & ""
            txt27EAdd2.Text = nds.Tables(0).Rows(i)(59) & "" 'dr(59) & ""
            txt27EAdd3.Text = nds.Tables(0).Rows(i)(60) & "" 'dr(60) & ""
            txt27EAdd4.Text = nds.Tables(0).Rows(i)(61) & "" 'dr(61) & ""
            txt27EAdd5.Text = nds.Tables(0).Rows(i)(62) & "" 'dr(62) & ""
            cboPR27EState.SelectedValue = nds.Tables(0).Rows(i)(63) & "" 'dr(63) & ""
            txtPR27EPin.Text = nds.Tables(0).Rows(i)(64) & "" 'dr(64) & ""
            txt27EEmail.Text = nds.Tables(0).Rows(i)(65) & "" 'dr(65) & ""
            txt27ESTD.Text = nds.Tables(0).Rows(i)(66) & "" 'dr(66) & ""
            txt27EPHONE.Text = nds.Tables(0).Rows(i)(67) & "" 'dr(67) & ""
            If nds.Tables(0).Rows(i)(68) & "" = True Then
                chk27EAddChg.Checked = True
            Else
                chk27EAddChg.Checked = False
            End If
            If nds.Tables(0).Rows(i)(69) & "" = True Then
                chkUseForm16.Checked = True
            Else
                chkUseForm16.Checked = False
            End If
            txtmobile.Text = nds.Tables(0).Rows(i)(81) & "" 'dr(81) & ""
            txt24PRPAN.Text = nds.Tables(0).Rows(i)(101) & "" 'dr(101) & ""
            txt26PRPAN.Text = nds.Tables(0).Rows(i)(102) & "" 'dr(102) & ""
            txt27PRPAN.Text = nds.Tables(0).Rows(i)(103) & "" 'dr(103) & ""
            txt27EPRPAN.Text = nds.Tables(0).Rows(i)(104) & "" 'dr(104) & ""

            Try
                frmGovDet.cboGovtState.SelectedValue = nds.Tables(0).Rows(i)(70) & "" 'dr(70) & ""
            Catch
                frmGovDet.cboGovtState.SelectedValue = 0
            End Try
            frmGovDet.txtPAOCode.Text = nds.Tables(0).Rows(i)(71) & "" 'dr(71) & ""
            frmGovDet.txtDDOCode.Text = nds.Tables(0).Rows(i)(72) & "" 'dr(72) & ""
            Try
                frmGovDet.cboMinistry.SelectedValue = nds.Tables(0).Rows(i)(73) & "" 'dr(73) & ""
            Catch
                frmGovDet.cboMinistry.SelectedValue = 0
            End Try

            frmGovDet.txtMinistryName.Text = nds.Tables(0).Rows(i)(74) & "" 'dr(74) & ""
            frmGovDet.txtPAORegNo.Text = nds.Tables(0).Rows(i)(75) & "" 'dr(75) & ""
            frmGovDet.txtDDORegNo.Text = nds.Tables(0).Rows(i)(76) & "" 'dr(76) & ""
        Next

    End Sub
    Public Sub controlEnabled()
        cmdAlternate.Enabled = True
        If cboGovtDetails.SelectedValue = "A" Or cboGovtDetails.SelectedValue = "S" Or cboGovtDetails.SelectedValue = "D" Or cboGovtDetails.SelectedValue = "E" Or cboGovtDetails.SelectedValue = "N" Or cboGovtDetails.SelectedValue = "L" Or cboGovtDetails.SelectedValue = "G" Or cboGovtDetails.SelectedValue = "H" Then
            cmdGovtDetails.Enabled = True
        End If
        Panel1.Enabled = True
        Panel2.Enabled = True
        Panel3.Enabled = True
        Panel4.Enabled = True
        Panel5.Enabled = True
        txtCoName.Enabled = True
        txtCoBrDiv.Enabled = True
        txtCoAdd1.Enabled = True
        txtCoAdd2.Enabled = True
        txtCoAdd3.Enabled = True
        txtCoAdd5.Enabled = True
        txtCoAdd4.Enabled = True
        cboCoState.Enabled = True
        txtCoPin.Enabled = True
        txtCoSTD.Enabled = True
        txtCoPhone.Enabled = True
        txtCoEmail.Enabled = True
        txtCoTAN.Enabled = True
        txtCoPAN.Enabled = True
        txtGSTIN.Enabled = True
        chkAddChg.Enabled = True
        cboGovtDetails.Enabled = True
        txtName24.Enabled = True
        txtDesg24.Enabled = True
        txt24Add1.Enabled = True
        txt24Add2.Enabled = True
        txt24Add3.Enabled = True
        txt24Add5.Enabled = True
        txt24Add4.Enabled = True
        cboPR24State.Enabled = True
        txtPR24Pin.Enabled = True
        txt24Email.Enabled = True
        txt24STD.Enabled = True
        txt24PHONE.Enabled = True
        chk24AddChg.Enabled = True
        chkUseForm16.Enabled = True
        txtName26.Enabled = True
        txtDesg26.Enabled = True
        txt26add1.Enabled = True
        txt26add2.Enabled = True
        txt26add3.Enabled = True
        txt26add5.Enabled = True
        txt26add4.Enabled = True
        cboPR26State.Enabled = True
        txtPR26Pin.Enabled = True
        txt26Email.Enabled = True
        txt26STD.Enabled = True
        txt26PHONE.Enabled = True
        chk26AddChg.Enabled = True
        txtName27.Enabled = True
        txtDesg27.Enabled = True
        txt27Add1.Enabled = True
        txt27Add2.Enabled = True
        txt27Add3.Enabled = True
        txt27Add5.Enabled = True
        txt27Add4.Enabled = True
        cboPR27State.Enabled = True
        txtPR27Pin.Enabled = True
        txt27Email.Enabled = True
        txt27STD.Enabled = True
        txt27PHONE.Enabled = True
        chk27AddChg.Enabled = True
        txtName27E.Enabled = True
        txtDesg27EQ.Enabled = True
        txt27EAdd1.Enabled = True
        txt27EAdd2.Enabled = True
        txt27EAdd3.Enabled = True
        txt27EAdd5.Enabled = True
        txt27EAdd4.Enabled = True
        cboPR27EState.Enabled = True
        txtPR27EPin.Enabled = True
        txt27EEmail.Enabled = True
        txt27ESTD.Enabled = True
        txt27EPHONE.Enabled = True
        chk27EAddChg.Enabled = True
        chkUseForm16.Enabled = True
        txtmobile.Enabled = True
        txt24PRPAN.Enabled = True
        txt26PRPAN.Enabled = True
        txt27PRPAN.Enabled = True
        txt27EPRPAN.Enabled = True
        cmdCopyAdd.Enabled = True
        cmdcopydetail.Enabled = True
        CmdCopyname.Enabled = True
    End Sub
    Public Sub Cleartext()
        txtCoName.Text = ""
        txtCoBrDiv.Text = ""
        txtCoAdd1.Text = ""
        txtCoAdd2.Text = ""
        txtCoAdd3.Text = ""
        txtCoAdd5.Text = ""
        txtCoAdd4.Text = ""
        cboCoState.SelectedValue = 0
        txtCoPin.Text = ""
        txtCoSTD.Text = ""
        txtCoPhone.Text = ""
        txtCoEmail.Text = ""
        txtCoTAN.Text = ""
        txtCoPAN.Text = ""
        chkAddChg.Checked = False
        cboGovtDetails.Text = ""
        txtName24.Text = ""
        txtDesg24.Text = ""
        txt24Add1.Text = ""
        txt24Add2.Text = ""
        txt24Add3.Text = ""
        txt24Add5.Text = ""
        txt24Add4.Text = ""
        cboPR24State.SelectedValue = 0
        txtPR24Pin.Text = ""
        txt24Email.Text = ""
        txt24STD.Text = ""
        txt24PHONE.Text = ""
        chk24AddChg.Checked = False
        txtName26.Text = ""
        txtDesg26.Text = ""
        txt26add1.Text = ""
        txt26add2.Text = ""
        txt26add3.Text = ""
        txt26add5.Text = ""
        txt26add4.Text = ""
        cboPR26State.SelectedValue = 0
        txtPR26Pin.Text = ""
        txt26Email.Text = ""
        txt26STD.Text = ""
        txt26PHONE.Text = ""
        chk26AddChg.Checked = False
        txtName27.Text = ""
        txtDesg27.Text = ""
        txt27Add1.Text = ""
        txt27Add2.Text = ""
        txt27Add3.Text = ""
        txt27Add5.Text = ""
        txt27Add4.Text = ""
        cboPR27State.SelectedValue = 0
        txtPR27Pin.Text = ""
        txt27Email.Text = ""
        txt27STD.Text = ""
        txt27PHONE.Text = ""
        chk27AddChg.Checked = False
        txtName27E.Text = ""
        txtDesg27EQ.Text = ""
        txt27EAdd1.Text = ""
        txt27EAdd2.Text = ""
        txt27EAdd3.Text = ""
        txt27EAdd5.Text = ""
        txt27EAdd4.Text = ""
        cboPR27EState.SelectedValue = 0
        txtPR27EPin.Text = ""
        txt27EEmail.Text = ""
        txt27ESTD.Text = ""
        txt27EPHONE.Text = ""
        chk27EAddChg.Checked = False
        chkUseForm16.Checked = False
        txtmobile.Text = ""
        txt24PRPAN.Text = ""
        txt26PRPAN.Text = ""
        txt27PRPAN.Text = ""
        txt27EPRPAN.Text = ""
        cboGovtDetails.SelectedValue = 1
        cboGovtDetails.Text = "Central Government"
        frmGovDet.txtDDOCode.Text = ""
        frmGovDet.txtPAOCode.Text = ""
        frmGovDet.txtPAORegNo.Text = ""
        frmGovDet.txtDDORegNo.Text = ""
        frmGovDet.txtMinistryName.Text = ""
        frmGovDet.txtAIN.Text = ""
        frmGovDet.cboGovtState.SelectedValue = 0
        frmGovDet.cboMinistry.SelectedValue = 0
    End Sub
    Public Sub controlDisabled()
        chk26AddChg.Enabled = False
        txtPR26Pin.Enabled = False
        txt26STD.Enabled = False
        txt26PHONE.Enabled = False
        txt26Email.Enabled = False
        cmdAlternate.Enabled = False
        cmdGovtDetails.Enabled = False
        Panel1.Enabled = False
        Panel2.Enabled = False
        Panel3.Enabled = False
        Panel4.Enabled = False
        Panel5.Enabled = False
        txtCoName.Enabled = False
        txtCoBrDiv.Enabled = False
        txtCoAdd1.Enabled = False
        txtCoAdd2.Enabled = False
        txtCoAdd3.Enabled = False
        txtCoAdd5.Enabled = False
        txtCoAdd4.Enabled = False
        cboCoState.Enabled = False
        txtCoPin.Enabled = False
        txtCoSTD.Enabled = False
        txtCoPhone.Enabled = False
        txtCoEmail.Enabled = False
        txtCoTAN.Enabled = False
        txtCoPAN.Enabled = False
        txtGSTIN.Enabled = False
        chkAddChg.Enabled = False
        cboGovtDetails.Enabled = False
        txtName24.Enabled = False
        txtDesg24.Enabled = False
        txt24Add1.Enabled = False
        txt24Add2.Enabled = False
        txt24Add3.Enabled = False
        txt24Add5.Enabled = False
        txt24Add4.Enabled = False
        cboPR24State.Enabled = False
        txtPR24Pin.Enabled = False
        txt24Email.Enabled = False
        txt24STD.Enabled = False
        txt24PHONE.Enabled = False
        chk24AddChg.Enabled = False
        txtName26.Enabled = False
        txtDesg26.Enabled = False
        txt26add1.Enabled = False
        txt26add2.Enabled = False
        txt26add3.Enabled = False
        txt26add5.Enabled = False
        txt26add4.Enabled = False
        cboPR26State.Enabled = False
        txtPR24Pin.Enabled = False
        txt24Email.Enabled = False
        txt24STD.Enabled = False
        txt24PHONE.Enabled = False
        chk24AddChg.Enabled = False
        txtName27.Enabled = False
        txtDesg27.Enabled = False
        txt27Add1.Enabled = False
        txt27Add2.Enabled = False
        txt27Add3.Enabled = False
        txt27Add5.Enabled = False
        txt27Add4.Enabled = False
        cboPR27State.Enabled = False
        txtPR27Pin.Enabled = False
        txt27Email.Enabled = False
        txt27STD.Enabled = False
        txt27PHONE.Enabled = False
        chk27AddChg.Enabled = False
        txtName27E.Enabled = False
        txtDesg27EQ.Enabled = False
        txt27EAdd1.Enabled = False
        txt27EAdd2.Enabled = False
        txt27EAdd3.Enabled = False
        txt27EAdd5.Enabled = False
        txt27EAdd4.Enabled = False
        cboPR27EState.Enabled = False
        txtPR27EPin.Enabled = False
        txt27EEmail.Enabled = False
        txt27ESTD.Enabled = False
        txt27EPHONE.Enabled = False
        chk27EAddChg.Enabled = False
        chkUseForm16.Enabled = False
        txtmobile.Enabled = False
        txt24PRPAN.Enabled = False
        txt26PRPAN.Enabled = False
        txt27PRPAN.Enabled = False
        txt27EPRPAN.Enabled = False
        cmdCopyAdd.Enabled = False
        cmdcopydetail.Enabled = False
        CmdCopyname.Enabled = False
    End Sub
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        ' Dim Olditem As ListViewItem
        If tempadd = 1 Then
            If SavedData() = False Then
                Exit Sub
            End If
            'NormalMode()
            tempadd = 0
            tempEdit = 0
            controlDisabled()
            lvwCo.Enabled = True
            cmdSave.Enabled = False
            cmdCoAdd.Enabled = True
            cmdBackup.Enabled = True
            cmdCoEdit.Enabled = True
            cmdTDS.Enabled = True
            cmdRestore.Enabled = True
            cmdExit.Text = "Exit"
            lvwCo.Clear()
            setupListView()
            connectlistview()
            If lvwCo.Items.Count > 0 Then
                lvwCo.Items(0).Selected = True
                lvwCo.Select()
            End If
        End If
        If tempEdit = 1 Then
            If BeforeSave() = False Then
                Exit Sub
            End If
            Dataupdate()
            'NormalMode()
            tempadd = 0
            tempEdit = 0
            controlDisabled()
            lvwCo.Enabled = True
            cmdSave.Enabled = False
            cmdCoAdd.Enabled = True
            cmdBackup.Enabled = True
            cmdCoEdit.Enabled = True
            cmdTDS.Enabled = True
            cmdRestore.Enabled = True
            lvwCo.Clear()
            setupListView()
            connectlistview()
            ' Olditem = lvwCo.Items(0)
            'lvwCo.SelectedItems(0).Selected = True
            ' oCoMst.FillCosInLvw(lvwCo)
            cmdExit.Text = "Exit"

        End If
        'connectcontols()
        lvwCo.Items(0).Selected = True
        lvwCo.Select()
        'SelectedId = lvwCo.SelectedItems(0).SubItems(1).Text
        If lvwCo.SelectedIndices.Count < 0 Then
            selectedcoid = lvwCo.SelectedItems(0).SubItems(1).Text
        End If
        ConnectData()
    End Sub
    Private Sub cmdAltCantact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAlternate.Click
        Dim frmAltDet As New frmAltDetails
        frmAltDet.Show()
    End Sub


    Private Function BeforeSave() As Boolean
        If Trim(txtCoName.Text) = vbNullString Then
            Call MsgBox("Company Name Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtCoName.Focus()
            Exit Function
        End If
        If Trim(txtCoAdd1.Text) = vbNullString Then
            Call MsgBox("Company Address Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtCoAdd1.Focus()
            Exit Function
        End If
        If Trim(cboCoState.Text) = vbNullString Then
            Call MsgBox("Company State Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            cboCoState.Focus()
            Exit Function
        End If
        If Trim(txtCoPin.Text) = vbNullString Then
            Call MsgBox("Company Pin Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtCoPin.Focus()
            Exit Function
        End If
        Dim PANErr As Integer

        If Trim(cboGovtDetails.Text) <> vbNullString And (cboGovtDetails.Items.Count > -1 And cboGovtDetails.Items.Count >= 7) Then
            PANErr = IsValidPAN(txtCoPAN.Text, True, False)
        End If

        ' PANErr = IsValidTAN(txtCoTAN.Text, False, False)
        If PANErr <> 0 Then
            Select Case PANErr
                Case 1
                    Call MsgBox("Length of PAN is invalid, please enter 10 digit valid PAN.", vbExclamation, "Invalid PAN")
                Case 2
                    Call MsgBox("The string you have entered is invalid, enter a valid PAN", vbExclamation, "Invalid PAN")
                Case 3
                    Call MsgBox("The format of the PAN number you entered is invalid. ", vbExclamation, "Invalid PAN")
                Case 4
                    Call MsgBox("The last character in PAN is invalid. Please check and enter correct data.", vbExclamation, "Invalid PAN")
            End Select
            BeforeSave = False
            txtCoPAN.Focus()
            Exit Function
        End If

        Dim TANErr As Integer
        TANErr = IsValidTAN(txtCoTAN.Text, False, False)
        If TANErr <> 0 Then
            Select Case TANErr
                Case 1
                    Call MsgBox("Length of TAN is invalid, please enter 10 digit valid TAN.", vbExclamation, "Invalid TAN")
                Case 2
                    Call MsgBox("The string you have entered is invalid, enter a valid TAN", vbExclamation, "Invalid TAN")
                Case 3
                    Call MsgBox("The format of the TAN number you entered is invalid. ", vbExclamation, "Invalid TAN")
                Case 4
                    Call MsgBox("The last character in TAN is invalid. Please check and enter correct data.", vbExclamation, "Invalid TAN")
            End Select
            BeforeSave = False
            txtCoTAN.Focus()
            Exit Function
        End If
        'If optGovt.Value = False And OptOthers.Value = False Then
        '        Call MsgBox("Select Status Of the Company!", vbInformation, "Caution")
        '        Cancel = True
        '        optGovt.SetFocus
        '        Exit Sub
        'End If

        If Trim(cboGovtDetails.Text) = vbNullString Then
            Call MsgBox("Select Deducter Type!", vbInformation, "Caution")
            BeforeSave = False
            cboGovtDetails.Focus()
            Exit Function
        End If


        ''If OptOthers.Value = True Then
        ''    If txtCoPAN = vbNullString Then
        ''        Call MsgBox("PAN for Other Company is Compulsory!", vbInformation, "Caution")  ' comment by jayshree
        ''        Cancel = True
        ''        txtCoPAN.SetFocus
        ''        Exit Sub
        ''    End If
        ''End If


        If Trim(txtName24.Text) = vbNullString Then
            Call MsgBox("Form No.24 Name Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtName24.Focus()
            Exit Function
        End If
        If Trim(txtDesg24.Text) = vbNullString Then
            Call MsgBox("No.24 Designation Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtDesg24.Focus()
            Exit Function
        End If
        If Trim(txtName26.Text) = vbNullString Then
            Call MsgBox("Form No.26 Name Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txt26add4.Focus()
            Exit Function
        End If
        If Trim(txtDesg26.Text) = vbNullString Then
            Call MsgBox("No.26 Designation Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtDesg26.Focus()
            Exit Function
        End If
        If Trim(txtName27.Text) = vbNullString Then
            Call MsgBox("Form No.27 Name Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtName27.Focus()
            Exit Function
        End If
        If Trim(txtDesg27.Text) = vbNullString Then
            Call MsgBox("No.27 Designation Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtDesg27.Focus()
            Exit Function
        End If

        If Trim(txtName27E.Text) = vbNullString Then
            Call MsgBox("Form No.27E Name Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtName27E.Focus()
            Exit Function
        End If

        If Trim(txtDesg27EQ.Text) = vbNullString Then
            Call MsgBox("Form No.27E Designation Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtDesg27EQ.Focus()
            Exit Function
        End If

        If Trim(txt24Add1.Text) = vbNullString And Trim(txt24Add2.Text) = vbNullString And Trim(txt24Add3.Text) = vbNullString And Trim(txt24Add5.Text) = vbNullString And Trim(txt24Add4.Text) = vbNullString Then
            Call MsgBox("Form24Q Person Repsonsilble Address Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txt24Add1.Focus()
            Exit Function
        End If
        If Trim(cboPR24State.Text) = vbNullString Then
            Call MsgBox("Form24Q State Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            cboPR24State.Focus()
            Exit Function
        End If
        If Trim(txtPR24Pin.Text) = vbNullString Then
            Call MsgBox("Form24Q Pin Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtPR24Pin.Focus()
            Exit Function
        End If

        If Trim(txt26add1.Text) = vbNullString And Trim(txt26add2.Text) = vbNullString And Trim(txt26add3.Text) = vbNullString And Trim(txt26add5.Text) = vbNullString And Trim(txt26add4.Text) = vbNullString Then
            Call MsgBox("Form26Q Person Repsonsilble Address Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txt26add1.Focus()
            Exit Function
        End If
        If Trim(cboPR26State.Text) = vbNullString Then
            Call MsgBox("Form26Q State Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            cboPR26State.Focus()
            Exit Function
        End If
        If Trim(txtPR26Pin.Text) = vbNullString Then
            Call MsgBox("Form26Q Pin Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtPR26Pin.Focus()
            Exit Function
        End If
        If Trim(txt27Add1.Text) = vbNullString And Trim(txt27Add2.Text) = vbNullString And Trim(txt27Add3.Text) = vbNullString And Trim(txt27Add5.Text) = vbNullString And Trim(txt27Add4.Text) = vbNullString Then
            Call MsgBox("Form27Q Person Repsonsilble Address Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txt27Add1.Focus()
            Exit Function
        End If
        If Trim(cboPR27State.Text) = vbNullString Then
            Call MsgBox("Form27Q State Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            cboPR27State.Focus()
            Exit Function
        End If
        If Trim(txtPR27Pin.Text) = vbNullString Then
            Call MsgBox("Form27Q Pin Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtPR27Pin.Focus()
            Exit Function
        End If
        If Trim(txt27EAdd1.Text) = vbNullString And Trim(txt27EAdd2.Text) = vbNullString And Trim(txt27EAdd3.Text) = vbNullString And Trim(txt27EAdd5.Text) = vbNullString And Trim(txt27EAdd4.Text) = vbNullString Then
            Call MsgBox("Form27EQ Person Repsonsilble Address Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txt27EAdd1.Focus()
            Exit Function
        End If


        If Trim(cboPR27EState.Text) = vbNullString Then
            Call MsgBox("Form27EQ State Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            cboPR27EState.Focus()
            Exit Function
        End If
        If Trim(txtPR27EPin.Text) = vbNullString Then
            Call MsgBox("Form27EQ Pin Cannot Be Blank!", vbInformation, "Caution")
            BeforeSave = False
            txtPR27EPin.Focus()
            Exit Function
        End If

        '       If grdAdd.TextMatrix(0, 6) = vbNullString Then
        '           Call MsgBox("Form24Q Pin Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd24Add.SetFocus
        '           Exit Sub
        '       End If
        '       If grdAdd.TextMatrix(1, 0) = vbNullString Then
        '           Call MsgBox("Form26Q Person Repsonsilble Address Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd26Add.SetFocus
        '           Exit Sub
        '       End If
        '       If grdAdd.TextMatrix(1, 5) = vbNullString Then
        '           Call MsgBox("Form26Q State Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd26Add.SetFocus
        '           Exit Sub
        '       End If
        '       If grdAdd.TextMatrix(1, 6) = vbNullString Then
        '           Call MsgBox("Form26Q Pin Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd26Add.SetFocus
        '           Exit Sub
        '       End If
        '       If grdAdd.TextMatrix(2, 0) = vbNullString Then
        '           Call MsgBox("Form27Q Person Repsonsilble Address Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd27Add.SetFocus
        '           Exit Sub
        '       End If
        '       If grdAdd.TextMatrix(2, 5) = vbNullString Then
        '           Call MsgBox("Form27Q State Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd27Add.SetFocus
        '           Exit Sub
        '       End If
        '       If grdAdd.TextMatrix(2, 6) = vbNullString Then
        '           Call MsgBox("Form27Q Pin Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd27Add.SetFocus
        '           Exit Sub
        '       End If
        '       If grdAdd.TextMatrix(3, 0) = vbNullString Then
        '           Call MsgBox("Form27EQ Person Repsonsilble Address Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd27EAdd.SetFocus
        '           Exit Sub
        '       End If
        '       If grdAdd.TextMatrix(3, 5) = vbNullString Then
        '           Call MsgBox("Form27EQ State Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd27EAdd.SetFocus
        '           Exit Sub
        '       End If
        '       If grdAdd.TextMatrix(3, 6) = vbNullString Then
        '           Call MsgBox("Form27EQ Pin Cannot Be Blank!", vbInformation, "Caution")
        '           Cancel = True
        '           cmd27EAdd.SetFocus
        '           Exit Sub
        '       End If
        If (cboGovtDetails.SelectedIndex > -1 And cboGovtDetails.SelectedIndex <= 7) Then 'Govt Deductor
            'check state name; mandatory for SEHN, no value for others
            If cboGovtDetails.SelectedIndex = 1 Or cboGovtDetails.SelectedIndex = 3 Or
        cboGovtDetails.SelectedIndex = 5 Or cboGovtDetails.SelectedIndex = 7 Then
                If frmGovDet.cboGovtState.SelectedIndex = -1 Then
                    BeforeSave = False
                    MsgBox("Please select State from Govt. Department Details form")
                    Exit Function
                End If
            Else
                'force blank for others...
                frmGovDet.cboGovtState.SelectedIndex = -1
            End If

            'check PAO code and DDO code, mandatory for A, optional for SDEGHLN
            If cboGovtDetails.SelectedIndex = 0 Then
                If Trim(frmGovDet.txtPAOCode.Text) = vbNullString Then
                    MsgBox("PAO code cannot be left blank, please fill details")
                    BeforeSave = False
                    Exit Function
                End If
                If Trim(frmGovDet.txtDDOCode.Text) = vbNullString Then
                    MsgBox("DDO code cannot be left blank, please fill details")
                    BeforeSave = False
                    Exit Function
                End If
            End If

            'check ministry name, mandatory for ADG, optional for EHLN
            If cboGovtDetails.SelectedIndex = 0 Or cboGovtDetails.SelectedIndex = 2 Or cboGovtDetails.SelectedIndex = 4 Then
                If frmGovDet.cboMinistry.SelectedIndex = -1 Then
                    MsgBox("Please select Ministry, cannot be left blank")
                    BeforeSave = False
                    Exit Function
                ElseIf frmGovDet.cboMinistry.SelectedValue = 99 Then
                    If Trim(frmGovDet.txtMinistryName.Text) = vbNullString Then
                        MsgBox("Ministry name cannot be left blank, please fill details")
                        BeforeSave = False
                        Exit Function
                    End If
                End If
            End If
        Else
            frmGovDet.cboGovtState.SelectedIndex = -1
            frmGovDet.txtPAOCode.Text = ""
            frmGovDet.txtDDOCode.Text = ""
            frmGovDet.txtPAORegNo.Text = ""
            frmGovDet.txtDDORegNo.Text = ""
            frmGovDet.cboMinistry.SelectedIndex = -1
        End If

        'Following changes done for FVU 3.0 - FY 10-11 onwards .. done by Nitin Betharia on 06/10/2010

        'Atleast one Email ID is required...
        If Trim(txtCoEmail.Text) = "" And Trim(txt24Email.Text) = "" And Trim(txt26Email.Text) = "" _
    And Trim(txt27Email.Text) = "" And Trim(txt27EEmail.Text) = "" Then
            MsgBox("Please provide at least one email id", vbInformation)
            BeforeSave = False
            Exit Function
        End If

        'Atleast one STD-Tel combo is required..first check the combination of STD+TelNo
        If Trim(txtCoSTD.Text) <> "" Then
            'STD is not blank
            If Trim(txtCoPhone.Text) = "" Then
                ' but telephone is blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Company Details)")
                BeforeSave = False
                Exit Function
            End If
        Else
            'STD is blank
            If Trim(txtCoPhone.Text) <> "" Then
                ' but telephone is not blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Company Details)")
                BeforeSave = False
                Exit Function
            End If
        End If

        If Trim(txtmobile.Text) = "" Then
            MsgBox("Please provide Mobile Number", vbInformation)
            BeforeSave = False
            Exit Function
        End If


        If Trim(txt24STD.Text) <> "" Then
            'STD is not blank
            If Trim(txt24PHONE.Text) = "" Then
                ' but telephone is blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Responsible Person)")
                BeforeSave = False
                Exit Function
            End If
        Else
            'STD is blank
            If Trim(txt24PHONE.Text) <> "" Then
                ' but telephone is not blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Responsible Person)")
                BeforeSave = False
                Exit Function
            End If
        End If
        If Trim(txt26STD.Text) <> "" Then
            'STD is not blank
            If Trim(txt26PHONE.Text) = "" Then
                ' but telephone is blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Responsible Person)")
                BeforeSave = False
                Exit Function
            End If
        Else
            'STD is blank
            If Trim(txt26PHONE.Text) <> "" Then
                ' but telephone is not blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Responsible Person)")
                BeforeSave = False
                Exit Function
            End If
        End If
        If Trim(txt27STD.Text) <> "" Then
            'STD is not blank
            If Trim(txt27PHONE.Text) = "" Then
                ' but telephone is blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Responsible Person)")
                BeforeSave = False
                Exit Function
            End If
        Else
            'STD is blank
            If Trim(txt27PHONE.Text) <> "" Then
                ' but telephone is not blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Responsible Person)")
                BeforeSave = False
                Exit Function
            End If
        End If
        If Trim(txt27ESTD.Text) <> "" Then
            'STD is not blank
            If Trim(txt27EPHONE.Text) = "" Then
                ' but telephone is blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Responsible Person)")
                BeforeSave = False
                Exit Function
            End If
        Else
            'STD is blank
            If Trim(txt27EPHONE.Text) <> "" Then
                ' but telephone is not blank
                MsgBox("Please provide both STD and Telephone no.; or remove both (Responsible Person)")
                BeforeSave = False
                Exit Function
            End If
        End If

        ' Now check that atleast phone no should exist..
        If Trim(txtCoPhone.Text) = "" And Trim(txt24PHONE.Text) = "" And Trim(txt26PHONE.Text) = "" _
    And Trim(txt27PHONE.Text) = "" And Trim(txt27EPHONE.Text) = "" Then
            MsgBox("Please provide at least one STD+Telephone No", vbInformation)
            BeforeSave = False
            Exit Function
        End If
        '       Dim MSGTXT As String
        '       For i = 0 To 3
        '           If Trim(grdAdd.TextMatrix(i, 11)) = "" Then
        '               Select Case i
        '                   Case 0
        '                       MsgBox "Please provide PAN of Responsible Person Form 24Q (Click Address )"
        '        Case 1
        '                       MsgBox "Please provide PAN of Responsible Person Form 26Q (Click Address )"
        '        Case 2
        '                       MsgBox "Please provide PAN of Responsible Person Form 27Q (Click Address )"
        '        Case 3
        '                       MsgBox "Please provide PAN of Responsible Person Form 27EQ (Click Address )"
        'End Select
        '               Cancel = True
        '               Exit Sub
        '           End If
        'Next i
        If Trim(txt24PRPAN.Text) = "" Then
            MsgBox("Please provide PAN of Responsible Person Form 24Q")
            BeforeSave = False
            Exit Function
        End If
        If Trim(txt26PRPAN.Text) = "" Then
            MsgBox("Please provide PAN of Responsible Person Form 26Q")
            BeforeSave = False
            Exit Function
        End If
        If Trim(txt27PRPAN.Text) = "" Then
            MsgBox("Please provide PAN of Responsible Person Form 27Q")
            BeforeSave = False
            Exit Function
        End If
        If Trim(txt27EPRPAN.Text) = "" Then
            MsgBox("Please provide PAN of Responsible Person Form 27EQ")
            BeforeSave = False
            Exit Function
        End If
        BeforeSave = True
    End Function

    Public Function SavedData() As Boolean
        Dim cmd As New OleDbCommand
        Dim sql As String = "SELECT TOP 1 CoID FROM CoMst ORDER BY CoID DESC"
        Dim cmd1 As New OleDbCommand(sql, cn)
        Dim sqlinsert As String
        Dim transaction As OleDbTransaction
        If BeforeSave() = False Then
            cmd1.Dispose()
            SavedData = False
            Exit Function
        End If
        cmd1.ExecuteNonQuery()
        dr3 = cmd1.ExecuteReader
        While dr3.Read()
            MaxID = dr3(0) + 1
        End While
        If MaxID = 0 Then MaxID = 1
        Dim CoStatus As String
        CoStatus = ""
        Select Case cboGovtDetails.SelectedIndex
            Case 0
                CoStatus = "A"
            Case 1
                CoStatus = "S"
            Case 2
                CoStatus = "D"
            Case 3
                CoStatus = "E"
            Case 4
                CoStatus = "G"
            Case 5
                CoStatus = "H"
            Case 6
                CoStatus = "L"
            Case 7
                CoStatus = "N"
            Case 8
                CoStatus = "K"
            Case 9
                CoStatus = "M"
            Case 10
                CoStatus = "P"
            Case 11
                CoStatus = "T"
            Case 12
                CoStatus = "J"
            Case 13
                CoStatus = "B"
            Case 14
                CoStatus = "Q"
            Case 15
                CoStatus = "F"
        End Select


        sqlinsert = "INSERT INTO CoMst(CoId,CoName,CoBrDiv,CoAdd1,CoAdd2,CoAdd3,CoAdd4,CoAdd5,CoStateID,CoPin,CoTAN,CoPAN,IsCoAddChg,CoEmail,CoStd,CoPhone,CoStatus,PR26Name,PR26Desg,PR27Name,PR27Desg,PR24Name,PR24Desg,PR27EName,PR27EDesg,PR24Add1,PR24Add2,PR24Add3,PR24Add4,PR24Add5,PR24StateID,PR24Pin,IsPR24AddChg,PR24Email,PR24Std,PR24Phone,PR26Add1,PR26Add2,PR26Add3,PR26Add4,PR26Add5,PR26StateID,PR26Pin,IsPR26AddChg,PR26Email,PR26Std,PR26Phone,PR27Add1,PR27Add2,PR27Add3,PR27Add4,PR27Add5,PR27StateID,PR27Pin,IsPR27AddChg,PR27Email,PR27Std,PR27Phone,PR27EAdd1,PR27EAdd2,PR27EAdd3,PR27EAdd4,PR27EAdd5,PR27EStateID,PR27EPin,IsPR27EAddChg,PR27EEmail,PR27EStd,PR27EPhone,UseForm16,GovtStateID,PAOCode,DDOCode,MinistryID,MinistryName,PAORegNo,DDORegNo,mobile,AIN,PR24PAN,PR26PAN,PR27PAN,PR27EPAN) values (" & MaxID & "," & IIf(txtCoName.Text = vbNullString, "Null", "'" & txtCoName.Text & "'") & "," _
          & IIf(txtCoBrDiv.Text = vbNullString, "Null", "'" & txtCoBrDiv.Text & "'") & "," _
          & IIf(txtCoAdd1.Text = vbNullString, "Null", "'" & txtCoAdd1.Text & "'") & "," & IIf(txtCoAdd2.Text = vbNullString, "Null", "'" & txtCoAdd2.Text & "'") & "," _
          & IIf(txtCoAdd3.Text = vbNullString, "Null", "'" & txtCoAdd3.Text & "'") & "," & IIf(txtCoAdd4.Text = vbNullString, "Null", "'" & txtCoAdd4.Text & "'") & "," _
          & IIf(txtCoAdd5.Text = vbNullString, "Null", "'" & txtCoAdd5.Text & "'") & "," & IIf(cboCoState.SelectedValue = Nothing, 0, cboCoState.SelectedValue) & ", " & IIf(txtCoPin.Text = vbNullString, "Null", "'" & txtCoPin.Text & "'") & "," _
          & IIf(txtCoTAN.Text = vbNullString, "Null", "'" & txtCoTAN.Text & "'") & "," & IIf(txtCoPAN.Text = vbNullString, "Null", "'" & txtCoPAN.Text & "'") & "," _
          & IIf(chkAddChg.Checked = True, True, False) & "," _
          & IIf(txtCoEmail.Text = vbNullString, "Null", "'" & txtCoEmail.Text & "'") & "," _
          & IIf(txtCoSTD.Text = vbNullString, "Null", "'" & txtCoSTD.Text & "'") & "," _
          & IIf(txtCoPhone.Text = vbNullString, 0, "'" & txtCoPhone.Text & "'") & "," _
          & IIf(cboGovtDetails.SelectedValue = Nothing, 0, "'" & CoStatus & "'") & "," _
          & IIf(txtName26.Text = vbNullString, "Null", "'" & txtName26.Text & "'") & "," _
          & IIf(txtDesg26.Text = vbNullString, "Null", "'" & txtDesg26.Text & "'") & "," _
          & IIf(txtName27.Text = vbNullString, "Null", "'" & txtName27.Text & "'") & "," _
          & IIf(txtDesg27.Text = vbNullString, "Null", "'" & txtDesg27.Text & "'") & "," _
          & IIf(txtName24.Text = vbNullString, "Null", "'" & txtName24.Text & "'") & "," _
          & IIf(txtDesg24.Text = vbNullString, "Null", "'" & txtDesg24.Text & "'") & "," _
          & IIf(txtName27E.Text = vbNullString, "Null", "'" & txtName27E.Text & "'") & "," _
          & IIf(txtDesg27EQ.Text = vbNullString, "Null", "'" & txtDesg27EQ.Text & "'") & "," _
          & IIf(txt24Add1.Text = vbNullString, "Null", "'" & txt24Add1.Text & "'") & "," _
          & IIf(txt24Add2.Text = vbNullString, "Null", "'" & txt24Add2.Text & "'") & "," _
          & IIf(txt24Add3.Text = vbNullString, "Null", "'" & txt24Add3.Text & "'") & "," _
          & IIf(txt24Add4.Text = vbNullString, "Null", "'" & txt24Add4.Text & "'") & "," _
          & IIf(txt24Add5.Text = vbNullString, "Null", "'" & txt24Add5.Text & "'") & "," _
          & IIf(cboPR24State.SelectedValue = Nothing, 0, cboPR24State.SelectedValue) & "," _
          & IIf(txtPR24Pin.Text = vbNullString, "Null", "'" & txtPR24Pin.Text & "'") & ", " & IIf(chk24AddChg.Checked = True, True, False) & "," _
          & IIf(txt24Email.Text = vbNullString, "Null", "'" & txt24Email.Text & "'") & "," _
          & IIf(txt24STD.Text = vbNullString, "Null", "'" & txt24STD.Text & "'") & "," _
          & IIf(txt24PHONE.Text = vbNullString, 0, txt24PHONE.Text) & "," _
          & IIf(txt26add1.Text = vbNullString, "Null", "'" & txt26add1.Text & "'") & "," _
          & IIf(txt26add2.Text = vbNullString, "Null", "'" & txt26add2.Text & "'") & "," _
          & IIf(txt26add3.Text = vbNullString, "Null", "'" & txt26add3.Text & "'") & "," _
          & IIf(txt26add4.Text = vbNullString, "Null", "'" & txt26add4.Text & "'") & "," _
          & IIf(txt26add5.Text = vbNullString, "Null", "'" & txt26add5.Text & "'") & "," _
          & IIf(cboPR26State.SelectedValue = Nothing, 0, cboPR26State.SelectedValue) & "," _
          & IIf(txtPR26Pin.Text = vbNullString, "Null", "'" & txtPR26Pin.Text & "'") & ", " & IIf(chk26AddChg.Checked = True, True, False) & "," _
          & IIf(txt26Email.Text = vbNullString, "Null", "'" & txt26Email.Text & "'") & "," _
          & IIf(txt26STD.Text = vbNullString, "Null", "'" & txt26STD.Text & "'") & "," _
          & IIf(txt26PHONE.Text = vbNullString, 0, "'" & txt26PHONE.Text & "'") & "," _
          & IIf(txt27Add1.Text = vbNullString, "Null", "'" & txt27Add1.Text & "'") & "," _
          & IIf(txt27Add2.Text = vbNullString, "Null", "'" & txt27Add2.Text & "'") & "," _
          & IIf(txt27Add3.Text = vbNullString, "Null", "'" & txt27Add3.Text & "'") & "," _
          & IIf(txt27Add4.Text = vbNullString, "Null", "'" & txt27Add4.Text & "'") & "," _
          & IIf(txt27Add5.Text = vbNullString, "Null", "'" & txt27Add5.Text & "'") & "," _
          & IIf(cboPR27State.SelectedValue = Nothing, 0, cboPR27State.SelectedValue) & "," _
          & IIf(txtPR27Pin.Text = vbNullString, "Null", "'" & txtPR27Pin.Text & "'") & ", " & IIf(chk27AddChg.Checked = True, True, False) & "," _
          & IIf(txt27Email.Text = vbNullString, "Null", "'" & txt27Email.Text & "'") & "," _
          & IIf(txt27STD.Text = vbNullString, "Null", "'" & txt27STD.Text & "'") & "," _
          & IIf(txt27PHONE.Text = vbNullString, 0, "'" & txt27PHONE.Text & "'") & "," _
          & IIf(txt27EAdd1.Text = vbNullString, "Null", "'" & txt27EAdd1.Text & "'") & "," _
          & IIf(txt27EAdd2.Text = vbNullString, "Null", "'" & txt27EAdd2.Text & "'") & "," _
          & IIf(txt27EAdd3.Text = vbNullString, "Null", "'" & txt27EAdd3.Text & "'") & "," _
          & IIf(txt27EAdd4.Text = vbNullString, "Null", "'" & txt27EAdd4.Text & "'") & "," _
          & IIf(txt27EAdd5.Text = vbNullString, "Null", "'" & txt27EAdd5.Text & "'") & "," _
          & IIf(cboPR27EState.SelectedValue = Nothing, 0, cboPR27EState.SelectedValue) & "," _
          & IIf(txtPR27EPin.Text = vbNullString, "Null", txtPR27EPin.Text) & ", " & IIf(chk27EAddChg.Checked = True, True, False) & "," _
          & IIf(txt27EEmail.Text = vbNullString, "Null", "'" & txt27EEmail.Text & "'") & "," _
          & IIf(txt27ESTD.Text = vbNullString, "Null", "'" & txt27ESTD.Text & "'") & "," _
          & IIf(txt27EPHONE.Text = vbNullString, 0, "'" & txt27EPHONE.Text & "'") & "," _
          & IIf(chkUseForm16.Checked = True, True, False) & "," _
          & IIf(frmGovDet.cboGovtState.Text = Nothing, "Null", frmGovDet.cboGovtState.SelectedValue) & "," _
          & IIf(frmGovDet.txtPAOCode.Text = vbNullString, "Null", "'" & frmGovDet.txtPAOCode.Text & "'") & "," _
          & IIf(frmGovDet.txtDDOCode.Text = vbNullString, "Null", "'" & frmGovDet.txtDDOCode.Text & "'") & "," _
          & IIf(frmGovDet.cboMinistry.Text = Nothing, "Null", frmGovDet.cboMinistry.SelectedValue) & "," _
          & IIf(frmGovDet.txtMinistryName.Text = vbNullString, "Null", "'" & frmGovDet.txtMinistryName.Text & "'") & "," _
          & IIf(frmGovDet.txtPAORegNo.Text = vbNullString, "Null", frmGovDet.txtPAORegNo.Text) & "," _
          & IIf(frmGovDet.txtDDORegNo.Text = vbNullString, "Null", "'" & frmGovDet.txtDDORegNo.Text & "'") & "," & IIf(txtmobile.Text = vbNullString, "Null", "'" & txtmobile.Text & "'") & "," _
          & IIf(frmGovDet.txtAIN.Text = vbNullString, "Null", "'" & frmGovDet.txtAIN.Text & "'") & "," _
          & IIf(txt24PRPAN.Text = vbNullString, "Null", "'" & txt24PRPAN.Text & "'") & "," _
          & IIf(txt26PRPAN.Text = vbNullString, "Null", "'" & txt26PRPAN.Text & "'") & "," _
          & IIf(txt27PRPAN.Text = vbNullString, "Null", "'" & txt27PRPAN.Text & "'") & "," _
          & IIf(txt27EPRPAN.Text = vbNullString, "Null", "'" & txt27EPRPAN.Text & "'") & ")"

        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction
        cmd.CommandText = sqlinsert
        Try
            cmd.ExecuteNonQuery()
            transaction.Commit()
            SavedData = True
        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show(ex.Message) 'Error MEssage
            SavedData = False
        End Try
        cmd.Dispose()
        transaction.Dispose()
    End Function
    Public Sub Dataupdate()
        'frmLogin.cn.Open()

        Dim transaction As OleDbTransaction
        Dim sqlupdate As String
        Dim cmd As New OleDbCommand

        sqlupdate = "Update CoMSt Set CoName = " & IIf(txtCoName.Text = vbNullString, "Null", "'" & txtCoName.Text & "'") & ", " _
            & " CoBrDiv = " & IIf(txtCoBrDiv.Text = vbNullString, "Null", "'" & txtCoBrDiv.Text & "'") & ", " _
            & " CoAdd1 = " & IIf(txtCoAdd1.Text = vbNullString, "Null", "'" & txtCoAdd1.Text & "'") & ", " _
            & " CoAdd2 = " & IIf(txtCoAdd2.Text = vbNullString, "Null", "'" & txtCoAdd2.Text & "'") & ", " _
            & " CoAdd3 = " & IIf(txtCoAdd3.Text = vbNullString, "Null", "'" & txtCoAdd3.Text & "'") & ", " _
            & " CoAdd4 = " & IIf(txtCoAdd4.Text = vbNullString, "Null", "'" & txtCoAdd4.Text & "'") & ", " _
            & " CoAdd5 = " & IIf(txtCoAdd5.Text = vbNullString, "Null", "'" & txtCoAdd5.Text & "'") & ", " _
            & " CoStateID =" & IIf(cboCoState.SelectedValue = Nothing, 0, cboCoState.SelectedValue) & "," _
            & " CoTAN = " & IIf(txtCoTAN.Text = vbNullString, "Null", "'" & txtCoTAN.Text & "'") & "," _
            & " CoPAN = " & IIf(txtCoPAN.Text = vbNullString, "Null", "'" & txtCoPAN.Text & "'") & "," _
            & " IsCoAddChg = " & IIf(chkAddChg.Checked = True, True, False) & "," _
            & " CoStatus = " & "'" & cboGovtDetails.SelectedValue & "'" & "," _
            & " CoEmail = " & IIf(txtCoEmail.Text = vbNullString, "Null", "'" & txtCoEmail.Text & "'") & "," _
            & " CoStd = " & IIf(txtCoSTD.Text = vbNullString, "Null", "'" & txtCoSTD.Text & "'") & "," _
            & " CoPhone = " & IIf(txtCoPhone.Text = vbNullString, "Null", "'" & txtCoPhone.Text & "'") & "," _
            & " PR26Name = " & IIf(txtName26.Text = vbNullString, "Null", "'" & txtName26.Text & "'") & "," _
            & " PR26Desg = " & IIf(txtDesg26.Text = vbNullString, "Null", "'" & txtDesg26.Text & "'") & "," _
            & " PR27Name = " & IIf(txtName27.Text = vbNullString, "Null", "'" & txtName27.Text & "'") & "," _
            & " PR27Desg = " & IIf(txtDesg27.Text = vbNullString, "Null", "'" & txtDesg27.Text & "'") & ","

        sqlupdate = sqlupdate & " PR24Name = " & IIf(txtName24.Text = vbNullString, "Null", "'" & txtName24.Text & "'") & "," _
        & " PR24Desg = " & IIf(txtDesg24.Text = vbNullString, "Null", "'" & txtDesg24.Text & "'") & "," _
        & " PR27EName = " & IIf(txtName27E.Text = vbNullString, "Null", "'" & txtName27E.Text & "'") & "," _
        & " PR27EDesg = " & IIf(txtDesg27EQ.Text = vbNullString, "Null", "'" & txtDesg27EQ.Text & "'") & "," _
        & " PR24Add1 = " & IIf(txt24Add1.Text = vbNullString, "Null", "'" & txt24Add1.Text & "'") & "," _
        & " PR24Add2 = " & IIf(txt24Add2.Text = vbNullString, "Null", "'" & txt24Add2.Text & "'") & "," _
        & " PR24Add3 = " & IIf(txt24Add3.Text = vbNullString, "Null", "'" & txt24Add3.Text & "'") & "," _
        & " PR24Add4 = " & IIf(txt24Add4.Text = vbNullString, "Null", "'" & txt24Add4.Text & "'") & "," _
        & " PR24Add5 = " & IIf(txt24Add5.Text = vbNullString, "Null", "'" & txt24Add5.Text & "'") & "," _
        & " PR24StateID =" & IIf(cboPR24State.SelectedValue = Nothing, 0, cboPR24State.SelectedValue) & "," _
        & " PR24Pin = " & IIf(txtPR24Pin.Text = vbNullString, "Null", txtPR24Pin.Text) & ", IsPR24AddChg = " & IIf(chk24AddChg.Checked = True, True, False) & "," _
        & " PR24Email = " & IIf(txt24Email.Text = vbNullString, "Null", "'" & txt24Email.Text & "'") & "," _
        & " PR24Std = " & IIf(txt24STD.Text = vbNullString, "Null", "'" & txt24STD.Text & "'") & "," _
        & " PR24Phone = " & IIf(txt24PHONE.Text = vbNullString, "Null", "'" & txt24PHONE.Text & "'") & ","

        sqlupdate = sqlupdate & " PR26Add1 = " & IIf(txt26add1.Text = vbNullString, "Null", "'" & txt26add1.Text & "'") & "," _
        & " PR26Add2 = " & IIf(txt26add2.Text = vbNullString, "Null", "'" & txt26add2.Text & "'") & "," _
        & " PR26Add3 = " & IIf(txt26add3.Text = vbNullString, "Null", "'" & txt26add3.Text & "'") & "," _
        & " PR26Add4 = " & IIf(txt26add4.Text = vbNullString, "Null", "'" & txt26add4.Text & "'") & "," _
        & " PR26Add5 = " & IIf(txt26add5.Text = vbNullString, "Null", "'" & txt26add5.Text & "'") & "," _
        & " PR26StateID = " & IIf(cboPR26State.SelectedValue = Nothing, 0, cboPR26State.SelectedValue) & "," _
        & " PR26Pin = " & IIf(txtPR26Pin.Text = vbNullString, "Null", txtPR26Pin.Text) & ", IsPR26AddChg = " & IIf(chk24AddChg.Checked = True, True, False) & "," _
        & " PR26Email = " & IIf(txt26Email.Text = vbNullString, "Null", "'" & txt26Email.Text & "'") & "," _
        & " PR26Std = " & IIf(txt26STD.Text = vbNullString, "Null", "'" & txt26STD.Text & "'") & "," _
        & " PR26Phone = " & IIf(txt26PHONE.Text = vbNullString, "Null", "'" & txt26PHONE.Text & "'") & ","

        sqlupdate = sqlupdate & " PR27Add1 = " & IIf(txt27Add1.Text = vbNullString, "Null", "'" & txt27Add1.Text & "'") & "," _
        & " PR27Add2 = " & IIf(txt27Add2.Text = vbNullString, "Null", "'" & txt27Add2.Text & "'") & "," _
        & " PR27Add3 = " & IIf(txt27Add3.Text = vbNullString, "Null", "'" & txt27Add3.Text & "'") & "," _
        & " PR27Add4 = " & IIf(txt27Add4.Text = vbNullString, "Null", "'" & txt27Add4.Text & "'") & "," _
        & " PR27Add5 = " & IIf(txt27Add5.Text = vbNullString, "Null", "'" & txt27Add5.Text & "'") & "," _
        & " PR27StateID = " & IIf(cboPR27State.SelectedValue = Nothing, 0, cboPR27State.SelectedValue) & "," _
        & " PR27Pin = " & IIf(txtPR27Pin.Text = vbNullString, "Null", txtPR27Pin.Text) & ", IsPR27AddChg = " & IIf(chk27AddChg.Checked = True, True, False) & "," _
        & " PR27Email = " & IIf(txt27Email.Text = vbNullString, "Null", "'" & txt27Email.Text & "'") & "," _
        & " PR27Std = " & IIf(txt27STD.Text = vbNullString, "Null", "'" & txt27STD.Text & "'") & "," _
        & " PR27Phone = " & IIf(txt27PHONE.Text = vbNullString, "Null", "'" & txt27PHONE.Text & "'") & ","


        sqlupdate = sqlupdate & " PR27EAdd1 = " & IIf(txt27EAdd1.Text = vbNullString, "Null", "'" & txt27EAdd1.Text & "'") & "," _
        & " PR27EAdd2 = " & IIf(txt27EAdd2.Text = vbNullString, "Null", "'" & txt27EAdd2.Text & "'") & "," _
        & " PR27EAdd3 = " & IIf(txt27EAdd3.Text = vbNullString, "Null", "'" & txt27EAdd3.Text & "'") & "," _
        & " PR27EAdd4 = " & IIf(txt27EAdd4.Text = vbNullString, "Null", "'" & txt27EAdd4.Text & "'") & "," _
        & " PR27EAdd5 = " & IIf(txt27EAdd5.Text = vbNullString, "Null", "'" & txt27EAdd5.Text & "'") & "," _
        & " PR27EStateID = " & IIf(cboPR27EState.SelectedValue = Nothing, -1, cboPR27EState.SelectedValue) & "," _
        & " PR27EPin = " & IIf(txtPR27EPin.Text = vbNullString, "Null", txtPR27EPin.Text) & ", IsPR27EAddChg = " & IIf(chk27EAddChg.Checked = True, True, False) & "," _
        & " PR27EEmail = " & IIf(txt27EEmail.Text = vbNullString, "Null", "'" & txt27EEmail.Text & "'") & "," _
        & " PR27EStd = " & IIf(txt27ESTD.Text = vbNullString, "Null", "'" & txt27ESTD.Text & "'") & "," _
        & " PR27EPhone = " & IIf(txt27EPHONE.Text = vbNullString, "", "'" & txt27EPHONE.Text & "'") & "," _
        & " UseForm16 = " & IIf(chkUseForm16.Checked = True, True, False) & "," _
        & " GovtStateID = " & IIf(frmGovDet.cboGovtState.Text = Nothing, "Null", frmGovDet.cboGovtState.SelectedValue) & "," _
        & " PAOCode = " & IIf(frmGovDet.txtPAOCode.Text = vbNullString, "Null", "'" & frmGovDet.txtPAOCode.Text & "'") & "," _
        & " DDOCode = " & IIf(frmGovDet.txtDDOCode.Text = vbNullString, "Null", "'" & frmGovDet.txtDDOCode.Text & "'") & "," _
        & " MinistryID = " & IIf(frmGovDet.cboMinistry.Text = Nothing, "Null", frmGovDet.cboMinistry.SelectedValue) & "," _
        & " MinistryName = " & IIf(frmGovDet.txtMinistryName.Text = vbNullString, "Null", "'" & frmGovDet.txtMinistryName.Text & "'") & "," _
        & " PAORegNo = " & IIf(frmGovDet.txtPAORegNo.Text = vbNullString, "Null", frmGovDet.txtPAORegNo.Text) & "," _
        & " DDORegNo = " & IIf(frmGovDet.txtDDORegNo.Text = vbNullString, "Null", "'" & frmGovDet.txtDDORegNo.Text & "'") & "," _
        & " Mobile = " & IIf(txtmobile.Text = vbNullString, "Null", "'" & txtmobile.Text & "'") & "," _
        & " AIN = " & IIf(frmGovDet.txtAIN.Text = vbNullString, "Null", "'" & frmGovDet.txtAIN.Text & "'") & "," _
        & " PR24PAN = " & IIf(txt24PRPAN.Text = vbNullString, "Null", "'" & txt24PRPAN.Text & "'") & "," _
        & " PR26PAN = " & IIf(txt26PRPAN.Text = vbNullString, "Null", "'" & txt26PRPAN.Text & "'") & "," _
        & " PR27PAN = " & IIf(txt27PRPAN.Text = vbNullString, "Null", "'" & txt27PRPAN.Text & "'") & "," _
        & " PR27EPAN = " & IIf(txt27EPRPAN.Text = vbNullString, "Null", "'" & txt27EPRPAN.Text & "'") _
        & " Where CoId = " & selectedcoid

        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction
        cmd.CommandText = sqlupdate

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
    Private Sub frmCoMst_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Me.EnterTab(e)

    End Sub
    Public Sub EnterTab(ByVal e As KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub frmCoMst_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        frmLogin.Close()
    End Sub
    Private Sub Fill_StatusDed()

        Dim nds As New DataSet

        Dim QueSt As String = "Select DeductorType,DeductorTypeDescription from DeductorTypeMst"
        nds = FetchDataSet(QueSt)
        If nds.Tables(0).Rows.Count > 0 Then
            cboGovtDetails.DataSource = nds.Tables(0)
            cboGovtDetails.ValueMember = "DeductorType"
            cboGovtDetails.DisplayMember = "DeductorTypeDescription"
        End If
        nds.Dispose()
        'frmLogin.cn.Close()
    End Sub


    Private Sub cmdGovtDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGovtDetails.Click
        If cboGovtDetails.SelectedValue = "A" Or cboGovtDetails.SelectedValue = "D" Or cboGovtDetails.SelectedValue = "L" Or cboGovtDetails.SelectedValue = "G" Then
            frmGovDet.cboGovtState.Enabled = False
            frmGovDet.txtPAOCode.Focus()
        Else
            frmGovDet.cboGovtState.Enabled = True
            frmGovDet.cboGovtState.Focus()
        End If
        frmGovDet.Show()
    End Sub

    Private Sub cmbDeduStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGovtDetails.SelectedIndexChanged
        If Mode <> "X" Then
            Try
                If cboGovtDetails.SelectedValue = "A" Or cboGovtDetails.SelectedValue = "S" Or cboGovtDetails.SelectedValue = "D" Or cboGovtDetails.SelectedValue = "E" Or cboGovtDetails.SelectedValue = "N" Or cboGovtDetails.SelectedValue = "L" Or cboGovtDetails.SelectedValue = "G" Or cboGovtDetails.SelectedValue = "H" Then
                    cmdGovtDetails.Enabled = True
                Else
                    cmdGovtDetails.Enabled = False
                End If
            Catch
            End Try
        End If
    End Sub
    Dim WithEvents tm As New Timer With {.Interval = 1000, .Enabled = True}

    Private Sub Tm_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tm.Tick
        ToolStripStatusLabel4.Text = Now.ToString("hh:mm:ss tt")
        ToolStripStatusLabel3.Text = Now.ToString("dd/MM/yy")
        CheckStatus()
    End Sub

    Private Sub CheckStatus()
        If Control.IsKeyLocked(Keys.Scroll) = True Then
            StatusScrl.Enabled = True
        Else
            StatusScrl.Enabled = False
        End If

        If Control.IsKeyLocked(Keys.CapsLock) = True Then
            StatusCaps.Enabled = True
        Else
            StatusCaps.Enabled = False
        End If

        If Control.IsKeyLocked(Keys.NumLock) = True Then
            StatusNum.Enabled = True
        Else
            StatusNum.Enabled = False
        End If

        If Control.IsKeyLocked(Keys.Insert) = True Then
            statusINS.Enabled = True
        Else
            statusINS.Enabled = False
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Private Sub txtCoName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoName.Enter
        txtCoName.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoName.Leave
        txtCoName.BackColor = Color.White
    End Sub

    Private Sub txtCoBrDiv_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoBrDiv.Leave
        txtCoBrDiv.BackColor = Color.White
    End Sub

    Private Sub txtCoBrDiv_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoBrDiv.Enter
        txtCoBrDiv.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoAdd1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd1.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoAdd1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd1.BackColor = Color.White
    End Sub

    Private Sub txtCoAdd2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd2.BackColor = Color.White
    End Sub

    Private Sub txtCoAdd2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd2.BackColor = Color.LightYellow
    End Sub



    Private Sub txtCoAdd3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd3.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoAdd3_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd3.BackColor = Color.White
    End Sub

    Private Sub txtCoAdd4_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd5.BackColor = Color.White
    End Sub

    Private Sub txtCoAdd4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd5.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoAdd5_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd4.BackColor = Color.White
    End Sub

    Private Sub txtCoAdd5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoAdd4.BackColor = Color.LightYellow
    End Sub


    Private Sub cboCoState_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCoState.Leave
        cboCoState.BackColor = Color.White
    End Sub

    Private Sub cboCoState_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCoState.Enter
        cboCoState.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoPin_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoPin.Leave
        txtCoPin.BackColor = Color.White
    End Sub

    Private Sub txtCoPin_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoPin.Enter
        txtCoPin.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoEmail_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoEmail.Enter
        txtCoEmail.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoEmail_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoEmail.Leave
        txtCoEmail.BackColor = Color.White
    End Sub

    Private Sub txtmobile_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtmobile.Leave
        txtmobile.BackColor = Color.White
    End Sub

    Private Sub txtmobile_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtmobile.Enter
        txtmobile.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoSTD_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoSTD.Leave
        txtCoSTD.BackColor = Color.White
    End Sub

    Private Sub txtCoSTD_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoSTD.Enter
        txtCoSTD.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoPhone_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoPhone.BackColor = Color.White
    End Sub

    Private Sub txtCoPhone_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtCoPhone.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoTAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoTAN.Leave
        txtCoTAN.BackColor = Color.White
    End Sub

    Private Sub txtCoTAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoTAN.Enter
        txtCoTAN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtCoPAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoPAN.Leave
        txtCoPAN.BackColor = Color.White
    End Sub

    Private Sub txtCoPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoPAN.Enter
        txtCoPAN.BackColor = Color.LightYellow
    End Sub

    Private Sub cboGovtDetails_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGovtDetails.Leave
        cboGovtDetails.BackColor = Color.White
    End Sub

    Private Sub cboGovtDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGovtDetails.Enter
        cboGovtDetails.BackColor = Color.LightYellow
    End Sub

    Private Sub txtName24_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName24.Enter
        txtName24.BackColor = Color.LightYellow
    End Sub

    Private Sub txtName24_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName24.Leave
        txtName24.BackColor = Color.White
    End Sub

    Private Sub txtDesg24_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles txtDesg24.Layout
        txtDesg24.BackColor = Color.White
    End Sub

    Private Sub txtDesg24_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesg24.Enter
        txtDesg24.BackColor = Color.LightYellow
    End Sub

    Private Sub txt24Add1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add1.BackColor = Color.LightYellow
    End Sub

    Private Sub txt24Add1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add1.BackColor = Color.White
    End Sub

    Private Sub txt24Add2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add2.BackColor = Color.LightYellow
    End Sub

    Private Sub txt24Add2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add2.BackColor = Color.White
    End Sub

    Private Sub txt24Add3_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add3.BackColor = Color.White
    End Sub

    Private Sub txt24Add3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add3.BackColor = Color.LightYellow
    End Sub

    Private Sub txt24Add4_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add5.BackColor = Color.White
    End Sub

    Private Sub txt24Add4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add5.BackColor = Color.LightYellow
    End Sub

    Private Sub txt24Add5_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add4.BackColor = Color.White
    End Sub

    Private Sub txt24Add5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt24Add4.BackColor = Color.LightYellow
    End Sub

    Private Sub cboPR24State_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPR24State.Enter
        cboPR24State.BackColor = Color.LightYellow
    End Sub

    Private Sub cboPR24State_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPR24State.Leave
        cboPR24State.BackColor = Color.White
    End Sub

    Private Sub txtPR24Pin_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPR24Pin.Enter
        txtPR24Pin.BackColor = Color.LightYellow
    End Sub

    Private Sub txtPR24Pin_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPR24Pin.Leave
        txtPR24Pin.BackColor = Color.White
    End Sub

    Private Sub txt24STD_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt24STD.Leave
        txt24STD.BackColor = Color.White
    End Sub

    Private Sub txt24STD_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt24STD.Enter
        txt24STD.BackColor = Color.LightYellow
    End Sub

    Private Sub txt24PHONE_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt24PHONE.Enter
        txt24PHONE.BackColor = Color.LightYellow
    End Sub

    Private Sub txt24PHONE_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt24PHONE.Leave
        txt24PHONE.BackColor = Color.White
    End Sub

    Private Sub txt24Email_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt24Email.Leave
        txt24Email.BackColor = Color.White
    End Sub

    Private Sub txt24Email_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt24Email.Enter
        txt24Email.BackColor = Color.LightYellow
    End Sub

    Private Sub txt24PRPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt24PRPAN.Enter
        txt24PRPAN.BackColor = Color.LightYellow
    End Sub

    Private Sub txt24PRPAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt24PRPAN.Leave
        txt24PRPAN.BackColor = Color.White
    End Sub

    Private Sub txtDesg26_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesg26.Leave
        txtDesg26.BackColor = Color.White
    End Sub

    Private Sub txtDesg26_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesg26.Enter
        txtDesg26.BackColor = Color.LightYellow
    End Sub

    Private Sub cboPR26State_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPR26State.Leave
        cboPR26State.BackColor = Color.White
    End Sub

    Private Sub cboPR26State_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPR26State.Enter
        cboPR26State.BackColor = Color.LightYellow
    End Sub

    Private Sub txtPR26Pin_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPR26Pin.Leave
        txtPR26Pin.BackColor = Color.White
    End Sub

    Private Sub txtPR26Pin_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPR26Pin.Enter
        txtPR26Pin.BackColor = Color.LightYellow
    End Sub

    Private Sub txt26STD_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt26STD.Enter
        txt26STD.BackColor = Color.LightYellow
    End Sub

    Private Sub txt26STD_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt26STD.Leave
        txt26STD.BackColor = Color.White
    End Sub

    Private Sub txt26PHONE_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt26PHONE.Enter
        txt26PHONE.BackColor = Color.LightYellow
    End Sub

    Private Sub txt26PHONE_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt26PHONE.Leave
        txt26PHONE.BackColor = Color.White
    End Sub

    Private Sub txt26Email_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt26Email.Enter
        txt26Email.BackColor = Color.LightYellow
    End Sub

    Private Sub txt26Email_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt26Email.Leave
        txt26Email.BackColor = Color.White
    End Sub

    Private Sub txt26PRPAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt26PRPAN.Leave
        txt26PRPAN.BackColor = Color.White
    End Sub

    Private Sub txt26PRPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt26PRPAN.Enter
        txt26PRPAN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtName27_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName27.Leave
        txtName27.BackColor = Color.White
    End Sub

    Private Sub txtName27_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName27.Enter
        txtName27.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDesg27_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesg27.Leave
        txtDesg27.BackColor = Color.White
    End Sub

    Private Sub txtDesg27_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesg27.Enter
        txtDesg27.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27Add1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add1.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27Add1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add1.BackColor = Color.White
    End Sub

    Private Sub txt27Add2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add2.BackColor = Color.White
    End Sub

    Private Sub txt27Add2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add2.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27Add3_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add3.BackColor = Color.White
    End Sub

    Private Sub txt27Add3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add3.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27Add4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add5.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27Add4_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add5.BackColor = Color.White
    End Sub

    Private Sub txt27Add5_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add4.BackColor = Color.White
    End Sub

    Private Sub txt27Add5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27Add4.BackColor = Color.LightYellow
    End Sub

    Private Sub cboPR27State_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPR27State.Leave
        cboPR27State.BackColor = Color.White
    End Sub

    Private Sub cboPR27State_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPR27State.Enter
        cboPR27State.BackColor = Color.LightYellow
    End Sub

    Private Sub txtPR27Pin_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPR27Pin.Leave
        txtPR27Pin.BackColor = Color.White
    End Sub

    Private Sub txtPR27Pin_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPR27Pin.Enter
        txtPR27Pin.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27STD_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27STD.Leave
        txt27STD.BackColor = Color.White
    End Sub

    Private Sub txt27STD_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27STD.Enter
        txt27STD.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27PHONE_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27PHONE.Leave
        txt27PHONE.BackColor = Color.White
    End Sub

    Private Sub txt27PHONE_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27PHONE.Enter
        txt27PHONE.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27Email_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27Email.Leave
        txt27Email.BackColor = Color.White
    End Sub

    Private Sub txt27Email_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27Email.Enter
        txt27Email.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27PRPAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27PRPAN.Leave
        txt27PRPAN.BackColor = Color.White
    End Sub

    Private Sub txt27PRPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27PRPAN.Enter
        txt27PRPAN.BackColor = Color.LightYellow
    End Sub

    Private Sub txtName27E_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName27E.Leave
        txtName27E.BackColor = Color.White
    End Sub

    Private Sub txtName27E_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName27E.Enter
        txtName27E.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDesg27EQ_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesg27EQ.Enter
        txtDesg27EQ.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDesg27EQ_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesg27EQ.Leave
        txtDesg27EQ.BackColor = Color.White
    End Sub

    Private Sub txt27EAdd1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd1.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27EAdd1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd1.BackColor = Color.White
    End Sub

    Private Sub txt27EAdd2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd2.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27EAdd2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd2.BackColor = Color.White
    End Sub

    Private Sub txt27EAdd3_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd3.BackColor = Color.White
    End Sub

    Private Sub txt27EAdd3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd3.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27EAdd4_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd5.BackColor = Color.White
    End Sub

    Private Sub txt27EAdd4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd5.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27EAdd5_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd4.BackColor = Color.White
    End Sub

    Private Sub txt27EAdd5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txt27EAdd4.BackColor = Color.LightYellow
    End Sub

    Private Sub cboPR27EState_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPR27EState.Enter
        cboPR27EState.BackColor = Color.LightYellow
    End Sub

    Private Sub cboPR27EState_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPR27EState.Leave
        cboPR27EState.BackColor = Color.White
    End Sub

    Private Sub txtPR27EPin_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPR27EPin.Leave
        txtPR27EPin.BackColor = Color.White
    End Sub

    Private Sub txtPR27EPin_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPR27EPin.Enter
        txtPR27EPin.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27ESTD_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27ESTD.Leave
        txt27ESTD.BackColor = Color.White
    End Sub

    Private Sub txt27ESTD_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27ESTD.Enter
        txt27ESTD.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27EPHONE_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27EPHONE.Leave
        txt27EPHONE.BackColor = Color.White
    End Sub

    Private Sub txt27EPHONE_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27EPHONE.Enter
        txt27EPHONE.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27EEmail_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27EEmail.Leave
        txt27EEmail.BackColor = Color.White
    End Sub

    Private Sub txt27EEmail_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27EEmail.Enter
        txt27EEmail.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27EPRPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27EPRPAN.Enter
        txt27EPRPAN.BackColor = Color.LightYellow
    End Sub

    Private Sub txt27EPRPAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt27EPRPAN.Leave
        txt27EPRPAN.BackColor = Color.White
    End Sub

    Private Sub txtDesg24_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesg24.Leave
        txtDesg24.BackColor = Color.White
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Label47.Text = Marqueestrings.left(Label3.Text)
        If (Label47.Left) + (Label47.Width) <= 0 Then
            Label47.Left = Me.Width
        End If
        Label47.Left = (Label47.Left) - 1
    End Sub
    Private Sub Panel6_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel6.MouseMove
        If drag Then
            Me.Top = Cursor.Position.Y - mousey
            Me.Left = Cursor.Position.X - mousex
        End If
    End Sub
    Private Sub Panel6_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel6.MouseDown
        drag = True
        mousex = Cursor.Position.X - Me.Left
        mousey = Cursor.Position.Y - Me.Top
    End Sub
    Private Sub Panel6_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel6.MouseUp
        drag = False
    End Sub
    Private Sub frmCoMst_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        drag = True
        mousex = Cursor.Position.X - Me.Left
        mousey = Cursor.Position.Y - Me.Top
    End Sub
    Private Sub frmCoMst_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        If drag Then
            Me.Top = Cursor.Position.Y - mousey
            Me.Left = Cursor.Position.X - mousex
        End If
    End Sub
    Private Sub frmCoMst_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
        drag = False
    End Sub

    Private Sub frmCoMst_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Timer1.Start()
    End Sub

    Private Sub Label47_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label47.MouseDown
        drag = True
        mousex = Cursor.Position.X - Me.Left
        mousey = Cursor.Position.Y - Me.Top
    End Sub

    Private Sub Label47_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label47.MouseUp
        drag = False
    End Sub

    Private Sub Label47_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label47.MouseMove
        If drag Then
            Me.Top = Cursor.Position.Y - mousey
            Me.Left = Cursor.Position.X - mousex
        End If
    End Sub

    Private Sub cmdImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport.Click
        'If Directory.Exists(Application.StartupPath.Substring(0, Application.StartupPath.Length - 4) & Val(Application.StartupPath.Substring(Application.StartupPath.Length - 4))) Then
        '    OpendgTDS.InitialDirectory = Application.StartupPath.Substring(0, Application.StartupPath.Length - 4) & (Val(Application.StartupPath.Substring(Application.StartupPath.Length - 4)) - 1)
        'Else
        '    OpendgTDS.InitialDirectory = Application.StartupPath
        'End If
        Dim str As String
        Dim pos As Integer
        Dim frm As New frmImport
        str = Application.StartupPath
        pos = InStr(str, "WizinTDS " & FY.Substring(FY.Length - 4))

        If Directory.Exists(str.Substring(0, pos - 1) & "WizinTDS " & (FY.Substring(FY.Length - 4)) - 1 & "\Database") Then
            OpendgTDS.InitialDirectory = str.Substring(0, pos - 1) & "WizinTDS " & (FY.Substring(FY.Length - 4)) - 1 & "\Database"
        Else
            'OpendgTDS.InitialDirectory = str.Substring(0, pos - 1) & "WizinTDS" & (FY.Substring(FY.Length - 4)) & "\Database"
            OpendgTDS.InitialDirectory = Application.StartupPath & "\Database"
        End If

        OpendgTDS.FileName = "WizinTDS.mdb"
        Dim i = OpendgTDS.ShowDialog()
        If i = 2 Then
            frm.Dispose()
            Exit Sub
        End If
        'OpendgTDS.ShowDialog()
        If Strings.Right(OpendgTDS.FileName, 12) = "WizinTDS.mdb" Then
            frm.Cpath = OpendgTDS.FileName
            frm.Ctitle = "WizinTDS.mdb"
            frm.ShowDialog()
            'oCoMst.FillCosInLvw lvwCo
            lvwCo.Items.Clear()
            connectlistview()
            lvwCo.Refresh()

            If lvwCo.Items.Count <= 0 Then
                EditMode()
            Else
                NormalMode()
            End If

        Else

        End If
        Exit Sub
CancelErr:
        If Err.Number = 32755 Then
        Else
            MsgBox(Err.Number & "-" & Err.Description, vbCritical, "Error!!  Call JAK Infosolutions")
        End If
    End Sub

    Private Sub cmdBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBackup.Click

        Try
            SaveDgTDS.Title = "Select Location for Backup"
            SaveDgTDS.FileName = "BkupWTDS" & Application.ProductName.Substring(Application.ProductName.Length - 2) & ".mdb"
            'SaveDgTDS.ShowDialog()
            Dim i = SaveDgTDS.ShowDialog()
            If i = 2 Then

                Exit Sub
            End If
            If (String.IsNullOrEmpty(SaveDgTDS.FileName)) Then
                MessageBox.Show("No filename selected for Backup, Process Aborted")
            Else
                If System.IO.File.Exists(SaveDgTDS.FileName) Then
                    If (MessageBox.Show(SaveDgTDS.FileName & "Already exists. Do you Want to Overwrite?", "Overwrite", MessageBoxButtons.YesNoCancel) = vbYes) Then
                        GoTo Backup
                    End If
                Else
                    GoTo Backup
                End If
            End If
            Exit Sub
Backup:

            Me.Cursor = Cursors.WaitCursor
            System.IO.File.Copy(Application.StartupPath & "\Database\WizinTDS.mdb", SaveDgTDS.FileName, True)
            'fs.CopyFile(Application.StartupPath & "\Database\WizinTDS.mdb", SaveDgTDS.FileName, True)
            Me.Cursor = Cursors.Default
            MessageBox.Show("BackUp Complete...")
            Exit Sub
        Catch ex As Exception
            If Err.Number = 32755 Then
                MessageBox.Show("Backup Aborted by User...", "BackUp Aborted", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("BackUp Failed!!! Contact JAK!!!" & ex.Message, "BackUp Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.Cursor = Cursors.Default
        End Try


    End Sub

    Private Sub cmdRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRestore.Click
        Try
            OpendgTDS.Title = "Select The File To Restore"
            OpendgTDS.FileName = "BkupWTDS" & Application.ProductName.Substring(Application.ProductName.Length - 2) & ".mdb"
            'OpendgTDS.ShowDialog()
            Dim i = OpendgTDS.ShowDialog()
            If i = 2 Then

                Exit Sub
            End If
            If (String.IsNullOrEmpty(SaveDgTDS.FileName)) Then
                MessageBox.Show("No filename selected for Restore, Process Aborted")
            Else
                If Not System.IO.File.Exists(SaveDgTDS.FileName) Then
                    MessageBox.Show("BackUp File DoesNot Exist in this location", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                If System.IO.File.Exists(Application.StartupPath & "\Database\WizinTDS.mdb") Then
                    If (MessageBox.Show(Application.StartupPath & "\WizinTDS.mdb Already Exist. Do you want to overwrite?", "File Already Exist", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes) Then
                        If MessageBox.Show("This restore will destroy your present database file" & vbCrLf & "Are you Absolutely sure of overwriting this file with backedup file", "Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning) = vbYes Then
                            GoTo Restore
                        End If
                    End If
                Else
                    GoTo Restore
                End If
            End If
            Exit Sub
Restore:
            Me.Cursor = Cursors.WaitCursor
            File.Copy(OpendgTDS.FileName, Application.StartupPath & "\WizinTDS.mdb", True)
            Me.Cursor = Cursors.Default
            MessageBox.Show("Restore Completed Successfully Exiting Wizin... Please Restart Software...", "Restore Successfull", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
            Exit Sub
        Catch ex As Exception
            If Err.Number = 32755 Then
                MessageBox.Show("Restore Aborted by User...", "Restore Aborted", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Restore Failed!!! Contact JAK!!!" & ex.Message, "Restore Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.Cursor = Cursors.Default
        End Try


    End Sub

    Private Sub frmCoMst_MarginChanged(sender As Object, e As EventArgs) Handles Me.MarginChanged

    End Sub

    Private Sub frmCoMst_MaximizedBoundsChanged(sender As Object, e As EventArgs) Handles Me.MaximizedBoundsChanged

    End Sub

    Private Sub cmdcopydetail_Click(sender As Object, e As EventArgs) Handles cmdcopydetail.Click
        CopyData()
    End Sub

    Private Sub CopyData()

        On Error Resume Next
        txt24Add1.Text = txtCoAdd1.Text
        txt24Add2.Text = txtCoAdd2.Text
        txt24Add3.Text = txtCoAdd3.Text
        txt24Add5.Text = txtCoAdd5.Text
        txt24Add4.Text = txtCoAdd4.Text
        cboPR24State.SelectedValue = cboCoState.SelectedValue
        txtPR24Pin.Text = txtCoPin.Text
        txt24Email.Text = txtCoEmail.Text
        txt24STD.Text = txtCoSTD.Text
        txt24PHONE.Text = txtCoPhone.Text
        chk24AddChg.Checked = IIf(chkAddChg.Checked, "True", "False")

    End Sub

    Private Sub CmdCopyname_Click(sender As Object, e As EventArgs) Handles CmdCopyname.Click
        txtName27.Text = txtName24.Text
        txtDesg27.Text = txtDesg24.Text
        txtName26.Text = txtName24.Text
        txtDesg26.Text = txtDesg24.Text
        txtName27E.Text = txtName24.Text
        txtDesg27EQ.Text = txtDesg24.Text
    End Sub

    Private Sub cmdCopyAdd_Click(sender As Object, e As EventArgs) Handles cmdCopyAdd.Click
        On Error Resume Next
        txt26add1.Text = txt24Add1.Text
        txt26add2.Text = txt24Add2.Text
        txt26add3.Text = txt24Add3.Text
        txt26add5.Text = txt24Add5.Text
        txt26add4.Text = txt24Add4.Text
        cboPR26State.SelectedValue = cboPR24State.SelectedValue
        txtPR26Pin.Text = txtPR24Pin.Text
        txt26Email.Text = txt24Email.Text
        txt26STD.Text = txt24STD.Text
        txt26PHONE.Text = txt24PHONE.Text
        chk26AddChg.Checked = IIf(chk24AddChg.Checked, "True", "False")
        txt26PRPAN.Text = txt24PRPAN.Text

        txt27Add1.Text = txt24Add1.Text
        txt27Add2.Text = txt24Add2.Text
        txt27Add3.Text = txt24Add3.Text
        txt27Add4.Text = txt24Add4.Text
        txt27Add5.Text = txt24Add5.Text
        cboPR27State.SelectedValue = cboPR24State.SelectedValue
        txtPR27Pin.Text = txtPR24Pin.Text
        txt27Email.Text = txt24Email.Text
        txt27STD.Text = txt24STD.Text
        txt27PHONE.Text = txt24PHONE.Text
        chk27AddChg.Checked = IIf(chk24AddChg.Checked, "True", "False")
        txt27PRPAN.Text = txt24PRPAN.Text

        txt27EAdd1.Text = txt24Add1.Text
        txt27EAdd2.Text = txt24Add2.Text
        txt27EAdd3.Text = txt24Add3.Text
        txt27EAdd5.Text = txt24Add5.Text
        txt27EAdd4.Text = txt24Add4.Text
        cboPR27EState.SelectedValue = cboPR24State.SelectedValue
        txtPR27EPin.Text = txtPR24Pin.Text
        txt27EEmail.Text = txt24Email.Text
        txt27ESTD.Text = txt24STD.Text
        txt27EPHONE.Text = txt24PHONE.Text
        chk27EAddChg.Checked = IIf(chk24AddChg.Checked, "True", "False")
        txt27EPRPAN.Text = txt24PRPAN.Text
    End Sub

    Private Sub cmdCoDel_Click(sender As Object, e As EventArgs) Handles cmdCoDel.Click

        If MsgBox("Are you sure you want to delete this company." _
    & vbCrLf & "All data for this company will be PERMANENTLY removed?" _
    , vbYesNo + vbQuestion + vbDefaultButton2, "DELETE COMPANY") = vbYes Then
            If MsgBox("Are you absoultely sure?" _
                           & vbCrLf & "DATA WILL BE PERMANENTLY REMOVED." _
                           , vbYesNo + vbCritical + vbDefaultButton2, "SURE?") = vbYes Then
                If Delete() = True Then
                    NormalMode()
                    lvwCo.Clear()
                    frmGovDet.Visible = False
                    setupListView()
                    connectlistview()

                    If lvwCo.Items.Count > 0 Then
                        lvwCo.Items(0).Selected = True
                        lvwCo.Select()
                    Else
                        Add()
                    End If
                End If
            End If

        End If
    End Sub

    Private Function Delete() As Boolean

        Dim sql As String
        Dim cmd As New OleDbCommand
        Dim transaction As OleDbTransaction

        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction

        Try
            sql = " Delete * from challan26Q where retnid in(select retnid from retnmst where coid=" & selectedcoid & ")"

            cmd.CommandText = sql
            cmd.ExecuteNonQuery()


            sql = " Delete * from challan24Q where retnid in(select retnid from retnmst where coid=" & selectedcoid & ")"
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            sql = " Delete * from deductee24Q where did in(select did from deductmst where coid=" & selectedcoid & ")"
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            sql = " Delete * from deductee26Q where did in(select did from deductmst where coid=" & selectedcoid & ")"
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            sql = "Delete * From deductMst Where CoID = " & selectedcoid
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            sql = "Delete * From BankMst Where CoID = " & selectedcoid
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            sql = "Delete * From retnMst Where CoID = " & selectedcoid
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            sql = "Delete * From CoMst Where CoID = " & selectedcoid
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            transaction.Commit()

            Delete = True

        Catch ex As Exception
            MessageBox.Show("Message:", ex.Message)

            transaction.Rollback()
            Delete = False
        End Try

        cmd.Dispose()
        transaction.Dispose()
    End Function
    Private Sub Add()
        If lvwCo.Items.Count < NoOfCo Then
            Mode = "A"

            tempadd = 1
            controlEnabled()
            Cleartext()
            cmdSave.Enabled = True
            cmdCoAdd.Enabled = False
            cmdBackup.Enabled = False
            cmdCoEdit.Enabled = False
            cmdRestore.Enabled = False
            cmdTDS.Enabled = False
            lvwCo.Enabled = False
            cmdExit.Text = "Cancel"
            'SelectedId = Nothing
            cmdAlternate.Enabled = False
            cboGovtDetails.Text = ""
            If txtCoName.Enabled = True Then
                txtCoName.Focus()
            End If

        Else
            Call MsgBox("Cannot Add New Company, you have rights to create" _
                        & vbCrLf & "only " & NoOfCo & " Companies, Please buy higher licence from " _
                        & vbCrLf & "JAK Infosolutions Pvt. Ltd., Nagpur. Phone: 0712-2250009, 2251515." _
                        , vbInformation)
        End If
        ShowLicInfo()
    End Sub
    Private Sub ShowLicInfo()
        ToolStripStatusLabel1.Text = "Licensed for " & NoOfCo & " Companies. Utilised : " & lvwCo.Items.Count
        'StatusBarTDS.Panels(1).Text = "Licensed for " & NoOfCo & " Companies. Utilised : " & lvwCo.Items.Count
    End Sub
    Private Sub NormalMode()
        On Error Resume Next    'some controls may not have enabled property
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            ctrl.Enabled = False
        Next

        With Me
            .lvwCo.Enabled = True
            .cmdCoAdd.Enabled = True
            .cmdCoEdit.Enabled = True
            .cmdCoDel.Enabled = True
            .cmdTDS.Enabled = False         'will change on selection of company..
            .cmdSave.Enabled = False
            .cmdExit.Text = "E&xit"
            .cmdExit.Enabled = True
            .cmdImport.Enabled = True
            .cmdBackup.Enabled = True
            .cmdRestore.Enabled = True
            .cmdGovtDetails.Enabled = False
            .cmdonline.Enabled = True
        End With
        'If lvwCo.Enabled = True Then lvwCo.SetFocus
        Mode = "X"
        ctrl = Nothing
    End Sub

    Private Sub EditMode()
        'On Error Resume Next    'some controls may not have enabled property
        'Dim ctrl As Control
        'For Each ctrl In Me.Controls
        '    ctrl.Enabled = True
        'Next
        'With Me
        '    .lvwCo.Enabled = False
        '    .cmdCoAdd.Enabled = False
        '    .cmdCoEdit.Enabled = False
        '    .cmdCoDel.Enabled = False
        '    .cmdTDS.Enabled = False         'will change on selection of company..
        '    .cmdSave.Enabled = True
        '    .cmdExit.Text = "&Cancel"
        '    .cmdExit.Enabled = True
        '    .cmdBackup.Enabled = False
        '    .cmdRestore.Enabled = False
        '    .cmdGovtDetails.Enabled = True
        'End With
        'If txtCoName.Enabled = True Then txtCoName.Focus()
        'ctrl = Nothing
    End Sub
    Private Sub txtCoName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCoName.KeyPress
        If Asc(e.KeyChar) = Asc("/") Or Asc(e.KeyChar) = Asc("\") Or Asc(e.KeyChar) = Asc("'") Or Asc(e.KeyChar) = Asc(Chr(34)) Then
            e.KeyChar = ""
        End If
    End Sub

    Private Sub TabPage2_Click_1(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub lvwCo_DoubleClick(sender As Object, e As EventArgs) Handles lvwCo.DoubleClick
        frmTDS.ShowDialog()
    End Sub

    Private Sub txtCoName_TextChanged(sender As Object, e As EventArgs) Handles txtCoName.TextChanged

    End Sub

    Private Sub txtCoBrDiv_TextChanged(sender As Object, e As EventArgs) Handles txtCoBrDiv.TextChanged

    End Sub

    Private Sub txtGSTIN_TextChanged(sender As Object, e As EventArgs) Handles txtGSTIN.TextChanged

    End Sub

    Private Sub txtGSTIN_Validating(sender As Object, e As CancelEventArgs) Handles txtGSTIN.Validating

        Dim i As Integer, HasNumbers As Boolean, ValidFormat As Boolean

        'Not blank, now check the length..
        If Len(Trim(txtGSTIN.Text)) <> 15 Then
            isvalidGSTNo = False      'Invalid Length
            MsgBox("Invalid Length of GSTIN, Please enter Valid GSTIN.")

            Exit Sub
        End If
        'length is ok..now check format..
        'Check for numbers
        For i = 1 To Len(txtGSTIN.Text)
            If IsNumeric(Strings.Mid(txtGSTIN.Text, i, 1)) = True Then
                HasNumbers = True
                Exit For
            End If
        Next
        If HasNumbers = False Then

        Else
            'there are numbers, it must be in AAAAA9999A format
            For i = 1 To Len(txtGSTIN.Text)
                Select Case i
                    Case 3, 4, 5, 6, 7, 12, 14  'Alphabets
                        If Asc(UCase(Strings.Mid(txtGSTIN.Text, i, 1))) >= Asc("A") And Asc(UCase(Strings.Mid(txtGSTIN.Text, i, 1))) <= Asc("Z") Then
                            ValidFormat = True
                        Else
                            ValidFormat = False
                            Exit For
                        End If
                    Case 1, 2, 8, 9, 10, 11, 13 'Numbers
                        If IsNumeric(Strings.Mid(txtGSTIN.Text, i, 1)) = True Then
                            ValidFormat = True

                        Else
                            ValidFormat = False
                            isvalidGSTNo = False
                            MsgBox("Please Enter Valid GSTIN")
                            Exit For
                        End If
                End Select
            Next i

        End If

        If ValidFormat = True Then
            'check the sixth char
            Select Case Strings.Mid(gstin, 6, 1)
                Case "P", "H", "C", "J", "F", "A", "T", "B", "L", "G"
                    ValidFormat = True

            End Select

        End If
        If ValidFormat = False Then
            MsgBox("GSTN Not in Proper Format.")
        End If
    End Sub

    Private Sub txtCoPin_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCoPin.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txtCoEmail_TextChanged(sender As Object, e As EventArgs) Handles txtCoEmail.TextChanged

    End Sub

    Private Sub txtCoEmail_Validating(sender As Object, e As CancelEventArgs) Handles txtCoEmail.Validating
        If Len(Trim(txtCoEmail.Text)) > 0 Then
            'check only when something is written
            If ValidEmail(txtCoEmail.Text) = False Then
                MsgBox("Invalid Email ID, please correct it", vbCritical)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub txtmobile_TextChanged(sender As Object, e As EventArgs) Handles txtmobile.TextChanged

    End Sub

    Private Sub txtmobile_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtmobile.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txtCoSTD_TextChanged(sender As Object, e As EventArgs) Handles txtCoSTD.TextChanged

    End Sub

    Private Sub txtCoSTD_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCoSTD.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txtCoPhone_TextChanged(sender As Object, e As EventArgs) Handles txtCoPhone.TextChanged

    End Sub

    Private Sub txtCoPhone_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCoPhone.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPR24Pin_TextChanged(sender As Object, e As EventArgs) Handles txtPR24Pin.TextChanged

    End Sub

    Private Sub txtPR24Pin_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR24Pin.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txt24STD_TextChanged(sender As Object, e As EventArgs) Handles txt24STD.TextChanged

    End Sub

    Private Sub txt24STD_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt24STD.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txt24PHONE_TextChanged(sender As Object, e As EventArgs) Handles txt24PHONE.TextChanged

    End Sub

    Private Sub txt26PHONE_TextChanged(sender As Object, e As EventArgs) Handles txt26PHONE.TextChanged

    End Sub

    Private Sub txt26PHONE_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt26PHONE.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txt24PHONE_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt24PHONE.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txtPR27Pin_TextChanged(sender As Object, e As EventArgs) Handles txtPR27Pin.TextChanged

    End Sub

    Private Sub txtPR27Pin_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR27Pin.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txt27STD_TextChanged(sender As Object, e As EventArgs) Handles txt27STD.TextChanged

    End Sub

    Private Sub txt27STD_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt27STD.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txt27PHONE_TextChanged(sender As Object, e As EventArgs) Handles txt27PHONE.TextChanged

    End Sub

    Private Sub txt27PHONE_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt27PHONE.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txtPR27EPin_TextChanged(sender As Object, e As EventArgs) Handles txtPR27EPin.TextChanged

    End Sub

    Private Sub txtPR27EPin_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR27EPin.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txtCoTAN_Validating(sender As Object, e As CancelEventArgs) Handles txtCoTAN.Validating
        'BeforeSave()
    End Sub

    Private Sub txt24Email_Validating(sender As Object, e As CancelEventArgs) Handles txt24Email.Validating
        If Len(Trim(txt24Email.Text)) > 0 Then
            'check only when something is written
            If ValidEmail(txt24Email.Text) = False Then
                MsgBox("Invalid Email ID, please correct it", vbCritical)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub txt26Email_TextChanged(sender As Object, e As EventArgs) Handles txt26Email.TextChanged

    End Sub

    Private Sub txt26Email_Validating(sender As Object, e As CancelEventArgs) Handles txt26Email.Validating
        If Len(Trim(txt26Email.Text)) > 0 Then
            'check only when something is written
            If ValidEmail(txt26Email.Text) = False Then
                MsgBox("Invalid Email ID, please correct it", vbCritical)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub txt27Email_TextChanged(sender As Object, e As EventArgs) Handles txt27Email.TextChanged

    End Sub

    Private Sub txt27Email_Validating(sender As Object, e As CancelEventArgs) Handles txt27Email.Validating
        If Len(Trim(txt27Email.Text)) > 0 Then
            'check only when something is written
            If ValidEmail(txt27Email.Text) = False Then
                MsgBox("Invalid Email ID, please correct it", vbCritical)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub txt27EEmail_TextChanged(sender As Object, e As EventArgs) Handles txt27EEmail.TextChanged

    End Sub

    Private Sub txt27EEmail_Validating(sender As Object, e As CancelEventArgs) Handles txt27EEmail.Validating
        If Len(Trim(txt27Email.Text)) > 0 Then
            'check only when something is written
            If ValidEmail(txt27Email.Text) = False Then
                MsgBox("Invalid Email ID, please correct it", vbCritical)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub cboCoState_KeyDown(sender As Object, e As KeyEventArgs) Handles cboCoState.KeyDown
        If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub cboGovtDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles cboGovtDetails.KeyDown
        'If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub cboPR24State_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPR24State.SelectedIndexChanged

    End Sub

    Private Sub cboPR24State_KeyDown(sender As Object, e As KeyEventArgs) Handles cboPR24State.KeyDown
        If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub cboPR26State_KeyDown(sender As Object, e As KeyEventArgs) Handles cboPR26State.KeyDown
        If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub cboPR27State_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPR27State.SelectedIndexChanged

    End Sub

    Private Sub cboPR27State_KeyDown(sender As Object, e As KeyEventArgs) Handles cboPR27State.KeyDown
        If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub cboPR27EState_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPR27EState.SelectedIndexChanged

    End Sub

    Private Sub cboPR27EState_KeyDown(sender As Object, e As KeyEventArgs) Handles cboPR27EState.KeyDown
        If e.KeyCode = Keys.Return Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub lvwCo_KeyDown(sender As Object, e As KeyEventArgs) Handles lvwCo.KeyDown
        If e.KeyCode = Keys.Return Then frmTDS.ShowDialog()
    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub frmCoMst_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        ControlPaint.DrawBorder(e.Graphics, Me.ClientRectangle, Color.Black, ButtonBorderStyle.Solid)
        Dim dpiX As Single = e.Graphics.DpiX
        Dim dpiY As Single = e.Graphics.DpiY
        Dim pfc As New PrivateFontCollection()
        If dpiX = 96 Then
            ' Me.Width = 927
            ' On Form1 shown, start applying font 
            'Dim CFontPath As String = Application.StartupPath
            'pfc.AddFontFile("C:\JAKINFO\WizinTDS 2018\WizinTDS2018\Resources\Fonts\Roboto.ttf")
            Dim allCtrl As New List(Of Control)
            For Each ctrl As Control In FindALLControlRecursive(allCtrl, Me)
                ' You need to define which control type to change it's font family; not recommendd to just change all controls' fonts, it will create a missy shape
                If TypeOf ctrl Is Label Then 'Or TypeOf ctrl Is Button Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is RadioButton Or TypeOf ctrl Is ProgressBar Or TypeOf ctrl Is GroupBox Then
                    ' Dim CurrentCtrlFontSize = ctrl.Font.Size ' get current object's font size before applying new font family
                    ctrl.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Bold)
                Else
                    ctrl.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
                End If
            Next
            allCtrl.Clear()
        Else

            ' On Form1 shown, start applying font 
            'Dim CFontPath As String = Application.StartupPath
            'pfc.AddFontFile("C:\JAKINFO\WizinTDS 2018\WizinTDS2018\Resources\Fonts\Roboto.ttf")
            Dim allCtrl As New List(Of Control)
            For Each ctrl As Control In FindALLControlRecursive(allCtrl, Me)
                ' You need to define which control type to change it's font family; not recommendd to just change all controls' fonts, it will create a missy shape
                If TypeOf ctrl Is Label Then 'Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is Button Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is RadioButton Or TypeOf ctrl Is ProgressBar Or TypeOf ctrl Is GroupBox Then
                    ' Dim CurrentCtrlFontSize = ctrl.Font.Size ' get current object's font size before applying new font family
                    ctrl.Font = New Font("Microsoft Sans Serif", 6, FontStyle.Bold)
                Else
                    ctrl.Font = New Font("Microsoft Sans Serif", 6, FontStyle.Regular)
                End If
            Next
            allCtrl.Clear()
        End If
    End Sub

    Private Sub Deductor_Paint(sender As Object, e As PaintEventArgs) Handles Deductor.Paint
        ControlPaint.DrawBorder(e.Graphics, Me.ClientRectangle, Color.Black, ButtonBorderStyle.Solid)
    End Sub

    Private Sub lvwCo_Paint(sender As Object, e As PaintEventArgs) Handles lvwCo.Paint
        ControlPaint.DrawBorder(e.Graphics, Me.ClientRectangle, Color.Black, ButtonBorderStyle.Solid)
    End Sub

    Private Sub StatusStrip1_Paint(sender As Object, e As PaintEventArgs) Handles StatusStrip1.Paint
        ControlPaint.DrawBorder(e.Graphics, Me.ClientRectangle, Color.Black, ButtonBorderStyle.Solid)
    End Sub

    Private Sub txtCoAdd1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub frmCoMst_Leave(sender As Object, e As EventArgs) Handles Me.Leave

    End Sub

    Private Sub txtCoPAN_TextChanged(sender As Object, e As EventArgs) Handles txtCoPAN.TextChanged

    End Sub

    Private Sub txt24PRPAN_Validating(sender As Object, e As CancelEventArgs) Handles txt24PRPAN.Validating
        Dim PANErr As String = IsValidPAN(txt24PRPAN.Text, True, False)
        If PANErr <> 0 Then

            MsgBox("Enter valid PAN of Person Responsible", vbCritical, "INVALID PAN")
            txt24PRPAN.Focus()
        End If
    End Sub

    Private Sub txt26PRPAN_TextChanged(sender As Object, e As EventArgs) Handles txt26PRPAN.TextChanged
        Dim PANErr As String = IsValidPAN(txt26PRPAN.Text, True, False)
    End Sub

    Private Sub txt27PRPAN_TextChanged(sender As Object, e As EventArgs) Handles txt27PRPAN.TextChanged
        Dim PANErr As String = IsValidPAN(txt27PRPAN.Text, True, False)
    End Sub

    Private Sub txt27EPRPAN_TextChanged(sender As Object, e As EventArgs) Handles txt27EPRPAN.TextChanged
        Dim PANErr As String = IsValidPAN(txt27PRPAN.Text, True, False)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtCoAdd1.TextChanged

    End Sub

    Private Sub txtCoAdd4_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub txtCoAdd2_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles txt24Add3.TextChanged

    End Sub

    Private Sub ToolStripStatusLabel1_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel1.Click

    End Sub

    Private Sub cboGovtDetails_Validating(sender As Object, e As CancelEventArgs) Handles cboGovtDetails.Validating
        Dim frm As New frmGovDetails
        If (cboGovtDetails.SelectedIndex > -1 And cboGovtDetails.SelectedIndex <= 7) Then
            If cboGovtDetails.SelectedIndex = 1 Or cboGovtDetails.SelectedIndex = 3 Or
                cboGovtDetails.SelectedIndex = 5 Or cboGovtDetails.SelectedIndex = 7 Then
                frm.cboGovtState.Enabled = True
            Else
                frm.cboGovtState.SelectedIndex = -1
                frm.cboGovtState.Enabled = False
            End If

            frm.ShowDialog()
        End If
    End Sub

    Private Sub cmdCoEdit_LostFocus(sender As Object, e As EventArgs) Handles cmdCoEdit.LostFocus
        txtCoName.Focus()
    End Sub

    Private Sub chkAddChg_LostFocus(sender As Object, e As EventArgs) Handles chkAddChg.LostFocus
        TabPage1.Focus()
    End Sub

    Private Sub chkAddChg_Leave(sender As Object, e As EventArgs) Handles chkAddChg.Leave
        MainTab.SelectedIndex = MainTab.SelectedIndex + 1
        If MainTab.SelectedIndex = 1 Then
            txtName24.Focus()
        End If
    End Sub

    Private Sub chkUseForm16_Leave(sender As Object, e As EventArgs) Handles chkUseForm16.Leave
        MainTab.SelectedIndex = MainTab.SelectedIndex + 1
        If MainTab.SelectedIndex = 1 Then
            txtName26.Focus()
        End If
    End Sub

    Private Sub chk26AddChg_Leave(sender As Object, e As EventArgs) Handles chk26AddChg.Leave
        MainTab.SelectedIndex = MainTab.SelectedIndex + 1
        If MainTab.SelectedIndex = 2 Then
            txtName27.Focus()
        End If
    End Sub

    Private Sub chk27AddChg_Leave(sender As Object, e As EventArgs) Handles chk27AddChg.Leave
        MainTab.SelectedIndex = MainTab.SelectedIndex + 1
        If MainTab.SelectedIndex = 4 Then
            txtName27E.Focus()
        End If
    End Sub

    Private Sub chk27EAddChg_Leave(sender As Object, e As EventArgs) Handles chk27EAddChg.Leave
        cmdSave.Focus()
    End Sub

    Private Sub cmdSave_Leave(sender As Object, e As EventArgs) Handles cmdSave.Leave
        lvwCo.Focus()
    End Sub

    Private Sub txtCoTAN_TextAlignChanged(sender As Object, e As EventArgs) Handles txtCoTAN.TextAlignChanged

    End Sub

    Private Sub txt27EPHONE_TextChanged(sender As Object, e As EventArgs) Handles txt27EPHONE.TextChanged

    End Sub

    Private Sub txt27EPHONE_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt27EPHONE.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txt27ESTD_TextChanged(sender As Object, e As EventArgs) Handles txt27ESTD.TextChanged

    End Sub

    Private Sub txt27ESTD_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt27ESTD.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txt26STD_TextChanged(sender As Object, e As EventArgs) Handles txt26STD.TextChanged

    End Sub

    Private Sub txt26STD_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt26STD.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub cboCoState_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboCoState.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboCoState.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub cboPR24State_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboPR24State.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboPR24State.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub cboPR26State_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboPR26State.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboPR26State.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub cboPR27EState_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboPR27EState.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboPR27EState.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub cboPR27State_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboPR27State.KeyPress
        If Char.IsLower(e.KeyChar) Then
            'Convert to uppercase, and put at the caret position in the TextBox.
            cboPR27State.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If
    End Sub

    Private Sub cboPR26State_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPR26State.SelectedIndexChanged

    End Sub

    Private Sub txt26PHONE_GotFocus(sender As Object, e As EventArgs) Handles txt26PHONE.GotFocus
        CtrlGotFocus(txt26PHONE)
    End Sub

    Private Sub txt26STD_GotFocus(sender As Object, e As EventArgs) Handles txt26STD.GotFocus
        CtrlGotFocus(txt26STD)
    End Sub

    Private Sub txtPR26Pin_GotFocus(sender As Object, e As EventArgs) Handles txtPR26Pin.GotFocus
        CtrlGotFocus(txtPR26Pin)
    End Sub

    Private Sub txt24STD_GotFocus(sender As Object, e As EventArgs) Handles txt24STD.GotFocus
        CtrlGotFocus(txt24STD)
    End Sub

    Private Sub txt24PHONE_GotFocus(sender As Object, e As EventArgs) Handles txt24PHONE.GotFocus
        CtrlGotFocus(txt24PHONE)
    End Sub

    Private Sub txtPR26Pin_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPR26Pin.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub

    Private Sub txtCoPin_GotFocus(sender As Object, e As EventArgs) Handles txtCoPin.GotFocus
        CtrlGotFocus(txtCoPin)
    End Sub

    Private Sub txt24Add1_GotFocus(sender As Object, e As EventArgs) Handles txt24Add1.GotFocus
        CtrlGotFocus(txt24Add1)
    End Sub

    Private Sub txt24Add2_GotFocus(sender As Object, e As EventArgs) Handles txt24Add2.GotFocus
        CtrlGotFocus(txt24Add2)
    End Sub

    Private Sub txt24Add3_GotFocus(sender As Object, e As EventArgs) Handles txt24Add3.GotFocus
        CtrlGotFocus(txt24Add3)
    End Sub

    Private Sub txt24Add4_GotFocus(sender As Object, e As EventArgs) Handles txt24Add4.GotFocus
        CtrlGotFocus(txt24Add4)
    End Sub

    Private Sub txt24Add5_GotFocus(sender As Object, e As EventArgs) Handles txt24Add5.GotFocus
        CtrlGotFocus(txt24Add5)
    End Sub

    Private Sub txt24Email_GotFocus(sender As Object, e As EventArgs) Handles txt24Email.GotFocus
        CtrlGotFocus(txt24Email)
    End Sub

    Private Sub txt24PRPAN_GotFocus(sender As Object, e As EventArgs) Handles txt24PRPAN.GotFocus
        CtrlGotFocus(txt24PRPAN)
    End Sub

    Private Sub txt26add1_GotFocus(sender As Object, e As EventArgs) Handles txt26add1.GotFocus
        CtrlGotFocus(txt26add1)
    End Sub

    Private Sub txt26add2_GotFocus(sender As Object, e As EventArgs) Handles txt26add2.GotFocus
        CtrlGotFocus(txt26add2)
    End Sub

    Private Sub txt26add3_GotFocus(sender As Object, e As EventArgs) Handles txt26add3.GotFocus
        CtrlGotFocus(txt26add3)
    End Sub

    Private Sub txt26add4_GotFocus(sender As Object, e As EventArgs) Handles txt26add4.GotFocus
        CtrlGotFocus(txt26add4)
    End Sub

    Private Sub txt26add5_GotFocus(sender As Object, e As EventArgs) Handles txt26add5.GotFocus
        CtrlGotFocus(txt26add5)
    End Sub

    Private Sub txt26Email_GotFocus(sender As Object, e As EventArgs) Handles txt26Email.GotFocus
        CtrlGotFocus(txt26Email)
    End Sub

    Private Sub txt26PRPAN_GotFocus(sender As Object, e As EventArgs) Handles txt26PRPAN.GotFocus
        CtrlGotFocus(txt26PRPAN)
    End Sub

    Private Sub txt27Add1_GotFocus(sender As Object, e As EventArgs) Handles txt27Add1.GotFocus
        CtrlGotFocus(txt27Add1)
    End Sub

    Private Sub txt27Add2_GotFocus(sender As Object, e As EventArgs) Handles txt27Add2.GotFocus
        CtrlGotFocus(txt27Add2)
    End Sub

    Private Sub txt27Add3_GotFocus(sender As Object, e As EventArgs) Handles txt27Add3.GotFocus
        CtrlGotFocus(txt27Add3)
    End Sub

    Private Sub txt27Add5_GotFocus(sender As Object, e As EventArgs) Handles txt27Add5.GotFocus
        CtrlGotFocus(txt27Add5)
    End Sub

    Private Sub txt27Add4_GotFocus(sender As Object, e As EventArgs) Handles txt27Add4.GotFocus
        CtrlGotFocus(txt27Add4)
    End Sub

    Private Sub txt27EAdd1_GotFocus(sender As Object, e As EventArgs) Handles txt27EAdd1.GotFocus
        CtrlGotFocus(txt27EAdd1)
    End Sub

    Private Sub txt27EAdd2_GotFocus(sender As Object, e As EventArgs) Handles txt27EAdd2.GotFocus
        CtrlGotFocus(txt27EAdd2)
    End Sub

    Private Sub txt27EAdd3_GotFocus(sender As Object, e As EventArgs) Handles txt27EAdd3.GotFocus
        CtrlGotFocus(txt27EAdd3)
    End Sub

    Private Sub txt27EAdd4_GotFocus(sender As Object, e As EventArgs) Handles txt27EAdd4.GotFocus
        CtrlGotFocus(txt27EAdd4)
    End Sub

    Private Sub txt27EAdd5_GotFocus(sender As Object, e As EventArgs) Handles txt27EAdd5.GotFocus
        CtrlGotFocus(txt27EAdd5)
    End Sub

    Private Sub txt27EEmail_GotFocus(sender As Object, e As EventArgs) Handles txt27EEmail.GotFocus
        CtrlGotFocus(txt27EEmail)
    End Sub

    Private Sub txt27Email_GotFocus(sender As Object, e As EventArgs) Handles txt27Email.GotFocus
        CtrlGotFocus(txt27Email)
    End Sub

    Private Sub txt27EPHONE_GotFocus(sender As Object, e As EventArgs) Handles txt27EPHONE.GotFocus
        CtrlGotFocus(txt27EPHONE)
    End Sub

    Private Sub txt27EPRPAN_GotFocus(sender As Object, e As EventArgs) Handles txt27EPRPAN.GotFocus
        CtrlGotFocus(txt27EPRPAN)
    End Sub

    Private Sub txt27ESTD_GotFocus(sender As Object, e As EventArgs) Handles txt27ESTD.GotFocus
        CtrlGotFocus(txt27ESTD)
    End Sub

    Private Sub txt27PHONE_GotFocus(sender As Object, e As EventArgs) Handles txt27PHONE.GotFocus
        CtrlGotFocus(txt27PHONE)
    End Sub

    Private Sub txt27PRPAN_GotFocus(sender As Object, e As EventArgs) Handles txt27PRPAN.GotFocus
        CtrlGotFocus(txt27PRPAN)
    End Sub

    Private Sub txt27STD_GotFocus(sender As Object, e As EventArgs) Handles txt27STD.GotFocus
        CtrlGotFocus(txt27STD)
    End Sub

    Private Sub txtCoAdd1_GotFocus(sender As Object, e As EventArgs) Handles txtCoAdd1.GotFocus
        CtrlGotFocus(txtCoAdd1)
    End Sub

    Private Sub txtCoAdd2_GotFocus(sender As Object, e As EventArgs) Handles txtCoAdd2.GotFocus
        CtrlGotFocus(txtCoAdd2)
    End Sub

    Private Sub txtCoAdd3_GotFocus(sender As Object, e As EventArgs) Handles txtCoAdd3.GotFocus
        CtrlGotFocus(txtCoAdd3)
    End Sub

    Private Sub txtCoAdd4_GotFocus(sender As Object, e As EventArgs) Handles txtCoAdd4.GotFocus
        CtrlGotFocus(txtCoAdd4)
    End Sub

    Private Sub txtCoAdd5_GotFocus(sender As Object, e As EventArgs) Handles txtCoAdd5.GotFocus
        CtrlGotFocus(txtCoAdd5)
    End Sub

    Private Sub txtCoBrDiv_GotFocus(sender As Object, e As EventArgs) Handles txtCoBrDiv.GotFocus
        CtrlGotFocus(txtCoBrDiv)
    End Sub

    Private Sub txtCoEmail_GotFocus(sender As Object, e As EventArgs) Handles txtCoEmail.GotFocus
        CtrlGotFocus(txtCoEmail)
    End Sub

    Private Sub txtCoName_GotFocus(sender As Object, e As EventArgs) Handles txtCoName.GotFocus
        CtrlGotFocus(txtCoName)
    End Sub

    Private Sub txtCoPAN_GotFocus(sender As Object, e As EventArgs) Handles txtCoPAN.GotFocus
        CtrlGotFocus(txtCoPAN)
    End Sub

    Private Sub txtCoPhone_GotFocus(sender As Object, e As EventArgs) Handles txtCoPhone.GotFocus
        CtrlGotFocus(txtCoPhone)
    End Sub

    Private Sub txtCoSTD_GotFocus(sender As Object, e As EventArgs) Handles txtCoSTD.GotFocus
        CtrlGotFocus(txtCoSTD)
    End Sub

    Private Sub txtCoTAN_GotFocus(sender As Object, e As EventArgs) Handles txtCoTAN.GotFocus
        CtrlGotFocus(txtCoTAN)
    End Sub

    Private Sub txtDesg24_GotFocus(sender As Object, e As EventArgs) Handles txtDesg24.GotFocus
        CtrlGotFocus(txtDesg24)
    End Sub

    Private Sub txtDesg26_GotFocus(sender As Object, e As EventArgs) Handles txtDesg26.GotFocus
        CtrlGotFocus(txtDesg26)
    End Sub

    Private Sub txtDesg27_GotFocus(sender As Object, e As EventArgs) Handles txtDesg27.GotFocus
        CtrlGotFocus(txtDesg27)
    End Sub

    Private Sub txtDesg27EQ_GotFocus(sender As Object, e As EventArgs) Handles txtDesg27EQ.GotFocus
        CtrlGotFocus(txtDesg27EQ)
    End Sub

    Private Sub txtGSTIN_GotFocus(sender As Object, e As EventArgs) Handles txtGSTIN.GotFocus
        CtrlGotFocus(txtGSTIN)
    End Sub

    Private Sub txtmobile_GotFocus(sender As Object, e As EventArgs) Handles txtmobile.GotFocus
        CtrlGotFocus(txtmobile)
    End Sub

    Private Sub txtName24_GotFocus(sender As Object, e As EventArgs) Handles txtName24.GotFocus
        CtrlGotFocus(txtName24)
    End Sub

    Private Sub txtName26_GotFocus(sender As Object, e As EventArgs) Handles txtName26.GotFocus
        CtrlGotFocus(txtName26)
    End Sub

    Private Sub txtName27_GotFocus(sender As Object, e As EventArgs) Handles txtName27.GotFocus
        CtrlGotFocus(txtName27)
    End Sub

    Private Sub txtName27E_GotFocus(sender As Object, e As EventArgs) Handles txtName27E.GotFocus
        CtrlGotFocus(txtName27E)
    End Sub

    Private Sub txtPR24Pin_GotFocus(sender As Object, e As EventArgs) Handles txtPR24Pin.GotFocus
        CtrlGotFocus(txtPR24Pin)
    End Sub

    Private Sub txtPR27EPin_GotFocus(sender As Object, e As EventArgs) Handles txtPR27EPin.GotFocus
        CtrlGotFocus(txtPR27EPin)
    End Sub

    Private Sub txtPR27Pin_GotFocus(sender As Object, e As EventArgs) Handles txtPR27Pin.GotFocus
        CtrlGotFocus(txtPR27Pin)
    End Sub

    Private Sub cboCoState_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCoState.SelectedIndexChanged

    End Sub

    Private Sub Panel6_Paint(sender As Object, e As PaintEventArgs) Handles Panel6.Paint

    End Sub

    Private Sub txt24PRPAN_TextChanged(sender As Object, e As EventArgs) Handles txt24PRPAN.TextChanged

    End Sub

    Private Sub txtCoPAN_Validated(sender As Object, e As EventArgs) Handles txtCoPAN.Validated

    End Sub

    Private Sub txt27PRPAN_Validating(sender As Object, e As CancelEventArgs) Handles txt27PRPAN.Validating
        Dim PANErr As String = IsValidPAN(txt27PRPAN.Text, True, False)
        If PANErr <> 0 Then

            MsgBox("Enter valid PAN of Person Responsible", vbCritical, "INVALID PAN")
            txt27PRPAN.Focus()
        End If
    End Sub



    Private Sub txt26PRPAN_Validating(sender As Object, e As CancelEventArgs) Handles txt26PRPAN.Validating
        Dim PANErr As String = IsValidPAN(txt26PRPAN.Text, True, False)
        If PANErr <> 0 Then
            MsgBox("Enter valid PAN of Person Responsible", vbCritical, "INVALID PAN")
            txt26PRPAN.Focus()
        End If
    End Sub

    Private Sub txt27EPRPAN_Validating(sender As Object, e As CancelEventArgs) Handles txt27EPRPAN.Validating
        Dim PANErr As String = IsValidPAN(txt27EPRPAN.Text, True, False)
        If PANErr <> 0 Then

            MsgBox("Enter valid PAN of Person Responsible", vbCritical, "INVALID PAN")
            txt27EPRPAN.Focus()
        End If
    End Sub

    Private Sub txt24Add1_LostFocus(sender As Object, e As EventArgs) Handles txt24Add1.LostFocus
        CtrlLostFocus(txt24Add1)
    End Sub

    Private Sub txt24Add2_LostFocus(sender As Object, e As EventArgs) Handles txt24Add2.LostFocus
        CtrlLostFocus(txt24Add2)
    End Sub

    Private Sub txt24Add3_LostFocus(sender As Object, e As EventArgs) Handles txt24Add3.LostFocus
        CtrlLostFocus(txt24Add3)
    End Sub

    Private Sub txt24Add4_LostFocus(sender As Object, e As EventArgs) Handles txt24Add4.LostFocus
        CtrlLostFocus(txt24Add4)
    End Sub

    Private Sub txt24Add5_LostFocus(sender As Object, e As EventArgs) Handles txt24Add5.LostFocus
        CtrlLostFocus(txt24Add5)
    End Sub

    Private Sub txt26add1_LostFocus(sender As Object, e As EventArgs) Handles txt26add1.LostFocus
        CtrlLostFocus(txt26add1)
    End Sub

    Private Sub txt26add2_LostFocus(sender As Object, e As EventArgs) Handles txt26add2.LostFocus
        CtrlLostFocus(txt26add2)
    End Sub

    Private Sub txt26add3_LostFocus(sender As Object, e As EventArgs) Handles txt26add3.LostFocus
        CtrlLostFocus(txt26add3)
    End Sub

    Private Sub txt26add4_LostFocus(sender As Object, e As EventArgs) Handles txt26add4.LostFocus
        CtrlLostFocus(txt26add4)
    End Sub

    Private Sub txt26add5_LostFocus(sender As Object, e As EventArgs) Handles txt26add5.LostFocus
        CtrlLostFocus(txt26add5)
    End Sub

    Private Sub txt27EAdd1_LostFocus(sender As Object, e As EventArgs) Handles txt27EAdd1.LostFocus
        CtrlLostFocus(txt27EAdd1)
    End Sub

    Private Sub txt27EAdd2_LostFocus(sender As Object, e As EventArgs) Handles txt27EAdd2.LostFocus
        CtrlLostFocus(txt27EAdd2)
    End Sub

    Private Sub txt27EAdd3_LostFocus(sender As Object, e As EventArgs) Handles txt27EAdd3.LostFocus
        CtrlLostFocus(txt27EAdd3)
    End Sub

    Private Sub txt27EAdd4_LostFocus(sender As Object, e As EventArgs) Handles txt27EAdd4.LostFocus
        CtrlLostFocus(txt27EAdd4)
    End Sub

    Private Sub txt27EAdd5_LostFocus(sender As Object, e As EventArgs) Handles txt27EAdd5.LostFocus
        CtrlLostFocus(txt27EAdd5)
    End Sub

    Private Sub txtCoAdd1_LostFocus(sender As Object, e As EventArgs) Handles txtCoAdd1.LostFocus
        CtrlLostFocus(txtCoAdd1)
    End Sub

    Private Sub txtCoAdd2_LostFocus(sender As Object, e As EventArgs) Handles txtCoAdd2.LostFocus
        CtrlLostFocus(txtCoAdd2)
    End Sub

    Private Sub txtCoAdd3_LostFocus(sender As Object, e As EventArgs) Handles txtCoAdd3.LostFocus
        CtrlLostFocus(txtCoAdd3)
    End Sub

    Private Sub txtCoAdd4_LostFocus(sender As Object, e As EventArgs) Handles txtCoAdd4.LostFocus
        CtrlLostFocus(txtCoAdd4)
    End Sub

    Private Sub txtCoAdd5_LostFocus(sender As Object, e As EventArgs) Handles txtCoAdd5.LostFocus
        CtrlLostFocus(txtCoAdd5)
    End Sub

    Private Sub txt27Add1_LostFocus(sender As Object, e As EventArgs) Handles txt27Add1.LostFocus
        CtrlLostFocus(txt27Add1)
    End Sub

    Private Sub txt27Add2_LostFocus(sender As Object, e As EventArgs) Handles txt27Add2.LostFocus
        CtrlLostFocus(txt27Add2)
    End Sub

    Private Sub txt27Add3_LostFocus(sender As Object, e As EventArgs) Handles txt27Add3.LostFocus
        CtrlLostFocus(txt27Add3)
    End Sub

    Private Sub txt27Add4_LostFocus(sender As Object, e As EventArgs) Handles txt27Add4.LostFocus
        CtrlLostFocus(txt27Add4)
    End Sub

    Private Sub txt27Add5_LostFocus(sender As Object, e As EventArgs) Handles txt27Add5.LostFocus
        CtrlLostFocus(txt27Add5)
    End Sub

    Private Sub txtName26_LostFocus(sender As Object, e As EventArgs) Handles txtName26.LostFocus
        CtrlLostFocus(txtName26)
    End Sub

    Private Sub txtGSTIN_LostFocus(sender As Object, e As EventArgs) Handles txtGSTIN.LostFocus
        CtrlLostFocus(txtGSTIN)
    End Sub
End Class