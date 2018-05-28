Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Security.Policy
Imports mshtml

'Imports WindowsApplication1.modMain

Public Class frmPanVer
    Dim oDed As ClsDeductMstObj
    Dim cmd As OleDbCommand
    Dim headadaptor, headadaptor1 As New OleDbDataAdapter
    Dim headcommand As New OleDbCommandBuilder
    Dim ds As New DataSet
    Dim dr, dr2, dr3 As OleDbDataReader
    Dim IsPANVeried As Boolean
    Dim fds As New DataSet
    Dim nm As String
    'Dim obj As modMain

    Private Sub frmPanVer_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub cboDeductee_Leave(ByVal sender As Object, ByVal e As EventArgs) Handles cboDeductee.Leave
        cboDeductee.BackColor = Color.White
    End Sub

    Private Sub cboDeductee_Enter(ByVal sender As Object, ByVal e As EventArgs) Handles cboDeductee.Enter
        cboDeductee.BackColor = Color.LightYellow
    End Sub

    Private Sub txtadress1_Leave(ByVal sender As Object, ByVal e As EventArgs) Handles txtDAdd1.Leave
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

    Private Sub txtAdress4_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd4.Leave
        txtDAdd4.BackColor = Color.White
    End Sub

    Private Sub txtAdress4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd4.Enter
        txtDAdd4.BackColor = Color.LightYellow
    End Sub

    Private Sub txtAdress5_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd5.Leave
        txtDAdd5.BackColor = Color.White
    End Sub

    Private Sub txtAdress5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDAdd5.Enter
        txtDAdd5.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbdeducState_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDState.Enter
        cboDState.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbdeducState_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDState.Leave
        cboDState.BackColor = Color.White
    End Sub

    Private Sub txtdeducPin_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPin.Leave
        txtDPin.BackColor = Color.White
    End Sub

    Private Sub txtdeducPin_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDPin.Enter
        txtDPin.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbCatofPAN_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbocat.Leave
        cbocat.BackColor = Color.White
    End Sub

    Private Sub cmbCatofPAN_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbocat.Enter
        cbocat.BackColor = Color.LightYellow
    End Sub

    Private Sub txtDedName2_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDName.Leave
        txtDName.BackColor = Color.White
    End Sub

    Private Sub txtDedName2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDName.Enter
        txtDName.BackColor = Color.LightYellow
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
    Private Sub frmPanVer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        IsPANVeried = False

        oDed = New ClsDeductMstObj
        'Fill Category list
        cbocat.Items.Add("VALID PAN")
        cbocat.Items.Add("PAN APPLIED")
        cbocat.Items.Add("PAN INVALID")
        cbocat.Items.Add("PAN NOT AVBL")
        cbocat.SelectedIndex = 0
        cboCategory.Items.Add("G - General/Other")
        cboCategory.Items.Add("W - Woman Assessee")
        cboCategory.Items.Add("S - Senior Citizen")
        cboCategory.Items.Add("O - Super Senior Citizen")
        cboCategory.SelectedIndex = 0
        fillcombo()
        FillDeductees()
    End Sub

    Public Sub fillcombo()
        'frmPanVer.cn.Open()
        Dim nds As New DataSet
        Dim QueSt As String = "Select StateID,StateName from StateMst"
        nds = FetchDataSet(QueSt)

        If nds.Tables(0).Rows.Count > 0 Then
            cboDState.DataSource = nds.Tables(0)
            cboDState.ValueMember = "StateID"
            cboDState.DisplayMember = "StateName"
        End If
        cboDState.SelectedIndex = -1
        nds.Dispose()

        'cboDState.Items.Clear()
        'fds = FetchDataSet("Select * from StateMst")
        'cboDState.DataSource = Nothing
        'cboDState.Items.Clear()
        'cboDState.DataSource = fds.Tables(0)

        'cboDState.ValueMember = "StateID"
        'cboDState.DisplayMember = "StateName"
        'cboDState.SelectedIndex = -1
        'fds.Dispose()
    End Sub

    Private Sub FillDeductees()
        'cboDeductee.Items.Add("[ SELECT DEDUCTEE ]")
        Dim fds As New DataSet
        'Dim Sql As String = "Select DId ,DName From DeductMst Where CoId = " & selectedcoid & " AND DPANCat=0 and PANVerified=false Order By DName"
        'cboDeductee.DataSource = Nothing
        'fds = FetchDataSet(Sql)
        'If fds.Tables(0).Rows.Count > 0 Then
        '    cboDeductee.DataSource = fds.Tables(0)
        '    'cboDeductee.ValueMember = "DId"
        '    cboDeductee.DisplayMember = "DName"
        'End If
        ''cboDeductee.ValueMember = "DId"
        'cboDeductee.SelectedIndex = -1
        'fds.Dispose()
        'cboDeductee.Items.Clear()
        fds = FetchDataSet("Select DId ,DName From DeductMst Where CoId = " & selectedcoid & " AND DPANCat=0 and PANVerified=false")
        cboDeductee.DataSource = fds.Tables(0)
        cboDeductee.DisplayMember = "DName"
        cboDeductee.ValueMember = "DId"
        cboDeductee.SelectedIndex = -1
        fds.Dispose()
        If cboDeductee.SelectedIndex < 0 Then
            txtDName.Clear()
            txtDPAN.Clear()
            txtDPin.Clear()
            txtDAdd1.Clear()
            txtDAdd2.Clear()
            txtDAdd3.Clear()
            txtDAdd4.Clear()
            txtDAdd5.Clear()
            cbocat.SelectedIndex = -1
            cboCategory.SelectedIndex = -1
            cboDState.SelectedIndex = -1
        End If

    End Sub

    Private Sub TableLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint
        'cboDeductee.BackColor = Color.Yellow
    End Sub
    'Dim WithEvents oDed As ClsDeductMstObj
    'Attribute oDed.VB_VarHelpID = -1
    'Dim MyOptBtn As HTMLOptionButtonElement
    Dim MyCombo As HtmlElement
    'Dim IsPANVeried As Boolean

    Private Sub cboDeductee_SelectedIndex1(sender As Object, e As EventArgs) Handles cboDeductee.SelectedIndexChanged
        If cboDeductee.Text <> "System.Data.DataRowView" And cboDeductee.SelectedIndex >= 0 Then
            ShowData(cboDeductee.Text)
            txtDName.Text = cboDeductee.Text
        End If

    End Sub

    Private Sub CleardeducteeCtrls()
        cboDeductee.Text = ""
        txtDAdd1.Text = ""
        txtDAdd2.Text = ""
        txtDPAN.Text = ""
        txtDAdd3.Text = ""
        txtDAdd4.Text = ""
        txtDAdd5.Text = ""
        txtDPin.Text = ""
        txtref.Text = ""
        cboDState.Text = ""
        cbocat.Text = ""
        'cmdSave.Text = "&Add"
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
            'txtDid.Text = .did.ToString()
            cboDState.SelectedValue = .DState
            cboDeductee.Text = nm
            cbocat.SelectedIndex = .Dcat
            txtDPin.Text = .DPin
            txtDPAN.Text = .DPan
            optCo.Checked = IIf(.DType = "C", True, False)
            optOther.Checked = IIf(.DType = "O", True, False)
            txtref.Text = .Dref
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
        End With
    End Sub

    Private Sub BigWindowSize()
        Me.Height = 11550
        Me.Width = 12000
    End Sub

    Private Sub SmallWindowSize()
        'Me.Height = 5190
        'Me.Width = 6690

    End Sub

    Private Sub cboDState_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDState.SelectedIndexChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub cboDeductee_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDeductee.SelectedIndexChanged
        cboDeductee.BackColor = Color.LightYellow
    End Sub

    Private Sub cmdShowPAN_Click(sender As Object, e As EventArgs) Handles cmdShowPAN.Click

        If cboDeductee.SelectedIndex < 0 Then
            MsgBox("No Deducteee selected!", 0 + 48, "Auto PAN Verification")
            cboDeductee.Focus()
            Exit Sub
        End If

        If txtDPin.Text = "" Then
            MsgBox("Pin code is mandetory!", 0 + 48, "Auto PAN Verification")
            Exit Sub
        End If
        Me.WindowState = FormWindowState.Maximized
        WebBrowser1.Size = New System.Drawing.Size(1200, 600)

        Form_Resize()
        'Call BigWindowSize()
        WebBrowser1.Navigate("https://onlineservices.tin.egov-nsdl.com/etaxnew/tdsnontds.jsp")
        'WebBrowser1.Navigate("https://onlineservices.tin.egov-nsdl.com/etaxnew/PopServlet")
    End Sub

    Private Sub cbocat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbocat.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Dispose()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dispose()
    End Sub

    Private Sub txtDAdd4_TextChanged(sender As Object, e As EventArgs) Handles txtDAdd4.TextChanged

    End Sub

    Private Sub cboCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCategory.SelectedIndexChanged

    End Sub

    Private Sub txtDAdd3_TextChanged(sender As Object, e As EventArgs) Handles txtDAdd3.TextChanged

    End Sub

    Private Sub cmdVerified_Click()
        ' sender As Object, e As EventArgs
        'On err GoTo ErrHandler
        Dim ds1 As New DataSet
        If cboDeductee.SelectedIndex < 0 Then
            Exit Sub
        End If

        If IsPANVeried Then
            '*************************
            'Updating Field implecitly
            '*************************
            ds1 = FetchDataSet("Update DeductMST SET PANVerified=True Where DID=" & cboDeductee.SelectedIndex & " and CoID=" & selectedcoid & " and PANVerified=False")
            'cn.Execute
            MsgBox("PAN Verified successfully!", 0 + 64, "PAN Verification")
        Else
            If MsgBox("PAN is not verified as per ITD database." & vbCrLf &
                "" & vbCrLf & "Do you wish to mark this deductee as verified?", 4 + 32, "Auto PAN Verification") = vbYes Then
                '*************************
                'Updating Field explecitly
                '*************************
                ds1 = FetchDataSet("Update DeductMST SET PANVerified=True Where DID=" & cboDeductee.SelectedItem(cboDeductee.SelectedIndex) & " and CoID=" & selectedcoid & " ")
                'Cnn.Execute
            Else
                IsPANVeried = False
            End If
        End If
        FillDeductees()
        SmallWindowSize()
        Exit Sub
    End Sub

    Private Sub TableLayoutPanel5_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Form_Resize()
        'center position on screen
        Me.Left = ((Screen.PrimaryScreen.Bounds.Width / 2) - (Me.Width / 2))
        Me.Top = ((Screen.PrimaryScreen.Bounds.Height / 2) - (Me.Height / 2))
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        'Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Object)
        'Dim MyLink As HtmlElement

        Dim HTML As System.Windows.Forms.HtmlDocument
        Dim OptElements As System.Windows.Forms.HtmlElementCollection
        Dim curElement, OptElement As System.Windows.Forms.HtmlElement
        Dim MyOptBtn As HTMLOptionButtonElement

        'Dim MyTAN, MyName As HTMLInputTextElement
        'Dim CtrlName As String
        'mshtml.HTMLDocument HTML = (mshtml.HTMLDocument)webBrowser1.Document
        'On err GoTo ErrHandler
        BigWindowSize()
        'oDed = oDed.Fetch(nm)
        If cboDeductee.SelectedIndex < -1 Then
            Exit Sub
        End If

        If (sender Is WebBrowser1) Then
            If e.Url.OriginalString = "https://onlineservices.tin.egov-nsdl.com/etaxnew/tdsnontds.jsp" Then
                HTML = WebBrowser1.Document
                curElement = HTML.getElementById("280")
                curElement.InvokeMember("CLICK")
            ElseIf StringS.Left(e.Url.OriginalString, 65) = "https://onlineservices.tin.egov-nsdl.com/etaxnew/PopServlet?rKey=" Then
                HTML = WebBrowser1.Document
                'Loop through all elements and fill the data
                For Each curElement In HTML.getElementsByTagName("input")
                    ' Debug.Print(curElement.Name)
                    If curElement.getAttribute("Name") = "PAN" Then 'curElement.Name = "PAN" Then
                        curElement.SetAttribute("value", txtDPAN.Text) ' curElement = oDed.DPan
                    ElseIf curElement.Name = "MajorHead_1" Then
                        curElement.SetAttribute("Checked", "yes")
                        If curElement.innerText = "0020" Then
                            MyOptBtn.checked = IIf(Mid(txtDPAN.Text, 4, 1) = "C", True, False)
                        ElseIf curElement.innerText = "0021" Then
                            MyOptBtn.checked = IIf(Mid(txtDPAN.Text, 4, 1) <> "C", True, False)
                        End If
                        MyOptBtn = Nothing
                    ElseIf curElement.Name = "MinorHead_1" Then
                        'If curElement.InnerText = "300" Then
                        'MyOptBtn = "curElement"
                        curElement.SetAttribute("Checked", "yes")
                        'End If
                    ElseIf curElement.Name = "Name" Then
                        curElement.InnerText = cboDeductee.Text
                    ElseIf curElement.Name = "Add_Line1" Then
                        curElement.InnerText = txtDAdd1.Text
                    ElseIf curElement.Name = "Add_Line2" Then
                        curElement.InnerText = txtDAdd2.Text
                    ElseIf curElement.Name = "Add_Line3" Then
                        curElement.InnerText = txtDAdd3.Text
                    ElseIf curElement.Name = "Add_Line4" Then
                        curElement.InnerText = txtDAdd4.Text
                    ElseIf curElement.Name = "Add_Line5" Then
                        curElement.innerText = "NAGPUR" 'oDed.DAdd5
                    ElseIf curElement.Name = "Add_PIN" Then
                        curElement.InnerText = txtDPin.Text
                        'ElseIf curElement.Name = "Add_EMAIL" Then
                        '    curElement.SetAttribute("value", vbNullString) 'curElement.Value = vbNullString
                        'ElseIf curElement.Name = "Add_MOBILE" Then
                        '    curElement.SetAttribute("value", vbNullString) 'curElement.Value = vbNullString
                    ElseIf curElement.Name = "Submit" Then
                        curElement.InvokeMember("click")
                    End If
                Next
                'select Combo items
                OptElements = HTML.GetElementsByTagName("Option")
                Dim MyAY, MyState As HtmlElement
                Dim FY As String
                Dim rst As New DataSet
                ' Dim MyStateName As String
                curElement = HTML.GetElementById("AssessYear_1")
                curElement = HTML.GetElementById("Add_State_1")
                rst = FetchDataSet("SELECT StateName from StateMst WHERE StateId=" & cboDState.SelectedValue)
                ' rst.Dispose()
                cboDState.DataSource = rst.Tables(0)
                cboDState.DisplayMember = "StateName"
                For Each OptElement In OptElements
                    ' Debug.Print(OptElement.Name)
                    Select Case OptElement.Parent.Name
                        Case "AssessYear_1"
                            'If OptElement.GetAttribute("Text") = (Left(AY, 4) & "-" & Right(AY, 2)) Then 'AY Change to FY BeCouse in Q4 it gives Error for self assesment Tax
                            If OptElement.GetAttribute("Text") = AY.Substring(0, 4) & "-" & AY.Substring(AY.Length - 2, 2) Then
                                '  OptElement.InnerHtml = True
                                OptElement.SetAttribute("selected", True)
                                ' OptElement.RaiseEvent("onchange")
                            End If

                        Case "FinancialYear"
                            If OptElement.GetAttribute("Text") = FY.Substring(0, 4) & "-" & FY.Substring(FY.Length - 2, 2) Then
                                ' OptElement.InnerHtml = True
                                OptElement.SetAttribute("selected", True)
                            End If
                        Case "Add_State_1"
                            If OptElement.GetAttribute("Text") = cboDState.Text Then
                                ' OptElement.InnerHtml = True
                                OptElement.SetAttribute("selected", True)
                            End If
                        Case "NaturePayment"
            '                            If OptElement.getAttribute("Text") = FrmChallan281.cmbSection.Text Then
            '                                OptElement.Selected = True
            '                            End If
                        Case "NetBankName_c"
                            If OptElement.GetAttribute("Text") = "Allahabad Bank" Then
                                ' OptElement.InnerHtml = True
                                OptElement.SetAttribute("selected", True)
                            End If
                    End Select
                Next
                '************************
                'Clicking Proceed Button
                '************************
                'For Each curElement In HTML.GetElementsByTagName("input")
                '    If curElement.GetAttribute("Text") = "Submit" Then
                '        curElement.InvokeMember("click") '  curElement.click()
                '    End If
                'Next
                'For Each curElement In HTML.GetElementsByTagName("input")
                '        'Debug.Print curElement.Name
                '        If curElement.GetAttribute("name") = "Submit" Then
                '            curElement.SetAttribute("selected", True)  '"NGPJ00254C"
                '        End If
                '    Next
            ElseIf Microsoft.VisualBasic.Strings.Left(e.Url.OriginalString, 60) = "https://onlineservices.tin.nsdl.com/etaxnew/SubmitTdsn?rKey=" Then
                    HTML = WebBrowser1.Document

                '************************
                'Verifying Deductee Name
                '************************
                For Each curElement In HTML.getElementsByTagName("input")
                    If curElement.Name = "Name" Then
                        If curElement.TagName = cboDeductee.Text Then
                            'Deductee if Verified
                            IsPANVeried = True
                            cmdVerified_Click()
                        End If
                    End If
                Next
            End If
        End If
        Me.WebBrowser1.Document.Body.Style = "font-size:x-small;"
        Exit Sub

ErrHandler:
        MessageBox.Show("Err.Description")

    End Sub

End Class