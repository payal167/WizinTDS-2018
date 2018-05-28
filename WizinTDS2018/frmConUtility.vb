Imports System.IO
Imports Microsoft.Win32
Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.ComponentModel


Public Class frmConUtility
    Public coname, cotan As String
    Dim tdsFilePath As String
    Public oCoMst As New clsCoMst
    'Dim fso As New Scripting.FileSystemObject, 
    Dim Qtr As String
    Public isError As Boolean
    Public PRNNO1 As String

    Private Sub frmConUtility_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        frmCoMst.EnterTab(e)
    End Sub

    Private Sub txtRecNO_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOldRRRNo.Leave
        txtOldRRRNo.BackColor = Color.White
    End Sub

    Private Sub txtRecNO_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOldRRRNo.Enter
        txtOldRRRNo.BackColor = Color.LightYellow
    End Sub

    Private Sub cmdOpen27A_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdprint.Click
        Dim Fname As String
        Fname = Application.StartupPath & "\e-TDS Files\" & oCoMst.CoName & "\27A_" & oCoMst.CoTAN & "_" & Microsoft.VisualBasic.Strings.Left(Qtr, Len(Qtr) - 1) & "_" & Microsoft.VisualBasic.Strings.Right(Qtr, 2) & "_" & Microsoft.VisualBasic.Strings.Left(FY, 4) & Microsoft.VisualBasic.Strings.Right(FY, 2) & ".pdf"
        ShellExecute(0&, Fname, "", "", "open", vbNormalFocus)
    End Sub

    Private Sub cmdShowStatFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowStatFile.Click
        Call OpenHtmlPage((Microsoft.VisualBasic.Strings.Left(cdgconvert.FileName, Len(cdgconvert.FileName) - 4) & "err.html"))
    End Sub

    Private Sub frmConUtility_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TopMost = True
        oCoMst = oCoMst.FetchCo(selectedcoid)
        Dim rsqt As New DataSet

        Qtr = frmTDS.ConvertWhich
        'oCoMst.FetchCo (selectedcoid)
        Select Case Strings.Left(frmTDS.ConvertWhich, 2)
            Case "24"
                Me.Text = "Converting Form 24Q (Quarter " & Strings.Right(frmTDS.ConvertWhich, 1) & ")"
                cmdprint.Text = "Print Form 27A"
            Case "26"
                Me.Text = "Converting Form 26Q (Quarter " & Strings.Right(frmTDS.ConvertWhich, 1) & ")"
                cmdprint.Text = "Print Form 27A"
            Case "27"
                If Mid(Qtr, 3, 1) = "E" Then
                    Me.Text = "Converting Form 27EQ (Quarter " & Strings.Right(frmTDS.ConvertWhich, 1) & ")"
                    cmdprint.Text = "Print Form 27A"
                Else
                    Me.Text = "Converting Form 27Q (Quarter " & Strings.Right(frmTDS.ConvertWhich, 1) & ")"
                    cmdprint.Text = "Print Form 27A"
                End If
        End Select

        cdgconvert.Title = "Save e-TDS File - Form No " & Qtr
        cdgconvert.Filter = "*.txt|*.txt"
        cdgconvert.DefaultExt = "txt"
        'Added on 14/05/05 for company specific folder creation..
        'oCoMst = selectedcoid
        tdsFilePath = Application.StartupPath & "\e-TDS Files\" & oCoMst.CoName
        If Not Directory.Exists(tdsFilePath) Then
            Directory.CreateDirectory(tdsFilePath)
        End If

        cdgconvert.InitialDirectory = tdsFilePath
        cdgconvert.FileName = tdsFilePath & "\F" & Qtr & "V" & ".txt" '& IIf(revised = "C", "_Revised", "") & ".txt"
        cdgconvert.DefaultExt = True

        'cdgConvert.ShowSave
        '   rsqt.Open "SELECT RetnMst.CoID, RetnMst.FrmType, RetnMst.PRN, strings.left([frmtype],Len([frmtype])-1) AS Form, Right([frmtype],1) AS qutr FROM RetnMst where coid= " & selectedcoid & " and ((strings.left([frmtype],Len([frmtype])-1))=" & Chr(39) & (strings.left([Qtr], Len([Qtr]) - 1)) & Chr(39) & ") AND ((Right([frmtype],1))<" & Right(Qtr, 1) & ")order by retnmst.frmtype DESC", Cnn, adOpenForwardOnly, adLockReadOnly
        '  Dim frm As New frmCOdetail

        ' txtOldRRRNo = frm.PRNNO
        'While Not rsqt.EOF
        'If Not IsNull(rsqt!PRN) Then
        '    txtOldRRRNo = rsqt!PRN
        '    lblQtrDisplay.Caption = "Receipt No.of" & " " & rsqt!FrmType
        '    'rsqt.Close
        '    'txtOldRRRNo.Enabled = False
        '    Exit Sub
        'End If
        ' rsqt.MoveNext
        'Wend

    End Sub

    Private Sub chkOldReceipt_CheckedChanged(sender As Object, e As EventArgs) Handles chkOldReceipt.CheckedChanged

    End Sub

    Private Sub chkOldReceipt_Click(sender As Object, e As EventArgs) Handles chkOldReceipt.Click
        If txtOldRRRNo.Text = "" Then
            If chkOldReceipt.Checked = 1 Then
                Label1.Enabled = True
                txtOldRRRNo.Enabled = True
            Else
                Label1.Enabled = False
                txtOldRRRNo.Enabled = False
                txtOldRRRNo.Text = ""
            End If
        Else
            chkOldReceipt.Checked = 1
            MessageBox.Show("Receipt No not blank so you can't change")
        End If
    End Sub

    Private Sub CMDCHECK_Click(sender As Object, e As EventArgs) Handles CMDCHECK.Click
        '  Dim frm As New frmCOdetail
        ' frm.Show
        'PRNNO1 = FRM.PRNNO
        '  txtOldRRRNo = frm.PRNNO
        'frmConUtility.Show vbModal, Me
        'Me.Show vbNormal
    End Sub

    '    Private Sub cmdConvert_Click(sender As Object, e As EventArgs) Handles cmdConvert.Click
    '        ' Dim fso As New Scripting.FileSystemObject
    '        Dim sqlac As String
    '        Dim Qtr As String
    '        ' Dim rs As New ADODB.Recordset
    '        Dim reply As Integer

    '        On Error GoTo canerr

    '        If txtOldRRRNo.Text = "" And chkOldReceipt.Checked = 1 Then
    '            MsgBox("Receipt No. can't blank please update Receipt No then convert Return", "", MessageBoxButtons.YesNo = DialogResult.Yes)
    '            Exit Sub
    '        End If

    '        'txtOldRRRNo = PRNNO
    '        lastrr = txtOldRRRNo.Text
    '        lastret = IIf(chkOldReceipt.Checked = True, "Y", "N")
    '        Qtr = frmTDS.ConvertWhich
    '        If chkOldReceipt.Checked = True Then
    '            Dim qtrdis As Integer
    '            Dim regqtr As Integer

    '            'regqtr = Right(Qtr, 1) - 1
    '            If lblQtrDisplay.Text <> vbNullString Then
    '                qtrdis = IIf(lblQtrDisplay.Text = "", 0, Strings.Right(lblQtrDisplay.Text, 1))
    '            Else
    '                qtrdis = lblQtrDisplay.Text

    '                regqtr = Strings.Right(Qtr, 1) - 1
    '                If (qtrdis <> regqtr) Then
    '                    reply = MsgBox("Regular statement filed for earlier period was not matched, Do you want to continue.....?", vbExclamation + vbYesNo, "Caution")
    '                    If reply = vbNo Then
    '                        Exit Sub
    '                    End If
    '                End If
    '            End If
    '        End If
    '        'Me.MousePointer = vbHourglass
    '        'oCoMst = oCoMst.FetchCo(selectedcoid)
    '        'ask whether return is revised...
    '        Dim revised As String
    '        ' revised = revised * 1
    '        Dim OldRRR As Double
    '        Dim TANApplNo As Double
    '        'Select Case MsgBox("Are you converting this return as revised return?", vbYesNo + vbExclamation + vbDefaultButton2, "REVISED?")
    '        'Case vbYes
    '        '   revised = "C"
    '        '   OldRRR = InputBox("Enter 14 Digit Original Receipt Number", "Original Receipt Number", 0)
    '        'Case vbNo
    '        '   revised = "R"
    '        '   OldRRR = 0
    '        'End Select
    '        revised = "R"
    '        OldRRR = 0
    '        TANApplNo = 0
    '        '   If oCoMst.CoTAN = "TANAPPLIED" Then
    '        '      TANApplNo = InputBox("Enter 14 Digit TAN Application Number", "TAN Application Number", 0)
    '        '   Else
    '        '      TANApplNo = 0
    '        '   End If

    '        '********Commented by Payal
    '        If BeforeConvert() = True Then Exit Sub
    '        IsAllPANVerified = True

    '        If cdgconvert.FileName = "" Then MsgBox("Cannot convert without filename") : GoTo cleanup
    '        Select Case Strings.Left(Qtr, 3)

    '            Case "24Q"
    '                Call Convert24Q(Qtr, cdgconvert.FileName, revised, OldRRR, TANApplNo)
    '            Case "26Q"
    '                Call Convert26Q(Qtr, cdgconvert.FileName, revised, OldRRR, TANApplNo)
    '            Case "27Q"
    '                Call Convert27Q(Qtr, cdgconvert.FileName, revised, OldRRR, TANApplNo)
    '            Case "27E"
    '                'Call Convert27EQ(Qtr, cdgconvert.FileName, revised, OldRRR, TANApplNo)
    '        End Select
    '        If CPan = True Then
    '            CPan = False
    '            Exit Sub
    '        End If
    '        '*********
    '        If IsAllPANVerified = False Then GoTo cleanup
    '        'IF PRN OF FORM ENTERED THEN FILE NOT CONVERTED DATE- 03/08/10
    '        ' sqlac = " SELECT RetnMst.RetnID, RetnMst.FrmType, RetnMst.DtOfFiling, RetnMst.PRN, CoMst.CoID FROM CoMst INNER JOIN RetnMst ON CoMst.CoID = RetnMst.CoID WHERE RetnMst.FrmType ='" & Qtr & "' AND CoMst.CoID=" & selectedcoid & ""
    '        ' rs.Open sqlac, Cnn, adOpenForwardOnly, adLockReadOnly
    '        Dim ds As New DataSet
    '        ds = FetchDataSet(" SELECT RetnMst.RetnID, RetnMst.FrmType, RetnMst.DtOfFiling, RetnMst.PRN, CoMst.CoID FROM CoMst INNER JOIN RetnMst ON CoMst.CoID = RetnMst.CoID WHERE RetnMst.FrmType ='" & Qtr & "' AND CoMst.CoID=" & selectedcoid & "")

    '        If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0)("prn").ToString()) Then
    '            MsgBox("Original return already submitted." _
    '        & vbCrLf & "Hence, Cannot convert file again.", vbOKOnly, "Duplicate File Converstion")
    '            Exit Sub
    '        End If
    '        If isError = True Then GoTo cleanup
    '        ' Me.MousePointer = vbDefault
    '        MsgBox("File Converted Sucessfully - Never take this file to NSDL centre." & vbCrLf &
    '               "Pass this file through NSDL-FVU for Validation and take the .FVU file to NSDL Centre." & vbCrLf & vbCrLf &
    '               "FileName:" & cdgconvert.FileName)

    '        'Check automatically the converted file after conversion
    '        If chkAutoLaunch.Checked = True Then
    '            cmdFVU_Click(sender, e)
    '        End If

    'cleanup:
    '        '  Me.MousePointer = vbDefault
    '        '    If rst.State = adStateOpen Then rst.Close
    '        'rst = Nothing
    '        'oRetnMst = Nothing
    '        'oCoMst = Nothing
    '        'ds.Dispose()
    '        Exit Sub
    'canerr:
    '        If Err.Number <> 32755 Then
    '            MsgBox(Err.Description, , Err.Number)
    '        End If
    '        GoTo cleanup
    '    End Sub

    Private Sub cmdcsi_Click(sender As Object, e As EventArgs) Handles cmdcsi.Click
        'Download CSI
        WebBrowser1.Visible = True
        'Me.Height = 10000
        'Me.Width = 15000
        'Me.Top = 1000
        'Me.strings.left = 1000
        WebBrowser1.Navigate("https://tin.tin.nsdl.com/oltas/index.html")
        'WebBrowser1.Navigate("https://tin.tin.nsdl.com/oltas/servlet/TanSearch")

    End Sub

    Private Sub cmdFVU_Click(sender As Object, e As EventArgs) Handles cmdFVU.Click
        On Error GoTo ErrHandler
        Dim WinSysDir As String, FVUHandle As Long
        Dim ParaStr As String
        'Dim sqlcsi As String = "select coname,cotan from company where coid =" & selectedcoid
        Dim csids As DataSet
        'Dim ocomst As New clsCoMst

        '  Dim fso As New FileSystemObject
        Dim CSIName As String
        'Exploring .csi file path as per version 4.0
        'csids = FetchDataSet(sqlcsi)
        ' coname = csids.Tables(0).Rows(0)(0).ToString()
        ' cotan = csids.Tables(0).Rows(0)(1).ToString()

        CSIName = Application.StartupPath & "\e-TDS files\" & oCoMst.CoName & "\" & oCoMst.CoTAN & Today().ToString("ddMMyy") & ".csi"
        cdgOpenCSI.FileName = ""
        If File.Exists(CSIName) = False Then
            Do While True
                'cdgOpenCSI.Filter = "*.csi"
                cdgOpenCSI.InitialDirectory = Application.StartupPath & "\e-TDS files"
                cdgOpenCSI.Title = "Select Challan status enquiry file"
                cdgOpenCSI.ShowDialog()
                If cdgOpenCSI.FileName = "" Then
                    If Not MsgBox("CSI file not selected, do you want to try again?", vbYesNo, "No File Selected") = vbYes Then
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
        Else
            cdgOpenCSI.FileName = CSIName
        End If
        If UCase(Strings.Left(Strings.Right(cdgOpenCSI.FileName, 20), 10)) <> oCoMst.CoTAN Then
            If MsgBox("TAN of selected file DOES NOT MATCH with TAN of deductor" & vbCrLf &
                "Do you want to continue with this file?", vbQuestion + vbYesNo + vbDefaultButton2, "TAN Mismatch") = vbNo Then
                cdgOpenCSI.FileName = ""
                Exit Sub

            End If
        End If

        WinSysDir = modTools.GetWinDir


        ParaStr = Chr(34) & cdgconvert.FileName & Chr(34) & Chr(32) & Chr(34) & Strings.Left(cdgconvert.FileName, Len(cdgconvert.FileName) - 4) & ".err" &
        Chr(34) & Chr(32) & Chr(34) & Strings.Left(cdgconvert.FileName, Len(cdgconvert.FileName) - 4) & ".fvu" & Chr(34) & Chr(32) & Chr(34) & "0" & Chr(34) &
        Chr(32) & Chr(34) & "5.7" & Chr(34) & Chr(32) & Chr(34) & "1" & Chr(34) & Chr(32) & Chr(34) &
        cdgOpenCSI.FileName & Chr(34)

        'Delete Error File if Exist
        Me.Cursor = Cursors.WaitCursor
        If File.Exists(Strings.Left(cdgconvert.FileName, Len(cdgconvert.FileName) - 4) & ".err") Then
            File.Delete(Strings.Left(cdgconvert.FileName, Len(cdgconvert.FileName) - 4) & ".err")
        End If

        If cdgOpenCSI.FileName = vbNullString Then
            'open fvu in interactive mode
            OpenFVUNew(Application.StartupPath & "\e-TDS Files\")
            'FVUHandle = OpenFVUNew(Application.StartupPath & "\e-TDS Files\")
            If MsgBox("Press Yes, when you see FVU running to auto fill data." & vbCrLf &
                "Press No for not filling data", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                RefillFVU()
            End If

        Else
            'open fvu in auto mode
            Call OpenFVUNew(Application.StartupPath & "\e-TDS Files\", ParaStr)
        End If

        'For i = 1 To 300000

        '    Application.DoEvents()
        'Next i

        If File.Exists(Strings.Left(cdgconvert.FileName, Strings.Len(cdgconvert.FileName) - 4) & ".err") Then
            If (MsgBox("Error file generated!" & vbCrLf &
                "" & vbCrLf & "Would you like to see error file?", 4 + 32, "Convert File Error!") = vbYes) Then
                Shell(0 & vbNullString, Strings.Left(cdgconvert.FileName, Strings.Len(cdgconvert.FileName) - 4) & ".err") ', vbNullString, vbNullString, vbNormalFocus
                '            OpenNotePad strings.left(cdgConvert.filename, Len(cdgConvert.filename) - 4) & ".ERR"
            End If
        End If
        Me.Cursor = Cursors.WaitCursor
        'Me.MousePointer = vbDefault
        'following 4 lines commented by nitin, we dont need them as we have automated the validation process
        '    If MsgBox("Press Yes, when you see FVU running to auto fill data." & vbCrLf & _
        '        "Press No for not filling data", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        '        RefillFVU
        '    End If

EX:     Exit Sub
ErrHandler:
        If Err.Number = 5 Then
            If MsgBox("FVU Still not loaded, do you want to retry?", vbYesNo) = vbYes Then
                Resume
            Else
                Exit Sub
            End If
        Else
            MsgBox(Err.Number & "-" & Err.Description)
            Exit Sub
        End If
    End Sub
    Private Sub RefillFVU()
        Interaction.AppActivate("TDS/TCS File Validation Utility") ', True
        '****************************************
        'Excel.Application.SendKeys cdgconvert.FileName, True
        '        Excel. "{TAB()}", True 
        '        Excel.Application.SendKeys vbTab, True
        '        Excel.Application.SendKeys vbTab, True
        '        Excel.Application.SendKeys vbTab, True
        '        Excel.Application.SendKeys tdsFilePath, True
        '        Excel.Application.SendKeys vbTab, True
        '        Excel.Application.SendKeys vbTab, True
        '        Excel.Application.SendKeys vbTab, True
        '        Excel.Application.SendKeys vbTab, True
        '        Excel.Application.SendKeys vbTab, True
        '        Excel.Application.SendKeys Chr(vbKeySpace), True
        CMDCHECK.Focus()
    End Sub

    Private Sub txtOldRRRNo_TextChanged(sender As Object, e As EventArgs) Handles txtOldRRRNo.TextChanged

    End Sub



    Private Sub cmdOpen27A_Click()
        'Dim fso As New FileSystemObject
        '    If fso.FileExists(App.Path & "\e-TDS Files\" & oCoMst.CoName & "\27A_" & oCoMst.CoTAN & "_" & strings.left(Qtr, Len(Qtr) - 1) & "_" & Right(Qtr, 2) & "_" & strings.left(FY, 4) & Right(FY, 2) & ".pdf") Then
        '        Shell App.Path & "\e-TDS Files\" & oCoMst.CoName & "\27A_" & oCoMst.CoTAN & "_" & strings.left(Qtr, Len(Qtr) - 1) & "_" & Right(Qtr, 2) & "_" & strings.left(FY, 4) & Right(FY, 2) & ".pdf", vbNormalFocus
        '    Else
        '        MsgBox "27A file does not exist!" & vbCrLf & _
        '            " " & vbCrLf & "First validate this Return.", 0 + 16
        '    End If
        Dim Fname As String
        Fname = Application.StartupPath & "\e-TDS Files\" & coname & "\27A_" & cotan & "_" & Strings.Left(Qtr, Len(Qtr) - 1) & "_" & Strings.Right(Qtr, 2) & "_" & Strings.Left(FY, 4) & Strings.Right(FY, 2) & ".pdf"

        Shell(0&, vbNullString, Fname, vbNullString)
    End Sub

    Private Sub cmdOpenTxtFile_Click()
        'If cdgConvert.filename = vbNullString
        '    With cdgConvert
        '        .DialogTitle = "Select Text File you want to view"
        '        .DefaultExt = "txt"
        '        .Filter = "Text Files (*.txt)|*.txt"
        '        .InitDir = App.Path & "\e-TDS Files"
        '        .ShowOpen
        '    End With
        'End If
        If cdgconvert.FileName = vbNullString Then Exit Sub
        Process.Start(cdgconvert.FileName(), "")
    End Sub

    '    Private Sub CMDPRINT_Click()
    '        frm27A.WhichQtr = Qtr
    '        frm27B.WhichQtr = Qtr
    '        Select Case strings.left(frmTDS.ConvertWhich, 3)
    '            Case "24Q"
    '                frm27A.WhichFrm = "24"
    '                Unload Me
    '        frm27A.Show vbModal
    '     Case "26Q"
    '                frm27A.WhichFrm = "26"
    '                Unload Me
    '         frm27A.Show vbModal
    '    Case "27Q"
    '                frm27A.WhichFrm = "27"
    '                Unload Me
    '         frm27A.Show vbModal
    '     Case "27E"
    '                frm27B.WhichFrm = "27"
    '                Unload Me
    '        frm27B.Show vbModal
    'End Select
    '    End Sub

    Private Sub cmdShowErrFile_Click()
        '   Dim fso As New FileSystemObject, ErrFileName As String
        '   cdgErrFile.ShowOpen
        '   cdgErrFile.Filter = "*.htm"
        '   'If fso.FileExists(ErrFileName) Then

        Call OpenHtmlPage((Strings.Left(cdgconvert.FileName, Len(cdgconvert.FileName) - 4) & "err.html"))
        '   End If

    End Sub

    Private Sub cmdShowStatFile_Click()
        Call OpenHtmlPage((Strings.Left(cdgconvert.FileName, Len(cdgconvert.FileName) - 4) & ".html"))
    End Sub

    Private Sub Command1_Click()
        'Dim FRRM As New frmConUtility
        'FRRM.txtOldRRRNo = Me.txtPRNNO


    End Sub

    Private Sub Form_Load()
        oCoMst = oCoMst.FetchCo(selectedcoid)

    End Sub

    Private Function BeforeConvert() As Boolean
        'Dim fso As New Scripting.FileSystemObject
        Me.Cursor = Cursors.WaitCursor
        If frmRegister.Mylock.RegisteredUser = False Then
            Call MsgBox("You are using an unregistered software so cannot proceed." _
        & vbCrLf & "For registration please contact JAK Infosolutions Pvt. Ltd., Nagpur." & vbCrLf & "Phone: 0712-2250009, 2251515." _
        , vbInformation, Application.ProductName)

            BeforeConvert = True
            GoTo cleanup
        End If
cleanup:
        Me.Cursor = Cursors.Default
        '    If rst.State = adStateOpen Then rst.Close
        'rst = Nothing
        'oRetnMst = Nothing
        'oCoMst = Nothing
        Exit Function
canerr:
        If Err.Number <> 32755 Then
            MsgBox(Err.Description, , Err.Number)
        End If
        GoTo cleanup
    End Function

    Private Sub txtOldRRRNo_Validate(Cancel As Boolean)
        If Len(Trim(txtOldRRRNo.Text)) > 0 Then
            If Len(txtOldRRRNo.Text) <> 15 Then
                MsgBox("Invalid receipt No, please correct it 15 Digit", vbCritical)
                Cancel = True
            End If
        End If
    End Sub

    Private Sub BindingNavigator1_RefreshItems(sender As Object, e As EventArgs) Handles BindingNavigator1.RefreshItems

    End Sub


    Private Sub optRevised_CheckedChanged(sender As Object, e As EventArgs) Handles optRevised.CheckedChanged

    End Sub

    Private Sub txtOldRRRNo_Validated(sender As Object, e As EventArgs) Handles txtOldRRRNo.Validated
        If Len(Trim(txtOldRRRNo.Text)) > 0 Then
            If Len(txtOldRRRNo.Text) <> 15 Then
                MsgBox("Invalid receipt No, please correct it 15 Digit", vbCritical)
                sender = True
            End If
        End If
    End Sub

    Private Sub optRevised_Click(sender As Object, e As EventArgs) Handles optRevised.Click

    End Sub
    Private Sub WebBrowser1_FileDownload(ByVal ActiveDocument As Boolean, Cancel As Boolean)
        WebBrowser1.Visible = False
        Me.Height = 7300
        Me.Width = 6500
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

        Dim MyLink As HtmlElement
        Dim HTML As HtmlDocument
        Dim MyElements As HtmlElementCollection
        Dim curElement As HtmlElement
        Dim OptElements As HtmlElementCollection
        Dim OptElement As HtmlElement
        Dim MyTAN As HtmlElement
        Dim MyName As HtmlElement
        Dim CtrlName As String
        Dim CURDT As String

        Dim cudaydt, cumntdt, cuyrdt As String

        ' On err GoTo ErrHandler
        If (sender Is WebBrowser1) Then
            If e.Url.OriginalString = "https://tin.tin.nsdl.com/oltas/index.html" Then  'OLTAS Challan Enquiry - Main Form
                HTML = WebBrowser1.Document
                'Loop through all elements and fill the data
                For Each curElement In HTML.GetElementsByTagName("input")
                    'Debug.Print curElement.Value
                    If curElement.GetAttribute("value") = "     TAN Based View      " Then
                        'If curElement.ToString = "     TAN Based View      " Then
                        curElement.InvokeMember("click")

                    End If
                    Debug.Print(curElement.GetAttribute("value"))
                Next
            End If
        End If

        If e.Url.OriginalString = "https://tin.tin.nsdl.com/oltas/servlet/TanSearch" Then   'WebAddress after 'TAN Based view is clicked.
            HTML = WebBrowser1.Document
            'Loop through all elements and fill the data
            For Each curElement In HTML.GetElementsByTagName("input")
                'Debug.Print curElement.Name
                If curElement.GetAttribute("id") = "TAN_NO" Then
                    curElement.SetAttribute("value", oCoMst.CoTAN)  '"NGPJ00254C"
                End If
            Next
            Dim MonthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month)

            CURDT = "30/06/" & Strings.Right(FY, 4) 'Format("30/06/" & Strings.Right(FY, 4), "dd/MM/yyyy")
            cudaydt = If(Len(Date.Now.Day()) = 1, "0" & Date.Now.Day, Date.Now.Day) ', "0" &  Date.Now.Day ,  Date.Now.Day
            cumntdt = MonthName 'IIf(Month(Today()) = "1", "JAN", IIf(Month(Today()) = "2", "FEB", IIf(Month(Of Date) = "3", "MAR", IIf(Month(Of Date) = "4", "APR", IIf(Month(Of Date) = "5", "MAY", IIf(Month(Of Date) = "6", "JUN", IIf(Month(Of Date) = "7", "JUL", IIf(Month(Of Date) = "8", "AUG", IIf(Month(Of Date) = "9", "SEP", IIf(Month(Of Date) = "10", "OCT", IIf(Month(Of Date) = "11", "NOV", "DEC")))))))))))
            cuyrdt = Year(Today())

            OptElements = HTML.GetElementsByTagName("Option")

            For Each OptElement In OptElements
                ' Debug.Print OptElement.parentElement.Name
                Select Case OptElement.Parent.Name
                    Case "TAN_FROM_DT_DD"
                        If OptElement.GetAttribute("Text") = "01" Then
                            OptElement.SetAttribute("selected", True)

                        End If
                    Case "TAN_FROM_DT_MM"
                        If OptElement.GetAttribute("Text") = "APR" Then
                            OptElement.SetAttribute("selected", "APR")
                        End If
                    Case "TAN_FROM_DT_YY"
                        If OptElement.GetAttribute("Text") = (Strings.Left(FY, 4)) Then
                            OptElement.SetAttribute("selected", (Strings.Left(FY, 4)))
                        End If
                    Case "TAN_TO_DT_DD"
                        If Today() > CURDT Then
                            If OptElement.GetAttribute("Text") = "30" Then
                                OptElement.SetAttribute("selected", "30")
                            End If
                        Else
                            If OptElement.GetAttribute("Text") = cudaydt Then
                                OptElement.SetAttribute("selected", cudaydt)
                            End If
                        End If
                    Case "TAN_TO_DT_MM"
                        If Today() > CURDT Then
                            If OptElement.GetAttribute("Text") = "JUN" Then
                                OptElement.SetAttribute("selected", "JUN")
                            End If
                        Else
                            If OptElement.GetAttribute("Text") = UCase(cumntdt) Then
                                OptElement.SetAttribute("selected", cumntdt)
                            End If
                        End If
                    Case "TAN_TO_DT_YY"
                        If Today() > CURDT Then
                            If OptElement.GetAttribute("Text") = Strings.Right(FY, 4) Then
                                OptElement.InvokeMember("click")
                                OptElement.SetAttribute("selected", Strings.Right(FY, 4))
                                OptElement.InvokeMember("click")
                                OptElement.SetAttribute("selected", Strings.Right(FY, 4))
                            End If
                        Else
                            If OptElement.GetAttribute("Text") = cuyrdt Then
                                OptElement.InvokeMember("click")
                                OptElement.SetAttribute("selected", cuyrdt)
                                OptElement.InvokeMember("click")
                                OptElement.SetAttribute("selected", cuyrdt)
                            End If
                        End If
                        curElement = HTML.GetElementById("DOWLOADFILE")
                        OptElement.InvokeMember("click")
                End Select
                '                    Me.Height = 7300
                '                    Me.Width = 6500
            Next
        End If
        Exit Sub

ErrHandler:
        MsgBox(Err.Description)
    End Sub

    Private Sub cmdOpenTxtFile_Click(sender As Object, e As EventArgs) Handles cmdOpenTxtFile.Click
        Dim FNM As String
        If cdgconvert.FileName = vbNullString Then

            With cdgconvert
                .Title = "Select Text File you want to view"
                .DefaultExt = "txt"
                .Filter = "Text Files (*.txt)|*.txt"
                .InitialDirectory = Application.StartupPath & "\e-TDS Files"
                .OpenFile()
            End With
        End If
        FNM = cdgconvert.FileName()
        If cdgconvert.FileName = vbNullString Then Exit Sub
        Process.Start(FNM)
    End Sub

    Private Sub WebBrowser1_FileDownload(sender As Object, e As EventArgs) Handles WebBrowser1.FileDownload
        'WebBrowser1.Visible = False
        'Me.Height = 7300
        'Me.Width = 6500
    End Sub

    Private Sub cmdConvert_Click(sender As Object, e As EventArgs) Handles cmdConvert.Click
        Dim oCoMst As New clsCoMst ', fso As New FileSystemObject
        Dim sqlac As String
        Dim Qtr As String

        Dim reply As Integer

        'On Error GoTo canerr

        If txtOldRRRNo.Text = "" And chkOldReceipt.Checked = True Then
            MsgBox("Receipt No. can't blank please update Receipt No then convert Return", vbOKOnly)
            Exit Sub
        End If


        lastrr = txtOldRRRNo.Text
        lastret = IIf(chkOldReceipt.Checked = True, "Y", "N")
        Qtr = frmTDS.ConvertWhich
        If chkOldReceipt.Checked = True Then
            Dim qtrdis As Integer
            Dim regqtr As Integer

            qtrdis = IIf(lblQtrDisplay.Text = "", 0, Strings.Right(lblQtrDisplay.Text, 1))

        End If
        Me.Cursor = Cursors.WaitCursor
        oCoMst = oCoMst.FetchCo(selectedcoid)
        'ask whether return is revised...
        Dim revised As String
        Dim OldRRR As Double
        Dim TANApplNo As Double

        revised = "R"
        OldRRR = 0
        TANApplNo = 0

        If BeforeConvert() = True Then Exit Sub


        If cdgconvert.FileName = "" Then MsgBox("Cannot convert without filename") : GoTo cleanup

        Select Case Strings.Left(Qtr, 3)

            Case "24Q"
                Call Convert24Q(Qtr, cdgconvert.FileName, revised, OldRRR, TANApplNo)
            Case "26Q"
                Call Convert26Q(Qtr, cdgconvert.FileName, revised, OldRRR, TANApplNo)
            Case "27Q"
                Call Convert27Q(Qtr, cdgconvert.FileName, revised, OldRRR, TANApplNo)
            Case "27E"
                Call Convert27EQ(Qtr, cdgconvert.FileName, revised, OldRRR, TANApplNo)
        End Select
        If CPan = True Then
            CPan = False
            Exit Sub
        End If
        If IsAllPANVerified = False Then GoTo cleanup
        'IF PRN OF FORM ENTERED THEN FILE NOT CONVERTED DATE- 03/08/10
        sqlac = " SELECT RetnMst.RetnID, RetnMst.FrmType, RetnMst.DtOfFiling, RetnMst.PRN, CoMst.CoID FROM CoMst INNER JOIN RetnMst ON CoMst.CoID = RetnMst.CoID WHERE RetnMst.FrmType ='" & Qtr & "' AND CoMst.CoID=" & selectedcoid & ""

        Dim ds As New DataSet
        ds = FetchDataSet(sqlac)
        If ds.Tables(0).Rows(0)("PRN").ToString() <> "" Then
            MsgBox("Original return already submitted." _
        & vbCrLf & "Hence, Cannot convert file again.", vbOKOnly, "Duplicate File Converstion")
            Exit Sub
        End If
        If isError = True Then GoTo cleanup
        Me.Cursor = Cursors.Default
        MsgBox("File Converted Sucessfully - Never take this file to NSDL centre." & vbCrLf &
          "Pass this file through NSDL-FVU for Validation and take the .FVU file to NSDL Centre." & vbCrLf & vbCrLf &
          "FileName:" & cdgconvert.FileName)

        'Check automatically the converted file after conversion
        If chkAutoLaunch.Checked = True Then
            cmdFVU_Click(sender, e)
        End If

cleanup:
        Me.Cursor = Cursors.Default



        'oCoMst = Nothing
        Exit Sub
canerr:
        If Err.Number <> 32755 Then
            MsgBox(Err.Description, , Err.Number)
        End If
        GoTo cleanup
EX: End Sub

    Private Sub cmdShowErrFile_Click(sender As Object, e As EventArgs) Handles cmdShowErrFile.Click
        Dim a As String
        'a = cdgconvert.FileName
        ' Call OpenHtmlPage((Strings.Left(cdgconvert.FileName, Len(cdgconvert.FileName) - 4) & ".err"))
        Call OpenHtmlPage((Microsoft.VisualBasic.Strings.Left(cdgconvert.FileName, Len(cdgconvert.FileName) - 4) & "err.html"))
    End Sub

    Private Sub frmConUtility_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Me.Dispose()
    End Sub
End Class