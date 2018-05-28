Imports System.IO
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Threading
Imports System.Globalization
Imports System.Drawing.Text

Module modMain
    Public selectedcoid As Long
    Public PrintRep As Boolean   'used for same pan/refno of deductee's list of Annexure
    Public rspan As New System.Data.DataSet  'used for same pan/refno of deductee's list of Annexure
    Public cn As New OleDbConnection
    Public NoOfCo As Integer 'Long
    Public SoftName As String
    Public DeleteAllowed As Boolean
    Public Frm27Qrt As Integer
    Public AllowCertificate As Boolean
    Public txtfind As String 'used for search deductee name add by jayhsree
    Public findsql1 As String 'used for search deductee name add by jayhsree
    Public findsql2 As String 'used for search deductee name add by jayhsree
    Public IsAllPANVerified As Boolean
    Public cboGovDetIndex As Integer
    Public RetnId As Long, F16ID As Long
    Private Resolution As New ResolutionChanger
    Private OldWidth As UInteger
    Private OldHeight As UInteger

    'API Declarations to be used with VBAccelerator Controls...
    'Private Declare Sub InitCommonControls Lib "comctl32" ()
    Public Sub Main()
        'InitCommonControls()          'Intialise the common controls..
        'ChangeDate("dd/MM/yyyy")
        Dim ODate = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\International", "sShortDate", "dd/MM/yyyy")
        If ODate <> "dd/MM/yyyy" Then
            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)
            Microsoft.Win32.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\International", "sShortDate", "dd/MM/yyyy")
            'MsgBox("Date of format is wrong..Please Restart software..!!")
            End
        End If
        OldHeight = CUInt(Screen.PrimaryScreen.Bounds.Height)
        OldWidth = CUInt(Screen.PrimaryScreen.Bounds.Width)
        Select Case Resolution.SetResolution(1280, 768)
            Case ResolutionChanger.ChangeResult.Success
                'MsgBox("The Resolution was changed", MsgBoxStyle.OkOnly)
            Case ResolutionChanger.ChangeResult.Restart
                MsgBox("Restart your system to activate the new resolution setting", MsgBoxStyle.OkOnly)
            Case ResolutionChanger.ChangeResult.Fail
                MsgBox("The resolution couldn't be changed", MsgBoxStyle.OkOnly)
            Case ResolutionChanger.ChangeResult.ResolutionNotSupported
                MsgBox("The requested resolution is not supported by your system", MsgBoxStyle.OkOnly)
        End Select
        Dim IsPathOk As Boolean
        Call SetTDSRates()
        Dim startpath As String
        ''check installation path..set in modTDSRate module.
        ''** the following code was commented for running the software in any path by ritesh
        ''MsgBox App.Path & vbCrLf & InstallPath
        'startpath = Application.StartupPath
        'startpath = Strings.Right(startpath, 9)
        'If startpath = "bin\Debug" Then
        '    startpath = Strings.Left(Application.StartupPath, 24)
        'Else
        '    startpath = Application.StartupPath
        'End If


        'If UCase(startpath) <> UCase(InstallPath) Then
        '    'Check for lan path..
        '    If Left(Application.StartupPath, 2) = "\\" Or Left(Application.StartupPath, 2) = "//" Then
        '        If UCase(Left(Application.StartupPath, 23)) <> UCase(Replace(InstallPath, ":", "")) Then
        '            IsPathOk = False
        '        Else
        '            IsPathOk = True
        '        End If
        '    Else
        '        IsPathOk = False
        '    End If
        '    If IsPathOk = False Then
        '        Call MsgBox("You have voilated certain permissions.  Please contact" _
        '       & vbCrLf & "JAK Infosolutions Pvt. Ltd. for taking the appropriate" _
        '       & vbCrLf & "license.  Phone: 0712-2250009, 5608288." _
        '       , vbCritical + vbDefaultButton1, Application.ProductName)
        '        End 'End program
        '    End If
        'End If
        ''** here the commented code ends
        'only one instance of the program should run on one machine..
        If PrevInstance() Then
            MsgBox("This Program is already running", vbCritical, "Double Instance of " & Application.ProductName)
            End
        End If
        'set the connection to the database,
        'Single user connection string, uses mdb exclusively.
        'Cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\easyTDS.mdb;Mode=Share Deny Read|Share Deny Write;Persist Security Info=False"
        'Multi user Connection string, uses mdb in shared mode.
        If File.Exists(Application.StartupPath & "\Database\WizinTDS.mdb") = False Then
            System.IO.File.Copy(Application.StartupPath & "\Database\WTDS.mdb", Application.StartupPath & "\Database\WizinTDS.mdb", False)
            'fs.CopyFile(Application.StartupPath & "\Database\WTDS.mdb", Application.StartupPath & "\Database\WizinTDS.mdb", False)
        End If
        Try
            cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Database\WizinTDS.mdb;Persist Security Info=False;Jet OLEDB:Database Password='apr01'"
            'Cnn.CursorLocation = CursorLocationEnum.adUseClient
            'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Gateway\supriya1\WizinTDS2018\WizinTDS2018\bin\Debug\Database\WizinTDS.mdb;Persist Security Info=False;Jet OLEDB:Database Password='apr01'"
            cn.Open()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        SoftName = "Wizin-TDS" & GetSetting("Wizin-TDS", Application.ProductName, "CDKey")
        'frmLogin.Show()
    End Sub
    Function PrevInstance() As Boolean
        If UBound(Diagnostics.Process.GetProcessesByName _
           (Diagnostics.Process.GetCurrentProcess.ProcessName)) _
           > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub CentreScreen(ByVal frm As Form)
        frm.Top = ((My.Computer.Screen.Bounds.Size.Height) / 2 - frm.Height / 2)
        frm.Left = ((My.Computer.Screen.Bounds.Size.Width / 2) - frm.Width / 2)
    End Sub

    'Public Sub ClearCtrls(ByVal frm As Form)
    '    Dim ctrl As Control
    '    For Each ctrl In frm.Controls
    '        If TypeOf ctrl Is TextBox Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is MaskedTextBox Or _
    '           TypeOf ctrl Is MyNumberBox Or TypeOf ctrl Is ComboBox Or TypeOf ctrl Is ComboBox Then
    '            ctrl.Text = vbNullString
    '        End If
    '        If TypeOf ctrl Is DataGrid Then
    '            ctrl.Text = ""
    '        End If
    '        If TypeOf ctrl Is CheckBox Then
    '            ctrl = CheckState.Unchecked
    '        End If
    '        If TypeOf ctrl Is RadioButton Then
    '            ctrl.Value = False
    '        End If
    '        If TypeOf ctrl Is ComboBox Then
    '            ctrl.SelectedIndex = -1
    '        End If

    '    Next ctrl

    'End Sub
    'Public Function EnableCtrls(ByVal frm As Form) As Object
    '    EnableCtrls = 0
    '    '#Const Compile_EnableCtrls = True
    '    If EnableCtrls Then
    '        EnableDisableCtrls(frm, True)
    '    End If ' Compile_EnableCtrls
    'End Function
    ''Public Function EnableCtrls(ByVal frm As Form)
    ''    Call EnableDisableCtrls(frm, True)
    ''End Function
    'Public Function DisableCtrls(ByVal frm As Form) As Object
    '    DisableCtrls = 0
    '    '#Const Compile_DisableCtrls = True
    '    If DisableCtrls Then
    '        EnableDisableCtrls(frm, False)
    '    End If ' Compile_DisableCtrls
    'End Function
    ''Public Function DisableCtrls(ByVal frm As Form)
    ''    Call EnableDisableCtrls(frm, False)
    ''End Function

    'Private Function EnableDisableCtrls(ByVal frm As Form, ByVal Enable As Boolean)
    '    Dim ctrl As Control
    '    For Each ctrl In frm.Controls
    '        If TypeOf ctrl Is TextBox Or _
    '              TypeOf ctrl Is DataGrid Or _
    '              TypeOf ctrl Is ComboBox Or _
    '              TypeOf ctrl Is RadioButton Or TypeOf ctrl Is CheckBox Or _
    '              TypeOf ctrl Is DateTimePicker Then
    '            ctrl.Enabled = Enable
    '        End If
    '        'if there is a treeview, disable it..if other are enabled and vice-versa
    '        If TypeOf ctrl Is TreeView Then ctrl.Enabled = Not Enable
    '    Next ctrl
    'End Function

    '' code to convert number into words...
    'Private Function IvalidNo(ByVal tmpstring As String) As Boolean
    '    Dim tmpbool As Boolean
    '    tmpbool = False
    '    If Len(Trim(tmpstring)) > 0 Then tmpbool = True
    '    IvalidNo = tmpbool
    'End Function

    'Public Function NumToWords(ByVal tmpVal As Decimal) As String
    '    Dim tmpamt As Decimal, Tmpamtstr As String, tmplen As Integer
    '    Dim tmparay() As Integer, lclcounter As Integer, tmpword As String
    '    Dim Paisestr As String, TMPTREND As String, Tmptrends As Object
    '    Tmptrends = {" ", " ", " Hundred ", " Thousand ", " ", " Lakh ", " ", " Crore", " ", " Billion ", " ", " Trillion "}
    '    tmpamt = Int(tmpVal)
    '    Tmpamtstr = Trim(Str(tmpamt))
    '    tmplen = Len(Tmpamtstr)
    '    Paisestr = IIf(InStr(Str(tmpVal), ".") = 0, "00", Mid(Str(tmpVal), InStr(Str(tmpVal), ".") + 1, 2))
    '    ReDim tmparay(tmplen - 1)
    '    For lclcounter = 0 To tmplen - 1
    '        tmparay(lclcounter) = Mid(Tmpamtstr, tmplen - lclcounter, 1)
    '    Next
    '    TMPTREND = ""
    '    For lclcounter = UBound(tmparay) To 0 Step -1
    '        Select Case lclcounter
    '            Case 0, 3, 5, 7, 9, 11, 13
    '                If (lclcounter = 0 Or lclcounter = 3 Or lclcounter = 5 Or lclcounter = 7 Or lclcounter = 9 Or lclcounter = 13 Or lclcounter = 11) And lclcounter + 1 <= UBound(tmparay) Then
    '                    TMPTREND = ReturnTenAmount(tmparay(lclcounter + 1), tmparay(lclcounter))
    '                Else
    '                    TMPTREND = ReturnOneAmount(tmparay(lclcounter))
    '                End If
    '            Case 2
    '                TMPTREND = ReturnOneAmount(tmparay(lclcounter))
    '        End Select
    '        If lclcounter <= UBound(Tmptrends) Then TMPTREND = TMPTREND & IIf(IvalidNo(TMPTREND) = True, Tmptrends(lclcounter), " ")
    '        tmpword = tmpword & TMPTREND
    '        TMPTREND = ""
    '    Next
    '    NumToWords = tmpword & IIf(Val(Paisestr) > 0, " And ", " ") & Get_Paise(Val(Mid(Paisestr, 1, 1)), Val(Mid(Paisestr, 2, 1)))
    '    Do While InStr(1, NumToWords, "  ") <> 0
    '        NumToWords = Replace(Trim(NumToWords), Space(2), Space(1))
    '    Loop
    'End Function

    'Private Function ReturnTenAmount(ByVal Tmpten As Integer, ByVal tmpones As Integer) As String
    '    Dim Tenarray As Object, Twnarray As Object, tmpVal As String
    '    Tenarray = {" ", "Eleven ", "Twelve ", "Thirteen ", "Forteen ", "Fifteen", "Sixteen ", "Seventeen ", "Eighteen ", "Nineteen "}
    '    Twnarray = {" ", "Ten ", "Twenty ", "Thirty ", "Forty ", "Fifty ", "Sixty ", "Seventy ", "Eighty ", "Ninety "}
    '    tmpVal = " "
    '    tmpVal = Twnarray(Tmpten) & " " & ReturnOneAmount(tmpones)
    '    If Tmpten = 1 And tmpones <> 0 Then
    '        tmpVal = Tenarray(tmpones)
    '    End If
    '    ReturnTenAmount = tmpVal
    'End Function

    'Private Function ReturnOneAmount(ByVal Tmpten As Integer) As String
    '    Dim Onearray As Object
    '    Dim tmpVal As String
    '    Onearray = {" ", "One ", "Two ", "Three ", "Four ", "Five ", "Six ", "Seven ", "Eight ", "Nine "}
    '    tmpVal = " "
    '    tmpVal = Onearray(Tmpten)
    '    ReturnOneAmount = tmpVal
    'End Function

    'Private Function Get_Paise(ByVal Tmpten As Integer, ByVal TmpOne As Integer) As String
    '    Dim tmpVal As String
    '    tmpVal = " PAISE " & ReturnTenAmount(Tmpten, TmpOne) & " Only."
    '    If Tmpten = 0 And TmpOne = 0 Then
    '        tmpVal = " Only."
    '    End If
    '    Get_Paise = tmpVal
    'End Function
    'Prakash 18/06/2008
    Function SpellRupee(ByVal MyNumber)
        Dim Dollars, Cents, Temp, Temp1, Dollars1, Temp2
        Dim DecimalPlace, Count

        'ReDim Place(9)
        'Place(2) = " THOUSAND "
        'Place(3) = " Million "
        'Place(4) = " Billion "
        'Place(5) = " Trillion "

        ' String representation of amount.
        MyNumber = Trim(Str(MyNumber))

        ' Position of decimal place 0 if none.
        DecimalPlace = InStr(MyNumber, ".")
        ' Convert cents and set MyNumber to dollar amount.
        If DecimalPlace > 0 Then
            Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) &
                "00", 2))
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If

        Count = 1
        Temp2 = ""
        If Len(MyNumber) > 7 Then
            Temp2 = GetHundreds(Left(MyNumber, Len(MyNumber) - 7)) & " CRORES "
            MyNumber = Right(MyNumber, 7)

        End If

        Temp1 = ""
        If Len(MyNumber) > 5 Then
            Temp1 = GetHundreds(Left(MyNumber, Len(MyNumber) - 5)) & " LAKHS "
            MyNumber = Right(MyNumber, 5)

        End If

        Do While MyNumber <> ""
            Temp = GetHundreds(Right(MyNumber, 3))
            ' If Temp <> "" Then Dollars1 = Temp & Place(Count) & Dollars1
            If Len(MyNumber) > 3 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
                MyNumber = ""
            End If
            Count = Count + 1
        Loop
        Dollars = Temp2 & Temp1 & Dollars1
        Select Case Dollars
            Case ""
                Dollars = "NIL RUPEES"
            Case "One"
                Dollars = "ONE RUPEE "
            Case Else
                Dollars = " RUPEES " & Dollars & " ONLY "
        End Select

        Select Case Cents
            Case ""
                Cents = " "
            Case "One"
                Cents = " "
            Case Else
                Cents = " And " & Cents & " PAISA"
        End Select

        SpellRupee = Dollars & Cents
    End Function




    ''*******************************************
    '' Converts a number from 100-999 into text *
    ''*******************************************

    Private Function GetHundreds(ByVal MyNumber)
        Dim Result As String

        If Val(MyNumber) = 0 Then Exit Function
        MyNumber = Right("000" & MyNumber, 3)

        ' Convert the hundreds place.
        If Mid(MyNumber, 1, 1) <> "0" Then
            Result = GetNumber(Mid(MyNumber, 1, 1)) & " HUNDRED "
        End If

        ' Convert the tens and ones place.
        If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & GetTens(Mid(MyNumber, 2))
        Else
            Result = Result & GetNumber(Mid(MyNumber, 3))
        End If

        GetHundreds = Result
    End Function



    ''*********************************************
    '' Converts a number from 10 to 99 into text. *
    ''*********************************************

    Private Function GetTens(ByVal TensText)
        Dim Result As String

        Result = ""           ' Null out the temporary function value.
        If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
            Select Case Val(TensText)
                Case 10 : Result = "TEN"
                Case 11 : Result = "ELEVEN"
                Case 12 : Result = "TWELVE"
                Case 13 : Result = "THIRTEEN"
                Case 14 : Result = "FOURTEEN"
                Case 15 : Result = "FIFTEEN"
                Case 16 : Result = "SIXTEEN"
                Case 17 : Result = "SEVENTEEN"
                Case 18 : Result = "EIGHTEEN"
                Case 19 : Result = "NINETEEN"
                Case Else
            End Select
        Else                                 ' If value between 20-99...
            Select Case Val(Left(TensText, 1))
                Case 2 : Result = "TWENTY "
                Case 3 : Result = "THIRTY "
                Case 4 : Result = "FORTY "
                Case 5 : Result = "FIFTY "
                Case 6 : Result = "SIXTY "
                Case 7 : Result = "SEVENTY "
                Case 8 : Result = "EIGHTY "
                Case 9 : Result = "NINETY "
                Case Else
            End Select
            Result = Result & GetNumber _
                (Right(TensText, 1))  ' Retrieve ones place.
        End If
        GetTens = Result
    End Function




    ''*******************************************
    '' Converts a number from 1 to 9 into text. *
    ''*******************************************

    Private Function GetNumber(ByVal Digit)
        Select Case Val(Digit)
            Case 1 : GetNumber = "ONE"
            Case 2 : GetNumber = "TWO"
            Case 3 : GetNumber = "THREE"
            Case 4 : GetNumber = "FOUR"
            Case 5 : GetNumber = "FIVE"
            Case 6 : GetNumber = "SIX"
            Case 7 : GetNumber = "SEVEN"
            Case 8 : GetNumber = "EIGHT"
            Case 9 : GetNumber = "NINE"
            Case Else : GetNumber = ""
        End Select
    End Function

    Public Function SetFormat(vFormat As String, Dval As String) As String
        Dim i As Integer

        For i = 1 To (Len(vFormat) - Len(Dval))
            Dval = "0" & Dval
        Next i
        SetFormat = Dval

    End Function

    Public Function SetFormat1(vFormat As String, Dval As String) As String
        Dim i As Integer

        For i = 1 To (Len(Dval) - Len(vFormat))
            Dval = "0" & vFormat
        Next i
        SetFormat1 = Dval

    End Function


    Function CheckPRNNo(ByVal FormNo As String) As Boolean
        Dim nds As New DataSet
        nds = FetchDataSet("Select * from RetnMst where coid=" & selectedcoid & " And FrmType='" & FormNo & "'")
        If nds.Tables(0).Rows.Count <= 0 Then
            CheckPRNNo = False
        ElseIf (nds.Tables(0).Rows(0)("prn").ToString() = "") Then
            CheckPRNNo = False
        Else
            If MsgBox("Return already filed for this quarter" & vbCrLf & "Do you still want to edit the return?", vbQuestion + vbYesNo, "Warning") = vbYes Then
                CheckPRNNo = False
            Else
                CheckPRNNo = True
            End If
        End If
        nds.Dispose()
    End Function

    Public Function FetchDataSet(SqlString As String) As DataSet
        Dim QueSt As String
        Dim headadaptor As New OleDbDataAdapter
        Dim cmd As New OleDbCommand
        Dim ds As New DataSet
        QueSt = SqlString
        Try
            cmd = New OleDbCommand(QueSt, cn)
            headadaptor = New OleDbDataAdapter
            ds = New DataSet
            headadaptor.SelectCommand = cmd
            headadaptor.Fill(ds)
            headadaptor.Dispose()
            cmd.Dispose()
            Return ds
        Catch ex As Exception
            MsgBox(ex.Message, "info")
            Return Nothing
        End Try
    End Function

    Public Sub EndMain()
        cn.Dispose()
        cn = Nothing
    End Sub


    Public Function FindALLControlRecursive(ByVal list As List(Of Control), ByVal parent As Control) As List(Of Control)
        ' function that returns all control in a form, parent or child regardless of control's type
        If parent Is Nothing Then
            Return list
        Else
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindALLControlRecursive(list, child)
        Next
        Return list
    End Function

End Module