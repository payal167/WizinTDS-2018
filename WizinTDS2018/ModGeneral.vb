

'Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports System.Windows.Forms
Module ModGeneral
    Public Const SWP_NOMOVE = 2
    Public Const SWP_NOSIZE = 1
    Public Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
    Public CtrlBkClr As System.Drawing.Color
    Public CtrlFrClr As System.Drawing.Color


    Public Enum MyKeypressEnum
        KeyPressDefault = 0
        KeypressUpperCase = 1
        KeyPressLowerCase = 2
        KeyPressProperCase = 3
        KeyPressNumberOnly = 4
        KeyPressAutoFind = 5
    End Enum

    'Variables/Constants for API
    Public Const CB_FINDSTRING = &H14C
    Public Const CB_ERR = (-1)

    Declare Function SendMessage Lib "user32" Alias _
                                     "SendMessageA" _
                                     (ByVal hwnd As Long,
                                      ByVal wMsg As Long,
                                      ByVal wParam As Long,
                                      ByVal lParam As Object) As Long


    Declare Function AppendMenu Lib "user32" Alias _
    "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags _
    As Long, ByVal wIDNewItem As Long, ByVal _
    lpNewItem As String) As Long

    Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, ByVal bRevert As Long) As Long

    Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hwnd As Long,
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

    Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" (ByVal lpPrevWndFunc _
    As Long, ByVal hwnd As Long, ByVal Msg As _
    Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Public Const WM_SYSCOMMAND = &H112
    Public Const MF_SEPARATOR = &H800&
    Public Const MF_STRING = &H0&
    Public Const GWL_WNDPROC = (-4)
    Public Const IDM_ABOUT As Long = 1010
    Public lProcOld As Long

    Public Declare Function GetTempPath Lib "kernel32" Alias _
        "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer _
        As String) As Long

    Public Const MAX_PATH = 260

    Public Declare Function SetWindowPos Lib "user32" _
         (ByVal hwnd As Long,
         ByVal hWndInsertAfter As Long,
         ByVal X As Long, ByVal Y As Long,
         ByVal cx As Long, ByVal cy As Long,
         ByVal wFlags As Long) As Long


    Public Sub CentreScreen(ByVal frm As Form)
        frm.Top = ((My.Computer.Screen.WorkingArea.Height / 2) - frm.Height / 2)
        frm.Left = ((My.Computer.Screen.WorkingArea.Width / 2) - frm.Width / 2)
    End Sub
    Public Sub chknum(ByVal txtbx As TextBox)
        If Trim(txtbx.Text) <> vbNullString Then
            If Len(txtbx) = 1 And Trim(txtbx.Text) = "." Then
                Exit Sub
            End If
            If Not (IsNumeric(txtbx)) Then
                MsgBox("Please enter numbers", vbApplicationModal + vbInformation, "Caution")
                txtbx.Text = CInt(txtbx.Text)
            Else
                If InStr(1, txtbx.Text, ",") > 0 Then
                    MsgBox("Please enter numbers", vbApplicationModal + vbInformation, "Caution")
                    txtbx.Text = Val(txtbx.Text)
                End If
            End If
        End If
    End Sub
    Public Sub EnableCtrls(ByVal frm As Form)
        Call EnableDisableCtrls(frm, True)
    End Sub

    Public Sub DisableCtrls(ByVal frm As Form)
        Call EnableDisableCtrls(frm, False)
    End Sub

    Public Sub EnableDisableCtrls(ByVal frm As Form, ByVal Enable As Boolean)
        Dim ctrl As Control
        For Each ctrl In frm.Controls
            If TypeOf ctrl Is TextBox Or
                  TypeOf ctrl Is DataGrid Or
                  TypeOf ctrl Is ComboBox Or
                  TypeOf ctrl Is RadioButton Or TypeOf ctrl Is CheckBox Or
                  TypeOf ctrl Is DateTimePicker Then
                ctrl.Enabled = Enable
            End If
            'if there is a treeview, disable it..if other are enabled and vice-versa
            If TypeOf ctrl Is TreeView Then ctrl.Enabled = Not Enable
        Next ctrl
    End Sub

    Public Function NormalMode(ByVal frm As Form)
        On Error Resume Next

        ''If (frm.cmds) = True Then
        'If frm.Controls.OfType(Button) Then
        '    'enable all buttons, and disable the required buttons..
        '    For I = 0 To frm.cmds.Count - 1
        '        frm.cmds(I).Enabled = True
        '        If frm.cmds(I).Caption = "&Save" Then
        '            frm.cmds(I).Enabled = False
        '        ElseIf frm.cmds(I).Caption = "&Cancel" Then
        '            frm.cmds(I).Caption = "&Exit"
        '        End If
        '    Next I
        'End If
    End Function

    'Public Function EditMode(ByVal frm As Form)
    'On Error Resume Next
    'Dim I As Long
    'If IsObject(frm.cmds) = True Then
    '    'disable all buttons, and enable the required buttons
    '    For I = 0 To frm.cmds.Count - 1
    '        frm.cmds(I).Enabled = False
    '        If frm.cmds(I).Caption = "&Save" Or frm.cmds(I).Caption = "&Refresh" Then
    '            frm.cmds(I).Enabled = True
    '        ElseIf frm.cmds(I).Caption = "E&xit" Then
    '            frm.cmds(I).Caption = "&Cancel"
    '            frm.cmds(I).Enabled = True
    '        End If
    '    Next I
    'End If
    'End Function

    Public Function NumToWords(ByVal tmpVal As Decimal) As String
        Dim tmpamt As Decimal, Tmpamtstr As String, tmplen As Integer
        Dim tmparay() As Integer, lclcounter As Integer, tmpword As String
        Dim Paisestr As String, TMPTREND As String, Tmptrends As Object ' tmpslitno As Object,
        Tmptrends = {" ", " ", " Hundred ", " Thousand ", " ", " Lakh ", " ", " Crore", " ", " Billion ", " ", " Trillion "}
        tmpamt = Int(tmpVal)
        Tmpamtstr = Trim(Str(tmpamt))
        tmplen = Len(Tmpamtstr)
        Paisestr = IIf(InStr(Str(tmpVal), ".") = 0, "00", Mid(Str(tmpVal), InStr(Str(tmpVal), ".") + 1, 2))
        ReDim tmparay(tmplen - 1)
        For lclcounter = 0 To tmplen - 1
            tmparay(lclcounter) = Mid(Tmpamtstr, tmplen - lclcounter, 1)
        Next
        TMPTREND = ""
        For lclcounter = UBound(tmparay) To 0 Step -1
            Select Case lclcounter
                Case 0, 3, 5, 7, 9, 11, 13
                    If (lclcounter = 0 Or lclcounter = 3 Or lclcounter = 5 Or lclcounter = 7 Or lclcounter = 9 Or lclcounter = 13 Or lclcounter = 11) And lclcounter + 1 <= UBound(tmparay) Then
                        TMPTREND = ReturnTenAmount(tmparay(lclcounter + 1), tmparay(lclcounter))
                    Else
                        TMPTREND = ReturnOneAmount(tmparay(lclcounter))
                    End If
                Case 2
                    TMPTREND = ReturnOneAmount(tmparay(lclcounter))
            End Select
            If lclcounter <= UBound(Tmptrends) Then TMPTREND = TMPTREND & IIf(IvalidNo(TMPTREND) = True, Tmptrends(lclcounter), " ")
            tmpword = tmpword & TMPTREND
            TMPTREND = ""
        Next
        NumToWords = tmpword & IIf(Val(Paisestr) > 0, " And ", " ") & Get_Paise(Val(Mid(Paisestr, 1, 1)), Val(Mid(Paisestr, 2, 1)))
        Do While InStr(1, NumToWords, "  ") <> 0
            NumToWords = Replace(Trim(NumToWords), Space(2), Space(1))
        Loop
    End Function

    Private Function ReturnTenAmount(ByVal Tmpten As Integer, ByVal tmpones As Integer) As String
        Dim Tenarray As Object, Twnarray As Object, tmpVal As String
        Tenarray = {" ", "Eleven ", "Twelve ", "Thirteen ", "Forteen ", "Fifteen", "Sixteen ", "Seventeen ", "Eighteen ", "Nineteen "}
        Twnarray = {" ", "Ten ", "Twenty ", "Thirty ", "Forty ", "Fifty ", "Sixty ", "Seventy ", "Eighty ", "Ninety "}
        tmpVal = " "
        tmpVal = Twnarray(Tmpten) & " " & ReturnOneAmount(tmpones)
        If Tmpten = 1 And tmpones <> 0 Then
            tmpVal = Tenarray(tmpones)
        End If
        ReturnTenAmount = tmpVal
    End Function

    Private Function ReturnOneAmount(ByVal Tmpten As Integer) As String
        Dim Onearray As Object
        Dim tmpVal As String
        Onearray = {" ", "One ", "Two ", "Three ", "Four ", "Five ", "Six ", "Seven ", "Eight ", "Nine "}
        tmpVal = " "
        tmpVal = Onearray(Tmpten)
        ReturnOneAmount = tmpVal
    End Function

    Private Function Get_Paise(ByVal Tmpten As Integer, ByVal TmpOne As Integer) As String
        Dim tmpVal As String
        tmpVal = " PAISE " & ReturnTenAmount(Tmpten, TmpOne) & " Only."
        If Tmpten = 0 And TmpOne = 0 Then
            tmpVal = " Only."
        End If
        Get_Paise = tmpVal
    End Function

    Public Sub CompactJetDatabase(ByVal Location As String,
        Optional ByVal BackupOriginal As Boolean = True, Optional ByVal pwd As String = "")

        On Error GoTo CompactErr

        Dim strBackupFile As String
        Dim strTempFile As String

        'Check the database exists
        If Len(Dir(Location)) Then

            ' If a backup is required, do it!
            If BackupOriginal = True Then
                strBackupFile = GetTemporaryPath() & "backup.mdb"
                If Len(Dir(strBackupFile)) Then Kill(strBackupFile)
                FileCopy(Location, strBackupFile)
            End If

            ' Create temporary filename
            strTempFile = GetTemporaryPath() & "temp.mdb"
            If Len(Dir(strTempFile)) Then Kill(strTempFile)

            ' Do the compacting via DBEngine
            '   If Trim(Pwd) = "" Then
            '      DBEngine.CompactDatabase Location, strTempFile
            '   Else
            '      DBEngine.CompactDatabase Location, strTempFile, , , ";Pwd=" & Pwd
            '   End If

            ' Remove the original database file
            Kill(Location)

            ' Copy the temporary now-compressed
            ' database file back to the original
            ' location
            FileCopy(strTempFile, Location)

            ' Delete the temporary file
            Kill(strTempFile)

        Else

        End If
        Exit Sub

CompactErr:
        MsgBox(Err.Number & ":" & Err.Description, vbCritical, "Unable to Compress Database")
        Exit Sub

    End Sub

    Public Function GetTemporaryPath()

        Dim strFolder As String
        Dim lngResult As Long

        strFolder = MAX_PATH.ToString()
        lngResult = GetTempPath(MAX_PATH, strFolder)

        If lngResult <> 0 Then
            GetTemporaryPath = strings.left(strFolder, InStr(strFolder,
             Chr(0)) - 1)
        Else
            GetTemporaryPath = ""
        End If

    End Function

    Public Sub SetTopmost(ByVal frm As Form, ByVal bTopmost As Boolean)
        Dim I As Long
        If bTopmost = True Then
            I = SetWindowPos(frm.Handle, HWND_TOPMOST,
                 0, 0, 0, 0, SWP_WNDFLAGS)
        Else
            I = SetWindowPos(frm.Handle, HWND_NOTOPMOST,
                 0, 0, 0, 0, SWP_WNDFLAGS)
        End If
    End Sub

    Public Function AutoFind(ByRef cboCurrent As ComboBox,
                             ByVal KeyAscii As Integer,
                             Optional ByVal LimitToList As Boolean = True)

        'Dim lCB As Long
        'Dim sFindString As String


        'On Error GoTo Err_handler
        'If KeyAscii = 8 Then
        '    If cboCurrent.SelectionStart <= 1 Then
        '        cboCurrent.Text = ""
        '        AutoFind = 0
        '        Exit Function
        '    End If
        '    If cboCurrent.SelectionLength = 0 Then
        '        sFindString = UCase(strings.left(cboCurrent.Text, Len(cboCurrent) - 1))
        '    Else
        '        sFindString = strings.left$(cboCurrent.Text, cboCurrent.SelectionStart - 1)
        '    End If
        'ElseIf KeyAscii < 32 Or KeyAscii > 127 Then
        '    Exit Function
        'Else
        '    If cboCurrent.SelectionLength = 0 Then
        '        sFindString = UCase(cboCurrent.Text & Chr(KeyAscii))
        '    Else
        '        sFindString = strings.left$(cboCurrent.Text, cboCurrent.SelectionStart) & Chr(KeyAscii)
        '    End If
        'End If
        'lCB = SendMessage(cboCurrent.Handle, CB_FINDSTRING, -1, sFindString) 'ByVal sFindString)

        'If lCB <> CB_ERR Then
        '    cboCurrent.SelectedIndex = lCB
        '    cboCurrent.SelectionStart = Len(sFindString)
        '    cboCurrent.SelectionLength = Len(cboCurrent.Text) - cboCurrent.SelectionStart
        '    AutoFind = 0
        '    Debug.Print(cboCurrent.SelectedIndex & "|" & frmTDS26Q.cboDedName.SelectedIndex)
        'Else
        '    If LimitToList = True Then
        AutoFind = 0
        '    Else
        AutoFind = KeyAscii
        '    End If
        'End If

Err_handler:
    End Function

    Public Function ValidEmail(ByVal sEmail As String, Optional ByRef sReason As String = "") As Boolean
        Dim sPrefix As String
        Dim sSuffix As String
        Dim sMiddle As String
        Dim nCharacter As Integer
        Dim sBuffer As String

        sEmail = Trim(sEmail)

        If Len(sEmail) < 8 Then
            'too short
            ValidEmail = False
            sReason = "Too Short"
            Exit Function
        End If

        If InStr(sEmail, "@") = 0 Then
            ' Missing @
            ValidEmail = False
            sReason = "Missing The @"
            Exit Function
        End If

        If InStr(InStr(sEmail, "@") + 1, sEmail, "@") <> 0 Then
            ' Too many @
            ValidEmail = False
            sReason = "Too Many @"
            Exit Function
        End If

        If InStr(sEmail, ".") = 0 Then
            ' missing dot
            ValidEmail = False
            sReason = "Missing the Period"
            Exit Function
        End If
        If InStr(sEmail, "@") = 1 Or InStr(sEmail, "@") = Len(sEmail) Or
            InStr(sEmail, ".") = 1 Or InStr(sEmail, ".") = Len(sEmail) Then
            ' either @ or . is at start or at the end, invalid format
            ValidEmail = False
            sReason = "Invalid Format"
            Exit Function
        End If

        For nCharacter = 1 To Len(sEmail)
            sBuffer = Mid(sEmail, nCharacter, 1)
            If Not (LCase(sBuffer) Like "[a-z]" Or sBuffer = "@" Or
                sBuffer = "." Or sBuffer = "-" Or sBuffer = "_" Or
                IsNumeric(sBuffer)) Then
                ' invalid character is there in it.
                ValidEmail = False
                sReason = "Invalid Character"
                Exit Function
            End If
        Next nCharacter
        nCharacter = 0

        On Error Resume Next
        sBuffer = Right(sEmail, 4)
        If InStr(sBuffer, ".") = 0 Then GoTo TooLong
        If strings.left(sBuffer, 1) = "." Then sBuffer = Right(sBuffer, 3)
        If strings.left(Right(sBuffer, 3), 1) = "." Then sBuffer = Right(sBuffer, 2)
        If strings.left(Right(sBuffer, 2), 1) = "." Then sBuffer = Right(sBuffer, 1)

        If Len(sBuffer) < 2 Then
            ' suffix is too short
            ValidEmail = False
            sReason = "Suffix Too Short"
            Exit Function
        End If

TooLong:

        If Len(sBuffer) > 3 Then
            ' suffix is too long
            ValidEmail = False
            sReason = "Suffix Too Long"
            Exit Function
        End If

        'everything is ok.
        sReason = ""
        ValidEmail = True
    End Function

    ' Number only function...
    Private Function NumberOnly(ByVal ctl As TextBox, ByVal KeyAscii As Integer,
        Optional ByVal Fraction As Integer = 2,
        Optional ByVal AllowMinus As Boolean = True) As Integer

        Dim DecimalPosition As Integer
        DecimalPosition = InStr(1, ctl.Text, ".")
        If KeyAscii < 48 Or KeyAscii > 57 Then
            ' Non numeric keys...
            If KeyAscii = 46 Then 'decimal/period
                If Fraction = 0 Then
                    ' decimal entry is not allowed
                    KeyAscii = 0
                Else
                    If DecimalPosition <> 0 Then
                        ' decimal is already there..
                        KeyAscii = 0
                    ElseIf (Len(ctl.Text) - ctl.SelectionStart) > Fraction Then
                        ' decimal is being inserted..
                        KeyAscii = 0
                    End If
                End If
            Else
                If KeyAscii = 127 Or KeyAscii = 8 Then
                    ' Backspace and delete key
                    ' do nothing..
                ElseIf KeyAscii = 45 Then
                    If AllowMinus = True Then
                        If InStr(1, ctl.Text, "-") <> 0 Then
                            ' Minus already there. Make it plus
                            ctl.Text = Mid(ctl.Text, 2, Len(ctl))
                            SendKeys.Send("{End}")
                            KeyAscii = 0
                        Else
                            ' No Minus sign, add it..
                            ctl.Text = "-" & ctl.Text
                            SendKeys.Send("{End}")
                            KeyAscii = 0   ' we have already added it..
                        End If
                    Else
                        ' No minus sign is allowed
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = 43 Then
                    ' Plus sign..
                    If InStr(1, ctl.Text, "-") <> 0 Then
                        'Minus is there, make it plus (ie remove it)
                        ctl.Text = Mid(ctl.Text, 2, Len(ctl))
                        SendKeys.Send("{End}")
                        KeyAscii = 0
                    Else
                        ' We dont want to add a plus sign
                        KeyAscii = 0
                    End If
                Else
                    ' All other keys..
                    KeyAscii = 0
                End If
            End If
        Else
            ' numeric keys..
            If Fraction <> 0 Then
                ' Decimal is allowed, check whether decimal is already there.
                If DecimalPosition <> 0 Then
                    ' Yes, now get how many chars are there after deci.
                    If (Len(ctl.Text) - DecimalPosition) >= Fraction Then
                        If (ctl.SelectionStart >= DecimalPosition) Then
                            'Filled..No more numbers are allowed
                            KeyAscii = 0
                        End If
                    End If
                End If
            End If
        End If
        NumberOnly = KeyAscii
    End Function

    Private Function NumberOnlyfrCombo(ByVal ctl As ComboBox, ByVal KeyAscii As Integer,
        Optional ByVal Fraction As Integer = 2,
        Optional ByVal AllowMinus As Boolean = True) As Integer

        Dim DecimalPosition As Integer
        DecimalPosition = InStr(1, ctl.Text, ".")
        If KeyAscii < 48 Or KeyAscii > 57 Then
            ' Non numeric keys...
            If KeyAscii = 46 Then 'decimal/period
                If Fraction = 0 Then
                    ' decimal entry is not allowed
                    KeyAscii = 0
                Else
                    If DecimalPosition <> 0 Then
                        ' decimal is already there..
                        KeyAscii = 0
                    ElseIf (Len(ctl.Text) - ctl.SelectionStart) > Fraction Then
                        ' decimal is being inserted..
                        KeyAscii = 0
                    End If
                End If
            Else
                If KeyAscii = 127 Or KeyAscii = 8 Then
                    ' Backspace and delete key
                    ' do nothing..
                ElseIf KeyAscii = 45 Then
                    If AllowMinus = True Then
                        If InStr(1, ctl.Text, "-") <> 0 Then
                            ' Minus already there. Make it plus
                            ctl.Text = Mid(ctl.Text, 2, Len(ctl))
                            SendKeys.Send("{End}")
                            KeyAscii = 0
                        Else
                            ' No Minus sign, add it..
                            ctl.Text = "-" & ctl.Text
                            SendKeys.Send("{End}")
                            KeyAscii = 0   ' we have already added it..
                        End If
                    Else
                        ' No minus sign is allowed
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = 43 Then
                    ' Plus sign..
                    If InStr(1, ctl.Text, "-") <> 0 Then
                        'Minus is there, make it plus (ie remove it)
                        ctl.Text = Mid(ctl.Text, 2, Len(ctl))
                        SendKeys.Send("{End}")
                        KeyAscii = 0
                    Else
                        ' We dont want to add a plus sign
                        KeyAscii = 0
                    End If
                Else
                    ' All other keys..
                    KeyAscii = 0
                End If
            End If
        Else
            ' numeric keys..
            If Fraction <> 0 Then
                ' Decimal is allowed, check whether decimal is already there.
                If DecimalPosition <> 0 Then
                    ' Yes, now get how many chars are there after deci.
                    If (Len(ctl.Text) - DecimalPosition) >= Fraction Then
                        If (ctl.SelectionStart >= DecimalPosition) Then
                            'Filled..No more numbers are allowed
                            KeyAscii = 0
                        End If
                    End If
                End If
            End If
        End If
        NumberOnlyfrCombo = KeyAscii
    End Function

    ' Number only function...
    'Public Function NumberOnly4Grid(ByRef ctl, ByVal KeyAscii As Integer,
    '    Optional ByVal Fraction As Integer = 2,
    '    Optional ByVal AllowMinus As Boolean = True) As Integer

    '    Dim DecimalPosition As Integer
    '    DecimalPosition = InStr(1, ctl.EditText, ".")
    '    If KeyAscii < 48 Or KeyAscii > 57 Then
    '        ' Non numeric keys...
    '        If KeyAscii = 46 Then 'decimal/period
    '            If Fraction = 0 Then
    '                ' decimal entry is not allowed
    '                KeyAscii = 0
    '            Else
    '                If DecimalPosition <> 0 Then
    '                    ' decimal is already there..
    '                    KeyAscii = 0
    '                ElseIf (Len(ctl.EditText) - ctl.EditSelStart) > Fraction Then
    '                    ' decimal is being inserted..
    '                    KeyAscii = 0
    '                End If
    '            End If
    '        Else
    '            If KeyAscii = 127 Or KeyAscii = 8 Then
    '                ' Backspace and delete key
    '                ' do nothing..
    '            ElseIf KeyAscii = 45 Then
    '                If AllowMinus = True Then
    '                    If InStr(1, ctl.EditText, "-") <> 0 Then
    '                        ' Minus already there. Make it plus
    '                        ctl.EditText = Mid(ctl, 2, Len(ctl))
    '                        'SendKeys "{End}"
    '                        KeyAscii = 0
    '                    Else
    '                        ' No Minus sign, add it..
    '                        ctl.EditText = "-" & ctl.EditText
    '                        'SendKeys "{End}"
    '                        KeyAscii = 0   ' we have already added it..
    '                    End If
    '                Else
    '                    ' No minus sign is allowed
    '                    KeyAscii = 0
    '                End If
    '            ElseIf KeyAscii = 43 Then
    '                ' Plus sign..
    '                If InStr(1, ctl.EditText, "-") <> 0 Then
    '                    'Minus is there, make it plus (ie remove it)
    '                    ctl.EditText = Mid(ctl.EditText, 2, Len(ctl.EditText))
    '                    SendKeys("{End}")
    '                    KeyAscii = 0
    '                Else
    '                    ' We dont want to add a plus sign
    '                    KeyAscii = 0
    '                End If
    '            Else
    '                ' All other keys..
    '                KeyAscii = 0
    '            End If
    '        End If
    '    Else
    '        ' numeric keys..
    '        If Fraction <> 0 Then
    '            ' Decimal is allowed, check whether decimal is already there.
    '            If DecimalPosition <> 0 Then
    '                ' Yes, now get how many chars are there after deci.
    '                If (Len(ctl.EditText) - DecimalPosition) >= Fraction Then
    '                    If (ctl.EditSelStart >= DecimalPosition) Then
    '                        'Filled..No more numbers are allowed
    '                        KeyAscii = 0
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End If
    '    NumberOnly4Grid = KeyAscii
    'End Function




    ' MyFormat function, to format the number string..
    Private Function MyFormat(ByVal txt As String, ByVal frmstr As String) As String
        On Error Resume Next
        MyFormat = Replace(Trim(Format(txt, Replace(frmstr, ",", " "))), " ", ",")
    End Function

    ' code to convert number into words...
    Private Function IvalidNo(ByVal tmpstring As String) As Boolean
        Dim tmpbool As Boolean
        tmpbool = False
        If Len(Trim(tmpstring)) > 0 Then tmpbool = True
        IvalidNo = tmpbool
    End Function

    Private Function NoToWordsNEW(ByVal tmpVal As Decimal) As String
        Dim tmpamt As Decimal, Tmpamtstr As String, tmplen As Integer
        Dim tmparay() As Integer, lclcounter As Integer, tmpword As String
        Dim Paisestr As String, TMPTREND As String, Tmptrends As Object 'tmpslitno As Object,
        tmpword = ""
        Tmptrends = {" ", " ", " Hundred ", " Thousand ", " ", " Lakh ", " ", " Crore", " ", " Billion ", " ", " Trillion "}
        tmpamt = Int(tmpVal)
        Tmpamtstr = Trim(Str(tmpamt))
        tmplen = Len(Tmpamtstr)
        Paisestr = IIf(InStr(Str(tmpVal), ".") = 0, "00", Mid(Str(tmpVal), InStr(Str(tmpVal), ".") + 1, 2))
        ReDim tmparay(tmplen - 1)
        For lclcounter = 0 To tmplen - 1
            tmparay(lclcounter) = Mid(Tmpamtstr, tmplen - lclcounter, 1)
        Next
        TMPTREND = ""
        For lclcounter = UBound(tmparay) To 0 Step -1
            Select Case lclcounter
                Case 0, 3, 5, 7, 9, 11, 13
                    If (lclcounter = 0 Or lclcounter = 3 Or lclcounter = 5 Or lclcounter = 7 Or lclcounter = 9 Or lclcounter = 13 Or lclcounter = 11) And lclcounter + 1 <= UBound(tmparay) Then
                        TMPTREND = ReturnTenAmount(tmparay(lclcounter + 1), tmparay(lclcounter))
                    Else
                        TMPTREND = ReturnOneAmount(tmparay(lclcounter))
                    End If
                Case 2
                    TMPTREND = ReturnOneAmount(tmparay(lclcounter))
            End Select
            If lclcounter <= UBound(Tmptrends) Then TMPTREND = TMPTREND & IIf(IvalidNo(TMPTREND) = True, Tmptrends(lclcounter), " ")
            tmpword = tmpword & TMPTREND
            TMPTREND = ""
        Next
        NoToWordsNEW = tmpword & IIf(Val(Paisestr) > 0, " And ", " ") & Get_Paise(Val(Mid(Paisestr, 1, 1)), Val(Mid(Paisestr, 2, 1)))
        Do While InStr(1, NoToWordsNEW, "  ") <> 0
            NoToWordsNEW = Replace(Trim(NoToWordsNEW), Space(2), Space(1))
        Loop
    End Function

    Public Function ProperCase(ByVal KEY As Integer, ByVal ctrl As TextBox) As Integer
        If ctrl.SelectionStart <> 0 Then
            If InStr(1, ".,/[()]; ", Mid(ctrl.Text, ctrl.SelectionStart, 1)) <> 0 Or
               Asc(Mid(ctrl.Text, ctrl.SelectionStart, 1)) = 10 Then
                'Ascii(10) = Enter key in Multiline Text Box..
                ProperCase = Asc(UCase(Chr(KEY)))
            Else
                ProperCase = KEY
            End If
        Else
            ProperCase = Asc(UCase(Chr(KEY)))
        End If
    End Function

    Public Sub CtrlGotFocus(ByVal ctrl As TextBox)
        On Error Resume Next
        CtrlBkClr = ctrl.BackColor
        CtrlFrClr = ctrl.ForeColor
        ctrl.BackColor = Color.LightYellow 'SystemColorConstants.vbInfoBackground
        ctrl.ForeColor = Color.Black 'SystemColorConstants.vbInfoText
        'ctrl. = 0
        ctrl.Select(0, Len(ctrl.Text))
    End Sub

    Public Sub CtrlGotFocusC(ByVal ctrl As ComboBox)
        On Error Resume Next
        CtrlBkClr = ctrl.BackColor
        CtrlFrClr = ctrl.ForeColor
        ctrl.BackColor = Color.LightYellow 'SystemColorConstants.vbInfoBackground
        ctrl.ForeColor = Color.Black 'SystemColorConstants.vbInfoText
        'ctrl. = 0
        ctrl.Select(0, Len(ctrl.Text))
    End Sub

    Public Sub CtrlGotFocusDate(ByVal ctrl As MaskedTextBox)
        On Error Resume Next
        CtrlBkClr = ctrl.BackColor
        CtrlFrClr = ctrl.ForeColor
        ctrl.BackColor = Color.LightYellow 'SystemColorConstants.vbInfoBackground
        ctrl.ForeColor = Color.Black 'SystemColorConstants.vbInfoText
        'ctrl. = 0
        ctrl.Select(0, Len(ctrl.Text))
    End Sub

    Public Sub CtrlLostFocus(ByVal ctrl As Control)
        On Error Resume Next
        ctrl.BackColor = CtrlBkClr
        ctrl.ForeColor = CtrlFrClr

    End Sub

   Public Function CtrlKeyPress(ByVal ctrl As Control, ByVal KeyAscii As Integer, ByVal What2Do As MyKeypressEnum,
                   Optional ByVal DecPlace As Integer = 2, Optional ByVal AllowMinus As Boolean = True,
                   Optional ByVal EnterKeyMoves2NextCtrl As Boolean = True, Optional ByVal NoSpecialChar As Boolean = False) As Integer
        'this is required first, as other procedures may make keyascii as zero.
        If EnterKeyMoves2NextCtrl = True Then
            If KeyAscii = 13 Then SendKeys.Send("{tab}") : Return CtrlKeyPress = 0 : Exit Function
        End If
        Select Case What2Do
            Case MyKeypressEnum.KeyPressDefault
                Return CtrlKeyPress = KeyAscii
            Case MyKeypressEnum.KeypressUpperCase
                Return Asc(UCase(Chr(KeyAscii)))
            Case MyKeypressEnum.KeyPressLowerCase
                Return CtrlKeyPress = Asc(LCase(Chr(KeyAscii)))
            Case MyKeypressEnum.KeyPressProperCase
                CtrlKeyPress = ProperCase(KeyAscii, ctrl)
            Case MyKeypressEnum.KeyPressNumberOnly
                If TypeOf ctrl Is TextBox Then
                    Return CtrlKeyPress = NumberOnly(ctrl, KeyAscii, DecPlace, AllowMinus)
                ElseIf TypeOf ctrl Is ComboBox Then
                    Return CtrlKeyPress = NumberOnlyfrCombo(ctrl, KeyAscii, DecPlace, AllowMinus)
                End If
            Case MyKeypressEnum.KeyPressAutoFind
                If TypeOf ctrl Is ComboBox Then
                    Return CtrlKeyPress = AutoFind(ctrl, KeyAscii, True)
                Else
                    MsgBox("This What2Do Type is not valid for Combo Box", vbExclamation, "Invalid Enum")
                    Return CtrlKeyPress = KeyAscii
                End If
        End Select
        If CtrlKeyPress = 34 Or CtrlKeyPress = 39 Then
            Return CtrlKeyPress = 0
        End If
        If NoSpecialChar = True Then
            If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                Return CtrlKeyPress = 0
            End If
        End If
    End Function

    'Public Sub FillCombo(Obj As Object, Id As String, Nm As String, TblName As String, Optional WhereFldName As String, Optional WhereId As Long)
    'Dim Cnn As New clsConnect
    'Dim rs As New ADODB.Recordset
    'Dim Sql As String
    '    Obj.Clear
    '    Sql = "Select " & Id & "  as Id,  " & Nm & " as Name from  " & TblName
    '    If WhereFldName <> vbNullString And IsNull(WhereId) = False Then
    '       Sql = Sql & " Where " & WhereFldName & " = " & WhereId
    '    End If
    '    Sql = Sql & " order By " & Nm & ""
    '    rs.Open Sql, Cnn.CnnObj, adOpenForwardOnly, adLockReadOnly
    '    If rs.RecordCount <> 0 Then
    '        While Not rs.EOF
    '            If TypeOf Obj Is ComboBox Then
    '                    Obj.AddItem rs!Name
    '                    Obj.ItemData(Obj.NewIndex) = rs!Id
    '            Else
    '
    '            End If
    '        rs.MoveNext
    '        Wend
    '        Obj.ListIndex = 0
    '    End If
    '
    'End Sub

    ' Number only function.for vsflexgrid..
    'Public Function NumberOnly4vsGrid(ByVal ctl As DataGrid, ByVal KeyAscii As Integer,
    '    Optional ByVal Fraction As Integer = 2,
    '    Optional ByVal AllowMinus As Boolean = True) As Integer

    '    Dim DecimalPosition As Integer
    '    DecimalPosition = InStr(1, ctl.EditText, ".")
    '    If KeyAscii < 48 Or KeyAscii > 57 Then
    '        ' Non numeric keys...
    '        If KeyAscii = 46 Then 'decimal/period
    '            If Fraction = 0 Then
    '                ' decimal entry is not allowed
    '                KeyAscii = 0
    '            Else
    '                If DecimalPosition <> 0 Then
    '                    ' decimal is already there..
    '                    KeyAscii = 0
    '                ElseIf (Len(ctl.EditText) - ctl.EditSelStart) > Fraction Then
    '                    ' decimal is being inserted..
    '                    KeyAscii = 0
    '                End If
    '            End If
    '        Else
    '            If KeyAscii = 127 Or KeyAscii = 8 Then
    '                ' Backspace and delete key
    '                ' do nothing..
    '            ElseIf KeyAscii = 45 Then
    '                If AllowMinus = True Then
    '                    If InStr(1, ctl.EditText, "-") <> 0 Then
    '                        ' Minus already there. Make it plus
    '                        ctl.EditText = Mid(ctl.EditText, 2, Len(ctl.EditText))
    '                        SendKeys.Send("{End}")
    '                        KeyAscii = 0
    '                    Else
    '                        ' No Minus sign, add it..
    '                        ctl.EditText = "-" & ctl.EditText
    '                        SendKeys.Send("{End}")
    '                        KeyAscii = 0   ' we have already added it..
    '                    End If
    '                Else
    '                    ' No minus sign is allowed
    '                    KeyAscii = 0
    '                End If
    '            ElseIf KeyAscii = 43 Then
    '                ' Plus sign..
    '                If InStr(1, ctl.EditText, "-") <> 0 Then
    '                    'Minus is there, make it plus (ie remove it)
    '                    ctl.EditText = Mid(ctl.EditText, 2, Len(ctl.editext))
    '                    SendKeys.Send("{End}")
    '                    KeyAscii = 0
    '                Else
    '                    ' We dont want to add a plus sign
    '                    KeyAscii = 0
    '                End If
    '            Else
    '                ' All other keys..
    '                KeyAscii = 0
    '            End If
    '        End If
    '    Else
    '        ' numeric keys..
    '        If Fraction <> 0 Then
    '            ' Decimal is allowed, check whether decimal is already there.
    '            If DecimalPosition <> 0 Then
    '                ' Yes, now get how many chars are there after deci.
    '                If (Len(ctl.EditText) - DecimalPosition) >= Fraction Then
    '                    If (ctl.EditSelStart >= DecimalPosition) Then
    '                        'Filled..No more numbers are allowed
    '                        KeyAscii = 0
    '                    End If
    '                End If
    '            End If
    '        End If

    '    End If
    '    NumberOnly4vsGrid = KeyAscii
    'End Function

    Public Sub FindInCombo(ByVal cmbMyComboBox As ComboBox, ByVal LimitTextToList As Boolean)
        Dim I As Long, txt As String
        txt = cmbMyComboBox.Text
        If LimitTextToList = True Then
            For I = 0 To cmbMyComboBox.Items.Count - 1
                cmbMyComboBox.SelectedIndex = I
                If txt = cmbMyComboBox.Text Then
                    cmbMyComboBox.SelectedIndex = I
                    Exit For
                End If
            Next I
        End If
    End Sub

    Public Function IsValidPAN(ByVal PAN As String, Optional ByVal AllowBlank As Boolean = False, Optional ByVal AllowPANAPPLIED As Boolean = False) As Integer
        Dim I As Integer, HasNumbers As Boolean, ValidFormat As Boolean
        'check for blank first..
        If AllowBlank = True Then
            If Len(Trim(PAN)) = 0 Then
                IsValidPAN = 0      'Valid
                Exit Function
            End If
        End If
        'Not blank, now check the length..
        If Len(Trim(PAN)) <> 10 Then
            IsValidPAN = 1      'Invalid Length
            Exit Function
        End If
        'length is ok..now check format..
        'Check for numbers
        For I = 1 To Len(PAN)
            If IsNumeric(Mid(PAN, I, 1)) = True Then
                HasNumbers = True
                Exit For
            End If
        Next I
        If HasNumbers = False Then
            'Does not have numbers, must be TANAPPLIED if allowed.
            If AllowPANAPPLIED = True Then
                If UCase(PAN) <> "PANAPPLIED" Then
                    IsValidPAN = 2      'Invalid String
                Else
                    IsValidPAN = 0      'Valid
                End If
                Exit Function
            Else
                'Numbers not found and TANapplied is also false, must be in proper format..
                'hence invalid..
                IsValidPAN = 3  'Not proper format
                Exit Function
            End If
        Else
            'there are numbers, it must be in AAAAA9999A format
            For I = 1 To Len(PAN)
                Select Case I
                    Case 1, 2, 3, 4, 5, 10  'Alphabets
                        If Asc(UCase(Mid(PAN, I, 1))) >= Asc("A") And Asc(UCase(Mid(PAN, I, 1))) <= Asc("Z") Then
                            ValidFormat = True
                        Else
                            ValidFormat = False
                            Exit For
                        End If
                    Case 6, 7, 8, 9 'Numbers
                        If IsNumeric(Mid(PAN, I, 1)) = True Then
                            ValidFormat = True
                        Else
                            ValidFormat = False
                            Exit For
                        End If
                End Select
            Next I
            If ValidFormat = True Then
                'check the fourth char
                Select Case Mid(PAN, 4, 1)
                    Case "P", "H", "C", "J", "F", "A", "T", "B", "L", "G"
                        IsValidPAN = 0
                    Case Else
                        IsValidPAN = 4
                End Select
                '        If Mid(PAN, 4, 1) <> "P" Or Mid(PAN, 4, 1) <> "C" Or Mid(PAN, 4, 1) <> "F" Or _
                '            Mid(PAN, 4, 1) <> "A" Or Mid(PAN, 4, 1) <> "B" Or Mid(PAN, 4, 1) <> "G" Or _
                '            Mid(PAN, 4, 1) <> "J" Then
                '            IsValidPAN = 4  'invalid 4th char
                '        Else
                '            IsValidPAN = 0  'Valid
                '        End If
            Else
                IsValidPAN = 3  'Not proper format.
            End If
        End If
    End Function

    Public Function IsValidTAN(ByVal TAN As String, Optional ByVal AllowBlank As Boolean = False, Optional ByVal AllowTANAPPLIED As Boolean = False) As Integer
        Dim I As Integer, HasNumbers As Boolean, ValidFormat As Boolean
        'check for blank first..
        If AllowBlank = True Then
            If Len(Trim(TAN)) = 0 Then
                IsValidTAN = 0      'Valid
                Exit Function
            End If
        End If
        'Not blank, now check the length..
        If Len(Trim(TAN)) <> 10 Then
            IsValidTAN = 1      'Invalid Length
            Exit Function
        End If
        'length is ok..now check format..
        'Check for numbers
        For I = 1 To Len(TAN)
            If IsNumeric(Mid(TAN, I, 1)) = True Then
                HasNumbers = True
                Exit For
            End If
        Next I
        If HasNumbers = False Then
            'Does not have numbers, must be TANAPPLIED if allowed.
            If AllowTANAPPLIED = True Then
                If UCase(TAN) <> "TANAPPLIED" Then
                    IsValidTAN = 2      'Invalid String
                Else
                    IsValidTAN = 0      'Valid
                End If
                Exit Function
            Else
                'Numbers not found and TANapplied is also false, must be in proper format..
                'hence invalid..
                IsValidTAN = 3  'Not proper format
                Exit Function
            End If
        Else
            'there are numbers, it must be in AAAAA9999A format
            For I = 1 To Len(TAN)
                Select Case I
                    Case 1, 2, 3, 4, 10  'Alphabets
                        If Asc(UCase(Mid(TAN, I, 1))) >= Asc("A") And Asc(UCase(Mid(TAN, I, 1))) <= Asc("Z") Then
                            ValidFormat = True
                        Else
                            ValidFormat = False
                            Exit For
                        End If
                    Case 5, 6, 7, 8, 9 'Numbers
                        If IsNumeric(Mid(TAN, I, 1)) = True Then
                            ValidFormat = True
                        Else
                            ValidFormat = False
                            Exit For
                        End If
                End Select
            Next I
            If ValidFormat = True Then
                Dim ModVal As Integer
                ModVal = CLng(Mid(TAN, 5, 5)) Mod 7
                If ModVal = 0 And Right(TAN, 1) = "A" Then
                    IsValidTAN = 0  'Valid
                ElseIf ModVal = 1 And Right(TAN, 1) = "B" Then
                    IsValidTAN = 0
                ElseIf ModVal = 2 And Right(TAN, 1) = "C" Then
                    IsValidTAN = 0
                ElseIf ModVal = 3 And Right(TAN, 1) = "D" Then
                    IsValidTAN = 0
                ElseIf ModVal = 4 And Right(TAN, 1) = "E" Then
                    IsValidTAN = 0
                ElseIf ModVal = 5 And Right(TAN, 1) = "F" Then
                    IsValidTAN = 0
                ElseIf ModVal = 6 And Right(TAN, 1) = "G" Then
                    IsValidTAN = 0
                Else
                    IsValidTAN = 4  'last char not proper
                End If
            Else
                IsValidTAN = 3  'Not proper format.
            End If
        End If
    End Function

    Public Function HasBlankCertDt(ByVal selectedcoid As Long, ByVal frmid As Long) As Boolean
        On Error GoTo cleanup
        'Cnn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Dim RetnID As Long
        Dim nds As New DataSet
        nds = FetchDataSet("SELECT RetnID from RetnMst where CoID = " & selectedcoid & " and FrmType = " & frmid)
        'rst.Open("Select RetnID from RetnMst where CoID = " & selectedcoid & " And FrmType = " & frmid, Cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If nds.Tables(0).Rows.Count <= 0 Then
            RetnID = -1
        Else
            RetnID = Convert.ToInt32(nds.Tables(0).Rows(0)("RetnID"))
        End If
        nds = New DataSet
        If frmid = 26 Then
            'rst1.Open("Select * from DeducteeTDS where RetnID = " & RetnID & " And isnull(CertificateDt) And reason <>'B'", Cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            nds = FetchDataSet("Select * from DeducteeTDS where RetnID = " & RetnID & " And isnull(CertificateDt) And reason <>'B'")
        ElseIf frmid = 27 Or frmid = 28 Or frmid = 29 Or frmid = 30 Then
            'rst1.Open("SELECT * from Deductee27 where RetnID = " & RetnID & " and isnull(certificateDt) and reason <>'B'", Cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            nds = FetchDataSet("SELECT * from Deductee27 where RetnID = " & RetnID & " and isnull(certificateDt) and reason <>'B'")
        ElseIf frmid = 31 Then
            'rst1.Open("SELECT * from Deductee27E where RetnID = " & RetnID & " and isnull(certificateDt) and reason <>'B'", Cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            nds = FetchDataSet("SELECT * from Deductee27E where RetnID = " & RetnID & " and isnull(certificateDt) and reason <>'B'")
        End If
        'rst1.MoveLast
        If nds.Tables(0).Rows.Count > 0 Then
            HasBlankCertDt = True
        Else
            HasBlankCertDt = False
        End If
cleanup:
        'If rst.State = ADODB.ObjectStateEnum.adStateOpen Then rst.Close()
        'If rst1.State = ADODB.ObjectStateEnum.adStateOpen Then rst1.Close()
        'rst = Nothing
        'rst1 = Nothing
        nds.Dispose()
        nds = Nothing
        RetnID = Nothing
    End Function

    Public Function GetBlankCertRec(ByVal selectedcoid As Long, ByVal frmid As Long) As DataSet
        On Error GoTo cleanup
        'Cnn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Dim RetnID As Long
        Dim nds As New DataSet
        'rst.Open("SELECT RetnID from RetnMst where CoID = " & selectedcoid & " and FrmType = " & frmid, Cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        nds = FetchDataSet("SELECT RetnID from RetnMst where CoID = " & selectedcoid & " and FrmType = " & frmid)
        If nds.Tables(0).Rows.Count <= 0 Then
            RetnID = -1
        Else
            RetnID = Convert.ToInt32(nds.Tables(0).Rows(0)("RetnID"))
        End If
        If frmid = 26 Then
            'rst1.Open("SELECT r.*,d.dname from DeducteeTDS as r,DeductMst as d where r.RetnID = " & RetnID & " And isnull(r.CertificateDt) And r.DID=d.DID", Cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            nds = FetchDataSet("SELECT r.*,d.dname from DeducteeTDS as r,DeductMst as d where r.RetnID = " & RetnID & " And isnull(r.CertificateDt) And r.DID=d.DID")
        ElseIf frmid = 27 Or frmid = 28 Or frmid = 29 Or frmid = 30 Then
            'rst1.Open("SELECT r.*,d.dname from Deductee27 as r,DeductMst as d where r.RetnID = " & RetnID & " and isnull(r.CertificateDt) and r.DID=d.DID", Cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            nds = FetchDataSet("SELECT r.*,d.dname from Deductee27 as r,DeductMst as d where r.RetnID = " & RetnID & " and isnull(r.CertificateDt) and r.DID=d.DID")
        End If
        GetBlankCertRec = nds
cleanup:
        '   If rst.State = adStateOpen Then rst.Close
        '   If rst1.State = adStateOpen Then rst1.Close
        '   Set rst = Nothing
        '   Set rst1 = Nothing
        nds.Dispose()
        nds = Nothing
        RetnID = Nothing

    End Function

End Module
