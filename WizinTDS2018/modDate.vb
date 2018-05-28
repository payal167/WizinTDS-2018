Imports System.Globalization
Imports System.Threading
Module modDate

    Public Const LOCALE_SLANGUAGE As Long = &H2
    Public Const LOCALE_SSHORTDATE As Long = &H1F

    Public Const DATE_LONGDATE As Long = &H2
    Public Const DATE_SHORTDATE As Long = &H1

    Public Const HWND_BROADCAST As Long = &HFFFF&
    Public Const WM_SETTINGCHANGE As Long = &H1A

    Public Declare Function PostMessage Lib "user32" _
   Alias "PostMessageA" _
  (ByVal hwnd As Long,
   ByVal wMsg As Long,
   ByVal wParam As Long, lParam As Long) As Long

    Public Declare Function EnumDateFormats Lib "kernel32" _
   Alias "EnumDateFormatsA" _
  (ByVal lpDateFmtEnumProc As Long,
   ByVal Locale As Long,
   ByVal dwFlags As Long) As Long

    Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As VariantType,
   Source As VariantType,
   ByVal Length As Long)

    Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

    Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long,
   ByVal LCType As Long,
   ByVal lpLCData As String,
   ByVal cchData As Long) As Long

    Public Declare Function SetLocaleInfo Lib "kernel32" _
    Alias "SetLocaleInfoA" _
   (ByVal Locale As Long,
    ByVal LCType As Long,
    ByVal lpLCData As String) As Long
    'Function ChangeDate(xFormat As String)
    '    Dim ODate = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Control-Panel\International", "sShortDate", "dd/MM/yyyy")
    '    If ODate <> "dd/MM/yyyy" Then
    '        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)
    '        Microsoft.Win32.Registry.SetValue("HKEY_CURRENT_USER\Control-Panel\International", "sShortDate", "dd/MM/yyyy")
    '        'MsgBox("Date of format is wrong..Please Restart software..!!")
    '        End
    '    End If

    '    'On Error Resume Next
    '    'Dim xCID As Long
    '    'Dim xChangedFormat As String

    '    'xCID = GetSystemDefaultLCID()

    '    'xChangedFormat = xFormat

    '    'If xChangedFormat <> "" Then

    '    '    Call SetLocaleInfo(xCID, LOCALE_SSHORTDATE, xChangedFormat)

    '    '    Call PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, 0&)

    '    '    'Call EnumDateFormats(AddressOf theEnumDates, xCID, DATE_SHORTDATE)
    '    'End If

    'End Function


    Public Function fGetUserLocaleInfo(ByVal lLocaleID As Long,
            ByVal lLCType As Long) As String

        Dim sReturn As String
        Dim lReturn As Long
        lReturn = GetLocaleInfo(lLocaleID, lLCType, sReturn, Len(sReturn))
        If lReturn Then

            sReturn = Space$(lReturn)
            If lReturn Then
                fGetUserLocaleInfo = Left$(sReturn, lReturn - 1)
            End If
        End If

    End Function


    Public Function theEnumDates(lDateFormatString As Long) As Long
        theEnumDates = 1
    End Function


    Private Function GetStrFromPointer(sString As Long) As String
        Dim lPos As Long
        Dim sBuffer As String
        sBuffer = Space$(128)
        Call CopyMemory(sBuffer, sString, Len(sBuffer))
        lPos = InStr(sBuffer, Chr(0))

        If lPos Then
            GetStrFromPointer = Left$(sBuffer, lPos - 1)
        End If
    End Function









End Module
