Imports System.IO

Module modTools
    Public Const MAX_PATH = 260
    Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

    Public Function GetWinDir() As String
        Dim strBuffer As String
        Dim lngReturn As Long
        Dim strWindowsDirectory As String

        strBuffer = Space$(MAX_PATH)
        lngReturn = GetWindowsDirectory(strBuffer, MAX_PATH)
        strWindowsDirectory = Left$(Trim(strBuffer), Len(Trim(strBuffer)) - 1)
        GetWinDir = Trim(strWindowsDirectory)
    End Function

    Public Function GetWinSysDir() As String
        Dim strBuffer As String
        Dim lngReturn As Long
        Dim strWindowsSystemDirectory As String

        strBuffer = Space$(MAX_PATH)
        lngReturn = GetSystemDirectory(strBuffer, MAX_PATH)
        strWindowsSystemDirectory = Left$(Trim(strBuffer), Len(Trim(strBuffer)) - 1)
        GetWinSysDir = Trim(strWindowsSystemDirectory)
    End Function

    Public Sub ShowCalculator()
        Dim Str As String ', fs As New FileSystemObject
        Str = GetWinSysDir() & "\Calc.Exe"
        If File.Exists(Str) = True Then
            Shell(Str, vbNormalFocus)
        Else
            Str = GetWinDir() & "\Calc.exe"
            If File.Exists(Str) = True Then
                Shell(Str, vbNormalFocus)
            Else
                MsgBox("File not found, make sure Calc.exe is at " & Str, vbExclamation, "Calc.Exe Not Found")
            End If
        End If
        Str = Nothing
    End Sub

    Public Sub OpenNotePad(filename As String)
        'Dim fs As New FileSystemObject
        Dim Str As String
        Str = GetWinSysDir() & "\NotePad.Exe"
        If File.Exists(Str) = True Then
            Shell(Str & " " & filename, vbNormalFocus)
        Else
            Str = GetWinDir() & "\NotePad.exe"
            If File.Exists(Str) = True Then
                Shell(Str & " " & filename, vbNormalFocus)
            Else
                MsgBox("File not found, make sure Notepad.exe is at " & Str, vbExclamation, "Notepad.Exe Not Found")
            End If
        End If
        Str = Nothing
    End Sub

    Public Sub OpenHtmlPage(filename As String)
        'ShellExecute(0&, vbNullString, filename, vbNullString, vbNullString, vbNormalFocus)
        Dim a As Long
        a = ShellExecute(0&, filename, "", "", "open", vbNormalFocus)
    End Sub

    Public Sub OpenEmailClient(emailAddress As String)
        ShellExecute(0&, vbNullString, "mailto:" & emailAddress, vbNullString, vbNullString, vbNormalFocus)
    End Sub


    'Public Function OpenFVUNew(FVUPathName As String, Optional FVUParam As String = vbNullString) As Long

    '    OpenFVUNew = ShellExecute(0&, "Open", "TDS_STANDALONE_FVU.jar", FVUParam, FVUPathName, 3)

    'End Function
    Public Sub OpenFVUNew(FVUPathName As String, Optional FVUParam As String = vbNullString)


        Dim ProcessInfo As New ProcessStartInfo()
        Dim process As New Process
        ProcessInfo = New ProcessStartInfo
        ProcessInfo.FileName = "javaw.exe"
        ProcessInfo.Arguments = "-jar TDS_STANDALONE_FVU.jar " & FVUParam
        ProcessInfo.WorkingDirectory = FVUPathName
        'ProcessInfo.CreateNoWindow = True
        ProcessInfo.UseShellExecute = False
        ProcessInfo.RedirectStandardError = True
        ProcessInfo.RedirectStandardOutput = True
        process = Process.Start(ProcessInfo)

        process.WaitForExit()

        process.Close()
    End Sub

End Module
