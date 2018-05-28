Module modErrorHandling
    '

    ' Define your custom errors here.  Be sure to use numbers
    ' greater than 512, to avoid conflicts with OLE error numbers.
    Public Const MyObjectError1 = 1000
    Public Const MyObjectError2 = 1010
    Public Const MyObjectErrorN = 1234
    Public Const MyUnhandledError = 9999

    ' This function will retrieve an error description from a resource
    ' file (.RES).  The ErrorNum is the index of the string
    ' in the resource file.  Called by RaiseError
    Private Function GetErrorTextFromResource(ByVal ErrorNum As Long) As String
        On Error GoTo GetErrorTextFromResourceError
        Dim strMsg As String

        ' get the string from a resource file
        GetErrorTextFromResource = ErrorNum 'loadresstring(ErrorNum)

        Exit Function
GetErrorTextFromResourceError:
        If err.Number <> 0 Then
            GetErrorTextFromResource = "An unknown error has occurred!"
        End If
    End Function

    'There are a number of methods for retrieving the error
    'message.  The following method uses a resource file to
    'retrieve strings indexed by the error number you are
    'raising.
    Public Sub RaiseError(ByVal ErrorNumber As Long, ByVal Source As String)
        Dim strErrorText As String

        strErrorText = GetErrorTextFromResource(ErrorNumber)

        'raise an error back to the client
        MsgBox(vbObjectError + ErrorNumber, Source, strErrorText)
    End Sub



End Module
