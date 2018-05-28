Imports System.IO
Module modTDSUtilieties
    Dim txstrm As StreamWriter
    Dim MainPath As String
    Dim DEDNAME As String
    Public Function AllPANVerified(RecSet As DataSet, Fname As String) As Boolean

        RecSet.Tables(0).Select("PANVerified=False") '= False
        RecSet.Tables(0).DefaultView.Sort = "DNAME"
        If RecSet.Tables(0).Rows.Count > 0 Then
            ' For i = 1 To RecSet.Tables(0).Rows.Count
            AllPANVerified = False
                MainPath = Application.StartupPath
                Dim FILE_NAME As String = MainPath & "\PAN.txt"

                If System.IO.File.Exists(FILE_NAME) = False Then
                    System.IO.File.Create(FILE_NAME).Dispose()
                End If
                Dim txstrm As New System.IO.StreamWriter(FILE_NAME, True)
                txstrm.WriteLine("List of Deductees where PAN is not verified")
            '  Next
            'txstrm.Dispose()
            '  txstrm.WriteLine()

            DEDNAME = RecSet.Tables(0).Rows(0)("DNAME").ToString()
            'WRITE THE FIRST RECORD - DONE TO GET UNIQUE NAMES
            txstrm.WriteLine(RecSet.Tables(0).Rows(0)("DNAME"))
            'RecSet.MoveNext
            For j = 1 To RecSet.Tables(0).Rows.Count
                If RecSet.Tables(0).Rows(0)("DName") <> DEDNAME Then
                    txstrm.WriteLine(RecSet.Tables(0).Rows(0)("DNAME"))
                    DEDNAME = RecSet.Tables(0).Rows(0)("DName")
                End If
                'RecSet.MoveNext
            Next
            txstrm.Close()

        Else
            AllPANVerified = True
        End If
    End Function


    Public Function Registry_Read(Key_Path, Key_Name) As VariantType
        On Error Resume Next
        Dim Registry As Object
        Registry = CreateObject("WScript.Shell")
        Registry_Read = Registry.RegRead(Key_Path & Key_Name)
    End Function


    Public Sub Registry_Write(Key_Path As String, Key_Name As String, Key_Value As VariantType)
        Dim Key_Type As String
        On Error Resume Next

        Dim Registry As Object
        Dim Registry_Value As VariantType
        Registry = CreateObject("WScript.Shell")

        Registry_Value = Registry_Read(Key_Path, Key_Name)

        If Key_Type = "" Then

            'REG_SZ is the default.
            Registry.RegWrite(Key_Path & Key_Name, Key_Value)

        Else

            Registry.RegWrite(Key_Path & Key_Name, Key_Value, Key_Type)

        End If

    End Sub

End Module
