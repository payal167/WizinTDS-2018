Imports System.IO
Public Class Cls_Find
#Const DebugMode = 1
#If DebugMode Then
    Private mlClassDebugID As Long
#End If
    Private rsfind As New DataSet
    'Public cn As VariantType
    Public gid As Integer
    Public mrkgid As String
    Public gdesc As String
    Public clsmstat As String

    Public Sub FillList(findsql As String, lvw As ListView)
        Dim lst As ListViewItem
        On Error GoTo FillListErr
        lvw.Items.Clear()
        rsfind = New DataSet
        rsfind = FetchDataSet(findsql) ', Cnn, adOpenStatic, adLockOptimistic
        If rsfind.Tables(0).Rows.Count = 0 Then ' And rsfind.EOF Then
            Exit Sub
        End If
        With lvw
            '.Columns.Clear()
            '.Columns.Add("Description")
            '.Columns.Add("ID")
            'rsfind.MoveFirst
            If clsmstat = "m" Then
                mrkgid = rsfind.Tables(0).Rows(0)(1).ToString()
                gdesc = rsfind.Tables(0).Rows(0)(0).ToString()
            Else
                gid = rsfind.Tables(0).Rows(0)(1).ToString()
                gdesc = rsfind.Tables(0).Rows(0)(0).ToString()
            End If

            For i = 0 To rsfind.Tables(0).Rows.Count - 1
                lst = .Items.Add(rsfind.Tables(0).Rows(i)(0).ToString())
                lst.SubItems.Add(rsfind.Tables(0).Rows(i)(1))
                'rsfind.MoveNext
            Next 'Loop
        End With
        rsfind = Nothing
        Exit Sub

FillListErr:
        Call RaiseError(MyUnhandledError, "Cls_Find:FillList Method")
    End Sub

End Class
