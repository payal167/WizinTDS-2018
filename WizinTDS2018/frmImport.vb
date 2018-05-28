
Public Class frmImport
    Public Cpath As String
    Public Ctitle As String
    Dim icnn As New OleDb.OleDbConnection

    Private Sub cmdok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdok.Click
        Dim reply As Integer, i As Integer, j As Integer, ic As Integer
        Dim rsto As New DataSet
        Dim rsfr As New DataSet
        Dim rsBK As New DataSet
        Dim rsBKc As New DataSet
        Dim sqli As String, sqlfname As String
        Dim sql As String
        Dim coid As Integer, did As Long, pcoid As Integer

        Dim cmd As New OleDb.OleDbCommand

        'On Error GoTo onerr
        'check the license on no of companies that can be used...
        'added by nitin on 06/07/2006.
        Dim cntr As Long
        For i = 0 To lvwImport.Items.Count - 1
            If lvwImport.Items(i).Checked = True Then
                cntr = cntr + 1
            End If
        Next i
        If (frmCoMst.lvwCo.Items.Count + cntr) > NoOfCo Then
            Call MsgBox("Present no. of companies + Companies Selected to be imported; will voilate the license." _
                    & vbCrLf & "Total Available license is " & NoOfCo & " Companies. Total license used is " & frmCoMst.lvwCo.Items.Count _
                    & " Companies." & vbCrLf & "Balance license is " & (NoOfCo - frmCoMst.lvwCo.Items.Count) & " Companies." _
                    & vbCrLf & "Please select lesser no. of companies for import." _
                    , vbCritical, "License Voilation")
            Exit Sub
        End If

        'regular import code...
        ' Dim Itm As ListItem, IsChecked As Boolean
        Dim Itm As ListViewItem, IsChecked As Boolean
        For Each Itm In lvwImport.Items
            If Itm.Checked = True Then
                IsChecked = True
                Exit For
            End If
        Next
        If IsChecked Then
            reply = MsgBox(" Do you want to import the selected companies", vbYesNo, "Confirmation")
        Else
            MsgBox("Nothing selected...")
            Exit Sub
        End If
        If reply = vbNo Then Exit Sub
        ProgressBar1.Visible = True
        'ProgressBar1.Value = 1
        'Cnn.BeginTrans()

        ' Company Master Importal
        coid = 0
        rsto = FetchDataSet("select max(coid) from comst")
        If rsto.Tables(0).Rows.Count > 0 And rsto.Tables(0).Rows(0)(0).ToString() <> "" Then 'And Not String.IsNullOrEmpty(rsto(0).Value)
            coid = IIf(rsto.Tables(0).Rows(0)(0).ToString() = "", 1, Val(rsto.Tables(0).Rows(0)(0)) + 1)
        Else
            coid = 1
        End If
        rsto = Nothing

        rsto = FetchDataSet(" select max(did) from deductmst")
        If rsto.Tables(0).Rows.Count > 0 And rsto.Tables(0).Rows(0)(0).ToString() <> "" Then
            If rsto.Tables(0).Rows(0)(0).ToString() = "" Then
                did = 1
            Else
                did = Val(rsto.Tables(0).Rows(0)(0)) + 1
            End If
            'did = IIf(rsto.Tables(0).Rows(0)(0).ToString() = "", 1, Val(rsto.Tables(0).Rows(0)(0)) + 1)
        Else
            did = 1
        End If
        rsto.Dispose()
        Dim dc As New DataColumn
        sqlfname = ""
        sqli = ""
        'StatusBar1.SimpleText = ""
        cmd.Connection = cn


        For j = 0 To lvwImport.Items.Count - 1
            If lvwImport.Items(j).Checked = True Then
                rsfr = FetchDataSetIMP("select * from comst where coname='" & lvwImport.Items(j).Text & "'")
                If (rsfr.Tables(0).Rows.Count > 0 And rsfr.Tables(0).Rows(0)(0).ToString() <> "") Then
                    pcoid = rsfr.Tables(0).Rows(0)(0)
                    'StatusBar1.SimpleText = ""
                    'StatusBar1.SimpleText = "Importing Data Of " & lvwImport.ListItems(j).Text
                    For Each dc In rsfr.Tables(0).Columns
                        sqlfname = sqlfname & dc.ColumnName & ","
                        If dc.ColumnName = "CoID" Then
                            sqli = sqli & coid & ", "
                        Else
                            Dim typ As TypeCode
                            typ = Type.GetTypeCode(dc.DataType)
                            If Type.GetTypeCode(dc.DataType) = TypeCode.Char Or Type.GetTypeCode(dc.DataType) = TypeCode.String Then
                                sqli = sqli & IIf(rsfr.Tables(0).Rows(0)(dc).ToString() = "", "Null", "'" & rsfr.Tables(0).Rows(0)(dc) & "'") & ","
                            Else
                                sqli = sqli & IIf(rsfr.Tables(0).Rows(0)(dc).ToString() = "", "Null", rsfr.Tables(0).Rows(0)(dc)) & ","
                            End If
                        End If
                    Next
                    sqli = Mid(sqli, 1, Len(sqli) - 1)
                    sqli = " Insert into comst (" & Strings.Left(sqlfname, Len(sqlfname) - 1) & ") values(" & sqli & ")"

                    cmd.CommandText = sqli
                    cmd.ExecuteNonQuery()

                    rsfr = Nothing
                    sqli = vbNullString
                    sqlfname = vbNullString
                    '*********************************************************************************
                    'bank Master Import

                    Dim rsBkChk As New DataSet
                    rsfr = FetchDataSetIMP(" select * from bankmst where coid=" & pcoid)
                    For i = 0 To rsfr.Tables(0).Rows.Count - 1
                        '
                        'StatusBar1.SimpleText = ""
                        'StatusBar1.SimpleText = "Importing Data of " & lvwImport.ListItems(j).Text
                        For Each dc In rsfr.Tables(0).Columns
                            sqlfname = sqlfname & dc.ColumnName & ","
                            If dc.ColumnName = "CoID" Then
                                sqli = sqli & coid & ","
                            Else
                                Dim typ As TypeCode
                                typ = Type.GetTypeCode(dc.DataType)
                                If Type.GetTypeCode(dc.DataType) = TypeCode.Char Or Type.GetTypeCode(dc.DataType) = TypeCode.String Then
                                    sqli = sqli & IIf(rsfr.Tables(0).Rows(i)(dc).ToString() = "", "Null", "'" & rsfr.Tables(0).Rows(i)(dc) & "'") & ","
                                Else
                                    sqli = sqli & IIf(rsfr.Tables(0).Rows(i)(dc).ToString() = "", "Null", rsfr.Tables(0).Rows(i)(dc)) & ","
                                End If
                            End If

                        Next dc
                        'as bank master BSRCode is common for all companies, check if already exist or not
                        'check for duplicate bsr here...
                        'If rsBkChk.State = ADODB.ObjectStateEnum.adStateOpen Then rsBkChk.Close()
                        rsBkChk = FetchDataSet("SELECT * FROM BankMst WHERE BankBrCode = '" & rsfr.Tables(0).Rows(i)("BankBrCode") & "'")
                        If rsBkChk.Tables(0).Rows.Count = 0 Then
                            sqli = Mid(sqli, 1, Len(sqli) - 1)
                            sqli = " Insert into bankmst (" & Strings.Left(sqlfname, Len(sqlfname) - 1) & ") values(" & sqli & ")"
                            cmd.CommandText = sqli

                            cmd.ExecuteNonQuery()


                        End If

                        sqli = vbNullString
                        sqlfname = vbNullString
                    Next

                    rsBK = Nothing
                    sqli = vbNullString
                    sqlfname = vbNullString

                    '*********************************************************************************
                    ' Deductee Master Import
                    rsfr = FetchDataSetIMP(" select * from deductmst where coid=" & pcoid)
                    'ProgressBar1.Value = 1
                    ProgressBar1.Maximum = rsfr.Tables(0).Rows.Count - 1
                    For i = 0 To rsfr.Tables(0).Rows.Count - 1

                        For Each dc In rsfr.Tables(0).Columns
                            sqlfname = sqlfname & dc.ColumnName & ","
                            If dc.ColumnName = "CoID" Then
                                sqli = sqli & coid & ","

                            ElseIf dc.ColumnName = "DId" Then
                                sqli = sqli & did & ","

                            Else
                                Dim typ As TypeCode
                                typ = Type.GetTypeCode(dc.DataType)
                                If Type.GetTypeCode(dc.DataType) = TypeCode.Char Or Type.GetTypeCode(dc.DataType) = TypeCode.String Then
                                    sqli = sqli & IIf(rsfr.Tables(0).Rows(i)(dc).ToString() = "", "Null", "'" & rsfr.Tables(0).Rows(i)(dc) & "'") & ","
                                Else
                                    sqli = sqli & IIf(rsfr.Tables(0).Rows(i)(dc).ToString() = "", "Null", rsfr.Tables(0).Rows(i)(dc)) & ","
                                End If
                            End If
                        Next dc
                        sqli = Mid(sqli, 1, Len(sqli) - 1)

                        sqli = " Insert into deductmst (" & Strings.Left(sqlfname, Len(sqlfname) - 1) & ") values(" & sqli & ")"
                        cmd.CommandText = sqli
                        cmd.ExecuteNonQuery()
                        did = did + 1
                        sqli = vbNullString
                        sqlfname = vbNullString
                        'rsfr.MoveNext()
                        If ProgressBar1.Value < ProgressBar1.Maximum Then
                            ProgressBar1.Value = ProgressBar1.Value + 1
                        End If
                    Next i

                    coid = coid + 1
                End If
                rsfr = Nothing
                sqli = vbNullString
            End If
        Next j

        ProgressBar1.Visible = False
        If MsgBox(" Import Completed Successfully!! Kindly, exit and restart this software.", vbExclamation, "Message") = vbOK Then
            frmCoMst = Nothing
            frmLogin = Nothing
        End If
        'StatusBar1.SimpleText = ""
        'Unload(Me)
        Me.Dispose()
        icnn.Dispose()
        rsto.Dispose()

        rsBKc.Dispose()

        cmd.Dispose()

        Exit Sub
onerr:

        icnn.Dispose()
        rsto.Dispose()

        rsBK.Dispose()
        rsBKc.Dispose()

        cmd.Dispose()
        If Err.Number = -2147467259 Then
            MsgBox("Company already imported !!", vbCritical, "Stop")
        Else
            MsgBox(Err.Description, , Err.Number)
        End If
        ProgressBar1.Visible = False
        Me.Dispose()

    End Sub
    Public Function FetchDataSetIMP(SqlString As String) As DataSet
        Dim QueSt As String
        Dim headadaptor As New OleDb.OleDbDataAdapter
        Dim cmd As New OleDb.OleDbCommand
        Dim ds As New DataSet
        QueSt = SqlString
        cmd = New OleDb.OleDbCommand(QueSt, icnn)
        headadaptor = New OleDb.OleDbDataAdapter
        ds = New DataSet
        headadaptor.SelectCommand = cmd
        headadaptor.Fill(ds)
        headadaptor.Dispose()
        cmd.Dispose()
        Return ds


    End Function
    Private Sub setupListView()
        'add columns to the listview
        lvwImport.Columns.Add("Deductor/Seller Name", 400, HorizontalAlignment.Left)
        lvwImport.Columns.Add("Co ID", 0, HorizontalAlignment.Left)
        'Display listview in details view
        lvwImport.View = View.Details
        'display grid lines
        lvwImport.GridLines = True
        'allow full row selection
        lvwImport.FullRowSelect = True
    End Sub
    Private Sub frmImport_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim sql As String, cmd As OleDb.OleDbCommand, Itm As ListViewItem
        Dim nds As New DataSet, headadaptor As New OleDb.OleDbDataAdapter
        setupListView()
        icnn.Close()
        icnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Cpath & ";Persist Security Info=False;Jet OLEDB:Database Password='apr01'"
        '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\WizinTDS.mdb;Persist Security Info=False;Jet OLEDB:Database Password='apr01'"
        icnn.Open()
        sql = " select CoName,CoID from comst order by coname"

        cmd = New OleDb.OleDbCommand(sql, icnn)
        headadaptor = New OleDb.OleDbDataAdapter
        nds = New DataSet
        headadaptor.SelectCommand = cmd
        headadaptor.Fill(nds)
        headadaptor.Dispose()
        cmd.Dispose()
        If nds.Tables(0).Rows.Count > 0 Then
            For i = 0 To nds.Tables(0).Rows.Count - 1
                Dim newitem As New ListViewItem()
                newitem.Text = nds.Tables(0).Rows(i)(0) 'first column
                newitem.SubItems.Add(nds.Tables(0).Rows(i)(1)) 'second column
                lvwImport.Items.Add(newitem)
            Next
        End If
        'For i = 0 To nds.Tables(0).Rows.Count - 1
        '    Itm = lvwImport.Items.Add(nds.Tables(0).Rows(i)("CoName").ToString())
        'Next
        nds.Dispose()
    End Sub



    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Dispose()
    End Sub

    Private Sub lvwImport_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvwImport.SelectedIndexChanged

    End Sub
End Class