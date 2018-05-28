'Imports System
'Imports System.Collections.Generic
'Imports System.ComponentModel
Imports System.Data
'Imports System.Drawing
'Imports System.Linq
'Imports System.Text
Imports System.Text.RegularExpressions
'Imports System.Windows.Forms
Imports System.IO
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Enable = Microsoft.Office.Interop.Excel.XlEnableSelection
'Imports word = Microsoft.Office.Interop.Word
Public Class Form1
    'Dim wordapp As New word.Application
    'Dim wrddoc As New word.Document
    'Dim cn As New OleDbConnection
    Dim cmd, cmd1 As New OleDbCommand
    Dim da, da1 As New OleDbDataAdapter
    Dim dr1 As OleDbDataReader
    Dim ds, ds1 As New DataSet
    Dim did, i, j, r, c, res, x, chid, comboindex As Integer
    Dim xlApp As New Excel.Application
    Dim xlwb As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim columnNumberZeroBased As Integer
    Public xlrange As Excel.Range
    Dim a, f As Integer
    Dim tdsrate As Double
    Dim dname As String
    Public obj
    Dim stri, stropen As String
    Dim dgcv As New DataGridViewComboBoxColumn
    Dim oRetnMst As New ClsRetnMstObj
    Public rtnid As Integer
    Public oFrmType As String
    Public FNAME As String
    Public fchallan As String
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        Dim ns As Integer
        ComboBox2.Items.Clear()
        Try
            OpenFileDialog1.Multiselect = False
            'OpenFileDialog1.FileName = ""
            'OpenFileDialog1.Filter = "*.xlsx"
            OpenFileDialog1.ShowDialog()

            stropen = OpenFileDialog1.FileName.ToString()
            'File.Open("E:\Sonal\1.mdb", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)
            Process.Start(stropen)


            ' FileOpen(1, stropen, OpenMode.Append, OpenAccess.Write, OpenShare.Shared)
            textBox1.Text = stropen

            'If (CheckBox1.Checked = True) Then
            'OpenFileDialog1.OpenFile()
            'End If

            Dim oexcel As Object
            oexcel = CreateObject("Excel.Application")

            xlwb = xlApp.Workbooks.Open(stropen, , False, , , , , , , True)

            'xlSheet = xlwb.Worksheets("Challan")
            ns = xlwb.Sheets.Count
            For i = 1 To ns
                xlSheet = xlwb.Sheets(i)
                ComboBox2.Items.Add(xlSheet.Name)
            Next
            'If (CheckBox1.Checked = True) Then
            '    OpenFileDialog1.OpenFile()
            'End If
            'MessageBox.Show(xlSheet.Name.ToString())


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'xlSheet = New Excel.Worksheet
        'xlrange = Excel.Range

        ' wrddoc = wordapp.Documents

        'obj = Split(Command, "^")
        'Dim RetnId As Integer

        RetnId = rtnid
        'RetnId = 74
        FNAME = FNAME
        oFrmType = oFrmType '= frmTDS24Q.quter '"24Q"
        If fchallan = "Challan" Then
            comboBox1.SelectedIndex = 0
            comboBox1.Enabled = False
        Else
            comboBox1.SelectedIndex = 1
            comboBox1.Enabled = False
        End If
        Button3.Visible = True
        'cn = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0; Data Source=" & Application.StartupPath & "\Database\WizinTDS.mdb;Persist Security Info=False; JET OLEDB:Database Password=apr01;")
        ''cn = New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0; Data Source=E:\Sonal\DeductMst.accdb;Persist Security Info=False")
        'cn.Open()

        'Dim nds As DataSet
        'Dim Qtr As String
        'nds = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " and FrmType= '24Q" & Qtr & "'")
    End Sub

    Public Function validatePan(ByVal s As String) As Boolean
        Dim x As Integer
        Dim HasNumbers, ValidFormat, IsNumeric As Boolean
        Dim n As String
        '//check for blank first..
        n = s.Substring(5, 4)

        If (s.Length = 10) Then

            ValidFormat = True      '//Valid

        Else

            ValidFormat = False
        End If

        IsNumeric = Integer.TryParse(n, x)

        If (IsNumeric = True) Then
            HasNumbers = True
        Else

            HasNumbers = False
        End If
        'length is ok..now check format..
        '//Check for numbers
        '   //  MessageBox.Show(n);
        If ValidFormat = True And HasNumbers = True Then
            If (Regex.IsMatch(s.Substring(0, 5).ToUpper(), "^[A-Z]+$")) And (Regex.IsMatch(s.Substring(9, 1).ToUpper(), "^[A-Z]+$")) Then
                'if(Regex.IsMatch(s.Substring(0,5).ToUpper(),@"^(A-Z]+$") && Regex.IsMatch(s.Substring(9,1).ToUpper(),@"^(A-Z]+$"))

                ValidFormat = True
            End If
        End If


        '//'there are numbers, it must be in AAAAA9999A format
        '    //for(i = 1; i< s.Length;i++)
        '    //{
        '    //    switch (i) 
        '    //    {
        '    //        case 1: 
        '    //        case 2: 
        '    //        case 3: 
        '    //        case 4: 
        '    //        case 5:
        '    //        case 10: //'Alphabets
        '    //        if(Convert.ToInt32(Encoding.ASCII.GetBytes(s.ToUpper().ElementAt(i).ToString())) >= Convert.ToInt32(Encoding.ASCII.GetBytes("A")) && (Convert.ToInt32(Encoding.ASCII.GetBytes(s.ToUpper().ElementAt(i).ToString())) <= Convert.ToInt32(Encoding.ASCII.GetBytes("Z"))))
        '    //        {
        '    //            ValidFormat = true;
        '    //            break;
        '    //        }
        '    //        else
        '    //        {
        '    //            ValidFormat = false;
        '    //            break;
        '    //        }

        '    //        case 6:
        '    //        case 7: 
        '    //        case 8: 
        '    //        case 9: //'Number s
        '    //        if (int.TryParse(n.ElementAt(i).ToString(),out x) == true)
        '    //        {
        '    //            ValidFormat = true;
        '    //            break;
        '    //        }
        '    //        else
        '    //        {
        '    //            ValidFormat = false;
        '    //            break;
        '    //        }
        '    //    }
        '    //}
        If ValidFormat = True Then

            '//'check the fourth char
            Dim pan As String
            pan = dataGridView2.Rows(i).Cells(1).Value.ToString().ToUpper()
            'Debug.Print(pan)
            Select Case pan.ElementAt(3).ToString()
                'Switch(pan.ElementAt(3).ToString())

                Case "P"
                Case "H"
                Case "C"
                Case "J"
                Case "F"
                Case "A"
                Case "T"
                Case "B"
                Case "L"
                Case "G"
                    ValidFormat = True
                Case Else

                    ValidFormat = False

            End Select
        End If
        If (ValidFormat = True And HasNumbers = True) Then
            Return True

        Else

            Return False  '//'Not proper format.
        End If


    End Function


    Public Function checkcid(ByVal cid As String, ByVal chdate As Date) As Boolean
        Dim sql As String


        sql = "SELECT * FROM Challan" & FNAME & " WHERE TranVouNo=" & cid & " AND DtOfChallan=#" & chdate & "#"

        'If obj(3) = 1 Then
        '    sql = sql & " 4 and 6"
        'ElseIf obj(3) = 2 Then
        '    sql = sql & " 7 and 9"
        'ElseIf obj(3) = 3 Then
        '    sql = sql & " 10 and 12"
        'ElseIf obj(3) = 4 Then
        '    sql = sql & " 1 and 3"
        'End If
        cmd1 = New OleDbCommand(sql, cn)
        da1 = New OleDbDataAdapter(cmd1)
        ds1 = New DataSet()
        da1.Fill(ds1)
        If ds1.Tables(0).Rows.Count > 0 Then
            If (String.IsNullOrEmpty(Convert.ToString(ds1.Tables(0).Rows(0)("ChallanId")))) Then
                chid = ""
                Return False
            Else
                chid = Convert.ToInt32(ds1.Tables(0).Rows(0)("ChallanId"))
                Return True
            End If
        Else
            sql = "SELECT * FROM Challan" & FNAME & " WHERE BankChallanNo=" & cid & " AND DtOfChallan=#" & chdate & "#"
            'If obj(3) = 1 Then
            '    sql = sql & " 4 and 6"
            'ElseIf obj(3) = 2 Then
            '    sql = sql & " 7 and 9"
            'ElseIf obj(3) = 3 Then
            '    sql = sql & " 10 and 12"
            'ElseIf obj(3) = 4 Then
            '    sql = sql & " 1 and 3"
            'End If
            cmd1 = New OleDbCommand(sql, cn)
            da1 = New OleDbDataAdapter(cmd1)
            ds1 = New DataSet()
            da1.Fill(ds1)
            If ds1.Tables(0).Rows.Count > 0 Then
                If (String.IsNullOrEmpty(Convert.ToString(ds1.Tables(0).Rows(0)("ChallanId")))) Then

                    Return False
                Else
                    chid = Convert.ToInt32(ds1.Tables(0).Rows(0)("ChallanId"))
                    Return True
                End If
            Else
                Return False
            End If
        End If

    End Function

    Public Function checkpan(ByVal pan As String) As Boolean
        Dim f As Integer
        'Dim nds As DataSet
        'If (cn.State = ConnectionState.Closed) Then
        '    cn.Open()
        'End If
        cmd1 = New OleDbCommand("SELECT * FROM DeductMst", cn)
        dr1 = cmd1.ExecuteReader()
        If (dr1.HasRows) Then
            cmd1 = New OleDbCommand("SELECT DPan FROM DeductMst WHERE DPan='" & pan & "' AND CoId =" & selectedcoid, cn)
            dr1 = cmd1.ExecuteReader()
            If (dr1.HasRows) Then
                f = 1
            Else
                f = 0
            End If

        Else
            f = 0
        End If
        If (f = 0) Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Function chkBank(ByVal brCode As Integer) As Boolean
        Dim flag As Integer
        Try
            'If brCode.ToString().Length = 7 Then
            '    flag = 1
            'Else
            '    flag = 0
            'End If
            'If flag = 1 Then
            cmd = New OleDbCommand("SELECT BankBrCode FROM BankMst")
            da = New OleDbDataAdapter()
            ds = New DataSet()
            da.SelectCommand = cmd
            da.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                cmd = New OleDbCommand("SELECT BankBrCode FROM BankMst WHERE BankBrCode=" & brCode)
                da = New OleDbDataAdapter()
                ds = New DataSet()
                da.SelectCommand = cmd
                da.Fill(ds)
                If (ds.Tables(0).Rows.Count > 0) Then
                    Return True
                Else
                    Return False
                End If
            End If
            'End If

        Catch ex As Exception
            MessageBox.Show("Please Check BSR Code")
        End Try

    End Function
    Public Function dedupdt() As Integer
        'Try
        '    cmd = New OleDbCommand("UPDATE DeductMst SET DId=,CoID=,DName=,DPan=,DAdd1=,DAdd2=,DAdd3=,DAdd4=,DAdd5=,DState=,DPin=,DType=,F26=,F27=,F24=,DPanRef=,DPANCat=,Category=,PANVerified=,DDesg= WHERE DPAN=")
        'Catch ex As Exception

        'End Try
    End Function
    Public Sub PanNameValidation()
        'Dim cmdchkPan As New OleDbCommand()
        'Dim drchkpan As OleDbDataReader
        'For Me.i = 0 To dataGridView2.Rows.Count - 2
        '    For Me.j = 1 To dataGridView2.Rows.Count - 2
        '        If (dataGridView2.Rows(i).Cells(1).Value.ToString() = dataGridView2.Rows(j).Cells(1).Value.ToString()) Then
        '            If Not (dataGridView2.Rows(i).Cells(0).Value.ToString() = dataGridView2.Rows(j).Cells(0).Value.ToString()) Then
        '                Throw New Exception(dataGridView2.Rows(i).Cells(1).Value.ToString() + "Same PAN Number for different Deductees... Please Check Excel")
        '            End If
        '        End If
        '        If (dataGridView2.Rows(i).Cells(0).Value.ToString() = dataGridView2.Rows(j).Cells(0).Value.ToString()) Then
        '            If Not (dataGridView2.Rows(i).Cells(1).Value.ToString() = dataGridView2.Rows(i).Cells(1).Value.ToString()) Then
        '                Throw New Exception(dataGridView2.Rows(i).Cells(1).Value.ToString() + "This deductee has been entered with a different PAN Number... Please Check Excel")
        '            End If
        '        End If
        '    Next
        '    cmdchkPan = New OleDbCommand("SELECT DName FROM DeductMst WHERE DPan='" & dataGridView2.Rows(i).Cells(1).Value.ToString() & "'", cn)
        '    drchkpan = cmdchkPan.ExecuteReader()
        '    If drchkpan.HasRows Then
        '        If Not dataGridView2.Rows(i).Cells(0).Value.ToString() = drchkpan.GetString(0) Then
        '            Throw New Exception("PAN Already allocated to" & dataGridView2.Rows(i).Cells(0).Value.ToString())
        '        End If
        '    End If
        'Next
    End Sub
    Public Function dedinst() As Integer

        ' Try
        cmd = New OleDbCommand("SELECT Max(DId) FROM DeductMst", cn)
        da = New OleDbDataAdapter()
        ds = New DataSet()
        da.SelectCommand = cmd
        da.Fill(ds)
        If (String.IsNullOrEmpty(ds.Tables(0).Rows(0)(0).ToString()) = True) Then
            a = 1
        Else

            a = Convert.ToInt32(ds.Tables(0).Rows(0)(0).ToString()) + 1
        End If

        stri = "INSERT INTO DeductMst(DId,CoID,DName,DPan,DAdd1,DAdd2,DAdd3,DAdd4,DAdd5,DState,DPin,DType,F26,F27,F24,DPANRef,DPANCat,Category,PANVerified,DDesgn) VALUES(" & a & "," & selectedcoid & ",'" & dataGridView2.Rows(i).Cells(0).Value.ToString() & "','" 'Id,CoId,,Name
        'validatePan(dataGridView2.Rows(i).Cells(1).Value.ToString())

        'If (validatePan(dataGridView2.Rows(i).Cells(1).Value.ToString()) = True) Then

        stri = stri & dataGridView2.Rows(i).Cells(1).Value.ToString() & "','" '//DPAN
        Debug.Print(dataGridView2.Rows(i).Cells(1).Value.ToString())
        'Else

        'MessageBox.Show(dataGridView2.Rows(i).Cells(1).Value.ToString() & "Cannot Accept the PAN Number... Please Check...")

        'Throw New Exception("Invalid PAN... Please Check your Excel data...")

        'End If


        stri = stri & dataGridView2.Rows(i).Cells(2).Value.ToString() & "','" '//Add1
        stri = stri & dataGridView2.Rows(i).Cells(3).Value.ToString() & "','" '//Add2
        stri = stri & dataGridView2.Rows(i).Cells(4).Value.ToString() & "','" '//Add3
        stri = stri & dataGridView2.Rows(i).Cells(5).Value.ToString() & "','" '//Add4
        stri = stri & dataGridView2.Rows(i).Cells(6).Value.ToString() & "'," '//Add5
        If dataGridView2.Rows(i).Cells(7).Value = 0 Then
            stri = stri & "19" & ","
        Else
            stri = stri & dataGridView2.Rows(i).Cells(7).Value.ToString() & "," '//State
        End If

        stri = stri & dataGridView2.Rows(i).Cells(8).Value.ToString() & ",'" '//Pin
        stri = stri & IIf(dataGridView2.Rows(i).Cells(9).Value.ToString() = "0", "O", dataGridView2.Rows(i).Cells(9).Value.ToString()) & "',No,No,No,'" '//Dtype
        ' stri = stri & dataGridView2.Rows(i).Cells(10).Value.ToString() & "',0,'" '//DPANRef.DPANCat 
        stri = stri & IIf(dataGridView2.Rows(i).Cells(10).Value.ToString() = "0", "''", dataGridView2.Rows(i).Cells(10).Value.ToString()) & "',0, '" '//DPANRef.DPANCat    for reference no 0
        stri = stri & IIf(dataGridView2.Rows(i).Cells(11).Value.ToString() = "0", "G", dataGridView2.Rows(i).Cells(11).Value.ToString()) & "',No,'" '//DCategory, PANVer
        stri = stri + dataGridView2.Rows(i).Cells(12).Value.ToString() & "')" '//Desg.
        Debug.Print(stri)
        cmd = New OleDbCommand(stri, cn)
        cmd.ExecuteNonQuery()
        'If (cmd.ExecuteNonQuery()) Then
        '    MessageBox.Show("Records Inserted..")
        'End If

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message & "Insertion failed in DeductMst")
        '    Close()
        'End Try
        Return 1
    End Function


    Private Sub comboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles comboBox1.SelectedIndexChanged
        dataGridView1.Columns.Add(dgcv)
        dgcv.MinimumWidth = 215
        If (dataGridView2.Columns.Count > 0) Then

            dataGridView2.Columns.Clear()
        End If

        'dataGridView1.Columns.Add(dgcv)



        If (dataGridView1.Rows.Count > 0) Then

            dataGridView1.Rows.Clear()
        End If

        If (comboBox1.SelectedIndex = 0) Then

            dataGridView1.Height = 400
            dataGridView1.Width = 435

            For i = 0 To 13
                dataGridView1.Rows.Add()
            Next

            'dataGridView1.Rows(0).Cells(0).Value = "Section                                     "
            dataGridView1.Rows(0).Cells(0).Value = "Tax Amount                              *"
            dataGridView1.Rows(1).Cells(0).Value = "Surcharge"
            dataGridView1.Rows(2).Cells(0).Value = "ECess"
            dataGridView1.Rows(3).Cells(0).Value = "Interest"
            dataGridView1.Rows(4).Cells(0).Value = "Others"
            dataGridView1.Rows(5).Cells(0).Value = "AInterest"
            dataGridView1.Rows(6).Cells(0).Value = "AOthers"
            dataGridView1.Rows(7).Cells(0).Value = "ChqDDNo"
            dataGridView1.Rows(8).Cells(0).Value = "IsBookEntry"
            dataGridView1.Rows(9).Cells(0).Value = "Bank Challan No.                      *"
            dataGridView1.Rows(10).Cells(0).Value = "Tran Vou No."
            dataGridView1.Rows(11).Cells(0).Value = "Bank Br. Code                          *"
            dataGridView1.Rows(12).Cells(0).Value = "Date Of Challan                        *"
            dataGridView1.Rows(13).Cells(0).Value = "Remark"



            'dataGridView2.Columns.Add("col1", "Section")
            dataGridView2.Columns.Add("col1", "Tax Amount")
            dataGridView2.Columns.Add("col2", "Surcharge")
            dataGridView2.Columns.Add("col3", "ECess")
            dataGridView2.Columns.Add("col4", "Interest")
            dataGridView2.Columns.Add("col5", "Others")
            dataGridView2.Columns.Add("col6", "AInterest")
            dataGridView2.Columns.Add("col7", "AOthers")
            dataGridView2.Columns.Add("col8", "ChqDDNo")
            dataGridView2.Columns.Add("col9", "IsBookEntry")
            dataGridView2.Columns.Add("col10", "BankChallanNo")
            dataGridView2.Columns.Add("col11", "Tran Vou No")
            dataGridView2.Columns.Add("col12", "BankBrCode")
            dataGridView2.Columns.Add("col13", "Date Of Challan")
            dataGridView2.Columns.Add("col14", "Remark")
        Else

            dataGridView1.Height = 450
            dataGridView1.Width = 450
            If FNAME = "24Q" Then
                For Me.i = 0 To 24

                    dataGridView1.Rows.Add()
                Next
            Else
                For Me.i = 0 To 25

                    dataGridView1.Rows.Add()
                Next
            End If
            dataGridView1.Rows(0).Cells(0).Value = "Deductee Name                       *"
            dataGridView1.Rows(1).Cells(0).Value = "Deductee PAN                         *"
            dataGridView1.Rows(2).Cells(0).Value = "Address1"
            dataGridView1.Rows(3).Cells(0).Value = "Address2"
            dataGridView1.Rows(4).Cells(0).Value = "Address3"
            dataGridView1.Rows(5).Cells(0).Value = "Address4"
            dataGridView1.Rows(6).Cells(0).Value = "Address5"
            dataGridView1.Rows(7).Cells(0).Value = "Deductee State"
            dataGridView1.Rows(8).Cells(0).Value = "Deductee Pin"
            dataGridView1.Rows(9).Cells(0).Value = "Deductee Type"
            dataGridView1.Rows(10).Cells(0).Value = "Deductee PAN Ref."
            dataGridView1.Rows(11).Cells(0).Value = "Deductee Category"
            dataGridView1.Rows(12).Cells(0).Value = "Deductee Designation"
            dataGridView1.Rows(13).Cells(0).Value = "Section                                     *"
            dataGridView1.Rows(14).Cells(0).Value = "Tax Amount                              *"
            dataGridView1.Rows(15).Cells(0).Value = "Surcharge"
            dataGridView1.Rows(16).Cells(0).Value = "ECess"
            dataGridView1.Rows(17).Cells(0).Value = "Amount Of Payment                 *"
            dataGridView1.Rows(18).Cells(0).Value = "Payment Date                          *"
            dataGridView1.Rows(19).Cells(0).Value = "Total tax Deducted                    *"
            dataGridView1.Rows(20).Cells(0).Value = "Date Of Deduction                    *"
            dataGridView1.Rows(21).Cells(0).Value = "Remark"
            dataGridView1.Rows(22).Cells(0).Value = "Cert No."
            dataGridView1.Rows(23).Cells(0).Value = "Challan Link                              *"
            dataGridView1.Rows(24).Cells(0).Value = "Date Of Challan                        *"
            If Not FNAME = "24Q" Then
                dataGridView1.Rows(25).Cells(0).Value = "Rate Of TDS"
            End If


            dataGridView2.Columns.Add("col1", "Deductee Name")
            dataGridView2.Columns.Add("col2", "Deductee PAN")
            dataGridView2.Columns.Add("col3", "Address1")
            dataGridView2.Columns.Add("col4", "Address2")
            dataGridView2.Columns.Add("col5", "Address3")
            dataGridView2.Columns.Add("col6", "Address4")
            dataGridView2.Columns.Add("col7", "Address5")
            dataGridView2.Columns.Add("col8", "Deductee State")
            dataGridView2.Columns.Add("col9", "Deductee Pin")
            dataGridView2.Columns.Add("col10", "Deductee Type")
            dataGridView2.Columns.Add("col11", "Deductee PAN Ref.")
            dataGridView2.Columns.Add("col12", "Deductee Category")
            dataGridView2.Columns.Add("col13", "Deductee Designation")
            dataGridView2.Columns.Add("col14", "Section")
            dataGridView2.Columns.Add("col15", "Tax Amount")
            dataGridView2.Columns.Add("col16", "Surcharge")
            dataGridView2.Columns.Add("col17", "ECess")
            dataGridView2.Columns.Add("col18", "Amount Of Payment")
            dataGridView2.Columns.Add("col19", "Payment Date ")
            dataGridView2.Columns.Add("col20", "Total tax Deducted")
            dataGridView2.Columns.Add("col21", "Date Of Deduction")
            dataGridView2.Columns.Add("col22", "Remark")
            dataGridView2.Columns.Add("col23", "Cert No.")
            dataGridView2.Columns.Add("col24", "Challan Link")
            dataGridView2.Columns.Add("col26", "Challan Date")
            If Not FNAME = "24Q" Then
                dataGridView2.Columns.Add("col26", "Rate Of TDS")
            End If

        End If


    End Sub
    Public Sub fetch()
        'Try

        Dim rtxt As Integer
        rtxt = Convert.ToInt32(txtRowNum.Text)

        'For Me.i = 0 To dataGridView1.RowCount - 1
        '    dataGridView2.Columns.Add(Convert.ToString(xlrange.Cells(rtxt - 1, i + 1).Value), Convert.ToString(xlrange.Cells(rtxt - 1, i + 1).Value))
        'Next
        'MessageBox.Show(dataGridView2.ColumnCount)
        For Me.i = Convert.ToInt32(txtRowNum.Text) To r

            dataGridView2.Rows.Add()
            For Me.j = 0 To dataGridView1.RowCount - 1
                Dim a As Integer
                a = dataGridView1.Rows(j).Cells(1).Value
                If dataGridView1.Rows(j).Cells(1).Value = 0 Then
                    dataGridView2.Rows(i - Convert.ToInt32(txtRowNum.Text)).Cells(j).Value = 0
                Else
                    If (String.IsNullOrEmpty(Convert.ToString(xlrange.Cells(i, a).Value))) Then
                        dataGridView2.Rows(i - Convert.ToInt32(txtRowNum.Text)).Cells(j).Value = 0
                    Else
                        dataGridView2.Rows(i - Convert.ToInt32(txtRowNum.Text)).Cells(j).Value = Convert.ToString(xlrange.Cells(i, a).Value.ToString())
                    End If
                End If
            Next
            Label5.Text = "Buffering Excel Data..." & i
        Next
        ' MessageBox.Show(i.ToString())
        'Catch
        '    MessageBox.Show("Data not fetched...")
        'End Try
        If comboBox1.SelectedIndex = 1 Then
            For i = 0 To dataGridView2.RowCount - 2
                dname = dataGridView2.Rows(i).Cells(0).Value.ToString()
                For j = 0 To dname.Length - 1
                    If Regex.IsMatch(dname, "[!@,#$%^&*()']") Then
                        dataGridView2.Rows(i).Cells(0).Value = Regex.Replace(dname, "[!@,#$%^&*()']", "_")

                    End If
                    'If (Regex.IsMatch(dataGridView2.Rows(i).Cells(0).Value.ToString(), "[!@#$%^&*()']")) Then
                    '    Regex.Replace(dataGridView2.Rows(i).Cells(0).Value.ToString(), "[!@#$%^&*()',;]", "_")
                    'End If
                Next
            Next
        End If

    End Sub

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        'Try

        '    Dim rtxt As Integer
        '    rtxt = Convert.ToInt32(txtRowNum.Text)

        '    For Me.i = 0 To dataGridView1.RowCount - 1
        '        dataGridView2.Columns.Add(Convert.ToString(xlrange.Cells(rtxt - 1, i + 1).Value), Convert.ToString(xlrange.Cells(rtxt - 1, i + 1).Value))
        '    Next
        '    MessageBox.Show(dataGridView2.ColumnCount)
        '    For Me.i = Convert.ToInt32(txtRowNum.Text) To r

        '        dataGridView2.Rows.Add()
        '        For Me.j = 1 To c

        '            If (String.IsNullOrEmpty(Convert.ToString(xlrange.Cells(i, j).Value))) Then

        '                dataGridView2.Rows(i - Convert.ToInt32(txtRowNum.Text)).Cells(j - 1).Value = 0

        '            Else

        '                dataGridView2.Rows(i - Convert.ToInt32(txtRowNum.Text)).Cells(j - 1).Value = Convert.ToString(xlrange.Cells(i, j).Value)
        '            End If
        '        Next
        '    Next

        'Catch
        '    MessageBox.Show("Data not fetched...")
        'End Try
        button1.Visible = False
        Label5.Visible = True
        fetch()

        If (comboBox1.SelectedIndex = 0) Then

            Dim dateinformat As String

            Dim IsDouble, IsDate, ValidChq As Boolean
            Dim datex As Date
            Dim taxamt, s, ecess, interest, others, Ainterest, Aothers, ttax, dx As Double
            Try
                For Me.i = 0 To dataGridView2.RowCount - 2
                    Try

                        IsDouble = Double.TryParse(dataGridView2.Rows(i).Cells(0).Value.ToString(), dx)
                        If IsDouble Then

                            If dataGridView2.Rows(i).Cells(10).Value.ToString().Length > 9 Then
                                Throw New Exception("Invalid Bank Challan Number" & dataGridView2.Rows(i).Cells(10).Value.ToString())
                                Call showError(i + 1, dataGridView1.Rows(dataGridView2.CurrentCell.ColumnIndex).Cells(1).Value, "Invalid Bank Challan Number" & dataGridView2.Rows(i).Cells(10).Value.ToString())
                                Exit Sub

                            Else
                                IsDate = Date.TryParse(dataGridView2.Rows(i).Cells(12).Value.ToString(), datex)
                                If Not IsDate Then
                                    Throw New Exception("Invalid Challan Date" & dataGridView2.Rows(i).Cells(12).Value.ToString())
                                    Call showError(i + 1, dataGridView1.Rows(dataGridView2.CurrentCell.ColumnIndex).Cells(1).Value, "Invalid Challan Date" & dataGridView2.Rows(i).Cells(12).Value.ToString())
                                    Exit Sub


                                End If
                            End If

                        Else
                            Throw New Exception("Please Enter Valid Tax Amount" & dataGridView2.Rows(i).Cells(0).Value.ToString())
                            Call showError(i + 1, dataGridView1.Rows(dataGridView2.CurrentCell.ColumnIndex).Cells(1).Value, "Please Enter Valid Tax Amount" & dataGridView2.Rows(i).Cells(1).Value.ToString())
                            Exit Sub

                        End If
                    Catch ex As Exception
                        Call showError(i + 1, dataGridView1.Rows(dataGridView2.CurrentCell.ColumnIndex).Cells(1).Value, ex.Message)
                        Exit Sub
                    End Try
                    cmd = New OleDbCommand("SELECT Max(ChallanId) FROM Challan" & FNAME, cn)
                    da = New OleDbDataAdapter()
                    ds = New DataSet()
                    da.SelectCommand = cmd
                    da.Fill(ds)
                    If (String.IsNullOrEmpty(ds.Tables(0).Rows(0)(0).ToString()) = True) Then
                        a = 1
                    Else
                        a = Convert.ToInt32(ds.Tables(0).Rows(0)(0).ToString()) + 1
                    End If

                    taxamt = Convert.ToDouble(dataGridView2.Rows(i).Cells(0).Value.ToString())
                    s = Convert.ToDouble(dataGridView2.Rows(i).Cells(1).Value.ToString())
                    ecess = Convert.ToDouble(dataGridView2.Rows(i).Cells(2).Value.ToString())
                    interest = Convert.ToDouble(dataGridView2.Rows(i).Cells(3).Value.ToString())
                    others = Convert.ToDouble(dataGridView2.Rows(i).Cells(4).Value.ToString())
                    Ainterest = Convert.ToDouble(dataGridView2.Rows(i).Cells(5).Value.ToString())
                    Aothers = Convert.ToDouble(dataGridView2.Rows(i).Cells(6).Value.ToString())

                    ttax = taxamt + s + ecess + interest + others + Ainterest + Aothers




                    stri = "INSERT INTO Challan" & FNAME & "(ChallanID,RetnId,TaxAmt,Surcharge,ECess,Interest,Others,TotalTax,AInterest,AOthers,ChqDDNo,IsBookEntry,BankChallanNo,TranVouNo,BankBrCode,DtOfChallan,Remark,AFees,MinorHead) VALUES(" & a & "," & RetnId & ",'" & dataGridView2.Rows(i).Cells(0).Value.ToString() & "'," '//ChallanID,RetnId,TaxAmt
                    stri = stri & dataGridView2.Rows(i).Cells(1).Value.ToString() & "," '//Surcharge
                    stri = stri & dataGridView2.Rows(i).Cells(2).Value.ToString() & "," '//ECess
                    stri = stri & dataGridView2.Rows(i).Cells(3).Value.ToString() & "," '//Interest
                    stri = stri & dataGridView2.Rows(i).Cells(4).Value.ToString() & "," '//Others
                    stri = stri & ttax & "," '//Total Tax
                    stri = stri & dataGridView2.Rows(i).Cells(5).Value.ToString() & "," '//AInterest
                    'stri = stri & ttax & "," '//Total Tax
                    stri = stri & dataGridView2.Rows(i).Cells(6).Value.ToString() & "," '//AOthers
                    stri = stri & dataGridView2.Rows(i).Cells(7).Value.ToString() & "," '//ChqDDNo,
                    If dataGridView2.Rows(i).Cells(8).Value.ToString() = "N" Or dataGridView2.Rows(i).Cells(8).Value.ToString() = "0" Then
                        stri = stri & "False,'" '//IsBookEntry
                    Else
                        stri = stri & "True,'"
                    End If
                    stri = stri & dataGridView2.Rows(i).Cells(9).Value.ToString() & "'," '//Bank Challan No.
                    If String.IsNullOrEmpty(Convert.ToString(dataGridView2.Rows(i).Cells(10).Value.ToString())) Then

                    Else
                        stri = stri & dataGridView2.Rows(i).Cells(10).Value.ToString() & "," '//TranVou No
                    End If

                    stri = stri & dataGridView2.Rows(i).Cells(11).Value.ToString() & ",#" '//Bank Br Code
                    'If (chkBank(Convert.ToInt32(dataGridView2.Rows(i).Cells(12).Value))) Then
                    dateinformat = Convert.ToDateTime(dataGridView2.Rows(i).Cells(12).Value.ToString()).ToShortDateString()
                    If Convert.ToInt32(Trim(dataGridView2.Rows(i).Cells(12).Value.ToString()).Substring(0, 2)) > 12 Then
                        dateinformat = Date.ParseExact(dateinformat, "dd/MM/yyyy", Nothing)
                    Else
                        dateinformat = Date.ParseExact(dateinformat, "MM/dd/yyyy", Nothing)
                    End If
                    stri = stri & dateinformat & "#,'" '//Date Of Challan
                    'Else
                    ' MessageBox.Show("Please Enter Bank Details in Bank Master for BSR Code" & dataGridView2.Rows(i).Cells(12).Value.ToString(), "BSR Code Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    'Call showError(i + 1, dataGridView1.Rows(dataGridView2.CurrentCell.ColumnIndex).Cells(1).Value, "Please Enter Bank Details in Bank Master for BSR Code" & dataGridView2.Rows(i).Cells(12).Value.ToString())
                    'cmd1 = New OleDbCommand("INSERT INTO BankMst VALUES('" & dataGridView2.Rows(i).Cells(12).Value.ToString() & "','','aaa','Ngp','','NAGPUR','MAHARASHTRA',''," & Convert.ToInt32(obj(0)) & ")", cn)
                    ' cmd1.ExecuteNonQuery()

                    'End If
                    'stri = stri & dataGridView2.Rows(i).Cells(13).Value.ToString() & "','" '//Date Of Challan
                    If (String.IsNullOrEmpty(Convert.ToString(dataGridView2.Rows(i).Cells(13).Value))) Then
                        stri = stri & "',0,'200')"
                    Else
                        stri = stri & dataGridView2.Rows(i).Cells(13).Value.ToString() & "',0,'200')" '//Remark
                    End If
                    cmd = New OleDbCommand(stri, cn)
                    res = cmd.ExecuteNonQuery()



                Next
            Catch ex As Exception

                MessageBox.Show(ex.Message)
            End Try

            If (res > 0) Then

                MessageBox.Show("Challan Import Successfull!!!")
                Label5.Visible = False
                button1.Visible = True
            End If
        End If

        If (comboBox1.SelectedIndex = 1) Then
            Dim IsDouble, IsDate, ValidChq As Boolean
            Dim datex As Date
            Dim taxamt, s, ecess, ttax, dx As Double

            Dim result As Boolean
            'Try

            'Call PanNameValidation()

            For i = 0 To dataGridView2.RowCount - 2
                Try
                    Label5.Text = "Validating Excel Row " & i
                    If Not validatePan(dataGridView2.Rows(i).Cells(1).Value.ToString()) Then   'PAN
                        Throw New Exception("Invalid PAN!!! " & dataGridView2.Rows(i).Cells(1).Value.ToString())
                    End If
                    If dataGridView2.Rows(i).Cells(0).Value.ToString().Length > 75 Then  'NAME OF DEDUCTEE
                        Throw New Exception("Too long Name !!!!" & dataGridView2.Rows(i).Cells(0).Value.ToString())
                    End If

                    If dataGridView2.Rows(i).Cells(2).Value.ToString().Length > 25 Then
                        Throw New Exception("Too long Address !!!!" & dataGridView2.Rows(i).Cells(2).Value.ToString())
                    End If

                    If dataGridView2.Rows(i).Cells(3).Value.ToString().Length > 25 Then
                        Throw New Exception("Too long Address !!!!" & dataGridView2.Rows(i).Cells(3).Value.ToString())
                    End If
                    If dataGridView2.Rows(i).Cells(4).Value.ToString().Length > 25 Then
                        Throw New Exception("Too long Address !!!!" & dataGridView2.Rows(i).Cells(4).Value.ToString())
                    End If
                    If dataGridView2.Rows(i).Cells(5).Value.ToString().Length > 25 Then
                        Throw New Exception("Too long Address !!!!" & dataGridView2.Rows(i).Cells(5).Value.ToString())
                    End If
                    If dataGridView2.Rows(i).Cells(6).Value.ToString().Length > 25 Then
                        Throw New Exception("Too long Address !!!!" & dataGridView2.Rows(i).Cells(6).Value.ToString())
                    End If

                    IsDouble = Double.TryParse(dataGridView2.Rows(i).Cells(14).Value.ToString(), dx) 'AMOUNT OF PAYMENT
                    If IsDouble Then
                        IsDouble = Double.TryParse(dataGridView2.Rows(i).Cells(15).Value.ToString(), dx)
                        If IsDouble Then
                            IsDouble = Double.TryParse(dataGridView2.Rows(i).Cells(16).Value.ToString(), dx)
                            If IsDouble Then
                                IsDouble = Double.TryParse(dataGridView2.Rows(i).Cells(17).Value.ToString(), dx) 'TDS
                                If IsDouble Then
                                    IsDouble = Double.TryParse(dataGridView2.Rows(i).Cells(19).Value.ToString(), dx)  'TOTAL TAX DEPOSITE
                                    If IsDouble Then
                                        IsDate = Date.TryParse(dataGridView2.Rows(i).Cells(18).Value.ToString(), datex)  'DATE OF PAYMENT
                                        If IsDate Then
                                            IsDate = Date.TryParse(dataGridView2.Rows(i).Cells(20).Value.ToString(), datex) 'DATE OF DEDUCTION
                                            If Not IsDate Then
                                                Throw New Exception("Please Enter valid Deduction date" & dataGridView2.Rows(i).Cells(20).Value.ToString())
                                            Else
                                                datex = Convert.ToDateTime(dataGridView2.Rows(i).Cells(20).Value.ToString())  'DATE OF CHALLAN
                                                If datex.Year <= Today.Year Then
                                                    If (Strings.Right(oFrmType, 1) = 1) Then
                                                        If datex.Month > 6 Then
                                                            Throw New Exception("Invalid Quarter")
                                                        End If
                                                    End If

                                                    If (Strings.Right(oFrmType, 1) = 2) Then
                                                        If datex.Month > 9 Then
                                                            Throw New Exception("Invalid Quarter")
                                                        End If
                                                    End If

                                                    If (Strings.Right(oFrmType, 1) = 3) Then
                                                        If datex.Month > 12 Then
                                                            Throw New Exception("Invalid Quarter")
                                                        End If
                                                    End If

                                                    If (Strings.Right(oFrmType, 1) = 4) Then
                                                        If datex.Month > 3 Then
                                                            Throw New Exception("Invalid Quarter")
                                                        End If
                                                    End If
                                                Else
                                                    Throw New Exception("Invalid Date Please Check Year...")
                                                End If

                                            End If
                                        Else
                                            Throw New Exception("Please Enter valid Payment date" & dataGridView2.Rows(i).Cells(18).Value.ToString())
                                        End If
                                    Else
                                        Throw New Exception("Please Enter Valid Payment Total Tax Deducted" & dataGridView2.Rows(i).Cells(19).Value.ToString())

                                    End If
                                Else
                                    Throw New Exception("Please Enter Valid Payment Amount" & dataGridView2.Rows(i).Cells(17).Value.ToString())

                                End If
                            Else
                                Throw New Exception("Please Enter Valid ECess" & dataGridView2.Rows(i).Cells(3).Value.ToString())
                            End If
                        Else

                            Throw New Exception("Please Enter Valid Surcharge" & dataGridView2.Rows(i).Cells(2).Value.ToString())
                        End If
                    Else
                        Throw New Exception("Please Enter Valid Tax Amount" & dataGridView2.Rows(i).Cells(1).Value.ToString())

                    End If
                    If Not String.IsNullOrEmpty(Convert.ToString(dataGridView2.Rows(i).Cells(21).Value)) Then
                        If Not (dataGridView2.Rows(i).Cells(21).Value.ToString() = "Y") Then
                            If Not (dataGridView2.Rows(i).Cells(20).Value = dataGridView2.Rows(i).Cells(18).Value) Then
                                Throw New Exception("Date of Payment Should be equal to Date of deduction..." & dataGridView2.Rows(i).Cells(18).Value.ToString() & "... Check Row Number " & i + Convert.ToInt32(txtRowNum.Text) & " From your Excel")

                            End If
                        End If
                    End If
                Catch ex As Exception

                    'Call showError(i + 1, dataGridView1.Rows(dataGridView2.CurrentCell.ColumnIndex).Cells(1).Value, ex.Message)
                    MessageBox.Show(ex.Message)
                    Exit Sub
                End Try
            Next
            For Me.i = 0 To dataGridView2.RowCount - 2
                Label5.Text = "Importing Record " & i
                result = checkpan(dataGridView2.Rows(i).Cells(1).Value.ToString())
                taxamt = Convert.ToDouble(dataGridView2.Rows(i).Cells(14).Value.ToString())
                s = Convert.ToDouble(dataGridView2.Rows(i).Cells(15).Value.ToString())
                ecess = Convert.ToDouble(dataGridView2.Rows(i).Cells(16).Value.ToString())
                ttax = taxamt + s + ecess

                tdsrate = (Convert.ToDouble(dataGridView2.Rows(i).Cells(14).Value.ToString()) * 100) / Convert.ToDouble(dataGridView2.Rows(i).Cells(17).Value.ToString())


                If (result = True) Then
                    cmd1 = New OleDbCommand("SELECT DId,CoID,DName FROM DeductMst WHERE CoID=" & selectedcoid & " AND DPan ='" & dataGridView2.Rows(i).Cells(1).Value.ToString() & "'", cn)
                    'Debug.Print(dataGridView2.Rows(i).Cells(1).Value.ToString)
                    dr1 = cmd1.ExecuteReader()

                    If (dr1.HasRows) Then

                    Else
                        dedinst()
                        GoTo insrt
                    End If
                    GoTo insrt
                Else
                    dedinst()
                End If


                cmd1 = New OleDbCommand("SELECT DId,CoID,DName FROM DeductMst WHERE CoID=" & selectedcoid & " AND DPan ='" & dataGridView2.Rows(i).Cells(1).Value.ToString() & "'", cn)
                da1 = New OleDbDataAdapter()
                ds1 = New DataSet()
                da1.SelectCommand = cmd1
                da1.Fill(ds1)
                did = Convert.ToInt32(ds1.Tables(0).Rows(0)(0))

                cmd1 = New OleDbCommand("SELECT RetnID,ChallanId,AmtOfPayment,DtOfPayment FROM Deductee" & FNAME & " WHERE RetnId =" & RetnId & " AND ChallanId=" & dataGridView2.Rows(i).Cells(23).Value.ToString() & " AND AmtOfPayment = " & dataGridView2.Rows(i).Cells(17).Value.ToString() & " AND DtOfPayment=#" & Convert.ToDateTime(dataGridView2.Rows(i).Cells(18).Value.ToString()).ToShortDateString() & "#", cn)
                '//MessageBox.Show(dataGridView2.Rows[i].Cells[Convert.ToInt32(dataGridView1.Rows[1].Cells[1].Value) - 1].Value.ToString());
                da1 = New OleDbDataAdapter()
                ds1 = New DataSet()
                da1.SelectCommand = cmd1
                da1.Fill(ds1)
                If (ds1.Tables(0).Rows.Count > 0) Then

                    If (Convert.ToInt32(ds1.Tables(0).Rows(0)("RetnID").ToString()) = RetnId) Then

                        If MessageBox.Show(dataGridView2.Rows(i).Cells(1).Value.ToString() & " Deductee with this PAN Number already exists for same company... Do you want to add this record???", "Duplicate Entry", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                            If MessageBox.Show(dataGridView2.Rows(i).Cells(1).Value.ToString() & " Deductee with this PAN Number already exists for same company... Do you want to Replace Existing record???", "Duplicate Entry", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                                If MessageBox.Show("Are you sure you want to replace existing record???", "Caution", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                                    cmd = New OleDbCommand("DELETE FROM Deductee" & oRetnMst.FrmType & "WHERE RetnId= " & RetnId)
                                    cmd.ExecuteNonQuery()
                                    GoTo insrt
                                Else
                                    GoTo insrt

                                End If
                            Else
                                GoTo insrt
                            End If
                        Else
                            Continue For
                        End If

                    End If

                Else

insrt:
                    Dim dateinformat As String

                    cmd1 = New OleDbCommand("SELECT DId,CoID,DName FROM DeductMst WHERE DPan ='" & dataGridView2.Rows(i).Cells(1).Value.ToString() & "' AND CoId=" & selectedcoid, cn)
                    da1 = New OleDbDataAdapter()
                    ds1 = New DataSet()
                    da1.SelectCommand = cmd1
                    da1.Fill(ds1)
                    did = Convert.ToInt32(ds1.Tables(0).Rows(0)(0))

                    cmd1 = New OleDbCommand("SELECT MAX(ID" & FNAME & ") FROM Deductee" & FNAME, cn)
                    da1 = New OleDbDataAdapter(cmd1)
                    ds1 = New DataSet()
                    da1.Fill(ds1)
                    If (String.IsNullOrEmpty(ds1.Tables(0).Rows(0)(0).ToString())) Then

                        x = 1
                    Else
                        x = Convert.ToInt32(ds1.Tables(0).Rows(0)(0)) + 1
                    End If

                    stri = ""
                    '//f = 1;
                    If FNAME = "24Q" Then
                        stri = "INSERT INTO Deductee" & FNAME & "(Id" & FNAME & ",RetnId,DId,DCode,Sec,AmtOfPayment,DtOfPayment,TaxAmt,Surcharge,ECess,TotalTaxDeposited,TotalTaxDeducted,DtOfDeduction,Remark,ChallanId,CertNo) VALUES(" & x & "," & RetnId & "," & did & ",'" & IIf(dataGridView2.Rows(i).Cells(9).Value.ToString() = 0, "O", dataGridView2.Rows(i).Cells(9).Value.ToString()) & "','" ' //Id,RetnID,DID,DCode
                    ElseIf oFrmType = "27EQ" Then
                        stri = "INSERT INTO Deductee" & FNAME & "(Id" & FNAME & ",RetnId,DId,DCode,Sec,AmtOfPayment,PurchAmt,DtOfPayment,RateOfTDS,TaxAmt,Surcharge,ECess,TotalTaxDeposited,TotalTaxDeducted,DtOfDeduction,Remark,ChallanId,CertNo) VALUES(" & x & "," & RetnId & "," & did & ",'" & IIf(dataGridView2.Rows(i).Cells(9).Value.ToString() = 0, "O", dataGridView2.Rows(i).Cells(9).Value.ToString()) & "','" ' //Id,RetnID,DID,DCode
                    Else
                        stri = "INSERT INTO Deductee" & FNAME & "(Id" & FNAME & ",RetnId,DId,DCode,Sec,AmtOfPayment,DtOfPayment,RateOfTDS,TaxAmt,Surcharge,ECess,TotalTaxDeposited,TotalTaxDeducted,DtOfDeduction,Remark,ChallanId,CertNo) VALUES(" & x & "," & RetnId & "," & did & ",'" & IIf(dataGridView2.Rows(i).Cells(9).Value.ToString() = 0, "O", dataGridView2.Rows(i).Cells(9).Value.ToString()) & "','" ' //Id,RetnID,DID,DCode
                    End If
                    ' stri = stri & dataGridView2.Rows(i).Cells(1).Value.ToString() & "','" '//DPan
                    stri = stri & dataGridView2.Rows(i).Cells(13).Value.ToString() & "'," '// Section
                    stri = stri & dataGridView2.Rows(i).Cells(17).Value.ToString() & ", " '//Payment Amount
                    If FNAME = "27EQ" Then
                        stri = stri & dataGridView2.Rows(i).Cells(17).Value.ToString() & "," '//Purchase Amount
                    End If
                    dateinformat = Convert.ToDateTime(dataGridView2.Rows(i).Cells(18).Value.ToString()).ToShortDateString()
                    If Convert.ToInt32(Trim(dataGridView2.Rows(i).Cells(18).Value.ToString()).Substring(0, 2)) > 12 Then
                        dateinformat = Date.ParseExact(dateinformat, "dd/MM/yyyy", Nothing)
                    Else
                        dateinformat = Date.ParseExact(dateinformat, "MM/dd/yyyy", Nothing)
                    End If
                    stri = stri & "# " & dateinformat & "#," '//Payment Date
                    If Not FNAME = "24Q" Then
                        stri = stri & dataGridView2.Rows(i).Cells(25).Value.ToString() & ","
                    End If

                    stri = stri & dataGridView2.Rows(i).Cells(14).Value.ToString() & "," '//Tax Amount
                    stri = stri & dataGridView2.Rows(i).Cells(15).Value.ToString() & "," '//Surcharge
                    stri = stri & dataGridView2.Rows(i).Cells(16).Value.ToString() & "," '//ECess 
                    stri = stri & ttax & "," '//Total Tax Deposited
                    stri = stri & dataGridView2.Rows(i).Cells(19).Value.ToString() & ",# " '//Total Tax Deducted
                    dateinformat = Convert.ToDateTime(dataGridView2.Rows(i).Cells(20).Value.ToString()).ToShortDateString()
                    If Convert.ToInt32(Trim(dataGridView2.Rows(i).Cells(20).Value.ToString()).Substring(0, 2)) > 12 Then
                        dateinformat = Date.ParseExact(dateinformat, "dd/MM/yyyy", Nothing)
                    Else
                        dateinformat = Date.ParseExact(dateinformat, "MM/dd/yyyy", Nothing)
                    End If
                    stri = stri & dateinformat & "#," '//Date Of Deduction
                    If Convert.ToString(dataGridView2.Rows(i).Cells(21).Value.ToString()) = "0" Then
                        stri = stri & "NULL,"
                    Else
                        stri = stri & "'" & dataGridView2.Rows(i).Cells(21).Value.ToString() & "'," '//Remark
                    End If
                    dateinformat = Convert.ToDateTime(dataGridView2.Rows(i).Cells(24).Value.ToString()).ToShortDateString()
                    If Convert.ToInt32(Trim(dataGridView2.Rows(i).Cells(24).Value.ToString()).Substring(0, 2)) > 12 Then
                        dateinformat = Date.ParseExact(dateinformat, "dd/MM/yyyy", Nothing)
                    Else
                        dateinformat = Date.ParseExact(dateinformat, "MM/dd/yyyy", Nothing)
                    End If
                    '***************comment bcoz challan id not present in challan but need to import deductee********

                    'If (checkcid(dataGridView2.Rows(i).Cells(23).Value.ToString(), dateinformat)) Then
                    '    stri = stri & chid & "," '//Challan Link
                    'Else
                    '    MessageBox.Show(dataGridView2.Rows(i).Cells(1).Value.ToString() & " This PAN not linked to any Challan")
                    '    Continue For
                    'End If
                    If (String.IsNullOrEmpty(dataGridView2.Rows(i).Cells(23).Value.ToString())) Then
                        stri = stri & "" & "," '//Challan Link
                    Else
                        checkcid(dataGridView2.Rows(i).Cells(23).Value.ToString(), dateinformat)
                        stri = stri & chid & ","
                        ' Continue For
                    End If
                    If Convert.ToString(dataGridView2.Rows(i).Cells(22).Value.ToString() = "0") Then
                        stri = stri & "NULL)"
                    Else
                        stri = stri & "'" & dataGridView2.Rows(i).Cells(22).Value.ToString() + "')" '//Cert No.
                    End If

                    'MessageBox.Show(stri)
                    cmd = New OleDbCommand(stri, cn)
                    res = cmd.ExecuteNonQuery()
                End If
                ds1.Dispose()

            Next
            If (res > 0) Then

                MessageBox.Show("Import Successfull!!!")
                'wrddoc.SaveAs("C:\Error File", "docx")
                button1.Visible = True
                Label5.Visible = False
                xlrange = Nothing
                xlSheet = Nothing
                xlwb.Close()
                xlApp.Quit()
            End If
            'Catch ex As Exception
            '    MessageBox.Show("Insertion Failed in DeductTrans" + ex.Message)
            'End Try
        End If
        'Me.Close()

    End Sub

    'Private Sub Form1_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
    '    Me.Close()
    'End Sub

    Private Sub txtRowNum_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRowNum.Leave
        Dim nrows As Integer
        nrows = xlrange.Rows.Count
        Try
            If (String.IsNullOrEmpty(txtRowNum.Text) Or (Convert.ToInt32(txtRowNum.Text) = 0)) Then

                MessageBox.Show("We need a non-zero row number... Please Provide the row number...")
                txtRowNum.Focus()
            End If
            If (Convert.ToInt32(txtRowNum.Text) > nrows) Then
                MessageBox.Show("There are only " & nrows & " Rows in your xl sheet... Please Check...")
                txtRowNum.Focus()
            End If
        Catch ex As FormatException
            MessageBox.Show("We need a non-zero row number... Please Provide the row number...")
            txtRowNum.Focus()
        End Try

    End Sub

    Private Sub dataGridView1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dataGridView1.KeyPress
        If (Char.IsNumber(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub

    Private Sub dataGridView1_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dataGridView1.CellEndEdit

        Try

            If (e.ColumnIndex = 1) Then

                If (Not String.IsNullOrEmpty(Convert.ToString(dataGridView1.CurrentCell.Value))) Then

                    If (Regex.IsMatch(dataGridView1.CurrentCell.Value.ToString(), "^[A-Za-z0]+$")) Then

                        dataGridView1.CurrentCell.Value = ""

                        Throw New Exception("DataType")
                    End If
                End If

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)
        End Try
        'Try

        '    dataGridView1.CurrentCell = dataGridView1.Rows(dataGridView1.CurrentRow.Index - 1).Cells(1)
        'Catch

        '    MessageBox.Show("Please Enter a non zero integer")

        'End Try
    End Sub
    Sub checkgrid()

        f = 1

        'For Me.i = 0 To dataGridView1.Rows.Count - 1
        '    If Convert.ToInt32(dataGridView1.Rows(i).Cells(1).Value) > c Then
        '        MessageBox.Show("There are not " & dataGridView1.Rows(i).Cells(1).Value & "columns in your excel sheet... Please check your column number for " & dataGridView1.Rows(i).Cells(1).Value)
        '        dataGridView1.Rows(i).Cells(1).Value = ""
        '        button1.Enabled = False
        '    End If
        'Next
        For Me.i = 0 To dataGridView1.Rows.Count - 1
            If (String.IsNullOrEmpty(dataGridView1.Rows(i).Cells(1).Value)) Then
                ' f = 1
                If (Convert.ToString(dataGridView1.Rows(i).Cells(0).Value).Contains("*")) Then
                    f = 0
                Else
                    f = 1
                    dataGridView1.Rows(i).Cells(1).Value = 0
                End If
            End If
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        'Dim f As Integer
        'f = 1

        'For i = 0 To dataGridView1.Rows.Count - 1
        '    If (String.IsNullOrEmpty(dataGridView1.Rows(i).Cells(1).Value) = False) Then
        '        If Convert.ToInt32(dataGridView1.Rows(i).Cells(1).Value) > c Then
        '            MessageBox.Show("There are not " & dataGridView1.Rows(i).Cells(1).Value & "columns in your excel sheet... Please check your column number for " & dataGridView1.Rows(i).Cells(0).Value)
        '            dataGridView1.Rows(i).Cells(1).Value = ""
        '            button1.Enabled = False
        '        End If
        '    Else
        '        If (Convert.ToString(dataGridView1.Rows(i).Cells(0).Value).Contains("*")) Then
        '            f = 0
        '            'dataGridView1.Rows(i).Cells(1).Value = 0

        '        Else
        '            dataGridView1.Rows(i).Cells(1).Value = 0
        '        End If
        '    End If

        'Next
        'For Me.i = 0 To dataGridView1.Rows.Count - 1

        'If (String.IsNullOrEmpty(Convert.ToString(dataGridView1.Rows(i).Cells(1).Value)) = True) Then
        'If (Convert.ToString(dataGridView1.Rows(i).Cells(0).Value).Contains("*")) Then
        '    f = 0
        '    'dataGridView1.Rows(i).Cells(1).Value = 0

        'Else
        '    dataGridView1.Rows(i).Cells(1).Value = 0
        'End If
        'End If
        'Next
        button1.Enabled = False
        checkgrid()
        'If (f = 0) Then

        '    MessageBox.Show("Please mention * Fields...")
        '    button1.Enabled = False
        'End If

        For Me.i = 0 To dataGridView1.Rows.Count - 1
            If (String.IsNullOrEmpty(Convert.ToString(dataGridView1.Rows(i).Cells(1).Value)) = False) Then
                button1.Enabled = True
            Else

                'dataGridView1.Focus()
                button1.Enabled = False
                Exit For
            End If
        Next
    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub


    Private Sub txtRowNum_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRowNum.TextChanged

    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectionChangeCommitted
        Dim xlsheets(100) As Excel.Worksheet

        For Me.i = 1 To xlwb.Sheets.Count
            xlsheets(i) = xlwb.Sheets(i)
            If (xlsheets(i).Name = ComboBox2.SelectedItem.ToString()) Then
                xlSheet = xlwb.Sheets(i)
            End If
        Next
        'xlSheet = xlwb.Sheets(ComboBox2.Text)
        xlSheet.Cells.Locked = False
        xlrange = xlSheet.UsedRange
        r = xlSheet.UsedRange.Rows.Count
        c = xlSheet.UsedRange.Columns.Count
        'xlrange.Cells(2, 1).Font.ColorIndex = 8
        'c = xlSheet.Columns.Count
        xlrange.Locked = False


        For i = 1 To c
            dgcv.Items.Add(GetStandardExcelColumnName(i))
        Next

        ' txtRowNum.Focus()
    End Sub
    Public Sub showError(ByVal row As Integer, ByVal column As Integer, ByVal msg As String)
        Process.Start(stropen)
        Try
            'xlrange.Cells(row + 1, column).Interior.Color = Color.Red
            xlrange.Rows(row + 1).Interior.Color = Color.Red

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        xlrange.Cells(row + 1, c + 1).Value = ""


    End Sub
    Public Function GetStandardExcelColumnName(ByVal columnNumberOneBased As Integer) As String
        Dim baseValue As Integer

        Dim ret As String
        baseValue = Convert.ToInt32(Convert.ToChar("A"))
        columnNumberZeroBased = columnNumberOneBased - 1
        ret = ""
        If (columnNumberOneBased > 26) Then

            ret = GetStandardExcelColumnName(columnNumberZeroBased / 26)
        End If
        Return ret + Convert.ToChar(baseValue + (columnNumberZeroBased Mod 26))

    End Function

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkHeader.CheckedChanged
        dgcv.Items.Clear()
        If ChkHeader.Checked = True Then
            'For i = 1 To c
            '    dgcv.Items.Add(xlrange.Cells(1, i).Value)
            'Next
            txtHeader.Visible = True
            txtHeader.Focus()

        Else
            txtHeader.Visible = False
            For i = 1 To c
                dgcv.Items.Add(GetStandardExcelColumnName(i))
            Next
        End If

    End Sub

    Private Sub txtHeader_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHeader.TextChanged
        dgcv.Items.Clear()

    End Sub

    Private Sub txtHeader_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHeader.Leave


        If String.IsNullOrEmpty(txtHeader.Text) Or txtHeader.Text = 0 Then
            MessageBox.Show("Enter a non zero Header Row number")
            txtHeader.Focus()
        Else
            If ChkHeader.Checked = True Then
                For i = 1 To c
                    'If String.IsNullOrEmpty(Convert.ToString(xlrange.Cells(Convert.ToInt32(txtHeader.Text), i).Value) = False) Then
                    dgcv.Items.Add(IIf(String.IsNullOrEmpty(xlrange.Cells(Convert.ToInt32(txtHeader.Text), i).Value), "", xlrange.Cells(Convert.ToInt32(txtHeader.Text), i).Value))
                    ' End If
                Next
            End If
        End If
    End Sub

    Private Sub dataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dataGridView1.CellContentClick
        'If e.ColumnIndex = 3 Then
        '    dataGridView1.CurrentRow.Cells(1).Value = columnNumberZeroBased
        'End If
    End Sub

    Private Sub dataGridView1_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dataGridView1.CellValueChanged
        'If e.ColumnIndex = 3 Then
        '    dataGridView1.CurrentRow.Cells(1).Value = columnNumberZeroBased
        'End If
        If (e.ColumnIndex = 2) Then
            ChkHeader.Enabled = False
        End If
        If e.ColumnIndex = 2 Then
            dataGridView1.CurrentRow.Cells(1).Value = dgcv.Items.IndexOf(dataGridView1.Rows(dataGridView1.CurrentRow.Index).Cells(2).Value) + 1
        End If
    End Sub

    Private Sub dataGridView1_CellLeave(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dataGridView1.CellLeave

    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Process.Start(Application.StartupPath & "\Help.docx")
    End Sub

End Class
