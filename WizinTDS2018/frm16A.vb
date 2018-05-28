Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.IO

Public Class frm16A
    Dim xlapp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    'Dim frmd As New frmMulSelDed
    Dim sqld As String
    Dim WithEvents o16A As New Form16Details
    Dim MainPath, SavePath As String
    Dim frmd As New frmMulSelDed
    'Dim fso As New FileSystemobject


    Private Sub cmdgen_Click(sender As Object, e As EventArgs) Handles cmdgen.Click
        On Error Resume Next
        Dim mFlag As Boolean
        Dim ctr As Integer
        mFlag = CheckBeforeGenerate()
        If mFlag = True Then
            MsgBox("Some fields remain blank." & vbCrLf &
            "Please fill all fields.", 0 + 16, "Data Require...")
            txtSignByName.Focus()
            Exit Sub
        End If
        Dim MultiDeducteeList As Long
        If cmbDeductee.Enabled = True Then
            MultiDeducteeList = cmbDeductee.Text
        Else
            MultiDeducteeList = sqld
        End If
        Dim ds As DataSet
        ds = FetchDataSet("sqld")
        For i = 0 To ds.Tables(0).Rows.Count - 1
            lblMsg.Text = "Generating form 16 for:" & cmbDeductee.Text
            'DeducteeName = MultiDeducteeList(i)
            ' F16ID = o16A.F16ID
            ' UpdateForm16Data(RetnId)
            FillForm16()
            MainPath = Application.StartupPath & "\Form16 Files\"
            SavePath = MainPath & o16A.mCoName & "\"
            Dim logInfo As System.IO.DirectoryInfo
            If Not Directory.Exists(MainPath) Then
                Directory.CreateDirectory(MainPath)
            End If
            If Not Directory.Exists(SavePath) Then
                Directory.CreateDirectory(SavePath)
            End If

            If File.Exists(SavePath & cmbDeductee.Text & ".xls") Then
                File.Delete(SavePath & cmbDeductee.Text & ".xls")
                'DeleteFile.exits(SavePath & cmbDeductee.Text & ".xls")
            End If
            xlBook.SaveAs(SavePath & cmbDeductee.Text & ".xls")
            'clean up....
            xlBook.Close()
        Next
        lblMsg.Text = "Process Finished..."
        xlapp.Quit()
        xlapp = Nothing
        xlBook = Nothing
        xlSheet = Nothing
        'DeducteeName = ""
        If chkOpenXL.Checked = 1 Then
            'open workbook...only if single deductee selection...
            If ds.Tables(0).Rows.Count - 1 <= 0 Then
                Dim XLHandle As Long
                XLHandle = OpenXLFile(SavePath & cmbDeductee.Text & ".xls")
                'AppActivate XLHandle, True
            End If
        End If

        If cmbDeductee.Enabled = False Then
            'open save folder here, when multiple form 16 are generated..
        End If
    End Sub

    Private Function CheckBeforeGenerate() As Boolean
        If Trim(txtSignByName.Text) = vbNullString Then
            CheckBeforeGenerate = True
        ElseIf Trim(txtSignByFatherName.Text) = vbNullString Then
            CheckBeforeGenerate = True
        ElseIf Trim(txtSignByCapacity.Text) = vbNullString Then
            CheckBeforeGenerate = True
        ElseIf Trim(txtPlace.Text) = vbNullString Then
            CheckBeforeGenerate = True
        ElseIf Trim(certidt.Text) = "__/__/__" Then
            CheckBeforeGenerate = True
        ElseIf cmbDeductee.Enabled = False Then
            If sqld = vbNullString Then
                CheckBeforeGenerate = True
            End If
        ElseIf cmbDeductee.SelectedIndex <= -1 Then
            CheckBeforeGenerate = True
        Else
            CheckBeforeGenerate = False
        End If
        Return CheckBeforeGenerate
    End Function

    Private Sub FillForm16()
        Dim rs, rsTmp, rsQrt As New DataSet
        Dim sql2, sql As String
        Dim NoOfRowsInserted As Long
        Dim i, TmpDID As Long
        Dim SqlQurt As String
        Dim mDname As String, TmpCellAdd As Long, AddOfGrossTotInc As Long
        Dim Addof_TaxDeductedAtSourceUs192_2 As Long
        GenFromXl()
        xlapp = New Excel.Application
        xlBook = xlapp.Workbooks.Add
        xlSheet = xlBook.Worksheets("Sheet1")
        xlapp.Visible = True
        xlSheet = xlBook.Sheets("Sheet1")
        xlSheet = xlBook.ActiveSheet
        xlSheet.Name = "Export sheet"
        rs = FetchDataSet("SELECT Form16Details.*, Form16MoreDetails.*, Form16Details.RetnID, Form16Details.DId" &
   " FROM Form16Details INNER JOIN Form16MoreDetails ON Form16Details.F16ID = Form16MoreDetails.F16ID" &
   " WHERE (((Form16Details.RetnID)=" & o16A.RetnID & ") AND ((Form16Details.F16Id)=" & o16A.F16ID & "))")


        TmpDID = o16A.did 'rs.Fields("Form16Details.DID")

        sql2 = "SELECT * from DeductMst WHERE DId=" & TmpDID & " And dname in ('" & cmbDeductee.Text & "')"

        'rstTmp.Open sql2, cnn, adOpenKeyset, adLockReadOnly
        With xlSheet
            'Basic Page Formatting...
            .Range("A1", "F999").Font.Size = 8
            .Range(.Cells(1, 1), .Cells(1, 6)).ColumnWidth = 13.75
            'Start writing form 16 with data...
            .Cells(1, 1) = "FORM NO. 16"
            With .Range(.Cells(1, 1), .Cells(1, 6))
                .Merge()
                .Font.Size = 20
                .Font.Bold = True
                .Font.Underline = True
            End With
            'Line 2
            .Cells(2, 1) = "[(See rule 31(1)(a)]"
            With .Range(.Cells(2, 1), .Cells(2, 6))
                .Merge()
            End With
            'Line 3
            .Cells(3, 1) = "PART A"
            With .Range(.Cells(3, 1), .Cells(3, 6))
                .Merge()
                .Font.Size = 10
                .Font.Bold = True
            End With
            'Line 4
            .Cells(4, 1) = "Certificate under section 203 of the Income Tax Act, 1961 for Tax Deducted at Source on Salary"
            With .Range(.Cells(4, 1), .Cells(4, 6))
                .Merge()
                .Font.Size = 10
                .Font.Bold = True
            End With
            'Line 5
            .Cells(5, 1) = "Name and Addres of the Employer"
            With .Range(.Cells(5, 1), .Cells(5, 3))
                .Merge()
            End With
            .Cells(5, 4) = " Name and Designation of the Employee"
            With .Range(.Cells(5, 4), .Cells(5, 6))
                .Merge()
            End With
            'Line 6
            .Cells(6, 1) = o16A.mCoName
            With .Range(.Cells(6, 1), .Cells(6, 3))
                .Merge()
            End With
            .Cells(6, 4) = cmbDeductee.Text 'rstTmp!DName
            With .Range(.Cells(6, 4), .Cells(6, 6))
                .Merge()
            End With
            mDname = cmbDeductee.Text 'rstTmp!DName
            'Line 7
            .Cells(7, 1) = o16A.mCoAdd1 & " " & o16A.mCoAdd2 & " " & o16A.mCoAdd3
            With .Range(.Cells(7, 1), .Cells(8, 3))
                .Merge()
            End With
            .Cells(7, 4) = o16A.DDesgn 'rst!DDesgn
            With .Range(.Cells(7, 4), .Cells(8, 6))
                .Merge()
            End With
            'Line 8
            .Cells(8, 1) = o16A.mCoAdd4 & " " & o16A.mCoAdd5 & "-" & o16A.mCoPin
            With .Range(.Cells(8, 1), .Cells(8, 3))
                .Merge()
            End With
            'Line 9
            .Cells(9, 1) = " PAN "
            With .Range(.Cells(9, 1), .Cells(9, 2))
                .Merge()
            End With
            .Cells(9, 3) = " TAN "
            With .Range(.Cells(9, 3), .Cells(9, 4))
                .Merge()
            End With
            .Cells(9, 5) = "PAN/GIR NO"
            With .Range(.Cells(9, 5), .Cells(9, 6))
                .Merge()
            End With
            'Line 10
            .Cells(10, 1) = o16A.mCoPAN
            With .Range(.Cells(10, 1), .Cells(10, 2))
                .Merge()
            End With
            .Cells(10, 3) = o16A.mCoTAN
            With .Range(.Cells(10, 3), .Cells(10, 4))
                .Merge()
            End With
            .Cells(10, 5) = o16A.mCoPAN 'rstTmp!DPan
            With .Range(.Cells(10, 5), .Cells(10, 6))
                .Merge()
            End With
            'Line 11
            .Cells(11, 1) = "CIT (TDS)"
            With .Range(.Cells(11, 1), .Cells(11, 3))
                .Merge()
                .Font.Bold = True
            End With
            .Cells(11, 4) = "Period"
            With .Range(.Cells(11, 4), .Cells(11, 5))
                .Font.Bold = True
                .Merge()
            End With
            .Cells(11, 6) = "Assessment Year"
            With .Range(.Cells(11, 6), .Cells(13, 6))
                .Font.Bold = True
                .Merge()
                .WrapText = True
            End With
            'Line 12 & 13
            Dim strsql As String
            Dim headadaptor As New OleDbDataAdapter
            Dim cmd As New OleDbCommand
            Dim ds As New DataSet
            'If rstTmp.State = adStateOpen Then rstTmp.Close
            strsql = "SELECT CoMst.CITTDSAddtess, CoMst.CITTDSCity, CoMst.CITTDSPin From CoMst WHERE CoMst.CoID=" & selectedcoid
            'rstTmp.Open strsql, cnn, adOpenKeyset, adLockReadOnly
            cmd = New OleDbCommand(strsql, cn)
            headadaptor = New OleDbDataAdapter
            ds = New DataSet
            headadaptor.SelectCommand = cmd
            headadaptor.Fill(ds)
            'rst.fill(strsql)
            .Cells(12, 1) = "Address - " & ds.Tables(0).Rows(0)("CITTDSAddtess").ToString() ' ( '& rstTmp(0)
            With .Range(.Cells(12, 1), .Cells(13, 3))
                .Merge()
            End With
            .Cells(12, 4) = "FROM"
            With .Range(.Cells(12, 4), .Cells(13, 4))
                .Merge()
            End With
            With .Range(.Cells(12, 5), .Cells(13, 5))
                .Merge()
            End With
            .Cells(12, 5) = "TO"
            'Line 14
            .Cells(14, 1) = "City - " & ds.Tables(0).Rows(0)("CITTDSCity").ToString 'rstTmp(1)
            .Cells(14, 2) = "Pin Code - "
            .Cells(14, 3) = ds.Tables(0).Rows(0)("CITTDSPin").ToString 'rstTmp(2)
            .Cells(14, 4) = o16A.EmpFromDt 'Format(rst!EmpFromDt, "dd-mmm-yy")
            .Cells(14, 5) = o16A.EmpFromDt 'Format(rst!EmpToDt, "dd-mmm-yy")
            '.Cells(14, 6) = mAYear

            'Create a border around these cells.
            .Range(.Cells(5, 1), .Cells(14, 6)).BorderAround()
            '.Range(.Cells(5, 1), .Cells(14, 6)).Borders(xlInsideHorizontal).Weight = xlThin
            '.Range(.Cells(5, 1), .Cells(14, 6)).Borders(xlInsideVertical).Weight = xlThin
            'Line 15
            .Cells(15, 1) = "Summary of tax deducted at source"
            With .Range(.Cells(15, 1), .Cells(15, 6))
                .Merge()
                .Font.Bold = True
            End With
            'Line 16 & 17
            With .Range(.Cells(16, 1), .Cells(16, 6))
                .WrapText = True
            End With
            .Cells(16, 1) = "Quarter"
            .Cells(16, 2) = "Receipt Numbers of original statements of TDS under sub-section (3) of section 200"
            .Cells(16, 4) = "Amount of tax deducted in respect of the employee"
            .Cells(16, 6) = "Amount of tax deposited remitted in respect of the employee"
            With .Range(.Cells(16, 1), .Cells(17, 1))
                .Merge()
            End With
            With .Range(.Cells(16, 2), .Cells(17, 3))
                .Merge()
                .WrapText = True
                .RowHeight = 25
            End With
            With .Range(.Cells(16, 4), .Cells(17, 5))
                .Merge()
            End With
            With .Range(.Cells(16, 6), .Cells(17, 6))
                .Merge()
            End With
            'Center all the data above this line...
            .Range(.Cells(1, 1), .Cells(17, 6)).HorizontalAlignment = HorizontalAlignment.Center
            'Line 18 to 21
            .Cells(18, 1) = "1st Quarter"
            .Cells(19, 1) = "2nd Quarter"
            .Cells(20, 1) = "3rd Quarter"
            .Cells(21, 1) = "4th Quarter"
            .Range(.Cells(18, 2), .Cells(18, 3)).Merge()
            .Range(.Cells(19, 2), .Cells(19, 3)).Merge()
            .Range(.Cells(20, 2), .Cells(20, 3)).Merge()
            .Range(.Cells(21, 2), .Cells(21, 3)).Merge()
            .Range(.Cells(18, 4), .Cells(18, 5)).Merge()
            .Range(.Cells(19, 4), .Cells(19, 5)).Merge()
            .Range(.Cells(20, 4), .Cells(20, 5)).Merge()
            .Range(.Cells(21, 4), .Cells(21, 5)).Merge()
            'Qtry return PRN data...
            SetTDSRates()
            rsQrt = FetchDataSet(" SELECT r.retnid, r.dtoffiling,r.prn,r.rprn,r.frmtype,r.NewReceiptNo FROM  retnmst AS r, comst AS c  Where  c.coid = r.coid   and c.coid=" & selectedcoid & "  and (#30/Jun/" & Year(FromDate) & "# between #" & Format(o16A.EmpFromDt) & "# and #" & Format(o16A.EmpToDt) & "#" _
                 & " or #" & Format(o16A.EmpToDt) & "# > #30/Jun/" & Year(FromDate) & " #) and r.frmtype='24Q1'" _
                 & " Union All  SELECT r.retnid, r.dtoffiling,r.prn,r.rprn,r.frmtype,r.NewReceiptNo FROM  retnmst AS r, comst AS c  Where  c.coid = r.coid   and c.coid=" & selectedcoid & "  and (#30/09/" & Year(FromDate) & "# between #" & Format(o16A.EmpFromDt) & "# and #" & Format(o16A.EmpToDt) & "#" _
                 & " or #" & Format(o16A.EmpToDt) & "# between #01/07/" & Year(FromDate) & " # and  #30/Sep/" & Year(FromDate) & "#) and r.frmtype='24Q2'" _
                 & " Union All  SELECT r.retnid, r.dtoffiling,r.prn,r.rprn,r.frmtype,r.NewReceiptNo FROM  retnmst AS r, comst AS c  Where  c.coid = r.coid   and c.coid=" & selectedcoid & "  and (#31/12/" & Year(FromDate) & "# between #" & Format(o16A.EmpFromDt) & "# and #" & Format(o16A.EmpToDt) & "#" _
                 & " or #" & Format(o16A.EmpToDt) & "# between #01/10/" & Year(FromDate) & " # and  #31/Dec/" & Year(FromDate) & "#)and r.frmtype='24Q3'" _
                 & " Union All SELECT r.retnid, r.dtoffiling,r.prn,r.rprn,r.frmtype,r.NewReceiptNo FROM  retnmst AS r, comst AS c  Where  c.coid = r.coid   and c.coid=" & selectedcoid & "  and (#31/03/" & Year(ToDate) & "# between #" & Format(o16A.EmpFromDt) & "# and #" & Format(o16A.EmpToDt) & "#" _
                 & " or #" & Format(o16A.EmpToDt) & "# between #01/01/" & Year(ToDate) & " # and  #31/Mar/" & Year(ToDate) & "#) and r.frmtype='24Q4'")
            ' rsQrt.Open SqlQurt, cnn, adOpenStatic, adLockOptimistic
            'While Not rsQrt.EOF
            Dim s1 As Integer
            For s1 = 0 To rsQrt.Tables(0).Rows.Count - 1
                If Not String.IsNullOrEmpty(rsQrt.Tables(0).Rows(s1)("PRN")) Then
                    If Trim(rsQrt.Tables(0).Rows(0)("FrmType")) = 1 Then
                        .Cells(18, 2) = rsQrt.Tables(0).Rows(s1)("NewReceiptNo")
                    ElseIf Trim(rsQrt.Tables(0).Rows("FrmType")(1)) = 2 Then
                        .Cells(19, 2) = rsQrt.Tables(0).Rows(s1)("NewReceiptNo")
                    ElseIf Trim(rsQrt.Tables(0).Rows("FrmType")(1)) = 3 Then
                        .Cells(20, 2) = rsQrt.Tables(0).Rows(s1)("NewReceiptNo")
                    ElseIf Trim(rsQrt.Tables(0).Rows("FrmType")(1)) = 4 Then
                        .Cells(21, 2) = rsQrt.Tables(0).Rows(s1)("NewReceiptNo")
                    End If
                End If
                'rsQrt.MoveNext
                'Wend
            Next
            rsQrt.Dispose()

            ' Qtrly TDS deducted & deposited
            rsQrt = FetchDataSet("Select DeductMst.DName, Sum(Deductee24Q.TotalTaxDeposited) As SumOfTotalTaxDeposited, Sum(Deductee24Q.TotalTaxDeducted) As SumOfTotalTaxDeducted, Form16Details.F16ID, 'Q1' AS FrmType FROM ((DeductMst INNER JOIN Deductee24Q ON DeductMst.DId = Deductee24Q.DId) INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId) LEFT JOIN RetnMst ON Deductee24Q.RetnID = RetnMst.RetnID" _
                & " Where (((Deductee24Q.dtofDeduction)" & " between #01/Apr/" & Year(FromDate) & " # and  #30/Jun/" & Year(FromDate) & "#))GROUP BY DeductMst.DName, Form16Details.F16ID, DeductMst.CoID Having (((Form16Details.F16ID) =" & o16A.F16ID & ") And ((DeductMst.CoID) = " & selectedcoid & " ))" _
                 & " Union All" _
                 & " SELECT DeductMst.DName, Sum(Deductee24Q.TotalTaxDeposited) AS SumOfTotalTaxDeposited, Sum(Deductee24Q.TotalTaxDeducted) AS SumOfTotalTaxDeducted, Form16Details.F16ID, 'Q2' AS FrmType FROM ((DeductMst INNER JOIN Deductee24Q ON DeductMst.DId = Deductee24Q.DId) INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId) LEFT JOIN RetnMst ON Deductee24Q.RetnID = RetnMst.RetnID" _
                 & " WHERE (((Deductee24Q.DtOfDeduction)" & " between #01/Jul/" & Year(FromDate) & " # and  #30/Sept/" & Year(FromDate) & "#))GROUP BY DeductMst.DName, Form16Details.F16ID, DeductMst.CoID Having (((Form16Details.F16ID) =" & o16A.F16ID & ") And ((DeductMst.CoID) = " & selectedcoid & " ))" _
                 & " Union All" _
                 & " SELECT DeductMst.DName, Sum(Deductee24Q.TotalTaxDeposited) AS SumOfTotalTaxDeposited, Sum(Deductee24Q.TotalTaxDeducted) AS SumOfTotalTaxDeducted, Form16Details.F16ID, 'Q3' AS FrmType FROM ((DeductMst INNER JOIN Deductee24Q ON DeductMst.DId = Deductee24Q.DId) INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId) LEFT JOIN RetnMst ON Deductee24Q.RetnID = RetnMst.RetnID" _
                 & " WHERE (((Deductee24Q.DtOfDeduction)" & " between #01/Oct/" & Year(FromDate) & " # and  #31/Dec/" & Year(FromDate) & "#))GROUP BY DeductMst.DName, Form16Details.F16ID, DeductMst.CoID Having (((Form16Details.F16ID) =" & o16A.F16ID & ") And ((DeductMst.CoID) = " & selectedcoid & " ))" _
                 & " Union All" _
                 & " SELECT DeductMst.DName, Sum(Deductee24Q.TotalTaxDeposited) AS SumOfTotalTaxDeposited, Sum(Deductee24Q.TotalTaxDeducted) AS SumOfTotalTaxDeducted, Form16Details.F16ID, 'Q4' AS FrmType FROM ((DeductMst INNER JOIN Deductee24Q ON DeductMst.DId = Deductee24Q.DId) INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId) LEFT JOIN RetnMst ON Deductee24Q.RetnID = RetnMst.RetnID" _
                 & " WHERE (((Deductee24Q.DtOfDeduction)" & " between #01/Jan/" & Year(ToDate) & " # and  #31/Mar/" & Year(ToDate) & "#)) GROUP BY DeductMst.DName, Form16Details.F16ID, DeductMst.CoID Having (((Form16Details.F16ID) =" & o16A.F16ID & ") And ((DeductMst.CoID) = " & selectedcoid & " ))" _
                 & " Union All" _
                 & " SELECT DeductMst.DName, Sum([F16Challan].[taxamt]+[f16challan].[surcharge]+[f16challan].[ecess]) AS TotalTaxDeposited, Sum([F16Challan].[taxamt]+[f16challan].[surcharge]+[f16challan].[ecess])AS TotalTaxDeducted, Form16Details.F16ID, 'Q1' AS FrmType FROM RetnMst INNER JOIN ((DeductMst INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId) INNER JOIN F16Challan ON Form16Details.F16ID = F16Challan.F16ID) ON RetnMst.RetnID = Form16Details.RetnID" _
                 & " WHERE (((F16Challan.DtOfChallan)" & " between #01/Apr/" & Year(FromDate) & " # and  #30/Jun/" & Year(FromDate) & "#)) GROUP BY DeductMst.DName, Form16Details.F16ID, DeductMst.CoID Having (((Form16Details.F16ID) =" & o16A.F16ID & ") And ((DeductMst.CoID) = " & selectedcoid & "))" _
                 & " Union All" _
                 & " SELECT DeductMst.DName, Sum([F16Challan].[taxamt]+[f16challan].[surcharge]+[f16challan].[ecess]) AS TotalTaxDeposited, Sum([F16Challan].[taxamt]+[f16challan].[surcharge]+[f16challan].[ecess]) AS TotalTaxDeducted, Form16Details.F16ID, 'Q2' AS FrmType FROM RetnMst INNER JOIN ((DeductMst INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId) INNER JOIN F16Challan ON Form16Details.F16ID = F16Challan.F16ID) ON RetnMst.RetnID = Form16Details.RetnID" _
                 & " WHERE (((F16Challan.DtOfChallan)" & " between #01/jul/" & Year(FromDate) & " # and  #30/sept/" & Year(FromDate) & "#)) GROUP BY DeductMst.DName, Form16Details.F16ID, DeductMst.CoID Having (((Form16Details.F16ID) = " & o16A.F16ID & ") And ((DeductMst.CoID) = " & selectedcoid & "))" _
                 & " Union All" _
                 & " SELECT DeductMst.DName, Sum([F16Challan].[taxamt]+[f16challan].[surcharge]+[f16challan].[ecess]) AS TotalTaxDeposited, Sum([F16Challan].[taxamt]+[f16challan].[surcharge]+[f16challan].[ecess]) AS TotalTaxDeducted, Form16Details.F16ID, 'Q3' AS FrmType FROM RetnMst INNER JOIN ((DeductMst INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId) INNER JOIN F16Challan ON Form16Details.F16ID = F16Challan.F16ID) ON RetnMst.RetnID = Form16Details.RetnID" _
                 & " WHERE (((F16Challan.DtOfChallan)" & " between #01/oct/" & Year(FromDate) & " # and  #31/dec/" & Year(FromDate) & "#)) GROUP BY DeductMst.DName, Form16Details.F16ID, DeductMst.CoID Having (((Form16Details.F16ID) = " & o16A.F16ID & ") And ((DeductMst.CoID) = " & selectedcoid & "))" _
                 & " Union All" _
                 & " SELECT DeductMst.DName, Sum([F16Challan].[taxamt]+[f16challan].[surcharge]+[f16challan].[ecess]) AS TotalTaxDeposited, Sum([F16Challan].[taxamt]+[f16challan].[surcharge]+[f16challan].[ecess]) AS TotalTaxDeducted, Form16Details.F16ID, 'Q4' AS FrmType FROM RetnMst INNER JOIN ((DeductMst INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId) INNER JOIN F16Challan ON Form16Details.F16ID = F16Challan.F16ID) ON RetnMst.RetnID = Form16Details.RetnID" _
                 & " WHERE (((F16Challan.DtOfChallan)" & " between #01/jan/" & Year(ToDate) & " # and  #31/Mar/" & Year(ToDate) & "#)) GROUP BY DeductMst.DName, Form16Details.F16ID, DeductMst.CoID Having (((Form16Details.F16ID) = " & o16A.F16ID & ") And ((DeductMst.CoID) = " & selectedcoid & "))")
            '            rsQrt.Open SqlQurt, cnn, adOpenStatic, adLockOptimistic
            'While Not rsQrt.EOF
            Dim s2 As Integer
            For s2 = 0 To rsQrt.Tables(0).Rows.Count - 1
                If Not String.IsNullOrEmpty(rsQrt.Tables(0).Rows(s2)("SumOfTotalTaxDeducted")) Then
                    If rsQrt.Tables(0).Rows(s1)("FrmType")(1) = 1 Then
                        .Cells(18, 4) = .Cells(18, 4) + rsQrt.Tables(0).Rows(s2)("SumOfTotalTaxDeducted")
                        .Range(.Cells(18, 4), .Cells(18, 5)).Merge()
                        .Cells(18, 6) = .Cells(18, 6) + rsQrt.Tables(0).Rows(s2)("SumOfTotalTaxDeducted")
                    ElseIf rsQrt.Tables(0).Rows(s1)("FrmType")(1) = 2 Then
                        .Cells(19, 4) = .Cells(19, 4) + rsQrt.Tables(0).Rows(s2)("SumOfTotalTaxDeducted")
                        .Range(.Cells(19, 4), .Cells(19, 5)).Merge()
                        .Cells(19, 6) = .Cells(19, 6) + rsQrt.Tables(0).Rows(s2)("SumOfTotalTaxDeducted")
                    ElseIf rsQrt.Tables(0).Rows(s1)("FrmType")(1) = 3 Then
                        .Cells(20, 4) = .Cells(20, 4) + rsQrt.Tables(0).Rows(s2)("SumOfTotalTaxDeducted")
                        .Range(.Cells(20, 4), .Cells(20, 5)).Merge()
                        .Cells(20, 6) = .Cells(20, 6) + rsQrt.Tables(0).Rows(s2)("SumOfTotalTaxDeducted")
                    ElseIf rsQrt.Tables(0).Rows(s1)("FrmType")(1) = 4 Then
                        .Cells(21, 4) = .Cells(21, 4) + rsQrt.Tables(0).Rows(s2)("SumOfTotalTaxDeducted")
                        .Range(.Cells(21, 4), .Cells(21, 5)).Merge()
                        .Cells(21, 6) = .Cells(21, 6) + rsQrt.Tables(0).Rows(s2)("SumOfTotalTaxDeducted")
                    End If
                End If
            Next
            'rsQrt.MoveNext
            'Wend
            ' rsQrt.Close
            ' Set rsQrt = Nothing
            rsQrt.Dispose()
            'Line 22
            .Cells(22, 1) = "TOTAL:"
            .Cells(22, 4) = "=SUM(D18:D21)"
            .Cells(22, 6) = "=SUM(F18:F21)"
            .Range(.Cells(22, 1), .Cells(22, 3)).Merge()
            .Range(.Cells(22, 4), .Cells(22, 5)).Merge()

            'Create a border around these cells.
            .Range(.Cells(16, 1), .Cells(22, 6)).BorderAround()
            ' .Range(.Cells(16, 1), .Cells(22, 6)).Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 
            ' .Range(.Cells(16, 1), .Cells(22, 6)).Borders(xlInsideVertical).Weight = xlThin

            'Line 23
            .Cells(23, 1) = "PART B"
            With .Range(.Cells(23, 1), .Cells(23, 6))
                .Merge()
                .Font.Size = 10
                .Font.Bold = True
                .HorizontalAlignment = HorizontalAlignment.Center
            End With
            'Line 24
            .Cells(24, 1) = "DETAILS OF SALARY PAID AND ANY OTHER INCOME AND TAX DEDUCTED"
            With .Range(.Cells(24, 1), .Cells(24, 6))
                .Merge()
                .Font.Size = 10
                .Font.Bold = True
                .HorizontalAlignment = HorizontalAlignment.Center
            End With
            'Line 25
            .Cells(25, 1) = "1. Gross Salary"
            With .Range(.Cells(25, 1), .Cells(25, 3))
                .Merge()
                .Font.Bold = True
            End With
            'Line 26
            .Cells(26, 1) = "a. Salary as per provisions contained in section 17(1)"
            With .Range(.Cells(26, 1), .Cells(26, 3))
                .Merge()
            End With
            .Cells(26, 4) = o16A.Gross1
            'Line 27
            .Cells(27, 1) = "b. Value of perquisites u/s 17(2)."
            With .Range(.Cells(27, 1), .Cells(27, 3))
                .Merge()
            End With
            .Cells(27, 4) = o16A.Gross2
            'Line 28
            .Cells(28, 1) = "(as per Form No. 12BB, whereever applicable."
            With .Range(.Cells(28, 1), .Cells(28, 3))
                .Merge()
            End With
            'Line 29
            .Cells(29, 1) = "c. Profits in lieu of salary u/s 17(3)"
            With .Range(.Cells(29, 1), .Cells(29, 3))
                .Merge()
            End With
            .Cells(29, 4) = Format(o16A.Gross3, "#0.00")
            'Line 30
            .Cells(30, 1) = "(as per Form No. 12BB, whereever applicable."
            With .Range(.Cells(30, 1), .Cells(30, 3))
                .Merge()
            End With
            'Line 31
            .Cells(31, 1) = "d. Total: including Previous Employer Salary"
            With .Range(.Cells(31, 1), .Cells(31, 3))
                .Merge()
                .Font.Bold = True
            End With
            .Cells(31, 4) = o16A.TotalSalaryPreEmp
            Dim rspr As New DataSet
            Dim sqlpr As String
            rspr = FetchDataSet("SELECT Form16Details.TotalSalaryPreEmp, Form16Details.TDSAmtPreEmp From Form16Details WHERE Form16Details.F16ID=" & o16A.F16ID)

            .Cells(31, 5) = "=sum(d26:d29)+" & Val(o16A.TotalSalaryPreEmp)
            'End With
            ''Fill Allowances and other Incomes
            '' Call FillAllowanceOtherIncome
            Dim mRow As Integer
            mRow = 32

            With xlSheet
                .Cells(mRow, 1) = "2. Less: Allowances to the extent exempt u/s 10"
                With .Range(.Cells(mRow, 1), .Cells(mRow, 3))
                    .Merge()
                    .Font.Bold = True
                End With
                mRow = mRow + 1
                .Cells(mRow, 1) = "Allowances"
                With .Range(.Cells(mRow, 1), .Cells(mRow, 2))
                    .Merge()
                    .Font.Bold = True
                End With
                .Cells(mRow, 3) = "Rs."
                With .Range(.Cells(mRow, 3), .Cells(mRow, 3))
                    .Font.Bold = True
                End With
                mRow = mRow + 1
                'allowances details...
                Dim StartRow As Integer, BalanceRow As Integer, BalanceRow1 As Integer

                StartRow = mRow
                'rst.Filter = "TypeOfDetail='A'"
                'rst.MoveFirst
                'Do While Not rst.EOF
                '    .Cells(mRow, 1) = rst!Particulars
                '    .Cells(mRow, 3) = rst!GrossAmt
                '    mRow = mRow + 1
                '    rst.MoveNext
                'Loop
                mRow = mRow + 1
                .Cells(mRow, 1) = "TOTAL:"
                .Cells(mRow, 3) = "=SUM(C" & StartRow & ":c" & mRow - 1 & ")"
                .Cells(mRow, 5) = "=C" & mRow
                mRow = mRow + 1
                .Cells(mRow, 1) = "3. Balance (1-2)"
                .Cells(mRow, 5) = "=E31-E" & mRow - 1
                BalanceRow = mRow
                'rst.Filter = ""
                mRow = mRow + 1
                .Cells(mRow, 1) = "4. Deductions " : mRow = mRow + 1
                .Cells(mRow, 1) = "(a) Entertainment Allowance"
                .Cells(mRow, 3) = o16A.Sec16ii : mRow = mRow + 1
                .Cells(mRow, 1) = "(b) Tax on Employment"
                .Cells(mRow, 3) = o16A.Sec16iii : mRow = mRow + 1
                .Cells(mRow, 1) = "5. Aggregate of 4(a) and 4(b)"
                .Cells(mRow, 5) = "=SUM(C" & mRow - 2 & ":C" & mRow - 1 & ")" : mRow = mRow + 1
                .Cells(mRow, 1) = "6. Income Chargable under the Head Salaries (3-5)"
                .Cells(mRow, 6) = "=E" & BalanceRow & "-E" & mRow - 1
                BalanceRow = mRow
                mRow = mRow + 1
                .Cells(mRow, 1) = "Income"
                With .Range(.Cells(mRow, 1), .Cells(mRow, 2))
                    .Merge()
                    .Font.Bold = True
                End With
                .Cells(mRow, 3) = "Rs."
                With .Range(.Cells(mRow, 3), .Cells(mRow, 3))
                    .Font.Bold = True
                End With
                mRow = mRow + 1
                'Other Income...
                StartRow = mRow
                'rst.Filter = "TypeOfDetail='O'"
                'rst.MoveFirst
                'Do While Not rst.EOF
                '    .Cells(mRow, 1) = rst!Particulars
                '    .Cells(mRow, 3) = rst!GrossAmt
                '    mRow = mRow + 1
                '    rst.MoveNext
                'Loop
                mRow = mRow + 1
                .Cells(mRow, 1) = "TOTAL:"
                .Cells(mRow, 3) = "=SUM(C" & StartRow & ":c" & mRow - 1 & ")"
                .Cells(mRow, 6) = "=C" & mRow
                mRow = mRow + 1
                .Cells(mRow, 1) = "8. Gross Total Income (6+7)"
                .Cells(mRow, 6) = "=F" & BalanceRow & "+F" & mRow - 1
                BalanceRow = mRow
                'rst.Filter = "" : mRow = mRow + 1
                .Cells(mRow, 1) = "9. Deductions under Chapter VI A" : mRow = mRow + 1
                .Cells(mRow, 1) = "A. Sections 80C, 80CCC & 80CCD" : mRow = mRow + 1

                .Cells(mRow, 4) = "Gross Amount"
                With .Range(.Cells(mRow, 4), .Cells(mRow, 5))
                    .Font.Bold = True
                End With
                .Cells(mRow, 5) = "Deductible Amount"
                mRow = mRow + 1
                '80C...
                StartRow = mRow
                'rst.Filter = "TypeOfDetail='E'"
                'rst.MoveFirst
                'Do While Not rst.EOF
                '    .Cells(mRow, 1) = rst!Particulars
                '    .Cells(mRow, 4) = rst!GrossAmt
                '    .Cells(mRow, 5) = rst!DeductibleAmt
                '    mRow = mRow + 1
                '    rst.MoveNext
                'Loop
                mRow = mRow + 1
                .Cells(mRow, 1) = "TOTAL:"
                .Cells(mRow, 4) = "=SUM(D" & StartRow & ":D" & mRow - 1 & ")"
                .Cells(mRow, 5) = "=SUM(E" & StartRow & ":E" & mRow - 1 & ")"
                .Cells(mRow, 6) = "=E" & mRow
                BalanceRow1 = mRow
                mRow = mRow + 2
                .Cells(mRow, 1) = "Note :1. Aggregate amount deductible under section 80C, 80CCC & 80CCD(1) shall not exceed one lakh rupees " : mRow = mRow + 1
                'commented on dt 11/05/14 for ver 5.0
                '             .Cells(mRow, 1) = "shall not exceed one lakh rupees other sections (e.g., 80E, 80G etc.)": mRow = mRow + 1
                '             .Cells(mRow, 1) = "2. aggregate amount deductible under the three sections , i.e., 80C, 80CCC and 80CCD, ": mRow = mRow + 1

                'Chapter VI A Other...............
                .Cells(mRow, 1) = "Other sections (e.g., 80E, 80G, 80TTA, etc.) under chapter VI A" : mRow = mRow + 1
                .Cells(mRow, 3) = "Gross Amount"
                With .Range(.Cells(mRow, 3), .Cells(mRow, 5))
                    .Font.Bold = True
                End With
                .Cells(mRow, 4) = "Qualifying Amt"
                .Cells(mRow, 5) = "Deductible Amount"
                mRow = mRow + 1
                StartRow = mRow
                ' rst.Filter = "TypeOfDetail='V' OR TypeOfDetail='F' or TypeOfDetail='G'"
                'rst.MoveFirst
                'Do While Not rst.EOF
                '    .Cells(mRow, 1) = rst!Particulars
                '    .Cells(mRow, 3) = rst!GrossAmt
                '    .Cells(mRow, 4) = rst!QualifyAmt
                '    .Cells(mRow, 5) = rst!DeductibleAmt
                '    mRow = mRow + 1
                '    rst.MoveNext
                'Loop
                'rst.Filter = ""
                mRow = mRow + 1
                .Cells(mRow, 1) = "TOTAL:"
                .Cells(mRow, 3) = "=SUM(C" & StartRow & ":C" & mRow - 1 & ")"
                .Cells(mRow, 4) = "=SUM(D" & StartRow & ":D" & mRow - 1 & ")"
                .Cells(mRow, 5) = "=SUM(E" & StartRow & ":E" & mRow - 1 & ")"
                .Cells(mRow, 6) = "=E" & mRow
                mRow = mRow + 1
                .Cells(mRow, 1) = "10. Aggregate of deductible amounts under Chapter VI-A (a+b+c)"
                .Cells(mRow, 6) = "=F" & BalanceRow1 & "+F" & mRow - 1 : mRow = mRow + 1
                .Cells(mRow, 1) = "11. Total income (8—10)"
                .Cells(mRow, 6) = "=F" & BalanceRow & "-F" & mRow - 1 : mRow = mRow + 1
                .Cells(mRow, 1) = "12. Tax on total income Rs."
                .Cells(mRow, 6) = o16A.TaxAmt : mRow = mRow + 1
                .Cells(mRow, 1) = "13. Surcharge (on tax computed at S. No. 12) Rs."
                .Cells(mRow, 6) = o16A.Surcharge : mRow = mRow + 1
                .Cells(mRow, 1) = "14. Education Cess @3% "
                .Cells(mRow, 6) = o16A.ECess : mRow = mRow + 1
                .Cells(mRow, 1) = "15. Tax payable (12+13+14) Rs."
                .Cells(mRow, 6) = "=SUM(F" & mRow - 3 & ":F" & mRow - 1 & ")" : mRow = mRow + 1
                .Cells(mRow, 1) = "16. Relief u/s 89 (attach details)"
                .Cells(mRow, 6) = o16A.Relief89 : mRow = mRow + 1
                .Cells(mRow, 1) = "17. Tax Payable"
                .Cells(mRow, 6) = "=F" & mRow - 2 & " - F" & mRow - 1 : mRow = mRow + 1
                .Cells(mRow, 1) = "18. Less : (a) Tax deducted at source u/s 192(1)incluing Tds by Previous Employer "
                .Cells(mRow, 5) = "=F22-E" & mRow + 2 & "+" & Val(o16A.TDSAmtPreEmp) : mRow = mRow + 1
                .Cells(mRow, 1) = "(b) Tax paid by the employer on behalf of the employee " ': mRow = mRow + 1
                .Cells(mRow, 5) = o16A.TDSOnPerks : mRow = mRow + 1
                .Cells(mRow, 1) = "u/s 192(1A) on perquisites u/s 17(2)"
                Dim sqlperk As String
                Dim rsperk As New DataSet
                .Cells(mRow, 6) = "=sum(E" & mRow - 2 & ":E" & mRow : mRow = mRow + 1
                .Cells(mRow, 1) = "Tax Payble /Refundable (17-18)"
                .Cells(mRow, 6) = "=f" & mRow - 4 & "-f" & mRow - 1
                'Format cells
                .Range(.Cells(25, 1), .Cells(mRow, 6)).BorderAround()
                '.Range(.Cells(25, 3), .Cells(mRow, 6)).Borders(xlInsideVertical).Weight = xlThin
                mRow = mRow + 1
                BalanceRow = mRow
                .Cells(mRow, 1) = "VERIFICATION"
                With .Range(.Cells(mRow, 1), .Cells(mRow, 6))
                    .Merge()
                    .Font.Bold = True
                    .HorizontalAlignment = HorizontalAlignment.Center
                End With
                mRow = mRow + 1
                Dim DeclRow As Integer
                DeclRow = mRow
                ' Format this row, data will be filled after challan details...
                With .Range(.Cells(mRow, 1), .Cells(mRow, 6))
                    .Merge()
                    .WrapText = True
                    .RowHeight = 56
                    .HorizontalAlignment = HorizontalAlignment.Center
                End With
                mRow = mRow + 2
                .Cells(mRow, 1) = "Place : " & txtPlace.Text : mRow = mRow + 1
                .Cells(mRow, 1) = "Date : " & Format(certidt.Text) : mRow = mRow + 1
                .Cells(mRow, 4) = "Signature of the person responsible for deduction of tax" : mRow = mRow + 1
                .Cells(mRow, 4) = "Full Name   :" & txtSignByName.Text : mRow = mRow + 1
                .Cells(mRow, 4) = "Designation :" & txtSignByCapacity.Text
                .Range(.Cells(BalanceRow, 1), .Cells(mRow, 6)).BorderAround()
                mRow = mRow + 2
                If rspr.Tables(0).Rows.Count > 0 And rspr.Tables(0).Rows.Count <> 0 Or rspr.Tables(0).Rows.Count <> vbNull Then
                    .Cells(mRow, 1) = "The above figures includes : " : mRow = mRow + 1
                    .Cells(mRow, 1) = "Salary from previous employer(S) Rs. " & rspr.Tables(0).Rows(0)("TotalSalaryPreEmp") : mRow = mRow + 1
                    .Cells(mRow, 1) = "TDS by previous employer(S) Rs. " & rspr.Tables(0).Rows(0)("TDSAmtPreEmp") : mRow = mRow + 1
                End If

                'now the challan details...
                .HPageBreaks.Add(.Cells(mRow, 1))
                mRow = mRow + 1
                .Cells(mRow, 1) = "=" & Chr(34) & "Name: " & Chr(34) & " & d6"
                mRow = mRow + 1
                BalanceRow = mRow

                If o16A.mCoStatus = "A" Or o16A.mCoStatus = "S" Or o16A.mCoStatus = "D" Or o16A.mCoStatus = "E" Or o16A.mCoStatus = "G" Or
                    o16A.mCoStatus = "H" Or o16A.mCoStatus = "L" Or o16A.mCoStatus = "N" Then
                    'This is a Govt Co. - Show Annexure A only..
                    .Cells(mRow, 1) = "ANNEXURE - A"
                    .Range(.Cells(mRow, 1), .Cells(mRow, 5)).Merge()
                    mRow = mRow + 1
                    .Cells(mRow, 1) = "DETAILS OF TAX DEDUCTED AND DEPOSITED IN THE CENTRAL GOVERNMENT ACCOUNT"
                    .Range(.Cells(mRow, 1), .Cells(mRow, 5)).Merge()
                    mRow = mRow + 1
                    .Cells(mRow, 1) = "THROUGH BOOK ENTRY"
                    .Range(.Cells(mRow, 1), .Cells(mRow, 5)).Merge()
                    mRow = mRow + 1
                    .Cells(mRow, 1) = "(The Employer to provide payment wise details of tax deducted and deposited with respect to the employees)"
                    .Range(.Cells(mRow, 1), .Cells(mRow, 5)).RowHeight = 25
                    .Range(.Cells(mRow, 1), .Cells(mRow, 5)).Merge()
                    mRow = mRow + 1
                    .Cells(mRow, 1) = "Sr.No."
                    .Range(.Cells(mRow, 1), .Cells(mRow + 1, 1)).Merge() : .Range(.Cells(mRow, 2), .Cells(mRow + 1, 2)).Merge()
                    .Cells(mRow, 2) = "Tax Deposited in respect of Employees (Rs)."
                    '                    .Range(Cells(mRow, 2), Cells(mRow + 1, 2)).Merge
                    .Cells(mRow, 3) = "Book identification number (BIN)"
                    .Range(.Cells(mRow, 3), .Cells(mRow, 5)).Merge()
                    mRow = mRow + 1
                    .Cells(mRow, 3) = "Receipt No. of form No. 24G"
                    .Cells(mRow, 4) = "DDO Sequence No. in the Book Adjustment Mini Statement"
                    .Cells(mRow, 5) = "Date on which tax deposited (dd/mm/yyyy)"
                Else
                    'This is a non-Govt Co. - Show Annexure B only..
                    .Cells(mRow, 1) = "ANNEXURE - B"
                    .Range(.Cells(mRow, 1), .Cells(mRow, 5)).Merge()
                    mRow = mRow + 1
                    .Cells(mRow, 1) = "DETAILS OF TAX DEDUCTED AND DEPOSITED IN THE CENTRAL GOVERNMENT ACCOUNT"
                    .Range(.Cells(mRow, 1), .Cells(mRow, 5)).Merge()
                    mRow = mRow + 1
                    .Cells(mRow, 1) = "(The Employer to provide payment wise details of tax deducted and deposited with respect to the employees)"
                    .Range(.Cells(mRow, 1), .Cells(mRow, 5)).RowHeight = 25
                    .Range(.Cells(mRow, 1), .Cells(mRow, 5)).Merge()
                    mRow = mRow + 1
                    .Cells(mRow, 1) = "Sr.No."
                    .Range(.Cells(mRow, 1), .Cells(mRow + 1, 1)).Merge()
                    .Cells(mRow, 2) = "Tax Deposited in respect of Employees (Rs)."
                    .Range(.Cells(mRow, 2), .Cells(mRow + 1, 2)).Merge()      '
                    .Cells(mRow, 3) = "Challan identification number (CIN)"
                    .Range(.Cells(mRow, 3), .Cells(mRow, 5)).Merge()     '
                    mRow = mRow + 1
                    '.Range(Cells(mRow - 1, 2), Cells(mRow, 2)).Merge
                    .Cells(mRow, 3) = "BSR code of the Bank Branch"
                    .Cells(mRow, 4) = "Date on which tax deposited (dd/mm/yyyy)"
                    .Cells(mRow, 5) = "Challan Serial Number"
                End If
                With .Range(.Cells(BalanceRow, 1), .Cells(mRow, 5))
                    .HorizontalAlignment = HorizontalAlignment.Center
                    .WrapText = True
                    .Font.Bold = True
                End With
                mRow = mRow + 1
                Dim CRow As Long, TotalTDSFromChallan As Double
                Dim rstTmp As DataSet
                'If rstTmp.State = adStateOpen Then rstTmp.Close

                rstTmp = FetchDataSet("SELECT D24.RetnID, D24.DId, D24.TaxAmt, D24.Surcharge, D24.ECess, C24.ChqDDNo, C24.BankBrCode," &
                 " C24.DtOfChallan, C24.BankChallanNo FROM Challan24Q AS C24 INNER JOIN Deductee24Q AS D24 ON " &
                 " C24.ChallanID = D24.ChallanId WHERE (((D24.DId)=" & TmpDID & ")) " &
                 " Union All select 0 as expr1, 0 as expr2, F.TaxAmt, F.Surcharge, F.ECess, F.ChqDDNo, " &
                 " F.BankBrCode, F.DtOfChallan, F.BankChallanNo FROM F16Challan as F where F.F16ID=" & F16ID & " Order by DtOfChallan")

                'rstTmp.Open strsql, cnn, adOpenKeyset, adLockReadOnly

                Dim cRowStartsFrom As Long, SrNoCtr As Long
                Dim mCol As Integer
                Addof_TaxDeductedAtSourceUs192_2 = mRow
                cRowStartsFrom = mRow
                TotalTDSFromChallan = 0 : SrNoCtr = 1
                Dim s As Integer
                For s = 0 To rstTmp.Tables(0).Rows.Count - 1
                    .Cells(mRow, 1) = SrNoCtr
                    .Cells(mRow, 2) = rstTmp.Tables(0).Rows(s)("TaxAmt") + rstTmp.Tables(0).Rows(s)("Surcharge") + rstTmp.Tables(0).Rows(s)("ECess")
                    .Cells(mRow, 3) = rstTmp.Tables(0).Rows(s)("BankBrCode") & ""
                    .Cells(mRow, 4) = IIf(o16A.mCoStatus = "A" Or o16A.mCoStatus = "S" Or o16A.mCoStatus = "D" Or o16A.mCoStatus = "E" Or o16A.mCoStatus = "G" Or
                    o16A.mCoStatus = "H" Or o16A.mCoStatus = "L" Or o16A.mCoStatus = "N", "", Format(rstTmp.Tables(0).Rows(s)("DtOfChallan")))
                    .Cells(mRow, 5) = IIf(o16A.mCoStatus = "A" Or o16A.mCoStatus = "S" Or o16A.mCoStatus = "D" Or o16A.mCoStatus = "E" Or o16A.mCoStatus = "G" Or
                    o16A.mCoStatus = "H" Or o16A.mCoStatus = "L" Or o16A.mCoStatus = "N", Format(rstTmp.Tables(0).Rows(s)("DtOfChallan")), rstTmp.Tables(0).Rows(s)("BankChallanNo"))
                    mRow = mRow + 1
                    SrNoCtr = SrNoCtr + 1
                    TotalTDSFromChallan = TotalTDSFromChallan + rstTmp.Tables(0).Rows(s)("TaxAmt") + rstTmp.Tables(0).Rows(s)("Surcharge") + rstTmp.Tables(0).Rows(s)("ECess")
                    'rstTmp.Move
                Next
                'Loop
                .Cells(mRow, 1) = "TOTAL:"
                .Cells(mRow, 2) = IIf(TotalTDSFromChallan = 0, 0, "=SUM(B" & cRowStartsFrom & ":B" & mRow - 1 & ")")
                .Range(.Cells(BalanceRow, 1), .Cells(mRow, 5)).BorderAround()
                ' .Range(.Cells(BalanceRow, 1), .Cells(mRow, 5)).Borders(xlInsideHorizontal).Weight = xlThin
                ' .Range(.Cells(BalanceRow, 1), .Cells(mRow, 5)).Borders(xlInsideVertical).Weight = xlThin
                'Declaration above challan details but filled here...
                .Cells(DeclRow, 1) = "I " & txtSignByName.Text & " son/daughter of " & txtSignByFatherName.Text & " working in the capacity of " &
             txtSignByCapacity.Text & " (Desig.) do hereby certify that a sum of " & TotalTDSFromChallan & " (in words) " & SpellRupee(TotalTDSFromChallan) & " has been deducted at " &
             "source and paid to the credit of the Central Government. I further certify that the information given above is true and correct based on the books of account, documents and other available records."
            End With
        End With
    End Sub

    Private Sub cmdexit_Click(sender As Object, e As EventArgs) Handles cmdexit.Click
        Close()
    End Sub

    Private Sub cmdc_Click(sender As Object, e As EventArgs) Handles cmdc.Click
        Dim sql As String
        sql = (" SELECT DeductMst.*, Form16Details.* FROM DeductMst INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId  WHERE Form16Details.RetnID =" & RetnId)
        cmbDeductee.Items.Clear()
        cmbDeductee.SelectedIndex = -1
        frmd.sql1 = sql
        frmd.strflg = True
        frmd.Show()
        sqld = frmd.sql1
        If sqld = "" Then
            cmbDeductee.Enabled = True
            chkOpenXL.Enabled = True
        Else
            cmbDeductee.Enabled = False
            chkOpenXL.Enabled = False
        End If
        frmd.Hide()
    End Sub

    Private Sub cmbDeductee_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDeductee.SelectedIndexChanged
        If cmbDeductee.SelectedIndex >= 0 And cmbDeductee.Text <> "System.Data.DataRowView" Then
            o16A = o16A.Fetch(cmbDeductee.Text)
            txtSignByName.Text = o16A.PR24Name & ""
            txtSignByFatherName.Text = o16A.SignByFatherName & ""
            txtSignByCapacity.Text = o16A.PR24Desg & ""
            txtPlace.Text = o16A.PlaceOfForm
            certidt.Text = o16A.DateOfForm
            '.Text = oBank.State
            'cmdSave.Text = "&Save"
        Else
            'ClearBankCtrls()
        End If
    End Sub
    'Dim rs As New DataSet
    '    ' ReDim arrF16ID(0)
    '    cmbDeductee.Items.Clear()
    '    rs = FetchDataSet("SELECT * FROM FORM16DETAILS WHERE DID=" & cmbDeductee.Items(cmbDeductee.SelectedIndex).ToString)
    '    If cmbDeductee.Text <> "System.Data.DataRowView" And cmbDeductee.Text <> "" Then
    '        'ShowData(cmbDeductee.Text)
    '        'txtDName.Text = cmbDeductee.Text
    '    End If
    '    '        While Not rs.EOF
    '    '            txtSignByFatherName.Text = IIf(IsNull(rs!SignByFatherName), vbNullString, rs!SignByFatherName)
    '    '            txtPlace.Text = IIf(IsNull(rs!place), vbNullString, rs!place)
    '    '            certidt.Text = Format(IIf(IsNull(rs!DateOfCertificate), Date, rs!DateOfCertificate), "dd/mm/yy")
    '    '            RetnId = IIf(IsNull(rs!RetnId), vbNullString, rs!RetnId)
    '    '            F16ID = IIf(IsNull(rs!F16ID), vbNullString, rs!F16ID)
    '    '            arrF16ID(0) = F16ID
    '    '            Exit Sub
    '    'Wend

    '    'If rs.State = 1 Then rs.Close
    '    'Set rs = Nothing
    'End Sub



    Private Sub frm16A_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Main()
        o16A = New Form16Details
        filldeductee()
        'Dim DM As New frm16A
        'Dim ds, ds1 As New DataSet
        'ds = FetchDataSet("SELECT * FROM CoMst WHERE CoID=" & selectedcoid)
        'If ds.Tables(0).Rows.Count > 0 Then
        '    With DM
        '        txtSignByName.Text = ds.Tables(0).Rows(0)("PR24Name") 'rst!PR24Name
        '        txtSignByCapacity.Text = ds.Tables(0).Rows(0)("PR24Desg") 'rst!PR24Desg

        '    End With
        'Else
        'End If
        'ds.Dispose()
        'ds1 = FetchDataSet("SELECT SignByFatherName,Place, DateOfCertificate, F16ID " _
        '       & "FROM FORM16DETAILS WHERE RETNID=" & RetnId)

        ''Do While Not ds.Tables(0).Rows(0).co
        ''    txtSignByFatherName.Text = 'IIf(IsNull(rs!SignByFatherName), vbNullString, rs!SignByFatherName)
        ''    txtPlace.Text = IIf(IsNull(rs!place), vbNullString, rs!place)
        ''    certidt.Text = Format(IIf(IsNull(rs!DateOfCertificate), Date, rs!DateOfCertificate), "dd/mm/yy")
        ''    'F16ID = rs!F16ID
        ''    Exit Do
        ''Loop
        ''If rs.State = 1 Then rs.Close
        ''Set rs = Nothing

        'txtSignByName.TabIndex = 0
        'txtSignByName.TabStop = True
        'txtSignByFatherName.TabStop = True
        'txtSignByCapacity.TabStop = True
        'txtPlace.TabStop = True
        'certidt.TabStop = True
    End Sub

    Private Sub filldeductee()
        selectedcoid = 4
        'Dim rs As New DataSet
        'cmbDeductee.Items.Clear()
        'rs = FetchDataSet("SELECT DeductMst.*, Form16Details.* FROM DeductMst INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId where DeductMst.CoId=" & selectedcoid)
        'cmbDeductee.DataSource = rs.Tables(0)
        'cmbDeductee.DisplayMember = "DName"
        ''cmbDeductee.ValueMember = "DId"
        Dim nds As New DataSet
        Dim Sql As String = "SELECT DeductMst.*, Form16Details.* FROM DeductMst INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId where DeductMst.CoId=" & selectedcoid
        cmbDeductee.DataSource = Nothing
        nds = FetchDataSet(Sql)
        If nds.Tables(0).Rows.Count > 0 Then
            cmbDeductee.DataSource = nds.Tables(0)
            'cmbDeductee.ValueMember = "DId"
            cmbDeductee.DisplayMember = "DName"
        End If
        cmbDeductee.SelectedIndex = 0
        nds.Dispose()
    End Sub

    Private Sub UpdateForm16Data(mRetID As Long)
        Dim rs As New DataSet
        Dim mSql As String


        mSql = "UPDATE FORM16DETAILS SET SignByName='"

        mSql = "UPDATE FORM16DETAILS SET SignByFatherName='" & txtSignByFatherName.Text & "', " _
               & "Place='" & txtPlace.Text & "', " _
               & "DateOfCertificate=#" & Format(certidt.Text) & "# " _
               & " WHERE RETNID=" & o16A.RetnID
        'rs.Open mSql, cnn, adOpenStatic, adLockOptimistic
        'If rs.State = 1 Then rs.Close
        rs = FetchDataSet("mSql")
        rs.Dispose()
    End Sub

    Private Sub GenFromXl()
        Dim m As Integer, i As Integer
        On Error GoTo ErrHandler
        xlapp = Nothing : xlBook = Nothing : xlSheet = Nothing
        xlapp = New Excel.Application : xlBook = xlapp.Workbooks.Add : xlSheet = xlBook.Worksheets("sheet1")
        Exit Sub
ErrHandler:
        MsgBox(Err.Description, 0 + 16, "Error...")
    End Sub

    Public Function OpenXLFile(XLFileName As String) As Long
        OpenXLFile = Shell(0&, "Open", XLFileName, vbNullString)
    End Function

    Private Sub chkOpenXL_CheckedChanged(sender As Object, e As EventArgs) Handles chkOpenXL.CheckedChanged

    End Sub
End Class