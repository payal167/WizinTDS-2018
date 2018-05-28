Imports System.IO
Imports System.Reflection

Module mod27EQConvert

    'Dim fso As New FileSystemObject
    Dim eTDSPath As String, eFileName As String
    'Public CPan As Boolean
    Dim TStrm As IO.StreamWriter
    Dim oCoMst As New clsCoMst
    Dim Value18 As Date
    ', TStrm As TextStream
    Public Sub Convert27EQ(frmid As String, filename As String, revised As String, oldRRRNo As Double, TANApplNo As Double)
        Dim rstChln As New DataSet, rstDeductee As New DataSet
        Dim rstRetn As New DataSet, rstCSum As New DataSet, rstDSum As New DataSet
        Dim rstC As New DataSet, rstD As New DataSet
        Dim rstChlnDed As New DataSet
        Dim LNo As Long, RecNo As Long, RecNoDed As Long
        Dim SumC As Double ', SumD As Double
        Dim FH(18), BH(70), CD(41), DD(43)
        Dim strFH As String, strBH As String, strCD As String, strDD As String
        Dim i As Long
        Dim ChlnTotalinDeductee As Double
        oCoMst = oCoMst.FetchCo(selectedcoid)
        ', fs As New FileSystemObject
        'Check whether the return exists..
        rstRetn = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " AND FrmType = '" & frmid & "'") ', Cnn
        If rstRetn.Tables(0).Rows.Count <= 0 Then
            Call MsgBox("There is no data for this return form.  Kindly create" _
            & vbCrLf & "this return by using Data entry option and then" _
            & vbCrLf & "create e-TDS file using this option." _
            , vbExclamation + vbDefaultButton1, "RETURN NOT FOUND")
            GoTo cleanup
            Exit Sub
        End If

        oCoMst = oCoMst.FetchCo(selectedcoid)
        'Start the conversion
        'Open the text file..
        eTDSPath = Application.StartupPath & "\e-TDS Files"
        eFileName = eTDSPath & "\" & oCoMst.CoName & "\F" & frmid & "V" & ".txt"
        '  TStrm = fs.CreateTextFile(filename, True)
        'Dim TStrm As New System.IO.StreamWriter(eFileName, True)
        TStrm = File.CreateText(eFileName)
        'Get the related data..
        rstC = FetchDataSet("SELECT * FROM Challan27EQ WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn      'Challan Records
        rstD = FetchDataSet("SELECT DT.*, DM.* FROM Deductee27EQ AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" &
      rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn   'Deductee Records
        rstCSum = FetchDataSet("SELECT sum(TaxAmt) as TotAmt,sum(Surcharge) as TotSc, sum(Ecess) as TotEcess," &
               " Sum(Interest) as TotInt, sum(Others) as TotOth, sum(AFees) as TotFees FROM Challan27EQ WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn      'Challan Records
        'If oCoMst.chkpan = True Then
        If AllPANVerified(rstD, Left(filename, Len(filename) - 5) & "_PAN_VERIFY.CSV") = False Then
            If MsgBox("There are unverified PANs in the return." & vbCrLf &
        "Do you want to take risk of converting the TDS file?", vbYesNo + vbDefaultButton2, "UNVERIFIED PAN EXISTS") = vbNo Then
                IsAllPANVerified = False
                TStrm.Close()
                Call OpenNotePad(Left(filename, Len(filename) - 5) & "_PAN_VERIFY.CSV")
                GoTo cleanup
            Else
                IsAllPANVerified = True
            End If
        Else
            IsAllPANVerified = True
        End If
        'End If
        'START WRITING THE TEXT FILE NOW.
        '*****************************************************************************************
        'FILE HEADER RECORD...
        FH(1) = 1
        FH(2) = "FH"
        FH(3) = "TC1"
        FH(4) = "R"
        FH(5) = Format(Today(), "ddMMyyyy")
        FH(6) = 1
        FH(7) = "D"
        FH(8) = UCase(oCoMst.CoTAN)
        FH(9) = 1
        Dim versionNumber As String
        Dim a
        versionNumber = Assembly.GetExecutingAssembly.FullName
        versionNumber = Strings.Left(versionNumber, 29)
        a = Split(versionNumber, ",")

        FH(10) = a(0) + " " + Strings.Right(a(1), 7)
        FH(11) = vbNullString
        FH(12) = vbNullString
        FH(13) = vbNullString
        FH(14) = vbNullString
        FH(15) = vbNullString
        FH(16) = vbNullString
        FH(17) = vbNullString
        FH(18) = vbNullString   'Added on 28/01/2013 for new FVU 3.3
        strFH = vbNullString
        For i = 1 To UBound(FH)
            strFH = strFH & FH(i)
            If i <= UBound(FH) - 1 Then strFH = strFH & "^"
        Next i
        TStrm.WriteLine(strFH)
        '*****************************************************************************************
        'BATCH HEADER RECORD...
        'Calculate Challan Total, and Deductee Total
        Dim ChallanTotal As Double
        ChallanTotal = Val(rstCSum.Tables(0).Rows(0)("TotAmt").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("totsc").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("totecess").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("totint").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("tototh").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("totfees").ToString())
        If String.IsNullOrEmpty(ChallanTotal) = True Then
            SumC = 0
        Else
            SumC = CDbl(ChallanTotal)
        End If
        BH(1) = 2
        BH(2) = "BH"
        BH(3) = 1
        BH(4) = rstC.Tables(0).Rows.Count
        BH(5) = "27EQ"
        BH(6) = vbNullString
        BH(7) = vbNullString
        BH(8) = vbNullString
        BH(9) = lastrr 'IIf(IsNull(lastrr), "", vbNullString) 'vbNullString
        BH(10) = vbNullString
        BH(11) = vbNullString
        BH(12) = vbNullString
        BH(13) = UCase(oCoMst.CoTAN)
        BH(14) = vbNullString
        BH(15) = IIf(Len(Trim(oCoMst.CoPAN)) = 0, "PANNOTREQD", UCase(oCoMst.CoPAN))    'Changed on 03/10/09, as per new format ver. 4.0
        BH(16) = Left(AY, 4) & Right(AY, 2)     'eg. AY format 200607
        BH(17) = Left(FY, 4) & Right(FY, 2)     'eg. FY format 200506
        BH(18) = "Q" & Right(frmid, 1)
        BH(19) = oCoMst.CoName
        BH(20) = oCoMst.CoBrDiv
        BH(21) = oCoMst.CoAdd1
        BH(22) = oCoMst.CoAdd2
        BH(23) = oCoMst.CoAdd3
        BH(24) = oCoMst.CoAdd4
        BH(25) = oCoMst.CoAdd5
        BH(26) = oCoMst.CoStateID
        BH(27) = oCoMst.CoPin
        BH(28) = oCoMst.CoEmail
        BH(29) = oCoMst.CoStd
        BH(30) = oCoMst.CoPhone
        BH(31) = IIf(oCoMst.IsCoAddChg = True, "Y", "N")
        BH(32) = oCoMst.CoStatus
        BH(33) = oCoMst.PR27Ename
        BH(34) = oCoMst.PR27EDesg
        BH(35) = oCoMst.PR27EAdd1
        BH(36) = oCoMst.PR27EAdd2
        BH(37) = oCoMst.PR27EAdd3
        BH(38) = oCoMst.PR27EAdd4
        BH(39) = oCoMst.PR27EAdd5
        BH(40) = oCoMst.PR27EStateID
        BH(41) = oCoMst.PR27EPin
        BH(42) = oCoMst.PR27EEmail
        BH(43) = oCoMst.Comobile  'vbNullString   'mobile no - not provided in this release
        BH(44) = oCoMst.PR27EStd
        BH(45) = oCoMst.PR27EPhone
        BH(46) = IIf(oCoMst.IsPR27EAddChg = True, "Y", "N")
        BH(47) = Format(ChallanTotal, "0.00")
        BH(48) = vbNullString
        BH(49) = vbNullString
        BH(50) = vbNullString
        BH(51) = "N"
        BH(52) = lastret  'IIf(IsNull(lastrr), "", vbNullString)   'vbNullString
        BH(53) = vbNullString
        BH(54) = IIf(oCoMst.GovtStateID = "", "", Format(oCoMst.GovtStateID, "00"))
        BH(55) = oCoMst.PAOCode
        BH(56) = oCoMst.DDOCode
        BH(57) = IIf(oCoMst.MinistryID = "", "", Format(oCoMst.MinistryID, "00"))
        BH(58) = IIf(Val(BH(57)) = 99, oCoMst.MinistryName, vbNullString)
        BH(59) = oCoMst.PR27PAN     'Trim(oCoMst.TANRegNo)
        BH(60) = IIf(oCoMst.PAORegNo = "", "", Format(oCoMst.PAORegNo, "00"))
        BH(61) = oCoMst.DDORegNo
        BH(62) = oCoMst.CoSTDAlt
        BH(63) = oCoMst.CoPhoneAlt
        BH(64) = oCoMst.CoEmailAlt
        BH(65) = oCoMst.PR27ESTDAlt
        BH(66) = oCoMst.PR27EPhoneAlt
        BH(67) = oCoMst.PR27EEmailAlt
        BH(68) = oCoMst.AIN
        BH(69) = oCoMst.gstin
        BH(70) = vbNullString

        strBH = ""
        For i = 1 To UBound(BH)
            strBH = strBH & BH(i)
            If i <= UBound(BH) - 1 Then strBH = strBH & "^"
        Next i
        TStrm.WriteLine(strBH)
        '*****************************************************************************************
        'CHALLAN DETAIL RECORD
        If rstC.Tables(0).Rows.Count > 0 Then
            'rstC.MoveFirst
        End If
        LNo = 3 : RecNo = 1
        Dim CTotal As Long
        'Do While Not rstC.Tables(0).Rows.Count
        For c = 0 To rstC.Tables(0).Rows.Count - 1
            rstChlnDed = FetchDataSet("SELECT DT.*, DM.* FROM Deductee27EQ AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString() &
      " AND ChallanID=" & rstC.Tables(0).Rows(c)("ChallanID").ToString()) ', Cnn   'Deductee Records
            'fees added on 17/05/14
            'CTotal = (CLng(IIf(IsNull(rstC!TaxAmt), 0, rstC!TaxAmt)) + CLng(IIf(IsNull(rstC!Surcharge), 0, rstC!Surcharge)) + CLng(IIf(IsNull(rstC!ECess), 0, rstC!ECess)) + CLng(IIf(IsNull(rstC!Interest), 0, rstC!Interest)) + CLng(IIf(IsNull(rstC!Others), 0, rstC!Others)))
            'CTotal = CLng(IIf((rstC.Tables(0).Rows(0)("TaxAmt").ToString()), 0, rstC.Tables(0).Rows(0)("TaxAmt").ToString())) +
            '         CLng(IIf((rstC.Tables(0).Rows(0)("Surcharge").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharge").ToString())) +
            '         CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString()) +
            '         CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString())) +
            '         CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString())) +
            '         CLng(IIf((rstC.Tables(0).Rows(0)("AFees").ToString()), 0, rstC.Tables(0).Rows(0)("AFees").ToString())))
            CTotal = Val(rstC.Tables(0).Rows(c)("TaxAmt").ToString()) +
                     Val(rstC.Tables(0).Rows(c)("Surcharge").ToString()) +
                     Val(rstC.Tables(0).Rows(c)("ECess").ToString()) +
                     Val(rstC.Tables(0).Rows(c)("Interest").ToString()) +
                     Val(rstC.Tables(0).Rows(c)("Others").ToString()) +
                     Val(rstC.Tables(0).Rows(c)("AFees").ToString())
            CD(1) = LNo
            CD(2) = "CD"
            CD(3) = 1
            CD(4) = RecNo
            CD(5) = rstChlnDed.Tables(0).Rows.Count
            CD(6) = IIf(CTotal <= 0, "Y", "N")
            CD(7) = vbNullString
            CD(8) = vbNullString
            CD(9) = vbNullString
            CD(10) = vbNullString
            CD(11) = vbNullString
            CD(12) = IIf((CD(6) = "Y" Or (oCoMst.CoStatus = "C" And rstC.Tables(0).Rows(c)("IsBookEntry").ToString() = True)), vbNullString, rstC.Tables(0).Rows(c)("BankChallanNo").ToString())
            CD(13) = vbNullString
            'changed on 03/10/09 as per new format ver 4.0
            If BH(32) = "A" Or BH(32) = "S" Or BH(32) = "D" Or BH(32) = "E" Or BH(32) = "G" Or
        BH(32) = "H" Or BH(32) = "L" Or BH(32) = "N" Then
                CD(14) = rstC.Tables(0).Rows(c)("TranVouNo").ToString()
            Else
                CD(14) = vbNullString
            End If
            '    CD(14) = IIf(oCoMst.CoStatus = "C" And rstC!IsBookEntry = True, rstC!BankChallanNo, vbNullString)
            CD(15) = vbNullString
            CD(16) = IIf(String.IsNullOrEmpty((rstC.Tables(0).Rows(c)("BankBrCode").ToString())), vbNullString, Format(Val(rstC.Tables(0).Rows(c)("BankBrCode").ToString()), "0000000"))
            CD(17) = vbNullString

            If CD(6) = "Y" Then
                Select Case Right(frmid, 1)
                    Case 1
                        Value18 = "30/06/" & Left(FY, 4)
                    Case 2
                        Value18 = "30/09/" & Left(FY, 4)
                    Case 3
                        Value18 = "31/12/" & Left(FY, 4)
                    Case 4
                        Value18 = "31/03/" & Right(FY, 4)
                End Select
            Else
                Value18 = rstC.Tables(0).Rows(c)("DtOfChallan").ToString()
            End If
            CD(18) = Format(Value18, "ddMMyyyy")
            CD(19) = vbNullString
            CD(20) = vbNullString
            CD(21) = vbNullString     'IIf(IsNull(rstC!CollCode), vbNullString, rstC!CollCode)
            CD(22) = Format(Val(rstC.Tables(0).Rows(c)("TaxAmt").ToString()), "0.00")
            CD(23) = Format(Val(rstC.Tables(0).Rows(c)("Surcharge").ToString()), "0.00")
            CD(24) = Format(Val(rstC.Tables(0).Rows(c)("ECess").ToString()), "0.00")
            CD(25) = Format(Val(rstC.Tables(0).Rows(c)("Interest").ToString()), "0.00")
            CD(26) = Format(Val(rstC.Tables(0).Rows(c)("Others").ToString()), "0.00")
            CD(39) = Format(Val(rstC.Tables(0).Rows(c)("AFees").ToString()), "0.00")
            CD(27) = Format((Val(CD(22)) + Val(CD(23)) + Val(CD(24)) + Val(CD(25)) + Val(CD(26)) + Val(CD(39))), "0.00")
            CD(28) = vbNullString

            'If rstDSum.State = adStateOpen Then
            rstDSum.Dispose()
            ' End If
            rstDSum = FetchDataSet("SELECT sum(TotalTaxDeposited) as TotDep, sum(TaxAmt) as TotTax, " &
       "sum(Surcharge) as TotSur, sum(Ecess) as TotEcess, sum(0) as TotInt, " &
       "sum(0) as TotOth FROM Deductee27EQ WHERE ChallanID=" & rstC.Tables(0).Rows(c)("ChallanID").ToString()) ', Cnn


            CD(29) = Format(Val(rstDSum.Tables(0).Rows(0)("totdep").ToString()), "0.00")
            CD(30) = Format(Val(rstDSum.Tables(0).Rows(0)("tottax").ToString()), "0.00")
            CD(31) = Format(Val(rstDSum.Tables(0).Rows(0)("totsur").ToString()), "0.00")
            CD(32) = Format(Val(rstDSum.Tables(0).Rows(0)("totecess").ToString()), "0.00")
            CD(33) = Format(Val(CD(30)) + Val(CD(31)) + Val(CD(32)), "0.00")
            CD(34) = Format(Val(rstC.Tables(0).Rows(c)("AInterest").ToString()), "0.00")
            CD(35) = Format(Val(rstC.Tables(0).Rows(c)("AOthers").ToString()), "0.00")
            CD(36) = vbNullString    'rstC!ChqDDNo for FVU3.8
            CD(37) = IIf(rstC.Tables(0).Rows(c)("IsBookEntry").ToString(), "Y", IIf(Val(CD(27)) = 0, "", "N"))  'Value changed to NULL in case of NIL challan FUV3.9
            CD(38) = vbNullString
            'CD(39) = already declared above total challan CD(27)
            CD(40) = IIf(Val(CD(27)) = 0 Or CD(37) = "Y", "", rstC.Tables(0).Rows(c)("MinorHead").ToString() & "")
            CD(41) = vbNullString

            strCD = ""
            For i = 1 To UBound(CD)
                strCD = strCD & CD(i)
                If i <= UBound(CD) - 1 Then strCD = strCD & "^"
            Next i
            TStrm.WriteLine(strCD)
            LNo = LNo + 1
            'Insert Deductee Details for this challan
            '*****************************************************************************************
            'DEDUCTEE DETAIL RECORD
            RecNoDed = 1      'Line no not reset as it will continue from challan detail...
            'open the recordset
            If rstChlnDed.Tables(0).Rows.Count > 0 Then
                'rstChlnDed.MoveFirst
            End If
            For k = 0 To rstChlnDed.Tables(0).Rows.Count - 1
                DD(1) = LNo
                DD(2) = "DD"
                DD(3) = 1
                DD(4) = CD(4)
                DD(5) = RecNoDed
                DD(6) = "O"
                DD(7) = vbNullString
                DD(8) = IIf(rstChlnDed.Tables(0).Rows(k)("DCode").ToString() = "O", 2, 1)
                DD(9) = vbNullString
                DD(10) = rstChlnDed.Tables(0).Rows(k)("DPan").ToString()
                DD(11) = vbNullString
                DD(12) = rstChlnDed.Tables(0).Rows(k)("Dpanref").ToString()
                DD(13) = IIf(String.IsNullOrEmpty(rstChlnDed.Tables(0).Rows(k)("DName").ToString()), "", rstChlnDed.Tables(0).Rows(k)("DName").ToString())
                DD(14) = Format(Val(rstChlnDed.Tables(0).Rows(k)("TaxAmt").ToString()), "0.00")
                DD(15) = Format(Val(rstChlnDed.Tables(0).Rows(k)("Surcharge").ToString()), "0.00")
                DD(16) = Format(Val(rstChlnDed.Tables(0).Rows(k)("ECess").ToString()), "0.00")
                DD(17) = Format(Val(DD(14)) + Val(DD(15)) + Val(DD(16)), "0.00")
                DD(18) = vbNullString
                DD(19) = Format(Val(rstChlnDed.Tables(0).Rows(k)("TotalTaxDeposited").ToString()), "0.00")
                DD(20) = vbNullString
                DD(21) = Format(Val(rstChlnDed.Tables(0).Rows(k)("PurchAmt").ToString()), "0.00")
                DD(22) = Format(Val(rstChlnDed.Tables(0).Rows(k)("AmtOfPayment").ToString()), "0.00")
                Dim dt As Date
                dt = rstChlnDed.Tables(0).Rows(k)("DtOfPayment")
                DD(23) = Format(dt, "ddMMyyyy")
                If rstChlnDed.Tables(0).Rows(k)("DtOfDeduction").ToString = "" Then
                    DD(24) = vbNullString
                Else

                    dt = rstChlnDed.Tables(0).Rows(k)("DtOfDeduction")
                    DD(24) = Format(dt, "ddMMyyyy")
                End If
                DD(25) = vbNullString
                DD(26) = Format(Val(rstChlnDed.Tables(0).Rows(k)("RateOfTDS").ToString()), "##0.0000")
                DD(27) = vbNullString
                DD(28) = vbNullString    'CD(37)
                DD(29) = vbNullString
                DD(30) = Trim(rstChlnDed.Tables(0).Rows(k)("Remark").ToString())
                DD(31) = IIf(rstChlnDed.Tables(0).Rows(k)("dpancat").ToString() <> 0, IIf(rstChlnDed.Tables(0).Rows(k)("CollNonRes").ToString() = True, "Y", "N"), "") 'vbNullString
                DD(32) = IIf(DD(31) = "Y", IIf(rstChlnDed.Tables(0).Rows(k)("PerEstInd").ToString() = True, "Y", "N"), "") 'vbNullString
                DD(33) = rstChlnDed.Tables(0).Rows(k)("Sec").ToString()
                DD(34) = rstChlnDed.Tables(0).Rows(k)("CertNo").ToString()
                DD(35) = vbNullString
                DD(36) = vbNullString
                DD(37) = vbNullString
                DD(38) = vbNullString
                DD(39) = vbNullString
                DD(40) = vbNullString
                DD(41) = vbNullString
                DD(42) = vbNullString
                DD(43) = vbNullString
                RecNoDed = RecNoDed + 1
                If RecNoDed > 999999999 Then RecNoDed = 1
                strDD = ""
                For i = 1 To UBound(DD)
                    strDD = strDD & DD(i)
                    If i <= UBound(DD) - 1 Then strDD = strDD & "^"
                Next i
                TStrm.WriteLine(strDD)
                'warn if PAN Available and remark = "C"
                If IsValidPAN(rstChlnDed.Tables(0).Rows(k)("DPan").ToString()) = 0 Then
                    If DD(30) = "C" Then
                        If MsgBox("WARNING: Please remove remark 'C' of Deductee:" & DD(13) & " and then convert the file before filing. " & vbCrLf &
                    "Do You want to Continue", vbYesNo, "WARNING") = vbNo Then
                            CPan = True
                        End If
                    End If
                End If
                LNo = LNo + 1
                'rstChlnDed.MoveNext
            Next k
            'If rstChlnDed.State = adStateOpen Then 
            'rstChlnDed.Dispose()
            'rstC.MoveNext
            RecNo = RecNo + 1
        Next c
        TStrm.Close()
cleanup:
        'CLOSE THE FILE..

        'AND THE CONNECTIONS ALSO..
        'If rstC.State = adStateOpen Then 
        rstC.Dispose()
        rstD.Dispose()
        rstRetn.Dispose()
        rstCSum.Dispose()
        rstDSum.Dispose()
        rstC = Nothing
        rstD = Nothing
        rstRetn = Nothing
        rstCSum = Nothing
        rstDSum = Nothing
        'fs = Nothing
       ' oCoMst = Nothing
        frmConUtility.isError = False
        Exit Sub
    End Sub

    Public Sub Convert27ETxt(frmid As String, filename As String, revised As String, oldRRRNo As Double, TANApplNo As Double)
        Dim rstC, rstD, rstRetn, rstDSum, rstCSum As New DataSet
        Dim LNo As Long, RecNo As Long
        Dim SumC As Double, SumD As Double
        If frmid <> "F27E" Then
            MsgBox("Wrong Parameter, Call JAK Infosolutions P Ltd", vbCritical)
            GoTo cleanup
        Else
            'Parameter ok, check if return exists..
            rstRetn = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " AND FrmType = 31")
            If rstRetn.Tables(0).Rows.Count <= 0 Then
                Call MsgBox("There is no data for this return form.  Kindly create" _
            & vbCrLf & "this return by using Data entry option and then" _
            & vbCrLf & "create e-TDS file using this option." _
            , vbExclamation + vbDefaultButton1, "RETURN NOT FOUND")
                GoTo cleanup
            End If
        End If
        oCoMst = oCoMst.FetchCo(selectedcoid)
        eTDSPath = Application.StartupPath & "\e-TDS Files"
        eFileName = eTDSPath & "\" & oCoMst.CoName & "\" & frmid & "V" & ".txt"
        'Do the conversion
        Dim TStrm As New System.IO.StreamWriter(eFileName, True)
        'TStrm = fso.CreateTextFile(filename, True)
        'Get the related data..
        rstC = FetchDataSet("SELECT * FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID"))      'Challan Records
        rstD = FetchDataSet("SELECT DT.*, DM.* FROM Deductee27E AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" &
      rstRetn.Tables(0).Rows(0)("RetnID"))  'Deductee Records
        rstCSum = FetchDataSet("SELECT sum(Amt) as TotAmt,sum(Surcharges) as TotSc, sum(Ecess) as TotEcess," &
               " Sum(Interest) as TotInt, sum(Others) as TotOth FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID"))    'Challan Records
        rstDSum = FetchDataSet("SELECT sum(AmtOfTDS) as TotD FROM Deductee27E WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID")) 'Deductee Records
        'START WRITING THE TEXT FILE NOW.
        Dim txtTANAppNo As Double
        'If oCoMst.CoTAN = "TANAPPLIED" Then
        '   txtTANAppNo = Format(InputBox("Please Enter your TAN Application No", "TANAPPLIED Number", 0), "00000000000000")
        'End If
        'FILE HEADER RECORD...COMMON FOR F27E AND F27
        If rstD.Tables(0).Rows.Count = 0 And rstC.Tables(0).Rows.Count = 0 Then
            ' FILE TYPE CHANGED FROM XNS TO NS3 AS PER NEW FORMAT DT 22/4/2005
            TStrm.WriteLine("000000001FHNS3" & revised & Format(Today(), "ddMMyyyy") & "000000001" & oCoMst.CoTAN & "000000001")
        ElseIf rstD.Tables(0).Rows.Count > 999999999 Or rstC.Tables(0).Rows.Count > 999999999 Then
            ' FILE TYPE CHANGED FROM ENS TO NS3 AS PER NEW FORMAT DT 22/4/2005
            TStrm.WriteLine("000000001FHNS3" & revised & Format(Today(), "ddMMyyyy") & "000000001" & oCoMst.CoTAN & "000000001")
        Else
            TStrm.WriteLine("000000001FHNS3" & revised & Format(Today(), "ddMMyyyy") & "000000001" & oCoMst.CoTAN & "000000001")
        End If
        If frmid = "F27E" Then
            Dim ChallanTotal As Double
            ChallanTotal = (IIf((rstCSum.Tables(0).Rows(0)("TotAmt")), 0, rstCSum.Tables(0).Rows(0)("TotAmt")) +
                    IIf((rstCSum.Tables(0).Rows(0)("totsc").ToString()), 0, rstCSum.Tables(0).Rows(0)("totsc").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("totecess").ToString()), 0, rstCSum.Tables(0).Rows(0)("totecess").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("totint").ToString()), 0, rstCSum.Tables(0).Rows(0)("totint").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("tototh").ToString()), 0, rstCSum.Tables(0).Rows(0)("tototh").ToString()))
            '    ChallanTotal = (rstCSum!TotAmt + rstCSum!totsc + rstCSum!Totecess + rstCSum!TotInt + rstCSum!TotOth)
            If String.IsNullOrEmpty(ChallanTotal) = True Then
                SumC = 0
            Else
                SumC = CDbl(ChallanTotal * 100)
            End If
            If String.IsNullOrEmpty(rstDSum.Tables(0).Rows(0)("totd").ToString()) = True Then
                SumD = 0
            Else
                SumD = CDbl(rstDSum.Tables(0).Rows(0)("totd").ToString() * 100)
            End If

            'BATCH HEADER RECORD.

            TStrm.WriteLine("000000002BH000000001" & Format(IIf(rstC.Tables(0).Rows.Count > 999999999, "999999999", rstC.Tables(0).Rows.Count), "000000000") &
   Format(IIf(rstD.Tables(0).Rows.Count > 999999999, "999999999", rstD.Tables(0).Rows.Count), "000000000") & Format("27E", "!@@@@") & Space(8) &
   oCoMst.CoTAN & oCoMst.CoPAN & Left(AY, 4) & Right(AY, 2) & Left(FY, 4) & Right(FY, 2) &
   Format(oCoMst.CoName, "!" & New String("@", 75)) & Format(oCoMst.CoAdd1, "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd2), Space(25), IIf(oCoMst.CoAdd2 = vbNullString, Space(25), oCoMst.CoAdd2)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd3), Space(25), IIf(oCoMst.CoAdd3 = vbNullString, Space(25), oCoMst.CoAdd3)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd4), Space(25), IIf(oCoMst.CoAdd4 = vbNullString, Space(25), oCoMst.CoAdd4)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd5), Space(25), IIf(oCoMst.CoAdd5 = vbNullString, Space(25), oCoMst.CoAdd5)), "!" & New String("@", 25)) &
   Format(oCoMst.CoStateID, "00") & Format(oCoMst.CoPin, "000000") &
   IIf(oCoMst.IsCoAddChg = True, "Y", "N") & oCoMst.CoStatus &
   Format("Y", "!@@") & Format(oCoMst.PR27Ename, "!" & New String("@", 75)) &
   Format(oCoMst.PR27EDesg, "!" & New String("@", 20)) &
   Format(SumC, New String("0", 14)) &
   Format(SumD, New String("0", 14)) & Format(0, New String("0", 14) &
   Space(10) & Format(txtTANAppNo, New String("0", 14)) &
   Format(oldRRRNo, New String("0", 14))))   '' revised return number to be incrop...
        ElseIf frmid = "F27" Then
        End If
        'CHALLAN DETAIL RECORD
        If rstC.Tables(0).Rows.Count > 0 Then
            'rstC.MoveFirst
        End If
        LNo = 3 : RecNo = 1
        Dim CTotal As Long
        Do While Not rstC.Tables(0).Rows.Count
            CTotal = (CLng(IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess)").ToString()) + CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString())))) * 100
            TStrm.WriteLine(Format(LNo, "000000000") & "CD" & "000000001" & Format(RecNo, "000000000") & Format(rstC.Tables(0).Rows(0)("Sec").ToString(), "!@@@@@") &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)!Interest), 0, rstC.Tables(0).Rows(0)("Interest").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString()) * 100), New String("0", 14)) &
      Format(CTotal, New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), 0, rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstC.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstC.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(rstC.Tables(0).Rows(0)("DtOfVoucher").ToString(), "ddMMyyyy") &
      Format(IIf((rstC.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstC.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstC.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      IIf(rstC.Tables(0).Rows(0)("IsBookEntry").ToString() = True, "Y", "N") & rstC.Tables(0).Rows(0)("CollCode").ToString())
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            '  rstC.MoveNext
        Loop
        'DEDUCTEE DETAIL RECORD
        If rstD.Tables(0).Rows.Count > 0 Then
            'rstD.MoveFirst
        End If
        RecNo = 1      'Line no not reset as it will continue from challan detail...
        Do While Not rstD.Tables(0).Rows.Count
            ' new filler added after pin no as per new format dt 22/04/2005, isbookentry added..
            '
            TStrm.WriteLine(Format(LNo, "000000000") & "DD" & "000000001" & Format(RecNo, "000000000") & Format(rstD.Tables(0).Rows(0)("Sec").ToString(), "!@@@@@") &
      Format(IIf(rstD.Tables(0).Rows(0)("DType").ToString() = "O", 2, 1), "00") &
      Format(IIf((rstD.Tables(0).Rows(0)("DPan").ToString()), Space(10), IIf(rstD.Tables(0).Rows(0)("DPan").ToString() = vbNullString, Space(10), rstD.Tables(0).Rows(0)("DPan").ToString())), "!@@@@@@@@@@") &
      Format(rstD.Tables(0).Rows(0)("DName").ToString(), "!" & New String("@", 75)) & Format(rstD.Tables(0).Rows(0)("DAdd1").ToString(), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd2").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd2").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd2").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd3").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd3").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd3").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd4").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd4").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd4").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd5").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd5").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd5").ToString())), "!" & New String("@", 25)) &
      Format(rstD.Tables(0).Rows(0)("DState").ToString(), "00") & Format(rstD.Tables(0).Rows(0)("DPin").ToString(), "000000") &
      Format(rstD.Tables(0).Rows(0)("PurchAmt").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("AmtOfPay").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("DtOfPay").ToString(), "ddMMyyyy") & Space(1) &
      Format(IIf(rstD.Tables(0).Rows(0)("RateOfTDS").ToString() >= 100, 0, rstD.Tables(0).Rows(0)("RateOfTDS").ToString()) * 100, "0000") & Space(1) &
      Format(rstD.Tables(0).Rows(0)("AmtOfTDS").ToString() * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstD.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstD.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstD.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstD.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("CertificateDt").ToString()), Space(8), rstD.Tables(0).Rows(0)("CertificateDt").ToString()), "ddMMyyyy") &
      "X" & New String("0", 14))
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then RecNo = 1
            ' rstD.MoveNext
        Loop
        TStrm.Close()
cleanup:
        'CLOSE THE FILE..

        'AND THE CONNECTIONS ALSO..
        rstC.Dispose()
        rstD.Dispose()
        rstRetn.Dispose()
        rstCSum.Dispose()
        rstDSum.Dispose()
        rstC = Nothing
        rstD = Nothing
        rstRetn = Nothing
        rstCSum = Nothing
        rstDSum = Nothing
        filename = Nothing
        oCoMst = Nothing

    End Sub

    Public Sub Convert2Txt24(filename As String, revised As String, oldRRRNo As Double, TANApplNo As Double)
        Dim rstC As New DataSet, rstD As New DataSet, rstP As New DataSet
        Dim rstRetn As New DataSet
        Dim rstCSum As New DataSet, rstDSum As New DataSet
        Dim LNo As Long, RecNo As Long
        Dim SumC As Double, SumD As Double
        Dim DRec As String
        Dim Total209 As Double, Total211 As Double, Total213 As Double, Total217 As Double
        Dim Total218 As Double, Total223 As Double, Total225 As Double
        Dim Total228 As Double, Total229 As Double

        'Parameter ok, check if return exists..
        rstRetn = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " AND FrmType = 24") ', Cnn
        If rstRetn.Tables(0).Rows.Count <= 0 Then
            Call MsgBox("There is no data for this return form.  Kindly create" _
         & vbCrLf & "this return by using Data entry option and then" _
         & vbCrLf & "create e-TDS file using this option." _
         , vbExclamation + vbDefaultButton1, "RETURN NOT FOUND")
            GoTo cleanup
        End If
        oCoMst = oCoMst.FetchCo(selectedcoid)
        eTDSPath = Application.StartupPath & "\e-TDS Files"
        eFileName = eTDSPath & "\" & "24-" & oCoMst.CoName & ".TXT"

        ' TStrm = fso.CreateTextFile(filename, True)
        Dim TStrm As New System.IO.StreamWriter(eFileName, True)
        'Get the related data..
        rstC = FetchDataSet("SELECT * FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn      'Challan Records
        rstD = FetchDataSet("SELECT DT.*, DM.* FROM DeducteeSAL AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" &
      rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn   'Deductee Records

        '''for checking by namrata, fetches 2 records
        'even when retn id is not matching..eg. fetch 2 records of retn
        ''id = 13 when calling for retnid=2
        Dim sql As String

        rstP = FetchDataSet("SELECT P.*, D.* FROM RetnMst AS R, PerqSAL AS P, DeductMst as D " &
      "where  P.DId = D.DId and p.retnid=r.retnid and r.RetnID=" &
      rstRetn.Tables(0).Rows(0)("RetnID").ToString())
        '   rstP.Open sql, Cnn   'Preq records..
        '
        'rstP.Open "SELECT P.*, D.* FROM RetnMst AS R INNER JOIN " & _
        '    "(PerqSAL AS P INNER JOIN DeductMst AS D ON P.DId = D.DId) ON R.RetnID = P.RetnID" & _
        '    "WHERE (((R.RetnID)=" & rstRetn!RetnId & "));", Cnn

        rstCSum = FetchDataSet("SELECT sum(Amt) as TotAmt,sum(Surcharges) as TotSc, sum(Ecess) as TotEcess," &
               " Sum(Interest) as TotInt, sum(Others) as TotOth FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn      'Challan Records
        rstDSum = FetchDataSet("SELECT sum(TDSAmt) as TotAmt,sum(TDSEcess)as TotEcess,sum(TDSSurcharge) as TotSC FROM DeducteeSAL WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn  'Deductee Records

        'START WRITING THE TEXT FILE NOW.
        'FILE HEADER RECORD
        If rstD.Tables(0).Rows.Count = 0 And rstC.Tables(0).Rows.Count = 0 Then
            ' FILE TYPE CHANGED TO SL3 from XSA as per new format dated 22/4/2005 by nitin
            TStrm.WriteLine("000000001FHSL3" & revised & Format(Today(), "ddMMyyyy") & "000000001" & oCoMst.CoTAN & "000000001")
        ElseIf rstD.Tables(0).Rows.Count > 999999999 Or rstC.Tables(0).Rows.Count > 999999999 Then
            ' FILE TYPE CHANGED TO SL3 from ESA as per new format dated 22/4/2005 by nitin
            TStrm.WriteLine("000000001FHSL3" & revised & Format(Today(), "ddMMyyyy") & "000000001" & oCoMst.CoTAN & "000000001")
        Else
            'FILE TYPE CHANGED FROM SL1 TO SL3 AS PER NEW FORMATS DT. 22/04/2005. ON 03/05/2005 BY NITIN..
            TStrm.WriteLine("000000001FHSL3" & revised & Format(Today(), "ddMMyyyy") & "000000001" & oCoMst.CoTAN & "000000001")
        End If
        Dim ChallanTotal As Double
        Dim DeducteeTotal As Double
        ChallanTotal = (IIf((rstCSum.Tables(0).Rows(0)("TotAmt").ToString()), 0, rstCSum.Tables(0).Rows(0)("TotAmt").ToString()) +
                 IIf((rstCSum.Tables(0).Rows(0)("totsc").ToString()), 0, rstCSum.Tables(0).Rows(0)("totsc").ToString()) +
                 IIf((rstCSum.Tables(0).Rows(0)("totecess").ToString()), 0, rstCSum.Tables(0).Rows(0)("totecess").ToString()) +
                 IIf((rstCSum.Tables(0).Rows(0)("totint").ToString()), 0, rstCSum.Tables(0).Rows(0)("totint").ToString()) +
                 IIf((rstCSum.Tables(0).Rows(0)("tototh").ToString()), 0, rstCSum.Tables(0).Rows(0)("tototh").ToString()))
        DeducteeTotal = (IIf((rstDSum.Tables(0).Rows(0)("TotAmt").ToString()), 0, rstDSum.Tables(0).Rows(0)("TotAmt").ToString()) +
                  IIf((rstDSum.Tables(0).Rows(0)("totsc").ToString()), 0, rstDSum.Tables(0).Rows(0)("totsc").ToString()) +
                  IIf((rstDSum.Tables(0).Rows(0)("totecess").ToString()), 0, rstDSum.Tables(0).Rows(0)("totecess").ToString()))
        If String.IsNullOrEmpty(ChallanTotal) = True Then
            SumC = 0
        Else
            SumC = CLng(ChallanTotal * 100)
        End If
        If String.IsNullOrEmpty(DeducteeTotal) = True Then
            SumD = 0
        Else
            SumD = CLng(DeducteeTotal * 100)
        End If

        'BATCH HEADER RECORD.
        TStrm.WriteLine("000000002BH000000001" & Format(IIf(rstC.Tables(0).Rows.Count > 999999999, "999999999", rstC.Tables(0).Rows.Count), "000000000") &
Format(IIf(rstD.Tables(0).Rows.Count > 999999999, "999999999", rstD.Tables(0).Rows.Count), "000000000") & Format(IIf(rstP.Tables(0).Rows.Count > 999999999, "999999999", rstP.Tables(0).Rows.Count), "000000000") &
Space(8) &
oCoMst.CoTAN & oCoMst.CoPAN & Left(AY, 4) & Right(AY, 2) & Left(FY, 4) & Right(FY, 2) &
Format(oCoMst.CoName, "!" & New String("@", 75)) &
IIf(oCoMst.CoStatus = "O", "0000000002", "0000000001") &
Format(oCoMst.CoAdd1, "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd2), Space(25), IIf(oCoMst.CoAdd2 = vbNullString, Space(25), oCoMst.CoAdd2)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd3), Space(25), IIf(oCoMst.CoAdd3 = vbNullString, Space(25), oCoMst.CoAdd3)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd4), Space(25), IIf(oCoMst.CoAdd4 = vbNullString, Space(25), oCoMst.CoAdd4)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd5), Space(25), IIf(oCoMst.CoAdd5 = vbNullString, Space(25), oCoMst.CoAdd5)), "!" & New String("@", 25)) &
Format(oCoMst.CoStateID, "00") & Format(oCoMst.CoPin, "000000") &
IIf(oCoMst.IsCoAddChg = True, "Y", "N") &
Format(oCoMst.PRName24, "!" & New String("@", 75)) &
Format(oCoMst.PRDesg24, "!" & New String("@", 20)) &
Format(oCoMst.PR24Add1, "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.PR24Add2), Space(25), IIf(oCoMst.PR24Add2 = vbNullString, Space(25), oCoMst.PR24Add2)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.PR24Add3), Space(25), IIf(oCoMst.PR24Add3 = vbNullString, Space(25), oCoMst.PR24Add3)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.PR24Add4), Space(25), IIf(oCoMst.PR24Add4 = vbNullString, Space(25), oCoMst.PR24Add4)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.PR24Add5), Space(25), IIf(oCoMst.PR24Add5 = vbNullString, Space(25), oCoMst.PR24Add5)), "!" & New String("@", 25)) &
Format(oCoMst.PR24StateID, "00") & Format(oCoMst.PR24Pin, "000000") &
IIf(oCoMst.IsPR24AddChg = True, "Y", "N") &
Format(SumC, New String("0", 14)) &
Format(SumD, New String("0", 14)) & Format(0, New String("0", 14)) & Space(10) &
Format(IIf(TANApplNo = 0, 0, TANApplNo), New String("0", 14)) & Format(IIf(oldRRRNo = 0, 0, oldRRRNo), New String("0", 14)))

        'CHALLAN DETAIL RECORD, IMPORTANT NOTE: this is different from form 27E,27,
        'section field is not used here...
        Dim CTotal As Long
        If rstC.Tables(0).Rows.Count > 0 Then
            'rstC.MoveFirst
        End If
        LNo = 3 : RecNo = 1
        Do While Not rstC.Tables(0).Rows.Count
            CTotal = (CLng(IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString()))) * 100
            TStrm.WriteLine(Format(LNo, "000000000") & "CD" & "000000001" & Format(RecNo, "000000000") &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString()) * 100), New String("0", 14)) &
      Format(CTotal, New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), 0, rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstC.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstC.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(rstC.Tables(0).Rows(0)("DtOfVoucher").ToString(), "ddMMyyyy") &
      Format(IIf((rstC.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstC.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstC.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      IIf(rstC.Tables(0).Rows(0)("IsBookEntry").ToString() = True, "Y", "N"))

            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            ' rstC.MoveNext
        Loop
        'DEDUCTEE DETAIL RECORD
        Dim rno As Long
        If rstD.Tables(0).Rows.Count > 0 Then
            'rstD.MoveFirst
        End If
        RecNo = 1      'Line no not reset as it will continue from challan detail...
        rno = rstD.Tables(0).Rows.Count
        Dim esr() As String
        ReDim Preserve esr(rno)
        Do While Not rstD.Tables(0).Rows.Count

            'Do some calcuations...
            Total209 = rstD.Tables(0).Rows(0)("TotalSal").ToString() + rstD.Tables(0).Rows(0)("TotalRent").ToString() + rstD.Tables(0).Rows(0)("PFAmt").ToString()
            Total211 = Total209 - rstD.Tables(0).Rows(0)("Deduction16").ToString()
            If rstD.Tables(0).Rows(0)("OtherIncomeFlg").ToString() = "P" Then
                Total213 = Total211 + rstD.Tables(0).Rows(0)("OtherIncomeAmt").ToString()
            Else
                Total213 = Total211 - rstD.Tables(0).Rows(0)("OtherIncomeAmt").ToString()
            End If
            Total217 = rstD.Tables(0).Rows(0)("80gAmt").ToString() + rstD.Tables(0).Rows(0)("80ggAmt").ToString() + rstD.Tables(0).Rows(0)("6aamt").ToString()
            Total218 = Total213 - Total217
            Total223 = rstD.Tables(0).Rows(0)("TaxOnIncome").ToString() - (rstD.Tables(0).Rows(0)("88Rebate").ToString() + rstD.Tables(0).Rows(0)("88BRebate").ToString() + rstD.Tables(0).Rows(0)("88crebate").ToString() + IIf((rstD.Tables(0).Rows(0)("88Drebate").ToString()), 0, rstD.Tables(0).Rows(0)("88Drebate").ToString())) + rstD.Tables(0).Rows(0)("SurchargeAmt").ToString()
            Total225 = Total223 - rstD.Tables(0).Rows(0)("89Relief").ToString()
            Total228 = rstD.Tables(0).Rows(0)("TDSAmt").ToString() + rstD.Tables(0).Rows(0)("TDSSurcharge").ToString() + IIf((rstD.Tables(0).Rows(0)("TDSECess").ToString()), 0, rstD.Tables(0).Rows(0)("TDSECess").ToString())
            Total229 = Math.Abs(Total225 - Total228)

            '---
            'esr(RecNo - 1, 0) = RecNo
            'esr(RecNo - 1, 1) = rstD.Tables(0).Rows(0)("DName").ToString()

            'variable drec used, because of error during desing time,
            'too many lines to concatinate.
            DRec = Format(LNo, "000000000") & "DD" & "000000001" & Format(RecNo, "000000000") & Format(RecNo, "000000000") &
      Format(IIf((rstD.Tables(0).Rows(0)("DPan").ToString()), Space(10), IIf(rstD.Tables(0).Rows(0)("DPan").ToString() = vbNullString, Space(10), rstD.Tables(0).Rows(0)("DPan").ToString())), "!@@@@@@@@@@") &
      Format(rstD.Tables(0).Rows(0)("DName").ToString(), "!" & New String("@", 75)) &
      Format(rstD.Tables(0).Rows(0)("FromDt").ToString(), "ddMMyyyy") & Format(rstD.Tables(0).Rows(0)("todt").ToString(), "ddMMyyyy") &
      Format(rstD.Tables(0).Rows(0)("TotalSal").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("TotalRent").ToString() * 100, New String("0", 14)) & Format(rstD.Tables(0).Rows(0)("PFAmt").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("ExPerqAmt").ToString() * 100, New String("0", 14)) & Format((Total209 * 100), New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("Deduction16").ToString() * 100, New String("0", 14)) &
      Format((Total211) * 100, New String("0", 14)) &
      Format(IIf(rstD.Tables(0).Rows(0)("OtherIncomeAmt").ToString() = 0, "P", rstD.Tables(0).Rows(0)("OtherIncomeFlg").ToString()), "@") & Format(rstD.Tables(0).Rows(0)("OtherIncomeAmt").ToString() * 100, New String("0", 13)) &
      Format((Total213) * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("80GAmt").ToString() * 100, New String("0", 14)) & Format(rstD.Tables(0).Rows(0)("80GGAmt").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("6AAmt").ToString() * 100, New String("0", 14)) &
      Format((Total217) * 100, New String("0", 14)) &
      Format((Total218) * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("TaxOnIncome").ToString() * 100, New String("0", 14)) & Format(rstD.Tables(0).Rows(0)("88rebate").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("88Brebate").ToString() * 100, New String("0", 14))
            DRec = DRec & Format(rstD.Tables(0).Rows(0)("88Crebate").ToString() * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("88drebate").ToString()), 0, rstD.Tables(0).Rows(0)("88drebate").ToString()) * 100, New String("0", 14)) &
      Format(Total223 * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("89Relief").ToString() * 100, New String("0", 14)) &
      Format(Total225 * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("TDSAmt").ToString() * 100, New String("0", 14)) & Format(rstD.Tables(0).Rows(0)("TDSSurcharge").ToString() * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("TDSECess").ToString()), 0, rstD.Tables(0).Rows(0)("TDSECess").ToString()) * 100, New String("0", 14)) &
      Format(Total228 * 100, New String("0", 14)) &
      Format(IIf(Total229 = 0, "P", rstD.Tables(0).Rows(0)("RefundFlag").ToString()), "@") &
      Format(Total229 * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("Remark").ToString()), Space(75), rstD.Tables(0).Rows(0)("Remark").ToString()), "!" & New String("@", 75))
            TStrm.WriteLine(DRec)
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            'rstD.MoveNext
        Loop
        'PREQ. DETAIL RECORD
        Dim PreqVal As Double, i As Long, psr As Long
        If rstP.Tables(0).Rows.Count > 0 Then 'rstP.MoveFirst

        End If
        RecNo = 1      'Line no not reset as it will continue from challan detail...
        Do While Not rstP.Tables(0).Rows.Count
            'variable drec used, because of error during desing time,
            'too many lines to concatinate.
            For i = 0 To UBound(esr)
                'If UCase(Trim(rstP.Tables(0).Rows(0)("DName").ToString())) = UCase(Trim(esr(i, 1))) Then
                '    psr = Trim(esr(i, 0))
                '    Exit For
                'End If
            Next i
            DRec = Format(LNo, "000000000") & "PD" & "000000001" & Format(RecNo, "000000000") &
      Format(rstP.Tables(0).Rows(0)("DName").ToString(), "!" & New String("@", 75)) & Format(psr, "000000000") &
      Format(rstP.Tables(0).Rows(0)("UnFurnishAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("AsIfUnFurnishAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("CostOfFurnture").ToString() * 100, New String("0", 14)) &
      Format((rstP.Tables(0).Rows(0)("CostOfFurnture").ToString() * 0.1) * 100, New String("0", 14)) &
      Format((rstP.Tables(0).Rows(0)("AsIfUnFurnishAmt").ToString() + (rstP.Tables(0).Rows(0)("CostOfFurnture").ToString() * 0.1)) * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("RentAmt").ToString() * 100, New String("0", 14))
            If rstP.Tables(0).Rows(0)("UnFurnishAmt").ToString() > 0 Then
                PreqVal = rstP.Tables(0).Rows(0)("UnFurnishAmt").ToString() - rstP.Tables(0).Rows(0)("RentAmt").ToString()
            Else
                PreqVal = ((rstP.Tables(0).Rows(0)("AsIfUnFurnishAmt").ToString() + (rstP.Tables(0).Rows(0)("CostOfFurnture").ToString() * 0.1)) - rstP.Tables(0).Rows(0)("RentAmt").ToString())
            End If
            DRec = DRec & Format(PreqVal * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("ConveyanceAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("SalaryForPersonal").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("TravellingAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("OtherAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("EmployerPFAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("PFInterest").ToString() * 100, New String("0", 14)) &
      Format(((PreqVal + rstP.Tables(0).Rows(0)("ConveyanceAmt").ToString() + rstP.Tables(0).Rows(0)("SalaryForPersonal").ToString() + rstP.Tables(0).Rows(0)("TravellingAmt ").ToString() +
      rstP.Tables(0).Rows(0)("OtherAmt").ToString() + rstP.Tables(0).Rows(0)("EmployerPFAmt").ToString() + rstP.Tables(0).Rows(0)("PFInterest").ToString()) * 100), New String("0", 14))
            TStrm.WriteLine(DRec)
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            'rstP.MoveNext
        Loop
        'CLOSE THE FILE..
        TStrm.Close()
cleanup:
        'AND THE CONNECTIONS ALSO..
        rstC.Dispose()
        rstD.Dispose()
        rstP.Dispose()
        rstRetn.Dispose()
        rstCSum.Dispose()
        rstDSum.Dispose()
        rstC = Nothing
        rstD = Nothing
        rstP = Nothing
        rstRetn = Nothing
        rstCSum = Nothing
        rstDSum = Nothing
        filename = Nothing
        oCoMst = Nothing
    End Sub
    Public Sub Convert27Txt(frmid As String, filename As String, revised As String)
        Dim rstC As New DataSet, rstD As New DataSet
        Dim rstRetn As New DataSet
        Dim rstCSum As New DataSet, rstDSum As New DataSet
        Dim LNo As Long, RecNo As Long
        Dim SumC As Double, SumD As Double
        If frmid <> "F27" And frmid <> "F28" And frmid <> "F29" And frmid <> "F30" Then
            MsgBox("Wrong Parameter, Call JAK Infosolutions P Ltd", vbCritical)
            GoTo cleanup
        Else
            'Parameter ok, check if return exists..
            rstRetn = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " AND FrmType = " & Right(frmid, 2)) ', Cnn
            If rstRetn.Tables(0).Rows.Count <= 0 Then
                Call MsgBox("There is no data for this return form.  Kindly create" _
            & vbCrLf & "this return by using Data entry option and then" _
            & vbCrLf & "create e-TDS file using this option." _
            , vbExclamation + vbDefaultButton1, "RETURN NOT FOUND")
                GoTo cleanup
            End If
        End If
        oCoMst = oCoMst.FetchCo(selectedcoid)
        eTDSPath = Application.StartupPath & "\e-TDS Files"
        Dim FrmNo As String
        Select Case frmid
            Case "F27"
                FrmNo = "F27-Q1"
            Case "F28"
                FrmNo = "F27-Q2"
            Case "F29"
                FrmNo = "F27-Q3"
            Case "F30"
                FrmNo = "F27-Q4"
        End Select
        eFileName = eTDSPath & "\" & FrmNo & "-" & oCoMst.CoName & ".TXT"
        'Do the conversion
        ' TStrm = fso.CreateTextFile(filename, True)
        Dim TStrm As New System.IO.StreamWriter(eFileName, True)
        'Get the related data..
        rstC = FetchDataSet("SELECT * FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn      'Challan Records
        rstD = FetchDataSet("SELECT DT.*, DM.* FROM Deductee27 AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" &
      rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn   'Deductee Records
        rstCSum = FetchDataSet("SELECT sum(Amt) as TotC FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn      'Challan Records
        rstDSum = FetchDataSet("SELECT sum(AmtOfTDS) as TotD FROM Deductee27 WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn  'Deductee Records
        'START WRITING THE TEXT FILE NOW.
        'FILE HEADER RECORD...COMMON FOR F27E AND F27
        If rstD.Tables(0).Rows.Count = 0 And rstC.Tables(0).Rows.Count = 0 Then
            'Header when deductee record count is zero..
            TStrm.WriteLine("000000001FHXNS" & revised & Format(Today(), "ddMMyyyy") & "000000001" & oCoMst.CoTAN & "000000001")
        ElseIf rstD.Tables(0).Rows.Count > 999999999 Or rstC.Tables(0).Rows.Count > 999999999 Then
            'Header when deductee records exceeds 999999999
            TStrm.WriteLine("000000001FHENS" & revised & Format(Today(), "ddMMyyyy") & "000000001" & oCoMst.CoTAN & "000000001")
        Else
            'Normal header
            TStrm.WriteLine("000000001FHNS1" & revised & Format(Today(), "ddMMyyyy") & "000000001" & oCoMst.CoTAN & "000000001")
        End If
        If String.IsNullOrEmpty(rstCSum.Tables(0).Rows(0)("totc").ToString()) = True Then
            SumC = 0
        Else
            SumC = CDbl(rstCSum.Tables(0).Rows(0)("totc").ToString() * 100)
        End If
        If String.IsNullOrEmpty(rstDSum.Tables(0).Rows(0)("totd").ToString()) = True Then
            SumD = 0
        Else
            SumD = CDbl(rstDSum.Tables(0).Rows(0)("totd").ToString() * 100)
        End If
        'BATCH HEADER RECORD.
        TStrm.WriteLine("000000002BH000000001" & Format(IIf(rstC.Tables(0).Rows.Count > 999999999, "999999999", rstC.Tables(0).Rows.Count), "000000000") &
   Format(IIf(rstD.Tables(0).Rows.Count > 999999999, "999999999", rstD.Tables(0).Rows.Count), "000000000") & Format("27", "!@@@@") & Space(8) &
   oCoMst.CoTAN & oCoMst.CoPAN & Left(AY, 4) & Right(AY, 2) & Left(FY, 4) & Right(FY, 2) &
   Format(oCoMst.CoName, "!" & New String("@", 75)) & Format(oCoMst.CoAdd1, "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd2), Space(25), IIf(oCoMst.CoAdd2 = vbNullString, Space(25), oCoMst.CoAdd2)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd3), Space(25), IIf(oCoMst.CoAdd3 = vbNullString, Space(25), oCoMst.CoAdd3)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd4), Space(25), IIf(oCoMst.CoAdd4 = vbNullString, Space(25), oCoMst.CoAdd4)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd5), Space(25), IIf(oCoMst.CoAdd5 = vbNullString, Space(25), oCoMst.CoAdd5)), "!" & New String("@", 25)) &
   Format(oCoMst.CoStateID, "00") & Format(oCoMst.CoPin, "000000") &
   IIf(oCoMst.IsCoAddChg = True, "Y", "N") & oCoMst.CoStatus &
   Format(Right(FrmNo, 2), "!@@") & Format(oCoMst.PR27Ename, "!" & New String("@", 75)) &
   Format(oCoMst.PR27EDesg, "!" & New String("@", 20)) &
   Format(SumC, New String("0", 14)) &
   Format(SumD, New String("0", 14)) & Format(0, New String("0", 14)) & Space(10) & Format(0, New String("0", 14)))
        'CHALLAN DETAIL RECORD
        If rstC.Tables(0).Rows.Count > 0 Then
            'rstC.MoveFirst
        End If
        LNo = 3 : RecNo = 1
        Do While Not rstC.Tables(0).Rows.Count
            TStrm.WriteLine(Format(LNo, "000000000") & "CD" & "000000001" & Format(RecNo, "000000000") & Format(rstC.Tables(0).Rows(0)("Sec").ToString(), "!@@@@@") &
      Format(CLng(rstC.Tables(0).Rows(0)("amt").ToString() * 100), New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstC.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstC.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      Format(rstC.Tables(0).Rows(0)("DtOfVoucher").ToString(), "ddMMyyyy") &
      Format(IIf((rstC.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstC.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstC.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@"))
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            'rstC.MoveNext
        Loop
        'DEDUCTEE DETAIL RECORD
        If rstD.Tables(0).Rows.Count > 0 Then
            'rstD.MoveFirst
        End If
        RecNo = 1      'Line no not reset as it will continue from challan detail...
        Do While Not rstD.Tables(0).Rows.Count
            TStrm.WriteLine(Format(LNo, "000000000") & "DD" & "000000001" & Format(RecNo, "000000000") & Format(rstD.Tables(0).Rows(0)("Sec").ToString(), "!@@@@@") &
      Format(IIf(rstD.Tables(0).Rows(0)("DType").ToString() = "O", 2, 1), "00") &
      Format(IIf((rstD.Tables(0).Rows(0)("DPan").ToString()), Space(10), IIf(rstD.Tables(0).Rows(0)("DPan").ToString() = vbNullString, Space(10), rstD.Tables(0).Rows(0)("DPan").ToString())), "!@@@@@@@@@@") &
      Format(rstD.Tables(0).Rows(0)("DName").ToString(), "!" & New String("@", 75)) & Format(rstD.Tables(0).Rows(0)("DAdd1").ToString(), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd2").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd2").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd2").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd3").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd3").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd3").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd4").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd4").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd4").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd5").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd5").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd5").ToString())), "!" & New String("@", 25)) &
      Format(rstD.Tables(0).Rows(0)("DState").ToString(), "00") & Format(rstD.Tables(0).Rows(0)("DPin").ToString(), "000000") & Format(rstD.Tables(0).Rows(0)("AmtOfPay").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("DtOfPay").ToString(), "ddMMyyyy") & Format(rstD.Tables(0).Rows(0)("RateOfTDS").ToString() * 100, "0000") & "N" &
      Format(rstD.Tables(0).Rows(0)("AmtOfTDS").ToString() * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstD.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstD.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstD.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstD.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("CertificateDt").ToString()), Space(8), rstD.Tables(0).Rows(0)("CertificateDt").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("Reason").ToString()), Space(1), IIf(rstD.Tables(0).Rows(0)("Reason").ToString() = vbNullString, Space(1), rstD.Tables(0).Rows(0)("Reason").ToString())), "@") &
    New String("0", 14))
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            'rstD.MoveNext
        Loop
        TStrm.Close()
cleanup:
        'CLOSE THE FILE..
        'AND THE CONNECTIONS ALSO..
        rstC.Dispose()
        rstD.Dispose()
        ' rstP.Dispose()
        rstRetn.Dispose()
        rstCSum.Dispose()
        rstDSum.Dispose()
        rstC = Nothing
        rstD = Nothing
        rstRetn = Nothing
        rstCSum = Nothing
        rstDSum = Nothing
        filename = Nothing
        oCoMst = Nothing
    End Sub
End Module